# Outlook AI Assistant
# Copyright (C) 2025  JohnSmith321
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.
"""
Spam / newsletter classifier using Claude AI.

Optimizations:
  - Uses Haiku (fast/cheap model) instead of Opus
  - Batch scan: groups 10 emails per API call
  - Persistent cache: skips emails already classified
  - Truncated body: 400 chars max (enough for spam detection)
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List

from outlook_client import EmailMessage
from ai_client import AIClient
import config

logger = config.get_logger(__name__)

# ---------------------------------------------------------------------------
# Scan cache (persists across sessions)
# ---------------------------------------------------------------------------

_CACHE_PATH: Path = config.SCAN_CACHE_FILE


def _load_cache() -> Dict[str, str]:
    """Load entry_id → label cache from disk."""
    try:
        return json.loads(_CACHE_PATH.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.debug("Cannot load scan cache: %s", exc)
        return {}


def _save_cache(cache: Dict[str, str]) -> None:
    """Persist cache to disk."""
    try:
        _CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:
        logger.warning("Cannot save scan cache: %s", exc)


# ---------------------------------------------------------------------------
# Result container
# ---------------------------------------------------------------------------

@dataclass
class ScanResult:
    total: int = 0
    spam_ids: List[str] = field(default_factory=list)
    newsletter_ids: List[str] = field(default_factory=list)
    normal_ids: List[str] = field(default_factory=list)
    classifications: Dict[str, str] = field(default_factory=dict)  # entry_id → type
    cached_count: int = 0  # how many came from cache

    def display(self) -> str:
        lines = [
            f"📊 Kết quả quét {self.total} email:",
            f"  🔴 Spam / Quảng cáo   : {len(self.spam_ids)}",
            f"  📰 Newsletter / Bản tin: {len(self.newsletter_ids)}",
            f"  ✅ Email bình thường   : {len(self.normal_ids)}",
        ]
        if self.cached_count:
            lines.append(f"  💾 Từ cache (không tốn token): {self.cached_count}")
        if self.spam_ids:
            lines.append("\nSpam được phát hiện:")
            for eid in self.spam_ids[:10]:
                lines.append(f"  • (id={eid[:16]}...)")
        if self.newsletter_ids:
            lines.append("\nNewsletter được phát hiện:")
            for eid in self.newsletter_ids[:10]:
                lines.append(f"  • (id={eid[:16]}...)")
        if len(self.spam_ids) + len(self.newsletter_ids) > 0:
            lines.append(
                "\n➡  Dùng nút [🗑️ Xóa Spam] hoặc [📰 Chuyển Newsletter] để xử lý."
            )
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# Batch classifier (Haiku, 10 emails per call)
# ---------------------------------------------------------------------------

BATCH_SIZE = 10

_SYSTEM_PROMPT = """Bạn là bộ lọc email thông minh. Phân loại TỪNG email thành MỘT trong 3 loại:
- "spam"       : email rác, quảng cáo không liên quan, phishing, offer giả mạo
- "newsletter" : bản tin định kỳ, thông báo từ dịch vụ đã đăng ký, marketing email từ thương hiệu đã biết
- "normal"     : email công việc, cá nhân, thông báo quan trọng cần xử lý

Trả về JSON array, mỗi phần tử tương ứng với email cùng thứ tự:
[{"id": 1, "type": "spam|newsletter|normal"}, ...]
Chỉ trả về JSON, KHÔNG thêm markdown hay giải thích."""


def _format_email_for_batch(idx: int, email: EmailMessage) -> str:
    """Format one email for inclusion in a batch prompt."""
    return (
        f"--- Email {idx} ---\n"
        f"From: {email.sender} <{email.sender_email}>\n"
        f"Subject: {email.subject}\n"
        f"Body: {email.body[:400]}\n"
    )


def _classify_batch(emails: List[EmailMessage], ai: AIClient) -> List[str]:
    """Classify a batch of emails in a single API call using Haiku."""
    user = "\n".join(
        _format_email_for_batch(i + 1, e) for i, e in enumerate(emails)
    )
    try:
        raw = ai.chat_fast(_SYSTEM_PROMPT, user, max_tokens=1024)
        # Extract JSON array
        m = re.search(r"\[.*\]", raw, re.DOTALL)
        if m:
            items = json.loads(m.group())
            results = []
            for item in items:
                t = item.get("type", "normal")
                results.append(t if t in ("spam", "newsletter", "normal") else "normal")
            # Pad if Claude returned fewer items
            while len(results) < len(emails):
                results.append("normal")
            return results[:len(emails)]
        logger.warning("No JSON array found in spam batch response")
    except Exception as exc:
        logger.warning("Batch classify failed, defaulting all to normal: %s", exc)
    return ["normal"] * len(emails)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

class SpamCleaner:
    """Scan a list of emails and classify as spam / newsletter / normal."""

    def __init__(self, ai: AIClient) -> None:
        self._ai = ai

    def scan(
        self,
        emails: List[EmailMessage],
        progress_cb: Callable[[int, int], None] | None = None,
    ) -> ScanResult:
        """
        Classify all emails using batch calls + cache.
        progress_cb(current, total) called after each batch.
        """
        cache = _load_cache()
        result = ScanResult(total=len(emails))

        # Split into cached and uncached
        uncached: List[EmailMessage] = []
        for email in emails:
            if email.entry_id in cache:
                classification = cache[email.entry_id]
                result.classifications[email.entry_id] = classification
                result.cached_count += 1
                if classification == "spam":
                    result.spam_ids.append(email.entry_id)
                elif classification == "newsletter":
                    result.newsletter_ids.append(email.entry_id)
                else:
                    result.normal_ids.append(email.entry_id)
            else:
                uncached.append(email)

        # Report cached progress immediately if nothing left to scan
        if progress_cb and result.cached_count and not uncached:
            progress_cb(result.cached_count, len(emails))

        # Batch classify uncached emails
        done = result.cached_count
        for batch_start in range(0, len(uncached), BATCH_SIZE):
            batch = uncached[batch_start:batch_start + BATCH_SIZE]
            labels = _classify_batch(batch, self._ai)

            for email, label in zip(batch, labels):
                result.classifications[email.entry_id] = label
                cache[email.entry_id] = label

                if label == "spam":
                    result.spam_ids.append(email.entry_id)
                elif label == "newsletter":
                    result.newsletter_ids.append(email.entry_id)
                else:
                    result.normal_ids.append(email.entry_id)

            done += len(batch)
            if progress_cb:
                progress_cb(done, len(emails))

        _save_cache(cache)
        return result
