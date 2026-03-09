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

Classifies each email as 'spam', 'newsletter', or 'normal'.
Results are used by the GUI to offer delete / move actions.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List

from outlook_client import EmailMessage
from ai_client import AIClient

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

    def display(self) -> str:
        lines = [
            f"📊 Kết quả quét {self.total} email:",
            f"  🔴 Spam / Quảng cáo   : {len(self.spam_ids)}",
            f"  📰 Newsletter / Bản tin: {len(self.newsletter_ids)}",
            f"  ✅ Email bình thường   : {len(self.normal_ids)}",
        ]
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
# Classifier
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = """Bạn là bộ lọc email thông minh. Phân loại email thành MỘT trong 3 loại:
- "spam"       : email rác, quảng cáo không liên quan, phishing, offer giả mạo
- "newsletter" : bản tin định kỳ, thông báo từ dịch vụ đã đăng ký, marketing email từ thương hiệu đã biết
- "normal"     : email công việc, cá nhân, thông báo quan trọng cần xử lý

Chỉ trả về JSON, không giải thích thêm:
{"type": "spam|newsletter|normal", "reason": "lý do ngắn (< 20 từ)"}"""


def _classify_one(email: EmailMessage, ai: AIClient) -> str:
    """Return 'spam', 'newsletter', or 'normal' for a single email."""
    user = (
        f"From: {email.sender} <{email.sender_email}>\n"
        f"Subject: {email.subject}\n"
        f"Body (first 400 chars):\n{email.body[:400]}"
    )
    try:
        raw = ai.chat(_SYSTEM_PROMPT, user)
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            data = json.loads(m.group())
            t = data.get("type", "normal")
            if t in ("spam", "newsletter", "normal"):
                return t
    except Exception:
        pass
    return "normal"


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
        Classify all emails. progress_cb(current, total) called after each one.
        Returns a ScanResult with entry_ids grouped by type.
        """
        result = ScanResult(total=len(emails))

        for i, email in enumerate(emails, 1):
            classification = _classify_one(email, self._ai)
            result.classifications[email.entry_id] = classification

            if classification == "spam":
                result.spam_ids.append(email.entry_id)
            elif classification == "newsletter":
                result.newsletter_ids.append(email.entry_id)
            else:
                result.normal_ids.append(email.entry_id)

            if progress_cb:
                progress_cb(i, len(emails))

        return result
