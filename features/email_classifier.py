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
Feature: Automatic email classification.

Uses Claude to analyse email metadata + body and return:
  - Priority  : Urgent / Normal / Low
  - Category  : Work / Personal / Newsletter / Finance / HR / IT / Other
  - Action    : Read-Only / Reply-Needed / Meeting-Request / Task-Request / Spam
  - Summary   : One-sentence description
"""

from __future__ import annotations

import json
from dataclasses import dataclass

from ai_client import AIClient
from outlook_client import EmailMessage

_SYSTEM = """Bạn là trợ lý AI chuyên phân loại email doanh nghiệp.
Phân tích email được cung cấp và trả về JSON với cấu trúc sau (KHÔNG thêm markdown hay giải thích):
{
  "priority": "<Urgent | Normal | Low>",
  "category": "<Work | Personal | Newsletter | Finance | HR | IT | Other>",
  "action": "<Read-Only | Reply-Needed | Meeting-Request | Task-Request | Spam>",
  "summary": "<một câu tóm tắt ngắn gọn nội dung email bằng tiếng Việt>"
}

Quy tắc:
- Urgent: cần hành động trong 24h, deadline gấp, sự cố hệ thống, lãnh đạo yêu cầu
- Reply-Needed: người gửi cần phản hồi rõ ràng
- Meeting-Request: email mời họp, đề xuất thời gian gặp mặt
- Task-Request: yêu cầu thực hiện công việc, giao nhiệm vụ
"""


@dataclass
class ClassificationResult:
    priority: str
    category: str
    action: str
    summary: str

    def display(self) -> str:
        priority_emoji = {"Urgent": "🔴", "Normal": "🟡", "Low": "🟢"}.get(
            self.priority, "⚪"
        )
        action_emoji = {
            "Reply-Needed": "↩️",
            "Meeting-Request": "📅",
            "Task-Request": "✅",
            "Spam": "🚫",
            "Read-Only": "👁️",
        }.get(self.action, "")
        return (
            f"Mức độ ưu tiên : {priority_emoji} {self.priority}\n"
            f"Danh mục       : {self.category}\n"
            f"Hành động      : {action_emoji} {self.action}\n"
            f"Tóm tắt        : {self.summary}"
        )


class EmailClassifier:
    def __init__(self, ai: AIClient) -> None:
        self._ai = ai

    def classify(self, email: EmailMessage) -> ClassificationResult:
        """Classify a single email and return a ClassificationResult."""
        user_prompt = (
            f"Người gửi: {email.sender} <{email.sender_email}>\n"
            f"Chủ đề: {email.subject}\n"
            f"Thời gian nhận: {email.received_time.strftime('%Y-%m-%d %H:%M')}\n"
            f"Nội dung:\n{email.body[:3000]}"
        )

        raw = self._ai.chat(system=_SYSTEM, user=user_prompt)
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            # Fallback: extract first JSON block
            import re
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            data = json.loads(m.group()) if m else {}

        return ClassificationResult(
            priority=data.get("priority", "Normal"),
            category=data.get("category", "Other"),
            action=data.get("action", "Read-Only"),
            summary=data.get("summary", "Không thể phân tích."),
        )

    def classify_bulk(
        self, emails: list[EmailMessage]
    ) -> list[tuple[EmailMessage, ClassificationResult]]:
        """Classify a list of emails one-by-one. Returns pairs (email, result)."""
        results = []
        for email in emails:
            try:
                result = self.classify(email)
            except Exception as exc:
                result = ClassificationResult(
                    priority="Normal",
                    category="Other",
                    action="Read-Only",
                    summary=f"Lỗi phân tích: {exc}",
                )
            results.append((email, result))
        return results
