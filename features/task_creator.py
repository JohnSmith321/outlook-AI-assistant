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
Feature: Create Outlook tasks from email content.

Uses Claude to extract action items from an email and creates
corresponding tasks in Microsoft Outlook.
"""

from __future__ import annotations

import json
import datetime
import re
from dataclasses import dataclass, field
from typing import List, Optional

from ai_client import AIClient
from outlook_client import EmailMessage, OutlookClient, OutlookTask

_SYSTEM = """Bạn là trợ lý AI giúp trích xuất các công việc cần làm (task) từ email doanh nghiệp.
Phân tích email và trả về JSON với cấu trúc sau (KHÔNG thêm markdown hay giải thích):
{
  "tasks": [
    {
      "subject": "<tiêu đề task ngắn gọn, rõ ràng>",
      "body": "<mô tả chi tiết task, bao gồm context từ email>",
      "due_date": "<YYYY-MM-DD hoặc null nếu không xác định>",
      "importance": "<High | Normal | Low>",
      "categories": "<Work | Personal | Finance | IT | HR>"
    }
  ]
}

Quy tắc:
- Chỉ tạo task cho các hành động cụ thể, có thể thực hiện được
- Nếu email không có task nào → trả về "tasks": []
- due_date: suy luận từ context (e.g. "ngày mai", "cuối tuần", "15/3")
- Mỗi task phải độc lập và có đủ thông tin để thực hiện
"""

_IMPORTANCE_MAP = {"High": 2, "Normal": 1, "Low": 0}


@dataclass
class ExtractedTask:
    subject: str
    body: str
    due_date: Optional[datetime.datetime]
    importance: int
    categories: str


@dataclass
class TaskCreationResult:
    extracted: List[ExtractedTask] = field(default_factory=list)
    created_count: int = 0
    errors: List[str] = field(default_factory=list)

    def display(self) -> str:
        if not self.extracted:
            return "Không tìm thấy task nào trong email này."
        lines = [f"Tìm thấy {len(self.extracted)} task:\n"]
        for i, t in enumerate(self.extracted, 1):
            due = t.due_date.strftime("%d/%m/%Y") if t.due_date else "Chưa xác định"
            imp = {2: "🔴 Cao", 1: "🟡 Trung bình", 0: "🟢 Thấp"}.get(t.importance, "")
            lines.append(
                f"  {i}. {t.subject}\n"
                f"     Hạn: {due} | Ưu tiên: {imp} | Danh mục: {t.categories}\n"
                f"     {t.body[:120]}...\n" if len(t.body) > 120 else
                f"  {i}. {t.subject}\n"
                f"     Hạn: {due} | Ưu tiên: {imp} | Danh mục: {t.categories}\n"
                f"     {t.body}\n"
            )
        lines.append(f"\n✅ Đã tạo {self.created_count}/{len(self.extracted)} task trong Outlook.")
        if self.errors:
            lines.append("Lỗi: " + "; ".join(self.errors))
        return "".join(lines)


def _parse_date(date_str: str) -> Optional[datetime.datetime]:
    if not date_str or date_str.lower() == "null":
        return None
    try:
        return datetime.datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        return None


class TaskCreator:
    def __init__(self, ai: AIClient, outlook: OutlookClient) -> None:
        self._ai = ai
        self._outlook = outlook

    def extract_and_create(self, email: EmailMessage) -> TaskCreationResult:
        """Extract tasks from email using Claude and create them in Outlook."""
        user_prompt = (
            f"Người gửi: {email.sender} <{email.sender_email}>\n"
            f"Chủ đề: {email.subject}\n"
            f"Thời gian nhận: {email.received_time.strftime('%Y-%m-%d %H:%M')}\n"
            f"Ngày hiện tại: {datetime.date.today().isoformat()}\n"
            f"Nội dung:\n{email.body[:4000]}"
        )

        raw = self._ai.chat(system=_SYSTEM, user=user_prompt)

        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            data = json.loads(m.group()) if m else {"tasks": []}

        result = TaskCreationResult()
        for t in data.get("tasks", []):
            extracted = ExtractedTask(
                subject=t.get("subject", "Task từ email"),
                body=t.get("body", ""),
                due_date=_parse_date(t.get("due_date")),
                importance=_IMPORTANCE_MAP.get(t.get("importance", "Normal"), 1),
                categories=t.get("categories", "Work"),
            )
            result.extracted.append(extracted)

            # Create in Outlook
            try:
                self._outlook.create_task(
                    OutlookTask(
                        subject=extracted.subject,
                        body=f"[Tạo từ email: {email.subject}]\n\n{extracted.body}",
                        due_date=extracted.due_date,
                        importance=extracted.importance,
                        categories=extracted.categories,
                    )
                )
                result.created_count += 1
            except Exception as exc:
                result.errors.append(str(exc))

        return result
