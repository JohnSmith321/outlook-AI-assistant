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
Feature (Advanced): Daily work schedule suggestion.

Analyses today's emails and tasks to suggest an optimised
work schedule for the day with time blocks and priorities.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from typing import List

from ai_client import AIClient
from outlook_client import EmailMessage

_SYSTEM = """Bạn là trợ lý AI lên kế hoạch công việc hàng ngày thông minh.
Dựa trên danh sách email/task được cung cấp, hãy đề xuất lịch làm việc tối ưu cho ngày hôm nay.

Định dạng đầu ra:
📅 KẾ HOẠCH NGÀY [ngày]
════════════════════════

🌅 Buổi sáng (8:00 – 12:00)
┌─────────────────────────────────────────┐
│ 08:00 – 08:30 │ <công việc>             │
│ 08:30 – 09:30 │ <công việc>             │
│ ...                                     │
└─────────────────────────────────────────┘

☀️ Buổi chiều (13:00 – 17:30)
┌─────────────────────────────────────────┐
│ 13:00 – 14:00 │ <công việc>             │
│ ...                                     │
└─────────────────────────────────────────┘

🌙 Ngoài giờ (nếu cần)
• <công việc không khẩn cấp>

💡 Gợi ý:
• <lời khuyên về năng suất, ưu tiên>

Nguyên tắc xếp lịch:
- Công việc quan trọng + phức tạp → sáng sớm (khi não minh mẫn nhất)
- Họp → cố gắng gộp cùng thời điểm
- Email/Reply → 2 lần/ngày (9h và 15h)
- Giữa mỗi khối 90 phút → 15 phút nghỉ
- Không nhồi quá nhiều vào một ngày
"""


@dataclass
class ScheduleResult:
    schedule_text: str
    date: datetime.date

    def display(self) -> str:
        return self.schedule_text


class DailyScheduler:
    def __init__(self, ai: AIClient) -> None:
        self._ai = ai

    def suggest_schedule(
        self,
        emails: List[EmailMessage],
        extra_notes: str = "",
    ) -> ScheduleResult:
        """
        Generate a daily schedule suggestion based on today's emails.

        Args:
            emails: List of today's (or recent) emails to process.
            extra_notes: Free-text notes from the user (existing meetings, deadlines, etc.)
        """
        today = datetime.date.today()
        now = datetime.datetime.now()

        # Build email digest
        email_lines = []
        for i, e in enumerate(emails[:20], 1):  # max 20 emails
            email_lines.append(
                f"{i}. [{e.importance_label}] {e.subject}\n"
                f"   Từ: {e.sender} | {e.received_time.strftime('%H:%M')}\n"
                f"   Tóm tắt: {e.body[:200].replace(chr(10), ' ')}\n"
            )

        email_digest = "\n".join(email_lines) if email_lines else "Không có email mới."

        user_prompt = (
            f"Ngày hôm nay: {today.strftime('%A, %d/%m/%Y')}\n"
            f"Giờ hiện tại: {now.strftime('%H:%M')}\n\n"
            f"DANH SÁCH EMAIL/TASK CẦN XỬ LÝ:\n"
            f"{email_digest}\n"
        )
        if extra_notes:
            user_prompt += f"\nGHI CHÚ BỔ SUNG:\n{extra_notes}\n"

        schedule = self._ai.chat(system=_SYSTEM, user=user_prompt, stream=True)
        return ScheduleResult(schedule_text=schedule, date=today)


# Monkey-patch EmailMessage with a helper property
def _importance_label(self) -> str:
    return {2: "KHẨN", 1: "Bình thường", 0: "Thấp"}.get(self.importance, "")

EmailMessage.importance_label = property(_importance_label)
