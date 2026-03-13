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
Feature: Create Outlook calendar events from email content.

Uses Claude to detect meeting requests in emails and creates
corresponding appointments/meetings in Microsoft Outlook Calendar.
"""

from __future__ import annotations

import json
import datetime
import re
from dataclasses import dataclass, field
from typing import List, Optional

import config
from ai_client import AIClient
from outlook_client import EmailMessage, OutlookClient, CalendarEvent

logger = config.get_logger(__name__)

_SYSTEM = """Bạn là trợ lý AI giúp phát hiện và trích xuất thông tin cuộc họp từ email doanh nghiệp.
Phân tích email và trả về JSON với cấu trúc sau (KHÔNG thêm markdown hay giải thích):
{
  "has_meeting": true/false,
  "events": [
    {
      "subject": "<tiêu đề cuộc họp>",
      "body": "<mô tả, agenda, nội dung cuộc họp>",
      "start": "<YYYY-MM-DD HH:MM hoặc null>",
      "end": "<YYYY-MM-DD HH:MM hoặc null>",
      "location": "<địa điểm hoặc link meeting hoặc chuỗi rỗng>",
      "required_attendees": "<email1; email2 hoặc chuỗi rỗng>",
      "duration_minutes": <số phút, mặc định 60 nếu không xác định>
    }
  ]
}

Quy tắc:
- Nếu không có thông tin cuộc họp → has_meeting: false, events: []
- Suy luận thời gian từ context ("thứ Hai tuần này", "9h sáng mai", "15/3 lúc 2pm")
- Ngày hiện tại sẽ được cung cấp trong prompt
- Nếu không có giờ kết thúc → thêm duration_minutes vào giờ bắt đầu
- required_attendees: trích xuất từ email To/CC hoặc nội dung
"""


@dataclass
class ExtractedEvent:
    subject: str
    body: str
    start: Optional[datetime.datetime]
    end: Optional[datetime.datetime]
    location: str
    required_attendees: str


@dataclass
class CalendarCreationResult:
    has_meeting: bool = False
    extracted: List[ExtractedEvent] = field(default_factory=list)
    created_count: int = 0
    errors: List[str] = field(default_factory=list)

    def display(self) -> str:
        if not self.has_meeting or not self.extracted:
            return "Không phát hiện thông tin cuộc họp trong email này."
        lines = [f"Phát hiện {len(self.extracted)} cuộc họp:\n"]
        for i, e in enumerate(self.extracted, 1):
            start_str = e.start.strftime("%d/%m/%Y %H:%M") if e.start else "Chưa xác định"
            end_str = e.end.strftime("%H:%M") if e.end else "?"
            lines.append(
                f"  {i}. 📅 {e.subject}\n"
                f"     Thời gian: {start_str} – {end_str}\n"
                f"     Địa điểm: {e.location or 'Chưa xác định'}\n"
                f"     Tham dự: {e.required_attendees or 'Chưa xác định'}\n"
            )
        lines.append(f"\n✅ Đã tạo {self.created_count}/{len(self.extracted)} sự kiện trong Outlook Calendar.")
        if self.errors:
            lines.append("Lỗi: " + "; ".join(self.errors))
        return "".join(lines)


def _parse_dt(dt_str: str) -> Optional[datetime.datetime]:
    if not dt_str or dt_str.lower() == "null":
        return None
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(dt_str, fmt)
        except ValueError:
            continue
    return None


class CalendarCreator:
    def __init__(self, ai: AIClient, outlook: OutlookClient) -> None:
        self._ai = ai
        self._outlook = outlook

    def extract_and_create(self, email: EmailMessage) -> CalendarCreationResult:
        """Detect meeting info in email and create Outlook calendar events."""
        today = datetime.date.today()
        now = datetime.datetime.now()
        received_str = (
            email.received_time.strftime('%Y-%m-%d %H:%M')
            if email.received_time else "N/A"
        )
        user_prompt = (
            f"Người gửi: {email.sender} <{email.sender_email}>\n"
            f"Chủ đề: {email.subject}\n"
            f"Thời gian nhận: {received_str}\n"
            f"Ngày hiện tại: {today.isoformat()} ({today.strftime('%A')})\n"
            f"Giờ hiện tại: {now.strftime('%H:%M')}\n"
            f"Nội dung:\n{email.body[:config.EMAIL_BODY_TRUNCATE]}"
        )

        raw = self._ai.chat(system=_SYSTEM, user=user_prompt)

        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            logger.warning("JSON decode failed for calendar extraction, trying regex fallback")
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            if m:
                data = json.loads(m.group())
            else:
                logger.warning("Regex fallback also failed, no events extracted")
                data = {"has_meeting": False, "events": []}

        result = CalendarCreationResult(has_meeting=data.get("has_meeting", False))

        for ev in data.get("events", []):
            start = _parse_dt(ev.get("start"))
            end = _parse_dt(ev.get("end"))
            if start and not end:
                try:
                    dur = int(ev.get("duration_minutes", 60))
                except (TypeError, ValueError):
                    dur = 60
                end = start + datetime.timedelta(minutes=dur)

            extracted = ExtractedEvent(
                subject=ev.get("subject", "Cuộc họp từ email"),
                body=f"[Tạo từ email: {email.subject}]\n\n{ev.get('body', '')}",
                start=start,
                end=end,
                location=ev.get("location", ""),
                required_attendees=ev.get("required_attendees", ""),
            )
            result.extracted.append(extracted)

            try:
                self._outlook.create_calendar_event(
                    CalendarEvent(
                        subject=extracted.subject,
                        body=extracted.body,
                        start=extracted.start,
                        end=extracted.end,
                        location=extracted.location,
                        required_attendees=extracted.required_attendees,
                    )
                )
                result.created_count += 1
            except Exception as exc:
                result.errors.append(str(exc))

        return result
