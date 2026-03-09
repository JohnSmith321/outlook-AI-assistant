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
Feature (Advanced): Rewrite email in professional style.

Rewrites email content in professional Vietnamese or English,
preserving the original intent while improving tone and clarity.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from ai_client import AIClient
from outlook_client import EmailMessage

Language = Literal["vi", "en"]

_SYSTEM_VI = """Bạn là chuyên gia soạn thảo email doanh nghiệp chuyên nghiệp bằng tiếng Việt.
Nhiệm vụ: Viết lại email theo văn phong chuyên nghiệp, lịch sự và rõ ràng.

Yêu cầu:
- Giữ nguyên ý nghĩa và thông tin gốc
- Sử dụng kính ngữ phù hợp (Kính gửi, Trân trọng,...)
- Cấu trúc rõ ràng: mở đầu → nội dung chính → kết thúc
- Không sử dụng ngôn ngữ thông thường, tiếng lóng
- Câu văn ngắn gọn, súc tích
- Kết thúc bằng lời chào trang trọng

Chỉ trả về nội dung email đã viết lại, KHÔNG giải thích hay bình luận.
"""

_SYSTEM_EN = """You are a professional business email writing expert.
Task: Rewrite the provided email in clear, professional, and polished English.

Requirements:
- Preserve the original intent and key information
- Use appropriate formal greetings (Dear Mr./Ms., Sincerely,...)
- Clear structure: opening → main content → closing
- No slang, colloquial expressions, or abbreviations
- Concise sentences with active voice preferred
- End with a professional sign-off

Return ONLY the rewritten email content, without explanation or commentary.
"""


@dataclass
class RewriteResult:
    original_subject: str
    rewritten_body: str
    language: Language

    def display(self) -> str:
        lang_label = "Tiếng Việt" if self.language == "vi" else "English"
        return (
            f"Email gốc: {self.original_subject}\n"
            f"Ngôn ngữ: {lang_label}\n"
            f"{'─'*50}\n"
            f"{self.rewritten_body}"
        )


class EmailRewriter:
    def __init__(self, ai: AIClient) -> None:
        self._ai = ai

    def rewrite(self, email: EmailMessage, language: Language = "vi") -> RewriteResult:
        """Rewrite email content in the specified language."""
        system = _SYSTEM_VI if language == "vi" else _SYSTEM_EN

        if language == "vi":
            user_prompt = (
                f"Hãy viết lại email sau theo văn phong chuyên nghiệp tiếng Việt:\n\n"
                f"Chủ đề gốc: {email.subject}\n"
                f"Người gửi: {email.sender}\n\n"
                f"Nội dung gốc:\n{email.body[:4000]}"
            )
        else:
            user_prompt = (
                f"Please rewrite the following email in professional English:\n\n"
                f"Original Subject: {email.subject}\n"
                f"From: {email.sender}\n\n"
                f"Original Content:\n{email.body[:4000]}"
            )

        rewritten = self._ai.chat(system=system, user=user_prompt, stream=True)
        return RewriteResult(
            original_subject=email.subject,
            rewritten_body=rewritten,
            language=language,
        )
