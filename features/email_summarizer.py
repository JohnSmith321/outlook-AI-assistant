"""
Feature (Advanced): Summarize an entire email thread.

Collects all emails in a conversation thread and produces a structured
Vietnamese summary: key points, decisions made, and next actions.
"""

from __future__ import annotations

from dataclasses import dataclass

from ai_client import AIClient
from outlook_client import EmailMessage, EmailThread

_SYSTEM = """Bạn là trợ lý AI chuyên tóm tắt luồng email doanh nghiệp bằng tiếng Việt.
Phân tích toàn bộ chuỗi email theo thứ tự thời gian và trả về bản tóm tắt có cấu trúc.

Định dạng đầu ra:
📋 TÓM TẮT LUỒNG EMAIL
─────────────────────
📌 Chủ đề chính:
<1-2 câu mô tả nội dung tổng quát>

👥 Người tham gia:
<danh sách tên – vai trò nếu rõ>

🔑 Các điểm chính:
• <điểm 1>
• <điểm 2>
• ...

✅ Quyết định / Kết luận:
• <quyết định 1 nếu có>

⏭️ Hành động tiếp theo:
• <hành động 1 – người chịu trách nhiệm nếu rõ>

⚠️ Vấn đề chưa giải quyết (nếu có):
• <vấn đề 1>

Sử dụng ngôn ngữ ngắn gọn, chuyên nghiệp. Tối đa 400 từ.
"""


@dataclass
class SummaryResult:
    summary: str
    email_count: int
    thread_topic: str

    def display(self) -> str:
        header = f"Tổng cộng {self.email_count} email trong luồng: \"{self.thread_topic}\"\n\n"
        return header + self.summary


class EmailSummarizer:
    def __init__(self, ai: AIClient) -> None:
        self._ai = ai

    def summarize_email(self, email: EmailMessage) -> SummaryResult:
        """Summarize a single email."""
        user_prompt = (
            f"EMAIL ĐƠN LẺ:\n"
            f"Từ: {email.sender} <{email.sender_email}>\n"
            f"Chủ đề: {email.subject}\n"
            f"Thời gian: {email.received_time.strftime('%Y-%m-%d %H:%M')}\n"
            f"Nội dung:\n{email.body[:5000]}"
        )
        summary = self._ai.chat(system=_SYSTEM, user=user_prompt, stream=True)
        return SummaryResult(
            summary=summary,
            email_count=1,
            thread_topic=email.subject,
        )

    def summarize_thread(self, thread: EmailThread) -> SummaryResult:
        """Summarize an entire email thread in chronological order."""
        if not thread.messages:
            return SummaryResult(
                summary="Không có email nào trong luồng này.",
                email_count=0,
                thread_topic=thread.topic,
            )

        # Build combined text, oldest first (thread.messages is already sorted)
        parts = [f"LUỒNG EMAIL: {thread.topic}\n{'='*50}"]
        for i, msg in enumerate(thread.messages, 1):
            parts.append(
                f"\n[Email {i}/{len(thread.messages)}]\n"
                f"Từ: {msg.sender} <{msg.sender_email}>\n"
                f"Thời gian: {msg.received_time.strftime('%Y-%m-%d %H:%M')}\n"
                f"Nội dung:\n{msg.body[:1500]}\n"
                f"{'─'*40}"
            )

        combined = "\n".join(parts)
        # Limit to ~8000 chars to stay within token budget
        combined = combined[:8000]

        summary = self._ai.chat(system=_SYSTEM, user=combined, stream=True)
        return SummaryResult(
            summary=summary,
            email_count=len(thread.messages),
            thread_topic=thread.topic,
        )
