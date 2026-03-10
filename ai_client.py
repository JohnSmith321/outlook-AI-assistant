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
Claude AI client wrapper.

Provides a single entry-point for all AI calls in the application so that
model name and common parameters are managed in one place.
"""

from __future__ import annotations

import anthropic
import config


class AIClient:
    """Thin wrapper around the Anthropic Python SDK."""

    def __init__(self) -> None:
        self._client = anthropic.Anthropic(api_key=config.get_api_key())
        self._model = config.CLAUDE_MODEL
        self._model_fast = config.CLAUDE_MODEL_FAST
        self._max_tokens = config.MAX_TOKENS

    def chat(
        self,
        system: str,
        user: str,
        max_tokens: int | None = None,
        stream: bool = False,
    ) -> str:
        """
        Send a single user message with a system prompt and return the text reply.

        Args:
            system: System prompt for Claude.
            user: User message content.
            max_tokens: Override default max_tokens if needed.
            stream: If True, stream and accumulate the full response.

        Returns:
            The text content of Claude's reply.
        """
        tokens = max_tokens or self._max_tokens

        if stream:
            with self._client.messages.stream(
                model=self._model,
                max_tokens=tokens,
                system=system,
                messages=[{"role": "user", "content": user}],
            ) as s:
                return s.get_final_message().content[0].text

        response = self._client.messages.create(
            model=self._model,
            max_tokens=tokens,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
        for block in response.content:
            if block.type == "text":
                return block.text
        return ""

    def chat_fast(
        self,
        system: str,
        user: str,
        max_tokens: int | None = None,
    ) -> str:
        """Like chat() but uses the fast/cheap model (Haiku) for simple tasks."""
        tokens = max_tokens or self._max_tokens
        response = self._client.messages.create(
            model=self._model_fast,
            max_tokens=tokens,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
        for block in response.content:
            if block.type == "text":
                return block.text
        return ""
