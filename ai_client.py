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
AI client wrapper — supports Anthropic Claude and OpenAI-compatible APIs.

Provides a single entry-point for all AI calls in the application so that
model name and common parameters are managed in one place.

Set AI_PROVIDER=anthropic (default) or AI_PROVIDER=openai in .env to switch.
OpenAI mode works with any OpenAI-compatible endpoint: OpenAI, Ollama,
LM Studio, OpenRouter, Groq, etc.
"""

from __future__ import annotations

import config

logger = config.get_logger(__name__)


class AIClient:
    """Unified AI client — delegates to Anthropic or OpenAI backend."""

    def __init__(self) -> None:
        self._provider = config.AI_PROVIDER
        self._max_tokens = config.MAX_TOKENS

        if self._provider == "openai":
            from openai import OpenAI
            self._openai = OpenAI(
                api_key=config.get_api_key(),
                base_url=config.OPENAI_BASE_URL,
            )
            self._model = config.OPENAI_MODEL
            self._model_fast = config.OPENAI_MODEL_FAST or config.OPENAI_MODEL
        else:
            import anthropic
            self._anthropic = anthropic.Anthropic(api_key=config.get_api_key())
            self._model = config.CLAUDE_MODEL
            self._model_fast = config.CLAUDE_MODEL_FAST

    # ------------------------------------------------------------------
    # Public API (unchanged interface for all callers)
    # ------------------------------------------------------------------

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
            system: System prompt.
            user: User message content.
            max_tokens: Override default max_tokens if needed.
            stream: If True, stream and accumulate the full response.

        Returns:
            The text content of the AI reply.
        """
        tokens = max_tokens or self._max_tokens
        logger.debug("chat() provider=%s model=%s max_tokens=%d stream=%s",
                      self._provider, self._model, tokens, stream)

        if self._provider == "openai":
            return self._openai_chat(self._model, system, user, tokens, stream)
        return self._anthropic_chat(self._model, system, user, tokens, stream)

    def chat_fast(
        self,
        system: str,
        user: str,
        max_tokens: int | None = None,
    ) -> str:
        """Like chat() but uses the fast/cheap model for simple tasks."""
        tokens = max_tokens or self._max_tokens
        logger.debug("chat_fast() provider=%s model=%s max_tokens=%d",
                      self._provider, self._model_fast, tokens)

        if self._provider == "openai":
            return self._openai_chat(self._model_fast, system, user, tokens, False)
        return self._anthropic_chat(self._model_fast, system, user, tokens, False)

    # ------------------------------------------------------------------
    # Anthropic backend
    # ------------------------------------------------------------------

    def _anthropic_chat(
        self, model: str, system: str, user: str, tokens: int, stream: bool,
    ) -> str:
        if stream:
            with self._anthropic.messages.stream(
                model=model,
                max_tokens=tokens,
                system=system,
                messages=[{"role": "user", "content": user}],
            ) as s:
                return s.get_final_message().content[0].text

        response = self._anthropic.messages.create(
            model=model,
            max_tokens=tokens,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
        for block in response.content:
            if block.type == "text":
                return block.text
        return ""

    # ------------------------------------------------------------------
    # OpenAI-compatible backend
    # ------------------------------------------------------------------

    def _openai_chat(
        self, model: str, system: str, user: str, tokens: int, stream: bool,
    ) -> str:
        messages = [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ]

        if stream:
            chunks = []
            response = self._openai.chat.completions.create(
                model=model,
                max_tokens=tokens,
                messages=messages,
                stream=True,
            )
            for chunk in response:
                delta = chunk.choices[0].delta
                if delta.content:
                    chunks.append(delta.content)
            return "".join(chunks)

        response = self._openai.chat.completions.create(
            model=model,
            max_tokens=tokens,
            messages=messages,
        )
        return response.choices[0].message.content or ""
