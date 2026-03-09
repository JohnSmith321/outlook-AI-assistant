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
