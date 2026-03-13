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
"""Configuration management for Outlook AI Assistant."""

import logging
import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env file from project root
_env_path = Path(__file__).parent / ".env"
load_dotenv(_env_path)


# ---------------------------------------------------------------------------
# AI Provider configuration
# ---------------------------------------------------------------------------

AI_PROVIDER = os.environ.get("AI_PROVIDER", "anthropic").lower()  # "anthropic" or "openai"


def get_api_key() -> str:
    """Return the API key for the configured provider."""
    if AI_PROVIDER == "openai":
        key = os.environ.get("OPENAI_API_KEY", "")
        if not key:
            raise ValueError(
                "OPENAI_API_KEY is not set. "
                "Copy .env.example to .env and fill in your key. "
                "For Ollama, set OPENAI_API_KEY=ollama"
            )
        return key
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        raise ValueError(
            "ANTHROPIC_API_KEY is not set. "
            "Copy .env.example to .env and fill in your key."
        )
    return key


# Model names (overridable via .env)
CLAUDE_MODEL = os.environ.get(
    "CLAUDE_MODEL", "claude-opus-4-6"
)
CLAUDE_MODEL_FAST = os.environ.get(
    "CLAUDE_MODEL_FAST", "claude-haiku-4-5-20251001"
)

# OpenAI-compatible provider settings
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o")
OPENAI_MODEL_FAST = os.environ.get("OPENAI_MODEL_FAST", "")  # empty = same as OPENAI_MODEL

# Max tokens for AI responses
_max_tokens_raw = os.environ.get("MAX_TOKENS", "4096")
try:
    MAX_TOKENS = int(_max_tokens_raw)
    if MAX_TOKENS <= 0:
        raise ValueError("must be positive")
except ValueError:
    MAX_TOKENS = 4096
    logging.warning("Invalid MAX_TOKENS=%r, using default 4096", _max_tokens_raw)

# Max chars of email body sent to AI (shared across all features)
EMAIL_BODY_TRUNCATE = 4000

# Spam scan cache file (persists results across sessions)
SCAN_CACHE_FILE = Path(__file__).parent / ".scan_cache.json"

# Number of emails to load from Inbox at startup
EMAIL_LOAD_LIMIT = 50

# Outlook folder constants (olDefaultFolders enum)
OUTLOOK_INBOX = 6
OUTLOOK_TASKS = 13
OUTLOOK_CALENDAR = 9


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

def get_logger(name: str) -> logging.Logger:
    """Return a named logger with a standard format. Call once per module."""
    logger = logging.getLogger(name)
    if not logger.handlers:
        handler = logging.StreamHandler()
        handler.setFormatter(
            logging.Formatter("[%(levelname)s] %(name)s: %(message)s")
        )
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)
    return logger
