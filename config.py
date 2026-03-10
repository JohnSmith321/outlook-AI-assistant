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

import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env file from project root
_env_path = Path(__file__).parent / ".env"
load_dotenv(_env_path)


def get_api_key() -> str:
    """Return the Anthropic API key from environment."""
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        raise ValueError(
            "ANTHROPIC_API_KEY is not set. "
            "Copy .env.example to .env and fill in your key."
        )
    return key


# Claude models
CLAUDE_MODEL = "claude-opus-4-6"            # complex tasks (summarize, rewrite, schedule)
CLAUDE_MODEL_FAST = "claude-haiku-4-5-20251001"  # simple tasks (classify, spam scan)

# Max tokens for Claude responses
MAX_TOKENS = 4096

# Spam scan cache file (persists results across sessions)
SCAN_CACHE_FILE = Path(__file__).parent / ".scan_cache.json"

# Number of emails to load from Inbox at startup
EMAIL_LOAD_LIMIT = 50

# Outlook folder constants (olDefaultFolders enum)
OUTLOOK_INBOX = 6
OUTLOOK_TASKS = 13
OUTLOOK_CALENDAR = 9
