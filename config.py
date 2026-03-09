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


# Claude model to use throughout the app
CLAUDE_MODEL = "claude-opus-4-6"

# Max tokens for Claude responses
MAX_TOKENS = 4096

# Number of emails to load from Inbox at startup
EMAIL_LOAD_LIMIT = 50

# Outlook folder constants (olDefaultFolders enum)
OUTLOOK_INBOX = 6
OUTLOOK_TASKS = 13
OUTLOOK_CALENDAR = 9
