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
Email organizer: reorganize emails by sender domain and year.

Logic:
  • Extract domain from sender email address
  • Humanise domain to an org/folder name
  • Group by (org_name, year)
  • Target path in Outlook: [Store root] / Organized / OrgName / Year
  • Also reads and displays existing Outlook rules
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from outlook_client import EmailMessage

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Well-known personal / generic email providers → folder name override
_PERSONAL_DOMAINS: Dict[str, str] = {
    "gmail.com": "Gmail",
    "googlemail.com": "Gmail",
    "yahoo.com": "Yahoo",
    "yahoo.co.jp": "Yahoo",
    "hotmail.com": "Hotmail",
    "hotmail.co.uk": "Hotmail",
    "outlook.com": "Outlook",
    "live.com": "Live",
    "msn.com": "MSN",
    "icloud.com": "iCloud",
    "me.com": "iCloud",
    "protonmail.com": "ProtonMail",
    "proton.me": "ProtonMail",
    "yandex.com": "Yandex",
    "yandex.ru": "Yandex",
    "aol.com": "AOL",
    "mail.com": "Mail",
    "zoho.com": "Zoho",
}

# Root folder name under which everything is organised
ORGANIZED_ROOT = "Organized"


# ---------------------------------------------------------------------------
# Domain / name helpers
# ---------------------------------------------------------------------------

def extract_domain(sender_email: str) -> str:
    """Extract lowercased domain from an email address."""
    addr = sender_email.strip().lower()
    # strip display-name wrapper like "Name <addr>"
    m = re.search(r"<([^>]+)>", addr)
    if m:
        addr = m.group(1)
    if "@" in addr:
        return addr.split("@", 1)[1]
    return "unknown"


def domain_to_folder_name(domain: str) -> str:
    """
    Convert a domain to a human-readable folder name.
    E.g.  viettel.vn → Viettel
          mail.google.com → Google
          company-xyz.co.uk → Company-Xyz
    Personal/generic providers use their well-known brand name.
    """
    if domain == "unknown":
        return "Không xác định"
    if domain in _PERSONAL_DOMAINS:
        return _PERSONAL_DOMAINS[domain]

    # Strip leading subdomains: keep only the 2nd-level domain
    parts = domain.split(".")
    if len(parts) >= 2:
        core = parts[-2]          # e.g. "viettel" from "viettel.vn"
    else:
        core = parts[0]

    # Capitalise nicely: "company-xyz" → "Company-Xyz"
    return "-".join(word.capitalize() for word in core.replace("_", "-").split("-"))


# ---------------------------------------------------------------------------
# Planning
# ---------------------------------------------------------------------------

@dataclass
class OrganizePlan:
    """Maps (folder_path, year) → list of emails to move there."""
    # key: (org_name, year_str) → emails
    groups: Dict[Tuple[str, str], List[EmailMessage]] = field(default_factory=dict)

    def total_emails(self) -> int:
        return sum(len(v) for v in self.groups.values())

    def folder_count(self) -> int:
        return len(self.groups)

    def display_preview(self, max_rows: int = 25) -> str:
        if not self.groups:
            return "Không có email nào để tổ chức."
        lines = [
            f"📂 Kế hoạch tổ chức: {self.total_emails()} email → {self.folder_count()} thư mục",
            "",
            f"{'Thư mục đích':<40} {'Số email':>8}",
            "─" * 50,
        ]
        count = 0
        for (org, year), emails in sorted(self.groups.items()):
            path = f"{ORGANIZED_ROOT} / {org} / {year}"
            lines.append(f"{path:<40} {len(emails):>8}")
            count += 1
            if count >= max_rows:
                remaining = self.folder_count() - max_rows
                lines.append(f"  ... và {remaining} thư mục khác")
                break
        return "\n".join(lines)


def plan_organization(emails: List[EmailMessage]) -> OrganizePlan:
    """
    Group emails by (org_name, year).
    Does NOT touch Outlook — pure Python grouping logic.
    """
    plan = OrganizePlan()
    groups: Dict[Tuple[str, str], List[EmailMessage]] = defaultdict(list)

    for email in emails:
        domain = extract_domain(email.sender_email)
        org_name = domain_to_folder_name(domain)
        year = str(email.received_time.year) if email.received_time else "Unknown"
        groups[(org_name, year)].append(email)

    plan.groups = dict(groups)
    return plan


# ---------------------------------------------------------------------------
# Rules display helper
# ---------------------------------------------------------------------------

def format_rules(rules: List[dict]) -> str:
    """Format Outlook rules list for display in the output pane."""
    if not rules:
        return "📋 Không tìm thấy rule Outlook nào trong tài khoản mặc định."

    lines = [f"📋 Danh sách Outlook Rules ({len(rules)} rules):"]
    for r in rules:
        status = "✅" if r.get("enabled") else "⛔"
        name = r.get("name", "(no name)")
        order = r.get("execution_order", "?")
        lines.append(f"  {status} [{order:>2}] {name}")
    lines.append(
        "\nGợi ý: Kiểm tra và điều chỉnh rules trong Outlook → File → Manage Rules & Alerts"
        " để tự động chuyển email mới vào đúng thư mục."
    )
    return "\n".join(lines)
