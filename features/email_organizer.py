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
Email organizer: reorganize emails by sender domain/year, archive old emails.

Organize logic:
  • Company domain  → Organized / CompanyName / Year
  • Personal domain → Organized / BrandName / SenderName / Year
    (gmail.com, yahoo.com, etc. get per-individual-sender sub-folders)

Newsletter logic:
  • Company domain  → Newsletter / CompanyName
  • Personal domain → Newsletter / BrandName / SenderName

Archive logic:
  • Emails older than ARCHIVE_CUTOFF_YEARS move to an Archive PST:
    Archive root / Year   (flat by year, no domain grouping)
"""

from __future__ import annotations

import datetime
import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from outlook_client import EmailMessage

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

ORGANIZED_ROOT = "Organized"
ARCHIVE_ROOT = "Archive"
ARCHIVE_CUTOFF_YEARS = 2

PST_WARN_GB = 47.0    # show warning at this size
PST_LIMIT_GB = 50.0   # show error at this size

# Well-known personal / generic email providers → brand folder name
_PERSONAL_DOMAINS: Dict[str, str] = {
    "gmail.com": "Gmail",
    "googlemail.com": "Gmail",
    "yahoo.com": "Yahoo",
    "yahoo.co.jp": "Yahoo",
    "yahoo.co.uk": "Yahoo",
    "hotmail.com": "Hotmail",
    "hotmail.co.uk": "Hotmail",
    "hotmail.fr": "Hotmail",
    "outlook.com": "Outlook",
    "live.com": "Live",
    "live.co.uk": "Live",
    "msn.com": "MSN",
    "icloud.com": "iCloud",
    "me.com": "iCloud",
    "mac.com": "iCloud",
    "protonmail.com": "ProtonMail",
    "proton.me": "ProtonMail",
    "yandex.com": "Yandex",
    "yandex.ru": "Yandex",
    "aol.com": "AOL",
    "mail.com": "Mail",
    "zoho.com": "Zoho",
}

# Chars forbidden in Outlook folder names
_INVALID_FOLDER_CHARS = re.compile(r'[\\/:*?"<>|~]')


# ---------------------------------------------------------------------------
# Domain / name helpers
# ---------------------------------------------------------------------------

def extract_domain(sender_email: str) -> str:
    """Extract lowercased domain from an email address."""
    addr = sender_email.strip().lower()
    m = re.search(r"<([^>]+)>", addr)
    if m:
        addr = m.group(1)
    if "@" in addr:
        return addr.split("@", 1)[1].strip()
    return "unknown"


def is_personal_domain(domain: str) -> bool:
    """Return True if the domain is a personal webmail provider."""
    return domain in _PERSONAL_DOMAINS


def domain_to_folder_name(domain: str) -> str:
    """
    Convert domain to a human-readable folder name.
      viettel.vn       → Viettel
      mail.google.com  → Google
      company-xyz.co.uk → Company-Xyz
    Personal providers return their brand name (Gmail, Yahoo, …).
    """
    if domain == "unknown":
        return "Không xác định"
    if domain in _PERSONAL_DOMAINS:
        return _PERSONAL_DOMAINS[domain]

    parts = domain.split(".")
    core = parts[-2] if len(parts) >= 2 else parts[0]
    return "-".join(w.capitalize() for w in core.replace("_", "-").split("-"))


def clean_folder_name(name: str, max_len: int = 60) -> str:
    """Strip chars that Outlook rejects in folder names, truncate to max_len."""
    cleaned = _INVALID_FOLDER_CHARS.sub("", name.strip())
    return cleaned[:max_len] or "Unknown"


# ---------------------------------------------------------------------------
# Path-computing helpers (relative paths, without the root prefix)
# ---------------------------------------------------------------------------

def _get_organize_rel_path(email: EmailMessage) -> Tuple[str, ...]:
    """
    Return relative path tuple under ORGANIZED_ROOT for an email.
      Company : (CompanyName, Year)
      Personal: (BrandName, SenderName, Year)
    """
    domain = extract_domain(email.sender_email)
    org = domain_to_folder_name(domain)
    year = str(email.received_time.year) if email.received_time else "Unknown"

    if is_personal_domain(domain):
        sender = clean_folder_name(email.sender)
        return (org, sender, year)
    return (org, year)


def get_organize_path(email: EmailMessage) -> Tuple[str, ...]:
    """Full path parts (including ORGANIZED_ROOT) for the email organizer."""
    return (ORGANIZED_ROOT,) + _get_organize_rel_path(email)


def get_newsletter_path(email: EmailMessage) -> Tuple[str, ...]:
    """
    Full path parts (starting with 'Newsletter') for the newsletter mover.
      Company : (Newsletter, CompanyName)
      Personal: (Newsletter, BrandName, SenderName)
    """
    domain = extract_domain(email.sender_email)
    org = domain_to_folder_name(domain)

    if is_personal_domain(domain):
        sender = clean_folder_name(email.sender)
        return ("Newsletter", org, sender)
    return ("Newsletter", org)


# ---------------------------------------------------------------------------
# Organize plan
# ---------------------------------------------------------------------------

@dataclass
class OrganizePlan:
    """Groups emails by their target folder path tuple."""
    # key: relative path tuple (without ORGANIZED_ROOT), e.g. ("Viettel", "2025")
    groups: Dict[Tuple[str, ...], List[EmailMessage]] = field(default_factory=dict)

    def total_emails(self) -> int:
        return sum(len(v) for v in self.groups.values())

    def folder_count(self) -> int:
        return len(self.groups)

    def display_preview(self, max_rows: int = 30) -> str:
        if not self.groups:
            return "Không có email nào để tổ chức."
        lines = [
            f"📂 Kế hoạch tổ chức: {self.total_emails()} email → {self.folder_count()} thư mục",
            f"  (Cần xác nhận trước khi chuyển — AI không tự động di chuyển)",
            "",
            f"{'Thư mục đích':<48} {'Số email':>8}",
            "─" * 58,
        ]
        count = 0
        for path_parts, emails in sorted(self.groups.items()):
            full = f"{ORGANIZED_ROOT} / " + " / ".join(path_parts)
            lines.append(f"{full:<48} {len(emails):>8}")
            count += 1
            if count >= max_rows:
                lines.append(f"  ... và {self.folder_count() - max_rows} thư mục khác")
                break
        return "\n".join(lines)


def plan_organization(emails: List[EmailMessage]) -> OrganizePlan:
    """
    Group emails by their target path (pure Python, does not touch Outlook).
    Personal-domain emails get an extra SenderName level.
    """
    groups: Dict[Tuple[str, ...], List[EmailMessage]] = defaultdict(list)
    for email in emails:
        rel = _get_organize_rel_path(email)
        groups[rel].append(email)
    plan = OrganizePlan()
    plan.groups = dict(groups)
    return plan


# ---------------------------------------------------------------------------
# Archive plan
# ---------------------------------------------------------------------------

@dataclass
class ArchivePlan:
    """Groups old emails by year for archiving."""
    groups: Dict[str, List[EmailMessage]] = field(default_factory=dict)  # year → emails
    cutoff_date: datetime.datetime = field(
        default_factory=lambda: datetime.datetime.now()
    )

    def total_emails(self) -> int:
        return sum(len(v) for v in self.groups.values())

    def display_preview(self) -> str:
        if not self.groups:
            return (
                f"Không có email nào trước {self.cutoff_date.strftime('%d/%m/%Y')} "
                f"(giới hạn {ARCHIVE_CUTOFF_YEARS} năm)."
            )
        lines = [
            f"📦 Kế hoạch archive: {self.total_emails()} email "
            f"(trước {self.cutoff_date.strftime('%d/%m/%Y')})",
            f"  (Cần xác nhận — AI không tự động di chuyển)",
            "",
            f"{'Năm':<8} {'Số email':>8}",
            "─" * 20,
        ]
        for year in sorted(self.groups.keys()):
            lines.append(f"{year:<8} {len(self.groups[year]):>8}")
        return "\n".join(lines)


def plan_archive(
    emails: List[EmailMessage],
    cutoff_years: int = ARCHIVE_CUTOFF_YEARS,
) -> ArchivePlan:
    """
    Return emails older than cutoff_years grouped by year.
    Does NOT touch Outlook.
    """
    cutoff = datetime.datetime.now() - datetime.timedelta(days=cutoff_years * 365)
    groups: Dict[str, List[EmailMessage]] = defaultdict(list)

    for email in emails:
        rt = email.received_time
        if rt is None:
            continue
        # Handle both naive and timezone-aware datetimes
        if rt.tzinfo is not None:
            naive = rt.replace(tzinfo=None)
        else:
            naive = rt
        if naive < cutoff:
            groups[str(rt.year)].append(email)

    plan = ArchivePlan(cutoff_date=cutoff)
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
        "\nGợi ý: Vào Outlook → File → Manage Rules & Alerts để tạo rule mới"
        " hoặc điều chỉnh rule hiện có cho email tự động vào đúng thư mục."
    )
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# PST size helpers
# ---------------------------------------------------------------------------

def format_pst_sizes(sizes: List[dict]) -> str:
    """Format PST size info for display."""
    if not sizes:
        return "Không có thông tin kích thước PST."
    lines = ["💾 Kích thước file PST:"]
    for s in sizes:
        bar_filled = int(min(s["size_gb"], PST_LIMIT_GB) / PST_LIMIT_GB * 20)
        bar = "█" * bar_filled + "░" * (20 - bar_filled)
        pct = s["size_gb"] / PST_LIMIT_GB * 100
        warn = " ⚠️" if s["size_gb"] >= PST_WARN_GB else ""
        lines.append(
            f"  {s['name'][:30]:<30} [{bar}] {s['size_gb']:>5.1f} GB ({pct:.0f}%){warn}"
        )
    return "\n".join(lines)
