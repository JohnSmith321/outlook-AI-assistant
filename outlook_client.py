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
Outlook COM client using pywin32.

Wraps all Outlook automation so that features only call clean Python methods
without touching COM objects directly.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from typing import List, Optional

try:
    import win32com.client
    import pywintypes
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ---------------------------------------------------------------------------
# Data classes (plain Python – no COM dependency)
# ---------------------------------------------------------------------------

@dataclass
class EmailMessage:
    entry_id: str
    subject: str
    sender: str
    sender_email: str
    received_time: datetime.datetime
    body: str
    conversation_topic: str = ""
    importance: int = 1          # 0=Low, 1=Normal, 2=High
    unread: bool = False


@dataclass
class FolderInfo:
    """Represents an Outlook folder from any PST/store."""
    display_name: str        # e.g. "Inbox"
    full_path: str           # e.g. "Viettel / Inbox / Projects"
    store_name: str          # PST/account display name
    entry_id: str            # folder EntryID for retrieval
    store_id: str            # store EntryID (needed for GetFolderFromID)
    item_count: int = 0      # number of items in folder

    def label(self) -> str:
        """Short label for the dropdown UI."""
        return f"{self.store_name}  ›  {self.full_path}"


@dataclass
class EmailThread:
    topic: str
    messages: List[EmailMessage] = field(default_factory=list)


@dataclass
class OutlookTask:
    subject: str
    body: str = ""
    due_date: Optional[datetime.datetime] = None
    importance: int = 1          # 0=Low, 1=Normal, 2=High
    categories: str = ""


@dataclass
class CalendarEvent:
    subject: str
    body: str = ""
    start: Optional[datetime.datetime] = None
    end: Optional[datetime.datetime] = None
    location: str = ""
    required_attendees: str = ""
    optional_attendees: str = ""


# ---------------------------------------------------------------------------
# Client
# ---------------------------------------------------------------------------

class OutlookClient:
    """High-level Outlook automation client."""

    def __init__(self) -> None:
        if not HAS_WIN32:
            raise RuntimeError(
                "pywin32 is not installed. Run: pip install pywin32"
            )
        self._app = win32com.client.Dispatch("Outlook.Application")
        self._ns = self._app.GetNamespace("MAPI")

    # ------------------------------------------------------------------
    # Folder enumeration (multi-PST / multi-account support)
    # ------------------------------------------------------------------

    def get_all_folders(self, mail_only: bool = True) -> List[FolderInfo]:
        """
        Return a flat list of all folders across every store (PST / account).

        Args:
            mail_only: If True, skip Calendar/Tasks/Contacts folders and
                       only include folders that can hold mail items.
        """
        # olItemType constants for mail folders
        _MAIL_TYPES = {0}   # olMailItem = 0

        folders: List[FolderInfo] = []
        for store in self._ns.Stores:
            try:
                store_name = store.DisplayName
                store_id = store.StoreID
                root = store.GetRootFolder()
                self._recurse_folders(
                    root, store_name, store_id,
                    parent_path="", results=folders,
                    mail_only=mail_only,
                )
            except Exception:
                continue
        return folders

    def _recurse_folders(
        self,
        folder,
        store_name: str,
        store_id: str,
        parent_path: str,
        results: List[FolderInfo],
        mail_only: bool,
    ) -> None:
        try:
            name = folder.Name
            path = f"{parent_path} / {name}" if parent_path else name

            # Filter: only include folders that hold mail (DefaultItemType == 0)
            try:
                item_type = folder.DefaultItemType
            except Exception:
                item_type = -1

            if not mail_only or item_type == 0:
                try:
                    count = folder.Items.Count
                except Exception:
                    count = 0
                results.append(FolderInfo(
                    display_name=name,
                    full_path=path,
                    store_name=store_name,
                    entry_id=folder.EntryID,
                    store_id=store_id,
                    item_count=count,
                ))

            # Recurse into sub-folders
            for sub in folder.Folders:
                self._recurse_folders(
                    sub, store_name, store_id, path, results, mail_only
                )
        except Exception:
            pass

    def get_default_inbox_info(self) -> Optional[FolderInfo]:
        """Return FolderInfo for the default Inbox (fallback)."""
        try:
            inbox = self._ns.GetDefaultFolder(6)
            store = inbox.Store
            return FolderInfo(
                display_name="Inbox",
                full_path="Inbox",
                store_name=store.DisplayName,
                entry_id=inbox.EntryID,
                store_id=store.StoreID,
                item_count=inbox.Items.Count,
            )
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Email reading
    # ------------------------------------------------------------------

    def get_inbox_emails(self, limit: int = 50) -> List[EmailMessage]:
        """Return the most recent *limit* emails from the default Inbox."""
        inbox = self._ns.GetDefaultFolder(6)  # olFolderInbox
        return self._read_folder_items(inbox, limit)

    def get_emails_from_folder(
        self, folder_info: FolderInfo, limit: int = 50
    ) -> List[EmailMessage]:
        """Return the most recent *limit* emails from any folder across any store."""
        try:
            folder = self._ns.GetFolderFromID(folder_info.entry_id, folder_info.store_id)
            return self._read_folder_items(folder, limit)
        except Exception as exc:
            raise RuntimeError(
                f"Cannot read folder '{folder_info.full_path}': {exc}"
            ) from exc

    def _read_folder_items(self, folder, limit: int) -> List[EmailMessage]:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # newest first
        emails: List[EmailMessage] = []
        count = 0
        for item in items:
            if count >= limit:
                break
            try:
                if item.Class != 43:   # 43 = olMail
                    continue
                emails.append(self._mail_item_to_email(item))
                count += 1
            except Exception:
                continue
        return emails

    def get_email_by_id(self, entry_id: str) -> Optional[EmailMessage]:
        """Retrieve a single email by its EntryID."""
        try:
            item = self._ns.GetItemFromID(entry_id)
            return self._mail_item_to_email(item)
        except Exception:
            return None

    def get_thread_emails(
        self,
        conversation_topic: str,
        folder_info: Optional[FolderInfo] = None,
        limit: int = 20,
    ) -> EmailThread:
        """
        Return all emails in a conversation thread by topic.
        Searches the given folder, or the default Inbox if none provided.
        """
        if folder_info:
            try:
                folder = self._ns.GetFolderFromID(
                    folder_info.entry_id, folder_info.store_id
                )
            except Exception:
                folder = self._ns.GetDefaultFolder(6)
        else:
            folder = self._ns.GetDefaultFolder(6)

        items = folder.Items
        items.Sort("[ReceivedTime]", False)  # oldest first

        thread = EmailThread(topic=conversation_topic)
        count = 0
        for item in items:
            if count >= limit:
                break
            try:
                if item.Class != 43:
                    continue
                if item.ConversationTopic == conversation_topic:
                    thread.messages.append(self._mail_item_to_email(item))
                    count += 1
            except Exception:
                continue
        return thread

    def _mail_item_to_email(self, item) -> EmailMessage:
        received = item.ReceivedTime
        if hasattr(received, 'timestamp'):
            dt = received
        else:
            # pywintypes.datetime → Python datetime
            dt = datetime.datetime(
                received.year, received.month, received.day,
                received.hour, received.minute, received.second
            )
        return EmailMessage(
            entry_id=item.EntryID,
            subject=item.Subject or "(no subject)",
            sender=item.SenderName or "",
            sender_email=item.SenderEmailAddress or "",
            received_time=dt,
            body=item.Body or "",
            conversation_topic=item.ConversationTopic or "",
            importance=item.Importance,
            unread=item.UnRead,
        )

    # ------------------------------------------------------------------
    # Task creation
    # ------------------------------------------------------------------

    def create_task(self, task: OutlookTask) -> bool:
        """Create a new task in Outlook Tasks folder. Returns True on success."""
        try:
            item = self._app.CreateItem(3)   # 3 = olTaskItem
            item.Subject = task.subject
            item.Body = task.body
            item.Importance = task.importance
            if task.categories:
                item.Categories = task.categories
            if task.due_date:
                item.DueDate = task.due_date
            item.Save()
            return True
        except Exception as exc:
            raise RuntimeError(f"Failed to create task: {exc}") from exc

    # ------------------------------------------------------------------
    # Calendar / meeting creation
    # ------------------------------------------------------------------

    def create_calendar_event(self, event: CalendarEvent) -> bool:
        """Create a new appointment/meeting in the Calendar. Returns True on success."""
        try:
            item = self._app.CreateItem(1)   # 1 = olAppointmentItem
            item.Subject = event.subject
            item.Body = event.body
            item.Location = event.location
            if event.start:
                item.Start = event.start
            if event.end:
                item.End = event.end
            if event.required_attendees:
                item.RequiredAttendees = event.required_attendees
            if event.optional_attendees:
                item.OptionalAttendees = event.optional_attendees
            item.Save()
            return True
        except Exception as exc:
            raise RuntimeError(f"Failed to create calendar event: {exc}") from exc

    # ------------------------------------------------------------------
    # Email management (delete / move / folder creation)
    # ------------------------------------------------------------------

    def delete_emails(self, entry_ids: List[str]) -> tuple:
        """
        Move emails to Deleted Items (soft delete).
        Returns (success_count, fail_count).
        """
        success, fail = 0, 0
        for entry_id in entry_ids:
            try:
                item = self._ns.GetItemFromID(entry_id)
                item.Delete()
                success += 1
            except Exception:
                fail += 1
        return success, fail

    def move_email(
        self,
        entry_id: str,
        target_entry_id: str,
        target_store_id: str,
    ) -> bool:
        """Move a single email to a target folder. Returns True on success."""
        try:
            item = self._ns.GetItemFromID(entry_id)
            target = self._ns.GetFolderFromID(target_entry_id, target_store_id)
            item.Move(target)
            return True
        except Exception:
            return False

    def move_emails(
        self,
        entry_ids: List[str],
        target_entry_id: str,
        target_store_id: str,
    ) -> tuple:
        """Move multiple emails to a target folder. Returns (success, fail)."""
        success, fail = 0, 0
        for eid in entry_ids:
            if self.move_email(eid, target_entry_id, target_store_id):
                success += 1
            else:
                fail += 1
        return success, fail

    def get_or_create_folder_path(
        self, store_id: str, path_parts: List[str]
    ) -> FolderInfo:
        """
        Get or create a nested folder path under a store's root.

        Args:
            store_id:    StoreID of the target PST / account.
            path_parts:  Folder names to traverse/create, e.g.
                         ['Organized', 'Viettel', '2025']

        Returns FolderInfo for the deepest folder in the path.
        """
        target_store = None
        for store in self._ns.Stores:
            if store.StoreID == store_id:
                target_store = store
                break
        if target_store is None:
            raise RuntimeError(f"Store not found: {store_id[:20]}...")

        current = target_store.GetRootFolder()
        for part in path_parts:
            found = None
            try:
                for sub in current.Folders:
                    if sub.Name == part:
                        found = sub
                        break
            except Exception:
                pass
            current = found if found is not None else current.Folders.Add(part)

        try:
            count = current.Items.Count
        except Exception:
            count = 0

        return FolderInfo(
            display_name=current.Name,
            full_path=" / ".join(path_parts),
            store_name=target_store.DisplayName,
            entry_id=current.EntryID,
            store_id=target_store.StoreID,
            item_count=count,
        )

    def get_newsletter_folder(self, store_id: str) -> FolderInfo:
        """Get or create a 'Newsletter' subfolder directly under the store root."""
        return self.get_or_create_folder_path(store_id, ["Newsletter"])

    # ------------------------------------------------------------------
    # Outlook Rules
    # ------------------------------------------------------------------

    def get_outlook_rules(self) -> List[dict]:
        """
        Return rule info dicts from the default store.
        Each dict has: name, enabled, execution_order.
        """
        rules_info: List[dict] = []
        try:
            rules = self._ns.DefaultStore.GetRules()
            for i in range(1, rules.Count + 1):
                try:
                    rule = rules.Item(i)
                    rules_info.append(
                        {
                            "name": rule.Name,
                            "enabled": bool(rule.Enabled),
                            "execution_order": rule.ExecutionOrder,
                        }
                    )
                except Exception:
                    continue
        except Exception:
            pass
        return rules_info
