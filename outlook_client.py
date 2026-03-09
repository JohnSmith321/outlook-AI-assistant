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
    # Email reading
    # ------------------------------------------------------------------

    def get_inbox_emails(self, limit: int = 50) -> List[EmailMessage]:
        """Return the most recent *limit* emails from the Inbox."""
        inbox = self._ns.GetDefaultFolder(6)  # olFolderInbox
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)  # descending

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

    def get_thread_emails(self, conversation_topic: str, limit: int = 20) -> EmailThread:
        """Return all emails in a conversation thread by topic."""
        inbox = self._ns.GetDefaultFolder(6)
        items = inbox.Items
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
