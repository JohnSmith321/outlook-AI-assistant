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
Outlook AI Assistant — Main entry point.

Launches a tkinter GUI that connects to Microsoft Outlook via COM
and provides AI-powered email management features powered by Claude.
"""

from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog, filedialog
import sys
import os

# Add project root to path so sub-modules can import each other
sys.path.insert(0, os.path.dirname(__file__))

import config
from ai_client import AIClient
from outlook_client import OutlookClient, EmailMessage, FolderInfo
from features.email_classifier import EmailClassifier
from features.task_creator import TaskCreator
from features.calendar_creator import CalendarCreator
from features.email_summarizer import EmailSummarizer
from features.email_rewriter import EmailRewriter
from features.scheduler import DailyScheduler
from features.spam_cleaner import SpamCleaner, ScanResult
from features.email_organizer import (
    plan_organization, plan_archive, format_rules, format_pst_sizes,
    get_newsletter_path, get_organize_path,
    ORGANIZED_ROOT, ARCHIVE_ROOT, ARCHIVE_CUTOFF_YEARS,
    PST_WARN_GB, PST_LIMIT_GB,
)


# ---------------------------------------------------------------------------
# Colour / style constants
# ---------------------------------------------------------------------------
BG_DARK = "#1e1e2e"
BG_MID = "#2a2a3e"
BG_PANEL = "#252535"
FG_TEXT = "#cdd6f4"
FG_ACCENT = "#89b4fa"
FG_SUCCESS = "#a6e3a1"
FG_WARN = "#fab387"
FG_ERROR = "#f38ba8"
FONT_MONO = ("Consolas", 10)
FONT_UI = ("Segoe UI", 10)
FONT_TITLE = ("Segoe UI", 12, "bold")


class OutlookAIApp(tk.Tk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Outlook AI Assistant  •  Powered by Claude")
        self.geometry("1300x780")
        self.configure(bg=BG_DARK)
        self.resizable(True, True)

        # --- Services ---
        self._ai: AIClient | None = None
        self._outlook: OutlookClient | None = None
        self._emails: list[EmailMessage] = []
        self._selected_email: EmailMessage | None = None
        self._folders: list[FolderInfo] = []
        self._current_folder: FolderInfo | None = None

        # --- Spam / newsletter scan state ---
        self._scan_result: ScanResult | None = None
        # entry_id → EmailMessage snapshot taken at scan time
        self._scanned_emails: dict[str, EmailMessage] = {}

        # Build UI then boot services
        self._build_ui()
        self.after(100, self._boot)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        # Top bar
        top = tk.Frame(self, bg=BG_DARK)
        top.pack(fill=tk.X, padx=10, pady=(8, 0))

        tk.Label(
            top,
            text="  Outlook AI Assistant",
            font=("Segoe UI", 14, "bold"),
            fg=FG_ACCENT,
            bg=BG_DARK,
        ).pack(side=tk.LEFT)

        self._status_var = tk.StringVar(value="Đang khởi động...")
        tk.Label(
            top,
            textvariable=self._status_var,
            font=FONT_UI,
            fg=FG_WARN,
            bg=BG_DARK,
        ).pack(side=tk.RIGHT, padx=10)

        # Main paned window
        paned = tk.PanedWindow(
            self, orient=tk.HORIZONTAL, bg=BG_DARK, sashwidth=6, sashpad=2
        )
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        # Left panel – email list
        left = tk.Frame(paned, bg=BG_PANEL, bd=1, relief=tk.FLAT)
        self._build_email_list(left)
        paned.add(left, minsize=380)

        # Right panel – detail + output
        right = tk.Frame(paned, bg=BG_DARK)
        self._build_right_panel(right)
        paned.add(right, minsize=500)

        # Bottom action bar
        self._build_action_bar()

    def _build_email_list(self, parent: tk.Frame) -> None:
        hdr = tk.Frame(parent, bg=BG_MID)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=" Thư mục", font=FONT_TITLE, fg=FG_ACCENT, bg=BG_MID).pack(
            side=tk.LEFT, padx=8, pady=6
        )
        tk.Button(
            hdr,
            text="Làm mới",
            command=self._reload_emails,
            font=FONT_UI,
            bg="#313244",
            fg=FG_TEXT,
            bd=0,
            padx=8,
        ).pack(side=tk.RIGHT, padx=6, pady=4)

        # Folder selector dropdown
        folder_frame = tk.Frame(parent, bg=BG_PANEL)
        folder_frame.pack(fill=tk.X, padx=6, pady=(4, 0))
        tk.Label(
            folder_frame, text="📁", font=FONT_UI, bg=BG_PANEL, fg=FG_ACCENT
        ).pack(side=tk.LEFT)
        self._folder_var = tk.StringVar(value="Đang tải thư mục...")
        self._folder_combo = ttk.Combobox(
            folder_frame,
            textvariable=self._folder_var,
            font=FONT_UI,
            state="readonly",
            width=38,
        )
        style = ttk.Style()
        style.configure("TCombobox", fieldbackground="#313244", background="#313244",
                        foreground=FG_TEXT, selectbackground="#45475a")
        self._folder_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
        self._folder_combo.bind("<<ComboboxSelected>>", self._on_folder_change)

        # Search bar
        sf = tk.Frame(parent, bg=BG_PANEL)
        sf.pack(fill=tk.X, padx=6, pady=4)
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._filter_list())
        tk.Entry(
            sf,
            textvariable=self._search_var,
            font=FONT_UI,
            bg="#313244",
            fg=FG_TEXT,
            insertbackground=FG_TEXT,
            bd=0,
        ).pack(fill=tk.X, ipady=4, padx=2)

        # Treeview  (extended = Ctrl/Shift multi-select)
        cols = ("priority", "sender", "subject", "time")
        self._tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="extended")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background=BG_PANEL,
            foreground=FG_TEXT,
            fieldbackground=BG_PANEL,
            rowheight=24,
            font=FONT_UI,
        )
        style.configure("Treeview.Heading", background=BG_MID, foreground=FG_ACCENT, font=FONT_UI)
        style.map("Treeview", background=[("selected", "#45475a")])

        self._tree.heading("priority", text="AI")
        self._tree.heading("sender", text="Người gửi")
        self._tree.heading("subject", text="Chủ đề")
        self._tree.heading("time", text="Thời gian")
        self._tree.column("priority", width=65, minwidth=65, anchor=tk.CENTER)
        self._tree.column("sender", width=120, minwidth=80)
        self._tree.column("subject", width=220, minwidth=100)
        self._tree.column("time", width=85, minwidth=70, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._tree.bind("<<TreeviewSelect>>", self._on_email_select)
        self._tree.tag_configure("unread", foreground="#89dceb", font=("Segoe UI", 10, "bold"))
        self._tree.tag_configure("urgent", foreground=FG_WARN)
        self._tree.tag_configure("pri_urgent", foreground=FG_ERROR)
        self._tree.tag_configure("pri_normal", foreground=FG_TEXT)
        self._tree.tag_configure("pri_low", foreground="#6c7086")

    def _build_right_panel(self, parent: tk.Frame) -> None:
        # Email detail pane
        detail_frame = tk.LabelFrame(
            parent, text=" Chi tiết email ", font=FONT_UI,
            bg=BG_DARK, fg=FG_ACCENT, bd=1
        )
        detail_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 4))

        self._detail_text = scrolledtext.ScrolledText(
            detail_frame, height=10, font=FONT_MONO,
            bg=BG_PANEL, fg=FG_TEXT, insertbackground=FG_TEXT,
            wrap=tk.WORD, state=tk.DISABLED, bd=0
        )
        self._detail_text.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        # AI output pane
        out_frame = tk.LabelFrame(
            parent, text=" Kết quả AI ", font=FONT_UI,
            bg=BG_DARK, fg=FG_ACCENT, bd=1
        )
        out_frame.pack(fill=tk.BOTH, expand=True)

        self._output_text = scrolledtext.ScrolledText(
            out_frame, height=16, font=FONT_MONO,
            bg=BG_PANEL, fg=FG_SUCCESS, insertbackground=FG_TEXT,
            wrap=tk.WORD, state=tk.DISABLED, bd=0
        )
        self._output_text.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

    def _build_action_bar(self) -> None:
        container = tk.Frame(self, bg=BG_MID)
        container.pack(fill=tk.X, padx=10, pady=(0, 8))

        # Row 1 – analysis features
        row1 = tk.Frame(container, bg=BG_MID)
        row1.pack(fill=tk.X)
        row1_buttons = [
            ("Phân loại Email", self._run_classify, "#313244", FG_ACCENT),
            ("★ Phân loại tất cả", self._run_classify_all, "#1e3a5f", FG_ACCENT),
            ("Tạo Task", self._run_create_task, "#313244", FG_SUCCESS),
            ("Tạo Lịch họp", self._run_create_meeting, "#313244", "#cba6f7"),
            ("Tóm tắt Thread", self._run_summarize, "#313244", "#89dceb"),
            ("Viết lại (VI)", lambda: self._run_rewrite("vi"), "#313244", FG_WARN),
            ("Viết lại (EN)", lambda: self._run_rewrite("en"), "#313244", "#f9e2af"),
            ("Gợi ý lịch ngày", self._run_schedule, "#313244", "#a6e3a1"),
        ]
        for label, cmd, bg, fg in row1_buttons:
            tk.Button(
                row1, text=label, command=cmd,
                font=FONT_UI, bg=bg, fg=fg,
                activebackground="#45475a", activeforeground=fg,
                bd=0, padx=10, pady=5, cursor="hand2",
            ).pack(side=tk.LEFT, padx=3, pady=(5, 2))

        # Row 2 – management features
        row2 = tk.Frame(container, bg=BG_MID)
        row2.pack(fill=tk.X)
        row2_buttons = [
            ("🔍 Quét Spam/NL", self._run_spam_scan, "#3b1f1f", "#f38ba8"),
            ("🗑️ Xóa Spam", self._run_delete_spam, "#3b1f1f", FG_ERROR),
            ("📰 Chuyển Newsletter", self._run_move_newsletter, "#1f2d3b", "#89dceb"),
            ("📂 Tổ chức Email", self._run_organize, "#1f3b2a", FG_SUCCESS),
            ("📦 Archive Cũ", self._run_archive, "#2a2a1f", "#f9e2af"),
            ("💾 Kiểm tra PST", self._run_check_pst, "#1f2a2a", FG_ACCENT),
        ]
        for label, cmd, bg, fg in row2_buttons:
            tk.Button(
                row2, text=label, command=cmd,
                font=FONT_UI, bg=bg, fg=fg,
                activebackground="#45475a", activeforeground=fg,
                bd=0, padx=10, pady=5, cursor="hand2",
            ).pack(side=tk.LEFT, padx=3, pady=(2, 5))

        self._progress = ttk.Progressbar(row2, mode="indeterminate", length=120)
        self._progress.pack(side=tk.RIGHT, padx=10, pady=(2, 5))

    # ------------------------------------------------------------------
    # Boot / Initialisation
    # ------------------------------------------------------------------

    def _boot(self) -> None:
        threading.Thread(target=self._init_services, daemon=True).start()

    def _init_services(self) -> None:
        self._set_status("Đang kết nối Outlook...", FG_WARN)
        try:
            self._outlook = OutlookClient()
        except Exception as exc:
            self._set_status(f"Lỗi Outlook: {exc}", FG_ERROR)
            return

        self._set_status("Đang kết nối Claude AI...", FG_WARN)
        try:
            self._ai = AIClient()
        except ValueError as exc:
            self._set_status(str(exc), FG_ERROR)
            return

        self._set_status("Đang tải danh sách thư mục...", FG_WARN)
        threading.Thread(target=self._load_folders_thread, daemon=True).start()

    def _load_folders_thread(self) -> None:
        try:
            folders = self._outlook.get_all_folders(mail_only=True)
            # Put Inbox folders first
            inbox_folders = [f for f in folders if "inbox" in f.display_name.lower()
                             or "hộp thư đến" in f.display_name.lower()
                             or f.display_name.lower() in ("inbox", "hộp thư đến")]
            other_folders = [f for f in folders if f not in inbox_folders]
            self._folders = inbox_folders + other_folders

            labels = [f.label() for f in self._folders]
            self.after(0, self._populate_folder_combo, labels)

            # Auto-select default inbox
            default = self._outlook.get_default_inbox_info()
            if default and self._folders:
                # Find matching folder by entry_id
                match = next(
                    (f for f in self._folders if f.entry_id == default.entry_id),
                    self._folders[0],
                )
                self._current_folder = match
                self.after(0, self._folder_var.set, match.label())

            self._set_status("Sẵn sàng", FG_SUCCESS)
            self._reload_emails()
        except Exception as exc:
            self._set_status(f"Lỗi tải thư mục: {exc}", FG_ERROR)
            # Fallback: just load default inbox
            self._reload_emails()

    def _populate_folder_combo(self, labels: list[str]) -> None:
        self._folder_combo["values"] = labels

    def _on_folder_change(self, _event=None) -> None:
        selected_label = self._folder_var.get()
        folder = next((f for f in self._folders if f.label() == selected_label), None)
        if folder and folder != self._current_folder:
            self._current_folder = folder
            self._reload_emails()

    def _reload_emails(self) -> None:
        if not self._outlook:
            return
        threading.Thread(target=self._load_emails_thread, daemon=True).start()

    def _load_emails_thread(self) -> None:
        self._set_status("Đang tải email...", FG_WARN)
        try:
            if self._current_folder:
                self._emails = self._outlook.get_emails_from_folder(
                    self._current_folder, limit=config.EMAIL_LOAD_LIMIT
                )
                folder_label = self._current_folder.display_name
            else:
                self._emails = self._outlook.get_inbox_emails(limit=config.EMAIL_LOAD_LIMIT)
                folder_label = "Inbox"
            self.after(0, self._populate_list, self._emails)
            self._set_status(
                f"[{folder_label}] {len(self._emails)} email", FG_SUCCESS
            )
            # Passive PST size check on every folder load
            self._passive_pst_check()
        except Exception as exc:
            self._set_status(f"Lỗi tải email: {exc}", FG_ERROR)

    def _populate_list(self, emails: list[EmailMessage]) -> None:
        # Preserve existing priority labels across redraws
        existing = {
            self._tree.item(iid)["values"][0]: iid
            for iid in self._tree.get_children()
        } if self._tree.get_children() else {}
        prior_map: dict[str, str] = {
            iid: self._tree.set(iid, "priority")
            for iid in self._tree.get_children()
        }

        self._tree.delete(*self._tree.get_children())
        for e in emails:
            tag = "unread" if e.unread else ("urgent" if e.importance == 2 else "")
            priority_label = prior_map.get(e.entry_id, "")
            self._tree.insert(
                "", tk.END,
                iid=e.entry_id,
                values=(
                    priority_label,
                    e.sender[:22],
                    e.subject[:55],
                    e.received_time.strftime("%d/%m %H:%M"),
                ),
                tags=(tag,),
            )

    def _filter_list(self) -> None:
        q = self._search_var.get().lower()
        filtered = [
            e for e in self._emails
            if q in e.sender.lower() or q in e.subject.lower()
        ] if q else self._emails
        self._populate_list(filtered)

    # ------------------------------------------------------------------
    # Email selection
    # ------------------------------------------------------------------

    def _get_selected_emails(self) -> list[EmailMessage]:
        """Return EmailMessage objects for all currently selected rows."""
        sel = self._tree.selection()
        return [e for e in self._emails if e.entry_id in sel]

    def _on_email_select(self, _event=None) -> None:
        emails = self._get_selected_emails()
        if not emails:
            return
        self._selected_email = emails[0]
        if len(emails) == 1:
            self._show_email_detail(emails[0])
        else:
            self._write_detail(
                f"Đã chọn {len(emails)} email:\n"
                + "\n".join(f"  • {e.subject[:70]}" for e in emails)
            )

    def _show_email_detail(self, email: EmailMessage) -> None:
        text = (
            f"Từ     : {email.sender} <{email.sender_email}>\n"
            f"Chủ đề : {email.subject}\n"
            f"Nhận   : {email.received_time.strftime('%A, %d/%m/%Y %H:%M')}\n"
            f"{'─'*60}\n"
            f"{email.body[:2000]}"
        )
        self._write_detail(text)

    # ------------------------------------------------------------------
    # Feature runners  (each spawns a background thread)
    # ------------------------------------------------------------------

    def _guard(self, require_single: bool = False) -> bool:
        """Return True if services are ready and at least one email is selected."""
        if not self._ai or not self._outlook:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ chưa khởi tạo xong.")
            return False
        if not self._get_selected_emails():
            messagebox.showinfo("Chọn email", "Vui lòng chọn ít nhất một email.")
            return False
        if require_single and len(self._get_selected_emails()) > 1:
            messagebox.showinfo("Chọn 1 email", "Tính năng này chỉ xử lý 1 email. Vui lòng chọn 1 email.")
            return False
        return True

    def _run_classify(self) -> None:
        if not self._guard(require_single=True):
            return
        self._run_in_thread(self._classify_thread)

    def _classify_thread(self) -> None:
        self._set_status("Đang phân loại email...", FG_WARN)
        try:
            clf = EmailClassifier(self._ai)
            result = clf.classify(self._selected_email)
            self._update_tree_priority(self._selected_email.entry_id, result)
            self._write_output(result.display())
            self._set_status("Phân loại xong", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi: {exc}", error=True)
            self._set_status("Lỗi phân loại", FG_ERROR)

    def _run_classify_all(self) -> None:
        if not self._ai:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ AI chưa sẵn sàng.")
            return
        if not self._emails:
            messagebox.showinfo("Không có email", "Danh sách email rỗng.")
            return
        self._run_in_thread(self._classify_all_thread)

    def _classify_all_thread(self) -> None:
        emails = self._emails
        total = len(emails)
        clf = EmailClassifier(self._ai)
        lines = []
        for i, email in enumerate(emails, 1):
            self._set_status(f"Đang phân loại {i}/{total}...", FG_WARN)
            try:
                result = clf.classify(email)
                self._update_tree_priority(email.entry_id, result)
                lines.append(
                    f"{i:>2}. [{result.priority:<6}] {email.subject[:50]}\n"
                    f"      → {result.action} | {result.category} | {result.summary[:60]}"
                )
            except Exception as exc:
                lines.append(f"{i:>2}. [LỖI] {email.subject[:50]} — {exc}")
        self._write_output(
            f"✅ Đã phân loại {total} email:\n{'─'*60}\n" + "\n".join(lines)
        )
        self._set_status(f"Phân loại xong {total} email", FG_SUCCESS)

    def _update_tree_priority(self, entry_id: str, result) -> None:
        """Update the priority column cell in the treeview (thread-safe via after)."""
        label_map = {"Urgent": "🔴 Urgent", "Normal": "🟡 Normal", "Low": "🟢 Low"}
        label = label_map.get(result.priority, result.priority)

        def _do():
            try:
                self._tree.set(entry_id, "priority", label)
                # Update row colour tag
                existing_tags = list(self._tree.item(entry_id, "tags"))
                for t in ("pri_urgent", "pri_normal", "pri_low"):
                    if t in existing_tags:
                        existing_tags.remove(t)
                tag = {"Urgent": "pri_urgent", "Normal": "pri_normal", "Low": "pri_low"}.get(
                    result.priority, ""
                )
                if tag:
                    existing_tags.append(tag)
                self._tree.item(entry_id, tags=existing_tags)
            except Exception:
                pass
        self.after(0, _do)

    def _run_create_task(self) -> None:
        if not self._guard():
            return
        self._run_in_thread(self._create_task_thread)

    def _create_task_thread(self) -> None:
        emails = self._get_selected_emails()
        creator = TaskCreator(self._ai, self._outlook)
        total_created, lines = 0, []
        for i, email in enumerate(emails, 1):
            if len(emails) > 1:
                self._set_status(f"Đang tạo task {i}/{len(emails)}...", FG_WARN)
            else:
                self._set_status("Đang tạo task...", FG_WARN)
            try:
                result = creator.extract_and_create(email)
                total_created += result.created_count
                lines.append(f"📧 {email.subject[:55]}\n{result.display()}")
            except Exception as exc:
                lines.append(f"📧 {email.subject[:55]}\n  Lỗi: {exc}")
        self._write_output("\n\n".join(lines))
        self._set_status(f"Đã tạo {total_created} task từ {len(emails)} email", FG_SUCCESS)

    def _run_create_meeting(self) -> None:
        if not self._guard():
            return
        self._run_in_thread(self._create_meeting_thread)

    def _create_meeting_thread(self) -> None:
        emails = self._get_selected_emails()
        creator = CalendarCreator(self._ai, self._outlook)
        total_created, lines = 0, []
        for i, email in enumerate(emails, 1):
            if len(emails) > 1:
                self._set_status(f"Đang tạo lịch {i}/{len(emails)}...", FG_WARN)
            else:
                self._set_status("Đang tạo lịch họp...", FG_WARN)
            try:
                result = creator.extract_and_create(email)
                total_created += result.created_count
                lines.append(f"📧 {email.subject[:55]}\n{result.display()}")
            except Exception as exc:
                lines.append(f"📧 {email.subject[:55]}\n  Lỗi: {exc}")
        self._write_output("\n\n".join(lines))
        self._set_status(f"Đã tạo {total_created} sự kiện từ {len(emails)} email", FG_SUCCESS)

    def _run_summarize(self) -> None:
        if not self._guard(require_single=True):
            return
        self._run_in_thread(self._summarize_thread)

    def _summarize_thread(self) -> None:
        self._set_status("Đang tóm tắt luồng email...", FG_WARN)
        try:
            summarizer = EmailSummarizer(self._ai)
            email = self._selected_email

            # Try to get thread from current folder; fall back to single email
            if email.conversation_topic and self._outlook:
                thread = self._outlook.get_thread_emails(
                    email.conversation_topic, folder_info=self._current_folder
                )
                if len(thread.messages) > 1:
                    result = summarizer.summarize_thread(thread)
                else:
                    result = summarizer.summarize_email(email)
            else:
                result = summarizer.summarize_email(email)

            self._write_output(result.display())
            self._set_status("Tóm tắt xong", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi: {exc}", error=True)
            self._set_status("Lỗi tóm tắt", FG_ERROR)

    def _run_rewrite(self, lang: str) -> None:
        if not self._guard(require_single=True):
            return
        self._run_in_thread(lambda: self._rewrite_thread(lang))

    def _rewrite_thread(self, lang: str) -> None:
        lang_label = "tiếng Việt" if lang == "vi" else "English"
        self._set_status(f"Đang viết lại ({lang_label})...", FG_WARN)
        try:
            rewriter = EmailRewriter(self._ai)
            result = rewriter.rewrite(self._selected_email, language=lang)
            self._write_output(result.display())
            self._set_status("Viết lại xong", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi: {exc}", error=True)
            self._set_status("Lỗi viết lại", FG_ERROR)

    def _run_schedule(self) -> None:
        if not self._ai:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ AI chưa sẵn sàng.")
            return

        notes = simpledialog.askstring(
            "Gợi ý lịch ngày",
            "Nhập ghi chú bổ sung (họp cố định, deadline,...):\n(Để trống nếu không có)",
            parent=self,
        )
        self._run_in_thread(lambda: self._schedule_thread(notes or ""))

    def _schedule_thread(self, notes: str) -> None:
        self._set_status("Đang lên kế hoạch ngày...", FG_WARN)
        try:
            scheduler = DailyScheduler(self._ai)
            result = scheduler.suggest_schedule(self._emails[:20], extra_notes=notes)
            self._write_output(result.display())
            self._set_status("Đã lên kế hoạch xong", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi: {exc}", error=True)
            self._set_status("Lỗi lên kế hoạch", FG_ERROR)

    # ------------------------------------------------------------------
    # Spam / Newsletter management
    # ------------------------------------------------------------------

    def _run_spam_scan(self) -> None:
        if not self._ai or not self._outlook:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ chưa khởi tạo xong.")
            return
        if not self._emails:
            messagebox.showinfo("Không có email", "Danh sách email rỗng.")
            return
        self._run_in_thread(self._spam_scan_thread)

    def _spam_scan_thread(self) -> None:
        emails = list(self._emails)
        cleaner = SpamCleaner(self._ai)

        def on_progress(current: int, _total: int) -> None:
            self._set_status(f"Đang quét {current}/{_total}...", FG_WARN)

        try:
            result = cleaner.scan(emails, progress_cb=on_progress)
            self._scan_result = result
            # Snapshot email objects for later use by newsletter mover
            self._scanned_emails = {e.entry_id: e for e in emails}
            self._write_output(result.display())
            self._set_status(
                f"Quét xong: {len(result.spam_ids)} spam, "
                f"{len(result.newsletter_ids)} newsletter",
                FG_SUCCESS,
            )
        except Exception as exc:
            self._write_output(f"Lỗi quét: {exc}", error=True)
            self._set_status("Lỗi quét", FG_ERROR)

    def _run_delete_spam(self) -> None:
        if not self._outlook:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ chưa sẵn sàng.")
            return
        if not self._scan_result or not self._scan_result.spam_ids:
            messagebox.showinfo(
                "Chưa quét",
                "Chưa có kết quả quét.\nHãy nhấn [🔍 Quét Spam/NL] trước.",
            )
            return
        count = len(self._scan_result.spam_ids)
        if not messagebox.askyesno(
            "Xác nhận xóa Spam",
            f"Sẽ chuyển {count} email spam vào Deleted Items.\n\nTiếp tục?",
        ):
            return
        self._run_in_thread(self._delete_spam_thread)

    def _delete_spam_thread(self) -> None:
        self._set_status("Đang xóa spam...", FG_WARN)
        try:
            ids = list(self._scan_result.spam_ids)
            success, fail = self._outlook.delete_emails(ids)
            # Remove deleted from local list & tree
            deleted_set = set(ids[:success])
            self._emails = [e for e in self._emails if e.entry_id not in deleted_set]
            self.after(0, self._populate_list, self._emails)
            self._scan_result.spam_ids.clear()
            self._write_output(
                f"🗑️ Đã xóa {success} email spam vào Deleted Items."
                + (f"\n  ({fail} lỗi, không xóa được.)" if fail else "")
            )
            self._set_status(f"Đã xóa {success} spam", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi xóa spam: {exc}", error=True)
            self._set_status("Lỗi xóa spam", FG_ERROR)

    def _run_move_newsletter(self) -> None:
        if not self._outlook:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ chưa sẵn sàng.")
            return
        if not self._scan_result or not self._scan_result.newsletter_ids:
            messagebox.showinfo(
                "Chưa quét",
                "Chưa có kết quả quét.\nHãy nhấn [🔍 Quét Spam/NL] trước.",
            )
            return
        count = len(self._scan_result.newsletter_ids)
        if not messagebox.askyesno(
            "Chuyển Newsletter",
            f"Sẽ chuyển {count} email newsletter vào thư mục 'Newsletter'.\n\nTiếp tục?",
        ):
            return
        self._run_in_thread(self._move_newsletter_thread)

    def _move_newsletter_thread(self) -> None:
        self._set_status("Đang chuyển newsletter...", FG_WARN)
        try:
            store_id = None
            if self._current_folder:
                store_id = self._current_folder.store_id
            else:
                inbox = self._outlook.get_default_inbox_info()
                if inbox:
                    store_id = inbox.store_id
            if not store_id:
                raise RuntimeError("Không xác định được store để tạo thư mục.")

            ids = list(self._scan_result.newsletter_ids)
            success, fail = 0, 0
            moved_ids: set[str] = set()

            for i, entry_id in enumerate(ids, 1):
                self._set_status(f"Đang chuyển newsletter {i}/{len(ids)}...", FG_WARN)
                email = self._scanned_emails.get(entry_id)
                if email is None:
                    fail += 1
                    continue
                # Determine per-email folder: Newsletter / OrgName [/ SenderName]
                path_parts = list(get_newsletter_path(email))
                try:
                    folder = self._outlook.get_or_create_folder_path(store_id, path_parts)
                    if self._outlook.move_email(entry_id, folder.entry_id, folder.store_id):
                        success += 1
                        moved_ids.add(entry_id)
                    else:
                        fail += 1
                except Exception:
                    fail += 1

            self._emails = [e for e in self._emails if e.entry_id not in moved_ids]
            self.after(0, self._populate_list, self._emails)
            self._scan_result.newsletter_ids.clear()
            self._write_output(
                f"📰 Đã chuyển {success} newsletter vào Newsletter / [Tổ chức] / [Người gửi]."
                + (f"\n  ({fail} lỗi)" if fail else "")
            )
            self._set_status(f"Đã chuyển {success} newsletter", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi chuyển newsletter: {exc}", error=True)
            self._set_status("Lỗi chuyển newsletter", FG_ERROR)

    # ------------------------------------------------------------------
    # Email organizer
    # ------------------------------------------------------------------

    def _run_organize(self) -> None:
        if not self._outlook:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ chưa sẵn sàng.")
            return
        if not self._emails:
            messagebox.showinfo("Không có email", "Danh sách email rỗng.")
            return

        # Build and preview the plan before committing
        plan = plan_organization(self._emails)
        preview = plan.display_preview()

        if not messagebox.askyesno(
            "Tổ chức Email",
            f"{preview}\n\n"
            f"Email sẽ được chuyển vào '{ORGANIZED_ROOT}/[Tổ chức]/[Năm]' "
            f"trong thư mục gốc của store hiện tại.\n\n"
            f"Tiếp tục?",
        ):
            return
        self._run_in_thread(self._organize_thread)

    def _organize_thread(self) -> None:
        self._set_status("Đang tổ chức email...", FG_WARN)
        try:
            plan = plan_organization(self._emails)

            store_id = None
            if self._current_folder:
                store_id = self._current_folder.store_id
            else:
                inbox = self._outlook.get_default_inbox_info()
                if inbox:
                    store_id = inbox.store_id
            if not store_id:
                raise RuntimeError("Không xác định được store.")

            total = plan.total_emails()
            moved_total, fail_total = 0, 0
            done = 0

            for path_parts, emails in plan.groups.items():
                try:
                    target_folder = self._outlook.get_or_create_folder_path(
                        store_id, [ORGANIZED_ROOT] + list(path_parts)
                    )
                except Exception as exc:
                    fail_total += len(emails)
                    continue

                for email in emails:
                    ok = self._outlook.move_email(
                        email.entry_id,
                        target_folder.entry_id,
                        target_folder.store_id,
                    )
                    if ok:
                        moved_total += 1
                    else:
                        fail_total += 1
                    done += 1
                    if done % 10 == 0:
                        self._set_status(f"Đang chuyển {done}/{total}...", FG_WARN)

            # Remove all moved emails from local list
            self._emails.clear()
            self.after(0, self._populate_list, self._emails)

            # Fetch and display Outlook rules
            rules = self._outlook.get_outlook_rules()
            rules_text = format_rules(rules)

            self._write_output(
                f"📂 Tổ chức xong!\n"
                f"  Đã chuyển : {moved_total} email\n"
                f"  Lỗi       : {fail_total} email\n"
                f"  Thư mục   : {plan.folder_count()}\n\n"
                + rules_text
            )
            self._set_status(f"Tổ chức xong: {moved_total} email", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi tổ chức: {exc}", error=True)
            self._set_status("Lỗi tổ chức", FG_ERROR)

    # ------------------------------------------------------------------
    # Archive old emails
    # ------------------------------------------------------------------

    def _run_archive(self) -> None:
        if not self._outlook:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ chưa sẵn sàng.")
            return
        if not self._emails:
            messagebox.showinfo("Không có email", "Danh sách email rỗng.")
            return

        plan = plan_archive(self._emails, cutoff_years=ARCHIVE_CUTOFF_YEARS)
        if not plan.groups:
            messagebox.showinfo(
                "Không có email cũ",
                plan.display_preview(),
            )
            return

        years = sorted(plan.groups.keys())
        preview = plan.display_preview()
        year_files = "\n".join(f"  • Outlook_Archive_{y}.pst" for y in years)
        if not messagebox.askyesno(
            "Archive Email Cũ",
            f"{preview}\n\n"
            f"Mỗi năm sẽ được lưu vào một file PST riêng:\n{year_files}\n\n"
            f"Chọn thư mục lưu các file PST archive.\nTiếp tục?",
        ):
            return

        archive_dir = filedialog.askdirectory(
            title="Chọn thư mục lưu file PST Archive",
        )
        if not archive_dir:
            return

        self._run_in_thread(lambda: self._archive_thread(plan, archive_dir))

    def _archive_thread(self, plan, archive_dir: str) -> None:
        self._set_status("Đang archive email cũ...", FG_WARN)
        try:
            total = plan.total_emails()
            moved_total, fail_total = 0, 0
            done = 0
            result_lines = []

            for year, emails in sorted(plan.groups.items()):
                pst_path = os.path.join(archive_dir, f"Outlook_Archive_{year}.pst")
                try:
                    store_id = self._outlook.get_or_open_pst(
                        pst_path, display_name=f"Archive {year}"
                    )
                    year_folder = self._outlook.get_or_create_folder_path(
                        store_id, [ARCHIVE_ROOT]
                    )
                except Exception as exc:
                    fail_total += len(emails)
                    result_lines.append(f"  ❌ {year}: lỗi tạo PST — {exc}")
                    continue

                yr_moved, yr_fail = 0, 0
                for email in emails:
                    ok = self._outlook.move_email(
                        email.entry_id,
                        year_folder.entry_id,
                        year_folder.store_id,
                    )
                    if ok:
                        moved_total += 1
                        yr_moved += 1
                    else:
                        fail_total += 1
                        yr_fail += 1
                    done += 1
                    if done % 10 == 0:
                        self._set_status(f"Đang archive {done}/{total}...", FG_WARN)

                result_lines.append(
                    f"  📦 {year}: {yr_moved} email → Outlook_Archive_{year}.pst"
                    + (f" ({yr_fail} lỗi)" if yr_fail else "")
                )

            # Refresh local list
            archived_ids = {e.entry_id for emails in plan.groups.values() for e in emails}
            self._emails = [e for e in self._emails if e.entry_id not in archived_ids]
            self.after(0, self._populate_list, self._emails)

            self._write_output(
                f"📦 Archive xong!\n"
                f"  Đã chuyển : {moved_total} email\n"
                f"  Lỗi       : {fail_total} email\n"
                f"  Thư mục   : {archive_dir}\n\n"
                + "\n".join(result_lines)
                + f"\n\nMỗi năm lưu vào file PST riêng trong thư mục '{ARCHIVE_ROOT}'."
            )
            self._set_status(f"Archive xong: {moved_total} email", FG_SUCCESS)
            # Check PST sizes after archiving
            self._passive_pst_check()
        except Exception as exc:
            self._write_output(f"Lỗi archive: {exc}", error=True)
            self._set_status("Lỗi archive", FG_ERROR)

    # ------------------------------------------------------------------
    # PST size monitoring
    # ------------------------------------------------------------------

    def _run_check_pst(self) -> None:
        if not self._outlook:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ chưa sẵn sàng.")
            return
        try:
            sizes = self._outlook.get_store_sizes()
            report = format_pst_sizes(sizes)
            self._write_output(report)

            # Show alert if any store exceeds warning threshold
            alerts = [s for s in sizes if s["size_gb"] >= PST_WARN_GB]
            if alerts:
                lines = []
                for s in alerts:
                    level = "⛔ NGUY HIỂM" if s["size_gb"] >= PST_LIMIT_GB else "⚠️ CẢNH BÁO"
                    lines.append(f"{level}: {s['name'][:40]} — {s['size_gb']:.1f} GB")
                messagebox.showwarning(
                    "Cảnh báo kích thước PST",
                    "\n".join(lines)
                    + f"\n\nFile PST trên {PST_LIMIT_GB:.0f} GB có thể bị lỗi."
                    + "\nHãy archive hoặc nén PST ngay.",
                )
                self._set_status(
                    f"⚠️ PST lớn: {alerts[0]['name'][:25]} {alerts[0]['size_gb']:.1f} GB",
                    FG_ERROR,
                )
            else:
                self._set_status("PST OK", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi kiểm tra PST: {exc}", error=True)

    def _passive_pst_check(self) -> None:
        """Non-blocking PST size check — only shows a warning if thresholds exceeded."""
        if not self._outlook:
            return
        try:
            sizes = self._outlook.get_store_sizes()
            critical = [s for s in sizes if s["size_gb"] >= PST_LIMIT_GB]
            warn = [s for s in sizes if PST_WARN_GB <= s["size_gb"] < PST_LIMIT_GB]

            if critical:
                s = critical[0]
                msg = (
                    f"⛔ {s['name'][:40]}\n{s['size_gb']:.1f} GB / {PST_LIMIT_GB:.0f} GB giới hạn\n\n"
                    f"File PST đang ở mức nguy hiểm, có nguy cơ bị hỏng dữ liệu.\n"
                    f"Hãy archive hoặc nén PST ngay lập tức!"
                )
                self.after(0, lambda m=msg: messagebox.showerror("PST quá lớn!", m))
            elif warn:
                s = warn[0]
                msg = (
                    f"⚠️ {s['name'][:40]}\n{s['size_gb']:.1f} GB / {PST_LIMIT_GB:.0f} GB\n\n"
                    f"File PST đang tiếp cận giới hạn {PST_LIMIT_GB:.0f} GB.\n"
                    f"Hãy archive email cũ để giảm kích thước."
                )
                self.after(0, lambda m=msg: messagebox.showwarning("Cảnh báo kích thước PST", m))
        except Exception:
            pass  # Silent fail — passive check should never crash

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _run_in_thread(self, fn) -> None:
        self._progress.start(12)
        def wrapper():
            try:
                fn()
            finally:
                self.after(0, self._progress.stop)
        threading.Thread(target=wrapper, daemon=True).start()

    def _set_status(self, msg: str, colour: str = FG_TEXT) -> None:
        def _update():
            self._status_var.set(msg)
            for w in self.winfo_children():
                if isinstance(w, tk.Frame):
                    for lbl in w.winfo_children():
                        if isinstance(lbl, tk.Label) and lbl.cget("textvariable") == str(self._status_var):
                            lbl.configure(fg=colour)
        self.after(0, _update)

    def _write_detail(self, text: str) -> None:
        def _update():
            self._detail_text.configure(state=tk.NORMAL)
            self._detail_text.delete("1.0", tk.END)
            self._detail_text.insert(tk.END, text)
            self._detail_text.configure(state=tk.DISABLED)
        self.after(0, _update)

    def _write_output(self, text: str, error: bool = False) -> None:
        colour = FG_ERROR if error else FG_SUCCESS

        def _update():
            self._output_text.configure(state=tk.NORMAL, fg=colour)
            self._output_text.delete("1.0", tk.END)
            self._output_text.insert(tk.END, text)
            self._output_text.configure(state=tk.DISABLED)
        self.after(0, _update)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = OutlookAIApp()
    app.mainloop()
