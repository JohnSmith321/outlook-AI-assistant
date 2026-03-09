"""
Outlook AI Assistant — Main entry point.

Launches a tkinter GUI that connects to Microsoft Outlook via COM
and provides AI-powered email management features powered by Claude.
"""

from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import sys
import os

# Add project root to path so sub-modules can import each other
sys.path.insert(0, os.path.dirname(__file__))

import config
from ai_client import AIClient
from outlook_client import OutlookClient, EmailMessage
from features.email_classifier import EmailClassifier
from features.task_creator import TaskCreator
from features.calendar_creator import CalendarCreator
from features.email_summarizer import EmailSummarizer
from features.email_rewriter import EmailRewriter
from features.scheduler import DailyScheduler


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
        tk.Label(hdr, text=" Hộp thư đến", font=FONT_TITLE, fg=FG_ACCENT, bg=BG_MID).pack(
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

        # Treeview
        cols = ("sender", "subject", "time")
        self._tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
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

        self._tree.heading("sender", text="Người gửi")
        self._tree.heading("subject", text="Chủ đề")
        self._tree.heading("time", text="Thời gian")
        self._tree.column("sender", width=140, minwidth=80)
        self._tree.column("subject", width=240, minwidth=100)
        self._tree.column("time", width=90, minwidth=70, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._tree.bind("<<TreeviewSelect>>", self._on_email_select)
        self._tree.tag_configure("unread", foreground="#89dceb", font=("Segoe UI", 10, "bold"))
        self._tree.tag_configure("urgent", foreground=FG_WARN)

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
        bar = tk.Frame(self, bg=BG_MID)
        bar.pack(fill=tk.X, padx=10, pady=(0, 8))

        buttons = [
            ("Phân loại Email", self._run_classify, "#313244", FG_ACCENT),
            ("Tạo Task", self._run_create_task, "#313244", FG_SUCCESS),
            ("Tạo Lịch họp", self._run_create_meeting, "#313244", "#cba6f7"),
            ("Tóm tắt Thread", self._run_summarize, "#313244", "#89dceb"),
            ("Viết lại (VI)", lambda: self._run_rewrite("vi"), "#313244", FG_WARN),
            ("Viết lại (EN)", lambda: self._run_rewrite("en"), "#313244", "#f9e2af"),
            ("Gợi ý lịch ngày", self._run_schedule, "#313244", "#a6e3a1"),
        ]
        for label, cmd, bg, fg in buttons:
            tk.Button(
                bar, text=label, command=cmd,
                font=FONT_UI, bg=bg, fg=fg,
                activebackground="#45475a", activeforeground=fg,
                bd=0, padx=12, pady=6, cursor="hand2",
            ).pack(side=tk.LEFT, padx=4, pady=6)

        self._progress = ttk.Progressbar(bar, mode="indeterminate", length=120)
        self._progress.pack(side=tk.RIGHT, padx=10)

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

        self._set_status("Sẵn sàng", FG_SUCCESS)
        self._reload_emails()

    def _reload_emails(self) -> None:
        if not self._outlook:
            return
        threading.Thread(target=self._load_emails_thread, daemon=True).start()

    def _load_emails_thread(self) -> None:
        self._set_status("Đang tải email...", FG_WARN)
        try:
            self._emails = self._outlook.get_inbox_emails(limit=config.EMAIL_LOAD_LIMIT)
            self.after(0, self._populate_list, self._emails)
            self._set_status(f"Đã tải {len(self._emails)} email", FG_SUCCESS)
        except Exception as exc:
            self._set_status(f"Lỗi tải email: {exc}", FG_ERROR)

    def _populate_list(self, emails: list[EmailMessage]) -> None:
        self._tree.delete(*self._tree.get_children())
        for e in emails:
            tag = "unread" if e.unread else ("urgent" if e.importance == 2 else "")
            self._tree.insert(
                "", tk.END,
                iid=e.entry_id,
                values=(
                    e.sender[:25],
                    e.subject[:60],
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

    def _on_email_select(self, _event=None) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        entry_id = sel[0]
        email = next((e for e in self._emails if e.entry_id == entry_id), None)
        if email:
            self._selected_email = email
            self._show_email_detail(email)

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

    def _guard(self) -> bool:
        """Return True if services are ready and an email is selected."""
        if not self._ai or not self._outlook:
            messagebox.showwarning("Chưa sẵn sàng", "Dịch vụ chưa khởi tạo xong.")
            return False
        if not self._selected_email:
            messagebox.showinfo("Chọn email", "Vui lòng chọn một email trước.")
            return False
        return True

    def _run_classify(self) -> None:
        if not self._guard():
            return
        self._run_in_thread(self._classify_thread)

    def _classify_thread(self) -> None:
        self._set_status("Đang phân loại email...", FG_WARN)
        try:
            clf = EmailClassifier(self._ai)
            result = clf.classify(self._selected_email)
            self._write_output(result.display())
            self._set_status("Phân loại xong", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi: {exc}", error=True)
            self._set_status("Lỗi phân loại", FG_ERROR)

    def _run_create_task(self) -> None:
        if not self._guard():
            return
        self._run_in_thread(self._create_task_thread)

    def _create_task_thread(self) -> None:
        self._set_status("Đang tạo task...", FG_WARN)
        try:
            creator = TaskCreator(self._ai, self._outlook)
            result = creator.extract_and_create(self._selected_email)
            self._write_output(result.display())
            self._set_status(f"Đã tạo {result.created_count} task", FG_SUCCESS)
        except Exception as exc:
            self._write_output(f"Lỗi: {exc}", error=True)
            self._set_status("Lỗi tạo task", FG_ERROR)

    def _run_create_meeting(self) -> None:
        if not self._guard():
            return
        self._run_in_thread(self._create_meeting_thread)

    def _create_meeting_thread(self) -> None:
        self._set_status("Đang tạo lịch họp...", FG_WARN)
        try:
            creator = CalendarCreator(self._ai, self._outlook)
            result = creator.extract_and_create(self._selected_email)
            self._write_output(result.display())
            self._set_status(
                f"Đã tạo {result.created_count} sự kiện" if result.has_meeting else "Không có cuộc họp",
                FG_SUCCESS,
            )
        except Exception as exc:
            self._write_output(f"Lỗi: {exc}", error=True)
            self._set_status("Lỗi tạo lịch", FG_ERROR)

    def _run_summarize(self) -> None:
        if not self._guard():
            return
        self._run_in_thread(self._summarize_thread)

    def _summarize_thread(self) -> None:
        self._set_status("Đang tóm tắt luồng email...", FG_WARN)
        try:
            summarizer = EmailSummarizer(self._ai)
            email = self._selected_email

            # Try to get thread; fall back to single email
            if email.conversation_topic and self._outlook:
                thread = self._outlook.get_thread_emails(email.conversation_topic)
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
        if not self._guard():
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
