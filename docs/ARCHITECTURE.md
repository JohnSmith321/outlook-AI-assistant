# Kiến trúc Hệ thống – Outlook AI Assistant

## Tổng quan

Outlook AI Assistant là ứng dụng desktop Python tích hợp hai hệ thống bên ngoài:
1. **Microsoft Outlook** (qua COM/pywin32) – nguồn dữ liệu email và đích tạo task/calendar/folder/PST
2. **Anthropic Claude API** – engine AI để phân tích, phân loại và sinh nội dung

```
┌─────────────────────────────────────────────────────────────────┐
│                       NGƯỜI DÙNG (GUI)                          │
│                    main.py  (tkinter)                           │
└──────────────────────┬──────────────────────────────────────────┘
                       │
         ┌─────────────┼──────────────┐
         ▼             ▼              ▼
  ┌─────────────┐ ┌──────────┐ ┌───────────────┐
  │OutlookClient│ │ AIClient │ │ Feature Layer │
  │(pywin32 COM)│ │(Anthropic│ │  (8 modules)  │
  └──────┬──────┘ │  SDK)    │ └───────┬───────┘
         │        └─────┬────┘         │
         ▼              ▼              │
  ┌────────────┐  ┌──────────────┐    │
  │  Microsoft │  │  Claude API  │    │
  │  Outlook   │  │ (claude-opus │◄───┘
  │ (COM/MAPI) │  │   -4-6)      │
  └────────────┘  └──────────────┘
```

---

## Các thành phần

### 1. Presentation Layer – `main.py`

**Công nghệ**: Python `tkinter` (built-in, không cần cài thêm)

**Trách nhiệm**:
- Hiển thị danh sách email trong Inbox
- Nhận thao tác người dùng (click nút)
- Gọi các Feature module trong background thread
- Hiển thị kết quả AI trong real-time

**Pattern**: MVC – View (tkinter) + thin Controller (event handlers), Model là các service classes

**Thiết kế GUI**:
```
┌────── PanedWindow (horizontal) ──────────────────────────────┐
│  Left (380px)              │  Right (expandable)             │
│  ┌──────────────────────┐  │  ┌────────────────────────────┐ │
│  │ Header + Reload btn  │  │  │ Email Detail (ScrolledText)│ │
│  ├──────────────────────┤  │  ├────────────────────────────┤ │
│  │ Search Entry         │  │  │ AI Output (ScrolledText)   │ │
│  ├──────────────────────┤  │  │                            │ │
│  │ Treeview (emails)    │  │  │                            │ │
│  │                      │  │  └────────────────────────────┘ │
│  └──────────────────────┘  │                                  │
└────────────────────────────┴──────────────────────────────────┘
│  Hàng 1: [Phân loại] [Tạo Task] [Tạo Lịch họp] [Tóm tắt]   │
│  Hàng 2: [Viết lại VI] [Viết lại EN] [Gợi ý lịch ngày]      │
│  Hàng 3: [Quét Spam/Newsletter] [Xóa Spam] [Di chuyển NL]   │
│  Hàng 4: [Sắp xếp Email] [Archive Email cũ] [Kiểm tra PST]  │
└──────────────────────────────────────────────────────────────┘
```

### 2. Integration Layer

#### `outlook_client.py` – Outlook COM Bridge
**Công nghệ**: `pywin32` (win32com.client)

**Data Classes** (thuần Python, không phụ thuộc COM):
- `EmailMessage` – thông tin email: subject, sender, body, time, entry_id, store_id
- `EmailThread` – tập hợp email cùng conversation topic
- `OutlookTask` – dữ liệu cho Outlook TaskItem
- `CalendarEvent` – dữ liệu cho Outlook AppointmentItem

**COM Objects sử dụng**:
| COM Object / Method | Mục đích |
|---------------------|----------|
| `Application` | Entry point Outlook |
| `Namespace("MAPI")` | Truy cập MAPI namespace |
| `GetDefaultFolder(6)` – olFolderInbox | Inbox |
| `GetDefaultFolder(13)` – olFolderTasks | Tasks |
| `GetDefaultFolder(9)` – olFolderCalendar | Calendar |
| `CreateItem(1)` – olAppointmentItem | Tạo cuộc hẹn |
| `CreateItem(3)` – olTaskItem | Tạo task |
| `namespace.Stores` | Liệt kê tất cả PST/mailbox |
| `store.FilePath` | Đường dẫn file PST |
| `namespace.AddStoreEx(path, 2)` | Mở hoặc tạo PST mới |
| `folder.Folders.Add(name)` | Tạo subfolder |
| `item.Move(folder)` | Di chuyển email |

**API mở rộng** (multi-PST):
```python
get_all_stores() -> List[StoreInfo]
get_store_sizes() -> List[Dict]          # size_gb, name, path
get_or_open_pst(pst_path, display_name) -> str  # store_id
get_or_create_folder_path(store_id, path_parts) -> COMFolder
move_email(entry_id, store_id, folder_path) -> bool
```

#### `ai_client.py` – Claude API Wrapper
**Công nghệ**: `anthropic` Python SDK

**Model**: `claude-opus-4-6`
**Features used**: Basic messages, streaming (cho response dài)

**Interface**:
```python
ai.chat(system=..., user=..., stream=False) -> str
```

### 3. Feature Layer – `features/`

Mỗi module thực hiện một hoặc nhiều tính năng độc lập. Pattern chung:
1. Nhận `EmailMessage` (hoặc danh sách) từ caller
2. Build prompt (system + user)
3. Gọi `AIClient.chat()`
4. Parse response (JSON hoặc plain text, có regex fallback)
5. Thực hiện action (create task/event/move email) nếu cần
6. Trả về `*Result` dataclass

**Phân tích & Soạn thảo (AI-driven)**:
| Module | Input | Claude Output | Outlook Action |
|--------|-------|---------------|----------------|
| `email_classifier.py` | EmailMessage | JSON classification | Không |
| `task_creator.py` | EmailMessage | JSON task list | Tạo TaskItem |
| `calendar_creator.py` | EmailMessage | JSON event list | Tạo AppointmentItem |
| `email_summarizer.py` | EmailMessage / EmailThread | Plain text summary | Không |
| `email_rewriter.py` | EmailMessage | Plain text rewrite | Không |
| `scheduler.py` | List[EmailMessage] | Plain text schedule | Không |

**Quản lý & Dọn dẹp (hành động trực tiếp)**:
| Module | Input | Claude Output | Outlook Action |
|--------|-------|---------------|----------------|
| `spam_cleaner.py` | List[EmailMessage] | JSON per-email label | Xóa / Di chuyển email |
| `email_organizer.py` | List[EmailMessage] | (không dùng Claude) | Di chuyển email vào folder theo sender/năm |

**Hằng số quan trọng** (`features/email_organizer.py`):
```python
ORGANIZED_ROOT = "Organized"       # Root folder sắp xếp email
ARCHIVE_ROOT = "Archive"           # Folder trong file PST archive
ARCHIVE_CUTOFF_YEARS = 2           # Archive email cũ hơn 2 năm
PST_WARN_GB = 47.0                 # Ngưỡng cảnh báo vàng
PST_LIMIT_GB = 50.0                # Ngưỡng cảnh báo đỏ
```

---

## Luồng dữ liệu (Data Flow)

### Phân loại email
```
User click "Phân loại"
    → main.py._run_classify()
    → [background thread]
    → EmailClassifier.classify(email)
        → Build system + user prompt
        → AIClient.chat() → Claude API
        → Parse JSON response
        → Return ClassificationResult
    → [main thread] _write_output(result.display())
```

### Tạo Task
```
User click "Tạo Task"
    → main.py._run_create_task()
    → [background thread]
    → TaskCreator.extract_and_create(email)
        → Claude extracts task list as JSON
        → For each task:
            → OutlookClient.create_task(OutlookTask)
                → win32com CreateItem(3) → Save()
        → Return TaskCreationResult
    → [main thread] _write_output(result.display())
```

### Quét Spam & Newsletter
```
User click "Quét Spam/Newsletter"
    → main.py._run_scan()
    → [background thread]
    → SpamCleaner.scan(emails)
        → For each email:
            → Claude classifies as spam/newsletter/normal (JSON)
        → Return ScanResult(spam_ids, newsletter_ids, normal_ids)
    → [main thread] _write_output(report); enable Xóa Spam, Di chuyển NL buttons
```

### Sắp xếp Email
```
User click "Sắp xếp Email"
    → main.py._run_organize()
    → preview OrganizePlan (không dùng Claude)
    → User xác nhận
    → [background thread]
    → _organize_thread(plan)
        → For each (path_parts, emails) in plan.groups:
            → OutlookClient.get_or_create_folder_path(store_id, [ORGANIZED_ROOT] + list(path_parts))
            → For each email: OutlookClient.move_email(...)
```

### Archive Email cũ
```
User click "Archive Email cũ"
    → main.py._run_archive()
    → preview danh sách file PST theo năm
    → filedialog.askdirectory() → archive_dir
    → [background thread]
    → _archive_thread(plan, archive_dir)
        → For each year in plan.groups:
            → pst_path = archive_dir / "Outlook_Archive_{year}.pst"
            → store_id = OutlookClient.get_or_open_pst(pst_path, f"Archive {year}")
            → year_folder = get_or_create_folder_path(store_id, ["Archive"])
            → For each email: move_email(entry_id, store_id=None, folder=year_folder)
```

---

## Threading Model

Mọi AI call và Outlook operation chạy trên **background daemon thread** để GUI không bị block:

```python
threading.Thread(target=feature_fn, daemon=True).start()
```

Cập nhật UI được marshal về main thread qua `self.after(0, callback)`.

**Lưu ý đặc biệt với passive PST check**:
```python
# _passive_pst_check chạy trên background thread
# Dùng default arg để tránh lambda closure late-binding
self.after(0, lambda m=msg: messagebox.showerror("PST quá lớn!", m))
```

---

## Prompt Engineering Strategy

### Nguyên tắc chung
- **JSON output** cho các feature cần structured data (classifier, task, calendar, spam scan)
- **Plain text** cho các feature output trực tiếp cho người dùng (summary, rewrite, schedule)
- System prompt luôn bao gồm: vai trò AI + format output mong muốn + quy tắc xử lý
- User prompt bao gồm: metadata email (sender, subject, time) + body (truncated 3000-5000 chars)
- **Regex fallback**: `re.search(r"\{.*\}", raw, re.DOTALL)` để extract JSON khi Claude thêm markdown

### Ví dụ – Email Classifier System Prompt
```
Bạn là trợ lý AI chuyên phân loại email doanh nghiệp.
Phân tích email và trả về JSON...
{priority, category, action, summary}
Quy tắc: Urgent nếu cần hành động trong 24h...
```

### Ví dụ – Spam Cleaner System Prompt
```
Bạn là trợ lý AI phân loại email.
Phân loại email thành một trong ba loại: spam / newsletter / normal
Trả về JSON: {"label": "spam"|"newsletter"|"normal", "reason": "..."}
```

---

## Bảo mật

| Vấn đề | Giải pháp |
|--------|-----------|
| API Key | Lưu trong `.env` (không commit), load qua `python-dotenv` |
| Email body | Truncate tới 3000-5000 chars trước khi gửi lên API |
| COM object | Chỉ truy cập local, không expose ra ngoài |
| Error handling | Mọi exception đều được bắt, log ra UI, không crash app |
| PST path | Normalize với `os.path.normcase(os.path.abspath())` để tránh mở duplicate |

---

## Dependencies

| Package | Version | Mục đích |
|---------|---------|----------|
| `anthropic` | ≥0.40.0 | Claude API client |
| `pywin32` | ≥306 | Outlook COM automation |
| `python-dotenv` | ≥1.0.0 | Load `.env` file |
| `tkinter` | built-in | GUI framework |
