# Kiến trúc Hệ thống – Outlook AI Assistant

## Tổng quan

Outlook AI Assistant là ứng dụng desktop Python tích hợp hai hệ thống bên ngoài:
1. **Microsoft Outlook** (qua COM/pywin32) – nguồn dữ liệu email và đích tạo task/calendar
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
  │(pywin32 COM)│ │(Anthropic│ │  (6 modules)  │
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
│  Action Bar: [Phân loại] [Task] [Lịch] [Tóm tắt] [VI] [EN] [📅]
└──────────────────────────────────────────────────────────────┘
```

### 2. Integration Layer

#### `outlook_client.py` – Outlook COM Bridge
**Công nghệ**: `pywin32` (win32com.client)

**Data Classes** (thuần Python, không phụ thuộc COM):
- `EmailMessage` – thông tin email: subject, sender, body, time, entry_id
- `EmailThread` – tập hợp email cùng conversation topic
- `OutlookTask` – dữ liệu cho Outlook TaskItem
- `CalendarEvent` – dữ liệu cho Outlook AppointmentItem

**COM Objects sử dụng**:
| COM Object | Outlook Constant | Mục đích |
|------------|-----------------|----------|
| `Application` | – | Entry point Outlook |
| `Namespace("MAPI")` | – | Truy cập folders |
| `GetDefaultFolder(6)` | olFolderInbox | Inbox |
| `GetDefaultFolder(13)` | olFolderTasks | Tasks |
| `GetDefaultFolder(9)` | olFolderCalendar | Calendar |
| `CreateItem(1)` | olAppointmentItem | Tạo cuộc hẹn |
| `CreateItem(3)` | olTaskItem | Tạo task |

#### `ai_client.py` – Claude API Wrapper
**Công nghệ**: `anthropic` Python SDK

**Model**: `claude-opus-4-6`
**Features used**: Basic messages, streaming (cho response dài)

**Interface**:
```python
ai.chat(system=..., user=..., stream=False) -> str
```

### 3. Feature Layer – `features/`

Mỗi module thực hiện một tính năng độc lập. Pattern chung:
1. Nhận `EmailMessage` từ caller
2. Build prompt (system + user)
3. Gọi `AIClient.chat()`
4. Parse response (JSON hoặc plain text)
5. Thực hiện action (create task/event) nếu cần
6. Trả về `*Result` dataclass

| Module | Input | Claude Output | Outlook Action |
|--------|-------|---------------|----------------|
| `email_classifier.py` | EmailMessage | JSON classification | Không |
| `task_creator.py` | EmailMessage | JSON task list | Tạo TaskItem |
| `calendar_creator.py` | EmailMessage | JSON event list | Tạo AppointmentItem |
| `email_summarizer.py` | EmailMessage / EmailThread | Plain text summary | Không |
| `email_rewriter.py` | EmailMessage | Plain text rewrite | Không |
| `scheduler.py` | List[EmailMessage] | Plain text schedule | Không |

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

---

## Threading Model

Mọi AI call và Outlook operation chạy trên **background daemon thread** để GUI không bị block:

```python
threading.Thread(target=feature_fn, daemon=True).start()
```

Cập nhật UI được marshal về main thread qua `self.after(0, callback)`.

---

## Prompt Engineering Strategy

### Nguyên tắc chung
- **JSON output** cho các feature cần structured data (classifier, task, calendar)
- **Plain text** cho các feature output trực tiếp cho người dùng (summary, rewrite, schedule)
- System prompt luôn bao gồm: vai trò AI + format output mong muốn + quy tắc xử lý
- User prompt bao gồm: metadata email (sender, subject, time) + body (truncated 3000-5000 chars)

### Ví dụ – Email Classifier System Prompt
```
Bạn là trợ lý AI chuyên phân loại email doanh nghiệp.
Phân tích email và trả về JSON...
{priority, category, action, summary}
Quy tắc: Urgent nếu cần hành động trong 24h...
```

---

## Bảo mật

| Vấn đề | Giải pháp |
|--------|-----------|
| API Key | Lưu trong `.env` (không commit), load qua `python-dotenv` |
| Email body | Truncate tới 3000-5000 chars trước khi gửi lên API |
| COM object | Chỉ truy cập local, không expose ra ngoài |
| Error handling | Mọi exception đều được bắt, log ra UI, không crash app |

---

## Dependencies

| Package | Version | Mục đích |
|---------|---------|----------|
| `anthropic` | ≥0.40.0 | Claude API client |
| `pywin32` | ≥306 | Outlook COM automation |
| `python-dotenv` | ≥1.0.0 | Load `.env` file |
| `tkinter` | built-in | GUI framework |
