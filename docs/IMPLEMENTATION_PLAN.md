# Kế hoạch Triển khai – Outlook AI Assistant

## Tổng quan dự án

| Thông tin | Chi tiết |
|-----------|----------|
| Tên dự án | Outlook AI Assistant |
| Ngôn ngữ | Python 3.10+ |
| LLM | Claude Opus 4.6 (Anthropic) |
| Tích hợp Outlook | pywin32 (COM/MAPI) |
| GUI | tkinter (built-in) |
| Tổng thời gian ước tính | 5 ngày |

---

## Giai đoạn 1: Thiết lập hạ tầng (Ngày 1 – buổi sáng)

### 1.1 Khởi tạo dự án
- [x] Tạo cấu trúc thư mục (`features/`, `docs/`)
- [x] Tạo `requirements.txt`
- [x] Tạo `.env.example`
- [x] Tạo `.gitignore` (loại trừ `.env`, `__pycache__`, `venv/`)

### 1.2 Module nền tảng
- [x] `config.py` – quản lý cấu hình và API key
- [x] `outlook_client.py` – wrapper toàn bộ Outlook COM operations
  - [x] `get_inbox_emails(limit)` – đọc email từ Inbox
  - [x] `get_thread_emails(topic)` – đọc email theo conversation
  - [x] `create_task(OutlookTask)` – tạo task
  - [x] `create_calendar_event(CalendarEvent)` – tạo lịch
- [x] `ai_client.py` – wrapper Claude API
  - [x] `chat(system, user, stream)` – gửi message và nhận response

**Tiêu chí hoàn thành**: Có thể đọc email từ Outlook và gọi Claude API thành công

---

## Giai đoạn 2: Tính năng cơ bản (Ngày 1 – buổi chiều → Ngày 2)

### 2.1 Phân loại email (`features/email_classifier.py`)
- [x] Thiết kế system prompt phân loại (priority, category, action, summary)
- [x] Implement `EmailClassifier.classify(email)`
- [x] Parse JSON response, xử lý lỗi parse
- [x] `ClassificationResult.display()` – format kết quả đẹp

**Test case**: Email mời họp → action=Meeting-Request, priority=Urgent

### 2.2 Tạo Task (`features/task_creator.py`)
- [x] Thiết kế prompt trích xuất task từ email
- [x] Implement `TaskCreator.extract_and_create(email)`
- [x] Parse danh sách task (JSON array)
- [x] Xử lý ngày tháng từ text tự nhiên → datetime
- [x] Gọi `OutlookClient.create_task()` cho mỗi task

**Test case**: Email giao nhiệm vụ với deadline → tạo đúng task với due date

### 2.3 Tạo Lịch họp (`features/calendar_creator.py`)
- [x] Thiết kế prompt phát hiện meeting info
- [x] Implement `CalendarCreator.extract_and_create(email)`
- [x] Parse thông tin cuộc họp (subject, start, end, location, attendees)
- [x] Tự động tính end time nếu chỉ có start + duration
- [x] Gọi `OutlookClient.create_calendar_event()`

**Test case**: Email "Họp lúc 10h thứ Hai, phòng A3" → tạo đúng sự kiện calendar

---

## Giai đoạn 3: Tính năng nâng cao (Ngày 2 – buổi chiều)

### 3.1 Tóm tắt Thread (`features/email_summarizer.py`)
- [x] Thiết kế prompt tóm tắt có cấu trúc (điểm chính, quyết định, hành động)
- [x] `summarize_email(email)` – tóm tắt email đơn lẻ
- [x] `summarize_thread(thread)` – tóm tắt cả luồng email
- [x] Dùng streaming để nhận response dài
- [x] Format output với emoji và cấu trúc rõ ràng

### 3.2 Viết lại Email (`features/email_rewriter.py`)
- [x] System prompt riêng cho VI và EN
- [x] `rewrite(email, language)` – viết lại theo ngôn ngữ chỉ định
- [x] Giữ nguyên thông tin gốc, cải thiện văn phong
- [x] Dùng streaming

### 3.3 Gợi ý Lịch ngày (`features/scheduler.py`)
- [x] System prompt xây dựng kế hoạch ngày với time blocks
- [x] `suggest_schedule(emails, extra_notes)` – gợi ý dựa trên email hôm nay
- [x] Format theo bảng thời gian buổi sáng/chiều
- [x] Hỗ trợ ghi chú bổ sung từ người dùng

---

## Giai đoạn 4: GUI & Tích hợp (Ngày 3)

### 4.1 Giao diện chính (`main.py`)
- [x] Thiết kế layout: PanedWindow (email list | detail + output)
- [x] Email list: Treeview với columns (sender, subject, time)
- [x] Search/filter email theo text
- [x] Email detail panel: hiển thị nội dung email được chọn
- [x] AI output panel: hiển thị kết quả AI
- [x] Action bar: 7 nút tương ứng 6 tính năng + schedule

### 4.2 UX & Threading
- [x] Mọi AI call chạy trên background thread (không block GUI)
- [x] Progress bar trong khi AI đang xử lý
- [x] Status bar hiển thị trạng thái hiện tại
- [x] Dark theme (catppuccin-inspired color palette)
- [x] Màu sắc phân biệt: unread (cyan bold), urgent (orange)

### 4.3 Error Handling
- [x] Kiểm tra Outlook và AI service trước khi cho phép dùng tính năng
- [x] Hiển thị lỗi rõ ràng trong output panel
- [x] Bắt exception cụ thể (COM errors, API errors, JSON parse errors)

---

## Giai đoạn 5: Multi-PST & Nested Folders (Ngày 4 – buổi sáng)

### 5.1 Mở rộng OutlookClient
- [x] `get_all_stores()` – liệt kê tất cả PST/mailbox đang mở trong Outlook
- [x] `get_store_sizes()` – đọc kích thước từng PST qua `os.path.getsize(store.FilePath)`
- [x] `get_or_open_pst(pst_path, display_name)` – mở hoặc tạo PST mới qua `ns.AddStoreEx()`
- [x] `get_or_create_folder_path(store_id, path_parts)` – tạo nested folder tùy ý
- [x] `move_email(entry_id, store_id, folder_path)` – di chuyển email giữa các folder/PST
- [x] Batch email loading từ nhiều folder/PST

**Tiêu chí hoàn thành**: Có thể tạo PST mới, tạo folder con, và di chuyển email giữa PST

---

## Giai đoạn 6: Spam/Newsletter/Organizer/Archive/PST Monitoring (Ngày 4–5)

### 6.1 Spam & Newsletter Scanner (`features/spam_cleaner.py`)
- [x] `ScanResult` dataclass: spam_ids, newsletter_ids, normal_ids, classifications
- [x] System prompt phân loại 3 loại: spam / newsletter / normal
- [x] `SpamCleaner.scan(emails)` – gọi Claude cho từng email, parse JSON, regex fallback
- [x] Thống kê kết quả (số lượng từng loại)
- [x] GUI: nút "Quét Spam/Newsletter" gọi scan, hiển thị báo cáo tóm tắt

**Test case**: Email quảng cáo → spam; email bản tin định kỳ → newsletter; email công việc → normal

### 6.2 Xóa Spam & Di chuyển Newsletter (`features/spam_cleaner.py`)
- [x] `delete_spam(scan_result)` – xóa tất cả email được đánh dấu spam
- [x] `move_newsletters(scan_result, target_folder)` – di chuyển newsletter đến folder chỉ định
  - [x] Folder Newsletter phân theo sender: `Newsletter/Gmail/Tên người gửi/`
  - [x] Hỗ trợ các personal domain: gmail.com, yahoo.com, hotmail.com, v.v.
- [x] GUI: nút "Xóa Spam" và "Di chuyển Newsletter" (disabled nếu chưa quét)

### 6.3 Sắp xếp Email theo Sender (`features/email_organizer.py`)
- [x] `_get_organize_rel_path(email)` – phân loại theo domain:
  - Company domain → `Organized/TênCôngTy/Năm/`
  - Personal domain → `Organized/BrandName/TênNgườiGửi/Năm/`
- [x] `plan_organize(emails)` – nhóm email theo path key (Dict[Tuple[str, ...], List])
- [x] `OrganizePlan` dataclass với tổng hợp thống kê
- [x] GUI: nút "Sắp xếp Email" gọi plan, hiển thị preview, xác nhận rồi di chuyển
- [x] `_organize_thread` dùng variable-length path để tránh crash với 3-tuple

**Test case**: Email từ abc@gmail.com → `Organized/Gmail/abc/2025/`; email từ user@company.vn → `Organized/Company/2025/`

### 6.4 Archive Email cũ ra PST (`features/email_organizer.py`)
- [x] `plan_archive(emails, cutoff_years=2)` – lọc email cũ hơn cutoff, nhóm theo năm
- [x] Xử lý timezone-aware datetime khi so sánh với cutoff
- [x] GUI: nút "Archive Email cũ" → preview danh sách file PST → chọn thư mục → di chuyển
- [x] Một file PST riêng cho mỗi năm: `Outlook_Archive_{year}.pst`
- [x] Email trong mỗi PST phân folder theo sender: `Archive/CôngTy/` hoặc `Archive/Brand/localpart/`
- [x] Dùng email address (local part) thay vì display name để tránh tạo folder trùng
- [x] Folder cache tránh tạo lại folder trùng lặp trong cùng 1 lần archive
- [x] `filedialog.askdirectory()` cho user chọn thư mục xuất

**Test case**: Email từ abc@gmail.com năm 2023 → `Outlook_Archive_2023.pst/Archive/Gmail/abc/`

### 6.5 Kiểm tra kích thước PST (`features/email_organizer.py`)
- [x] `format_pst_sizes(sizes)` – render bảng + bar chart ASCII cho từng PST
- [x] Ngưỡng cảnh báo: 47 GB (⚠️ warning), 50 GB (⛔ critical)
- [x] GUI: nút "Kiểm tra PST" → hiển thị báo cáo chi tiết
- [x] Passive check: tự động kiểm tra sau mỗi lần load email → popup nếu vượt ngưỡng
- [x] Lambda closure safety: dùng default arg binding `lambda m=msg: ...` tránh late-binding bug

**Test case**: PST có kích thước 48 GB → hiển thị cảnh báo màu vàng; >50 GB → popup lỗi đỏ

---

## Giai đoạn 7: Tài liệu (Hoàn thiện)

- [x] `docs/README.md` – hướng dẫn cài đặt và sử dụng
- [x] `docs/ARCHITECTURE.md` – mô tả kiến trúc hệ thống
- [x] `docs/IMPLEMENTATION_PLAN.md` – kế hoạch triển khai (file này)
- [x] `docs/TEST_PLAN.md` – kế hoạch kiểm thử
- [x] `docs/DEV_DIARY.md` – nhật ký phát triển
- [x] Root `README.md` – hướng dẫn đầy đủ với tất cả tính năng mới

---

## Rủi ro và Giải pháp

| Rủi ro | Xác suất | Giải pháp |
|--------|----------|-----------|
| Outlook COM không kết nối được | Thấp | Kiểm tra Outlook đang chạy, hiển thị lỗi rõ ràng |
| Claude API rate limit | Thấp | SDK tự retry, hiển thị thông báo chờ |
| JSON parse lỗi từ Claude | Trung bình | Regex fallback để extract JSON từ response |
| Timezone của Outlook vs Python | Trung bình | Convert pywintypes.datetime → Python datetime |
| Email body quá dài → vượt token limit | Cao | Truncate body trước khi gửi (3000-5000 chars) |
| PST file path normalize (tránh mở trùng) | Trung bình | `os.path.normcase(os.path.abspath(path))` để so sánh |
| Lambda closure late-binding | Trung bình | Dùng default arg: `lambda m=msg: messagebox.showerror(...)` |
| Variable-length folder path (2-tuple vs 3-tuple) | Trung bình | Dùng `Tuple[str, ...]` thay vì fixed-length destructuring |

---

## Định nghĩa hoàn thành (Definition of Done)

Một tính năng được coi là hoàn chỉnh khi:
1. ✅ Code hoạt động đúng với email thực từ Outlook
2. ✅ Xử lý được các edge case (email rỗng, không có task/meeting,...)
3. ✅ Error handling: không crash khi có lỗi, hiển thị message thân thiện
4. ✅ Output được hiển thị đẹp trong GUI
5. ✅ Không block GUI thread

---

## Metrics đánh giá thành công

| Metric | Mục tiêu |
|--------|---------|
| Tỉ lệ phân loại email chính xác | ≥ 85% |
| Tỉ lệ phát hiện spam chính xác | ≥ 90% |
| Thời gian response mỗi tính năng | < 15 giây |
| Tỉ lệ tạo task/event thành công | ≥ 95% |
| Tỉ lệ di chuyển email thành công | ≥ 99% |
| Không có crash trong demo 30 phút | 100% |
