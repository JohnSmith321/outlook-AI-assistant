# Outlook AI Assistant

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Python](https://img.shields.io/badge/Python-3.10+-green.svg)](https://www.python.org/)
[![Claude](https://img.shields.io/badge/LLM-Claude%20Opus%204.6-orange.svg)](https://www.anthropic.com/)

Trợ lý AI tự động hóa công việc email trong Microsoft Outlook, tích hợp Claude (Anthropic).

## Tính năng

**Phân tích & Soạn thảo**
- **Chọn thư mục linh hoạt** — hỗ trợ nhiều tài khoản / file PST (Inbox, Viettel, Archive...)
- Tự động phân loại email đơn lẻ hoặc **hàng loạt** (ưu tiên, danh mục, hành động)
- Tạo Task Outlook từ một hoặc **nhiều email** cùng lúc
- Tạo lịch họp Outlook Calendar từ một hoặc **nhiều email** cùng lúc
- Tóm tắt luồng email có cấu trúc *(nâng cao)*
- Viết lại email chuyên nghiệp (tiếng Việt / tiếng Anh) *(nâng cao)*
- Gợi ý kế hoạch làm việc hàng ngày *(nâng cao)*

**Quản lý & Dọn dẹp**
- **Quét Spam/Newsletter** — AI phân loại từng email thành spam / newsletter / bình thường
- **Xóa Spam** — chuyển email spam vào Deleted Items (cần xác nhận)
- **Chuyển Newsletter** — tự động tạo thư mục `Newsletter / Tổ chức [/ Người gửi]` và phân loại
- **Tổ chức Email** — xếp email vào `Organized / Tổ chức / Năm` (cần xác nhận trước khi di chuyển)
- **Archive Email Cũ** — chuyển email > 2 năm vào file PST archive theo từng năm
- **Kiểm tra PST** — theo dõi kích thước file PST, cảnh báo ở 47 GB, báo lỗi ở 50 GB

## Cài đặt nhanh

```bash
pip install -r requirements.txt
copy .env.example .env   # Điền ANTHROPIC_API_KEY
python main.py
```

Xem [docs/README.md](docs/README.md) để biết hướng dẫn chi tiết.

---

## Hướng dẫn sử dụng (User Manual)

### Yêu cầu trước khi chạy
- Microsoft Outlook đã cài đặt và **đang mở**
- File `.env` đã có `ANTHROPIC_API_KEY`

### Khởi động ứng dụng

```bash
python main.py
```

Khi khởi động, ứng dụng sẽ tự động:
1. Kết nối Microsoft Outlook qua COM (tự động phát hiện tất cả PST / tài khoản)
2. Kết nối Claude AI API
3. Tải danh sách email từ Inbox mặc định

Thanh trạng thái (góc phải trên) chuyển sang **xanh lá** = sẵn sàng.

### Giao diện chính

```
┌──────────────────────────────────────────────────────────────────┐
│  Outlook AI Assistant  •  Powered by Claude           [Status]   │
├─────────────────────────┬────────────────────────────────────────┤
│  📁 [Chọn thư mục ▼]    │                                        │
│  🔍 [Search bar]        │  Chi tiết email được chọn              │
│  ┌─────────────────────┐│  ─────────────────────────────────     │
│  │AI │ Người gửi│ Chủ đề││                                        │
│  │🔴 │ ...      │ ...   ││  Kết quả AI                            │
│  │🟡 │ ...      │ ...   ││  ─────────────────────────────────     │
│  │🟢 │ ...      │ ...   ││  Output từ tính năng AI                │
│  └─────────────────────┘│                                        │
│  [Làm mới]              │                                        │
├─────────────────────────┴────────────────────────────────────────┤
│ [Phân loại] [★ Phân loại tất cả] [Tạo Task] [Tạo Lịch]          │
│ [Tóm tắt Thread] [Viết lại (VI)] [Viết lại (EN)] [📅 Lịch ngày] │
│ [🔍 Quét Spam/NL] [🗑️ Xóa Spam] [📰 Chuyển Newsletter]          │
│ [📂 Tổ chức Email] [📦 Archive Cũ] [💾 Kiểm tra PST]    [====]  │
└──────────────────────────────────────────────────────────────────┘
```

### Hướng dẫn từng tính năng

#### 0. Chọn thư mục (multi-PST)
> Duyệt email từ bất kỳ thư mục nào trong mọi tài khoản / file PST đang mở trong Outlook.

1. Nhấn vào dropdown **📁 [Chọn thư mục]** ở góc trái trên
2. Danh sách hiển thị tất cả thư mục dạng `Tài khoản  ›  Đường dẫn` (ví dụ: `Viettel  ›  Inbox / Projects`)
3. Chọn thư mục → danh sách email tự động tải lại
4. Dùng **[Làm mới]** để reload email từ thư mục hiện tại

#### 1. Phân loại Email (đơn lẻ)
> Phân tích email và xác định mức độ ưu tiên, danh mục, hành động cần làm.

1. Click vào email trong danh sách
2. Nhấn **[Phân loại Email]**
3. Kết quả hiển thị trong khung phải và cột **AI** trong danh sách:
   - 🔴 **Urgent** / 🟡 **Normal** / 🟢 **Low**
   - Danh mục: Work / Personal / Newsletter / Finance / HR / IT
   - Hành động: Reply-Needed / Meeting-Request / Task-Request / Read-Only

#### 2. ★ Phân loại tất cả (batch)
> Phân loại toàn bộ email đang hiển thị trong danh sách chỉ với một cú click.

1. Tải thư mục cần phân loại (dùng dropdown hoặc Làm mới)
2. Nhấn **[★ Phân loại tất cả]**
3. AI lần lượt phân tích từng email — cột **AI** cập nhật real-time:
   - 🔴 Urgent · 🟡 Normal · 🟢 Low
4. Kết quả lưu trong bộ nhớ phiên làm việc, không mất khi cuộn danh sách

> **Lưu ý**: Batch classification gọi API cho mỗi email; số lượng lớn sẽ tốn thời gian và token.

#### 3. Tạo Task (đơn lẻ hoặc nhiều email)
> Trích xuất công việc cần làm từ email và tạo thẳng vào Outlook Tasks.

**Một email:**
1. Click chọn email
2. Nhấn **[Tạo Task]** → AI nhận diện task, deadline, mức ưu tiên
3. Task xuất hiện ngay trong **Outlook → Tasks**

**Nhiều email:**
1. Giữ **Ctrl** + click để chọn nhiều email riêng lẻ
2. Giữ **Shift** + click để chọn một dải email liên tiếp
3. Nhấn **[Tạo Task]** → AI xử lý tuần tự từng email, tạo task riêng cho mỗi cái

#### 4. Tạo Lịch họp (đơn lẻ hoặc nhiều email)
> Phát hiện thông tin cuộc họp và tạo sự kiện vào Outlook Calendar.

**Một email:**
1. Chọn email mời họp → **[Tạo Lịch họp]**
2. AI trích xuất: tiêu đề, thời gian bắt đầu/kết thúc, địa điểm, người tham dự
3. Sự kiện xuất hiện ngay trong **Outlook → Calendar**

**Nhiều email:**
1. Ctrl+click hoặc Shift+click để chọn nhiều email mời họp
2. Nhấn **[Tạo Lịch họp]** → AI tạo sự kiện riêng cho từng email

#### 5. Tóm tắt Thread *(Nâng cao)*
> Tóm tắt toàn bộ luồng hội thoại email thành bản tóm tắt có cấu trúc.

1. Chọn bất kỳ email trong luồng
2. Nhấn **[Tóm tắt Thread]**
3. Nhận bản tóm tắt gồm: chủ đề chính, điểm chính, quyết định, hành động tiếp theo

#### 6. Viết lại Email *(Nâng cao)*
> Viết lại nội dung email theo văn phong chuyên nghiệp.

1. Chọn email cần viết lại
2. Nhấn **[Viết lại (VI)]** để viết lại bằng **tiếng Việt** chuyên nghiệp
3. Hoặc **[Viết lại (EN)]** để viết lại bằng **tiếng Anh** chuyên nghiệp
4. Copy nội dung từ khung kết quả để sử dụng

#### 7. Gợi ý Lịch ngày *(Nâng cao)*
> Đề xuất kế hoạch làm việc tối ưu cho ngày hôm nay.

1. Nhấn **[📅 Gợi ý lịch ngày]**
2. Nhập ghi chú bổ sung nếu có (họp cố định, deadline quan trọng...)
3. Nhận lịch làm việc theo khung giờ buổi sáng / chiều / tối

#### 8. Quét Spam / Newsletter
> AI phân tích từng email trong danh sách và phân loại thành 3 nhóm.

1. Tải thư mục cần quét
2. Nhấn **[🔍 Quét Spam/NL]**
3. AI lần lượt phân tích từng email, kết quả hiển thị khi xong:
   - 🔴 Spam / Quảng cáo
   - 📰 Newsletter / Bản tin
   - ✅ Email bình thường
4. Dùng **[🗑️ Xóa Spam]** hoặc **[📰 Chuyển Newsletter]** để xử lý

> **Lưu ý**: Quét gọi AI cho mỗi email. Nên quét sau khi đã lọc thư mục cần xử lý.

#### 9. Xóa Spam
> Chuyển email spam vào Deleted Items (không xóa vĩnh viễn).

1. Đã **Quét Spam/NL** trước
2. Nhấn **[🗑️ Xóa Spam]** → hộp thoại xác nhận hiện số lượng
3. Xác nhận → email spam chuyển vào Deleted Items
4. Danh sách email tự động làm mới

#### 10. Chuyển Newsletter
> Tự động phân loại newsletter vào thư mục con theo tổ chức và người gửi.

1. Đã **Quét Spam/NL** trước
2. Nhấn **[📰 Chuyển Newsletter]** → xác nhận
3. Ứng dụng tạo thư mục theo cấu trúc:
   - Email từ domain công ty: `Newsletter / TênCôngTy`
   - Email từ domain cá nhân (Gmail, Yahoo...): `Newsletter / Gmail / TênNgườiGửi`

#### 11. Tổ chức Email
> Sắp xếp email theo người gửi và năm vào thư mục `Organized`.

1. Tải thư mục cần tổ chức
2. Nhấn **[📂 Tổ chức Email]** → xem trước kế hoạch di chuyển
3. **Xác nhận** → ứng dụng di chuyển email vào:
   - Domain công ty: `Organized / TênCôngTy / Năm`
   - Domain cá nhân: `Organized / Gmail / TênNgườiGửi / Năm`
4. Sau khi hoàn tất, hiển thị danh sách Outlook Rules hiện có

> **Quan trọng**: AI luôn hỏi xác nhận trước khi di chuyển — không tự động move.

#### 12. Archive Email Cũ
> Chuyển email cũ hơn 2 năm vào các file PST riêng theo từng năm.

1. Tải thư mục cần archive
2. Nhấn **[📦 Archive Cũ]** → xem trước số lượng email theo năm
3. Xác nhận → chọn **thư mục** lưu các file PST
4. Ứng dụng tạo một file PST riêng cho mỗi năm:
   ```
   📁 Thư mục đã chọn/
   ├── Outlook_Archive_2023.pst  (email năm 2023)
   ├── Outlook_Archive_2022.pst  (email năm 2022)
   └── Outlook_Archive_2021.pst  (email năm 2021)
   ```
5. Mỗi file PST được tự động thêm vào Outlook, email chuyển vào thư mục `Archive`

> **Lưu ý**: File PST được tự động mở trong Outlook nếu chưa có. Mỗi năm một file giúp dễ quản lý và backup riêng lẻ.

#### 13. Kiểm tra PST
> Xem kích thước tất cả file PST và nhận cảnh báo nếu gần đầy.

1. Nhấn **[💾 Kiểm tra PST]**
2. Kết quả hiển thị thanh tiến trình kích thước mỗi store:
   ```
   TênPST   [████████████░░░░░░░░]  30.5 GB (61%)
   ```
3. Cảnh báo tự động:
   - ⚠️ **47 GB**: Sắp đầy, nên archive bớt email
   - ⛔ **50 GB**: Nguy hiểm, Outlook có thể báo lỗi hoặc hỏng dữ liệu

> Ứng dụng cũng tự động kiểm tra kích thước PST mỗi khi tải email (cảnh báo thụ động).

### Xử lý lỗi thường gặp

| Thông báo | Nguyên nhân | Giải pháp |
|-----------|-------------|-----------|
| `ANTHROPIC_API_KEY is not set` | Chưa tạo file `.env` | Copy `.env.example` → `.env`, điền API key |
| `Lỗi kết nối Outlook` | Outlook chưa mở | Mở Microsoft Outlook trước khi chạy |
| `Vui lòng chọn một email trước` | Chưa chọn email | Click vào email trong danh sách |
| `Dịch vụ chưa khởi tạo xong` | App đang boot | Chờ status bar chuyển xanh lá |
| `Cannot read folder '...'` | PST bị detach hoặc lỗi | Kiểm tra Outlook đã load đủ PST chưa |
| `Chưa có kết quả quét` | Chưa chạy Quét Spam/NL | Nhấn [🔍 Quét Spam/NL] trước |

---

## FAQ

**Q: Ứng dụng có hỗ trợ nhiều tài khoản Outlook không?**
A: Có. Dropdown "Chọn thư mục" liệt kê tất cả thư mục từ mọi tài khoản và file PST đang mở trong Outlook — bao gồm Exchange, Gmail (qua IMAP), và file `.pst` thêm thủ công.

**Q: Phân loại hàng loạt mất bao lâu?**
A: Khoảng 2–5 giây mỗi email tùy độ dài. 50 email ≈ 2–4 phút. Kết quả hiện real-time trong cột AI khi xử lý xong từng email.

**Q: Task và lịch họp được tạo ở đâu?**
A: Trực tiếp trong Outlook — Task vào **Outlook → Tasks**, lịch họp vào **Outlook → Calendar** của tài khoản mặc định.

**Q: Dữ liệu email có gửi ra ngoài không?**
A: Nội dung email (subject + body, tối đa ~400 ký tự cho spam scan, ~3000 ký tự cho các tính năng khác) được gửi đến Anthropic API để phân tích. Không lưu trên server của Anthropic sau khi xử lý. Xem [Anthropic Privacy Policy](https://www.anthropic.com/privacy).

**Q: Có thể dùng với file PST archive không?**
A: Có, miễn là file PST đó đang được mở trong Outlook (File → Open & Export → Open Outlook Data File). Ứng dụng tự động phát hiện tất cả PST đang active.

**Q: Tại sao một số thư mục không hiện trong dropdown?**
A: Dropdown chỉ hiển thị thư mục chứa **email** (`DefaultItemType = 0`). Các thư mục Calendar, Tasks, Contacts bị ẩn vì không phù hợp để đọc email.

**Q: Ứng dụng có lưu kết quả phân loại không?**
A: Kết quả chỉ tồn tại trong phiên làm việc hiện tại (lưu trong RAM). Khi đóng ứng dụng, nhãn ưu tiên sẽ mất. Tính năng lưu persistent là roadmap tương lai.

**Q: Newsletter folder được tạo ở đâu?**
A: Trong cùng store (PST/tài khoản) với thư mục đang xem. Cấu trúc: `Newsletter / TênCôngTy` cho domain công ty, `Newsletter / Gmail / TênNgườiGửi` cho Gmail/Yahoo/... Thư mục con được tạo tự động nếu chưa có.

**Q: Archive tạo bao nhiêu file PST?**
A: Mỗi năm một file riêng — ví dụ email năm 2023 vào `Outlook_Archive_2023.pst`, năm 2022 vào `Outlook_Archive_2022.pst`. Tất cả được tạo trong thư mục bạn chọn và tự động thêm vào Outlook profile.

**Q: Giới hạn kích thước PST là bao nhiêu?**
A: Outlook hỗ trợ PST tối đa 50 GB (Unicode format). Ứng dụng cảnh báo ở **47 GB** và báo nguy hiểm ở **50 GB**. Nên dùng tính năng Archive để giữ PST dưới 40 GB.

---

## Cấu trúc dự án

```
Outlook_AI/
├── main.py              # GUI (tkinter)
├── config.py            # Cấu hình
├── ai_client.py         # Claude API wrapper
├── outlook_client.py    # Outlook COM wrapper
├── features/            # Feature modules
│   ├── email_classifier.py
│   ├── task_creator.py
│   ├── calendar_creator.py
│   ├── email_summarizer.py
│   ├── email_rewriter.py
│   ├── scheduler.py
│   ├── spam_cleaner.py   # Spam / newsletter scanner
│   └── email_organizer.py # Organize, archive, PST monitor
└── docs/                # Tài liệu dự án
```

## Tech Stack
- **Python** 3.10+
- **LLM**: Claude Opus 4.6 (Anthropic)
- **Outlook integration**: pywin32 (COM/MAPI) — hỗ trợ multi-PST
- **GUI**: tkinter (dark theme, Catppuccin palette)

## License

This project is licensed under the **GNU General Public License v3.0**.
See the [LICENSE](LICENSE) file for details.

> You are free to use, modify, and distribute this software under the terms of GPL-3.0.
> Any derivative work must also be distributed under the same license.
