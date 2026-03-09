# Outlook AI Assistant

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Python](https://img.shields.io/badge/Python-3.10+-green.svg)](https://www.python.org/)
[![Claude](https://img.shields.io/badge/LLM-Claude%20Opus%204.6-orange.svg)](https://www.anthropic.com/)

Trợ lý AI tự động hóa công việc email trong Microsoft Outlook, tích hợp Claude (Anthropic).

## Tính năng
- Tự động phân loại email (ưu tiên, danh mục, hành động)
- Tạo Task Outlook từ nội dung email
- Tạo lịch họp Outlook Calendar từ email mời họp
- Tóm tắt luồng email có cấu trúc *(nâng cao)*
- Viết lại email chuyên nghiệp (tiếng Việt / tiếng Anh) *(nâng cao)*
- Gợi ý kế hoạch làm việc hàng ngày *(nâng cao)*

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
1. Kết nối Microsoft Outlook qua COM
2. Kết nối Claude AI API
3. Tải danh sách email từ Inbox

Thanh trạng thái (góc phải trên) chuyển sang **xanh lá** = sẵn sàng.

### Giao diện chính

```
┌─────────────────────────────────────────────────────────────────┐
│  Outlook AI Assistant  •  Powered by Claude          [Status]   │
├────────────────────────┬────────────────────────────────────────┤
│  🔍 [Search bar]       │                                        │
│  ┌────────────────────┐│  Chi tiết email được chọn             │
│  │ Sender │Subject│Time││  ────────────────────────────────     │
│  │ ...    │ ...   │... ││                                        │
│  │ ...    │ ...   │... ││  Kết quả AI                           │
│  └────────────────────┘│  ────────────────────────────────      │
│  [Làm mới]             │  Output từ tính năng AI               │
├────────────────────────┴────────────────────────────────────────┤
│ [Phân loại] [Tạo Task] [Tạo Lịch] [Tóm tắt] [VI] [EN] [📅]    │
└─────────────────────────────────────────────────────────────────┘
```

### Hướng dẫn từng tính năng

#### 1. Phân loại Email
> Phân tích email và xác định mức độ ưu tiên, danh mục, hành động cần làm.

1. Click vào email trong danh sách
2. Nhấn **[Phân loại Email]**
3. Kết quả hiển thị trong khung phải:
   - 🔴 **Urgent** / 🟡 **Normal** / 🟢 **Low**
   - Danh mục: Work / Personal / Newsletter / Finance / HR / IT
   - Hành động: Reply-Needed / Meeting-Request / Task-Request / Read-Only

#### 2. Tạo Task
> Trích xuất công việc cần làm từ email và tạo thẳng vào Outlook Tasks.

1. Chọn email giao việc / yêu cầu thực hiện
2. Nhấn **[Tạo Task]**
3. AI tự động nhận diện các task, deadline, mức ưu tiên
4. Task xuất hiện ngay trong **Outlook → Tasks**

#### 3. Tạo Lịch họp
> Phát hiện thông tin cuộc họp và tạo sự kiện vào Outlook Calendar.

1. Chọn email mời họp
2. Nhấn **[Tạo Lịch họp]**
3. AI trích xuất: tiêu đề, thời gian bắt đầu/kết thúc, địa điểm, người tham dự
4. Sự kiện xuất hiện ngay trong **Outlook → Calendar**

#### 4. Tóm tắt Thread *(Nâng cao)*
> Tóm tắt toàn bộ luồng hội thoại email thành bản tóm tắt có cấu trúc.

1. Chọn bất kỳ email trong luồng
2. Nhấn **[Tóm tắt Thread]**
3. Nhận bản tóm tắt gồm: chủ đề chính, điểm chính, quyết định, hành động tiếp theo

#### 5. Viết lại Email *(Nâng cao)*
> Viết lại nội dung email theo văn phong chuyên nghiệp.

1. Chọn email cần viết lại
2. Nhấn **[Viết lại (VI)]** để viết lại bằng **tiếng Việt** chuyên nghiệp
3. Hoặc **[Viết lại (EN)]** để viết lại bằng **tiếng Anh** chuyên nghiệp
4. Copy nội dung từ khung kết quả để sử dụng

#### 6. Gợi ý Lịch ngày *(Nâng cao)*
> Đề xuất kế hoạch làm việc tối ưu cho ngày hôm nay.

1. Nhấn **[Gợi ý lịch ngày]**
2. Nhập ghi chú bổ sung nếu có (họp cố định, deadline quan trọng...)
3. Nhận lịch làm việc theo khung giờ buổi sáng / chiều / tối

### Xử lý lỗi thường gặp

| Thông báo | Nguyên nhân | Giải pháp |
|-----------|-------------|-----------|
| `ANTHROPIC_API_KEY is not set` | Chưa tạo file `.env` | Copy `.env.example` → `.env`, điền API key |
| `Lỗi kết nối Outlook` | Outlook chưa mở | Mở Microsoft Outlook trước khi chạy |
| `Vui lòng chọn một email trước` | Chưa chọn email | Click vào email trong danh sách |
| `Dịch vụ chưa khởi tạo xong` | App đang boot | Chờ status bar chuyển xanh lá |

---

## Cấu trúc dự án

```
Outlook_AI/
├── main.py              # GUI (tkinter)
├── config.py            # Cấu hình
├── ai_client.py         # Claude API wrapper
├── outlook_client.py    # Outlook COM wrapper
├── features/            # 6 feature modules
└── docs/                # Tài liệu dự án
```

## Tech Stack
- **Python** 3.10+
- **LLM**: Claude Opus 4.6 (Anthropic)
- **Outlook integration**: pywin32 (COM/MAPI)
- **GUI**: tkinter (dark theme)

## License

This project is licensed under the **GNU General Public License v3.0**.
See the [LICENSE](LICENSE) file for details.

> You are free to use, modify, and distribute this software under the terms of GPL-3.0.
> Any derivative work must also be distributed under the same license.
