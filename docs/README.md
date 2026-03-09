# Outlook AI Assistant

Trợ lý AI tích hợp với Microsoft Outlook, giúp tự động hóa các công việc email hàng ngày thông qua sức mạnh của Claude (Anthropic).

---

## Tính năng

### Tính năng cơ bản
| Tính năng | Mô tả |
|-----------|-------|
| **Phân loại Email** | Tự động phân tích và phân loại email theo mức độ ưu tiên, danh mục, và hành động cần thực hiện |
| **Tạo Task** | Trích xuất các công việc cần làm từ email và tạo task trực tiếp trong Outlook |
| **Tạo Lịch họp** | Phát hiện thông tin cuộc họp trong email và tạo sự kiện vào Outlook Calendar |

### Tính năng nâng cao
| Tính năng | Mô tả |
|-----------|-------|
| **Tóm tắt Thread** | Tóm tắt toàn bộ luồng email một cách có cấu trúc (điểm chính, quyết định, hành động tiếp theo) |
| **Viết lại Email (VI/EN)** | Viết lại nội dung email theo văn phong chuyên nghiệp bằng tiếng Việt hoặc tiếng Anh |
| **Gợi ý lịch ngày** | Đề xuất kế hoạch làm việc tối ưu cho ngày dựa trên email và task hiện có |

---

## Yêu cầu hệ thống

- **OS**: Windows 10/11
- **Python**: 3.10 trở lên
- **Microsoft Outlook**: đã cài đặt và đăng nhập tài khoản
- **Anthropic API Key**: đăng ký tại [console.anthropic.com](https://console.anthropic.com)

---

## Cài đặt

### 1. Clone repository
```bash
git clone https://github.com/<your-username>/outlook-ai-assistant.git
cd outlook-ai-assistant
```

### 2. Tạo virtual environment
```bash
python -m venv venv
venv\Scripts\activate
```

### 3. Cài đặt dependencies
```bash
pip install -r requirements.txt
```

### 4. Cấu hình API Key
```bash
copy .env.example .env
# Mở file .env và điền ANTHROPIC_API_KEY của bạn
```

Nội dung file `.env`:
```
ANTHROPIC_API_KEY=sk-ant-api03-...
```

### 5. Chạy ứng dụng
```bash
python main.py
```

---

## Hướng dẫn sử dụng

### Giao diện chính

```
┌─────────────────────────────────────────────────────────────┐
│  Outlook AI Assistant  •  Powered by Claude         Status  │
├──────────────────────┬──────────────────────────────────────┤
│  Hộp thư đến         │  Chi tiết email                     │
│  ┌────────────────┐  │  ─────────────────────────────────  │
│  │ Email list     │  │  Nội dung email được chọn           │
│  │ (treeview)     │  ├──────────────────────────────────── │
│  │                │  │  Kết quả AI                         │
│  └────────────────┘  │  Output từ các tính năng AI         │
├──────────────────────┴──────────────────────────────────────┤
│ [Phân loại] [Tạo Task] [Tạo Lịch] [Tóm tắt] [VI] [EN] [📅]│
└─────────────────────────────────────────────────────────────┘
```

### Các bước sử dụng

1. **Mở ứng dụng** → Outlook và Claude AI sẽ tự động kết nối
2. **Chọn email** từ danh sách bên trái
3. **Nhấn nút** tương ứng với tính năng muốn sử dụng
4. **Xem kết quả** trong khung "Kết quả AI" bên phải

### Chi tiết các nút

| Nút | Tính năng | Đầu ra |
|-----|-----------|--------|
| **Phân loại Email** | Phân tích email đang chọn | Mức ưu tiên, danh mục, hành động gợi ý |
| **Tạo Task** | Tạo task Outlook từ email | Danh sách task đã tạo trong Outlook |
| **Tạo Lịch họp** | Tạo sự kiện từ email | Sự kiện đã tạo trong Outlook Calendar |
| **Tóm tắt Thread** | Tóm tắt luồng email | Tóm tắt có cấu trúc bằng tiếng Việt |
| **Viết lại (VI)** | Viết lại bằng tiếng Việt | Nội dung email chuyên nghiệp |
| **Viết lại (EN)** | Viết lại bằng tiếng Anh | Professional English email |
| **Gợi ý lịch ngày** | Lên kế hoạch hôm nay | Lịch làm việc theo giờ |

---

## Cấu trúc dự án

```
Outlook_AI/
├── main.py                    # GUI chính (tkinter)
├── config.py                  # Cấu hình ứng dụng
├── ai_client.py               # Wrapper Claude API
├── outlook_client.py          # Wrapper Outlook COM (pywin32)
├── features/
│   ├── __init__.py
│   ├── email_classifier.py    # Phân loại email
│   ├── task_creator.py        # Tạo Outlook task
│   ├── calendar_creator.py    # Tạo Outlook calendar event
│   ├── email_summarizer.py    # Tóm tắt luồng email
│   ├── email_rewriter.py      # Viết lại email
│   └── scheduler.py           # Gợi ý lịch ngày
├── requirements.txt
├── .env.example
└── docs/
    ├── README.md
    ├── ARCHITECTURE.md
    ├── IMPLEMENTATION_PLAN.md
    ├── TEST_PLAN.md
    └── DEV_DIARY.md
```

---

## Xử lý lỗi thường gặp

| Lỗi | Nguyên nhân | Giải pháp |
|-----|-------------|-----------|
| `ANTHROPIC_API_KEY is not set` | Chưa tạo file `.env` | Copy `.env.example` → `.env` và điền API key |
| `pywin32 is not installed` | Thiếu dependency | `pip install pywin32` |
| `Lỗi kết nối Outlook` | Outlook chưa mở | Mở Microsoft Outlook trước khi chạy app |
| `Rate limit error` | Gọi API quá nhiều | Chờ vài giây và thử lại |

---

## License

MIT License – xem file `LICENSE` để biết thêm chi tiết.
