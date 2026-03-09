# Outlook AI Assistant

Trợ lý AI tích hợp với Microsoft Outlook, giúp tự động hóa các công việc email hàng ngày thông qua sức mạnh của Claude (Anthropic).

---

## Tính năng

### Phân tích & Soạn thảo (AI-driven)

| Tính năng | Mô tả |
|-----------|-------|
| **1. Phân loại Email** | Tự động phân tích và phân loại email theo mức độ ưu tiên (Urgent/Normal/Low), danh mục (Work/Personal/Newsletter), và hành động cần thực hiện |
| **2. Tạo Task** | Trích xuất các công việc cần làm từ email và tạo task trực tiếp trong Outlook Tasks với due date và mô tả đầy đủ |
| **3. Tạo Lịch họp** | Phát hiện thông tin cuộc họp trong email (thời gian, địa điểm, người tham dự) và tạo sự kiện vào Outlook Calendar |
| **4. Tóm tắt Email/Thread** | Tóm tắt email đơn lẻ hoặc toàn bộ luồng hội thoại thành các điểm chính, quyết định và hành động tiếp theo |
| **5. Viết lại Email (VI)** | Viết lại nội dung email theo văn phong chuyên nghiệp bằng tiếng Việt (kính ngữ, câu văn rõ ràng) |
| **6. Viết lại Email (EN)** | Viết lại nội dung email theo văn phong chuyên nghiệp bằng tiếng Anh (Dear/Sincerely format) |
| **7. Gợi ý lịch ngày** | Đề xuất kế hoạch làm việc tối ưu cho ngày dựa trên email và task hiện có, hỗ trợ ghi chú bổ sung |

### Quản lý & Dọn dẹp

| Tính năng | Mô tả |
|-----------|-------|
| **8. Quét Spam/Newsletter** | Dùng AI để quét và phân loại toàn bộ email trong danh sách thành: spam, newsletter, hoặc email bình thường |
| **9. Xóa Spam** | Xóa vĩnh viễn tất cả email được phân loại là spam sau khi đã quét (yêu cầu quét trước) |
| **10. Di chuyển Newsletter** | Di chuyển email newsletter vào thư mục phân loại theo tên người gửi: `Newsletter/Gmail/TênNgườiGửi/` |
| **11. Sắp xếp Email** | Tự động tổ chức email theo người gửi và năm: `Organized/TênCôngTy/Năm/` hoặc `Organized/Gmail/TênNgười/Năm/` |
| **12. Archive Email cũ** | Di chuyển email cũ hơn 2 năm ra các file PST riêng biệt theo từng năm: `Outlook_Archive_2023.pst`, `Outlook_Archive_2022.pst`, ... |
| **13. Kiểm tra kích thước PST** | Hiển thị kích thước tất cả PST/mailbox đang mở với bar chart ASCII, cảnh báo khi gần đạt giới hạn 50 GB |

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
│ Hàng 1: [Phân loại] [Tạo Task] [Tạo Lịch họp] [Tóm tắt]  │
│ Hàng 2: [Viết lại VI] [Viết lại EN] [Gợi ý lịch ngày]     │
│ Hàng 3: [Quét Spam/Newsletter] [Xóa Spam] [Di chuyển NL]  │
│ Hàng 4: [Sắp xếp Email] [Archive Email cũ] [Kiểm tra PST] │
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
| **Quét Spam/Newsletter** | Quét toàn bộ email trong danh sách | Báo cáo: X spam, Y newsletter, Z normal |
| **Xóa Spam** | Xóa email spam (phải quét trước) | Xác nhận số email đã xóa |
| **Di chuyển Newsletter** | Chuyển newsletter vào folder riêng | Cấu trúc Newsletter/Brand/Sender/ |
| **Sắp xếp Email** | Tổ chức email theo sender/năm | Preview → xác nhận → di chuyển |
| **Archive Email cũ** | Lưu trữ email >2 năm ra PST | Một file .pst riêng cho mỗi năm |
| **Kiểm tra PST** | Hiển thị kích thước tất cả PST | Bar chart + cảnh báo nếu >47 GB |

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
│   ├── email_classifier.py    # Tính năng 1: Phân loại email
│   ├── task_creator.py        # Tính năng 2: Tạo Outlook task
│   ├── calendar_creator.py    # Tính năng 3: Tạo Outlook calendar event
│   ├── email_summarizer.py    # Tính năng 4: Tóm tắt email/thread
│   ├── email_rewriter.py      # Tính năng 5+6: Viết lại email VI/EN
│   ├── scheduler.py           # Tính năng 7: Gợi ý lịch ngày
│   ├── spam_cleaner.py        # Tính năng 8+9+10: Quét/Xóa spam, Di chuyển newsletter
│   └── email_organizer.py     # Tính năng 11+12+13: Sắp xếp, Archive, Kiểm tra PST
├── requirements.txt
├── .env.example
└── docs/
    ├── README.md              # Tài liệu này
    ├── ARCHITECTURE.md        # Kiến trúc hệ thống
    ├── IMPLEMENTATION_PLAN.md # Kế hoạch triển khai
    ├── TEST_PLAN.md           # Kế hoạch kiểm thử
    └── DEV_DIARY.md           # Nhật ký phát triển
```

---

## Xử lý lỗi thường gặp

| Lỗi | Nguyên nhân | Giải pháp |
|-----|-------------|-----------|
| `ANTHROPIC_API_KEY is not set` | Chưa tạo file `.env` | Copy `.env.example` → `.env` và điền API key |
| `pywin32 is not installed` | Thiếu dependency | `pip install pywin32` |
| `Lỗi kết nối Outlook` | Outlook chưa mở | Mở Microsoft Outlook trước khi chạy app |
| `Rate limit error` | Gọi API quá nhiều | Chờ vài giây và thử lại |
| `Không tạo được PST` | Không có quyền ghi vào thư mục | Chọn thư mục khác có quyền ghi |

---

## License

MIT License – xem file `LICENSE` để biết thêm chi tiết.
