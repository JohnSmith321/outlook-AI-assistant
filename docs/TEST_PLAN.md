# Kế hoạch Kiểm thử – Outlook AI Assistant

## Mục tiêu kiểm thử

Đảm bảo tất cả tính năng hoạt động đúng, ổn định và thân thiện với người dùng trong môi trường Windows với Microsoft Outlook đã cài đặt.

---

## Phạm vi kiểm thử

### Trong phạm vi
- Kết nối và đọc email từ Outlook Inbox
- 6 tính năng AI (classify, task, calendar, summarize, rewrite, schedule)
- GUI interaction và display
- Error handling

### Ngoài phạm vi
- Gửi email tự động
- Đọc email từ shared mailbox / delegated mailbox
- Kiểm thử hiệu năng dưới tải cao
- Bảo mật penetration testing

---

## Test Cases

### TC-01: Khởi động ứng dụng

| Bước | Hành động | Kết quả mong đợi |
|------|-----------|-----------------|
| 1 | Chạy `python main.py` | Cửa sổ GUI mở ra |
| 2 | Chờ kết nối | Status bar hiển thị "Đang kết nối Outlook..." |
| 3 | Sau ~3s | Status bar hiển thị "Sẵn sàng" màu xanh |
| 4 | Kiểm tra email list | Danh sách email từ Inbox hiển thị trong Treeview |

**Điều kiện tiên quyết**: Outlook đang mở, `.env` có API key hợp lệ

---

### TC-02: Khởi động không có API Key

| Bước | Hành động | Kết quả mong đợi |
|------|-----------|-----------------|
| 1 | Xóa hoặc để trống `ANTHROPIC_API_KEY` trong `.env` | – |
| 2 | Chạy `main.py` | Status bar hiển thị lỗi màu đỏ |
| 3 | Nhấn nút "Phân loại Email" | Thông báo "Dịch vụ chưa khởi tạo xong" |

---

### TC-03: Đọc và hiển thị email

| Bước | Hành động | Kết quả mong đợi |
|------|-----------|-----------------|
| 1 | Ứng dụng đã khởi động | Email list có dữ liệu |
| 2 | Click vào một email | Chi tiết email hiển thị bên phải (From, Subject, Body) |
| 3 | Nhập text vào ô search | Danh sách lọc real-time theo sender/subject |
| 4 | Xóa text search | Danh sách trở về đầy đủ |
| 5 | Click "Làm mới" | Reload email từ Outlook |

---

### TC-04: Phân loại email

**Email test 1 – Khẩn cấp từ sếp**:
```
Subject: [KHẨN] Báo cáo Q1 nộp trước 5pm hôm nay
Body: Anh/chị cần nộp báo cáo quý 1 trước 5 giờ chiều hôm nay.
```
Kết quả mong đợi: `priority=Urgent, action=Task-Request, category=Work`

**Email test 2 – Newsletter**:
```
Subject: Bản tin công nghệ tuần này
Body: Chào bạn, đây là bản tin công nghệ hàng tuần...
```
Kết quả mong đợi: `priority=Low, action=Read-Only, category=Newsletter`

**Email test 3 – Mời họp**:
```
Subject: Họp team dự án ABC – Thứ Tư 10h
Body: Kính mời anh/chị tham dự cuộc họp team dự án ABC vào lúc 10:00 sáng thứ Tư...
```
Kết quả mong đợi: `action=Meeting-Request, priority=Normal/Urgent`

---

### TC-05: Tạo Task từ email

**Email test**:
```
Subject: Nhiệm vụ tháng 3
Body: Bạn cần hoàn thành 3 việc sau trước ngày 31/3:
      1. Cập nhật tài liệu kỹ thuật dự án X
      2. Review code PR #42
      3. Gửi báo cáo tiến độ cho khách hàng
```

| Bước | Kiểm tra |
|------|---------|
| Click "Tạo Task" | Output hiển thị 3 task với tên rõ ràng |
| Mở Outlook Tasks | 3 task mới xuất hiện trong danh sách |
| Kiểm tra task 1 | Subject, body, due date, category đúng |
| Email không có task | Output: "Không tìm thấy task nào" |

---

### TC-06: Tạo Lịch họp từ email

**Email test**:
```
Subject: Họp kickoff dự án AI
Body: Kính mời tham dự buổi họp kickoff dự án AI Assistant vào:
      - Thời gian: Thứ Hai, 15/3/2025 lúc 14:00
      - Địa điểm: Phòng họp B2 / Link Zoom: https://zoom.us/j/...
      - Thời lượng: 90 phút
      - Tham dự: nguyen@company.com, tran@company.com
```

| Bước | Kiểm tra |
|------|---------|
| Click "Tạo Lịch họp" | Output hiển thị thông tin sự kiện |
| Mở Outlook Calendar | Sự kiện mới xuất hiện đúng ngày/giờ |
| Kiểm tra sự kiện | Subject, location, attendees đúng |
| Email thông thường (không có meeting) | Output: "Không phát hiện thông tin cuộc họp" |

---

### TC-07: Tóm tắt email

**Email test – Đơn lẻ**:
- Email dài với nhiều nội dung → Summary ngắn gọn, có cấu trúc
- Phải có: Chủ đề chính, Điểm chính, Hành động tiếp theo

**Email test – Luồng (Thread)**:
- Email có conversation topic với 3+ email → Tóm tắt toàn bộ luồng
- Hiển thị "X email trong luồng"

---

### TC-08: Viết lại email

**Test Vietnamese**:
- Input: Email thông thường với ngôn ngữ thông thường
- Output: Email có kính ngữ, câu văn rõ ràng, kết thúc bằng "Trân trọng"

**Test English**:
- Input: Email tiếng Việt hoặc tiếng Anh thô
- Output: Professional English email với "Dear", "Sincerely"

---

### TC-09: Gợi ý lịch ngày

| Bước | Kiểm tra |
|------|---------|
| Click "Gợi ý lịch ngày" | Hộp thoại nhập ghi chú xuất hiện |
| Nhập "Họp 9h, deadline báo cáo lúc 15h" | – |
| Click OK | Output hiển thị lịch theo thời gian buổi sáng/chiều |
| Kiểm tra nội dung | Ưu tiên đúng, gợi ý hợp lý, không nhồi quá nhiều |

---

### TC-10: Kiểm thử Edge Cases

| Trường hợp | Kết quả mong đợi |
|------------|-----------------|
| Email có body rỗng | Không crash, hiển thị thông báo phù hợp |
| Email rất dài (>10.000 chars) | Truncate an toàn, không lỗi token |
| Click nút khi chưa chọn email | Popup "Vui lòng chọn một email trước" |
| Claude API timeout | Hiển thị lỗi, không hang |
| Outlook bị đóng giữa chừng | Hiển thị lỗi COM, không crash |
| Inbox rỗng (0 email) | Treeview rỗng, không crash |

---

### TC-11: Kiểm thử GUI

| Thao tác | Kết quả mong đợi |
|---------|-----------------|
| Thay đổi kích thước cửa sổ | Panels co giãn đúng |
| Kéo PanedWindow divider | Email list và output thay đổi kích thước |
| Click nhanh nhiều nút | Không crash, chỉ chạy task cuối |
| Scroll email list dài | Scroll mượt, không lag |

---

## Môi trường kiểm thử

| Môi trường | Chi tiết |
|-----------|----------|
| OS | Windows 10/11 (64-bit) |
| Python | 3.10, 3.11, 3.12 |
| Outlook | Microsoft 365 Apps, Outlook 2019 |
| Anthropic API | claude-opus-4-6 (production) |

---

## Tiêu chí Pass/Fail

| Metric | Pass |
|--------|------|
| TC cơ bản (TC-03 đến TC-06) | 100% pass |
| TC nâng cao (TC-07 đến TC-09) | ≥ 90% pass |
| TC edge cases (TC-10) | ≥ 80% pass |
| Không có crash trong 30 phút demo | 100% |
| Thời gian response trung bình | < 15 giây |

---

## Báo cáo Bug

Khi phát hiện bug, ghi lại:
1. **Môi trường**: OS, Python version, Outlook version
2. **Steps to reproduce**: Các bước tái hiện lỗi
3. **Expected**: Kết quả mong đợi
4. **Actual**: Kết quả thực tế
5. **Screenshot/Log**: Nếu có
6. **Severity**: Critical / High / Medium / Low
