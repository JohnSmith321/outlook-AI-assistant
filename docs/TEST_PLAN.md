# Kế hoạch Kiểm thử – Outlook AI Assistant

## Mục tiêu kiểm thử

Đảm bảo tất cả tính năng hoạt động đúng, ổn định và thân thiện với người dùng trong môi trường Windows với Microsoft Outlook đã cài đặt.

---

## Phạm vi kiểm thử

### Trong phạm vi
- Kết nối và đọc email từ Outlook Inbox
- 13 tính năng: classify, task, calendar, summarize, rewrite (VI/EN), schedule, spam scan, delete spam, move newsletter, organize, archive, PST check
- GUI interaction và display
- Error handling và edge cases

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

### TC-10: Quét Spam & Newsletter

**Chuẩn bị**: Inbox có ít nhất 5 email gồm 2 spam, 2 newsletter, 1 email bình thường

| Bước | Hành động | Kết quả mong đợi |
|------|-----------|-----------------|
| 1 | Click "Quét Spam/Newsletter" | Progress bar chạy |
| 2 | Chờ scan hoàn tất | Output hiển thị báo cáo: X spam, Y newsletter, Z normal |
| 3 | Kiểm tra email spam | Đúng là email quảng cáo/lừa đảo |
| 4 | Kiểm tra email newsletter | Đúng là email bản tin định kỳ |
| 5 | Nút "Xóa Spam" và "Di chuyển Newsletter" | Được enable sau khi quét |

**Email test – Spam**:
```
Subject: Bạn đã trúng thưởng 1 tỷ đồng!
Body: Click vào link sau để nhận thưởng...
```
Kết quả mong đợi: Phân loại là `spam`

**Email test – Newsletter**:
```
Subject: Weekly Digest – Tech News
From: newsletter@techblog.com
Body: Đây là bản tin hàng tuần của TechBlog...
```
Kết quả mong đợi: Phân loại là `newsletter`

---

### TC-11: Xóa Spam

**Điều kiện tiên quyết**: Đã chạy TC-10, có ít nhất 1 email spam

| Bước | Hành động | Kết quả mong đợi |
|------|-----------|-----------------|
| 1 | Click "Xóa Spam" | Hộp xác nhận xuất hiện |
| 2 | Xác nhận xóa | Progress bar chạy |
| 3 | Hoàn tất | Output: "Đã xóa X email spam" |
| 4 | Kiểm tra Outlook | Các email spam không còn trong Inbox |
| 5 | Chưa quét → click "Xóa Spam" | Nút bị disabled / thông báo "Chưa quét" |

---

### TC-12: Di chuyển Newsletter

**Điều kiện tiên quyết**: Đã chạy TC-10, có ít nhất 1 email newsletter

| Bước | Hành động | Kết quả mong đợi |
|------|-----------|-----------------|
| 1 | Click "Di chuyển Newsletter" | Progress bar chạy |
| 2 | Hoàn tất | Output: "Đã di chuyển X email newsletter" |
| 3 | Kiểm tra Outlook folder | Folder `Newsletter/` được tạo trong Inbox |
| 4 | Email từ newsletter@techblog.com | Nằm trong `Newsletter/TechBlog/newsletter/` |
| 5 | Email từ abc@gmail.com | Nằm trong `Newsletter/Gmail/abc/` |

---

### TC-13: Sắp xếp Email

| Bước | Hành động | Kết quả mong đợi |
|------|-----------|-----------------|
| 1 | Click "Sắp xếp Email" | Output hiển thị preview: X email → Y nhóm |
| 2 | Xác nhận | Progress bar chạy, email được di chuyển |
| 3 | Email từ user@companyabc.vn | Vào `Organized/Companyabc/2025/` |
| 4 | Email từ sender@gmail.com | Vào `Organized/Gmail/sender/2025/` |
| 5 | Email từ info@yahoo.com | Vào `Organized/Yahoo/info/2025/` |
| 6 | Inbox sau khi sắp xếp | Ít email hơn, email đã được phân loại vào subfolder |

---

### TC-14: Archive Email cũ

**Điều kiện tiên quyết**: Inbox có email từ năm 2023 trở về trước

| Bước | Hành động | Kết quả mong đợi |
|------|-----------|-----------------|
| 1 | Click "Archive Email cũ" | Output hiển thị preview: X email → các năm 2021, 2022, 2023 |
| 2 | Hiển thị danh sách file sẽ tạo | `Outlook_Archive_2023.pst`, `Outlook_Archive_2022.pst`, ... |
| 3 | Hộp thoại chọn thư mục xuất hiện | `filedialog.askdirectory()` |
| 4 | Chọn thư mục và xác nhận | Progress bar chạy |
| 5 | Kiểm tra thư mục đã chọn | Xuất hiện file `Outlook_Archive_2023.pst` |
| 6 | Mở file PST trong Outlook | Có folder `Archive/TênCôngTy/` hoặc `Archive/Gmail/localpart/` |
| 7 | Email cùng sender trong 1 folder | Không tạo folder trùng (dùng email address, không dùng display name) |
| 8 | Email năm 2022 trong PST 2023 | Không có (mỗi PST chỉ chứa đúng năm đó) |
| 9 | Hủy chọn thư mục | Không có file nào được tạo |

---

### TC-15: Kiểm tra kích thước PST

| Bước | Hành động | Kết quả mong đợi |
|------|-----------|-----------------|
| 1 | Click "Kiểm tra PST" | Output hiển thị bảng kích thước tất cả PST/mailbox |
| 2 | Kiểm tra bar chart | Thanh bar tỷ lệ với kích thước thực |
| 3 | PST nhỏ hơn 47 GB | Hiển thị màu xanh/bình thường |
| 4 | PST từ 47-50 GB | Hiển thị ⚠️ cảnh báo vàng |
| 5 | PST lớn hơn 50 GB | Hiển thị ⛔ cảnh báo đỏ + popup |
| 6 | Passive check sau reload | Nếu vượt ngưỡng → popup tự động xuất hiện |

---

### TC-16: Kiểm thử Edge Cases

| Trường hợp | Kết quả mong đợi |
|------------|-----------------|
| Email có body rỗng | Không crash, hiển thị thông báo phù hợp |
| Email rất dài (>10.000 chars) | Truncate an toàn, không lỗi token |
| Click nút khi chưa chọn email | Popup "Vui lòng chọn một email trước" |
| Claude API timeout | Hiển thị lỗi, không hang |
| Outlook bị đóng giữa chừng | Hiển thị lỗi COM, không crash |
| Inbox rỗng (0 email) | Treeview rỗng, không crash |
| Archive khi không có email cũ | Output: "Không có email nào cần archive" |
| Chọn cùng thư mục archive 2 lần | PST mở lại, không tạo duplicate |
| Scan spam với inbox rỗng | Output: "Không có email nào để quét" |

---

### TC-17: Kiểm thử GUI

| Thao tác | Kết quả mong đợi |
|---------|-----------------|
| Thay đổi kích thước cửa sổ | Panels co giãn đúng |
| Kéo PanedWindow divider | Email list và output thay đổi kích thước |
| Click nhanh nhiều nút | Không crash, chỉ chạy task cuối |
| Scroll email list dài | Scroll mượt, không lag |
| 4 hàng nút (action bar) | Tất cả nút hiển thị đủ, không bị cắt |

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
| TC spam/newsletter (TC-10 đến TC-12) | ≥ 90% pass |
| TC organizer/archive/PST (TC-13 đến TC-15) | ≥ 95% pass |
| TC edge cases (TC-16) | ≥ 80% pass |
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
