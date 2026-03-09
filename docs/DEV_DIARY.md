# Nhật ký Phát triển – Outlook AI Assistant

## Giới thiệu

File này ghi lại hành trình phát triển dự án Outlook AI Assistant — từ ý tưởng ban đầu, những quyết định thiết kế, khó khăn gặp phải và bài học rút ra trong quá trình xây dựng.

---

## Ngày 1 – Thiết lập nền tảng

### Quyết định công nghệ đầu tiên

Khi bắt đầu dự án, câu hỏi đầu tiên là: **dùng giao diện gì?**

Có ba lựa chọn:
1. **tkinter** (built-in) – không cần cài thêm, đơn giản, nhưng giao diện cơ bản
2. **PyQt6** – đẹp hơn nhiều nhưng phải cài thêm (~50MB) và license phức tạp
3. **CLI thuần** – nhanh nhất để code nhưng khó demo

Tôi chọn **tkinter** vì không cần phụ thuộc ngoài và phù hợp với scope demo. Tuy nhiên để GUI trông chuyên nghiệp hơn, tôi thiết kế một dark theme tùy chỉnh dựa trên bảng màu Catppuccin – kết quả khá bất ngờ khi tkinter với màu sắc phù hợp trông không tệ chút nào.

### Vật lộn với pywin32 và Outlook COM

Phần khó nhất của ngày đầu không phải là AI mà là **pywin32**.

**Vấn đề 1: pywintypes.datetime**
Outlook trả về `pywintypes.datetime` cho timestamp, không phải Python's `datetime.datetime`. Chúng trông giống nhau nhưng không compatible. Phải convert thủ công:
```python
dt = datetime.datetime(
    received.year, received.month, received.day,
    received.hour, received.minute, received.second
)
```

**Vấn đề 2: COM object không serializable**
Ban đầu tôi cố gắng pass COM objects giữa các threads. COM objects không thread-safe! Giải pháp: tạo `EmailMessage` dataclass thuần Python ngay khi đọc từ COM, sau đó pass dataclass này đi khắp nơi. COM objects chỉ sống trong một thread duy nhất.

**Bài học**: Luôn deserialize dữ liệu từ COM/external systems thành plain Python objects càng sớm càng tốt.

---

## Ngày 1 (tiếp) – Prompt Engineering

### Lần đầu thiết kế prompt cho email classifier

Prompt đầu tiên của tôi rất đơn giản và... cho kết quả tệ:
```
Phân loại email sau thành: Urgent/Normal/Low
```

Claude trả về plain text như "Email này có vẻ là Normal priority vì..." — hoàn toàn không parse được!

**Iteration 2**: Thêm yêu cầu JSON output và ví dụ format:
```
Trả về JSON: {"priority": "...", "category": "..."}
KHÔNG thêm markdown hay giải thích.
```

Kết quả tốt hơn nhiều. Nhưng vẫn thỉnh thoảng Claude thêm ```json``` block ở đầu.

**Iteration 3**: Thêm regex fallback để extract JSON từ bất kỳ format nào:
```python
m = re.search(r"\{.*\}", raw, re.DOTALL)
data = json.loads(m.group()) if m else {}
```

**Bài học**: Khi dùng LLM để parse structured data, luôn:
1. Nói rõ format trong prompt ("KHÔNG thêm markdown")
2. Có fallback mechanism
3. Validate output sau khi parse

---

## Ngày 2 – Feature Implementation

### Khó khăn với Calendar Creator – Natural Language Date Parsing

Email thường nói "thứ Hai tuần này", "ngày mai lúc 2h chiều", "cuối tháng" thay vì ISO date. Tôi phải cho Claude context ngày hiện tại và thứ trong tuần để nó tự suy luận.

Trick quan trọng:
```python
today = datetime.date.today()
user_prompt = (
    f"Ngày hiện tại: {today.isoformat()} ({today.strftime('%A')})\n"
    ...
)
```

Claude rất giỏi suy luận "thứ Hai tuần này" → ngày cụ thể khi có ngày hôm nay làm reference.

### Streaming vs Non-streaming

Ban đầu tôi dùng non-streaming cho tất cả. Nhưng tính năng tóm tắt thread và gợi ý lịch ngày có thể tạo response rất dài (800-1500 tokens) → timeout issues.

Giải pháp: dùng streaming cho các feature có output dài:
```python
with client.messages.stream(...) as s:
    return s.get_final_message().content[0].text
```

Lưu ý: `get_final_message()` vẫn đợi toàn bộ response hoàn tất. Nếu muốn hiển thị từng token real-time thì cần thêm logic phức tạp hơn (future improvement).

---

## Ngày 2 (tiếp) – Threading

### Bug khó debug: GUI bị đơ

Khi tôi gọi AI API trực tiếp trong button handler:
```python
def on_click():
    result = ai.chat(...)  # BLOCKING! GUI đơ 10-30 giây
    output.set(result)
```

GUI đơ hoàn toàn, không click được, không scroll được.

Giải pháp: Background threads + `self.after()` để update UI:
```python
def on_click():
    threading.Thread(target=worker, daemon=True).start()

def worker():
    result = ai.chat(...)
    self.after(0, lambda: output.set(result))  # safe UI update
```

**Bài học quan trọng**: Tkinter (và hầu hết GUI frameworks) không thread-safe. Mọi UI update phải chạy trên main thread. `after(0, callback)` là cách marshal về main thread.

---

## Ngày 3 – Polish & Documentation

### Thiết kế lại color scheme

Dark theme ban đầu dùng màu xám đơn điệu. Sau khi thêm màu Catppuccin (cyan cho unread, orange cho urgent, green cho success, blue cho accent), giao diện sinh động hơn rất nhiều và người dùng có thể nhận biết trạng thái nhanh hơn.

### Xử lý email có body rỗng

Phát hiện một bug nhỏ: một số email (tự động generated) có body rỗng hoặc chỉ có HTML không đọc được. Giải pháp đơn giản: truncate và default rỗng đã handle sẵn trong `EmailMessage` dataclass với `body = item.Body or ""`.

---

## Tổng kết – Bài học rút ra

### Kỹ thuật

1. **COM Automation với pywin32**: Deserialize COM objects ngay lập tức, đừng pass COM objects qua threads
2. **Prompt Engineering**: JSON output cần "KHÔNG thêm markdown", luôn có fallback parser
3. **Threading với tkinter**: Dùng `after(0, fn)` cho mọi UI update từ background thread
4. **Streaming**: Dùng streaming cho responses dài, `get_final_message()` để lấy kết quả cuối

### Thiết kế

5. **Data Classes**: Tách biệt domain objects (`EmailMessage`) khỏi service objects (`OutlookClient`)
6. **Layered Architecture**: Mỗi layer có một trách nhiệm rõ ràng, dễ test và thay thế
7. **Error Messages**: Lỗi kỹ thuật (`COM Error 0x80040119`) không có nghĩa với người dùng → wrap thành message thân thiện

### Về LLM nói chung

8. **LLM không phải deterministic**: Cùng một prompt có thể cho output format khác nhau → luôn có fallback
9. **Context matters**: Cho Claude biết ngày hôm nay, múi giờ, ngôn ngữ mong muốn → kết quả tốt hơn nhiều
10. **Truncation is necessary**: Email dài có thể vượt quá token budget → cắt body xuống 3000-5000 chars, ưu tiên phần đầu

---

## Những gì sẽ làm nếu có thêm thời gian

- [ ] **Real-time streaming display** – hiển thị từng token thay vì đợi full response
- [ ] **Smart reply** – gợi ý câu trả lời cho email cần reply
- [ ] **Email triage dashboard** – phân tích toàn bộ inbox theo category/priority với chart
- [ ] **Multi-account support** – hỗ trợ nhiều tài khoản Outlook
- [ ] **Offline mode** – cache kết quả phân loại để không gọi API lại
- [ ] **Keyboard shortcuts** – Ctrl+1 classify, Ctrl+2 task, v.v.
- [ ] **Export results** – xuất tóm tắt và kế hoạch ngày ra file Word/PDF

---

## Cảm nhận cá nhân

Dự án này là lần đầu tiên tôi kết hợp **Windows COM automation** với **LLM API** và GUI trong cùng một ứng dụng. Khó khăn nhất không phải là code AI mà là làm cho 3 hệ thống (Outlook COM, Claude API, tkinter) hoạt động hài hòa với nhau, đặc biệt là vấn đề threading.

Claude Opus 4.6 thực sự ấn tượng trong việc hiểu context email tiếng Việt và trích xuất thông tin có cấu trúc. Prompt engineering không khó bằng tôi nghĩ — chỉ cần rõ ràng về format output và có fallback là đủ.

Tổng thể đây là một bài học thực tế quý giá về việc xây dựng **AI-augmented desktop application** trong môi trường doanh nghiệp thực.
