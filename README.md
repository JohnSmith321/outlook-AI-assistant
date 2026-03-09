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
