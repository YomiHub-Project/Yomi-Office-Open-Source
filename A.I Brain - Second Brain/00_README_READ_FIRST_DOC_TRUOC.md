# 00 — ĐỌC TRƯỚC KHI LÀM BẤT CỨ VIỆC GÌ (READ FIRST)

> **Quy định bắt buộc của chủ dự án (25/11/2026):** mọi phiên làm việc sau — người hoặc A.I (Claude, Google Antigravity, Gemini, Copilot, Codex…) — **phải đọc toàn bộ thư mục `A.I Brain - Second Brain` và `TECHNICAL_UPDATE_NOTES.md` trước khi sửa mã nguồn**, và **phải ghi thêm** kinh nghiệm mới vào đây khi kết thúc phiên.

## Dự án là gì
**Yomi-Office Open Source** (trước v0.1.8 tên là "Office Open Source") — bộ ứng dụng văn phòng mã nguồn mở cho **Android 6.0+** và **Windows 7+**: Word, Excel, PowerPoint, PDF (đọc, OCR, ghép, chuyển đổi), CamScanner, MathType, Trợ lý A.I, đồng bộ Cloud, DeX.
- Repo phát hành: https://github.com/YomiHub-Project/Yomi-Office-Open-Source (Releases = nguồn cập nhật chính của ứng dụng).
- Nguồn cập nhật phụ: Google Drive `1kkqNrHFGQXLbf7PPYYc6HWgD1mFOdIyr`, MediaFire `seqly6bv55yr8`.
- Yêu cầu nằm trong 2 Google Sheet của chủ dự án: **Project GitHub** (~600 repo tham chiếu, 24 mục I–XXIV) và **Project - Office Open Source Project** (sheet *Office Feature Check* 33 mục, *Office Apps Creat & Update*, *App Feature Set*, *Template Office*, *Update App Link*).

## Đọc theo thứ tự
1. `01_architecture` — các khối mã và cách chúng nối nhau.
2. `10_project_directory_structure` — đường dẫn thật trên máy (đã đổi chỗ, đừng dùng đường dẫn cũ).
3. `02_procedure` + `03_build_release` — làm một bản cập nhật từ đầu đến cuối.
4. `04_technical_lesson` — **các lỗi đã từng mất nhiều thời gian**. Đọc kỹ.
5. `05_testing`, `06_tools_used`, `07_performance_rules`.
6. `08_ai_session_tips`, `09_agent_orchestration` — làm việc với nhiều A.I cùng lúc.
7. `11_project_handover` — bàn giao; `12_session_log` — nhật ký các phiên.

## 5 điều quan trọng nhất
1. Mã nguồn sống ở **`vNNN\src\src`** (giao diện dùng chung) + **`vNNN\src\platform\{android,windows}`**. Thư mục `app\` (Flutter) là **di sản v0.0.x, không dùng nữa**.
2. Mỗi bản mới: chép `v(N-1)` → `vN`, sửa trong `vN`, đầu ra vào `..\output\vX.Y.Z\{Android,Windows}`.
3. Script `.ps1` có tiếng Việt **phải lưu UTF-8 có BOM**.
4. Không bao giờ báo "đã xong" khi chưa chạy thử: dùng `#selftest`, thiết bị thật/emulator, ảnh chụp + video.
5. Ghi mọi thay đổi vào `TECHNICAL_UPDATE_NOTES.md` (mục, lỗi, dòng lệnh, công cụ, kết quả kiểm thử, ngày).
