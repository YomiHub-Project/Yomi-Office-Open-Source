# 08 — Mẹo làm việc với A.I

- Mở đầu phiên: đọc Brain → hỏi xác nhận ngắn các điểm mâu thuẫn (dùng câu hỏi trắc nghiệm) → lập danh sách công việc.
- Máy ảo Linux của Claude (`device_bash`) thấy các thư mục đã kết nối ở `$HOME/mnt/<tên>`; **không** chạy được chương trình Windows, **không** ra GitHub. Việc Windows: thả job `.ps1` vào `vN\jobs` (runner phải đang chạy: `RUN_vNNN_BUILD_RUNNER.bat`).
- Lệnh `device_bash` tối đa 180 s: build dài phải chạy qua runner rồi đọc log.
- Không in danh sách gói cài trên điện thoại hay dữ liệu cá nhân ra log (dùng bộ lọc chặt khi `pm list packages`).
- Computer-use: Antigravity/Word/Excel/PowerPoint/GitHub Desktop ở tier "full"; Android Studio/VS Code/PowerShell/File Explorer chỉ "click" (không gõ phím) → lệnh shell đi qua runner.
- Tin nhắn dài gửi Antigravity: ghi clipboard khi Antigravity đang ở trước, dán, **zoom kiểm tra** trước khi Enter; tin bị "Queued" khi tác nhân con đang chạy → bấm "Send Now".
