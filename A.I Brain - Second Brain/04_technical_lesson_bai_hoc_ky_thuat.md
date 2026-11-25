# 04 — Bài học kỹ thuật (đã mất thời gian — đừng lặp lại)

| # | Vấn đề | Nguyên nhân | Cách xử lý |
|---|---|---|---|
| 1 | PowerShell báo `Missing closing '}'` với script trông đúng | `.ps1` tiếng Việt lưu UTF-8 **không BOM** → Windows PowerShell 5.1 đọc theo ANSI | Luôn lưu `.ps1` UTF-8 **có BOM** (python: `encoding='utf-8-sig'`) |
| 2 | Gradle: `Conflicting configuration … ndk abiFilters … splits abi` | Không được đặt `ndk.abiFilters` khi bật `splits.abi` | Chỉ dùng `splits { abi { reset(); include … } }` |
| 3 | Gradle `--offline` thiếu jar AGP 8.2.1 | Cache `D:\yh\gradle-home` chưa đủ | Script tự chạy lại online |
| 4 | Tính năng "có trong mã" nhưng không chạy | `Bridge.scanFiles()` (v0.1.6) không được JS gọi tới — Feature Check #28 chưa từng hoạt động | Khi nhận một mục "đã làm", **grep cả phía gọi**; thêm selftest cho mục đó |
| 5 | Gọi JS↔Java đồng bộ làm đứng giao diện | `@JavascriptInterface String f()` chạy trên luồng JS | Việc nặng: `void f(String id)` + `new Thread` + `a.callback(id, json)` |
| 6 | Logo mới bị thay lại logo cũ sau build | Script v0.1.7 chép `mipmap-*` từ `app\` (Flutter cũ) đè lên | Không chép mipmap từ `app\`; logo nằm trong `vN\...\res\mipmap-*` |
| 7 | Word có ảnh treo vĩnh viễn (v0.1.6→0.1.7) | `crash\dump.ini` của LibreOffice làm bật hộp thoại Crash Report ẩn | Dọn báo cáo lỗi cũ trước khi khởi động engine |
| 8 | Basic tool mất tham số cuối | `TextInputStream.isEOF()` của LO báo hết sớm | Đọc toàn bộ tệp tham số một lần |
| 9 | Nghi lỗi `_3dSheets` gắn vĩnh viễn vào sheet | Kiểm tra: `toRef()` xóa ngay sau khi dùng — **không lỗi** | Luôn viết test chứng minh trước khi "sửa" |
| 10 | Spread `{...obj}` trong mã giao diện | WebView Android 6 (Chrome 44) — Babel có thể bỏ sót tùy cấu hình | Dùng `Object.assign` |
| 11 | Đường dẫn dự án cũ viết cứng (`D:\A.I Project\Office Open Source Project`) | Dự án đã chuyển vào `D:\A.I Project\A.I Project - Office Open Source Project\…` | Tính `$PROJ` từ `$PSScriptRoot`, không viết cứng |
| 12 | Máy ảo của Claude không ra GitHub (HTTP 403 proxy) | Chính sách mạng | git/GitHub/tải logo làm trên Windows (runner hoặc Antigravity) |
| 13 | Ghi clipboard bị chặn khi File Explorer ở trước | Ứng dụng tier "click" | Đưa ứng dụng tier "full" (Antigravity) lên trước rồi mới ghi clipboard; kiểm tra lại nội dung trước khi gửi |
| 14 | Job runner chạy tệp `.ps1` đang ghi dở → "is not recognized…" | Runner quét `jobs\` mỗi 2 s, có thể lấy tệp khi `printf … ; cat >> …` chưa xong | Ghi vào `jobs\<tên>.tmp` rồi `mv` thành `.ps1` (đổi tên là nguyên tử) |
| 15 | Video kiểm thử lưu vào thư mục lạ `all_proxy\` | Biến PowerShell **không phân biệt hoa thường**: vòng `foreach ($v in …)` ghi đè `$V` (thư mục video) | Đặt tên biến vòng lặp khác hẳn (`$pv`); không dùng tên chỉ khác hoa/thường |
| 16 | Gradle "Timeout waiting to lock journal cache" | Gradle daemon **mồ côi** của Android Studio (chạy từ sáng, AS đã đóng) giữ khóa `D:\yh\gradle-home\caches\journal-1` | Kiểm tra `Win32_Process` (java.exe, cha còn sống không); dừng daemon mồ côi rồi build lại |
| 17 | Dừng tiến trình kiểm thử có thể giết nhầm việc của tác nhân khác | `Stop-Process scrcpy` dừng mọi scrcpy (cả bản quay Note 9 của Antigravity) | Lưu PID của chính mình (`Start-Process -PassThru`) và chỉ dừng đúng PID đó |
| 18 | Emulator TV `YomiHub_TV_API34` là **x86 32-bit** — v0.1.8 không còn APK x86 | TV thật dùng chip ARM | Thử "trình chiếu lên TV" bằng emulator x86_64 + `settings put global overlay_display_devices 1920x1080/240` (màn hình phụ), chụp bằng `screencap -d <displayId>` |
| 19 | Quay video emulator (`adb emu screenrecord`) không ra tệp | (a) tối đa **180 s** mỗi lần (`--time-limit 1800` bị từ chối); (b) lệnh console emulator **cắt đường dẫn tại dấu cách** | Quay từng đoạn 175 s vào `%TEMP%\oos_seg_*` (không dấu cách) trong `Start-Job`, xong ghép bằng `ffmpeg -f concat`; thiết bị thật dùng scrcpy `--record x.mkv` |
| 20 | Runner ghi `.done` ngay dù job còn chạy (job vòng lặp lưu tiến độ) → tưởng job đã thoát, chạy thêm bản thứ hai → lưu trùng | Không tin `.done` với job chạy dài; kiểm tra tiến trình thật | `Get-CimInstance Win32_Process` lọc `CommandLine -match '<tên job>'`; dừng đúng PID |
