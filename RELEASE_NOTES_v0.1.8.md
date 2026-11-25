# Yomi-Office Open Source v0.1.8 — Ghi chú phát hành

*Ngày phát hành: 25/11/2026 · Bản trước: v0.1.7 (22/10/2026)*

## 1. Tóm tắt
- **Đổi tên** thành **Yomi-Office Open Source**, logo mới (Yomi-Logo) trên Android, Windows, giao diện, trang chủ.
- **Bộ cài mới theo thiết bị**: Android chia **9 APK** theo phiên bản Android × CPU; Windows có bộ cài riêng cho **Windows 7/8.1** và **Windows 10/11** ở dạng **Setup .exe, .msi, bản portable** và gói **.appxbundle**.
- **APK nhỏ hơn ~73%**: 77–94 MB mỗi APK (v0.1.7: 315 MB) vì mỗi APK chỉ chứa thư viện LibreOffice của đúng 1 loại CPU.
- **Phát hành trên GitHub Releases**; ứng dụng tự kiểm tra cập nhật ở GitHub (nguồn chính), Google Drive, MediaFire.

## 2. Tính năng mới so với v0.1.7
| # | Tính năng | Chi tiết |
|---|---|---|
| 1 | **Tab "Trên thiết bị"** | Khi mở ứng dụng, tự quét máy và liệt kê mọi tệp Word / Excel / PowerPoint (Android: chạy nền, không làm đứng giao diện; Windows: quét thư mục người dùng). Nút *Quét lại*, nút *Cấp quyền truy cập tệp* trên Android 11+. Tùy chọn → Chung: bật/tắt quét khi mở ứng dụng |
| 2 | **Free Claude Code** trong Trợ lý A.I | Kết nối proxy Free Claude Code (mặc định `http://127.0.0.1:8082`) — dùng miễn phí các mô hình NVIDIA NIM, OpenRouter, Groq, Ollama… Token proxy không bắt buộc; điện thoại nhập địa chỉ IP máy tính chạy proxy |
| 3 | **Cập nhật từ GitHub Releases** | Nguồn kiểm tra cập nhật chính; tự chọn đúng bộ cài (APK theo Android + CPU; Setup/MSI/portable/appx theo Windows) |
| 4 | **Bộ cài .msi** cho Windows | Cài đặt doanh nghiệp/GPO; cập nhật tự động dùng lại .msi |
| 5 | **APK theo CPU** | arm64-v8a, armeabi-v7a, x86_64 cho Android 6–7, 8–9, 10+ |
| 6 | **Tự kiểm tra (#selftest) mở rộng** | 49 mục (thêm 5 mục v0.1.8: tên/phiên bản, hàm liên kết nhiều sheet/3D/nhiều bảng tính, cập nhật GitHub, Free Claude Code, quét tệp) |

## 3. Lỗi đã sửa
1. **Quét tệp Office khi mở ứng dụng không hoạt động** (có từ v0.1.6): mã Android có nhưng giao diện không gọi tới; nay đã nối, chạy nền và không tự mở màn hình xin quyền mỗi lần mở app.
2. **Cập nhật phần mềm bỏ sót nguồn chính GitHub** và không nhận tệp `.msi`.
3. **Logo cũ** còn ở thanh bên, màn hình chờ, favicon, About; **chuỗi phiên bản cũ ghi cứng** ("v0.1.7", "v0.1.6", "v0.1.3") ở thẻ DeX, VBA Immediate Window, thông tin tài liệu.
4. Script build Android cũ chép đè biểu tượng cũ lên biểu tượng mới.
5. Hộp **Tùy chọn** tràn ngang trên màn hình dọc điện thoại; tab Cập nhật ghi sai nguồn (thiếu GitHub) — thêm ô "Kho GitHub".
6. `licenses.txt` (Giới thiệu → Giấy phép mã nguồn mở) ghi sai tên/phiên bản.
7. Script build/kiểm thử: lỗi mã hóa UTF-8 không BOM trên PowerShell 5.1, xung đột `abiFilters`/`splits` của Gradle, đường dẫn dự án viết cứng.

## 4. Kiểm thử
| Thiết bị | Kết quả |
|---|---|
| Emulator Android 6.0 (API 23, x86_64) | selftest **49/49**; mở 8/8 tệp thật (.doc, .docx, .xlsm, .pptx) không văng; quét tệp 8/8 |
| Emulator Android 14 (API 34, x86_64) + màn hình phụ TV | selftest **49/49**; mở tệp thật không văng |
| Samsung Galaxy Note 9 (Android 10, arm64) | selftest **49/49**; mở **23/23** tệp Test Office không văng; quét 105 tệp; DeX, màn hình phụ TV, Quick Share, Trợ lý A.I (gồm Free Claude Code), PDF, MathType, CamScanner: đạt |
| Windows 7/8.1, 10/11 (máy x64) | Bản portable và .msi của Win7-8.1 x86/x64, Win10-11 x86/x64 cài, chạy, gỡ đạt; ARM64 kiểm tra gói đúng kiến trúc; chuyển DOCX→PDF, quét tệp đạt; mở **22/22** tệp Test Office cạnh Microsoft Office 16 (ảnh đối chiếu + video) |

## 5. Hạn chế đã biết
- Sửa chung thời gian thực với Google Drive / OneDrive cần OAuth Client ID riêng.
- WPS Office và MathType là phần mềm đóng nguồn: chỉ có tính năng tương đương bằng mã nguồn mở.
- AirDrop là giao thức của Apple: dùng Quick Share / chia sẻ qua Wi-Fi.
- `.appxbundle` chỉ cài trên Windows 10 1809+ / 11. Windows 7 không có bản ARM.
- Bản x86 32-bit cho Android đã bỏ (không còn thiết bị thực tế); Android TV x86 dùng bản x86_64 nếu TV hỗ trợ.
- Android: engine Native chỉ để xem; chỉnh sửa bằng trình soạn Web.
- Bảng tính rất lớn (vd 15.8 MB .xlsm, 6.8 MB .xlsx) mở bằng engine Native trên điện thoại có thể mất hơn 35 giây (kiểm thử trên Galaxy Note 9).
- Hàm Excel mới từ các website Excel Functions, đối chiếu ảnh Office Screen: chưa làm ở bản này.
- Một số bảng tính hiển thị chưa giống MS Office (vd thiếu chữ ở hàng tiêu đề có ô gộp/bảng, form control đè lên ô) — đã ghi nhận khi so sánh 22 tệp với Microsoft Office 16, sẽ sửa ở v0.1.9.
- Gói `.appx`/`.appxbundle` ký bằng chứng chỉ tự tạo: cần cài `Yomi-Office-Open-Source-0.1.8-DevCert.cer` (Trusted People) trước khi cài.

## 6. Bộ cài
Xem trang Releases: https://github.com/YomiHub-Project/Yomi-Office-Open-Source/releases/tag/v0.1.8 (SHA-256 trong `SHA256_v0.1.8.md`).
