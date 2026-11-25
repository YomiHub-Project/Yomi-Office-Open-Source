# 07 — Quy tắc thực hiện

1. **Trung thực**: mục nào chưa làm/không làm được thì ghi rõ lý do (vd WPS/MathType đóng nguồn; AirDrop là giao thức Apple; sửa chung thời gian thực cần OAuth Client ID; `.appxbundle` chỉ Win10 1809+; Windows 7 không có ARM).
2. **Không đè việc của người khác**: chia thư mục rõ ràng (xem `09_agent_orchestration`).
3. **Tối ưu dung lượng**: APK chia theo ABI (mỗi APK chỉ 1 bộ `.so` LibreOffice thay vì 4); bỏ x86 32-bit.
4. **Tương thích Android 6**: giao diện bản legacy/universal là ES5; tránh cú pháp ES2018+ trong `src` nếu không chắc Babel xử lý.
5. **Không làm chậm khi khởi động**: việc nền (quét tệp, kiểm tra cập nhật) chạy trễ vài giây, có giới hạn thời gian/số lượng.
6. **Mỗi thay đổi có kiểm thử**; mỗi lỗi sửa có selftest hoặc ảnh chứng minh.
7. **Giấy phép**: dự án AGPL (vd Free Claude Code) chỉ **kết nối qua HTTP**, không nhúng mã.
8. **Bảo mật**: khóa API chỉ lưu trên máy (Windows DPAPI); không đưa khóa/token vào repo hay log.
