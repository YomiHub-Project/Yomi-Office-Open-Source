# 02 — Quy trình một bản cập nhật (Procedure)

1. **Đọc** A.I Brain + `TECHNICAL_UPDATE_NOTES.md` + 2 Google Sheet (xuất ra .csv/.txt nếu cần).
2. **Tóm tắt yêu cầu bằng tiếng Việt và hỏi xác nhận** những điểm mâu thuẫn trước khi làm (vd v0.1.8: repo đích, cách chia APK, Win7).
3. **Tạo `vN`**: chép `v(N-1)\src\src` (nhỏ, ~2.5 MB) và robocopy phần còn lại (`platform` ~10 GB) trên Windows: `robocopy v(N-1)\src vN\src /E /XD v(N-1)\src\src /MT:16`.
4. **Đổi phiên bản** ở: `core/util.js` (`VERSION`, `APP_NAME`), `android/app/build.gradle` (versionName, versionCode), script build Windows, appx manifest.
5. **Làm từng mục** của *Office Feature Check* / *Office Apps Creat & Update*; mỗi mục ghi 1 dòng trong `vN\coord\TIEN_DO_<tên>.md`.
6. **Build** qua runner (`RUN_vNNN_BUILD_RUNNER.bat` → đặt job .ps1 vào `vN\jobs`, log ở `vN\joblogs`).
7. **Kiểm thử**: `#selftest`, thiết bị thật + emulator Android 6/10/TV, các bản Windows, toàn bộ `Test Office` so với MS Office gốc; ảnh + video.
8. **Ghi chép**: `TECHNICAL_UPDATE_NOTES.md`, `RELEASE_NOTES_vX.Y.Z.md`, A.I Brain, `SHA256`.
9. **Phát hành**: đẩy mã nguồn + Release (gắn toàn bộ tệp trong `output\vX.Y.Z`) + README (nút Download → `/releases/latest`), kiểm tra lại danh sách tệp trên GitHub.

## Lưu tiến độ tự động 5 phút/lần (quy định của chủ dự án, 25/11/2026 16:20)
- Đầu mỗi phiên: thả `vN\checkpoint_loop.ps1` vào `vN\jobs\000_checkpoint_loop.ps1` (runner phải đang chạy). Mỗi 5 phút nó ghi `vN\coord\TIEN_DO_HIEN_TAI.md` (trạng thái job, bộ cài trong output, ảnh/video, ghi chú mới nhất của từng tác nhân), thêm 1 dòng vào `vN\coord\CHECKPOINT_LOG.md` và sao lưu `coord\*.md`, TECHNICAL_UPDATE_NOTES, RELEASE_NOTES, A.I Brain, `src\src` (zip) vào `vN\checkpoints\<yyyyMMdd_HHmm>\` (giữ 36 bản = 3 giờ).
- Dừng: tạo tệp `vN\coord\STOP_CHECKPOINT`.
- Phiên bị gián đoạn: đọc `TIEN_DO_HIEN_TAI.md` + `CHECKPOINT_LOG.md` để biết đang ở đâu; khôi phục mã nguồn từ `checkpoints\…\src_src.zip` nếu cần.
