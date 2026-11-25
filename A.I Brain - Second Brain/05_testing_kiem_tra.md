# 05 — Kiểm thử

## Công cụ có sẵn trên máy
- `#selftest`: mở ứng dụng với hash `selftest` (Android: `am start -n com.officeopensource.office_open_source/.MainActivity --es hash selftest`), đọc logcat `OOSTEST PASS|FAIL … / OOSTEST DONE | x passed, y failed`.
- Route kiểm thử: `#tool=ai`, `#tool=scan`, `#tool=pdfedit`, `#tool=devscan` (v0.1.8), `#dex=tvhome`.
- Script: `vN\src\platform\test_android_vNNN.ps1 -Serial <adb serial> | -Avd <tên> [-Dex]` — tự chọn APK theo API + CPU, quay video (scrcpy `.mkv` → ffmpeg `.mp4`), chụp ảnh, mở từng tệp `Test Office`.
- Thiết bị thật: Samsung Galaxy Note 9 SM-N960N (Android 10, arm64, có DeX) serial `<adb-serial>`.
- Emulator (`%USERPROFILE%\.android\avd`): `OOS_Android6_API23`, `YH_Android10_API29`, `YomiHub_TV_API34` (TV — thử trình chiếu DeX), `YomiHub056_Phone_API34`.
- Công cụ: adb (`%LOCALAPPDATA%\Android\Sdk\platform-tools`), emulator, scrcpy `D:\yh\tools\scrcpy`, ffmpeg `D:\yh\tools\ffmpeg-9.0.2-essentials_build\bin`.
- Tệp mẫu: `<gốc>\Test Office` (24 tệp: 1 .doc, 4 .docx, 2 .xlsx, 4 .xlsm, 13 .pptx; có 1 tệp khóa `~$…` phải bỏ qua).

## Kiểm thử đơn vị nhanh trong máy ảo
`esbuild --bundle --platform=node` một entry nhỏ import trực tiếp `excel/model.js`, `excel/formula.js`, `core/update.js`… (đặt `NODE_PATH` tới `v005\deps\node_modules`); bỏ `import` UI bằng stub khi cần.

## Tiêu chí
- Mỗi mục sheet có ít nhất 1 bằng chứng: PASS selftest, ảnh, hoặc log.
- Không ghi "đạt" khi chỉ build được; ghi rõ chưa thử ở đâu.
