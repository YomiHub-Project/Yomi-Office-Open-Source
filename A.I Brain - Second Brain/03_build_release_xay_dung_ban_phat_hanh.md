# 03 — Build & phát hành

## Giao diện web
```
set NODE_PATH=<gốc dự án>\v005\deps\node_modules
cd vN\src
node tools/build.js --profile=universal --out=dist        (ES5, ~90 s lần đầu, Babel có cache)
node tools/build.js --profile=modern    --out=dist-modern (ES2017, vài giây)
node tools/build.js --profile=legacy    --out=dist-legacy (ES5 + asm.js OCR)
```
- Chạy trên **Windows** (nhanh). Trong máy ảo Linux của Claude, đọc `node_modules` qua ổ dùng chung quá chậm (>3 phút chưa xong) — chỉ dùng máy ảo để kiểm thử đơn vị nhỏ bằng `esbuild --bundle` (binary Linux trong `src\esbuild-linux-x64-0.24.0.tgz`).
- Cảnh báo `[icons] missing: options, chart-` có từ trước, không chặn build.

## Android (v0.1.8)
- `vN\src\platform\build_vNNN_android.ps1`: chép dist vào `app\src\{legacy,universal,modern}\assets\www` (flavor android8 dùng `universal`), Gradle 8.3 (`D:\yh\gradle-home`), JDK `D:\yh\tools\jdk`, SDK `D:\yh\tools\android-sdk`.
- **9 APK** = 3 flavor (Android6-7 API23+, Android8-9 API26+, Android10plus API29+) × 3 ABI (`splits.abi`: arm64-v8a, armeabi-v7a, x86_64). versionCode = 18000 + 10·flavor + ABI.
- Thư viện LibreOffice Android: `<gốc>\v016\lo\android\{jniLibs,assets}` (sourceSets trong build.gradle, đường dẫn tương đối `../../../../v016`).
- Ký bằng `~/.android/debug.keystore` (cùng chữ ký các bản trước → cài đè được).
- Tên tệp: `Yomi-Office-Open-Source-<ver>-<Android6-7|Android8-9|Android10plus>-<abi>.apk`.

## Windows (Antigravity phụ trách ở v0.1.8)
- `build_vNNN_windows.ps1`: 5 cấu hình Win7-8.1 x64/x86 (LibreOffice 24.8) + Win10-11 x64/x86/ARM64; mỗi cấu hình `-Setup.exe`, `.msi` (WiX 3.14 `D:\yh\tools\wix314`), `-portable.zip`; thêm `.appxbundle` (Win10 1809+).
- Tên tệp: `Yomi-Office-Open-Source-<ver>-<edition>-Setup.exe | <edition>.msi | <edition>-portable.zip`, `Yomi-Office-Open-Source-<ver>.appxbundle`.

## Quy ước tên = hợp đồng với tính năng Cập nhật
`core/update.js` (`FILE_RE`, `pickFile`) chọn bộ cài theo tên. **Đổi tên tệp phát hành thì phải sửa `update.js` và chạy selftest "v018 Cập nhật phần mềm"**.
