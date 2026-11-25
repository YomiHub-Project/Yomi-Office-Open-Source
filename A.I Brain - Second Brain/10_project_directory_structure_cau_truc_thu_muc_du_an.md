# 10 — Cấu trúc thư mục (đường dẫn thật, 25/11/2026)

```
D:\A.I Project\A.I Project - Office Open Source Project\
├─ output\                       ← bộ cài các phiên bản: v0.0.1 … v0.1.7, v0.1.8\{Android,Windows}, version.json
└─ Office Open Source Project\    ← GỐC DỰ ÁN ($PROJ)
   ├─ A.I Brain - Second Brain\  ← thư mục này
   ├─ TECHNICAL_UPDATE_NOTES.md   ← nhật ký kỹ thuật bắt buộc
   ├─ RELEASE_NOTES_vX.Y.Z.md
   ├─ RUN_vNNN_BUILD_RUNNER.bat   ← chạy runner job của phiên bản
   ├─ Test Office\               ← 24 tệp Word/Excel/PowerPoint để so sánh với MS Office
   ├─ logo\                      ← logo (Yomi-Office Open Source.png)
   ├─ v005\deps\node_modules     ← phụ thuộc build giao diện
   ├─ v016\lo\                   ← LibreOffice (Android jniLibs/assets, Windows)
   ├─ v017\, v018\               ← mỗi phiên bản: src\src (giao diện), src\platform\{android,windows}, build\, jobs\, joblogs\, coord\, shots\, video\, test\
   ├─ app\                       ← Flutter cũ (v0.0.x) — KHÔNG dùng
   └─ desktop_client\            ← launcher C# cũ + assets
D:\yh\tools\                     ← jdk, android-sdk, git, wix314, scrcpy, ffmpeg, flutter…
D:\A.I Project\GitHub\          ← một số repo tham chiếu đã tải
```
