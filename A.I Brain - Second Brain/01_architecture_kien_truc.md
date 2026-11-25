# 01 — Kiến trúc (Architecture)

```
             ┌──────────────── vNNN\src\src  (giao diện dùng chung, ES2017 modules) ────────────────┐
             │ core/  shell, ui, host (cầu nối nền tảng), store, update, devscan, cloud, dex, selftest │
             │ word/  excel/ (formula, functions*, vba/)  ppt/  pdf/  scan/ (CamScanner)  math/  ai/   │
             └──────────────┬───────────────────────────────────────────────────────────────────────┘
        tools/build.js (esbuild → Babel ES5 → postcss) → dist (universal, ES5) / dist-modern (ES2017) / dist-legacy (ES5, asm.js OCR)
                            │
   ┌────────────────────────┴───────────────┐                    ┌────────────────────────────────────────┐
   │ Android  platform\android (Java)        │                    │ Windows platform\windows (C#, .NET 4)    │
   │ MainActivity: WebView tải www/          │                    │ Launcher + máy chủ http 127.0.0.1 (/api)│
   │ Bridge = window.OOSNative (JS↔Java)     │                    │ WebView2/Edge app window                │
   │ NativeDocActivity: LibreOfficeKit (.so) │                    │ Native: LibreOffice 24.8 (Win7–8.1)     │
   │ flavors: legacy/android8/modern × ABI   │                    │         LibreOffice 26.x (Win10–11)     │
   └─────────────────────────────────────────┘                    └────────────────────────────────────────┘
```

## Hai "engine" mở tài liệu
- **Web** (mặc định cho chỉnh sửa trên Android): trình soạn thảo JS tự viết (Word/Excel/PowerPoint), đọc/ghi OOXML, ODF, RTF, .doc/.xls/.ppt cũ.
- **Native** (LibreOfficeKit): hiển thị chuẩn Microsoft Office. Trên Android chỉ xem; trên Windows sửa được.
- Bố cục Ribbon/menu Tệp/Tùy chọn dùng chung: `core/ui_layout.json`, `core/ui_settings.json` (v0.1.7).

## Cầu nối nền tảng (`core/host.js`)
- `Host.kind`: `android` | `windows` | `web`.
- Android: `window.OOSNative.<method>(…, id)` → Java chạy nền → `window.__oosCb(id, json)`.
- Windows: `fetch('/api/<tên>')` tới máy chủ cục bộ của launcher.
- Khi thêm tính năng cần quyền máy: viết hàm trong `Host`, có **đường dự phòng** khi host chưa hỗ trợ (vd `scanOffice()` → tự duyệt `/api/list_dir`).

## Các mô-đun hay sửa
| Mô-đun | Việc |
|---|---|
| `core/update.js` | Cập nhật phần mềm: GitHub Releases (chính) → Drive/MediaFire (phụ); `myEdition()` + `pickFile()` chọn bộ cài đúng edition |
| `core/devscan.js` (v0.1.8) | Quét tệp Word/Excel/PowerPoint trên máy, tab "Trên thiết bị" |
| `ai/ai.js` | Trợ lý A.I: ngoại tuyến, Web (không/có đăng nhập), khóa API (lưu trên máy), Free Claude Code (proxy 127.0.0.1:8082) |
| `excel/formula.js` | Parser + Calc; tham chiếu sheet, 3D, `[Book.xlsx]Sheet!A1` (qua `window.OOS.docs` / `OOS_EXTERNAL_WORKBOOKS`) |
| `core/selftest.js` | `#selftest` — kiểm thử tự động trong WebView thật, in `OOSTEST PASS/FAIL` |
