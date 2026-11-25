# Office Open Source Project — Ghi chú kỹ thuật cập nhật

**BẮT BUỘC ĐỌC TRƯỚC MỖI LẦN CẬP NHẬT.**

Chủ dự án yêu cầu tạo và duy trì tài liệu này ngày 25/10/2026. Quy định được ghi thêm trong `AGENTS.md` ở gốc dự án. Mỗi người hoặc tác nhân tự động cập nhật ứng dụng phải đọc toàn bộ tài liệu, kiểm tra trạng thái và bổ sung kết quả sau khi làm việc.

---

## v0.1.8 — 25/11/2026 (Claude + Google Antigravity `/swarm`)

**Người thực hiện:** Claude (giao diện dùng chung, Android, tài liệu, phát hành) và Google Antigravity (Windows, đồng bộ mã nguồn tham chiếu, so sánh với MS Office, kiểm thử Note 9). Phân công: `v018\coord\PHAN_CONG_SWARM.md`.
**Quyết định đã chốt với chủ dự án:** repo `YomiHub-Project/Yomi-Office-Open-Source`; Android 9 APK (Android 6-7 / 8-9 / 10+ × arm64-v8a / armeabi-v7a / x86_64); Windows Win7–8.1 (x64, x86) + Win10–11 (x64, x86, ARM64) × (Setup .exe, .msi, portable .zip) + .appxbundle.

### A. Nhật ký thay đổi (Claude)
| Thời gian | Hạng mục | Lỗi / thay đổi | Tệp & dòng lệnh | Công cụ | Kết quả kiểm thử |
|---|---|---|---|---|---|
| 25/11/2026 14:18 | Tạo `v018`, chép `src/src` từ v017 | — | `cp -r v017/src/src v018/src/src` | device_bash | 85 tệp, 2.5 MB |
| 25/11/2026 14:20 | Đổi tên ứng dụng (Office Apps Creat & Update, Feature Check #2) | Tên cũ "Office Open Source" ở 91 chỗ (hộp thoại, tiêu đề, metadata docx/xlsx/pptx/pdf/odf/rtf, macro mẫu) | `sed 's/Office Open Source/Yomi-Office Open Source/g'` trên 40 tệp; `core/util.js`: `VERSION='0.1.8'`, `APP_NAME='Yomi-Office Open Source'`; `core/unify.js` dự phòng 0.1.8 | sed, grep | 91/91 thay thế, không còn tên cũ hiển thị |
| 25/11/2026 14:24 | Cập nhật phần mềm (Feature Check #21, sheet Update App Link) | **Lỗi:** chỉ quét Drive + MediaFire, bỏ sót nguồn chính GitHub Releases; không nhận tên bộ cài mới, không nhận `.msi` | `core/update.js`: thêm `DEFAULT_SOURCES.github`, `parseGitHubRelease()`, `listGitHub()` (API `/releases/latest`), `srcName()`; `FILE_RE` nhận `Yomi-Office-Open-Source-*` và `.msi`; `myEdition()/pickFile()` nhận edition Android theo CPU (`Android10plus-arm64-v8a`…) và `ed.msi`; `core/bstage.js:283` ghi nguồn | node (kiểm thử đơn vị `t.mjs`) | 7/7 trường hợp chọn đúng bộ cài (2 Android, Setup, MSI, portable, Win7 x86, appxbundle). Lần chạy đầu MSI chọn nhầm .exe → đã sửa, chạy lại đạt |
| 25/11/2026 14:28 | Tích hợp Free Claude Code vào cài đặt A.I (Office Apps Creat & Update) | Chưa có | `ai/ai.js`: nhà cung cấp `freeclaude` (Anthropic Messages, `http://127.0.0.1:8082/v1/messages`, token tùy chọn gửi ở `x-api-key` + `Bearer`, địa chỉ sửa được); ô "Token proxy (không bắt buộc)" trong Cài đặt A.I; dòng trạng thái hiện địa chỉ proxy. Dùng `Object.assign` (không dùng spread) để giữ tương thích WebView Android 6 | node + máy chủ giả lập FCC cổng 8082 (`fcc.mjs`) | 2/2: không token → trả lời; có token `tok123` → header đúng, trả lời đúng |
| 25/11/2026 14:33 | Quét tệp Office trên máy khi mở ứng dụng (Feature Check #28) | **Thiếu tính năng:** trước đây chỉ có danh sách Gần đây/Đã ghim | Mới `core/devscan.js` (`scan`, `walkWindows`, `normalize`, `startupScan`; bỏ qua thư mục hệ thống, tệp khóa `~$`; giới hạn 6000 thư mục / độ sâu 7 / 45 s / 5000 tệp); `core/host.js`: `Host.scanOffice()` (Android `OOSNative.scanOffice`, Windows `/api/scan_office`, rỗng → tự duyệt); `core/shell.js`: tab **Trên thiết bị** + nút Quét lại + tiến trình; Tùy chọn → Chung: "Quét tệp … khi mở ứng dụng"; `main.js`: `startupScan()` sau 4 s | node + host giả lập đọc thư mục thật `Test Office` | 23/23 tệp thật (4 Word, 6 Excel, 13 PowerPoint); bỏ qua đúng `node_modules`, `AppData`, `~$*.docx` |
| 25/11/2026 14:42 | Android: bộ cài theo CPU × phiên bản Android | v0.1.7: 1 APK/flavor chứa thư viện LibreOffice của **4 CPU** (719 MB chưa nén, APK 315 MB) | `app/build.gradle`: versionName 0.1.8; flavor legacy/android8/modern có `BuildConfig.EDITION` (Android6-7/Android8-9/Android10plus), versionCode 18010/18020/18030 (+1/2/3 theo ABI qua `versionCodeOverride`); `splits.abi` arm64-v8a/armeabi-v7a/x86_64; bỏ x86 32-bit | Gradle 8.3 | Lần 1: lỗi "Conflicting configuration … ndk abiFilters … splits" → bỏ `ndk.abiFilters`; offline thiếu jar AGP 8.2.1 → script tự build online |
| 25/11/2026 14:42 | Android: Feature Check #28 (quét tệp) | **Lỗi tồn đọng từ v0.1.6:** `Bridge.scanFiles()` có nhưng JS không gọi; chạy đồng bộ (đứng giao diện); tự mở màn hình xin quyền mỗi lần | `Bridge.java`: `scanOffice(id)` chạy luồng nền + callback, `requestFilesAccess()`, `hasFullAccess()`, `doScan(ask)`; thêm `mtime`, `dir`; `info()` trả `edition` (= EDITION + CPU thật của APK qua `nativeLibraryDir`) | — | chờ kiểm thử trên Note 9 |
| 25/11/2026 14:42 | Android: tên + logo | Tên "Office Open Source", icon cũ; script v0.1.7 chép mipmap từ `app\` Flutter cũ đè lên | `res/values/strings.xml` app_name = Yomi-Office Open Source; `mipmap-*/ic_launcher.png` 48–192 px từ logo Yomi (PIL, lề 6%); `build_v018_android.ps1` bỏ bước chép mipmap cũ | python PIL | ảnh xem trước `v018/shots/android/icon_preview.png` đạt |
| 25/11/2026 14:43 | Script PowerShell | **Lỗi:** `.ps1` tiếng Việt lưu UTF-8 không BOM → PowerShell 5.1 đọc sai, báo "Missing closing '}'" | Lưu `.ps1` bằng UTF-8 **có BOM** (`utf-8-sig`) | python | chạy lại đạt |
| 25/11/2026 14:46 | Excel: liên kết nhiều sheet / nhiều bảng tính (Feature Check #30) | Nghi `_3dSheets` gắn vĩnh viễn vào sheet → kiểm tra: **không lỗi** (bị xóa ngay ở `formula.js:336`) | Thêm selftest v018 | esbuild + node (`xl_test.mjs`) | 9/9 đạt: 3D, Sheet!A1, VLOOKUP sheet khác, `[Book2.xlsx]Data!A1`, `'[Book2.xlsx]Data'!B1` (tiếng Việt) |
| 25/11/2026 14:47 | Tự kiểm tra (#selftest) | Thêm 5 mục v018: tên/phiên bản, liên kết nhiều sheet/bảng tính, Cập nhật GitHub, Free Claude Code, quét tệp; route `#tool=devscan`; log `OOSDEVSCAN` | `core/selftest.js`, `main.js`, `core/devscan.js` | node --check | cú pháp đạt; chạy trên thiết bị ở bước kiểm thử |
| 25/11/2026 15:11 | Build 9 APK (lần 1 thành công) | Gradle bị khóa bởi daemon mồ côi của Android Studio (PID 34428, chạy từ 7:13) — Antigravity dừng daemon | `build_v018_android.ps1` (job 012) | Gradle 8.3, aapt | 9/9 APK: 76.9–94.0 MB (v0.1.7: 315 MB → giảm ~72–75%); aapt: label "Yomi-Office Open Source", minSdk 23/26/29, native-code đúng 1 CPU, versionCode 18011–18033 |
| 25/11/2026 15:21 | Kiểm thử emulator Android 6.0 (API 23, x86_64) | Video scrcpy = 0 MB trên emulator (sẽ dùng `adb emu screenrecord`) | `test_android_v018.ps1 -Avd OOS_Android6_API23` | adb, emulator, scrcpy | **19/20**: cài đặt, phiên bản 0.1.8, quét tệp 6/6, Trợ lý A.I, **selftest 49/49**, mở 6/6 tệp .doc/.docx/.xlsm (Native) không văng |
| 25/11/2026 15:25 | Lỗi giao diện phát hiện khi kiểm thử | Logo cũ ở thanh bên/màn hình chờ/favicon/About; thẻ DeX ghi cứng "(v0.1.7)"; nhiều chuỗi "v0.1.6/0.1.3" | `static/assets/office-logo-icon.png`, `office-logo-full.png` (dựng bằng System.Drawing, Segoe UI), `favicon.png`, `app.ico`; `static/index.html` (title, splash); `shell.js:556`, `dex.js` (4 chỗ), `host.js:361`, `suites.js` (2 chỗ) → `VERSION` | PIL, PowerShell System.Drawing | build lại (job 040) |
| 25/11/2026 15:31 | Kiểm thử emulator Android 14 (API 34) | Script chọn APK arm64 (giả lập) thay vì x86_64 → emulator quá tải "System UI isn't responding" (không phải ứng dụng văng); màn hình phụ overlay không chụp được | Sửa chọn CPU theo thứ tự của thiết bị; TV dùng `adb emu multidisplay add` | adb | **20/22**: selftest 49/49, quét 4/4, mở 4/4 tệp; 2 FAIL thuộc công cụ kiểm thử → chạy lại sau build 040 |

### B. Trạng thái từng mục — sheet *Office Feature Check*
| # | Yêu cầu | Trạng thái | Bằng chứng / ghi chú |
|---|---|---|---|
| 1 | Thực hiện sheet Office Apps Creat & Update | Đang làm | Xem bảng B bên dưới (8 mục). |
| 2 | Tên ứng dụng "Yomi-Office Open Source" (cả tên sau khi cài) | Xong (Android) · Windows: Antigravity | 91 chuỗi giao diện + `APP_NAME`; Android `app_name`; aapt xác nhận label trên 9 APK; tiêu đề trang, màn hình chờ, logo mới. |
| 3 | Rà các project trong sheet Project GitHub, tối ưu & cập nhật | Một phần | Antigravity đồng bộ: 179 clone mới, 52 cập nhật, 217 đã mới nhất, 71 lỗi (58 timeout đang thử lại clone nông, 11 repo không còn tồn tại/riêng tư, 2 đường dẫn không hợp lệ trên Windows). Tích hợp mới ở bản này: Free Claude Code. |
| 4 | Ổn định & hoạt động offline của App Feature Set | Đạt trên emulator | selftest 49/49 trên Android 6 & 14: Word, Excel, PowerPoint, PDF (tạo/mở/trích chữ/DOCX/mật khẩu/nén), OCR vie+eng offline, CamScanner, MathType (KaTeX) — các mục này chạy không cần mạng. |
| 5 | Tính năng, độ ổn định Word/Excel/PowerPoint | Đạt trên emulator | Mở 6/6 (Android 6) và 4/4 (Android 14) tệp thật bằng engine Native, không văng; selftest Word/Excel/PPT khứ hồi DOCX/XLSX/PPTX đạt. |
| 6 | Cập nhật hàm Excel từ các website Excel Functions | Chưa làm ở bản này | Không thêm hàm mới ở v0.1.8. Bộ hàm hiện có từ v0.0.5–v0.0.6 vẫn đạt selftest. |
| 7 | Tích hợp LibreOffice/OpenOffice/Google/OnlyOffice/WPS | Giữ nguyên v0.1.7 | Engine Native = LibreOffice (Android 26.2, Windows 24.8/26.x). WPS là phần mềm đóng nguồn: chỉ có tương đương (JSA macro). Không bổ sung mới ở bản này. |
| 8 | MathType | Đạt selftest | `mathtype: KaTeX render` PASS. MathType gốc đóng nguồn: dùng KaTeX/LaTeX/MathML. |
| 9 | So với Word/Excel/PowerPoint trên máy, bổ sung tính năng thiếu | Antigravity (mục 29) | So sánh bằng `test_compare_v018.ps1` trên Windows. |
| 10 | Tính năng con trỏ chuột Word/Excel/PowerPoint | Giữ nguyên v0.1.7 | Sửa con trỏ Excel ở v0.1.7; bản này không đổi. |
| 11 | Bổ sung tất cả hàm Excel | Chưa làm ở bản này | Như mục 6. |
| 12 | Bộ tính năng A.I | Đạt | selftest A.I offline + Web + API; thêm Free Claude Code (thử proxy giả lập 2/2, selftest đạt). |
| 13 | Đối chiếu ảnh Office Screen, bổ sung tính năng | Chưa làm ở bản này | Không mở lại tài liệu Office Screen trong phiên này. |
| 14 | Word/Excel/PowerPoint đầy đủ & offline | Đạt (phần đã kiểm) | Xem mục 4–5. |
| 15 | Undo/Redo, AutoSave, chọn nhiều ô Excel | Đạt selftest | `excel: undo/redo` PASS; chọn nhiều ô & AutoSave giữ nguyên từ v0.1.7, Antigravity kiểm thủ công trên Note 9. |
| 16 | Tối ưu tích hợp LibreOffice…Craft offline | Giữ nguyên v0.1.7 | — |
| 17 | OOXML, ODF, Craft trong Word/Excel/PowerPoint | Đạt selftest | DOCX/XLSX/PPTX khứ hồi, `.ods/.odt/.odp` khứ hồi, RTF, XML 2003/SYLK/DIF/MHT PASS. |
| 18 | Tương thích Microsoft Office có sẵn | Antigravity (mục 29) | — |
| 19 | VBA & Macro | Đạt selftest | VBA Excel (.xlsm khứ hồi, class module), Word VBA (.docm), PowerPoint VBA (.pptm), JSA PASS. |
| 20 | A.I: web không đăng nhập / đăng nhập / khóa API | Đạt | Có từ v0.1.7; selftest v007 A.I PASS. |
| 21 | Cập nhật phần mềm trong Tùy chọn | **Đã sửa** | Thiếu nguồn chính GitHub Releases → đã thêm; nhận tên bộ cài mới + .msi; 7/7 kiểm thử đơn vị + selftest v018 PASS. |
| 22 | CamScanner | Đạt selftest | Nhận mép, nắn phối cảnh, lọc, xuất PDF PASS. |
| 23 | PDF: tích hợp mục I, V, VI, VIII, XIII | Giữ nguyên + đạt selftest | — |
| 24 | Giao diện Word/Excel/PowerPoint | Đạt (ảnh) | Ảnh chụp màn hình Android 6/14, Note 9. |
| 25 | Tách trang, chia trang, thước Word & Excel | Đạt selftest | `v008 Word phân trang`, `v008 Excel ngắt trang`, `Word thước kẻ` PASS. |
| 26 | Đồng bộ Google Drive/OneDrive/Dropbox/MediaFire + sửa chung | Một phần | Selftest đám mây (máy chủ giả lập) PASS. Sửa chung thời gian thực cần OAuth Client ID của chủ dự án — chưa có. |
| 27 | Giao diện trình đọc PDF theo tài liệu ảnh | Giữ nguyên v0.0.8 | — |
| 28 | Quét toàn bộ tệp trên máy, lọc Word/Excel/PowerPoint | **Đã sửa (lỗi tồn từ v0.1.6)** | Hàm Android có nhưng chưa từng được gọi → nối giao diện, quét nền, tab "Trên thiết bị", Windows `/api/scan_office`. Android 6: 6/6, Android 14: 4/4, máy ảo: 23/23. |
| 29 | Chạy Test Office so với MS Office gốc, ảnh chụp | Antigravity | — |
| 30 | Hàm liên kết nhiều sheet / nhiều bảng tính | Đạt | 9/9 (node) + selftest 6/6 trên Android 6 & 14. Nghi lỗi `_3dSheets` → kiểm tra: không lỗi. |
| 31 | Lưu khóa API trên máy | Đạt | Có từ v0.1.7 (localStorage + Host.saveConfig; Windows DPAPI). |
| 32 | AirDrop & QuickShare | Giới hạn | AirDrop là giao thức riêng của Apple — dùng Quick Share/chia sẻ Android + LAN Share. |
| 33 | Android DeX, trình chiếu lên TV, ảnh chụp | Đang kiểm | Emulator màn hình phụ 1920×1080 + Note 9 (Antigravity). DeX thật cần cáp HDMI/TV. |

### C. Trạng thái — sheet *Office Apps Creat & Update*
| Yêu cầu | Trạng thái | Ghi chú |
|---|---|---|
| Tương thích & hiển thị tệp MS Office | Đạt phần đã kiểm | Mở tệp thật bằng engine Native không văng; so sánh chi tiết với MS Office: Antigravity (mục 29). |
| Lỗi con trỏ chuột khi mở Excel | Giữ bản sửa v0.1.7 | — |
| Đồng nhất giao diện WebView & Native | Giữ bản v0.1.7 | — |
| Tích hợp tất cả tính năng các bộ Office (bắt buộc) | Một phần | Không thể làm trọn trong một bản; xem mục 3, 7, 16. |
| PowerPoint Animation | Giữ bản v0.1.7 | 12 hiệu ứng + ghi vào .pptx; trình chiếu chưa phát hoạt hình từng đối tượng (hạn chế cũ). |
| Đổi logo sang Yomi-Logo | **Xong** | Android mipmap 48–192 px, logo trong giao diện, favicon, app.ico; Windows: Antigravity. |
| README theo mẫu + nút Download + Trang chủ/Chính sách | Đang làm | Phần phát hành. |
| Tích hợp Free Claude Code vào cài đặt A.I | **Xong** | Proxy `http://127.0.0.1:8082/v1/messages`, token tùy chọn, địa chỉ sửa được (AGPL: chỉ kết nối, không nhúng mã). |

### D. Windows (Antigravity)
_(hợp nhất từ `v018\coord\TIEN_DO_antigravity.md` khi hoàn tất)_

### E. Kiểm thử thiết bị thật Samsung Galaxy Note 9 (Antigravity)
_(hợp nhất khi hoàn tất)_

### F. Phát hành GitHub
_(ghi khi hoàn tất)_
