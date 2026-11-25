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
| 25/11/2026 15:35 | Build lại web + 9 APK (job 040) | Logo mới + phiên bản động | `build_v018_android.ps1` | Gradle (tăng dần ~6 phút) | 9/9 APK, 78–93 MB |
| 25/11/2026 15:51 | Emulator Android 6 (lượt 3, APK cuối) | Video: `adb emu screenrecord` tối đa 180 s + console cắt đường dẫn có dấu cách | Quay từng đoạn 175 s vào `%TEMP%` (Start-Job) + `ffmpeg -f concat` | adb, ffmpeg | **21/23**: selftest 49/49, quét 8/8, mở 8/8 tệp (.doc/.docx/.xlsm/.pptx), DeX desktop hiển thị (thanh tác vụ tiếng Việt, dock ứng dụng); 2 FAIL = video (script cũ) |
| 25/11/2026 16:00 | Emulator Android 14 x86_64 + TV | CPU chọn đúng x86_64 (không còn "System UI isn't responding") | — | adb, ffmpeg | **22/23**: selftest 49/49, mở 6/6, Presenter Console "MÀN HÌNH NGOÀI: ĐANG CHIẾU", **video 2.4 MB (3 đoạn)**; FAIL: `adb emu multidisplay` chưa tạo màn hình phụ |
| 25/11/2026 16:01 | Chuỗi phiên bản cứng trong DeX | Ghi chú Presenter "v0.1.6", báo cáo mẫu "V0.1.6", slide mẫu, URL `office-open-source/office-v0.1.6`, tên tệp mẫu | `core/dex.js` (6 chỗ) → `VERSION` / repo YomiHub | sed/python, node --check | build riêng cho Android (`-Suffix -a`, job 049) để không đè `dist` đang dùng cho đóng gói Windows (DeX chỉ có trên Android) |
| 25/11/2026 16:05 | Note 9 thật (Antigravity) | — | `test_android_v018.ps1 -Serial 281038ac31217ece -Dex -Tv` | adb | **37/39**: selftest 49/49, mở **23/23** tệp Test Office không văng; FAIL 1: script đọc chuỗi tiếng Anh "Display" nhưng OneUI tiếng Việt in "Lớp phủ #1" — màn hình phụ 1920×1080 **đã tạo** (mDisplayId=6) → Claude sửa script đếm `mDisplayId`; FAIL 2: scrcpy 0 MB → Antigravity quay bằng `screenrecord` khi kiểm thử thủ công |
| 25/11/2026 16:24 | Emulator Android 6 lượt cuối (APK 16:15, có sửa dex.js) | — | job 049 | adb, ffmpeg | **19/19 đạt** + **video** `android_emu_android6_final_v0.1.8.mp4` (3 đoạn, 2.4 MB) |
| 25/11/2026 16:20 | `licenses.txt` (nút "Giấy phép mã nguồn mở") + dòng đầu `app.js` ghi cứng "Office Open Source v0.1.1/v0.1.6" | `tools/build.js` đọc `APP_NAME`/`VERSION` từ `core/util.js`; vá trực tiếp 6 thư mục dist | python, node --check | 6/6 dist đã đúng "Yomi-Office Open Source v0.1.8" |
| 25/11/2026 16:27 | **Lưu tiến độ tự động 5 phút/lần** (chủ dự án yêu cầu 16:20) | Lần chạy đầu thoát sau 1 vòng (chưa rõ nguyên nhân; bản thử 3 vòng chạy đúng) → chạy lại với nhật ký lý do thoát | `v018/checkpoint_loop.ps1` → job `001_checkpoint_loop`: `coord/TIEN_DO_HIEN_TAI.md`, `coord/CHECKPOINT_LOG.md`, `checkpoints/<giờ>/` (giữ 36 bản) | PowerShell, ZipFile | bản 16:21, 16:27 đã lưu |
| 25/11/2026 16:28 | Build 9 APK cuối (dist-a có licenses mới) | — | job 061 `-SkipWeb -Suffix -a` | Gradle | đang chạy |
| 25/11/2026 16:35 | Sửa theo kết quả kiểm thử Note 9 (Antigravity) | (1) Tab Tùy chọn → Cập nhật vẫn ghi "Nguồn: Google Drive và MediaFire"; (2) hộp Tùy chọn 740 px tràn ngang trên màn hình dọc; (3) thiếu ô cấu hình nguồn GitHub | `tools/settings_gen.py:95` + `core/ui_settings.json` (mô tả nguồn mới); `core/update.js optionsSection`: ô "Kho GitHub (nguồn chính)", input `maxWidth:100%`; `styles/zz-backstage.css`: `@media (max-width: 780px)` (hộp 96vw, danh mục thành hàng tab cuộn ngang, nhãn/ô xuống dòng); vá chuỗi trong `dist*/app.js` cho Windows | python, sed | build cuối job 064 |
| 25/11/2026 16:35 | Hạn chế đã biết (không sửa ở v0.1.8) | Note 9: `t06.xlsm` (15.8 MB) và `t18.xlsx` (6.8 MB) mở bằng engine Native > 35 s | — | — | ghi vào RELEASE_NOTES |
| 25/11/2026 16:39 | Build 9 APK CUỐI (job 064) | Có: CSS Tùy chọn màn hình hẹp, nguồn GitHub, ô "Kho GitHub", licenses v0.1.8, dex.js | `build_v018_android.ps1 -Suffix -a` | Gradle | 9/9 APK; dist-a kiểm tra đủ 5 thay đổi |
| 25/11/2026 16:44 | Kiểm thử khói APK cuối, emulator Android 14 (job 066) | — | `test_android_v018.ps1 -Avd YomiHub056_Phone_API34 -MaxFiles 3` | adb | cài đặt, phiên bản 0.1.8 (18033, x86_64), quét tệp, A.I: PASS (selftest đang chạy) |
| 25/11/2026 16:42 | `version.json` + `SHA256_v0.1.8.md` | Khóa Windows `-exe/-msi/-zip` để update.js chọn theo tên tệp (MSI đúng) | `v018/publish/make_manifest_v018.ps1` | node (5/5 đúng) | chạy khi Windows xong |
| 25/11/2026 16:44 | **Đẩy lên GitHub — đợt 1** (job 070) | Mã nguồn + README + docs + Release v0.1.8 + 17 tệp đã có (9 APK + 8 bộ cài Windows), 2.35 GB | `v018/publish/publish_v018.ps1` | git (GitHub Desktop, GCM), curl, GitHub REST | ĐANG CHẠY — chạy lại sau khi Windows xong để bổ sung tệp còn thiếu |
| 25/11/2026 16:44 | Phối hợp | Đã gửi `/swarm` nhờ Antigravity kiểm lại hộp Tùy chọn trên Note 9 bằng APK cuối + báo thời gian xong Windows | — | computer-use | chờ |

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
- Host Windows (`v018\src\platform\windows\*.cs`, 9 tệp C#): đổi tên/logo, `/api/info` trả `"msi"`, cập nhật qua `msiexec /i … /passive`, `POST /api/scan_office`, Free Claude Code, `ListGitHub`; biên dịch Roslyn `csc.exe` 0 lỗi.
- `build_v018_windows.ps1` (UTF-8 BOM, `$PROJ` động): sửa WiX `CNDL0230` (GUID MD5 cho KeyPath HKCU), `CNDL0117`, `-sice:ICE60`; loader ARM64; thêm bản Universal (All-AnyCPU → `Yomi-Office-Open-Source-0.1.8-Setup.exe/.msi/-portable.zip`, không ghi edition, giống quy ước bản phổ thông v0.1.7).
- Tiến độ lúc 17:26: đã có Win7-8.1 x86/x64, Win10-11 x86/x64/ARM64, Universal (Setup + .msi + portable); đang đóng gói lại các bản build trước 16:20 (`licenses.txt` cũ) và `.appxbundle`. Kiểm thử cài đặt + so sánh 23 tệp với MS Office: chờ (`test_compare_v018.ps1`).
_(Kết quả kiểm thử Windows sẽ hợp nhất khi có `coord\WIN_DONE.txt`.)_

### E. Kiểm thử thiết bị thật Samsung Galaxy Note 9 (Antigravity)
Thiết bị: SM-N960N, Android 10 (API 29), arm64-v8a, serial `281038ac31217ece`. Nguồn: `v018\coord\TIEN_DO_antigravity.md`.

| Hạng mục | Kết quả | Ghi chú / bằng chứng |
|---|---|---|
| Script `test_android_v018.ps1 -Serial 281038ac31217ece -Dex -Tv` (15:45–15:59, APK job 040) | **37 đạt / 2 lỗi** | 2 lỗi đều do script kiểm thử: regex đếm màn hình phụ không khớp chuỗi OneUI tiếng Việt ("Lớp phủ #1", màn hình 1920×1080 thực tế đã tạo); scrcpy cho video 0 MB |
| Cài đặt, định danh gói | Đạt | versionName 0.1.8-android10, versionCode 18032, arm64-v8a |
| `#selftest` | **49/49** | gồm 5 mục mới v0.1.8 |
| Tab "Trên thiết bị" | Đạt | quét 105 tệp (Word 74, Excel 18, PowerPoint 13) |
| Mở 23 tệp Test Office | **23/23 không văng** | 21 hiển thị nhanh; `t06.xlsm` (15.8 MB) và `t18.xlsx` (6.8 MB) > 35 s bằng engine Native → hạn chế đã biết |
| Word / Excel / PowerPoint, PDF, MathType, CamScanner | Đạt (thủ công) | ảnh `manual_01b…06` trong `v018\shots\android` |
| Trợ lý A.I 3 chế độ + Free Claude Code | Đạt | `manual_07c_free_claude_code.png` |
| Quick Share | Đạt | Share Sheet có "Chia sẻ nhanh" |
| DeX desktop / trackpad / TV home | Đạt | `android_device_40_dex_desktop.png`, `android_device_52_tv_home.png` |
| Video toàn phiên | Đạt (quay bù `adb shell screenrecord` + ffmpeg) | `v018\video\android\android_device_v0.1.8.mp4` (14.1 MB) |
| Tùy chọn trên màn hình dọc + mô tả nguồn cập nhật | Lỗi lúc 16:11 → **đã sửa** (build 16:38), Antigravity xác minh lại: Đạt | `note9_final_options_chung.png`, `note9_final_options_capnhat.png` |


### F. Phát hành GitHub
| Lúc | Đợt | Kết quả |
|---|---|---|
| 16:47 | Đợt 1 (job 070) | Đẩy mã nguồn `9945e17` (257 tệp); tạo Release v0.1.8 (id 457911185); 17/17 tệp (9 APK + 8 bộ cài Windows); `releases/latest` = v0.1.8 |
| 17:04 | Đợt 1b (job 093, chủ dự án yêu cầu) | Commit `4ccaccb` (README, ghi chú phát hành + kỹ thuật); cập nhật nội dung Release; thêm Win10-11-x64 Setup/portable, Win10-11-x86.msi → **20/20 tệp, 2.96 GB**; GitHub Pages bật (`main /docs`) |
| chờ | Đợt 2 (job 090, tự chạy khi có `WIN_DONE.txt`) | `version.json` + `SHA256_v0.1.8.md`, các bộ cài Windows còn lại/đóng gói lại, `.appxbundle` |

`publish_v018.ps1` mới: `-MinAgeSec` bỏ qua tệp vừa ghi (tránh tải tệp đang build); khi Release đã có thì cập nhật nội dung theo `RELEASE_NOTES_v0.1.8.md`; tải lại tệp khi kích thước khác.
