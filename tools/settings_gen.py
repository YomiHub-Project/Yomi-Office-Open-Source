# Office Open Source v0.1.7 - shared Backstage + Options spec used by BOTH the WebView editor and the Native (LibreOfficeKit) editor.
# Output: src/core/ui_settings.json  (embedded as a resource in the Windows launcher, imported by the web bundle)
import json, re, sys

def chk(key, t, d=False, nat=None, only=None): return dict(k='chk', key=key, t=t, d=d, nat=nat, only=only)
def num(key, t, d, mn, mx, nat=None, only=None): return dict(k='num', key=key, t=t, d=d, min=mn, max=mx, nat=nat, only=only)
def sel(key, t, d, opts, nat=None, only=None): return dict(k='sel', key=key, t=t, d=d, opts=opts, nat=nat, only=only)
def txt(key, t, d='', nat=None): return dict(k='text', key=key, t=t, d=d, nat=nat)
def btn(act, t, icon): return dict(k='btn', act=act, t=t, i=icon)
def note(t, only=None): return dict(k='note', t=t, only=only)
def sec(h, *items, sub=None, only=None): return dict(h=h, sub=sub, only=only, items=[i for i in items])

FONTS = [['Calibri', 'Calibri'], ['Times New Roman', 'Times New Roman'], ['Arial', 'Arial'], ['Roboto', 'Roboto'], ['Segoe UI', 'Segoe UI']]

options = [
  dict(id='general', t='Chung', secs=[
    sec('Tùy chọn chung khi làm việc với Office Open Source'),
    sec('Tùy chọn giao diện người dùng',
        chk('ribbonCollapsed', 'Thu gọn Ribbon khi mở tài liệu', False, nat='ribbon.min'),
        sel('theme', 'Chủ đề Office:', 'light', [['light', 'Sáng'], ['dark', 'Tối'], ['auto', 'Theo hệ thống']]),
        chk('glass', 'Hiệu ứng Liquid Glass (kính mờ)', True),
        chk('touchMode', 'Chế độ Cảm ứng (nút Ribbon lớn hơn, dễ chạm)', False)),
    sec('Cá nhân hóa bản sao Office Open Source của bạn',
        txt('userName', 'Tên người dùng:', '', nat='user'),
        txt('initials', 'Viết tắt:', '', nat='initials')),
    sec('Tùy chọn khởi động',
        sel('officeEngine', 'Mở tệp Word / Excel / PowerPoint bằng:', 'native', [['native', 'Native – LibreOffice core (khuyên dùng)'], ['web', 'Web – trình soạn Office Open Source']]),
        chk('newWindowPerFile', 'Mở mỗi tệp Word / Excel / PowerPoint / PDF trong cửa sổ mới', True)),
  ]),
  dict(id='display', t='Hiển thị', tx='Công thức', secs=[
    sec('Tùy chọn tính toán', note('Tính toán tự động / thủ công: Công thức → Tính toán. Hiện công thức: Công thức → Kiểm tra công thức.'), only='excel'),
    sec('Hiển thị',
        chk('rulers', 'Hiện thước (Word)', True, nat='word.ruler'),
        chk('gridlines', 'Hiện đường lưới (Excel)', True),
        num('zoomDefault', 'Thu phóng mặc định khi mở (%, 0 = tự động)', 0, 0, 400, nat='zoom.default'),
        chk('cursorBlink', 'Con trỏ văn bản nhấp nháy', True)),
    sec('Vùng an toàn & Chống che khuất',
        chk('antiObscuration', 'Tự động chống che khuất bởi thanh trạng thái & thanh điều hướng', True),
        num('safePadTop', 'Đệm viền trên bổ sung (px)', 0, 0, 100),
        num('safePadBottom', 'Đệm viền dưới bổ sung (px)', 0, 0, 120)),
  ]),
  dict(id='proofing', t='Kiểm lỗi', secs=[
    sec('Khi sửa chính tả trong các chương trình',
        chk('spellAuto', 'Kiểm tra chính tả khi nhập', True, nat='spell.auto'),
        sub='Bật/tắt nhanh ở Xem lại → Chính tả.'),
  ]),
  dict(id='save', t='Lưu', secs=[
    sec('Lưu tài liệu',
        chk('autosave', 'Lưu thông tin Tự động khôi phục', True),
        num('autosaveMin', 'Lưu thông tin Tự động khôi phục mỗi (phút):', 2, 1, 120, nat='autosave.minutes'),
        chk('autosaveLive', 'Tự động lưu tệp gốc sau mỗi thay đổi (Tự động lưu)', False, nat='autosave.live'),
        txt('saveDir', 'Vị trí tệp cục bộ mặc định:', '', nat='save.dir'),
        num('recentMax', 'Số tệp gần đây tối đa:', 30, 5, 200)),
  ]),
  dict(id='lang', t='Ngôn ngữ', secs=[
    sec('Ngôn ngữ', note('Ngôn ngữ soạn thảo mặc định: Tiếng Việt (vi-VN). Ngôn ngữ giao diện: Tiếng Việt.')),
  ]),
  dict(id='advanced', t='Nâng cao', secs=[
    sec('Chuột & con trỏ',
        chk('mouse.wheelZoomCtrl', 'Ctrl + lăn chuột để phóng to/thu nhỏ', True),
        chk('mouse.smoothScroll', 'Cuộn mượt', True),
        num('mouse.scrollLines', 'Số dòng cuộn mỗi nấc lăn chuột', 3, 1, 20),
        chk('mouse.rightClickMenus', 'Menu chuột phải', True),
        chk('mouse.dragDrop', 'Kéo thả văn bản / ô / đối tượng', True),
        chk('mouse.doubleClickEdit', 'Bấm đúp để sửa ô / đối tượng', True),
        chk('mouse.miniToolbar', 'Hiện thanh công cụ nhỏ khi chọn văn bản', True),
        chk('mouse.hoverPreview', 'Xem trước trực tiếp khi rê chuột', True),
        sel('mouse.cursorStyle', 'Kiểu con trỏ', 'default', [['default', 'Mặc định của hệ thống'], ['large', 'Con trỏ lớn'], ['precise', 'Con trỏ chính xác (dấu cộng)']]),
        chk('mouse.middleAutoscroll', 'Bấm nút giữa (con lăn) để cuộn tự động', True),
        sel('mouse.caretWidth', 'Kích thước con trỏ văn bản', '1', [['1', 'Mảnh'], ['2', 'Vừa'], ['3', 'Dày']]),
        chk('mouse.clickHighlight', 'Hiện vệt sáng khi bấm chuột', False)),
    sec('Tùy chọn chỉnh sửa',
        chk('mouse.typeToEdit', 'Excel: gõ trực tiếp vào ô đang chọn', True),
        sel('mouse.enterMove', 'Excel: sau khi nhấn Enter, di chuyển vùng chọn', 'down', [['down', 'Xuống'], ['right', 'Sang phải'], ['up', 'Lên'], ['left', 'Sang trái'], ['none', 'Không di chuyển']]),
        chk('mouse.wheelZoomNoCtrl', 'Excel: lăn chuột để phóng to/thu nhỏ', False),
        chk('mouse.clickAndType', 'Word: Bấm-và-Gõ (bấm đúp vào chỗ trống để gõ)', True)),
    sec('Phông chữ mặc định',
        sel('wordFont', 'Phông chữ mặc định Word', 'Calibri', FONTS),
        num('wordSize', 'Cỡ chữ mặc định Word', 11, 6, 72),
        sel('excelFont', 'Phông chữ mặc định Excel', 'Calibri', FONTS)),
  ]),
  dict(id='ribbon', t='Tùy chỉnh Ribbon', secs=[
    sec('Tùy chỉnh Ribbon', dict(k='tabs', key='ribbonHide', nat='ribbon.hide', d=''), sub='Chọn các tab hiển thị trên Ribbon (áp dụng khi mở tài liệu tiếp theo).'),
  ]),
  dict(id='addins', t='Bổ trợ', secs=[
    sec('Trợ lý A.I', note('Web (ChatGPT, Gemini, Copilot, Claude, DeepSeek, Grok) hoặc khóa API'), btn('ai', 'Cài đặt A.I…', 'l:sparkles')),
    sec('Phông chữ', note('Phông tích hợp, Font Pack và phông của máy'), btn('fonts', 'Quản lý phông chữ…', 'l:type')),
    sec('Đám mây', note('Google Drive · OneDrive · Dropbox · MediaFire'), btn('cloud', 'Tài khoản & đồng bộ…', 'l:cloud')),
  ]),
  dict(id='update', t='Cập nhật', secs=[
    sec('Cập nhật phần mềm',
        chk('update.auto', 'Tự động kiểm tra phiên bản mới khi mở ứng dụng', True, nat='update.auto'),
        chk('update.autoInstall', 'Tự động tải và cài đặt bản cập nhật', False),
        btn('updateNow', 'Kiểm tra cập nhật ngay', 'update'),
        sub='Nguồn: GitHub Releases (chính) · Google Drive và MediaFire (phụ) của Yomi-Office Open Source.'),
  ]),
  dict(id='dex', t='DeX & Trình chiếu', secs=[
    sec('Android DeX, chiếu màn hình & PC Link',
        btn('dex', 'Trung tâm Cài đặt & Điều khiển DeX…', 'l:monitor'),
        chk('dex.enabled', 'Bắt buộc bật giao diện DeX Desktop Mode', False),
        sel('dex.resolution', 'Độ phân giải hiển thị DeX', 'auto', [['auto', 'Tự động'], ['1080p', 'Full HD (1920 × 1080)'], ['1440p', '2K QHD (2560 × 1440)'], ['4k', '4K UHD (3840 × 2160)']]),
        sel('dex.uiScale', 'Tỷ lệ thu phóng giao diện', '100', [['75', '75%'], ['85', '85%'], ['100', '100%'], ['115', '115%'], ['125', '125%'], ['150', '150%']]),
        chk('dex.autoLaunchOnHdmi', 'Tự động bật DeX khi cắm cáp HDMI / DisplayPort', True),
        chk('dex.wakeLock', 'Giữ màn hình luôn sáng khi kết nối DeX', True),
        sel('dex.keyboardProfile', 'Sơ đồ phím tắt bàn phím ngoài', 'windows', [['windows', 'Chuẩn Windows'], ['dex', 'Chuẩn Samsung DeX'], ['mac', 'Chuẩn Apple Mac']]),
        chk('dex.webLinkEnabled', 'Máy chủ Web Link qua Wi-Fi nội bộ', False),
        num('dex.webLinkPort', 'Cổng máy chủ Web Link', 8080, 1024, 65535)),
  ]),
  dict(id='trust', t='Trung tâm Tin cậy', secs=[
    sec('Trung tâm Tin cậy',
        chk('macrosAllowed', 'Cho phép chạy macro trong tài liệu (.docm, .xlsm, .pptm)', False, nat='__macros'),
        sub='Macro trong tài liệu (VBA / Basic) có thể chứa mã độc. Chỉ bật khi tin cậy nguồn tệp.'),
  ]),
]

# Backstage: same navigation, same tiles, same file-type list in both editors.
nav = [['home', 'Trang chủ', 'home'], ['new', 'Mới', 'blankpage'], ['open', 'Mở', 'open'], None,
       ['info', 'Thông tin', 'info'], ['save', 'Lưu', 'save'], ['saveas', 'Lưu thành', 'saveas'], ['print', 'In', 'print'],
       ['share', 'Chia sẻ', 'share'], ['export', 'Xuất', 'export'], ['close', 'Đóng', 'close'], None,
       ['account', 'Tài khoản', 'account'], ['options', 'Tùy chọn', 'options']]

tiles = {
  'open': [['browse', 'folder', 'Duyệt', 'Mở tệp từ máy tính này'],
           ['pc', 'pc', 'Máy tính này', 'Tài liệu'],
           ['cloudOpen', 'l:cloud', 'Đám mây', 'Google Drive · OneDrive · Dropbox · MediaFire']],
  'info': [['reveal', 'open', 'Mở vị trí tệp', 'Hiện tệp trong thư mục chứa nó'],
           ['protect', 'protect', 'Bảo vệ tài liệu', 'Đặt mật khẩu hoặc đánh dấu bản cuối cùng (chỉ đọc)'],
           ['macros', 'macros', 'Macro', 'Cho phép/chặn macro VBA, Basic trong tài liệu'],
           ['inspect', 'l:search-check', 'Kiểm tra tài liệu', 'Thống kê nội dung, siêu dữ liệu, xóa thông tin cá nhân'],
           ['versions', 'refresh', 'Phiên bản', 'Xem và khôi phục các bản tự động lưu']],
  'saveas': [['recent', 'refresh', 'Gần đây', None], ['pc', 'pc', 'Máy tính này', None], ['addplace', 'addplace', 'Thêm vị trí', None],
             ['browse', 'folder', 'Duyệt', None], ['cloudSave', 'l:cloud', 'Đám mây', None]],
  'share': [['lanshare', 'lanshare', 'Gửi qua Wi-Fi (AirDrop / Quick Share)', 'Tạo liên kết tải về cho máy khác cùng mạng, có thể cùng sửa'],
            ['email', 'email', 'Gửi email dưới dạng tệp đính kèm', 'Mở thư mới và đính kèm bản sao tài liệu'],
            ['sendpdf', 'pdf', 'Gửi dưới dạng PDF', 'Xuất PDF rồi chọn nơi gửi']],
  'export': [['pdf', 'pdf', 'Tạo tài liệu PDF', 'Giữ nguyên bố cục, phông chữ và hình ảnh; khó bị sửa đổi']],
}
# file types: [ext, label, where]  where: b = both editors, n = native engine only, w = web engine only
formats = {
  'word': [['docx', 'Tài liệu Word', 'b'], ['docm', 'Tài liệu Word có macro', 'b'], ['doc', 'Tài liệu Word 97-2003', 'b'], ['dotx', 'Mẫu Word', 'b'],
           ['odt', 'Văn bản OpenDocument', 'b'], ['rtf', 'Định dạng Rich Text', 'b'], ['pdf', 'PDF', 'b'], ['html', 'Trang web', 'b'],
           ['txt', 'Văn bản thuần', 'b'], ['epub', 'Sách điện tử EPUB', 'n'], ['md', 'Markdown', 'w'], ['mht', 'Trang web một tệp', 'w'], ['xml', 'Tài liệu Word XML', 'w']],
  'excel': [['xlsx', 'Sổ làm việc Excel', 'b'], ['xlsm', 'Sổ làm việc Excel có macro', 'b'], ['xls', 'Sổ làm việc Excel 97-2003', 'n'], ['xltx', 'Mẫu Excel', 'b'],
            ['ods', 'Bảng tính OpenDocument', 'b'], ['csv', 'CSV (phân tách bằng dấu phẩy)', 'b'], ['pdf', 'PDF', 'b'], ['html', 'Trang web', 'b'],
            ['tsv', 'Văn bản phân tách Tab', 'w'], ['json', 'Dữ liệu JSON', 'w'], ['xml2003', 'Bảng tính XML 2003', 'w'], ['slk', 'SYLK', 'w'], ['dif', 'DIF', 'w'], ['mht', 'Trang web một tệp', 'w']],
  'ppt': [['pptx', 'Bản trình bày PowerPoint', 'b'], ['pptm', 'Bản trình bày PowerPoint có macro', 'b'], ['ppt', 'Bản trình bày PowerPoint 97-2003', 'n'], ['potx', 'Mẫu PowerPoint', 'b'],
          ['ppsx', 'Trình chiếu PowerPoint', 'b'], ['odp', 'Bản trình bày OpenDocument', 'b'], ['pdf', 'PDF', 'b'], ['png', 'Ảnh PNG (trang chiếu hiện tại)', 'b'],
          ['jpg', 'Ảnh JPEG (trang chiếu hiện tại)', 'w'], ['svg', 'Ảnh SVG', 'n'], ['html', 'Trình chiếu HTML', 'w'], ['md', 'Dàn ý Markdown', 'w'], ['rtf', 'Dàn ý RTF', 'w'], ['video', 'Video trình chiếu (*.webm)', 'w']],
}
for o in options:
    for s in o['secs']:
        for i in s['items']:
            for k in [k for k, v in list(i.items()) if v is None]: del i[k]
        for k in [k for k, v in list(s.items()) if v is None]: del s[k]
spec = dict(options=options, nav=nav, tiles=tiles, formats=formats)
out = sys.argv[1] if len(sys.argv) > 1 else 'ui_settings.json'
json.dump(spec, open(out, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))

# lucide glyphs used by the spec -> add to ui_layout.json "lucide" so the native renderer can draw them
if len(sys.argv) > 3:
    icons_js, layout = sys.argv[2], sys.argv[3]
    src = open(icons_js, encoding='utf-8').read()
    L = json.load(open(layout, encoding='utf-8'))
    need = set(re.findall(r'"l:([a-z0-9-]+)"', json.dumps(spec)))
    for n in need:
        m = re.search('"' + re.escape(n) + r'":("(?:[^"\\]|\\.)*")', src)
        if m: L['lucide'][n] = json.loads(m.group(1))
        else: print('missing lucide', n)
    json.dump(L, open(layout, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print('ok', out)
