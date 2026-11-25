// PDF ribbon: Home, Edit pages, Convert, Tools, View, Help
import { h, ic } from '../core/util.js';
import { toast, showMenu } from '../core/ui.js';
import * as T from './tools.js';
import { helpTab } from '../core/help.js';
import * as U from './ui8.js';

const sel = d => (d.sel.length ? d.sel : [d.cur || 0]);
const need = (d, fn) => () => { if (!d.pdf) return toast('Mở một PDF trước', 'err'); fn(); };

export function ribbonDef(doc) {
  // v0.0.8: tab order after Foxit PDF Reader — Trang đầu · Chú thích · Xem · Biểu mẫu · Bảo vệ & Ký · Sắp xếp · Chuyển đổi · Công cụ · Chia sẻ · Trợ giúp
  const help = helpTab(doc, 'pdf');
  help.groups.push({ label: 'Trình đọc PDF', items: [
    { type: 'big', label: 'Đặt làm trình đọc PDF mặc định', icon: 'badge-check', run: () => import('../core/host.js').then(H => { const Host = H.default; if (Host.kind === 'android') toast('Mở một tệp PDF bằng "Mở bằng" → chọn Yomi-Office Open Source → "Luôn luôn"'); else Host.openUrl('ms-settings:defaultapps'); }) },
    { type: 'big', label: 'Phím tắt PDF', icon: 'keyboard', run: () => import('../core/ui.js').then(u => u.dialog({ title: 'Phím tắt trình đọc PDF', width: 460, body: h('table.kv', [['Trang sau / trước', 'PgDn / PgUp (← → ở chế độ một/hai trang)'], ['Trang đầu / cuối', 'Ctrl+Home / Ctrl+End'], ['Tìm kiếm', 'Ctrl+F'], ['Phóng to / thu nhỏ', 'Ctrl + lăn chuột'], ['Dừng tự cuộn', 'Esc'], ['In', 'Ctrl+P']].map(r => h('tr', h('td', r[0]), h('td', r[1])))), buttons: [{ label: 'Đóng', primary: true }] })) }
  ] });
  return { tabs: [homeTab(doc), editTab(doc), viewTab(doc), formTab(doc), protectTab(doc), pagesTab(doc), convertTab(doc), toolsTab(doc), shareTab(doc), help] };
}
const PRO = (fn, a) => () => import('./pro.js').then(P => P[fn].apply(null, a));
// v0.0.7: Edit & Protect (annotate, sign, forms, compress, password, properties, outline, search, translate)
function editTab(d) {
  return { id: 'edit', label: 'Chú thích', en: 'Comment', groups: [
    { label: 'Công cụ', items: [{ type: 'col', items: [
      { type: 'btn', label: 'Tay', icon: 'hand', active: () => d.toolMode === 'hand', run: () => U.setTool(d, 'hand') },
      { type: 'btn', label: 'Chọn chữ', icon: 'text-cursor', active: () => d.toolMode === 'select', run: () => U.setTool(d, 'select') },
      { type: 'btn', label: 'Ảnh chụp nhanh', icon: 'camera', active: () => d.toolMode === 'snap', run: () => U.setTool(d, 'snap') }] }] },
    { label: 'Đánh dấu văn bản', items: [
      { type: 'big', label: 'Tô sáng', icon: 'highlighter', run: need(d, PRO('annotate', [d, 'highlight'])) },
      { type: 'col', items: [
        { type: 'btn', label: 'Bút chì', icon: 'pen-line', run: need(d, PRO('annotate', [d, 'pen'])) },
        { type: 'btn', label: 'Máy đánh chữ', icon: 'type', run: need(d, PRO('annotate', [d, 'text'])) },
        { type: 'btn', label: 'Hình / mũi tên', icon: 'square', run: need(d, PRO('annotate', [d, 'rect'])) }] },
      { type: 'big', label: 'Tẩy trắng', icon: 'eraser', run: need(d, PRO('annotate', [d, 'whiteout'])) }
    ] },
    { label: 'Con dấu', items: [
      { type: 'big', label: 'Đóng dấu', icon: 'stamp', run: need(d, PRO('annotate', [d, 'stamp'])) },
      { type: 'big', label: 'Dấu thời gian', icon: 'clock', run: need(d, () => U.timeStamp(d)) }
    ] },
    { label: 'Đo', items: [{ type: 'big', label: 'Đo khoảng cách', icon: 'ruler', active: () => d.toolMode === 'measure', run: need(d, () => U.setTool(d, 'measure')) }] },
    { label: 'Quản lý chú thích', items: [
      { type: 'big', label: 'Danh sách chú thích', icon: 'message-square', run: need(d, () => U.togglePanel(d, 'comments')) },
      { type: 'big', label: 'Tệp đính kèm', icon: 'paperclip', run: need(d, () => U.togglePanel(d, 'attach')) }
    ] },
    { label: 'Quét', items: [{ type: 'big', label: 'CamScanner', icon: 'scan-line', run: () => import('../scan/scan.js').then(S => S.openScanner(d.shell)) }] }
  ] };
}
function formTab(d) {
  return { id: 'form', label: 'Biểu mẫu', en: 'Form', groups: [
    { label: 'Điền', items: [
      { type: 'big', label: 'Điền biểu mẫu', icon: 'clipboard-list', run: need(d, PRO('formsDialog', [d])) },
      { type: 'big', label: 'Đặt lại biểu mẫu', icon: 'rotate-ccw', run: need(d, () => U.resetForm(d)) }
    ] },
    { label: 'Dữ liệu', items: [
      { type: 'big', label: 'Biểu mẫu → bảng tính', icon: 'sheet', run: need(d, () => U.exportFormData(d, 'sheet')) },
      { type: 'col', items: [
        { type: 'btn', label: 'Nhập dữ liệu (XFDF/JSON/CSV)', icon: 'download', run: need(d, () => U.importFormData(d)) },
        { type: 'btn', label: 'Xuất dữ liệu XFDF', icon: 'upload', run: need(d, () => U.exportFormData(d, 'xfdf')) },
        { type: 'btn', label: 'Xuất dữ liệu CSV', icon: 'file-spreadsheet', run: need(d, () => U.exportFormData(d, 'csv')) }] }
    ] },
    { label: 'Ký & Đánh dấu', items: [
      { type: 'big', label: 'Ký tên', icon: 'signature', run: need(d, PRO('annotate', [d, 'sign'])) },
      { type: 'big', label: 'Thêm chữ', icon: 'type', run: need(d, PRO('annotate', [d, 'text'])) }
    ] }
  ] };
}
function protectTab(d) {
  return { id: 'protect', label: 'Bảo vệ & Ký', en: 'Protect', groups: [
    { label: 'Ký', items: [
      { type: 'big', label: 'Ký tên', icon: 'signature', run: need(d, PRO('annotate', [d, 'sign'])) },
      { type: 'big', label: 'Đóng dấu thời gian', icon: 'clock', run: need(d, () => U.timeStamp(d)) }
    ] },
    { label: 'Mật khẩu', items: [
      { type: 'big', label: 'Đặt mật khẩu', icon: 'lock', run: need(d, PRO('protectDialog', [d])) },
      { type: 'big', label: 'Gỡ mật khẩu', icon: 'lock-open', run: need(d, PRO('unlockDialog', [d])) }
    ] },
    { label: 'Che thông tin', items: [{ type: 'big', label: 'Xoá trắng', icon: 'eraser', run: need(d, PRO('annotate', [d, 'whiteout'])) }] },
    { label: 'Tài liệu', items: [
      { type: 'big', label: 'Thuộc tính', icon: 'info', run: need(d, () => U.propertiesDialog(d)) },
      { type: 'big', label: 'PDF/A lưu trữ', icon: 'badge-check', run: need(d, () => U.toPdfA(d)) }
    ] }
  ] };
}
function shareTab(d) {
  return { id: 'share', label: 'Chia sẻ', en: 'Share', groups: [
    { label: 'Gửi', items: [
      { type: 'big', label: 'Email', icon: 'mail', run: need(d, () => U.shareMenu(d)[0].run()) },
      { type: 'big', label: 'Đám mây', icon: 'cloud-upload', run: need(d, () => U.shareMenu(d)[1].run()) },
      { type: 'big', label: 'Sao chép văn bản', icon: 'copy', run: need(d, () => U.shareMenu(d)[2].run()) }
    ] }
  ] };
}
function homeTab(d) {
  return { id: 'home', label: 'Trang đầu', en: 'Home', groups: [
    { label: 'Tệp', items: [
      { type: 'big', label: 'Mở', icon: 'folder-open', sc: 'Ctrl+O', run: () => d.shell.openDialog('pdf') },
      { type: 'drop', big: true, label: 'Lưu', icon: 'save', menu: () => U.saveMenu(d) },
      { type: 'big', label: 'In', icon: 'printer', sc: 'Ctrl+P', run: need(d, () => d.print()) }
    ] },
    { label: 'Công cụ', items: [
      { type: 'big', label: 'Tay', icon: 'hand', active: () => d.toolMode === 'hand', run: () => U.setTool(d, 'hand') },
      { type: 'big', label: 'Chọn', icon: 'text-cursor', active: () => d.toolMode === 'select', run: () => U.setTool(d, 'select') },
      { type: 'big', label: 'Ảnh chụp nhanh', icon: 'camera', active: () => d.toolMode === 'snap', run: need(d, () => U.setTool(d, 'snap')) },
      { type: 'big', label: 'Sao chép chữ', icon: 'clipboard', run: need(d, async () => { const t = window.getSelection().toString() || await T.extractText(d); try { await navigator.clipboard.writeText(t); toast('Đã sao chép ' + t.length + ' ký tự'); } catch (e) { toast('Dùng PDF → TXT để lưu văn bản'); } }) }
    ] },
    { label: 'Thu phóng', items: [
      { type: 'drop', big: true, label: 'Thu phóng', icon: 'zoom-in', menu: () => U.zoomMenu(d) },
      { type: 'drop', big: true, label: 'Vừa trang', icon: 'scan', menu: () => [{ label: 'Vừa trang', run: () => d.setFit('page') }, { label: 'Vừa chiều rộng', run: () => d.setFit('width') }, { label: 'Vừa vùng có nội dung', run: () => d.setFit('visible') }, { label: 'Kích thước thực', run: () => d.shell.setZoom(100) }] },
      { type: 'col', items: [
        { type: 'btn', label: 'Phóng to', icon: 'zoom-in', run: () => d.shell.setZoom((d.zoom || 100) + 25) },
        { type: 'btn', label: 'Thu nhỏ', icon: 'zoom-out', run: () => d.shell.setZoom((d.zoom || 100) - 25) },
        { type: 'btn', label: 'Thu phóng vùng', icon: 'scan-search', active: () => d.toolMode === 'zoom', run: need(d, () => U.setTool(d, 'zoom')) }] }
    ] },
    { label: 'Xem', items: [
      { type: 'big', label: 'Dàn lại chữ', icon: 'wrap-text', active: () => !!d.reflow, run: need(d, () => U.reflowView(d)) },
      { type: 'drop', big: true, label: 'Xoay khung nhìn', icon: 'rotate-cw', menu: () => [{ label: 'Xoay phải', icon: 'rotate-cw', run: () => U.rotateView(d, 90) }, { label: 'Xoay trái', icon: 'rotate-ccw', run: () => U.rotateView(d, -90) }] }
    ] },
    { label: 'Chú thích', items: [
      { type: 'big', label: 'Máy đánh chữ', icon: 'type', run: need(d, PRO('annotate', [d, 'text'])) },
      { type: 'big', label: 'Tô sáng', icon: 'highlighter', run: need(d, PRO('annotate', [d, 'highlight'])) },
      { type: 'big', label: 'Điền & Ký', icon: 'signature', run: need(d, PRO('annotate', [d, 'sign'])) }
    ] },
    { label: 'Văn bản', items: [
      { type: 'big', label: 'OCR nhận dạng chữ', icon: 'scan-text', run: () => T.ocrDialog(d) },
      { type: 'big', label: 'Tìm kiếm', icon: 'search', sc: 'Ctrl+F', run: need(d, () => U.togglePanel(d, 'search')) }
    ] },
    { label: 'A.I', items: [{ type: 'big', label: 'Trợ lý A.I', icon: 'sparkles', run: need(d, async () => { const t = await T.extractText(d); import('../ai/ai.js').then(m => m.openAi(d.shell, d, 'Tóm tắt nội dung sau:\n' + t.slice(0, 6000))); }) }] }
  ] };
}
function pagesTab(d) {
  return { id: 'pages', label: 'Sắp xếp trang', en: 'Organize', groups: [
    { label: 'Xoay', items: [
      { type: 'big', label: 'Xoay phải', icon: 'rotate-cw', run: need(d, () => T.rotatePages(d, sel(d), 90)) },
      { type: 'big', label: 'Xoay trái', icon: 'rotate-ccw', run: need(d, () => T.rotatePages(d, sel(d), 270)) },
      { type: 'big', label: 'Xoay tất cả', icon: 'refresh-cw', run: need(d, () => T.rotatePages(d, d.pageInfo.map((_, i) => i), 90)) }
    ] },
    { label: 'Trang', items: [
      { type: 'big', label: 'Xoá trang', icon: 'trash-2', run: need(d, () => T.deletePages(d, sel(d))) },
      { type: 'big', label: 'Lên trên', icon: 'arrow-up', run: need(d, () => { const i = d.cur || 0; if (i > 0) T.movePage(d, i, i - 1); }) },
      { type: 'big', label: 'Xuống dưới', icon: 'arrow-down', run: need(d, () => { const i = d.cur || 0; if (i < d.pdf.numPages - 1) T.movePage(d, i, i + 1); }) }
    ] },
    { label: 'Ghép & Tách', items: [
      { type: 'big', label: 'Chèn PDF khác', icon: 'file-plus', run: () => T.mergeDialog(d) },
      { type: 'big', label: 'Tách / Trích trang', icon: 'scissors', run: () => T.splitDialog(d) }
    ] },
    { label: 'Trang (v0.0.8)', items: [
      { type: 'big', label: 'Chèn trang trống', icon: 'file-plus-2', run: need(d, () => U.insertBlank(d)) },
      { type: 'big', label: 'Cắt cúp lề', icon: 'crop', run: need(d, () => U.cropDialog(d)) },
      { type: 'big', label: 'Đảo thứ tự', icon: 'arrow-down-up', run: need(d, () => U.reverseOrder(d)) },
      { type: 'big', label: 'Sửa tệp hỏng', icon: 'wrench', run: need(d, () => U.repairPdf(d)) }
    ] },
    { label: 'Hoàn tác', items: [{ type: 'col', items: [
      { type: 'btn', label: 'Hoàn tác', icon: 'undo-2', sc: 'Ctrl+Z', disabled: () => !d.canUndo(), run: () => d.undo() },
      { type: 'btn', label: 'Làm lại', icon: 'redo-2', sc: 'Ctrl+Y', disabled: () => !d.canRedo(), run: () => d.redo() }
    ] }] }
  ] };
}
function convertTab(d) {
  return { id: 'convert', label: 'Chuyển đổi', en: 'Convert', groups: [
    { label: 'Từ PDF', items: [
      { type: 'big', label: 'PDF → Word', icon: 'file-type-2', run: () => T.exportAs(d, 'docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Word') },
      { type: 'big', label: 'PDF → TXT', icon: 'file-text', run: () => T.exportAs(d, 'txt', 'text/plain', 'TXT') },
      { type: 'big', label: 'PDF → PowerPoint', icon: 'presentation', run: need(d, () => U.toPptx(d)) },
      { type: 'big', label: 'PDF → Excel', icon: 'sheet', run: need(d, () => U.toXlsx(d)) },
      { type: 'big', label: 'PDF → PDF/A', icon: 'badge-check', run: need(d, () => U.toPdfA(d)) },
      { type: 'big', label: 'PDF → HTML', icon: 'globe', run: () => T.exportAs(d, 'html', 'text/html', 'HTML') },
      { type: 'drop', big: true, label: 'PDF → Ảnh', icon: 'image', menu: () => [{ label: 'PNG (chất lượng cao)', run: () => T.pdfToImages(d, 'png') }, { label: 'JPG (nhẹ)', run: () => T.pdfToImages(d, 'jpg') }] }
    ] },
    { label: 'Sang PDF', items: [
      { type: 'big', label: 'Ảnh → PDF', icon: 'image-plus', run: () => T.imagesToPdfDialog(d) },
      { type: 'big', label: 'Làm phẳng PDF', icon: 'layers', run: () => T.exportAs(d, 'pdf-flat', 'application/pdf', 'PDF phẳng') }
    ] }
  ] };
}
function toolsTab(d) {
  return { id: 'tools', label: 'Công cụ', en: 'Tools', groups: [
    { label: 'Nhận dạng', items: [{ type: 'big', label: 'OCR (Việt + Anh)', icon: 'scan-text', run: () => T.ocrDialog(d) }] },
    { label: 'Trích xuất', items: [
      { type: 'big', label: 'Trích ảnh', icon: 'images', run: () => T.extractImages(d) },
      { type: 'big', label: 'Trích văn bản', icon: 'text-select', run: () => T.exportAs(d, 'txt', 'text/plain', 'TXT') }
    ] },
    { label: 'Đóng dấu', items: [
      { type: 'big', label: 'Hình mờ', icon: 'droplets', run: () => T.watermarkDialog(d) },
      { type: 'big', label: 'Đánh số trang', icon: 'list-ordered', run: () => T.pageNumbersDialog(d) }
    ] },
    { label: 'Bảng Công cụ', items: [{ type: 'big', label: 'Tất cả công cụ', icon: 'wrench', active: () => d.el.classList.contains('pdf-tools-on'), run: () => U.toggleTools(d) }] },
    { label: 'Tài liệu', items: [{ type: 'big', label: 'Thuộc tính', icon: 'info', run: need(d, () => U.propertiesDialog(d)) }, { type: 'big', label: 'Nén PDF', icon: 'archive', run: need(d, PRO('compressDialog', [d])) }] },
    { label: 'A.I', items: [{ type: 'big', label: 'Tóm tắt (A.I)', icon: 'sparkles', run: need(d, async () => { const t = await T.extractText(d); import('../ai/ai.js').then(m => m.openAi(d.shell, d, 'Tóm tắt nội dung sau:\n' + t.slice(0, 6000))); }) }] }
  ] };
}
function viewTab(d) {
  const L = (l, label, icon) => ({ type: 'btn', label, icon, active: () => d.layout === l && d.fitMode !== 'grid', run: () => { if (d.fitMode === 'grid') d.setFit('width'); U.setLayout(d, l); } });
  return { id: 'view', label: 'Xem', en: 'View', groups: [
    { label: 'Chế độ đọc', items: [
      { type: 'big', label: 'Chế độ đọc', icon: 'book-open', active: () => d.el.classList.contains('pdf-readmode'), run: () => { const on = !d.el.classList.contains('pdf-readmode'); d.el.classList.toggle('pdf-readmode', on); if (on) { d.el.classList.add('pdf-nothumbs'); d.el.classList.remove('pdf-tools-on'); d.setFit('page'); toast('Chế độ đọc: bấm lại nút hoặc Esc để thoát'); } else d.relayout(); } },
      { type: 'big', label: 'Dàn lại chữ', icon: 'wrap-text', active: () => !!d.reflow, run: need(d, () => U.reflowView(d)) },
      { type: 'big', label: 'Chế độ đêm', icon: 'moon', active: () => d.el.classList.contains('pdf-dark'), run: () => { d.el.classList.toggle('pdf-dark'); d.refreshUi(); } },
      { type: 'big', label: 'Đảo thứ tự xem', icon: 'arrow-down-up', active: () => !!d.reverse, run: need(d, () => U.toggleReverse(d)) }
    ] },
    { label: 'Bố cục trang', items: [
      { type: 'col', items: [L('single', 'Một trang', 'file'), L('cont', 'Liên tục', 'scroll-text'), L('twoCont', 'Hai trang', 'book-open')] },
      { type: 'col', items: [L('twoCover', 'Hai trang có bìa', 'book-open-text'), { type: 'btn', label: 'Lưới trang', icon: 'layout-grid', active: () => d.fitMode === 'grid', run: () => d.setFit(d.fitMode === 'grid' ? 'width' : 'grid') }] }
    ] },
    { label: 'Xoay khung nhìn', items: [
      { type: 'col', items: [{ type: 'btn', label: 'Xoay phải', icon: 'rotate-cw', run: need(d, () => U.rotateView(d, 90)) }, { type: 'btn', label: 'Xoay trái', icon: 'rotate-ccw', run: need(d, () => U.rotateView(d, -90)) }] }
    ] },
    { label: 'Trợ giúp đọc', items: [
      { type: 'big', label: 'Tự cuộn', icon: 'arrow-down-to-line', active: () => !!d._auto, run: need(d, () => d._auto ? U.stopAuto(d) : U.autoScroll(d)) },
      { type: 'big', label: 'Đọc to', icon: 'volume-2', active: () => !!d._speaking, run: need(d, () => U.readAloud(d)) },
      { type: 'big', label: 'Đếm từ', icon: 'list-ordered', run: need(d, () => U.wordCount(d)) }
    ] },
    { label: 'Bảng điều hướng', items: [{ type: 'col', items: [
      { type: 'btn', label: 'Trang', icon: 'layout-grid', run: () => U.togglePanel(d, 'pages') },
      { type: 'btn', label: 'Dấu trang', icon: 'bookmark', run: () => U.togglePanel(d, 'bookmarks') },
      { type: 'btn', label: 'Thông tin', icon: 'info', run: () => U.togglePanel(d, 'info') }] },
      { type: 'big', label: 'Bảng Công cụ', icon: 'wrench', active: () => d.el.classList.contains('pdf-tools-on'), run: () => U.toggleTools(d) }
    ] }
  ] };
}
