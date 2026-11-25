import { breaksMenu, printAreaMenu } from './pages.js';
import { h, ic, $ } from '../core/util.js';
import { toast, showMenu, colorPicker, dialog, promptDlg, alertDlg } from '../core/ui.js';
import { settings } from '../core/store.js';
import * as X from './cmds.js';
import { FONTS, SIZES, tableDesignTab, cellStylesMenu, cfMenu, TABLE_STYLES, tableStylePreview, THEMES, applyTheme, setPage, fnCategoryMenu } from './features.js';
import { NUMBER_FORMATS, format } from './numfmt.js';
import { CHART_TYPES } from './charts.js';
import { SHAPES, shapeIcon } from '../core/shapes.js';
import { drawTab } from '../core/ink.js';
import { helpTab } from '../core/help.js';
import { symbolDialog, qrDialog, iconsDialog } from '../core/insert.js';
import { MAX_ROWS, MAX_COLS, rangeAddr } from './model.js';

const st = d => X.curStyle(d);
const fullWB = d => d;

export function ribbonDef(doc) {
  const d = doc;
  const tabs = [
    homeTab(d), insertTab(d), drawTab(d, d.objects.ink), pageLayoutTab(d), formulasTab(d), dataTab(d), reviewTab(d), viewTab(d), automateTab(d), developerTab(d), helpTab(d, 'excel')
  ];
  const s = d.wb.cur;
  if (s.tables.some(t => s.sel.r >= t.r1 && s.sel.r <= t.r2 && s.sel.c >= t.c1 && s.sel.c <= t.c2)) tabs.push(tableDesignTab(d));
  const o = d.objects.selected;
  if (o && o.type === 'chart') tabs.push(chartDesignTab(d, o));
  if (o && (o.type === 'image')) tabs.push(pictureTab(d, o));
  if (o && (o.type === 'shape' || o.type === 'textbox')) tabs.push(shapeTab(d, o));
  return { tabs };
}

function homeTab(d) {
  return { id: 'home', label: 'Trang đầu', en: 'Home', groups: [
    { label: 'Bảng tạm', items: [
      { type: 'split', big: true, label: 'Dán', icon: 'clipboard-paste', title: 'Dán (Ctrl+V)', desc: 'Dán nội dung từ bộ nhớ tạm. Bấm mũi tên để chọn: chỉ giá trị, công thức, định dạng, hoán đổi…', run: () => X.paste(d, 'all'), menu: () => X.pasteMenu(d) },
      { type: 'col', items: [
        { type: 'btn', label: 'Cắt', icon: 'scissors', sc: 'Ctrl+X', run: () => X.copy(d, true) },
        { type: 'btn', label: 'Sao chép', icon: 'copy', sc: 'Ctrl+C', run: () => X.copy(d, false) },
        { type: 'btn', label: 'Sao định dạng', icon: 'paintbrush', title: 'Format Painter — bấm một lần để áp dụng 1 lần, bấm đúp để khóa', active: () => !!d.painter, run: () => formatPainter(d) }
      ] }
    ] },
    { label: 'Phông chữ', launcher: () => X.formatCellsDialog(d, 'font'), items: [
      { type: 'col', items: [
        { type: 'row', items: [
          { type: 'combo', width: 128, title: 'Phông chữ', fontBox: true, get: () => st(d).font || settings().excelFont || 'Calibri', set: (a, v) => X.setStyle(d, { font: v }, 'Phông chữ'), options: () => FONTS, optStyle: v => 'font-family:' + v },
          { type: 'combo', width: 40, title: 'Cỡ chữ', get: () => String(st(d).sz || settings().excelSize || 11), set: (a, v) => { const n = parseFloat(v); if (n > 0 && n < 410) X.setStyle(d, { sz: n }, 'Cỡ chữ'); }, options: () => SIZES.map(String) },
          { type: 'btn', icon: 'a-arrow-up', title: 'Tăng cỡ chữ (Ctrl+Shift+>)', run: () => X.fontSizeStep(d, 1) },
          { type: 'btn', icon: 'a-arrow-down', title: 'Giảm cỡ chữ (Ctrl+Shift+<)', run: () => X.fontSizeStep(d, -1) }
        ] },
        { type: 'row', items: [
          { type: 'btn', icon: 'bold', title: 'In đậm', sc: 'Ctrl+B', active: () => st(d).b, run: () => X.toggleStyle(d, 'b') },
          { type: 'btn', icon: 'italic', title: 'In nghiêng', sc: 'Ctrl+I', active: () => st(d).i, run: () => X.toggleStyle(d, 'i') },
          { type: 'split', icon: 'underline', title: 'Gạch chân', sc: 'Ctrl+U', active: () => st(d).u, run: () => X.toggleStyle(d, 'u', 'single'), menu: () => [{ label: 'Gạch chân đơn', run: () => X.setStyle(d, { u: 'single' }) }, { label: 'Gạch chân kép', run: () => X.setStyle(d, { u: 'double' }) }] },
          { type: 'btn', icon: 'strikethrough', title: 'Gạch ngang', sc: 'Ctrl+5', active: () => st(d).s, run: () => X.toggleStyle(d, 's') },
          { type: 'split', icon: 'grid-2x2', title: 'Viền', run: () => X.applyBorder(d, d.lastBorder || 'bottom'), menu: () => X.borderMenu(d) },
          { type: 'color', icon: 'paint-bucket', title: 'Màu tô', colorBar: () => d.lastFill || '#ffff00', run: () => X.setStyle(d, { fill: d.lastFill || '#ffff00' }, 'Màu tô'), pick: (a, c) => { d.lastFill = c; X.setStyle(d, { fill: c === 'transparent' ? null : c }, 'Màu tô'); }, colorOpts: { none: 'Không tô màu' } },
          { type: 'color', icon: 'baseline', title: 'Màu phông chữ', colorBar: () => d.lastColor || '#c00000', run: () => X.setStyle(d, { color: d.lastColor || '#c00000' }, 'Màu chữ'), pick: (a, c) => { d.lastColor = c; X.setStyle(d, { color: c }, 'Màu chữ'); }, colorOpts: { auto: 'Tự động' } }
        ] }
      ] }
    ] },
    { label: 'Căn lề', launcher: () => X.formatCellsDialog(d, 'align'), items: [
      { type: 'col', items: [
        { type: 'row', items: [
          { type: 'btn', icon: 'align-vertical-justify-start', title: 'Căn trên', active: () => st(d).va === 'top', run: () => X.setAlign(d, undefined, 'top') },
          { type: 'btn', icon: 'align-vertical-justify-center', title: 'Căn giữa (dọc)', active: () => st(d).va === 'middle', run: () => X.setAlign(d, undefined, 'middle') },
          { type: 'btn', icon: 'align-vertical-justify-end', title: 'Căn dưới', active: () => !st(d).va || st(d).va === 'bottom', run: () => X.setAlign(d, undefined, null) },
          { type: 'drop', icon: 'rotate-ccw', title: 'Hướng chữ', menu: () => [[45, 'Xoay ngược chiều kim đồng hồ'], [-45, 'Xoay theo chiều kim đồng hồ'], [255, 'Văn bản dọc'], [90, 'Xoay lên'], [-90, 'Xoay xuống'], [0, 'Bỏ xoay']].map(x => ({ label: x[1], run: () => X.setStyle(d, { rot: x[0] || null }, 'Hướng chữ') })) },
          { type: 'btn', label: 'Ngắt dòng', icon: 'wrap-text', active: () => st(d).wrap, run: () => X.setStyle(d, { wrap: st(d).wrap ? null : true }, 'Ngắt dòng') }
        ] },
        { type: 'row', items: [
          { type: 'btn', icon: 'align-left', title: 'Căn trái', active: () => st(d).ha === 'left', run: () => X.setAlign(d, st(d).ha === 'left' ? null : 'left') },
          { type: 'btn', icon: 'align-center', title: 'Căn giữa', active: () => st(d).ha === 'center', run: () => X.setAlign(d, st(d).ha === 'center' ? null : 'center') },
          { type: 'btn', icon: 'align-right', title: 'Căn phải', active: () => st(d).ha === 'right', run: () => X.setAlign(d, st(d).ha === 'right' ? null : 'right') },
          { type: 'btn', icon: 'indent-decrease', title: 'Giảm thụt lề', run: () => X.indent(d, -1) },
          { type: 'btn', icon: 'indent-increase', title: 'Tăng thụt lề', run: () => X.indent(d, 1) },
          { type: 'split', label: 'Gộp & Căn giữa', icon: 'table-cells-merge', run: () => X.merge(d, 'center'), menu: () => [{ label: 'Gộp & Căn giữa', icon: 'table-cells-merge', run: () => X.merge(d, 'center') }, { label: 'Gộp ngang', icon: 'rows-3', run: () => X.merge(d, 'across') }, { label: 'Gộp ô', icon: 'square', run: () => X.merge(d, 'merge') }, { label: 'Bỏ gộp ô', icon: 'table-cells-split', run: () => X.merge(d, 'unmerge') }] }
        ] }
      ] }
    ] },
    { label: 'Số', launcher: () => X.formatCellsDialog(d, 'number'), items: [
      { type: 'col', items: [
        { type: 'row', items: [{ type: 'combo', width: 128, title: 'Định dạng số', get: () => { const nf = st(d).nf; const m = NUMBER_FORMATS.find(x => x[0] === nf); return m ? m[1].replace(/\s*\(.*\)/, '') : nf || 'Chung'; }, set: (a, v) => { const m = NUMBER_FORMATS.find(x => x[1] === v || x[1].replace(/\s*\(.*\)/, '') === v); X.setNumFmt(d, m ? m[0] : v); }, options: () => NUMBER_FORMATS.map(x => [x[1], x[1] + '   ' + format(1234.5, x[0]).t]) }] },
        { type: 'row', items: [
          { type: 'split', icon: 'banknote', title: 'Định dạng số kế toán', run: () => X.setNumFmt(d, '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'), menu: () => [['#,##0 "₫"', '₫ Tiếng Việt'], ['"$"#,##0.00', '$ English (US)'], ['#,##0.00 "€"', '€ Euro'], ['"£"#,##0.00', '£ English (UK)'], ['"¥"#,##0', '¥ 日本語']].map(x => ({ label: x[1], run: () => X.setNumFmt(d, x[0]) })).concat([{ sep: true }, { label: 'Định dạng kế toán khác…', run: () => X.formatCellsDialog(d, 'number') }]) },
          { type: 'btn', icon: 'percent', title: 'Kiểu phần trăm (Ctrl+Shift+%)', run: () => X.setNumFmt(d, '0%') },
          { type: 'btn', icon: 'hash', title: 'Kiểu dấu phẩy (phân cách hàng nghìn)', run: () => X.setNumFmt(d, '#,##0.00') },
          { type: 'btn', icon: 'decimals-arrow-right', title: 'Tăng số thập phân', run: () => X.changeDecimals(d, 1) },
          { type: 'btn', icon: 'decimals-arrow-left', title: 'Giảm số thập phân', run: () => X.changeDecimals(d, -1) }
        ] }
      ] }
    ] },
    { label: 'Kiểu', items: [
      { type: 'drop', big: true, label: 'Định dạng có điều kiện', icon: 'table-properties', menu: () => cfMenu(d) },
      { type: 'drop', big: true, label: 'Định dạng thành Bảng', icon: 'table', menu: () => TABLE_STYLES.map(ts => ({ label: ts.name, swatch: ts.head, run: () => X.formatAsTable(d, ts.id) })) },
      { type: 'drop', big: true, label: 'Kiểu ô', icon: 'palette', menu: () => cellStylesMenu(d) }
    ] },
    { label: 'Ô', items: [
      { type: 'drop', big: true, label: 'Chèn', icon: 'square-plus', menu: () => [{ label: 'Chèn ô…', icon: 'plus', run: () => X.insertDialog(d) }, { label: 'Chèn hàng trang tính', icon: 'between-horizontal-start', run: () => X.insertSelRows(d) }, { label: 'Chèn cột trang tính', icon: 'between-vertical-start', run: () => X.insertSelCols(d) }, { label: 'Chèn trang tính', icon: 'file-plus', run: () => X.insertSheet(d) }] },
      { type: 'drop', big: true, label: 'Xóa', icon: 'square-minus', menu: () => [{ label: 'Xóa ô…', icon: 'minus', run: () => X.deleteDialog(d) }, { label: 'Xóa hàng trang tính', icon: 'trash-2', run: () => X.deleteSelRows(d) }, { label: 'Xóa cột trang tính', icon: 'trash-2', run: () => X.deleteSelCols(d) }, { label: 'Xóa trang tính', icon: 'file-x', run: () => X.deleteSheet(d, d.wb.cur) }] },
      { type: 'drop', big: true, label: 'Định dạng', icon: 'settings-2', menu: () => [{ header: 'Kích thước ô' }, { label: 'Chiều cao hàng…', icon: 'move-vertical', run: () => X.rowHeightDialog(d) }, { label: 'Tự động khớp chiều cao hàng', icon: 'arrow-up-down', run: () => X.autoFit(d, 'r') }, { label: 'Độ rộng cột…', icon: 'move-horizontal', run: () => X.colWidthDialog(d) }, { label: 'Tự động khớp độ rộng cột', icon: 'arrow-left-right', run: () => X.autoFit(d, 'c') }, { header: 'Hiển thị' }, { label: 'Ẩn & Bỏ ẩn', icon: 'eye', sub: [{ label: 'Ẩn hàng', run: () => X.hideRows(d, true) }, { label: 'Ẩn cột', run: () => X.hideCols(d, true) }, { label: 'Ẩn trang tính', run: () => { d.wb.cur.hidden = true; d.switchSheet(1); d.renderTabs(); } }, { label: 'Bỏ ẩn hàng', run: () => X.hideRows(d, false) }, { label: 'Bỏ ẩn cột', run: () => X.hideCols(d, false) }, { label: 'Bỏ ẩn trang tính…', run: () => X.unhideSheetDialog(d) }] }, { header: 'Sắp xếp trang tính' }, { label: 'Đổi tên trang tính', icon: 'pencil', run: () => d.renameSheet(d.wb.cur) }, { label: 'Di chuyển hoặc sao chép…', icon: 'copy', run: () => X.moveCopySheetDialog(d, d.wb.cur) }, { label: 'Màu tab', icon: 'palette', run: () => colorPicker(d.tabsBar, c => { d.wb.cur.tabColor = c === 'transparent' ? null : c; d.renderTabs(); d.markDirty(); }, { none: 'Không màu' }) }, { header: 'Bảo vệ' }, { label: d.wb.cur.protected ? 'Bỏ bảo vệ trang tính' : 'Bảo vệ trang tính…', icon: 'lock', run: () => X.toggleProtect(d) }, { label: 'Khóa/mở khóa ô', icon: 'lock-open', run: () => X.setStyle(d, { unlocked: st(d).unlocked ? null : true }, 'Khóa ô') }, { label: 'Định dạng ô…', icon: 'paintbrush', run: () => X.formatCellsDialog(d) }] }
    ] },
    { label: 'Chỉnh sửa', items: [
      { type: 'col', items: [
        { type: 'split', label: 'Tự động tính tổng', icon: 'sigma', run: () => X.autoSum(d, 'SUM'), menu: () => [['SUM', 'Tổng', 'sigma'], ['AVERAGE', 'Trung bình', 'divide'], ['COUNT', 'Đếm số', 'hash'], ['MAX', 'Lớn nhất', 'arrow-up-to-line'], ['MIN', 'Nhỏ nhất', 'arrow-down-to-line']].map(x => ({ label: x[1], icon: x[2], run: () => X.autoSum(d, x[0]) })).concat([{ sep: true }, { label: 'Hàm khác…', icon: 'function-square', run: () => X.insertFunctionDialog(d) }]) },
        { type: 'drop', label: 'Điền', icon: 'arrow-down-to-line', menu: () => [{ label: 'Xuống', icon: 'arrow-down', shortcut: 'Ctrl+D', run: () => X.fillDir(d, 'down') }, { label: 'Phải', icon: 'arrow-right', shortcut: 'Ctrl+R', run: () => X.fillDir(d, 'right') }, { label: 'Lên', icon: 'arrow-up', run: () => X.fillDir(d, 'up') }, { label: 'Trái', icon: 'arrow-left', run: () => X.fillDir(d, 'left') }, { label: 'Chuỗi…', icon: 'list-ordered', run: () => X.seriesDialog(d) }, { label: 'Điền nhanh (Flash Fill)', icon: 'zap', shortcut: 'Ctrl+E', run: () => X.flashFill(d) }] },
        { type: 'drop', label: 'Xóa', icon: 'eraser', menu: () => [{ label: 'Xóa tất cả', icon: 'eraser', run: () => X.clear(d, 'all') }, { label: 'Xóa định dạng', icon: 'remove-formatting', run: () => X.clear(d, 'formats') }, { label: 'Xóa nội dung', icon: 'delete', run: () => X.clear(d, 'contents') }, { label: 'Xóa chú thích & ghi chú', icon: 'message-square-x', run: () => X.clear(d, 'comments') }, { label: 'Xóa liên kết', icon: 'unlink', run: () => X.clear(d, 'links') }] }
      ] },
      { type: 'drop', big: true, label: 'Sắp xếp & Lọc', icon: 'arrow-down-wide-narrow', menu: () => [{ label: 'Sắp xếp A đến Z', icon: 'arrow-down-a-z', run: () => X.quickSort(d, false) }, { label: 'Sắp xếp Z đến A', icon: 'arrow-down-z-a', run: () => X.quickSort(d, true) }, { label: 'Sắp xếp tùy chỉnh…', icon: 'list-filter', run: () => X.sortDialog(d) }, { sep: true }, { label: 'Lọc', icon: 'filter', checked: !!d.wb.cur.filter || undefined, shortcut: 'Ctrl+Shift+L', run: () => X.toggleFilter(d) }, { label: 'Xóa bộ lọc', icon: 'filter-x', run: () => X.clearFilters(d) }, { label: 'Áp dụng lại', icon: 'refresh-cw', run: () => X.reapplyFilter(d) }] },
      { type: 'drop', big: true, label: 'Tìm & Chọn', icon: 'search', menu: () => [{ label: 'Tìm…', icon: 'search', shortcut: 'Ctrl+F', run: () => X.findDialog(d) }, { label: 'Thay thế…', icon: 'replace', shortcut: 'Ctrl+H', run: () => X.findDialog(d, true) }, { label: 'Đi tới…', icon: 'navigation', shortcut: 'Ctrl+G', run: () => X.gotoDialog(d) }, { label: 'Đi tới Đặc biệt…', icon: 'scan', run: () => X.gotoSpecialDialog(d) }, { sep: true }, { label: 'Công thức', run: () => X.selectSpecial(d, 'formulas') }, { label: 'Chú thích', run: () => X.selectSpecial(d, 'comments') }, { label: 'Định dạng có điều kiện', run: () => X.selectSpecial(d, 'cf') }, { label: 'Hằng số', run: () => X.selectSpecial(d, 'constants') }, { label: 'Xác thực dữ liệu', run: () => X.selectSpecial(d, 'dv') }, { label: 'Chọn đối tượng', icon: 'mouse-pointer-2', run: () => X.selectSpecial(d, 'objects') }] }
    ] },
    { label: 'A.I', items: [
      { type: 'big', label: 'Copilot (A.I)', icon: 'sparkles', desc: 'Trợ lý A.I: phân tích dữ liệu, giải thích công thức, tạo công thức từ mô tả. Chạy ngoại tuyến hoặc với khóa API của bạn.', run: () => import('../ai/ai.js').then(m => m.openAi(d.shell, d)) },
      { type: 'big', label: 'Phân tích dữ liệu', icon: 'chart-no-axes-combined', run: () => X.quickAnalysis(d) }
    ] }
  ] };
}
function formatPainter(d) {
  const s = d.wb.cur; const src = Object.assign({}, X.curStyle(d));
  if (d.painter) { d.painter = null; d.refreshUi(); return; }
  d.painter = { style: src };
  toast('Sao định dạng: chọn ô đích để áp dụng định dạng');
  const once = () => { setTimeout(() => { if (!d.painter) return; const st2 = d.painter.style; d.painter = null; X.setStyle(d, Object.assign({ b: null, i: null, u: null, s: null, color: null, fill: null, nf: null, ha: null, va: null, bd: null, font: null, sz: null, wrap: null }, st2), 'Sao định dạng'); d.refreshUi(); d.grid.scroller.removeEventListener('mouseup', once); }, 30); };
  d.grid.scroller.addEventListener('mouseup', once);
}

function insertTab(d) {
  return { id: 'insert', label: 'Chèn', en: 'Insert', groups: [
    { label: 'Bảng', items: [
      { type: 'big', label: 'PivotTable', icon: 'table-2', desc: 'Tóm tắt dữ liệu bằng bảng tổng hợp (PivotTable).', run: () => import('./features.js').then(m => m.pivotDialog(d)) },
      { type: 'big', label: 'PivotTable đề xuất', icon: 'sparkles', run: () => import('./features.js').then(m => m.pivotDialog(d)) },
      { type: 'big', label: 'Bảng', icon: 'table', sc: 'Ctrl+T', run: () => X.formatAsTable(d, 'Medium2') }
    ] },
    { label: 'Hình minh họa', items: [
      { type: 'big', label: 'Ảnh', icon: 'image', desc: 'Chèn ảnh từ thiết bị (PNG, JPG, GIF, SVG).', run: () => d.objects.insertImage() },
      { type: 'drop', big: true, label: 'Hình dạng', icon: 'shapes', menuWidth: 300, menu: () => SHAPES.map(sh => ({ label: sh[1], el: null, icon: null, run: () => d.objects.addShape(sh[0]) })) },
      { type: 'big', label: 'Biểu tượng', icon: 'smile', run: () => iconsDialog(url => d.objects.insertImage(url)) },
      { type: 'big', label: 'Mã QR', icon: 'qr-code', desc: 'Tạo mã QR từ văn bản/liên kết (QR Code generator – Nayuki/Arase).', run: () => qrDialog(url => d.objects.insertImage(url), (() => { const s = d.wb.cur; const v = s.value(s.sel.r, s.sel.c); return v == null ? '' : String(v); })()) }
    ] },
    { label: 'Biểu đồ', launcher: () => import('./features.js').then(m => m.recommendedCharts(d)), items: [
      { type: 'big', label: 'Biểu đồ đề xuất', icon: 'chart-no-axes-combined', run: () => import('./features.js').then(m => m.recommendedCharts(d)) },
      { type: 'col', items: [
        { type: 'row', items: [
          { type: 'drop', icon: 'chart-column', title: 'Cột / Thanh', menu: () => CHART_TYPES.filter(t => /column|bar/.test(t[0])).map(t => ({ label: t[1], icon: t[2], run: () => X.insertChart(d, t[0]) })) },
          { type: 'drop', icon: 'chart-line', title: 'Đường / Vùng', menu: () => CHART_TYPES.filter(t => /line|area/.test(t[0])).map(t => ({ label: t[1], icon: t[2], run: () => X.insertChart(d, t[0]) })) },
          { type: 'drop', icon: 'chart-pie', title: 'Tròn / Vành khuyên', menu: () => CHART_TYPES.filter(t => /pie|doughnut/.test(t[0])).map(t => ({ label: t[1], icon: t[2], run: () => X.insertChart(d, t[0]) })) }
        ] },
        { type: 'row', items: [
          { type: 'drop', icon: 'chart-scatter', title: 'Phân tán', menu: () => [{ label: 'Phân tán (XY)', icon: 'chart-scatter', run: () => X.insertChart(d, 'scatter') }] },
          { type: 'btn', icon: 'chart-column-stacked', title: 'Cột chồng', run: () => X.insertChart(d, 'column-stacked') },
          { type: 'btn', icon: 'chart-candlestick', title: 'Kết hợp cột + đường', run: () => X.insertChart(d, 'combo') }
        ] }
      ] },
      { type: 'big', label: 'Trang biểu đồ', icon: 'file-chart-column', desc: 'Tạo biểu đồ trên trang tính riêng (F11).', run: () => X.insertChart(d, 'column', true) }
    ] },
    { label: 'Biểu đồ thu nhỏ', items: [
      { type: 'col', items: [{ type: 'btn', label: 'Đường', icon: 'chart-spline', run: () => import('./features.js').then(m => m.sparklines(d, 'line')) }, { type: 'btn', label: 'Cột', icon: 'chart-column-big', run: () => import('./features.js').then(m => m.sparklines(d, 'column')) }, { type: 'btn', label: 'Thắng/Thua', icon: 'chart-column-decreasing', run: () => import('./features.js').then(m => m.sparklines(d, 'winloss')) }] }
    ] },
    { label: 'Liên kết', items: [{ type: 'big', label: 'Liên kết', icon: 'link', sc: 'Ctrl+K', run: () => X.linkDialog(d) }] },
    { label: 'Chú thích', items: [{ type: 'big', label: 'Chú thích', icon: 'message-square-plus', run: () => X.commentDialog(d) }] },
    { label: 'Văn bản', items: [
      { type: 'big', label: 'Hộp văn bản', icon: 'type', run: () => d.objects.addTextbox() },
      { type: 'big', label: 'Đầu trang & Chân trang', icon: 'panel-top', run: () => import('./features.js').then(m => m.pageSetupDialog(d)) },
      { type: 'big', label: 'WordArt', icon: 'whole-word', run: () => d.objects.addTextbox(true) }
    ] },
    { label: 'Ký hiệu', items: [
      { type: 'big', label: 'Phương trình (MathType)', icon: 'sigma-square', desc: 'Mở trình soạn phương trình MathType: chèn công thức toán học.', run: () => import('../math/math.js').then(m => m.equationEditor('').then(lx => { if (lx) d.objects.addEquation(lx); })) },
      { type: 'big', label: 'Ký hiệu', icon: 'omega', run: () => symbolDialog(ch => { const s = d.wb.cur; const v = s.value(s.sel.r, s.sel.c); d.setInput(s.sel.r, s.sel.c, (v == null ? '' : String(v)) + ch); }) }
    ] }
  ] };
}

function pageLayoutTab(d) {
  const ps = () => d.wb.cur.pageSetup;
  return { id: 'pagelayout', label: 'Bố trí Trang', en: 'Page Layout', groups: [
    { label: 'Chủ đề', items: [
      { type: 'drop', big: true, label: 'Chủ đề', icon: 'swatch-book', menu: () => THEMES.map(t => ({ label: t.name, swatch: t.colors[0], checked: d.wb.props.theme === t.name || undefined, run: () => applyTheme(d, t) })) },
      { type: 'col', items: [
        { type: 'drop', label: 'Màu', icon: 'palette', menu: () => THEMES.map(t => ({ label: t.name, swatch: t.colors[0], run: () => applyTheme(d, Object.assign({}, t, { font: d.wb.props.themeFont || 'Calibri' })) })) },
        { type: 'drop', label: 'Phông', icon: 'type', menu: () => FONTS.slice(0, 12).map(f => ({ label: f, style: 'font-family:' + f, run: () => { import('../core/store.js').then(m => m.saveSettings({ excelFont: f })); d.grid.measureCache = {}; d.grid.redraw(); } })) }
      ] }
    ] },
    { label: 'Thiết lập Trang', launcher: () => import('./features.js').then(m => m.pageSetupDialog(d)), items: [
      { type: 'drop', big: true, label: 'Lề', icon: 'square-dashed', menu: () => [['normal', 'Bình thường'], ['wide', 'Rộng'], ['narrow', 'Hẹp']].map(x => ({ label: x[1], checked: ps().margins === x[0] || undefined, run: () => setPage(d, { margins: x[0] }) })) },
      { type: 'drop', big: true, label: 'Hướng', icon: 'rectangle-vertical', menu: () => [['portrait', 'Dọc'], ['landscape', 'Ngang']].map(x => ({ label: x[1], checked: ps().orientation === x[0] || undefined, run: () => setPage(d, { orientation: x[0] }) })) },
      { type: 'drop', big: true, label: 'Khổ giấy', icon: 'file', menu: () => ['A4', 'A3', 'A5', 'Letter', 'Legal'].map(x => ({ label: x, checked: ps().size === x || undefined, run: () => setPage(d, { size: x }) })) },
      { type: 'drop', big: true, label: 'Vùng in', icon: 'printer', menu: () => printAreaMenu(d) },
      { type: 'drop', big: true, label: 'Ngắt trang', icon: 'separator-horizontal', menu: () => breaksMenu(d) },
      { type: 'big', label: 'Nền', icon: 'image', run: () => toast('Nền trang tính: dùng Chèn > Ảnh rồi Đưa ra sau') },
      { type: 'big', label: 'In tiêu đề', icon: 'heading', run: async () => { const v = await promptDlg('Các hàng lặp lại ở đầu mỗi trang (vd: 1:2):', ps().printTitles || '', 'In tiêu đề'); if (v != null) setPage(d, { printTitles: v }); } }
    ] },
    { label: 'Co giãn cho Vừa', items: [{ type: 'col', items: [
      { type: 'drop', label: 'Chiều rộng: ' + (ps().fit === 'width' || ps().fit === 'page' ? '1 trang' : 'Tự động'), icon: 'move-horizontal', menu: () => [{ label: 'Tự động', run: () => setPage(d, { fit: 'none' }) }, { label: '1 trang', run: () => setPage(d, { fit: 'width' }) }] },
      { type: 'drop', label: 'Chiều cao: ' + (ps().fit === 'page' ? '1 trang' : 'Tự động'), icon: 'move-vertical', menu: () => [{ label: 'Tự động', run: () => setPage(d, { fit: ps().fit === 'page' ? 'width' : ps().fit }) }, { label: '1 trang', run: () => setPage(d, { fit: 'page' }) }] }
    ] }] },
    { label: 'Tùy chọn Trang tính', items: [
      { type: 'col', items: [{ type: 'btn', label: 'Đường lưới: Xem', icon: 'grid-3x3', active: () => d.wb.cur.showGrid, run: () => { d.wb.cur.showGrid = !d.wb.cur.showGrid; d.grid.redraw(); d.markDirty(); } }, { type: 'btn', label: 'Đường lưới: In', icon: 'printer', active: () => ps().gridlines, run: () => setPage(d, { gridlines: !ps().gridlines }) }] },
      { type: 'col', items: [{ type: 'btn', label: 'Tiêu đề: Xem', icon: 'heading', active: () => d.wb.cur.showHead, run: () => { d.wb.cur.showHead = !d.wb.cur.showHead; d.grid.resize(); d.markDirty(); } }, { type: 'btn', label: 'Tiêu đề: In', icon: 'printer', active: () => ps().headings, run: () => setPage(d, { headings: !ps().headings }) }] }
    ] },
    { label: 'Sắp xếp', items: [
      { type: 'btn', label: 'Đưa lên trước', icon: 'bring-to-front', disabled: () => !d.objects.selected, run: () => d.objects.order(d.objects.selected, 1) },
      { type: 'btn', label: 'Đưa ra sau', icon: 'send-to-back', disabled: () => !d.objects.selected, run: () => d.objects.order(d.objects.selected, -1) },
      { type: 'btn', label: 'Xoay', icon: 'rotate-cw', disabled: () => !d.objects.selected, run: () => d.objects.mutate('Xoay', () => { d.objects.selected.rot = ((d.objects.selected.rot || 0) + 90) % 360; }) }
    ] }
  ] };
}

function formulasTab(d) {
  const cat = (label, icon, c) => ({ type: 'drop', big: true, label, icon, menuWidth: 240, menu: () => fnCategoryMenu(d, c) });
  return { id: 'formulas', label: 'Công thức', en: 'Formulas', groups: [
    { label: 'Thư viện hàm', items: [
      { type: 'big', label: 'Chèn hàm', icon: 'function-square', sc: 'Shift+F3', desc: 'Tìm và chèn trong hơn 500 hàm Excel (toán, thống kê, tài chính, tra cứu, ngày giờ, văn bản, kỹ thuật…).', run: () => X.insertFunctionDialog(d) },
      { type: 'split', big: true, label: 'Tự động tính tổng', icon: 'sigma', run: () => X.autoSum(d, 'SUM'), menu: () => [['SUM', 'Tổng'], ['AVERAGE', 'Trung bình'], ['COUNT', 'Đếm số'], ['MAX', 'Lớn nhất'], ['MIN', 'Nhỏ nhất']].map(x => ({ label: x[1], run: () => X.autoSum(d, x[0]) })) },
      { type: 'drop', big: true, label: 'Dùng gần đây', icon: 'history', menu: () => (settings().xlRecentFns || ['SUM', 'AVERAGE', 'IF', 'VLOOKUP', 'COUNTIF']).map(n => ({ label: n, run: () => d.grid.beginEdit('enter', '=' + n + '(') })) },
      cat('Tài chính', 'landmark', 'Tài chính'), cat('Lô-gic', 'git-branch', 'Lô-gic'), cat('Văn bản', 'type', 'Văn bản'), cat('Ngày & Giờ', 'calendar-clock', 'Ngày & Giờ'), cat('Tra cứu & Tham chiếu', 'search', 'Tra cứu & Tham chiếu'), cat('Toán & Lượng giác', 'pi', 'Toán học & Lượng giác'),
      { type: 'drop', big: true, label: 'Hàm khác', icon: 'ellipsis', menu: () => ['Thống kê', 'Kỹ thuật', 'Khối (Cube)', 'Thông tin', 'Tương thích', 'Web', 'Cơ sở dữ liệu'].map(c => ({ label: c, sub: () => fnCategoryMenu(d, c) })) }
    ] },
    { label: 'Tên đã Xác định', items: [
      { type: 'big', label: 'Quản lý Tên', icon: 'tags', run: () => X.nameManager(d) },
      { type: 'col', items: [
        { type: 'btn', label: 'Xác định Tên', icon: 'tag', run: () => X.defineNameDialog(d) },
        { type: 'drop', label: 'Dùng trong Công thức', icon: 'function-square', menu: () => { const n = Object.keys(d.wb.names); return n.length ? n.map(x => ({ label: d.wb.names[x].display || x, run: () => { if (d.grid.editing) { const ed = d.grid.editor; ed.value += x; d.grid.onEditorInput(); } else d.grid.beginEdit('enter', '=' + x); } })) : [{ label: '(chưa có tên)', disabled: true }]; } },
        { type: 'btn', label: 'Tạo từ Vùng chọn', icon: 'table-properties', run: () => X.createNamesFromSelection(d) }
      ] }
    ] },
    { label: 'Kiểm tra Công thức', items: [
      { type: 'col', items: [
        { type: 'btn', label: 'Truy tìm Ô Tiền lệ', icon: 'arrow-right-from-line', run: () => import('./features.js').then(m => m.tracePrecedents(d)) },
        { type: 'btn', label: 'Truy tìm Ô Phụ thuộc', icon: 'arrow-left-from-line', run: () => import('./features.js').then(m => m.traceDependents(d)) },
        { type: 'btn', label: 'Loại bỏ Mũi tên', icon: 'x', run: () => import('./features.js').then(m => m.removeArrows(d)) }
      ] },
      { type: 'col', items: [
        { type: 'btn', label: 'Hiện Công thức', icon: 'function-square', sc: 'Ctrl+`', active: () => d.wb.cur.showFormulas, run: () => X.toggleShowFormulas(d) },
        { type: 'btn', label: 'Kiểm tra Lỗi', icon: 'triangle-alert', run: () => import('./features.js').then(m => m.errorChecking(d)) },
        { type: 'btn', label: 'Đánh giá Công thức', icon: 'list-checks', run: () => import('./features.js').then(m => m.evaluateFormula(d)) }
      ] },
      { type: 'big', label: 'Cửa sổ Theo dõi', icon: 'eye', run: () => import('./features.js').then(m => m.watchWindow(d)) }
    ] },
    { label: 'Tính toán', items: [
      { type: 'drop', big: true, label: 'Tùy chọn Tính toán', icon: 'calculator', menu: () => [{ label: 'Tự động', checked: d.wb.calcMode !== 'manual' || undefined, run: () => { d.wb.calcMode = 'auto'; d.recalc(true); d.grid.redraw(); } }, { label: 'Thủ công', checked: d.wb.calcMode === 'manual' || undefined, run: () => { d.wb.calcMode = 'manual'; toast('Tính toán thủ công: nhấn F9 để tính lại'); } }] },
      { type: 'col', items: [{ type: 'btn', label: 'Tính Ngay', icon: 'calculator', sc: 'F9', run: () => { d.recalc(true); d.grid.redraw(); toast('Đã tính toán lại sổ làm việc'); } }, { type: 'btn', label: 'Tính Trang tính', icon: 'sheet', run: () => { d.recalc(true); d.grid.redraw(); } }] }
    ] }
  ] };
}

function dataTab(d) {
  return { id: 'data', label: 'Dữ liệu', en: 'Data', groups: [
    { label: 'Lấy & Chuyển đổi Dữ liệu', items: [
      { type: 'drop', big: true, label: 'Lấy Dữ liệu', icon: 'database', menu: () => [{ label: 'Từ Văn bản/CSV…', icon: 'file-text', run: () => importData(d, '.csv,.txt,.tsv') }, { label: 'Từ Sổ làm việc Excel…', icon: 'sheet', run: () => importData(d, '.xlsx,.xlsm') }, { label: 'Từ JSON…', icon: 'braces', run: () => importData(d, '.json') }, { label: 'Từ Bảng/Vùng', icon: 'table', run: () => X.formatAsTable(d, 'Medium2') }] },
      { type: 'col', items: [{ type: 'btn', label: 'Từ Văn bản/CSV', icon: 'file-text', run: () => importData(d, '.csv,.txt,.tsv') }, { type: 'btn', label: 'Từ Bảng/Vùng', icon: 'table', run: () => X.formatAsTable(d, 'Medium2') }, { type: 'btn', label: 'Từ Sổ làm việc', icon: 'sheet', run: () => importData(d, '.xlsx') }] }
    ] },
    { label: 'Truy vấn & Kết nối', items: [{ type: 'big', label: 'Làm mới Tất cả', icon: 'refresh-cw', sc: 'Ctrl+Alt+F5', run: () => import('./features.js').then(m => m.refreshAll(d)) }] },
    { label: 'Sắp xếp & Lọc', items: [
      { type: 'col', items: [{ type: 'btn', icon: 'arrow-down-a-z', title: 'Sắp xếp A → Z', run: () => X.quickSort(d, false) }, { type: 'btn', icon: 'arrow-down-z-a', title: 'Sắp xếp Z → A', run: () => X.quickSort(d, true) }] },
      { type: 'big', label: 'Sắp xếp', icon: 'list-filter', run: () => X.sortDialog(d) },
      { type: 'big', label: 'Lọc', icon: 'filter', active: () => !!d.wb.cur.filter, run: () => X.toggleFilter(d) },
      { type: 'col', items: [{ type: 'btn', label: 'Xóa', icon: 'filter-x', disabled: () => !d.wb.cur.filter, run: () => X.clearFilters(d) }, { type: 'btn', label: 'Áp dụng lại', icon: 'refresh-cw', disabled: () => !d.wb.cur.filter, run: () => X.reapplyFilter(d) }] }
    ] },
    { label: 'Công cụ Dữ liệu', items: [
      { type: 'big', label: 'Văn bản thành Cột', icon: 'columns-3', run: () => import('./features.js').then(m => m.textToColumns(d)) },
      { type: 'col', items: [
        { type: 'btn', label: 'Điền Nhanh', icon: 'zap', sc: 'Ctrl+E', run: () => X.flashFill(d) },
        { type: 'btn', label: 'Loại bỏ Trùng lặp', icon: 'copy-minus', run: () => import('./features.js').then(m => m.removeDuplicates(d)) },
        { type: 'split', label: 'Xác thực Dữ liệu', icon: 'shield-check', run: () => import('./features.js').then(m => m.dvDialog(d)), menu: () => [{ label: 'Xác thực Dữ liệu…', run: () => import('./features.js').then(m => m.dvDialog(d)) }, { label: 'Khoanh tròn Dữ liệu Không hợp lệ', run: () => import('./features.js').then(m => m.circleInvalid(d)) }] }
      ] },
      { type: 'col', items: [{ type: 'btn', label: 'Hợp nhất', icon: 'combine', run: () => import('./features.js').then(m => m.consolidateDialog(d)) }] }
    ] },
    { label: 'Dự báo', items: [
      { type: 'drop', big: true, label: 'Phân tích What-If', icon: 'circle-help', menu: () => [
        { label: 'Tìm Mục tiêu (Goal Seek)…', icon: 'target', run: () => import('./features.js').then(m => m.goalSeek(d)) },
        { label: 'Bộ giải Solver (LibreOffice)…', icon: 'calculator', run: () => import('../core/suites.js').then(m => m.openSolverDialog(d)) }
      ] },
      { type: 'big', label: 'Bộ giải Solver', icon: 'calculator', title: 'Giải bài toán tối ưu quy hoạch mục tiêu (LibreOffice)', run: () => import('../core/suites.js').then(m => m.openSolverDialog(d)) },
      { type: 'big', label: 'Trang Dự báo', icon: 'chart-line', run: () => import('./features.js').then(m => m.forecastSheet(d)) }
    ] },
    { label: 'Đại cương', items: [
      { type: 'col', items: [
        { type: 'btn', label: 'Nhóm', icon: 'group', run: () => import('./features.js').then(m => m.group(d, true)) },
        { type: 'btn', label: 'Rã Nhóm', icon: 'ungroup', run: () => import('./features.js').then(m => m.group(d, false)) },
        { type: 'btn', label: 'Tổng phụ', icon: 'sigma', run: () => import('./features.js').then(m => m.subtotalDialog(d)) }
      ] },
      { type: 'col', items: [{ type: 'btn', label: 'Hiện chi tiết', icon: 'plus', run: () => import('./features.js').then(m => m.outlineShow(d, true)) }, { type: 'btn', label: 'Ẩn chi tiết', icon: 'minus', run: () => import('./features.js').then(m => m.outlineShow(d, false)) }] }
    ] }
  ] };
}
async function importData(d, accept) {
  const Host = (await import('../core/host.js')).default;
  const files = await Host.openFiles(accept); if (!files.length) return;
  const f = files[0]; const { readCsv, readXlsx } = await import('./xlsx.js'); const { extOf, utf8Decode, baseName } = await import('../core/util.js');
  const ext = extOf(f.name);
  d.change('Nhập dữ liệu', d.wb.cur, 'book', () => {});
  const u = d.undoStack.pop();
  const before = JSON.stringify((await import('./model.js')).wbToJSON(d.wb));
  if (ext === 'xlsx' || ext === 'xlsm') { const { Workbook } = await import('./model.js'); const tmp = new Workbook(); await readXlsx(tmp, f.bytes); tmp.sheets.forEach(s => { s.wb = d.wb; let n = s.name; while (d.wb.sheet(n)) n += '_'; s.name = n; d.wb.sheets.push(s); }); }
  else if (ext === 'json') { try { const data = JSON.parse(utf8Decode(f.bytes)); const s = d.wb.addSheet(baseName(f.name).slice(0, 28)); const rows = Array.isArray(data) ? data : data.data || [data]; const keys = rows.length && typeof rows[0] === 'object' && !Array.isArray(rows[0]) ? Object.keys(rows[0]) : null; if (keys) { keys.forEach((k, j) => s.setCell(0, j, { v: k })); rows.forEach((r, i) => keys.forEach((k, j) => s.setCell(i + 1, j, { v: r[k] == null ? null : (typeof r[k] === 'object' ? JSON.stringify(r[k]) : r[k]) }))); } else rows.forEach((r, i) => (Array.isArray(r) ? r : [r]).forEach((v, j) => s.setCell(i, j, { v }))); } catch (e) { toast('JSON không hợp lệ', 'err'); return; } }
  else readCsv(d.wb, f.bytes, ext === 'tsv' ? '\t' : null, f.name);
  const { wbFromJSON } = await import('./model.js');
  d.undoStack.push({ label: 'Nhập dữ liệu', undo: () => { const w = wbFromJSON(JSON.parse(before)); d.wb.sheets = w.sheets; d.wb.sheets.forEach(s => { s.wb = d.wb; }); d.wb.active = Math.min(d.wb.active, d.wb.sheets.length - 1); d.renderTabs(); }, redo: () => toast('Hãy nhập lại tệp') });
  d.wb.active = d.wb.sheets.length - 1; d.afterChange(); d.renderTabs(); d.activateSheet(d.wb.cur);
  toast('Đã nhập ' + f.name, 'ok');
}

function reviewTab(d) {
  return { id: 'review', label: 'Xem lại', en: 'Review', groups: [
    { label: 'Soát lỗi', items: [
      { type: 'big', label: 'Chính tả', icon: 'spell-check', sc: 'F7', run: () => import('../ai/ai.js').then(m => m.proofread(d)) },
      { type: 'col', items: [{ type: 'btn', label: 'Từ điển đồng nghĩa', icon: 'book-open', run: () => import('../ai/ai.js').then(m => m.openAi(d.shell, d, 'Gợi ý từ đồng nghĩa cho: ' + String(d.wb.cur.value(d.wb.cur.sel.r, d.wb.cur.sel.c) || ''))) }, { type: 'btn', label: 'Thống kê Sổ làm việc', icon: 'chart-bar', run: () => d.shell.inspect(d) }] }
    ] },
    { label: 'Hiệu suất', items: [{ type: 'big', label: 'Kiểm tra Hiệu suất', icon: 'gauge', run: () => { const t0 = Date.now(); d.recalc(true); const ms = Date.now() - t0; alertDlg('Tính toán lại toàn bộ sổ làm việc mất ' + ms + ' ms.\nSố ô: ' + d.wb.sheets.reduce((n, s) => n + s.cells.size, 0) + '\nKhông phát hiện ô định dạng thừa.', 'Kiểm tra Hiệu suất'); } }] },
    { label: 'Trợ năng', items: [{ type: 'big', label: 'Kiểm tra Trợ năng', icon: 'accessibility', run: () => d.shell.inspect(d) }] },
    { label: 'Ngôn ngữ', items: [{ type: 'big', label: 'Dịch', icon: 'languages', run: () => import('../ai/ai.js').then(m => m.openAi(d.shell, d, 'Dịch sang tiếng Anh: ' + String(d.wb.cur.value(d.wb.cur.sel.r, d.wb.cur.sel.c) || ''))) },
      { type: 'big', label: 'Chuyển mã', icon: 'replace-all', title: 'Chuyển TCVN3 (.VnTime) / VNI → Unicode trong vùng chọn', run: () => import('../core/fonts.js').then(async F => { const mode = await F.convertDialog(); if (!mode) return; const s = d.wb.cur; const rg = d.grid.selRange(); const u = s.used(); let n = 0; d.change('Chuyển mã', s, rg, () => { for (let r = rg.r1; r <= Math.min(rg.r2, u.r2); r++) for (let c = rg.c1; c <= Math.min(rg.c2, u.c2); c++) { const x = s.get(r, c); if (!x || typeof x.v !== 'string' || x.f) continue; const ff = (s.styleAt(r, c) || {}).font || ''; const kind = mode === 'auto' ? (F.detectLegacy(x.v, ff) || 'tcvn3') : mode; const nv = F.convertLegacy(x.v, kind); if (nv !== x.v) { s.setCell(r, c, { v: nv }); n++; } } }); toast('Đã chuyển mã ' + n + ' ô', 'ok'); }) }] },
    { label: 'Chú thích', items: [
      { type: 'big', label: 'Chú thích Mới', icon: 'message-square-plus', run: () => X.commentDialog(d) },
      { type: 'col', items: [{ type: 'btn', label: 'Xóa', icon: 'message-square-x', run: () => X.clear(d, 'comments') }, { type: 'btn', label: 'Trước', icon: 'arrow-left', run: () => X.nextComment(d, -1) }, { type: 'btn', label: 'Tiếp', icon: 'arrow-right', run: () => X.nextComment(d, 1) }] },
      { type: 'big', label: 'Hiện Chú thích', icon: 'messages-square', run: () => X.commentsPane(d) }
    ] },
    { label: 'Ghi chú', items: [{ type: 'drop', big: true, label: 'Ghi chú', icon: 'sticky-note', menu: () => [{ label: 'Ghi chú Mới', run: () => X.noteDialog(d) }, { label: 'Sửa Ghi chú', run: () => X.noteDialog(d) }, { label: 'Hiện tất cả Ghi chú', run: () => X.commentsPane(d) }] }] },
    { label: 'Bảo vệ', items: [
      { type: 'big', label: d.wb.cur.protected ? 'Bỏ bảo vệ Trang tính' : 'Bảo vệ Trang tính', icon: 'lock', active: () => !!d.wb.cur.protected, run: () => X.toggleProtect(d) },
      { type: 'big', label: 'Bảo vệ Sổ làm việc', icon: 'shield', active: () => !!d.wb.props.protected, run: () => { d.wb.props.protected = !d.wb.props.protected; toast(d.wb.props.protected ? 'Đã khóa cấu trúc sổ làm việc (không thêm/xóa trang tính)' : 'Đã bỏ khóa cấu trúc'); d.markDirty(); } },
      { type: 'big', label: 'Cho phép Sửa Vùng', icon: 'lock-open', run: () => { X.setStyle(d, { unlocked: true }, 'Cho phép sửa vùng'); toast('Vùng chọn được phép sửa khi trang tính bị bảo vệ'); } }
    ] },
    { label: 'Mực', items: [{ type: 'big', label: 'Ẩn Mực', icon: 'eye-off', active: () => d.inkHidden, run: () => { d.inkHidden = !d.inkHidden; d.objects.renderInk(); } }] }
  ] };
}

function viewTab(d) {
  return { id: 'view', label: 'Xem', en: 'View', groups: [
    { label: 'Dạng xem Sổ làm việc', items: [
      { type: 'big', label: 'Bình thường', icon: 'grid-3x3', active: () => !d.wb.cur.showPageBreaks, run: () => { d.wb.cur.showPageBreaks = false; d.grid.redraw(); } },
      { type: 'big', label: 'Xem trước Ngắt trang', icon: 'square-dashed', active: () => !!d.wb.cur.showPageBreaks, run: () => { d.wb.cur.showPageBreaks = true; d.grid.redraw(); } },
      { type: 'big', label: 'Bố trí Trang', icon: 'file', run: () => d.shell.backstage(d, 'print') },
      { type: 'big', label: 'Android DeX', icon: 'monitor', title: 'Chế độ Android DeX Desktop Mode & Chiếu màn hình ngoài / PC Link', run: () => import('../core/dex.js').then(m => m.openDexGuideDialog()) },
      { type: 'drop', big: true, label: 'Bảo vệ mắt', icon: 'sun', title: 'Chế độ bảo vệ mắt (WPS Eye Protection)', menu: () => import('../core/suites.js').then(m => m.eyeProtectionMenu(d)) }
    ] },
    { label: 'Hiện', items: [
      { type: 'col', items: [
        { type: 'btn', label: 'Đường lưới', icon: 'grid-3x3', active: () => d.wb.cur.showGrid, run: () => { d.wb.cur.showGrid = !d.wb.cur.showGrid; d.grid.redraw(); } },
        { type: 'btn', label: 'Thanh Công thức', icon: 'function-square', active: () => !d.formulaBar.classList.contains('hidden'), run: () => { d.formulaBar.classList.toggle('hidden'); d.grid.resize(); } },
        { type: 'btn', label: 'Đầu đề', icon: 'heading', active: () => d.wb.cur.showHead, run: () => { d.wb.cur.showHead = !d.wb.cur.showHead; d.grid.resize(); } }
      ] },
      { type: 'big', label: 'Bộ định hướng', icon: 'compass', title: 'Bộ định hướng đối tượng (LibreOffice Navigator)', run: () => import('../core/suites.js').then(m => m.openNavigator(d)) }
    ] },
    { label: 'Thu phóng', items: [
      { type: 'big', label: 'Thu phóng', icon: 'zoom-in', run: () => showMenu([200, 150, 125, 100, 85, 75, 50, 25].map(z => ({ label: z + '%', run: () => d.shell.setZoom(z) })), { x: 300, y: 150 }) },
      { type: 'big', label: '100%', icon: 'scan', run: () => d.shell.setZoom(100) },
      { type: 'big', label: 'Thu phóng theo Vùng chọn', icon: 'scan-search', run: () => { const rg = d.grid.selRange(); const s = d.wb.cur; let w = 0, hh = 0; for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 200); c++) w += s.colWidth(c); for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 500); r++) hh += s.rowHeight(r); const z = Math.max(10, Math.min(400, Math.floor(Math.min((d.grid.W - 60) / w, (d.grid.H - 40) / hh) * 100))); d.shell.setZoom(z); d.grid.scrollTo(rg.r1, rg.c1); } }
    ] },
    { label: 'Cửa sổ', items: [
      { type: 'big', label: 'Cửa sổ Mới', icon: 'app-window', run: () => import('../core/host.js').then(m => m.default.newWindow()) },
      { type: 'drop', big: true, label: 'Cố định Ngăn', icon: 'snowflake', menu: () => [{ label: 'Cố định Ngăn', desc: 'Giữ hàng và cột phía trên/trái ô hiện tại', run: () => import('./features.js').then(m => m.freeze(d, 'cell')) }, { label: 'Cố định Hàng Trên cùng', run: () => import('./features.js').then(m => m.freeze(d, 'row')) }, { label: 'Cố định Cột Đầu tiên', run: () => import('./features.js').then(m => m.freeze(d, 'col')) }, { label: 'Bỏ Cố định Ngăn', run: () => import('./features.js').then(m => m.freeze(d, 'none')) }] },
      { type: 'col', items: [
        { type: 'btn', label: 'Ẩn', icon: 'eye-off', run: () => { if (d.wb.sheets.filter(s => !s.hidden).length < 2) return toast('Sổ làm việc phải có ít nhất một trang tính hiển thị', 'err'); d.wb.cur.hidden = true; d.switchSheet(1); d.renderTabs(); } },
        { type: 'btn', label: 'Bỏ ẩn', icon: 'eye', disabled: () => !d.wb.sheets.some(s => s.hidden), run: () => X.unhideSheetDialog(d) }
      ] },
      { type: 'drop', big: true, label: 'Chuyển Cửa sổ', icon: 'layers', menu: () => d.shell.docs.map(x => ({ label: x.name + ' — ' + x.meta.label, checked: x === d || undefined, run: () => d.shell.showDoc(x) })) }
    ] },
    { label: 'Android DeX', items: [
      { type: 'big', label: 'Cài đặt DeX', icon: 'monitor', title: 'Trung tâm Cài đặt & Điều khiển Android DeX', run: () => import('../core/dex.js').then(D => D.openDexControlCenterDialog()) },
      { type: 'col', items: [
        { type: 'btn', label: 'Bàn rê chuột ảo', icon: 'mouse-pointer', run: () => import('../core/dex.js').then(D => D.openVirtualTrackpad()) },
        { type: 'btn', label: 'Chia đôi màn hình', icon: 'columns', run: () => import('../core/dex.js').then(D => D.tileWindow('left')) }
      ] }
    ] },
    { label: 'Macro', items: [{ type: 'drop', big: true, label: 'Macro', icon: 'square-code', title: 'Macro VBA (Alt+F8)', menu: () => [
      { label: 'Xem Macro (Alt+F8)', icon: 'list', run: () => import('./vba/index.js').then(m => m.macrosDialog(d)) },
      { label: d._vbaRec ? 'Dừng Ghi' : 'Ghi Macro…', icon: d._vbaRec ? 'square' : 'circle-dot', run: () => import('./vba/index.js').then(m => m.toggleRecord(d)) },
      { label: 'Dùng Tham chiếu Tương đối', icon: 'move', run: () => import('./vba/index.js').then(m => m.toggleRecord(d, true)) },
      { sep: true }, { label: 'Office Scripts (JavaScript)…', icon: 'file-code', run: () => import('./automate.js').then(m => m.scriptsPane(d)) }] }] }
  ] };
}
// v0.0.6: Developer tab (Nhà phát triển) — Visual Basic Editor, Macros, macro recorder, macro security info
function developerTab(d) {
  const V = f => () => import('./vba/index.js').then(m => f(m));
  return { id: 'developer', label: 'Nhà phát triển', en: 'Developer', groups: [
    { label: 'Mã', items: [
      { type: 'big', label: 'Visual Basic', icon: 'square-code', title: 'Mở trình soạn VBA (Alt+F11)', run: V(m => m.openEditor(d)) },
      { type: 'big', label: 'Macro', icon: 'list-video', title: 'Danh sách macro (Alt+F8)', run: V(m => m.macrosDialog(d)) },
      { type: 'col', items: [
        { type: 'btn', label: 'Ghi Macro', icon: 'circle-dot', title: 'Ghi Macro / Dừng Ghi', active: () => !!d._vbaRec, run: V(m => m.toggleRecord(d)) },
        { type: 'btn', label: 'Tham chiếu Tương đối', icon: 'move', run: V(m => m.toggleRecord(d, true)) },
        { type: 'btn', label: 'Cửa sổ Lệnh (Ctrl+G)', icon: 'terminal', title: 'Immediate Window — Thực thi lệnh trực tiếp', run: () => import('../core/suites.js').then(m => m.openImmediateWindow(d)) }
      ] },
      { type: 'big', label: 'Bảo mật Macro', icon: 'shield-check', run: () => alertDlg('Macro VBA trong Yomi-Office Open Source chạy trong môi trường cách ly: không truy cập tệp hệ thống, không chạy chương trình ngoài (Shell/CreateObject bị chặn), không gửi dữ liệu ra mạng. Macro chỉ chạy khi bạn bấm Chạy.', 'Bảo mật Macro') }
    ] },
    { label: 'esbuild (TS / TSX)', items: [
      { type: 'big', label: 'TypeScript / TSX', icon: 'file-code', title: 'Trình biên dịch & Thực thi TypeScript / TSX Macro bằng esbuild (ts vs tsx)', run: () => import('../core/suites.js').then(m => m.openEsbuildEditor(d, 'ts')) }
    ] },
    { label: 'Bổ trợ', items: [{ type: 'big', label: 'Office Scripts', icon: 'file-code', run: () => import('./automate.js').then(m => m.scriptsPane(d)) }] },
    { label: 'Chạy nhanh', items: [{ type: 'custom', render: () => { const box = h('div.rb-col'); import('./vba/index.js').then(m => { const ms = m.listMacros(d).slice(0, 6); if (!ms.length) box.appendChild(h('span.muted.small', 'Chưa có macro')); ms.forEach((x, i) => { if (i % 3 === 0) box.appendChild(h('div.rb-row')); const b = h('button.rb-btn', ic('play', 14), h('span.rb-lbl', x.name)); b.addEventListener('click', () => m.runMacro(d, x.name)); box.lastChild.appendChild(b); }); }); return box; } }] }
  ] };
}
function automateTab(d) {
  return { id: 'automate', label: 'Tự động hóa', en: 'Automate', groups: [
    { label: 'Office Scripts', items: [
      { type: 'big', label: 'Tập lệnh Mới', icon: 'file-code', run: () => import('./automate.js').then(m => m.scriptEditor(d)) },
      { type: 'big', label: 'Xem Tập lệnh', icon: 'folder-code', run: () => import('./automate.js').then(m => m.scriptsPane(d)) }
    ] },
    { label: 'Thư viện Office Scripts', items: [{ type: 'custom', render: () => { const box = h('div.rb-col'); import('./automate.js').then(m => m.BUILTIN.slice(0, 6).forEach((sc, i) => { if (i % 3 === 0) box.appendChild(h('div.rb-row')); const b = h('button.rb-btn', ic('file-code', 16), h('span.rb-lbl', sc.name)); b.addEventListener('click', () => m.runScript(d, sc.code, sc.name)); box.lastChild.appendChild(b); })); return box; } }] },
    { label: 'Tự động hóa Công việc', items: [{ type: 'big', label: 'Ghi lại Thao tác', icon: 'circle-dot', run: () => import('./automate.js').then(m => m.recordActions(d)) }] }
  ] };
}

function chartDesignTab(d, o) {
  const mut = (label, fn) => d.objects.mutate(label, fn);
  return { id: 'chartdesign', label: 'Thiết kế Biểu đồ', en: 'Chart Design', contextual: true, groups: [
    { label: 'Bố trí Biểu đồ', items: [
      { type: 'drop', big: true, label: 'Thêm Thành phần', icon: 'plus-square', menu: () => [{ label: (o.title === false ? 'Hiện' : 'Ẩn') + ' tiêu đề', run: () => mut('Tiêu đề', () => { o.title = o.title === false ? '' : false; }) }, { label: (o.legend === false ? 'Hiện' : 'Ẩn') + ' chú giải', run: () => mut('Chú giải', () => { o.legend = o.legend === false; }) }, { label: (o.labels ? 'Ẩn' : 'Hiện') + ' nhãn dữ liệu', run: () => mut('Nhãn', () => { o.labels = !o.labels; }) }] },
      { type: 'big', label: 'Sửa tiêu đề', icon: 'heading', run: () => promptDlg('Tiêu đề biểu đồ:', o.title || '', 'Tiêu đề').then(v => { if (v != null) mut('Tiêu đề', () => { o.title = v; }); }) }
    ] },
    { label: 'Kiểu Biểu đồ', items: [{ type: 'gallery', width: 360, items: () => [PALETTE_SETS[0], PALETTE_SETS[1], PALETTE_SETS[2], PALETTE_SETS[3], PALETTE_SETS[4]].map((p, i) => ({ title: 'Màu ' + (i + 1), preview: () => h('div.pal-prev', p.map(c => h('span', { style: { background: c } }))), active: () => JSON.stringify(o.colors) === JSON.stringify(p), run: () => mut('Màu biểu đồ', () => { o.colors = p; }) })) }] },
    { label: 'Dữ liệu', items: [{ type: 'big', label: 'Chuyển Hàng/Cột', icon: 'arrow-left-right', run: () => mut('Chuyển hàng/cột', () => { o.byRow = !o.byRow; }) }, { type: 'big', label: 'Chọn Dữ liệu', icon: 'table', run: () => d.objects.chartDataDialog(o) }] },
    { label: 'Kiểu', items: [{ type: 'drop', big: true, label: 'Đổi Kiểu Biểu đồ', icon: 'chart-column', menu: () => CHART_TYPES.map(t => ({ label: t[1], icon: t[2], checked: o.kind === t[0] || undefined, run: () => mut('Đổi kiểu', () => { o.kind = t[0]; }) })) }] },
    { label: 'Vị trí', items: [{ type: 'big', label: 'Lưu thành Ảnh', icon: 'image-down', run: () => d.objects.saveChartImage(o) }, { type: 'big', label: 'Văn bản Thay thế', icon: 'accessibility', run: () => d.objects.altText(o) }] }
  ] };
}
const PALETTE_SETS = [['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'], ['#0F6CBD', '#8764B8', '#00B7C3', '#E3008C', '#FFB900', '#107C10'], ['#1F3864', '#2E75B5', '#9DC3E6', '#BDD7EE', '#DEEBF7', '#5B9BD5'], ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050'], ['#404040', '#7F7F7F', '#A6A6A6', '#BFBFBF', '#D9D9D9', '#595959']];
function pictureTab(d, o) {
  const mut = (label, fn) => d.objects.mutate(label, fn);
  return { id: 'picformat', label: 'Định dạng Ảnh', en: 'Picture Format', contextual: true, groups: [
    { label: 'Điều chỉnh', items: [{ type: 'drop', big: true, label: 'Hiệu ứng', icon: 'wand', menu: () => [['none', 'Không'], ['grayscale(1)', 'Đen trắng'], ['sepia(1)', 'Nâu cổ điển'], ['brightness(1.2)', 'Sáng hơn'], ['contrast(1.3)', 'Tương phản cao'], ['blur(2px)', 'Làm mờ']].map(x => ({ label: x[1], run: () => mut('Hiệu ứng ảnh', () => { o.filter = x[0] === 'none' ? null : x[0]; }) })) }] },
    { label: 'Kiểu Ảnh', items: [{ type: 'drop', big: true, label: 'Viền Ảnh', icon: 'square', menu: () => [['', 'Không viền'], ['1px solid #000', 'Mảnh đen'], ['3px solid #fff', 'Khung trắng'], ['4px solid #4472c4', 'Khung xanh'], ['6px double #7f7f7f', 'Khung kép']].map(x => ({ label: x[1], run: () => mut('Viền ảnh', () => { o.border = x[0] || null; }) })) }] },
    { label: 'Trợ năng', items: [{ type: 'big', label: 'Văn bản Thay thế', icon: 'accessibility', run: () => d.objects.altText(o) }] },
    { label: 'Sắp xếp', items: [{ type: 'btn', label: 'Đưa lên trước', icon: 'bring-to-front', run: () => d.objects.order(o, 1) }, { type: 'btn', label: 'Đưa ra sau', icon: 'send-to-back', run: () => d.objects.order(o, -1) }, { type: 'btn', label: 'Xoay', icon: 'rotate-cw', run: () => mut('Xoay', () => { o.rot = ((o.rot || 0) + 90) % 360; }) }] },
    { label: 'Kích thước', items: [{ type: 'btn', label: 'Đặt lại kích thước', icon: 'scaling', run: () => mut('Kích thước', () => { const img = new Image(); img.src = o.src; if (img.naturalWidth) { o.w = img.naturalWidth; o.h = img.naturalHeight; } }) }] }
  ] };
}
function shapeTab(d, o) {
  const mut = (label, fn) => d.objects.mutate(label, fn);
  return { id: 'shapeformat', label: 'Định dạng Hình', en: 'Shape Format', contextual: true, groups: [
    { label: 'Kiểu Hình', items: [
      { type: 'color', big: true, label: 'Tô Hình', icon: 'paint-bucket', colorBar: () => o.fill || '#4472c4', run: () => {}, pick: (a, c) => mut('Tô hình', () => { o.fill = c; }), colorOpts: { none: 'Không tô' } },
      { type: 'color', big: true, label: 'Viền Hình', icon: 'square', colorBar: () => o.stroke || '#2f528f', run: () => {}, pick: (a, c) => mut('Viền hình', () => { o.stroke = c; }), colorOpts: { none: 'Không viền' } },
      { type: 'color', big: true, label: 'Màu chữ', icon: 'baseline', colorBar: () => o.color || '#000', run: () => {}, pick: (a, c) => mut('Màu chữ', () => { o.color = c; }) }
    ] },
    { label: 'Văn bản', items: [{ type: 'big', label: 'Sửa Văn bản', icon: 'type', run: () => d.objects.editObj(o, d.objects.inner.querySelector('[data-id="' + o.id + '"]')) }, { type: 'drop', big: true, label: 'Cỡ chữ', icon: 'a-large-small', menu: () => [10, 11, 12, 14, 18, 24, 32, 40, 54].map(z => ({ label: z + '', run: () => mut('Cỡ chữ', () => { o.sz = z; }) })) }] },
    { label: 'Sắp xếp', items: [{ type: 'btn', label: 'Đưa lên trước', icon: 'bring-to-front', run: () => d.objects.order(o, 1) }, { type: 'btn', label: 'Đưa ra sau', icon: 'send-to-back', run: () => d.objects.order(o, -1) }, { type: 'btn', label: 'Xoay', icon: 'rotate-cw', run: () => mut('Xoay', () => { o.rot = ((o.rot || 0) + 90) % 360; }) }] }
  ] };
}
