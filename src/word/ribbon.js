// Word ribbon: Home, Insert, Draw, Design, Layout, References, Mailings, Review, View, Help (+ contextual Table/Picture)
import { h, ic } from '../core/util.js';
import { toast, showMenu, colorPicker, dialog, promptDlg } from '../core/ui.js';
import { settings } from '../core/store.js';
import * as W from './cmds.js';
import { helpTab } from '../core/help.js';
import { symbolDialog, qrDialog, iconsDialog } from '../core/insert.js';
import { SHAPES } from '../core/shapes.js';
import { drawTab } from '../core/ink.js';

const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const st = d => d.queryState();

export function ribbonDef(doc) {
  const tabs = [homeTab(doc), insertTab(doc), drawTab(doc, doc._ink || (doc._ink = makeInk(doc))), designTab(doc), layoutTab(doc), referencesTab(doc), mailingsTab(doc), reviewTab(doc), viewTab(doc), developerTab(doc), helpTab(doc, 'word')];
  return { tabs };
}
function makeInk(doc) {
  // Word ink overlays the whole page as absolutely positioned SVG
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('class', 'wd-ink');
  doc.sheet.appendChild(svg); doc._inkSvg = svg; doc.wordInk = doc.wordInk || [];
  const catch_ = h('div.wd-inkcatch'); catch_.style.display = 'none'; doc.sheet.appendChild(catch_);
  const { InkTool, renderInk } = window.__inkLib || {};
  const render = () => { import('../core/ink.js').then(m => { m.renderInk(svg, doc.wordInk, 1); }); };
  doc.renderInk = render; doc.clearInk = () => { W.withUndo(doc, 'Xóa nét vẽ', () => { doc.wordInk = []; render(); }); };
  let tool = null;
  const ink = { get tool() { return tool; }, pen: { color: '#000', width: 2, opacity: 1 }, setTool(t, pen) { tool = t; if (pen) this.pen = pen; catch_.style.display = t ? 'block' : 'none'; } };
  import('../core/ink.js').then(m => { const it = new m.InkTool({ toLocal: e => { const r = svg.getBoundingClientRect(); const z = doc.zoom / 100; return [(e.clientX - r.left) / z, (e.clientY - r.top) / z]; }, getStrokes: () => doc.wordInk, scale: () => doc.zoom / 100, begin: () => doc.pushHistory(), commit: () => { doc.markDirty(); }, redraw: render }); it.attach(catch_); ink.setTool = (t, pen) => { tool = t; if (pen) ink.pen = pen; it.tool = t; it.pen = ink.pen; catch_.style.display = t ? 'block' : 'none'; }; it.tool = null; });
  return ink;
}

function homeTab(d) {
  const FONTS = W.FONTS;
  return { id: 'home', label: 'Trang đầu', en: 'Home', groups: [
    { label: 'Bảng tạm', items: [
      { type: 'split', big: true, label: 'Dán', icon: 'clipboard-paste', sc: 'Ctrl+V', run: () => { d.body.focus(); document.execCommand('paste'); }, menu: () => [{ label: 'Giữ định dạng nguồn', icon: 'clipboard-paste', run: () => { d.body.focus(); document.execCommand('paste'); } }, { label: 'Chỉ giữ văn bản', icon: 'type', run: () => W.pasteSpecial(d) }] },
      { type: 'col', items: [{ type: 'btn', label: 'Cắt', icon: 'scissors', sc: 'Ctrl+X', run: () => { d.body.focus(); document.execCommand('cut'); } }, { type: 'btn', label: 'Sao chép', icon: 'copy', sc: 'Ctrl+C', run: () => { d.body.focus(); document.execCommand('copy'); } }, { type: 'btn', label: 'Sao định dạng', icon: 'paintbrush', run: () => W.formatPainter(d) }] }
    ] },
    { label: 'Phông chữ', launcher: () => W.fontDialog(d), items: [
      { type: 'col', items: [
        { type: 'row', items: [
          { type: 'combo', width: 128, title: 'Phông chữ', fontBox: true, get: () => st(d).font || 'Calibri', set: (a, v) => W.setFont(d, v), options: () => FONTS, optStyle: v => 'font-family:' + v },
          { type: 'combo', width: 40, title: 'Cỡ chữ', get: () => String(st(d).size || 11), set: (a, v) => W.setSize(d, parseFloat(v)), options: () => SIZES.map(String) },
          { type: 'btn', icon: 'a-arrow-up', title: 'Tăng cỡ chữ', run: () => W.sizeStep(d, 1) },
          { type: 'btn', icon: 'a-arrow-down', title: 'Giảm cỡ chữ', run: () => W.sizeStep(d, -1) },
          { type: 'btn', icon: 'remove-formatting', title: 'Xóa mọi định dạng', run: () => W.clearFormat(d) }
        ] },
        { type: 'row', items: [
          { type: 'btn', icon: 'bold', title: 'Đậm', sc: 'Ctrl+B', active: () => st(d).bold, run: () => W.bold(d) },
          { type: 'btn', icon: 'italic', title: 'Nghiêng', sc: 'Ctrl+I', active: () => st(d).italic, run: () => W.italic(d) },
          { type: 'split', icon: 'underline', title: 'Gạch chân', sc: 'Ctrl+U', active: () => st(d).underline, run: () => W.underline(d), menu: () => [{ label: 'Gạch chân đơn', run: () => W.underline(d) }] },
          { type: 'btn', icon: 'strikethrough', title: 'Gạch ngang', active: () => st(d).strike, run: () => W.strike(d) },
          { type: 'btn', icon: 'subscript', title: 'Chỉ số dưới', active: () => st(d).sub, run: () => W.sub(d) },
          { type: 'btn', icon: 'superscript', title: 'Chỉ số trên', active: () => st(d).sup, run: () => W.sup(d) },
          { type: 'drop', icon: 'case-sensitive', title: 'Đổi kiểu chữ hoa/thường', menu: () => [['upper', 'CHỮ HOA'], ['lower', 'chữ thường'], ['title', 'Viết Hoa Đầu Từ'], ['sentence', 'Viết hoa đầu câu']].map(x => ({ label: x[1], run: () => changeCase(d, x[0]) })) },
          { type: 'color', icon: 'highlighter', title: 'Màu tô sáng', colorBar: () => d._hl || '#ffff00', run: () => W.highlight(d, d._hl || '#ffff00'), pick: (a, c) => { d._hl = c; W.highlight(d, c === 'transparent' ? 'transparent' : c); }, colorOpts: { none: 'Không tô sáng' } },
          { type: 'color', icon: 'baseline', title: 'Màu phông chữ', colorBar: () => d._fc || '#c00000', run: () => W.foreColor(d, d._fc || '#c00000'), pick: (a, c) => { d._fc = c; W.foreColor(d, c); }, colorOpts: { auto: 'Tự động' } }
        ] }
      ] }
    ] },
    { label: 'Đoạn văn', launcher: () => W.paragraphDialog(d), items: [
      { type: 'col', items: [
        { type: 'row', items: [
          { type: 'split', icon: 'list', title: 'Danh sách dấu đầu dòng', active: () => st(d).ul, run: () => W.list(d, 'ul'), menu: () => ['•', '◦', '▪', '–', '➤'].map(b => ({ label: 'Dấu "' + b + '"', run: () => W.list(d, 'ul') })) },
          { type: 'split', icon: 'list-ordered', title: 'Danh sách đánh số', active: () => st(d).ol, run: () => W.list(d, 'ol'), menu: () => [{ label: '1. 2. 3.', run: () => W.list(d, 'ol') }, { label: 'a. b. c.', run: () => W.list(d, 'ol') }] },
          { type: 'btn', icon: 'list-tree', title: 'Danh sách nhiều cấp', run: () => W.list(d, 'ul') },
          { type: 'btn', icon: 'indent-decrease', title: 'Giảm thụt lề', run: () => W.outdent(d) },
          { type: 'btn', icon: 'indent-increase', title: 'Tăng thụt lề', run: () => W.indent(d) },
          { type: 'btn', icon: 'arrow-down-up', title: 'Sắp xếp', run: () => sortParas(d) },
          { type: 'btn', icon: 'pilcrow', title: 'Hiện/ẩn ký hiệu định dạng', active: () => d.showMarks, run: () => W.toggleMarks(d) }
        ] },
        { type: 'row', items: [
          { type: 'btn', icon: 'align-left', title: 'Căn trái', sc: 'Ctrl+L', active: () => st(d).alignL, run: () => W.align(d, 'Left') },
          { type: 'btn', icon: 'align-center', title: 'Căn giữa', sc: 'Ctrl+E', active: () => st(d).alignC, run: () => W.align(d, 'Center') },
          { type: 'btn', icon: 'align-right', title: 'Căn phải', sc: 'Ctrl+R', active: () => st(d).alignR, run: () => W.align(d, 'Right') },
          { type: 'btn', icon: 'align-justify', title: 'Căn đều hai bên', sc: 'Ctrl+J', active: () => st(d).alignJ, run: () => W.align(d, 'Full') },
          { type: 'drop', icon: 'rows-3', title: 'Giãn dòng và khoảng cách đoạn', menu: () => [['1', '1,0'], ['1.15', '1,15'], ['1.5', '1,5'], ['2', '2,0'], ['2.5', '2,5'], ['3', '3,0']].map(x => ({ label: x[1], run: () => W.lineSpacing(d, x[0]) })).concat([{ sep: true }, { label: 'Thêm cách trước đoạn', run: () => W.spaceBefore(d, 12) }, { label: 'Thêm cách sau đoạn', run: () => W.spaceAfter(d, 12) }, { label: 'Bỏ cách sau đoạn', run: () => W.spaceAfter(d, 0) }]) },
          { type: 'color', icon: 'paint-bucket', title: 'Tô nền đoạn', colorBar: () => d._shade || '#d9e1f2', run: () => shade(d, d._shade || '#d9e1f2'), pick: (a, c) => { d._shade = c; shade(d, c); }, colorOpts: { none: 'Không màu' } },
          { type: 'drop', icon: 'grid-2x2', title: 'Viền', menu: () => [['bottom', 'Viền dưới'], ['all', 'Tất cả viền'], ['outside', 'Viền ngoài'], ['none', 'Không viền']].map(x => ({ label: x[1], run: () => paraBorder(d, x[0]) })) }
        ] }
      ] }
    ] },
    { label: 'Kiểu', items: [{ type: 'gallery', width: 360, items: () => W.stylesGallery(d), more: () => W.STYLES.map(s => ({ label: s[1], run: () => W.applyStyle(d, s[0]) })) }] },
    { label: 'Chỉnh sửa', items: [
      { type: 'col', items: [{ type: 'btn', label: 'Tìm', icon: 'search', sc: 'Ctrl+F', run: () => W.findDialog(d) }, { type: 'btn', label: 'Thay thế', icon: 'replace', sc: 'Ctrl+H', run: () => W.findDialog(d, true) }, { type: 'btn', label: 'Chọn tất cả', icon: 'box-select', run: () => { d.body.focus(); document.execCommand('selectAll'); } }] }
    ] },
    { label: 'A.I', items: [{ type: 'big', label: 'Copilot (A.I)', icon: 'sparkles', run: () => import('../ai/ai.js').then(m => m.openAi(d.shell, d)) }, { type: 'big', label: 'Soát lỗi A.I', icon: 'spell-check', run: () => import('../ai/ai.js').then(m => m.proofread(d)) }] }
  ] };
}
function changeCase(d, mode) { W.withUndo(d, 'Đổi kiểu chữ', () => { const sel = window.getSelection(); const t = sel.toString(); if (!t) return; let r; if (mode === 'upper') r = t.toUpperCase(); else if (mode === 'lower') r = t.toLowerCase(); else if (mode === 'title') r = t.toLowerCase().replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase()); else r = t.charAt(0).toUpperCase() + t.slice(1).toLowerCase(); document.execCommand('insertText', false, r); }); }
function shade(d, c) { W.withUndo(d, 'Tô nền đoạn', () => W.selectedBlocks(d).forEach(b => b.style.background = c === 'transparent' ? '' : c)); }
function paraBorder(d, kind) { W.withUndo(d, 'Viền đoạn', () => W.selectedBlocks(d).forEach(b => { if (kind === 'none') b.style.border = b.style.borderBottom = ''; else if (kind === 'bottom') b.style.borderBottom = '1px solid #000'; else { b.style.border = '1px solid #000'; b.style.padding = '2px 6px'; } })); }
function sortParas(d) { W.withUndo(d, 'Sắp xếp', () => { const blocks = W.selectedBlocks(d); if (blocks.length < 2) return toast('Chọn nhiều đoạn để sắp xếp', 'err'); const sorted = blocks.slice().sort((a, b) => a.textContent.localeCompare(b.textContent, 'vi')); const parent = blocks[0].parentNode; sorted.forEach(b => parent.insertBefore(b, blocks[blocks.length - 1].nextSibling)); }); }

function insertTab(d) {
  return { id: 'insert', label: 'Chèn', en: 'Insert', groups: [
    { label: 'Trang', items: [
      { type: 'drop', big: true, label: 'Trang bìa', icon: 'book-image', menu: () => [{ label: 'Trang bìa đơn giản', run: () => W.insertCover(d) }, { label: 'Trang bìa VISICONS', run: () => W.insertCover(d) }] },
      { type: 'big', label: 'Trang trống', icon: 'file', run: () => W.insertPageBreak(d, true) },
      { type: 'big', label: 'Ngắt trang', icon: 'separator-horizontal', sc: 'Ctrl+Enter', run: () => W.insertPageBreak(d) }
    ] },
    { label: 'Bảng', items: [{ type: 'drop', big: true, label: 'Bảng', icon: 'table', menu: () => [{ label: 'Chèn bảng…', icon: 'table', run: () => W.insertTableDialog(d) }, { label: 'Bảng nhanh 2×2', run: () => W.insertHtml(d, quickTable(2, 2)) }, { label: 'Bảng nhanh 3×3', run: () => W.insertHtml(d, quickTable(3, 3)) }, { label: 'Bảng nhanh 5×4', run: () => W.insertHtml(d, quickTable(4, 5)) }] }] },
    { label: 'Hình minh họa', items: [
      { type: 'big', label: 'Ảnh', icon: 'image', run: () => W.insertImage(d) },
      { type: 'drop', big: true, label: 'Hình dạng', icon: 'shapes', menu: () => SHAPES.slice(0, 24).map(s => ({ label: s[1], run: () => insertShape(d, s[0]) })) },
      { type: 'big', label: 'Biểu tượng', icon: 'smile', run: () => iconsDialog(url => W.insertImageUrl(d, url)) },
      { type: 'big', label: 'Biểu đồ', icon: 'chart-column', run: () => insertChartImg(d) },
      { type: 'big', label: 'Mã QR', icon: 'qr-code', run: () => qrDialog(url => W.insertImageUrl(d, url)) },
      { type: 'drop', big: true, label: 'Thẻ thông minh', icon: 'credit-card', title: 'Thẻ Craft Project (Callout, KPI, Trích dẫn)', menu: () => [
        { label: 'Khối lưu ý (Callout Box)', icon: 'info', run: () => import('../core/suites.js').then(m => m.insertCraftCard(d, 'callout')) },
        { label: 'Thẻ số liệu KPI (Stat Card)', icon: 'chart-bar', run: () => import('../core/suites.js').then(m => m.insertCraftCard(d, 'stat')) },
        { label: 'Khối trích dẫn (Quote Card)', icon: 'quote', run: () => import('../core/suites.js').then(m => m.insertCraftCard(d, 'quote')) }
      ] }
    ] },
    { label: 'Liên kết', items: [{ type: 'big', label: 'Liên kết', icon: 'link', sc: 'Ctrl+K', run: () => W.linkDialog(d) }, { type: 'big', label: 'Dấu trang', icon: 'bookmark', run: () => bookmark(d) }] },
    { label: 'Chú thích', items: [{ type: 'big', label: 'Chú thích', icon: 'message-square-plus', run: () => W.addComment(d) }] },
    { label: 'Đầu trang & Chân trang', items: [
      { type: 'drop', big: true, label: 'Đầu trang', icon: 'panel-top', menu: () => [{ label: 'Nhập đầu trang…', run: () => headerFooter(d, 'header') }] },
      { type: 'drop', big: true, label: 'Chân trang', icon: 'panel-bottom', menu: () => [{ label: 'Nhập chân trang…', run: () => headerFooter(d, 'footer') }, { label: 'Số trang', run: () => W.insertField(d, 'page') }] },
      { type: 'drop', big: true, label: 'Số trang', icon: 'hash', menu: () => [{ label: 'Chèn số trang hiện tại', run: () => W.insertField(d, 'page') }] }
    ] },
    { label: 'Văn bản', items: [
      { type: 'big', label: 'Hộp văn bản', icon: 'square-dashed-bottom-code', run: () => W.insertHtml(d, '<div class="wd-textbox" style="border:1px solid #bbb;padding:8px;margin:8px 0;display:inline-block;min-width:200px">Nội dung hộp văn bản</div>') },
      { type: 'drop', big: true, label: 'Thành phần nhanh', icon: 'blocks', menu: () => [{ label: 'Trường: Ngày', run: () => W.insertField(d, 'date') }, { label: 'Trường: Giờ', run: () => W.insertField(d, 'time') }, { label: 'Trường: Tác giả', run: () => W.insertField(d, 'author') }, { label: 'Trường: Tên tệp', run: () => W.insertField(d, 'filename') }, { label: 'Mục lục', run: () => W.insertTOC(d) }] },
      { type: 'big', label: 'WordArt', icon: 'whole-word', run: () => insertWordArt(d) },
      { type: 'big', label: 'Chữ hoa đầu đoạn', icon: 'pilcrow-square', run: () => dropCap(d) },
      { type: 'drop', big: true, label: 'Ngày & Giờ', icon: 'calendar-clock', menu: () => [['date', 'Ngày'], ['time', 'Giờ'], ['datetime', 'Ngày và Giờ']].map(x => ({ label: x[1], run: () => W.insertField(d, x[0]) })) }
    ] },
    { label: 'Ký hiệu', items: [
      { type: 'big', label: 'Phương trình (MathType)', icon: 'sigma-square', run: () => import('../math/math.js').then(m => m.equationEditor('').then(lx => { if (lx) W.insertEquation(d, lx); })) },
      { type: 'drop', big: true, label: 'Ký hiệu', icon: 'omega', menu: () => [{ label: 'Thêm ký hiệu…', run: () => symbolDialog(ch => W.insertSymbol(d, ch)) }, { sep: true }].concat(['©', '®', '™', '°', '±', '≈', '≠', '→', '★', '§'].map(c => ({ label: c, run: () => W.insertSymbol(d, c) }))) }
    ] }
  ] };
}
function quickTable(r, c) { let h = '<table class="wd-table"><tbody>'; for (let i = 0; i < r; i++) { h += '<tr>'; for (let j = 0; j < c; j++) h += (i === 0 ? '<th><br></th>' : '<td><br></td>'); h += '</tr>'; } return h + '</tbody></table><p><br></p>'; }
function insertShape(d, shape) { import('../core/shapes.js').then(m => { const svg = m.shapeSvg(shape, 160, 100, '#4472c4', '#2f528f', 1.5); W.insertHtml(d, '<span class="wd-shape" contenteditable="false" style="display:inline-block;width:160px;height:100px;vertical-align:middle">' + svg + '</span>&nbsp;'); }); }
function insertChartImg(d) { import('../excel/charts.js').then(m => { const data = { cats: ['Quý 1', 'Quý 2', 'Quý 3', 'Quý 4'], series: [{ name: 'Doanh thu', values: [120, 150, 138, 170] }] }; const div = document.createElement('div'); const svg = m.renderChart(div, data, { kind: 'column', w: 480, h: 300, title: 'Biểu đồ' }); W.insertHtml(d, '<span class="wd-chart" contenteditable="false" style="display:block;max-width:480px">' + div.innerHTML + '</span><p><br></p>'); }); }
function bookmark(d) { promptDlg('Tên dấu trang:', '', 'Dấu trang').then(v => { if (v) W.insertHtml(d, '<a id="' + v.replace(/\s/g, '_') + '" class="wd-bookmark"></a>'); }); }
function headerFooter(d, which) { promptDlg('Nội dung ' + (which === 'header' ? 'đầu trang' : 'chân trang') + ':', d['_' + which] || '', which === 'header' ? 'Đầu trang' : 'Chân trang').then(v => { if (v == null) return; d['_' + which] = v; d.markDirty(); toast('Đã đặt ' + (which === 'header' ? 'đầu trang' : 'chân trang') + ' — hiển thị khi in'); }); }
function insertWordArt(d) { promptDlg('Văn bản nghệ thuật:', 'Văn bản', 'WordArt').then(v => { if (v) W.insertHtml(d, '<span class="wd-wordart" style="font-size:36px;font-weight:bold;color:#4472c4;text-shadow:2px 2px 3px rgba(0,0,0,.3)">' + (v || '') + '</span>'); }); }
function dropCap(d) { const b = W.selectedBlocks ? d : d; W.withUndo(d, 'Chữ hoa đầu đoạn', () => { const bl = (W.currentBlock ? null : null); const block = window.getSelection().anchorNode; let p = block; if (p && p.nodeType === 3) p = p.parentElement; while (p && !/^P$/.test(p.tagName)) p = p.parentElement; if (p && p.textContent) { const first = p.textContent[0]; p.innerHTML = '<span class="wd-dropcap">' + first + '</span>' + p.innerHTML.replace(first, ''); } }); }

function designTab(d) {
  const THEMES = [['Office', 'Calibri', '#4472c4'], ['Liquid Glass', 'Segoe UI', '#0f6cbd'], ['VISICONS', 'Times New Roman', '#1f3864'], ['Trang nhã', 'Garamond', '#7030a0'], ['Hiện đại', 'Century Gothic', '#2e75b5'], ['Tối giản', 'Arial', '#404040']];
  return { id: 'design', label: 'Thiết kế', en: 'Design', groups: [
    { label: 'Định dạng Tài liệu', items: [
      { type: 'gallery', width: 300, items: () => THEMES.map(t => ({ title: t[0], preview: () => h('div.theme-prev', { style: { fontFamily: t[1] } }, h('div', { style: { color: t[2], fontWeight: 'bold', fontSize: '13px' } }, 'Aa'), h('div', { style: { fontSize: '8px', color: '#666' } }, t[0])), run: () => applyDocTheme(d, t) })) },
      { type: 'drop', big: true, label: 'Màu', icon: 'palette', menu: () => THEMES.map(t => ({ label: t[0], swatch: t[2], run: () => applyDocTheme(d, t) })) },
      { type: 'drop', big: true, label: 'Phông chữ', icon: 'type', menu: () => W.FONTS.slice(0, 14).map(f => ({ label: f, style: 'font-family:' + f, run: () => { d.body.style.fontFamily = f; d.markDirty(); } })) },
      { type: 'drop', big: true, label: 'Giãn dòng Đoạn', icon: 'rows-3', menu: () => [['1', '1,0'], ['1.15', '1,15'], ['1.5', '1,5'], ['2', '2,0']].map(x => ({ label: x[1], run: () => { d.body.style.lineHeight = x[0]; d.markDirty(); } })) }
    ] },
    { label: 'Nền Trang', items: [
      { type: 'drop', big: true, label: 'Hình mờ (Watermark)', icon: 'droplets', menu: () => [{ label: 'MẬT', run: () => watermark(d, 'MẬT') }, { label: 'BẢN NHÁP', run: () => watermark(d, 'BẢN NHÁP') }, { label: 'KHẨN', run: () => watermark(d, 'KHẨN') }, { label: 'Tùy chỉnh…', run: () => promptDlg('Nội dung hình mờ:', '', 'Hình mờ').then(v => v && watermark(d, v)) }, { label: 'Xóa hình mờ', run: () => watermark(d, null) }] },
      { type: 'color', big: true, label: 'Màu Trang', icon: 'paint-bucket', colorBar: () => d._pageColor || '#ffffff', run: () => {}, pick: (a, c) => { d._pageColor = c; d.sheet.style.background = c === 'transparent' ? '#fff' : c; d.body.style.background = 'transparent'; d.markDirty(); }, colorOpts: { none: 'Không màu' } },
      { type: 'drop', big: true, label: 'Viền Trang', icon: 'square', menu: () => [['1px solid #000', 'Mảnh'], ['3px solid #4472c4', 'Đậm xanh'], ['3px double #000', 'Kép'], ['', 'Không viền']].map(x => ({ label: x[1], run: () => { d.body.style.outline = x[0]; d.body.style.outlineOffset = x[0] ? '-16px' : ''; d.markDirty(); } })) }
    ] }
  ] };
}
function applyDocTheme(d, t) { d.body.style.fontFamily = t[1]; d.el.style.setProperty('--wd-accent', t[2]); $$all(d, 'h1,h2,h3,h4').forEach(hd => hd.style.color = t[2]); d.markDirty(); toast('Đã áp dụng chủ đề ' + t[0]); }
function $$all(d, s) { return Array.prototype.slice.call(d.body.querySelectorAll(s)); }
function watermark(d, text) { let wm = d.sheet.querySelector('.wd-watermark'); if (!text) { if (wm) wm.remove(); d.markDirty(); return; } if (!wm) { wm = h('div.wd-watermark'); d.sheet.appendChild(wm); } wm.textContent = text; d.markDirty(); }

function layoutTab(d) {
  return { id: 'layout', label: 'Bố cục', en: 'Layout', groups: [
    { label: 'Thiết lập Trang', launcher: () => W.pageSetupDialog(d), items: [
      { type: 'drop', big: true, label: 'Lề', icon: 'square-dashed', menu: () => [['normal', 'Bình thường'], ['narrow', 'Hẹp'], ['moderate', 'Vừa'], ['wide', 'Rộng']].map(x => ({ label: x[1], checked: d.margins === x[0] || undefined, run: () => W.setPageQuick(d, 'margins', x[0]) })).concat([{ sep: true }, { label: 'Lề tùy chỉnh...', icon: 'ruler', run: () => W.pageSetupDialog(d) }]) },
      { type: 'drop', big: true, label: 'Hướng', icon: 'rectangle-vertical', menu: () => [['portrait', 'Dọc'], ['landscape', 'Ngang']].map(x => ({ label: x[1], checked: d.orientation === x[0] || undefined, run: () => W.setPageQuick(d, 'orientation', x[0]) })) },
      { type: 'drop', big: true, label: 'Khổ giấy', icon: 'file', menu: () => ['A4', 'A3', 'A5', 'Letter', 'Legal'].map(x => ({ label: x, checked: d.pageSize === x || undefined, run: () => W.setPageQuick(d, 'pageSize', x) })) },
      { type: 'drop', big: true, label: 'Cột', icon: 'columns-3', menu: () => [[1, 'Một'], [2, 'Hai'], [3, 'Ba']].map(x => ({ label: x[1], run: () => W.columns(d, x[0]) })) },
      { type: 'drop', big: true, label: 'Ngắt', icon: 'separator-horizontal', menu: () => [{ label: 'Ngắt trang', desc: 'Bắt đầu trang mới tại vị trí con trỏ (Ctrl+Enter)', run: () => W.insertPageBreak(d) }, { label: 'Trang trống', run: () => W.insertPageBreak(d, true) }, { label: 'Ngắt cột', run: () => W.insertHtml(d, '<span style="break-after:column"></span>') }, '-', { label: 'Đi tới trang…', run: () => W.gotoPageDialog(d) }, { label: 'Xóa tất cả ngắt trang', run: () => W.withUndo(d, 'Xóa ngắt trang', () => Array.prototype.slice.call(d.body.querySelectorAll('.wd-pagebreak')).forEach(x => x.parentNode.removeChild(x))) }] }
    ] },
    { label: 'Đoạn văn', items: [
      { type: 'col', items: [{ type: 'btn', label: 'Thụt lề trái +', icon: 'indent-increase', run: () => W.indent(d) }, { type: 'btn', label: 'Thụt lề trái −', icon: 'indent-decrease', run: () => W.outdent(d) }] },
      { type: 'col', items: [{ type: 'btn', label: 'Cách trước +6pt', icon: 'arrow-up', run: () => W.spaceBefore(d, 6) }, { type: 'btn', label: 'Cách sau +6pt', icon: 'arrow-down', run: () => W.spaceAfter(d, 6) }] }
    ] },
    { label: 'Sắp xếp', items: [
      { type: 'drop', big: true, label: 'Vị trí ảnh', icon: 'layout-template', menu: () => [{ label: 'Căn trái', run: () => floatImg(d, 'left') }, { label: 'Căn giữa', run: () => floatImg(d, 'center') }, { label: 'Căn phải', run: () => floatImg(d, 'right') }, { label: 'Nội tuyến (inline)', run: () => floatImg(d, 'none') }] },
      { type: 'drop', big: true, label: 'Ngắt chữ quanh ảnh', icon: 'wrap-text', menu: () => [{ label: 'Vuông (chữ bao quanh)', run: () => wrapImg(d, 'left') }, { label: 'Trên dưới', run: () => wrapImg(d, 'none') }] }
    ] }
  ] };
}
function floatImg(d, how) { const img = d._selImg; if (!img) return toast('Chọn một ảnh trước', 'err'); W.withUndo(d, 'Vị trí ảnh', () => { img.style.float = ''; img.style.display = ''; img.style.margin = ''; if (how === 'center') { img.style.display = 'block'; img.style.margin = '8px auto'; } else if (how === 'left') { img.style.float = 'left'; img.style.margin = '4px 12px 4px 0'; } else if (how === 'right') { img.style.float = 'right'; img.style.margin = '4px 0 4px 12px'; } }); }
function wrapImg(d, side) { floatImg(d, side === 'left' ? 'left' : 'none'); }

function referencesTab(d) {
  return { id: 'references', label: 'Tham chiếu', en: 'References', groups: [
    { label: 'Mục lục', items: [
      { type: 'drop', big: true, label: 'Mục lục', icon: 'list-tree', menu: () => [{ label: 'Mục lục tự động', run: () => W.insertTOC(d) }, { label: 'Cập nhật mục lục', run: () => { const toc = d.body.querySelector('.wd-toc'); if (toc) toc.remove(); W.insertTOC(d); } }, { label: 'Xóa mục lục', run: () => { const toc = d.body.querySelector('.wd-toc'); if (toc) { toc.remove(); d.markDirty(); } } }] },
      { type: 'big', label: 'Cập nhật', icon: 'refresh-cw', run: () => { const toc = d.body.querySelector('.wd-toc'); if (toc) { toc.remove(); W.insertTOC(d); } else toast('Chưa có mục lục'); } }
    ] },
    { label: 'Chú thích cuối trang', items: [
      { type: 'big', label: 'Chèn chú thích chân', icon: 'footprints', run: () => footnote(d) },
      { type: 'big', label: 'Chú thích cuối tài liệu', icon: 'book-marked', run: () => footnote(d, true) }
    ] },
    { label: 'Trích dẫn & Thư mục', items: [
      { type: 'drop', big: true, label: 'Chèn trích dẫn', icon: 'quote', menu: () => [{ label: 'Thêm nguồn mới…', run: () => citation(d) }] },
      { type: 'big', label: 'Thư mục', icon: 'library', run: () => W.insertHtml(d, '<h2>Tài liệu tham khảo</h2><p>[1] Tác giả, <i>Tiêu đề</i>, Nhà xuất bản, Năm.</p>') }
    ] },
    { label: 'Chú thích ảnh', items: [{ type: 'big', label: 'Chèn chú thích', icon: 'captions', run: () => caption(d) }] },
    { label: 'Chỉ mục', items: [{ type: 'big', label: 'Đánh dấu mục', icon: 'bookmark-plus', run: () => bookmark(d) }] }
  ] };
}
function footnote(d, end) { const n = (d._footnotes = (d._footnotes || 0) + 1); promptDlg('Nội dung chú thích ' + n + ':', '', 'Chú thích chân trang').then(v => { if (v == null) return; W.insertHtml(d, '<sup class="wd-fnref">[' + n + ']</sup>'); const note = h('div.wd-footnote', '[' + n + '] ' + v); d.body.appendChild(note); d.markDirty(); }); }
function citation(d) { formDlg('Thêm nguồn', [{ key: 'author', label: 'Tác giả' }, { key: 'title', label: 'Tiêu đề' }, { key: 'year', label: 'Năm' }]).then(v => { if (v) W.insertText(d, ' (' + v.author + ', ' + v.year + ')'); }); }
function caption(d) { promptDlg('Chú thích (Hình/Bảng):', 'Hình 1: ', 'Chú thích').then(v => { if (v) W.insertHtml(d, '<p class="wd-caption" style="text-align:center;font-size:10pt;color:#555"><i>' + (v || '') + '</i></p>'); }); }

function mailingsTab(d) {
  return { id: 'mailings', label: 'Thư tín', en: 'Mailings', groups: [
    { label: 'Tạo', items: [{ type: 'big', label: 'Phong bì', icon: 'mail', run: () => envelope(d) }, { type: 'big', label: 'Nhãn', icon: 'tag', run: () => labels(d) }] },
    { label: 'Bắt đầu Trộn Thư', items: [
      { type: 'drop', big: true, label: 'Bắt đầu Trộn Thư', icon: 'mails', menu: () => [{ label: 'Thư (Letters)', run: () => toast('Trộn thư: dùng các trường «Tên», «Địa chỉ» rồi Hoàn tất & Trộn') }, { label: 'Thư điện tử', run: () => toast('Soạn mẫu rồi dùng Hoàn tất & Trộn') }, { label: 'Nhãn', run: () => labels(d) }] },
      { type: 'big', label: 'Chọn người nhận', icon: 'users', run: () => mergeRecipients(d) }
    ] },
    { label: 'Viết & Chèn Trường', items: [
      { type: 'drop', big: true, label: 'Chèn Trường Trộn', icon: 'braces', menu: () => (d._mergeFields || ['Tên', 'Địa chỉ', 'Điện thoại', 'Email']).map(f => ({ label: f, run: () => W.insertHtml(d, '<span class="wd-mergefield">«' + f + '»</span>') })) },
      { type: 'big', label: 'Dòng Chào hỏi', icon: 'hand', run: () => W.insertHtml(d, '<p>Kính gửi <span class="wd-mergefield">«Tên»</span>,</p>') }
    ] },
    { label: 'Hoàn tất', items: [{ type: 'big', label: 'Hoàn tất & Trộn', icon: 'merge', run: () => mergeFinish(d) }] }
  ] };
}
function envelope(d) { toast('Phong bì: chọn khổ giấy nhỏ ở Bố cục > Khổ giấy rồi nhập địa chỉ'); }
function labels(d) { W.insertTableDialog(d); }
function mergeRecipients(d) { promptDlg('Dán dữ liệu người nhận (mỗi dòng một người, các cột cách nhau bằng Tab; dòng đầu là tên trường):', '', 'Người nhận trộn thư', { multiline: true, rows: 8 }).then(v => { if (!v) return; const lines = v.trim().split('\n'); d._mergeFields = lines[0].split('\t'); d._mergeData = lines.slice(1).map(l => { const cells = l.split('\t'); const o = {}; d._mergeFields.forEach((f, i) => o[f] = cells[i] || ''); return o; }); toast('Đã nạp ' + d._mergeData.length + ' người nhận, ' + d._mergeFields.length + ' trường'); d.refreshUi(); }); }
function mergeFinish(d) { if (!d._mergeData || !d._mergeData.length) return toast('Chọn người nhận và chèn trường trộn trước', 'err'); const tmpl = d.body.innerHTML; let out = ''; d._mergeData.forEach((row, i) => { let h = tmpl; d._mergeFields.forEach(f => { h = h.split('«' + f + '»').join(esc(row[f] || '')); }); h = h.replace(/<span class="wd-mergefield">([^<]*)<\/span>/g, '$1'); out += h + (i < d._mergeData.length - 1 ? '<div class="wd-pagebreak"></div>' : ''); }); W.withUndo(d, 'Trộn thư', () => { d.body.innerHTML = out; }); toast('Đã trộn ' + d._mergeData.length + ' thư', 'ok'); }

function reviewTab(d) {
  return { id: 'review', label: 'Xem lại', en: 'Review', groups: [
    { label: 'Soát lỗi', items: [
      { type: 'big', label: 'Chính tả & Ngữ pháp', icon: 'spell-check', sc: 'F7', run: () => import('../ai/ai.js').then(m => m.proofread(d)) },
      { type: 'col', items: [
        { type: 'btn', label: 'Từ điển đồng nghĩa', icon: 'book-a', run: () => import('../ai/ai.js').then(m => m.openAi(d.shell, d, 'Gợi ý từ đồng nghĩa cho: ' + window.getSelection().toString())) },
        { type: 'btn', label: 'Đếm số từ & Thống kê', icon: 'calculator', title: 'Thống kê chi tiết tài liệu (LibreOffice)', run: () => import('../core/suites.js').then(m => m.openWordCountStatistics(d)) },
        { type: 'btn', label: 'Tự động sửa lỗi', icon: 'type', title: 'Tự động sửa lỗi & gõ tắt (LibreOffice AutoCorrect)', run: () => import('../core/suites.js').then(m => m.openAutoCorrectDialog()) }
      ] }
    ] },
    { label: 'Ngôn ngữ', items: [{ type: 'big', label: 'Dịch', icon: 'languages', run: () => import('../ai/ai.js').then(m => m.openAi(d.shell, d, 'Dịch đoạn sau sang tiếng Anh:\n' + window.getSelection().toString())) },
      { type: 'big', label: 'Chuyển mã', icon: 'replace-all', title: 'Chuyển TCVN3 (.VnTime) / VNI → Unicode', run: () => import('../core/fonts.js').then(async F => { const mode = await F.convertDialog(); if (!mode) return; d.exec('Chuyển mã', () => { const n = F.convertWordLegacy(d, mode); toast(n ? 'Đã chuyển mã ' + n + ' đoạn văn bản sang Unicode' : 'Không tìm thấy văn bản mã cũ (TCVN3/VNI)', n ? 'ok' : null); }); }) }] },
    { label: 'Chú thích', items: [
      { type: 'big', label: 'Chú thích Mới', icon: 'message-square-plus', run: () => W.addComment(d) },
      { type: 'col', items: [{ type: 'btn', label: 'Xóa', icon: 'message-square-x', run: () => { const sel = window.getSelection(); let n = sel.anchorNode; if (n && n.nodeType === 3) n = n.parentElement; const c = n && n.closest('.wd-comment'); if (c) { const cid = c.dataset.cid; d.comments = d.comments.filter(x => x.id !== cid); const p = c.parentNode; while (c.firstChild) p.insertBefore(c.firstChild, c); p.removeChild(c); d.markDirty(); } } }, { type: 'btn', label: 'Hiển thị', icon: 'messages-square', run: () => W.commentsPane(d) }] }
    ] },
    { label: 'Theo dõi', items: [
      { type: 'big', label: 'Theo dõi Thay đổi', icon: 'git-compare', sc: 'Ctrl+Shift+E', active: () => d.trackChanges, run: () => W.toggleTrack(d) },
      { type: 'col', items: [{ type: 'btn', label: 'Chấp nhận tất cả', icon: 'check', run: () => W.acceptChanges(d, true) }, { type: 'btn', label: 'Từ chối tất cả', icon: 'x', run: () => W.rejectChanges(d) }] }
    ] },
    { label: 'Bảo vệ', items: [{ type: 'big', label: 'Hạn chế Chỉnh sửa', icon: 'lock', active: () => d.readOnly, run: () => { d.setReadOnly(!d.readOnly); toast(d.readOnly ? 'Đã khóa chỉnh sửa (chỉ đọc)' : 'Đã mở khóa'); d.refreshUi(); } }] },
    { label: 'Trợ năng', items: [{ type: 'big', label: 'Kiểm tra Trợ năng', icon: 'accessibility', run: () => d.shell.inspect(d) }] }
  ] };
}

function developerTab(d) {
  const V = f => () => import('../excel/vba/index.js').then(m => f(m));
  return { id: 'developer', label: 'Nhà phát triển', en: 'Developer', groups: [
    { label: 'Mã', items: [
      { type: 'big', label: 'Visual Basic', icon: 'square-code', title: 'Mở trình soạn VBA (Alt+F11)', run: V(m => m.openEditor(d)) },
      { type: 'big', label: 'Macro', icon: 'list-video', title: 'Danh sách macro (Alt+F8)', run: V(m => m.macrosDialog(d)) },
      { type: 'col', items: [
        { type: 'btn', label: 'Thư viện Macro', icon: 'library', title: 'Mô-đun VBA mẫu từ các dự án GitHub (AcuteMSWord, nsfc-md2docx)', run: () => import('../excel/vba/library.js').then(m => m.libraryDialog(d, () => import('../excel/vba/index.js').then(x => x.openEditor(d)))) },
        { type: 'btn', label: 'Bảo mật Macro', icon: 'shield-check', run: () => import('../core/ui.js').then(u => u.alertDlg('Macro VBA trong Yomi-Office Open Source chạy trong môi trường cách ly: không truy cập tệp hệ thống, không chạy chương trình ngoài, không gửi dữ liệu ra mạng. Macro chỉ chạy khi bạn bấm Chạy. Lưu dạng .docm để giữ macro.', 'Bảo mật Macro')) }
      ] }
    ] },
    { label: 'Biểu mẫu (OnlyOffice)', items: [
      { type: 'drop', big: true, label: 'Trường biểu mẫu', icon: 'file-check', menu: () => [
        { label: 'Hộp kiểm (Checkbox)', icon: 'check-square', run: () => import('../core/suites.js').then(m => m.insertFormControl(d, 'checkbox')) },
        { label: 'Danh sách chọn (Dropdown)', icon: 'list', run: () => import('../core/suites.js').then(m => m.insertFormControl(d, 'dropdown')) },
        { label: 'Trường ngày tháng (Date)', icon: 'calendar', run: () => import('../core/suites.js').then(m => m.insertFormControl(d, 'date')) },
        { label: 'Vùng văn bản (Text Field)', icon: 'type', run: () => import('../core/suites.js').then(m => m.insertFormControl(d, 'text')) }
      ] }
    ] },
    { label: 'esbuild (TS / TSX)', items: [
      { type: 'big', label: 'TypeScript / TSX', icon: 'file-code', title: 'Trình biên dịch & Thực thi TypeScript / TSX Macro bằng esbuild (ts vs tsx)', run: () => import('../core/suites.js').then(m => m.openEsbuildEditor(d, 'tsx')) }
    ] },
    { label: 'Chạy nhanh', items: [{ type: 'custom', render: () => { const box = h('div.rb-col'); import('../excel/vba/index.js').then(m => { const ms = m.listMacros(d).slice(0, 6); if (!ms.length) box.appendChild(h('span.muted.small', 'Chưa có macro')); ms.forEach((x, i) => { if (i % 3 === 0) box.appendChild(h('div.rb-row')); const b = h('button.rb-btn', ic('play', 14), h('span.rb-lbl', x.name)); b.addEventListener('click', () => m.runMacro(d, x.name)); box.lastChild.appendChild(b); }); }); return box; } }] }
  ] };
}
function viewTab(d) {
  return { id: 'view', label: 'Xem', en: 'View', groups: [
    { label: 'Dạng xem', items: [
      { type: 'big', label: 'Chế độ Đọc', icon: 'book-open', active: () => d.viewMode === 'read', run: () => d.setView('read') },
      { type: 'big', label: 'Bố trí In', icon: 'file', active: () => d.viewMode === 'print', run: () => d.setView('print') },
      { type: 'big', label: 'Bố trí Web', icon: 'globe', active: () => d.viewMode === 'web', run: () => d.setView('web') },
      { type: 'big', label: 'Android DeX', icon: 'monitor', title: 'Chế độ Android DeX Desktop Mode & Chiếu màn hình ngoài / PC Link', run: () => import('../core/dex.js').then(m => m.openDexGuideDialog()) },
      { type: 'drop', big: true, label: 'Bảo vệ mắt', icon: 'sun', title: 'Chế độ bảo vệ mắt (WPS Eye Protection)', menu: () => import('../core/suites.js').then(m => m.eyeProtectionMenu(d)) }
    ] },
    { label: 'Hiển thị', items: [
      { type: 'col', items: [
        { type: 'btn', label: 'Thước', icon: 'ruler', active: () => !d.rulerObj.hidden, run: () => { d.rulerObj.toggle(); } },
        { type: 'btn', label: 'Ký hiệu định dạng', icon: 'pilcrow', active: () => d.showMarks, run: () => W.toggleMarks(d) },
        { type: 'btn', label: 'Ngăn Dẫn hướng', icon: 'panel-left', run: () => W.navPane(d) }
      ] },
      { type: 'big', label: 'Bộ định hướng', icon: 'compass', title: 'Bộ định hướng đối tượng (LibreOffice Navigator)', run: () => import('../core/suites.js').then(m => m.openNavigator(d)) }
    ] },
    { label: 'Thu phóng', items: [
      { type: 'big', label: 'Thu phóng', icon: 'zoom-in', run: () => showMenu([200, 150, 125, 100, 75, 50].map(z => ({ label: z + '%', run: () => d.shell.setZoom(z) })), { x: 300, y: 150 }) },
      { type: 'big', label: '100%', icon: 'scan', run: () => d.shell.setZoom(100) },
      { type: 'big', label: 'Một trang', icon: 'file', run: () => fitPage(d) },
      { type: 'big', label: 'Chiều rộng Trang', icon: 'move-horizontal', run: () => fitWidth(d) }
    ] },
    { label: 'Cửa sổ', items: [
      { type: 'big', label: 'Cửa sổ Mới', icon: 'app-window', run: () => import('../core/host.js').then(m => m.default.newWindow()) },
      { type: 'drop', big: true, label: 'Chuyển Cửa sổ', icon: 'layers', menu: () => d.shell.docs.map(x => ({ label: x.name + ' — ' + x.meta.label, checked: x === d || undefined, run: () => d.shell.showDoc(x) })) }
    ] },
    { label: 'Android DeX', items: [
      { type: 'big', label: 'Cài đặt DeX', icon: 'monitor', title: 'Trung tâm Cài đặt & Điều khiển Android DeX', run: () => import('../core/dex.js').then(D => D.openDexControlCenterDialog()) },
      { type: 'col', items: [
        { type: 'btn', label: 'Bàn rê chuột ảo', icon: 'mouse-pointer', run: () => import('../core/dex.js').then(D => D.openVirtualTrackpad()) },
        { type: 'btn', label: 'Chia đôi màn hình', icon: 'columns', run: () => import('../core/dex.js').then(D => D.tileWindow('left')) }
      ] }
    ] }
  ] };
}
function fitPage(d) { const avail = d.scroller.clientHeight - 40; d.shell.setZoom(Math.floor(avail / d.pageH * 100)); }
function fitWidth(d) { const avail = d.scroller.clientWidth - 40; d.shell.setZoom(Math.floor(avail / d.pageW * 100)); }
