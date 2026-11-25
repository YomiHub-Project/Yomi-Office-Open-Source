// Shared insert dialogs: Symbols, Icons, QR code
import { h, ic, icHtml, esc, bytesToDataUrl } from './util.js';
import { dialog, toast } from './ui.js';
import ICONS from './icons.gen.js';
import qrcode from 'qrcode-generator';

const SYMBOL_SETS = {
  'Thông dụng': '©®™§¶†‡•…‰′″‹›«»“”‘’–—±×÷≈≠≤≥∞√∑∏∫∂∆∇πΩµ°′″℃℉€£¥₫$¢%‰№℮',
  'Mũi tên': '←↑→↓↔↕↖↗↘↙⇐⇑⇒⇓⇔⇕➔➜➤➥➨⟵⟶⟷↺↻⤴⤵',
  'Toán học': '±∓×÷√∛∜∞∝∑∏∐∫∬∭∮∂∆∇∈∉∋∌∩∪⊂⊃⊆⊇⊄⊅∅∀∃∄¬∧∨⊕⊗⊥∥∠∡∢≈≅≡≢≤≥≪≫≺≻⊢⊨ℕℤℚℝℂℵ',
  'Hy Lạp': 'αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ',
  'Hình học': '■□▪▫▲△▶▷▼▽◀◁◆◇○●◎◉★☆♠♣♥♦✓✔✗✘☐☑☒⬛⬜🔲🔳',
  'Tiền tệ': '₫$€£¥₩₹₽₺₴₦₱฿₪¢',
  'Tiếng Việt': 'ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬĐÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴàáảãạăằắẳẵặâầấẩẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵ',
  'Phân số & chỉ số': '½⅓⅔¼¾⅕⅖⅗⅘⅙⅚⅛⅜⅝⅞⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎',
  'Biểu tượng': '☎☏✆✉✂✎✏✐☀☁☂☃☄★☆☯☮☢☣♻⚠⚡⚙⚑⚐♀♂⚥☺☻☹♪♫♬'
};
export function symbolDialog(onPick) {
  let cur = Object.keys(SYMBOL_SETS)[0];
  const tabs = h('div.sym-tabs'); const grid = h('div.sym-grid'); const prev = h('div.sym-prev');
  const draw = () => { grid.innerHTML = ''; Array.from(SYMBOL_SETS[cur]).forEach(ch => { const b = h('button.sym', ch); b.title = 'U+' + ch.codePointAt(0).toString(16).toUpperCase(); b.addEventListener('click', () => { prev.textContent = ch + '  ' + b.title; onPick(ch); }); grid.appendChild(b); }); };
  Object.keys(SYMBOL_SETS).forEach(k => { const b = h('button.seg-b' + (k === cur ? '.on' : ''), k); b.addEventListener('click', () => { cur = k; Array.prototype.forEach.call(tabs.children, x => x.classList.toggle('on', x === b)); draw(); }); tabs.appendChild(b); });
  draw();
  return dialog({ title: 'Ký hiệu', width: 560, body: h('div', tabs, grid, prev, h('div.hint', 'Bấm vào ký hiệu để chèn. Có thể chèn nhiều lần.')), buttons: [{ label: 'Đóng', primary: true }] });
}
export function iconsDialog(onPick) {
  const names = Object.keys(ICONS);
  const search = h('input.inp', { placeholder: 'Tìm biểu tượng (tiếng Anh: home, user, chart…)' });
  const grid = h('div.icon-grid');
  let color = '#0f6cbd';
  const col = h('input', { type: 'color', value: color }); col.addEventListener('change', () => { color = col.value; draw(); });
  const draw = () => { grid.innerHTML = ''; const q = search.value.trim().toLowerCase(); names.filter(n => !q || n.indexOf(q) >= 0).slice(0, 400).forEach(n => { const b = h('button.icon-cell', { title: n, html: icHtml(n, 26) }); b.style.color = color; b.addEventListener('click', () => { const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 24 24" fill="none" stroke="' + color + '" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' + ICONS[n] + '</svg>'; onPick('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg)); document.querySelector('.dlg-back .dlg-x').click(); }); grid.appendChild(b); }); };
  search.addEventListener('input', draw); draw();
  return dialog({ title: 'Chèn biểu tượng', width: 640, body: h('div', h('div.form-row.inline', search, h('label.chk', col, h('span', ' Màu'))), grid), buttons: [{ label: 'Đóng' }] });
}
export function qrSvg(text, opts) {
  opts = opts || {};
  const qr = qrcode(0, opts.ecc || 'M'); qr.addData(unescape(encodeURIComponent(text)), 'Byte'); qr.make();
  const n = qr.getModuleCount(), m = 4; let d = '';
  for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) if (qr.isDark(r, c)) d += 'M' + (c + m) + ',' + (r + m) + 'h1v1h-1z';
  return '<svg xmlns="http://www.w3.org/2000/svg" width="' + (n + m * 2) * 8 + '" height="' + (n + m * 2) * 8 + '" viewBox="0 0 ' + (n + m * 2) + ' ' + (n + m * 2) + '" shape-rendering="crispEdges"><rect width="100%" height="100%" fill="' + (opts.bg || '#fff') + '"/><path d="' + d + '" fill="' + (opts.fg || '#000') + '"/></svg>';
}
export function qrDialog(onPick, initial) {
  const ta = h('textarea.inp', { rows: 3, placeholder: 'Nhập liên kết, văn bản, số điện thoại, thông tin Wi-Fi…' }); ta.value = initial || 'https://';
  const fg = h('input', { type: 'color', value: '#000000' }), bg = h('input', { type: 'color', value: '#ffffff' });
  const ecc = h('select.inp.small', ['L', 'M', 'Q', 'H'].map(x => h('option', { value: x }, 'Sửa lỗi ' + x))); ecc.value = 'M';
  const prev = h('div.qr-prev');
  const upd = () => { try { prev.innerHTML = qrSvg(ta.value || ' ', { fg: fg.value, bg: bg.value, ecc: ecc.value }); } catch (e) { prev.textContent = 'Nội dung quá dài cho mã QR'; } };
  [ta, fg, bg, ecc].forEach(x => x.addEventListener('input', upd)); ecc.addEventListener('change', upd); upd();
  return dialog({ title: 'Tạo mã QR', width: 460, body: h('div.form', ta, h('div.form-row.inline', h('label.chk', fg, h('span', ' Màu mã')), h('label.chk', bg, h('span', ' Màu nền')), ecc), prev, h('div.hint', 'Bộ mã hóa QR mã nguồn mở (MIT) — hoạt động ngoại tuyến.')),
    buttons: [{ label: 'Chèn', primary: true, run: () => { try { onPick('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(qrSvg(ta.value, { fg: fg.value, bg: bg.value, ecc: ecc.value }))); } catch (e) { toast('Không tạo được mã QR', 'err'); return false; } return true; } }, { label: 'Hủy', value: null }] });
}
