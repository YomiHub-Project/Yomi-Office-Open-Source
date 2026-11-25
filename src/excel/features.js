// Excel features: Format Cells, tables, conditional formatting, data validation, insert function, pivot, data tools, auditing, page setup & print
import { h, ic, esc, $, $$, clamp, deepClone, debounce, keyName } from '../core/util.js';
import { toast, showMenu, dialog, promptDlg, confirmDlg, formDlg, colorPicker, alertDlg, THEME_COLORS, STD_COLORS, closeMenus, showPopup } from '../core/ui.js';
import { settings } from '../core/store.js';
import { Err, E, key, colName, addr, rangeAddr, parseRangeAddr, norm, inRange, internStyle, mergeStyle, MAX_ROWS, MAX_COLS, dateToSerial } from './model.js';
import { FN, parse, cmp, formulaRefs, shiftFormula } from './formula.js';
import { format, parseInput, NUMBER_FORMATS, isDateFormat } from './numfmt.js';
import { functionCatalog, FUNCTION_META } from './functions.js';
import { sheetToHtml } from './xlsx.js';
import * as P8 from './pages.js';

const sel = doc => doc.wb.cur.sel;
const bound = doc => doc.boundingSel();
function clipRange(s, rg) { const u = s.used(); return { r1: rg.r1, c1: rg.c1, r2: rg.r2 >= MAX_ROWS - 1 ? Math.max(u.r2, rg.r1) : rg.r2, c2: rg.c2 >= MAX_COLS - 1 ? Math.max(u.c2, rg.c1) : rg.c2 }; }
function curSt(doc) { const s = doc.wb.cur; return s.styleAt(s.sel.r, s.sel.c) || {}; }
function apply(doc, patch, label) { return import('./cmds.js').then(m => m.setStyle(doc, patch, label)); }
import { FONT_LIST } from '../core/fonts.js';
export const FONTS = FONT_LIST;
export const SIZES = [8, 9, 10, 10.5, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

// ---------------------------------------------------------------- Format Cells dialog (Ctrl+1)
export function formatCellsDialog(doc, tab) {
  const st0 = Object.assign({}, curSt(doc)); const st = Object.assign({}, st0);
  const s = doc.wb.cur; const v0 = s.value(s.sel.r, s.sel.c);
  const tabs = [['number', 'Số'], ['align', 'Căn lề'], ['font', 'Phông chữ'], ['border', 'Viền'], ['fill', 'Tô nền'], ['protect', 'Bảo vệ']];
  const head = h('div.fc-tabs'); const page = h('div.fc-page');
  const show = id => { $$('.fc-tab', head).forEach(b => b.classList.toggle('on', b.dataset.t === id)); page.innerHTML = ''; page.appendChild(pages[id]()); };
  tabs.forEach(t => { const b = h('button.fc-tab', { dataset: { t: t[0] } }, t[1]); b.addEventListener('click', () => show(t[0])); head.appendChild(b); });
  const preview = () => { const el = $('.fc-sample', page); if (el) { const f = format(typeof v0 === 'number' ? v0 : 1234.5678, st.nf); el.textContent = f.t; el.style.color = f.c || st.color || '#000'; } };
  const pages = {
    number() {
      const cats = [['General', 'Chung'], ['number', 'Số'], ['currency', 'Tiền tệ'], ['accounting', 'Kế toán'], ['date', 'Ngày'], ['time', 'Thời gian'], ['percent', 'Phần trăm'], ['fraction', 'Phân số'], ['scientific', 'Khoa học'], ['text', 'Văn bản'], ['special', 'Đặc biệt'], ['custom', 'Tùy chỉnh']];
      const presets = {
        number: ['0', '0.00', '#,##0', '#,##0.00', '#,##0.00;[Red]-#,##0.00', '#,##0;(#,##0)'],
        currency: ['#,##0 "₫"', '#,##0.00 "₫"', '"$"#,##0.00', '"$"#,##0.00;[Red]-"$"#,##0.00', '#,##0.00 "€"', '"¥"#,##0'],
        accounting: ['_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)', '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)', '_-* #,##0 "₫"_-;-* #,##0 "₫"_-;_-* "-"?? "₫"_-;_-@_-'],
        date: ['dd/mm/yyyy', 'd/m/yyyy', 'dd-mm-yy', 'dddd, d mmmm yyyy', 'd mmmm yyyy', 'mmm-yy', 'yyyy-mm-dd', '"Ngày "dd" tháng "mm" năm "yyyy'],
        time: ['hh:mm', 'hh:mm:ss', 'h:mm AM/PM', 'h:mm:ss AM/PM', '[h]:mm:ss', 'dd/mm/yyyy hh:mm'],
        percent: ['0%', '0.0%', '0.00%'], fraction: ['# ?/?', '# ??/??', '# ?/4', '# ?/8', '# ?/10', '# ?/100'], scientific: ['0.00E+00', '0.0E+00', '0E+00'], text: ['@'],
        special: ['00000', '000-000-0000', '0000 000 000', '[<=9999999]###-####;(###) ###-####'],
        custom: NUMBER_FORMATS.map(x => x[0]).concat(['0.0', '#,##0.000', '0.00_);(0.00)', '"Tổng: "#,##0', '[Blue]#,##0;[Red]-#,##0;0'])
      };
      const list = h('div.fc-list'); const code = h('input.inp', { value: st.nf || 'General' });
      code.addEventListener('input', () => { st.nf = code.value === 'General' ? null : code.value; preview(); });
      const catList = h('div.fc-cats');
      const showCat = c => {
        $$('.fc-cat', catList).forEach(b => b.classList.toggle('on', b.dataset.c === c)); list.innerHTML = '';
        if (c === 'General') { st.nf = null; code.value = 'General'; preview(); return; }
        (presets[c] || []).forEach(p => { const it = h('div.fc-item' + (st.nf === p ? '.on' : ''), format(typeof v0 === 'number' ? (c === 'date' || c === 'time' ? (v0 > 1 ? v0 : 46312.5) : v0) : (c === 'date' || c === 'time' ? 46312.5 : -1234.5678), p).t + '   ', h('span.muted', p)); it.addEventListener('click', () => { st.nf = p; code.value = p; $$('.fc-item', list).forEach(x => x.classList.remove('on')); it.classList.add('on'); preview(); }); list.appendChild(it); });
      };
      cats.forEach(c => { const b = h('div.fc-cat', { dataset: { c: c[0] } }, c[1]); b.addEventListener('click', () => showCat(c[0])); catList.appendChild(b); });
      const box = h('div.fc-num', catList, h('div.fc-right', h('div.lbl', 'Mẫu'), h('div.fc-sample'), h('div.lbl', 'Loại / Mã định dạng'), code, list));
      setTimeout(() => { showCat(!st.nf ? 'General' : /%/.test(st.nf) ? 'percent' : isDateFormat(st.nf) ? (/[ymd]/.test(st.nf) ? 'date' : 'time') : /₫|\$|€/.test(st.nf) ? 'currency' : 'custom'); code.value = st.nf || 'General'; preview(); }, 0);
      return box;
    },
    align() {
      const ha = h('select.inp.small', [['general', 'Chung'], ['left', 'Trái (Thụt lề)'], ['center', 'Giữa'], ['right', 'Phải'], ['fill', 'Lấp đầy'], ['justify', 'Căn đều'], ['centerAcross', 'Giữa qua vùng chọn']].map(o => h('option', { value: o[0] }, o[1]))); ha.value = st.ha || 'general';
      const va = h('select.inp.small', [['top', 'Trên'], ['middle', 'Giữa'], ['bottom', 'Dưới']].map(o => h('option', { value: o[0] }, o[1]))); va.value = st.va || 'bottom';
      const ind = h('input.inp.small', { type: 'number', min: 0, max: 15, value: st.indent || 0 });
      const wrap = h('input', { type: 'checkbox' }); wrap.checked = !!st.wrap; const shrink = h('input', { type: 'checkbox' }); shrink.checked = !!st.shrink;
      const mergeC = h('input', { type: 'checkbox' }); const s2 = doc.wb.cur; mergeC.checked = !!s2.mergeAt(s2.sel.r, s2.sel.c);
      ha.addEventListener('change', () => { st.ha = ha.value === 'general' ? null : ha.value; }); va.addEventListener('change', () => { st.va = va.value === 'bottom' ? null : va.value; });
      ind.addEventListener('change', () => { st.indent = +ind.value || null; }); wrap.addEventListener('change', () => { st.wrap = wrap.checked || null; }); shrink.addEventListener('change', () => { st.shrink = shrink.checked || null; });
      mergeC.addEventListener('change', () => { st._merge = mergeC.checked ? 'merge' : 'unmerge'; });
      return h('div.form', h('div.form-sec', 'Căn lề văn bản'), h('div.form-row.inline', h('label.lbl', 'Ngang:'), ha), h('div.form-row.inline', h('label.lbl', 'Dọc:'), va), h('div.form-row.inline', h('label.lbl', 'Thụt lề:'), ind),
        h('div.form-sec', 'Điều khiển văn bản'), h('label.chk', wrap, h('span', 'Ngắt dòng văn bản (Wrap text)')), h('label.chk', shrink, h('span', 'Thu nhỏ cho vừa')), h('label.chk', mergeC, h('span', 'Gộp ô')));
    },
    font() {
      const fn = h('select.inp', FONTS.map(f => h('option', { value: f, style: { fontFamily: f } }, f))); fn.value = st.font || settings().excelFont || 'Calibri';
      const sz = h('select.inp.small', SIZES.map(z => h('option', { value: z }, z))); sz.value = st.sz || 11;
      const sty = h('select.inp.small', [['n', 'Thường'], ['i', 'Nghiêng'], ['b', 'Đậm'], ['bi', 'Đậm nghiêng']].map(o => h('option', { value: o[0] }, o[1]))); sty.value = (st.b ? 'b' : '') + (st.i ? 'i' : '') || 'n';
      const und = h('select.inp.small', [['', 'Không'], ['single', 'Đơn'], ['double', 'Kép']].map(o => h('option', { value: o[0] }, o[1]))); und.value = st.u || '';
      const strike = h('input', { type: 'checkbox' }); strike.checked = !!st.s;
      const col = h('input', { type: 'color', value: st.color || '#000000' });
      const smp = h('div.fc-fontsample', 'AaBbCc ÁÂĂ ợ 123');
      const upd = () => { st.font = fn.value; st.sz = +sz.value; st.b = /b/.test(sty.value) || null; st.i = /i/.test(sty.value) || null; st.u = und.value || null; st.s = strike.checked || null; st.color = col.value === '#000000' ? null : col.value; smp.style.fontFamily = fn.value; smp.style.fontSize = sz.value + 'pt'; smp.style.fontWeight = st.b ? 'bold' : ''; smp.style.fontStyle = st.i ? 'italic' : ''; smp.style.textDecoration = (st.u ? 'underline ' : '') + (st.s ? 'line-through' : ''); smp.style.color = col.value; };
      [fn, sz, sty, und, strike, col].forEach(x => x.addEventListener('change', upd)); setTimeout(upd, 0);
      return h('div.form', h('div.form-row', h('label.lbl', 'Phông chữ:'), fn), h('div.form-row.inline', h('label.lbl', 'Kiểu:'), sty), h('div.form-row.inline', h('label.lbl', 'Cỡ:'), sz), h('div.form-row.inline', h('label.lbl', 'Gạch chân:'), und), h('div.form-row.inline', h('label.lbl', 'Màu:'), col), h('label.chk', strike, h('span', 'Gạch ngang')), h('div.lbl', 'Xem trước'), smp);
    },
    border() {
      const bstyle = h('select.inp.small', [['thin', 'Mảnh'], ['medium', 'Vừa'], ['thick', 'Đậm'], ['dashed', 'Nét đứt'], ['dotted', 'Chấm'], ['double', 'Kép'], ['hair', 'Rất mảnh']].map(o => h('option', { value: o[0] }, o[1])));
      const bcol = h('input', { type: 'color', value: '#000000' });
      const box = h('div.fc-bbox');
      const bd = Object.assign({}, st.bd || {});
      const drawBox = () => { box.style.borderTop = bd.t ? '2px solid ' + bd.t.c : '1px dashed #ccc'; box.style.borderBottom = bd.b ? '2px solid ' + bd.b.c : '1px dashed #ccc'; box.style.borderLeft = bd.l ? '2px solid ' + bd.l.c : '1px dashed #ccc'; box.style.borderRight = bd.r ? '2px solid ' + bd.r.c : '1px dashed #ccc'; };
      const tog = side => { const v = { s: bstyle.value, c: bcol.value }; bd[side] = bd[side] ? null : v; st.bd = Object.assign({}, bd); st._bdTouched = true; drawBox(); };
      const btn = (side, label, icon) => h('button.btn', { onclick: () => tog(side) }, ic(icon, 14), ' ' + label);
      drawBox();
      return h('div.form', h('div.form-row.inline', h('label.lbl', 'Kiểu đường:'), bstyle), h('div.form-row.inline', h('label.lbl', 'Màu:'), bcol),
        h('div', h('button.btn', { onclick: () => { ['t', 'b', 'l', 'r'].forEach(k => { bd[k] = null; }); st.bd = {}; st._bdTouched = true; st._bdPreset = 'none'; drawBox(); } }, 'Không'), h('button.btn', { onclick: () => { st._bdPreset = 'outside'; st._bdStyle = { s: bstyle.value, c: bcol.value }; toast('Sẽ áp dụng viền ngoài'); } }, 'Viền ngoài'), h('button.btn', { onclick: () => { st._bdPreset = 'all'; st._bdStyle = { s: bstyle.value, c: bcol.value }; toast('Sẽ áp dụng viền trong + ngoài'); } }, 'Trong + ngoài')),
        h('div.fc-bwrap', btn('t', 'Trên', 'panel-top'), box, btn('b', 'Dưới', 'panel-bottom'), btn('l', 'Trái', 'panel-left'), btn('r', 'Phải', 'panel-right')));
    },
    fill() {
      const wrap = h('div.fc-fill');
      const cur = h('div.fc-fillcur', { style: { background: st.fill || '#fff' } }, 'Màu hiện tại');
      wrap.appendChild(h('button.btn', { onclick: () => { st.fill = null; cur.style.background = '#fff'; } }, 'Không màu'));
      THEME_COLORS.concat([STD_COLORS]).forEach(row => wrap.appendChild(h('div.cp-row', row.map(c => h('button.cp-sw', { style: { background: c }, onclick: () => { st.fill = c; cur.style.background = c; } })))));
      const more = h('input', { type: 'color', value: st.fill || '#ffffff' }); more.addEventListener('change', () => { st.fill = more.value; cur.style.background = more.value; });
      wrap.appendChild(h('label.chk', more, h('span', ' Màu khác…'))); wrap.appendChild(cur);
      return wrap;
    },
    protect() {
      const lk = h('input', { type: 'checkbox' }); lk.checked = !st.unlocked; const hid = h('input', { type: 'checkbox' }); hid.checked = !!st.hideF;
      lk.addEventListener('change', () => { st.unlocked = !lk.checked || null; }); hid.addEventListener('change', () => { st.hideF = hid.checked || null; });
      return h('div.form', h('label.chk', lk, h('span', 'Đã khóa (Locked)')), h('label.chk', hid, h('span', 'Ẩn công thức (Hidden)')), h('div.hint', 'Khóa ô hoặc ẩn công thức chỉ có hiệu lực khi bạn bảo vệ trang tính (Xem lại > Bảo vệ trang tính).'));
    }
  };
  show(tab || 'number');
  dialog({ title: 'Định dạng ô', width: 640, body: h('div.fc', head, page), cls: 'fc-dlg', buttons: [{ label: 'OK', primary: true, run: () => {
    const patch = {};
    ['nf', 'ha', 'va', 'indent', 'wrap', 'shrink', 'font', 'sz', 'b', 'i', 'u', 's', 'color', 'fill', 'unlocked', 'hideF'].forEach(k => { if (st[k] !== st0[k]) patch[k] = st[k] == null ? null : st[k]; });
    if (st._bdTouched && !st._bdPreset) patch.bd = Object.assign({ t: null, b: null, l: null, r: null }, st.bd);
    const ops = [];
    if (Object.keys(patch).length) ops.push(() => apply(doc, patch, 'Định dạng ô'));
    if (st._bdPreset) ops.push(() => import('./cmds.js').then(m => m.applyBorder(doc, st._bdPreset, st._bdStyle && st._bdStyle.c, st._bdStyle && st._bdStyle.s)));
    if (st._merge) ops.push(() => import('./cmds.js').then(m => m.merge(doc, st._merge)));
    ops.reduce((p, f) => p.then(f), Promise.resolve());
    return true;
  } }, { label: 'Hủy', value: null }] });
}

// ---------------------------------------------------------------- Table styles & Format as Table
export const TABLE_STYLES = [
  { id: 'Light1', name: 'Sáng 1', head: '#ffffff', headText: '#000', band: '#f2f2f2', line: '#000' },
  { id: 'Medium2', name: 'Xanh dương 2', head: '#4472c4', headText: '#fff', band: '#d9e1f2', line: '#8ea9db' },
  { id: 'Medium3', name: 'Cam 3', head: '#ed7d31', headText: '#fff', band: '#fce4d6', line: '#f4b084' },
  { id: 'Medium4', name: 'Xám 4', head: '#a5a5a5', headText: '#fff', band: '#ededed', line: '#c9c9c9' },
  { id: 'Medium5', name: 'Vàng 5', head: '#ffc000', headText: '#fff', band: '#fff2cc', line: '#ffd966' },
  { id: 'Medium6', name: 'Xanh nhạt 6', head: '#5b9bd5', headText: '#fff', band: '#ddebf7', line: '#9bc2e6' },
  { id: 'Medium7', name: 'Xanh lá 7', head: '#70ad47', headText: '#fff', band: '#e2efda', line: '#a9d08e' },
  { id: 'Dark1', name: 'Tối 1', head: '#000000', headText: '#fff', band: '#d9d9d9', line: '#595959' },
  { id: 'Dark9', name: 'Tối xanh 9', head: '#203764', headText: '#fff', band: '#b4c6e7', line: '#305496' },
  { id: 'Medium10', name: 'Đỏ 10', head: '#c00000', headText: '#fff', band: '#f8cbad', line: '#f4b084' },
  { id: 'Medium11', name: 'Tím 11', head: '#7030a0', headText: '#fff', band: '#e4d4f0', line: '#b18fd2' },
  { id: 'Visicons', name: 'VISICONS', head: '#1f3864', headText: '#fff', band: '#eef3fb', line: '#8ea9db' }
];
export function tableStylePreview(ts) {
  const t = h('div.ts-prev');
  for (let r = 0; r < 4; r++) { const row = h('div.ts-row'); for (let c = 0; c < 4; c++) row.appendChild(h('span', { style: { background: r === 0 ? ts.head : (r % 2 ? ts.band : '#fff'), borderBottom: '1px solid ' + ts.line } })); t.appendChild(row); }
  return t;
}
export function formatAsTable(doc, styleId) {
  const s = doc.wb.cur;
  import('./cmds.js').then(m => {
    let rg = m.currentRegion(doc);
    if (s.tables.some(t => t.r1 <= rg.r2 && t.r2 >= rg.r1 && t.c1 <= rg.c2 && t.c2 >= rg.c1)) { const t = s.tables.find(t => inRange(t, s.sel.r, s.sel.c)); if (t && styleId) { doc.change('Kiểu bảng', s, 'sheet', () => { t.style = styleId; applyTableStyle(s, t); }); } else toast('Vùng này đã là Bảng'); return; }
    const go = sid => {
      doc.change('Định dạng thành Bảng', s, 'sheet', () => {
        // ensure headers
        for (let c = rg.c1; c <= rg.c2; c++) { const v = s.value(rg.r1, c); if (v == null || v === '') s.setCell(rg.r1, c, { v: 'Cột' + (c - rg.c1 + 1) }); }
        const name = 'Bảng' + (doc.wb.sheets.reduce((n, x) => n + x.tables.length, 0) + 1);
        const t = { name, r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, style: sid, banded: true, header: true, totals: false, filter: true };
        s.tables.push(t);
        applyTableStyle(s, t);
        if (!s.filter) s.filter = { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, crit: {} };
      });
      toast('Đã tạo Bảng ' + rangeAddr(rg) + ' — tab Thiết kế Bảng đã xuất hiện', 'ok');
    };
    if (styleId) go(styleId); else showMenu([{ header: 'Chọn kiểu bảng' }].concat(TABLE_STYLES.map(ts => ({ label: ts.name, swatch: ts.head, run: () => go(ts.id) }))), { x: 300, y: 160 });
  });
}
export function applyTableStyle(s, t) {
  const ts = TABLE_STYLES.find(x => x.id === t.style) || TABLE_STYLES[1];
  for (let r = t.r1; r <= t.r2; r++) for (let c = t.c1; c <= t.c2; c++) {
    const x = s.get(r, c); const base = x && x.s;
    const isHead = r === t.r1 && t.header !== false, isTot = t.totals && r === t.r2;
    let patch;
    if (isHead) patch = { fill: ts.head, color: ts.headText, b: true, bd: { b: { s: 'thin', c: ts.line } } };
    else if (isTot) patch = { b: true, bd: { t: { s: 'double', c: ts.line } }, fill: null };
    else patch = { fill: t.banded !== false && (r - t.r1) % 2 === 0 ? ts.band : null, color: null, b: null, bd: { b: { s: 'thin', c: ts.line } } };
    s.setCell(r, c, { s: mergeStyle(base, patch) });
  }
}
export function tableDesignTab(doc) {
  const s = doc.wb.cur; const t = () => s.tables.find(x => inRange(x, s.sel.r, s.sel.c));
  const toggle = (k, label) => ({ type: 'btn', label, icon: 'square-check', active: () => t() && t()[k] !== false && !!t()[k], run: () => { const tb = t(); if (!tb) return; doc.change(label, s, 'sheet', () => { tb[k] = !tb[k]; if (k === 'totals') { if (tb.totals) { tb.r2++; s.setCell(tb.r2, tb.c1, { v: 'Tổng' }); for (let c = tb.c1 + 1; c <= tb.c2; c++) if (typeof s.value(tb.r2 - 1, c) === 'number') s.setCell(tb.r2, c, { f: 'SUBTOTAL(109,' + rangeAddr({ r1: tb.r1 + 1, c1: c, r2: tb.r2 - 1, c2: c }) + ')' }); } else { for (let c = tb.c1; c <= tb.c2; c++) s.setCell(tb.r2, c, { v: null, f: null, s: null }); tb.r2--; } } if (k === 'filter') { s.filter = tb.filter ? { r1: tb.r1, c1: tb.c1, r2: tb.r2, c2: tb.c2, crit: {} } : null; } applyTableStyle(s, tb); }); } });
  return { id: 'tabledesign', label: 'Thiết kế Bảng', en: 'Table Design', contextual: true, groups: [
    { label: 'Thuộc tính', items: [{ type: 'custom', render: () => h('div.rb-col', h('div', { style: { fontSize: '11px' } }, 'Tên bảng:'), (() => { const i = h('input.rb-combo-inp', { value: t() ? t().name : '', style: { width: '110px', border: '1px solid #ccc' } }); i.addEventListener('change', () => { const tb = t(); if (tb) { tb.name = i.value.replace(/\s+/g, '_'); doc.markDirty(); } }); return i; })()) }, { type: 'btn', label: 'Đổi kích thước bảng', icon: 'scaling', run: async () => { const tb = t(); if (!tb) return; const v = await promptDlg('Vùng mới cho bảng:', rangeAddr(tb), 'Đổi kích thước bảng'); const rg = v && parseRangeAddr(v); if (rg) doc.change('Đổi kích thước bảng', s, 'sheet', () => { Object.assign(tb, rg); applyTableStyle(s, tb); }); } }] },
    { label: 'Công cụ', items: [{ type: 'btn', label: 'Tóm tắt bằng PivotTable', icon: 'table-2', run: () => pivotDialog(doc) }, { type: 'btn', label: 'Loại bỏ trùng lặp', icon: 'copy-minus', run: () => removeDuplicates(doc) }, { type: 'btn', label: 'Chuyển thành vùng', icon: 'grid-3x3', run: () => { const tb = t(); if (tb) doc.change('Chuyển thành vùng', s, 'props', () => { s.tables = s.tables.filter(x => x !== tb); }); } }] },
    { label: 'Tùy chọn kiểu bảng', items: [{ type: 'col', items: [toggle('header', 'Hàng tiêu đề'), toggle('totals', 'Hàng tổng'), toggle('banded', 'Hàng xen kẽ')] }, { type: 'col', items: [toggle('filter', 'Nút lọc')] }] },
    { label: 'Kiểu bảng', items: [{ type: 'gallery', width: 380, items: () => TABLE_STYLES.map(ts => ({ title: ts.name, preview: () => tableStylePreview(ts), active: () => t() && t().style === ts.id, run: () => formatAsTable(doc, ts.id) })) }] }
  ] };
}

// ---------------------------------------------------------------- cell styles gallery
export const CELL_STYLES = [
  ['Bình thường', {}], ['Tốt', { fill: '#c6efce', color: '#006100' }], ['Xấu', { fill: '#ffc7ce', color: '#9c0006' }], ['Trung tính', { fill: '#ffeb9c', color: '#9c5700' }],
  ['Tính toán', { fill: '#f2f2f2', color: '#fa7d00', b: true, bd: { t: { s: 'thin', c: '#7f7f7f' }, b: { s: 'thin', c: '#7f7f7f' }, l: { s: 'thin', c: '#7f7f7f' }, r: { s: 'thin', c: '#7f7f7f' } } }],
  ['Ô kiểm tra', { fill: '#a5a5a5', color: '#fff', b: true }], ['Văn bản giải thích', { i: true, color: '#7f7f7f' }], ['Nhập liệu', { fill: '#ffcc99', color: '#3f3f76' }], ['Ô liên kết', { color: '#fa7d00', bd: { b: { s: 'double', c: '#ff8001' } } }], ['Ghi chú', { fill: '#ffffcc', bd: { t: { s: 'thin', c: '#b2b2b2' }, b: { s: 'thin', c: '#b2b2b2' }, l: { s: 'thin', c: '#b2b2b2' }, r: { s: 'thin', c: '#b2b2b2' } } }], ['Văn bản cảnh báo', { color: '#ff0000' }],
  ['Tiêu đề', { sz: 18, color: '#44546a', font: 'Calibri Light' }], ['Tiêu đề 1', { sz: 15, b: true, color: '#44546a', bd: { b: { s: 'thick', c: '#4472c4' } } }], ['Tiêu đề 2', { sz: 13, b: true, color: '#44546a', bd: { b: { s: 'thick', c: '#a2b8e1' } } }], ['Tiêu đề 3', { b: true, color: '#44546a', bd: { b: { s: 'medium', c: '#8ea9db' } } }], ['Tổng', { b: true, bd: { t: { s: 'thin', c: '#4472c4' }, b: { s: 'double', c: '#4472c4' } } }],
  ['Nhấn 1', { fill: '#4472c4', color: '#fff' }], ['Nhấn 2', { fill: '#ed7d31', color: '#fff' }], ['Nhấn 3', { fill: '#a5a5a5', color: '#fff' }], ['Nhấn 4', { fill: '#ffc000', color: '#fff' }], ['Nhấn 5', { fill: '#5b9bd5', color: '#fff' }], ['Nhấn 6', { fill: '#70ad47', color: '#fff' }],
  ['Dấu phẩy', { nf: '#,##0.00' }], ['Tiền tệ', { nf: '#,##0 "₫"' }], ['Phần trăm', { nf: '0%' }]
];
export function cellStylesMenu(doc) {
  return CELL_STYLES.map(([n, st]) => ({ label: n, el: null, swatch: st.fill || '#fff', style: (st.b ? 'font-weight:bold;' : '') + (st.i ? 'font-style:italic;' : '') + (st.color ? 'color:' + st.color + ';' : ''), run: () => apply(doc, n === 'Bình thường' ? { fill: null, color: null, b: null, i: null, bd: null, sz: null, font: null, nf: null } : Object.assign({ fill: null, color: null, b: null, i: null }, st), 'Kiểu ô: ' + n) }));
}

// ---------------------------------------------------------------- Conditional formatting
// rule: {type, range, op, v1, v2, fill, color, b, text, formula, stopIfTrue, ...}
export function cfStyleImpl(doc, s, r, c, v) {
  let out = null;
  for (const rule of s.cf) {
    if (!inRange(rule.range, r, c)) continue;
    const res = evalRule(doc, s, rule, r, c, v);
    if (res) { out = Object.assign(out || {}, res); if (rule.stopIfTrue) break; }
  }
  return out;
}
function rangeStats(doc, s, rule) {
  if (rule._stats && rule._rev === s.rev + ':' + doc.wb.rev) return rule._stats;
  const vals = []; const R = rule.range; const u = s.used();
  for (let r = R.r1; r <= Math.min(R.r2, u.r2); r++) for (let c = R.c1; c <= Math.min(R.c2, u.c2); c++) { const v = s.value(r, c); if (typeof v === 'number') vals.push(v); }
  const sorted = vals.slice().sort((a, b) => a - b);
  const st = { min: sorted[0], max: sorted[sorted.length - 1], avg: vals.reduce((a, b) => a + b, 0) / (vals.length || 1), sorted, n: vals.length, sd: 0 };
  if (vals.length > 1) { st.sd = Math.sqrt(vals.reduce((a, b) => a + (b - st.avg) * (b - st.avg), 0) / (vals.length - 1)); }
  const counts = new Map(); for (let r = R.r1; r <= Math.min(R.r2, u.r2); r++) for (let c = R.c1; c <= Math.min(R.c2, u.c2); c++) { const v = s.value(r, c); if (v == null || v === '') continue; const k = typeof v === 'string' ? v.toLowerCase() : v; counts.set(k, (counts.get(k) || 0) + 1); }
  st.counts = counts;
  rule._stats = st; rule._rev = s.rev + ':' + doc.wb.rev;
  return st;
}
function mix(c1, c2, t) { const p = x => [parseInt(x.slice(1, 3), 16), parseInt(x.slice(3, 5), 16), parseInt(x.slice(5, 7), 16)]; const a = p(c1), b = p(c2); return '#' + a.map((v, i) => ('0' + Math.round(v + (b[i] - v) * t).toString(16)).slice(-2)).join(''); }
function evalRule(doc, s, rule, r, c, v) {
  const fmt = () => ({ fill: rule.fill || null, color: rule.color || null, b: rule.b || null });
  const n = typeof v === 'number' ? v : null;
  const val = x => { if (typeof x === 'string' && /^=/.test(x)) return doc.calc.evalFormula(shiftFormula(x.slice(1), r - rule.range.r1, c - rule.range.c1), s, r, c); const p = parseInput(String(x)); return p.v; };
  switch (rule.type) {
    case 'cell': {
      if (v == null || v === '') return null;
      const a = val(rule.v1), b = rule.v2 != null ? val(rule.v2) : null;
      const k = cmp(v, a);
      const ok = rule.op === 'gt' ? k > 0 : rule.op === 'lt' ? k < 0 : rule.op === 'ge' ? k >= 0 : rule.op === 'le' ? k <= 0 : rule.op === 'eq' ? k === 0 : rule.op === 'ne' ? k !== 0 : rule.op === 'between' ? (cmp(v, a) >= 0 && cmp(v, b) <= 0) : rule.op === 'notbetween' ? !(cmp(v, a) >= 0 && cmp(v, b) <= 0) : false;
      return ok ? fmt() : null;
    }
    case 'text': { if (v == null) return null; const t = String(v).toLowerCase(), q = String(rule.text || '').toLowerCase(); const ok = rule.op === 'contains' ? t.indexOf(q) >= 0 : rule.op === 'notcontains' ? t.indexOf(q) < 0 : rule.op === 'begins' ? t.indexOf(q) === 0 : t.slice(-q.length) === q; return ok ? fmt() : null; }
    case 'date': { if (n == null) return null; const today = Math.floor((Date.now() - new Date().getTimezoneOffset() * 60000) / 864e5) + 25569; const d = Math.floor(n); const ok = rule.op === 'today' ? d === today : rule.op === 'yesterday' ? d === today - 1 : rule.op === 'tomorrow' ? d === today + 1 : rule.op === 'last7' ? d > today - 7 && d <= today : rule.op === 'thisweek' ? Math.floor((d + 1) / 7) === Math.floor((today + 1) / 7) : false; return ok ? fmt() : null; }
    case 'duplicate': case 'unique': { if (v == null || v === '') return null; const st = rangeStats(doc, s, rule); const cnt = st.counts.get(typeof v === 'string' ? v.toLowerCase() : v) || 0; return (rule.type === 'duplicate' ? cnt > 1 : cnt === 1) ? fmt() : null; }
    case 'top': case 'bottom': { if (n == null) return null; const st = rangeStats(doc, s, rule); let k = rule.n || 10; if (rule.percent) k = Math.max(1, Math.floor(st.n * k / 100)); const thr = rule.type === 'top' ? st.sorted[st.n - k] : st.sorted[k - 1]; return (rule.type === 'top' ? n >= thr : n <= thr) ? fmt() : null; }
    case 'aboveavg': case 'belowavg': { if (n == null) return null; const st = rangeStats(doc, s, rule); return (rule.type === 'aboveavg' ? n > st.avg : n < st.avg) ? fmt() : null; }
    case 'formula': { const res = doc.calc.evalFormula(shiftFormula(String(rule.formula).replace(/^=/, ''), r - rule.range.r1, c - rule.range.c1), s, r, c); return res === true || (typeof res === 'number' && res !== 0) ? fmt() : null; }
    case 'blank': return (v == null || v === '') ? fmt() : null;
    case 'error': return v instanceof Err ? fmt() : null;
    case 'databar': { if (n == null) return null; const st = rangeStats(doc, s, rule); const lo = Math.min(0, st.min), hi = st.max; return { bar: { ratio: hi === lo ? 1 : (n - lo) / (hi - lo), color: rule.color || '#638ec6' } }; }
    case 'scale2': case 'scale3': {
      if (n == null) return null; const st = rangeStats(doc, s, rule); const cols = rule.colors || (rule.type === 'scale3' ? ['#f8696b', '#ffeb84', '#63be7b'] : ['#ffffff', '#63be7b']);
      if (st.max === st.min) return { fill: cols[cols.length - 1] };
      if (rule.type === 'scale2') return { fill: mix(cols[0], cols[1], (n - st.min) / (st.max - st.min)) };
      const mid = st.sorted[Math.floor((st.n - 1) / 2)];
      return { fill: n <= mid ? mix(cols[0], cols[1], mid === st.min ? 1 : (n - st.min) / (mid - st.min)) : mix(cols[1], cols[2], st.max === mid ? 1 : (n - mid) / (st.max - mid)) };
    }
    case 'icons': { if (n == null) return null; const st = rangeStats(doc, s, rule); const t = st.max === st.min ? 1 : (n - st.min) / (st.max - st.min); return { color: t >= 0.67 ? '#00b050' : t >= 0.33 ? '#ffc000' : '#c00000', icon: t >= 0.67 ? 'up' : t >= 0.33 ? 'mid' : 'down' }; }
  }
  return null;
}
export function addCfRule(doc, rule, rg) {
  const s = doc.wb.cur; rg = rg || clipRange(s, doc.grid.selRange());
  doc.change('Định dạng có điều kiện', s, 'props', () => { s.cf.push(Object.assign({ range: rg }, rule)); });
}
const PRESET_FMT = [['Tô đỏ nhạt, chữ đỏ đậm', { fill: '#ffc7ce', color: '#9c0006' }], ['Tô vàng, chữ vàng đậm', { fill: '#ffeb9c', color: '#9c5700' }], ['Tô xanh, chữ xanh đậm', { fill: '#c6efce', color: '#006100' }], ['Tô đỏ nhạt', { fill: '#ffc7ce' }], ['Chữ đỏ', { color: '#9c0006' }], ['Viền đỏ (đậm)', { color: '#c00000', b: true }]];
async function cfRuleDialog(doc, base, title, fields) {
  const v = await formDlg(title, fields.concat([{ key: 'fmt', type: 'select', label: 'Định dạng với:', options: PRESET_FMT.map((p, i) => [String(i), p[0]]), value: '0' }]));
  if (!v) return;
  const f = PRESET_FMT[+v.fmt][1];
  addCfRule(doc, Object.assign({}, base, { v1: v.v1, v2: v.v2, text: v.text, n: v.n, percent: v.percent, formula: v.formula, op: v.op || base.op }, f));
}
export function cfMenu(doc) {
  const cell = (op, label, two) => ({ label, run: () => cfRuleDialog(doc, { type: 'cell', op }, label, two ? [{ key: 'v1', label: 'Từ:' }, { key: 'v2', label: 'Đến:' }] : [{ key: 'v1', label: 'Định dạng ô ' + label.toLowerCase() + ':' }]) });
  return [
    { label: 'Tô sáng quy tắc ô', icon: 'highlighter', sub: [cell('gt', 'Lớn hơn…'), cell('lt', 'Nhỏ hơn…'), cell('between', 'Giữa…', true), cell('eq', 'Bằng…'), { label: 'Văn bản có chứa…', run: () => cfRuleDialog(doc, { type: 'text', op: 'contains' }, 'Văn bản có chứa', [{ key: 'text', label: 'Định dạng ô chứa văn bản:' }]) }, { label: 'Ngày xảy ra…', run: () => cfRuleDialog(doc, { type: 'date' }, 'Ngày xảy ra', [{ key: 'op', type: 'select', label: 'Khi ngày là:', options: [['today', 'Hôm nay'], ['yesterday', 'Hôm qua'], ['tomorrow', 'Ngày mai'], ['last7', 'Trong 7 ngày qua'], ['thisweek', 'Tuần này']], value: 'today' }]) }, { label: 'Giá trị trùng lặp…', run: () => cfRuleDialog(doc, { type: 'duplicate' }, 'Giá trị trùng lặp', []) }] },
    { label: 'Quy tắc trên/dưới', icon: 'arrow-up-down', sub: [{ label: '10 mục đầu…', run: () => cfRuleDialog(doc, { type: 'top' }, '10 mục đầu', [{ key: 'n', type: 'number', label: 'Số mục:', value: 10 }]) }, { label: '10% đầu…', run: () => cfRuleDialog(doc, { type: 'top', percent: true }, '10% đầu', [{ key: 'n', type: 'number', label: '%:', value: 10 }]) }, { label: '10 mục cuối…', run: () => cfRuleDialog(doc, { type: 'bottom' }, '10 mục cuối', [{ key: 'n', type: 'number', label: 'Số mục:', value: 10 }]) }, { label: 'Trên trung bình…', run: () => cfRuleDialog(doc, { type: 'aboveavg' }, 'Trên trung bình', []) }, { label: 'Dưới trung bình…', run: () => cfRuleDialog(doc, { type: 'belowavg' }, 'Dưới trung bình', []) }] },
    { label: 'Thanh dữ liệu', icon: 'bar-chart-horizontal', sub: [['#638ec6', 'Xanh dương'], ['#63c384', 'Xanh lá'], ['#ff555a', 'Đỏ'], ['#ffb628', 'Cam'], ['#008aef', 'Xanh da trời'], ['#d6007b', 'Tím']].map(x => ({ label: x[1], swatch: x[0], run: () => addCfRule(doc, { type: 'databar', color: x[0] }) })) },
    { label: 'Thang màu', icon: 'palette', sub: [{ label: 'Xanh - Vàng - Đỏ', run: () => addCfRule(doc, { type: 'scale3', colors: ['#63be7b', '#ffeb84', '#f8696b'] }) }, { label: 'Đỏ - Vàng - Xanh', run: () => addCfRule(doc, { type: 'scale3', colors: ['#f8696b', '#ffeb84', '#63be7b'] }) }, { label: 'Trắng - Xanh lá', run: () => addCfRule(doc, { type: 'scale2', colors: ['#ffffff', '#63be7b'] }) }, { label: 'Trắng - Đỏ', run: () => addCfRule(doc, { type: 'scale2', colors: ['#ffffff', '#f8696b'] }) }, { label: 'Xanh dương - Trắng - Đỏ', run: () => addCfRule(doc, { type: 'scale3', colors: ['#5a8ac6', '#fcfcff', '#f8696b'] }) }] },
    { label: 'Bộ biểu tượng (màu theo mức)', icon: 'traffic-cone', run: () => addCfRule(doc, { type: 'icons' }) },
    { sep: true },
    { label: 'Quy tắc mới (dùng công thức)…', icon: 'plus', run: () => cfRuleDialog(doc, { type: 'formula' }, 'Quy tắc định dạng mới', [{ key: 'formula', label: 'Định dạng các giá trị khi công thức này đúng (vd =$C2>100):', value: '=' }]) },
    { label: 'Xóa quy tắc', icon: 'eraser', sub: [{ label: 'Xóa quy tắc khỏi ô đã chọn', run: () => { const s = doc.wb.cur; const rg = doc.grid.selRange(); doc.change('Xóa quy tắc', s, 'props', () => { s.cf = s.cf.filter(r => !(r.range.r1 <= rg.r2 && r.range.r2 >= rg.r1 && r.range.c1 <= rg.c2 && r.range.c2 >= rg.c1)); }); } }, { label: 'Xóa quy tắc khỏi toàn bộ trang tính', run: () => { const s = doc.wb.cur; doc.change('Xóa quy tắc', s, 'props', () => { s.cf = []; }); } }] },
    { label: 'Quản lý quy tắc…', icon: 'list-checks', run: () => cfManager(doc) }
  ];
}
function cfManager(doc) {
  const s = doc.wb.cur; const body = h('div');
  const names = { cell: 'Giá trị ô', text: 'Văn bản', date: 'Ngày', duplicate: 'Trùng lặp', unique: 'Duy nhất', top: 'Trên cùng', bottom: 'Dưới cùng', aboveavg: 'Trên TB', belowavg: 'Dưới TB', formula: 'Công thức', databar: 'Thanh dữ liệu', scale2: 'Thang 2 màu', scale3: 'Thang 3 màu', icons: 'Biểu tượng', blank: 'Ô trống', error: 'Lỗi' };
  const draw = () => { body.innerHTML = ''; if (!s.cf.length) body.appendChild(h('div.muted', 'Trang tính này chưa có quy tắc nào.')); s.cf.forEach((r, i) => body.appendChild(h('div.cf-row', h('span.swatch', { style: { background: r.fill || r.color || '#eee' } }), h('span', names[r.type] + (r.op ? ' ' + r.op : '') + (r.v1 != null ? ' ' + r.v1 : '') + (r.text ? ' "' + r.text + '"' : '') + (r.formula ? ' ' + r.formula : '')), h('span.muted', rangeAddr(r.range)), h('button.btn', { onclick: () => { if (i > 0) { doc.change('Đổi thứ tự quy tắc', s, 'props', () => { const t = s.cf[i]; s.cf[i] = s.cf[i - 1]; s.cf[i - 1] = t; }); draw(); } } }, ic('arrow-up', 13)), h('button.btn', { onclick: () => { doc.change('Xóa quy tắc', s, 'props', () => s.cf.splice(i, 1)); draw(); } }, ic('trash-2', 13))))); };
  draw();
  dialog({ title: 'Trình quản lý quy tắc định dạng có điều kiện', width: 620, body, buttons: [{ label: 'Đóng', primary: true }] });
}

// ---------------------------------------------------------------- Data validation
export function dvCheck(doc, dv, text, s, r, c) {
  const p = parseInput(text); const v = p.v;
  const num = x => { const q = /^=/.test(String(x)) ? doc.calc.evalFormula(String(x).slice(1), s, r, c) : parseInput(String(x)).v; return typeof q === 'number' ? q : NaN; };
  switch (dv.type) {
    case 'list': { const items = dvItems(doc, dv, s, r, c).map(x => String(x).toLowerCase()); return items.indexOf(String(v).toLowerCase()) >= 0; }
    case 'whole': if (typeof v !== 'number' || v % 1) return false; return cmpOp(v, dv.op, num(dv.v1), num(dv.v2));
    case 'decimal': case 'date': case 'time': if (typeof v !== 'number') return false; return cmpOp(v, dv.op, num(dv.v1), num(dv.v2));
    case 'length': return cmpOp(String(v).length, dv.op, num(dv.v1), num(dv.v2));
    case 'custom': { const res = doc.calc.evalFormula(String(dv.formula || '').replace(/^=/, ''), s, r, c); return res === true || (typeof res === 'number' && res !== 0); }
  }
  return true;
}
function cmpOp(v, op, a, b) { switch (op) { case 'between': return v >= a && v <= b; case 'notbetween': return v < a || v > b; case 'eq': return v === a; case 'ne': return v !== a; case 'gt': return v > a; case 'lt': return v < a; case 'ge': return v >= a; case 'le': return v <= a; } return true; }
export function dvItems(doc, dv, s, r, c) {
  const src = String(dv.source || '');
  if (/^=/.test(src)) { const res = doc.calc.evalFormula(src.slice(1), s, r, c); if (Array.isArray(res)) return [].concat.apply([], res).filter(x => x != null && x !== ''); return res == null ? [] : [res]; }
  return src.split(/[,;]/).map(x => x.trim()).filter(Boolean);
}
export async function dvDialog(doc) {
  const s = doc.wb.cur; const rg = clipRange(s, doc.grid.selRange());
  const ex = s.dv.find(d => inRange(d.range, s.sel.r, s.sel.c)) || {};
  const v = await formDlg('Xác thực dữ liệu (Data Validation)', [
    { type: 'section', label: 'Thiết đặt' },
    { key: 'type', type: 'select', label: 'Cho phép:', value: ex.type || 'any', options: [['any', 'Mọi giá trị'], ['whole', 'Số nguyên'], ['decimal', 'Số thập phân'], ['list', 'Danh sách'], ['date', 'Ngày'], ['time', 'Thời gian'], ['length', 'Độ dài văn bản'], ['custom', 'Tùy chỉnh (công thức)']] },
    { key: 'op', type: 'select', label: 'Dữ liệu:', value: ex.op || 'between', options: [['between', 'giữa'], ['notbetween', 'không giữa'], ['eq', 'bằng'], ['ne', 'khác'], ['gt', 'lớn hơn'], ['lt', 'nhỏ hơn'], ['ge', 'lớn hơn hoặc bằng'], ['le', 'nhỏ hơn hoặc bằng']] },
    { key: 'v1', label: 'Tối thiểu / Giá trị:', value: ex.v1 || '' }, { key: 'v2', label: 'Tối đa:', value: ex.v2 || '' },
    { key: 'source', label: 'Nguồn danh sách (vd: Có,Không hoặc =$A$1:$A$10):', value: ex.source || '' },
    { key: 'formula', label: 'Công thức (Tùy chỉnh):', value: ex.formula || '' },
    { type: 'section', label: 'Thông báo nhập & cảnh báo lỗi' },
    { key: 'inMsg', label: 'Thông báo khi chọn ô:', value: ex.inMsg || '' },
    { key: 'errStyle', type: 'select', label: 'Kiểu cảnh báo:', value: ex.errStyle || 'stop', options: [['stop', 'Dừng (Stop)'], ['warning', 'Cảnh báo'], ['info', 'Thông tin']] },
    { key: 'errMsg', label: 'Thông báo lỗi:', value: ex.errMsg || '' }
  ], { width: 520 });
  if (!v) return;
  doc.change('Xác thực dữ liệu', s, 'props', () => {
    s.dv = s.dv.filter(d => !(d.range.r1 <= rg.r2 && d.range.r2 >= rg.r1 && d.range.c1 <= rg.c2 && d.range.c2 >= rg.c1));
    if (v.type !== 'any') s.dv.push({ range: rg, type: v.type, op: v.op, v1: v.v1, v2: v.v2, source: v.source, formula: v.formula, inMsg: v.inMsg, errStyle: v.errStyle, errMsg: v.errMsg });
  });
}
export function circleInvalid(doc) {
  const s = doc.wb.cur; let n = 0; const bad = [];
  s.dv.forEach(dv => { const R = clipRange(s, dv.range); for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) { const v = s.value(r, c); if (v == null) continue; if (!dvCheck(doc, dv, format(v, (s.styleAt(r, c) || {}).nf).t, s, r, c) && !dvCheck(doc, dv, String(v), s, r, c)) { bad.push({ r1: r, c1: c, r2: r, c2: c }); n++; } } });
  if (!n) return toast('Không có dữ liệu không hợp lệ', 'ok');
  s.sel.ranges = bad; s.sel.r = bad[0].r1; s.sel.c = bad[0].c1; doc.onSelection(); doc.grid.redraw();
  toast('Đã chọn ' + n + ' ô dữ liệu không hợp lệ', 'err');
}

// ---------------------------------------------------------------- Insert Function dialog
let fnDescCache = null;
function fnDescriptions() {
  if (fnDescCache) return fnDescCache;
  fnDescCache = {};
  return fnDescCache;
}
export function insertFunctionDialog(doc, preset) {
  const cat = functionCatalog(fnDescriptions());
  const cats = ['Tất cả', 'Dùng gần đây'].concat(Array.from(new Set(cat.map(f => f.cat))).sort());
  const recent = settings().xlRecentFns || ['SUM', 'AVERAGE', 'IF', 'VLOOKUP', 'XLOOKUP', 'COUNTIF', 'SUMIF', 'MAX', 'MIN', 'ROUND'];
  const search = h('input.inp', { placeholder: 'Nhập mô tả ngắn hoặc tên hàm, rồi bấm Đi (vd: tổng, tra cứu, ngày)' });
  const catSel = h('select.inp', cats.map(c => h('option', { value: c }, c))); catSel.value = preset || 'Dùng gần đây';
  const list = h('div.fn-list'); const info = h('div.fn-info');
  let chosen = null;
  const draw = () => {
    const q = search.value.trim().toLowerCase();
    let items = catSel.value === 'Tất cả' ? cat : catSel.value === 'Dùng gần đây' ? recent.map(n => cat.find(f => f.name === n)).filter(Boolean) : cat.filter(f => f.cat === catSel.value);
    if (q) items = cat.filter(f => f.name.toLowerCase().indexOf(q) >= 0 || (f.desc || '').toLowerCase().indexOf(q) >= 0);
    list.innerHTML = '';
    items.slice(0, 600).forEach((f, i) => { const it = h('div.fn-item', f.name); it.addEventListener('click', () => pick(f, it)); it.addEventListener('dblclick', () => { pick(f, it); done(); }); list.appendChild(it); if (i === 0) pick(f, it); });
    if (!items.length) info.textContent = 'Không tìm thấy hàm.';
  };
  const pick = (f, el) => { chosen = f; $$('.fn-item', list).forEach(x => x.classList.remove('on')); el.classList.add('on'); info.innerHTML = '<b>' + f.name + '(' + esc(f.sig) + ')</b><br>' + esc(f.desc || '') + '<br><span class="muted">Nhóm: ' + esc(f.cat) + '</span>'; };
  search.addEventListener('input', debounce(draw, 120)); catSel.addEventListener('change', () => { search.value = ''; draw(); });
  const total = h('div.muted', 'Thư viện có ' + cat.length + ' hàm Excel, tất cả tính toán ngoại tuyến.');
  const body = h('div.form', h('label.lbl', 'Tìm hàm:'), search, h('label.lbl', 'Hoặc chọn nhóm:'), catSel, h('label.lbl', 'Chọn hàm:'), list, info, total);
  let closeFn = null;
  const done = () => {
    if (!chosen) return;
    const st = settings(); st.xlRecentFns = [chosen.name].concat(recent.filter(n => n !== chosen.name)).slice(0, 10); import('../core/store.js').then(m => m.saveSettings({ xlRecentFns: st.xlRecentFns }));
    if (closeFn) closeFn(true);
    functionArgsDialog(doc, chosen);
  };
  draw();
  dialog({ title: 'Chèn hàm', width: 560, body, onOpen: (b, close) => { closeFn = close; }, buttons: [{ label: 'OK', primary: true, run: () => { done(); return true; } }, { label: 'Hủy', value: null }] });
}
function functionArgsDialog(doc, f) {
  const args = String(f.sig || '').split(/,\s*/).filter(a => a && a !== '...');
  const s = doc.wb.cur; const r = s.sel.r, c = s.sel.c;
  if (!args.length) { doc.grid.beginEdit('enter', '=' + f.name + '()'); return; }
  const inputs = args.slice(0, 8).map(a => h('input.inp', { placeholder: a.replace(/[[\]]/g, '') }));
  const result = h('div.fn-res');
  const formula = () => { const vals = inputs.map(i => i.value.trim()); while (vals.length && !vals[vals.length - 1]) vals.pop(); return f.name + '(' + vals.join(',') + ')'; };
  const upd = () => { const res = doc.calc.evalFormula(formula(), s, r, c); result.textContent = 'Kết quả công thức = ' + (res instanceof Err ? res.code : Array.isArray(res) ? '{mảng ' + res.length + '×' + res[0].length + '}' : format(res, null).t); };
  inputs.forEach(i => i.addEventListener('input', upd));
  // point-and-click: focus an arg then select cells on grid
  const body = h('div.form', h('b', f.name), h('div.muted', f.desc), args.slice(0, 8).map((a, i) => h('div.form-row', h('label.lbl', a), inputs[i])), result, h('div.hint', 'Mẹo: nhập địa chỉ như A1:A10, giá trị, hoặc công thức lồng nhau.'));
  dialog({ title: 'Đối số hàm', width: 520, body, buttons: [{ label: 'OK', primary: true, run: () => { doc.setInput(r, c, '=' + formula()); return true; } }, { label: 'Hủy', value: null }] });
  setTimeout(upd, 50);
}
export function fnCategoryMenu(doc, catName) {
  const cat = functionCatalog(fnDescriptions()).filter(f => f.cat === catName);
  return cat.map(f => ({ label: f.name, title: f.desc, run: () => functionArgsDialog(doc, f) })).concat([{ sep: true }, { label: 'Chèn hàm…', icon: 'function-square', run: () => insertFunctionDialog(doc, catName) }]);
}

// ---------------------------------------------------------------- PivotTable (static summary with refresh)
export async function pivotDialog(doc) {
  const s = doc.wb.cur;
  const cm = await import('./cmds.js');
  const rg = cm.currentRegion(doc);
  const headers = []; for (let c = rg.c1; c <= rg.c2; c++) headers.push([String(c), String(s.value(rg.r1, c) == null ? colName(c) : s.value(rg.r1, c))]);
  if (rg.r2 <= rg.r1) return toast('Chọn vùng dữ liệu có tiêu đề cột để tạo PivotTable', 'err');
  const v = await formDlg('Tạo PivotTable', [
    { type: 'html', el: h('div.hint', 'Nguồn: ' + s.name + '!' + rangeAddr(rg)) },
    { key: 'rows', type: 'select', label: 'Hàng (Rows):', options: headers, value: headers[0][0] },
    { key: 'cols', type: 'select', label: 'Cột (Columns):', options: [['', '(không)']].concat(headers), value: '' },
    { key: 'vals', type: 'select', label: 'Giá trị (Values):', options: headers, value: headers[headers.length - 1][0] },
    { key: 'agg', type: 'select', label: 'Tóm tắt theo:', options: [['sum', 'Tổng (Sum)'], ['count', 'Đếm (Count)'], ['avg', 'Trung bình'], ['max', 'Lớn nhất'], ['min', 'Nhỏ nhất']], value: 'sum' },
    { key: 'dest', type: 'select', label: 'Đặt PivotTable tại:', options: [['new', 'Trang tính mới'], ['here', 'Bên phải dữ liệu']], value: 'new' }
  ]);
  if (!v) return;
  const def = { src: { sheet: s.name, r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, rows: +v.rows, cols: v.cols === '' ? null : +v.cols, vals: +v.vals, agg: v.agg };
  doc.change('Tạo PivotTable', s, 'book', () => {
    let ts, r0, c0;
    if (v.dest === 'new') { ts = doc.wb.addSheet(cm && 'Pivot' + (doc.wb.sheets.length), doc.wb.active); r0 = 2; c0 = 0; doc.wb.active = doc.wb.sheets.indexOf(ts); }
    else { ts = s; r0 = rg.r1; c0 = rg.c2 + 2; }
    def.dest = { sheet: ts.name, r: r0, c: c0 };
    (ts.pivots = ts.pivots || []).push(def);
    buildPivot(doc, def);
  });
  doc.renderTabs(); doc.activateSheet(doc.wb.cur);
  toast('Đã tạo PivotTable. Dữ liệu > Làm mới tất cả để cập nhật.', 'ok');
}
export function buildPivot(doc, def) {
  const src = doc.wb.sheet(def.src.sheet), ts = doc.wb.sheet(def.dest.sheet); if (!src || !ts) return;
  const { r1, c1, r2, c2 } = def.src;
  const rowKeys = [], colKeys = [], data = new Map();
  const keyStr = v => v == null ? '(trống)' : (v instanceof Err ? v.code : String(v));
  for (let r = r1 + 1; r <= r2; r++) {
    const rk = keyStr(src.value(r, def.rows)), ck = def.cols != null ? keyStr(src.value(r, def.cols)) : 'Tổng';
    if (rowKeys.indexOf(rk) < 0) rowKeys.push(rk); if (colKeys.indexOf(ck) < 0) colKeys.push(ck);
    const k = rk + '\u0001' + ck; if (!data.has(k)) data.set(k, []);
    data.get(k).push(src.value(r, def.vals));
  }
  rowKeys.sort((a, b) => cmp(parseInput(a).v, parseInput(b).v)); colKeys.sort((a, b) => cmp(parseInput(a).v, parseInput(b).v));
  const agg = list => { const n = list.filter(x => typeof x === 'number'); switch (def.agg) { case 'count': return list.filter(x => x != null && x !== '').length; case 'avg': return n.length ? n.reduce((a, b) => a + b, 0) / n.length : 0; case 'max': return n.length ? Math.max.apply(null, n) : 0; case 'min': return n.length ? Math.min.apply(null, n) : 0; default: return n.reduce((a, b) => a + b, 0); } };
  const { r: R0, c: C0 } = def.dest;
  // clear old area
  if (def.last) for (let r = def.last.r1; r <= def.last.r2; r++) for (let c = def.last.c1; c <= def.last.c2; c++) ts.cells.delete(key(r, c));
  const hdrSt = internStyle({ b: true, fill: '#d9e1f2', bd: { b: { s: 'thin', c: '#8ea9db' } } }), totSt = internStyle({ b: true, fill: '#d9e1f2', bd: { t: { s: 'thin', c: '#8ea9db' } } });
  const aggName = { sum: 'Tổng', count: 'Đếm', avg: 'Trung bình', max: 'Lớn nhất', min: 'Nhỏ nhất' }[def.agg];
  ts.setCell(R0 - 1 >= 0 ? R0 - 1 : R0, C0, { v: aggName + ' của ' + String(src.value(r1, def.vals)), s: internStyle({ b: true }) });
  ts.setCell(R0, C0, { v: String(src.value(r1, def.rows) || 'Nhãn hàng'), s: hdrSt });
  colKeys.forEach((ck, j) => ts.setCell(R0, C0 + 1 + j, { v: def.cols != null ? parseInput(ck).v : aggName, s: hdrSt }));
  const multi = def.cols != null; if (multi) ts.setCell(R0, C0 + 1 + colKeys.length, { v: 'Tổng cộng', s: hdrSt });
  rowKeys.forEach((rk, i) => {
    ts.setCell(R0 + 1 + i, C0, { v: parseInput(rk).v });
    let rowAll = [];
    colKeys.forEach((ck, j) => { const list = data.get(rk + '\u0001' + ck) || []; rowAll = rowAll.concat(list); ts.setCell(R0 + 1 + i, C0 + 1 + j, { v: list.length ? agg(list) : null, s: internStyle({ nf: '#,##0.##' }) }); });
    if (multi) ts.setCell(R0 + 1 + i, C0 + 1 + colKeys.length, { v: agg(rowAll), s: internStyle({ b: true, nf: '#,##0.##' }) });
  });
  const tr = R0 + 1 + rowKeys.length;
  ts.setCell(tr, C0, { v: 'Tổng cộng', s: totSt });
  colKeys.forEach((ck, j) => { let all = []; rowKeys.forEach(rk => { all = all.concat(data.get(rk + '\u0001' + ck) || []); }); ts.setCell(tr, C0 + 1 + j, { v: agg(all), s: internStyle({ b: true, fill: '#d9e1f2', nf: '#,##0.##', bd: { t: { s: 'thin', c: '#8ea9db' } } }) }); });
  if (multi) { let all = []; data.forEach(l => { all = all.concat(l); }); ts.setCell(tr, C0 + 1 + colKeys.length, { v: agg(all), s: totSt }); }
  def.last = { r1: Math.max(0, R0 - 1), c1: C0, r2: tr, c2: C0 + colKeys.length + (multi ? 1 : 0) };
  for (let c = C0; c <= def.last.c2; c++) if (!ts.colW[c]) ts.colW[c] = 110;
}
export function refreshAll(doc) {
  doc.change('Làm mới tất cả', doc.wb.cur, 'book', () => { doc.wb.sheets.forEach(s => (s.pivots || []).forEach(p => buildPivot(doc, p))); });
  doc.recalc(true); toast('Đã làm mới dữ liệu và tính toán lại', 'ok');
}

// ---------------------------------------------------------------- Data tools
export async function textToColumns(doc) {
  const s = doc.wb.cur; const rg = clipRange(s, doc.grid.selRange());
  if (rg.c1 !== rg.c2) return toast('Chọn một cột để tách', 'err');
  const v = await formDlg('Chuyển văn bản thành cột', [{ key: 'd', type: 'select', label: 'Dấu phân cách:', options: [[',', 'Dấu phẩy (,)'], [';', 'Chấm phẩy (;)'], ['\t', 'Tab'], [' ', 'Khoảng trắng'], ['-', 'Gạch ngang (-)'], ['/', 'Gạch chéo (/)'], ['other', 'Khác…']], value: ',' }, { key: 'o', label: 'Ký tự khác:', value: '' }, { key: 'merge', type: 'checkbox', label: 'Coi các dấu phân cách liên tiếp là một', value: true }]);
  if (!v) return;
  const d = v.d === 'other' ? v.o : v.d; if (!d) return;
  let maxW = 1; const rows = [];
  for (let r = rg.r1; r <= rg.r2; r++) { const t = s.value(r, rg.c1); let parts = t == null ? [''] : String(t).split(d); if (v.merge) parts = parts.filter((x, i) => x !== '' || i === 0); rows.push(parts); maxW = Math.max(maxW, parts.length); }
  doc.change('Văn bản thành cột', s, { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c1 + maxW - 1 }, () => rows.forEach((parts, i) => parts.forEach((p, j) => s.setCell(rg.r1 + i, rg.c1 + j, { v: parseInput(p.trim()).v, f: null }))));
}
export async function removeDuplicates(doc) {
  const s = doc.wb.cur; const cm = await import('./cmds.js'); const rg = cm.currentRegion(doc);
  const cols = []; for (let c = rg.c1; c <= rg.c2; c++) cols.push({ key: 'c' + c, type: 'checkbox', label: 'Cột ' + colName(c) + ' — ' + String(s.value(rg.r1, c) == null ? '' : s.value(rg.r1, c)), value: true });
  const v = await formDlg('Loại bỏ trùng lặp', [{ key: 'head', type: 'checkbox', label: 'Dữ liệu có tiêu đề', value: true }].concat(cols));
  if (!v) return;
  const use = []; for (let c = rg.c1; c <= rg.c2; c++) if (v['c' + c]) use.push(c);
  const start = rg.r1 + (v.head ? 1 : 0); const seen = new Set(); const keep = [];
  for (let r = start; r <= rg.r2; r++) { const k = use.map(c => { const x = s.value(r, c); return typeof x === 'string' ? x.toLowerCase() : String(x); }).join('\u0001'); if (!seen.has(k)) { seen.add(k); keep.push(r); } }
  const removed = rg.r2 - start + 1 - keep.length;
  doc.change('Loại bỏ trùng lặp', s, 'sheet', () => {
    const rowsData = keep.map(r => { const a = []; for (let c = rg.c1; c <= rg.c2; c++) a.push(s.cells.get(key(r, c)) || null); return a; });
    for (let r = start; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) s.cells.delete(key(r, c));
    rowsData.forEach((a, i) => a.forEach((x, j) => { if (x) s.cells.set(key(start + i, rg.c1 + j), x); }));
  });
  alertDlg(removed + ' giá trị trùng lặp đã được tìm thấy và loại bỏ; ' + keep.length + ' giá trị duy nhất được giữ lại.', 'Loại bỏ trùng lặp');
}
export async function goalSeek(doc) {
  const s = doc.wb.cur;
  const v = await formDlg('Tìm mục tiêu (Goal Seek)', [{ key: 'set', label: 'Đặt ô (chứa công thức):', value: addr(s.sel.r, s.sel.c) }, { key: 'to', type: 'number', label: 'Thành giá trị:', value: '' }, { key: 'by', label: 'Bằng cách thay đổi ô:', value: '' }]);
  if (!v) return;
  const a = parseRangeAddr(v.set), b = parseRangeAddr(v.by); if (!a || !b) return toast('Địa chỉ không hợp lệ', 'err');
  const cell = s.get(a.r1, a.c1); if (!cell || !cell.f) return toast('Ô đích phải chứa công thức', 'err');
  const target = +v.to; const orig = s.value(b.r1, b.c1);
  const f = x => { s.setCell(b.r1, b.c1, { v: x, f: null }); doc.calc.recalc(); const y = s.value(a.r1, a.c1); return typeof y === 'number' ? y - target : NaN; };
  let x0 = typeof orig === 'number' ? orig : 0, x1 = x0 === 0 ? 1 : x0 * 1.1; let f0 = f(x0), f1 = f(x1); let ok = false;
  for (let i = 0; i < 100; i++) { if (!isFinite(f1) || !isFinite(f0)) break; if (Math.abs(f1) < 1e-9) { ok = true; break; } const d = f1 - f0; if (d === 0) break; const x2 = x1 - f1 * (x1 - x0) / d; x0 = x1; f0 = f1; x1 = x2; f1 = f(x1); }
  const found = x1;
  s.setCell(b.r1, b.c1, { v: orig, f: null }); doc.calc.recalc();
  if (ok) { doc.change('Tìm mục tiêu', s, { r1: b.r1, c1: b.c1, r2: b.r1, c2: b.c1 }, () => s.setCell(b.r1, b.c1, { v: +found.toPrecision(15), f: null })); alertDlg('Đã tìm thấy lời giải.\nGiá trị đích: ' + target + '\nGiá trị hiện tại: ' + format(s.value(a.r1, a.c1)).t + '\n' + v.by + ' = ' + found, 'Trạng thái Tìm mục tiêu'); }
  else alertDlg('Tìm mục tiêu với ô ' + v.set + ' có thể không tìm thấy lời giải.', 'Trạng thái Tìm mục tiêu');
}
export async function forecastSheet(doc) {
  const s = doc.wb.cur; const cm = await import('./cmds.js'); const rg = cm.currentRegion(doc);
  if (rg.c2 - rg.c1 < 1) return toast('Chọn 2 cột: thời gian (ngày/số) và giá trị', 'err');
  const head = typeof s.value(rg.r1, rg.c1 + 1) === 'string';
  const xs = [], ys = []; for (let r = rg.r1 + (head ? 1 : 0); r <= rg.r2; r++) { const x = s.value(r, rg.c1), y = s.value(r, rg.c1 + 1); if (typeof x === 'number' && typeof y === 'number') { xs.push(x); ys.push(y); } }
  if (xs.length < 3) return toast('Cần ít nhất 3 điểm dữ liệu số', 'err');
  const n = xs.length, mx = xs.reduce((a, b) => a + b) / n, my = ys.reduce((a, b) => a + b) / n; let sxy = 0, sxx = 0; xs.forEach((x, i) => { sxy += (x - mx) * (ys[i] - my); sxx += (x - mx) * (x - mx); });
  const b = sxy / sxx, a = my - b * mx; const step = (xs[n - 1] - xs[0]) / (n - 1);
  const v = await formDlg('Tạo trang tính dự báo', [{ key: 'k', type: 'number', label: 'Số kỳ dự báo:', value: 6 }]); if (!v) return;
  doc.change('Trang tính dự báo', s, 'book', () => {
    const ns = doc.wb.addSheet('Dự báo ' + (doc.wb.sheets.length), doc.wb.active + 1);
    const nf = (s.styleAt(rg.r1 + 1, rg.c1) || {}).nf;
    ns.setCell(0, 0, { v: 'Thời gian', s: internStyle({ b: true, fill: '#d9e1f2' }) }); ns.setCell(0, 1, { v: 'Giá trị', s: internStyle({ b: true, fill: '#d9e1f2' }) }); ns.setCell(0, 2, { v: 'Dự báo', s: internStyle({ b: true, fill: '#d9e1f2' }) });
    xs.forEach((x, i) => { ns.setCell(i + 1, 0, { v: x, s: nf ? internStyle({ nf }) : null }); ns.setCell(i + 1, 1, { v: ys[i] }); });
    ns.setCell(n, 2, { v: ys[n - 1] });
    for (let k = 1; k <= v.k; k++) { const x = xs[n - 1] + step * k; ns.setCell(n + k, 0, { v: x, s: nf ? internStyle({ nf }) : null }); ns.setCell(n + k, 2, { f: 'FORECAST.LINEAR(A' + (n + k + 1) + ',$B$2:$B$' + (n + 1) + ',$A$2:$A$' + (n + 1) + ')' }); }
    ns.colW[0] = 100; ns.colW[1] = 100; ns.colW[2] = 100;
    ns.objects.push({ id: 'ch' + Date.now(), type: 'chart', kind: 'line', src: 'A1:C' + (n + v.k + 1), x: 340, y: 20, w: 520, h: 300, title: 'Dự báo' });
    doc.wb.active = doc.wb.sheets.indexOf(ns);
  });
  doc.renderTabs(); doc.activateSheet(doc.wb.cur);
}
export function group(doc, on) {
  const s = doc.wb.cur; const rg = doc.grid.selRange();
  const byCols = rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1;
  doc.change(on ? 'Nhóm' : 'Bỏ nhóm', s, 'props', () => {
    if (byCols) for (let c = rg.c1; c <= rg.c2; c++) s.outlineC[c] = Math.max(0, (s.outlineC[c] || 0) + (on ? 1 : -1)) || undefined;
    else for (let r = rg.r1; r <= Math.min(rg.r2, s.used().r2 + 1); r++) s.outlineR[r] = Math.max(0, (s.outlineR[r] || 0) + (on ? 1 : -1)) || undefined;
  });
  toast(on ? 'Đã nhóm. Dùng Ẩn chi tiết / Hiện chi tiết để thu gọn.' : 'Đã bỏ nhóm');
}
export function outlineShow(doc, show) {
  const s = doc.wb.cur; const r0 = s.sel.r;
  // find contiguous group around selection
  if (!s.outlineR[r0]) { const c0 = s.sel.c; if (s.outlineC[c0]) { let a = c0, b = c0; while (s.outlineC[a - 1]) a--; while (s.outlineC[b + 1]) b++; doc.change(show ? 'Hiện chi tiết' : 'Ẩn chi tiết', s, 'props', () => { for (let c = a; c <= b; c++) { if (show) delete s.hiddenC[c]; else s.hiddenC[c] = true; } }); return; } return toast('Chọn một ô trong nhóm (Dữ liệu > Nhóm)', 'err'); }
  let a = r0, b = r0; while (s.outlineR[a - 1]) a--; while (s.outlineR[b + 1]) b++;
  doc.change(show ? 'Hiện chi tiết' : 'Ẩn chi tiết', s, 'props', () => { for (let r = a; r <= b; r++) { if (show) delete s.hiddenR[r]; else s.hiddenR[r] = true; } });
}
export async function subtotalDialog(doc) {
  const s = doc.wb.cur; const cm = await import('./cmds.js'); const rg = cm.currentRegion(doc);
  const heads = []; for (let c = rg.c1; c <= rg.c2; c++) heads.push([String(c), String(s.value(rg.r1, c))]);
  const v = await formDlg('Tổng phụ (Subtotal)', [{ key: 'by', type: 'select', label: 'Tại mỗi thay đổi trong:', options: heads, value: heads[0][0] }, { key: 'fn', type: 'select', label: 'Dùng hàm:', options: [['9', 'Tổng'], ['2', 'Đếm'], ['1', 'Trung bình'], ['4', 'Lớn nhất'], ['5', 'Nhỏ nhất']], value: '9' }, { key: 'col', type: 'select', label: 'Thêm tổng phụ vào:', options: heads, value: heads[heads.length - 1][0] }]);
  if (!v) return;
  const by = +v.by, col = +v.col;
  doc.change('Tổng phụ', s, 'sheet', () => {
    const rows = []; for (let r = rg.r1 + 1; r <= rg.r2; r++) { const a = []; for (let c = rg.c1; c <= rg.c2; c++) a.push(s.cells.get(key(r, c)) || null); rows.push(a); }
    for (let r = rg.r1 + 1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) s.cells.delete(key(r, c));
    let out = rg.r1 + 1, grpStart = out, prev = rows.length ? (rows[0][by - rg.c1] || {}).v : null;
    const bold = internStyle({ b: true });
    const flush = (label) => { s.setCell(out, by, { v: String(label) + ' Tổng', s: bold }); s.setCell(out, col, { f: 'SUBTOTAL(' + v.fn + ',' + colName(col) + (grpStart + 1) + ':' + colName(col) + out + ')', s: bold }); s.outlineR[out] = undefined; out++; };
    rows.forEach((a, i) => { const k = (a[by - rg.c1] || {}).v; if (i && k !== prev) { flush(prev); grpStart = out; } a.forEach((x, j) => { if (x) s.cells.set(key(out, rg.c1 + j), x); }); s.outlineR[out] = 1; out++; prev = k; });
    flush(prev);
    s.setCell(out, by, { v: 'Tổng cộng', s: bold }); s.setCell(out, col, { f: 'SUBTOTAL(' + v.fn + ',' + colName(col) + (rg.r1 + 2) + ':' + colName(col) + out + ')', s: bold });
  });
}
export async function consolidateDialog(doc) {
  const v = await formDlg('Hợp nhất (Consolidate)', [{ key: 'fn', type: 'select', label: 'Hàm:', options: [['SUM', 'Tổng'], ['AVERAGE', 'Trung bình'], ['COUNT', 'Đếm'], ['MAX', 'Lớn nhất'], ['MIN', 'Nhỏ nhất']], value: 'SUM' }, { key: 'refs', label: 'Các tham chiếu (phân cách bằng dấu ;) vd: Sheet1!B2:D10;Sheet2!B2:D10', value: '' }]);
  if (!v || !v.refs) return;
  const refs = v.refs.split(';').map(x => x.trim()).filter(Boolean); const first = parseRangeAddr(refs[0]); if (!first) return toast('Tham chiếu không hợp lệ', 'err');
  const s = doc.wb.cur, r0 = s.sel.r, c0 = s.sel.c; const R = first.r2 - first.r1, C = first.c2 - first.c1;
  doc.change('Hợp nhất', s, { r1: r0, c1: c0, r2: r0 + R, c2: c0 + C }, () => { for (let i = 0; i <= R; i++) for (let j = 0; j <= C; j++) s.setCell(r0 + i, c0 + j, { f: v.fn + '(' + refs.map(rf => { const m = /^(.*!)?(.+)$/.exec(rf); const rg = parseRangeAddr(m[2]); return (m[1] || '') + addr(rg.r1 + i, rg.c1 + j); }).join(',') + ')' }); });
}

// ---------------------------------------------------------------- Formula auditing
export function tracePrecedents(doc) {
  const s = doc.wb.cur; const cell = s.get(s.sel.r, s.sel.c); if (!cell || !cell.f) return toast('Ô hiện tại không chứa công thức', 'err');
  doc.arrows = (doc.arrows || []).concat(formulaRefs(cell.f).filter(x => !x.ref.sheet || x.ref.sheet === s.name).map(x => ({ from: { r: x.ref.r1, c: x.ref.c1, r2: x.ref.r2, c2: x.ref.c2 }, to: { r: s.sel.r, c: s.sel.c } })));
  installArrowDraw(doc); doc.grid.redraw();
}
export function traceDependents(doc) {
  const s = doc.wb.cur; const r0 = s.sel.r, c0 = s.sel.c; const arrows = [];
  s.cells.forEach((x, k) => { if (!x.f) return; formulaRefs(x.f).forEach(rf => { const R = rf.ref; if ((!R.sheet || R.sheet === s.name) && r0 >= Math.min(R.r1, R.r2) && r0 <= Math.max(R.r1, R.r2) && c0 >= Math.min(R.c1, R.c2) && c0 <= Math.max(R.c1, R.c2)) arrows.push({ from: { r: r0, c: c0 }, to: { r: Math.floor(k / 20000), c: k % 20000 } }); }); });
  if (!arrows.length) return toast('Không có ô nào phụ thuộc vào ô này');
  doc.arrows = (doc.arrows || []).concat(arrows); installArrowDraw(doc); doc.grid.redraw();
}
export function removeArrows(doc) { doc.arrows = []; doc.grid.redraw(); }
function installArrowDraw(doc) {
  if (doc._arrowHook) return; doc._arrowHook = true;
  const prev = doc.afterDraw;
  doc.afterDraw = () => {
    if (prev) prev.call(doc);
    if (!doc.arrows || !doc.arrows.length) return;
    const g = doc.grid, ctx = g.ctx;
    ctx.save(); ctx.strokeStyle = '#2f5496'; ctx.fillStyle = '#2f5496'; ctx.lineWidth = 1.5;
    doc.arrows.forEach(a => {
      const f = g.cellRect(a.from.r, a.from.c, a.from.r2, a.from.c2), t = g.cellRect(a.to.r, a.to.c);
      if (a.from.r2 != null && (a.from.r2 !== a.from.r || a.from.c2 !== a.from.c)) ctx.strokeRect(f.x + 1, f.y + 1, f.w - 2, f.h - 2);
      const x1 = f.x + Math.min(f.w, g.colW(a.from.c)) / 2, y1 = f.y + Math.min(f.h, g.rowH(a.from.r)) / 2, x2 = t.x + t.w / 2, y2 = t.y + t.h / 2;
      ctx.beginPath(); ctx.arc(x1, y1, 3, 0, Math.PI * 2); ctx.fill();
      ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
      const ang = Math.atan2(y2 - y1, x2 - x1); ctx.beginPath(); ctx.moveTo(x2, y2); ctx.lineTo(x2 - 9 * Math.cos(ang - 0.4), y2 - 9 * Math.sin(ang - 0.4)); ctx.lineTo(x2 - 9 * Math.cos(ang + 0.4), y2 - 9 * Math.sin(ang + 0.4)); ctx.closePath(); ctx.fill();
    });
    ctx.restore();
  };
}
export function errorChecking(doc) {
  const errs = []; doc.wb.sheets.forEach(s => s.cells.forEach((x, k) => { if (x.v instanceof Err) errs.push({ s, r: Math.floor(k / 20000), c: k % 20000, x }); }));
  if (!errs.length) return alertDlg('Đã hoàn tất kiểm tra lỗi cho toàn bộ sổ làm việc. Không tìm thấy lỗi.', 'Kiểm tra lỗi');
  const help = { '#DIV/0!': 'Chia cho số không hoặc ô trống.', '#N/A': 'Không tìm thấy giá trị (hàm tra cứu).', '#NAME?': 'Tên hàm hoặc tên vùng không đúng chính tả.', '#REF!': 'Tham chiếu ô không hợp lệ (ô đã bị xóa).', '#VALUE!': 'Sai kiểu dữ liệu đối số (vd: cộng văn bản).', '#NUM!': 'Giá trị số không hợp lệ.', '#SPILL!': 'Vùng tràn kết quả mảng bị chặn bởi ô có dữ liệu.', '#CALC!': 'Kết quả mảng rỗng.' };
  const body = h('div', errs.slice(0, 200).map(e => { const row = h('div.find-row', h('span', e.s.name + '!' + addr(e.r, e.c)), h('b', { style: { color: '#c00000' } }, e.x.v.code), h('span', '=' + (e.x.f || '')), h('span.muted', help[e.x.v.code] || '')); row.addEventListener('click', () => { if (e.s !== doc.wb.cur) doc.activateSheet(e.s); doc.grid.select(e.r, e.c); doc.grid.scrollTo(e.r, e.c); }); return row; }));
  dialog({ title: 'Kiểm tra lỗi — ' + errs.length + ' lỗi', width: 680, body, buttons: [{ label: 'Đóng', primary: true }] });
}
export function evaluateFormula(doc) {
  const s = doc.wb.cur; const cell = s.get(s.sel.r, s.sel.c); if (!cell || !cell.f) return toast('Ô hiện tại không chứa công thức', 'err');
  // step-by-step: evaluate each reference and innermost function call
  const steps = []; let f = cell.f;
  formulaRefs(cell.f).forEach(rf => { const t = cell.f.substr(rf.p, rf.len); const v = doc.calc.evalFormula(t, s, s.sel.r, s.sel.c); steps.push(t + ' = ' + (Array.isArray(v) ? '{' + v.length + '×' + v[0].length + '}' : v instanceof Err ? v.code : format(v).t)); });
  const re = /([A-Z][A-Z0-9.]*)\(([^()]*)\)/; let guard = 0;
  while (re.test(f) && guard++ < 20) { const m = re.exec(f); const v = doc.calc.evalFormula(m[0], s, s.sel.r, s.sel.c); const txt = v instanceof Err ? v.code : Array.isArray(v) ? '{…}' : typeof v === 'string' ? '"' + v + '"' : format(v).t; steps.push(m[0] + ' → ' + txt); f = f.replace(m[0], typeof v === 'number' ? String(v) : typeof v === 'string' ? '"' + v.replace(/"/g, '""') + '"' : txt); }
  const final = s.value(s.sel.r, s.sel.c);
  dialog({ title: 'Đánh giá công thức — ' + addr(s.sel.r, s.sel.c), width: 620, body: h('div', h('div.fn-info', '=' + cell.f), h('ol.eval-steps', steps.map(x => h('li', x))), h('div', h('b', 'Kết quả: '), final instanceof Err ? final.code : format(final, cell.s && cell.s.nf).t)), buttons: [{ label: 'Đóng', primary: true }] });
}
export function watchWindow(doc) {
  doc.watches = doc.watches || [];
  const s = doc.wb.cur; const k = s.name + '!' + addr(s.sel.r, s.sel.c); if (doc.watches.indexOf(k) < 0) doc.watches.push(k);
  const old = $('.xl-watch', doc.el); if (old) old.parentNode.removeChild(old);
  const pane = h('div.pane.xl-watch', h('div.pane-head', h('span.t', 'Cửa sổ Theo dõi'), h('button.pane-x', { onclick: () => pane.parentNode.removeChild(pane) }, ic('x', 16))));
  const body = h('div.pane-body'); pane.appendChild(body);
  const draw = () => { body.innerHTML = ''; doc.watches.forEach(w => { const m = /^(.*)!(.+)$/.exec(w); const sh = doc.wb.sheet(m[1]); if (!sh) return; const a = parseRangeAddr(m[2]); const v = sh.value(a.r1, a.c1); const cell = sh.get(a.r1, a.c1); body.appendChild(h('div.cm-card', h('b', w), h('div', 'Giá trị: ' + (v instanceof Err ? v.code : format(v, cell && cell.s && cell.s.nf).t)), cell && cell.f ? h('div.muted', '=' + cell.f) : null, h('button.btn', { onclick: () => { doc.watches = doc.watches.filter(x => x !== w); draw(); } }, 'Bỏ theo dõi'))); }); };
  draw(); doc.el.appendChild(pane);
  const prev = doc.afterChange.bind(doc); if (!doc._watchHooked) { doc._watchHooked = true; const orig = doc.afterChange; doc.afterChange = function (n) { orig.call(doc, n); if (pane.parentNode) draw(); }; }
}

// ---------------------------------------------------------------- Freeze panes
export function freeze(doc, mode) {
  const s = doc.wb.cur;
  doc.change('Cố định ngăn', s, 'props', () => {
    if (mode === 'none') s.freeze = { r: 0, c: 0 };
    else if (mode === 'row') s.freeze = { r: 1 + Math.floor(doc.grid.scroller.scrollTop / (20 * doc.grid.z)) * 0, c: 0 };
    else if (mode === 'col') s.freeze = { r: 0, c: 1 };
    else s.freeze = { r: s.sel.r, c: s.sel.c };
  });
  doc.grid.scroller.scrollTop = 0; doc.grid.scroller.scrollLeft = 0; doc.grid.updateSpacer(); doc.grid.redraw();
}

// ---------------------------------------------------------------- Page setup & printing
const PAPER = { A4: [210, 297], A3: [297, 420], A5: [148, 210], Letter: [216, 279], Legal: [216, 356] };
const MARGINS = { normal: [19, 18, 19, 18], wide: [25, 25, 25, 25], narrow: [13, 6, 13, 6] };
export function pageSetupDialog(doc) {
  const s = doc.wb.cur; const ps = s.pageSetup || (s.pageSetup = {});
  return formDlg('Thiết lập trang', [
    { key: 'orientation', type: 'radio', label: 'Hướng', value: ps.orientation, options: [['portrait', 'Dọc'], ['landscape', 'Ngang']] },
    { key: 'size', type: 'select', label: 'Khổ giấy', value: ps.size, options: Object.keys(PAPER).map(k => [k, k + ' (' + PAPER[k][0] + '×' + PAPER[k][1] + ' mm)']) },
    { key: 'margins', type: 'select', label: 'Lề', value: ps.margins, options: [['normal', 'Bình thường'], ['wide', 'Rộng'], ['narrow', 'Hẹp']] },
    { key: 'fit', type: 'select', label: 'Co giãn', value: ps.fit || 'none', options: [['none', 'Không co giãn (theo tỷ lệ)'], ['width', 'Vừa tất cả cột trên một trang'], ['rows1', 'Vừa tất cả hàng trên một trang'], ['page', 'Vừa trang tính trên một trang']] },
    { key: 'printArea', label: 'Vùng in (để trống = toàn bộ):', value: ps.printArea || '' },
    { key: 'printTitles', label: 'Hàng lặp lại ở đầu mỗi trang (vd: 1:2):', value: ps.printTitles || '' },
    { key: 'printTitlesCols', label: 'Cột lặp lại ở bên trái mỗi trang (vd: A:B):', value: ps.printTitlesCols || '' },
    { key: 'scale', label: 'Tỷ lệ in (%) khi không co giãn:', value: String(ps.scale || 100) },
    { key: 'order', type: 'radio', label: 'Thứ tự trang', value: ps.order || 'downThenOver', options: [['downThenOver', 'Xuống rồi ngang'], ['overThenDown', 'Ngang rồi xuống']] },
    { key: 'centerH', type: 'checkbox', label: 'Căn giữa theo chiều ngang', value: !!ps.centerH },
    { key: 'gridlines', type: 'checkbox', label: 'In đường lưới', value: ps.gridlines },
    { key: 'headings', type: 'checkbox', label: 'In tiêu đề hàng/cột', value: ps.headings },
    { key: 'header', label: 'Đầu trang:', value: ps.header || '' },
    { key: 'footer', label: 'Chân trang:', value: ps.footer || 'Trang &[Page]' }
  ], { width: 480 }).then(v => {
    if (!v) return;
    v.scale = Math.max(10, Math.min(400, parseInt(v.scale) || 100));
    doc.change('Thiết lập trang', s, 'props', () => { Object.assign(s.pageSetup, v); });
    s.showPageBreaks = true;
    doc.grid.redraw();
  });
}
export function setPage(doc, patch) {
  const s = doc.wb.cur;
  doc.change('Thiết lập trang', s, 'props', () => {
    if (!s.pageSetup) s.pageSetup = {};
    Object.assign(s.pageSetup, patch);
  });
  s.showPageBreaks = true;
  if (doc.grid) doc.grid.redraw();
  if (doc.refreshUi) doc.refreshUi();
  doc.markDirty();
  toast('Đã cập nhật thiết lập trang tính');
}
export function printOptions(doc) {
  const s = doc.wb.cur; const ps = s.pageSetup;
  const sel = (label, val, opts, key2) => { const i = h('select.inp', opts.map(o => h('option', { value: o[0] }, o[1]))); i.value = val; i.addEventListener('change', () => { ps[key2] = i.value; doc.markDirty(); const pv = $('.print-prev'); if (pv) printPreview(doc, pv); }); return h('div.form-row', h('label.lbl', label), i); };
  return h('div.form',
    sel('Vùng in', ps.what || 'sheet', [['sheet', 'In trang tính hiện hoạt'], ['book', 'In toàn bộ sổ làm việc'], ['selection', 'In vùng chọn']], 'what'),
    sel('Hướng', ps.orientation, [['portrait', 'Hướng dọc'], ['landscape', 'Hướng ngang']], 'orientation'),
    sel('Khổ giấy', ps.size, Object.keys(PAPER).map(k => [k, k]), 'size'),
    sel('Lề', ps.margins, [['normal', 'Lề bình thường'], ['wide', 'Lề rộng'], ['narrow', 'Lề hẹp']], 'margins'),
    sel('Co giãn', ps.fit || 'none', [['none', 'Không co giãn (theo tỷ lệ)'], ['page', 'Vừa trang tính trên một trang'], ['width', 'Vừa tất cả cột trên một trang'], ['rows1', 'Vừa tất cả hàng trên một trang']], 'fit'),
    sel('Tỷ lệ', String(ps.scale || 100), [['50', '50%'], ['75', '75%'], ['85', '85%'], ['90', '90%'], ['100', '100%'], ['110', '110%'], ['125', '125%'], ['150', '150%']], 'scale'),
    sel('Thứ tự trang', ps.order || 'downThenOver', [['downThenOver', 'Xuống rồi sang ngang'], ['overThenDown', 'Ngang rồi xuống dưới']], 'order'),
    h('div.muted', { style: { margin: '6px 0' } }, 'Số trang sẽ in: ' + P8.pageCount(doc)),
    h('button.btn', { onclick: () => pageSetupDialog(doc) }, 'Thiết lập trang…'));
}
export function renderPrint(doc, area) { return P8.renderPrint(doc, area); }
export function printPreview(doc, box) { return P8.printPreview(doc, box); }
export function setPrintArea(doc, range) { return P8.setPrintArea(doc, range); }
export function addToPrintArea(doc, range) { return P8.addToPrintArea(doc, range); }
export function clearPrintArea(doc) { return P8.clearPrintArea(doc); }
export function printAreaMenu(doc) { return P8.printAreaMenu(doc); }

// ---------------------------------------------------------------- Themes (Page Layout)
export const THEMES = [
  { name: 'Office', font: 'Calibri', colors: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'] },
  { name: 'Office 2013–2022', font: 'Calibri', colors: ['#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000', '#4472C4', '#70AD47'] },
  { name: 'Liquid Glass', font: 'Segoe UI', colors: ['#0F6CBD', '#8764B8', '#00B7C3', '#E3008C', '#FFB900', '#107C10'] },
  { name: 'VISICONS', font: 'Times New Roman', colors: ['#1F3864', '#C00000', '#2E75B5', '#BF8F00', '#548235', '#7030A0'] },
  { name: 'Retrospect', font: 'Calibri Light', colors: ['#E48312', '#BD582C', '#865640', '#9B8357', '#C2BC80', '#94A088'] },
  { name: 'Slice', font: 'Century Gothic', colors: ['#052F61', '#A50E82', '#14967C', '#6A9E1F', '#E87D37', '#C62324'] }
];
export function applyTheme(doc, t) {
  doc.change('Chủ đề: ' + t.name, doc.wb.cur, 'book', () => { doc.wb.props.theme = t.name; doc.wb.props.themeColors = t.colors; doc.wb.props.themeFont = t.font; });
  import('../core/store.js').then(m => m.saveSettings({ excelFont: t.font }));
  doc.grid.measureCache = {}; doc.grid.redraw(); doc.objects.render();
}

// ---------------------------------------------------------------- Charts
export async function insertChart(doc, kind, newSheet) {
  const cm = await import('./cmds.js');
  const s = doc.wb.cur; let rg = cm.currentRegion(doc);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && !doc.grid.cellHasContent(rg.r1, rg.c1)) return toast('Chọn vùng dữ liệu (có tiêu đề) trước khi chèn biểu đồ', 'err');
  const src = rangeAddr(rg);
  if (newSheet) {
    doc.change('Trang biểu đồ', s, 'book', () => { const ns = doc.wb.addSheet('Biểu đồ' + (doc.wb.sheets.length), doc.wb.active + 1); ns.showGrid = false; ns.objects.push({ id: 'ch' + Date.now(), type: 'chart', kind, src: "'" + s.name + "'!" + src, x: 10, y: 10, w: 900, h: 520, title: '' }); doc.wb.active = doc.wb.sheets.indexOf(ns); });
    doc.renderTabs(); doc.activateSheet(doc.wb.cur); return;
  }
  doc.objects.addChart(kind, src);
}
export async function recommendedCharts(doc) {
  const { CHART_TYPES, renderChart, chartData } = await import('./charts.js');
  const cm = await import('./cmds.js'); const s = doc.wb.cur; const rg = cm.currentRegion(doc); const src = rangeAddr(rg);
  const grid = h('div.chart-gallery');
  let chosen = 'column';
  CHART_TYPES.forEach(t => { const card = h('div.chart-card' + (t[0] === chosen ? '.on' : ''), h('div.chart-card-prev'), h('div', t[1])); const pv = card.firstChild; setTimeout(() => renderChart(pv, chartData(s, src, { kind: t[0] }), { kind: t[0], w: 200, h: 120, title: false, legend: false }), 10); card.addEventListener('click', () => { chosen = t[0]; $$('.chart-card', grid).forEach(x => x.classList.remove('on')); card.classList.add('on'); }); card.addEventListener('dblclick', () => { $('.dlg-back .btn.primary').click(); }); grid.appendChild(card); });
  const ok = await dialog({ title: 'Biểu đồ được đề xuất — dữ liệu ' + src, width: 760, body: grid, buttons: [{ label: 'OK', primary: true, value: true }, { label: 'Hủy', value: false }] });
  if (ok) doc.objects.addChart(chosen, src);
}
export function sparklines(doc, kind) {
  const s = doc.wb.cur;
  promptDlg('Vùng dữ liệu cho biểu đồ thu nhỏ (vd: B2:F2):', '', 'Chèn biểu đồ thu nhỏ (' + kind + ')').then(v => {
    const rg = v && parseRangeAddr(v); if (!rg) return;
    const vals = []; for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const x = s.value(r, c); vals.push(typeof x === 'number' ? x : 0); }
    const w = 120, hh = 24; const mn = Math.min.apply(null, vals), mx = Math.max.apply(null, vals);
    let body = '';
    if (kind === 'line') body = '<polyline fill="none" stroke="#4472c4" stroke-width="1.5" points="' + vals.map((x, i) => (i / Math.max(1, vals.length - 1) * (w - 4) + 2).toFixed(1) + ',' + (hh - 2 - (x - mn) / (mx - mn || 1) * (hh - 4)).toFixed(1)).join(' ') + '"/>';
    else body = vals.map((x, i) => { const bw = (w - 4) / vals.length; const y0 = kind === 'winloss' ? hh / 2 : hh - 2; const bh = kind === 'winloss' ? (hh / 2 - 2) * (x >= 0 ? 1 : -1) : (x - Math.min(0, mn)) / (mx - Math.min(0, mn) || 1) * (hh - 4); return '<rect x="' + (2 + i * bw + 1).toFixed(1) + '" y="' + Math.min(y0, y0 - bh).toFixed(1) + '" width="' + (bw - 2).toFixed(1) + '" height="' + Math.abs(bh).toFixed(1) + '" fill="' + (x < 0 ? '#c00000' : '#4472c4') + '"/>'; }).join('');
    const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + hh + '">' + body + '</svg>');
    const g = doc.grid; const rc = g.cellRect(s.sel.r, s.sel.c); const z = g.z;
    doc.objects.add({ id: 'sp' + Date.now(), type: 'image', src: url, alt: 'Biểu đồ thu nhỏ ' + v, x: (rc.x - g.headW() + g.scroller.scrollLeft) / z + 2, y: (rc.y - g.headH() + g.scroller.scrollTop) / z + 1, w: s.colWidth(s.sel.c) - 4, h: s.rowHeight(s.sel.r) - 2 });
  });
}
