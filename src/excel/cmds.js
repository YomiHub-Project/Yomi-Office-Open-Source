// Excel commands: formatting, structure, clipboard, fill, sort & filter, find, names, protection, context menus
import { h, ic, esc, $, $$, clamp, deepClone, utf8Decode, keyName } from '../core/util.js';
import { toast, showMenu, dialog, promptDlg, confirmDlg, formDlg, colorPicker, alertDlg } from '../core/ui.js';
import { settings } from '../core/store.js';
import Host from '../core/host.js';
import { computePages } from './pages.js';
import { Err, E, key, colName, addr, rangeAddr, parseRangeAddr, parseAddr, norm, inRange, internStyle, mergeStyle, MAX_ROWS, MAX_COLS, DEF_COL_W, DEF_ROW_H, dateToSerial, serialToParts } from './model.js';
import { FN, shiftFormula, adjustFormulaForInsert, parse, cmp, tokenize, formulaRefs } from './formula.js';
import { format, parseInput, NUMBER_FORMATS, isDateFormat } from './numfmt.js';
import { PT2PX } from './grid.js';
export * from './features.js';
import { cfStyleImpl, formatCellsDialog, formatAsTable, dvCheck } from './features.js';

// ---------------------------------------------------------------- sheet state snapshots (for undo of structural ops)
export function sheetState(s) {
  const cells = []; s.cells.forEach((x, k) => cells.push([k, Object.assign({}, x)]));
  return { cells, props: sheetProps(s) };
}
export function applySheetState(s, st) { s.cells = new Map(st.cells.map(([k, x]) => [k, Object.assign({}, x)])); applySheetProps(s, st.props); }
const PROPS = ['colW', 'rowH', 'hiddenR', 'hiddenC', 'colS', 'rowS', 'merges', 'freeze', 'cf', 'dv', 'filter', 'tables', 'objects', 'outlineR', 'outlineC', 'protected', 'tabColor', 'name', 'showGrid', 'showHead', 'pageSetup', 'ink'];
export function sheetProps(s) { const o = {}; PROPS.forEach(p => { o[p] = deepClone(s[p]); }); return o; }
export function applySheetProps(s, o) {
  PROPS.forEach(p => { if (o[p] !== undefined) s[p] = deepClone(o[p]); });
  for (const k in s.colS) s.colS[k] = internStyle(s.colS[k]); for (const k in s.rowS) s.rowS[k] = internStyle(s.rowS[k]);
  if (s._ax) { s._ax.r.dirty = true; s._ax.c.dirty = true; }
}

// ---------------------------------------------------------------- iteration helpers
function selRanges(doc) { return doc.wb.cur.sel.ranges; }
function clipRange(s, rg) { const u = s.used(); return { r1: rg.r1, c1: rg.c1, r2: rg.r2 >= MAX_ROWS - 1 ? Math.max(u.r2, rg.r1) : rg.r2, c2: rg.c2 >= MAX_COLS - 1 ? Math.max(u.c2, rg.c1) : rg.c2 }; }
function eachSelCell(doc, fn) { const s = doc.wb.cur; selRanges(doc).forEach(rg => { const R = clipRange(s, rg); for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) fn(r, c); }); }
function bound(doc) { return doc.boundingSel(); }
function isFullCol(rg) { return rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1; }
function isFullRow(rg) { return rg.c1 === 0 && rg.c2 >= MAX_COLS - 1; }

// ---------------------------------------------------------------- style helpers
export function curStyle(doc) { const s = doc.wb.cur; return s.styleAt(s.sel.r, s.sel.c) || {}; }
export function setStyle(doc, patch, label) {
  const s = doc.wb.cur;
  doc.change(label || 'Định dạng', s, selRanges(doc).some(rg => isFullCol(rg) || isFullRow(rg)) ? 'sheet' : bound(doc), () => {
    selRanges(doc).forEach(rg => {
      if (isFullCol(rg) && isFullRow(rg)) { /* whole sheet: apply to used cells and set column styles */ for (let c = 0; c <= Math.max(s.used().c2, 30); c++) s.colS[c] = mergeStyle(s.colS[c], patch); }
      else if (isFullCol(rg)) for (let c = rg.c1; c <= rg.c2; c++) s.colS[c] = mergeStyle(s.colS[c], patch);
      else if (isFullRow(rg)) for (let r = rg.r1; r <= rg.r2; r++) s.rowS[r] = mergeStyle(s.rowS[r], patch);
      const R = clipRange(s, rg);
      for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) {
        const cell = s.get(r, c);
        if (!cell && (isFullCol(rg) || isFullRow(rg))) continue;
        const base = cell && cell.s ? cell.s : (s.rowS[r] || s.colS[c] || null);
        s.setCell(r, c, { s: mergeStyle(base, patch) });
      }
    });
    if (patch.sz || patch.wrap || patch.font) for (const rg of selRanges(doc)) { const R = clipRange(s, rg); for (let r = R.r1; r <= Math.min(R.r2, R.r1 + 500); r++) autoRowHeight(doc, s, r); }
  });
  doc.lastRepeatable = () => setStyle(doc, patch, label);
}
export function toggleStyle(doc, k, v) { const cur = curStyle(doc); setStyle(doc, { [k]: cur[k] ? null : (v || true) }, { b: 'In đậm', i: 'In nghiêng', u: 'Gạch chân', s: 'Gạch ngang' }[k]); }
export function setNumFmt(doc, nf) { setStyle(doc, { nf: nf === 'General' ? null : nf }, 'Định dạng số'); }
export function changeDecimals(doc, d) {
  const st = curStyle(doc); let nf = st.nf || 'General';
  const s = doc.wb.cur; const v = s.value(s.sel.r, s.sel.c);
  if (nf === 'General') { const t = typeof v === 'number' ? String(v) : '0'; const dec = (t.split('.')[1] || '').length; nf = dec ? '0.' + '0'.repeat(dec) : '0'; }
  if (d > 0) nf = /\.0+/.test(nf) ? nf.replace(/\.(0+)/, m => m + '0') : nf.replace(/(0)(?![^"]*")([^0]*)$/, '$1.0$2');
  else nf = nf.replace(/\.(0+)/, (m, z) => z.length > 1 ? '.' + z.slice(1) : '');
  setNumFmt(doc, nf);
}
export function fontSizeStep(doc, dir) {
  const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const cur = curStyle(doc).sz || settings().excelSize || 11;
  let n = dir > 0 ? sizes.find(x => x > cur) : sizes.slice().reverse().find(x => x < cur);
  if (!n) n = dir > 0 ? cur + 10 : Math.max(1, cur - 1);
  setStyle(doc, { sz: n }, 'Cỡ chữ');
}
export function autoRowHeight(doc, s, r) {
  if (s.rowHManual && s.rowHManual[r]) return;
  const u = s.used(); let maxH = DEF_ROW_H, any = false;
  const g = doc.grid;
  for (let c = 0; c <= u.c2; c++) {
    const cell = s.get(r, c); if (!cell) continue;
    const st = cell.s || {};
    const sz = (st.sz || settings().excelSize || 11);
    let lines = 1;
    if (cell.v != null) {
      const text = format(cell.v, st.nf).t;
      if (st.wrap) { const fontStr = g.font(Object.assign({}, st, {})).replace(/^[\d.]+px/, ''); lines = g.wrapLines(text, g.font(st), s.colWidth(c) * g.z - 8).length; }
      else lines = String(text).split('\n').length === 1 ? 1 : 1;
    }
    const hgt = Math.ceil(sz * PT2PX * 1.25 * lines + 4);
    if (hgt > maxH) maxH = hgt;
    any = true;
  }
  if (!any) return;
  if (maxH === DEF_ROW_H) delete s.rowH[r]; else s.rowH[r] = maxH;
}

// ---------------------------------------------------------------- borders
export function applyBorder(doc, kind, color, style) {
  const s = doc.wb.cur; const b = { s: style || doc._borderStyle || 'thin', c: color || doc._borderColor || '#000000' };
  const thick = { s: 'medium', c: b.c }, dbl = { s: 'double', c: b.c };
  doc.change('Viền', s, bound(doc), () => {
    selRanges(doc).forEach(rg0 => {
      const rg = clipRange(s, rg0);
      for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
        const cell = s.get(r, c); const st = cell && cell.s ? cell.s : null; let bd = {};
        const top = r === rg.r1, bot = r === rg.r2, lef = c === rg.c1, rig = c === rg.c2;
        switch (kind) {
          case 'none': bd = { t: null, b: null, l: null, r: null, d: null, u: null }; break;
          case 'all': bd = { t: b, b: b, l: b, r: b }; break;
          case 'outside': if (top) bd.t = b; if (bot) bd.b = b; if (lef) bd.l = b; if (rig) bd.r = b; break;
          case 'thickbox': if (top) bd.t = thick; if (bot) bd.b = thick; if (lef) bd.l = thick; if (rig) bd.r = thick; break;
          case 'inside': if (!top) bd.t = b; if (!bot) bd.b = b; if (!lef) bd.l = b; if (!rig) bd.r = b; break;
          case 'top': if (top) bd.t = b; break; case 'bottom': if (bot) bd.b = b; break; case 'left': if (lef) bd.l = b; break; case 'right': if (rig) bd.r = b; break;
          case 'bottomdouble': if (bot) bd.b = dbl; break; case 'thickbottom': if (bot) bd.b = thick; break;
          case 'topbottom': if (top) bd.t = b; if (bot) bd.b = b; break; case 'topthickbottom': if (top) bd.t = b; if (bot) bd.b = thick; break; case 'topdoublebottom': if (top) bd.t = b; if (bot) bd.b = dbl; break;
          case 'diagdown': bd.u = b; break; case 'diagup': bd.d = b; break;
        }
        if (!Object.keys(bd).length) continue;
        s.setCell(r, c, { s: mergeStyle(st, { bd }) });
      }
    });
  });
  doc.lastBorder = kind;
}
export const BORDER_MENU = [['bottom', 'Viền dưới', 'panel-bottom'], ['top', 'Viền trên', 'panel-top'], ['left', 'Viền trái', 'panel-left'], ['right', 'Viền phải', 'panel-right'], null, ['none', 'Không viền', 'square-dashed'], ['all', 'Tất cả viền', 'grid-3x3'], ['outside', 'Viền ngoài', 'square'], ['thickbox', 'Viền ngoài đậm', 'square'], ['inside', 'Viền trong', 'plus'], null, ['bottomdouble', 'Viền dưới kép', 'equal'], ['thickbottom', 'Viền dưới đậm', 'minus'], ['topbottom', 'Viền trên và dưới', 'rows-2'], ['topthickbottom', 'Viền trên và dưới đậm', 'rows-2'], ['topdoublebottom', 'Viền trên và dưới kép', 'rows-2'], null, ['diagdown', 'Đường chéo xuống', 'slash'], ['diagup', 'Đường chéo lên', 'slash']];
export function borderMenu(doc) {
  return BORDER_MENU.map(b => b ? { label: b[1], icon: b[2], run: () => applyBorder(doc, b[0]) } : { sep: true }).concat([{ sep: true },
    { label: 'Màu đường viền', icon: 'palette', sub: [{ el: h('div', { style: { padding: '4px' } }, 'Chọn màu…') }], run: null },
    { label: 'Kiểu đường viền', icon: 'minus', sub: [['thin', 'Mảnh'], ['medium', 'Vừa'], ['thick', 'Đậm'], ['dashed', 'Nét đứt'], ['dotted', 'Chấm'], ['double', 'Kép']].map(x => ({ label: x[1], checked: (doc._borderStyle || 'thin') === x[0] || undefined, run: () => { doc._borderStyle = x[0]; toast('Kiểu viền: ' + x[1]); } })) },
    { label: 'Thêm viền khác…', icon: 'square-dashed-bottom', run: () => formatCellsDialog(doc, 'border') }]).map(it => it.label === 'Màu đường viền' ? { label: it.label, icon: it.icon, run: () => colorPicker(doc.ribbon.el, c => { doc._borderColor = c; toast('Màu viền đã chọn'); }) } : it);
}

// ---------------------------------------------------------------- alignment / merge
export function setAlign(doc, ha, va) { const p = {}; if (ha !== undefined) p.ha = ha; if (va !== undefined) p.va = va; setStyle(doc, p, 'Căn lề'); }
export function indent(doc, d) { const cur = curStyle(doc).indent || 0; setStyle(doc, { indent: Math.max(0, cur + d) || null, ha: curStyle(doc).ha || 'left' }, 'Thụt lề'); }
export function merge(doc, mode) {
  const s = doc.wb.cur;
  doc.change(mode === 'unmerge' ? 'Bỏ gộp ô' : 'Gộp ô', s, bound(doc), () => {
    selRanges(doc).forEach(rg0 => {
      const rg = clipRange(s, rg0);
      s.merges = s.merges.filter(m => !(m.r1 <= rg.r2 && m.r2 >= rg.r1 && m.c1 <= rg.c2 && m.c2 >= rg.c1));
      if (mode === 'unmerge') return;
      const rows = mode === 'across' ? range(rg.r1, rg.r2).map(r => ({ r1: r, r2: r, c1: rg.c1, c2: rg.c2 })) : [rg];
      rows.forEach(m => {
        if (m.r1 === m.r2 && m.c1 === m.c2) return;
        // keep top-left value only
        let kept = false;
        for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) { if (r === m.r1 && c === m.c1) continue; const x = s.get(r, c); if (x && (x.v != null || x.f)) { kept = true; s.setCell(r, c, { v: null, f: null }); } }
        s.merges.push({ r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 });
        if (mode === 'center') { const x = s.get(m.r1, m.c1); s.setCell(m.r1, m.c1, { s: mergeStyle(x && x.s, { ha: 'center' }) }); }
      });
    });
  });
}
function range(a, b) { const o = []; for (let i = a; i <= b; i++) o.push(i); return o; }

// ---------------------------------------------------------------- clear / delete content
export function clear(doc, what) {
  const s = doc.wb.cur;
  doc.change({ all: 'Xóa tất cả', contents: 'Xóa nội dung', formats: 'Xóa định dạng', comments: 'Xóa chú thích', links: 'Xóa liên kết' }[what] || 'Xóa', s, selRanges(doc).some(rg => isFullCol(rg) || isFullRow(rg)) ? 'sheet' : bound(doc), () => {
    selRanges(doc).forEach(rg0 => {
      const rg = clipRange(s, rg0);
      const ks = []; s.cells.forEach((x, k) => { const r = Math.floor(k / 20000), c = k % 20000; if (inRange(rg, r, c)) ks.push(k); });
      ks.forEach(k => {
        const x = s.cells.get(k); const r = Math.floor(k / 20000), c = k % 20000;
        if (what === 'all') s.cells.delete(k);
        else if (what === 'contents') s.setCell(r, c, { v: null, f: null });
        else if (what === 'formats') { s.setCell(r, c, { s: null }); }
        else if (what === 'comments') s.setCell(r, c, { note: null });
        else if (what === 'links') { s.setCell(r, c, { link: null }); }
      });
      if (what === 'all' || what === 'formats') { if (isFullCol(rg0)) for (let c = rg0.c1; c <= rg0.c2; c++) delete s.colS[c]; if (isFullRow(rg0)) for (let r = rg0.r1; r <= rg0.r2; r++) delete s.rowS[r]; s.merges = s.merges.filter(m => !(m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2)); }
    });
  });
}

// ---------------------------------------------------------------- insert / delete rows, columns, cells
function shiftAll(doc, s, axis, at, n) {
  // move cells
  const cells = new Map();
  s.cells.forEach((x, k) => {
    let r = Math.floor(k / 20000), c = k % 20000;
    if (axis === 'r') { if (n < 0 && r >= at && r < at - n) return; if (r >= at) r += n; }
    else { if (n < 0 && c >= at && c < at - n) return; if (c >= at) c += n; }
    if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) return;
    cells.set(key(r, c), x);
  });
  s.cells = cells;
  const shiftMap = (m) => { const o = {}; for (const k in m) { let i = +k; if (n < 0 && i >= at && i < at - n) continue; if (i >= at) i += n; o[i] = m[k]; } return o; };
  if (axis === 'r') { s.rowH = shiftMap(s.rowH); s.hiddenR = shiftMap(s.hiddenR); s.rowS = shiftMap(s.rowS); s.outlineR = shiftMap(s.outlineR); }
  else { s.colW = shiftMap(s.colW); s.hiddenC = shiftMap(s.hiddenC); s.colS = shiftMap(s.colS); s.outlineC = shiftMap(s.outlineC); }
  const adj = (rg) => {
    const a1 = axis === 'r' ? 'r1' : 'c1', a2 = axis === 'r' ? 'r2' : 'c2';
    if (n > 0) { if (rg[a1] >= at) rg[a1] += n; if (rg[a2] >= at) rg[a2] += n; }
    else { const e = at - n - 1; const f = v => v < at ? v : v > e ? v + n : at; const nv1 = f(rg[a1]), nv2 = rg[a2] > e ? rg[a2] + n : (rg[a2] < at ? rg[a2] : at - 1); if (nv2 < nv1) return null; rg[a1] = nv1; rg[a2] = nv2; }
    return rg;
  };
  s.merges = s.merges.map(m => adj(Object.assign({}, m))).filter(Boolean).filter(m => m.r1 !== m.r2 || m.c1 !== m.c2);
  s.cf = s.cf.map(c => { const rg = adj(Object.assign({}, c.range)); return rg ? Object.assign({}, c, { range: rg }) : null; }).filter(Boolean);
  s.dv = s.dv.map(c => { const rg = adj(Object.assign({}, c.range)); return rg ? Object.assign({}, c, { range: rg }) : null; }).filter(Boolean);
  s.tables = s.tables.map(t => adj(Object.assign({}, t))).filter(Boolean);
  if (s.filter) { const f = adj(Object.assign({}, s.filter)); s.filter = f; }
  if (axis === 'r' && s.freeze.r > at) s.freeze.r = Math.max(at, s.freeze.r + n);
  if (axis === 'c' && s.freeze.c > at) s.freeze.c = Math.max(at, s.freeze.c + n);
  // formulas everywhere in the workbook
  doc.wb.sheets.forEach(sh => sh.cells.forEach(x => { if (x.f) x.f = adjustFormulaForInsert(x.f, sh.name, s.name, axis, at, n); }));
  for (const nm in doc.wb.names) { const d = doc.wb.names[nm]; if (d && d.f) { d.f = adjustFormulaForInsert(d.f.replace(/^=/, ''), s.name, s.name, axis, at, n); d._ast = null; } }
  s.objects.forEach(o => { if (o.anchor) { if (axis === 'r' && o.anchor.r >= at) o.anchor.r = Math.max(0, o.anchor.r + n); if (axis === 'c' && o.anchor.c >= at) o.anchor.c = Math.max(0, o.anchor.c + n); } if (o.type === 'chart' && o.src) { const rg = parseRangeAddr(o.src); if (rg) { const a = adj(rg); if (a) o.src = rangeAddr(a); } } });
  if (s._ax) { s._ax.r.dirty = true; s._ax.c.dirty = true; }
}
export function insertRows(doc, at, n) { const s = doc.wb.cur; doc.change('Chèn hàng', s, 'book', () => shiftAll(doc, s, 'r', at, n)); }
export function insertCols(doc, at, n) { const s = doc.wb.cur; doc.change('Chèn cột', s, 'book', () => shiftAll(doc, s, 'c', at, n)); }
export function deleteRows(doc, at, n) { const s = doc.wb.cur; doc.change('Xóa hàng', s, 'book', () => shiftAll(doc, s, 'r', at, -n)); }
export function deleteCols(doc, at, n) { const s = doc.wb.cur; doc.change('Xóa cột', s, 'book', () => shiftAll(doc, s, 'c', at, -n)); }
export function insertSelRows(doc, above) { const rg = doc.grid.selRange(); insertRows(doc, rg.r1, rg.r2 - rg.r1 + 1); }
export function insertSelCols(doc) { const rg = doc.grid.selRange(); insertCols(doc, rg.c1, rg.c2 - rg.c1 + 1); }
export function deleteSelRows(doc) { const rg = clipRange(doc.wb.cur, doc.grid.selRange()); deleteRows(doc, rg.r1, rg.r2 - rg.r1 + 1); }
export function deleteSelCols(doc) { const rg = clipRange(doc.wb.cur, doc.grid.selRange()); deleteCols(doc, rg.c1, rg.c2 - rg.c1 + 1); }
// shift cells within a band (insert/delete cells)
function shiftCells(doc, rg, dir, insert) {
  const s = doc.wb.cur; const u = s.used();
  const moved = new Map();
  const w = rg.c2 - rg.c1 + 1, ht = rg.r2 - rg.r1 + 1;
  s.cells.forEach((x, k) => {
    let r = Math.floor(k / 20000), c = k % 20000;
    if (dir === 'down' && c >= rg.c1 && c <= rg.c2 && r >= rg.r1) { if (!insert && r <= rg.r2) return; r += insert ? ht : -ht; }
    else if (dir === 'right' && r >= rg.r1 && r <= rg.r2 && c >= rg.c1) { if (!insert && c <= rg.c2) return; c += insert ? w : -w; }
    moved.set(key(r, c), x);
  });
  s.cells = moved;
  const f = [];
  doc.wb.sheets.forEach(sh => sh.cells.forEach(x => { if (x.f) f.push(x); }));
}
export function insertDialog(doc) {
  const rg = doc.grid.selRange();
  if (isFullRow(rg)) return insertSelRows(doc); if (isFullCol(rg)) return insertSelCols(doc);
  formDlg('Chèn', [{ key: 'm', type: 'radio', value: 'down', options: [['right', 'Dịch ô sang phải'], ['down', 'Dịch ô xuống dưới'], ['row', 'Toàn bộ hàng'], ['col', 'Toàn bộ cột']] }], { width: 320 }).then(v => {
    if (!v) return;
    if (v.m === 'row') return insertSelRows(doc); if (v.m === 'col') return insertSelCols(doc);
    doc.change('Chèn ô', doc.wb.cur, 'sheet', () => shiftCells(doc, rg, v.m === 'right' ? 'right' : 'down', true));
  });
}
export function deleteDialog(doc) {
  const rg = doc.grid.selRange();
  if (isFullRow(rg)) return deleteSelRows(doc); if (isFullCol(rg)) return deleteSelCols(doc);
  formDlg('Xóa', [{ key: 'm', type: 'radio', value: 'up', options: [['left', 'Dịch ô sang trái'], ['up', 'Dịch ô lên trên'], ['row', 'Toàn bộ hàng'], ['col', 'Toàn bộ cột']] }], { width: 320 }).then(v => {
    if (!v) return;
    if (v.m === 'row') return deleteSelRows(doc); if (v.m === 'col') return deleteSelCols(doc);
    doc.change('Xóa ô', doc.wb.cur, 'sheet', () => shiftCells(doc, rg, v.m === 'left' ? 'right' : 'down', false));
  });
}
export function hideRows(doc, v) { const s = doc.wb.cur; doc.change(v ? 'Ẩn hàng' : 'Bỏ ẩn hàng', s, 'props', () => selRanges(doc).forEach(rg => { const R = clipRange(s, rg); for (let r = R.r1; r <= R.r2; r++) { if (v) s.hiddenR[r] = true; else { delete s.hiddenR[r]; if (s.rowH[r] === 0) delete s.rowH[r]; } } })); }
export function hideCols(doc, v) { const s = doc.wb.cur; doc.change(v ? 'Ẩn cột' : 'Bỏ ẩn cột', s, 'props', () => selRanges(doc).forEach(rg => { const R = clipRange(s, rg); for (let c = R.c1; c <= R.c2; c++) { if (v) s.hiddenC[c] = true; else { delete s.hiddenC[c]; if (s.colW[c] === 0) delete s.colW[c]; } } })); }
export async function rowHeightDialog(doc) { const s = doc.wb.cur; const v = await promptDlg('Chiều cao hàng (điểm):', (s.rowHeight(s.sel.r) * 0.75).toFixed(2), 'Chiều cao hàng'); if (v == null) return; const px = Math.round(+v / 0.75); if (isNaN(px)) return; doc.change('Chiều cao hàng', s, 'props', () => selRanges(doc).forEach(rg => { const R = clipRange(s, rg); for (let r = R.r1; r <= R.r2; r++) { s.rowH[r] = px; (s.rowHManual = s.rowHManual || {})[r] = 1; } })); }
export async function colWidthDialog(doc) { const s = doc.wb.cur; const v = await promptDlg('Độ rộng cột (ký tự):', (s.colWidth(s.sel.c) / 7).toFixed(2), 'Độ rộng cột'); if (v == null) return; const px = Math.round(+v * 7 + 5); if (isNaN(px)) return; doc.change('Độ rộng cột', s, 'props', () => selRanges(doc).forEach(rg => { const R = clipRange(s, rg); for (let c = R.c1; c <= R.c2; c++) s.colW[c] = px; })); }
export function autoFit(doc, axis, i) {
  const s = doc.wb.cur; const g = doc.grid;
  const targets = []; const rg = doc.grid.selRange();
  if (i == null) { if (axis === 'c') for (let c = rg.c1; c <= Math.min(rg.c2, s.used().c2); c++) targets.push(c); else for (let r = rg.r1; r <= Math.min(rg.r2, s.used().r2); r++) targets.push(r); }
  else if (axis === 'c' && isFullCol(rg) && i >= rg.c1 && i <= rg.c2) for (let c = rg.c1; c <= rg.c2; c++) targets.push(c);
  else if (axis === 'r' && isFullRow(rg) && i >= rg.r1 && i <= rg.r2) for (let r = rg.r1; r <= rg.r2; r++) targets.push(r);
  else targets.push(i);
  doc.change(axis === 'c' ? 'Tự động khớp độ rộng cột' : 'Tự động khớp chiều cao hàng', s, 'props', () => {
    if (axis === 'c') {
      const u = s.used();
      targets.forEach(c => {
        let w = 0;
        for (let r = u.r1; r <= u.r2; r++) {
          const cell = s.get(r, c); if (!cell || cell.v == null) continue; if (s.mergeAt(r, c)) continue;
          const st = cell.s || {}; const t = format(cell.v, st.nf).t;
          const z = g.z; const fs = g.font(st);
          const tw = Math.max.apply(null, String(t).split('\n').map(x => g.measure(x, fs))) / z;
          if (tw > w) w = tw;
        }
        if (w) s.colW[c] = Math.ceil(w + 10); else delete s.colW[c];
      });
    } else targets.forEach(r => { if (s.rowHManual) delete s.rowHManual[r]; delete s.rowH[r]; autoRowHeight(doc, s, r); });
  });
}

// ---------------------------------------------------------------- sheets
export function insertSheet(doc) { doc.change('Chèn trang tính', doc.wb.cur, 'book', () => { const i = doc.wb.active + 1; doc.wb.addSheet(doc.wb.nextName(), i); doc.wb.active = i; }); doc.renderTabs(); doc.activateSheet(doc.wb.sheets[doc.wb.active]); }
export async function deleteSheet(doc, s) {
  if (s.cells.size && !(await confirmDlg('Trang tính "' + s.name + '" có dữ liệu. Xóa vĩnh viễn trang tính này?', 'Xóa trang tính', 'Xóa'))) return;
  doc.change('Xóa trang tính', s, 'book', () => { const i = doc.wb.sheets.indexOf(s); const nm = s.name; doc.wb.sheets.splice(i, 1); doc.wb.active = Math.max(0, Math.min(i, doc.wb.sheets.length - 1)); doc.wb.sheets.forEach(sh => sh.cells.forEach(x => { if (x.f && x.f.indexOf(nm) >= 0) x.f = x.f.replace(new RegExp("('?)" + nm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "\\1!(\\$?[A-Z]+\\$?\\d+(:\\$?[A-Z]+\\$?\\d+)?)", 'g'), '#REF!'); })); });
  doc.renderTabs(); doc.activateSheet(doc.wb.cur);
}
export function renameSheet(doc, s, name) {
  const old = s.name;
  doc.change('Đổi tên trang tính', s, 'book', () => {
    s.name = name;
    const q = n => /^[A-Za-z_][\w.]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'";
    doc.wb.sheets.forEach(sh => sh.cells.forEach(x => { if (x.f) x.f = x.f.split(q(old) + '!').join(q(name) + '!').split("'" + old + "'!").join(q(name) + '!'); }));
  });
  doc.renderTabs();
}
export function moveSheet(doc, from, to) { if (from === to) return; doc.change('Di chuyển trang tính', doc.wb.cur, 'book', () => { const [s] = doc.wb.sheets.splice(from, 1); doc.wb.sheets.splice(to, 0, s); doc.wb.active = to; }); doc.renderTabs(); doc.activateSheet(doc.wb.cur); }
export async function moveCopySheetDialog(doc, s) {
  const opts = doc.wb.sheets.map((x, i) => [String(i), x.name]).concat([[String(doc.wb.sheets.length), '(chuyển tới cuối)']]);
  const v = await formDlg('Di chuyển hoặc sao chép', [{ key: 'to', type: 'select', label: 'Trước trang tính:', options: opts, value: String(doc.wb.sheets.indexOf(s) + 1) }, { key: 'copy', type: 'checkbox', label: 'Tạo bản sao' }]);
  if (!v) return;
  const to = +v.to;
  if (v.copy) {
    doc.change('Sao chép trang tính', s, 'book', () => {
      const ns = doc.wb.addSheet(uniqueName(doc, s.name + ' (2)'), Math.min(to, doc.wb.sheets.length));
      applySheetState(ns, sheetState(s)); ns.name = uniqueName(doc, s.name + ' (2)'); doc.wb.active = doc.wb.sheets.indexOf(ns);
    });
    doc.renderTabs(); doc.activateSheet(doc.wb.cur);
  } else moveSheet(doc, doc.wb.sheets.indexOf(s), Math.min(to > doc.wb.sheets.indexOf(s) ? to - 1 : to, doc.wb.sheets.length - 1));
}
function uniqueName(doc, base) { let n = base, i = 2; while (doc.wb.sheets.some(s => s.name === n)) n = base.replace(/ \(\d+\)$/, '') + ' (' + (i++) + ')'; return n; }
export async function unhideSheetDialog(doc) { const hid = doc.wb.sheets.filter(s => s.hidden); const v = await formDlg('Bỏ ẩn trang tính', [{ key: 's', type: 'select', label: 'Bỏ ẩn trang tính:', options: hid.map(s => [s.name, s.name]), value: hid[0].name }]); if (!v) return; const s = doc.wb.sheet(v.s); s.hidden = false; doc.renderTabs(); doc.activateSheet(s); doc.markDirty(); }
export async function toggleProtect(doc) {
  const s = doc.wb.cur;
  if (s.protected) {
    if (s.protected.pw) { const p = await promptDlg('Mật khẩu để bỏ bảo vệ:', '', 'Bỏ bảo vệ trang tính', { type: 'password' }); if (p !== s.protected.pw) return toast('Mật khẩu không đúng', 'err'); }
    s.protected = null; toast('Đã bỏ bảo vệ trang tính'); doc.markDirty(); doc.refreshUi(); return;
  }
  const v = await formDlg('Bảo vệ trang tính', [{ key: 'pw', label: 'Mật khẩu (không bắt buộc):', type: 'password' }, { key: 'fmt', type: 'checkbox', label: 'Cho phép định dạng ô', value: false }, { key: 'sel', type: 'checkbox', label: 'Cho phép chọn ô bị khóa', value: true }, { type: 'html', el: h('div.hint', 'Mặc định mọi ô đều bị khóa. Để cho phép sửa một số ô: chọn ô → Định dạng ô (Ctrl+1) → Bảo vệ → bỏ chọn "Đã khóa" trước khi bảo vệ.') }]);
  if (!v) return; s.protected = { pw: v.pw || '', fmt: v.fmt }; toast('Đã bảo vệ trang tính', 'ok'); doc.markDirty(); doc.refreshUi();
}

// ---------------------------------------------------------------- clipboard
export function copy(doc, cut, e) {
  const s = doc.wb.cur; const rg = clipRange(s, doc.grid.selRange());
  if (doc.wb.cur.sel.ranges.length > 1) { toast('Không thể sao chép nhiều vùng chọn cùng lúc', 'err'); return; }
  const rows = [], htmlRows = [];
  const cells = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row = [], hrow = [];
    if (s.hiddenR[r] && !cut) { continue; }
    for (let c = rg.c1; c <= rg.c2; c++) {
      const cell = s.get(r, c); const v = s.value(r, c);
      const t = v == null ? '' : format(v, cell && cell.s && cell.s.nf).t;
      row.push(/[\t\n"]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t);
      const st = cell && cell.s || {};
      const css = [st.b ? 'font-weight:bold' : '', st.i ? 'font-style:italic' : '', st.color ? 'color:' + st.color : '', st.fill ? 'background:' + st.fill : '', st.ha ? 'text-align:' + st.ha : '', st.sz ? 'font-size:' + st.sz + 'pt' : ''].filter(Boolean).join(';');
      hrow.push('<td' + (css ? ' style="' + css + '"' : '') + (typeof v === 'number' ? ' x:num="' + v + '"' : '') + '>' + esc(t) + '</td>');
      cells.push([r - rg.r1, c - rg.c1, cell ? Object.assign({}, cell) : null]);
    }
    rows.push(row.join('\t')); htmlRows.push('<tr>' + hrow.join('') + '</tr>');
  }
  const text = rows.join('\r\n'), html = '<table>' + htmlRows.join('') + '</table>';
  doc.clip = { sheet: s, range: rg, cells, cut, text, merges: s.merges.filter(m => m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2).map(m => ({ r1: m.r1 - rg.r1, c1: m.c1 - rg.c1, r2: m.r2 - rg.r1, c2: m.c2 - rg.c1 })), colW: range(rg.c1, rg.c2).map(c => s.colW[c]) };
  doc.grid.copyRange = { sheet: s, range: rg }; doc.grid.redraw();
  if (e && e.clipboardData) { e.clipboardData.setData('text/plain', text); e.clipboardData.setData('text/html', html); }
  else { const ta = h('textarea', { style: { position: 'fixed', left: '-2000px', top: '0' } }); ta.value = text; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); } catch (er) {} ta.parentNode.removeChild(ta); doc.grid.focusSink(); }
  toast(cut ? 'Đã cắt — chọn đích rồi Dán (Ctrl+V)' : 'Đã sao chép ' + (rg.r2 - rg.r1 + 1) + '×' + (rg.c2 - rg.c1 + 1) + ' ô');
}
export function pasteFromEvent(doc, e) {
  const cd = e.clipboardData; if (!cd) return;
  const text = cd.getData('text/plain'); const html = cd.getData('text/html');
  // image paste
  if (cd.items) for (const it of Array.prototype.slice.call(cd.items)) if (/^image\//.test(it.type)) { const f = it.getAsFile(); const r = new FileReader(); r.onload = () => doc.objects.insertImage(r.result); r.readAsDataURL(f); return; }
  if (doc.clip && doc.clip.text === text.replace(/\r?\n$/, '').replace(/\n/g, '\r\n').replace(/\r\r/g, '\r') || (doc.clip && text.replace(/\r/g, '') === doc.clip.text.replace(/\r/g, ''))) return paste(doc, 'all');
  pasteExternal(doc, text, html);
}
export function paste(doc, mode) {
  const cl = doc.clip;
  if (!cl) { if (navigator.clipboard && navigator.clipboard.readText) { navigator.clipboard.readText().then(t => pasteExternal(doc, t, '')).catch(() => toast('Dùng Ctrl+V để dán từ ứng dụng khác', 'err')); return; } return toast('Bộ nhớ tạm trống. Dùng Ctrl+V để dán từ ứng dụng khác.'); }
  const s = doc.wb.cur; const dst = doc.grid.selRange();
  const rows = cl.range.r2 - cl.range.r1 + 1, cols = cl.range.c2 - cl.range.c1 + 1;
  const tr = mode === 'transpose';
  const H = tr ? cols : rows, W = tr ? rows : cols;
  // repeat to fill larger destination when it's a multiple
  const repR = Math.max(1, Math.floor((dst.r2 - dst.r1 + 1) / H)), repC = Math.max(1, Math.floor((dst.c2 - dst.c1 + 1) / W));
  const target = { r1: dst.r1, c1: dst.c1, r2: dst.r1 + H * repR - 1, c2: dst.c1 + W * repC - 1 };
  doc.change(mode === 'values' ? 'Dán giá trị' : mode === 'formats' ? 'Dán định dạng' : 'Dán', s, cl.cut && cl.sheet === s ? 'sheet' : target, () => {
    if (cl.cut) { const src = cl.sheet; for (let r = cl.range.r1; r <= cl.range.r2; r++) for (let c = cl.range.c1; c <= cl.range.c2; c++) src.cells.delete(key(r, c)); }
    for (let ir = 0; ir < repR; ir++) for (let ic2 = 0; ic2 < repC; ic2++) cl.cells.forEach(([dr, dc, cell]) => {
      const r = target.r1 + ir * H + (tr ? dc : dr), c = target.c1 + ic2 * W + (tr ? dr : dc);
      if (r >= MAX_ROWS || c >= MAX_COLS) return;
      const old = s.get(r, c) || {};
      if (!cell) { if (mode === 'all' || mode === 'transpose' || !mode) s.cells.delete(key(r, c)); return; }
      const shiftR = r - (cl.range.r1 + dr), shiftC = c - (cl.range.c1 + dc);
      if (mode === 'formats') { s.setCell(r, c, { s: cell.s || null }); return; }
      if (mode === 'values') { s.setCell(r, c, { v: cell.v, f: null }); return; }
      if (mode === 'formulas') { s.setCell(r, c, { v: cell.v, f: cell.f ? (cl.cut ? cell.f : shiftFormula(cell.f, shiftR, shiftC)) : null }); return; }
      if (mode === 'nobord') { s.setCell(r, c, Object.assign({}, cell, { f: cell.f ? shiftFormula(cell.f, shiftR, shiftC) : null, s: cell.s ? mergeStyle(cell.s, { bd: null }) : null })); return; }
      if (mode === 'link') { s.setCell(r, c, { f: (cl.sheet !== s ? q(cl.sheet.name) + '!' : '') + addr(cl.range.r1 + dr, cl.range.c1 + dc, true, true), v: null }); return; }
      const nc = Object.assign({}, cell); if (nc.f) nc.f = cl.cut ? cell.f : shiftFormula(cell.f, shiftR, shiftC);
      if (mode === 'transpose' && nc.f) nc.f = shiftFormula(cell.f, shiftR, shiftC);
      delete nc._ast; delete nc._src; delete nc.spillTo;
      s.cells.set(key(r, c), nc);
    });
    if (mode !== 'values' && mode !== 'formulas' && !tr) cl.merges.forEach(m => s.merges.push({ r1: target.r1 + m.r1, c1: target.c1 + m.c1, r2: target.r1 + m.r2, c2: target.c1 + m.c2 }));
    if (mode === 'colw') cl.colW.forEach((w, i) => { if (w != null) s.colW[target.c1 + i] = w; });
    s.sel.ranges = [target];
  });
  if (cl.cut) { doc.clip = null; doc.grid.copyRange = null; }
  doc.pasteOpts = true;
}
function q(n) { return /^[A-Za-z_][\w.]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'"; }
export function pasteExternal(doc, text, html) {
  const s = doc.wb.cur; const dst = doc.grid.selRange();
  let grid = null, styles = null;
  if (html && /<table/i.test(html)) {
    const div = document.createElement('div'); div.innerHTML = html.replace(/<style[\s\S]*?<\/style>/gi, '');
    const trs = div.querySelectorAll('tr'); grid = []; styles = [];
    Array.prototype.forEach.call(trs, tr => { const row = [], srow = []; Array.prototype.forEach.call(tr.querySelectorAll('td,th'), td => { row.push(td.textContent.replace(/ /g, ' ').trim()); const cs = td.getAttribute('style') || ''; const st = {}; if (/font-weight:\s*(bold|[6-9]00)/i.test(cs) || td.querySelector('b,strong') || td.tagName === 'TH') st.b = true; if (/font-style:\s*italic/i.test(cs) || td.querySelector('i,em')) st.i = true; const bg = /background(?:-color)?:\s*([^;]+)/i.exec(cs); if (bg && !/transparent|white|#fff\b|#ffffff/i.test(bg[1])) st.fill = bg[1].trim(); const col = /(?:^|;)\s*color:\s*([^;]+)/i.exec(cs); if (col) st.color = col[1].trim(); srow.push(Object.keys(st).length ? st : null); const span = +td.getAttribute('colspan') || 1; for (let k = 1; k < span; k++) { row.push(''); srow.push(null); } }); grid.push(row); styles.push(srow); });
  }
  if (!grid) grid = text.replace(/\r?\n$/, '').split(/\r?\n/).map(l => splitTsv(l));
  if (!grid.length) return;
  const target = { r1: dst.r1, c1: dst.c1, r2: dst.r1 + grid.length - 1, c2: dst.c1 + Math.max.apply(null, grid.map(r => r.length)) - 1 };
  doc.change('Dán', s, target, () => {
    grid.forEach((row, i) => row.forEach((t, j) => {
      const r = dst.r1 + i, c = dst.c1 + j;
      const p = /^=/.test(t) ? null : parseInput(t);
      const st = styles && styles[i] && styles[i][j];
      const old = s.get(r, c);
      const patch = p ? { v: p.v, f: null } : { f: normalizeFormula(t.slice(1)), v: null };
      let ns = old && old.s; if (p && p.nf) ns = mergeStyle(ns, { nf: p.nf }); if (st) ns = mergeStyle(ns, st);
      patch.s = ns || null;
      s.setCell(r, c, patch);
    }));
    s.sel.ranges = [target];
  });
}
function splitTsv(line) {
  const out = []; let cur = '', qt = false;
  for (let i = 0; i < line.length; i++) { const ch = line[i]; if (qt) { if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') qt = false; else cur += ch; } else if (ch === '"' && cur === '') qt = true; else if (ch === '\t') { out.push(cur); cur = ''; } else cur += ch; }
  out.push(cur); return out;
}
export function pasteMenu(doc) {
  return [
    { label: 'Dán (P)', icon: 'clipboard-paste', run: () => paste(doc, 'all') },
    { label: 'Công thức (F)', icon: 'function-square', run: () => paste(doc, 'formulas') },
    { label: 'Giữ độ rộng cột nguồn (W)', icon: 'move-horizontal', run: () => paste(doc, 'colw') },
    { label: 'Không viền (B)', icon: 'square-dashed', run: () => paste(doc, 'nobord') },
    { label: 'Hoán đổi hàng/cột (T)', icon: 'arrow-down-up', run: () => paste(doc, 'transpose') },
    { sep: true },
    { label: 'Chỉ giá trị (V)', icon: 'hash', run: () => paste(doc, 'values') },
    { label: 'Chỉ định dạng (R)', icon: 'paintbrush', run: () => paste(doc, 'formats') },
    { label: 'Dán liên kết (N)', icon: 'link', run: () => paste(doc, 'link') },
    { label: 'Dán dưới dạng ảnh', icon: 'image', run: () => doc.objects.pasteAsPicture() },
    { sep: true },
    { label: 'Dán đặc biệt…', icon: 'clipboard-list', run: () => pasteSpecialDialog(doc) }
  ];
}
async function pasteSpecialDialog(doc) {
  const v = await formDlg('Dán đặc biệt', [
    { key: 'what', type: 'radio', label: 'Dán', value: 'all', options: [['all', 'Tất cả'], ['formulas', 'Công thức'], ['values', 'Giá trị'], ['formats', 'Định dạng'], ['colw', 'Độ rộng cột'], ['nobord', 'Tất cả trừ viền']] },
    { key: 'op', type: 'radio', label: 'Phép toán', value: 'none', options: [['none', 'Không'], ['add', 'Cộng'], ['sub', 'Trừ'], ['mul', 'Nhân'], ['div', 'Chia']] },
    { key: 'skip', type: 'checkbox', label: 'Bỏ qua ô trống' }, { key: 'tr', type: 'checkbox', label: 'Hoán đổi (Transpose)' }
  ], { width: 420 });
  if (!v) return;
  if (v.op !== 'none' && doc.clip) {
    const s = doc.wb.cur; const dst = doc.grid.selRange(); const cl = doc.clip;
    doc.change('Dán đặc biệt', s, { r1: dst.r1, c1: dst.c1, r2: dst.r1 + cl.range.r2 - cl.range.r1, c2: dst.c1 + cl.range.c2 - cl.range.c1 }, () => cl.cells.forEach(([dr, dc, cell]) => {
      if (!cell || typeof cell.v !== 'number') return; const r = dst.r1 + dr, c = dst.c1 + dc; const cur = s.value(r, c); const a = typeof cur === 'number' ? cur : 0; const b = cell.v;
      const res = v.op === 'add' ? a + b : v.op === 'sub' ? a - b : v.op === 'mul' ? a * b : (b === 0 ? E.DIV0 : a / b); s.setCell(r, c, { v: res, f: null });
    }));
    return;
  }
  paste(doc, v.tr ? 'transpose' : v.what);
}

// ---------------------------------------------------------------- moving & filling
export function moveRange(doc, rg, r0, c0, copyMode) {
  const s = doc.wb.cur; const dr = r0 - rg.r1, dc = c0 - rg.c1;
  doc.change(copyMode ? 'Sao chép bằng kéo thả' : 'Di chuyển ô', s, 'sheet', () => {
    const items = [];
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const x = s.cells.get(key(r, c)); items.push([r, c, x]); if (!copyMode) s.cells.delete(key(r, c)); }
    items.forEach(([r, c, x]) => { const k = key(r + dr, c + dc); if (x) { const n = Object.assign({}, x); delete n._ast; if (n.f && copyMode) n.f = shiftFormula(n.f, dr, dc); s.cells.set(k, n); } else s.cells.delete(k); });
    if (!copyMode) {
      // references to moved cells follow them
      doc.wb.sheets.forEach(sh => sh.cells.forEach(x => { if (x.f) x.f = remapRefs(x.f, sh.name, s.name, rg, dr, dc); }));
      s.merges = s.merges.map(m => (m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2) ? { r1: m.r1 + dr, c1: m.c1 + dc, r2: m.r2 + dr, c2: m.c2 + dc } : m);
    }
    s.sel.ranges = [{ r1: rg.r1 + dr, c1: rg.c1 + dc, r2: rg.r2 + dr, c2: rg.c2 + dc }]; s.sel.r = rg.r1 + dr; s.sel.c = rg.c1 + dc;
  });
}
function remapRefs(f, fs, sheetName, rg, dr, dc) {
  try {
    const toks = tokenize(f); let out = '', last = 0;
    for (const t of toks) {
      if (t.t !== 'ref') continue; const R = t.v; if ((R.sheet || fs).toLowerCase() !== sheetName.toLowerCase()) continue;
      if (R.r1 >= rg.r1 && R.r2 <= rg.r2 && R.c1 >= rg.c1 && R.c2 <= rg.c2) {
        const len = formulaRefs(f.slice(t.p)).length ? formulaRefs(f.slice(t.p))[0].len : 0; if (!len) continue;
        const n = Object.assign({}, R, { r1: R.r1 + dr, r2: R.r2 + dr, c1: R.c1 + dc, c2: R.c2 + dc });
        out += f.slice(last, t.p) + refTxt(n); last = t.p + len;
      }
    }
    return out + f.slice(last);
  } catch (e) { return f; }
}
function refTxt(R) { const sh = R.sheet ? q(R.sheet) + '!' : ''; const a = (R.ac1 ? '$' : '') + colName(R.c1) + (R.ar1 ? '$' : '') + (R.r1 + 1); return sh + a + (R.kind === 'range' ? ':' + (R.ac2 ? '$' : '') + colName(R.c2) + (R.ar2 ? '$' : '') + (R.r2 + 1) : ''); }

const LISTS = [
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Chủ nhật', 'Thứ hai', 'Thứ ba', 'Thứ tư', 'Thứ năm', 'Thứ sáu', 'Thứ bảy'], ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'],
  ['Tháng 1', 'Tháng 2', 'Tháng 3', 'Tháng 4', 'Tháng 5', 'Tháng 6', 'Tháng 7', 'Tháng 8', 'Tháng 9', 'Tháng 10', 'Tháng 11', 'Tháng 12'], ['Quý 1', 'Quý 2', 'Quý 3', 'Quý 4'], ['Q1', 'Q2', 'Q3', 'Q4']
];
function listOf(v) { if (typeof v !== 'string') return null; for (const L of LISTS) { const i = L.findIndex(x => x.toLowerCase() === v.toLowerCase()); if (i >= 0) return { L, i }; } return null; }
// series generator for a 1-D source vector -> function(k) giving k-th value after the source (k from 0)
function seriesFor(vals, cells, ctrl) {
  const n = vals.length;
  const allNum = vals.every(v => typeof v === 'number');
  if (allNum && !ctrl) {
    if (n === 1) return k => ({ v: vals[0] });
    // linear trend
    const xs = vals.map((_, i) => i); const mx = (n - 1) / 2, my = vals.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0; xs.forEach((x, i) => { sxy += (x - mx) * (vals[i] - my); sxx += (x - mx) * (x - mx); });
    const b = sxx ? sxy / sxx : 0, a = my - b * mx;
    return k => ({ v: +(a + b * (n + k)).toPrecision(15) });
  }
  if (allNum && ctrl && n === 1) return k => ({ v: vals[0] + k + 1 });
  const L = vals.map(listOf);
  if (L.every(x => x && x.L === L[0].L)) { const step = n > 1 ? L[1].i - L[0].i : 1; return k => ({ v: L[0].L[(((L[n - 1].i + step * (k + 1)) % L[0].L.length) + L[0].L.length) % L[0].L.length] }); }
  // text with trailing number: "Item 1" -> "Item 2"
  const m = vals.map(v => typeof v === 'string' ? /^(.*?)(\d+)(\D*)$/.exec(v) : null);
  if (m.every(x => x) && m.every(x => x[1] === m[0][1] && x[3] === m[0][3])) {
    const nums = m.map(x => +x[2]); const step = n > 1 ? nums[n - 1] - nums[n - 2] : 1; const w = m[0][2].length;
    return k => { let d = String(nums[n - 1] + step * (k + 1)); if (/^0/.test(m[0][2])) while (d.length < w) d = '0' + d; return { v: m[0][1] + d + m[0][3] }; };
  }
  // repeat (copy)
  return k => ({ cell: cells[k % n] });
}
export function fillPreview(doc, src, dst) {
  if (dst.r1 === src.r1 && dst.r2 === src.r2 && dst.c1 === src.c1 && dst.c2 === src.c2) return '';
  const s = doc.wb.cur; const down = dst.r2 > src.r2, up = dst.r1 < src.r1, right = dst.c2 > src.c2;
  const r = down ? dst.r2 : up ? dst.r1 : src.r1, c = right ? dst.c2 : dst.c1 < src.c1 ? dst.c1 : src.c1;
  const tmp = computeFill(doc, src, dst, true);
  const v = tmp.get(key(r, c)); if (!v) return '';
  if (v.f) return '=' + v.f;
  return v.v == null ? '' : format(v.v, v.s && v.s.nf).t;
}
function computeFill(doc, src, dst, preview, ctrl) {
  const s = doc.wb.cur; const out = new Map();
  const vertical = dst.r1 !== src.r1 || dst.r2 !== src.r2;
  if (vertical) {
    for (let c = src.c1; c <= src.c2; c++) {
      const cells = [], vals = []; for (let r = src.r1; r <= src.r2; r++) { const x = s.get(r, c); cells.push(x ? Object.assign({}, x) : null); vals.push(x && !x.f ? x.v : (x && x.f ? null : null)); }
      const hasF = cells.some(x => x && x.f); const gen = hasF || vals.some(v => v == null) ? null : seriesFor(vals, cells, ctrl);
      const forward = dst.r2 > src.r2;
      const count = forward ? dst.r2 - src.r2 : src.r1 - dst.r1;
      for (let k = 0; k < count; k++) {
        const r = forward ? src.r2 + 1 + k : src.r1 - 1 - k;
        const si = forward ? k % cells.length : cells.length - 1 - (k % cells.length);
        const base = cells[si]; const srcR = src.r1 + si;
        let nc;
        if (gen && forward) { const g = gen(k); nc = g.cell ? Object.assign({}, g.cell) : Object.assign({}, base || {}, { v: g.v }); }
        else if (gen && !forward && vals.every(v => typeof v === 'number') && vals.length > 1) { const step = vals[1] - vals[0]; nc = Object.assign({}, base || {}, { v: vals[0] - step * (k + 1) }); }
        else { nc = base ? Object.assign({}, base) : null; if (nc && nc.f) nc.f = shiftFormula(base.f, r - srcR, 0); }
        if (nc) { delete nc._ast; delete nc._src; delete nc.spillTo; }
        out.set(key(r, c), nc);
      }
    }
  } else {
    for (let r = src.r1; r <= src.r2; r++) {
      const cells = [], vals = []; for (let c = src.c1; c <= src.c2; c++) { const x = s.get(r, c); cells.push(x ? Object.assign({}, x) : null); vals.push(x && !x.f ? x.v : null); }
      const hasF = cells.some(x => x && x.f); const gen = hasF || vals.some(v => v == null) ? null : seriesFor(vals, cells, ctrl);
      const forward = dst.c2 > src.c2; const count = forward ? dst.c2 - src.c2 : src.c1 - dst.c1;
      for (let k = 0; k < count; k++) {
        const c = forward ? src.c2 + 1 + k : src.c1 - 1 - k;
        const si = forward ? k % cells.length : cells.length - 1 - (k % cells.length);
        const base = cells[si]; const srcC = src.c1 + si; let nc;
        if (gen && forward) { const g = gen(k); nc = g.cell ? Object.assign({}, g.cell) : Object.assign({}, base || {}, { v: g.v }); }
        else { nc = base ? Object.assign({}, base) : null; if (nc && nc.f) nc.f = shiftFormula(base.f, 0, c - srcC); }
        if (nc) { delete nc._ast; delete nc._src; }
        out.set(key(r, c), nc);
      }
    }
  }
  return out;
}
export function autoFill(doc, src, dst, ctrl) {
  const s = doc.wb.cur;
  if (dst.r1 === src.r1 && dst.r2 === src.r2 && dst.c1 === src.c1 && dst.c2 === src.c2) return;
  // shrinking selection = clear
  if (dst.r2 < src.r2 || dst.c2 < src.c2) { if (dst.r1 === src.r1 && dst.c1 === src.c1 && (dst.r2 <= src.r2 && dst.c2 <= src.c2)) { doc.change('Xóa bằng kéo', s, src, () => { for (let r = dst.r2 + 1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) s.setCell(r, c, { v: null, f: null }); for (let c = dst.c2 + 1; c <= src.c2; c++) for (let r = src.r1; r <= src.r2; r++) s.setCell(r, c, { v: null, f: null }); s.sel.ranges = [dst]; }); return; } }
  doc.change('Tự động điền', s, { r1: Math.min(src.r1, dst.r1), c1: Math.min(src.c1, dst.c1), r2: Math.max(src.r2, dst.r2), c2: Math.max(src.c2, dst.c2) }, () => {
    const m = computeFill(doc, src, dst, false, ctrl);
    m.forEach((x, k) => { if (x) s.cells.set(k, x); else s.cells.delete(k); });
    s.sel.ranges = [dst];
  });
  // auto-fill options (like Excel's smart tag)
  doc.fillOptions = { src, dst };
  const tag = h('button.xl-filltag', { title: 'Tùy chọn Tự động điền' }, ic('clipboard-list', 14));
  const rc = doc.grid.cellRect(dst.r2, dst.c2); tag.style.left = (rc.x + rc.w + 2) + 'px'; tag.style.top = (rc.y + rc.h + 2) + 'px';
  const old = $('.xl-filltag', doc.grid.el); if (old) old.parentNode.removeChild(old);
  doc.grid.el.appendChild(tag);
  tag.addEventListener('click', () => showMenu([
    { label: 'Sao chép ô', run: () => { doc.undo(); fillCopy(doc, src, dst); } },
    { label: 'Điền chuỗi (Fill Series)', run: () => { doc.undo(); autoFill(doc, src, dst, true); } },
    { label: 'Chỉ điền định dạng', run: () => { doc.undo(); fillFormatOnly(doc, src, dst); } },
    { label: 'Điền không định dạng', run: () => { doc.undo(); fillNoFormat(doc, src, dst); } },
    { label: 'Điền nhanh (Flash Fill)', run: () => { doc.undo(); flashFill(doc); } }
  ], { anchor: tag }));
  setTimeout(() => { if (tag.parentNode) tag.parentNode.removeChild(tag); }, 8000);
}
function fillCopy(doc, src, dst) { const s = doc.wb.cur; doc.change('Sao chép ô', s, dst, () => { for (let r = dst.r1; r <= dst.r2; r++) for (let c = dst.c1; c <= dst.c2; c++) { const sr = src.r1 + ((r - src.r1) % (src.r2 - src.r1 + 1) + (src.r2 - src.r1 + 1)) % (src.r2 - src.r1 + 1), sc = src.c1 + ((c - src.c1) % (src.c2 - src.c1 + 1) + (src.c2 - src.c1 + 1)) % (src.c2 - src.c1 + 1); const x = s.get(sr, sc); if (x) { const n = Object.assign({}, x); delete n._ast; if (n.f) n.f = shiftFormula(x.f, r - sr, c - sc); s.cells.set(key(r, c), n); } else s.cells.delete(key(r, c)); } }); }
function fillFormatOnly(doc, src, dst) { const s = doc.wb.cur; doc.change('Điền định dạng', s, dst, () => { for (let r = dst.r1; r <= dst.r2; r++) for (let c = dst.c1; c <= dst.c2; c++) { const x = s.get(src.r1 + (r - dst.r1) % (src.r2 - src.r1 + 1), src.c1 + (c - dst.c1) % (src.c2 - src.c1 + 1)); s.setCell(r, c, { s: x && x.s || null }); } }); }
function fillNoFormat(doc, src, dst) { const s = doc.wb.cur; const keep = new Map(); for (let r = dst.r1; r <= dst.r2; r++) for (let c = dst.c1; c <= dst.c2; c++) { const x = s.get(r, c); keep.set(key(r, c), x && x.s); } autoFill(doc, src, dst); keep.forEach((st, k) => { const x = s.cells.get(k); if (x) { x.s = st; } }); doc.grid.redraw(); }
export function autoFillDown(doc) {
  const s = doc.wb.cur; const rg = doc.grid.selRange();
  const nb = rg.c1 > 0 ? rg.c1 - 1 : rg.c2 + 1;
  let r = rg.r2; while (r + 1 < MAX_ROWS && doc.grid.cellHasContent(r + 1, nb)) r++;
  if (r === rg.r2) { const nb2 = rg.c2 + 1; while (r + 1 < MAX_ROWS && doc.grid.cellHasContent(r + 1, nb2)) r++; }
  if (r > rg.r2) autoFill(doc, rg, { r1: rg.r1, c1: rg.c1, r2: r, c2: rg.c2 });
}
export function fillDir(doc, dir) {
  const s = doc.wb.cur; const rg = clipRange(s, doc.grid.selRange());
  doc.change('Điền ' + dir, s, rg, () => {
    if (dir === 'down' || dir === 'up') for (let c = rg.c1; c <= rg.c2; c++) { const sr = dir === 'down' ? rg.r1 : rg.r2; const x = s.get(sr, c); for (let r = rg.r1; r <= rg.r2; r++) { if (r === sr) continue; if (x) { const n = Object.assign({}, x); delete n._ast; if (n.f) n.f = shiftFormula(x.f, r - sr, 0); s.cells.set(key(r, c), n); } else s.cells.delete(key(r, c)); } }
    else for (let r = rg.r1; r <= rg.r2; r++) { const sc = dir === 'right' ? rg.c1 : rg.c2; const x = s.get(r, sc); for (let c = rg.c1; c <= rg.c2; c++) { if (c === sc) continue; if (x) { const n = Object.assign({}, x); delete n._ast; if (n.f) n.f = shiftFormula(x.f, 0, c - sc); s.cells.set(key(r, c), n); } else s.cells.delete(key(r, c)); } }
  });
}
export async function seriesDialog(doc) {
  const s = doc.wb.cur; const rg = doc.grid.selRange();
  const v = await formDlg('Chuỗi (Series)', [{ key: 'dir', type: 'radio', label: 'Chuỗi theo', value: rg.r2 > rg.r1 ? 'col' : 'row', options: [['row', 'Hàng'], ['col', 'Cột']] }, { key: 'type', type: 'radio', label: 'Kiểu', value: 'linear', options: [['linear', 'Tuyến tính'], ['growth', 'Tăng trưởng'], ['date', 'Ngày']] }, { key: 'step', type: 'number', label: 'Giá trị bước:', value: 1 }, { key: 'stop', type: 'number', label: 'Giá trị dừng (tùy chọn):', value: '' }]);
  if (!v) return;
  const start = s.value(rg.r1, rg.c1); if (typeof start !== 'number') return toast('Ô đầu tiên phải là số hoặc ngày', 'err');
  doc.change('Điền chuỗi', s, rg, () => {
    let val = start; let i = 0;
    const cells = []; if (v.dir === 'col') for (let r = rg.r1; r <= (v.stop != null && rg.r1 === rg.r2 ? MAX_ROWS - 1 : rg.r2); r++) cells.push([r, rg.c1]); else for (let c = rg.c1; c <= (v.stop != null && rg.c1 === rg.c2 ? MAX_COLS - 1 : rg.c2); c++) cells.push([rg.r1, c]);
    for (const [r, c] of cells) { if (v.stop != null && v.stop !== '' && (v.step > 0 ? val > v.stop : val < v.stop)) break; if (i > 100000) break; s.setCell(r, c, { v: val, f: null }); val = v.type === 'growth' ? val * v.step : val + v.step; i++; }
  });
}
export function flashFill(doc) {
  // learn pattern from examples above in the same column using neighbor text (split / join / case)
  const s = doc.wb.cur; const c = s.sel.c; let r0 = s.sel.r;
  const ex = []; for (let r = r0 - 1; r >= 0 && ex.length < 3; r--) { const v = s.value(r, c); if (v == null || v === '') break; ex.unshift(r); }
  if (!ex.length) return toast('Nhập ví dụ mẫu ở ô phía trên rồi chạy Điền nhanh (Ctrl+E)', 'err');
  const srcCols = []; for (let cc = c - 1; cc >= Math.max(0, c - 5); cc--) srcCols.push(cc);
  const candidates = [];
  srcCols.forEach(sc => {
    const sv = r => String(s.value(r, sc) == null ? '' : s.value(r, sc));
    const tests = [
      ['upper', t => t.toUpperCase()], ['lower', t => t.toLowerCase()], ['proper', t => t.toLowerCase().replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase())],
      ['first', t => t.split(/\s+/)[0]], ['last', t => t.split(/\s+/).pop()], ['notfirst', t => t.split(/\s+/).slice(1).join(' ')], ['notlast', t => t.split(/\s+/).slice(0, -1).join(' ')],
      ['beforeat', t => t.split('@')[0]], ['afterat', t => t.split('@')[1] || ''], ['digits', t => t.replace(/\D/g, '')], ['initials', t => t.split(/\s+/).map(w => w[0] || '').join('').toUpperCase()],
      ['beforecomma', t => t.split(',')[0].trim()], ['aftercomma', t => (t.split(',')[1] || '').trim()], ['beforedash', t => t.split('-')[0].trim()], ['afterdash', t => t.split('-').slice(1).join('-').trim()]
    ];
    tests.forEach(([nm, fn]) => { if (ex.every(r => fn(sv(r)) === String(s.value(r, c)))) candidates.push({ fn, sc }); });
  });
  // two-column join  "A B"
  if (!candidates.length && c >= 2) { const sep = String(s.value(ex[0], c)).replace(String(s.value(ex[0], c - 2)), '').replace(String(s.value(ex[0], c - 1)), ''); if (ex.every(r => String(s.value(r, c)) === String(s.value(r, c - 2)) + sep + String(s.value(r, c - 1)))) candidates.push({ join: sep }); }
  if (!candidates.length) return toast('Điền nhanh không nhận ra mẫu. Hãy nhập thêm ví dụ.', 'err');
  const cand = candidates[0];
  let r = ex[ex.length - 1] + 1; const end = (() => { let e = r; const ref = cand.sc != null ? cand.sc : c - 1; while (e < MAX_ROWS && doc.grid.cellHasContent(e, ref)) e++; return e - 1; })();
  doc.change('Điền nhanh', s, { r1: r, c1: c, r2: end, c2: c }, () => { for (; r <= end; r++) { const v = cand.join != null ? String(s.value(r, c - 2) == null ? '' : s.value(r, c - 2)) + cand.join + String(s.value(r, c - 1) == null ? '' : s.value(r, c - 1)) : cand.fn(String(s.value(r, cand.sc) == null ? '' : s.value(r, cand.sc))); s.setCell(r, c, { v: parseInput(v).v, f: null }); } });
  toast('Điền nhanh: đã điền ' + (end - ex[ex.length - 1]) + ' ô', 'ok');
}

// ---------------------------------------------------------------- formulas helpers
export function normalizeFormula(f) {
  // upper-case function names & refs outside strings; close missing parens
  let out = '', inS = false;
  for (let i = 0; i < f.length; i++) { const ch = f[i]; if (ch === '"') inS = !inS; out += ch; }
  out = out.replace(/("[^"]*")|([A-Za-z_][\w.]*)(?=\s*\()/g, (m, s1, fn) => s1 || (FN[fn.toUpperCase()] ? fn.toUpperCase() : fn))
    .replace(/("[^"]*")|(\$?\b[a-z]{1,3}\$?\d+\b)/g, (m, s1, ref) => s1 || ref.toUpperCase());
  const open = (out.replace(/"[^"]*"/g, '').match(/\(/g) || []).length, close = (out.replace(/"[^"]*"/g, '').match(/\)/g) || []).length;
  if (open > close) out += ')'.repeat(open - close);
  return out;
}
export function inferFormulaFormat(doc, f) {
  if (/^(TODAY|DATE|EDATE|EOMONTH|WORKDAY(\.INTL)?|DATEVALUE)\(/i.test(f)) return 'dd/mm/yyyy';
  if (/^NOW\(/i.test(f)) return 'dd/mm/yyyy hh:mm';
  if (/^TIME(VALUE)?\(/i.test(f)) return 'hh:mm';
  // inherit format from first referenced cell for simple arithmetic/SUM
  const refs = formulaRefs(f);
  if (refs.length && /^(SUM|AVERAGE|MIN|MAX|ROUND|SUBTOTAL)?\(?[\sA-Z0-9$:+\-*/,.()!']*\)?$/i.test(f)) { const R = refs[0].ref; const s = R.sheet ? doc.wb.sheet(R.sheet) : doc.wb.cur; if (s) { const st = s.styleAt(R.r1, R.c1); if (st && st.nf && !/%/.test(st.nf)) return st.nf; } }
  return null;
}
export function autoSum(doc, fn) {
  const s = doc.wb.cur; const rg = doc.grid.selRange();
  const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
  if (single) {
    const r = rg.r1, c = rg.c1;
    // look up for contiguous numbers, else left
    let r1 = r - 1; while (r1 >= 0 && typeof s.value(r1, c) === 'number') r1--;
    let src;
    if (r1 < r - 1) src = { r1: r1 + 1, c1: c, r2: r - 1, c2: c };
    else { let c1 = c - 1; while (c1 >= 0 && typeof s.value(r, c1) === 'number') c1--; if (c1 < c - 1) src = { r1: r, c1: c1 + 1, r2: r, c2: c - 1 }; }
    const text = '=' + fn + '(' + (src ? rangeAddr(src) : '') + ')';
    doc.grid.beginEdit('enter', text);
    if (src) { doc.grid.editor.setSelectionRange(text.length - 1 - rangeAddr(src).length, text.length - 1); doc.grid.pointRef = { start: text.length - 1 - rangeAddr(src).length, end: text.length - 1, r1: src.r1, c1: src.c1, r2: src.r2, c2: src.c2 }; }
    else doc.grid.editor.setSelectionRange(text.length - 1, text.length - 1);
    doc.grid.updateEditRefs(); doc.grid.redraw();
    return;
  }
  // multi-cell selection: totals below each column (or right of each row)
  doc.change('Tự động tính ' + fn, s, { r1: rg.r1, c1: rg.c1, r2: rg.r2 + 1, c2: rg.c2 + 1 }, () => {
    const lastRowEmpty = (() => { for (let c = rg.c1; c <= rg.c2; c++) if (s.value(rg.r2, c) != null) return false; return true; })();
    if (rg.r2 > rg.r1 || lastRowEmpty) {
      const tr = lastRowEmpty ? rg.r2 : rg.r2 + 1, er = lastRowEmpty ? rg.r2 - 1 : rg.r2;
      for (let c = rg.c1; c <= rg.c2; c++) s.setCell(tr, c, { f: fn + '(' + rangeAddr({ r1: rg.r1, c1: c, r2: er, c2: c }) + ')', v: null });
    } else for (let r = rg.r1; r <= rg.r2; r++) s.setCell(r, rg.c2 + 1, { f: fn + '(' + rangeAddr({ r1: r, c1: rg.c1, r2: r, c2: rg.c2 }) + ')', v: null });
  });
}
export function todayText() { const d = new Date(); return d.getDate() + '/' + (d.getMonth() + 1) + '/' + d.getFullYear(); }
export function nowTimeText() { const d = new Date(); return d.getHours() + ':' + ('0' + d.getMinutes()).slice(-2); }
export function toggleShowFormulas(doc) { const s = doc.wb.cur; s.showFormulas = !s.showFormulas; doc.grid.redraw(); doc.refreshUi(); }

// ---------------------------------------------------------------- sort & filter
export function sortRange(doc, keys, rg0, hasHeader) {
  const s = doc.wb.cur;
  let rg = rg0 || currentRegion(doc);
  const head = hasHeader == null ? detectHeader(s, rg) : hasHeader;
  const R = { r1: rg.r1 + (head ? 1 : 0), c1: rg.c1, r2: rg.r2, c2: rg.c2 };
  if (R.r2 <= R.r1) return;
  doc.change('Sắp xếp', s, 'sheet', () => {
    const rows = [];
    for (let r = R.r1; r <= R.r2; r++) { const cells = []; for (let c = R.c1; c <= R.c2; c++) cells.push(s.cells.get(key(r, c)) || null); rows.push({ r, cells, h: s.rowH[r] }); }
    rows.sort((a, b) => {
      for (const k of keys) {
        const ia = k.c - R.c1; const va = a.cells[ia] ? a.cells[ia].v : null, vb = b.cells[ia] ? b.cells[ia].v : null;
        if (va == null && vb == null) continue; if (va == null) return 1; if (vb == null) return -1;
        let d;
        if (k.by === 'color') { const ca = a.cells[ia] && a.cells[ia].s && a.cells[ia].s.fill || '', cb = b.cells[ia] && b.cells[ia].s && b.cells[ia].s.fill || ''; d = ca === k.color ? (cb === k.color ? 0 : -1) : (cb === k.color ? 1 : 0); }
        else if (k.list) { const ia2 = k.list.indexOf(String(va)), ib = k.list.indexOf(String(vb)); d = (ia2 < 0 ? 999 : ia2) - (ib < 0 ? 999 : ib); }
        else d = cmp(va instanceof Err ? '￿' : va, vb instanceof Err ? '￿' : vb);
        if (d) return k.desc ? -d : d;
      }
      return a.r - b.r;
    });
    rows.forEach((row, i) => {
      const r = R.r1 + i;
      row.cells.forEach((x, j) => { const k = key(r, R.c1 + j); if (x) { const n = Object.assign({}, x); delete n._ast; if (n.f) n.f = shiftFormula(x.f, r - row.r, 0); s.cells.set(k, n); } else s.cells.delete(k); });
    });
  });
  toast('Đã sắp xếp ' + (R.r2 - R.r1 + 1) + ' hàng');
}
export function quickSort(doc, desc) { const s = doc.wb.cur; const rg = currentRegion(doc); sortRange(doc, [{ c: s.sel.c, desc }], rg); }
export function currentRegion(doc, r, c) {
  const s = doc.wb.cur; const g = doc.grid; const sel = doc.grid.selRange();
  if (r == null && (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) && !isFullCol(sel) && !isFullRow(sel)) return sel;
  r = r == null ? s.sel.r : r; c = c == null ? s.sel.c : c;
  let R = { r1: r, c1: c, r2: r, c2: c }; let changed = true;
  const has = (rr, cc) => rr >= 0 && cc >= 0 && g.cellHasContent(rr, cc);
  while (changed) {
    changed = false;
    if (R.r1 > 0 && range(R.c1 - 1, R.c2 + 1).some(cc => has(R.r1 - 1, cc))) { R.r1--; changed = true; }
    if (range(R.c1 - 1, R.c2 + 1).some(cc => has(R.r2 + 1, cc))) { R.r2++; changed = true; }
    if (R.c1 > 0 && range(R.r1 - 1, R.r2 + 1).some(rr => has(rr, R.c1 - 1))) { R.c1--; changed = true; }
    if (range(R.r1 - 1, R.r2 + 1).some(rr => has(rr, R.c2 + 1))) { R.c2++; changed = true; }
    if (R.r2 - R.r1 > 100000) break;
  }
  return R;
}
function detectHeader(s, rg) {
  if (rg.r2 === rg.r1) return false;
  let textTop = 0, numBelow = 0; for (let c = rg.c1; c <= rg.c2; c++) { const a = s.value(rg.r1, c), b = s.value(rg.r1 + 1, c); if (typeof a === 'string') textTop++; if (typeof b === 'number') numBelow++; }
  const st = s.styleAt(rg.r1, rg.c1);
  return textTop === rg.c2 - rg.c1 + 1 && (numBelow > 0 || (st && st.b)) || !!(st && st.b && textTop);
}
export async function sortDialog(doc) {
  const s = doc.wb.cur; const rg = currentRegion(doc); const head = detectHeader(s, rg);
  const colOpts = range(rg.c1, rg.c2).map(c => [String(c), head ? 'Cột ' + colName(c) + ' — ' + String(s.value(rg.r1, c) == null ? '' : s.value(rg.r1, c)) : 'Cột ' + colName(c)]);
  const levels = [{ c: s.sel.c, desc: false }];
  const body = h('div');
  const hdr = h('input', { type: 'checkbox' }); hdr.checked = head;
  const list = h('div.sort-levels');
  const draw = () => { list.innerHTML = ''; levels.forEach((lv, i) => {
    const sel = h('select.inp.small', colOpts.map(o => h('option', { value: o[0] }, o[1]))); sel.value = String(lv.c); sel.style.width = '240px';
    sel.addEventListener('change', () => { lv.c = +sel.value; });
    const ord = h('select.inp.small', h('option', { value: 'a' }, 'A → Z / Nhỏ → Lớn'), h('option', { value: 'd' }, 'Z → A / Lớn → Nhỏ')); ord.value = lv.desc ? 'd' : 'a'; ord.addEventListener('change', () => { lv.desc = ord.value === 'd'; });
    list.appendChild(h('div.form-row.inline', h('span', { style: { width: '90px' } }, i ? 'Sau đó theo' : 'Sắp xếp theo'), sel, ord, i ? h('button.btn', { onclick: () => { levels.splice(i, 1); draw(); } }, ic('trash-2', 14)) : null));
  }); };
  draw();
  body.appendChild(h('div', h('button.btn', { onclick: () => { levels.push({ c: rg.c1, desc: false }); draw(); } }, ic('plus', 14), ' Thêm mức'), h('label.chk', { style: { display: 'inline-flex', marginLeft: '12px' } }, hdr, h('span', 'Dữ liệu có tiêu đề'))));
  body.appendChild(list);
  body.appendChild(h('div.hint', 'Vùng sắp xếp: ' + rangeAddr(rg)));
  const ok = await dialog({ title: 'Sắp xếp (Sort)', width: 600, body, buttons: [{ label: 'OK', primary: true, value: true }, { label: 'Hủy', value: false }] });
  if (ok) sortRange(doc, levels, rg, hdr.checked);
}
export function toggleFilter(doc) {
  const s = doc.wb.cur;
  if (s.filter) { doc.change('Tắt bộ lọc', s, 'sheet', () => { for (const r in s.hiddenR) if (s.filterHidden && s.filterHidden[r]) delete s.hiddenR[r]; s.filter = null; s.filterHidden = null; }); return; }
  const rg = currentRegion(doc);
  if (rg.r1 === rg.r2 && !doc.grid.cellHasContent(rg.r1, rg.c1)) return toast('Chọn một ô trong vùng dữ liệu trước khi bật Bộ lọc', 'err');
  doc.change('Bật bộ lọc', s, 'props', () => { s.filter = { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, crit: {} }; });
  toast('Đã bật Bộ lọc — bấm mũi tên ở hàng tiêu đề để lọc/sắp xếp');
}
export function applyFilter(doc) {
  const s = doc.wb.cur; const f = s.filter; if (!f) return;
  s.filterHidden = s.filterHidden || {};
  for (const r in s.filterHidden) { delete s.hiddenR[r]; } s.filterHidden = {};
  // extend to data below
  let r2 = f.r2; while (doc.grid.cellHasContent(r2 + 1, f.c1) || range(f.c1, f.c2).some(c => doc.grid.cellHasContent(r2 + 1, c))) r2++; f.r2 = r2;
  f.active = false;
  for (let r = f.r1 + 1; r <= f.r2; r++) {
    let show = true;
    for (const c in f.crit) { const cr = f.crit[c]; if (!cr) continue; f.active = true; const v = s.value(r, +c); if (!filterMatch(cr, v, s, r, +c)) { show = false; break; } }
    if (!show) { s.hiddenR[r] = true; s.filterHidden[r] = true; }
  }
  if (s._ax) s._ax.r.dirty = true;
}
function filterMatch(cr, v, s, r, c) {
  if (cr.values) { const t = v == null ? '' : format(v, (s.styleAt(r, c) || {}).nf).t; return cr.values.indexOf(t) >= 0; }
  if (cr.color) { const st = s.styleAt(r, c); return (st && st.fill) === cr.color; }
  if (cr.top) return true;
  if (cr.cond) { const preds = cr.cond.map(cd => condPred(cd)); return cr.and ? preds.every(p => p(v)) : preds.some(p => p(v)); }
  return true;
}
function condPred(cd) {
  const t = cd.v; const n = parseInput(t).v;
  return v => {
    switch (cd.op) {
      case 'eq': return String(v == null ? '' : v).toLowerCase() === String(t).toLowerCase();
      case 'ne': return String(v == null ? '' : v).toLowerCase() !== String(t).toLowerCase();
      case 'gt': return typeof v === 'number' && v > n; case 'ge': return typeof v === 'number' && v >= n; case 'lt': return typeof v === 'number' && v < n; case 'le': return typeof v === 'number' && v <= n;
      case 'begins': return String(v || '').toLowerCase().indexOf(String(t).toLowerCase()) === 0;
      case 'ends': { const a = String(v || '').toLowerCase(), b = String(t).toLowerCase(); return a.slice(-b.length) === b; }
      case 'contains': return String(v || '').toLowerCase().indexOf(String(t).toLowerCase()) >= 0;
      case 'notcontains': return String(v || '').toLowerCase().indexOf(String(t).toLowerCase()) < 0;
      case 'above': case 'below': return true;
    }
    return true;
  };
}
export function filterButtonHit(doc, hit, p) {
  const s = doc.wb.cur; const f = s.filter; if (!f || hit.r !== f.r1 || hit.c < f.c1 || hit.c > f.c2) return false;
  const rc = doc.grid.cellRect(hit.r, hit.c); const z = doc.grid.z;
  if (p.x < rc.x + rc.w - 18 * z || p.y < rc.y) return false;
  filterMenu(doc, hit.c, rc);
  return true;
}
function filterMenu(doc, c, rc) {
  const s = doc.wb.cur; const f = s.filter;
  // unique values (as displayed)
  const vals = new Map();
  for (let r = f.r1 + 1; r <= Math.max(f.r2, s.used().r2); r++) {
    if (s.filterHidden && s.filterHidden[r] && !(f.crit[c] && Object.keys(f.crit).length === 1)) { if (Object.keys(f.crit).some(k => +k !== c)) continue; }
    const v = s.value(r, c); const t = v == null ? '' : format(v, (s.styleAt(r, c) || {}).nf).t;
    if (!vals.has(t)) vals.set(t, v);
  }
  const keys = Array.from(vals.keys()).sort((a, b) => cmp(vals.get(a), vals.get(b)));
  const cur = f.crit[c] && f.crit[c].values;
  const box = h('div.flt');
  const search = h('input.inp', { placeholder: 'Tìm kiếm' });
  const list = h('div.flt-list');
  const checks = [];
  const all = h('input', { type: 'checkbox' }); all.checked = !cur;
  list.appendChild(h('label.chk', all, h('span', '(Chọn tất cả)')));
  keys.forEach(k => { const cb = h('input', { type: 'checkbox' }); cb.checked = !cur || cur.indexOf(k) >= 0; checks.push([k, cb]); list.appendChild(h('label.chk.flt-item', cb, h('span', k === '' ? '(Trống)' : k))); });
  all.addEventListener('change', () => checks.forEach(([, cb]) => { if (cb.parentNode.style.display !== 'none') cb.checked = all.checked; }));
  search.addEventListener('input', () => { const q2 = search.value.toLowerCase(); checks.forEach(([k, cb]) => { cb.parentNode.style.display = k.toLowerCase().indexOf(q2) >= 0 ? '' : 'none'; }); });
  const apply = (crit) => { doc.change('Lọc', s, 'props', () => { f.crit[c] = crit; applyFilter(doc); }); import('../core/ui.js').then(m => m.closeMenus()); };
  const sortRg = { r1: f.r1, c1: f.c1, r2: Math.max(f.r2, s.used().r2), c2: f.c2 };
  box.appendChild(h('button.menu-item', { onclick: () => { sortRange(doc, [{ c }], sortRg, true); import('../core/ui.js').then(m => m.closeMenus()); } }, h('span.mi-ic', ic('arrow-down-a-z', 15)), h('span.mi-label', 'Sắp xếp A → Z')));
  box.appendChild(h('button.menu-item', { onclick: () => { sortRange(doc, [{ c, desc: true }], sortRg, true); import('../core/ui.js').then(m => m.closeMenus()); } }, h('span.mi-ic', ic('arrow-down-z-a', 15)), h('span.mi-label', 'Sắp xếp Z → A')));
  box.appendChild(h('div.menu-sep'));
  box.appendChild(h('button.menu-item' + (f.crit[c] ? '' : '.disabled'), { onclick: () => apply(null) }, h('span.mi-ic', ic('filter-x', 15)), h('span.mi-label', 'Xóa bộ lọc khỏi cột ' + colName(c))));
  const isNum = keys.some(k => typeof vals.get(k) === 'number');
  box.appendChild(h('button.menu-item', { onclick: () => customFilterDialog(doc, c, isNum) }, h('span.mi-ic', ic('list-filter', 15)), h('span.mi-label', (isNum ? 'Bộ lọc số' : 'Bộ lọc văn bản') + ' (tùy chỉnh)…')));
  box.appendChild(search); box.appendChild(list);
  box.appendChild(h('div.flt-foot', h('button.btn.primary', { onclick: () => { const sel2 = checks.filter(([, cb]) => cb.checked).map(([k]) => k); apply(sel2.length === checks.length ? null : { values: sel2 }); } }, 'OK'), h('button.btn', { onclick: () => import('../core/ui.js').then(m => m.closeMenus()) }, 'Hủy')));
  import('../core/ui.js').then(m => m.showPopup(box, { x: rc.x + doc.grid.el.getBoundingClientRect().left, y: rc.y + rc.h + doc.grid.el.getBoundingClientRect().top, cls: 'flt-pop' }));
  setTimeout(() => search.focus(), 50);
}
async function customFilterDialog(doc, c, isNum) {
  const ops = [['eq', 'bằng'], ['ne', 'không bằng'], ['gt', 'lớn hơn'], ['ge', 'lớn hơn hoặc bằng'], ['lt', 'nhỏ hơn'], ['le', 'nhỏ hơn hoặc bằng'], ['begins', 'bắt đầu bằng'], ['ends', 'kết thúc bằng'], ['contains', 'chứa'], ['notcontains', 'không chứa']];
  const v = await formDlg('Bộ lọc tùy chỉnh — cột ' + colName(c), [{ key: 'op1', type: 'select', label: 'Hiện các hàng có giá trị', options: ops, value: isNum ? 'gt' : 'contains' }, { key: 'v1', label: 'Giá trị' }, { key: 'and', type: 'radio', value: 'and', options: [['and', 'Và (And)'], ['or', 'Hoặc (Or)']] }, { key: 'op2', type: 'select', label: '', options: [['', '(không)']].concat(ops), value: '' }, { key: 'v2', label: 'Giá trị' }]);
  if (!v) return;
  const cond = [{ op: v.op1, v: v.v1 }]; if (v.op2) cond.push({ op: v.op2, v: v.v2 });
  const s = doc.wb.cur; doc.change('Lọc tùy chỉnh', s, 'props', () => { s.filter.crit[c] = { cond, and: v.and === 'and' }; applyFilter(doc); });
}
export function clearFilters(doc) { const s = doc.wb.cur; if (!s.filter) return; doc.change('Xóa bộ lọc', s, 'props', () => { s.filter.crit = {}; applyFilter(doc); }); }
export function reapplyFilter(doc) { const s = doc.wb.cur; if (!s.filter) return; doc.change('Áp dụng lại bộ lọc', s, 'props', () => applyFilter(doc)); }

// ---------------------------------------------------------------- find / replace / go to
export function findDialog(doc, replaceMode) {
  const s0 = doc.wb.cur;
  const fi = h('input.inp', { value: doc._lastFind || '' }); const ri = h('input.inp');
  const mc = h('input', { type: 'checkbox' }), whole = h('input', { type: 'checkbox' }), inF = h('select.inp.small', h('option', { value: 'v' }, 'Giá trị'), h('option', { value: 'f' }, 'Công thức'), h('option', { value: 'n' }, 'Chú thích'));
  const scope = h('select.inp.small', h('option', { value: 's' }, 'Trang tính'), h('option', { value: 'w' }, 'Sổ làm việc'));
  const results = h('div.find-results');
  const body = h('div.form', h('label.lbl', 'Tìm gì:'), fi, replaceMode ? h('div', h('label.lbl', 'Thay bằng:'), ri) : null,
    h('div.form-row.inline', h('label.lbl', 'Trong:'), scope), h('div.form-row.inline', h('label.lbl', 'Tìm trong:'), inF),
    h('label.chk', mc, h('span', 'Khớp chữ hoa/thường')), h('label.chk', whole, h('span', 'Khớp toàn bộ nội dung ô')), results);
  const matcher = () => { const t = fi.value; if (!t) return null; const re = new RegExp((whole.checked ? '^' : '') + t.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + (whole.checked ? '$' : ''), mc.checked ? '' : 'i'); return re; };
  const textOf = (s, x, r, c) => inF.value === 'f' ? (x.f ? '=' + x.f : x.v == null ? '' : String(x.v)) : inF.value === 'n' ? (x.note ? x.note.text || '' : '') : (x.v == null ? '' : format(x.v, x.s && x.s.nf).t);
  const all = () => { const re = matcher(); if (!re) return []; const out = []; (scope.value === 'w' ? doc.wb.sheets : [doc.wb.cur]).forEach(s => { const ks = Array.from(s.cells.keys()).sort((a, b) => (Math.floor(a / 20000) - Math.floor(b / 20000)) || (a % 20000 - b % 20000)); ks.forEach(k => { const x = s.cells.get(k); const r = Math.floor(k / 20000), c = k % 20000; if (re.test(textOf(s, x, r, c))) out.push({ s, r, c }); }); }); return out; };
  const go = (m) => { if (m.s !== doc.wb.cur) doc.activateSheet(m.s); doc.grid.select(m.r, m.c); doc.grid.scrollTo(m.r, m.c); };
  const findNext = () => { doc._lastFind = fi.value; const list = all(); if (!list.length) { toast('Không tìm thấy "' + fi.value + '"', 'err'); return; } const s = doc.wb.cur, cr = s.sel.r, cc = s.sel.c; const n = list.find(m => m.s === s && (m.r > cr || (m.r === cr && m.c > cc))) || list.find(m => doc.wb.sheets.indexOf(m.s) > doc.wb.active) || list[0]; go(n); };
  const findAll = () => { const list = all(); results.innerHTML = ''; results.appendChild(h('div.muted', list.length + ' ô được tìm thấy')); list.slice(0, 300).forEach(m => { const x = m.s.get(m.r, m.c); const row = h('div.find-row', h('span', m.s.name), h('span', addr(m.r, m.c)), h('span', textOf(m.s, x).slice(0, 60))); row.addEventListener('click', () => go(m)); results.appendChild(row); }); };
  const replaceOne = () => { const re = matcher(); if (!re) return; const s = doc.wb.cur; const x = s.get(s.sel.r, s.sel.c); if (x && re.test(textOf(s, x))) { const nt = (x.f ? '=' + x.f : String(x.v == null ? '' : x.v)).replace(new RegExp(re.source.replace(/^\^|\$$/g, ''), re.flags + 'g'), ri.value); doc.setInput(s.sel.r, s.sel.c, nt); } findNext(); };
  const replaceAll = () => { const re = matcher(); if (!re) return; const list = all(); if (!list.length) return toast('Không tìm thấy', 'err'); doc.change('Thay thế tất cả', doc.wb.cur, 'book', () => list.forEach(m => { const x = m.s.get(m.r, m.c); const src = x.f ? '=' + x.f : String(x.v == null ? '' : x.v); const nt = src.replace(new RegExp(re.source.replace(/^\^|\$$/g, ''), re.flags + 'g'), ri.value); if (/^=/.test(nt)) m.s.setCell(m.r, m.c, { f: nt.slice(1), v: null }); else m.s.setCell(m.r, m.c, { v: parseInput(nt).v, f: null }); })); toast('Đã thay thế ' + list.length + ' ô', 'ok'); };
  const btns = [{ label: 'Tìm tiếp', primary: true, run: () => { findNext(); return false; } }, { label: 'Tìm tất cả', run: () => { findAll(); return false; } }];
  if (replaceMode) btns.push({ label: 'Thay thế', run: () => { replaceOne(); return false; } }, { label: 'Thay thế tất cả', run: () => { replaceAll(); return false; } });
  btns.push({ label: 'Đóng', value: null });
  dialog({ title: replaceMode ? 'Tìm và Thay thế' : 'Tìm', width: 520, body, buttons: btns });
}
export async function gotoDialog(doc) {
  const names = Object.keys(doc.wb.names);
  const v = await formDlg('Đi tới (Go To)', [{ key: 'ref', label: 'Tham chiếu (ví dụ: B20, Sheet2!A1:C5, hoặc tên):', value: '' }, { type: 'html', el: names.length ? h('div.hint', 'Tên đã đặt: ' + names.join(', ')) : h('div') }, { type: 'html', el: h('button.btn', { onclick: () => { import('../core/ui.js').then(m => { document.querySelector('.dlg-back .dlg-x').click(); gotoSpecialDialog(doc); }); } }, 'Đặc biệt (Special)…') }]);
  if (v && v.ref) gotoRef(doc, v.ref);
}
export function gotoRef(doc, ref) {
  ref = String(ref || '').trim(); if (!ref) return;
  const nm = doc.wb.names[ref.toUpperCase()] || doc.wb.names[ref];
  if (nm) ref = String(nm.f || nm).replace(/^=/, '');
  let s = doc.wb.cur; const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(ref);
  if (m) { s = doc.wb.sheet((m[1] || m[2]).replace(/''/g, "'")); ref = m[3]; if (!s) return toast('Không tìm thấy trang tính', 'err'); }
  const rg = parseRangeAddr(ref); if (!rg) return toast('Tham chiếu không hợp lệ: ' + ref, 'err');
  if (s !== doc.wb.cur) doc.activateSheet(s);
  s.sel.ranges = [rg]; s.sel.r = rg.r1; s.sel.c = rg.c1;
  doc.grid.scrollTo(rg.r1, rg.c1); doc.onSelection(); doc.grid.redraw();
}
export async function gotoSpecialDialog(doc) {
  const v = await formDlg('Đi tới Đặc biệt', [{ key: 'w', type: 'radio', value: 'blanks', options: [['comments', 'Chú thích'], ['constants', 'Hằng số'], ['formulas', 'Công thức'], ['errors', 'Lỗi'], ['blanks', 'Ô trống'], ['region', 'Vùng hiện tại'], ['last', 'Ô cuối cùng'], ['visible', 'Chỉ ô hiển thị'], ['cf', 'Định dạng có điều kiện'], ['dv', 'Xác thực dữ liệu'], ['objects', 'Đối tượng']] }], { width: 360 });
  if (!v) return; selectSpecial(doc, v.w);
}
export function selectSpecial(doc, w) {
  const s = doc.wb.cur; const base = doc.grid.selRange(); const single = base.r1 === base.r2 && base.c1 === base.c2;
  const area = single ? s.used() : clipRange(s, base);
  if (w === 'region') { const rg = currentRegion(doc, s.sel.r, s.sel.c); s.sel.ranges = [rg]; doc.onSelection(); doc.grid.redraw(); return; }
  if (w === 'last') { const u = s.used(); doc.grid.select(u.r2, u.c2); doc.grid.scrollTo(u.r2, u.c2); return; }
  if (w === 'objects') { doc.objects.selectAll(); return; }
  const cells = [];
  for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) {
    const x = s.get(r, c);
    const ok = w === 'comments' ? x && x.note : w === 'constants' ? x && !x.f && x.v != null : w === 'formulas' ? x && x.f : w === 'errors' ? x && x.v instanceof Err : w === 'blanks' ? !x || x.v == null : w === 'visible' ? !s.hiddenR[r] && !s.hiddenC[c] : w === 'cf' ? s.cf.some(cf => inRange(cf.range, r, c)) : w === 'dv' ? s.dv.some(d => inRange(d.range, r, c)) : false;
    if (ok) cells.push({ r1: r, c1: c, r2: r, c2: c });
  }
  if (!cells.length) return toast('Không tìm thấy ô nào phù hợp', 'err');
  // merge horizontally adjacent
  const merged = []; cells.forEach(cl => { const last = merged[merged.length - 1]; if (last && last.r1 === cl.r1 && last.c2 === cl.c1 - 1) last.c2 = cl.c2; else merged.push(Object.assign({}, cl)); });
  s.sel.ranges = merged.slice(0, 5000); s.sel.r = merged[0].r1; s.sel.c = merged[0].c1; doc.onSelection(); doc.grid.redraw();
  toast('Đã chọn ' + cells.length + ' ô');
}

// ---------------------------------------------------------------- hyperlinks
export async function linkDialog(doc) {
  const s = doc.wb.cur; const cell = s.get(s.sel.r, s.sel.c) || {};
  const v = await formDlg('Chèn liên kết', [{ key: 'text', label: 'Văn bản hiển thị:', value: cell.v != null ? String(cell.v) : '' }, { key: 'url', label: 'Địa chỉ (URL, email, hoặc vị trí trong sổ như Sheet2!A1):', value: cell.link ? cell.link.url || cell.link.loc || '' : 'https://' }, { key: 'tip', label: 'Mẹo màn hình (ScreenTip):', value: cell.link && cell.link.tip || '' }]);
  if (!v) return;
  doc.change('Chèn liên kết', s, { r1: s.sel.r, c1: s.sel.c, r2: s.sel.r, c2: s.sel.c }, () => {
    const isLoc = /^[^:/]+![A-Z]+\d+/i.test(v.url) || /^#/.test(v.url);
    s.setCell(s.sel.r, s.sel.c, { v: v.text || v.url, f: null, link: isLoc ? { loc: v.url.replace(/^#/, ''), tip: v.tip } : { url: /^[\w.+-]+@[\w-]+\.[\w.]+$/.test(v.url) ? 'mailto:' + v.url : v.url, tip: v.tip }, s: mergeStyle(cell.s, { color: '#0563C1', u: 'single' }) });
  });
}
export function followLink(doc, link) { if (link.loc) gotoRef(doc, link.loc); else if (link.url) Host.openUrl(link.url); }

// ---------------------------------------------------------------- notes & comments
export async function noteDialog(doc) {
  const s = doc.wb.cur; const cell = s.get(s.sel.r, s.sel.c) || {};
  const v = await promptDlg('Ghi chú cho ô ' + addr(s.sel.r, s.sel.c) + ':', cell.note ? cell.note.text : '', cell.note ? 'Sửa ghi chú' : 'Ghi chú mới', { multiline: true, rows: 5 });
  if (v == null) return;
  doc.change('Ghi chú', s, { r1: s.sel.r, c1: s.sel.c, r2: s.sel.r, c2: s.sel.c }, () => s.setCell(s.sel.r, s.sel.c, { note: v ? { text: v, author: settings().userName } : null }));
}
export async function commentDialog(doc) {
  const s = doc.wb.cur; const cell = s.get(s.sel.r, s.sel.c) || {};
  const v = await promptDlg('Bắt đầu cuộc hội thoại (@đề cập ai đó):', '', 'Chú thích mới', { multiline: true, rows: 4 });
  if (!v) return;
  const thread = (cell.note && cell.note.thread ? cell.note.thread.slice() : []).concat([{ author: settings().userName, text: v, time: new Date().toLocaleString() }]);
  doc.change('Chú thích', s, { r1: s.sel.r, c1: s.sel.c, r2: s.sel.r, c2: s.sel.c }, () => s.setCell(s.sel.r, s.sel.c, { note: { thread } }));
  commentsPane(doc);
}
export function commentsPane(doc) {
  const old = $('.xl-cpane', doc.el); if (old) { old.parentNode.removeChild(old); return; }
  const pane = h('div.pane.xl-cpane', h('div.pane-head', h('span.t', 'Chú thích & Ghi chú'), h('button.pane-x', { onclick: () => pane.parentNode.removeChild(pane) }, ic('x', 16))));
  const body = h('div.pane-body'); pane.appendChild(body);
  const draw = () => {
    body.innerHTML = '';
    body.appendChild(h('button.btn.primary', { onclick: () => commentDialog(doc).then(draw) }, ic('message-square-plus', 14), ' Chú thích mới'));
    doc.wb.sheets.forEach(s => s.cells.forEach((x, k) => { if (!x.note) return; const r = Math.floor(k / 20000), c = k % 20000;
      const card = h('div.cm-card', h('div.cm-head', h('b', s.name + '!' + addr(r, c))), x.note.thread ? x.note.thread.map(m => h('div.cm-msg', h('b', m.author), h('span.muted', ' ' + m.time), h('div', m.text))) : h('div.cm-msg', h('b', x.note.author || 'Ghi chú'), h('div', x.note.text)),
        h('div.cm-act', h('button.btn', { onclick: () => { if (s !== doc.wb.cur) doc.activateSheet(s); doc.grid.select(r, c); doc.grid.scrollTo(r, c); } }, 'Đi tới'), h('button.btn', { onclick: () => { doc.change('Xóa chú thích', s, { r1: r, c1: c, r2: r, c2: c }, () => s.setCell(r, c, { note: null })); draw(); } }, 'Xóa')));
      body.appendChild(card); }));
  };
  draw();
  doc.el.appendChild(pane);
}
export function nextComment(doc, dir) {
  const list = []; doc.wb.sheets.forEach((s, si) => s.cells.forEach((x, k) => { if (x.note) list.push({ s, si, r: Math.floor(k / 20000), c: k % 20000 }); }));
  if (!list.length) return toast('Không có chú thích nào');
  list.sort((a, b) => a.si - b.si || a.r - b.r || a.c - b.c);
  const s = doc.wb.cur, si = doc.wb.active; const curIdx = list.findIndex(m => m.si === si && m.r === s.sel.r && m.c === s.sel.c);
  const n = list[(curIdx + (dir || 1) + list.length) % list.length] || list[0];
  if (n.s !== doc.wb.cur) doc.activateSheet(n.s); doc.grid.select(n.r, n.c); doc.grid.scrollTo(n.r, n.c); doc.showNotePopup(n.r, n.c);
}

// ---------------------------------------------------------------- context menus
export function contextItems(doc, zone) {
  const s = doc.wb.cur; const cell = s.get(s.sel.r, s.sel.c);
  const base = [
    { label: 'Cắt', icon: 'scissors', shortcut: 'Ctrl+X', run: () => copy(doc, true) },
    { label: 'Sao chép', icon: 'copy', shortcut: 'Ctrl+C', run: () => copy(doc, false) },
    { label: 'Tùy chọn dán', icon: 'clipboard-paste', sub: () => pasteMenu(doc) },
    { sep: true }
  ];
  if (zone === 'col') return base.concat([
    { label: 'Chèn cột', icon: 'between-vertical-start', run: () => insertSelCols(doc) },
    { label: 'Xóa cột', icon: 'trash-2', run: () => deleteSelCols(doc) },
    { label: 'Xóa nội dung', icon: 'eraser', run: () => clear(doc, 'contents') },
    { sep: true },
    { label: 'Định dạng ô…', icon: 'paintbrush', run: () => formatCellsDialog(doc) },
    { label: 'Độ rộng cột…', icon: 'move-horizontal', run: () => colWidthDialog(doc) },
    { label: 'Tự động khớp độ rộng', icon: 'arrow-left-right', run: () => autoFit(doc, 'c') },
    { label: 'Ẩn', icon: 'eye-off', run: () => hideCols(doc, true) }, { label: 'Bỏ ẩn', icon: 'eye', run: () => hideCols(doc, false) }
  ]);
  if (zone === 'row') return base.concat([
    { label: 'Chèn hàng', icon: 'between-horizontal-start', run: () => insertSelRows(doc) },
    { label: 'Xóa hàng', icon: 'trash-2', run: () => deleteSelRows(doc) },
    { label: 'Xóa nội dung', icon: 'eraser', run: () => clear(doc, 'contents') },
    { sep: true },
    { label: 'Định dạng ô…', icon: 'paintbrush', run: () => formatCellsDialog(doc) },
    { label: 'Chiều cao hàng…', icon: 'move-vertical', run: () => rowHeightDialog(doc) },
    { label: 'Tự động khớp chiều cao', icon: 'arrow-up-down', run: () => autoFit(doc, 'r') },
    { label: 'Ẩn', icon: 'eye-off', run: () => hideRows(doc, true) }, { label: 'Bỏ ẩn', icon: 'eye', run: () => hideRows(doc, false) }
  ]);
  const dvList = doc.dvListFor ? doc.dvListFor(s.sel.r, s.sel.c) : null;
  return base.concat([
    { label: 'Tra cứu thông minh / Hỏi A.I', icon: 'sparkles', run: () => import('../ai/ai.js').then(m => m.openAi(doc.shell, doc, 'Giải thích dữ liệu trong vùng ' + rangeAddr(doc.grid.selRange()))) },
    { sep: true },
    { label: 'Chèn…', icon: 'plus', run: () => insertDialog(doc) },
    { label: 'Xóa…', icon: 'minus', run: () => deleteDialog(doc) },
    { label: 'Xóa nội dung', icon: 'eraser', shortcut: 'Delete', run: () => clear(doc, 'contents') },
    { label: 'Phân tích nhanh (Quick Analysis)', icon: 'zap', run: () => quickAnalysis(doc) },
    { sep: true },
    { label: 'Lọc', icon: 'filter', sub: [{ label: 'Lọc theo giá trị của ô đã chọn', run: () => filterBySelected(doc, 'value') }, { label: 'Lọc theo màu nền của ô', run: () => filterBySelected(doc, 'color') }, { label: 'Áp dụng lại', run: () => reapplyFilter(doc) }] },
    { label: 'Sắp xếp', icon: 'arrow-up-down', sub: [{ label: 'Sắp xếp A → Z', icon: 'arrow-down-a-z', run: () => quickSort(doc, false) }, { label: 'Sắp xếp Z → A', icon: 'arrow-down-z-a', run: () => quickSort(doc, true) }, { label: 'Đưa màu ô đã chọn lên đầu', run: () => { const st = s.styleAt(s.sel.r, s.sel.c); sortRange(doc, [{ c: s.sel.c, by: 'color', color: st && st.fill }]); } }, { label: 'Sắp xếp tùy chỉnh…', run: () => sortDialog(doc) }] },
    { label: 'Lấy dữ liệu từ Bảng/Phạm vi', icon: 'table', run: () => formatAsTable(doc) },
    { sep: true },
    { label: 'Chú thích mới', icon: 'message-square-plus', run: () => commentDialog(doc) },
    { label: cell && cell.note ? 'Sửa ghi chú' : 'Ghi chú mới', icon: 'sticky-note', run: () => noteDialog(doc) },
    cell && cell.note ? { label: 'Xóa chú thích/ghi chú', icon: 'message-square-x', run: () => clear(doc, 'comments') } : null,
    { sep: true },
    { label: 'Định dạng ô…', icon: 'paintbrush', shortcut: 'Ctrl+1', run: () => formatCellsDialog(doc) },
    { label: 'Chọn từ danh sách thả xuống…', icon: 'list', run: () => pickFromList(doc) },
    { label: 'Xác định tên…', icon: 'tag', run: () => defineNameDialog(doc) },
    { label: cell && cell.link ? 'Sửa liên kết…' : 'Liên kết', icon: 'link', shortcut: 'Ctrl+K', run: () => linkDialog(doc) },
    cell && cell.link ? { label: 'Mở liên kết', icon: 'external-link', run: () => followLink(doc, cell.link) } : null,
    cell && cell.link ? { label: 'Loại bỏ liên kết', icon: 'unlink', run: () => clear(doc, 'links') } : null,
    { sep: true },
    { label: 'Vùng in & Ngắt trang', icon: 'printer', sub: [
      { label: 'Đặt vùng in (vùng chọn hiện tại)', icon: 'crop', run: () => import('./pages.js').then(m => m.setPrintArea(doc)) },
      { label: 'Thêm vào vùng in', icon: 'plus-square', run: () => import('./pages.js').then(m => m.addToPrintArea(doc)) },
      { label: 'Xóa vùng in', icon: 'trash-2', run: () => import('./pages.js').then(m => m.clearPrintArea(doc)) },
      { sep: true },
      { label: 'Chèn ngắt trang', icon: 'separator-horizontal', run: () => import('./pages.js').then(m => m.insertBreak(doc)) },
      { label: 'Loại bỏ ngắt trang', icon: 'minus', run: () => import('./pages.js').then(m => m.removeBreak(doc)) },
      { label: 'Đặt lại tất cả ngắt trang', icon: 'rotate-ccw', run: () => import('./pages.js').then(m => m.resetBreaks(doc)) },
      { sep: true },
      { label: 'Xem trước Ngắt trang (Bật/Tắt)', icon: 'square-dashed', run: () => { doc.wb.cur.showPageBreaks = !doc.wb.cur.showPageBreaks; doc.grid.redraw(); if (doc.refreshUi) doc.refreshUi(); } }
    ] }
  ].filter(Boolean));
}
function filterBySelected(doc, by) {
  const s = doc.wb.cur; if (!s.filter) toggleFilter(doc); const f = s.filter; if (!f) return;
  const c = s.sel.c, r = s.sel.r; const st = s.styleAt(r, c);
  doc.change('Lọc', s, 'props', () => { f.crit[c] = by === 'color' ? { color: st && st.fill } : { values: [format(s.value(r, c), (st || {}).nf).t] }; applyFilter(doc); });
}
function pickFromList(doc) {
  const s = doc.wb.cur; const c = s.sel.c; const vals = new Set();
  for (let r = Math.max(0, s.sel.r - 500); r <= s.sel.r + 500; r++) { if (r === s.sel.r) continue; const v = s.value(r, c); if (typeof v === 'string' && v) vals.add(v); }
  const list = Array.from(vals).sort();
  if (!list.length) return toast('Không có giá trị văn bản nào trong cột');
  const rc = doc.grid.cellRect(s.sel.r, c); const er = doc.grid.el.getBoundingClientRect();
  showMenu(list.slice(0, 200).map(v => ({ label: v, run: () => doc.setInput(s.sel.r, c, v) })), { x: er.left + rc.x, y: er.top + rc.y + rc.h, minWidth: Math.max(160, rc.w) });
}
export function quickAnalysis(doc) {
  const s = doc.wb.cur; const rg = clipRange(s, doc.grid.selRange());
  showMenu([
    { header: 'Định dạng' }, { label: 'Thanh dữ liệu', icon: 'bar-chart-horizontal', run: () => import('./features.js').then(m => m.addCfRule(doc, { type: 'databar', color: '#638ec6' }, rg)) }, { label: 'Thang màu', icon: 'palette', run: () => import('./features.js').then(m => m.addCfRule(doc, { type: 'scale3' }, rg)) }, { label: 'Lớn hơn trung bình', icon: 'trending-up', run: () => import('./features.js').then(m => m.addCfRule(doc, { type: 'aboveavg', fill: '#ffc7ce', color: '#9c0006' }, rg)) },
    { header: 'Biểu đồ' }, { label: 'Cột', icon: 'chart-column', run: () => import('./features.js').then(m => m.insertChart(doc, 'column')) }, { label: 'Đường', icon: 'chart-line', run: () => import('./features.js').then(m => m.insertChart(doc, 'line')) }, { label: 'Tròn', icon: 'chart-pie', run: () => import('./features.js').then(m => m.insertChart(doc, 'pie')) },
    { header: 'Tổng' }, { label: 'Tổng cuối mỗi cột', icon: 'sigma', run: () => autoSum(doc, 'SUM') }, { label: 'Trung bình', icon: 'divide', run: () => autoSum(doc, 'AVERAGE') }, { label: 'Đếm', icon: 'hash', run: () => autoSum(doc, 'COUNT') },
    { header: 'Bảng' }, { label: 'Định dạng thành Bảng', icon: 'table', run: () => import('./features.js').then(m => m.formatAsTable(doc)) }, { label: 'PivotTable', icon: 'table-2', run: () => import('./features.js').then(m => m.pivotDialog(doc)) }
  ], { x: doc.grid.el.getBoundingClientRect().left + doc.grid.cellRect(rg.r2, rg.c2).x + 20, y: doc.grid.el.getBoundingClientRect().top + doc.grid.cellRect(rg.r2, rg.c2).y + 20 });
}

// ---------------------------------------------------------------- names
export async function defineNameDialog(doc, existing) {
  const s = doc.wb.cur; const rg = doc.grid.selRange();
  const sugg = (() => { const v = s.value(rg.r1 > 0 ? rg.r1 - 1 : rg.r1, rg.c1); return typeof v === 'string' ? v.replace(/[^\wÀ-ỹ]+/g, '_').replace(/^(\d)/, '_$1') : ''; })();
  const v = await formDlg(existing ? 'Sửa tên' : 'Tên mới', [{ key: 'name', label: 'Tên:', value: existing || sugg }, { key: 'ref', label: 'Tham chiếu tới:', value: existing ? String(doc.wb.names[existing].f || doc.wb.names[existing]) : '=' + q(s.name) + '!' + rangeAddr({ r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2') }, { key: 'cmt', label: 'Chú thích:', value: '' }],
    { validate: o => !/^[A-Za-z_\\À-ỹ][\w.À-ỹ]*$/.test(o.name) || parseAddr(o.name) ? 'Tên không hợp lệ (bắt đầu bằng chữ, không trùng địa chỉ ô)' : null });
  if (!v) return;
  doc.change('Đặt tên', s, 'book', () => { if (existing && existing !== v.name.toUpperCase()) delete doc.wb.names[existing]; doc.wb.names[v.name.toUpperCase()] = { f: v.ref.replace(/^=/, ''), comment: v.cmt, display: v.name }; });
}
export function nameManager(doc) {
  const body = h('div');
  const draw = () => {
    body.innerHTML = '';
    const tbl = h('table.kv.nm-table', h('tr', h('th', 'Tên'), h('th', 'Giá trị'), h('th', 'Tham chiếu tới'), h('th', '')));
    Object.keys(doc.wb.names).sort().forEach(n => {
      const d = doc.wb.names[n]; const f = String(d.f || d); let val = doc.calc.evalFormula(f, doc.wb.cur, 0, 0); if (Array.isArray(val)) val = '{' + val.slice(0, 3).map(r => r.slice(0, 4).join(',')).join(';') + '…}'; else if (val instanceof Err) val = val.code;
      tbl.appendChild(h('tr', h('td', d.display || n), h('td', String(val == null ? '' : val).slice(0, 40)), h('td', '=' + f), h('td', h('button.btn', { onclick: () => defineNameDialog(doc, n).then(draw) }, 'Sửa'), h('button.btn', { onclick: () => { doc.change('Xóa tên', doc.wb.cur, 'book', () => { delete doc.wb.names[n]; }); draw(); } }, 'Xóa'))));
    });
    body.appendChild(tbl);
    if (!Object.keys(doc.wb.names).length) body.appendChild(h('div.muted', 'Chưa có tên nào.'));
  };
  draw();
  dialog({ title: 'Trình quản lý Tên (Name Manager)', width: 720, body, buttons: [{ label: 'Mới…', run: () => { defineNameDialog(doc).then(draw); return false; } }, { label: 'Đóng', primary: true }] });
}
export function createNamesFromSelection(doc) {
  const s = doc.wb.cur; const rg = clipRange(s, doc.grid.selRange());
  doc.change('Tạo tên từ vùng chọn', s, 'book', () => { let n = 0; for (let c = rg.c1; c <= rg.c2; c++) { const t = s.value(rg.r1, c); if (typeof t !== 'string' || !t) continue; const nm = t.replace(/[^\wÀ-ỹ]+/g, '_').replace(/^(\d)/, '_$1'); doc.wb.names[nm.toUpperCase()] = { f: q(s.name) + '!$' + colName(c) + '$' + (rg.r1 + 2) + ':$' + colName(c) + '$' + (rg.r2 + 1), display: nm }; n++; } toast('Đã tạo ' + n + ' tên từ hàng đầu tiên', 'ok'); });
}

// ---------------------------------------------------------------- tables auto-extend
export function extendTable(doc, s, r, c) {
  s.tables.forEach(t => { if (r === t.r2 + 1 && c >= t.c1 && c <= t.c2) t.r2 = r; if (c === t.c2 + 1 && r >= t.r1 && r <= t.r2 && r === t.r1) { t.c2 = c; } });
}

// ---------------------------------------------------------------- page breaks (v0.0.8: real page model in pages.js)
export function pageBreaks(doc, s) { return computePages(s); }

// validation hook used on input
export function validate(doc, s, r, c, text) {
  const dv = s.dv.find(d => inRange(d.range, r, c)); if (!dv || text === '' || text == null) return true;
  const ok = dvCheck(doc, dv, text, s, r, c);
  if (!ok) { const st = dv.errStyle || 'stop'; import('../core/ui.js').then(m => m.alertDlg(dv.errMsg || 'Giá trị bạn nhập không hợp lệ.\nNgười dùng đã giới hạn các giá trị có thể nhập vào ô này.', dv.errTitle || 'Yomi-Office Open Source')); if (st === 'stop') return false; }
  return true;
}
export function cfStyle(doc, s, r, c, v) { return cfStyleImpl(doc, s, r, c, v); }
export function extraCommands(doc) {
  return [
    { label: 'Đi tới… (Ctrl+G)', icon: 'navigation', run: () => gotoDialog(doc) },
    { label: 'Tìm và thay thế (Ctrl+H)', icon: 'replace', run: () => findDialog(doc, true) },
    { label: 'Hiện công thức (Ctrl+`)', icon: 'function-square', run: () => toggleShowFormulas(doc) },
    { label: 'Điền nhanh (Ctrl+E)', icon: 'zap', run: () => flashFill(doc) },
    { label: 'Phân tích nhanh', icon: 'zap', run: () => quickAnalysis(doc) }
  ];
}
