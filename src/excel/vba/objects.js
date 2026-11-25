// Yomi-Office Open Source v0.0.6 — Excel object model + VBA built-in functions for the VBA engine (./interp.js)
import { VbaError, VbArray, Collection, Dictionary, NOT_FOUND, toNum, toStr, toDate, truthy, cmp, bankRound } from './interp.js';
import { parseRangeAddr, rangeAddr, addr, colName, colIndex, key, Err, MAX_ROWS, MAX_COLS, mergeStyle, norm, jsDateToSerial, serialToJsDate } from '../model.js';
import { parseInput, format } from '../numfmt.js';
import { FN } from '../formula.js';

// ---- colours: VBA uses BGR longs (RGB(r,g,b) = r + g*256 + b*65536)
export function rgbLong(r, g, b) { return (r & 255) + (g & 255) * 256 + (b & 255) * 65536; }
function longToHex(n) { n = toNum(n); const r = n & 255, g = (n >> 8) & 255, b = (n >> 16) & 255; return '#' + [r, g, b].map(x => (x < 16 ? '0' : '') + x.toString(16)).join(''); }
function hexToLong(h) { if (!h || h[0] !== '#') return 16777215; const s = h.length === 4 ? h.slice(1).split('').map(c => c + c).join('') : h.slice(1, 7); const r = parseInt(s.slice(0, 2), 16), g = parseInt(s.slice(2, 4), 16), b = parseInt(s.slice(4, 6), 16); return rgbLong(r, g, b); }
const QB = [0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0, 0x808080, 0xFF0000, 0x00FF00, 0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF];
const COLORINDEX = [null, '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080', '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF', '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99', '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696', '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'];

export const CONSTS = {
  vbcrlf: '\r\n', vbnewline: '\r\n', vbcr: '\r', vblf: '\n', vbtab: '\t', vbnullstring: '', vbnullchar: '\0', vbback: '\b',
  vbokonly: 0, vbokcancel: 1, vbabortretryignore: 2, vbyesnocancel: 3, vbyesno: 4, vbretrycancel: 5, vbcritical: 16, vbquestion: 32, vbexclamation: 48, vbinformation: 64, vbdefaultbutton1: 0, vbdefaultbutton2: 256, vbsystemmodal: 4096,
  vbok: 1, vbcancel: 2, vbabort: 3, vbretry: 4, vbignore: 5, vbyes: 6, vbno: 7,
  vbbinarycompare: 0, vbtextcompare: 1, vbuppercase: 1, vblowercase: 2, vbpropercase: 3,
  vbsunday: 1, vbmonday: 2, vbtuesday: 3, vbwednesday: 4, vbthursday: 5, vbfriday: 6, vbsaturday: 7,
  vbempty: 0, vbnull: 1, vbinteger: 2, vblong: 3, vbsingle: 4, vbdouble: 5, vbcurrency: 6, vbdate: 7, vbstring: 8, vbobject: 9, vberror: 10, vbboolean: 11, vbvariant: 12, vbarray: 8192,
  vbblack: 0, vbred: 255, vbgreen: 65280, vbyellow: 65535, vbblue: 16711680, vbmagenta: 16711935, vbcyan: 16776960, vbwhite: 16777215,
  xlup: -4162, xldown: -4121, xltoleft: -4159, xltoright: -4161,
  xlleft: -4131, xlright: -4152, xlcenter: -4108, xlgeneral: 1, xljustify: -4130, xltop: -4160, xlbottom: -4107, xlcenteracrossselection: 7,
  xlascending: 1, xldescending: 2, xlyes: 1, xlno: 2, xlguess: 0,
  xlcalculationautomatic: -4105, xlcalculationmanual: -4135, xlcalculationsemiautomatic: 2,
  xlcontinuous: 1, xldash: -4115, xldot: -4118, xldouble: -4119, xlnone: -4142, xlthin: 2, xlmedium: -4138, xlthick: 4, xlhairline: 1,
  xledgeleft: 7, xledgetop: 8, xledgebottom: 9, xledgeright: 10, xlinsidevertical: 11, xlinsidehorizontal: 12, xldiagonaldown: 5, xldiagonalup: 6,
  xlpasteall: -4104, xlpastevalues: -4163, xlpasteformats: -4122, xlpasteformulas: -4123, xlshiftdown: -4121, xlshifttoright: -4161, xlshiftup: -4162, xlshifttoleft: -4159,
  xlcelltypeblanks: 4, xlcelltypeconstants: 2, xlcelltypeformulas: -4123, xlcelltypelastcell: 11, xlcelltypevisible: 12,
  xlwhole: 1, xlpart: 2, xlvalues: -4163, xlformulas: -4123, xlbyrows: 1, xlbycolumns: 2, xlnext: 1, xlprevious: 2,
  xlsheetvisible: -1, xlsheethidden: 0, xlsheetveryhidden: 2, xlautomatic: -4105, xlsolid: 1, xlthemecolordark1: 1, xlthemecolorlight1: 2,
  xlopenxmlworkbook: 51, xlopenxmlworkbookmacroenabled: 52, xlcsv: 6, xlwbatworksheet: -4167, xlworksheet: -4167, xlr1c1: -4150, xla1: 1,
  xlunderlinestylesingle: 2, xlunderlinestyledouble: -4119, xlunderlinestylenone: -4142, msoTrue: -1, msofalse: 0, msotrue: -1
};

// ---------------------------------------------------------------- objects
class VObj { }
export class VRange extends VObj {
  constructor(m, s, rg) { super(); this.m = m; this.s = s; this.rg = norm(rg); }
  get rows() { return this.rg.r2 - this.rg.r1 + 1; } get cols() { return this.rg.c2 - this.rg.c1 + 1; }
  cell(i) { const c = this.cols; return { r: this.rg.r1 + Math.floor(i / c), c: this.rg.c1 + (i % c) }; }
  each() { const out = []; const n = Math.min(this.rows * this.cols, 2000000); for (let i = 0; i < n; i++) { const p = this.cell(i); out.push(new VRange(this.m, this.s, { r1: p.r, c1: p.c, r2: p.r, c2: p.c })); } return out; }
  __value() { return this.m.getValue(this); }
}
class VRows extends VObj { constructor(rng, isCol) { super(); this.rng = rng; this.isCol = isCol; } each() { const r = this.rng, out = []; if (this.isCol) for (let c = r.rg.c1; c <= r.rg.c2; c++) out.push(new VRange(r.m, r.s, { r1: r.rg.r1, c1: c, r2: r.rg.r2, c2: c })); else for (let i = r.rg.r1; i <= r.rg.r2; i++) out.push(new VRange(r.m, r.s, { r1: i, c1: r.rg.c1, r2: i, c2: r.rg.c2 })); return out; } }
class VFont extends VObj { constructor(rng) { super(); this.rng = rng; } }
class VInterior extends VObj { constructor(rng) { super(); this.rng = rng; } }
class VBorders extends VObj { constructor(rng, which) { super(); this.rng = rng; this.which = which; } }
export { VObj, VRows, VSheets, VBook, VApp, VWsf, VFont, VInterior, VBorders };
export class VSheet extends VObj { constructor(m, s) { super(); this.m = m; this.s = s; } }
class VSheets extends VObj { constructor(m) { super(); this.m = m; } each() { return this.m.wb().sheets.map(s => new VSheet(this.m, s)); } }
class VBook extends VObj { constructor(m) { super(); this.m = m; } }
class VApp extends VObj { constructor(m) { super(); this.m = m; } }
class VWsf extends VObj { constructor(m) { super(); this.m = m; } }
class VDebug extends VObj { constructor(m) { super(); this.m = m; } }

// ---------------------------------------------------------------- the model bridge
export class ExcelModel {
  constructor(doc, ui) { this.doc = doc; this.ui = ui; this.screenUpdating = true; this.statusBar = false; this.app = new VApp(this); this.book = new VBook(this); this.dbg = new VDebug(this); this.wsf = new VWsf(this); this.dirty = new Set(); }
  wb() { return this.doc.wb; }
  sheetObj(s) { return new VSheet(this, s); }
  range(s, a, b) {
    if (a instanceof VRange && b instanceof VRange) return new VRange(this, a.s, norm({ r1: Math.min(a.rg.r1, b.rg.r1), c1: Math.min(a.rg.c1, b.rg.c1), r2: Math.max(a.rg.r2, b.rg.r2), c2: Math.max(a.rg.c2, b.rg.c2) }));
    if (a instanceof VRange) return a;
    let t = toStr(a).trim(); let sh = s;
    const bang = t.lastIndexOf('!'); if (bang > 0) { const nm = t.slice(0, bang).replace(/^'|'$/g, ''); sh = this.wb().sheet(nm); if (!sh) throw new VbaError('Không có trang tính "' + nm + '"', 0, 9); t = t.slice(bang + 1); }
    if (b != null) t = t + ':' + toStr(b instanceof VRange ? rangeAddr(b.rg) : b);
    let rg = parseRangeAddr(t.replace(/\$/g, ''));
    if (!rg) { const nm = this.wb().names && (this.wb().names[t] || this.wb().names[t.toUpperCase()]); if (nm) { const ref = String(nm.ref || nm).replace(/^=/, ''); return this.range(sh, ref); } const tbl = this.findTable(t); if (tbl) return new VRange(this, tbl.s, tbl.rg); throw new VbaError('Địa chỉ ô không hợp lệ: "' + t + '"', 0, 1004); }
    return new VRange(this, sh, rg);
  }
  findTable(name) { for (const s of this.wb().sheets) for (const t of s.tables || []) if (String(t.name).toLowerCase() === name.toLowerCase()) return { s, rg: { r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 } }; return null; }
  activeSheet() { return this.wb().cur; }
  selRange() { const s = this.activeSheet(); const r = this.doc.grid.selRange(); return new VRange(this, s, r); }
  activeCell() { const s = this.activeSheet(); return new VRange(this, s, { r1: s.sel.r, c1: s.sel.c, r2: s.sel.r, c2: s.sel.c }); }
  touch(s) { this.dirty.add(s); }
  // ---- values
  cellVal(s, r, c) { const v = s.value(r, c); if (v instanceof Err) return { __cverr: v.code }; if (v === null || v === undefined || v === '') return undefined; const cell = s.get(r, c); if (typeof v === 'number' && cell && cell.s && cell.s.nf && /[dmyhs]/i.test(cell.s.nf) && !/[#0]/.test(cell.s.nf.replace(/"[^"]*"/g, ''))) return serialToJsDate(v); return v; }
  getValue(R) { if (R.rows === 1 && R.cols === 1) return this.cellVal(R.s, R.rg.r1, R.rg.c1); const rows = []; for (let r = R.rg.r1; r <= R.rg.r2; r++) { const row = []; for (let c = R.rg.c1; c <= R.rg.c2; c++) row.push(this.cellVal(R.s, r, c)); rows.push(row); } return VbArray.from2D(rows); }
  setOne(s, r, c, v, asFormula) {
    if (v instanceof VRange) v = this.getValue(v);
    if (v && v.__cverr) { s.setCell(r, c, { v: null, f: null }); return; }
    if (v === undefined || v === null || v === '') { s.setCell(r, c, { v: null, f: null }); return; }
    if (typeof v === 'string' && v[0] === '=' && v.length > 1) { s.setCell(r, c, { f: v.slice(1), v: null }); return; }
    if (v instanceof Date) { const cell = s.get(r, c); s.setCell(r, c, { v: jsDateToSerial(v), f: null, s: (cell && cell.s && cell.s.nf) ? cell.s : mergeStyle(cell && cell.s, { nf: (v.getHours() || v.getMinutes()) ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy' }) }); return; }
    if (typeof v === 'boolean') { s.setCell(r, c, { v, f: null }); return; }
    if (typeof v === 'string' && !asFormula) { const cell = s.get(r, c); if (cell && cell.s && cell.s.nf === '@') { s.setCell(r, c, { v, f: null }); return; } const p = parseInput(v); const patch = { v: p.v, f: null }; if (p.nf && (!cell || !cell.s || !cell.s.nf)) patch.s = mergeStyle(cell && cell.s, { nf: p.nf }); s.setCell(r, c, patch); return; }
    s.setCell(r, c, { v, f: null });
  }
  setValue(R, v, asFormula) {
    this.touch(R.s);
    if (v instanceof VRange) v = this.getValue(v);
    if (v instanceof VbArray) { const m = v.to2D(); for (let r = 0; r < R.rows; r++) for (let c = 0; c < R.cols; c++) { const row = m[r] || (m.length === 1 ? m[0] : null); const x = row ? (row[c] !== undefined || row.length > c ? row[c] : (row.length === 1 ? row[0] : undefined)) : undefined; if (r >= m.length && m.length !== 1) continue; this.setOne(R.s, R.rg.r1 + r, R.rg.c1 + c, x, asFormula); } return; }
    for (let r = R.rg.r1; r <= R.rg.r2; r++) for (let c = R.rg.c1; c <= R.rg.c2; c++) {
      let val = v; if (asFormula && typeof v === 'string' && v[0] === '=' && (r !== R.rg.r1 || c !== R.rg.c1)) { const { shiftFormula } = this.fmod; val = '=' + shiftFormula(v.slice(1), r - R.rg.r1, c - R.rg.c1); }
      this.setOne(R.s, r, c, val, asFormula);
    }
  }
  style(R, patch) { this.touch(R.s); const lim = this.limitRange(R); for (let r = lim.r1; r <= lim.r2; r++) for (let c = lim.c1; c <= lim.c2; c++) { const x = R.s.get(r, c); R.s.setCell(r, c, { s: mergeStyle(x && x.s, patch) }); } if (lim.whole) { if (lim.whole === 'col') for (let c = R.rg.c1; c <= R.rg.c2; c++) R.s.colS[c] = mergeStyle(R.s.colS[c], patch); else for (let r = R.rg.r1; r <= R.rg.r2; r++) R.s.rowS[r] = mergeStyle(R.s.rowS[r], patch); } }
  limitRange(R) { const u = R.s.used(); const rg = Object.assign({}, R.rg); let whole = null; if (rg.r2 - rg.r1 > 50000) { rg.r2 = Math.max(rg.r1, Math.min(rg.r2, u.r2)); whole = 'col'; } if (rg.c2 - rg.c1 > 2000) { rg.c2 = Math.max(rg.c1, Math.min(rg.c2, u.c2)); whole = whole || 'row'; } rg.whole = whole; return rg; }
  styleOf(R) { return R.s.styleAt(R.rg.r1, R.rg.c1) || {}; }
  endOf(R, dir) {
    const s = R.s; let r = R.rg.r1, c = R.rg.c1; const has = (rr, cc) => { const v = s.value(rr, cc); return v !== null && v !== undefined && v !== ''; };
    const d = { [-4162]: [-1, 0], [-4121]: [1, 0], [-4159]: [0, -1], [-4161]: [0, 1] }[toNum(dir)] || [1, 0];
    const lim = d[0] ? (d[0] > 0 ? Math.min(MAX_ROWS - 1, s.used().r2 + 1) : 0) : (d[1] > 0 ? Math.min(MAX_COLS - 1, s.used().c2 + 1) : 0);
    const next = () => { r += d[0]; c += d[1]; };
    const inb = () => r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS;
    const u = s.used(); if (d[0] < 0 && r > u.r2 + 1) r = u.r2 + 1; if (d[1] < 0 && c > u.c2 + 1) c = u.c2 + 1;
    const curHas = has(r, c); next();
    if (!inb()) { r -= d[0]; c -= d[1]; return new VRange(this, s, { r1: r, c1: c, r2: r, c2: c }); }
    if (curHas && has(r, c)) { while (inb() && has(r, c)) next(); r -= d[0]; c -= d[1]; }
    else { while (inb() && !has(r, c) && ((d[0] > 0 && r <= lim) || (d[0] < 0 && r >= 0) || (d[1] > 0 && c <= lim) || (d[1] < 0 && c >= 0))) next(); if (!inb() || !has(r, c)) { if (d[0] > 0) r = MAX_ROWS - 1; else if (d[0] < 0) r = 0; if (d[1] > 0) c = MAX_COLS - 1; else if (d[1] < 0) c = 0; } }
    r = Math.max(0, Math.min(MAX_ROWS - 1, r)); c = Math.max(0, Math.min(MAX_COLS - 1, c));
    return new VRange(this, s, { r1: r, c1: c, r2: r, c2: c });
  }
  currentRegion(R) { const s = R.s; let { r1, c1, r2, c2 } = R.rg; const has = (r, c) => r >= 0 && c >= 0 && s.value(r, c) != null && s.value(r, c) !== ''; let grown = true; while (grown) { grown = false; for (let c = c1 - 1; c <= c2 + 1; c++) { if (has(r1 - 1, c)) { r1--; grown = true; break; } } for (let c = c1 - 1; c <= c2 + 1; c++) { if (has(r2 + 1, c)) { r2++; grown = true; break; } } for (let r = r1 - 1; r <= r2 + 1; r++) { if (has(r, c1 - 1)) { c1--; grown = true; break; } } for (let r = r1 - 1; r <= r2 + 1; r++) { if (has(r, c2 + 1)) { c2++; grown = true; break; } } } return new VRange(this, s, { r1: Math.max(0, r1), c1: Math.max(0, c1), r2, c2 }); }
  // ---- WorksheetFunction / Evaluate: run through the real formula engine
  evaluate(text, s) { const calc = this.doc.calc; s = s || this.activeSheet(); this.doc.recalc(true); const v = calc.evalFormula(String(text).replace(/^=/, ''), s, 0, 0); return this.fromCalc(v); }
  fromCalc(v) { if (v instanceof Err) throw new VbaError('Hàm trang tính trả về lỗi ' + v.code, 0, 1004); if (v && v.sheet && v.r1 != null) return new VRange(this, v.sheet, { r1: v.r1, c1: v.c1, r2: v.r2, c2: v.c2 }); if (Array.isArray(v)) { if (v.length === 1 && v[0].length === 1) return v[0][0]; return VbArray.from2D(v.map(r => r.map(x => x instanceof Err ? { __cverr: x.code } : x))); } return v; }
  callWsf(name, args) {
    const fn = name.toUpperCase(); if (!FN[fn] && !FN['_XLFN.' + fn]) throw new VbaError('WorksheetFunction không có hàm ' + fn, 0, 438);
    this.doc.recalc(true);
    const parts = Array.from(args).map(a => this.lit(a));
    const v = this.doc.calc.evalFormula(fn + '(' + parts.join(',') + ')', this.activeSheet(), 0, 0);
    return this.fromCalc(v);
  }
  lit(a) {
    if (a === undefined) return '';
    if (a instanceof VRange) return "'" + a.s.name.replace(/'/g, "''") + "'!" + rangeAddr(a.rg);
    if (a instanceof VbArray) return '{' + a.to2D().map(r => r.map(x => this.lit(x)).join(',')).join(';') + '}';
    if (typeof a === 'string') return '"' + a.replace(/"/g, '""') + '"';
    if (typeof a === 'boolean') return a ? 'TRUE' : 'FALSE';
    if (a instanceof Date) return String(jsDateToSerial(a));
    if (a == null) return '0';
    return String(a);
  }
}

// ---------------------------------------------------------------- built-in functions
const pad2 = n => (n < 10 ? '0' : '') + n;
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
export function vbFormat(v, f) {
  if (f == null || f === '') return toStr(v);
  const lf = String(f).toLowerCase();
  if (v instanceof Date || /^(short date|long date|general date|short time|long time|medium date)$/.test(lf) || (/[dmyhns]/i.test(f) && !/[#0]/.test(f) && (v instanceof Date || typeof v === 'number'))) {
    const d = v instanceof Date ? v : toDate(v);
    if (lf === 'short date') return pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) + '/' + d.getFullYear();
    if (lf === 'long date') return DAYS[d.getDay()] + ', ' + d.getDate() + ' ' + MONTHS[d.getMonth()] + ' ' + d.getFullYear();
    if (lf === 'short time') return pad2(d.getHours()) + ':' + pad2(d.getMinutes());
    if (lf === 'long time') return pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds());
    if (lf === 'general date') return toStr(d);
    return String(f).replace(/yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|n|ss|s|am\/pm/gi, t => { const l = t.toLowerCase();
      switch (l) { case 'yyyy': return d.getFullYear(); case 'yy': return pad2(d.getFullYear() % 100); case 'mmmm': return MONTHS[d.getMonth()]; case 'mmm': return MONTHS[d.getMonth()].slice(0, 3); case 'mm': return /h/i.test(f) && !/[dy]/i.test(f) ? pad2(d.getMinutes()) : pad2(d.getMonth() + 1); case 'm': return d.getMonth() + 1; case 'dddd': return DAYS[d.getDay()]; case 'ddd': return DAYS[d.getDay()].slice(0, 3); case 'dd': return pad2(d.getDate()); case 'd': return d.getDate(); case 'hh': return pad2(d.getHours()); case 'h': return d.getHours(); case 'nn': return pad2(d.getMinutes()); case 'n': return d.getMinutes(); case 'ss': return pad2(d.getSeconds()); case 's': return d.getSeconds(); case 'am/pm': return d.getHours() < 12 ? 'AM' : 'PM'; } return t; });
  }
  if (lf === 'currency') return format(toNum(v), '#,##0.00 ₫').t;
  if (lf === 'fixed') return toNum(v).toFixed(2);
  if (lf === 'standard') return format(toNum(v), '#,##0.00').t;
  if (lf === 'percent') return (toNum(v) * 100).toFixed(2) + '%';
  if (lf === 'general number') return toStr(toNum(v));
  if (lf === 'yes/no') return truthy(v) ? 'Yes' : 'No';
  if (lf === 'true/false') return truthy(v) ? 'True' : 'False';
  if (lf === 'on/off') return truthy(v) ? 'On' : 'Off';
  if (/^[@&<>!]+$/.test(f)) return toStr(v);
  try { return format(typeof v === 'string' && !isNaN(+v) ? +v : v, f).t; } catch (e) { return toStr(v); }
}
function fixArgs(a, n) { while (a.length < n) a.push(undefined); return a; }
export function builtins(m, interp) {
  const ui = m.ui;
  const B = {
    msgbox: async a => ui.msgbox(toStr(m.v(a[0])), a[1] == null ? 0 : toNum(a[1]), a[2] == null ? 'Microsoft Excel' : toStr(a[2])),
    inputbox: async a => { const r = await ui.inputbox(toStr(a[0]), a[1] == null ? 'Microsoft Excel' : toStr(a[1]), a[2] == null ? '' : toStr(a[2])); return r == null ? '' : r; },
    len: a => { const v = m.v(a[0]); return v == null ? 0 : toStr(v).length; }, lenb: a => toStr(m.v(a[0])).length * 2,
    left: a => toStr(m.v(a[0])).slice(0, Math.max(0, toNum(a[1]))), right: a => { const s = toStr(m.v(a[0])); const n = Math.max(0, toNum(a[1])); return n ? s.slice(-n) : ''; },
    mid: a => { const s = toStr(m.v(a[0])); const st = toNum(a[1]) - 1; return a[2] == null ? s.slice(st) : s.substr(st, Math.max(0, toNum(a[2]))); },
    ucase: a => toStr(m.v(a[0])).toUpperCase(), lcase: a => toStr(m.v(a[0])).toLowerCase(),
    trim: a => toStr(m.v(a[0])).replace(/^ +| +$/g, ''), ltrim: a => toStr(m.v(a[0])).replace(/^ +/, ''), rtrim: a => toStr(m.v(a[0])).replace(/ +$/, ''),
    instr: a => { let st = 1, s1, s2, cmpm = 0; if (a.length >= 3 && typeof m.v(a[0]) === 'number') { st = toNum(a[0]); s1 = toStr(m.v(a[1])); s2 = toStr(m.v(a[2])); cmpm = a[3] ? toNum(a[3]) : 0; } else { s1 = toStr(m.v(a[0])); s2 = toStr(m.v(a[1])); cmpm = a[2] ? toNum(a[2]) : 0; } if (cmpm === 1) { s1 = s1.toLowerCase(); s2 = s2.toLowerCase(); } if (s2 === '') return st; return s1.indexOf(s2, st - 1) + 1; },
    instrrev: a => { const s1 = toStr(m.v(a[0])), s2 = toStr(m.v(a[1])); const st = a[2] == null || toNum(a[2]) < 0 ? s1.length : toNum(a[2]); return s1.lastIndexOf(s2, st - s2.length) + 1; },
    replace: a => { const s = toStr(m.v(a[0])), f = toStr(a[1]), r = toStr(a[2]); if (!f) return s; const st = a[3] ? toNum(a[3]) : 1; const cnt = a[4] == null ? -1 : toNum(a[4]); const ci = a[5] && toNum(a[5]) === 1; let out = s.slice(st - 1), n = 0; const re = new RegExp(f.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), ci ? 'gi' : 'g'); out = out.replace(re, x => (cnt < 0 || n++ < cnt) ? r : x); return out; },
    split: a => { const s = toStr(m.v(a[0])); const d = a[1] == null ? ' ' : toStr(a[1]); if (s === '') return VbArray.from1D([]); let parts = d === '' ? [s] : s.split(d); if (a[2] != null && toNum(a[2]) > 0) { const n = toNum(a[2]); if (parts.length > n) parts = parts.slice(0, n - 1).concat([parts.slice(n - 1).join(d)]); } return VbArray.from1D(parts); },
    join: a => { const arr = a[0]; const d = a[1] == null ? ' ' : toStr(a[1]); return arr instanceof VbArray ? arr.data.map(toStr).join(d) : toStr(arr); },
    strreverse: a => toStr(m.v(a[0])).split('').reverse().join(''), space: a => ' '.repeat(Math.max(0, toNum(a[0]))), string: a => toStr(a[1]).charAt(0).repeat(Math.max(0, toNum(a[0]))),
    strcomp: a => { let x = toStr(m.v(a[0])), y = toStr(m.v(a[1])); if (a[2] && toNum(a[2]) === 1) { x = x.toLowerCase(); y = y.toLowerCase(); } return x < y ? -1 : x > y ? 1 : 0; },
    strconv: a => { const s = toStr(m.v(a[0])); const k = toNum(a[1]); return k === 1 ? s.toUpperCase() : k === 2 ? s.toLowerCase() : k === 3 ? s.toLowerCase().replace(/(^|\s)(\S)/g, (x, p, c) => p + c.toUpperCase()) : s; },
    chr: a => String.fromCharCode(toNum(a[0])), chrw: a => String.fromCharCode(toNum(a[0])), asc: a => toStr(m.v(a[0])).charCodeAt(0) || 0, ascw: a => toStr(m.v(a[0])).charCodeAt(0) || 0,
    cstr: a => toStr(m.v(a[0])), cint: a => bankRound(toNum(m.v(a[0]))), clng: a => bankRound(toNum(m.v(a[0]))), clnglng: a => bankRound(toNum(m.v(a[0]))), cbyte: a => bankRound(toNum(m.v(a[0]))) & 255,
    cdbl: a => toNum(m.v(a[0])), csng: a => toNum(m.v(a[0])), ccur: a => Math.round(toNum(m.v(a[0])) * 10000) / 10000, cdec: a => toNum(m.v(a[0])), cbool: a => truthy(m.v(a[0])), cdate: a => toDate(m.v(a[0])), cvar: a => m.v(a[0]), cverr: a => ({ __cverr: '#N/A' }),
    val: a => { const s = toStr(m.v(a[0])).replace(/\s/g, ''); const x = /^[+-]?(\d*\.?\d+(e[+-]?\d+)?|&h[0-9a-f]+)/i.exec(s); return x ? (/^&h/i.test(x[0]) ? parseInt(x[0].slice(2), 16) : parseFloat(x[0])) : 0; },
    str: a => { const n = toNum(m.v(a[0])); return (n >= 0 ? ' ' : '') + toStr(n); },
    int: a => Math.floor(toNum(m.v(a[0]))), fix: a => Math.trunc(toNum(m.v(a[0]))), abs: a => Math.abs(toNum(m.v(a[0]))), sgn: a => Math.sign(toNum(m.v(a[0]))),
    sqr: a => Math.sqrt(toNum(m.v(a[0]))), exp: a => Math.exp(toNum(m.v(a[0]))), log: a => Math.log(toNum(m.v(a[0]))), sin: a => Math.sin(toNum(a[0])), cos: a => Math.cos(toNum(a[0])), tan: a => Math.tan(toNum(a[0])), atn: a => Math.atan(toNum(a[0])),
    round: a => { const n = toNum(m.v(a[0])), d = a[1] == null ? 0 : toNum(a[1]); const f = Math.pow(10, d); return bankRound(n * f) / f; },
    rnd: () => Math.random(), randomize: () => undefined,
    hex: a => Math.trunc(toNum(a[0])).toString(16).toUpperCase(), oct: a => Math.trunc(toNum(a[0])).toString(8),
    now: () => new Date(), date: () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }, time: () => { const d = new Date(); return new Date(1899, 11, 30, d.getHours(), d.getMinutes(), d.getSeconds()); }, timer: () => { const d = new Date(); return (d - new Date(d.getFullYear(), d.getMonth(), d.getDate())) / 1000; },
    year: a => toDate(m.v(a[0])).getFullYear(), month: a => toDate(m.v(a[0])).getMonth() + 1, day: a => toDate(m.v(a[0])).getDate(), hour: a => toDate(m.v(a[0])).getHours(), minute: a => toDate(m.v(a[0])).getMinutes(), second: a => toDate(m.v(a[0])).getSeconds(),
    weekday: a => { const d = toDate(m.v(a[0])).getDay(); const fd = a[1] == null ? 1 : toNum(a[1]); return ((d - (fd - 1) + 7) % 7) + 1; },
    weekdayname: a => DAYS[(toNum(a[0]) - 1 + 7) % 7], monthname: a => { const n = MONTHS[toNum(a[0]) - 1]; return a[1] && truthy(a[1]) ? n.slice(0, 3) : n; },
    dateserial: a => new Date(toNum(a[0]), toNum(a[1]) - 1, toNum(a[2])), timeserial: a => new Date(1899, 11, 30, toNum(a[0]), toNum(a[1]), toNum(a[2])),
    datevalue: a => { const d = toDate(m.v(a[0])); d.setHours(0, 0, 0, 0); return d; }, timevalue: a => { const d = toDate(m.v(a[0])); return new Date(1899, 11, 30, d.getHours(), d.getMinutes(), d.getSeconds()); },
    dateadd: a => { const iv = toStr(a[0]).toLowerCase(), n = toNum(a[1]); const d = new Date(+toDate(m.v(a[2]))); switch (iv) { case 'yyyy': d.setFullYear(d.getFullYear() + n); break; case 'q': d.setMonth(d.getMonth() + 3 * n); break; case 'm': d.setMonth(d.getMonth() + n); break; case 'ww': d.setDate(d.getDate() + 7 * n); break; case 'd': case 'y': case 'w': d.setDate(d.getDate() + n); break; case 'h': d.setHours(d.getHours() + n); break; case 'n': d.setMinutes(d.getMinutes() + n); break; case 's': d.setSeconds(d.getSeconds() + n); break; } return d; },
    datediff: a => { const iv = toStr(a[0]).toLowerCase(); const d1 = toDate(m.v(a[1])), d2 = toDate(m.v(a[2])); switch (iv) { case 'yyyy': return d2.getFullYear() - d1.getFullYear(); case 'q': return Math.floor(d2.getMonth() / 3) - Math.floor(d1.getMonth() / 3) + 4 * (d2.getFullYear() - d1.getFullYear()); case 'm': return d2.getMonth() - d1.getMonth() + 12 * (d2.getFullYear() - d1.getFullYear()); case 'ww': case 'w': return Math.trunc((d2 - d1) / 6048e5); case 'h': return Math.trunc((d2 - d1) / 36e5); case 'n': return Math.trunc((d2 - d1) / 6e4); case 's': return Math.trunc((d2 - d1) / 1e3); default: { const a1 = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate()), a2 = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate()); return Math.round((a2 - a1) / 864e5); } } },
    datepart: a => { const iv = toStr(a[0]).toLowerCase(), d = toDate(m.v(a[1])); return { yyyy: d.getFullYear(), q: Math.floor(d.getMonth() / 3) + 1, m: d.getMonth() + 1, d: d.getDate(), w: d.getDay() + 1, h: d.getHours(), n: d.getMinutes(), s: d.getSeconds(), y: Math.round((d - new Date(d.getFullYear(), 0, 1)) / 864e5) + 1 }[iv]; },
    format: a => vbFormat(m.v(a[0]), a[1] == null ? null : toStr(a[1])), formatnumber: a => format(toNum(m.v(a[0])), '#,##0' + (a[1] == null || toNum(a[1]) > 0 ? '.' + '0'.repeat(a[1] == null ? 2 : toNum(a[1])) : '')).t,
    formatpercent: a => (toNum(m.v(a[0])) * 100).toFixed(a[1] == null ? 2 : toNum(a[1])) + '%', formatcurrency: a => format(toNum(m.v(a[0])), '#,##0.00 ₫').t, formatdatetime: a => vbFormat(toDate(m.v(a[0])), ['general date', 'long date', 'short date', 'long time', 'short time'][a[1] == null ? 0 : toNum(a[1])]),
    isnumeric: a => { const v = m.v(a[0]); if (typeof v === 'number') return true; if (typeof v === 'boolean') return true; if (typeof v !== 'string' || v.trim() === '') return false; return !isNaN(+v.replace(/,/g, '')); },
    isempty: a => { const v = m.v(a[0]); return v === undefined || v === ''; }, isnull: a => a[0] === null && !(a[0] instanceof Object), isdate: a => { const v = m.v(a[0]); if (v instanceof Date) return true; try { toDate(v); return typeof v === 'string'; } catch (e) { return false; } },
    isarray: a => a[0] instanceof VbArray, isobject: a => a[0] instanceof VObj || a[0] instanceof Collection || a[0] instanceof Dictionary || a[0] === null, iserror: a => !!(m.v(a[0]) && m.v(a[0]).__cverr), ismissing: a => a[0] === undefined,
    ubound: a => { const arr = a[0]; if (!(arr instanceof VbArray)) throw new VbaError('UBound cần một mảng', 0, 13); const d = a[1] == null ? 1 : toNum(a[1]); if (arr.bounds[d - 1][1] < arr.bounds[d - 1][0]) throw new VbaError('Mảng rỗng (Subscript out of range)', 0, 9); return arr.bounds[d - 1][1]; },
    lbound: a => { const arr = a[0]; if (!(arr instanceof VbArray)) throw new VbaError('LBound cần một mảng', 0, 13); const d = a[1] == null ? 1 : toNum(a[1]); return arr.bounds[d - 1][0]; },
    iif: a => truthy(m.v(a[0])) ? a[1] : a[2], choose: a => a[toNum(a[0])], switch: a => { for (let i = 0; i + 1 < a.length; i += 2) if (truthy(m.v(a[i]))) return a[i + 1]; return null; },
    rgb: a => rgbLong(toNum(a[0]), toNum(a[1]), toNum(a[2])), qbcolor: a => { const c = QB[toNum(a[0]) & 15]; return rgbLong((c >> 16) & 255, (c >> 8) & 255, c & 255); },
    typename: a => m.typeName(a[0]), vartype: a => { const v = a[0]; if (v === undefined) return 0; if (v === null) return 1; if (typeof v === 'number') return Number.isInteger(v) ? 3 : 5; if (typeof v === 'string') return 8; if (typeof v === 'boolean') return 11; if (v instanceof Date) return 7; if (v instanceof VbArray) return 8204; return 9; },
    filter: a => { const arr = a[0]; const f = toStr(a[1]); const inc = a[2] == null ? true : truthy(a[2]); return VbArray.from1D(arr.data.filter(x => (toStr(x).toLowerCase().indexOf(f.toLowerCase()) >= 0) === inc)); },
    environ: a => ({ username: 'User', computername: 'PC', os: 'Windows_NT' })[toStr(a[0]).toLowerCase()] || '',
    doevents: () => new Promise(r => setTimeout(() => r(0), 0)),
    beep: () => undefined, sendkeys: () => undefined, shell: () => 0, appactivate: () => undefined,
    createobject: a => m.create(toStr(a[0]).toLowerCase()), getobject: () => null,
    callbyname: async a => { const o = a[0], name = toStr(a[1]).toLowerCase(); const ct = toNum(a[2]); if (ct === 4 || ct === 8) { await m.setMember(o, name, null, a[3]); return; } return m.member(o, name, Array.from(a).slice(3)); },
    evaluate: a => m.evaluate(toStr(a[0])),
    range: a => m.range(m.activeSheet(), a[0], a[1]),
    cells: a => m.member(m.sheetObj(m.activeSheet()), 'cells', a),
    rows: a => m.member(m.sheetObj(m.activeSheet()), 'rows', a), columns: a => m.member(m.sheetObj(m.activeSheet()), 'columns', a),
    worksheets: a => a.length ? m.member(new VSheets(m), '', a) : new VSheets(m), sheets: a => a.length ? m.member(new VSheets(m), '', a) : new VSheets(m),
    activesheet: () => m.sheetObj(m.activeSheet()), activecell: () => m.activeCell(), selection: () => m.selRange(),
    activeworkbook: () => m.book, thisworkbook: () => m.book, workbooks: a => a.length ? m.book : { __coll: [m.book] }, application: () => m.app, worksheetfunction: () => m.wsf, debug: () => m.dbg,
    intersect: a => { let r = null; for (const x of a) { if (!(x instanceof VRange)) continue; if (!r) { r = x.rg; continue; } r = { r1: Math.max(r.r1, x.rg.r1), c1: Math.max(r.c1, x.rg.c1), r2: Math.min(r.r2, x.rg.r2), c2: Math.min(r.c2, x.rg.c2) }; if (r.r1 > r.r2 || r.c1 > r.c2) return null; } return r ? new VRange(m, a[0].s, r) : null; },
    union: a => { let r = null; for (const x of a) { if (!(x instanceof VRange)) continue; r = r ? { r1: Math.min(r.r1, x.rg.r1), c1: Math.min(r.c1, x.rg.c1), r2: Math.max(r.r2, x.rg.r2), c2: Math.max(r.c2, x.rg.c2) } : Object.assign({}, x.rg); } return r ? new VRange(m, a[0].s, r) : null; },
    calculate: () => { m.doc.recalc(true); }
  };
  B.err = undefined; delete B.err;
  return B;
}

// ---------------------------------------------------------------- member access dispatch
export function installDispatch(m, interp) {
  const B = builtins(m, interp);
  m.v = v => (v instanceof VRange ? m.getValue(v) : v);
  m.create = cls => { cls = String(cls).toLowerCase(); if (cls === 'collection') return new Collection(); if (/dictionary$/.test(cls)) return new Dictionary(); if (cls === 'scripting.filesystemobject') throw new VbaError('FileSystemObject không khả dụng trong chế độ an toàn', 0, 429); if (cls === 'vbscript.regexp' || cls === 'regexp') return new VRegExp(); return new Collection(); };
  m.typeName = v => { if (v && v.__user) return v.mod.type === 'class' ? v.mod.name : m.typeName(v.hostObj); if (v === undefined) return 'Empty'; if (v === null) return 'Nothing'; if (typeof v === 'number') return Number.isInteger(v) ? 'Long' : 'Double'; if (typeof v === 'string') return 'String'; if (typeof v === 'boolean') return 'Boolean'; if (v instanceof Date) return 'Date'; if (v instanceof VbArray) return 'Variant()'; if (v instanceof VRange) return 'Range'; if (v instanceof VSheet) return 'Worksheet'; if (v instanceof VSheets) return 'Sheets'; if (v instanceof VBook) return 'Workbook'; if (v instanceof VApp) return 'Application'; if (v instanceof Collection) return 'Collection'; if (v instanceof Dictionary) return 'Dictionary'; if (v && v.__cverr) return 'Error'; return 'Object'; };
  m.member = (o, name, args, ip, wantValue, isStmt) => member(m, o, name, args || [], ip);
  m.setMember = (o, name, args, val) => setMember(m, o, name, args || [], val);
  const host = {
    value: v => (v instanceof VRange ? m.getValue(v) : v),
    create: m.create, typeName: m.typeName,
    bracket: t => m.range(m.activeSheet(), t),
    same: (a, b) => { if (a instanceof VRange && b instanceof VRange) return a.s === b.s && rangeAddr(a.rg) === rangeAddr(b.rg); if (a instanceof VSheet && b instanceof VSheet) return a.s === b.s; return a === b; },
    hasDefault: o => o instanceof VRange,
    each: coll => { if (coll instanceof VbArray) return coll.data.slice(); if (coll && coll.each) return coll.each(); if (coll && coll.__coll) return coll.__coll; if (coll == null) throw new VbaError('For Each cần một tập hợp (Object required)', 0, 424); throw new VbaError('Không thể lặp For Each trên giá trị này', 0, 451); },
    async ident(lc, args, ip, sc, isStmt) {
      if (Object.prototype.hasOwnProperty.call(CONSTS, lc)) return CONSTS[lc];
      const f = B[lc]; if (f) return await f(args || [], ip);
      // sheet code names (Sheet1.Range...) and workbook-level names
      const sh = m.wb().sheets.find(s => (s.codeName || s.name).toLowerCase().replace(/\s/g, '') === lc); if (sh) return m.sheetObj(sh);
      return NOT_FOUND;
    },
    async setIdent(lc, val) { if (lc === 'activecell') { m.setValue(m.activeCell(), val); return true; } if (lc === 'selection') { m.setValue(m.selRange(), val); return true; } return false; },
    async member(o, name, args, ip) { return member(m, o, name, args || [], ip); },
    async setMember(o, name, args, val) { return setMember(m, o, name, args || [], val); }
  };
  return host;
}
class VRegExp { constructor() { this.pattern = ''; this.global = false; this.ignorecase = false; this.multiline = false; } re() { return new RegExp(this.pattern, (this.global ? 'g' : '') + (this.ignorecase ? 'i' : '') + (this.multiline ? 'm' : '')); } }

function idx1(a, n) { return toNum(a) - 1 + (n || 0); }
export function member(m, o, name, args, ip) {
  if (o === undefined || o === null) throw new VbaError('Biến đối tượng chưa được gán (Object variable not set) khi truy cập ".' + name + '"', 0, 91);
  if (o instanceof Collection || o instanceof Dictionary) { if (name === 'count' && o instanceof Collection) return o.items.length; return o.vcall(name, Array.from(args)); }
  if (o instanceof VRegExp) { switch (name) { case 'pattern': case 'global': case 'ignorecase': case 'multiline': return o[name]; case 'test': return o.re().test(toStr(m.v(args[0]))); case 'replace': return toStr(m.v(args[0])).replace(o.re(), toStr(args[1]).replace(/\$(\d)/g, '$$$1')); case 'execute': { const s = toStr(m.v(args[0])); const out = []; const re = o.re(); let x; if (o.global) { while ((x = re.exec(s))) { out.push(matchObj(x)); if (!x[0]) re.lastIndex++; } } else if ((x = re.exec(s))) out.push(matchObj(x)); const c = new Collection(); out.forEach(i => c.items.push(i) && c.keys.push(null)); return c; } } }
  if (o && o.__match) { const x = o.__match; if (name === 'value' || name === '') return x[0]; if (name === 'firstindex') return x.index; if (name === 'length') return x[0].length; if (name === 'submatches') return VbArray.from1D(x.slice(1)); }
  if (typeof o === 'string' && name === '') return o;
  if (o instanceof VRange) return rangeMember(m, o, name, args);
  if (o instanceof VRows) { const r = o.rng; if (name === 'count') return o.isCol ? r.cols : r.rows; if (name === '' || name === 'item') { const i = toNum(args[0]) - 1; return o.isCol ? new VRange(m, r.s, { r1: r.rg.r1, c1: r.rg.c1 + i, r2: r.rg.r2, c2: r.rg.c1 + i }) : new VRange(m, r.s, { r1: r.rg.r1 + i, c1: r.rg.c1, r2: r.rg.r1 + i, c2: r.rg.c2 }); } const whole = o.isCol ? new VRange(m, r.s, { r1: 0, c1: r.rg.c1, r2: MAX_ROWS - 1, c2: r.rg.c2 }) : new VRange(m, r.s, { r1: r.rg.r1, c1: 0, r2: r.rg.r2, c2: MAX_COLS - 1 }); return rangeMember(m, whole, name, args); }
  if (o instanceof VFont) { const st = m.styleOf(o.rng.r); switch (name) { case 'bold': return !!st.b; case 'italic': return !!st.i; case 'underline': return st.u ? 2 : -4142; case 'strikethrough': return !!st.s; case 'size': return st.sz || 11; case 'name': return st.font || 'Calibri'; case 'color': return hexToLong(st.color || '#000000'); case 'colorindex': return 1; } throw new VbaError('Font không có thuộc tính ' + name, 0, 438); }
  if (o instanceof VInterior) { const st = m.styleOf(o.rng); switch (name) { case 'color': return st.fill ? hexToLong(st.fill) : 16777215; case 'colorindex': return st.fill ? Math.max(1, COLORINDEX.indexOf(String(st.fill).toUpperCase())) : -4142; case 'pattern': return st.fill ? 1 : -4142; } throw new VbaError('Interior không có thuộc tính ' + name, 0, 438); }
  if (o instanceof VBorders) { if (name === '' || name === 'item') return new VBorders(o.rng, toNum(args[0])); if (name === 'linestyle') return 1; if (name === 'weight') return 2; if (name === 'color') return 0; return undefined; }
  if (o instanceof VSheet) return sheetMember(m, o, name, args);
  if (o instanceof VSheets) { const wb = m.wb(); if (name === 'count') return wb.sheets.length; if (name === '' || name === 'item') { const k = args[0]; const s = typeof k === 'number' ? wb.sheets[k - 1] : wb.sheet(toStr(k)); if (!s) throw new VbaError('Không có trang tính "' + toStr(k) + '" (Subscript out of range)', 0, 9); return m.sheetObj(s); }
    if (name === 'add') { let at = wb.sheets.length; const named = args.named || {}; const before = named.before || args[0], after = named.after || args[1]; if (before instanceof VSheet) at = wb.sheets.indexOf(before.s); else if (after instanceof VSheet) at = wb.sheets.indexOf(after.s) + 1; else if (args[0] === undefined && args[1] === undefined && !named.before && !named.after) at = wb.active; const s = wb.addSheet(wb.nextName(), at); m.structural = true; m.doc.wb.active = wb.sheets.indexOf(s); return m.sheetObj(s); }
    if (name === 'select') return; throw new VbaError('Worksheets không có thành viên ' + name, 0, 438); }
  if (o instanceof VBook) { switch (name) { case 'worksheets': case 'sheets': return args.length ? member(m, new VSheets(m), '', args) : new VSheets(m); case 'activesheet': return m.sheetObj(m.activeSheet()); case 'name': return (m.doc.name || 'Book1') + '.xlsx'; case 'fullname': return m.doc.path || m.doc.name; case 'path': return m.doc.path ? String(m.doc.path).replace(/[\\/][^\\/]*$/, '') : ''; case 'save': m.saveRequested = 'save'; return; case 'saveas': m.saveRequested = 'saveas'; return; case 'close': return; case 'saved': return !m.doc.dirty; case 'names': return { __coll: [] }; case 'activate': return; }
    throw new VbaError('Workbook không có thành viên ' + name, 0, 438); }
  if (o instanceof VApp) { switch (name) { case 'screenupdating': return m.screenUpdating; case 'displayalerts': case 'enableevents': return true; case 'calculation': return m.wb().calcMode === 'manual' ? -4135 : -4105; case 'statusbar': return m.statusBar; case 'username': return m.ui.userName || 'User'; case 'version': return '16.0'; case 'name': return 'Microsoft Excel'; case 'operatingsystem': return navigator.userAgent.indexOf('Android') >= 0 ? 'Android' : 'Windows (64-bit) NT 10.00'; case 'worksheetfunction': return m.wsf; case 'activesheet': return m.sheetObj(m.activeSheet()); case 'activecell': return m.activeCell(); case 'selection': return m.selRange(); case 'activeworkbook': case 'thisworkbook': return m.book; case 'workbooks': return { __coll: [m.book] }; case 'worksheets': case 'sheets': return args.length ? member(m, new VSheets(m), '', args) : new VSheets(m); case 'range': return m.range(m.activeSheet(), args[0], args[1]); case 'cells': return sheetMember(m, m.sheetObj(m.activeSheet()), 'cells', args); case 'inputbox': { const conv = r => { if (r == null) return false; const t = args.named && args.named.type != null ? toNum(args.named.type) : (args[7] != null ? toNum(args[7]) : 2); if (t === 1) return toNum(r); if (t === 8) return m.range(m.activeSheet(), r); return r; }; const r = m.ui.inputbox(toStr(args[0]), args[1] == null ? 'Nhập' : toStr(args[1]), args[2] == null ? '' : toStr(args[2])); return r && r.then ? r.then(conv) : conv(r); } case 'run': return ip.run(toStr(args[0]), Array.from(args).slice(1)); case 'wait': case 'ontime': case 'onkey': return true; case 'calculate': case 'calculatefull': m.doc.recalc(true); return; case 'cutcopymode': return false; case 'intersect': return (builtins(m, ip)).intersect(args); case 'union': return (builtins(m, ip)).union(args); case 'evaluate': return m.evaluate(toStr(args[0])); case 'goto': { const r = args[0] instanceof VRange ? args[0] : m.range(m.activeSheet(), args[0]); m.gotoRange = r; return; } case 'volatile': return; case 'quit': return; case 'getopenfilename': return false; case 'caller': return null; case 'max': case 'min': case 'sum': return m.callWsf(name, args); }
    throw new VbaError('Application không có thành viên ' + name, 0, 438); }
  if (o instanceof VWsf) return m.callWsf(name, args);
  if (o instanceof VDebug) { if (name === 'print') { m.ui.print(Array.from(args).map(x => toStr(m.v(x))).join(' ')); return; } if (name === 'assert') return; }
  if (o && o.__coll) { if (name === 'count') return o.__coll.length; if (name === '' || name === 'item') return o.__coll[toNum(args[0]) - 1]; }
  if (o && o.__vbaErr) return o.__err(name);
  throw new VbaError('Đối tượng không hỗ trợ thuộc tính/phương thức "' + name + '" (' + m.typeName(o) + ')', 0, 438);
}
function matchObj(x) { return { __match: x }; }

function rangeMember(m, R, name, a) {
  const s = R.s, rg = R.rg;
  switch (name) {
    case '': case 'value': case 'value2': if (a.length) return rangeMember(m, R, 'cells', a); return m.getValue(R);
    case 'item': case 'cells': {
      if (!a.length) return R;
      let r, c; if (a.length === 1) { const i = toNum(a[0]) - 1; r = rg.r1 + Math.floor(i / R.cols); c = rg.c1 + (i % R.cols); } else { r = rg.r1 + toNum(a[0]) - 1; c = rg.c1 + (typeof a[1] === 'string' && isNaN(+a[1]) ? colIndex(a[1]) : toNum(a[1]) - 1); }
      if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) throw new VbaError('Ô nằm ngoài trang tính (1004)', 0, 1004);
      return new VRange(m, s, { r1: r, c1: c, r2: r, c2: c });
    }
    case 'range': return (() => { const sub = m.range(s, a[0], a[1]); return new VRange(m, s, { r1: rg.r1 + sub.rg.r1, c1: rg.c1 + sub.rg.c1, r2: rg.r1 + sub.rg.r2, c2: rg.c1 + sub.rg.c2 }); })();
    case 'text': { const v = s.value(rg.r1, rg.c1); const st = s.styleAt(rg.r1, rg.c1); try { return v == null ? '' : (v instanceof Err ? v.code : format(v, st && st.nf).t); } catch (e) { return toStr(v); } }
    case 'formula': case 'formulalocal': case 'formula2': { const x = s.get(rg.r1, rg.c1); return x && x.f ? '=' + x.f : (x && x.v != null ? toStr(x.v) : ''); }
    case 'formular1c1': { const x = s.get(rg.r1, rg.c1); return x && x.f ? '=' + x.f : ''; }
    case 'hasformula': { const x = s.get(rg.r1, rg.c1); return !!(x && x.f); }
    case 'row': return rg.r1 + 1; case 'column': return rg.c1 + 1;
    case 'rows': return a.length ? member(m, new VRows(R, false), '', a) : new VRows(R, false);
    case 'columns': return a.length ? member(m, new VRows(R, true), '', a) : new VRows(R, true);
    case 'count': case 'countlarge': return R.rows * R.cols;
    case 'address': { const named = a.named || {}; const ra = named.rowabsolute != null ? truthy(named.rowabsolute) : (a[0] == null ? true : truthy(a[0])); const ca = named.columnabsolute != null ? truthy(named.columnabsolute) : (a[1] == null ? true : truthy(a[1])); const one = (r, c) => addr(r, c, ra, ca); const t = one(rg.r1, rg.c1) + (R.rows * R.cols > 1 ? ':' + one(rg.r2, rg.c2) : ''); return (named.external && truthy(named.external)) || (a[3] && truthy(a[3])) ? "[" + (m.doc.name || 'Book1') + "]" + s.name + '!' + t : t; }
    case 'offset': { const dr = a[0] == null ? 0 : toNum(a[0]), dc = a[1] == null ? 0 : toNum(a[1]); return new VRange(m, s, { r1: rg.r1 + dr, c1: rg.c1 + dc, r2: rg.r2 + dr, c2: rg.c2 + dc }); }
    case 'resize': { const nr = a[0] == null ? R.rows : toNum(a[0]), nc = a[1] == null ? R.cols : toNum(a[1]); return new VRange(m, s, { r1: rg.r1, c1: rg.c1, r2: rg.r1 + nr - 1, c2: rg.c1 + nc - 1 }); }
    case 'end': return m.endOf(R, a[0]);
    case 'entirerow': return new VRange(m, s, { r1: rg.r1, c1: 0, r2: rg.r2, c2: MAX_COLS - 1 });
    case 'entirecolumn': return new VRange(m, s, { r1: 0, c1: rg.c1, r2: MAX_ROWS - 1, c2: rg.c2 });
    case 'currentregion': return m.currentRegion(R);
    case 'worksheet': case 'parent': return m.sheetObj(s);
    case 'font': return new VFont({ r: R });
    case 'interior': return new VInterior(R);
    case 'borders': return a.length ? new VBorders(R, toNum(a[0])) : new VBorders(R, 0);
    case 'numberformat': case 'numberformatlocal': return (m.styleOf(R).nf) || 'General';
    case 'horizontalalignment': { const h = m.styleOf(R).ha; return h === 'center' ? -4108 : h === 'right' ? -4152 : h === 'left' ? -4131 : 1; }
    case 'verticalalignment': { const v = m.styleOf(R).va; return v === 'top' ? -4160 : v === 'middle' || v === 'center' ? -4108 : -4107; }
    case 'wraptext': return !!m.styleOf(R).wrap;
    case 'columnwidth': return Math.round(s.colWidth(rg.c1) / 7 * 100) / 100;
    case 'rowheight': return Math.round(s.rowHeight(rg.r1) * 0.75 * 100) / 100;
    case 'width': return s.colWidth(rg.c1) * 0.75; case 'height': return s.rowHeight(rg.r1) * 0.75;
    case 'hidden': return R.rows > 50000 ? !!s.hiddenC[rg.c1] : !!s.hiddenR[rg.r1];
    case 'mergecells': return !!s.mergeAt(rg.r1, rg.c1);
    case 'mergearea': { const mm = s.mergeAt(rg.r1, rg.c1); return mm ? new VRange(m, s, mm) : R; }
    case 'comment': { const x = s.get(rg.r1, rg.c1); return x && x.note ? { __coll: [], text: x.note.text || x.note } : null; }
    case 'select': case 'activate': { m.touch(s); m.doc.wb.active = m.wb().sheets.indexOf(s); if (name === 'activate' && R.rows * R.cols === 1) { s.sel.r = rg.r1; s.sel.c = rg.c1; return; } s.sel = { ranges: [Object.assign({}, rg)], r: rg.r1, c: rg.c1 }; m.selMoved = true; return; }
    case 'clearcontents': m.touch(s); { const lim = m.limitRange(R); for (let r = lim.r1; r <= lim.r2; r++) for (let c = lim.c1; c <= lim.c2; c++) s.setCell(r, c, { v: null, f: null }); } return;
    case 'clear': m.touch(s); { const lim = m.limitRange(R); for (let r = lim.r1; r <= lim.r2; r++) for (let c = lim.c1; c <= lim.c2; c++) s.cells.delete(key(r, c)); s.rev++; } return;
    case 'clearformats': m.touch(s); { const lim = m.limitRange(R); for (let r = lim.r1; r <= lim.r2; r++) for (let c = lim.c1; c <= lim.c2; c++) s.setCell(r, c, { s: null }); } return;
    case 'clearcomments': case 'clearnotes': for (let r = rg.r1; r <= Math.min(rg.r2, s.used().r2); r++) for (let c = rg.c1; c <= Math.min(rg.c2, s.used().c2); c++) s.setCell(r, c, { note: null }); return;
    case 'merge': m.touch(s); s.merges = s.merges.filter(x => x.r2 < rg.r1 || x.r1 > rg.r2 || x.c2 < rg.c1 || x.c1 > rg.c2); s.merges.push(Object.assign({}, rg)); return;
    case 'unmerge': m.touch(s); s.merges = s.merges.filter(x => x.r2 < rg.r1 || x.r1 > rg.r2 || x.c2 < rg.c1 || x.c1 > rg.c2); return;
    case 'autofit': m.touch(s); if (R.rows > 50000 || name === 'autofit') { for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 200); c++) { let w = 20; const u = s.used(); for (let r = 0; r <= u.r2; r++) { const v = s.value(r, c); if (v != null) w = Math.max(w, String(v).length * 7.5 + 10); } s.colW[c] = Math.min(600, Math.round(w)); } if (s._ax) s._ax.c.dirty = true; } return;
    case 'copy': { const named = a.named || {}; const dest = named.destination || a[0]; if (dest instanceof VRange) { copyRange(m, R, dest, true); return; } m.clip = R; return; }
    case 'cut': { const dest = (a.named || {}).destination || a[0]; if (dest instanceof VRange) { copyRange(m, R, dest, true); for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) s.cells.delete(key(r, c)); s.rev++; return; } m.clip = R; return; }
    case 'pastespecial': { if (!m.clip) return; const t = a[0] == null ? -4104 : toNum(a[0]); copyRange(m, m.clip, R, t !== -4163, t === -4163 || t === -4104, t === -4122); return; }
    case 'delete': case 'insert': { m.touch(s); m.structural = true; const entireRow = rg.c1 === 0 && rg.c2 >= MAX_COLS - 1, entireCol = rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1; const ins = name === 'insert';
      if (entireRow || (!entireCol && (a[0] == null ? true : toNum(a[0]) === -4162 || toNum(a[0]) === -4121))) { m.doc.structOp ? m.doc.structOp(s, ins ? 'insertRows' : 'deleteRows', rg.r1, R.rows) : shiftCells(s, 'r', rg.r1, ins ? R.rows : -R.rows); }
      else { m.doc.structOp ? m.doc.structOp(s, ins ? 'insertCols' : 'deleteCols', rg.c1, R.cols) : shiftCells(s, 'c', rg.c1, ins ? R.cols : -R.cols); }
      return; }
    case 'find': { const what = toStr(m.v(a[0])).toLowerCase(); const named = a.named || {}; const whole = named.lookat != null ? toNum(named.lookat) === 1 : false; const u = s.used(); const lim = { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, u.r2), c2: Math.min(rg.c2, u.c2) };
      for (let r = lim.r1; r <= lim.r2; r++) for (let c = lim.c1; c <= lim.c2; c++) { const v = s.value(r, c); if (v == null) continue; const t = String(v instanceof Err ? v.code : v).toLowerCase(); if (whole ? t === what : t.indexOf(what) >= 0) return new VRange(m, s, { r1: r, c1: c, r2: r, c2: c }); } return null; }
    case 'replace': { const what = toStr(a[0]), rep = toStr(a[1]); const u = s.used(); m.touch(s); let n = 0; for (let r = rg.r1; r <= Math.min(rg.r2, u.r2); r++) for (let c = rg.c1; c <= Math.min(rg.c2, u.c2); c++) { const x = s.get(r, c); if (x && typeof x.v === 'string' && !x.f && x.v.indexOf(what) >= 0) { s.setCell(r, c, { v: x.v.split(what).join(rep) }); n++; } } return n > 0; }
    case 'sort': { m.touch(s); const named = a.named || {}; const k1 = named.key1 || a[0]; const o1 = named.order1 != null ? toNum(named.order1) : (a[1] != null ? toNum(a[1]) : 1); const hdr = named.header != null ? toNum(named.header) === 1 : false; const col = k1 instanceof VRange ? k1.rg.c1 : rg.c1;
      const r1 = rg.r1 + (hdr ? 1 : 0); const rows = []; for (let r = r1; r <= Math.min(rg.r2, s.used().r2); r++) { const row = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(s.get(r, c) ? Object.assign({}, s.get(r, c)) : null); rows.push(row); }
      rows.sort((x, y) => { const vx = x[col - rg.c1] ? x[col - rg.c1].v : null, vy = y[col - rg.c1] ? y[col - rg.c1].v : null; if (vx == null) return 1; if (vy == null) return -1; const r = cmp(vx, vy); return o1 === 2 ? -r : r; });
      rows.forEach((row, i) => row.forEach((x, j) => { const k = key(r1 + i, rg.c1 + j); if (x) s.cells.set(k, x); else s.cells.delete(k); })); s.rev++; return; }
    case 'removeduplicates': { m.touch(s); const cols = a[0] instanceof VbArray ? a[0].data.map(x => toNum(x) - 1) : [toNum(a[0] == null ? 1 : a[0]) - 1]; const hdr = (a.named && a.named.header != null ? toNum(a.named.header) : (a[1] != null ? toNum(a[1]) : 2)) === 1; const seen = new Set(); const keep = []; for (let r = rg.r1 + (hdr ? 1 : 0); r <= Math.min(rg.r2, s.used().r2); r++) { const k = cols.map(c => String(s.value(r, rg.c1 + c))).join('\u0001'); if (seen.has(k)) continue; seen.add(k); const row = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(s.get(r, c) ? Object.assign({}, s.get(r, c)) : null); keep.push(row); } for (let r = rg.r1 + (hdr ? 1 : 0); r <= Math.min(rg.r2, s.used().r2); r++) for (let c = rg.c1; c <= rg.c2; c++) s.cells.delete(key(r, c)); keep.forEach((row, i) => row.forEach((x, j) => { if (x) s.cells.set(key(rg.r1 + (hdr ? 1 : 0) + i, rg.c1 + j), x); })); s.rev++; return; }
    case 'autofilter': return;
    case 'specialcells': { const t = toNum(a[0]); const u = s.used(); const out = []; for (let r = Math.max(rg.r1, u.r1); r <= Math.min(rg.r2, u.r2); r++) for (let c = Math.max(rg.c1, u.c1); c <= Math.min(rg.c2, u.c2); c++) { const x = s.get(r, c); const blank = !x || x.v == null || x.v === ''; if ((t === 4 && blank) || (t === 2 && x && !x.f && !blank) || (t === -4123 && x && x.f)) out.push({ r, c }); } if (t === 11) return new VRange(m, s, { r1: u.r2, c1: u.c2, r2: u.r2, c2: u.c2 }); if (!out.length) throw new VbaError('Không tìm thấy ô nào (No cells were found)', 0, 1004); const b = { r1: Math.min.apply(null, out.map(p => p.r)), c1: Math.min.apply(null, out.map(p => p.c)), r2: Math.max.apply(null, out.map(p => p.r)), c2: Math.max.apply(null, out.map(p => p.c)) }; const res = new VRange(m, s, b); res.cellsList = out; res.each = () => out.map(p => new VRange(m, s, { r1: p.r, c1: p.c, r2: p.r, c2: p.c })); return res; }
    case 'calculate': m.doc.recalc(true); return;
    case 'show': return;
    case 'name': return '';
    case 'areas': return { __coll: [R] };
    case 'printout': case 'printpreview': return;
    case 'isempty': { const v = s.value(rg.r1, rg.c1); return v == null || v === ''; }
    case 'indentlevel': return m.styleOf(R).indent || 0;
    case 'orientation': return 0; case 'locked': return true; case 'shrinktofit': return false;
    case 'characters': return { __coll: [], text: toStr(m.getValue(R)) };
    case 'hyperlinks': return { __coll: [], add: true };
  }
  throw new VbaError('Range không có thành viên "' + name + '"', 0, 438);
}
function shiftCells(s, axis, at, n) { const moved = new Map(); s.cells.forEach((x, k) => { const r = Math.floor(k / 20000), c = k % 20000; if (axis === 'r') { if (r >= at) { if (n < 0 && r < at - n) return; moved.set(key(r + n, c), x); } else moved.set(k, x); } else { if (c >= at) { if (n < 0 && c < at - n) return; moved.set(key(r, c + n), x); } else moved.set(k, x); } }); s.cells = moved; s.rev++; }
function copyRange(m, src, dst, withFormat, values, formatsOnly) {
  const S = src.s, D = dst.s; m.touch(D); const { shiftFormula } = m.fmod;
  const tr = dst.rows > 1 && dst.rows % src.rows === 0 ? dst.rows : src.rows, tc = dst.cols > 1 && dst.cols % src.cols === 0 ? dst.cols : src.cols;
  const buf = [];
  for (let i = 0; i < tr; i++) for (let j = 0; j < tc; j++) { const sr = src.rg.r1 + (i % src.rows), sc = src.rg.c1 + (j % src.cols); const x = S.get(sr, sc); buf.push([i, j, x ? Object.assign({}, x) : null, sr, sc]); }
  buf.forEach(([i, j, x, sr, sc]) => {
    const r = dst.rg.r1 + i, c = dst.rg.c1 + j;
    if (!x) { if (!formatsOnly) D.setCell(r, c, { v: null, f: null, s: withFormat ? null : undefined }); return; }
    if (formatsOnly) { D.setCell(r, c, { s: x.s || null }); return; }
    const patch = {}; if (values && !withFormat) { patch.v = S.value(sr, sc); patch.f = null; } else if (x.f && !values) { patch.f = shiftFormula(x.f, r - sr, c - sc); patch.v = null; } else { patch.v = x.f ? S.value(sr, sc) : x.v; patch.f = null; }
    if (withFormat) patch.s = x.s || null;
    D.setCell(r, c, patch);
  });
}
function sheetMember(m, o, name, a) {
  const s = o.s, wb = m.wb();
  switch (name) {
    case 'name': case 'codename': return s.name;
    case 'range': return m.range(s, a[0], a[1]);
    case 'cells': if (!a.length) return new VRange(m, s, { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }); return rangeMember(m, new VRange(m, s, { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }), 'cells', a);
    case 'rows': { const all = new VRange(m, s, { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }); if (!a.length) return new VRows(all, false); if (typeof a[0] === 'string') { const p = toStr(a[0]).split(':'); return new VRange(m, s, { r1: toNum(p[0]) - 1, c1: 0, r2: toNum(p[1] || p[0]) - 1, c2: MAX_COLS - 1 }); } return member(m, new VRows(all, false), '', a); }
    case 'columns': { const all = new VRange(m, s, { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }); if (!a.length) return new VRows(all, true); if (typeof a[0] === 'string' && isNaN(+a[0])) { const p = toStr(a[0]).split(':'); return new VRange(m, s, { r1: 0, c1: colIndex(p[0]), r2: MAX_ROWS - 1, c2: colIndex(p[1] || p[0]) }); } return member(m, new VRows(all, true), '', a); }
    case 'usedrange': { const u = s.used(); return new VRange(m, s, u); }
    case 'activate': case 'select': m.doc.wb.active = wb.sheets.indexOf(s); m.sheetChanged = true; return;
    case 'index': return wb.sheets.indexOf(s) + 1;
    case 'visible': return s.hidden ? 0 : -1;
    case 'delete': { if (wb.sheets.length <= 1) throw new VbaError('Sổ làm việc phải có ít nhất một trang tính', 0, 1004); const i = wb.sheets.indexOf(s); wb.sheets.splice(i, 1); wb.active = Math.min(wb.active, wb.sheets.length - 1); m.structural = true; return; }
    case 'protect': s.protected = { pw: a[0] ? toStr(a[0]) : '' }; return; case 'unprotect': s.protected = null; return;
    case 'protectcontents': return !!s.protected;
    case 'calculate': m.doc.recalc(true); return;
    case 'parent': return m.book;
    case 'next': { const i = wb.sheets.indexOf(s); return wb.sheets[i + 1] ? m.sheetObj(wb.sheets[i + 1]) : null; }
    case 'previous': { const i = wb.sheets.indexOf(s); return wb.sheets[i - 1] ? m.sheetObj(wb.sheets[i - 1]) : null; }
    case 'copy': { const named = a.named || {}; const before = named.before || a[0], after = named.after || a[1]; const { wbToJSON, wbFromJSON } = m.mmod; const tmp = { sheets: [s], active: 0, names: {}, calcMode: 'auto', props: {} }; const j = wbToJSON(tmp); const w2 = wbFromJSON(j); const ns = w2.sheets[0]; ns.wb = wb; let nm = s.name + ' (2)'; let k = 2; while (wb.sheet(nm)) nm = s.name + ' (' + (++k) + ')'; ns.name = nm; let at = wb.sheets.length; if (before instanceof VSheet) at = wb.sheets.indexOf(before.s); else if (after instanceof VSheet) at = wb.sheets.indexOf(after.s) + 1; wb.sheets.splice(at, 0, ns); wb.active = at; m.structural = true; return; }
    case 'move': { const named = a.named || {}; const before = named.before || a[0], after = named.after || a[1]; const i = wb.sheets.indexOf(s); wb.sheets.splice(i, 1); let at = wb.sheets.length; if (before instanceof VSheet) at = wb.sheets.indexOf(before.s); else if (after instanceof VSheet) at = wb.sheets.indexOf(after.s) + 1; wb.sheets.splice(at, 0, s); m.structural = true; return; }
    case 'tab': return { __tab: s };
    case 'paste': return; case 'printout': return; case 'shapes': return { __coll: [] }; case 'listobjects': return { __coll: (s.tables || []).map(t => new VRange(m, s, t)) };
    case 'evaluate': return m.evaluate(toStr(a[0]), s);
    case 'standardwidth': return 8.43;
  }
  throw new VbaError('Worksheet không có thành viên "' + name + '"', 0, 438);
}

export function setMember(m, o, name, args, val) {
  if (o instanceof Collection || o instanceof Dictionary) { o.vcall(name, Array.from(args), val, true); return; }
  if (o instanceof VRegExp) { o[name] = name === 'pattern' ? toStr(val) : truthy(val); return; }
  if (o instanceof VRange) {
    if (args && args.length && (name === '' || name === 'cells' || name === 'item' || name === 'value')) { const r = rangeMember(m, o, 'cells', args); m.setValue(r, val); return; }
    const R = o; const s = R.s; m.touch(s);
    switch (name) {
      case '': case 'value': case 'value2': m.setValue(R, val); return;
      case 'formula': case 'formulalocal': case 'formula2': case 'formular1c1': { const v = toStr(m.v(val)); m.setValue(R, v[0] === '=' ? v : v, true); return; }
      case 'numberformat': case 'numberformatlocal': m.style(R, { nf: toStr(val) }); return;
      case 'horizontalalignment': { const v = toNum(val); m.style(R, { ha: v === -4108 || v === 7 ? 'center' : v === -4152 ? 'right' : v === -4131 ? 'left' : null }); return; }
      case 'verticalalignment': { const v = toNum(val); m.style(R, { va: v === -4160 ? 'top' : v === -4108 ? 'middle' : null }); return; }
      case 'wraptext': m.style(R, { wrap: truthy(val) || null }); return;
      case 'indentlevel': m.style(R, { indent: toNum(val) || null }); return;
      case 'columnwidth': { const w = Math.round(toNum(val) * 7 + 5); for (let c = R.rg.c1; c <= Math.min(R.rg.c2, R.rg.c1 + 16383); c++) s.colW[c] = w; if (s._ax) s._ax.c.dirty = true; return; }
      case 'rowheight': { const hgt = Math.round(toNum(val) / 0.75); for (let r = R.rg.r1; r <= Math.min(R.rg.r2, R.rg.r1 + 100000); r++) s.rowH[r] = hgt; if (s._ax) s._ax.r.dirty = true; return; }
      case 'hidden': { const hide = truthy(val); if (R.rows > 50000) { for (let c = R.rg.c1; c <= R.rg.c2; c++) { if (hide) s.hiddenC[c] = true; else delete s.hiddenC[c]; } if (s._ax) s._ax.c.dirty = true; } else { for (let r = R.rg.r1; r <= R.rg.r2; r++) { if (hide) s.hiddenR[r] = true; else delete s.hiddenR[r]; } if (s._ax) s._ax.r.dirty = true; } return; }
      case 'mergecells': if (truthy(val)) rangeMember(m, R, 'merge', []); else rangeMember(m, R, 'unmerge', []); return;
      case 'name': { const wb = m.wb(); wb.names = wb.names || {}; wb.names[toStr(val)] = { ref: "'" + s.name + "'!" + rangeAddr(R.rg) }; return; }
      case 'locked': case 'shrinktofit': case 'orientation': return;
    }
    throw new VbaError('Không thể gán thuộc tính Range.' + name, 0, 438);
  }
  if (o instanceof VRows) { const r = o.rng; const whole = o.isCol ? new VRange(m, r.s, { r1: 0, c1: r.rg.c1, r2: MAX_ROWS - 1, c2: r.rg.c2 }) : new VRange(m, r.s, { r1: r.rg.r1, c1: 0, r2: r.rg.r2, c2: MAX_COLS - 1 }); return setMember(m, whole, name, args, val); }
  if (o instanceof VFont) { const R = o.rng.r; const p = {}; switch (name) { case 'bold': p.b = truthy(val) || null; break; case 'italic': p.i = truthy(val) || null; break; case 'underline': { const v = typeof val === 'boolean' ? (val ? 2 : -4142) : toNum(val); p.u = v === -4142 || v === 0 ? null : (v === -4119 ? 'double' : 'single'); break; } case 'strikethrough': p.s = truthy(val) || null; break; case 'size': p.sz = toNum(val); break; case 'name': p.font = toStr(val); break; case 'color': p.color = longToHex(val); break; case 'colorindex': p.color = COLORINDEX[toNum(val)] || null; break; case 'themecolor': case 'tintandshade': return; default: throw new VbaError('Font không có thuộc tính ' + name, 0, 438); } m.style(R, p); return; }
  if (o instanceof VInterior) { switch (name) { case 'color': m.style(o.rng, { fill: longToHex(val) }); return; case 'colorindex': { const v = toNum(val); m.style(o.rng, { fill: v === -4142 || v === 0 ? null : (COLORINDEX[v] || null) }); return; } case 'pattern': if (toNum(val) === -4142) m.style(o.rng, { fill: null }); return; case 'themecolor': case 'tintandshade': case 'patterncolorindex': return; } throw new VbaError('Interior không có thuộc tính ' + name, 0, 438); }
  if (o instanceof VBorders) {
    const R = o.rng; const rg = R.rg; const w = o.which;
    const get = () => ({ s: 'thin', c: '#000000' });
    if (name === 'linestyle' || name === 'weight' || name === 'color' || name === 'colorindex') {
      const none = name === 'linestyle' && toNum(val) === -4142;
      const st = name === 'weight' ? ({ 1: 'hair', 2: 'thin', [-4138]: 'medium', 4: 'thick' })[toNum(val)] || 'thin' : name === 'linestyle' ? ({ 1: 'thin', [-4115]: 'dashed', [-4118]: 'dotted', [-4119]: 'double' })[toNum(val)] || 'thin' : null;
      const col = name === 'color' ? longToHex(val) : null;
      const lim = m.limitRange(R);
      for (let r = lim.r1; r <= lim.r2; r++) for (let c = lim.c1; c <= lim.c2; c++) {
        const x = R.s.get(r, c); const bd = Object.assign({}, (x && x.s && x.s.bd) || {});
        const sides = []; if (!w) sides.push('t', 'b', 'l', 'r'); else { if (w === 8 && r === rg.r1) sides.push('t'); if (w === 9 && r === rg.r2) sides.push('b'); if (w === 7 && c === rg.c1) sides.push('l'); if (w === 10 && c === rg.c2) sides.push('r'); if (w === 12 && r < rg.r2) sides.push('b'); if (w === 11 && c < rg.c2) sides.push('r'); if (w === 5) sides.push('u'); if (w === 6) sides.push('d'); }
        sides.forEach(sd => { if (none) delete bd[sd]; else bd[sd] = Object.assign(bd[sd] || get(), st ? { s: st } : {}, col ? { c: col } : {}); });
        R.s.setCell(r, c, { s: mergeStyle(x && x.s, { bd: Object.keys(bd).length ? bd : null }) });
      }
      m.touch(R.s); return;
    }
    return;
  }
  if (o instanceof VSheet) { const s = o.s; switch (name) { case 'name': { const nm = toStr(val); if (m.wb().sheets.some(x => x !== s && x.name.toLowerCase() === nm.toLowerCase())) throw new VbaError('Tên trang tính đã tồn tại', 0, 1004); const old = s.name; s.name = nm; m.renamed = m.renamed || []; m.renamed.push([old, nm]); m.structural = true; return; } case 'visible': s.hidden = !truthy(val) || toNum(val) === 2; m.structural = true; return; case 'tab': return; } if (name === 'cells' || name === 'range' || name === '') { const r = member(m, o, name === '' ? 'cells' : name, args); return setMember(m, r, '', [], val); } throw new VbaError('Không thể gán Worksheet.' + name, 0, 438); }
  if (o instanceof VApp) { switch (name) { case 'screenupdating': m.screenUpdating = truthy(val); return; case 'statusbar': m.statusBar = val; if (val && val !== false) m.ui.status(toStr(val)); return; case 'calculation': m.wb().calcMode = toNum(val) === -4135 ? 'manual' : 'auto'; return; case 'displayalerts': case 'enableevents': case 'cutcopymode': case 'displaystatusbar': case 'askToUpdateLinks': case 'calculatebeforesave': case 'cursor': return; } return; }
  if (o && o.__tab) { if (name === 'color') o.__tab.tabColor = longToHex(val); return; }
  if (o instanceof VBook) return;
  throw new VbaError('Không thể gán thuộc tính "' + name + '" (' + m.typeName(o) + ')', 0, 438);
}
