// Spreadsheet data model: Workbook / Sheet / Cell, addressing helpers, style interning
export const MAX_ROWS = 1048576, MAX_COLS = 16384;
export const DEF_COL_W = 64, DEF_ROW_H = 20;

export function colName(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; }
export function colIndex(name) { let c = 0; name = name.toUpperCase(); for (let i = 0; i < name.length; i++) c = c * 26 + (name.charCodeAt(i) - 64); return c - 1; }
export function addr(r, c, absR, absC) { return (absC ? '$' : '') + colName(c) + (absR ? '$' : '') + (r + 1); }
export function parseAddr(s) { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s.trim()); if (!m) return null; return { r: +m[2] - 1, c: colIndex(m[1]) }; }
export function rangeAddr(r) { if (!r) return ''; const a = addr(r.r1, r.c1), b = addr(r.r2, r.c2); return a === b ? a : a + ':' + b; }
export function parseRangeAddr(s) {
  s = String(s || '').trim().replace(/^.*!/, '');
  let m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(s);
  if (m) { const r1 = +m[2] - 1, c1 = colIndex(m[1]); const r2 = m[3] ? +m[4] - 1 : r1, c2 = m[3] ? colIndex(m[3]) : c1; return norm({ r1, c1, r2, c2 }); }
  m = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(s); if (m) return norm({ r1: 0, c1: colIndex(m[1]), r2: MAX_ROWS - 1, c2: colIndex(m[2]) });
  m = /^\$?(\d+):\$?(\d+)$/.exec(s); if (m) return norm({ r1: +m[1] - 1, c1: 0, r2: +m[2] - 1, c2: MAX_COLS - 1 });
  return null;
}
export function norm(r) { return { r1: Math.min(r.r1, r.r2), c1: Math.min(r.c1, r.c2), r2: Math.max(r.r1, r.r2), c2: Math.max(r.c1, r.c2) }; }
export function inRange(rg, r, c) { return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2; }
export function key(r, c) { return r * 20000 + c; }   // numeric key (MAX_COLS < 20000)
export function unkey(k) { return { r: Math.floor(k / 20000), c: k % 20000 }; }

// ---- errors
export class Err { constructor(code) { this.code = code; } toString() { return this.code; } }
export const E = { DIV0: new Err('#DIV/0!'), NA: new Err('#N/A'), NAME: new Err('#NAME?'), NULL: new Err('#NULL!'), NUM: new Err('#NUM!'), REF: new Err('#REF!'), VALUE: new Err('#VALUE!'), SPILL: new Err('#SPILL!'), CALC: new Err('#CALC!'), GETTING: new Err('#GETTING_DATA') };
export function errOf(code) { for (const k in E) if (E[k].code === code) return E[k]; return new Err(code); }
export function isErr(v) { return v instanceof Err; }

// ---- styles (interned, immutable)
const styleCache = {};
export function internStyle(s) {
  if (!s) return null;
  const clean = {}; let n = 0;
  Object.keys(s).sort().forEach(k => { const v = s[k]; if (v === undefined || v === null || v === false || v === '') return; clean[k] = v; n++; });
  if (!n) return null;
  const j = JSON.stringify(clean);
  return styleCache[j] || (styleCache[j] = Object.freeze(clean));
}
export function mergeStyle(base, patch) {
  const o = Object.assign({}, base || {});
  for (const k in patch) { if (patch[k] === null || patch[k] === undefined) delete o[k]; else if (k === 'bd' && patch.bd !== null) o.bd = Object.assign({}, o.bd || {}, patch.bd); else o[k] = patch[k]; }
  if (o.bd) { for (const s of ['t', 'b', 'l', 'r', 'd', 'u']) if (o.bd[s] === null) delete o.bd[s]; if (!Object.keys(o.bd).length) delete o.bd; }
  return internStyle(o);
}

// ---- Sheet
export class Sheet {
  constructor(wb, name) {
    this.wb = wb; this.name = name;
    this.cells = new Map();          // key -> {v, f, s, note, link, arr}
    this.colW = {}; this.rowH = {}; this.hiddenR = {}; this.hiddenC = {};
    this.colS = {}; this.rowS = {};  // whole column/row styles
    this.merges = []; this.freeze = { r: 0, c: 0 };
    this.cf = []; this.dv = []; this.filter = null; this.tables = [];
    this.objects = [];               // charts, images, shapes, textboxes, equations
    this.ink = [];                   // Draw tab strokes
    this.tabColor = null; this.hidden = false; this.protected = null;
    this.showGrid = true; this.showHead = true; this.zoom = 100;
    this.sel = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], r: 0, c: 0 };
    this.scroll = { x: 0, y: 0 };
    this.outlineR = {}; this.outlineC = {}; // group levels
    this.pageSetup = { orientation: 'portrait', size: 'A4', margins: 'normal', fitWidth: 0, printArea: null, printTitles: null, gridlines: false, headings: false };
    this.spill = new Map();          // key -> {src:key, v}
    this.rev = 0;
  }
  get(r, c) { return this.cells.get(key(r, c)); }
  ensure(r, c) { const k = key(r, c); let x = this.cells.get(k); if (!x) { x = {}; this.cells.set(k, x); } return x; }
  value(r, c) { const x = this.cells.get(key(r, c)); if (x) return x.v === undefined ? null : x.v; const sp = this.spill.get(key(r, c)); return sp ? sp.v : null; }
  styleAt(r, c) { const x = this.cells.get(key(r, c)); return (x && x.s) || this.rowS[r] || this.colS[c] || null; }
  colWidth(c) { return this.hiddenC[c] ? 0 : (this.colW[c] != null ? this.colW[c] : DEF_COL_W); }
  rowHeight(r) { return this.hiddenR[r] ? 0 : (this.rowH[r] != null ? this.rowH[r] : DEF_ROW_H); }
  // bounding box of used cells
  used() {
    let r2 = -1, c2 = -1, r1 = Infinity, c1 = Infinity;
    this.cells.forEach((x, k) => { if (x.v == null && !x.f && !x.s) return; const r = Math.floor(k / 20000), c = k % 20000; if (r > r2) r2 = r; if (c > c2) c2 = c; if (r < r1) r1 = r; if (c < c1) c1 = c; });
    this.merges.forEach(m => { if (m.r2 > r2) r2 = m.r2; if (m.c2 > c2) c2 = m.c2; });
    if (r2 < 0) return { r1: 0, c1: 0, r2: 0, c2: 0, empty: true };
    return { r1, c1, r2, c2 };
  }
  mergeAt(r, c) { for (const m of this.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m; return null; }
  setCell(r, c, patch) {
    const k = key(r, c); let x = this.cells.get(k);
    if (!x) { if (patch.v == null && !patch.f && !patch.s && !patch.note && !patch.link) return; x = {}; this.cells.set(k, x); }
    for (const p in patch) { if (patch[p] === undefined || patch[p] === null || patch[p] === '') delete x[p]; else x[p] = patch[p]; }
    if (x.v == null && !x.f && !x.s && !x.note && !x.link) this.cells.delete(k);
    this.rev++;
  }
}

export class Workbook {
  constructor() { this.sheets = []; this.active = 0; this.names = {}; this.calcMode = 'auto'; this.rev = 0; this.props = {}; }
  addSheet(name, at) {
    name = name || this.nextName();
    const s = new Sheet(this, name);
    if (at == null) this.sheets.push(s); else this.sheets.splice(at, 0, s);
    return s;
  }
  nextName() { let i = this.sheets.length + 1; while (this.sheet('Sheet' + i)) i++; return 'Sheet' + i; }
  sheet(name) { if (typeof name === 'number') return this.sheets[name]; const n = String(name).toLowerCase(); return this.sheets.find(s => s.name.toLowerCase() === n) || null; }
  get cur() { return this.sheets[this.active]; }
}

// ---- serialization (native snapshot for autosave)
export function wbToJSON(wb) {
  return {
    v: 6, active: wb.active, names: wb.names, calcMode: wb.calcMode, props: wb.props, vba: wb.vba || null, codeName: wb.codeName || null,
    sheets: wb.sheets.map(s => {
      const cells = [];
      s.cells.forEach((x, k) => cells.push([k, x.v instanceof Err ? { e: x.v.code } : x.v, x.f || 0, x.s || 0, x.note || 0, x.link || 0]));
      return { name: s.name, codeName: s.codeName || null, cells, colW: s.colW, rowH: s.rowH, hiddenR: s.hiddenR, hiddenC: s.hiddenC, colS: s.colS, rowS: s.rowS, merges: s.merges, freeze: s.freeze, cf: s.cf, dv: s.dv, filter: s.filter, tables: s.tables, objects: s.objects, ink: s.ink, tabColor: s.tabColor, hidden: s.hidden, protected: s.protected, showGrid: s.showGrid, showHead: s.showHead, zoom: s.zoom, pageSetup: s.pageSetup, outlineR: s.outlineR, outlineC: s.outlineC };
    })
  };
}
export function wbFromJSON(j) {
  const wb = new Workbook();
  wb.names = j.names || {}; wb.calcMode = j.calcMode || 'auto'; wb.props = j.props || {}; if (j.vba) wb.vba = j.vba; if (j.codeName) wb.codeName = j.codeName;
  (j.sheets || []).forEach(o => {
    const s = wb.addSheet(o.name);
    (o.cells || []).forEach(a => { const x = {}; if (a[1] != null) x.v = a[1] && a[1].e ? errOf(a[1].e) : a[1]; if (a[2]) x.f = a[2]; if (a[3]) x.s = internStyle(a[3]); if (a[4]) x.note = a[4]; if (a[5]) x.link = a[5]; s.cells.set(a[0], x); });
    if (o.codeName) s.codeName = o.codeName;
    ['colW', 'rowH', 'hiddenR', 'hiddenC', 'merges', 'freeze', 'cf', 'dv', 'filter', 'tables', 'objects', 'ink', 'tabColor', 'hidden', 'protected', 'showGrid', 'showHead', 'zoom', 'pageSetup', 'outlineR', 'outlineC'].forEach(k => { if (o[k] !== undefined && o[k] !== null) s[k] = o[k]; });
    for (const k in (o.colS || {})) s.colS[k] = internStyle(o.colS[k]);
    for (const k in (o.rowS || {})) s.rowS[k] = internStyle(o.rowS[k]);
  });
  if (!wb.sheets.length) wb.addSheet();
  wb.active = Math.min(j.active || 0, wb.sheets.length - 1);
  return wb;
}

// ---- dates (Excel 1900 system)
const EPOCH = Date.UTC(1899, 11, 30);
export function dateToSerial(y, m, d) { return (Date.UTC(y, m - 1, d) - EPOCH) / 864e5; }
export function jsDateToSerial(dt) { return (Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate(), dt.getHours(), dt.getMinutes(), dt.getSeconds(), dt.getMilliseconds()) - EPOCH) / 864e5; }
export function serialToParts(n) {
  const ms = Math.round(n * 864e5) + EPOCH; const d = new Date(ms);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), H: d.getUTCHours(), M: d.getUTCMinutes(), S: d.getUTCSeconds(), ms: d.getUTCMilliseconds(), dow: d.getUTCDay() };
}
export function serialToJsDate(n) { const p = serialToParts(n); return new Date(p.y, p.m - 1, p.d, p.H, p.M, p.S); }
