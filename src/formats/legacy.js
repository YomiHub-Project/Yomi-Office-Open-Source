// Yomi-Office Open Source v0.0.6 — read legacy Microsoft Office 97-2003 binary files (.doc, .xls, .ppt) offline.
// [MS-DOC] piece table text, [MS-XLS] BIFF8 cell values (+ SST, RK/MULRK, formulas' cached results, merged cells, sheet names),
// [MS-PPT] slide texts. Formatting of these old formats is simplified; the content (text/numbers) is preserved.
import { readCfb } from '../excel/vba/ovba.js';
import { esc } from '../core/util.js';

const u16 = (b, o) => b[o] | (b[o + 1] << 8);
const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
function stream(files, name) { const k = Object.keys(files).find(p => p.toLowerCase() === name.toLowerCase() || p.toLowerCase().endsWith('/' + name.toLowerCase())); return k ? files[k]() : null; }
const CP1252 = [8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 141, 381, 143, 144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 353, 8250, 339, 157, 382, 376];
const ansi = (b, o, n) => { let s = ''; for (let i = 0; i < n; i++) { const c = b[o + i]; s += String.fromCharCode(c >= 128 && c < 160 ? CP1252[c - 128] : c); } return s; };
const utf16 = (b, o, n) => { let s = ''; for (let i = 0; i < n; i++) s += String.fromCharCode(u16(b, o + 2 * i)); return s; };

// ================================================================ .doc
export function readDoc(bytes) {
  const files = readCfb(bytes); const wd = stream(files, 'WordDocument'); if (!wd) throw new Error('Không phải tệp Word 97-2003');
  const flags = u16(wd, 0x0A); const table = stream(files, (flags & 0x0200) ? '1Table' : '0Table'); if (!table) throw new Error('Thiếu luồng Table');
  if (flags & 0x0100) throw new Error('Tệp .doc được mã hóa (có mật khẩu) — chưa hỗ trợ');
  // FibRgFcLcb97: fcClx/lcbClx at offset 0x01A2 (fcClx) in the FIB
  const fcClx = u32(wd, 0x01A2), lcbClx = u32(wd, 0x01A6);
  let p = fcClx; const end = fcClx + lcbClx; let text = '';
  while (p < end) {
    const t = table[p];
    if (t === 1) { const cb = u16(table, p + 1); p += 3 + cb; continue; }          // Prc (property modifiers)
    if (t !== 2) break;
    const lcb = u32(table, p + 1); const pt = p + 5; const n = (lcb - 4) / 12;
    for (let i = 0; i < n; i++) {
      const cpS = u32(table, pt + 4 * i), cpE = u32(table, pt + 4 * (i + 1)); const pcd = pt + 4 * (n + 1) + 8 * i;
      let fc = u32(table, pcd + 2); const compressed = !!(fc & 0x40000000); fc &= 0x3FFFFFFF; const len = cpE - cpS;
      text += compressed ? ansi(wd, fc / 2, len) : utf16(wd, fc, len);
    }
    break;
  }
  // Word special characters → HTML paragraphs
  const paras = text.replace(/\u0013[^\u0014\u0015]*\u0014?([^\u0015]*)\u0015/g, '$1').replace(/[\u0001-\u0006\u0008\u0013-\u0015]/g, '').split(/[\r\u000c]/);
  // tables: every cell ends with \u0007 and every row ends with an extra paragraph holding a single \u0007 (TTP mark)
  let html = ''; let rows = null, row = [], cell = '';
  const flushTable = () => { if (!rows) return; if (row.length) rows.push(row); html += '<table class="wd-table">' + rows.map(r => '<tr>' + r.map(c => '<td>' + (esc(c).replace(/\n/g, '<br>') || '<br>') + '</td>').join('') + '</tr>').join('') + '</table>'; rows = null; row = []; };
  for (const raw of paras) {
    if (raw.indexOf('\u0007') >= 0) {
      if (!rows) rows = [];
      let prevBell = rows._prevBell || false;
      for (const ch of raw) { if (ch === '\u0007') { if (prevBell && cell === '') { rows.push(row); row = []; prevBell = false; continue; } row.push(cell); cell = ''; prevBell = true; } else { cell += ch; prevBell = false; } }
      rows._prevBell = prevBell; if (cell) { cell += '\n'; }
      continue;
    }
    if (rows && cell) { cell += raw + '\n'; continue; }
    flushTable();
    const t = raw.replace(/\u000b/g, '\n');
    html += '<p>' + (t ? esc(t).replace(/\n/g, '<br>') : '<br>') + '</p>';
  }
  flushTable();
  return { html };
}

// ================================================================ .xls (BIFF8)
export function readXls(wb, bytes, helpers) {
  const files = readCfb(bytes); const b = stream(files, 'Workbook') || stream(files, 'Book'); if (!b) throw new Error('Không phải tệp Excel 97-2003');
  const sst = []; const sheets = []; let cur = null; const fmts = {}; const xfFmt = []; let p = 0; let date1904 = false;
  const rk = v => { let n; if (v & 2) n = v >> 2; else { const buf = new DataView(new ArrayBuffer(8)); buf.setUint32(4, v & 0xFFFFFFFC, true); buf.setUint32(0, 0, true); n = buf.getFloat64(0, true); } return v & 1 ? n / 100 : n; };
  const f64 = o => new DataView(b.buffer, b.byteOffset + o, 8).getFloat64(0, true);
  // XLUnicodeRichExtendedString reader that can continue across CONTINUE records
  let recs = []; while (p + 4 <= b.length) { const id = u16(b, p), len = u16(b, p + 2); recs.push({ id, o: p + 4, len }); p += 4 + len; }
  const readSST = (start) => {
    // gather SST + CONTINUE payloads with boundaries
    const parts = [recs[start]]; for (let i = start + 1; i < recs.length && recs[i].id === 0x003C; i++) parts.push(recs[i]);
    let pi = 0, off = parts[0].o + 8; const total = u32(b, parts[0].o + 4);
    const ensure = () => { if (off >= parts[pi].o + parts[pi].len && pi + 1 < parts.length) { pi++; off = parts[pi].o; } };
    for (let k = 0; k < total; k++) {
      ensure(); const cch = u16(b, off); off += 2; let fl = b[off++]; let rt = 0, sz = 0; if (fl & 8) { rt = u16(b, off); off += 2; } if (fl & 4) { sz = u32(b, off); off += 4; }
      let s = ''; let left = cch;
      while (left > 0) { ensure(); const avail = parts[pi].o + parts[pi].len - off; const wide = fl & 1; const can = Math.min(left, wide ? avail >> 1 : avail); s += wide ? utf16(b, off, can) : ansi(b, off, can); off += can * (wide ? 2 : 1); left -= can; if (left > 0) { pi++; off = parts[pi].o; fl = b[off++]; } }
      off += rt * 4 + sz; while (pi + 1 < parts.length && off >= parts[pi].o + parts[pi].len) { const over = off - (parts[pi].o + parts[pi].len); pi++; off = parts[pi].o + over; }
      sst.push(s);
    }
  };
  const bounds = [];
  recs.forEach((r, i) => {
    if (r.id === 0x0085) { const pos = u32(b, r.o); const vis = b[r.o + 4]; const type = b[r.o + 5]; const cch = b[r.o + 6]; const wide = b[r.o + 7] & 1; const name = wide ? utf16(b, r.o + 8, cch) : ansi(b, r.o + 8, cch); if (type === 0) bounds.push({ pos, name, hidden: vis !== 0 }); }
    else if (r.id === 0x00FC) readSST(i);
    else if (r.id === 0x0022) date1904 = u16(b, r.o) === 1;
    else if (r.id === 0x041E) { const ix = u16(b, r.o); const cch = u16(b, r.o + 2); const wide = b[r.o + 4] & 1; fmts[ix] = wide ? utf16(b, r.o + 5, cch) : ansi(b, r.o + 5, cch); }
    else if (r.id === 0x00E0) xfFmt.push(u16(b, r.o + 2));
  });
  const BUILTIN = { 14: 'dd/mm/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 22: 'dd/mm/yyyy hh:mm', 9: '0%', 10: '0.00%', 3: '#,##0', 4: '#,##0.00', 2: '0.00', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'hh:mm', 21: 'hh:mm:ss' };
  const nfOf = xf => { const id = xfFmt[xf]; return id == null ? null : (fmts[id] || BUILTIN[id] || null); };
  bounds.forEach(bd => {
    const s = wb.addSheet(bd.name); if (bd.hidden) s.hidden = true;
    let i = recs.findIndex(r => r.o - 4 === bd.pos); if (i < 0) return;
    const set = (r, c, v, xf) => { const patch = { v }; const nf = nfOf(xf); if (nf && typeof v === 'number') patch.s = helpers.internStyle({ nf }); s.setCell(r, c, patch); };
    for (i = i + 1; i < recs.length; i++) {
      const r = recs[i]; if (r.id === 0x000A) break;
      const row = u16(b, r.o), col = u16(b, r.o + 2), xf = u16(b, r.o + 4);
      switch (r.id) {
        case 0x00FD: set(row, col, sst[u32(b, r.o + 6)], xf); break;                         // LABELSST
        case 0x0203: set(row, col, f64(r.o + 6), xf); break;                                 // NUMBER
        case 0x027E: set(row, col, rk(u32(b, r.o + 6)), xf); break;                          // RK
        case 0x00BD: { const last = u16(b, r.o + r.len - 2); for (let c = col, q = r.o + 4; c <= last; c++, q += 6) set(row, c, rk(u32(b, q + 2)), u16(b, q)); break; } // MULRK
        case 0x0205: { const v = b[r.o + 6], isErr = b[r.o + 7]; set(row, col, isErr ? null : !!v, xf); break; } // BOOLERR
        case 0x0204: { const cch = u16(b, r.o + 6); set(row, col, ansi(b, r.o + 8, cch), xf); break; } // LABEL (BIFF5)
        case 0x0006: { const lo = u32(b, r.o + 6), hi = u32(b, r.o + 10); if ((hi >>> 16) === 0xFFFF) { const t = lo & 0xFF; if (t === 1) set(row, col, !!((lo >> 16) & 0xFF), xf); else if (t === 0) { const nx = recs[i + 1]; if (nx && nx.id === 0x0207) { const cch = u16(b, nx.o); const wide = b[nx.o + 2] & 1; set(row, col, wide ? utf16(b, nx.o + 3, cch) : ansi(b, nx.o + 3, cch), xf); } } } else set(row, col, f64(r.o + 6), xf); break; } // FORMULA (cached value)
        case 0x00E5: { const n = u16(b, r.o); for (let k = 0; k < n; k++) { const q = r.o + 2 + 8 * k; s.merges.push({ r1: u16(b, q), r2: u16(b, q + 2), c1: u16(b, q + 4), c2: u16(b, q + 6) }); } break; } // MERGECELLS
        case 0x007D: { const c1 = u16(b, r.o), c2 = u16(b, r.o + 2), w = u16(b, r.o + 4); for (let c = c1; c <= Math.min(c2, c1 + 255); c++) s.colW[c] = Math.round(w / 256 * 7 + 5); break; } // COLINFO
      }
    }
  });
  return wb;
}

// ================================================================ .ppt
export function readPpt(bytes) {
  const files = readCfb(bytes); const d = stream(files, 'PowerPoint Document'); if (!d) throw new Error('Không phải tệp PowerPoint 97-2003');
  const slides = []; let cur = null;
  // pass 1: SlideListWithText (RT 0x0FF0, instance 0) — SlidePersistAtom (0x03F3) starts a slide, text atoms follow it
  const walk = (o, end, depth, inList) => {
    while (o + 8 <= end) {
      const verInst = u16(d, o), type = u16(d, o + 2), len = u32(d, o + 4); const ver = verInst & 0xF, inst = verInst >> 4; const body = o + 8;
      if (type === 0x0FF0) { walk(body, body + len, depth + 1, inst === 0); o = body + len; continue; }
      if (inList && type === 0x03F3) { cur = { texts: [] }; slides.push(cur); }
      if (inList && cur && type === 0x0FA0) cur.texts.push(utf16(d, body, len / 2));
      if (inList && cur && type === 0x0FA8) cur.texts.push(ansi(d, body, len));
      if (ver === 0xF && depth < 6 && type !== 0x0FF0) walk(body, body + len, depth + 1, inList);
      o = body + len;
    }
  };
  walk(0, d.length, 0, false);
  // pass 2 (files written by newer tools keep text inside each Slide container 0x03EE)
  if (!slides.some(s => s.texts.length)) {
    slides.length = 0;
    const walk2 = (o, end, depth, sl) => { while (o + 8 <= end) { const verInst = u16(d, o), type = u16(d, o + 2), len = u32(d, o + 4); const body = o + 8; let s2 = sl; if (type === 0x03EE) { s2 = { texts: [] }; slides.push(s2); } if (s2 && type === 0x0FA0) s2.texts.push(utf16(d, body, len / 2)); if (s2 && type === 0x0FA8) s2.texts.push(ansi(d, body, len)); if ((verInst & 0xF) === 0xF && depth < 10) walk2(body, body + len, depth + 1, s2); o = body + len; } };
    walk2(0, d.length, 0, null);
  }
  return slides.filter(s => s.texts.length).map(s => s.texts.map(t => t.replace(/\r/g, '\n').trim()).filter(Boolean));
}
