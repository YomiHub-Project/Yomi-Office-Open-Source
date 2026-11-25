// XLSX / CSV import & export (Office Open XML SpreadsheetML), HTML rendering of sheets
import { JSZipLib, utf8Decode, utf8Encode, esc, extOf, b64FromBytes, bytesToDataUrl, dataUrlToBytes, baseName } from '../core/util.js';
import { Err, errOf, key, colName, colIndex, addr, rangeAddr, parseRangeAddr, parseAddr, internStyle, MAX_ROWS, MAX_COLS, DEF_COL_W, DEF_ROW_H } from './model.js';
import { shiftFormula, FN } from './formula.js';
import { format, parseInput, isDateFormat } from './numfmt.js';

const BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??', 14: 'dd/mm/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'dd/mm/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@' };
const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
const DEF_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

function parseXml(text) { return new DOMParser().parseFromString(text, 'application/xml'); }
function kids(el, name) { const out = []; if (!el) return out; for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (n.localName === name || n.nodeName === name)) out.push(n); return out; }
function kid(el, name) { if (!el) return null; for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (n.localName === name || n.nodeName === name)) return n; return null; }
function all(el, name) { if (!el) return []; const l = el.getElementsByTagNameNS ? el.getElementsByTagNameNS('*', name) : el.getElementsByTagName(name); return Array.prototype.slice.call(l); }
function attr(el, n) { if (!el) return null; const v = el.getAttribute(n); return v == null ? null : v; }
function tint(hex, t) {
  if (!t) return hex;
  let r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
  // HLS tint like Excel
  const f = c => t < 0 ? c * (1 + t) : c * (1 - t) + 255 * t;
  r = Math.round(f(r)); g = Math.round(f(g)); b = Math.round(f(b));
  return [r, g, b].map(x => ('0' + Math.max(0, Math.min(255, x)).toString(16)).slice(-2)).join('').toUpperCase();
}
function colorOf(el, theme) {
  if (!el) return null;
  if (attr(el, 'auto') === '1') return null;
  let hex = null;
  const rgb = attr(el, 'rgb'); if (rgb) hex = rgb.length === 8 ? rgb.slice(2) : rgb;
  const th = attr(el, 'theme'); if (th != null && !hex) hex = (theme || DEF_THEME)[+th] || null;
  const ix = attr(el, 'indexed'); if (ix != null && !hex) { if (+ix === 64 || +ix === 65) return null; hex = INDEXED[+ix] || null; }
  if (!hex) return null;
  const t = attr(el, 'tint'); if (t) hex = tint(hex, +t);
  return '#' + hex.toUpperCase();
}

// ================================================================ READ
export async function readXlsx(wb, bytes) {
  const JSZip = await JSZipLib();
  const zip = await JSZip.loadAsync(bytes);
  const text = async p => { const f = zip.file(p.replace(/^\//, '')); return f ? f.async('string') : null; };
  const rels = async p => { const d = p.replace(/[^/]*$/, ''), n = p.replace(/^.*\//, ''); const t = await text(d + '_rels/' + n + '.rels'); const m = {}; if (t) all(parseXml(t), 'Relationship').forEach(r => { let tg = attr(r, 'Target'); if (attr(r, 'TargetMode') !== 'External') tg = resolvePath(d, tg); m[attr(r, 'Id')] = { target: tg, type: attr(r, 'Type'), external: attr(r, 'TargetMode') === 'External' }; }); return m; };
  const wbXml = parseXml(await text('xl/workbook.xml'));
  const wbRels = await rels('xl/workbook.xml');
  // theme colors
  let theme = DEF_THEME.slice();
  const themeRel = Object.keys(wbRels).map(k => wbRels[k]).find(r => /theme$/.test(r.type));
  if (themeRel) { const tt = await text(themeRel.target); if (tt) { const tx = parseXml(tt); const cs = all(tx, 'clrScheme')[0]; if (cs) { const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']; theme = order.map((n, i) => { const e = kid(cs, n); if (!e) return DEF_THEME[i]; const c = e.firstElementChild || kid(e, 'srgbClr') || kid(e, 'sysClr'); return c ? (attr(c, 'val') && c.localName === 'srgbClr' ? attr(c, 'val') : attr(c, 'lastClr') || DEF_THEME[i]) : DEF_THEME[i]; }); } } }
  // styles
  const styles = parseStyles(await text('xl/styles.xml'), theme);
  // shared strings
  const sst = [];
  const sstT = await text('xl/sharedStrings.xml');
  if (sstT) all(parseXml(sstT), 'si').forEach(si => { const rs = kids(si, 'r'); sst.push(rs.length ? rs.map(r => textOf(kid(r, 't'))).join('') : textOf(kid(si, 't'))); });
  // defined names
  all(wbXml, 'definedName').forEach(dn => { const n = attr(dn, 'name'); if (/^_xlnm\./.test(n)) return; wb.names[n.toUpperCase()] = { f: dn.textContent.replace(/^=/, ''), display: n }; });

  const sheetsEl = all(wbXml, 'sheet');
  const act = all(wbXml, 'workbookView')[0];
  let targetActiveIdx = (act && attr(act, 'activeTab')) ? +attr(act, 'activeTab') : 0;
  if (targetActiveIdx < 0 || targetActiveIdx >= sheetsEl.length) targetActiveIdx = 0;
  const activeSheetName = (sheetsEl[targetActiveIdx] && attr(sheetsEl[targetActiveIdx], 'name')) || (sheetsEl[0] && attr(sheetsEl[0], 'name'));

  for (let idx = 0; idx < sheetsEl.length; idx++) {
    const se = sheetsEl[idx];
    const sName = attr(se, 'name') || ('Sheet' + (idx + 1));
    const s = wb.addSheet(sName);
    const state = attr(se, 'state');
    if (state === 'hidden' || state === 'veryHidden') s.hidden = true;
    const rid = attr(se, 'r:id') || (se.getAttributeNS && se.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')) || se.getAttribute('id');
    const rel = rid ? wbRels[rid] : null;
    if (!rel) continue;

    const sheetPath = rel.target;
    const loadThisSheet = async () => {
      const xt = await text(sheetPath);
      if (!xt) return;
      const cn = /<sheetPr[^>]*codeName="([^"]+)"/.exec(xt);
      if (cn) s.codeName = cn[1];
      const sx = parseXml(xt);
      const srels = await rels(sheetPath);
      readSheet(s, sx, styles, sst, srels, theme);
      for (const id in srels) {
        const r = srels[id];
        if (/comments$/.test(r.type)) {
          const ct = await text(r.target);
          if (ct) {
            const cx = parseXml(ct);
            const authors = all(cx, 'author').map(a => a.textContent);
            all(cx, 'comment').forEach(cm => {
              const a = parseAddr(attr(cm, 'ref'));
              if (!a) return;
              const t = all(cm, 't').map(x => x.textContent).join('');
              const au = authors[+attr(cm, 'authorId')] || '';
              s.setCell(a.r, a.c, { note: { text: t.replace(new RegExp('^' + au.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ':\\s*'), ''), author: au } });
            });
          }
        }
        if (/drawing$/.test(r.type)) await readDrawing(zip, s, r.target, text, rels);
      }
    };

    const isPrioritySheet = (sheetsEl.length <= 5) || (sName === activeSheetName) || (idx < 2 && !s.hidden);
    if (isPrioritySheet) {
      await loadThisSheet();
    } else {
      s._lazyLoad = loadThisSheet;
    }

    all(wbXml, 'definedName').forEach(dn => { if (attr(dn, 'name') === '_xlnm.Print_Area' && +attr(dn, 'localSheetId') === idx) s.pageSetup.printArea = dn.textContent.replace(/^.*!/, '').replace(/\$/g, ''); });
    all(wbXml, 'definedName').forEach(dn => { if (attr(dn, 'name') === '_xlnm.Print_Titles' && +attr(dn, 'localSheetId') === idx) dn.textContent.split(',').forEach(part => { const r = part.replace(/^.*!/, '').replace(/\$/g, ''); if (/^\d+:\d+$/.test(r)) s.pageSetup.printTitles = r; else if (/^[A-Z]+:[A-Z]+$/i.test(r)) s.pageSetup.printTitlesCols = r.toUpperCase(); }); });
  }

  wb.active = Math.max(0, wb.sheets.findIndex(s => s.name === activeSheetName && !s.hidden));
  if (wb.active < 0 || wb.active >= wb.sheets.length) {
    wb.active = Math.max(0, wb.sheets.findIndex(s => !s.hidden));
  }
  if (wb.active < 0 || wb.active >= wb.sheets.length) wb.active = 0;

  // v0.0.6: VBA macros (.xlsm / .xlsb-free projects) — read the source of every module from xl/vbaProject.bin
  const vbaRel = Object.keys(wbRels).map(k => wbRels[k]).find(r => /vbaProject$/.test(r.type));
  const vbaFile = zip.file(vbaRel ? vbaRel.target.replace(/^\//, '') : 'xl/vbaProject.bin');
  if (vbaFile) { try { const { readVbaProject } = await import('./vba/ovba.js'); const pj = readVbaProject(await vbaFile.async('uint8array')); const { vbaToJsa } = await import('./vba/jsa.js'); pj.modules.forEach(m => { const js = vbaToJsa(m.code); if (js != null) { m.type = 'jsa'; m.code = js; } }); { const sf = zip.file('oos/vba-source.json'); if (sf) { const { applySidecar } = await import('./vba/ovba.js'); pj.modules = applySidecar(pj.modules, await sf.async('string')); } } pj.modules.forEach(m => { delete m.rawCode; }); wb.vba = { modules: pj.modules, shortcuts: {}, codepage: pj.codepage }; } catch (e) { console.warn('VBA', e); wb.vba = { modules: [], shortcuts: {}, error: String(e.message || e) }; } }
  { const wp = /<workbookPr[^>]*codeName="([^"]+)"/.exec(await text('xl/workbook.xml') || ''); if (wp) wb.codeName = wp[1]; }
  return wb;
}
function resolvePath(base, t) {
  if (/^\//.test(t)) return t.slice(1);
  const raw = base.endsWith('/') ? base + t : base + '/' + t;
  const parts = raw.split('/');
  const out = [];
  parts.forEach(p => {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  });
  return out.join('/');
}
function textOf(t) { return t ? t.textContent.replace(/_x([0-9A-F]{4})_/g, (m, hx) => String.fromCharCode(parseInt(hx, 16))) : ''; }
function parseStyles(xml, theme) {
  const out = { xfs: [], dxfs: [] };
  if (!xml) return out;
  const x = parseXml(xml);
  const numFmts = {}; all(x, 'numFmt').forEach(n => { numFmts[attr(n, 'numFmtId')] = attr(n, 'formatCode'); });
  const fonts = all(kid(x.documentElement, 'fonts'), 'font').map(f => {
    const o = {}; if (kid(f, 'b') && attr(kid(f, 'b'), 'val') !== '0') o.b = true; if (kid(f, 'i') && attr(kid(f, 'i'), 'val') !== '0') o.i = true;
    const u = kid(f, 'u'); if (u) o.u = attr(u, 'val') === 'double' ? 'double' : 'single'; if (kid(f, 'strike') && attr(kid(f, 'strike'), 'val') !== '0') o.s = true;
    const sz = kid(f, 'sz'); if (sz) o.sz = +attr(sz, 'val'); const nm = kid(f, 'name'); if (nm) o.font = attr(nm, 'val');
    const c = colorOf(kid(f, 'color'), theme); if (c && c !== '#000000') o.color = c;
    const va = kid(f, 'vertAlign'); if (va) o.va2 = attr(va, 'val');
    return o;
  });
  const fills = all(kid(x.documentElement, 'fills'), 'fill').map(f => { const p = kid(f, 'patternFill'); if (p && attr(p, 'patternType') && attr(p, 'patternType') !== 'none') return colorOf(kid(p, 'fgColor'), theme) || colorOf(kid(p, 'bgColor'), theme); const g = kid(f, 'gradientFill'); if (g) { const st = all(g, 'stop')[0]; return st ? colorOf(kid(st, 'color'), theme) : null; } return null; });
  const borders = all(kid(x.documentElement, 'borders'), 'border').map(b => {
    const o = {}; const side = (n, k) => { const e = kid(b, n); if (e && attr(e, 'style')) o[k] = { s: attr(e, 'style'), c: colorOf(kid(e, 'color'), theme) || '#000000' }; };
    side('left', 'l'); side('right', 'r'); side('top', 't'); side('bottom', 'b');
    const d = kid(b, 'diagonal'); if (d && attr(d, 'style')) { const v = { s: attr(d, 'style'), c: colorOf(kid(d, 'color'), theme) || '#000000' }; if (attr(b, 'diagonalDown') === '1') o.u = v; if (attr(b, 'diagonalUp') === '1') o.d = v; }
    return Object.keys(o).length ? o : null;
  });
  const cellXfs = kid(x.documentElement, 'cellXfs');
  out.xfs = kids(cellXfs, 'xf').map(xf => {
    const st = {};
    const nfId = attr(xf, 'numFmtId'); if (nfId && nfId !== '0') st.nf = numFmts[nfId] || BUILTIN_FMT[nfId] || null;
    const fo = fonts[+attr(xf, 'fontId') || 0]; if (fo) Object.assign(st, fo);
    const fi = fills[+attr(xf, 'fillId') || 0]; if (fi) st.fill = fi;
    const bo = borders[+attr(xf, 'borderId') || 0]; if (bo) st.bd = bo;
    const al = kid(xf, 'alignment');
    if (al) { const hz = attr(al, 'horizontal'); if (hz && hz !== 'general') st.ha = hz === 'centerContinuous' ? 'centerAcross' : hz === 'distributed' ? 'justify' : hz; const vt = attr(al, 'vertical'); if (vt) st.va = vt === 'center' ? 'middle' : vt === 'bottom' ? null : vt; if (attr(al, 'wrapText') === '1') st.wrap = true; if (attr(al, 'indent')) st.indent = +attr(al, 'indent'); if (attr(al, 'textRotation')) st.rot = +attr(al, 'textRotation'); if (attr(al, 'shrinkToFit') === '1') st.shrink = true; }
    const pr = kid(xf, 'protection'); if (pr) { if (attr(pr, 'locked') === '0') st.unlocked = true; if (attr(pr, 'hidden') === '1') st.hideF = true; }
    if (st.sz === 11) delete st.sz;
    if (st.font === 'Calibri') delete st.font;
    delete st.va2;
    return internStyle(st);
  });
  out.dxfs = all(kid(x.documentElement, 'dxfs'), 'dxf').map(d => { const o = {}; const f = kid(d, 'font'); if (f) { const c = colorOf(kid(f, 'color'), theme); if (c) o.color = c; if (kid(f, 'b')) o.b = true; } const fl = kid(d, 'fill'); if (fl) { const p = kid(fl, 'patternFill'); if (p) o.fill = colorOf(kid(p, 'bgColor'), theme) || colorOf(kid(p, 'fgColor'), theme); } return o; });
  out.defaultFont = fonts[0] || {};
  return out;
}
function readSheet(s, x, styles, sst, srels, theme) {
  const root = x.documentElement;
  // views
  const sv = all(root, 'sheetView')[0];
  if (sv) {
    if (attr(sv, 'showGridLines') === '0') s.showGrid = false; if (attr(sv, 'showRowColHeaders') === '0') s.showHead = false;
    if (attr(sv, 'zoomScale')) s.zoom = +attr(sv, 'zoomScale');
    const pane = kid(sv, 'pane'); if (pane && (attr(pane, 'state') === 'frozen' || attr(pane, 'state') === 'frozenSplit')) s.freeze = { r: +attr(pane, 'ySplit') || 0, c: +attr(pane, 'xSplit') || 0 };
    const sel = kid(sv, 'selection'); if (sel && attr(sel, 'activeCell')) { const a = parseAddr(attr(sel, 'activeCell')); if (a) { s.sel = { ranges: [{ r1: a.r, c1: a.c, r2: a.r, c2: a.c }], r: a.r, c: a.c }; } }
  }
  const fmtPr = kid(root, 'sheetFormatPr'); let defRowH = DEF_ROW_H; if (fmtPr && attr(fmtPr, 'defaultRowHeight')) defRowH = Math.round(+attr(fmtPr, 'defaultRowHeight') * 4 / 3);
  if (fmtPr && attr(fmtPr, 'defaultColWidth')) s.defColWx = +attr(fmtPr, 'defaultColWidth');
  // columns
  all(kid(root, 'cols'), 'col').forEach(c => {
    const a = +attr(c, 'min') - 1, b = Math.min(+attr(c, 'max') - 1, MAX_COLS - 1);
    const w = attr(c, 'width') ? Math.round(+attr(c, 'width') * 7 + 5) : null;
    const st = attr(c, 's') ? styles.xfs[+attr(c, 's')] : null;
    if (b - a > 2000) { if (st && b >= 16000) {} return; }
    for (let i = a; i <= b; i++) { if (w != null && attr(c, 'customWidth') !== '0') s.colW[i] = w; else if (w != null) s.colW[i] = w; if (attr(c, 'hidden') === '1') s.hiddenC[i] = true; if (st) s.colS[i] = st; if (attr(c, 'outlineLevel')) s.outlineC[i] = +attr(c, 'outlineLevel'); }
  });
  // rows & cells (optimized single-pass iterator for large datasets with 100k+ cells)
  const shared = {};
  const sd = kid(root, 'sheetData');
  let ri = 0;
  for (let row = sd ? sd.firstChild : null; row; row = row.nextSibling) {
    if (row.nodeType !== 1 || (row.localName !== 'row' && row.nodeName !== 'row')) continue;
    const rAttr = row.getAttribute('r');
    const r = rAttr ? +rAttr - 1 : ri;
    ri++;
    const ht = row.getAttribute('ht');
    if (ht && (row.getAttribute('customHeight') === '1' || Math.abs(+ht * 4 / 3 - DEF_ROW_H) > 1)) s.rowH[r] = Math.round(+ht * 4 / 3);
    if (row.getAttribute('hidden') === '1') s.hiddenR[r] = true;
    const rS = row.getAttribute('s');
    if (row.getAttribute('customFormat') === '1' && rS) s.rowS[r] = styles.xfs[+rS];
    const outline = row.getAttribute('outlineLevel');
    if (outline) s.outlineR[r] = +outline;

    let ci = 0;
    for (let c = row.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1 || (c.localName !== 'c' && c.nodeName !== 'c')) continue;
      const cR = c.getAttribute('r');
      const a = cR ? parseAddr(cR) : { r, c: ci };
      ci = a.c + 1;
      const t = c.getAttribute('t');
      const sId = c.getAttribute('s');
      const cell = {};
      if (sId) { const st = styles.xfs[+sId]; if (st) cell.s = st; }

      let vText = null, fEl = null, isEl = null;
      for (let ch = c.firstChild; ch; ch = ch.nextSibling) {
        if (ch.nodeType !== 1) continue;
        const ln = ch.localName || ch.nodeName;
        if (ln === 'v') vText = ch.textContent;
        else if (ln === 'f') fEl = ch;
        else if (ln === 'is') isEl = ch;
      }

      let v = vText;
      if (t === 's') v = v != null ? sst[+v] : null;
      else if (t === 'inlineStr') { v = isEl ? all(isEl, 't').map(textOf).join('') : ''; }
      else if (t === 'b') v = v === '1';
      else if (t === 'e') v = v ? errOf(v) : null;
      else if (t === 'str') v = v;
      else if (v != null) v = +v;
      if (v != null) cell.v = v;
      if (fEl) {
        const ft = fEl.getAttribute('t');
        let f = fEl.textContent;
        if (ft === 'shared') {
          const si = fEl.getAttribute('si');
          if (f) shared[si] = { f, r: a.r, c: a.c };
          else if (shared[si]) f = shiftFormula(shared[si].f, a.r - shared[si].r, a.c - shared[si].c);
        }
        if (f) cell.f = f.replace(/_xlfn\.(_xlws\.)?/g, '').replace(/_xlpm\./g, '');
      }
      if (cell.v != null || cell.f || cell.s) s.cells.set(key(a.r, a.c), cell);
    }
  }
  all(kid(root, 'mergeCells'), 'mergeCell').forEach(m => { const rg = parseRangeAddr(attr(m, 'ref')); if (rg) s.merges.push(rg); });
  // hyperlinks
  all(kid(root, 'hyperlinks'), 'hyperlink').forEach(hl => {
    const a = parseAddr(String(attr(hl, 'ref')).split(':')[0]); if (!a) return;
    const rid = attr(hl, 'r:id'); const url = rid && srels[rid] ? srels[rid].target : null; const loc = attr(hl, 'location');
    const x2 = s.cells.get(key(a.r, a.c)) || {}; x2.link = url ? { url } : { loc }; s.cells.set(key(a.r, a.c), x2);
  });
  // data validation
  all(kid(root, 'dataValidations'), 'dataValidation').forEach(dv => {
    const type = attr(dv, 'type'); if (!type) return;
    const f1 = kid(dv, 'formula1') ? kid(dv, 'formula1').textContent : '', f2 = kid(dv, 'formula2') ? kid(dv, 'formula2').textContent : '';
    String(attr(dv, 'sqref')).split(' ').forEach(sq => { const rg = parseRangeAddr(sq); if (!rg) return; s.dv.push({ range: rg, type: type === 'textLength' ? 'length' : type, op: ({ between: 'between', notBetween: 'notbetween', equal: 'eq', notEqual: 'ne', greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le' })[attr(dv, 'operator') || 'between'], v1: f1, v2: f2, source: type === 'list' ? (/^"/.test(f1) ? f1.replace(/^"|"$/g, '') : '=' + f1) : '', formula: type === 'custom' ? '=' + f1 : '', errMsg: attr(dv, 'error'), inMsg: attr(dv, 'prompt'), errStyle: attr(dv, 'errorStyle') || 'stop' }); });
  });
  // conditional formatting
  all(root, 'conditionalFormatting').forEach(cf => {
    const ranges = String(attr(cf, 'sqref')).split(' ').map(parseRangeAddr).filter(Boolean);
    kids(cf, 'cfRule').forEach(rule => {
      const t = attr(rule, 'type'); const dxf = styles.dxfs[+attr(rule, 'dxfId')] || {}; const fs = all(rule, 'formula').map(f => f.textContent);
      let o = null;
      if (t === 'cellIs') o = { type: 'cell', op: ({ greaterThan: 'gt', lessThan: 'lt', between: 'between', equal: 'eq', notEqual: 'ne', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le', notBetween: 'notbetween' })[attr(rule, 'operator')], v1: '=' + fs[0], v2: fs[1] != null ? '=' + fs[1] : null };
      else if (t === 'expression') o = { type: 'formula', formula: '=' + fs[0] };
      else if (t === 'dataBar') { const c = all(rule, 'color')[0]; o = { type: 'databar', color: colorOf(c, theme) || '#638ec6' }; }
      else if (t === 'colorScale') { const cs = all(rule, 'color').map(c => colorOf(c, theme)); o = { type: cs.length === 3 ? 'scale3' : 'scale2', colors: cs }; }
      else if (t === 'duplicateValues') o = { type: 'duplicate' }; else if (t === 'uniqueValues') o = { type: 'unique' };
      else if (t === 'top10') o = { type: attr(rule, 'bottom') === '1' ? 'bottom' : 'top', n: +attr(rule, 'rank') || 10, percent: attr(rule, 'percent') === '1' };
      else if (t === 'aboveAverage') o = { type: attr(rule, 'aboveAverage') === '0' ? 'belowavg' : 'aboveavg' };
      else if (t === 'containsText') o = { type: 'text', op: 'contains', text: attr(rule, 'text') };
      else if (t === 'containsBlanks') o = { type: 'blank' }; else if (t === 'containsErrors') o = { type: 'error' };
      else if (t === 'iconSet') o = { type: 'icons' };
      if (o) ranges.forEach(rg => s.cf.push(Object.assign({ range: rg, fill: dxf.fill, color: dxf.color, b: dxf.b, stopIfTrue: attr(rule, 'stopIfTrue') === '1' }, o)));
    });
  });
  const af = kid(root, 'autoFilter'); if (af) { const rg = parseRangeAddr(attr(af, 'ref')); if (rg) s.filter = Object.assign(rg, { crit: {} }); }
  const ps = kid(root, 'pageSetup'); if (ps) { if (attr(ps, 'orientation')) s.pageSetup.orientation = attr(ps, 'orientation') === 'landscape' ? 'landscape' : 'portrait'; const pz = { 9: 'A4', 8: 'A3', 11: 'A5', 1: 'Letter', 5: 'Legal' }[attr(ps, 'paperSize')]; if (pz) s.pageSetup.size = pz; }
  // v0.0.8: fit-to-page, scale, order, margins, header/footer, manual page breaks
  { const P = s.pageSetup; const fp = kid(kid(root, 'sheetPr'), 'pageSetUpPr');
    if (ps) { const fw = attr(ps, 'fitToWidth'), fh = attr(ps, 'fitToHeight'); if (fp && attr(fp, 'fitToPage') === '1') { const w = fw == null ? 1 : +fw, hh = fh == null ? 1 : +fh; P.fit = w && hh ? 'page' : w ? 'width' : hh ? 'rows1' : 'none'; } else if (attr(ps, 'scale')) P.scale = +attr(ps, 'scale'); if (attr(ps, 'pageOrder') === 'overThenDown') P.order = 'overThenDown'; }
    const pm = kid(root, 'pageMargins'); if (pm) { const mm = k => Math.round((+attr(pm, k) || 0) * 25.4); const m = [mm('top'), mm('right'), mm('bottom'), mm('left')]; const named = { normal: [19, 18, 19, 18], wide: [25, 25, 25, 25], narrow: [13, 6, 13, 6] }; let nm = null; for (const k in named) if (named[k].every((v, i) => Math.abs(v - m[i]) <= 1)) nm = k; P.margins = nm || m; }
    const po = kid(root, 'printOptions'); if (po) { if (attr(po, 'gridLines') === '1') P.gridlines = true; if (attr(po, 'headings') === '1') P.headings = true; if (attr(po, 'horizontalCentered') === '1') P.centerH = true; }
    const hfn = kid(root, 'headerFooter'); const hft = n => { const e = kid(hfn, n); return e ? e.textContent.replace(/&[LCR]/g, '').replace(/&P/g, '&[Page]').replace(/&N/g, '&[Pages]').replace(/&A/g, '&[Tab]').replace(/&D/g, '&[Date]').replace(/&T/g, '&[Time]').replace(/&F/g, '&[File]').replace(/&"[^"]*"/g, '').replace(/&\d+/g, '') : ''; };
    if (hfn) { const hh = hft('oddHeader'), ff = hft('oddFooter'); if (hh) P.header = hh; if (ff) P.footer = ff; }
    const brks = n => { const e = kid(root, n); return e ? all(e, 'brk').filter(b => attr(b, 'man') !== '0').map(b => +attr(b, 'id')).filter(x => x > 0) : []; };
    const rbs = brks('rowBreaks'), cbs = brks('colBreaks'); if (rbs.length) P.rowBreaks = rbs; if (cbs.length) P.colBreaks = cbs; }
  const tc = kid(kid(root, 'sheetPr'), 'tabColor'); if (tc) s.tabColor = colorOf(tc, theme);
  const prot = kid(root, 'sheetProtection'); if (prot && attr(prot, 'sheet') === '1') s.protected = { pw: '' };
}
async function readDrawing(zip, s, path, text, rels) {
  const t = await text(path); if (!t) return; const x = parseXml(t); const dr = await rels(path);
  const anchors = all(x, 'twoCellAnchor').concat(all(x, 'oneCellAnchor'));
  for (const an of anchors) {
    const from = kid(an, 'from'); const to = kid(an, 'to');
    const fc = +(kid(from, 'col') || {}).textContent || 0, fr = +(kid(from, 'row') || {}).textContent || 0;
    const tc = to ? +kid(to, 'col').textContent : fc + 5, tr = to ? +kid(to, 'row').textContent : fr + 10;
    const px = c => { let x2 = 0; for (let i = 0; i < c; i++) x2 += s.colWidth(i); return x2; }, py = r => { let y = 0; for (let i = 0; i < r; i++) y += s.rowHeight(i); return y; };
    const foff = +(kid(from, 'colOff') || {}).textContent / 9525 || 0, froff = +(kid(from, 'rowOff') || {}).textContent / 9525 || 0;
    const x1 = px(fc) + foff, y1 = py(fr) + froff; const w = Math.max(40, px(tc) - x1 + (to ? +(kid(to, 'colOff') || {}).textContent / 9525 || 0 : 0)), hh = Math.max(30, py(tr) - y1 + (to ? +(kid(to, 'rowOff') || {}).textContent / 9525 || 0 : 0));
    const pic = all(an, 'pic')[0];
    if (pic) { const blip = all(pic, 'blip')[0]; const rid = blip && (attr(blip, 'r:embed') || blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed')); const rel = rid && dr[rid]; if (rel) { const f = zip.file(rel.target); if (f) { const b = await f.async('uint8array'); const ext = extOf(rel.target); s.objects.push({ id: 'img' + Math.random().toString(36).slice(2), type: 'image', src: bytesToDataUrl(b, ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : ext === 'svg' ? 'image/svg+xml' : 'image/jpeg'), x: x1, y: y1, w, h: hh, alt: (all(pic, 'cNvPr')[0] && attr(all(pic, 'cNvPr')[0], 'descr')) || '' }); } } continue; }
    const gf = all(an, 'graphicFrame')[0];
    if (gf) { const ch = all(gf, 'chart')[0]; const rid = ch && (attr(ch, 'r:id') || ch.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')); const rel = rid && dr[rid]; if (rel) { const ct = await text(rel.target); if (ct) { const cx = parseXml(ct); const plot = all(cx, 'plotArea')[0]; let kind = 'column'; if (plot) { if (all(plot, 'pieChart').length) kind = 'pie'; else if (all(plot, 'doughnutChart').length) kind = 'doughnut'; else if (all(plot, 'lineChart').length) kind = 'line'; else if (all(plot, 'areaChart').length) kind = 'area'; else if (all(plot, 'scatterChart').length) kind = 'scatter'; else if (all(plot, 'barChart').length) { const bd = all(plot, 'barDir')[0]; kind = bd && attr(bd, 'val') === 'bar' ? 'bar' : 'column'; } } const refs = all(cx, 'f').map(f => f.textContent); const title = all(all(cx, 'title')[0] || cx, 't').map(z => z.textContent).join(''); let src = null; if (refs.length) { const rr = refs.map(r2 => parseRangeAddr(r2.replace(/^.*!/, '').replace(/\$/g, ''))).filter(Boolean); if (rr.length) src = rangeAddr({ r1: Math.min.apply(null, rr.map(z => z.r1)), c1: Math.min.apply(null, rr.map(z => z.c1)), r2: Math.max.apply(null, rr.map(z => z.r2)), c2: Math.max.apply(null, rr.map(z => z.c2)) }); } s.objects.push({ id: 'ch' + Math.random().toString(36).slice(2), type: 'chart', kind, src, x: x1, y: y1, w, h: hh, title }); } } }
    const sp = all(an, 'sp')[0];
    if (sp && !pic) { const txt = all(sp, 't').map(z => z.textContent).join('\n'); if (txt) s.objects.push({ id: 'tb' + Math.random().toString(36).slice(2), type: 'textbox', text: txt, x: x1, y: y1, w, h: hh }); }
  }
}

// ================================================================ WRITE
const NEW_FNS = ['XLOOKUP', 'XMATCH', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'LET', 'LAMBDA', 'IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'CONCAT', 'TEXTJOIN', 'XOR', 'IFNA', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'RANK.AVG', 'MODE.SNGL', 'MODE.MULT', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'TEXTBEFORE', 'TEXTAFTER', 'TEXTSPLIT', 'VSTACK', 'HSTACK', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'EXPAND', 'MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY', 'ISOMITTED', 'DAYS', 'ISOWEEKNUM', 'CEILING.MATH', 'FLOOR.MATH', 'AGGREGATE', 'NUMBERVALUE', 'UNICHAR', 'UNICODE', 'ARRAYTOTEXT', 'VALUETOTEXT', 'REGEXTEST', 'REGEXEXTRACT', 'REGEXREPLACE', 'FORECAST.LINEAR', 'NETWORKDAYS.INTL', 'WORKDAY.INTL', 'COVARIANCE.P', 'COVARIANCE.S', 'GROUPBY', 'PIVOTBY', 'PERCENTOF', 'TRIMRANGE', 'IMAGE', 'SHEET', 'SHEETS', 'FORMULATEXT', 'ISFORMULA', 'CONFIDENCE.NORM', 'CONFIDENCE.T', 'T.DIST', 'T.INV', 'CHISQ.DIST', 'BINOM.DIST', 'POISSON.DIST', 'EXPON.DIST', 'GAMMA.DIST', 'LOGNORM.DIST', 'WEIBULL.DIST', 'F.DIST', 'BETA.DIST', 'ERF.PRECISE', 'ERFC.PRECISE', 'GAMMALN.PRECISE', 'BITAND', 'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT', 'IMCOSH', 'IMSINH', 'IMTAN', 'ARABIC', 'BASE', 'DECIMAL', 'COMBINA', 'PERMUTATIONA', 'PHI', 'GAUSS', 'PDURATION', 'RRI', 'ACOT', 'ACOTH', 'COT', 'COTH', 'CSC', 'CSCH', 'SEC', 'SECH', 'ENCODEURL', 'FILTERXML', 'WEBSERVICE', 'DBCS'];
const WS_FNS = { FILTER: 1, SORT: 1 };
const PARAM_FNS = ['LET', 'LAMBDA', 'MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY'];
function xlfn(f) {
  return f.replace(/("(?:[^"]|"")*")|\b([A-Z][A-Z0-9.]*)(?=\()/g, (m, s, name) => { if (s) return s; if (NEW_FNS.indexOf(name) >= 0) return '_xlfn.' + (WS_FNS[name] ? '_xlws.' : '') + name; return name; });
}
function xmlEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ''); }
function hex(c) { if (!c) return null; c = String(c).replace('#', ''); if (c.length === 3) c = c.split('').map(x => x + x).join(''); if (!/^[0-9a-f]{6}$/i.test(c)) { const named = { red: 'FF0000', blue: '0000FF', green: '008000', black: '000000', white: 'FFFFFF', yellow: 'FFFF00', transparent: null }; return named[c.toLowerCase()] ? 'FF' + named[c.toLowerCase()] : null; } return 'FF' + c.toUpperCase(); }

export async function writeXlsx(wb, opts) {
  opts = opts || {}; const macro = !!opts.macro;
  if (macro) wb.sheets.forEach((s, i) => { if (!s.codeName) { let k = i + 1, nm; do { nm = 'Sheet' + k++; } while (wb.sheets.some(x => x !== s && x.codeName === nm)); s.codeName = nm; } });
  const JSZip = await JSZipLib(); const zip = new JSZip();
  // style registry
  const fonts = ['<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const numFmts = []; const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']; const xfMap = new Map(); const dxfs = [];
  const fmtId = nf => { if (!nf || nf === 'General') return 0; for (const k in BUILTIN_FMT) if (BUILTIN_FMT[k] === nf && +k !== 14 && +k !== 22) return +k; let i = numFmts.findIndex(n => n[1] === nf); if (i < 0) { numFmts.push([164 + numFmts.length, nf]); i = numFmts.length - 1; } return numFmts[i][0]; };
  const regist = (arr, xml) => { let i = arr.indexOf(xml); if (i < 0) { arr.push(xml); i = arr.length - 1; } return i; };
  const fontXml = st => '<font>' + (st.b ? '<b/>' : '') + (st.i ? '<i/>' : '') + (st.s ? '<strike/>' : '') + (st.u ? '<u' + (st.u === 'double' ? ' val="double"' : '') + '/>' : '') + '<sz val="' + (st.sz || 11) + '"/>' + (hex(st.color) ? '<color rgb="' + hex(st.color) + '"/>' : '<color theme="1"/>') + '<name val="' + xmlEsc(st.font || 'Calibri') + '"/><family val="2"/></font>';
  const sideXml = (n, b) => b ? '<' + n + ' style="' + (b.s === 'hair' || b.s === 'dotted' || b.s === 'dashed' || b.s === 'thin' || b.s === 'medium' || b.s === 'thick' || b.s === 'double' || b.s === 'mediumDashed' || b.s === 'dashDot' ? b.s : 'thin') + '"><color rgb="' + (hex(b.c) || 'FF000000') + '"/></' + n + '>' : '<' + n + '/>';
  const xfOf = st => {
    if (!st) return 0; if (xfMap.has(st)) return xfMap.get(st);
    const fi = st.b || st.i || st.u || st.s || st.sz || st.color || st.font ? regist(fonts, fontXml(st)) : 0;
    const fl = st.fill && hex(st.fill) ? regist(fills, '<fill><patternFill patternType="solid"><fgColor rgb="' + hex(st.fill) + '"/><bgColor indexed="64"/></patternFill></fill>') : 0;
    const bd = st.bd ? regist(borders, '<border' + (st.bd.u ? ' diagonalDown="1"' : '') + (st.bd.d ? ' diagonalUp="1"' : '') + '>' + sideXml('left', st.bd.l) + sideXml('right', st.bd.r) + sideXml('top', st.bd.t) + sideXml('bottom', st.bd.b) + sideXml('diagonal', st.bd.u || st.bd.d) + '</border>') : 0;
    const nf = fmtId(st.nf);
    const al = st.ha || st.va || st.wrap || st.indent || st.rot || st.shrink ? '<alignment' + (st.ha ? ' horizontal="' + (st.ha === 'centerAcross' ? 'centerContinuous' : st.ha) + '"' : '') + (st.va ? ' vertical="' + (st.va === 'middle' ? 'center' : st.va) + '"' : '') + (st.wrap ? ' wrapText="1"' : '') + (st.indent ? ' indent="' + st.indent + '"' : '') + (st.rot ? ' textRotation="' + st.rot + '"' : '') + (st.shrink ? ' shrinkToFit="1"' : '') + '/>' : '';
    const pr = st.unlocked || st.hideF ? '<protection' + (st.unlocked ? ' locked="0"' : '') + (st.hideF ? ' hidden="1"' : '') + '/>' : '';
    const xml = '<xf numFmtId="' + nf + '" fontId="' + fi + '" fillId="' + fl + '" borderId="' + bd + '" xfId="0"' + (nf ? ' applyNumberFormat="1"' : '') + (fi ? ' applyFont="1"' : '') + (fl ? ' applyFill="1"' : '') + (bd ? ' applyBorder="1"' : '') + (al ? ' applyAlignment="1"' : '') + (pr ? ' applyProtection="1"' : '') + (al || pr ? '>' + al + pr + '</xf>' : '/>');
    const i = regist(xfs, xml); xfMap.set(st, i); return i;
  };
  const dxfOf = rule => { const x = '<dxf>' + (rule.color || rule.b ? '<font>' + (rule.b ? '<b/>' : '') + (hex(rule.color) ? '<color rgb="' + hex(rule.color) + '"/>' : '') + '</font>' : '') + (rule.fill && hex(rule.fill) ? '<fill><patternFill><bgColor rgb="' + hex(rule.fill) + '"/></patternFill></fill>' : '') + '</dxf>'; return regist(dxfs, x); };
  const sst = [], sstMap = new Map(); let sstCount = 0;
  const si = s => { sstCount++; let i = sstMap.get(s); if (i === undefined) { i = sst.length; sst.push(s); sstMap.set(s, i); } return i; };
  const ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>', '<Default Extension="png" ContentType="image/png"/>', '<Default Extension="jpeg" ContentType="image/jpeg"/>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>', '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>', '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>', '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>', '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'];
  const wbRels = []; const sheetsXml = []; let imgN = 0, chartN = 0, drawN = 0, cmtN = 0;
  wb.sheets.forEach((s, idx) => {
    const n = idx + 1; const srels = [];
    ct.push('<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
    wbRels.push('<Relationship Id="rId' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>');
    sheetsXml.push('<sheet name="' + xmlEsc(s.name) + '" sheetId="' + n + '"' + (s.hidden ? ' state="hidden"' : '') + ' r:id="rId' + n + '"/>');
    const u = s.used();
    // rows
    const byRow = new Map();
    s.cells.forEach((x, k) => { const r = Math.floor(k / 20000); if (!byRow.has(r)) byRow.set(r, []); byRow.get(r).push(k); });
    Object.keys(s.rowH).concat(Object.keys(s.hiddenR)).concat(Object.keys(s.rowS)).forEach(r => { if (!byRow.has(+r)) byRow.set(+r, []); });
    const rowsXml = [];
    Array.from(byRow.keys()).sort((a, b) => a - b).forEach(r => {
      const ks = byRow.get(r).sort((a, b) => a - b);
      const cells = ks.map(k => {
        const x = s.cells.get(k); const c = k % 20000; const ref = colName(c) + (r + 1);
        const sAttr = x.s ? ' s="' + xfOf(x.s) + '"' : '';
        let v = x.v, out;
        if (x.f) {
          const arr = x.spillTo ? ' cm="1"' : '';
          const fx = '<f' + (x.spillTo ? ' t="array" ref="' + ref + ':' + colName(x.spillTo.c2) + (x.spillTo.r2 + 1) + '"' : '') + '>' + xmlEsc(xlfn(x.f)) + '</f>';
          if (typeof v === 'string') return '<c r="' + ref + '"' + sAttr + arr + ' t="str">' + fx + '<v>' + xmlEsc(v) + '</v></c>';
          if (typeof v === 'boolean') return '<c r="' + ref + '"' + sAttr + arr + ' t="b">' + fx + '<v>' + (v ? 1 : 0) + '</v></c>';
          if (v instanceof Err) return '<c r="' + ref + '"' + sAttr + arr + ' t="e">' + fx + '<v>' + xmlEsc(v.code) + '</v></c>';
          return '<c r="' + ref + '"' + sAttr + arr + '>' + fx + (typeof v === 'number' && isFinite(v) ? '<v>' + v + '</v>' : '') + '</c>';
        }
        if (v == null) return '<c r="' + ref + '"' + sAttr + '/>';
        if (typeof v === 'number') return '<c r="' + ref + '"' + sAttr + '><v>' + v + '</v></c>';
        if (typeof v === 'boolean') return '<c r="' + ref + '"' + sAttr + ' t="b"><v>' + (v ? 1 : 0) + '</v></c>';
        if (v instanceof Err) return '<c r="' + ref + '"' + sAttr + ' t="e"><v>' + xmlEsc(v.code) + '</v></c>';
        return '<c r="' + ref + '"' + sAttr + ' t="s"><v>' + si(String(v)) + '</v></c>';
      }).join('');
      const ht = s.rowH[r] != null ? ' ht="' + (s.rowH[r] * 0.75).toFixed(2) + '" customHeight="1"' : '';
      const rs = s.rowS[r] ? ' s="' + xfOf(s.rowS[r]) + '" customFormat="1"' : '';
      rowsXml.push('<row r="' + (r + 1) + '"' + ht + (s.hiddenR[r] ? ' hidden="1"' : '') + rs + (s.outlineR[r] ? ' outlineLevel="' + s.outlineR[r] + '"' : '') + '>' + cells + '</row>');
    });
    // cols
    const colIdx = Array.from(new Set(Object.keys(s.colW).concat(Object.keys(s.hiddenC)).concat(Object.keys(s.colS)).map(Number))).sort((a, b) => a - b);
    const colsXml = colIdx.length ? '<cols>' + colIdx.map(c => '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + (((s.colW[c] != null ? s.colW[c] : DEF_COL_W) - 5) / 7).toFixed(2) + '"' + (s.colW[c] != null ? ' customWidth="1"' : '') + (s.hiddenC[c] ? ' hidden="1"' : '') + (s.colS[c] ? ' style="' + xfOf(s.colS[c]) + '"' : '') + (s.outlineC[c] ? ' outlineLevel="' + s.outlineC[c] + '"' : '') + '/>').join('') + '</cols>' : '';
    const fr = s.freeze;
    const pane = fr.r || fr.c ? '<pane' + (fr.c ? ' xSplit="' + fr.c + '"' : '') + (fr.r ? ' ySplit="' + fr.r + '"' : '') + ' topLeftCell="' + colName(fr.c) + (fr.r + 1) + '" activePane="' + (fr.r && fr.c ? 'bottomRight' : fr.r ? 'bottomLeft' : 'topRight') + '" state="frozen"/>' : '';
    const view = '<sheetViews><sheetView' + (idx === wb.active ? ' tabSelected="1"' : '') + (s.showGrid ? '' : ' showGridLines="0"') + (s.showHead ? '' : ' showRowColHeaders="0"') + (s.zoom && s.zoom !== 100 ? ' zoomScale="' + Math.round(s.zoom) + '"' : '') + ' workbookViewId="0">' + pane + '<selection' + (pane ? ' pane="' + (fr.r && fr.c ? 'bottomRight' : fr.r ? 'bottomLeft' : 'topRight') + '"' : '') + ' activeCell="' + addr(s.sel.r, s.sel.c) + '" sqref="' + addr(s.sel.r, s.sel.c) + '"/></sheetView></sheetViews>';
    const merges = s.merges.length ? '<mergeCells count="' + s.merges.length + '">' + s.merges.map(m => '<mergeCell ref="' + rangeAddr(m) + '"/>').join('') + '</mergeCells>' : '';
    // hyperlinks
    const links = []; s.cells.forEach((x, k) => { if (!x.link) return; const ref = addr(Math.floor(k / 20000), k % 20000); if (x.link.url) { const rid = 'rIdL' + links.length; srels.push('<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + xmlEsc(x.link.url) + '" TargetMode="External"/>'); links.push('<hyperlink ref="' + ref + '" r:id="' + rid + '"/>'); } else if (x.link.loc) links.push('<hyperlink ref="' + ref + '" location="' + xmlEsc(x.link.loc) + '"/>'); });
    const dvXml = s.dv.length ? '<dataValidations count="' + s.dv.length + '">' + s.dv.map(d => { const t = d.type === 'length' ? 'textLength' : d.type; const op = ({ between: 'between', notbetween: 'notBetween', eq: 'equal', ne: 'notEqual', gt: 'greaterThan', lt: 'lessThan', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual' })[d.op] || 'between'; const f1 = d.type === 'list' ? (/^=/.test(d.source || '') ? d.source.slice(1) : '"' + (d.source || '') + '"') : d.type === 'custom' ? String(d.formula || '').replace(/^=/, '') : String(d.v1 || '').replace(/^=/, ''); return '<dataValidation type="' + t + '"' + (d.type !== 'list' && d.type !== 'custom' ? ' operator="' + op + '"' : '') + ' allowBlank="1" showInputMessage="1" showErrorMessage="1"' + (d.errStyle && d.errStyle !== 'stop' ? ' errorStyle="' + (d.errStyle === 'info' ? 'information' : d.errStyle) + '"' : '') + (d.errMsg ? ' error="' + xmlEsc(d.errMsg) + '"' : '') + (d.inMsg ? ' prompt="' + xmlEsc(d.inMsg) + '"' : '') + ' sqref="' + rangeAddr(d.range) + '"><formula1>' + xmlEsc(f1) + '</formula1>' + (d.v2 && /between/.test(d.op) ? '<formula2>' + xmlEsc(String(d.v2).replace(/^=/, '')) + '</formula2>' : '') + '</dataValidation>'; }).join('') + '</dataValidations>' : '';
    let prio = 1;
    const cfXml = s.cf.map(r => {
      let rule = '';
      const v = x2 => xmlEsc(String(x2 == null ? '' : x2).replace(/^=/, '') || '0');
      if (r.type === 'cell') rule = '<cfRule type="cellIs" dxfId="' + dxfOf(r) + '" priority="' + prio++ + '" operator="' + ({ gt: 'greaterThan', lt: 'lessThan', between: 'between', eq: 'equal', ne: 'notEqual', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual', notbetween: 'notBetween' })[r.op] + '"><formula>' + (isNaN(+String(r.v1).replace(/^=/, '')) && !/^=/.test(String(r.v1)) ? '"' + xmlEsc(r.v1) + '"' : v(r.v1)) + '</formula>' + (r.v2 != null && /between/.test(r.op) ? '<formula>' + v(r.v2) + '</formula>' : '') + '</cfRule>';
      else if (r.type === 'formula') rule = '<cfRule type="expression" dxfId="' + dxfOf(r) + '" priority="' + prio++ + '"><formula>' + v(r.formula) + '</formula></cfRule>';
      else if (r.type === 'databar') rule = '<cfRule type="dataBar" priority="' + prio++ + '"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="' + (hex(r.color) || 'FF638EC6') + '"/></dataBar></cfRule>';
      else if (r.type === 'scale2' || r.type === 'scale3') { const cs = r.colors || (r.type === 'scale3' ? ['#f8696b', '#ffeb84', '#63be7b'] : ['#ffffff', '#63be7b']); rule = '<cfRule type="colorScale" priority="' + prio++ + '"><colorScale><cfvo type="min"/>' + (cs.length === 3 ? '<cfvo type="percentile" val="50"/>' : '') + '<cfvo type="max"/>' + cs.map(c => '<color rgb="' + hex(c) + '"/>').join('') + '</colorScale></cfRule>'; }
      else if (r.type === 'duplicate') rule = '<cfRule type="duplicateValues" dxfId="' + dxfOf(r) + '" priority="' + prio++ + '"/>';
      else if (r.type === 'unique') rule = '<cfRule type="uniqueValues" dxfId="' + dxfOf(r) + '" priority="' + prio++ + '"/>';
      else if (r.type === 'top' || r.type === 'bottom') rule = '<cfRule type="top10" dxfId="' + dxfOf(r) + '" priority="' + prio++ + '" rank="' + (r.n || 10) + '"' + (r.percent ? ' percent="1"' : '') + (r.type === 'bottom' ? ' bottom="1"' : '') + '/>';
      else if (r.type === 'aboveavg' || r.type === 'belowavg') rule = '<cfRule type="aboveAverage" dxfId="' + dxfOf(r) + '" priority="' + prio++ + '"' + (r.type === 'belowavg' ? ' aboveAverage="0"' : '') + '/>';
      else if (r.type === 'text') rule = '<cfRule type="containsText" dxfId="' + dxfOf(r) + '" priority="' + prio++ + '" operator="containsText" text="' + xmlEsc(r.text) + '"><formula>NOT(ISERROR(SEARCH("' + xmlEsc(r.text) + '",' + addr(r.range.r1, r.range.c1) + ')))</formula></cfRule>';
      if (!rule) return '';
      return '<conditionalFormatting sqref="' + rangeAddr(r.range) + '">' + rule + '</conditionalFormatting>';
    }).join('');
    const af = s.filter ? '<autoFilter ref="' + rangeAddr(s.filter) + '"/>' : '';
    const ps = s.pageSetup || {};
    // v0.0.8: margins, scale/fit, page order, header/footer and manual page breaks round-trip with Excel
    const MG = Array.isArray(ps.margins) ? ps.margins : ({ normal: [19, 18, 19, 18], wide: [25, 25, 25, 25], narrow: [13, 6, 13, 6] }[ps.margins] || [19, 18, 19, 18]);
    const inch = mm => (mm / 25.4).toFixed(3).replace(/\.?0+$/, '');
    const fitW = ps.fit === 'width' || ps.fit === 'page' ? 1 : 0, fitH = ps.fit === 'page' || ps.fit === 'rows1' ? 1 : 0, isFit = ps.fit && ps.fit !== 'none';
    const hfx = t => xmlEsc(String(t || '').replace(/&\[Page\]/g, '&P').replace(/&\[Pages\]/g, '&N').replace(/&\[Tab\]/g, '&A').replace(/&\[Date\]/g, '&D').replace(/&\[Time\]/g, '&T').replace(/&\[File\]/g, '&F'));
    const brk = (arr, max) => '<brk id="' + arr + '" max="' + max + '" man="1"/>';
    const rb = (ps.rowBreaks || []).filter(x => x > 0), cb = (ps.colBreaks || []).filter(x => x > 0);
    const pageXml = (ps.centerH ? '<printOptions horizontalCentered="1"' + (ps.gridlines ? ' gridLines="1"' : '') + (ps.headings ? ' headings="1"' : '') + '/>' : (ps.gridlines || ps.headings ? '<printOptions' + (ps.gridlines ? ' gridLines="1"' : '') + (ps.headings ? ' headings="1"' : '') + '/>' : '')) +
      '<pageMargins left="' + inch(MG[3]) + '" right="' + inch(MG[1]) + '" top="' + inch(MG[0]) + '" bottom="' + inch(MG[2]) + '" header="0.3" footer="0.3"/><pageSetup paperSize="' + ({ A4: 9, A3: 8, A5: 11, Letter: 1, Legal: 5 }[ps.size] || 9) + '"' + (!isFit && ps.scale && +ps.scale !== 100 ? ' scale="' + Math.round(+ps.scale) + '"' : '') + (isFit ? ' fitToWidth="' + fitW + '" fitToHeight="' + fitH + '"' : '') + (ps.order === 'overThenDown' ? ' pageOrder="overThenDown"' : '') + ' orientation="' + (ps.orientation || 'portrait') + '"/>' +
      (ps.header || ps.footer ? '<headerFooter>' + (ps.header ? '<oddHeader>&amp;C' + hfx(ps.header) + '</oddHeader>' : '') + (ps.footer ? '<oddFooter>&amp;C' + hfx(ps.footer) + '</oddFooter>' : '') + '</headerFooter>' : '') +
      (rb.length ? '<rowBreaks count="' + rb.length + '" manualBreakCount="' + rb.length + '">' + rb.map(r => brk(r, 16383)).join('') + '</rowBreaks>' : '') +
      (cb.length ? '<colBreaks count="' + cb.length + '" manualBreakCount="' + cb.length + '">' + cb.map(c => brk(c, 1048575)).join('') + '</colBreaks>' : '');
    // drawing (images & charts)
    let drawingXml = '';
    const objs = s.objects.filter(o => o.type === 'image' || o.type === 'chart');
    if (objs.length) {
      drawN++; const dn = drawN; const drels = [];
      srels.push('<Relationship Id="rIdD1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + dn + '.xml"/>');
      drawingXml = '<drawing r:id="rIdD1"/>';
      ct.push('<Override PartName="/xl/drawings/drawing' + dn + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
      const anchors = objs.map((o, i) => {
        const pos = posToCell(s, o.x, o.y), pos2 = posToCell(s, o.x + o.w, o.y + o.h);
        const anchor = '<xdr:from><xdr:col>' + pos.c + '</xdr:col><xdr:colOff>' + Math.round(pos.dx * 9525) + '</xdr:colOff><xdr:row>' + pos.r + '</xdr:row><xdr:rowOff>' + Math.round(pos.dy * 9525) + '</xdr:rowOff></xdr:from><xdr:to><xdr:col>' + pos2.c + '</xdr:col><xdr:colOff>' + Math.round(pos2.dx * 9525) + '</xdr:colOff><xdr:row>' + pos2.r + '</xdr:row><xdr:rowOff>' + Math.round(pos2.dy * 9525) + '</xdr:rowOff></xdr:to>';
        if (o.type === 'image') {
          imgN++; const m = /^data:image\/(png|jpe?g|gif)/.exec(o.src || ''); const ext = m ? (m[1] === 'jpg' ? 'jpeg' : m[1]) : 'png';
          zip.file('xl/media/image' + imgN + '.' + ext, dataUrlToBytes(o.src)); if (ext === 'gif') ct.push('<Default Extension="gif" ContentType="image/gif"/>');
          drels.push('<Relationship Id="rIdI' + i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image' + imgN + '.' + ext + '"/>');
          return '<xdr:twoCellAnchor editAs="oneCell">' + anchor + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (i + 2) + '" name="Picture ' + (i + 1) + '" descr="' + xmlEsc(o.alt || '') + '"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdI' + i + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + Math.round(o.w * 9525) + '" cy="' + Math.round(o.h * 9525) + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>';
        }
        chartN++;
        zip.file('xl/charts/chart' + chartN + '.xml', chartXml(s, o));
        ct.push('<Override PartName="/xl/charts/chart' + chartN + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
        drels.push('<Relationship Id="rIdC' + i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + chartN + '.xml"/>');
        return '<xdr:twoCellAnchor>' + anchor + '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + (i + 2) + '" name="Chart ' + (i + 1) + '" descr="' + xmlEsc(o.alt || '') + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC' + i + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
      });
      zip.file('xl/drawings/drawing' + dn + '.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + anchors.join('') + '</xdr:wsDr>');
      zip.file('xl/drawings/_rels/drawing' + dn + '.xml.rels', relsXml(drels));
    }
    // notes (comments + legacy VML so Excel shows them)
    const notes = []; s.cells.forEach((x, k) => { if (x.note) notes.push([Math.floor(k / 20000), k % 20000, x.note]); });
    let legacy = '';
    if (notes.length) {
      cmtN++;
      const authors = Array.from(new Set(notes.map(n => n[2].author || (n[2].thread && n[2].thread[0].author) || 'Yomi-Office Open Source')));
      zip.file('xl/comments' + cmtN + '.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>' + authors.map(a => '<author>' + xmlEsc(a) + '</author>').join('') + '</authors><commentList>' + notes.map(([r, c, n]) => { const au = n.author || (n.thread && n.thread[0].author) || 'Yomi-Office Open Source'; const txt = n.thread ? n.thread.map(m => m.author + ': ' + m.text).join('\n') : n.text; return '<comment ref="' + addr(r, c) + '" authorId="' + authors.indexOf(au) + '"><text><r><rPr><b/><sz val="9"/></rPr><t xml:space="preserve">' + xmlEsc(au) + ':</t></r><r><rPr><sz val="9"/></rPr><t xml:space="preserve">\n' + xmlEsc(txt) + '</t></r></text></comment>'; }).join('') + '</commentList></comments>');
      ct.push('<Override PartName="/xl/comments' + cmtN + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>');
      srels.push('<Relationship Id="rIdN1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments' + cmtN + '.xml"/>');
      srels.push('<Relationship Id="rIdV1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' + cmtN + '.vml"/>');
      zip.file('xl/drawings/vmlDrawing' + cmtN + '.vml', '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="' + cmtN + '"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' + notes.map(([r, c], i) => '<v:shape id="_x0000_s' + (1025 + i) + '" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:' + (i + 1) + ';visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>' + (c + 1) + ', 15, ' + r + ', 2, ' + (c + 3) + ', 15, ' + (r + 4) + ', 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>' + r + '</x:Row><x:Column>' + c + '</x:Column></x:ClientData></v:shape>').join('') + '</xml>');
      legacy = '<legacyDrawing r:id="rIdV1"/>';
    }
    const prot = s.protected ? '<sheetProtection sheet="1" objects="1" scenarios="1"' + (s.protected.fmt ? ' formatCells="0"' : '') + '/>' : '';
    const fitPr = ps.fit && ps.fit !== 'none';
    const sheetPr = s.tabColor || macro || fitPr ? '<sheetPr' + (macro ? ' codeName="' + xmlEsc(s.codeName) + '"' : '') + '>' + (s.tabColor ? '<tabColor rgb="' + hex(s.tabColor) + '"/>' : '') + (fitPr ? '<pageSetUpPr fitToPage="1"/>' : '') + '</sheetPr>' : '';
    const dim = '<dimension ref="' + (u.empty ? 'A1' : rangeAddr(u)) + '"/>';
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + sheetPr + dim + view + '<sheetFormatPr defaultRowHeight="15"/>' + colsXml + '<sheetData>' + rowsXml.join('') + '</sheetData>' + prot + af + merges + cfXml + dvXml + (links.length ? '<hyperlinks>' + links.join('') + '</hyperlinks>' : '') + pageXml + drawingXml + legacy + '</worksheet>';
    zip.file('xl/worksheets/sheet' + n + '.xml', xml);
    if (srels.length) zip.file('xl/worksheets/_rels/sheet' + n + '.xml.rels', relsXml(srels));
  });
  // defined names
  const dn = Object.keys(wb.names).map(n => '<definedName name="' + xmlEsc(wb.names[n].display || n) + '">' + xmlEsc(xlfn(String(wb.names[n].f || wb.names[n]).replace(/^=/, ''))) + '</definedName>').concat(wb.sheets.map((s, i) => (s.pageSetup && s.pageSetup.printArea ? '<definedName name="_xlnm.Print_Area" localSheetId="' + i + '">' + qn(s.name) + '!' + absRef(s.pageSetup.printArea) + '</definedName>' : '') + (s.pageSetup && (s.pageSetup.printTitles || s.pageSetup.printTitlesCols) ? '<definedName name="_xlnm.Print_Titles" localSheetId="' + i + '">' + [s.pageSetup.printTitlesCols ? qn(s.name) + '!' + String(s.pageSetup.printTitlesCols).replace(/\$/g, '').split(':').map(x => '$' + x).join(':') : '', s.pageSetup.printTitles ? qn(s.name) + '!' + String(s.pageSetup.printTitles).replace(/\$/g, '').split(':').map(x => '$' + x).join(':') : ''].filter(Boolean).join(',') + '</definedName>' : '')).filter(Boolean));
  zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr' + (macro ? ' codeName="ThisWorkbook"' : '') + ' defaultThemeVersion="164011"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300" activeTab="' + wb.active + '"/></bookViews><sheets>' + sheetsXml.join('') + '</sheets>' + (dn.length ? '<definedNames>' + dn.join('') + '</definedNames>' : '') + '<calcPr calcId="191029"' + (wb.calcMode === 'manual' ? ' calcMode="manual"' : '') + '/></workbook>');
  wbRels.push('<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>', '<Relationship Id="rIdSS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>', '<Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>');
  if (macro) {
    const { writeVbaProject } = await import('./vba/ovba.js');
    const user = ((wb.vba && wb.vba.modules) || []).filter(m => m.name !== '__Immediate');
    const docMods = [{ name: 'ThisWorkbook', type: 'workbook', code: (user.find(m => m.type === 'workbook') || {}).code || '' }].concat(wb.sheets.map(s => ({ name: s.codeName, type: 'sheet', code: (user.find(m => m.type === 'sheet' && m.name === s.codeName) || {}).code || '' })));
    const { jsaToVba } = await import('./vba/jsa.js');
    const mods = docMods.concat(user.filter(m => m.type !== 'workbook' && m.type !== 'sheet').map(m => ({ name: m.name, type: m.type === 'class' ? 'class' : 'module', code: m.type === 'jsa' ? jsaToVba(m.code) : m.code })));
    zip.file('xl/vbaProject.bin', writeVbaProject(mods));
    { const { vbaSidecar } = await import('./vba/ovba.js'); zip.file('oos/vba-source.json', vbaSidecar(docMods.concat(user.filter(m => m.type !== 'workbook' && m.type !== 'sheet').map(m => ({ name: m.name, type: m.type || 'module', code: m.code, wcode: m.type === 'jsa' ? jsaToVba(m.code) : m.code }))))); ct.push('<Default Extension="json" ContentType="application/json"/>'); }
    wbRels.push('<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>');
    ct.push('<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>');
    for (let i = 0; i < ct.length; i++) ct[i] = ct[i].replace('spreadsheetml.sheet.main+xml', 'XX').replace('application/vnd.openxmlformats-officedocument.XX', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml');
  }
  zip.file('xl/_rels/workbook.xml.rels', relsXml(wbRels));
  zip.file('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sstCount + '" uniqueCount="' + sst.length + '">' + sst.map(s => '<si><t xml:space="preserve">' + xmlEsc(s) + '</t></si>').join('') + '</sst>');
  zip.file('xl/metadata.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>');
  zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + (numFmts.length ? '<numFmts count="' + numFmts.length + '">' + numFmts.map(n => '<numFmt numFmtId="' + n[0] + '" formatCode="' + xmlEsc(n[1]) + '"/>').join('') + '</numFmts>' : '') + '<fonts count="' + fonts.length + '">' + fonts.join('') + '</fonts><fills count="' + fills.length + '">' + fills.join('') + '</fills><borders count="' + borders.length + '">' + borders.join('') + '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="' + dxfs.length + '">' + dxfs.join('') + '</dxfs></styleSheet>');
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + Array.from(new Set(ct)).join('') + '</Types>');
  zip.file('_rels/.rels', relsXml(['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>', '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>']));
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Yomi-Office Open Source</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>');
  zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application><AppVersion>16.0300</AppVersion></Properties>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
function qn(n) { return /^[A-Za-z_][\w.]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'"; }
function absRef(a) { return a.replace(/\$/g, '').replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'); }
function relsXml(list) { return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + list.join('') + '</Relationships>'; }
function posToCell(s, x, y) {
  let c = 0, acc = 0; while (c < MAX_COLS && acc + s.colWidth(c) <= x) { acc += s.colWidth(c); c++; } const dx = x - acc;
  let r = 0, accY = 0; while (r < MAX_ROWS && accY + s.rowHeight(r) <= y) { accY += s.rowHeight(r); r++; } const dy = y - accY;
  return { c, r, dx, dy };
}
function chartXml(s, o) {
  const rg = parseRangeAddr(o.src || 'A1:B2') || { r1: 0, c1: 0, r2: 1, c2: 1 };
  const sh = qn(s.name);
  const ref = (r1, c1, r2, c2) => sh + '!$' + colName(c1) + '$' + (r1 + 1) + ':$' + colName(c2) + '$' + (r2 + 1);
  const head = typeof s.value(rg.r1, rg.c1 + 1) === 'string';
  const catCol = typeof s.value(rg.r1 + 1, rg.c1) !== 'number' || o.kind === 'pie' || o.kind === 'doughnut';
  const series = []; const dataR1 = rg.r1 + (head ? 1 : 0);
  for (let c = rg.c1 + (catCol ? 1 : 0); c <= rg.c2; c++) series.push(c);
  const kind = o.kind || 'column';
  const ser = (c, i) => '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/>' + (head ? '<c:tx><c:strRef><c:f>' + ref(rg.r1, c, rg.r1, c) + '</c:f></c:strRef></c:tx>' : '') + (catCol ? (kind === 'scatter' ? '<c:xVal><c:numRef><c:f>' + ref(dataR1, rg.c1, rg.r2, rg.c1) + '</c:f></c:numRef></c:xVal>' : '<c:cat><c:strRef><c:f>' + ref(dataR1, rg.c1, rg.r2, rg.c1) + '</c:f></c:strRef></c:cat>') : '') + (kind === 'scatter' ? '<c:yVal>' : '<c:val>') + '<c:numRef><c:f>' + ref(dataR1, c, rg.r2, c) + '</c:f></c:numRef>' + (kind === 'scatter' ? '</c:yVal><c:smooth val="0"/>' : '</c:val>') + (kind === 'line' ? '<c:smooth val="0"/>' : '') + '</c:ser>';
  const sers = series.map(ser).join('');
  const axes = '<c:axId val="111"/><c:axId val="222"/>';
  let plot;
  if (kind === 'pie') plot = '<c:pieChart><c:varyColors val="1"/>' + sers + '<c:firstSliceAng val="0"/></c:pieChart>';
  else if (kind === 'doughnut') plot = '<c:doughnutChart><c:varyColors val="1"/>' + sers + '<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>';
  else if (kind === 'line') plot = '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + sers + '<c:marker val="1"/>' + axes + '</c:lineChart>';
  else if (kind === 'area') plot = '<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>' + sers + axes + '</c:areaChart>';
  else if (kind === 'scatter') plot = '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>' + sers + axes + '</c:scatterChart>';
  else plot = '<c:barChart><c:barDir val="' + (kind === 'bar' ? 'bar' : 'col') + '"/><c:grouping val="' + (o.stacked ? 'stacked' : 'clustered') + '"/><c:varyColors val="0"/>' + sers + (o.stacked ? '<c:overlap val="100"/>' : '') + '<c:gapWidth val="150"/>' + axes + '</c:barChart>';
  const axXml = kind === 'pie' || kind === 'doughnut' ? '' : (kind === 'scatter' ? '<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:valAx>' : '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + (kind === 'bar' ? 'l' : 'b') + '"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>') + '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + (kind === 'bar' ? 'b' : 'l') + '"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="111"/><c:crosses val="autoZero"/></c:valAx>';
  const title = o.title ? '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>' + xmlEsc(o.title) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>' : '<c:autoTitleDeleted val="1"/>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>' + title + '<c:plotArea><c:layout/>' + plot + axXml + '</c:plotArea>' + (o.legend === false ? '' : '<c:legend><c:legendPos val="' + (o.kind === 'pie' || o.kind === 'doughnut' ? 'r' : 'b') + '"/><c:overlay val="0"/></c:legend>') + '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>';
}

// ================================================================ CSV
export function readCsv(wb, bytes, sep, name) {
  let t = utf8Decode(bytes); if (t.charCodeAt(0) === 0xFEFF) t = t.slice(1);
  if (!sep) { const first = t.split(/\r?\n/)[0] || ''; const cnt = ch => (first.match(new RegExp('\\' + ch, 'g')) || []).length; sep = cnt('\t') > cnt(',') && cnt('\t') > cnt(';') ? '\t' : cnt(';') > cnt(',') ? ';' : ','; }
  const s = wb.addSheet(name ? baseName(name).slice(0, 31) : 'Sheet1');
  let r = 0, c = 0, cur = '', q = false; const flush = () => { if (cur !== '') { const p = parseInput(cur); const cell = { v: /^=/.test(cur) ? cur : p.v }; if (p.nf) cell.s = internStyle({ nf: p.nf }); s.cells.set(key(r, c), cell); } cur = ''; };
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) { if (ch === '"' && t[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; continue; }
    if (ch === '"' && cur === '') { q = true; continue; }
    if (ch === sep) { flush(); c++; continue; }
    if (ch === '\n' || ch === '\r') { if (ch === '\r' && t[i + 1] === '\n') i++; flush(); r++; c = 0; if (r >= MAX_ROWS) break; continue; }
    cur += ch;
  }
  flush();
  // auto column widths
  const u = s.used(); for (let cc = 0; cc <= Math.min(u.c2, 50); cc++) { let w = 0; for (let rr = 0; rr <= Math.min(u.r2, 200); rr++) { const v = s.value(rr, cc); if (v != null) w = Math.max(w, String(v).length); } if (w > 9) s.colW[cc] = Math.min(400, w * 7 + 12); }
  return s;
}
export function writeCsv(s, sep) {
  const u = s.used(); const lines = [];
  for (let r = 0; r <= u.r2; r++) { const row = []; for (let c = 0; c <= u.c2; c++) { const cell = s.get(r, c); const v = s.value(r, c); let t = v == null ? '' : v instanceof Err ? v.code : (cell && cell.s && cell.s.nf && typeof v === 'number' && isDateFormat(cell.s.nf)) ? format(v, cell.s.nf).t : typeof v === 'number' ? String(v) : String(v); if (t.indexOf(sep) >= 0 || /["\r\n]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"'; row.push(t); } lines.push(row.join(sep)); }
  return lines.join('\r\n');
}

// ================================================================ HTML (print / export / clipboard)
export function sheetToHtml(s, doc, opts) {
  opts = opts || {};
  const u = opts.area || s.used(); const out = [];
  const css = st => { if (!st) return ''; const a = []; if (st.b) a.push('font-weight:bold'); if (st.i) a.push('font-style:italic'); if (st.u || st.s) a.push('text-decoration:' + (st.u ? 'underline ' : '') + (st.s ? 'line-through' : '')); if (st.sz) a.push('font-size:' + st.sz + 'pt'); if (st.font) a.push("font-family:'" + st.font + "'"); if (st.color) a.push('color:' + st.color); if (st.fill && st.fill !== 'transparent') a.push('background:' + st.fill); if (st.ha && st.ha !== 'general') a.push('text-align:' + (st.ha === 'centerAcross' ? 'center' : st.ha)); if (st.va) a.push('vertical-align:' + (st.va === 'middle' ? 'middle' : st.va)); if (st.wrap) a.push('white-space:pre-wrap'); if (st.indent) a.push('padding-left:' + (st.indent * 9 + 3) + 'px'); if (st.bd) { const b = (k, side) => { const x = st.bd[k]; if (x) a.push('border-' + side + ':' + (x.s === 'thick' ? '3px' : x.s === 'medium' ? '2px' : '1px') + ' ' + (x.s === 'double' ? 'double' : /dash/.test(x.s) ? 'dashed' : /dot/.test(x.s) ? 'dotted' : 'solid') + ' ' + (x.c || '#000')); }; b('t', 'top'); b('b', 'bottom'); b('l', 'left'); b('r', 'right'); } return a.join(';'); };
  const grid = opts.gridlines ? 'border:1px solid #d4d4d4;' : '';
  const r1 = u.r1 || 0, c1 = u.c1 || 0, r2 = Math.min(u.r2, r1 + (opts.maxRows || 1e6)), c2 = u.c2;
  out.push('<table style="border-collapse:collapse;font-family:Calibri,Carlito,Arial,sans-serif;font-size:11pt"><colgroup>' + (opts.headings ? '<col style="width:34px">' : ''));
  for (let c = c1; c <= c2; c++) out.push('<col style="width:' + s.colWidth(c) + 'px">');
  out.push('</colgroup>');
  if (opts.headings) { out.push('<tr><td></td>'); for (let c = c1; c <= c2; c++) out.push('<td style="background:#f2f2f2;text-align:center;border:1px solid #ccc">' + colName(c) + '</td>'); out.push('</tr>'); }
  const covered = new Set();
  for (let r = r1; r <= r2; r++) {
    if (s.hiddenR[r]) continue;
    out.push('<tr style="height:' + s.rowHeight(r) + 'px">' + (opts.headings ? '<td style="background:#f2f2f2;text-align:center;border:1px solid #ccc">' + (r + 1) + '</td>' : ''));
    for (let c = c1; c <= c2; c++) {
      if (s.hiddenC[c]) continue;
      if (covered.has(r * 20000 + c)) continue;
      const m = s.mergeAt(r, c); let span = '';
      if (m && m.r1 === r && m.c1 === c) { span = (m.r2 > m.r1 ? ' rowspan="' + (m.r2 - m.r1 + 1) + '"' : '') + (m.c2 > m.c1 ? ' colspan="' + (m.c2 - m.c1 + 1) + '"' : ''); for (let rr = m.r1; rr <= m.r2; rr++) for (let cc = m.c1; cc <= m.c2; cc++) if (rr !== r || cc !== c) covered.add(rr * 20000 + cc); }
      const cell = s.get(r, c); const st = s.styleAt(r, c); const v = s.value(r, c);
      let txt = v == null ? '' : format(v, st && st.nf).t; const cf = doc && doc.cfStyle ? doc.cfStyle(s, r, c, v) : null;
      let style = grid + 'padding:1px 4px;' + css(st) + (cf && cf.fill ? ';background:' + cf.fill : '') + (cf && cf.color ? ';color:' + cf.color : '');
      if (typeof v === 'number' && !(st && st.ha)) style += ';text-align:right';
      out.push('<td' + span + ' style="' + style + '">' + esc(txt).replace(/\n/g, '<br>') + '</td>');
    }
    out.push('</tr>');
  }
  out.push('</table>');
  return out.join('');
}
