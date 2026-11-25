// v0.0.8 — Excel pagination: real page model (paper size, margins, scaling, print area, print titles, manual page breaks)
// used by Page Break Preview (grid), Print Preview (multi-page with navigation), printing and XLSX <rowBreaks>/<colBreaks>.
import { h, ic, $, clamp } from '../core/util.js';
import { toast, promptDlg, formDlg } from '../core/ui.js';
import { parseRangeAddr, rangeAddr, colName } from './model.js';
import { sheetToHtml } from './xlsx.js';

export const PAPER = { A4: [210, 297], A3: [297, 420], A5: [148, 210], Letter: [216, 279], Legal: [216, 356] };
export const MARGINS = { normal: [19, 18, 19, 18], wide: [25, 25, 25, 25], narrow: [13, 6, 13, 6] }; // mm: top right bottom left
const MM = 96 / 25.4;

// printable area of one page in CSS px (96 dpi)
export function paperBox(ps) {
  const P = PAPER[ps.size] || PAPER.A4; const land = ps.orientation === 'landscape';
  const W = (land ? P[1] : P[0]), H = (land ? P[0] : P[1]);
  const m = Array.isArray(ps.margins) ? ps.margins : (MARGINS[ps.margins] || MARGINS.normal);
  return { Wmm: W, Hmm: H, m, w: (W - m[1] - m[3]) * MM, h: (H - m[0] - m[2]) * MM - (ps.header ? 18 : 0) - (ps.footer ? 18 : 0) };
}
function titleRows(ps) { const t = String(ps.printTitles || '').replace(/\$/g, '').trim(); const m = /^(\d+)(?::(\d+))?$/.exec(t); if (!m) return null; const a = +m[1] - 1, b = (m[2] ? +m[2] : +m[1]) - 1; return a >= 0 && b >= a ? [a, b] : null; }
function titleCols(ps) { const t = String(ps.printTitlesCols || '').replace(/\$/g, '').trim().toUpperCase(); const m = /^([A-Z]+)(?::([A-Z]+))?$/.exec(t); if (!m) return null; const ci = s => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }; const a = ci(m[1]), b = ci(m[2] || m[1]); return b >= a ? [a, b] : null; }

// area to print: selection / print area / used range (always at least A1)
export function printArea(s, sel) {
  if (sel) return sel;
  const ps = s.pageSetup || {}; const pa = ps.printArea && parseRangeAddr(String(ps.printArea).split(',')[0]);
  if (pa) return pa;
  const u = s.used(); return { r1: Math.max(0, u.r1 || 0) === 0 ? 0 : u.r1, c1: 0, r2: Math.max(0, u.r2), c2: Math.max(0, u.c2) };
}

// split [a..b] into chunks that fit `avail` (size(i) px each, already scaled); manual = Set of indices that must start a new page
function chunks(a, b, size, avail, manual, reserve) {
  const out = []; let start = a, acc = reserve || 0;
  for (let i = a; i <= b; i++) {
    const sz = size(i);
    const forced = i > start && manual.has(i);
    if (forced || (i > start && acc + sz > avail + 0.5)) { out.push([start, i - 1, forced]); start = i; acc = reserve || 0; }
    acc += sz;
  }
  out.push([start, b, false]);
  return out;
}

// full page model of a sheet
export function computePages(s, opts) {
  opts = opts || {};
  const ps = s.pageSetup || (s.pageSetup = {});
  const area = printArea(s, opts.area);
  const box = paperBox(ps);
  const tr = titleRows(ps), tc = titleCols(ps);
  let totW = 0, totH = 0; for (let c = area.c1; c <= area.c2; c++) totW += s.colWidth(c); for (let r = area.r1; r <= area.r2; r++) totH += s.rowHeight(r);
  let scale = ps.scale ? clamp(+ps.scale / 100, 0.1, 4) : 1;
  if (ps.fit === 'width') scale = Math.min(1, box.w / Math.max(1, totW));
  else if (ps.fit === 'page') scale = Math.min(1, box.w / Math.max(1, totW), box.h / Math.max(1, totH));
  else if (ps.fit === 'cols1') scale = Math.min(1, box.w / Math.max(1, totW));
  else if (ps.fit === 'rows1') scale = Math.min(1, box.h / Math.max(1, totH));
  const manR = new Set((ps.fit === 'page' ? [] : ps.rowBreaks || []).filter(r => r > area.r1 && r <= area.r2));
  const manC = new Set((ps.fit === 'page' || ps.fit === 'width' ? [] : ps.colBreaks || []).filter(c => c > area.c1 && c <= area.c2));
  let tH = 0; if (tr) for (let r = tr[0]; r <= tr[1]; r++) tH += s.rowHeight(r) * scale;
  let tW = 0; if (tc) for (let c = tc[0]; c <= tc[1]; c++) tW += s.colWidth(c) * scale;
  const rowsC = chunks(area.r1, area.r2, r => s.rowHeight(r) * scale, box.h, manR, 0);
  // rows after the first page need room for repeated title rows
  if (tr && rowsC.length > 1) { const first = rowsC[0]; const rest = chunks(first[1] + 1, area.r2, r => s.rowHeight(r) * scale, box.h, manR, tH); rowsC.length = 1; rest.forEach(x => rowsC.push(x)); }
  const colsC = chunks(area.c1, area.c2, c => s.colWidth(c) * scale, box.w, manC, 0);
  if (tc && colsC.length > 1) { const first = colsC[0]; const rest = chunks(first[1] + 1, area.c2, c => s.colWidth(c) * scale, box.w, manC, tW); colsC.length = 1; rest.forEach(x => colsC.push(x)); }
  const pages = []; const overThenDown = ps.order === 'overThenDown';
  if (overThenDown) rowsC.forEach(R => colsC.forEach(C => pages.push({ r1: R[0], r2: R[1], c1: C[0], c2: C[1] })));
  else colsC.forEach(C => rowsC.forEach(R => pages.push({ r1: R[0], r2: R[1], c1: C[0], c2: C[1] })));
  pages.forEach((p, i) => { p.n = i + 1; p.titleRows = tr && (p.r1 > tr[1]) ? tr : null; p.titleCols = tc && (p.c1 > tc[1]) ? tc : null; });
  return {
    area, scale, box, pages,
    rowStarts: rowsC.slice(1).map(x => ({ at: x[0], manual: manR.has(x[0]) })),
    colStarts: colsC.slice(1).map(x => ({ at: x[0], manual: manC.has(x[0]) })),
    // legacy shape used by older callers: break indices + end sentinel
    rows: rowsC.slice(1).map(x => x[0]).concat([area.r2 + 1]), cols: colsC.slice(1).map(x => x[0]).concat([area.c2 + 1])
  };
}

// ---------------------------------------------------------------- manual page breaks (Bố trí Trang → Ngắt trang)
export function insertBreak(doc) {
  const s = doc.wb.cur; const r = s.sel.r, c = s.sel.c;
  if (r === 0 && c === 0) return toast('Chọn ô bên dưới/bên phải vị trí muốn ngắt trang (không phải ô A1)', 'err');
  doc.change('Chèn ngắt trang', s, 'props', () => {
    const ps = s.pageSetup; ps.rowBreaks = (ps.rowBreaks || []).slice(); ps.colBreaks = (ps.colBreaks || []).slice();
    if (r > 0 && ps.rowBreaks.indexOf(r) < 0) ps.rowBreaks.push(r);
    if (c > 0 && ps.colBreaks.indexOf(c) < 0) ps.colBreaks.push(c);
    ps.rowBreaks.sort((a, b) => a - b); ps.colBreaks.sort((a, b) => a - b);
    if (ps.fit === 'page') ps.fit = 'none';
  });
  s.showPageBreaks = true; doc.grid.redraw();
  toast('Đã chèn ngắt trang ' + (r > 0 ? 'trên hàng ' + (r + 1) : '') + (r > 0 && c > 0 ? ' và ' : '') + (c > 0 ? 'trước cột ' + colName(c) : ''));
}
export function removeBreak(doc) {
  const s = doc.wb.cur; const r = s.sel.r, c = s.sel.c; const ps = s.pageSetup;
  const hasR = (ps.rowBreaks || []).indexOf(r) >= 0, hasC = (ps.colBreaks || []).indexOf(c) >= 0;
  if (!hasR && !hasC) return toast('Ô hiện tại không nằm ngay sau một ngắt trang thủ công', 'err');
  doc.change('Loại bỏ ngắt trang', s, 'props', () => { ps.rowBreaks = (ps.rowBreaks || []).filter(x => x !== r); ps.colBreaks = (ps.colBreaks || []).filter(x => x !== c); });
  doc.grid.redraw();
}
export function resetBreaks(doc) {
  const s = doc.wb.cur; doc.change('Đặt lại tất cả ngắt trang', s, 'props', () => { s.pageSetup.rowBreaks = []; s.pageSetup.colBreaks = []; });
  doc.grid.redraw(); toast('Đã đặt lại tất cả ngắt trang');
}
// "Chia trang": split the print area every N rows (and/or N columns)
export async function splitEvery(doc) {
  const s = doc.wb.cur; const a = printArea(s);
  const v = await formDlg('Chia trang đều', [
    { key: 'rows', label: 'Số hàng mỗi trang (0 = tự động):', value: '40' },
    { key: 'cols', label: 'Số cột mỗi trang (0 = tự động):', value: '0' }
  ], { width: 380 }); if (!v) return;
  const nr = Math.max(0, parseInt(v.rows) || 0), nc = Math.max(0, parseInt(v.cols) || 0);
  doc.change('Chia trang', s, 'props', () => {
    const ps = s.pageSetup; ps.rowBreaks = []; ps.colBreaks = [];
    if (nr) for (let r = a.r1 + nr; r <= a.r2; r += nr) ps.rowBreaks.push(r);
    if (nc) for (let c = a.c1 + nc; c <= a.c2; c += nc) ps.colBreaks.push(c);
    if (ps.fit === 'page' || (nc && ps.fit === 'width')) ps.fit = 'none';
  });
  s.showPageBreaks = true; doc.grid.redraw();
  toast('Đã chia thành ' + computePages(s).pages.length + ' trang');
}
export function setPrintArea(doc, range) {
  const s = doc.wb.cur;
  const rg = range || doc.grid.selRange();
  const addr = rangeAddr(rg);
  doc.change('Đặt vùng in', s, 'props', () => {
    if (!s.pageSetup) s.pageSetup = {};
    s.pageSetup.printArea = addr;
  });
  s.showPageBreaks = true;
  doc.grid.redraw();
  if (doc.refreshUi) doc.refreshUi();
  doc.markDirty();
  toast('Đã đặt vùng in: ' + addr);
}

export function addToPrintArea(doc, range) {
  const s = doc.wb.cur;
  const rg = range || doc.grid.selRange();
  const newAddr = rangeAddr(rg);
  doc.change('Thêm vào vùng in', s, 'props', () => {
    if (!s.pageSetup) s.pageSetup = {};
    const cur = s.pageSetup.printArea;
    s.pageSetup.printArea = cur ? (cur + ',' + newAddr) : newAddr;
  });
  s.showPageBreaks = true;
  doc.grid.redraw();
  if (doc.refreshUi) doc.refreshUi();
  doc.markDirty();
  toast('Đã thêm vào vùng in: ' + newAddr);
}

export function clearPrintArea(doc) {
  const s = doc.wb.cur;
  doc.change('Xóa vùng in', s, 'props', () => {
    if (s.pageSetup) s.pageSetup.printArea = null;
  });
  doc.grid.redraw();
  if (doc.refreshUi) doc.refreshUi();
  doc.markDirty();
  toast('Đã xóa vùng in (in toàn bộ trang tính)');
}

export function printAreaMenu(doc) {
  return [
    { label: 'Đặt vùng in (vùng chọn)', icon: 'crop', run: () => setPrintArea(doc) },
    { label: 'Thêm vào vùng in', icon: 'plus-square', run: () => addToPrintArea(doc) },
    { label: 'Xóa vùng in', icon: 'trash-2', run: () => clearPrintArea(doc) },
    { sep: true },
    { label: 'Xem trước Ngắt trang & Vùng in', icon: 'square-dashed', checked: !!doc.wb.cur.showPageBreaks || undefined, run: () => { doc.wb.cur.showPageBreaks = !doc.wb.cur.showPageBreaks; doc.grid.redraw(); if (doc.refreshUi) doc.refreshUi(); } }
  ];
}

export function breaksMenu(doc) {
  return [
    { label: 'Chèn ngắt trang', icon: 'separator-horizontal', run: () => insertBreak(doc) },
    { label: 'Loại bỏ ngắt trang', icon: 'minus', run: () => removeBreak(doc) },
    { label: 'Đặt lại tất cả ngắt trang', icon: 'rotate-ccw', run: () => resetBreaks(doc) },
    { sep: true },
    { label: 'Chia trang đều theo số hàng/cột…', icon: 'rows-3', run: () => splitEvery(doc) },
    { label: 'Xem trước Ngắt trang (Bật/Tắt)', icon: 'square-dashed', checked: !!doc.wb.cur.showPageBreaks || undefined, run: () => { doc.wb.cur.showPageBreaks = !doc.wb.cur.showPageBreaks; doc.grid.redraw(); if (doc.refreshUi) doc.refreshUi(); } }
  ];
}

// ---------------------------------------------------------------- hit-testing for mouse dragging in Page Break Preview
export function hitPageBreak(grid, p) {
  const s = grid.sheet;
  if (!s || !s.showPageBreaks) return null;
  let pb;
  try { pb = computePages(s); } catch (e) { return null; }
  if (!pb || !pb.area) return null;
  const a = pb.area;
  const rStart = grid.cellRect(a.r1, a.c1);
  const rEnd = grid.cellRect(a.r2, a.c2);
  const x1 = rStart.x, y1 = rStart.y;
  const x2 = rEnd.x + rEnd.w, y2 = rEnd.y + rEnd.h;

  const tol = 6;
  const xMid = (x1 + x2) / 2, yMid = (y1 + y2) / 2;

  // 1. check resize handles of print area (corners & midpoints)
  const handles = [
    { part: 'corner-tl', x: x1, y: y1, cursor: 'nwse-resize' },
    { part: 'corner-tr', x: x2, y: y1, cursor: 'nesw-resize' },
    { part: 'corner-bl', x: x1, y: y2, cursor: 'nesw-resize' },
    { part: 'corner-br', x: x2, y: y2, cursor: 'nwse-resize' },
    { part: 'edge-top', x: xMid, y: y1, cursor: 'row-resize' },
    { part: 'edge-bottom', x: xMid, y: y2, cursor: 'row-resize' },
    { part: 'edge-left', x: x1, y: yMid, cursor: 'col-resize' },
    { part: 'edge-right', x: x2, y: yMid, cursor: 'col-resize' }
  ];
  for (const h of handles) {
    if (Math.abs(p.x - h.x) <= tol + 3 && Math.abs(p.y - h.y) <= tol + 3) {
      return { type: 'printArea', part: h.part, cursor: h.cursor, area: a, pb };
    }
  }

  // 2. check print area borders
  if (Math.abs(p.y - y1) <= tol && p.x >= x1 - tol && p.x <= x2 + tol) {
    return { type: 'printArea', part: 'edge-top', cursor: 'row-resize', area: a, pb };
  }
  if (Math.abs(p.y - y2) <= tol && p.x >= x1 - tol && p.x <= x2 + tol) {
    return { type: 'printArea', part: 'edge-bottom', cursor: 'row-resize', area: a, pb };
  }
  if (Math.abs(p.x - x1) <= tol && p.y >= y1 - tol && p.y <= y2 + tol) {
    return { type: 'printArea', part: 'edge-left', cursor: 'col-resize', area: a, pb };
  }
  if (Math.abs(p.x - x2) <= tol && p.y >= y1 - tol && p.y <= y2 + tol) {
    return { type: 'printArea', part: 'edge-right', cursor: 'col-resize', area: a, pb };
  }

  // 3. check horizontal row breaks
  if (pb.rowStarts && pb.rowStarts.length) {
    for (const b of pb.rowStarts) {
      const ry = grid.cellRect(b.at, a.c1).y;
      if (Math.abs(p.y - ry) <= tol && p.x >= x1 - tol && p.x <= x2 + tol) {
        return { type: 'rowBreak', at: b.at, manual: b.manual, cursor: 'row-resize', y: ry, area: a, pb };
      }
    }
  }

  // 4. check vertical col breaks
  if (pb.colStarts && pb.colStarts.length) {
    for (const b of pb.colStarts) {
      const cx = grid.cellRect(a.r1, b.at).x;
      if (Math.abs(p.x - cx) <= tol && p.y >= y1 - tol && p.y <= y2 + tol) {
        return { type: 'colBreak', at: b.at, manual: b.manual, cursor: 'col-resize', x: cx, area: a, pb };
      }
    }
  }

  return null;
}

// ---------------------------------------------------------------- grid overlay for Page Break Preview
export function drawBreaks(grid, ctx, X, Y, R, C) {
  const s = grid.sheet; let pb; try { pb = computePages(s); } catch (e) { return; }
  const a = pb.area, z = grid.z || 1;
  const ax1 = X(Math.max(a.c1, C.c1)), ax2 = X(Math.min(a.c2, C.c2) + 1), ay1 = Y(Math.max(a.r1, R.r1)), ay2 = Y(Math.min(a.r2, R.r2) + 1);
  const vx1 = X(C.c1), vx2 = X(C.c2 + 1), vy1 = Y(R.r1), vy2 = Y(R.r2 + 1);

  // shade everything outside the print area (like Excel Page Break Preview)
  ctx.fillStyle = 'rgba(128,128,128,0.28)';
  if (a.r2 < R.r1 || a.r1 > R.r2 || a.c2 < C.c1 || a.c1 > C.c2) {
    ctx.fillRect(vx1, vy1, vx2 - vx1, vy2 - vy1);
  } else {
    if (ay1 > vy1) ctx.fillRect(vx1, vy1, vx2 - vx1, ay1 - vy1);
    if (ay2 < vy2) ctx.fillRect(vx1, ay2, vx2 - vx1, vy2 - ay2);
    if (ax1 > vx1) ctx.fillRect(vx1, ay1, ax1 - vx1, ay2 - ay1);
    if (ax2 < vx2) ctx.fillRect(ax2, ay1, vx2 - ax2, ay2 - ay1);
  }

  // page number watermark
  ctx.save(); ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillStyle = 'rgba(43,87,154,0.30)';
  pb.pages.forEach(p => {
    if (p.r2 < R.r1 || p.r1 > R.r2 || p.c2 < C.c1 || p.c1 > C.c2) return;
    const x1 = X(p.c1), x2 = X(p.c2 + 1), y1 = Y(p.r1), y2 = Y(p.r2 + 1);
    const fs = Math.max(14, Math.min(64 * z, (x2 - x1) / 4, (y2 - y1) / 3));
    ctx.font = 'bold ' + Math.round(fs) + 'px Calibri, Arial, sans-serif';
    ctx.fillText('Trang ' + p.n, (x1 + x2) / 2, (y1 + y2) / 2);
  });
  ctx.restore();

  const hit = grid.pageBreakHit;
  const ds = grid.dragState;

  // Print area boundary box
  const boxX1 = X(a.c1), boxY1 = Y(a.r1);
  const boxX2 = X(a.c2 + 1), boxY2 = Y(a.r2 + 1);
  const boxW = boxX2 - boxX1, boxH = boxY2 - boxY1;

  if (!(a.r2 < R.r1 || a.r1 > R.r2 || a.c2 < C.c1 || a.c1 > C.c2)) {
    ctx.save();
    ctx.strokeStyle = '#2b579a';
    ctx.lineWidth = 2.5;
    ctx.setLineDash([]);
    ctx.strokeRect(boxX1 + 1, boxY1 + 1, boxW - 2, boxH - 2);

    // Draw 8 resize handles on print area border
    const hs = Math.max(6, Math.min(10, Math.round(7 * z)));
    const hList = [
      [boxX1, boxY1], [(boxX1 + boxX2) / 2, boxY1], [boxX2, boxY1],
      [boxX1, (boxY1 + boxY2) / 2],                   [boxX2, (boxY1 + boxY2) / 2],
      [boxX1, boxY2], [(boxX1 + boxX2) / 2, boxY2], [boxX2, boxY2]
    ];
    ctx.fillStyle = '#ffffff';
    ctx.strokeStyle = '#1a5fb4';
    ctx.lineWidth = 2;
    for (const [hx, hy] of hList) {
      if (hx >= vx1 - 10 && hx <= vx2 + 10 && hy >= vy1 - 10 && hy <= vy2 + 10) {
        ctx.fillRect(hx - hs / 2, hy - hs / 2, hs, hs);
        ctx.strokeRect(hx - hs / 2 + 0.5, hy - hs / 2 + 0.5, hs, hs);
      }
    }
    ctx.restore();
  }

  // Draw row breaks
  pb.rowStarts.forEach(b => {
    if (b.at < R.r1 || b.at > R.r2 + 1) return;
    const y = Math.round(Y(b.at)) + 0.5;
    const isHovered = hit && hit.type === 'rowBreak' && hit.at === b.at;
    ctx.save();
    ctx.strokeStyle = isHovered ? '#0078d4' : '#2b579a';
    ctx.lineWidth = isHovered ? 4 : (b.manual ? 2.5 : 1.5);
    ctx.setLineDash(b.manual ? [] : [7, 4]);
    ctx.beginPath();
    ctx.moveTo(boxX1, y);
    ctx.lineTo(boxX2, y);
    ctx.stroke();
    ctx.restore();
  });

  // Draw col breaks
  pb.colStarts.forEach(b => {
    if (b.at < C.c1 || b.at > C.c2 + 1) return;
    const x = Math.round(X(b.at)) + 0.5;
    const isHovered = hit && hit.type === 'colBreak' && hit.at === b.at;
    ctx.save();
    ctx.strokeStyle = isHovered ? '#0078d4' : '#2b579a';
    ctx.lineWidth = isHovered ? 4 : (b.manual ? 2.5 : 1.5);
    ctx.setLineDash(b.manual ? [] : [7, 4]);
    ctx.beginPath();
    ctx.moveTo(x, boxY1);
    ctx.lineTo(x, boxY2);
    ctx.stroke();
    ctx.restore();
  });

  // Active dragging indicators
  if (ds) {
    ctx.save();
    if (ds.type === 'pageBreakRow') {
      const dy = Math.round(Y(ds.targetRow)) + 0.5;
      ctx.strokeStyle = '#0078d4';
      ctx.lineWidth = 3;
      ctx.setLineDash([5, 3]);
      ctx.beginPath();
      ctx.moveTo(boxX1, dy);
      ctx.lineTo(boxX2, dy);
      ctx.stroke();
    } else if (ds.type === 'pageBreakCol') {
      const dx = Math.round(X(ds.targetCol)) + 0.5;
      ctx.strokeStyle = '#0078d4';
      ctx.lineWidth = 3;
      ctx.setLineDash([5, 3]);
      ctx.beginPath();
      ctx.moveTo(dx, boxY1);
      ctx.lineTo(dx, boxY2);
      ctx.stroke();
    } else if (ds.type === 'printAreaResize' && ds.preview) {
      const pv = ds.preview;
      const px1 = X(pv.c1), py1 = Y(pv.r1);
      const px2 = X(pv.c2 + 1), py2 = Y(pv.r2 + 1);
      ctx.fillStyle = 'rgba(0, 120, 212, 0.14)';
      ctx.fillRect(px1, py1, px2 - px1, py2 - py1);
      ctx.strokeStyle = '#0078d4';
      ctx.lineWidth = 3;
      ctx.setLineDash([6, 3]);
      ctx.strokeRect(px1 + 1, py1 + 1, px2 - px1 - 2, py2 - py1 - 2);
    }
    ctx.restore();
  }
}

// ---------------------------------------------------------------- page HTML (print + preview)
function hf(text, s, n, total) { return String(text || '').replace(/&\[Page\]/g, n).replace(/&\[Pages\]/g, total).replace(/&\[Tab\]/g, s.name).replace(/&\[Date\]/g, new Date().toLocaleDateString('vi-VN')).replace(/&\[Time\]/g, new Date().toLocaleTimeString('vi-VN')).replace(/&\[File\]/g, ''); }
export function pageHtml(doc, s, p, pb, total) {
  const ps = s.pageSetup || {}; const o = { print: true, gridlines: ps.gridlines, headings: ps.headings };
  const cs = p.titleCols ? [{ c1: p.titleCols[0], c2: p.titleCols[1] }, { c1: p.c1, c2: p.c2 }] : [{ c1: p.c1, c2: p.c2 }];
  const rs = p.titleRows ? [{ r1: p.titleRows[0], r2: p.titleRows[1] }, { r1: p.r1, r2: p.r2 }] : [{ r1: p.r1, r2: p.r2 }];
  // rows of tables side by side when title columns repeat
  const body = rs.map(R => '<div class="xl-pg-row">' + cs.map(C => sheetToHtml(s, doc, Object.assign({}, o, { area: { r1: R.r1, r2: R.r2, c1: C.c1, c2: C.c2 } }))).join('') + '</div>').join('');
  const head = ps.header ? '<div class="xl-pg-hf">' + escH(hf(ps.header, s, p.n, total)) + '</div>' : '';
  const foot = ps.footer ? '<div class="xl-pg-hf xl-pg-foot">' + escH(hf(ps.footer, s, p.n, total)) + '</div>' : '';
  return head + '<div class="xl-pg-body" style="zoom:' + pb.scale.toFixed(4) + ';' + (ps.centerH ? 'text-align:center;' : '') + '">' + body + '</div>' + foot;
}
function escH(t) { return String(t).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])); }

export function renderPrint(doc, area) {
  const s0 = doc.wb.cur; const ps0 = s0.pageSetup;
  const sheets = ps0.what === 'book' ? doc.wb.sheets.filter(x => !x.hidden) : [s0];
  const box = paperBox(ps0);
  area.appendChild(h('style', '@page { size: ' + box.Wmm + 'mm ' + box.Hmm + 'mm; margin: ' + box.m.map(x => x + 'mm').join(' ') + '; } #print-area .xl-pg { page-break-after: always; break-after: page; overflow: hidden; } #print-area .xl-pg:last-child { page-break-after: auto; break-after: auto; } #print-area table { border-collapse: collapse; } #print-area .xl-pg-row { white-space: nowrap; } #print-area .xl-pg-row > table { display: inline-table; vertical-align: top; } #print-area .xl-pg-hf { text-align: center; font: 10pt Calibri, Arial, sans-serif; height: 18px; } #print-area .xl-pg-foot { margin-top: 4px; }'));
  sheets.forEach(sh => {
    const pb = computePages(sh, { area: ps0.what === 'selection' && sh === s0 ? doc.grid.selRange() : null });
    pb.pages.forEach(p => { const d = h('div.xl-pg'); d.innerHTML = pageHtml(doc, sh, p, pb, pb.pages.length); area.appendChild(d); });
  });
}

// multi-page print preview with page navigation (backstage → In)
export function printPreview(doc, box) {
  box.innerHTML = '';
  const s = doc.wb.cur; const ps = s.pageSetup; const pb = computePages(s, { area: ps.what === 'selection' ? doc.grid.selRange() : null });
  const pp = paperBox(ps); const k = Math.min(2.4, 560 / pp.Hmm, 440 / pp.Wmm);
  let cur = clamp(doc._ppPage || 0, 0, pb.pages.length - 1);
  const page = h('div.xl-pp-page', { style: { width: pp.Wmm * k + 'px', height: pp.Hmm * k + 'px', position: 'relative', overflow: 'hidden' } });
  const inner = h('div.xl-pp-inner', { style: { position: 'absolute', left: pp.m[3] * k + 'px', top: pp.m[0] * k + 'px', width: pp.w + 'px', transformOrigin: '0 0', WebkitTransformOrigin: '0 0', transform: 'scale(' + (k / MM) + ')', WebkitTransform: 'scale(' + (k / MM) + ')' } });
  page.appendChild(inner);
  const lbl = h('span.xl-pp-lbl');
  const show = () => { cur = clamp(cur, 0, pb.pages.length - 1); doc._ppPage = cur; inner.innerHTML = pageHtml(doc, s, pb.pages[cur], pb, pb.pages.length); lbl.textContent = 'Trang ' + (cur + 1) + ' / ' + pb.pages.length + (pb.scale < 0.999 || pb.scale > 1.001 ? ' · tỷ lệ ' + Math.round(pb.scale * 100) + '%' : ''); };
  const nav = h('div.xl-pp-nav',
    h('button.btn', { title: 'Trang đầu', onclick: () => { cur = 0; show(); } }, ic('chevrons-left', 14)),
    h('button.btn', { title: 'Trang trước', onclick: () => { cur--; show(); } }, ic('chevron-left', 14)), lbl,
    h('button.btn', { title: 'Trang sau', onclick: () => { cur++; show(); } }, ic('chevron-right', 14)),
    h('button.btn', { title: 'Trang cuối', onclick: () => { cur = pb.pages.length - 1; show(); } }, ic('chevrons-right', 14)));
  box.appendChild(page); box.appendChild(nav); show();
}
export function pageCount(doc) { try { return computePages(doc.wb.cur).pages.length; } catch (e) { return 1; } }
