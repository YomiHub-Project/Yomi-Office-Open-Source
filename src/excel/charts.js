// SVG chart renderer (column, bar, line, area, pie, doughnut, scatter, stacked) shared by Excel, Word and PowerPoint
import { parseRangeAddr, Err } from './model.js';
import { format } from './numfmt.js';

export const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];
export const CHART_TYPES = [
  ['column', 'Cột nhóm (Clustered Column)', 'chart-column'], ['column-stacked', 'Cột chồng (Stacked Column)', 'chart-column-stacked'], ['bar', 'Thanh ngang (Bar)', 'chart-bar'], ['line', 'Đường (Line)', 'chart-line'], ['area', 'Vùng (Area)', 'chart-area'],
  ['pie', 'Tròn (Pie)', 'chart-pie'], ['doughnut', 'Vành khuyên (Doughnut)', 'circle-dot'], ['scatter', 'Phân tán (Scatter XY)', 'chart-scatter'], ['combo', 'Kết hợp cột + đường', 'chart-column']
];
const NS = 'http://www.w3.org/2000/svg';
function el(tag, attrs, parent) { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); if (parent) parent.appendChild(e); return e; }
function txt(parent, x, y, s, attrs) { const t = el('text', Object.assign({ x, y, 'font-size': 11, fill: '#595959', 'font-family': 'Segoe UI, Roboto, Arial, sans-serif' }, attrs || {}), parent); t.textContent = s; return t; }
function niceScale(min, max, ticks) {
  if (min === max) { max = min + 1; min = min - 1; }
  const range = max - min; const rough = range / (ticks || 5); const p = Math.pow(10, Math.floor(Math.log(rough) / Math.LN10)); const f = rough / p;
  const step = (f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10) * p;
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}
// extract chart data from a sheet range: {cats:[], series:[{name, values:[]}] , nf}
export function chartData(sheet, src, opts) {
  const rg = parseRangeAddr(src || ''); if (!rg) return { cats: [], series: [] };
  const v = (r, c) => { const x = sheet.value(r, c); return x instanceof Err ? null : x; };
  const firstRowText = (() => { let n = 0; for (let c = rg.c1; c <= rg.c2; c++) if (typeof v(rg.r1, c) === 'string') n++; return n >= Math.max(1, rg.c2 - rg.c1); })();
  const firstColText = (() => { for (let r = rg.r1 + (firstRowText ? 1 : 0); r <= rg.r2; r++) { const x = v(r, rg.c1); if (x != null && typeof x !== 'number') return true; } return false; })() || (opts && opts.kind === 'scatter');
  const byCol = !(opts && opts.byRow);
  const dr1 = rg.r1 + (firstRowText ? 1 : 0), dc1 = rg.c1 + (firstColText ? 1 : 0);
  const out = { cats: [], series: [] };
  const st = sheet.styleAt(dr1, dc1); out.nf = st && st.nf;
  const catNf = (sheet.styleAt(dr1, rg.c1) || {}).nf;
  if (byCol) {
    for (let r = dr1; r <= rg.r2; r++) out.cats.push(firstColText ? fmtCat(v(r, rg.c1), catNf) : String(r - dr1 + 1));
    for (let c = dc1; c <= rg.c2; c++) { const vals = []; for (let r = dr1; r <= rg.r2; r++) { const x = v(r, c); vals.push(typeof x === 'number' ? x : null); } out.series.push({ name: firstRowText ? String(v(rg.r1, c) == null ? '' : v(rg.r1, c)) : 'Chuỗi ' + (c - dc1 + 1), values: vals }); }
    if (opts && opts.kind === 'scatter') { out.xs = []; for (let r = dr1; r <= rg.r2; r++) { const x = v(r, rg.c1); out.xs.push(typeof x === 'number' ? x : r - dr1 + 1); } }
  } else {
    for (let c = dc1; c <= rg.c2; c++) out.cats.push(firstRowText ? fmtCat(v(rg.r1, c)) : String(c - dc1 + 1));
    for (let r = dr1; r <= rg.r2; r++) { const vals = []; for (let c = dc1; c <= rg.c2; c++) { const x = v(r, c); vals.push(typeof x === 'number' ? x : null); } out.series.push({ name: firstColText ? String(v(r, rg.c1)) : 'Chuỗi ' + (r - dr1 + 1), values: vals }); }
  }
  return out;
}
function fmtCat(x, nf) { if (x == null) return ''; if (typeof x === 'number' && nf) return format(x, nf).t; return String(x); }
function fmtVal(v, nf) { if (v == null) return ''; if (nf && !/[dmyhs]/i.test(nf.replace(/"[^"]*"/g, ''))) return format(v, nf).t; const a = Math.abs(v); if (a >= 1e9) return (v / 1e9).toFixed(1).replace(/\.0$/, '') + ' tỷ'; if (a >= 1e6) return (v / 1e6).toFixed(1).replace(/\.0$/, '') + ' tr'; if (a >= 1e4) return (v / 1e3).toFixed(0) + 'k'; return format(v, '#,##0.##').t.replace(/\.$/, ''); }

// render into container element (svg sized to container)
export function renderChart(container, data, o) {
  container.innerHTML = '';
  const W = Math.max(120, o.w || container.clientWidth || 480), H = Math.max(90, o.h || container.clientHeight || 300);
  const svg = el('svg', { width: W, height: H, viewBox: '0 0 ' + W + ' ' + H, xmlns: NS }, container);
  el('rect', { x: 0, y: 0, width: W, height: H, fill: o.bg || '#fff', stroke: o.border === false ? 'none' : '#d9d9d9' }, svg);
  const pal = o.colors || PALETTE;
  const kind = o.kind || 'column';
  let top = 10, bottom = H - 10, left = 10, right = W - 10;
  if (o.title !== '' && o.title !== false) { txt(svg, W / 2, 24, o.title || (data.series.length === 1 ? data.series[0].name : 'Tiêu đề biểu đồ'), { 'text-anchor': 'middle', 'font-size': 15, fill: '#404040' }); top = 38; }
  const series = data.series.filter(s => s.values.some(v => v != null));
  if (!series.length) { txt(svg, W / 2, H / 2, 'Không có dữ liệu số trong vùng nguồn', { 'text-anchor': 'middle' }); return svg; }
  // legend
  const showLegend = o.legend !== false && (series.length > 1 || kind === 'pie' || kind === 'doughnut');
  if (showLegend) {
    const items = kind === 'pie' || kind === 'doughnut' ? data.cats : series.map(s => s.name);
    if (kind === 'pie' || kind === 'doughnut') {
      const lx = W - Math.min(170, W * 0.35); let ly = top + 6; right = lx - 8;
      items.slice(0, 14).forEach((n, i) => { el('rect', { x: lx, y: ly, width: 10, height: 10, fill: pal[i % pal.length] }, svg); txt(svg, lx + 15, ly + 9, String(n).slice(0, 22)); ly += 17; });
    } else {
      let lx = left + 10, ly = bottom - 4; const parts = [];
      items.forEach((n, i) => parts.push([n, i]));
      const totalW = parts.reduce((a, p) => a + 26 + String(p[0]).slice(0, 20).length * 6.2, 0); lx = Math.max(left, (W - totalW) / 2);
      parts.forEach(([n, i]) => { el('rect', { x: lx, y: ly - 9, width: 10, height: 10, fill: pal[i % pal.length] }, svg); txt(svg, lx + 14, ly, String(n).slice(0, 20)); lx += 26 + String(n).slice(0, 20).length * 6.2; });
      bottom -= 22;
    }
  }
  if (kind === 'pie' || kind === 'doughnut') {
    const vals = series[0].values.map(v => Math.max(0, v || 0)); const tot = vals.reduce((a, b) => a + b, 0) || 1;
    const cx = (left + right) / 2, cy = (top + bottom) / 2, R = Math.max(10, Math.min(right - left, bottom - top) / 2 - 6), r0 = kind === 'doughnut' ? R * 0.55 : 0;
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + v / tot * Math.PI * 2; const large = a1 - a0 > Math.PI ? 1 : 0;
      const p = (a, rr) => [cx + rr * Math.cos(a), cy + rr * Math.sin(a)];
      const [x1, y1] = p(a0, R), [x2, y2] = p(a1, R), [x3, y3] = p(a1, r0), [x4, y4] = p(a0, r0);
      const d = vals.filter(x => x > 0).length === 1 && v > 0 ? 'M' + (cx - R) + ',' + cy + 'a' + R + ',' + R + ' 0 1,0 ' + 2 * R + ',0a' + R + ',' + R + ' 0 1,0 ' + (-2 * R) + ',0' : 'M' + x1 + ',' + y1 + 'A' + R + ',' + R + ' 0 ' + large + ',1 ' + x2 + ',' + y2 + (r0 ? 'L' + x3 + ',' + y3 + 'A' + r0 + ',' + r0 + ' 0 ' + large + ',0 ' + x4 + ',' + y4 + 'Z' : 'L' + cx + ',' + cy + 'Z');
      el('path', { d, fill: pal[i % pal.length], stroke: '#fff', 'stroke-width': 1.5 }, svg);
      if (v / tot > 0.04 && o.labels !== false) { const [lx, ly] = p((a0 + a1) / 2, r0 ? (R + r0) / 2 : R * 0.65); txt(svg, lx, ly + 4, Math.round(v / tot * 100) + '%', { 'text-anchor': 'middle', fill: '#fff', 'font-weight': 'bold' }); }
      a0 = a1;
    });
    return svg;
  }
  const stacked = /stacked/.test(kind) || o.stacked;
  const horiz = kind === 'bar';
  // value range
  let vmin = 0, vmax = 0;
  if (stacked) { data.cats.forEach((_, i) => { let p = 0, n = 0; series.forEach(s => { const v = s.values[i] || 0; if (v >= 0) p += v; else n += v; }); vmax = Math.max(vmax, p); vmin = Math.min(vmin, n); }); }
  else series.forEach(s => s.values.forEach(v => { if (v == null) return; if (v > vmax) vmax = v; if (v < vmin) vmin = v; }));
  if (kind === 'line' || kind === 'scatter') { const all = [].concat.apply([], series.map(s => s.values.filter(v => v != null))); const mn = Math.min.apply(null, all); if (mn > 0 && mn > (vmax - mn) * 0.5) vmin = mn - (vmax - mn) * 0.1; }
  const sc = niceScale(vmin, vmax, 5);
  // axes area
  const labelW = Math.max.apply(null, [sc.min, sc.max].map(v => fmtVal(v, data.nf).length)) * 6.5 + 8;
  left += horiz ? Math.min(140, Math.max.apply(null, data.cats.map(c => String(c).length)) * 6.3 + 8) : labelW;
  bottom -= horiz ? 18 : 20;
  const pw = Math.max(10, right - left), ph = Math.max(10, bottom - top);
  const vy = v => horiz ? left + (v - sc.min) / (sc.max - sc.min) * pw : bottom - (v - sc.min) / (sc.max - sc.min) * ph;
  // gridlines + value labels
  for (let v = sc.min; v <= sc.max + sc.step / 2; v += sc.step) {
    const p = vy(v);
    if (horiz) { el('line', { x1: p, y1: top, x2: p, y2: bottom, stroke: '#e5e5e5' }, svg); txt(svg, p, bottom + 14, fmtVal(+v.toPrecision(10), data.nf), { 'text-anchor': 'middle', 'font-size': 10 }); }
    else { el('line', { x1: left, y1: p, x2: right, y2: p, stroke: '#e5e5e5' }, svg); txt(svg, left - 5, p + 4, fmtVal(+v.toPrecision(10), data.nf), { 'text-anchor': 'end', 'font-size': 10 }); }
  }
  const n = data.cats.length || 1;
  const band = (horiz ? ph : pw) / n;
  // category labels
  const every = Math.max(1, Math.ceil(n / Math.max(1, (horiz ? ph / 16 : pw / 60))));
  data.cats.forEach((c, i) => { if (i % every) return; if (horiz) txt(svg, left - 5, top + band * (i + 0.5) + 4, String(c).slice(0, 22), { 'text-anchor': 'end', 'font-size': 10 }); else if (kind !== 'scatter') txt(svg, left + band * (i + 0.5), bottom + 14, String(c).slice(0, 14), { 'text-anchor': 'middle', 'font-size': 10 }); });
  el('line', horiz ? { x1: left, y1: top, x2: left, y2: bottom, stroke: '#bfbfbf' } : { x1: left, y1: vy(Math.max(0, sc.min)), x2: right, y2: vy(Math.max(0, sc.min)), stroke: '#bfbfbf' }, svg);
  const zero = vy(Math.max(sc.min, Math.min(0, sc.max)));
  const isBar = kind === 'column' || kind === 'bar' || kind === 'column-stacked' || kind === 'combo';
  if (isBar) {
    const barSeries = kind === 'combo' ? series.slice(0, Math.max(1, series.length - 1)) : series;
    const groupW = band * 0.68; const bw = stacked ? groupW : groupW / barSeries.length;
    const pos = data.cats.map(() => 0), neg = data.cats.map(() => 0);
    barSeries.forEach((s, si) => s.values.forEach((v, i) => {
      if (v == null) return;
      let a, b;
      if (stacked) { if (v >= 0) { a = pos[i]; pos[i] += v; b = pos[i]; } else { a = neg[i]; neg[i] += v; b = neg[i]; } }
      else { a = 0; b = v; }
      const p1 = vy(Math.max(sc.min, Math.min(a, b))), p2 = vy(Math.min(sc.max, Math.max(a, b)));
      const off = band * i + (band - groupW) / 2 + (stacked ? 0 : bw * si);
      const r = horiz ? { x: Math.min(p1, p2), y: top + off, width: Math.abs(p2 - p1), height: bw * 0.92 } : { x: left + off, y: Math.min(p1, p2), width: bw * 0.92, height: Math.abs(p2 - p1) };
      const rect = el('rect', Object.assign(r, { fill: pal[si % pal.length] }), svg);
      el('title', {}, rect).textContent = s.name + ' — ' + data.cats[i] + ': ' + fmtVal(v, data.nf);
      if (o.labels) txt(svg, horiz ? r.x + r.width + 3 : r.x + r.width / 2, horiz ? r.y + r.height / 2 + 4 : r.y - 3, fmtVal(v, data.nf), { 'text-anchor': horiz ? 'start' : 'middle', 'font-size': 9 });
    }));
    if (kind !== 'combo') return svg;
  }
  const lineSeries = kind === 'combo' ? series.slice(-1) : series;
  const xs = kind === 'scatter' && data.xs ? data.xs : null;
  let xsc = null; if (xs) { xsc = niceScale(Math.min.apply(null, xs), Math.max.apply(null, xs), 6); for (let v = xsc.min; v <= xsc.max + xsc.step / 2; v += xsc.step) txt(svg, left + (v - xsc.min) / (xsc.max - xsc.min) * pw, bottom + 14, fmtVal(+v.toPrecision(10)), { 'text-anchor': 'middle', 'font-size': 10 }); }
  const px = i => xs ? left + (xs[i] - xsc.min) / (xsc.max - xsc.min) * pw : left + band * (i + 0.5);
  lineSeries.forEach((s, si0) => {
    const si = kind === 'combo' ? series.length - 1 : si0; const col = pal[si % pal.length];
    const pts = []; s.values.forEach((v, i) => { if (v != null) pts.push([px(i), vy(v)]); });
    if (!pts.length) return;
    if (kind === 'area') el('path', { d: 'M' + pts[0][0] + ',' + zero + pts.map(p => 'L' + p[0] + ',' + p[1]).join('') + 'L' + pts[pts.length - 1][0] + ',' + zero + 'Z', fill: col, 'fill-opacity': 0.55, stroke: col }, svg);
    else if (kind !== 'scatter') el('path', { d: pts.map((p, i) => (i ? 'L' : 'M') + p[0] + ',' + p[1]).join(''), fill: 'none', stroke: col, 'stroke-width': 2.25, 'stroke-linejoin': 'round' }, svg);
    pts.forEach((p, i) => { const c = el('circle', { cx: p[0], cy: p[1], r: kind === 'scatter' ? 4 : 3, fill: kind === 'scatter' ? col : '#fff', stroke: col, 'stroke-width': 1.5 }, svg); if (o.labels) txt(svg, p[0], p[1] - 7, fmtVal(s.values[i], data.nf), { 'text-anchor': 'middle', 'font-size': 9 }); });
  });
  return svg;
}
// rasterize SVG chart to PNG data URL (for DOCX/PPTX/print fallbacks)
export function chartToPng(svg, scale) {
  return new Promise(res => {
    const xml = new XMLSerializer().serializeToString(svg);
    const img = new Image(); const w = +svg.getAttribute('width'), hh = +svg.getAttribute('height'); scale = scale || 2;
    img.onload = () => { const c = document.createElement('canvas'); c.width = w * scale; c.height = hh * scale; const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(img, 0, 0, c.width, c.height); res(c.toDataURL('image/png')); };
    img.onerror = () => res(null);
    img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(xml)));
  });
}
