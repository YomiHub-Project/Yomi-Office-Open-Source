// Draw tab: pen / highlighter / eraser ink on an SVG layer (mouse + touch), shared by Word/Excel/PowerPoint
import { h, ic } from './util.js';
import { colorPicker } from './ui.js';

export const PENS = [
  { id: 'pen-black', name: 'Bút đen', color: '#000000', width: 2, opacity: 1 },
  { id: 'pen-red', name: 'Bút đỏ', color: '#e81123', width: 2, opacity: 1 },
  { id: 'pen-blue', name: 'Bút xanh', color: '#0063b1', width: 2, opacity: 1 },
  { id: 'pen-green', name: 'Bút xanh lá', color: '#107c10', width: 2, opacity: 1 },
  { id: 'hl-yellow', name: 'Bút tô sáng vàng', color: '#ffff00', width: 14, opacity: 0.45, hl: true },
  { id: 'hl-green', name: 'Bút tô sáng xanh', color: '#00ff00', width: 14, opacity: 0.4, hl: true },
  { id: 'pencil', name: 'Bút chì', color: '#5f5f5f', width: 1.2, opacity: 0.85 }
];
const NS = 'http://www.w3.org/2000/svg';
export function strokePath(pts) {
  if (!pts.length) return '';
  if (pts.length < 3) return 'M' + pts.map(p => p[0].toFixed(1) + ',' + p[1].toFixed(1)).join('L');
  let d = 'M' + pts[0][0].toFixed(1) + ',' + pts[0][1].toFixed(1);
  for (let i = 1; i < pts.length - 1; i++) { const mx = (pts[i][0] + pts[i + 1][0]) / 2, my = (pts[i][1] + pts[i + 1][1]) / 2; d += 'Q' + pts[i][0].toFixed(1) + ',' + pts[i][1].toFixed(1) + ' ' + mx.toFixed(1) + ',' + my.toFixed(1); }
  const l = pts[pts.length - 1]; return d + 'L' + l[0].toFixed(1) + ',' + l[1].toFixed(1);
}
export function renderInk(svg, strokes, scale) {
  while (svg.firstChild) svg.removeChild(svg.firstChild);
  (strokes || []).forEach((s, i) => {
    const p = document.createElementNS(NS, 'path');
    p.setAttribute('d', strokePath(s.pts.map(q => [q[0] * scale, q[1] * scale])));
    p.setAttribute('fill', 'none'); p.setAttribute('stroke', s.color); p.setAttribute('stroke-width', s.width * scale);
    p.setAttribute('stroke-opacity', s.opacity); p.setAttribute('stroke-linecap', 'round'); p.setAttribute('stroke-linejoin', 'round');
    if (s.hl) p.setAttribute('style', 'mix-blend-mode:multiply');
    p.setAttribute('data-i', i);
    svg.appendChild(p);
  });
}
// attach drawing to an element; getStrokes() returns array to mutate; toLocal(e)->[x,y] in unscaled units
export class InkTool {
  constructor(opts) { this.o = opts; this.tool = null; this.pen = PENS[0]; }
  setTool(tool, pen) { this.tool = tool; if (pen) this.pen = pen; this.o.onMode && this.o.onMode(tool); }
  attach(surface) {
    const down = e => {
      if (!this.tool) return;
      if (e.type === 'mousedown' && e.button !== 0) return;
      e.preventDefault(); e.stopPropagation();
      const touch = e.type === 'touchstart';
      const p0 = this.o.toLocal(touch ? e.touches[0] : e);
      if (this.tool === 'eraser') { this.erase(p0); }
      const stroke = this.tool === 'eraser' ? null : { color: this.pen.color, width: this.pen.width, opacity: this.pen.opacity, hl: !!this.pen.hl, pts: [p0] };
      if (stroke) { this.o.begin && this.o.begin(); this.o.getStrokes().push(stroke); }
      const mv = ev => { const p = this.o.toLocal(touch ? ev.touches[0] : ev); if (stroke) { stroke.pts.push(p); this.o.redraw(); } else this.erase(p); if (ev.cancelable) ev.preventDefault(); };
      const up = () => { document.removeEventListener(touch ? 'touchmove' : 'mousemove', mv); document.removeEventListener(touch ? 'touchend' : 'mouseup', up); if (stroke && stroke.pts.length === 1) stroke.pts.push([stroke.pts[0][0] + 0.5, stroke.pts[0][1] + 0.5]); this.o.commit && this.o.commit(); this.o.redraw(); };
      document.addEventListener(touch ? 'touchmove' : 'mousemove', mv, touch ? { passive: false } : false); document.addEventListener(touch ? 'touchend' : 'mouseup', up);
    };
    surface.addEventListener('mousedown', down, true);
    surface.addEventListener('touchstart', down, { passive: false, capture: true });
  }
  erase(p) {
    const strokes = this.o.getStrokes(); const r = 8 / (this.o.scale ? this.o.scale() : 1);
    for (let i = strokes.length - 1; i >= 0; i--) { if (strokes[i].pts.some(q => Math.abs(q[0] - p[0]) < r && Math.abs(q[1] - p[1]) < r)) { this.o.begin && this.o.begin(); strokes.splice(i, 1); this.o.redraw(); } }
  }
}
// standard Draw ribbon tab
export function drawTab(app, ink) {
  const penBtn = pen => ({ type: 'btn', big: true, label: pen.name.replace('Bút ', '').replace('tô sáng ', 'Tô '), icon: pen.hl ? 'highlighter' : pen.id === 'pencil' ? 'pencil' : 'pen-tool', iconCls: 'pen-' + pen.id, active: () => ink.tool === 'pen' && ink.pen.id === pen.id, run: () => { ink.setTool(ink.tool === 'pen' && ink.pen.id === pen.id ? null : 'pen', pen); app.refreshUi(); } });
  return { id: 'draw', label: 'Vẽ', en: 'Draw', groups: [
    { label: 'Công cụ vẽ', items: [
      { type: 'big', label: 'Chọn', icon: 'mouse-pointer-2', active: () => !ink.tool, run: () => { ink.setTool(null); app.refreshUi(); } },
      { type: 'big', label: 'Tẩy', icon: 'eraser', active: () => ink.tool === 'eraser', run: () => { ink.setTool(ink.tool === 'eraser' ? null : 'eraser'); app.refreshUi(); } }
    ].concat(PENS.map(penBtn)).concat([
      { type: 'color', label: 'Màu bút', icon: 'palette', colorBar: () => ink.pen.color, run: () => {}, pick: (a, c) => { ink.pen = Object.assign({}, ink.pen, { color: c, id: 'custom' }); ink.setTool('pen'); } },
      { type: 'drop', label: 'Độ dày', icon: 'minus', menu: () => [1, 2, 3, 5, 8, 12, 18].map(w => ({ label: w + ' pt', checked: ink.pen.width === w || undefined, run: () => { ink.pen = Object.assign({}, ink.pen, { width: w, id: 'custom' }); ink.setTool('pen'); } })) }
    ]) },
    { label: 'Chỉnh sửa', items: [
      { type: 'big', label: 'Xóa mọi nét vẽ', icon: 'trash-2', run: () => { app.clearInk && app.clearInk(); } },
      { type: 'big', label: (app.inkHidden ? 'Hiện' : 'Ẩn') + ' nét vẽ', icon: 'eye-off', active: () => app.inkHidden, run: () => { app.inkHidden = !app.inkHidden; app.renderInk && app.renderInk(); } }
    ] },
    { label: 'Trợ giúp', items: [{ type: 'btn', label: 'Vẽ bằng chuột hoặc chạm. Chọn "Chọn" để thoát chế độ vẽ.', icon: 'info', run: () => {} }] }
  ] };
}
