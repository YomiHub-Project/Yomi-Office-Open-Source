// PowerPoint commands: slide ops, object editing, text, shapes, transitions, animations, slideshow, print
import { h, ic, esc, $, $$, clamp, uid, deepClone, keyName, chord, readFileAsDataUrl, bytesToDataUrl, KaTeX } from '../core/util.js';
import { toast, showMenu, contextMenu, dialog, promptDlg, confirmDlg, formDlg, colorPicker } from '../core/ui.js';
import { settings } from '../core/store.js';
import Host from '../core/host.js';
import { SW, SH } from './ppt.js';
import { shapeSvg } from '../core/shapes.js';
import { renderChart, chartData } from '../excel/charts.js';
import { THEMES, LAYOUTS, blankSlide, layoutObjects } from './templates.js';

// ---------------------------------------------------------------- rendering a slide to a container
export function renderSlideInto(doc, s, container, scale, interactive) {
  container.innerHTML = '';
  const theme = doc.theme;
  const bg = s.bg || theme.bg;
  container.style.background = bgCss(bg, theme);
  if (theme.accentBar && s.layout !== 'blank') { const bar = h('div.pp-accentbar', { style: { background: theme.colors[0], left: 0, top: 0, bottom: 0, width: (8 * scale) + 'px', position: 'absolute' } }); container.appendChild(bar); }
  s.objects.forEach(o => container.appendChild(renderObj(doc, s, o, scale, interactive)));
  if (interactive) {
    container.onmousedown = container.ontouchstart = e => { if (e.target === container || e.target.classList.contains('pp-accentbar')) { doc.selectedObj = null; deselectAll(container); doc.onSelection(); } };
  }
}
function bgCss(bg, theme) {
  if (!bg) return '#fff';
  if (typeof bg === 'string') return bg;
  if (bg.type === 'gradient') return 'linear-gradient(' + (bg.angle || 135) + 'deg,' + bg.stops.join(',') + ')';
  if (bg.type === 'image') return 'center/cover no-repeat url(' + bg.src + ')';
  return bg.color || '#fff';
}
function renderObj(doc, s, o, scale, interactive) {
  const el = h('div.pp-obj' + (o.type === 'text' ? '.pp-text' : ''), { dataset: { id: o.id } });
  el.style.left = o.x * scale + 'px'; el.style.top = o.y * scale + 'px'; el.style.width = o.w * scale + 'px'; el.style.height = o.h * scale + 'px';
  if (o.rot) el.style.webkitTransform = el.style.transform = 'rotate(' + o.rot + 'deg)';
  if (o.opacity != null) el.style.opacity = o.opacity;
  const body = h('div.pp-obj-body'); el.appendChild(body);
  if (o.type === 'text') {
    const t = h('div.pp-tf', { style: textStyle(doc, o, scale) });
    t.innerHTML = o.html || '';
    if (interactive) { t.contentEditable = 'false'; }
    if (!o.html && o.ph) { t.classList.add('ph'); t.innerHTML = o.ph === 'title' ? 'Bấm để thêm tiêu đề' : o.ph === 'subtitle' ? 'Bấm để thêm phụ đề' : 'Bấm để thêm văn bản'; }
    body.appendChild(t);
  } else if (o.type === 'shape') {
    body.innerHTML = shapeSvg(o.shape, o.w * scale, o.h * scale, o.fill || doc.theme.colors[0], o.stroke || 'none', (o.strokeW || 0) * scale);
    if (o.text) { const t = h('div.pp-shape-text', { style: { fontSize: (o.fontSize || 18) * scale + 'px', color: o.color || '#fff', fontWeight: o.b ? 'bold' : '' } }); t.innerHTML = o.text; body.appendChild(t); }
  } else if (o.type === 'image') { const img = h('img', { src: o.src, alt: o.alt || '', draggable: false }); img.style.width = '100%'; img.style.height = '100%'; img.style.objectFit = o.fit || 'contain'; if (o.filter) img.style.filter = o.filter; body.appendChild(img); if (o.border) body.style.border = o.border; }
  else if (o.type === 'chart') { renderChart(body, o.data || sampleChart(), { kind: o.kind || 'column', title: o.title, w: o.w * scale, h: o.h * scale, colors: doc.theme.colors, legend: o.legend }); }
  else if (o.type === 'table') body.appendChild(renderTable(o, scale));
  else if (o.type === 'equation') { const eq = h('div.pp-eq', { style: { fontSize: 30 * scale + 'px' } }); body.appendChild(eq); KaTeX().then(k => { try { k.render(o.latex || '', eq, { throwOnError: false, displayMode: true }); } catch (e) {} }); }
  else if (o.type === 'video') { if (o.src && scale > 0.3) { const v = h(o.audio ? 'audio' : 'video', { src: o.src, controls: true, preload: 'metadata', style: { width: '100%', height: o.audio ? 'auto' : '100%', background: o.audio ? 'transparent' : '#000' } }); body.appendChild(v); } else body.appendChild(h('div.pp-video', ic(o.audio ? 'volume-2' : 'play-circle', 48 * scale), h('div', o.name || 'Video'))); }
  if (interactive) bindObj(doc, s, o, el, body);
  return el;
}
function sampleChart() { return { cats: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Doanh thu', values: [120, 150, 138, 170] }] }; }
function textStyle(doc, o, scale) {
  const st = {};
  st.fontSize = (o.fontSize || (o.ph === 'title' ? 40 : o.ph === 'subtitle' ? 24 : 18)) * scale + 'px';
  st.fontFamily = o.font || doc.theme.font;
  st.color = o.color || (o.ph === 'title' ? doc.theme.titleColor : doc.theme.bodyColor);
  st.textAlign = o.align || (o.ph === 'title' ? 'left' : 'left');
  st.fontWeight = o.b || o.ph === 'title' ? 'bold' : 'normal';
  if (o.i) st.fontStyle = 'italic';
  st.lineHeight = o.lh || 1.2;
  st.justifyContent = o.valign === 'middle' ? 'center' : o.valign === 'bottom' ? 'flex-end' : 'flex-start';
  return st;
}
function renderTable(o, scale) {
  const tb = h('table.pp-tbl'); const theme = o.style || {};
  (o.rows || []).forEach((row, r) => { const tr = h('tr'); row.forEach(cell => { const td = h(r === 0 && o.header ? 'th' : 'td', { style: { fontSize: (o.fontSize || 16) * scale + 'px', padding: (6 * scale) + 'px' } }); td.innerHTML = cell || ''; tr.appendChild(td); }); tb.appendChild(tr); });
  return tb;
}

// ---------------------------------------------------------------- object interaction
function deselectAll(container) { $$('.pp-obj.sel', container).forEach(e => { e.classList.remove('sel'); $$('.pp-h', e).forEach(x => x.remove()); }); }
export function addHandles(el) { $$('.pp-h', el).forEach(x => x.remove()); ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].forEach(d => el.appendChild(h('div.pp-h.h-' + d, { dataset: { d } }))); el.appendChild(h('div.pp-h.rot', { dataset: { d: 'rot' } }, ic('rotate-cw', 10))); }
function isTextual(o) { return o.type === 'text' || o.type === 'shape'; }
// v0.0.6: PowerPoint-style text editing — a click (without dragging) on a text box / shape places the caret right where you clicked,
// a click on a placeholder clears the prompt text, dragging the box border/inside moves it, typing on a selected box starts editing.
// (v0.0.5 bug: mouseup always re-rendered the whole slide, destroying the element between the 2 clicks of a double-click → typing was impossible.)
function bindObj(doc, s, o, el, body) {
  const select = () => { if (doc.selectedObj === o && el.classList.contains('sel')) return false; doc.selectedObj = o; deselectAll(doc.canvas); el.classList.add('sel'); addHandles(el); doc.onSelection(); return true; };
  const start = e => {
    if (e.type === 'mousedown' && e.button !== 0) return;
    if (el.classList.contains('editing') && !(e.target.classList && e.target.classList.contains('pp-h'))) { e.stopPropagation(); return; }
    e.stopPropagation();
    const wasSel = !select();
    const d = e.target.getAttribute && e.target.getAttribute('data-d');
    const touch = e.type === 'touchstart'; const p0 = touch ? e.touches[0] : e; const sc = doc.scale;
    if (!touch) e.preventDefault();
    const o0 = { x: o.x, y: o.y, w: o.w, h: o.h, rot: o.rot || 0 }; const ratio = o.w / o.h;
    const cx = o.x + o.w / 2, cy = o.y + o.h / 2; let moved = false, hist = false;
    const mv = ev => {
      const p = touch ? ev.touches[0] : ev; const dx = (p.clientX - p0.clientX) / sc, dy = (p.clientY - p0.clientY) / sc;
      if (!moved && Math.abs(dx) * sc < 4 && Math.abs(dy) * sc < 4) return;
      if (!moved) { moved = true; if (!hist) { doc.pushHistory(); hist = true; } }
      if (d === 'rot') { const r = doc.canvas.getBoundingClientRect(); const ang = Math.atan2(p.clientY - (r.top + cy * sc), p.clientX - (r.left + cx * sc)) * 180 / Math.PI + 90; o.rot = Math.round(ang / (ev.shiftKey ? 1 : 15)) * (ev.shiftKey ? 1 : 15); }
      else if (!d) { o.x = o0.x + dx; o.y = o0.y + dy; if (!ev.altKey) { o.x = snap(o.x, [0, SW / 2 - o.w / 2, SW - o.w]); o.y = snap(o.y, [0, SH / 2 - o.h / 2, SH - o.h]); } }
      else { let x = o0.x, y = o0.y, w = o0.w, hh = o0.h; if (/e/.test(d)) w = Math.max(20, o0.w + dx); if (/s/.test(d)) hh = Math.max(16, o0.h + dy); if (/w/.test(d)) { w = Math.max(20, o0.w - dx); x = o0.x + o0.w - w; } if (/n/.test(d)) { hh = Math.max(16, o0.h - dy); y = o0.y + o0.h - hh; } if ((ev.shiftKey || o.type === 'image') && d.length === 2) { hh = w / ratio; if (/n/.test(d)) y = o0.y + o0.h - hh; } Object.assign(o, { x, y, w, h: hh }); }
      place(el, o, sc);
      if (ev.cancelable) ev.preventDefault();
    };
    const up = ev => {
      document.removeEventListener(touch ? 'touchmove' : 'mousemove', mv); document.removeEventListener(touch ? 'touchend' : 'mouseup', up);
      if (moved) {
        doc.pushHistory(); doc.markDirty(); doc.renderThumbs();
        if (o.type === 'chart' || o.type === 'shape' || d === 'rot') { doc.renderSlide(); const ne = doc.canvas.querySelector('[data-id="' + o.id + '"]'); if (ne) { doc.selectedObj = o; ne.classList.add('sel'); addHandles(ne); } }
        return;
      }
      // plain click: text boxes enter edit mode immediately (placeholder) or on the 2nd click (like PowerPoint)
      if (!d && isTextual(o) && !doc.readOnly && (wasSel || (o.type === 'text' && o.ph && !o.html))) { const p = touch ? (ev.changedTouches && ev.changedTouches[0]) : ev; editObj(doc, s, o, el, body, p ? { x: p.clientX, y: p.clientY } : null); }
    };
    document.addEventListener(touch ? 'touchmove' : 'mousemove', mv, touch ? { passive: false } : false); document.addEventListener(touch ? 'touchend' : 'mouseup', up);
  };
  el.addEventListener('mousedown', start); el.addEventListener('touchstart', start, { passive: false });
  el.addEventListener('dblclick', e => { e.stopPropagation(); if (!el.classList.contains('editing')) editObj(doc, s, o, el, body, { x: e.clientX, y: e.clientY }); });
  el.addEventListener('contextmenu', e => { e.stopPropagation(); if (el.classList.contains('editing')) return; select(); contextMenu(e, objMenu(doc, o)); });
}
function snap(v, points) { for (const p of points) if (Math.abs(v - p) < 8) return p; return v; }
function place(el, o, sc) { el.style.left = o.x * sc + 'px'; el.style.top = o.y * sc + 'px'; el.style.width = o.w * sc + 'px'; el.style.height = o.h * sc + 'px'; if (o.rot) el.style.transform = el.style.webkitTransform = 'rotate(' + o.rot + 'deg)'; }
function caretAt(x, y) {
  if (document.caretRangeFromPoint) return document.caretRangeFromPoint(x, y);
  if (document.caretPositionFromPoint) { const p = document.caretPositionFromPoint(x, y); if (!p) return null; const r = document.createRange(); r.setStart(p.offsetNode, p.offset); r.collapse(true); return r; }
  return null;
}
export function editObj(doc, s, o, el, body, pt, initialText) {
  if (isTextual(o)) {
    if (el.classList.contains('editing')) return;
    doc.pushHistory(); el.classList.add('editing'); doc.editingObj = o;
    if (!el.classList.contains('sel')) { deselectAll(doc.canvas); el.classList.add('sel'); addHandles(el); doc.selectedObj = o; }
    let tf = body.querySelector(o.type === 'text' ? '.pp-tf' : '.pp-shape-text');
    if (!tf) { tf = h('div.pp-shape-text', { style: { fontSize: (o.fontSize || 18) * doc.scale + 'px', color: o.color || '#fff' } }); body.appendChild(tf); tf.innerHTML = o.text || ''; }
    const wasPh = tf.classList.contains('ph');
    tf.classList.remove('ph'); if (wasPh || (o.type === 'text' && !o.html)) tf.innerHTML = '<br>';
    tf.contentEditable = 'true'; tf.setAttribute('spellcheck', 'false');
    tf.style.cursor = 'text'; tf.style.webkitUserSelect = tf.style.userSelect = 'text';
    try { tf.focus({ preventScroll: true }); } catch (e) { tf.focus(); }
    const sel = window.getSelection(); let rng = null;
    if (pt && !wasPh) { rng = caretAt(pt.x, pt.y); if (rng && !tf.contains(rng.startContainer)) rng = null; }
    if (!rng) { rng = document.createRange(); rng.selectNodeContents(tf); rng.collapse(false); }
    sel.removeAllRanges(); sel.addRange(rng);
    if (initialText) { document.execCommand('insertText', false, initialText); }
    let done = false;
    const commit = () => {
      if (done) return; done = true;
      tf.contentEditable = 'false'; el.classList.remove('editing'); doc.editingObj = null;
      let html = tf.innerHTML; if (/^(\s|<br\s*\/?>)*$/i.test(html)) html = '';
      const prev = o.type === 'text' ? (o.html || '') : (o.text || '');
      if (true) { if (o.type === 'text') o.html = html; else o.text = html; doc.pushHistory(); doc.markDirty(); }
      if (!html && o.type === 'text' && o.ph) { tf.classList.add('ph'); tf.innerHTML = o.ph === 'title' ? 'Bấm để thêm tiêu đề' : o.ph === 'subtitle' ? 'Bấm để thêm phụ đề' : 'Bấm để thêm văn bản'; }
      doc.renderThumbs(); if (doc.ribbon) doc.ribbon.refresh();
    };
    if (tf._oosBound) { tf._oosCommit = commit; } else {
    tf._oosBound = true; tf._oosCommit = commit;
    tf.addEventListener('blur', () => tf._oosCommit && tf._oosCommit());
    tf.addEventListener('keydown', e => { if (!el.classList.contains('editing')) return; const k = keyName(e); if (k === 'Escape') { e.preventDefault(); tf.blur(); if (doc.canvas) doc.canvas.focus && doc.canvas.focus(); } e.stopPropagation(); });
    tf.addEventListener('input', () => { if (!el.classList.contains('editing')) return; doc.markDirty(); clearTimeout(tf._t); tf._t = setTimeout(() => { if (o.type === 'text') o.html = tf.innerHTML; else o.text = tf.innerHTML; doc.renderThumbs(); }, 400); });
    ['mousedown', 'mouseup', 'click', 'touchstart'].forEach(ev => tf.addEventListener(ev, e => { if (el.classList.contains('editing')) e.stopPropagation(); }));
    }
  } else if (o.type === 'image') { altText(doc, o); }
  else if (o.type === 'chart') { chartDataDialog(doc, o); }
  else if (o.type === 'equation') { import('../math/math.js').then(m => m.equationEditor(o.latex).then(lx => { if (lx != null) doc.change('Sửa phương trình', () => o.latex = lx); })); }
  else if (o.type === 'table') tableEdit(doc, o);
}
// typing while a text box is selected (not editing) starts editing with that character — like PowerPoint
export function typeIntoSelected(doc, e) {
  const o = doc.selectedObj; if (!o || !isTextual(o) || doc.editingObj || doc.readOnly) return false;
  if (e.ctrlKey || e.metaKey || e.altKey) return false;
  const k = e.key; const isChar = k && k.length === 1; if (!isChar && k !== 'Enter' && k !== 'F2') return false;
  const el = doc.canvas.querySelector('[data-id="' + o.id + '"]'); if (!el) return false;
  e.preventDefault();
  if (isChar && o.type === 'text') { o.html = ''; }
  editObj(doc, doc.slide, o, el, el.querySelector('.pp-obj-body'), null, isChar ? k : '');
  return true;
}

// ---------------------------------------------------------------- object add/manipulate
export function addObj(doc, o) { doc.change('Chèn đối tượng', () => { o.id = uid('o'); doc.slide.objects.push(o); doc.selectedObj = o; }); setTimeout(() => { const el = doc.canvas.querySelector('[data-id="' + o.id + '"]'); if (el && (o.type === 'text' || o.type === 'shape')) editObj(doc, doc.slide, o, el, el.querySelector('.pp-obj-body')); }, 30); }
export function addTextbox(doc) { addObj(doc, { type: 'text', x: 240, y: 300, w: 500, h: 100, html: '', fontSize: 20 }); }
export function addShape(doc, shape) { doc.change('Chèn hình', () => { doc.slide.objects.push({ id: uid('o'), type: 'shape', shape, x: 440, y: 260, w: 400, h: 200, fill: doc.theme.colors[0], stroke: 'none' }); }); }
export async function addImage(doc) { const files = await Host.openFiles('image/*'); if (!files.length) return; const f = files[0]; const url = bytesToDataUrl(f.bytes, /\.png$/i.test(f.name) ? 'image/png' : /\.gif$/i.test(f.name) ? 'image/gif' : /\.svg$/i.test(f.name) ? 'image/svg+xml' : 'image/jpeg'); const img = new Image(); img.onload = () => { let w = img.naturalWidth, hh = img.naturalHeight; const sc = Math.min(1, 700 / w, 450 / hh); addObj(doc, { type: 'image', src: url, x: (SW - w * sc) / 2, y: (SH - hh * sc) / 2, w: w * sc, h: hh * sc, alt: '' }); }; img.src = url; }
export function addChart(doc, kind) { doc.change('Chèn biểu đồ', () => doc.slide.objects.push({ id: uid('o'), type: 'chart', kind, data: sampleChart(), x: 300, y: 180, w: 680, h: 400, title: 'Tiêu đề biểu đồ' })); }
export function addTable(doc) { formDlg('Chèn bảng', [{ key: 'r', type: 'number', label: 'Số hàng', value: 3 }, { key: 'c', type: 'number', label: 'Số cột', value: 4 }]).then(v => { if (!v) return; const rows = []; for (let r = 0; r < v.r; r++) { const row = []; for (let c = 0; c < v.c; c++) row.push(r === 0 ? 'Cột ' + (c + 1) : ''); rows.push(row); } doc.change('Chèn bảng', () => doc.slide.objects.push({ id: uid('o'), type: 'table', rows, header: true, x: 180, y: 220, w: 920, h: 300, fontSize: 18 })); }); }
export function addEquation(doc) { import('../math/math.js').then(m => m.equationEditor('').then(lx => { if (lx) addObj(doc, { type: 'equation', latex: lx, x: 400, y: 300, w: 480, h: 120 }); })); }
export function addVideo(doc) { toast('Chèn video: tính năng phát video sẽ mở tệp bằng trình phát hệ thống khi trình chiếu'); }
export function addIcon(doc, url) { addObj(doc, { type: 'image', src: url, x: 540, y: 280, w: 200, h: 200, alt: '', fit: 'contain' }); }
export function altText(doc, o) { promptDlg('Văn bản thay thế (Alt Text):', o.alt || '', 'Văn bản thay thế', { multiline: true }).then(v => { if (v != null) doc.change('Văn bản thay thế', () => o.alt = v); }); }
export function chartDataDialog(doc, o) { const d = o.data || sampleChart(); const txt = [d.series.map(s => s.name).join('\t')].concat(d.cats.map((c, i) => c + '\t' + d.series.map(s => s.values[i]).join('\t'))).join('\n'); promptDlg('Dữ liệu biểu đồ (dòng 1: tên chuỗi; mỗi dòng sau: Nhãn + giá trị, cách nhau bằng Tab):', 'Nhãn\t' + txt, 'Chỉnh sửa dữ liệu', { multiline: true, rows: 8 }).then(v => { if (v == null) return; const lines = v.trim().split('\n'); const names = lines[0].split('\t').slice(1); const cats = [], series = names.map(n => ({ name: n, values: [] })); lines.slice(1).forEach(l => { const cells = l.split('\t'); cats.push(cells[0]); names.forEach((n, i) => series[i].values.push(parseFloat(cells[i + 1]) || 0)); }); doc.change('Dữ liệu biểu đồ', () => o.data = { cats, series }); }); }
export function tableEdit(doc, o) { toast('Bấm đúp vào ô trong bảng để sửa — đang mở chế độ sửa từng ô'); }
export function objMenu(doc, o) {
  const m = [
    { label: 'Cắt', icon: 'scissors', run: () => { doc.clip = deepClone(o); removeObj(doc, o); } },
    { label: 'Sao chép', icon: 'copy', run: () => { doc.clip = deepClone(o); toast('Đã sao chép'); } },
    { label: 'Dán', icon: 'clipboard-paste', disabled: !doc.clip, run: () => pasteObj(doc) },
    { label: 'Nhân bản', icon: 'copy-plus', run: () => doc.change('Nhân bản', () => { const n = deepClone(o); n.id = uid('o'); n.x += 20; n.y += 20; doc.slide.objects.push(n); doc.selectedObj = n; }) },
    { sep: true },
    { label: 'Đưa lên trước', icon: 'bring-to-front', run: () => order(doc, o, 1) },
    { label: 'Đưa ra sau', icon: 'send-to-back', run: () => order(doc, o, -1) },
    { label: 'Xoay 90°', icon: 'rotate-cw', run: () => doc.change('Xoay', () => o.rot = ((o.rot || 0) + 90) % 360) },
    { sep: true }
  ];
  if (o.type === 'text' || o.type === 'shape') m.push({ label: 'Sửa văn bản', icon: 'type', run: () => { const el = doc.canvas.querySelector('[data-id="' + o.id + '"]'); if (el) editObj(doc, doc.slide, o, el, el.querySelector('.pp-obj-body')); } });
  if (o.type === 'shape') m.push({ label: 'Màu tô', icon: 'paint-bucket', run: () => colorPicker(doc.canvas, c => doc.change('Màu tô', () => o.fill = c)) });
  if (o.type === 'chart') m.push({ label: 'Sửa dữ liệu', icon: 'table', run: () => chartDataDialog(doc, o) }, { label: 'Đổi kiểu', icon: 'chart-column', sub: ['column', 'bar', 'line', 'pie', 'area'].map(k => ({ label: k, run: () => doc.change('Đổi kiểu', () => o.kind = k) })) });
  if (o.type === 'image') m.push({ label: 'Văn bản thay thế', icon: 'accessibility', run: () => altText(doc, o) });
  m.push({ sep: true }, { label: 'Xóa', icon: 'trash-2', run: () => removeObj(doc, o) });
  return m;
}
export function order(doc, o, dir) { doc.change('Thứ tự', () => { const arr = doc.slide.objects; const i = arr.indexOf(o); arr.splice(i, 1); if (dir > 0) arr.push(o); else arr.unshift(o); }); }
export function removeObj(doc, o) { doc.change('Xóa đối tượng', () => { doc.slide.objects = doc.slide.objects.filter(x => x !== o); doc.selectedObj = null; }); }
export function pasteObj(doc) { if (!doc.clip) return; const n = deepClone(doc.clip); n.id = uid('o'); n.x = (n.x || 100) + 20; n.y = (n.y || 100) + 20; doc.change('Dán', () => { doc.slide.objects.push(n); doc.selectedObj = n; }); }
export function align(doc, how) {
  const o = doc.selectedObj; if (!o) return;
  doc.change('Căn chỉnh', () => { if (how === 'left') o.x = 40; else if (how === 'hcenter') o.x = (SW - o.w) / 2; else if (how === 'right') o.x = SW - o.w - 40; else if (how === 'top') o.y = 40; else if (how === 'vcenter') o.y = (SH - o.h) / 2; else if (how === 'bottom') o.y = SH - o.h - 40; });
}

// ---------------------------------------------------------------- text formatting on selected text object
export function fmtText(doc, cmd, val) {
  const o = doc.selectedObj; const el = o && doc.canvas.querySelector('[data-id="' + o.id + '"] .pp-tf, [data-id="' + o.id + '"] .pp-shape-text');
  if (o && el && el.isContentEditable) { el.focus(); document.execCommand(cmd, false, val); if (o.type === 'text') o.html = el.innerHTML; else o.text = el.innerHTML; doc.markDirty(); return; }
  if (!o) return toast('Chọn một hộp văn bản', 'err');
  doc.change('Định dạng', () => { if (cmd === 'bold') o.b = !o.b; else if (cmd === 'italic') o.i = !o.i; else if (cmd === 'justifyLeft') o.align = 'left'; else if (cmd === 'justifyCenter') o.align = 'center'; else if (cmd === 'justifyRight') o.align = 'right'; });
}
export function setFontSize(doc, sz) { const o = doc.selectedObj; if (!o) return; doc.change('Cỡ chữ', () => o.fontSize = sz); }
export function sizeStep(doc, dir) { const o = doc.selectedObj; if (!o) return; doc.change('Cỡ chữ', () => o.fontSize = clamp((o.fontSize || 18) + dir * 4, 6, 200)); }
export function setFont(doc, f) { const o = doc.selectedObj; if (!o) return; doc.change('Phông chữ', () => o.font = f); }
export function textColor(doc, c) { const o = doc.selectedObj; if (!o) return; doc.change('Màu chữ', () => o.color = c); }
export function setAlign(doc, a) { const o = doc.selectedObj; if (!o) return; doc.change('Căn lề', () => o.align = a); }

// ---------------------------------------------------------------- slides
export function newSlide(doc, layout) {
  doc.change('Slide mới', () => { const s = blankSlide(layout || 'content', doc.theme); doc.slides.splice(doc.cur + 1, 0, s); doc.cur++; });
  doc.goto(doc.cur);
}
export function duplicateSlide(doc) { doc.change('Nhân bản slide', () => { (doc.sel.length ? doc.sel : [doc.cur]).sort((a, b) => b - a).forEach(i => { const n = deepClone(doc.slides[i]); n.objects.forEach(o => o.id = uid('o')); doc.slides.splice(i + 1, 0, n); }); doc.cur++; }); }
export function deleteSlide(doc) { if (doc.slides.length <= (doc.sel.length || 1)) { toast('Không thể xóa slide cuối cùng', 'err'); return; } doc.change('Xóa slide', () => { const del = (doc.sel.length ? doc.sel : [doc.cur]).slice().sort((a, b) => b - a); del.forEach(i => doc.slides.splice(i, 1)); doc.cur = clamp(doc.cur, 0, doc.slides.length - 1); doc.sel = [doc.cur]; }); }
export function moveSlide(doc, from, to) { if (from === to) return; doc.change('Di chuyển slide', () => { const [s] = doc.slides.splice(from, 1); doc.slides.splice(to, 0, s); doc.cur = to; }); }
export function changeLayout(doc, layout) { doc.change('Bố cục slide', () => { doc.slide.layout = layout; const existing = {}; doc.slide.objects.forEach(o => { if (o.ph) existing[o.ph] = o; }); doc.slide.objects = layoutObjects(layout, doc.theme, existing); }); }
export function newSlideMenu(doc) { return LAYOUTS.map(l => ({ label: l[1], icon: l[2], run: () => newSlide(doc, l[0]) })); }
export function layoutMenu(doc) { return LAYOUTS.map(l => ({ label: l[1], icon: l[2], checked: doc.slide.layout === l[0] || undefined, run: () => changeLayout(doc, l[0]) })); }
export function resetSlide(doc) { doc.change('Đặt lại slide', () => { const existing = {}; doc.slide.objects.forEach(o => { if (o.ph) existing[o.ph] = o; }); doc.slide.objects = layoutObjects(doc.slide.layout, doc.theme, existing); }); }
export function slideMenu(doc) {
  return [
    { label: 'Slide Mới', icon: 'plus', run: () => newSlide(doc) },
    { label: 'Nhân bản Slide', icon: 'copy-plus', run: () => duplicateSlide(doc) },
    { label: 'Xóa Slide', icon: 'trash-2', run: () => deleteSlide(doc) },
    { sep: true },
    { label: 'Bố cục', icon: 'layout-template', sub: () => layoutMenu(doc) },
    { label: 'Đặt lại', icon: 'rotate-ccw', run: () => resetSlide(doc) },
    { label: 'Nền slide…', icon: 'image', run: () => bgDialog(doc) },
    { sep: true },
    { label: 'Chuyển tiếp…', icon: 'wind', sub: () => transitionMenu(doc) },
    { label: 'Trình chiếu từ slide này', icon: 'play', run: () => startShow(doc, doc.cur) }
  ];
}
export function bgDialog(doc) {
  formDlg('Nền slide', [{ key: 'type', type: 'radio', value: 'solid', options: [['solid', 'Màu đơn'], ['gradient', 'Chuyển sắc'], ['theme', 'Theo chủ đề']] }, { key: 'color', type: 'color', label: 'Màu', value: '#ffffff' }, { key: 'all', type: 'checkbox', label: 'Áp dụng cho tất cả slide' }]).then(v => { if (!v) return; const bg = v.type === 'theme' ? null : v.type === 'gradient' ? { type: 'gradient', angle: 135, stops: [v.color, '#ffffff'] } : v.color; doc.change('Nền slide', () => { if (v.all) doc.slides.forEach(s => s.bg = bg); else doc.slide.bg = bg; }); });
}

// ---------------------------------------------------------------- transitions & animations
export const TRANSITIONS = [['none', 'Không', 'ban'], ['fade', 'Mờ dần', 'contrast'], ['push', 'Đẩy', 'arrow-right'], ['wipe', 'Lau', 'eraser'], ['split', 'Tách', 'split'], ['cut', 'Cắt', 'scissors'], ['zoom', 'Phóng', 'zoom-in'], ['morph', 'Biến hình', 'shapes'], ['fall', 'Rơi', 'arrow-down'], ['flip', 'Lật', 'flip-horizontal']];
export function transitionMenu(doc) { return TRANSITIONS.map(t => ({ label: t[1], icon: t[2], checked: (doc.slide.transition || 'none') === t[0] || undefined, run: () => setTransition(doc, t[0]) })); }
export function setTransition(doc, t, all) { doc.change('Chuyển tiếp', () => { if (all) doc.slides.forEach(s => s.transition = t === 'none' ? null : t); else doc.slide.transition = t === 'none' ? null : t; }); toast('Chuyển tiếp: ' + (TRANSITIONS.find(x => x[0] === t) || ['', t])[1]); }
export const ANIMS = [['none', 'Không'], ['fadeIn', 'Xuất hiện mờ'], ['flyLeft', 'Bay vào từ trái'], ['flyBottom', 'Bay vào từ dưới'], ['zoomIn', 'Phóng to'], ['bounce', 'Nảy'], ['spin', 'Xoay']];
export function setAnimation(doc, anim) { const o = doc.selectedObj; if (!o) return toast('Chọn một đối tượng để thêm hoạt ảnh', 'err'); doc.change('Hoạt ảnh', () => o.anim = anim === 'none' ? null : anim); toast('Hoạt ảnh: ' + (ANIMS.find(a => a[0] === anim) || ['', anim])[1]); }
export function animationMenu(doc) { return ANIMS.map(a => ({ label: a[1], checked: doc.selectedObj && (doc.selectedObj.anim || 'none') === a[0] || undefined, run: () => setAnimation(doc, a[0]) })); }

// ---------------------------------------------------------------- slideshow
export function startShow(doc, from) {
  const show = h('div.pp-show');
  const slideEl = h('div.pp-show-slide');
  show.appendChild(slideEl);
  const toolbar = h('div.pp-show-bar',
    h('button', { title: 'Trước', onclick: () => nav(-1) }, ic('chevron-left', 20)),
    h('span.pp-show-no'),
    h('button', { title: 'Sau', onclick: () => nav(1) }, ic('chevron-right', 20)),
    h('button', { title: 'Bút', onclick: () => togglePen() }, ic('pen-tool', 18)),
    h('button', { title: 'Thoát (Esc)', onclick: () => endShow(doc) }, ic('x', 20)));
  show.appendChild(toolbar);
  document.body.appendChild(show);
  doc._show = { el: show, idx: from || 0, animDone: false };
  const scale = () => Math.min(window.innerWidth / SW, window.innerHeight / SH);
  const render = () => {
    const s = doc.slides[doc._show.idx]; slideEl.innerHTML = '';
    slideEl.style.width = SW + 'px'; slideEl.style.height = SH + 'px'; const sc = scale();
    slideEl.style.webkitTransform = slideEl.style.transform = 'scale(' + sc + ')';
    slideEl.style.left = (window.innerWidth - SW * sc) / 2 + 'px'; slideEl.style.top = (window.innerHeight - SH * sc) / 2 + 'px';
    renderSlideInto(doc, s, slideEl, 1, false);
    if (s.transition) { slideEl.classList.remove('tr-anim'); void slideEl.offsetWidth; slideEl.className = 'pp-show-slide tr-anim tr-' + s.transition; }
    // animations: hide animated objects then reveal on click
    doc._show.anims = []; Array.prototype.forEach.call(slideEl.querySelectorAll('.pp-obj'), (el, i) => { const o = s.objects[i]; if (o && o.anim) { el.style.opacity = '0'; doc._show.anims.push({ el, anim: o.anim }); } });
    $('.pp-show-no', toolbar).textContent = (doc._show.idx + 1) + ' / ' + doc.slides.length;
    if (s.narration || s.timing) import('./record.js').then(R => R.onShowSlide(doc, s, () => nav(1)));
  };
  const nav = d => { if (d > 0 && doc._show.anims && doc._show.anims.length) { const a = doc._show.anims.shift(); a.el.style.opacity = ''; a.el.classList.add('anim-' + a.anim); return; } const ni = doc._show.idx + d; if (ni < 0) return; if (ni >= doc.slides.length) { endShow(doc); return; } doc._show.idx = ni; render(); };
  let pen = null;
  const togglePen = () => { if (pen) { pen.remove(); pen = null; return; } pen = h('canvas.pp-pen', { width: window.innerWidth, height: window.innerHeight }); show.appendChild(pen); const ctx = pen.getContext('2d'); ctx.strokeStyle = '#e81123'; ctx.lineWidth = 3; ctx.lineCap = 'round'; let drawing = false; pen.addEventListener('mousedown', e => { drawing = true; ctx.beginPath(); ctx.moveTo(e.clientX, e.clientY); }); pen.addEventListener('mousemove', e => { if (drawing) { ctx.lineTo(e.clientX, e.clientY); ctx.stroke(); } }); pen.addEventListener('mouseup', () => drawing = false); };
  doc._show.key = e => { const k = keyName(e); if (k === 'Escape') endShow(doc); else if (k === 'ArrowRight' || k === ' ' || k === 'Enter' || k === 'PageDown') { e.preventDefault(); nav(1); } else if (k === 'ArrowLeft' || k === 'PageUp' || k === 'Backspace') { e.preventDefault(); nav(-1); } else if (k === 'Home') { doc._show.idx = 0; render(); } else if (k === 'End') { doc._show.idx = doc.slides.length - 1; render(); } else if (k === 'b' || k === 'B') { show.classList.toggle('black'); } else if (k === 'w' || k === 'W') { show.classList.toggle('white'); } };
  document.addEventListener('keydown', doc._show.key, true);
  show.addEventListener('click', e => { if (e.target === show || e.target === slideEl) nav(1); });
  show.addEventListener('contextmenu', e => { e.preventDefault(); nav(-1); });
  window.addEventListener('resize', render);
  const el = document.documentElement; if (el.requestFullscreen) el.requestFullscreen().catch(() => {}); else if (el.webkitRequestFullscreen) el.webkitRequestFullscreen();
  render();
}
export function endShow(doc) { if (!doc._show) return; clearTimeout(doc._showTimer); if (doc._showAudio) { try { doc._showAudio.pause(); } catch (e) {} doc._showAudio = null; } document.removeEventListener('keydown', doc._show.key, true); doc._show.el.remove(); doc._show = null; if (document.exitFullscreen && document.fullscreenElement) document.exitFullscreen().catch(() => {}); else if (document.webkitExitFullscreen && document.webkitFullscreenElement) document.webkitExitFullscreen(); }
export function slideSorter(doc) {
  const old = $('.pp-sorter'); if (old) { old.remove(); return; }
  const back = h('div.pp-sorter'); const grid = h('div.pp-sorter-grid');
  doc.slides.forEach((s, i) => { const card = h('div.pp-sorter-card', { draggable: 'true', dataset: { i } }, h('span.pp-thumb-no', i + 1)); const mini = h('div.pp-mini'); renderSlideInto(doc, s, mini, 0.2); card.appendChild(mini); card.addEventListener('click', () => { doc.goto(i); back.remove(); }); card.addEventListener('dragstart', e => e.dataTransfer.setData('text/plain', String(i))); card.addEventListener('dragover', e => e.preventDefault()); card.addEventListener('drop', e => { e.preventDefault(); moveSlide(doc, +e.dataTransfer.getData('text/plain'), i); back.remove(); slideSorter(doc); }); grid.appendChild(card); });
  back.appendChild(h('div.pp-sorter-head', h('b', 'Sắp xếp Slide — kéo thả để đổi thứ tự'), h('button.btn', { onclick: () => back.remove() }, 'Đóng')));
  back.appendChild(grid); doc.el.appendChild(back);
}

// ---------------------------------------------------------------- keyboard
export function onKey(doc, e) {
  if (document.querySelector('.dlg-back') || doc._show) return;
  const c = chord(e), k = keyName(e);
  const editing = doc.canvas.querySelector('.editing');
  if (editing) return;
  const ae = document.activeElement; if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA' || ae.tagName === 'SELECT' || ae.isContentEditable)) return;
  if (c === 'ctrl+m') { e.preventDefault(); newSlide(doc); return; }
  if (c === 'f5') { e.preventDefault(); startShow(doc, 0); return; }
  if (c === 'shift+f5') { e.preventDefault(); startShow(doc, doc.cur); return; }
  if (c === 'ctrl+d' && doc.selectedObj) { e.preventDefault(); doc.change('Nhân bản', () => { const n = deepClone(doc.selectedObj); n.id = uid('o'); n.x += 20; n.y += 20; doc.slide.objects.push(n); doc.selectedObj = n; }); return; }
  if (c === 'ctrl+d') { e.preventDefault(); duplicateSlide(doc); return; }
  if (c === 'ctrl+c' && doc.selectedObj) { doc.clip = deepClone(doc.selectedObj); return; }
  if (c === 'ctrl+v' && doc.clip) { e.preventDefault(); pasteObj(doc); return; }
  if (c === 'ctrl+z') { e.preventDefault(); doc.undo(); return; }
  if (c === 'ctrl+y' || c === 'ctrl+shift+z') { e.preventDefault(); doc.redo(); return; }
  if (doc.selectedObj) {
    const o = doc.selectedObj;
    if (typeIntoSelected(doc, e)) return;
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); removeObj(doc, o); return; }
    if (/^Arrow/.test(k)) { e.preventDefault(); const step = e.ctrlKey ? 1 : 10; doc.pushHistory(); if (k === 'ArrowLeft') o.x -= step; if (k === 'ArrowRight') o.x += step; if (k === 'ArrowUp') o.y -= step; if (k === 'ArrowDown') o.y += step; doc.renderSlide(); doc.markDirty(); return; }
    if (k === 'Escape') { doc.selectedObj = null; doc.renderSlide(); doc.onSelection(); return; }
  } else {
    if (k === 'PageDown' || (k === 'ArrowDown' && document.activeElement !== doc.notesEl)) { e.preventDefault(); doc.goto(doc.cur + 1); }
    if (k === 'PageUp' || (k === 'ArrowUp' && document.activeElement !== doc.notesEl)) { e.preventDefault(); doc.goto(doc.cur - 1); }
  }
}

// ---------------------------------------------------------------- outline & export
export function outlineToSlides(text) {
  const slides = []; let cur = null;
  text.split('\n').forEach(line => {
    if (/^#\s+/.test(line) || (!/^\s/.test(line) && line.trim() && !cur)) { cur = blankSlide('content'); cur.objects[0].html = esc(line.replace(/^#\s+/, '')); cur.objects[1].html = ''; slides.push(cur); }
    else if (/^##\s+/.test(line)) { cur = blankSlide('content'); cur.objects[0].html = esc(line.replace(/^##\s+/, '')); slides.push(cur); }
    else if (line.trim() && cur) { const body = cur.objects[1]; body.html = (body.html || '') + '<div>• ' + esc(line.replace(/^\s*[-*]\s*/, '')) + '</div>'; }
  });
  return slides.length ? slides : [blankSlide('title')];
}
export function exportOutline(doc) { return doc.slides.map(s => { const title = plain(s.objects.find(o => o.ph === 'title') ? s.objects.find(o => o.ph === 'title').html : ''); const body = s.objects.filter(o => o.type === 'text' && o.ph !== 'title').map(o => plain(o.html).split('\n').map(l => l.trim() ? '- ' + l : '').filter(Boolean).join('\n')).join('\n'); return '# ' + title + '\n' + body; }).join('\n\n'); }
export function objText(o) { return plain(o.html || o.text || ''); }
export function plain(html) { const d = document.createElement('div'); d.innerHTML = html || ''; return (d.textContent || '').trim(); }
export function exportHtml(doc) {
  const slidesHtml = doc.slides.map((s, i) => { const div = document.createElement('div'); renderSlideInto(doc, s, div, 1, false); return '<section class="slide" data-tr="' + (s.transition || '') + '">' + div.innerHTML + '</section>'; }).join('');
  return '<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>' + esc(doc.name) + '</title><style>*{box-sizing:border-box;margin:0}body{background:#111;font-family:' + doc.theme.font + ',Arial,sans-serif;overflow:hidden}.slide{display:none;position:fixed;left:50%;top:50%;width:' + SW + 'px;height:' + SH + 'px;transform:translate(-50%,-50%) scale(var(--s,1));transform-origin:center}.slide.on{display:block}.pp-obj{position:absolute}.pp-tf,.pp-shape-text{width:100%;height:100%;display:flex;flex-direction:column;overflow:hidden}.nav{position:fixed;bottom:12px;left:50%;transform:translateX(-50%);color:#fff;z-index:10}</style></head><body>' + slidesHtml + '<div class="nav">◀ ▶ (phím mũi tên) · <span id="n"></span></div><script>var i=0,S=document.querySelectorAll(".slide");function fit(){var s=Math.min(innerWidth/' + SW + ',innerHeight/' + SH + ');document.documentElement.style.setProperty("--s",s)}function show(){S.forEach(function(x,k){x.classList.toggle("on",k===i)});document.getElementById("n").textContent=(i+1)+" / "+S.length}onresize=fit;onkeydown=function(e){if(e.key==="ArrowRight"||e.key===" "){i=Math.min(i+1,S.length-1)}else if(e.key==="ArrowLeft"){i=Math.max(0,i-1)}else return;show()};onclick=function(){i=Math.min(i+1,S.length-1);show()};fit();show();<\/script></body></html>';
}
export async function slideToPng(doc, idx, mime, scale) {
  scale = scale || 1;
  const div = document.createElement('div'); div.style.cssText = 'position:fixed;left:-9999px;width:' + SW + 'px;height:' + SH + 'px'; document.body.appendChild(div);
  renderSlideInto(doc, doc.slides[idx], div, 1, false);
  const xml = '<svg xmlns="http://www.w3.org/2000/svg" width="' + SW + '" height="' + SH + '"><foreignObject width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml" style="width:' + SW + 'px;height:' + SH + 'px;position:relative;font-family:' + doc.theme.font + ',Arial">' + div.innerHTML + '</div></foreignObject></svg>';
  document.body.removeChild(div);
  return new Promise(res => { const img = new Image(); img.onload = () => { const c = document.createElement('canvas'); c.width = SW * scale; c.height = SH * scale; const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(img, 0, 0, c.width, c.height); res(c.toDataURL(mime || 'image/png', 0.92)); }; img.onerror = () => res(null); img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(xml))); });
}
export function renderPrint(doc, area) {
  area.appendChild(h('style', '@page{size:' + SW + 'px ' + SH + 'px;margin:0} #print-area .pslide{page-break-after:always;width:' + SW + 'px;height:' + SH + 'px;position:relative;overflow:hidden} #print-area .pp-obj{position:absolute}'));
  doc.slides.forEach(s => { const div = h('div.pslide'); renderSlideInto(doc, s, div, 1, false); area.appendChild(div); });
}
export function printPreview(doc, box) {
  box.innerHTML = ''; const sc = Math.min(1, (box.clientWidth - 40) / SW);
  doc.slides.slice(0, 20).forEach((s, i) => { const wrap = h('div.pp-pp', { style: { width: SW * sc + 'px', height: SH * sc + 'px' } }); const div = h('div', { style: { width: SW + 'px', height: SH + 'px', position: 'relative', webkitTransform: 'scale(' + sc + ')', transform: 'scale(' + sc + ')', transformOrigin: 'top left' } }); renderSlideInto(doc, s, div, 1, false); wrap.appendChild(div); box.appendChild(wrap); });
}
export function printOptions(doc) { return h('div.form', h('div.muted', doc.slides.length + ' slide sẽ được in, mỗi slide một trang. Chọn "Lưu dưới dạng PDF" trong hộp thoại in để xuất PDF.')); }
export function extraCommands(doc) { return [{ label: 'Trình chiếu (F5)', icon: 'play', run: () => startShow(doc, 0) }, { label: 'Sắp xếp slide', icon: 'layout-grid', run: () => slideSorter(doc) }, { label: 'Slide mới', icon: 'plus', run: () => newSlide(doc) }]; }
