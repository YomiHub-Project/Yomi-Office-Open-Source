// Floating objects on a worksheet: charts, pictures, shapes, text boxes, equations + ink layer
import { h, ic, esc, clamp, uid, keyName, readFileAsDataUrl, KaTeX } from '../core/util.js';
import { contextMenu, toast, showMenu, promptDlg, dialog, formDlg, colorPicker } from '../core/ui.js';
import { renderChart, chartData, CHART_TYPES, PALETTE } from './charts.js';
import { rangeAddr, parseRangeAddr } from './model.js';
import { InkTool, renderInk } from '../core/ink.js';
import { shapeSvg, SHAPES } from '../core/shapes.js';
import Host from '../core/host.js';

export class Objects {
  constructor(doc) {
    this.doc = doc; this.selected = null;
    this.inner = h('div.xl-objs-inner');
    this.inkSvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); this.inkSvg.setAttribute('class', 'xl-ink');
    this.inner.appendChild(this.inkSvg);
    doc.grid.objLayer.appendChild(this.inner);
    this.ink = new InkTool({
      toLocal: e => { const r = this.inner.getBoundingClientRect(); const z = doc.grid.z; return [(e.clientX - r.left) / z, (e.clientY - r.top) / z]; },
      getStrokes: () => doc.wb.cur.ink, scale: () => doc.grid.z,
      begin: () => { if (!this._inkSnap) this._inkSnap = JSON.stringify(doc.wb.cur.ink); },
      commit: () => { const s = doc.wb.cur; const before = this._inkSnap; this._inkSnap = null; const after = JSON.stringify(s.ink); if (before !== after) { doc.undoStack.push({ label: 'Vẽ', undo: () => { s.ink = JSON.parse(before); this.renderInk(); }, redo: () => { s.ink = JSON.parse(after); this.renderInk(); }, sheet: s }); doc.markDirty(); } },
      redraw: () => this.renderInk(),
      onMode: m => { doc.grid.el.classList.toggle('inking', !!m); this.inkCatcher.style.display = m ? 'block' : 'none'; }
    });
    this.inkCatcher = h('div.xl-inkcatch'); this.inkCatcher.style.display = 'none';
    doc.grid.el.appendChild(this.inkCatcher);
    this.ink.attach(this.inkCatcher);
    document.addEventListener('keydown', e => this.onKey(e), true);
  }
  get list() { return this.doc.wb.cur.objects; }
  renderInk() { const z = this.doc.grid.z; this.inkSvg.style.display = this.doc.inkHidden ? 'none' : ''; this.inkSvg.setAttribute('width', 20000); this.inkSvg.setAttribute('height', 40000); renderInk(this.inkSvg, this.doc.wb.cur.ink, z); }
  render() {
    const z = this.doc.grid.z;
    Array.prototype.slice.call(this.inner.querySelectorAll('.xl-obj')).forEach(e => e.parentNode.removeChild(e));
    this.list.forEach(o => this.inner.appendChild(this.renderObj(o, z)));
    this.renderInk();
    this.doc.grid.positionObjects();
  }
  refreshCharts() { const z = this.doc.grid.z; this.list.forEach(o => { if (o.type !== 'chart') return; const el = this.inner.querySelector('[data-id="' + o.id + '"] .xl-obj-body'); if (el) this.drawChart(o, el, z); }); }
  drawChart(o, body, z) { renderChart(body, chartData(this.doc.wb.cur, o.src, o), { kind: o.kind, title: o.title, legend: o.legend, labels: o.labels, w: o.w * z, h: o.h * z, colors: o.colors || (this.doc.wb.props.themeColors) }); }
  renderObj(o, z) {
    const el = h('div.xl-obj' + (this.selected === o ? '.sel' : ''), { dataset: { id: o.id } });
    el.style.left = o.x * z + 'px'; el.style.top = o.y * z + 'px'; el.style.width = o.w * z + 'px'; el.style.height = o.h * z + 'px';
    if (o.rot) { el.style.webkitTransform = el.style.transform = 'rotate(' + o.rot + 'deg)'; }
    const body = h('div.xl-obj-body'); el.appendChild(body);
    if (o.type === 'chart') this.drawChart(o, body, z);
    else if (o.type === 'image') { const img = h('img', { src: o.src, alt: o.alt || '', draggable: false }); img.style.width = '100%'; img.style.height = '100%'; if (o.crop) img.style.objectFit = 'cover'; body.appendChild(img); if (o.border) body.style.border = o.border; }
    else if (o.type === 'shape') { body.innerHTML = shapeSvg(o.shape, o.w * z, o.h * z, o.fill || '#4472c4', o.stroke || '#2f528f', o.strokeW || 1.5); if (o.text) body.appendChild(h('div.xl-shape-text', { style: { color: o.color || '#fff' } }, o.text)); }
    else if (o.type === 'textbox') { const t = h('div.xl-tb', { style: { fontSize: (o.sz || 11) * 1.333 * z + 'px', color: o.color || '#000', background: o.fill || '#fff', border: o.noBorder ? '0' : '1px solid #bfbfbf', textAlign: o.align || 'left', fontWeight: o.b ? 'bold' : '', fontFamily: o.font || '' } }); t.textContent = o.text || ''; if (o.wordart) { t.style.background = 'transparent'; t.style.border = '0'; t.style.color = o.color || '#4472c4'; t.style.fontWeight = 'bold'; t.style.textShadow = '2px 2px 3px rgba(0,0,0,.35)'; t.style.fontSize = (o.sz || 32) * 1.333 * z + 'px'; } body.appendChild(t); }
    else if (o.type === 'equation') { const t = h('div.xl-eq'); t.style.fontSize = 16 * z + 'px'; body.appendChild(t); KaTeX().then(k => { try { k.render(o.latex || '', t, { throwOnError: false, displayMode: true }); } catch (e) { t.textContent = o.latex; } }); }
    // handles
    if (this.selected === o) ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].forEach(d => { const hd = h('div.xl-h.h-' + d, { dataset: { d } }); el.appendChild(hd); });
    this.bindObj(el, o);
    return el;
  }
  bindObj(el, o) {
    const start = e => {
      if (this.ink.tool) return;
      if (e.type === 'mousedown' && e.button !== 0) return;
      e.stopPropagation();
      const editing = el.querySelector('[contenteditable=true]'); if (editing && editing.contains(e.target)) return;
      if (this.doc.grid.editing) this.doc.grid.commit();
      if (this.selected !== o) { this.select(o); return; }
      e.preventDefault();
      const touch = e.type === 'touchstart'; const p0 = touch ? e.touches[0] : e; const z = this.doc.grid.z;
      const d = e.target.getAttribute && e.target.getAttribute('data-d');
      const o0 = { x: o.x, y: o.y, w: o.w, h: o.h }; const ratio = o.w / o.h;
      const before = JSON.stringify(this.list);
      const mv = ev => {
        const p = touch ? ev.touches[0] : ev; const dx = (p.clientX - p0.clientX) / z, dy = (p.clientY - p0.clientY) / z;
        if (!d) { o.x = Math.max(0, o0.x + dx); o.y = Math.max(0, o0.y + dy); }
        else {
          let x = o0.x, y = o0.y, w = o0.w, hh = o0.h;
          if (/e/.test(d)) w = Math.max(16, o0.w + dx); if (/s/.test(d)) hh = Math.max(16, o0.h + dy);
          if (/w/.test(d)) { w = Math.max(16, o0.w - dx); x = o0.x + o0.w - w; } if (/n/.test(d)) { hh = Math.max(16, o0.h - dy); y = o0.y + o0.h - hh; }
          if ((ev.shiftKey || o.type === 'image') && d.length === 2) { hh = w / ratio; if (/n/.test(d)) y = o0.y + o0.h - hh; }
          Object.assign(o, { x, y, w, h: hh });
        }
        el.style.left = o.x * z + 'px'; el.style.top = o.y * z + 'px'; el.style.width = o.w * z + 'px'; el.style.height = o.h * z + 'px';
        if (ev.cancelable) ev.preventDefault();
      };
      const up = () => { document.removeEventListener(touch ? 'touchmove' : 'mousemove', mv); document.removeEventListener(touch ? 'touchend' : 'mouseup', up); const after = JSON.stringify(this.list); if (after !== before) { this.pushUndo('Di chuyển/đổi kích thước đối tượng', before, after); this.render(); } };
      document.addEventListener(touch ? 'touchmove' : 'mousemove', mv, touch ? { passive: false } : false); document.addEventListener(touch ? 'touchend' : 'mouseup', up);
    };
    el.addEventListener('mousedown', start); el.addEventListener('touchstart', start, { passive: false });
    el.addEventListener('dblclick', e => { e.stopPropagation(); this.editObj(o, el); });
    el.addEventListener('contextmenu', e => { e.stopPropagation(); this.select(o); contextMenu(e, this.menu(o)); });
  }
  pushUndo(label, before, after) {
    const s = this.doc.wb.cur;
    this.doc.undoStack.push({ label, undo: () => { s.objects = JSON.parse(before); this.selected = null; this.render(); }, redo: () => { s.objects = JSON.parse(after); this.render(); }, sheet: s });
    this.doc.redoStack = []; this.doc.markDirty();
  }
  mutate(label, fn) { const before = JSON.stringify(this.list); fn(); this.pushUndo(label, before, JSON.stringify(this.list)); this.render(); this.doc.onSelection(); }
  select(o) { this.selected = o; this.render(); this.doc.updateContextualTabs(); if (this.doc.ribbon) { const ctx = o.type === 'chart' ? 'chartdesign' : o.type === 'image' ? 'picformat' : 'shapeformat'; this.doc.ribbon.select(ctx); } }
  deselect() { if (!this.selected) return; this.selected = null; this.render(); this.doc.updateContextualTabs(); }
  selectAll() { if (this.list.length) this.select(this.list[0]); }
  onKey(e) {
    if (!this.selected || this.doc.shell.cur !== this.doc) return;
    if (document.activeElement && document.activeElement.isContentEditable) return;
    const k = keyName(e); const o = this.selected;
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); e.stopPropagation(); this.remove(o); }
    else if (k === 'Escape') { this.deselect(); }
    else if (/^Arrow/.test(k)) { e.preventDefault(); e.stopPropagation(); const step = e.ctrlKey ? 1 : 8; this.mutate('Dịch chuyển', () => { if (k === 'ArrowLeft') o.x = Math.max(0, o.x - step); if (k === 'ArrowRight') o.x += step; if (k === 'ArrowUp') o.y = Math.max(0, o.y - step); if (k === 'ArrowDown') o.y += step; }); }
    else if ((e.ctrlKey || e.metaKey) && (k === 'c' || k === 'x')) { this.clip = JSON.parse(JSON.stringify(o)); if (k === 'x') this.remove(o); toast('Đã sao chép đối tượng'); e.preventDefault(); e.stopPropagation(); }
    else if ((e.ctrlKey || e.metaKey) && k === 'v' && this.clip) { e.preventDefault(); e.stopPropagation(); this.add(Object.assign({}, this.clip, { id: uid('o'), x: this.clip.x + 20, y: this.clip.y + 20 })); }
  }
  remove(o) { this.mutate('Xóa đối tượng', () => { const s = this.doc.wb.cur; s.objects = s.objects.filter(x => x !== o); this.selected = null; }); this.doc.updateContextualTabs(); }
  add(o) { this.mutate('Chèn đối tượng', () => { this.list.push(o); }); this.select(this.list[this.list.length - 1]); }
  defaultPos(w, hh) { const g = this.doc.grid, s = this.doc.wb.cur; const rc = g.cellRect(s.sel.r, s.sel.c); const z = g.z; return { x: (rc.x - g.headW() + g.scroller.scrollLeft) / z + 20, y: (rc.y - g.headH() + g.scroller.scrollTop) / z + 10, w: w || 480, h: hh || 288 }; }
  addChart(kind, src) { const p = this.defaultPos(); this.add(Object.assign({ id: uid('ch'), type: 'chart', kind, src, title: '' }, p)); }
  async insertImage(dataUrl) {
    if (!dataUrl) { const files = await Host.openFiles('image/*'); if (!files.length) return; const f = files[0]; const { bytesToDataUrl } = await import('../core/util.js'); dataUrl = bytesToDataUrl(f.bytes, /\.png$/i.test(f.name) ? 'image/png' : /\.gif$/i.test(f.name) ? 'image/gif' : /\.svg$/i.test(f.name) ? 'image/svg+xml' : 'image/jpeg'); }
    const img = new Image(); img.onload = () => { let w = img.naturalWidth, hh = img.naturalHeight; const sc = Math.min(1, 520 / w, 380 / hh); const p = this.defaultPos(w * sc, hh * sc); this.add(Object.assign({ id: uid('img'), type: 'image', src: dataUrl, alt: '' }, p)); }; img.src = dataUrl;
  }
  addShape(shape) { const p = this.defaultPos(140, 90); this.add(Object.assign({ id: uid('sh'), type: 'shape', shape, fill: (this.doc.wb.props.themeColors || PALETTE)[0], stroke: '#2f528f' }, p)); }
  addTextbox(wordart) { const p = this.defaultPos(wordart ? 360 : 200, wordart ? 80 : 60); this.add(Object.assign({ id: uid('tb'), type: 'textbox', text: wordart ? 'Văn bản nghệ thuật' : 'Nhập văn bản', wordart: !!wordart, sz: wordart ? 32 : 11 }, p)); setTimeout(() => { const el = this.inner.querySelector('[data-id="' + this.selected.id + '"]'); if (el) this.editObj(this.selected, el); }, 50); }
  addEquation(latex) { const p = this.defaultPos(300, 70); this.add(Object.assign({ id: uid('eq'), type: 'equation', latex }, p)); }
  async pasteAsPicture() { const s = this.doc.wb.cur; const cl = this.doc.clip; if (!cl) return toast('Sao chép vùng ô trước', 'err'); const { sheetToHtml } = await import('./xlsx.js'); const html = sheetToHtml(cl.sheet, this.doc, { area: cl.range }); const w = 400, hh = 200; const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + hh + '"><foreignObject width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml" style="font:11pt Calibri,Arial">' + html.replace(/<br>/g, '<br/>') + '</div></foreignObject></svg>'; this.add(Object.assign({ id: uid('img'), type: 'image', src: 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg), alt: 'Ảnh vùng ô' }, this.defaultPos(w, hh))); }
  editObj(o, el) {
    if (o.type === 'textbox') { const t = el.querySelector('.xl-tb'); t.contentEditable = 'true'; t.focus(); const sel = window.getSelection(); const r = document.createRange(); r.selectNodeContents(t); sel.removeAllRanges(); sel.addRange(r); t.addEventListener('blur', () => { t.contentEditable = 'false'; if (t.innerText !== o.text) this.mutate('Sửa hộp văn bản', () => { o.text = t.innerText; }); }, { once: true }); return; }
    if (o.type === 'shape') { promptDlg('Văn bản trong hình:', o.text || '', 'Thêm văn bản').then(v => { if (v != null) this.mutate('Văn bản hình', () => { o.text = v; }); }); return; }
    if (o.type === 'chart') { this.chartDataDialog(o); return; }
    if (o.type === 'equation') { import('../math/math.js').then(m => m.equationEditor(o.latex).then(lx => { if (lx != null) this.mutate('Sửa phương trình', () => { o.latex = lx; }); })); return; }
    if (o.type === 'image') { this.altText(o); }
  }
  async chartDataDialog(o) { const v = await promptDlg('Vùng dữ liệu biểu đồ (vd: A1:C10):', o.src || '', 'Chọn dữ liệu'); if (v && parseRangeAddr(v)) this.mutate('Dữ liệu biểu đồ', () => { o.src = v.toUpperCase(); }); }
  async altText(o) { const v = await promptDlg('Mô tả cho người khiếm thị (Alt Text):', o.alt || '', 'Văn bản thay thế', { multiline: true }); if (v != null) this.mutate('Văn bản thay thế', () => { o.alt = v; }); }
  menu(o) {
    const m = [
      { label: 'Cắt', icon: 'scissors', run: () => { this.clip = JSON.parse(JSON.stringify(o)); this.remove(o); } },
      { label: 'Sao chép', icon: 'copy', run: () => { this.clip = JSON.parse(JSON.stringify(o)); toast('Đã sao chép'); } },
      { label: 'Dán', icon: 'clipboard-paste', disabled: !this.clip, run: () => this.add(Object.assign({}, this.clip, { id: uid('o'), x: this.clip.x + 20, y: this.clip.y + 20 })) },
      { sep: true },
      { label: 'Đưa lên trước', icon: 'bring-to-front', run: () => this.order(o, 1) },
      { label: 'Đưa ra sau', icon: 'send-to-back', run: () => this.order(o, -1) },
      { label: 'Xoay 90°', icon: 'rotate-cw', run: () => this.mutate('Xoay', () => { o.rot = ((o.rot || 0) + 90) % 360; }) },
      { label: 'Văn bản thay thế (Alt Text)…', icon: 'accessibility', run: () => this.altText(o) },
      { sep: true }
    ];
    if (o.type === 'chart') m.push({ label: 'Đổi kiểu biểu đồ', icon: 'chart-column', sub: CHART_TYPES.map(t => ({ label: t[1], icon: t[2], checked: o.kind === t[0] || undefined, run: () => this.mutate('Đổi kiểu biểu đồ', () => { o.kind = t[0]; }) })) }, { label: 'Chọn dữ liệu…', icon: 'table', run: () => this.chartDataDialog(o) }, { label: 'Tiêu đề biểu đồ…', icon: 'heading', run: () => promptDlg('Tiêu đề:', o.title || '', 'Tiêu đề biểu đồ').then(v => { if (v != null) this.mutate('Tiêu đề', () => { o.title = v; }); }) }, { label: (o.labels ? 'Ẩn' : 'Hiện') + ' nhãn dữ liệu', icon: 'tag', run: () => this.mutate('Nhãn dữ liệu', () => { o.labels = !o.labels; }) }, { label: (o.legend === false ? 'Hiện' : 'Ẩn') + ' chú giải', icon: 'list', run: () => this.mutate('Chú giải', () => { o.legend = o.legend === false; }) }, { label: 'Lưu biểu đồ thành ảnh…', icon: 'image-down', run: () => this.saveChartImage(o) });
    if (o.type === 'shape' || o.type === 'textbox') m.push({ label: 'Màu tô…', icon: 'paint-bucket', run: () => colorPicker(document.body, c => this.mutate('Màu tô', () => { o.fill = c; }), { none: 'Không tô' }) }, { label: 'Sửa văn bản', icon: 'type', run: () => this.editObj(o, this.inner.querySelector('[data-id="' + o.id + '"]')) });
    if (o.type === 'image') m.push({ label: 'Thay ảnh…', icon: 'image', run: async () => { const files = await Host.openFiles('image/*'); if (files.length) { const { bytesToDataUrl } = await import('../core/util.js'); this.mutate('Thay ảnh', () => { o.src = bytesToDataUrl(files[0].bytes, /\.png$/i.test(files[0].name) ? 'image/png' : 'image/jpeg'); }); } } }, { label: 'Lưu ảnh…', icon: 'download', run: async () => { const { dataUrlToBytes } = await import('../core/util.js'); Host.saveFile('anh.' + (/png/.test(o.src) ? 'png' : 'jpg'), dataUrlToBytes(o.src), { saveAs: true }); } });
    m.push({ sep: true }, { label: 'Xóa', icon: 'trash-2', run: () => this.remove(o) });
    return m;
  }
  order(o, d) { this.mutate('Thứ tự', () => { const L = this.list; const i = L.indexOf(o); L.splice(i, 1); if (d > 0) L.push(o); else L.unshift(o); }); }
  async saveChartImage(o) { const el = this.inner.querySelector('[data-id="' + o.id + '"] svg'); if (!el) return; const { chartToPng } = await import('./charts.js'); const url = await chartToPng(el); const { dataUrlToBytes } = await import('../core/util.js'); Host.saveFile((o.title || 'bieu-do') + '.png', dataUrlToBytes(url), { saveAs: true }); }
  acceptsDrop(e) {
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f && /^image\//.test(f.type)) { e.preventDefault(); readFileAsDataUrl(f).then(u => this.insertImage(u)); return true; }
    return false;
  }
}
