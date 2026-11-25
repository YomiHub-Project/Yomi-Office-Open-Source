// PDF viewer + toolkit: render (pdf.js), structural edits (pdf-lib), OCR (Tesseract offline), convert, extract — works offline on Android & Windows
import { h, ic, $, $$, clamp, uid, esc, PDFJS, PDFLib, bytesToDataUrl, dataUrlToBytes, fmtSize } from '../core/util.js';
import { Doc, registerApp } from '../core/shell.js';
import { toast, contextMenu, dialog, confirmDlg, promptDlg, formDlg, busy } from '../core/ui.js';
import { settings } from '../core/store.js';
import Host from '../core/host.js';
import { ribbonDef as pdfRibbon } from './ribbon.js';
import * as T from './tools.js';
import * as U from './ui8.js';

export class PdfDoc extends Doc {
  static get kind() { return 'pdf'; }
  static templates() { return [{ id: 'blank', name: 'Công cụ PDF', thumb: () => h('div.tpl-thumb.app', { style: { background: '#d83b01' } }, ic('file-type', 44)) }]; }
  constructor(shell, opts) {
    super(shell, opts);
    this.bytes = null;          // current PDF bytes (Uint8Array) — source of truth
    this.pdf = null;            // pdf.js document
    this.pageInfo = [];         // [{num, w, h, canvas, textLayer, rendered, rendering}]
    this.scale = 1; this.zoom = 100; this.fitMode = 'width';
    this.sel = [];              // selected page indices (thumbnail multi-select)
    this.ocrText = {};          // pageIndex -> text
    this.undoStack = []; this.redoStack = [];
    this.buildUi();
  }
  async init() {
    // "new PDF" = empty toolkit landing; user opens or creates from images
    this.showLanding();
    this.afterLoadMeta();
  }
  async load(bytes, ext, name) {
    bytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    if (ext && /^(png|jpe?g|bmp|gif|webp)$/i.test(ext)) { bytes = await T.imagesToPdf([{ name, bytes }]); }
    await this.setBytes(bytes, true);
  }
  async restore(native) { if (native && native.bytes) await this.setBytes(native instanceof Uint8Array ? native : new Uint8Array(native.bytes), true); }
  snapshot() { return this.bytes ? { bytes: Array.from(this.bytes) } : null; }
  get saveBytes() { return this.bytes; }
  set saveBytes(v) { this.bytes = v; }
  canSaveExt(ext) { return ext === 'pdf'; }
  saveFormats() { return [['pdf', 'Tài liệu PDF', 'file-type'], ['pdf-flat', 'PDF (phẳng hoá/ảnh)', 'image'], ['txt', 'Văn bản (.txt)', 'file-text'], ['html', 'Trang web (.html)', 'globe'], ['docx', 'Tài liệu Word (.docx)', 'file-type-2']]; }
  async save(ext) {
    if (ext === 'pdf' || !ext) return this.bytes;
    if (ext === 'txt') return new TextEncoder().encode(await T.extractText(this));
    if (ext === 'html') return new TextEncoder().encode(await T.toHtml(this));
    if (ext === 'docx') return T.toDocx(this);
    if (ext === 'pdf-flat') return T.flatten(this);
    throw new Error('Không hỗ trợ');
  }

  // ---------------------------------------------------------------- bytes / pdf.js
  async setBytes(bytes, resetHistory) {
    this.bytes = bytes;
    if (resetHistory) { this.undoStack = [bytes]; this.redoStack = []; }
    const _b = busy('Đang mở PDF…');
    try {
      const pdfjs = await PDFJS();
      if (this.pdf) { try { await this.pdf.destroy(); } catch (e) {} }
      const task = pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false, password: this._password || undefined });
      // v0.0.7: password-protected PDFs
      task.onPassword = (update, reason) => { _b.done(); promptDlg(reason === 2 ? 'Sai mật khẩu — nhập lại:' : 'Tệp PDF được bảo vệ. Nhập mật khẩu để mở:', '', 'PDF có mật khẩu', { type: 'password' }).then(pw => { if (pw == null) { task.destroy(); return; } this._password = pw; update(pw); }); };
      this.pdf = await task.promise;
      this.pageInfo = [];
      for (let i = 1; i <= this.pdf.numPages; i++) this.pageInfo.push({ num: i, rendered: false, rendering: false });
      this.cur = Math.min(this.cur || 0, this.pdf.numPages - 1); this.sel = this.sel.filter(i => i < this.pdf.numPages);
      this.renderViewer();
      this.renderThumbs();
      this.afterLoadMeta();
      U.refreshPanels(this);
    } catch (e) { console.error(e); toast('Không đọc được PDF: ' + e.message, 'err'); }
    _b.done();
  }
  pushHistory() { if (this.bytes) { this.undoStack.push(this.bytes); if (this.undoStack.length > 40) this.undoStack.shift(); this.redoStack = []; } }
  async commit(newBytes, label) { this.pushHistory(); await this.setBytes(newBytes); this.markDirty(); if (label) toast(label); }
  canUndo() { return this.undoStack.length > 1; } canRedo() { return this.redoStack.length > 0; }
  async undo() { if (this.undoStack.length < 2) return; this.redoStack.push(this.undoStack.pop()); await this.setBytes(this.undoStack[this.undoStack.length - 1]); this.markDirty(); }
  async redo() { if (!this.redoStack.length) return; const b = this.redoStack.pop(); this.undoStack.push(b); await this.setBytes(b); this.markDirty(); }

  afterLoadMeta() {
    const n = this.pdf ? this.pdf.numPages : 0;
    this.sbPages.textContent = n ? (n + ' trang') : 'Chưa mở tệp';
    this.sbSize.textContent = this.bytes ? fmtSize(this.bytes.length) : '';
  }

  // ---------------------------------------------------------------- UI
  buildUi() {
    this.el.classList.add('pdf-doc');
    this.thumbs = h('div.pdf-thumbs');
    this.viewer = h('div.pdf-viewer');
    this.el.appendChild(this.thumbs);
    this.el.appendChild(this.viewer);
    this.sbPages = h('button.sb-item'); this.sbSize = h('button.sb-item');
    this.statusLeft.appendChild(this.sbPages); this.statusLeft.appendChild(this.sbSize);
    U.buildShell(this);
    this.viewer.addEventListener('scroll', () => this.onScroll());
    this.viewer.addEventListener('contextmenu', e => { const pc = e.target.closest && e.target.closest('.pdf-page'); if (pc) { e.preventDefault(); this.cur = +pc.dataset.i; contextMenu(e, T.pageMenu(this)); } });
  }
  ribbonDef() { return pdfRibbon(this); }
  viewButtons() {
    const mk = (icn, title, fn, on) => { const b = h('button.sb-view' + (on ? '.on' : ''), { title }, ic(icn, 14)); b.addEventListener('click', () => { fn(); $$('.sb-view', b.parentNode).forEach(x => x.classList.remove('on')); b.classList.add('on'); }); return b; };
    return h('div.sb-views', mk('file', 'Một trang', () => U.setLayout(this, 'single'), this.layout === 'single'), mk('scroll-text', 'Cuộn liên tục', () => U.setLayout(this, 'cont'), this.layout === 'cont'), mk('book-open', 'Hai trang liên tục', () => U.setLayout(this, 'twoCont'), this.layout === 'twoCont'), mk('book-open-text', 'Hai trang có trang bìa', () => U.setLayout(this, 'twoCover'), this.layout === 'twoCover'), mk('layout-grid', 'Sắp xếp trang (lưới)', () => this.setFit('grid'), this.fitMode === 'grid'));
  }
  printOptions() { return U.printOptions(this); }
  printPreview(box) { return U.printPreview(this, box); }
  renderPrint(area) { return U.renderPrint(this, area); }
  onShow() { if (!this.bytes) this.showLanding(); else this.relayout(); }
  onLayout() { this.relayout(); }
  setZoom(z) { this.zoom = clamp(z, 10, 500); this.fitMode = 'custom'; this.relayout(); }
  setFit(mode) { this.fitMode = mode; this.el.classList.toggle('pdf-grid', mode === 'grid'); this.relayout(); this.refreshUi(); }
  relayout() {
    if (!this.pdf || !this.pageInfo.length) return;
    const two = this.layout === 'twoCont' || this.layout === 'twoCover' || this.layout === 'two';
    const avail = this.viewer.clientWidth - 48;
    const first = this.pageInfo[this.cur || 0] && this.pageInfo[this.cur || 0]._baseW ? this.pageInfo[this.cur || 0] : this.pageInfo[0];
    const sw = ((this.viewRot || 0) % 180) ? 1 : 0;
    const dim = p => sw ^ (((p.rotate || 0) % 180) ? 1 : 0) ? [p._baseH || 792, p._baseW || 612] : [p._baseW || 612, p._baseH || 792];
    if (this.fitMode === 'width' || this.fitMode === 'page' || this.fitMode === 'grid' || this.fitMode === 'visible') {
      const [bw, bh] = dim(first); const cols = two ? 2 : 1;
      if (this.fitMode === 'width') this.scale = (avail - (cols - 1) * 12) / cols / bw;
      else if (this.fitMode === 'page') this.scale = Math.min((avail - (cols - 1) * 12) / cols / bw, (this.viewer.clientHeight - 48) / bh);
      else if (this.fitMode === 'visible') this.scale = (avail - (cols - 1) * 12) / cols / (bw * 0.86);
      else this.scale = (avail / (this.el.clientWidth < 760 ? 2 : 3) - 16) / bw;
      this.zoom = Math.round(this.scale * 100);
    } else this.scale = this.zoom / 100;
    this.pageInfo.forEach(p => { if (p.el && p._baseW) { const [w, hh] = dim(p); p.el.style.width = Math.round(w * this.scale) + 'px'; p.el.style.height = Math.round(hh * this.scale) + 'px'; p.rendered = false; } });
    if (this.shell._zoomCtl) this.shell._zoomCtl.update();
    this.onScroll(); U.updateNav(this);
  }
  showLanding() {
    this.thumbs.innerHTML = '';
    this.viewer.innerHTML = '';
    const tool = (icn, title, desc, fn) => { const b = h('button.pdf-tile', ic(icn, 30), h('div', h('b', title), h('div.muted', desc))); b.addEventListener('click', fn); return b; };
    const land = h('div.pdf-landing',
      h('div.pdf-land-head', ic('file-type', 48), h('h2', 'Công cụ PDF'), h('p.muted', 'Đọc, chỉnh sửa, OCR, chuyển đổi — hoạt động hoàn toàn ngoại tuyến.')),
      h('div.pdf-tiles',
        tool('folder-open', 'Mở PDF', 'Xem và chỉnh sửa tệp PDF', () => this.shell.openDialog('pdf')),
        tool('image-plus', 'Ảnh → PDF', 'Gộp nhiều ảnh thành một PDF', () => T.imagesToPdfDialog(this)),
        tool('combine', 'Ghép PDF', 'Nối nhiều tệp PDF', () => T.mergeDialog(this)),
        tool('scan-text', 'OCR nhận dạng chữ', 'Trích văn bản từ ảnh/PDF (Việt + Anh)', () => this.shell.openDialog('pdf'))
      ));
    this.viewer.appendChild(land);
  }

  // ---------------------------------------------------------------- rendering pages
  renderViewer() {
    this.viewer.innerHTML = '';
    this.el.classList.toggle('pdf-two', /^two/.test(this.layout)); this.el.classList.toggle('pdf-single', this.layout === 'single');
    const order = this.pageInfo.map((p, i) => i); if (this.reverse) order.reverse();
    let row = null; const two = /^two/.test(this.layout);
    order.forEach((i, k) => {
      const p = this.pageInfo[i];
      const el = h('div.pdf-page', { dataset: { i } }, h('div.pdf-page-no', 'Trang ' + (i + 1)));
      p.el = el; p.rendered = false; if (p.canvas) p.canvas = null; p.textLayer = null;
      if (two) { const startNew = this.layout === 'twoCover' ? (k === 0 || k % 2 === 1) : k % 2 === 0; if (startNew || !row) { row = h('div.pdf-spread'); this.viewer.appendChild(row); } row.appendChild(el); }
      else this.viewer.appendChild(el);
    });
    if (this.layout === 'single') this.showSingle();
    // measure first page then layout
    this.measurePages().then(() => this.relayout());
  }
  showSingle() { this.pageInfo.forEach((p, i) => { if (p.el) p.el.classList.toggle('pdf-on', i === (this.cur || 0)); }); }
  async measurePages() {
    const pdfjs = await PDFJS();
    for (const p of this.pageInfo) {
      const page = await this.pdf.getPage(p.num);
      const vp = page.getViewport({ scale: 1, rotation: 0 });
      p._baseW = vp.width; p._baseH = vp.height; p._page = page;
    }
  }
  onScroll() {
    if (!this.pageInfo.length) return;
    const vr = this.viewer.getBoundingClientRect(); const top = vr.top, bottom = vr.bottom; let best = -1, bestD = 1e9;
    this.pageInfo.forEach((p, i) => {
      if (!p.el || (this.layout === 'single' && i !== (this.cur || 0))) return;
      const r = p.el.getBoundingClientRect();
      const near = r.top < bottom + 600 && r.bottom > top - 600;
      if (near && !p.rendered && !p.rendering) this.renderPage(i);
      if (this.layout !== 'single') { const dd = Math.abs(r.top - top - 10); if (r.bottom > top + 80 && dd < bestD) { bestD = dd; best = i; } }
    });
    if (best >= 0 && best !== this.cur) { this.cur = best; this.highlightThumb(best); }
    if (this.sbPages) this.sbPages.textContent = this.pageInfo.length + ' trang';
    U.updateNav(this);
  }
  async renderPage(i) {
    const p = this.pageInfo[i]; if (!p || !p._page) return; p.rendering = true;
    try {
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const rot = ((p.rotate || 0) + (this.viewRot || 0)) % 360;
      const vp = p._page.getViewport({ scale: this.scale * dpr, rotation: rot });
      const canvas = p.canvas || h('canvas.pdf-cv'); p.canvas = canvas;
      canvas.width = Math.round(vp.width); canvas.height = Math.round(vp.height);
      canvas.style.width = '100%'; canvas.style.height = '100%';
      const ctx = canvas.getContext('2d');
      await p._page.render({ canvasContext: ctx, viewport: vp }).promise;
      if (!canvas.parentNode) p.el.appendChild(canvas);
      p.el.querySelector('.pdf-page-no') && (p.el.querySelector('.pdf-page-no').style.display = 'none');
      await this.renderTextLayer(i, rot);
      p.rendered = true;
    } catch (e) { /* ignore */ } p.rendering = false;
  }
  async renderTextLayer(i, rot) {
    if (!settings().pdf || settings().pdf.textLayer !== false) {
      const p = this.pageInfo[i];
      try {
        const tc = await p._page.getTextContent();
        const vp = p._page.getViewport({ scale: this.scale, rotation: rot });
        let layer = p.textLayer; if (!layer) { layer = h('div.pdf-textlayer'); p.textLayer = layer; p.el.appendChild(layer); }
        layer.innerHTML = ''; layer.style.width = vp.width + 'px'; layer.style.height = vp.height + 'px';
        tc.items.forEach(it => {
          if (!it.str) return;
          const tx = it.transform; const x = tx[4] * this.scale, y = vp.height - tx[5] * this.scale;
          const span = h('span', it.str); span.style.left = x + 'px'; span.style.top = (y - it.height * this.scale) + 'px'; span.style.fontSize = (it.height * this.scale) + 'px';
          layer.appendChild(span);
        });
      } catch (e) {}
    }
  }
  reRenderAll() { this.pageInfo.forEach(p => { p.rendered = false; }); this.onScroll(); }
  goto(i) {
    if (!this.pageInfo.length) return; i = clamp(i, 0, this.pageInfo.length - 1); const p = this.pageInfo[i];
    if (this.layout === 'single') { this.cur = i; this.showSingle(); this.viewer.scrollTop = 0; this.highlightThumb(i); this.onScroll(); return; }
    if (p && p.el) { const vr = this.viewer.getBoundingClientRect(), r = p.el.getBoundingClientRect(); this.viewer.scrollTop += r.top - vr.top - 12; this.cur = i; this.highlightThumb(i); U.updateNav(this); }
  }

  // ---------------------------------------------------------------- thumbnails
  renderThumbs() {
    this.thumbs.innerHTML = '';
    this.pageInfo.forEach((p, i) => {
      const th = h('div.pdf-thumb' + (i === this.cur ? '.on' : ''), { dataset: { i } }, h('span.pdf-thumb-no', i + 1));
      const holder = h('div.pdf-thumb-cv'); th.appendChild(holder); p.thumbHolder = holder;
      th.addEventListener('click', e => { if (e.ctrlKey || e.metaKey) { const k = this.sel.indexOf(i); if (k >= 0) this.sel.splice(k, 1); else this.sel.push(i); th.classList.toggle('msel'); } else { this.sel = [i]; this.goto(i); } });
      th.addEventListener('contextmenu', e => { e.preventDefault(); if (this.sel.indexOf(i) < 0) this.sel = [i]; this.cur = i; contextMenu(e, T.pageMenu(this)); });
      this.thumbs.appendChild(th);
      this.renderThumbImg(i);
    });
  }
  async renderThumbImg(i) {
    const p = this.pageInfo[i]; if (!p || !p._page || !p.thumbHolder) { if (p && !p._page) { await this.measureOne(i); } }
    const pg = p._page || await this.measureOne(i); if (!pg) return;
    try {
      const vp = pg.getViewport({ scale: 1, rotation: p.rotate || 0 });
      const s = 150 / vp.width; const v2 = pg.getViewport({ scale: s, rotation: p.rotate || 0 });
      const c = h('canvas'); c.width = v2.width; c.height = v2.height; c.style.width = '100%';
      await pg.render({ canvasContext: c.getContext('2d'), viewport: v2 }).promise;
      if (p.thumbHolder) { p.thumbHolder.innerHTML = ''; p.thumbHolder.appendChild(c); }
    } catch (e) {}
  }
  async measureOne(i) { const p = this.pageInfo[i]; if (!p) return null; if (!p._page) { p._page = await this.pdf.getPage(p.num); const vp = p._page.getViewport({ scale: 1 }); p._baseW = vp.width; p._baseH = vp.height; } return p._page; }
  highlightThumb(i) { $$('.pdf-thumb', this.thumbs).forEach(t => t.classList.toggle('on', +t.dataset.i === i)); const t = this.thumbs.children[i]; if (t) t.scrollIntoView({ block: 'nearest' }); }

  getText() { return Object.keys(this.ocrText).length ? Object.values(this.ocrText).join('\n\n') : ''; }
  stats() { return { 'Số trang': this.pdf ? this.pdf.numPages : 0, 'Kích thước': this.bytes ? fmtSize(this.bytes.length) : '—' }; }
  commands() { return T.extraCommands(this); }
  contextItems() { return T.pageMenu(this); }
  print() { if (this.pdf) this.shell.backstage(this, 'print'); else toast('Mở một PDF trước', 'err'); }
  destroy() { if (this.pdf) try { this.pdf.destroy(); } catch (e) {} }
}
registerApp('pdf', PdfDoc);
// v0.0.8: pages still rendering when a document is replaced/closed reject with "Transport destroyed" — expected, not an error
window.addEventListener('unhandledrejection', e => { const m = e.reason && (e.reason.message || String(e.reason)); if (m && /Transport destroyed|Worker was destroyed|Rendering cancelled|Worker was terminated/.test(m)) e.preventDefault(); });
