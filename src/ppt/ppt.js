// PowerPoint document: slide deck editor with draggable objects, thumbnails, transitions, animations, slideshow, PPTX read/write
import { h, ic, $, $$, esc, keyName, chord, clamp, uid, deepClone, bytesToDataUrl, KaTeX } from '../core/util.js';
import { Doc, registerApp } from '../core/shell.js';
import { toast, showMenu, contextMenu, dialog, promptDlg, confirmDlg, formDlg, colorPicker, longPress } from '../core/ui.js';
import { settings } from '../core/store.js';
import { ribbonDef as pptRibbon } from './ribbon.js';
import { unifyRibbon } from '../core/unify.js';
import * as P from './cmds.js';
import { readPptx, writePptx } from './pptx.js';
import { THEMES, LAYOUTS, templates as pTemplates, applyTemplate, blankSlide } from './templates.js';
import { InkTool, renderInk as inkRender, PENS } from '../core/ink.js';

export const SW = 1280, SH = 720; // slide canvas logical size (16:9)

export class PptDoc extends Doc {
  static get kind() { return 'ppt'; }
  static templates() { return pTemplates(); }
  constructor(shell, opts) {
    super(shell, opts);
    this.slides = []; this.cur = 0; this.sel = []; this.theme = THEMES[0]; this.ratio = '16:9';
    this.undoStack = []; this.redoStack = []; this.clip = null; this.zoom = 100;
    this.buildUi();
  }
  async init() {
    const t = this.opts.template;
    if (t && t.id !== 'blank') await applyTemplate(this, t);
    else { this.slides = [blankSlide('title')]; }
    this.afterLoad();
  }
  async load(bytes, ext) {
    if (ext === 'xml') { const { flatOpcToZip } = await import('../formats/extra.js'); bytes = await flatOpcToZip(new TextDecoder('utf-8').decode(bytes)); ext = 'pptx'; }
    if (ext === 'pptx' || ext === 'potx' || ext === 'pptm' || ext === 'ppsx' || ext === 'ppsm') { const r = await readPptx(bytes); this.slides = r.slides; this.theme = r.theme || THEMES[0]; this.ratio = r.ratio || '16:9'; if (r.vba) this.vba = r.vba; }
    else if (ext === 'ppt' || ext === 'pps' || ext === 'pot') { const { readPpt } = await import('../formats/legacy.js'); const list = readPpt(bytes); this.slides = list.map((texts, i) => { const s = blankSlide(i === 0 ? 'title' : 'content'); s.objects[0].html = esc(texts[0] || ''); if (s.objects[1]) s.objects[1].html = texts.slice(1).map(t => '<div>' + esc(t).replace(/\n/g, '<br>') + '</div>').join(''); return s; }); this.theme = THEMES[0]; }
    else if (ext === 'odp' || ext === 'otp') { const { readOdp } = await import('../formats/odf.js'); const r = await readOdp(bytes, SW, SH, blankSlide); this.slides = r.slides; this.theme = THEMES[0]; this.ratio = '16:9'; }
    else { const t = new TextDecoder().decode(bytes); this.slides = P.outlineToSlides(t); }
    if (!this.slides.length) this.slides = [blankSlide('title')];
    this.afterLoad();
  }
  async restore(native) { this.slides = native.slides; this.theme = native.theme || THEMES[0]; this.ratio = native.ratio; this.afterLoad(); }
  snapshot() { return { slides: this.slides, theme: this.theme, ratio: this.ratio }; }
  afterLoad() { this.cur = 0; this.renderThumbs(); this.renderSlide(); this.pushHistory(true); }
  canSaveExt(ext) { return ['pptx', 'odp', 'potx', 'ppsx', 'xml'].indexOf(ext) >= 0; }
  saveFormats() { return [['pptx', 'Bản trình bày PowerPoint', 'presentation'], ['pptm', 'Bản trình bày có Macro (*.pptm)', 'square-code'], ['pdf-print', 'PDF (qua In)', 'file-type'], ['html', 'Trình chiếu HTML', 'globe'], ['md', 'Dàn ý Markdown', 'hash'], ['png', 'Ảnh PNG (slide hiện tại)', 'image'], ['odp', 'Trình chiếu OpenDocument (*.odp)', 'presentation'], ['potx', 'Mẫu PowerPoint (*.potx)', 'file-plus'], ['ppsx', 'Trình chiếu PowerPoint (*.ppsx) — mở là chiếu', 'play'], ['xml', 'Bản trình bày PowerPoint XML (*.xml)', 'file-code'], ['rtf', 'Dàn ý / RTF (*.rtf)', 'file-text'], ['jpg', 'Ảnh JPEG (slide hiện tại)', 'image'], ['video', 'Video trình chiếu (*.webm / *.mp4)', 'film', 'webm']]; }
  async save(ext) {
    if (ext === 'pptx') return writePptx(this);
    if (ext === 'pptm') return writePptx(this, true);
    if (ext === 'odp') { const { writeOdp } = await import('../formats/odf.js'); return writeOdp(this, SW, SH); }
    if (ext === 'potx' || ext === 'ppsx') { const X = await import('../formats/extra.js'); return X.retypePackage(await writePptx(this), X.CT.pptx, X.CT[ext]); }
    if (ext === 'xml') { const X = await import('../formats/extra.js'); return X.zipToFlatOpc(await writePptx(this), 'PowerPoint.Show'); }
    if (ext === 'rtf') { const X = await import('../formats/extra.js'); return X.outlineRtf(this.slides, P.plain); }
    if (ext === 'jpg') { const url = await P.slideToPng(this, this.cur, 'image/jpeg'); return (await import('../core/util.js')).dataUrlToBytes(url); }
    if (ext === 'video') { const R = await import('./record.js'); return R.exportVideo(this); }
    if (ext === 'html') return new TextEncoder().encode(P.exportHtml(this));
    if (ext === 'md') return new TextEncoder().encode(P.exportOutline(this));
    if (ext === 'png') { const url = await P.slideToPng(this, this.cur); return (await import('../core/util.js')).dataUrlToBytes(url); }
    throw new Error('Không hỗ trợ');
  }
  get slide() { return this.slides[this.cur]; }

  buildUi() {
    this.el.classList.add('pp-doc');
    this.thumbs = h('div.pp-thumbs');
    this.stage = h('div.pp-stage');
    this.canvas = h('div.pp-canvas');
    this.stage.appendChild(this.canvas);
    this.notesEl = h('div.pp-notes', { contenteditable: 'true', 'data-ph': 'Bấm để thêm ghi chú' });
    const main = h('div.pp-main', this.stage, this.notesEl);
    this.el.appendChild(this.thumbs);
    this.el.appendChild(main);
    const NS = 'http://www.w3.org/2000/svg';
    this._inkSvg = document.createElementNS(NS, 'svg'); this._inkSvg.setAttribute('class', 'pp-ink');
    this._inkCatch = h('div.pp-inkcatch');
    this.sbSlideNo = h('button.sb-item'); this.sbTheme = h('span.sb-item'); this.sbLang = h('span.sb-item', ic('n:language', 14), ' Tiếng Việt'); this.sbAcc = h('button.sb-item', ic('n:access', 14), ' Trợ năng: Tốt'); this.sbAcc.addEventListener('click', () => this.shell.inspect(this)); this.sbNotes = h('button.sb-item', { title: 'Trang ghi chú' }, ic('n:notes', 14), ' Ghi chú');
    this.statusLeft.appendChild(this.sbSlideNo); this.statusLeft.appendChild(this.sbLang); this.statusLeft.appendChild(this.sbAcc); this.statusLeft.appendChild(this.sbNotes);   // v0.1.7: native order
    this.sbNotes.addEventListener('click', () => this.el.classList.toggle('pp-notes-off'));
    this.sbSlideNo.addEventListener('click', () => {});
    this.notesEl.addEventListener('input', () => { if (this.slide) { this.slide.notes = this.notesEl.innerHTML; this.markDirty(); } });
    this.notesEl.addEventListener('focus', () => {});
    document.addEventListener('keydown', e => { if (this.shell.cur !== this || this._show) return; const k = (e.altKey ? 'alt+' : '') + String(e.key || '').toLowerCase(); if (k === 'alt+f8' || k === 'alt+f11') { e.preventDefault(); import('../excel/vba/index.js').then(m => k === 'alt+f8' ? m.macrosDialog(this) : m.openEditor(this)); return; } P.onKey(this, e); }, false);
  }
  viewButtons() {
    const mk = (icn, title, fn, on) => { const b = h('button.sb-view' + (on ? '.on' : ''), { title }, ic(icn, 14)); b.addEventListener('click', fn); return b; };
    return h('div.sb-views', mk('n:normalview', 'Thường', () => {}, true), mk('n:sorter', 'Sắp xếp trang chiếu', () => P.slideSorter(this)), mk('n:readmode', 'Dạng xem đọc', () => P.startShow(this, this.cur)), mk('n:play', 'Trình chiếu từ trang hiện tại', () => P.startShow(this, this.cur)));
  }
  ribbonDef() { return unifyRibbon(this, pptRibbon(this)); }   // v0.1.7: same layout as the native editor
  onShow() { this.fit(); this.renderThumbs(); this.renderSlide(); this.onSelection(); }
  onLayout() { this.fit(); }
  focus() {}
  setZoom(z) { this.zoom = z; this.fit(z); }
  fit(forceZoom) {
    const r = this.stage.getBoundingClientRect(); if (!r.width) return;
    const z = forceZoom ? forceZoom / 100 : Math.min((r.width - 48) / SW, (r.height - 48) / SH);
    this.scale = z;
    this.canvas.style.width = SW + 'px'; this.canvas.style.height = SH + 'px';
    this.canvas.style.webkitTransform = this.canvas.style.transform = 'translate(-50%,-50%) scale(' + z + ')';
  }
  getText() { return this.slides.map((s, i) => '--- Slide ' + (i + 1) + ' ---\n' + s.objects.filter(o => o.type === 'text').map(o => P.objText(o)).join('\n') + (s.notes ? '\n[Ghi chú] ' + P.plain(s.notes) : '')).join('\n\n'); }
  stats() { return { 'Số slide': this.slides.length, 'Chủ đề': this.theme.name, 'Tỉ lệ': this.ratio, 'Đối tượng': this.slides.reduce((n, s) => n + s.objects.length, 0) }; }
  inspect() { const issues = []; this.slides.forEach((s, i) => { s.objects.forEach(o => { if (o.type === 'image' && !o.alt) issues.push('Slide ' + (i + 1) + ': ảnh thiếu văn bản thay thế.'); }); if (!s.objects.some(o => o.type === 'text' && o.ph === 'title')) issues.push('Slide ' + (i + 1) + ' không có tiêu đề (khó điều hướng).'); }); return issues; }

  // ---------------------------------------------------------------- rendering
  renderThumbs() {
    this.thumbs.innerHTML = '';
    this.slides.forEach((s, i) => {
      const th = h('div.pp-thumb' + (i === this.cur ? '.on' : '') + (this.sel.indexOf(i) >= 0 ? '.msel' : ''), { draggable: 'true', dataset: { i } }, h('span.pp-thumb-no', i + 1));
      const mini = h('div.pp-mini'); P.renderSlideInto(this, s, mini, 0.148); th.appendChild(mini);
      if (s.transition) th.appendChild(h('span.pp-thumb-tr', { title: 'Có hiệu ứng chuyển tiếp' }, ic('star', 11)));
      th.addEventListener('click', e => { if (e.ctrlKey || e.metaKey) { const k = this.sel.indexOf(i); if (k >= 0) this.sel.splice(k, 1); else this.sel.push(i); } else this.sel = [i]; this.goto(i); });
      th.addEventListener('contextmenu', e => { if (this.sel.indexOf(i) < 0) this.sel = [i]; this.goto(i); contextMenu(e, P.slideMenu(this)); });
      longPress(th, e => { this.sel = [i]; this.goto(i); contextMenu(e, P.slideMenu(this)); });
      th.addEventListener('dragstart', e => e.dataTransfer.setData('text/plain', 'slide:' + i));
      th.addEventListener('dragover', e => e.preventDefault());
      th.addEventListener('drop', e => { e.preventDefault(); const d = e.dataTransfer.getData('text/plain'); if (/^slide:/.test(d)) P.moveSlide(this, +d.slice(6), i); });
      this.thumbs.appendChild(th);
    });
    const add = h('button.pp-thumb-add', { title: 'Slide mới (Ctrl+M)' }, ic('plus', 20));
    add.addEventListener('click', () => P.newSlide(this));
    this.thumbs.appendChild(add);
  }
  renderSlide() {
    this.canvas.innerHTML = '';
    const s = this.slide; if (!s) return;
    P.renderSlideInto(this, s, this.canvas, 1, true);
    this.canvas.appendChild(this._inkSvg); this.canvas.appendChild(this._inkCatch);
    this.renderInk();
    const so = this.selectedObj; if (so && s.objects.indexOf(so) >= 0) { const se = this.canvas.querySelector('[data-id="' + so.id + '"]'); if (se) { se.classList.add('sel'); P.addHandles(se); } } else if (so) this.selectedObj = null;
    this.notesEl.innerHTML = s.notes || '';
    this.sbSlideNo.textContent = 'Trang chiếu ' + (this.cur + 1) + ' / ' + this.slides.length;
    this.sbTheme.textContent = this.theme.name;
    this.fit(this.zoom !== 100 && this._userZoom ? this.zoom : undefined);
  }
  goto(i) { if (i < 0 || i >= this.slides.length) return; this.cur = clamp(i, 0, this.slides.length - 1); this.selectedObj = null; this.renderThumbs(); this.renderSlide(); this.onSelection(); const th = this.thumbs.children[this.cur]; if (th) th.scrollIntoView({ block: 'nearest' }); }
  onSelection() { if (this.ribbon) this.ribbon.refresh(); this.updateContextualTabs(); }
  refreshUi() { if (this.ribbon) this.ribbon.refresh(); }
  updateContextualTabs() { const o = this.selectedObj; const key = o ? o.type : ''; if (this._ctxKey !== key && this.ribbon) { this._ctxKey = key; this.ribbon.def = this.ribbonDef(); this.ribbon.renderTabs(); if (!this.ribbon.def.tabs.find(t => t.id === this.ribbon.cur)) this.ribbon.select('home'); } }

  // ---------------------------------------------------------------- history
  pushHistory(initial) { const snap = JSON.stringify({ slides: this.slides, cur: this.cur }); if (!initial && this.undoStack.length && this.undoStack[this.undoStack.length - 1] === snap) return; this.undoStack.push(snap); if (this.undoStack.length > 120) this.undoStack.shift(); if (!initial) this.redoStack = []; }
  change(label, fn) { if (this.readOnly) return; this.pushHistory(); fn(); this.pushHistory(); this.markDirty(); this.renderThumbs(); this.renderSlide(); this.onSelection(); }
  undo() { if (this.undoStack.length < 2) return; const cur = this.undoStack.pop(); this.redoStack.push(cur); this.applySnap(this.undoStack[this.undoStack.length - 1]); }
  redo() { if (!this.redoStack.length) return; const snap = this.redoStack.pop(); this.undoStack.push(snap); this.applySnap(snap); }
  applySnap(snap) { const o = JSON.parse(snap); this.slides = o.slides; this.cur = clamp(o.cur, 0, this.slides.length - 1); this.selectedObj = null; this.markDirty(); this.renderThumbs(); this.renderSlide(); this.onSelection(); }
  canUndo() { return this.undoStack.length > 1; } canRedo() { return this.redoStack.length > 0; }

  // ---------------------------------------------------------------- ink (Draw tab)
  renderInk() { if (!this._inkSvg) return; this._inkSvg.setAttribute('width', SW); this._inkSvg.setAttribute('height', SH); this._inkSvg.setAttribute('viewBox', '0 0 ' + SW + ' ' + SH); inkRender(this._inkSvg, this.inkHidden ? [] : (this.slide ? (this.slide.ink || []) : []), 1); }
  getInk() {
    if (this._ink) return this._ink;
    const doc = this; let tool = null;
    const ink = { pen: Object.assign({}, PENS[0]), get tool() { return tool; }, setTool(t, pen) { tool = t; if (pen) this.pen = pen; doc._inkCatch.classList.toggle('on', !!t); if (doc._inkTool) { doc._inkTool.tool = t; doc._inkTool.pen = this.pen; } } };
    this._inkTool = new InkTool({
      toLocal: e => { const r = doc.canvas.getBoundingClientRect(); const z = doc.scale || 1; return [(e.clientX - r.left) / z, (e.clientY - r.top) / z]; },
      getStrokes: () => { if (!doc.slide.ink) doc.slide.ink = []; return doc.slide.ink; },
      scale: () => doc.scale || 1, begin: () => doc.pushHistory(), commit: () => { doc.markDirty(); doc.renderThumbs(); }, redraw: () => doc.renderInk()
    });
    this._inkTool.tool = null; this._inkTool.attach(this._inkCatch);
    this.clearInk = () => { doc.change('Xóa nét vẽ', () => { doc.slide.ink = []; }); };
    this._ink = ink; return ink;
  }

  contextItems() { return P.slideMenu(this); }
  commands() { return P.extraCommands(this); }
  renderPrint(area) { return P.renderPrint(this, area); }
  printPreview(box) { return P.printPreview(this, box); }
  printOptions() { return P.printOptions(this); }
  setReadOnly(v) { this.readOnly = v; this.canvas.classList.toggle('ro', v); }
  onBack() { if (this._show) { P.endShow(this); return true; } return false; }
  destroy() { if (this._show) P.endShow(this); }
}
registerApp('ppt', PptDoc);
