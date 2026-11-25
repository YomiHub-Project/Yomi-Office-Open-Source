// Word document: paginated WYSIWYG editor (contenteditable), full ribbon, styles, DOCX read/write, comments, track changes
import { h, ic, $, $$, esc, keyName, chord, clamp, debounce, uid, utf8Encode, utf8Decode, extOf, fmtDate, stripHtml } from '../core/util.js';
import { Doc, registerApp } from '../core/shell.js';
import { toast, showMenu, contextMenu, dialog, promptDlg, confirmDlg, formDlg, colorPicker, longPress, busy } from '../core/ui.js';
import { settings, saveSettings } from '../core/store.js';
import { ribbonDef as wordRibbon } from './ribbon.js';
import { unifyRibbon } from '../core/unify.js';
import * as W from './cmds.js';
import { readDocx, writeDocx, mdToHtml, htmlToMd } from './docx.js';
import { htmlToPlain } from './cmds.js';
import { Ruler } from './ruler.js';
import { templates as wTemplates, applyTemplate } from './templates.js';

const PAGE = { A4: [794, 1123], Letter: [816, 1056], A5: [559, 794], Legal: [816, 1344], A3: [1123, 1587] };
const MARG = { normal: [96, 96, 96, 96], narrow: [48, 48, 48, 48], moderate: [96, 72, 96, 72], wide: [96, 192, 96, 192] };

export class WordDoc extends Doc {
  static get kind() { return 'word'; }
  static templates() { return wTemplates(); }
  constructor(shell, opts) {
    super(shell, opts);
    this.pageSize = 'A4'; this.margins = 'normal'; this.orientation = 'portrait';
    this.undoStack = []; this.redoStack = []; this._histTimer = null;
    this.trackChanges = false; this.showMarks = false; this.comments = [];
    this.zoom = 100; this.viewMode = 'print';
    this.buildUi();
  }
  async init() {
    this.body.innerHTML = '';
    const t = this.opts.template;
    if (t && t.id !== 'blank') await applyTemplate(this, t);
    else this.body.appendChild(this.para(''));
    this.afterLoad();
  }
  async load(bytes, ext, name) {
    if (ext === 'xml' && /pkg:package/.test(new TextDecoder('utf-8').decode(bytes.subarray(0, 2000)))) { const { flatOpcToZip } = await import('../formats/extra.js'); bytes = await flatOpcToZip(utf8Decode(bytes)); ext = 'docx'; }
    if (ext === 'mht' || ext === 'mhtml') { const { readMht } = await import('../formats/extra.js'); const html = readMht(new TextDecoder('latin1').decode(bytes)); const m = /<body[^>]*>([\s\S]*?)<\/body>/i.exec(html); this.body.innerHTML = W.sanitize(m ? m[1] : html); W.normalizeDoc(this); this.afterLoad(); return; }
    if (ext === 'docx' || ext === 'dotx' || ext === 'docm' || ext === 'dotm') { const r = await readDocx(bytes); this.body.innerHTML = r.html; this.docDefaults = r.defaults; if (r.sectPr) this.applySect(r.sectPr); this.comments = r.comments || []; this.vba = r.vba || null; }
    else if (ext === 'html' || ext === 'htm') { const t = utf8Decode(bytes); const m = /<body[^>]*>([\s\S]*?)<\/body>/i.exec(t); this.body.innerHTML = W.sanitize(m ? m[1] : t); }
    else if (ext === 'md' || ext === 'markdown') this.body.innerHTML = mdToHtml(utf8Decode(bytes));
    else if (ext === 'rtf') this.body.innerHTML = W.rtfToHtml(new TextDecoder('latin1').decode(bytes));
    else if (ext === 'doc' || ext === 'dot' || ext === 'wps') { const t0 = new TextDecoder('latin1').decode(bytes.subarray(0, 6)); if (/^\{\\rtf/.test(t0)) this.body.innerHTML = W.rtfToHtml(new TextDecoder('latin1').decode(bytes)); else if (bytes[0] === 0xD0 && bytes[1] === 0xCF) { const { readDoc } = await import('../formats/legacy.js'); this.body.innerHTML = readDoc(bytes).html; } else this.body.innerHTML = '<p>' + esc(utf8Decode(bytes)) + '</p>'; }
    else if (ext === 'odt' || ext === 'ott') { const { readOdt } = await import('../formats/odf.js'); const r = await readOdt(bytes); this.body.innerHTML = r.html; if (r.sect) this.applySect(r.sect); }
    else { const t = utf8Decode(bytes); this.body.innerHTML = t.split(/\n{2,}/).map(p => '<p>' + esc(p).replace(/\n/g, '<br>') + '</p>').join('') || '<p></p>'; }
    if (!this.body.firstChild) this.body.appendChild(this.para(''));
    W.normalizeDoc(this);
    this.afterLoad();
  }
  async restore(native) { this.body.innerHTML = native.html; this.pageSize = native.pageSize || 'A4'; this.margins = native.margins || 'normal'; this.orientation = native.orientation || 'portrait'; this.comments = native.comments || []; this.afterLoad(); }
  snapshot() { return { html: this.body.innerHTML, pageSize: this.pageSize, margins: this.margins, orientation: this.orientation, comments: this.comments }; }
  afterLoad() { this.applyPageStyle(); this.updatePages(); this.updateRuler(); this.pushHistory(true); this.updateStatus(); }
  applySect(s) { if (s.size) this.pageSize = s.size; if (s.orientation) this.orientation = s.orientation; if (s.margins) this.margins = s.margins; }
  canSaveExt(ext) { return ['docx', 'docm', 'txt', 'md', 'html', 'odt', 'rtf', 'dotx', 'doc', 'mht', 'xml'].indexOf(ext) >= 0; }
  saveFormats() { return [['docx', 'Tài liệu Word', 'file-text'], ['docm', 'Tài liệu Word có Macro (*.docm)', 'square-code'], ['pdf-print', 'PDF (qua In)', 'file-type'], ['html', 'Trang web (HTML)', 'globe'], ['txt', 'Văn bản thuần', 'file'], ['md', 'Markdown', 'hash'], ['rtf', 'Rich Text Format', 'file-type'], ['odt', 'Văn bản OpenDocument (*.odt)', 'file-text'], ['dotx', 'Mẫu Word (*.dotx)', 'file-plus'], ['doc', 'Word 97-2003 (*.doc) — dạng RTF tương thích', 'file-text'], ['mht', 'Trang web một tệp (*.mht)', 'globe'], ['xml', 'Tài liệu Word XML (*.xml)', 'file-code']]; }
  async save(ext) {
    this.normalizeForSave();
    if (ext === 'docx') { if (this.vba && this.vba.modules && this.vba.modules.some(m => /\b(sub|function)\b/i.test(m.code)) && !this._vbaWarned) { this._vbaWarned = true; toast('Tài liệu có macro VBA — định dạng .docx không lưu macro. Hãy dùng "Lưu như" > Tài liệu Word có Macro (*.docm).', 'warn'); } return writeDocx(this); }
    if (ext === 'docm') return writeDocx(this, { macro: true });
    if (ext === 'html') return utf8Encode(W.exportHtml(this));
    if (ext === 'md') return utf8Encode(htmlToMd(this.body));
    if (ext === 'rtf') return utf8Encode(W.htmlToRtf(this.body));
    if (ext === 'odt') { const { writeOdt } = await import('../formats/odf.js'); return writeOdt(this); }
    if (ext === 'dotx') { const X = await import('../formats/extra.js'); return X.retypePackage(await writeDocx(this), X.CT.docx, X.CT.dotx); }
    if (ext === 'doc') return utf8Encode(W.htmlToRtf(this.body));
    if (ext === 'mht') { const X = await import('../formats/extra.js'); return utf8Encode(X.writeMht(this.body.innerHTML, this.name)); }
    if (ext === 'xml') { const X = await import('../formats/extra.js'); return X.zipToFlatOpc(await writeDocx(this), 'Word.Document'); }
    return utf8Encode(htmlToPlain(this.body));
  }
  normalizeForSave() { W.clearSelectionArtifacts(this); }

  // ---------------------------------------------------------------- UI
  buildUi() {
    this.el.classList.add('wd-doc');
    this.rulerObj = new Ruler(this); this.ruler = this.rulerObj.el;
    this.mainEl = h('div.wd-main');
    this.scroller = h('div.wd-scroll', { tabindex: -1 });
    this.pagesEl = h('div.wd-pages');
    this.sheet = h('div.wd-sheet');
    this.body = h('div.wd-body', { contenteditable: 'true', spellcheck: 'false', autocapitalize: 'off' });
    this.sheet.appendChild(this.body);
    this.pagesEl.appendChild(this.sheet);
    this.scroller.appendChild(this.pagesEl);
    this.el.appendChild(this.ruler);
    this.mainEl.appendChild(this.rulerObj.vel); this.mainEl.appendChild(this.scroller);
    this.el.appendChild(this.mainEl);
    // status items
    this.sbPage = h('button.sb-item'); this.sbWords = h('button.sb-item'); this.sbLang = h('span.sb-item', ic('n:language', 14), ' Tiếng Việt'); this.sbAcc = h('button.sb-item', { title: 'Kiểm tra trợ năng' }, ic('n:access', 14), ' Trợ năng: Tốt'); this.sbAcc.addEventListener('click', () => this.shell.inspect(this));   // v0.1.7: same segments as the native status bar
    this.statusLeft.appendChild(this.sbPage); this.statusLeft.appendChild(this.sbWords); this.statusLeft.appendChild(this.sbLang); this.statusLeft.appendChild(this.sbAcc);
    this.sbTrack = h('button.sb-item.hidden', ic('git-compare', 13), ' Theo dõi: Bật');
    this.statusLeft.appendChild(this.sbTrack);
    this.sbPage.addEventListener('click', () => W.gotoPageDialog(this));
    this.sbWords.addEventListener('click', () => W.wordCountDialog(this));
    this.bind();
  }
  viewButtons() {
    const mk = (icn, title, mode) => { const b = h('button.sb-view' + (this.viewMode === mode ? '.on' : ''), { title }, ic(icn, 14)); b.addEventListener('click', () => this.setView(mode)); return b; };
    return h('div.sb-views', mk('n:readmode', 'Chế độ đọc', 'read'), mk('n:printlayout', 'Bố cục in', 'print'), mk('n:weblayout', 'Bố cục web', 'web'));
  }
  setView(mode) { this.viewMode = mode; this.el.classList.toggle('view-web', mode === 'web'); this.el.classList.toggle('view-read', mode === 'read'); this.updatePages(); $$('.sb-view', this.statusRight.parentNode).forEach(b => {}); this.refreshUi(); }
  ribbonDef() { return unifyRibbon(this, wordRibbon(this)); }   // v0.1.7: same layout as the native editor
  onShow() { this.applyPageStyle(); this.updatePages(); this.autoFit(); this.updateRuler(); this.updateStatus(); this.onSelection(); }
  onLayout() { this.updatePages(); this.autoFit(); this.updateRuler(); }
  // phones / narrow windows: fit the page width to the screen unless the user picked a zoom
  autoFit() {
    if (this._userZoom || !this.pageW || !this.scroller) return;
    const avail = this.scroller.clientWidth - 16; if (avail <= 0) return;
    const z = avail < this.pageW ? Math.max(30, Math.floor(avail / this.pageW * 100)) : 100;
    if (z !== this.zoom) { this._autoZoom = true; this.setZoom(z); this._autoZoom = false; if (this.shell._zoomCtl && this.shell.cur === this) this.shell._zoomCtl.update(); }
  }
  focus() { this.body.focus(); this.restoreCaret(); }
  setZoom(z) {
    if (!this._autoZoom) this._userZoom = true;
    this.zoom = z;
    const k = z / 100, P = this.pagesEl;
    const pw = (this.pageW || 794) * k;
    const avail = this.scroller.clientWidth;
    if (avail > 0 && pw > avail - 32) {
      P.style.webkitTransformOrigin = P.style.transformOrigin = '0 0';
      const offsetL = Math.max(0, (P.offsetWidth - (this.pageW || 794)) / 2);
      const tx = 16 - offsetL * k;
      P.style.webkitTransform = P.style.transform = 'translateX(' + tx + 'px) scale(' + k + ')';
      P.style.marginBottom = -Math.round(P.offsetHeight * (1 - k)) + 'px';
      P.style.marginRight = Math.round(pw + 32 - avail) + 'px';
    } else if (k < 1) {
      P.style.webkitTransformOrigin = P.style.transformOrigin = '0 0';
      const baseW = this.pageW || P.offsetWidth;
      const offsetL = Math.max(0, (P.offsetWidth - baseW) / 2);
      const tx = Math.max(0, (avail - baseW * k) / 2 - offsetL * k);
      P.style.webkitTransform = P.style.transform = 'translateX(' + tx + 'px) scale(' + k + ')';
      P.style.marginBottom = -Math.round(P.offsetHeight * (1 - k)) + 'px';
      P.style.marginRight = '';
    } else {
      P.style.webkitTransformOrigin = P.style.transformOrigin = 'top center';
      P.style.webkitTransform = P.style.transform = 'scale(' + k + ')';
      P.style.marginBottom = '';
      P.style.marginRight = '';
    }
    this.updateRuler();
  }
  getText() { return htmlToPlain(this.body); }
  stats() { const t = this.getText(); const words = (t.match(/\S+/g) || []).length; return { 'Số trang': this.pageCount || 1, 'Số từ': words, 'Số ký tự': t.replace(/\s/g, '').length, 'Số ký tự (có cách)': t.length, 'Đoạn văn': $$('p,h1,h2,h3,h4,h5,h6,li', this.body).length, 'Chú thích': this.comments.length }; }
  inspect() { const issues = []; $$('img', this.body).forEach(im => { if (!im.alt) issues.push('Có ảnh thiếu văn bản thay thế (Alt Text).'); }); if (!$$('h1,h2,h3', this.body).length && this.getText().length > 1200) issues.push('Tài liệu dài nhưng chưa dùng Kiểu đề mục (Heading) — khó điều hướng với trình đọc màn hình.'); $$('table', this.body).forEach(tb => { if (!tb.querySelector('th')) issues.push('Có bảng không có hàng tiêu đề.'); }); return issues; }

  // ---------------------------------------------------------------- page geometry
  para(txt, tag) { const p = h(tag || 'p'); p.innerHTML = txt ? esc(txt) : '<br>'; return p; }
  applyPageStyle() {
    let [w, hh] = PAGE[this.pageSize] || PAGE.A4;
    if (this.orientation === 'landscape') { const t = w; w = hh; hh = t; }
    const m = margPx(this.margins);
    this.pageW = w; this.pageH = hh; this.marg = m;
    this.sheet.style.width = w + 'px';
    this.body.style.padding = m[0] + 'px ' + m[1] + 'px ' + m[2] + 'px ' + m[3] + 'px';
    this.body.style.minHeight = hh + 'px';
    this.el.style.setProperty('--wd-page-w', w + 'px');
    this.el.style.setProperty('--wd-content-h', (hh - m[0] - m[2]) + 'px');
    this.updatePageBg();
  }
  // Visual page breaks via a background gradient behind the single flowing body
  updatePageBg() {
    const hh = this.pageH, m = this.marg, gap = 24;
    const contentH = hh - m[0] - m[2];
    // draw page separators using repeating background on the sheet
    this.sheet.style.background = 'transparent';
  }
  // v0.0.8: real pagination. The body still flows as one editable column, but every block that would cross a page
  // boundary (bottom margin + gap + top margin of the next page) is pushed to the top of the next page, and a manual
  // page break (Ctrl+Enter) fills the rest of its page. Pushes are applied through a scoped <style> (nth-child paths),
  // so the document HTML that is saved / exported never contains layout spacers.
  updatePages() {
    const GAP = 24;
    if (!this._pgStyle) { this._pgStyle = h('style.wd-pg-style'); this.el.id = this.el.id || ('wd' + Math.random().toString(36).slice(2, 9)); }
    if (!this._pgStyle.parentNode) this.el.appendChild(this._pgStyle);
    const flowing = this.viewMode === 'web' || this.viewMode === 'read';
    this.sheet.classList.toggle('web', this.viewMode === 'web');
    this.sheet.classList.toggle('wd-paged', !flowing);
    if (flowing) { this._pgStyle.textContent = ''; this.pageCount = 1; this.pageStarts = [0]; this.updateStatus(); return; }
    const body = this.body, m = this.marg, pageH = this.pageH, contentH = pageH - m[0] - m[2];
    const scope = '#' + this.el.id + ' .wd-body';
    this._pgStyle.textContent = '';          // measure the natural flow
    body.style.minHeight = '0px';
    const pageTop = n => n * (pageH + GAP) + m[0], pageBot = n => pageTop(n) + contentH;
    const relTop = el => { let y = 0, e = el; while (e && e !== body) { y += e.offsetTop; e = e.offsetParent; } return y; };
    const path = el => { const p = []; let e = el; while (e && e !== body) { const par = e.parentNode; p.unshift(Array.prototype.indexOf.call(par.children, e) + 1); e = par; } return scope + ' > ' + p.map(i => ':nth-child(' + i + ')').join(' > '); };
    const cs = el => window.getComputedStyle(el);
    // units that may move independently: top-level blocks, list items of long lists, table rows
    const units = [];
    const multiCol = (parseInt(body.style.columnCount) || 1) > 1;
    Array.prototype.forEach.call(body.children, el => {
      if (el.classList.contains('wd-pagebreak')) { units.push({ el, brk: true }); return; }
      if (cs(el).position === 'absolute' || el.offsetHeight === 0) return;
      const tall = el.offsetHeight > contentH * 0.25;
      if (tall && (el.tagName === 'UL' || el.tagName === 'OL')) { Array.prototype.forEach.call(el.children, li => units.push({ el: li, kind: 'm' })); return; }
      if (tall && el.tagName === 'TABLE' && el.rows && el.rows.length > 1) { Array.prototype.forEach.call(el.rows, tr => units.push({ el: tr, kind: 'tr' })); return; }
      units.push({ el, kind: 'm' });
    });
    units.forEach(u => { u.top = relTop(u.el); u.h = u.el.offsetHeight; if (u.kind === 'm') { const st = cs(u.el); u.mt = parseFloat(st.marginTop) || 0; const prev = u.el.previousElementSibling; u.prevMb = prev ? (parseFloat(cs(prev).marginBottom) || 0) : 0; } else if (u.kind === 'tr') u.pt = Array.prototype.map.call(u.el.cells, c => parseFloat(cs(c).paddingTop) || 0); });
    const rules = []; let shift = 0, n = 0; const starts = [0];
    const pageOf = y => { while (y >= pageTop(n + 1) - m[0] - 0.5) { n++; starts.push(null); } };
    for (const u of units) {
      let top = u.top + shift; pageOf(top);
      if (u.brk) {
        // manual page break: fill the rest of the page so the next block starts on the next page
        const d = Math.max(0, pageTop(n + 1) - top);
        rules.push(path(u.el) + '{height:' + d.toFixed(1) + 'px !important;margin:0 !important;padding:0 !important}');
        shift += d - u.h; n++; starts.push(null); continue;
      }
      if (multiCol) continue;
      let d = 0;
      if (top < pageTop(n) - 0.5) d = pageTop(n) - top;                         // landed inside the top margin
      else if (top + u.h > pageBot(n) + 0.5 && u.h <= contentH && top > pageTop(n) + 0.5) d = pageTop(n + 1) - top; // would cross the page end
      if (d > 0.5) {
        if (u.kind === 'tr') rules.push(Array.prototype.map.call(u.el.cells, (c, i) => path(c) + '{padding-top:' + (u.pt[i] + d).toFixed(1) + 'px !important}').join(''));
        else rules.push(path(u.el) + '{margin-top:' + (Math.max(u.mt, u.prevMb) + d).toFixed(1) + 'px !important}');
        shift += d; pageOf(top + d);
      }
    }
    const last = units.length ? units[units.length - 1] : null;
    const endY = last ? last.top + shift + (last.brk ? 0 : last.h) : m[0];
    pageOf(Math.max(m[0], endY - 1));
    this.pageCount = Math.max(1, n + 1);
    this._pgStyle.textContent = rules.join('\n');
    let frames = this.sheet.querySelector('.wd-frames');
    if (!frames) { frames = h('div.wd-frames'); this.sheet.insertBefore(frames, this.body); }
    const want = this.pageCount;
    while (frames.children.length > want) frames.removeChild(frames.lastChild);
    while (frames.children.length < want) frames.appendChild(h('div.wd-page-frame', h('span.wd-page-num')));
    Array.prototype.forEach.call(frames.children, (f, i) => { f.style.top = (i * (pageH + GAP)) + 'px'; f.style.height = pageH + 'px'; f.firstChild.textContent = (i + 1) + ' / ' + want; });
    const total = want * pageH + (want - 1) * GAP;
    this.sheet.style.minHeight = total + 'px';
    body.style.minHeight = total + 'px';
    this.updateStatus();
    this.updateRuler();
  }
  pageAt(y) { return Math.max(1, Math.min(this.pageCount || 1, Math.floor(y / (this.pageH + 24)) + 1)); }
  gotoPage(n) { n = Math.max(1, Math.min(this.pageCount || 1, n)); this.scroller.scrollTop = Math.round(((n - 1) * (this.pageH + 24)) * (this.zoom / 100) + (this.pagesEl.offsetTop || 0)); this.curPage = n; this.updateStatus(); }
  updateRuler() { if (this.rulerObj) { if (this._rulerRaf) cancelAnimationFrame(this._rulerRaf); this._rulerRaf = requestAnimationFrame(() => { this._rulerRaf = 0; this.rulerObj.draw(); }); } }
  // custom margins [top, right, bottom, left] in px (96 dpi); live = while dragging on the ruler (no history entry)
  setMargins(m, live) { this.margins = m.map(x => Math.round(x * 100) / 100); this.applyPageStyle(); if (!live) { this.updatePages(); this.markDirty(); this.refreshUi(); } this.updateRuler(); }
  updateStatus() {
    const t = this.getText(); const words = (t.match(/\S+/g) || []).length;
    this.sbWords.textContent = words + ' từ';
    this.sbPage.textContent = 'Trang ' + (this.curPage || 1) + ' / ' + (this.pageCount || 1);
    this.sbTrack.classList.toggle('hidden', !this.trackChanges);
  }

  // ---------------------------------------------------------------- selection & editing
  onSelection() { if (this.ribbon) this.ribbon.refresh(); this.updateCurPage(); this.saveCaret(); this.updateRuler(); }
  refreshUi() { if (this.ribbon) this.ribbon.refresh(); this.updateStatus(); }
  saveCaret() { const s = window.getSelection(); if (s.rangeCount && this.body.contains(s.anchorNode)) this._savedRange = s.getRangeAt(0).cloneRange(); }
  restoreCaret() { if (this._savedRange) { const s = window.getSelection(); s.removeAllRanges(); try { s.addRange(this._savedRange); } catch (e) {} } else { W.caretToStart(this); } }
  updateCurPage(fromScroll) {
    // v0.0.8: page under the caret (or under the middle of the window while scrolling), measured against the real page frames
    let y;
    if (fromScroll) { const r = this.scroller.getBoundingClientRect(); y = (r.top + r.height / 3 - this.sheet.getBoundingClientRect().top) / (this.zoom / 100); }
    else { const s = window.getSelection(); if (!s.rangeCount) return; let n = s.anchorNode; if (n && n.nodeType === 3) n = n.parentElement; if (!n || !this.body.contains(n)) return; const rc = s.getRangeAt(0).getClientRects(); const t = rc && rc.length ? rc[0].top : n.getBoundingClientRect().top; y = (t - this.sheet.getBoundingClientRect().top) / (this.zoom / 100); }
    this.curPage = this.pageAt(y); this.updateStatus();
  }
  queryState() { return W.queryState(this); }

  bind() {
    const b = this.body;
    b.addEventListener('input', () => { this.markDirty(); this.scheduleHistory(); this.debouncedPages(); this.refreshUi(); W.autoFormat(this); });
    this.debouncedPages = debounce(() => this.updatePages(), 250);
    b.addEventListener('keydown', e => W.onKey(this, e));
    b.addEventListener('mouseup', () => this.onSelection());
    b.addEventListener('keyup', e => { if (/Arrow|Home|End|Page/.test(keyName(e))) this.onSelection(); });
    b.addEventListener('contextmenu', e => { if (!settings().mouse.rightClickMenus) return; contextMenu(e, W.contextItems(this)); });
    longPress(b, e => contextMenu(e, W.contextItems(this)));
    b.addEventListener('paste', e => W.onPaste(this, e));
    b.addEventListener('copy', e => W.onCopy(this, e));
    b.addEventListener('cut', e => { W.onCopy(this, e); setTimeout(() => this.markDirty(), 0); });
    b.addEventListener('dragover', e => e.preventDefault());
    b.addEventListener('drop', e => W.onDrop(this, e));
    b.addEventListener('click', e => { const a = e.target.closest('a'); if (a && (e.ctrlKey || e.metaKey)) { e.preventDefault(); import('../core/host.js').then(m => m.default.openUrl(a.href)); } if (e.target.tagName === 'IMG') W.selectImage(this, e.target); });
    this.scroller.addEventListener('scroll', () => { this.updateCurPage(true); if (this.rulerObj) this.rulerObj.draw(); });
    this.scroller.addEventListener('dblclick', e => { if (this.viewMode === 'web' || (e.target !== this.body && !e.target.classList.contains('wd-sheet') && !e.target.classList.contains('wd-pages') && !e.target.classList.contains('wd-page-frame'))) return; import('../core/mouse.js').then(M => { if (M.clickAndType(this, e)) { this.body.focus(); this.updatePages(); } }); });
    document.addEventListener('selectionchange', () => { if (this.shell.cur === this && this.body.contains(document.activeElement === this.body ? this.body : null)) this.updateRuler(); });
  }
  acceptsDrop(e) { return e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length && /^image\//.test(e.dataTransfer.files[0].type); }

  // ---------------------------------------------------------------- history
  scheduleHistory() { clearTimeout(this._histTimer); this._histTimer = setTimeout(() => this.pushHistory(), 500); }
  pushHistory(initial) {
    const html = this.body.innerHTML;
    if (!initial && this.undoStack.length && this.undoStack[this.undoStack.length - 1] === html) return;
    this.undoStack.push(html); if (this.undoStack.length > 120) this.undoStack.shift();
    if (!initial) this.redoStack = [];
  }
  exec(label, fn) { W.withUndo(this, label, fn); }
  undo() { clearTimeout(this._histTimer); if (this.undoStack.length < 2) return; this.saveCaret(); const cur = this.undoStack.pop(); this.redoStack.push(cur); this.body.innerHTML = this.undoStack[this.undoStack.length - 1]; this.afterRestore(); }
  redo() { if (!this.redoStack.length) return; const html = this.redoStack.pop(); this.undoStack.push(html); this.body.innerHTML = html; this.afterRestore(); }
  afterRestore() { this.updatePages(); this.markDirty(); this.refreshUi(); W.caretToStart(this); this.body.focus(); }
  canUndo() { return this.undoStack.length > 1; }
  canRedo() { return this.redoStack.length > 0; }

  contextItems() { return W.contextItems(this); }
  commands() { return W.extraCommands(this); }
  renderPrint(area) { return W.renderPrint(this, area); }
  printPreview(box) { return W.printPreview(this, box); }
  printOptions() { return W.printOptions(this); }
  showComments() { W.commentsPane(this); }
  setReadOnly(v) { this.readOnly = v; this.body.contentEditable = v ? 'false' : 'true'; }
  onBack() { if ($('.wd-pane', this.el)) { $('.wd-pane', this.el).remove(); return true; } return false; }
}
registerApp('word', WordDoc);
// margins may be a preset name or [top, right, bottom, left] px
export function margPx(v) { if (Array.isArray(v) && v.length === 4) return v.map(Number); return (MARG[v] || MARG.normal).slice(); }
export { PAGE, MARG };
