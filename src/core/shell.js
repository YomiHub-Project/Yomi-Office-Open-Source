// Application shell: window frame (title bar + QAT + search), backstage (File), status bar, document management
import { h, ic, esc, $, $$, extOf, baseName, chord, keyName, debounce, relTime, fmtSize, uid, emitter, APP_NAME, VERSION } from './util.js';
import { Ribbon, showMenu, toast, dialog, confirmDlg, promptDlg, formDlg, zoomControl, contextMenu, busy, closeMenus } from './ui.js';
import { settings, saveSettings, recentList, recentTouch, recentRemove, recentPin, snapshotGet } from './store.js';
import Host from './host.js';
import { initWinCaps } from './wincaps.js';
import { bsPage as ubPage, bsAction as ubAction, optionsDialog, SPEC as USPEC, friendlyDate } from './bstage.js';

export const KINDS = {
  word: { label: 'Word', vi: 'Văn bản', color: '#185abd', icon: 'file-text', exts: ['docx', 'dotx', 'docm', 'doc', 'dot', 'xml', 'mht', 'mhtml', 'txt', 'md', 'html', 'htm', 'rtf', 'odt', 'ott'], newExt: 'docx' },
  excel: { label: 'Excel', vi: 'Bảng tính', color: '#107c41', icon: 'sheet', exts: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xls', 'xlt', 'csv', 'tsv', 'ods', 'ots'], newExt: 'xlsx' },
  ppt: { label: 'PowerPoint', vi: 'Trình chiếu', color: '#c43e1c', icon: 'presentation', exts: ['pptx', 'pptm', 'potx', 'ppsx', 'ppt', 'pps', 'pot', 'odp', 'otp'], newExt: 'pptx' },
  pdf: { label: 'PDF', vi: 'PDF', color: '#d83b01', icon: 'file-type', exts: ['pdf', 'png', 'jpg', 'jpeg', 'bmp', 'gif', 'webp'], newExt: 'pdf' },
  mathtype: { label: 'MathType', vi: 'Toán học', color: '#68217a', icon: 'function-square', exts: ['tex', 'latex', 'mml', 'eqn'], newExt: 'tex' }
};
export function kindOfExt(ext) { for (const k in KINDS) if (KINDS[k].exts.indexOf(ext) >= 0) return k; return null; }

export const bus = emitter();
const registry = {}; // kind -> class
export function registerApp(kind, cls) { registry[kind] = cls; }

// ---------------------------------------------------------------- Base document
export class Doc {
  constructor(shell, opts) {
    this.shell = shell; this.opts = opts || {};
    this.id = opts.recentId || uid('doc');
    this.name = opts.name || 'Không tên';
    this.path = opts.path || null;
    this.dirty = false; this.zoom = 100; this.readOnly = false;
    this.el = h('div.doc-view');
    this.statusLeft = h('div.sb-left'); this.statusRight = h('div.sb-right');
  }
  get kind() { return this.constructor.kind; }
  get meta() { return KINDS[this.kind]; }
  markDirty() { if (!this.dirty) { this.dirty = true; this.shell.updateTitle(); } this.shell.scheduleAutosave(this); }
  setZoom(z) { this.zoom = z; }
  focus() {}
  afterCommand() { this.refreshUi(); }
  refreshUi() { if (this.ribbon) this.ribbon.refresh(); this.updateStatus && this.updateStatus(); }
  getText() { return ''; }
  snapshot() { return null; }
  commands() { return []; }
  // overridables: undo(), redo(), canUndo(), canRedo(), save(ext) -> bytes, exportFormats(), onShow(), onHide(), print(), find()
  saveFormats() { return [[this.meta.newExt, this.meta.label]]; }
}

// ---------------------------------------------------------------- Shell
export class Shell {
  constructor(root) {
    window.__oosShell = this;
    try { initWinCaps(); } catch (e) {}
    this.root = root; this.docs = []; this.cur = null; this.view = null;
    this.autosaveTimers = {};
    document.addEventListener('keydown', e => this.onKey(e));
    window.addEventListener('resize', debounce(() => this.cur && this.cur.onLayout && this.cur.onLayout(), 120));
    // Ctrl + wheel = zoom (mouse feature)
    document.addEventListener('wheel', e => {
      const xlRoll = this.cur && this.cur.kind === 'excel' && settings().mouse.wheelZoomNoCtrl && !e.shiftKey && e.target.closest && e.target.closest('.xl-grid,.xl-scroll');
      if (!xlRoll && (!(e.ctrlKey || e.metaKey) || !this.cur || !settings().mouse.wheelZoomCtrl)) return;
      if (e.target.closest && e.target.closest('.menu,.dlg')) return;
      e.preventDefault();
      const z = Math.max(10, Math.min(400, Math.round(this.cur.zoom + (e.deltaY < 0 ? 10 : -10))));
      this.setZoom(z);
    }, { passive: false });
    // drag & drop files onto window opens them
    document.addEventListener('dragover', e => { if (e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types || [], 'Files') >= 0) { e.preventDefault(); document.body.classList.add('drop-hint'); } });
    document.addEventListener('dragleave', e => { if (!e.relatedTarget) document.body.classList.remove('drop-hint'); });
    document.addEventListener('drop', async e => {
      document.body.classList.remove('drop-hint');
      if (!e.dataTransfer || !e.dataTransfer.files || !e.dataTransfer.files.length) return;
      if (this.cur && this.cur.acceptsDrop && this.cur.acceptsDrop(e)) return; // e.g. image into document
      e.preventDefault();
      for (const f of Array.prototype.slice.call(e.dataTransfer.files)) {
        const bytes = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(new Uint8Array(fr.result)); fr.readAsArrayBuffer(f); });
        this.openFile({ name: f.name, bytes });
      }
    });
    window.addEventListener('beforeunload', e => { if (this.docs.some(d => d.dirty)) { this.docs.forEach(d => d.dirty && this.autosaveNow(d)); } });
    window.OOS_onOpenFile = json => { try { const o = JSON.parse(json); import('./util.js').then(u => this.openFile({ name: o.name, bytes: u.bytesFromB64(o.b64), path: o.uri })); } catch (e) {} };
    window.OOS_onBack = () => this.onBack();
  }

  // ---------------- views
  showStart(kind) {
    if (this.cur && this.cur.onHide) this.cur.onHide();
    this.cur = null; this.view = 'start';
    this.root.innerHTML = '';
    this.root.appendChild(this.startView(kind));
    Host.setTitle(APP_NAME);
  }
  showDoc(doc) {
    if (this.cur && this.cur !== doc && this.cur.onHide) this.cur.onHide();
    this.cur = doc; this.view = 'doc';
    this.root.innerHTML = '';
    const frame = h('div.frame.k-' + doc.kind, { style: { '--accent': doc.meta.color } });
    frame.style.setProperty('--accent', doc.meta.color);
    frame.appendChild(this.titleBar(doc));
    if (!doc.ribbon) doc.ribbon = new Ribbon(doc, Object.assign({ onFile: () => this.backstage(doc), fileLabel: 'Tệp', collapseOnReclick: true, tail: t => this.ribbonTail(doc, t) }, doc.ribbonDef()));
    if (settings().ribbonCollapsed && !doc._rbInit) doc.ribbon.toggleCollapse(true);
    doc._rbInit = true;
    frame.appendChild(doc.ribbon.el);
    if (doc.formulaBar) frame.appendChild(doc.formulaBar);
    const content = h('div.doc-content', doc.el);
    frame.appendChild(content);
    frame.appendChild(this.statusBar(doc));
    this.root.appendChild(frame);
    this.updateTitle();
    if (doc.onShow) doc.onShow();
    doc.refreshUi();
    setTimeout(() => doc.focus(), 30);
  }
  titleBar(doc) {
    const s = settings();
    // v0.1.7: same switch as the native title bar - Tự động lưu = save the original file after every change (autosaveLive)
    const auto = h('label.qat-auto', { title: 'Tự động lưu: lưu ngay vào tệp gốc sau mỗi thay đổi' }, h('span', 'Tự động lưu'), h('span.switch' + (s.autosaveLive ? '.on' : ''), h('span.knob')), h('span.tb-state', s.autosaveLive ? 'Bật' : 'Tắt'));
    auto.addEventListener('click', () => { const v = !settings().autosaveLive; saveSettings({ autosaveLive: v }); auto.querySelector('.switch').classList.toggle('on', v); auto.querySelector('.tb-state').textContent = v ? 'Bật' : 'Tắt'; toast(v ? 'Đã bật Tự động lưu' : 'Đã tắt Tự động lưu'); if (v && doc.dirty) this.scheduleAutosave(doc); });
    const qb = (icon, title, fn, sc) => { const b = h('button.qat-b', { title: title + (sc ? ' (' + sc + ')' : '') }, ic(icon, 16)); b.addEventListener('mousedown', e => e.preventDefault()); b.addEventListener('click', fn); return b; };
    const undoB = qb('n:undo', 'Hoàn tác (Undo)', () => { doc.undo && doc.undo(); doc.refreshUi(); }, 'Ctrl+Z');
    const redoB = qb('n:redo', 'Làm lại (Redo)', () => { doc.redo && doc.redo(); doc.refreshUi(); }, 'Ctrl+Y');
    const undoDrop = h('button.qat-b.qat-dd', { title: 'Lịch sử hoàn tác' }, ic('chevron-down', 11));
    undoDrop.addEventListener('click', () => {
      const hist = doc.undoHistory ? doc.undoHistory() : [];
      showMenu(hist.length ? hist.slice(0, 20).map((x, i) => ({ label: (i + 1) + '. ' + x, run: () => { for (let k = 0; k <= i; k++) doc.undo(); doc.refreshUi(); } })) : [{ label: 'Không có thao tác để hoàn tác', disabled: true }], { anchor: undoDrop });
    });
    this._undoB = undoB; this._redoB = redoB;
    const title = h('div.tb-title', h('span.tb-name'), h('span.tb-app.tb-saved', ''), ic('chevron-down', 12));   // v0.1.7: "<tệp> • Đã lưu vào PC" like the native title bar
    title.addEventListener('click', () => this.renameDoc(doc));
    title.title = 'Bấm để đổi tên tài liệu';
    const search = h('div.tb-search', ic('search', 14), h('input', { placeholder: 'Tìm kiếm (Alt + Q)', id: 'tellme' }));
    const inp = search.querySelector('input');
    inp.addEventListener('input', () => this.tellMe(doc, inp));
    inp.addEventListener('focus', () => this.tellMe(doc, inp));
    inp.addEventListener('keydown', e => { if (keyName(e) === 'Escape') { inp.value = ''; inp.blur(); closeMenus(); doc.focus(); } if (keyName(e) === 'Enter') { const f = $('.menu .menu-item'); if (f) f.click(); } });
    const user = h('button.tb-user', { title: 'Tài khoản: ' + s.userName }, (s.initials || 'ND').slice(0, 2));
    user.addEventListener('click', () => this.backstage(doc, 'account'));
    const home = h('button.tb-home', { title: 'Màn hình chính Yomi-Office Open Source' }, ic('n:' + (doc.kind === 'excel' ? 'excel' : doc.kind === 'ppt' ? 'powerpoint' : 'word'), 20));
    home.addEventListener('click', () => this.goHome());
    const qatMore = h('button.qat-b.qat-dd', { title: 'Tùy chỉnh thanh công cụ truy nhập nhanh' }, ic('chevron-down', 12));
    qatMore.addEventListener('click', () => showMenu([
      { header: 'Thanh công cụ truy nhập nhanh' },
      { label: 'Mới (Ctrl+N)', icon: 'file-plus', run: () => this.newDoc(doc.kind) },
      { label: 'Mở (Ctrl+O)', icon: 'folder-open', run: () => this.openDialog() },
      { label: 'Lưu (Ctrl+S)', icon: 'save', run: () => this.save(doc) },
      { label: 'In (Ctrl+P)', icon: 'printer', run: () => this.backstage(doc, 'print') },
      { label: 'Trình chiếu/Xem toàn màn hình', icon: 'maximize', run: () => this.fullscreen() },
      { sep: true },
      { label: (doc.ribbon && doc.ribbon.collapsed ? 'Hiện' : 'Thu gọn') + ' Ribbon (Ctrl+F1)', icon: 'panel-top', run: () => doc.ribbon.toggleCollapse() }
    ], { anchor: qatMore }));
    return h('div.titlebar',
      home,
      h('div.qat', auto, qb('n:save', 'Lưu (Save)', () => this.save(doc), 'Ctrl+S'), undoB, undoDrop, redoB, qatMore),
      title, search,
      h('div.tb-right',
        h('div.dex-titlebar-tools',
          h('button.qat-b', { title: 'Chia đôi màn hình trái (50%) - DeX Snap Left', onclick: () => DEX.tileWindow('left') }, ic('columns', 15)),
          h('button.qat-b', { title: 'Chia đôi màn hình phải (50%) - DeX Snap Right', onclick: () => DEX.tileWindow('right') }, ic('columns', 15)),
          h('button.qat-b', { title: 'Cửa sổ nổi tự do (DeX Freeform Window)', onclick: () => DEX.tileWindow('window') }, ic('app-window', 15)),
          h('button.qat-b', { title: 'Không gian DeX Multi-Window & TV 4K Workspace', onclick: () => DEX.openDexDesktopEnvironment({ mode: 'multiwin' }) }, ic('monitor', 15))
        ),
        user));
  }
  ribbonTail(doc, tail) {
    // v0.1.7: same three buttons as the native ribbon: [Đang chỉnh sửa ▾] (Word) [Bình luận] [Chia sẻ]
    if (doc.kind === 'word') {
      const em = h('button.rb-editmode', ic('n:editor', 16), h('span', doc._reviewMode ? 'Đang xem lại' : 'Đang chỉnh sửa'), ic('chevron-down', 11));
      em.addEventListener('click', () => showMenu([
        { label: 'Chỉnh sửa', icon: 'n:editor', run: () => { doc._reviewMode = false; em.querySelector('span:not(.ic)').textContent = 'Đang chỉnh sửa'; if (doc.trackChanges && doc.trackChanges(false) === undefined) {} } },
        { label: 'Xem lại (theo dõi thay đổi)', icon: 'n:track', run: () => { doc._reviewMode = true; em.querySelector('span:not(.ic)').textContent = 'Đang xem lại'; if (doc.trackChanges) doc.trackChanges(true); } }
      ], { anchor: em }));
      tail.appendChild(em);
    }
    const b = h('button.rb-share', ic('n:share', 16), h('span', 'Chia sẻ'));
    b.addEventListener('click', () => this.backstage(doc, 'share'));
    const c = h('button.rb-comments', ic('n:comment', 16), h('span', 'Bình luận'));
    c.addEventListener('click', () => doc.showComments ? doc.showComments() : toast('Chú thích có trong tab Xem lại (Review)'));
    tail.appendChild(c); tail.appendChild(b);
  }
  statusBar(doc) {
    const z = zoomControl(() => doc.zoom, v => this.setZoom(v));
    this._zoomCtl = z;
    const views = doc.viewButtons ? doc.viewButtons() : null;
    const dexPill = DEX.createStatusBarItem(() => this.refreshUi());
    const sb = h('div.statusbar', doc.statusLeft, h('div.sb-spacer'), doc.statusRight, dexPill, views, z.el);
    return sb;
  }
  setZoom(z) { if (!this.cur) return; this.cur.setZoom(z); if (this._zoomCtl) this._zoomCtl.update(); }
  updateTitle() {
    const d = this.cur; if (!d) return;
    const n = $('.tb-name'); if (n) n.textContent = d.name + (d.fileExt ? '.' + d.fileExt : '');
    const sv = $('.tb-saved'); if (sv) sv.textContent = ' • ' + (d.dirty ? 'Chưa lưu' : d.path ? 'Đã lưu vào PC' : 'Chưa lưu');
    Host.setTitle(d.name + ' - ' + d.meta.label + ' | ' + APP_NAME);
    if (this._undoB) { this._undoB.classList.toggle('disabled', !(d.canUndo && d.canUndo())); this._redoB.classList.toggle('disabled', !(d.canRedo && d.canRedo())); }
  }
  async renameDoc(doc) {
    const v = await promptDlg('Tên tài liệu:', doc.name, 'Đổi tên');
    if (v) { doc.name = v.replace(/[\\/:*?"<>|]/g, '_'); doc.markDirty(); this.updateTitle(); }
  }
  tellMe(doc, inp) {
    const q = inp.value.trim().toLowerCase();
    const cmds = this.allCommands(doc);
    const hits = (q ? cmds.filter(c => (c.label + ' ' + (c.kw || '')).toLowerCase().indexOf(q) >= 0) : cmds.slice(0, 8)).slice(0, 12);
    showMenu(hits.length ? [{ header: q ? 'Kết quả cho "' + inp.value + '"' : 'Gợi ý' }].concat(hits.map(c => ({ label: c.label, icon: c.icon, run: () => { c.run(); doc.refreshUi(); } }))) : [{ label: 'Không tìm thấy lệnh', disabled: true }], { anchor: inp.parentNode, cls: 'keep-menu', minWidth: 320 });
    inp.focus();
  }
  allCommands(doc) {
    const out = [];
    const walk = (items, tab) => (items || []).forEach(it => {
      if (!it) return;
      if (it.items && (it.type === 'col' || it.type === 'row')) return walk(it.items, tab);
      if ((it.run || it.menu) && (it.label || it.title)) out.push({ label: (it.label || it.title) + ' — ' + tab, icon: it.icon, kw: it.title + ' ' + (it.kw || ''), run: () => it.run ? it.run(doc) : showMenu(it.menu(doc), { x: 200, y: 140 }) });
    });
    const def = doc.ribbon ? doc.ribbon.def : doc.ribbonDef();
    def.tabs.forEach(t => (typeof t.groups === 'function' ? t.groups() : t.groups).forEach(g => g && walk(g.items, t.label)));
    return out.concat(doc.commands ? doc.commands() : []);
  }

  // ---------------- documents
  newDoc(kind, template) {
    const Cls = registry[kind]; if (!Cls) return toast('Chưa hỗ trợ: ' + kind, 'err');
    const n = this.docs.filter(d => d.kind === kind).length + 1;
    const names = { word: 'Tài liệu', excel: 'Sổ làm việc', ppt: 'Bản trình bày', pdf: 'PDF', mathtype: 'Công thức toán' };
    const doc = new Cls(this, { name: (template && template.name) || (names[kind] + n), template });
    this.docs.push(doc);
    Promise.resolve(doc.init ? doc.init() : null).then(() => { this.showDoc(doc); }).catch(e => { console.error(e); toast('Lỗi tạo tài liệu: ' + e.message, 'err'); });
    return doc;
  }
  async openDialog(kind) {
    const exts = kind ? KINDS[kind].exts : [].concat.apply([], Object.keys(KINDS).map(k => KINDS[k].exts)).concat(['epub', 'fb2']);
    const files = await Host.openFiles(exts.map(e => '.' + e).join(','), true);
    for (const f of files) await this.openFile(f);
  }
  // v0.0.6: "one file = one window" (Word/Excel/PowerPoint/PDF). The first document opens in this window when it is empty
  // (Start screen); every further file opens in its own window. Setting: Tùy chọn > Chung > Mở mỗi tệp trong cửa sổ mới.
  wantsNewWindow(f) {
    if (settings().newWindowPerFile === false || !Host.canNewWindow() || f.here) return false;
    return this.docs.length > 0;
  }
  async openFile(f, recentId) {
    const ext = extOf(f.name); let kind = kindOfExt(ext);
    if ((ext === 'xml' || ext === 'mht' || ext === 'mhtml') && f.bytes) { const head = new TextDecoder('utf-8').decode(f.bytes.subarray(0, 4000)); if (/urn:schemas-microsoft-com:office:spreadsheet|progid="Excel/.test(head)) kind = 'excel'; else if (/PowerPoint\.Show|presentationml/.test(head)) kind = 'ppt'; }
    if (/^(epub|fb2|fbz)$/.test(ext) && f.bytes) { const R = await import('../reader/reader.js'); return R.openBook(this, f); }
    if (!kind) return toast('Định dạng .' + ext + ' chưa được hỗ trợ', 'err');
    // v0.1.6: Word / Excel / PowerPoint open in the native engine (LibreOffice core, full MS Office fidelity, no WebView)
    if ((kind === 'word' || kind === 'excel' || kind === 'ppt') && !f.forceWeb && settings().officeEngine !== 'web' && await Host.lokAvailable()) {
      const b0 = busy('Đang mở ' + f.name + ' bằng engine Native…');
      try { if (await Host.lokOpen(f)) { recentTouch({ id: 'lok_' + (f.path || f.name), name: f.name, kind, path: f.path || null, size: f.bytes ? f.bytes.length : 0 }, { ext, b64: null }); return null; } }
      finally { b0.done(); }
      toast('Engine Native chưa mở được tệp, chuyển sang giao diện Web', 'warn');
    }
    const existing = this.docs.find(d => (f.path && d.path === f.path));
    if (existing) { this.showDoc(existing); return existing; }
    if (this.wantsNewWindow(f)) { const ok = await Host.openInNewWindow(f); if (ok) { toast('Đã mở "' + f.name + '" trong cửa sổ mới'); return null; } }
    const Cls = registry[kind];
    const b = busy('Đang mở ' + f.name + '…');
    try {
      const doc = new Cls(this, { name: baseName(f.name), path: f.path || null, ext, recentId });
      await doc.load(f.bytes, ext, f.name);
      doc.fileExt = ext;
      this.docs.push(doc);
      this.showDoc(doc);
      doc._fileSize = f.bytes ? f.bytes.length : 0;
      recentTouch({ id: doc.id, name: f.name, kind, path: f.path || null, size: f.bytes.length }, { ext, b64: null });
      this.autosaveNow(doc, true);
      return doc;
    } catch (e) { console.error(e); toast('Không mở được ' + f.name + ': ' + (e.message || e), 'err', 5000); }
    finally { b.done(); }
  }
  async openRecent(rec) {
    // 1) real path (Windows path / Android content URI) 2) autosave snapshot
    if (rec.path) { try { const f = await Host.readPath(rec.path); if (f) return this.openFile(f, rec.id); } catch (e) {} }
    const snap = await snapshotGet(rec.id);
    if (snap && snap.native) {
      const Cls = registry[rec.kind];
      const doc = new Cls(this, { name: baseName(rec.name), recentId: rec.id, path: rec.path });
      await doc.restore(snap.native);
      this.docs.push(doc); this.showDoc(doc);
      return doc;
    }
    toast('Không tìm thấy tệp "' + rec.name + '" (đã di chuyển hoặc xóa).', 'err');
  }
  async save(doc, saveAs, ext, toPath) {
    if (!doc) return;
    if (doc.kind === 'pdf' && doc.saveBytes == null) return toast('Dùng các công cụ PDF để xuất tệp.');
    ext = ext || (saveAs ? null : (doc.fileExt && doc.canSaveExt && doc.canSaveExt(doc.fileExt) ? doc.fileExt : doc.meta.newExt));
    if (!ext) return this.backstage(doc, 'saveas');
    if (doc.cloud && !saveAs && !doc.path) { const C = await import('./cloud.js'); if (await C.syncUp(doc)) { doc.dirty = false; this.updateTitle(); toast('Đã lưu lên ' + C.PROVIDERS[doc.cloud.provider].name, 'ok'); } return; }
    const b = busy('Đang lưu…');
    try {
      const bytes = await doc.save(ext);
      const fmt = (doc.saveFormats ? doc.saveFormats() : []).find(f => f[0] === ext); const fileExt = (ext === 'video' && doc._lastVideoExt) || (fmt && fmt[3]) || ext;
      const name = doc.name + '.' + fileExt;
      const r = toPath ? await Host.saveFile(name, bytes, { path: toPath, saveAs: false }) : await Host.saveFile(name, bytes, { path: !saveAs && doc.fileExt === ext ? doc.path : null, saveAs: !!saveAs || !doc.path || doc.fileExt !== ext });
      if (!r) return;
      if (r.path) doc.path = r.path;
      if (r.name) doc.name = baseName(r.name);
      doc.fileExt = ext; doc.dirty = false; this.updateTitle();
      if (doc.cloud && !saveAs) import('./cloud.js').then(C => C.syncUp(doc)); // v0.0.8: keep the cloud copy in sync
      await recentTouch({ id: doc.id, name: doc.name + '.' + fileExt, kind: doc.kind, path: doc.path, size: bytes.length });
      this.autosaveNow(doc, true);
      toast('Đã lưu ' + doc.name + '.' + fileExt + (Host.kind === 'web' ? ' (thư mục Tải xuống)' : ''), 'ok');
    } catch (e) { console.error(e); toast('Lỗi khi lưu: ' + (e.message || e), 'err', 5000); }
    finally { b.done(); }
  }
  scheduleAutosave(doc) {
    if (settings().autosaveLive && doc.path && doc.fileExt && Host.kind !== 'web' && (!doc.canSaveExt || doc.canSaveExt(doc.fileExt))) {   // Tự động lưu (live) like the native editor
      clearTimeout(this._liveTimers && this._liveTimers[doc.id]); this._liveTimers = this._liveTimers || {};
      this._liveTimers[doc.id] = setTimeout(() => { if (doc.dirty) this.save(doc); }, 2500);
    }
    if (!settings().autosave) return;
    clearTimeout(this.autosaveTimers[doc.id]);
    this.autosaveTimers[doc.id] = setTimeout(() => this.autosaveNow(doc), (settings().autosaveSec || 30) * 1000);
  }
  async autosaveNow(doc, quiet) {
    try {
      const native = doc.snapshot(); if (!native) return;
      await recentTouch({ id: doc.id, name: doc.name + '.' + (doc.fileExt || doc.meta.newExt), kind: doc.kind, path: doc.path }, { native });
      doc.lastAutosave = Date.now();
      if (!quiet && this.cur === doc) { const s = $('.sb-autosave'); if (s) s.textContent = 'Đã tự động lưu ' + relTime(Date.now()); }
    } catch (e) { console.warn('autosave', e); }
  }
  async closeDoc(doc) {
    if (doc.dirty) {
      const r = await dialog({ title: APP_NAME, body: h('div.dlg-msg', 'Bạn có muốn lưu thay đổi cho "' + doc.name + '" không?'), buttons: [{ label: 'Lưu', primary: true, value: 'save' }, { label: 'Không lưu', value: 'no' }, { label: 'Hủy', value: null }] });
      if (!r) return false;
      if (r === 'save') { await this.save(doc); if (doc.dirty) return false; }
    }
    this.autosaveNow(doc, true);
    this.docs = this.docs.filter(d => d !== doc);
    if (doc.destroy) doc.destroy();
    if (this.cur === doc) { if (this.docs.length) this.showDoc(this.docs[this.docs.length - 1]); else this.showStart(doc.kind); }
    return true;
  }
  goHome() { this.showStart(this.cur ? this.cur.kind : null); }
  onBack() {
    // Android back button
    if (document.querySelector('.dlg-back')) { const x = $$('.dlg-back .dlg-x').pop(); if (x) x.click(); return true; }
    if (document.querySelector('.menu')) { closeMenus(); return true; }
    if (document.querySelector('.backstage')) { $('.backstage').remove(); return true; }
    if (this.cur && this.cur.onBack && this.cur.onBack()) return true;
    if (this.view === 'doc') { this.goHome(); return true; }
    return false;
  }
  fullscreen() {
    const el = document.documentElement;
    const req = el.requestFullscreen || el.webkitRequestFullscreen;
    if (document.fullscreenElement || document.webkitFullscreenElement) (document.exitFullscreen || document.webkitExitFullscreen).call(document);
    else if (req) req.call(el);
  }
  onKey(e) {
    const c = chord(e);
    const doc = this.cur;
    if (document.querySelector('.dlg-back')) return;
    if (c === 'ctrl+n') { e.preventDefault(); this.newDoc(doc ? doc.kind : 'word'); }
    else if (c === 'ctrl+o') { e.preventDefault(); this.openDialog(doc && doc.kind); }
    else if (c === 'ctrl+s' && doc) { e.preventDefault(); this.save(doc); }
    else if ((c === 'f12' || c === 'ctrl+shift+s') && doc) { e.preventDefault(); this.backstage(doc, 'saveas'); }
    else if (c === 'ctrl+p' && doc) { e.preventDefault(); this.backstage(doc, 'print'); }
    else if (c === 'ctrl+f1' && doc) { e.preventDefault(); doc.ribbon.toggleCollapse(); }
    else if ((c === 'alt+q') && doc) { e.preventDefault(); const t = $('#tellme'); if (t) t.focus(); }
    else if (c === 'ctrl+w' && doc) { e.preventDefault(); this.closeDoc(doc); }
    else if (c === 'ctrl+f4' && doc) { e.preventDefault(); this.closeDoc(doc); }
    else if (c === 'f11') { e.preventDefault(); this.fullscreen(); }
  }

  // ---------------- backstage (File menu)
  backstage(doc, page) {
    closeMenus();
    const old = $('.backstage'); if (old) old.parentNode.removeChild(old);
    const bs = h('div.backstage.k-' + doc.kind);
    bs.style.setProperty('--accent', doc.meta.color);
    const main = h('div.bs-main');
    const close = () => { if (bs.parentNode) bs.parentNode.removeChild(bs); doc.focus(); };
    const nav = h('div.bs-nav.ub-nav');   // v0.1.7: same navigation as the native editor (ui_settings.json)
    const pages = USPEC.nav;
    nav.appendChild(h('button.bs-back', { title: 'Quay lại (Esc)', onclick: close }, ic('n:back', 32)));
    const btns = {};
    let bottom = false;
    pages.forEach((p, i) => {
      if (!p) { if (i > 4) { bottom = true; nav.appendChild(h('div.bs-fill')); } else nav.appendChild(h('div.bs-sep')); return; }
      const b = h('button.bs-item' + (bottom ? '.bottom' : ''), ic('n:' + p[2], 16), h('span', p[1]));
      b.addEventListener('click', () => go(p[0]));
      btns[p[0]] = b; nav.appendChild(b);
    });
    let cur = 'info';
    const go = async (id) => {
      Object.keys(btns).forEach(k => btns[k].classList.toggle('on', k === id));
      main.innerHTML = '';
      if (id === 'save') { close(); return this.save(doc); }
      if (id === 'close') { close(); return this.closeDoc(doc); }
      if (id === 'options') { Object.keys(btns).forEach(k => btns[k].classList.toggle('on', k === cur)); go(cur); return optionsDialog(doc, this); }
      cur = id;
      main.appendChild(await ubPage(this, doc, id, close));
    };
    bs.appendChild(nav); bs.appendChild(main);
    bs.addEventListener('keydown', e => { if (keyName(e) === 'Escape') close(); });
    document.body.appendChild(bs);
    bs.tabIndex = -1; bs.focus();
    go(page === 'options' ? 'info' : (page || 'home'));
    if (page === 'options') optionsDialog(doc, this);
  }
  bsAction(doc, id) { return ubAction(this, doc, id); }
  async bsPageLegacy(doc, id, close) {
    const P = h('div.bs-page');
    const title = t => P.appendChild(h('h1.bs-h1', t));
    if (id === 'home' || id === 'new') {
      title(id === 'home' ? greeting() : 'Mới');
      P.appendChild(this.templateStrip(doc.kind, t => { close(); this.newDoc(doc.kind, t); }, id === 'new'));
      if (id === 'home') P.appendChild(await this.recentBlock(doc.kind, close));
    } else if (id === 'open') {
      title('Mở');
      P.appendChild(h('div.bs-open-src',
        h('button.bs-tile', { onclick: () => { close(); this.openDialog(doc.kind); } }, ic('hard-drive', 28), h('div', h('b', 'Duyệt (Browse)'), h('div.muted', Host.kind === 'android' ? 'Bộ nhớ máy, thẻ nhớ, Google Drive, OneDrive…' : 'Mở tệp từ máy tính này'))),
        h('button.bs-tile', { onclick: () => { close(); this.openDialog(); } }, ic('files', 28), h('div', h('b', 'Mọi loại tệp'), h('div.muted', 'docx, xlsx, pptx, pdf, csv, txt, md, html, ảnh…'))),
        h('button.bs-tile', { onclick: () => { close(); import('./cloud.js').then(C => C.cloudOpenDialog(this)); } }, ic('cloud', 28), h('div', h('b', 'Đám mây'), h('div.muted', 'Google Drive · OneDrive · Dropbox · MediaFire')))));
      P.appendChild(await this.recentBlock(null, close));
    } else if (id === 'info') {
      title('Thông tin');
      const st = doc.stats ? doc.stats() : {};
      P.appendChild(h('div.bs-info',
        h('div.bs-info-l',
          h('div.bs-card', ic('lock', 24), h('div', h('b', 'Bảo vệ tài liệu'), h('div.muted', 'Đánh dấu là bản cuối cùng (chỉ đọc) để tránh sửa đổi.'), h('button.btn', { onclick: () => { doc.readOnly = !doc.readOnly; doc.setReadOnly && doc.setReadOnly(doc.readOnly); toast(doc.readOnly ? 'Đã khóa chỉ đọc' : 'Đã mở khóa chỉnh sửa'); close(); } }, doc.readOnly ? 'Bỏ chỉ đọc' : 'Đánh dấu bản cuối cùng'))),
          h('div.bs-card', ic('search-check', 24), h('div', h('b', 'Kiểm tra tài liệu (Apache Tika)'), h('div.muted', 'Thống kê nội dung, siêu dữ liệu và xóa thông tin cá nhân.'), h('button.btn', { onclick: () => import('./suites.js').then(m => m.openTikaInspector(doc)) }, 'Kiểm tra'))),
          h('div.bs-card', ic('history', 24), h('div', h('b', 'Trung tâm Sao lưu (WPS Backup Center)'), h('div.muted', 'Lịch sử phiên bản tự động, bảo vệ an toàn ngoại tuyến.'), h('button.btn', { onclick: () => import('./suites.js').then(m => m.openBackupCenter(doc)) }, 'Mở sao lưu')))),
        h('div.bs-props', h('h3', 'Thuộc tính'),
          prop('Tên', doc.name), prop('Loại', doc.meta.label + (doc.fileExt ? ' (.' + doc.fileExt + ')' : '')), prop('Vị trí', doc.path || '(chưa lưu)'),
          Object.keys(st).map(k => prop(k, st[k])),
          prop('Tác giả', settings().userName), prop('Ứng dụng', APP_NAME + ' ' + VERSION))));
    } else if (id === 'saveas' || id === 'export') {
      title(id === 'saveas' ? 'Lưu Như' : 'Xuất');
      const fm = doc.saveFormats();
      const list = h('div.bs-formats');
      fm.forEach(f => {
        const b = h('button.bs-fmt', ic(f[2] || 'file', 22), h('div', h('b', f[1]), h('div.muted', '.' + (f[3] || f[0]))));
        b.addEventListener('click', () => { close(); if (f[0] === 'pdf-print') return this.printDoc(doc); this.save(doc, true, f[0]); });
        list.appendChild(b);
      });
      list.appendChild(h('button.bs-fmt', { onclick: () => { close(); this.printDoc(doc); } }, ic('file-type', 22), h('div', h('b', 'PDF (qua In → Lưu dưới dạng PDF)'), h('div.muted', 'Giữ nguyên bố cục trang'))));
      if (id === 'saveas') list.insertBefore(h('button.bs-fmt.bs-cloud', { onclick: () => { close(); import('./cloud.js').then(C => C.cloudSaveDialog(this, doc)); } }, ic('cloud-upload', 22), h('div', h('b', 'Lưu lên đám mây'), h('div.muted', 'Google Drive · OneDrive · Dropbox — tự đồng bộ, sửa cùng người khác'))), list.firstChild);
      P.appendChild(h('div.muted', { style: { marginBottom: '12px' } }, 'Chọn định dạng tệp. ' + (Host.kind === 'windows' ? 'Hộp thoại Lưu của Windows sẽ hiện ra để chọn thư mục.' : Host.kind === 'android' ? 'Trình chọn vị trí lưu của Android sẽ hiện ra.' : 'Tệp sẽ được tải xuống.')));
      P.appendChild(list);
    } else if (id === 'print') {
      title('In');
      P.appendChild(this.printPage(doc, close));
    } else if (id === 'share') {
      title('Chia sẻ');
      P.appendChild(h('div', { style: { marginBottom: '16px' } },
        h('button.btn.primary.big', {
          style: { display: 'flex', alignItems: 'center', gap: '8px', padding: '10px 18px', marginBottom: '6px' },
          onclick: () => {
            close();
            import('./cloud.js').then(C => C.airDropDialog(doc));
          }
        }, ic('wifi', 20), h('span', 'AirDrop & QuickShare (Mạng nội bộ / Đồng thời sửa)')),
        h('div.hint', 'Chia sẻ siêu tốc qua Wi-Fi / LAN và cho phép nhiều người cùng mở chỉnh sửa đồng thời.')
      ));
      const fm = doc.saveFormats();
      P.appendChild(h('div.muted', 'Gửi bản sao tài liệu qua ứng dụng khác (Zalo, Gmail, Drive…) hoặc lưu ra tệp để gửi:'));
      const row = h('div.bs-formats');
      fm.slice(0, 4).forEach(f => row.appendChild(h('button.bs-fmt', { onclick: async () => {
        const bytes = await doc.save(f[0]); const name = doc.name + '.' + f[0];
        if (!Host.share(name, bytes)) { await Host.saveFile(name, bytes, { saveAs: true }); }
        close();
      } }, ic('send', 22), h('div', h('b', 'Gửi dưới dạng ' + f[1]), h('div.muted', '.' + f[0])))));
      P.appendChild(row);
    } else if (id === 'account') {
      title('Tài khoản');
      const s = settings();
      const nm = h('input.inp', { value: s.userName }); const ini = h('input.inp', { value: s.initials, maxlength: 3 });
      P.appendChild(h('div.form', { style: { maxWidth: '420px' } }, h('label.lbl', 'Tên người dùng (hiển thị ở Tác giả, Chú thích, Theo dõi thay đổi)'), nm, h('label.lbl', 'Chữ viết tắt'), ini,
        h('button.btn.primary', { onclick: () => { saveSettings({ userName: nm.value || 'Người dùng', initials: ini.value || 'ND' }); toast('Đã lưu', 'ok'); } }, 'Lưu')));
      P.appendChild(h('div.bs-about', h('img', { src: 'assets/office-logo-full.png', alt: '' }), h('div', h('b', APP_NAME + ' ' + VERSION), h('div.muted', 'Bộ ứng dụng văn phòng mã nguồn mở: Word, Excel, PowerPoint, PDF, MathType, A.I'), h('div.muted', 'Nền tảng: ' + Host.kind + ' · ' + navigator.userAgent.replace(/^.*?\(|\).*$/g, '')), h('div.muted', 'Hoạt động hoàn toàn ngoại tuyến (offline).'))));
      P.appendChild(h('button.btn', { onclick: () => Host.openUrl('licenses.txt') }, 'Giấy phép mã nguồn mở'));
    } else if (id === 'options') {
      title('Tùy chọn');
      P.appendChild(optionsPage(doc, this));
    }
    return P;
  }
  printPage(doc, close) {
    const prev = h('div.print-prev');
    const settingsBox = h('div.print-set',
      h('button.btn.primary.big', { onclick: () => { close(); this.printDoc(doc); } }, ic('printer', 20), ' In'),
      h('div.muted', 'Chọn máy in hoặc "Lưu dưới dạng PDF" trong hộp thoại in của hệ thống.'),
      doc.printOptions ? doc.printOptions() : null);
    if (doc.printPreview) setTimeout(() => doc.printPreview(prev), 10);
    return h('div.print-wrap', settingsBox, prev);
  }
  printDoc(doc) {
    const area = h('div#print-area');
    document.body.appendChild(area);
    Promise.resolve(doc.renderPrint ? doc.renderPrint(area) : null).then(() => {
      document.body.classList.add('printing');
      setTimeout(() => {
        Host.print(doc.name);
        setTimeout(() => { document.body.classList.remove('printing'); if (area.parentNode) area.parentNode.removeChild(area); }, Host.kind === 'android' ? 4000 : 500);
      }, 250);
    });
  }
  inspect(doc) {
    const st = doc.stats ? doc.stats() : {};
    const issues = doc.inspect ? doc.inspect() : [];
    dialog({ title: 'Kiểm tra tài liệu', width: 520, body: h('div', h('h3', 'Thống kê'), h('table.kv', Object.keys(st).map(k => h('tr', h('td', k), h('td', String(st[k]))))),
      h('h3', 'Trợ năng (Accessibility)'), issues.length ? h('ul', issues.map(i => h('li', i))) : h('div.ok-line', ic('circle-check', 16), ' Không phát hiện vấn đề.')) });
  }

  // ---------------- start screen
  startView(kind) {
    const v = h('div.start');
    const rail = h('div.start-rail');
    const content = h('div.start-main');
    const sel = { k: kind || null };
    const railBtn = (k, label, icon, color) => {
      const b = h('button.rail-b' + (sel.k === k ? '.on' : ''), { style: color ? { color } : null }, ic(icon, 22), h('span', label));
      b.addEventListener('click', () => { sel.k = k; $$('.rail-b', rail).forEach(x => x.classList.remove('on')); b.classList.add('on'); render(); });
      return b;
    };
    rail.appendChild(h('img.rail-logo', { src: 'assets/office-logo-icon.png', alt: '' }));
    rail.appendChild(railBtn(null, 'Trang chủ', 'house'));
    rail.appendChild(railBtn('word', 'Word', 'file-text', KINDS.word.color));
    rail.appendChild(railBtn('excel', 'Excel', 'sheet', KINDS.excel.color));
    rail.appendChild(railBtn('ppt', 'PowerPoint', 'presentation', KINDS.ppt.color));
    rail.appendChild(railBtn('pdf', 'PDF', 'file-type', KINDS.pdf.color));
    rail.appendChild(railBtn('mathtype', 'MathType', 'function-square', KINDS.mathtype.color));
    const openB = h('button.rail-b', ic('folder-open', 22), h('span', 'Mở'));
    openB.addEventListener('click', () => this.openDialog(sel.k));
    rail.appendChild(openB);
    rail.appendChild(h('button.rail-b', { title: 'CamScanner — quét tài liệu bằng máy ảnh', onclick: () => import('../scan/scan.js').then(m => m.openScanner(this)) }, ic('scan-line', 22), h('span', 'Quét')));
    rail.appendChild(h('button.rail-b', { title: 'Đọc sách điện tử EPUB / FB2 / TXT', onclick: () => import('../reader/reader.js').then(m => m.openReaderDialog(this)) }, ic('book-open', 22), h('span', 'Đọc sách')));
    rail.appendChild(h('button.rail-b', { title: 'Android DeX Mode & Chiếu màn hình ngoài / PC Link', onclick: () => import('./dex.js').then(m => m.openDexGuideDialog()) }, ic('monitor', 22), h('span', 'DeX')));
    rail.appendChild(h('div.rail-sp'));
    rail.appendChild(h('button.rail-b', { onclick: () => import('../ai/ai.js').then(m => m.openAi(this, null)) }, ic('sparkles', 22), h('span', 'A.I')));
    rail.appendChild(h('button.rail-b', { onclick: () => optionsDialog(null, this) }, ic('settings', 22), h('span', 'Tùy chọn')));
    const render = async () => {
      content.innerHTML = '';
      v.style.setProperty('--accent', sel.k ? KINDS[sel.k].color : '#0f6cbd');
      const head = h('div.start-head', h('h1', greeting()), h('div.start-search', ic('search', 16), h('input', { placeholder: 'Tìm kiếm tệp gần đây' })));
      content.appendChild(head);
      const newHead = h('div.start-sec', h('h2', 'Mới'), h('span.sp'), sel.k ? h('a.more', { onclick: () => dialog({ title: 'Thêm mẫu', width: 900, body: this.templateStrip(sel.k, t => { closeAll(); this.newDoc(sel.k, t); }, true), buttons: [{ label: 'Đóng' }] }) }, 'Thêm mẫu →') : null);
      content.appendChild(newHead);
      if (sel.k) content.appendChild(this.templateStrip(sel.k, t => this.newDoc(sel.k, t)));
      else {
        const row = h('div.tpl-strip.apps');
        ['word', 'excel', 'ppt', 'pdf', 'mathtype'].forEach(k => {
          const b = h('button.tpl.app-tile', { style: { borderColor: KINDS[k].color } }, h('div.tpl-thumb.app', { style: { background: KINDS[k].color } }, ic(KINDS[k].icon, 44)), h('div.tpl-name', k === 'pdf' ? 'Công cụ PDF' : k === 'mathtype' ? 'MathType toán' : KINDS[k].label + ' trống'));
          b.addEventListener('click', () => this.newDoc(k));
          row.appendChild(b);
        });
        content.appendChild(row);

        // Android DeX & TV 4K Multi-Window Workspace Quick Launcher
        const dexBanner = h('div.dex-start-banner',
          h('div.dex-start-banner-content',
            h('div', { style: { display: 'flex', alignItems: 'center', gap: '12px' } },
              h('div.dex-banner-icon', ic('monitor', 26)),
              h('div',
                h('div.dex-banner-title', 'Không Gian Làm Việc Android DeX Desktop & TV 4K (v' + VERSION + ')'),
                h('div.dex-banner-desc', 'Đa nhiệm nhiều cửa sổ (Multi-Window 50/50), quản lý toàn bộ ứng dụng thiết bị ngoại vi và điều khiển phím/chuột tối ưu.')
              )
            ),
            h('div.dex-banner-actions',
              h('button.btn.primary', {
                onclick: () => DEX.openDexDesktopEnvironment({ mode: 'tvhome' })
              }, ic('monitor', 15), ' Vào DeX Desktop'),
              h('button.btn', {
                onclick: () => DEX.openDexDesktopEnvironment({ mode: 'multiwin' })
              }, ic('layout-grid', 15), ' Đa Nhiệm Multi-Window'),
              h('button.btn', {
                onclick: () => DEX.openDexDesktopEnvironment({ mode: 'tvapps' })
              }, ic('package', 15), ' Ứng Dụng Thiết Bị'),
              h('button.btn', {
                onclick: () => DEX.openVirtualTrackpad()
              }, ic('mouse-pointer', 15), ' Bàn Rê Chuột')
            )
          )
        );
        content.appendChild(dexBanner);
      }
      const rb = await this.recentBlock(sel.k, () => {}, head.querySelector('input'));
      content.appendChild(rb);
    };
    const closeAll = () => { $$('.dlg-back').forEach(d => d.parentNode && d.parentNode.removeChild(d)); };
    v.appendChild(rail); v.appendChild(content);
    render();
    return v;
  }
  templateStrip(kind, onPick, all) {
    const strip = h('div.tpl-strip' + (all ? '.all' : ''));
    const Cls = registry[kind];
    const tpls = (Cls && Cls.templates ? Cls.templates() : [{ id: 'blank', name: 'Trống' }]);
    (all ? tpls : tpls.slice(0, 8)).forEach(t => {
      const b = h('button.tpl' + (kind === 'word' ? '.portrait' : ''), h('div.tpl-thumb', t.thumb ? t.thumb() : h('div.tpl-blank')), h('div.tpl-name', t.name));
      b.addEventListener('click', () => onPick(t));
      strip.appendChild(b);
    });
    return strip;
  }
  async recentBlock(kind, close, searchInp) {
    const box = h('div.recent');
    const tabs = h('div.recent-tabs');
    let mode = 'recent';
    const list = h('div.recent-list');
    const draw = async () => {
      if (mode === 'device') return drawDevice();
      const all = (await recentList()).filter(r => !kind || r.kind === kind);
      const q = searchInp ? searchInp.value.trim().toLowerCase() : '';
      const items = all.filter(r => (mode === 'pinned' ? r.pinned : true) && (!q || r.name.toLowerCase().indexOf(q) >= 0));
      list.innerHTML = '';
      list.appendChild(h('div.recent-row.head', h('span.c-ic'), h('span.c-name', 'Tên'), h('span.c-date', 'Ngày sửa đổi')));
      if (!items.length) list.appendChild(h('div.recent-empty', mode === 'pinned' ? 'Ghim tệp để tìm nhanh: di chuột lên tệp và bấm biểu tượng ghim.' : 'Chưa có tài liệu gần đây. Tạo mới hoặc Mở một tệp để bắt đầu.'));
      items.forEach(r => {
        const k = KINDS[r.kind] || KINDS.word;
        const pin = h('button.pin' + (r.pinned ? '.on' : ''), { title: r.pinned ? 'Bỏ ghim' : 'Ghim' }, ic('pin', 14));
        pin.addEventListener('click', async e => { e.stopPropagation(); await recentPin(r.id, !r.pinned); draw(); });
        const row = h('div.recent-row', h('span.c-ic', { style: { color: k.color } }, ic(k.icon, 20)), h('span.c-name', h('div.rn', r.name), h('div.rp', r.path ? String(r.path).replace(/^content:\/\/[^/]+\/?/, 'Bộ nhớ › ').slice(0, 90) : 'Bản lưu trong ứng dụng')), h('span.c-date', friendlyDate(r.ts)), pin);
        row.addEventListener('click', () => { close && close(); this.openRecent(r); });
        row.addEventListener('contextmenu', e => contextMenu(e, [
          { label: 'Mở', icon: 'folder-open', run: () => { close && close(); this.openRecent(r); } },
          { label: r.pinned ? 'Bỏ ghim khỏi danh sách' : 'Ghim vào danh sách', icon: 'pin', run: async () => { await recentPin(r.id, !r.pinned); draw(); } },
          { label: 'Xóa khỏi danh sách', icon: 'trash-2', run: async () => { await recentRemove(r.id); draw(); } }
        ]));
        list.appendChild(row);
      });
    };
    // v0.1.8 — tab "Trên thiết bị": tệp Word/Excel/PowerPoint tìm được khi quét máy (Feature Check #28)
    const drawDevice = async (rescan) => {
      const DS = await import('./devscan.js'); const c = DS.cfg();
      const q = searchInp ? searchInp.value.trim().toLowerCase() : '';
      list.innerHTML = '';
      const info = h('div.recent-empty');
      const btn = h('button.btn', ic('refresh-cw', 13), ' Quét lại');
      const bar = h('div.dev-bar', { style: { display: 'flex', gap: '8px', alignItems: 'center', padding: '6px 4px' } }, btn, info);
      list.appendChild(bar);
      btn.addEventListener('click', () => drawDevice(true));
      if (Host.kind === 'android' && Host.scanFull === false) { const pb = h('button.btn', ic('shield-check', 13), ' Cấp quyền truy cập tệp'); pb.title = 'Cho phép đọc toàn bộ bộ nhớ để tìm mọi tệp Office (hiện chỉ quét Tài liệu và Tải xuống)'; pb.addEventListener('click', () => { Host.requestFilesAccess(); toast('Bật quyền rồi quay lại và bấm Quét lại'); }); bar.insertBefore(pb, info); }
      let files = c.files || [];
      if (rescan || !c.last) {
        info.textContent = 'Đang quét tệp Word, Excel, PowerPoint trên máy…'; btn.disabled = true;
        try { const r = await DS.scan(p => { info.textContent = 'Đang quét… ' + p.dirs + ' thư mục, ' + p.files + ' tệp'; }); if (r) files = r; else { info.textContent = 'Bản web không có quyền đọc ổ đĩa — dùng bản cài đặt Windows/Android, hoặc bấm Mở.'; btn.disabled = false; return; } }
        catch (e) { info.textContent = '⚠️ ' + e.message; btn.disabled = false; return; }
        btn.disabled = false; if (mode !== 'device') return;
      }
      const items = files.filter(f => (!kind || f.kind === kind) && (!q || f.name.toLowerCase().indexOf(q) >= 0));
      info.textContent = items.length + ' tệp' + (c.last ? ' · quét lúc ' + new Date(DS.cfg().last).toLocaleString('vi-VN') : '');
      list.appendChild(h('div.recent-row.head', h('span.c-ic'), h('span.c-name', 'Tên'), h('span.c-date', 'Ngày sửa đổi')));
      items.slice(0, 400).forEach(f => {
        const k = KINDS[f.kind] || KINDS.word;
        const row = h('div.recent-row', h('span.c-ic', { style: { color: k.color } }, ic(k.icon, 20)), h('span.c-name', h('div.rn', f.name), h('div.rp', String(f.dir || f.path || '').replace(/^content:\/\/[^/]+\/?/, 'Bộ nhớ › ').slice(0, 90))), h('span.c-date', f.mtime ? new Date(f.mtime).toLocaleDateString('vi-VN') : ''));
        row.addEventListener('click', () => { close && close(); this.openRecent({ name: f.name, kind: f.kind, path: f.path }); });
        list.appendChild(row);
      });
    };
    [['recent', 'Gần đây'], ['pinned', 'Đã ghim'], ['device', 'Trên thiết bị']].forEach(t => {
      const b = h('button.rt' + (t[0] === mode ? '.on' : ''), t[1]);
      b.addEventListener('click', () => { mode = t[0]; $$('.rt', tabs).forEach(x => x.classList.toggle('on', x === b)); draw(); });
      tabs.appendChild(b);
    });
    if (searchInp) searchInp.addEventListener('input', debounce(draw, 150));
    box.appendChild(tabs); box.appendChild(list);
    await draw();
    return box;
  }
}

function prop(k, v) { return h('div.prop', h('span.pk', k), h('span.pv', v == null ? '' : String(v))); }
export function greeting() { const hr = new Date().getHours(); return hr < 11 ? 'Chào buổi sáng' : hr < 13 ? 'Chào buổi trưa' : hr < 18 ? 'Chào buổi chiều' : 'Chào buổi tối'; }

import * as UPD from './update.js';
import * as CLOUD from './cloud.js';
import * as DEX from './dex.js';
export function optionsPage(doc, shell) {
  const s = settings();
  const m = s.mouse;
  const sec = (t, ...kids) => h('div.opt-sec', h('h3', t), kids);
  const chk = (label, get, set) => { const c = h('input', { type: 'checkbox' }); c.checked = !!get(); c.addEventListener('change', () => set(c.checked)); return h('label.chk', c, h('span', label)); };
  const num = (label, get, set, min, max) => { const i = h('input.inp.small', { type: 'number', value: get(), min, max }); i.addEventListener('change', () => set(+i.value)); return h('div.form-row.inline', h('label.lbl', label), i); };
  const sel = (label, opts, get, set) => { const i = h('select.inp.small', opts.map(o => h('option', { value: o[0] }, o[1]))); i.value = get(); i.addEventListener('change', () => set(i.value)); return h('div.form-row.inline', h('label.lbl', label), i); };
  const setM = (k, v) => { m[k] = v; saveSettings({ mouse: m }); applyMouseSettings(); };
  return h('div.options',
    sec('Chung',
      chk('Hiệu ứng Liquid Glass (kính mờ)', () => s.glass, v => { saveSettings({ glass: v }); document.body.classList.toggle('no-glass', !v); }),
      sel('Giao diện', [['light', 'Sáng'], ['dark', 'Tối'], ['auto', 'Theo hệ thống']], () => s.theme, v => { saveSettings({ theme: v }); applyTheme(); }),
      chk('Thu gọn Ribbon khi mở tài liệu', () => s.ribbonCollapsed, v => saveSettings({ ribbonCollapsed: v })),
      sel('Engine mở Word / Excel / PowerPoint', [['native', 'Native – LibreOffice core (chuẩn Microsoft Office, khuyên dùng)'], ['web', 'Web – trình soạn thảo Yomi-Office Open Source']], () => s.officeEngine || 'native', v => { saveSettings({ officeEngine: v }); toast(v === 'native' ? 'Tệp Office sẽ mở bằng engine Native' : 'Tệp Office sẽ mở bằng giao diện Web'); }),
      chk('Quét tệp Word / Excel / PowerPoint trên máy khi mở ứng dụng', () => (s.devscan || {}).onStart !== false, v => { const d = Object.assign({}, s.devscan || {}, { onStart: v }); saveSettings({ devscan: d }); }),
      chk('Mở mỗi tệp Word / Excel / PowerPoint / PDF trong cửa sổ mới', () => s.newWindowPerFile !== false, v => saveSettings({ newWindowPerFile: v })),
      chk('Chế độ Cảm ứng (nút Ribbon lớn hơn, dễ chạm)', () => !!s.touchMode, v => { saveSettings({ touchMode: v }); applyMouseSettings(); })),
    sec('Vùng an toàn & Chống che khuất (Safe Area & Anti-Obscuration)',
      chk('Tự động chống che khuất bởi thanh trạng thái & thanh điều hướng', () => s.antiObscuration !== false, v => {
        saveSettings({ antiObscuration: v });
        document.body.classList.toggle('anti-obscuration-off', !v);
        toast(v ? 'Đã bật bảo vệ vùng an toàn' : 'Đã tắt bảo vệ vùng an toàn');
      }),
      num('Đệm viền trên bổ sung (Top Padding - px)', () => s.safePadTop || 0, v => {
        saveSettings({ safePadTop: v });
        document.documentElement.style.setProperty('--extra-safe-top', (v || 0) + 'px');
      }, 0, 100),
      num('Đệm viền dưới bổ sung (Bottom Padding - px)', () => s.safePadBottom || 0, v => {
        saveSettings({ safePadBottom: v });
        document.documentElement.style.setProperty('--extra-safe-bottom', (v || 0) + 'px');
      }, 0, 120),
      h('div.hint', 'Tự động đảm bảo thanh công cụ, các nút bấm, hộp thoại Cài đặt không bao giờ bị thanh thông báo (trên) hoặc thanh điều hướng 3 nút (dưới) che lấp.')),
    UPD && UPD.optionsSection(sec, chk),
    CLOUD.optionsSection(sec),
    DEX.optionsSection(sec, chk, sel, num),
    sec('A.I',
      h('div.form-row.inline', h('label.lbl', 'Web (ChatGPT, Gemini, Copilot, Claude, DeepSeek, Grok, AI Hay) hoặc khóa API'), h('button.btn', { onclick: () => import('../ai/ai.js').then(A => A.aiSettings()) }, ic('sparkles', 14), ' Cài đặt A.I…'))),
    sec('Phông chữ',
      h('div.form-row.inline', h('label.lbl', 'Phông tích hợp, Font Pack và phông của máy'), h('button.btn', { onclick: () => import('./fonts.js').then(F => F.fontManager()) }, ic('type', 14), ' Quản lý phông chữ…'))),
    sec('Lưu',
      chk('Tự động lưu bản khôi phục (AutoSave)', () => s.autosave, v => saveSettings({ autosave: v })),
      num('Chu kỳ tự động lưu (giây)', () => s.autosaveSec, v => saveSettings({ autosaveSec: Math.max(5, v || 30) }), 5, 600),
      num('Số tệp gần đây tối đa', () => s.recentMax, v => saveSettings({ recentMax: Math.max(5, v || 30) }), 5, 200)),
    sec('Chuột & con trỏ (Mouse / Cursor)',
      chk('Ctrl + lăn chuột để phóng to/thu nhỏ', () => m.wheelZoomCtrl, v => setM('wheelZoomCtrl', v)),
      chk('Cuộn mượt (Smooth scrolling)', () => m.smoothScroll, v => setM('smoothScroll', v)),
      num('Số dòng cuộn mỗi nấc lăn chuột', () => m.scrollLines, v => setM('scrollLines', Math.max(1, Math.min(20, v || 3))), 1, 20),
      chk('Menu chuột phải (Context menu)', () => m.rightClickMenus, v => setM('rightClickMenus', v)),
      chk('Kéo thả văn bản / ô / đối tượng (Drag & drop)', () => m.dragDrop, v => setM('dragDrop', v)),
      chk('Bấm đúp để sửa ô / đối tượng', () => m.doubleClickEdit, v => setM('doubleClickEdit', v)),
      chk('Hiện thanh công cụ nhỏ khi chọn văn bản (Mini toolbar)', () => m.miniToolbar, v => setM('miniToolbar', v)),
      chk('Xem trước trực tiếp khi rê chuột (Live preview)', () => m.hoverPreview, v => setM('hoverPreview', v)),
      sel('Kiểu con trỏ', [['default', 'Mặc định của hệ thống'], ['large', 'Con trỏ lớn (dễ nhìn)'], ['precise', 'Con trỏ chính xác (dấu cộng)']], () => m.cursorStyle, v => setM('cursorStyle', v)),
      chk('Con trỏ văn bản nhấp nháy', () => s.cursorBlink, v => { saveSettings({ cursorBlink: v }); applyMouseSettings(); }),
      chk('Bấm nút giữa (con lăn) để cuộn tự động (Autoscroll)', () => m.middleAutoscroll !== false, v => setM('middleAutoscroll', v)),
      chk('Excel: lăn chuột để phóng to/thu nhỏ (Zoom on roll with IntelliMouse)', () => !!m.wheelZoomNoCtrl, v => setM('wheelZoomNoCtrl', v)),
      chk('Excel: gõ trực tiếp vào ô đang chọn (không cần bấm đúp)', () => m.typeToEdit !== false, v => setM('typeToEdit', v)),
      sel('Excel: sau khi nhấn Enter, di chuyển vùng chọn', [['down', 'Xuống'], ['right', 'Sang phải'], ['up', 'Lên'], ['left', 'Sang trái'], ['none', 'Không di chuyển']], () => m.enterMove || 'down', v => setM('enterMove', v)),
      chk('Word: Bấm-và-Gõ (bấm đúp vào chỗ trống trên trang để gõ)', () => m.clickAndType !== false, v => setM('clickAndType', v)),
      chk('Hiện vệt sáng khi bấm chuột (hữu ích khi trình chiếu/quay màn hình)', () => !!m.clickHighlight, v => setM('clickHighlight', v)),
      sel('Kích thước con trỏ văn bản (caret)', [['1', 'Mảnh'], ['2', 'Vừa'], ['3', 'Dày']], () => String(m.caretWidth || 1), v => setM('caretWidth', +v))),
    sec('Chỉnh sửa',
      sel('Phông chữ mặc định Word', [['Calibri', 'Calibri'], ['Times New Roman', 'Times New Roman'], ['Arial', 'Arial'], ['Roboto', 'Roboto'], ['Segoe UI', 'Segoe UI']], () => s.wordFont, v => saveSettings({ wordFont: v })),
      num('Cỡ chữ mặc định Word', () => s.wordSize, v => saveSettings({ wordSize: v || 11 }), 6, 72),
      sel('Phông chữ mặc định Excel', [['Calibri', 'Calibri'], ['Times New Roman', 'Times New Roman'], ['Arial', 'Arial'], ['Roboto', 'Roboto']], () => s.excelFont, v => saveSettings({ excelFont: v }))),
    h('div.muted', APP_NAME + ' ' + VERSION + ' — mọi thiết lập được lưu trên thiết bị này.'));
}
export function applySafeInsetsSettings() {
  const s = settings();
  document.body.classList.toggle('anti-obscuration-off', s.antiObscuration === false);
  if (s.safePadTop) document.documentElement.style.setProperty('--extra-safe-top', s.safePadTop + 'px');
  if (s.safePadBottom) document.documentElement.style.setProperty('--extra-safe-bottom', s.safePadBottom + 'px');
}
export function applyTheme() {
  const t = settings().theme;
  const dark = t === 'dark' || (t === 'auto' && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches);
  document.body.classList.toggle('dark', !!dark);
  document.body.classList.toggle('no-glass', !settings().glass);
}
export function applyMouseSettings() {
  const m = settings().mouse;
  document.body.classList.toggle('cur-large', m.cursorStyle === 'large');
  document.body.classList.toggle('cur-precise', m.cursorStyle === 'precise');
  document.body.classList.toggle('no-smooth', !m.smoothScroll);
  document.body.classList.toggle('no-blink', !settings().cursorBlink);
  document.body.classList.toggle('touch-mode', !!settings().touchMode);
  document.body.classList.toggle('caret-2', m.caretWidth === 2); document.body.classList.toggle('caret-3', m.caretWidth === 3);
  import('./mouse.js').then(M => M.installMouse(m));
}
