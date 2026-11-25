// v0.1.7: Backstage (Tệp) and Tùy chọn (Options) shared with the native (LibreOfficeKit) editor.
// Navigation, tiles, file-type lists and every option come from ui_settings.json - the native editor renders the very
// same file (UnifiedUi.cs), so both editors show identical pages, icons, order and settings (stored in settings.json).
import SPEC from './ui_settings.json';
import L from './ui_layout.json';
import { h, ic, APP_NAME, VERSION } from './util.js';
import { toast, dialog } from './ui.js';
import { settings, saveSettings } from './store.js';
import Host from './host.js';
import { toNative } from './unify.js';

export { SPEC };
const KIND = { word: 'word', excel: 'excel', ppt: 'ppt' };
const appIcon = k => 'n:' + (k === 'excel' ? 'excel' : k === 'ppt' ? 'powerpoint' : 'word');
const icon = (n, sz) => ic(n && n.indexOf(':') < 0 ? 'n:' + n : n, sz);
export function greeting() { const hr = new Date().getHours(); return hr < 11 ? 'Chào buổi sáng' : hr < 13 ? 'Chào buổi trưa' : hr < 18 ? 'Chào buổi chiều' : 'Chào buổi tối'; }
export function friendlyDate(ts) {
  const d = new Date(ts || Date.now()), t = new Date(); const p = n => (n < 10 ? '0' : '') + n;
  const hm = p(d.getHours()) + ':' + p(d.getMinutes());
  if (d.toDateString() === t.toDateString()) return 'Hôm nay lúc ' + hm;
  const y = new Date(t); y.setDate(t.getDate() - 1); if (d.toDateString() === y.toDateString()) return 'Hôm qua lúc ' + hm;
  return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear();
}

// ---------- settings get/set by path ("mouse.smoothScroll") ----------
function findItem(key) { for (const c of SPEC.options) for (const s of c.secs) for (const i of s.items) if (i.key === key) return i; return null; }
export function optGet(key) {
  const s = settings(); const parts = key.split('.'); let v = s;
  for (const p of parts) { v = v == null ? undefined : v[p]; }
  if (v === undefined) { const it = findItem(key); v = it ? it.d : undefined; }
  if (key === 'autosaveMin' && s.autosaveMin == null && s.autosaveSec) v = Math.max(1, Math.round(s.autosaveSec / 60));
  return v;
}
function optSet(patch) {
  const s = settings(); const top = {};
  Object.keys(patch).forEach(key => {
    const parts = key.split('.'); const v = patch[key];
    if (parts.length === 1) top[key] = v;
    else { const o = Object.assign({}, s[parts[0]] || {}); o[parts[1]] = v; top[parts[0]] = Object.assign(top[parts[0]] || o, { [parts[1]]: v }); }
  });
  if ('autosaveMin' in patch) top.autosaveSec = Math.max(1, +patch.autosaveMin || 2) * 60;
  if ('dex.enabled' in patch) top.dexForceDesktop = !!patch['dex.enabled'];
  saveSettings(top);
}
async function applyHooks(changed) {
  const S = await import('./shell.js');
  if ('theme' in changed) S.applyTheme();
  if ('glass' in changed) document.body.classList.toggle('no-glass', !changed.glass);
  if ('touchMode' in changed || 'cursorBlink' in changed || Object.keys(changed).some(k => k.indexOf('mouse.') === 0)) S.applyMouseSettings();
  if ('antiObscuration' in changed || 'safePadTop' in changed || 'safePadBottom' in changed) {
    document.body.classList.toggle('anti-obscuration-off', changed.antiObscuration === false);
    document.documentElement.style.setProperty('--extra-safe-top', (optGet('safePadTop') || 0) + 'px');
    document.documentElement.style.setProperty('--extra-safe-bottom', (optGet('safePadBottom') || 0) + 'px');
  }
  if (Object.keys(changed).some(k => k.indexOf('dex.') === 0)) {
    const D = await import('./dex.js');
    if ('dex.enabled' in changed) D.applyDexMode(!!changed['dex.enabled']);
    if ('dex.resolution' in changed) D.applyDexResolution(changed['dex.resolution']);
    if ('dex.uiScale' in changed) D.applyDexScale(+changed['dex.uiScale']);
    if ('dex.wakeLock' in changed) Host.setKeepScreenOn(!!changed['dex.wakeLock'] && D.isDexEnvironment());
    if ('dex.webLinkEnabled' in changed) { if (changed['dex.webLinkEnabled']) await Host.startWebLink(+optGet('dex.webLinkPort') || 8080); else await Host.stopWebLink(); }
  }
  if ('ribbonHide' in changed) toast('Tab Ribbon sẽ thay đổi khi mở tài liệu tiếp theo.');
}
export async function optAction(act, shell, doc) {
  if (act === 'ai') return import('../ai/ai.js').then(A => A.aiSettings());
  if (act === 'fonts') return import('./fonts.js').then(F => F.fontManager());
  if (act === 'cloud') return import('./cloud.js').then(C => C.accountsDialog());
  if (act === 'dex') return import('./dex.js').then(D => D.openDexControlCenterDialog());
  if (act === 'cloudOpen') return import('./cloud.js').then(C => C.cloudOpenDialog(shell || window.__oosShell));
  if (act === 'updateNow') return import('./update.js').then(U => U.checkNow(true));
}

// ---------- Tùy chọn: the same dialog as the native editor (category list on the left, page on the right, OK / Hủy) ----------
export function optionsDialog(doc, shell) {
  const kind = doc ? KIND[doc.kind] : null;
  const pend = {};
  const val = k => (k in pend ? pend[k] : optGet(k));
  const cats = h('div.uo-cats'), page = h('div.uo-page');
  const vis = x => !x.only || x.only === kind;
  const list = SPEC.options;
  const show = ci => {
    Array.prototype.slice.call(cats.children).forEach((b, i) => b.classList.toggle('on', i === ci));
    page.innerHTML = ''; page.scrollTop = 0;
    list[ci].secs.filter(vis).forEach(s => {
      page.appendChild(h('div.uo-head', s.h)); if (s.sub) page.appendChild(h('div.uo-sub', s.sub));
      s.items.filter(vis).forEach(it => {
        if (it.k === 'chk') { const c = h('input', { type: 'checkbox' }); c.checked = !!val(it.key); c.addEventListener('change', () => { pend[it.key] = c.checked; }); page.appendChild(h('label.uo-chk', c, h('span', it.t))); }
        else if (it.k === 'num') { const n = h('input.uo-num', { type: 'number', min: it.min, max: it.max, value: val(it.key) }); n.addEventListener('change', () => { pend[it.key] = Math.max(it.min, Math.min(it.max, +n.value || 0)); }); page.appendChild(h('div.uo-row', h('label.uo-lbl.w260', it.t), n)); }
        else if (it.k === 'sel') { const e = h('select.uo-sel', it.opts.map(o => h('option', { value: o[0] }, o[1]))); e.value = String(val(it.key)); e.addEventListener('change', () => { pend[it.key] = e.value; }); page.appendChild(h('div.uo-row', h('label.uo-lbl.w200', it.t), e)); }
        else if (it.k === 'text') { const t = h('input.uo-text', { value: val(it.key) || (it.key === 'userName' ? settings().userName : '') }); t.addEventListener('change', () => { pend[it.key] = t.value; }); page.appendChild(h('div.uo-row', h('label.uo-lbl.w200', it.t), t)); }
        else if (it.k === 'note') page.appendChild(h('div.uo-note', it.t));
        else if (it.k === 'btn') page.appendChild(h('button.uo-btn', { onclick: () => optAction(it.act, shell, doc) }, icon(it.i, 16), h('span', it.t)));
        else if (it.k === 'tabs') {
          const lay = L[kind || 'word']; const hid = String(val(it.key) || '');
          (lay ? lay.tabs : []).forEach((t, ti) => {
            const c = h('input', { type: 'checkbox' }); c.checked = hid.indexOf('|' + t.name + '|') < 0; c.disabled = ti === 0;
            c.addEventListener('change', () => { let x = String(val(it.key) || '').split('|' + t.name + '|').join(''); if (!c.checked) x += '|' + t.name + '|'; pend[it.key] = x; });
            page.appendChild(h('label.uo-chk.tight', c, h('span', t.name)));
          });
        }
      });
    });
  };
  list.forEach((c, i) => { const b = h('button.uo-cat', kind === 'excel' && c.tx ? c.tx : c.t); b.addEventListener('click', () => show(i)); cats.appendChild(b); });
  const body = h('div.uo-wrap', cats, page);
  setTimeout(() => show(0), 0);
  return dialog({ title: 'Tùy chọn', width: 740, cls: 'uo-dlg', body, buttons: [
    { label: 'OK', primary: true, run: async () => { if (Object.keys(pend).length) { optSet(pend); await applyHooks(pend); } } },
    { label: 'Hủy' }] });
}
// kept for older callers (start screen rail)
export function optionsPage(doc, shell) { const b = h('button.uo-btn', { onclick: () => optionsDialog(doc, shell) }, icon('options', 16), h('span', 'Tùy chọn…')); return b; }

// ---------- Backstage ----------
const H1 = t => h('div.ub-h1', t);
const H2 = t => h('div.ub-h2', t);
const P = (t, cls) => h('div.ub-p' + (cls || ''), t);
function tile(ic0, title, sub, run, w) {
  const b = h('button.ub-tile' + (sub ? '' : '.nosub') + (run ? '' : '.off'), { style: w ? { width: w + 'px' } : null },
    h('span.ub-ti', icon(ic0, sub ? 32 : 24)), h('span.ub-tt', h('span.ub-t1', title), sub ? h('span.ub-t2', sub) : null));
  if (run) b.addEventListener('click', run);
  return b;
}
function specTiles(page) { return SPEC.tiles[page] || []; }
function fmtList(doc) {
  const k = KIND[doc.kind];
  if (k) return SPEC.formats[k].map(f => ({ ext: f[0], label: f[1], where: f[2] }));
  return (doc.saveFormats ? doc.saveFormats() : []).map(f => ({ ext: f[0], label: f[1], where: 'b' }));
}
function fileExt(doc, ext) { const fm = (doc.saveFormats ? doc.saveFormats() : []).find(f => f[0] === ext); return (fm && fm[3]) || ext; }
async function saveFormat(shell, doc, f, close, toPath) {
  if (f.where === 'n') { close && close(); return toNative(doc, 'bs:fmt:' + f.ext, f.label); }
  if (f.ext === 'pdf' && !(doc.saveFormats ? doc.saveFormats() : []).some(x => x[0] === 'pdf')) { close && close(); return shell.printDoc(doc); }
  close && close(); return shell.save(doc, true, f.ext, toPath);
}

export async function bsAction(shell, doc, id, close) {
  const C = () => import('./cloud.js'), SU = () => import('./suites.js');
  close = close || (() => { const b = document.querySelector('.backstage'); if (b) b.remove(); });
  switch (id) {
    case 'browse': case 'pc': close(); return shell.openDialog();
    case 'cloudOpen': close(); return C().then(c => c.cloudOpenDialog(shell));
    case 'cloudSave': close(); return C().then(c => c.cloudSaveDialog(shell, doc));
    case 'reveal': if (Host.kind === 'windows' && doc.path) return Host.reveal(doc.path); return toast(doc.path ? doc.path : 'Tài liệu chưa được lưu vào thư mục nào.');
    case 'protect': doc.readOnly = !doc.readOnly; doc.setReadOnly && doc.setReadOnly(doc.readOnly); toast(doc.readOnly ? 'Đã đánh dấu bản cuối cùng (chỉ đọc)' : 'Đã mở khóa chỉnh sửa'); return close();
    case 'macros': { const v = !optGet('macrosAllowed'); optSet({ macrosAllowed: v }); toast('Đã ' + (v ? 'bật' : 'tắt') + ' macro. Mở lại tài liệu để áp dụng.'); return shell.backstage(doc, 'info'); }
    case 'inspect': return SU().then(m => m.openTikaInspector(doc));
    case 'versions': return SU().then(m => m.openBackupCenter(doc));
    case 'lanshare': close(); return C().then(c => c.airDropDialog(doc));
    case 'email': {
      const ext = doc.fileExt || doc.meta.newExt; const bytes = await doc.save(ext); const name = doc.name + '.' + fileExt(doc, ext);
      if (Host.share(name, bytes)) return;
      const r = await Host.saveFile(name, bytes, { saveAs: true }); if (!r) return;
      Host.openUrl('mailto:?subject=' + encodeURIComponent(name)); if (r.path && Host.kind === 'windows') Host.reveal(r.path); return;
    }
    case 'sendpdf': case 'pdf': close(); return shell.printDoc(doc);
    case 'options': return optionsDialog(doc, shell);
  }
  if (id.indexOf('page:') === 0) return shell.backstage(doc, id.slice(5));
  if (id.indexOf('fmt:') === 0) { const f = fmtList(doc).find(x => x.ext === id.slice(4)) || { ext: id.slice(4), where: 'b' }; return saveFormat(shell, doc, Object.assign({}, f, { where: f.where === 'n' ? 'b' : f.where }), close); }
  if (optAction(id, shell, doc)) return;
}

export async function bsPage(shell, doc, id, close) {
  const P0 = h('div.ub-page');
  const act = a => () => bsAction(shell, doc, a, close);
  const recent = async () => {
    P0.appendChild(H2('Gần đây'));
    const q = h('input.ub-search', { placeholder: 'Tìm kiếm tệp' });
    P0.appendChild(q);
    const blk = await shell.recentBlock(null, close, q);
    P0.appendChild(blk);
  };
  if (id === 'home' || id === 'new') {
    P0.appendChild(H1(id === 'home' ? greeting() : 'Mới'));
    P0.appendChild(H2('Mới'));
    const cur = KIND[doc.kind] || 'word';
    const order = [cur, cur === 'word' ? 'excel' : 'word', cur === 'ppt' ? 'excel' : 'ppt'];
    const names = { word: 'Tài liệu trống', excel: 'Sổ làm việc trống', ppt: 'Bản trình bày trống' };
    const row = h('div.ub-tpls');
    order.forEach(k => {
      const card = h('button.ub-tpl.k-' + k, h('span.ub-card', k === 'ppt' ? h('span.ub-cap', 'Bấm để thêm tiêu đề') : null, h('span.ub-cardic', icon(appIcon(k), 16))), h('span.ub-tname', names[k]));
      card.addEventListener('click', () => { close(); shell.newDoc(k); });
      row.appendChild(card);
    });
    const more = h('a.ub-more', { href: '#' }, 'Thêm mẫu  →');
    const all = h('div.ub-alltpl'); all.style.display = 'none';
    more.addEventListener('click', e => { e.preventDefault(); if (!all.firstChild) all.appendChild(shell.templateStrip(doc.kind, t => { close(); shell.newDoc(doc.kind, t); }, true)); all.style.display = all.style.display === 'none' ? '' : 'none'; });
    row.appendChild(more); P0.appendChild(row); P0.appendChild(all);
    if (id === 'home') await recent();
  } else if (id === 'open') {
    P0.appendChild(H1('Mở'));
    const col = h('div.ub-col');
    specTiles('open').forEach(t => col.appendChild(tile(t[1], t[2], t[0] === 'pc' ? (Host.kind === 'android' ? 'Bộ nhớ máy' : 'Tài liệu') : t[3], act(t[0]), 300)));
    P0.appendChild(col);
    await recent();
  } else if (id === 'info') {
    P0.appendChild(H1('Thông tin'));
    P0.appendChild(P(doc.name + (doc.fileExt ? '.' + doc.fileExt : ''), '.big'));
    P0.appendChild(P(doc.path ? String(doc.path).replace(/[\\/][^\\/]*$/, '') : '(chưa lưu)'));
    const left = h('div.ub-col');
    specTiles('info').forEach(t => {
      let title = t[2]; if (t[0] === 'macros') title = 'Macro: ' + (optGet('macrosAllowed') ? 'đang bật' : 'đang tắt');
      left.appendChild(tile(t[1], title, t[3], act(t[0]), 420));
    });
    // same properties, same order as the native editor
    const st = doc.stats ? doc.stats() : {};
    const props = h('div.ub-props', H2('Thuộc tính'));
    const prop = (k, v) => props.appendChild(h('div.ub-prop', h('span.k', k), h('span.v', String(v))));
    const fi = doc.path && Host.kind === 'windows' ? await Host.fileInfo(doc.path) : null;
    const p2 = n => (n < 10 ? '0' : '') + n, dt = ms => { const d = new Date(ms); return p2(d.getDate()) + '/' + p2(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + p2(d.getHours()) + ':' + p2(d.getMinutes()); };
    const size = fi ? fi.size : doc._fileSize;
    if (size) prop('Kích thước', size >= 1048576 ? (size / 1048576).toFixed(1).replace('.', ',') + ' MB' : Math.max(1, Math.floor(size / 1024)) + ' KB');
    if (fi) { prop('Ngày tạo', dt(fi.ctime)); prop('Lần sửa đổi cuối', dt(fi.mtime)); }
    if (doc.kind === 'word') { prop('Số trang', st['Số trang'] || 1); prop('Số từ', (st['Số từ'] || 0) + ' từ'); }
    else if (doc.kind === 'excel') prop('Số trang tính', doc.wb && doc.wb.sheets ? doc.wb.sheets.length : 1);
    else if (doc.kind === 'ppt') prop('Số trang chiếu', doc.slides ? doc.slides.length : 1);
    prop('Tác giả', settings().userName); prop('Định dạng', String(doc.fileExt || doc.meta.newExt).toUpperCase()); prop('Engine', APP_NAME + ' (Web)');
    if (doc.readOnly) prop('Trạng thái', 'Chỉ đọc');
    P0.appendChild(h('div.ub-two', left, props));
  } else if (id === 'saveas') {
    P0.appendChild(H1('Lưu thành'));
    const left = h('div.ub-col.narrow');
    const fm = fmtList(doc);
    let dir = null;
    const pathL = h('div.ub-path'), nameI = h('input.ub-name', { value: doc.name }), typeS = h('select.ub-type', fm.map((f, i) => h('option', { value: i }, f.label + ' (*.' + (f.ext === 'xml2003' ? 'xml' : f.ext === 'video' ? 'webm' : f.ext) + ')')));
    const ci = fm.findIndex(f => f.ext === (doc.fileExt || doc.meta.newExt)); typeS.value = String(Math.max(0, ci));
    const listEl = h('div.ub-folder');
    const saveB = h('button.ub-save', 'Lưu');
    const browse = async d => {
      if (Host.kind !== 'windows') { pathL.textContent = Host.kind === 'android' ? '📁  Trình chọn vị trí lưu của Android sẽ hiện ra' : '📁  Thư mục Tải xuống'; return; }
      const r = await Host.listDir(d || ''); if (!r || !r.ok) return;
      dir = r.dir; pathL.textContent = '📁  ' + dir; pathL.onclick = () => r.parent && browse(r.parent);
      listEl.innerHTML = ''; listEl.appendChild(h('div.ub-frow.head', h('span.n', 'Tên'), h('span.d', 'Ngày sửa đổi')));
      const exts = fm.map(f => '.' + fileExt(doc, f.ext));
      r.items.filter(it => it.dir || exts.some(e => it.name.toLowerCase().endsWith(e))).forEach(it => {
        const row = h('div.ub-frow', h('span.n', (it.dir ? '📁  ' : '📄  ') + it.name), h('span.d', friendlyDate(it.mtime)));
        row.addEventListener('dblclick', () => { if (it.dir) browse(it.path); else nameI.value = it.name.replace(/\.[^.]+$/, ''); });
        row.addEventListener('click', () => { if (!it.dir) nameI.value = it.name.replace(/\.[^.]+$/, ''); });
        listEl.appendChild(row);
      });
    };
    saveB.addEventListener('click', () => {
      const f = fm[+typeS.value || 0]; doc.name = nameI.value.trim() || doc.name;
      const target = dir && Host.kind === 'windows' ? dir.replace(/[\\/]$/, '') + '\\' + doc.name + '.' + fileExt(doc, f.ext) : null;
      saveFormat(shell, doc, f, close, target);
    });
    specTiles('saveas').forEach((t, i) => {
      const run = t[0] === 'recent' || t[0] === 'pc' ? () => browse(t[0] === 'pc' ? '' : (doc.path ? String(doc.path).replace(/[\\/][^\\/]*$/, '') : ''))
        : t[0] === 'addplace' ? async () => { if (Host.kind !== 'windows') return toast('Dùng "Duyệt" để chọn vị trí.'); const p = await Host.pickFolder(); if (p) browse(p); }
        : t[0] === 'browse' ? () => saveFormat(shell, doc, Object.assign({}, fm[+typeS.value || 0]), close) : act(t[0]);
      const b = tile(t[1], t[2], null, run, 260); if (i === 0) b.classList.add('sel'); left.appendChild(b);
    });
    const right = h('div.ub-right', pathL, h('div.ub-line', nameI, saveB), h('div.ub-line', typeS), listEl);
    P0.appendChild(h('div.ub-two', left, right));
    browse(doc.path ? String(doc.path).replace(/[\\/][^\\/]*$/, '') : '');
  } else if (id === 'print') {
    P0.appendChild(H1('In'));
    P0.appendChild(shell.printPage(doc, close));
  } else if (id === 'share') {
    P0.appendChild(H1('Chia sẻ'));
    const col = h('div.ub-col'); specTiles('share').forEach(t => col.appendChild(tile(t[1], t[2], t[3], act(t[0]), 520))); P0.appendChild(col);
  } else if (id === 'export') {
    P0.appendChild(H1('Xuất'));
    const col = h('div.ub-col'); specTiles('export').forEach(t => col.appendChild(tile(t[1], t[2], t[3], act(t[0]), 520))); P0.appendChild(col);
    P0.appendChild(H2('Thay đổi loại tệp'));
    const col2 = h('div.ub-col');
    fmtList(doc).forEach(f => col2.appendChild(tile(f.ext === 'pdf' ? 'pdf' : appIcon(KIND[doc.kind] || 'word'), f.label, '*.' + (f.ext === 'xml2003' ? 'xml' : f.ext === 'video' ? 'webm' : f.ext), () => saveFormat(shell, doc, f, close), 520)));
    P0.appendChild(col2);
  } else if (id === 'account') {
    P0.appendChild(H1('Tài khoản'));
    P0.appendChild(H2('Thông tin người dùng'));
    P0.appendChild(P('Tên của bạn (dùng cho ghi chú, theo dõi thay đổi):'));
    const nm = h('input.ub-name', { value: settings().userName });
    P0.appendChild(h('div.ub-line', nm, h('button.ub-btn', { onclick: () => { saveSettings({ userName: nm.value.trim() || 'Người dùng' }); toast('Đã lưu tên người dùng.', 'ok'); } }, 'Lưu tên')));
    P0.appendChild(H2('Thông tin sản phẩm'));
    P0.appendChild(P(APP_NAME + ' ' + VERSION, '.strong'));
    P0.appendChild(P('Engine Web: ' + APP_NAME + ' ' + VERSION + ' (HTML5, ' + (Host.kind === 'android' ? 'Android WebView' : Host.kind === 'windows' ? 'WebView2' : 'trình duyệt') + ')'));
    P0.appendChild(P('Ứng dụng mã nguồn mở, miễn phí. Giao diện tham khảo bố cục Microsoft 365.'));
    P0.appendChild(H2('Cập nhật phần mềm'));
    const U = await import('./update.js'); const c = U.cfg();
    P0.appendChild(P('Nguồn: GitHub Releases (chính), Google Drive và MediaFire (phụ) của Yomi-Office Open Source. Lần kiểm tra gần nhất: ' + (c.last ? friendlyDate(c.last) : 'chưa kiểm tra')));
    const cb = h('input', { type: 'checkbox' }); cb.checked = c.auto !== false; cb.addEventListener('change', () => { c.auto = cb.checked; saveSettings({ update: c }); });
    P0.appendChild(h('label.ub-chk', cb, h('span', 'Tự động kiểm tra khi mở tài liệu (mỗi ngày một lần)')));
    P0.appendChild(h('div.ub-col', tile('update', 'Cập nhật ngay', 'Kiểm tra phiên bản mới và tự động cài đặt', () => U.checkNow(true), 420)));
  }
  return P0;
}
