// v0.0.7 — Software update (Tùy chọn → Cập nhật phần mềm)
// On start-up the app scans the two public release folders (Google Drive + MediaFire). Each release folder holds the installers and
// a small version.json ({version, date, notes, files:{<edition>: <file name>}}). Without version.json the file names
// "OfficeOpenSource-X.Y.Z-<edition>.<ext>" are used. If a newer version exists the user is told and can install it automatically.
import { h, ic, esc, VERSION } from './util.js';
import { toast, dialog } from './ui.js';
import { settings, saveSettings } from './store.js';
import Host from './host.js';

export const DEFAULT_SOURCES = {
  github: 'YomiHub-Project/Yomi-Office-Open-Source', // v0.1.8 — nguồn chính: https://github.com/YomiHub-Project/Yomi-Office-Open-Source/releases
  drive: '1kkqNrHFGQXLbf7PPYYc6HWgD1mFOdIyr',   // https://drive.google.com/drive/folders/1kkqNrHFGQXLbf7PPYYc6HWgD1mFOdIyr
  mediafire: 'seqly6bv55yr8'                      // https://www.mediafire.com/folder/seqly6bv55yr8/Office_Open_Source
};
export function cfg() { const s = settings(); s.update = Object.assign({ auto: true, autoInstall: false, github: DEFAULT_SOURCES.github, drive: DEFAULT_SOURCES.drive, mediafire: DEFAULT_SOURCES.mediafire, last: 0, skip: '' }, s.update || {}); return s.update; }

export function cmpVer(a, b) { const x = String(a).split('.').map(n => parseInt(n, 10) || 0), y = String(b).split('.').map(n => parseInt(n, 10) || 0); for (let i = 0; i < Math.max(x.length, y.length); i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d > 0 ? 1 : -1; } return 0; }
// v0.1.8: bộ cài mới tên "Yomi-Office-Open-Source-X.Y.Z-<edition>.<ext>"; vẫn nhận tên cũ "OfficeOpenSource-..."
const FILE_RE = /^(?:OfficeOpenSource|Yomi-Office-Open-Source)-(\d+\.\d+\.\d+)(?:-(.+?))?\.(exe|msi|zip|apk|appxbundle|appx|msix|msixbundle)$/i;

// ---------------------------------------------------------------- folder listings
// Drive: the public "embedded folder view" is plain HTML (entries: id + title; sub-folders link to /drive/folders/<id>)
export function parseDrive(html) {
  const out = []; const re = /<div class="flip-entry" id="entry-([\w-]+)"[\s\S]*?<a href="([^"]+)"[\s\S]*?<div class="flip-entry-title">([^<]*)<\/div>/g; let m;
  while ((m = re.exec(html))) { const id = m[1], href = m[2].replace(/&amp;/g, '&'), name = decodeEnt(m[3]).trim(); const folder = /\/folders\//.test(href); out.push({ name, id, folder, src: 'drive', url: folder ? null : 'https://drive.usercontent.google.com/download?id=' + id + '&export=download&confirm=t' }); }
  return out;
}
function decodeEnt(s) { return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)); }
async function listDrive(id, depth) {
  try {
    const html = await Host.httpGet('https://drive.google.com/embeddedfolderview?id=' + encodeURIComponent(id));
    let items = parseDrive(html);
    if ((depth || 0) < 2) for (const f of items.filter(x => x.folder)) { try { items = items.concat((await listDrive(f.id, (depth || 0) + 1)).map(x => Object.assign(x, { dir: f.name }))); } catch (e) {} }
    return items.filter(x => !x.folder);
  } catch (e) { return []; }
}
// MediaFire: public JSON API (files + sub-folders); the download page is resolved to its direct link right before downloading
async function listMediaFire(key, depth) {
  try {
    const base = 'https://www.mediafire.com/api/1.5/folder/get_content.php?response_format=json&folder_key=' + encodeURIComponent(key);
    const text = await Host.httpGet(base + '&content_type=files&chunk_size=1000');
    const fj = JSON.parse(text);
    let items = (((fj.response && fj.response.folder_content) || {}).files || []).map(f => ({ name: f.filename, id: f.quickkey, src: 'mediafire', page: (f.links && f.links.normal_download) || ('https://www.mediafire.com/file/' + f.quickkey), size: +f.size || 0 }));
    if ((depth || 0) < 2) {
      try { const dj = JSON.parse(await Host.httpGet(base + '&content_type=folders')); for (const d of (((dj.response && dj.response.folder_content) || {}).folders || [])) { try { items = items.concat((await listMediaFire(d.folderkey, (depth || 0) + 1)).map(x => Object.assign(x, { dir: d.name }))); } catch (e) {} } } catch (e) {}
    }
    return items;
  } catch (e) { return []; }
}
// GitHub Releases (nguồn chính, sheet "Update App Link"): bản phát hành mới nhất + các tệp đính kèm
export function parseGitHubRelease(j) {
  if (!j || !j.tag_name) return { files: [], manifest: null };
  const version = String(j.tag_name).replace(/^v/i, '');
  const files = (j.assets || []).map(a => ({ name: a.name, id: String(a.id), src: 'github', url: a.browser_download_url, size: +a.size || 0 }));
  return { files, manifest: { version, date: (j.published_at || '').slice(0, 10), notes: j.body || '', src: 'github' } };
}
async function listGitHub(repo) {
  const text = await Host.httpGet('https://api.github.com/repos/' + repo + '/releases/latest');
  return parseGitHubRelease(JSON.parse(text));
}
export function mediafireDirect(html) { const m = /href="(https?:\/\/download[^"]+mediafire\.com\/[^"]+)"/.exec(html) || /"(https?:\/\/download\d*\.mediafire\.com\/[^"]+)"/.exec(html); return m ? decodeEnt(m[1]) : null; }
async function fileText(item) {
  if (item.src === 'drive' || item.src === 'github') return Host.httpGet(item.url);
  const html = await Host.httpGet(item.page); const d = mediafireDirect(html); if (!d) throw new Error('MediaFire: không tìm thấy liên kết tải'); return Host.httpGet(d);
}
async function directUrl(item) { if (item.src === 'drive' || item.src === 'github') return item.url; const d = mediafireDirect(await Host.httpGet(item.page)); if (!d) throw new Error('MediaFire: không tìm thấy liên kết tải'); return d; }

export function srcName(src) { return src === 'github' ? 'GitHub' : src === 'drive' ? 'Google Drive' : 'MediaFire'; }
// ---------------------------------------------------------------- which installer fits this copy of the app
let _ed = null;
export async function myEdition() {
  if (_ed) return _ed;
  let tag = ''; try { tag = (await (await fetch('edition.txt', { cache: 'no-store' })).text()).trim(); } catch (e) {}
  // v0.1.8: APK theo CPU × phiên bản Android tự báo edition, vd "Android10plus-arm64-v8a" (edition.txt chỉ có theo flavor)
  if (Host.kind === 'android') { try { const i = JSON.parse(window.OOSNative.info()); if (i && i.edition && /^Android\d/.test(i.edition)) tag = i.edition; } catch (e) {} }
  let info = {}; if (Host.kind === 'windows') { try { info = await (await fetch('/api/info')).json(); } catch (e) {} }
  const ed = { tag: tag || (Host.kind === 'android' ? 'android-universal' : 'web'), platform: Host.kind, portable: !!info.portable, appx: !!info.appx, msi: !!info.msi };
  // preference order of acceptable file editions (first match wins)
  const t = ed.tag; const want = [];
  if (/^Android\d/.test(t)) { want.push(t, ''); } // v0.1.8: bản theo CPU × phiên bản Android, ví dụ Android10plus-arm64-v8a
  else if (/^android-/.test(t)) { want.push(t, 'Android6plus-universal', t === 'android-legacy' ? 'Android6-7-legacy' : t === 'android-modern' ? 'Android10plus-modern' : 'android-universal', ''); }
  else if (ed.appx) want.push('appxbundle', 'Win10-11-x64-appx');
  else if (t) { const kind = ed.portable ? 'portable' : 'Setup'; want.push(t + '-' + kind, ed.portable ? 'portable' : 'Setup'); }
  ed.want = want; _ed = ed; return ed;
}
const ANDROID_NAMES = { 'android-legacy': 'Android6-7-legacy', 'android-universal': 'Android6plus-universal', 'android-modern': 'Android10plus-modern' };
export function pickFile(files, ed, manifest) {
  // 1) version.json mapping
  if (manifest && manifest.files) {
    const keys = [ed.tag + (ed.portable ? '-portable' : ed.appx ? '-appx' : '-setup'), ed.tag, ANDROID_NAMES[ed.tag], ed.appx ? 'appxbundle' : null, ed.platform === 'android' ? 'android' : ed.portable ? 'portable' : 'setup'].filter(Boolean);
    for (const k of keys) { const fn = manifest.files[k] || manifest.files[k.toLowerCase()]; if (fn) { const f = files.find(x => x.name === fn); if (f) return f; } }
  }
  // 2) by file name
  const ver = manifest && manifest.version;
  const cand = files.map(f => { const m = FILE_RE.exec(f.name); return m ? Object.assign({ ver: m[1], ed: m[2] || '', ext: m[3].toLowerCase() }, f) : null; }).filter(f => f && (!ver || f.ver === ver));
  const isAnd = ed.platform === 'android';
  let ok = cand.filter(f => isAnd ? f.ext === 'apk' : ed.appx ? /appx|msix/.test(f.ext) : ed.portable ? f.ext === 'zip' : ed.msi ? f.ext === 'msi' : f.ext === 'exe');
  if (ed.msi && !ok.length) cand.forEach(f => { if (f.ext === 'exe') ok.push(f); }); // không có .msi thì dùng Setup .exe
  const exact = isAnd ? (ANDROID_NAMES[ed.tag] || ed.tag) : ed.tag;
  return ok.find(f => exact && f.ed.toLowerCase() === (exact + (isAnd || ed.appx || (ed.msi && f.ext === 'msi') ? '' : ed.portable ? '-portable' : '-setup')).toLowerCase())
    || ok.find(f => exact && f.ed.toLowerCase().indexOf(exact.toLowerCase()) === 0)
    || ok.find(f => isAnd ? /universal/i.test(f.ed) || !f.ed : !f.ed || /^(Setup|portable)$/i.test(f.ed))
    || ok[0] || null;
}

// ---------------------------------------------------------------- check
export async function check(opts) {
  opts = opts || {}; const c = cfg(); const errors = []; let files = [];
  const tasks = []; let ghManifest = null;
  if (c.github) tasks.push(listGitHub(c.github).then(r => { files = files.concat(r.files); ghManifest = r.manifest; }, e => errors.push('GitHub: ' + e.message)));
  if (c.drive) tasks.push(listDrive(c.drive).then(x => { files = files.concat(x); }, e => errors.push('Google Drive: ' + e.message)));
  if (c.mediafire) tasks.push(listMediaFire(c.mediafire).then(x => { files = files.concat(x); }, e => errors.push('MediaFire: ' + e.message)));
  await Promise.all(tasks);
  // newest version.json wins (Drive preferred when equal); GitHub Release (nguồn chính) xét ở dưới
  let manifest = null;
  for (const vf of files.filter(f => /^version\.json$/i.test(f.name))) { try { const m = JSON.parse(await fileText(vf)); if (m && m.version && (!manifest || cmpVer(m.version, manifest.version) > 0)) manifest = m; } catch (e) { errors.push('version.json (' + vf.src + '): ' + e.message); } }
  // GitHub Release là nguồn chính: dùng khi mới hơn hoặc bằng bản ghi trong version.json
  if (ghManifest && (!manifest || cmpVer(ghManifest.version, manifest.version) >= 0)) manifest = Object.assign({}, manifest && manifest.version === ghManifest.version ? manifest : {}, ghManifest, { notes: ghManifest.notes || (manifest && manifest.notes) || '' });
  if (!manifest) { const vs = files.map(f => (FILE_RE.exec(f.name) || [])[1]).filter(Boolean).sort(cmpVer); if (vs.length) manifest = { version: vs[vs.length - 1], notes: '' }; }
  c.last = Date.now(); saveSettings({ update: c });
  const ed = await myEdition();
  const res = { current: VERSION, manifest, files, errors, edition: ed, newer: !!(manifest && cmpVer(manifest.version, VERSION) > 0) };
  if (res.newer) { const fv = files.filter(f => (FILE_RE.exec(f.name) || [])[1] === manifest.version || (manifest.files && Object.values(manifest.files).indexOf(f.name) >= 0)); res.file = pickFile(fv.length ? fv : files, ed, manifest); res.alt = fv.filter(f => f !== res.file && f.src !== (res.file && res.file.src) && f.name === (res.file && res.file.name)); }
  return res;
}

export async function install(res, onStatus) {
  const f = res.file; if (!f) throw new Error('Không có bộ cài phù hợp cho phiên bản này (' + res.edition.tag + ')');
  const tries = [f].concat(res.alt || []); let lastErr;
  for (const it of tries) {
    try { const url = await directUrl(it); onStatus && onStatus('Đang tải ' + it.name + ' từ ' + srcName(it.src) + '…');
      return await Host.downloadInstall(url, it.name, res.edition.portable ? 'portable' : res.edition.appx ? 'appx' : '', s => { if (onStatus && s && s.total) onStatus('Đang tải ' + it.name + ': ' + Math.round(100 * s.got / s.total) + '%'); else if (onStatus && s && s.state) onStatus(s.state === 'installing' ? 'Đang cài đặt…' : 'Đang tải ' + it.name + '…'); });
    } catch (e) { lastErr = e; }
  }
  throw lastErr;
}

export function showResult(res, manual) {
  if (!res.newer) {
    if (manual) dialog({ title: 'Cập nhật phần mềm', width: 480, body: h('div', h('p', '✅ Bạn đang dùng phiên bản mới nhất: ' + VERSION + '.'), res.manifest ? h('p.muted', 'Phiên bản trên máy chủ: ' + res.manifest.version) : h('p.muted', 'Chưa có bản phát hành nào trong thư mục cập nhật.'), res.errors.length ? h('p.muted', '⚠️ ' + res.errors.join(' · ')) : null), buttons: [{ label: 'Đóng', primary: true }] });
    return;
  }
  const status = h('div.muted');
  const m = res.manifest;
  const doInstall = async () => {
    try { status.textContent = 'Đang chuẩn bị…'; const r = await install(res, t => { status.textContent = t; });
      status.textContent = r && r.manual ? 'Đã mở liên kết tải về.' : Host.kind === 'android' ? 'Đã tải xong — bấm "Cài đặt" trên màn hình Android để hoàn tất.' : 'Đang cài đặt — ứng dụng sẽ khởi động lại.'; toast(status.textContent, 'ok', 6000);
    } catch (e) { status.textContent = '⚠️ ' + e.message; toast('Cập nhật lỗi: ' + e.message, 'err', 6000); }
  };
  dialog({ title: 'Có phiên bản mới: Yomi-Office Open Source ' + m.version, width: 560,
    body: h('div', h('p', 'Phiên bản hiện tại: ' + VERSION + ' → mới: ', h('b', m.version), m.date ? ' (' + m.date + ')' : ''),
      m.notes ? h('div.upd-notes', { style: { whiteSpace: 'pre-wrap', maxHeight: '260px', overflow: 'auto', border: '1px solid var(--line)', padding: '8px', borderRadius: '6px' } }, m.notes) : null,
      h('p.muted', res.file ? 'Bộ cài phù hợp: ' + res.file.name + ' (' + srcName(res.file.src) + ')' : '⚠️ Chưa tìm thấy bộ cài cho phiên bản ' + res.edition.tag), status),
    buttons: [res.file ? { label: 'Cập nhật ngay', primary: true, keep: true, run: () => { doInstall(); return false; } } : null,
      { label: 'Bỏ qua phiên bản này', run: () => { const c = cfg(); c.skip = m.version; saveSettings({ update: c }); } }, { label: 'Để sau' }].filter(Boolean) });
  if (cfg().autoInstall && res.file && !manual) doInstall();
}

export async function checkNow(manual) {
  const t = manual ? toast('Đang kiểm tra cập nhật trên GitHub, Google Drive và MediaFire…', null, 3000) : null;
  try { const res = await check(); if (!manual && res.newer && cfg().skip === res.manifest.version) return res; showResult(res, manual); return res; }
  catch (e) { if (manual) toast('Không kiểm tra được cập nhật: ' + e.message, 'err', 5000); return null; }
}
// called once at start-up
export function startupCheck() { const c = cfg(); if (!c.auto || Host.kind === 'web' || /[#&]selftest/.test(location.hash)) return; setTimeout(() => checkNow(false), 6000); }

export function optionsSection(sec, chk) {
  const c = cfg();
  const inp = (val, onch) => { const i = h('input.inp.small', { value: val, style: { width: '260px', maxWidth: '100%' } }); i.addEventListener('change', () => onch(i.value.trim())); return i; };
  const last = h('span.muted', c.last ? ' Lần kiểm tra gần nhất: ' + new Date(c.last).toLocaleString('vi-VN') : '');
  return sec('Cập nhật phần mềm',
    h('div.form-row.inline', h('label.lbl', 'Phiên bản hiện tại: ' + VERSION), h('button.btn', { onclick: () => checkNow(true) }, ic('refresh-cw', 14), ' Kiểm tra cập nhật ngay'), last),
    chk('Tự động kiểm tra phiên bản mới mỗi khi mở ứng dụng', () => c.auto, v => { c.auto = v; saveSettings({ update: c }); }),
    chk('Tự động tải và cài đặt bản cập nhật (Android vẫn cần bấm "Cài đặt" một lần)', () => c.autoInstall, v => { c.autoInstall = v; saveSettings({ update: c }); }),
    h('div.form-row.inline', h('label.lbl', 'Kho GitHub (nguồn chính)'), inp(c.github, v => { c.github = v || DEFAULT_SOURCES.github; saveSettings({ update: c }); })),
    h('div.form-row.inline', h('label.lbl', 'Thư mục Google Drive (ID)'), inp(c.drive, v => { c.drive = v || DEFAULT_SOURCES.drive; saveSettings({ update: c }); })),
    h('div.form-row.inline', h('label.lbl', 'Thư mục MediaFire (khóa)'), inp(c.mediafire, v => { c.mediafire = v || DEFAULT_SOURCES.mediafire; saveSettings({ update: c }); })));
}
