// Platform bridge: Android (native WebView shell), Windows (launcher with local server), plain web
import { b64FromBytes, bytesFromB64, extOf, mimeOf, loadScript, supportsWasm, h, utf8Encode, utf8Decode, VERSION } from './util.js';
import { toast } from './ui.js';
import PROFILE from './profile.gen.js';

const N = window.OOSNative; // injected by Android shell (addJavascriptInterface)
let cbSeq = 0; const cbs = {};
window.__oosCb = function (id, json) { const f = cbs[id]; delete cbs[id]; if (f) { let v = null; try { v = json ? JSON.parse(json) : null; } catch (e) { v = json; } f(v); } };
function nativeCall(method, args) {
  return new Promise(res => { const id = 'c' + (++cbSeq); cbs[id] = res; N[method].apply(N, args.concat([id])); });
}
function isWin() { return location.protocol === 'http:' && /^(127\.0\.0\.1|localhost)$/.test(location.hostname) && /[?&]host=win/.test(location.search + '&' + (sessionStorage.getItem('oos_host') || '')); }
if (/[?&]host=win/.test(location.search)) try { sessionStorage.setItem('oos_host', 'host=win'); } catch (e) {}

async function api(path, body) {
  const r = await fetch(path, body ? { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) } : {});
  if (!r.ok) throw new Error('HTTP ' + r.status);
  return r.json();
}

function pickWithInput(accept, multiple) {
  return new Promise(res => {
    const inp = h('input', { type: 'file', accept: accept || '', style: { position: 'fixed', left: '-1000px', top: '0' } });
    if (multiple) inp.multiple = true;
    document.body.appendChild(inp);
    let done = false;
    inp.addEventListener('change', async () => {
      done = true;
      const files = Array.prototype.slice.call(inp.files || []);
      const out = [];
      for (const f of files) {
        const bytes = await new Promise((ok, no) => { const r = new FileReader(); r.onload = () => ok(new Uint8Array(r.result)); r.onerror = no; r.readAsArrayBuffer(f); });
        out.push({ name: f.name, bytes, size: f.size });
      }
      if (inp.parentNode) inp.parentNode.removeChild(inp);
      res(out);
    });
    // detect cancel
    window.addEventListener('focus', function onF() { window.removeEventListener('focus', onF); setTimeout(() => { if (!done) { if (inp.parentNode) inp.parentNode.removeChild(inp); res([]); } }, 800); });
    inp.click();
  });
}

// Tesseract 4.1 compiled to asm.js (tesseract.js v2) — runs on any JS engine, no WebAssembly needed
let _asmW = null, _asmLang = null;
async function ocrAsm(dataUrl, lang, onProgress) {
  const base = location.href.replace(/[#?].*$/, '').replace(/[^/]*$/, '');
  if (!window.TesseractAsm) { const prev = window.Tesseract; await loadScript('vendor/tesseract-asm/tesseract.min.js'); window.TesseractAsm = window.Tesseract; window.Tesseract = prev; }
  if (!_asmW || _asmLang !== lang) {
    if (_asmW) try { await _asmW.terminate(); } catch (e) {}
    _asmW = window.TesseractAsm.createWorker({ workerPath: base + 'vendor/tesseract-asm/worker-compat.js', corePath: base + 'vendor/tesseract-asm/tesseract-core.asm.js', langPath: base + 'vendor/tesseract/lang', gzip: false, cacheMethod: 'none', workerBlobURL: false,
      logger: m => { if (onProgress && m && m.progress != null) onProgress(m.status, m.progress); } });
    await _asmW.load(); await _asmW.loadLanguage(lang); await _asmW.initialize(lang); _asmLang = lang;
  }
  // Chrome 44 cannot fetch() data: URLs and lacks canvas.toBlob → hand Tesseract a File (read with FileReader)
  const m = /^data:([^;,]+)?(;base64)?,/.exec(dataUrl) || [];
  const file = new File([bytesFromB64(dataUrl.slice(dataUrl.indexOf(',') + 1))], 'page.png', { type: m[1] || 'image/png' });
  const r = await _asmW.recognize(file);
  return { text: r.data.text, words: (r.data.words || []).map(w => ({ text: w.text, bbox: w.bbox, conf: w.confidence })), engine: 'asm.js' };
}

export const Host = {
  kind: N ? 'android' : isWin() ? 'windows' : 'web',
  get native() { return N; },
  info() {
    if (N && N.info) try { return JSON.parse(N.info()); } catch (e) {}
    return { platform: this.kind, ua: navigator.userAgent };
  },
  async openFiles(accept, multiple) {
    if (this.kind === 'windows') {
      try { const r = await api('/api/open', { filter: accept || '', multiple: !!multiple }); return (r.files || []).map(f => ({ name: f.name, path: f.path, bytes: bytesFromB64(f.b64) })); }
      catch (e) { console.warn(e); }
    }
    return pickWithInput(accept, multiple);
  },
  // returns {name, path} or null when cancelled
  async saveFile(name, bytes, opts) {
    opts = opts || {};
    const mime = opts.mime || mimeOf(extOf(name));
    if (typeof bytes === 'string') bytes = utf8Encode(bytes);
    const b64 = b64FromBytes(bytes);
    if (N && N.saveFile) {
      if (opts.path && !opts.saveAs && N.writeUri) { const ok = await nativeCall('writeUri', [opts.path, b64]); if (ok && ok.ok) return { name: ok.name || name, path: opts.path }; }
      const r = await nativeCall('saveFile', [name, mime, b64]);
      return r && r.ok ? { name: r.name || name, path: r.uri } : null;
    }
    if (this.kind === 'windows') {
      try { const r = await api('/api/save', { name, path: opts.saveAs ? '' : (opts.path || ''), b64, saveAs: !!opts.saveAs || !opts.path }); return r.ok ? { name: r.name, path: r.path } : null; }
      catch (e) { console.warn(e); }
    }
    const blob = new Blob([bytes], { type: mime });
    const a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.parentNode && a.parentNode.removeChild(a); }, 1500);
    return { name, path: null };
  },
  // v0.1.8: quét tệp Word/Excel/PowerPoint trên máy (Feature Check #28). null = host chưa hỗ trợ → devscan.js tự duyệt thư mục.
  async scanOffice() {
    if (N && N.scanOffice) { const r = await nativeCall('scanOffice', []); this.scanFull = !!(r && r.full); return r && r.ok ? r.files : null; }
    if (this.kind === 'windows') { try { const r = await api('/api/scan_office', {}); if (r && r.ok) return r.files; } catch (e) {} }
    return null;
  },
  requestFilesAccess() { if (N && N.requestFilesAccess) N.requestFilesAccess(); },
  async readPath(path) {
    if (!path) return null;
    if (N && N.readUri) { const r = await nativeCall('readUri', [path]); return r && r.ok ? { name: r.name, bytes: bytesFromB64(r.b64), path } : null; }
    if (this.kind === 'windows') { const r = await api('/api/read', { path }); return r.ok ? { name: r.name, bytes: bytesFromB64(r.b64), path } : null; }
    return null;
  },
  async pending() {
    try {
      if (N && N.takePendingFile) { const s = N.takePendingFile(); if (s) { const o = JSON.parse(s); return [{ name: o.name, bytes: bytesFromB64(o.b64), path: o.uri }]; } }
      if (this.kind === 'windows') { const r = await api('/api/pending'); return (r.files || []).filter(f => { if (f.act) { import('./bstage.js').then(B => B.optAction(f.act, window.__oosShell)); return false; } return true; }).map(f => { if (f.cmd) window.__oosPendingCmd = { name: f.name, id: f.cmd }; return { name: f.name, path: f.path, bytes: bytesFromB64(f.b64), forceWeb: !!f.forceWeb }; }); }
    } catch (e) { console.warn(e); }
    return [];
  },
  print(title) {
    if (N && N.print) { N.print(title || document.title); return; }
    window.print();
  },
  share(name, bytes, mime) {
    if (N && N.share) { N.share(name, mime || mimeOf(extOf(name)), b64FromBytes(bytes)); return true; }
    return false;
  },
  openUrl(url) {
    if (N && N.openUrl) { N.openUrl(url); return; }
    if (this.kind === 'windows') { api('/api/openurl', { url }).catch(() => window.open(url, '_blank')); return; }
    window.open(url, '_blank');
  },
  // v0.1.6: AirDrop / QuickShare over LAN (port 8765)
  async lanShare(name, bytes) {
    const b64 = b64FromBytes(bytes || new Uint8Array(0));
    if (N && N.lanShare) return JSON.parse(N.lanShare(name, b64));
    if (this.kind === 'windows') return api('/api/lan_share', { name, b64 });
    return null;
  },
  async lanReceived() {
    if (N && N.lanReceived) return JSON.parse(N.lanReceived() || '[]');
    if (this.kind === 'windows') { const r = await api('/api/lan_received'); return r.files || []; }
    return [];
  },
  // v0.1.6: native document engine (embedded LibreOffice core, rendered without WebView)
  _lok: null,
  async lokAvailable() {
    if (this._lok != null) return this._lok;
    try {
      if (N && N.lokAvailable) this._lok = !!N.lokAvailable();
      else if (this.kind === 'windows') { const r = await api('/api/lok_info'); this._lok = !!(r && r.available); }
      else this._lok = false;
    } catch (e) { this._lok = false; }
    return this._lok;
  },
  async lokOpen(f) {
    const b64 = f.path ? '' : (f.bytes ? b64FromBytes(f.bytes) : '');
    if (N && N.lokOpen) return !!N.lokOpen(f.path || '', f.name, b64);
    if (this.kind === 'windows') { try { const r = await api('/api/lok_open', { path: f.path || '', name: f.name, b64, cmd: f.cmd || '' }); return !!(r && r.ok); } catch (e) { return false; } }
    return false;
  },
  async lokConvert(f, fmt) {
    if (this.kind !== 'windows') return null;
    const r = await api('/api/lok_convert', { path: f.path || '', name: f.name, b64: f.path ? '' : b64FromBytes(f.bytes), fmt });
    return r && r.ok ? { name: r.name, bytes: bytesFromB64(r.b64) } : null;
  },
  // v0.0.6: open a file in a brand-new window (Windows: new app window; Android: new task in Recents; web: new tab)
  canNewWindow() { return !!((N && N.openWindow) || this.kind === 'windows' || this.kind === 'web'); },
  async openInNewWindow(f) {
    const b64 = f.bytes ? b64FromBytes(f.bytes) : '';
    if (N && N.openWindow) { N.openWindow(JSON.stringify({ name: f.name, b64, uri: f.path || '' }), ''); return true; }
    if (this.kind === 'windows') { try { const r = await api('/api/stash', { name: f.name, path: f.path || '', b64: f.path ? '' : b64 }); return !!r.ok; } catch (e) { console.warn(e); return false; } }
    try { const { idbPut } = await import('./store.js'); const tok = 'stash_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); await idbPut('docs', { id: tok, data: { name: f.name, b64, path: f.path || null }, ts: Date.now() }); const w = window.open(location.pathname + location.search + '#open=' + tok, '_blank'); return !!w; } catch (e) { return false; }
  },
  async takeStash(tok) {
    try {
      if (this.kind === 'windows') { const r = await api('/api/unstash', { token: tok }); return (r.files || []).map(f => ({ name: f.name, path: f.path || null, bytes: bytesFromB64(f.b64) })); }
      const { idbGet, idbDel } = await import('./store.js'); const r = await idbGet('docs', tok); if (!r) return []; await idbDel('docs', tok); return [{ name: r.data.name, path: r.data.path, bytes: bytesFromB64(r.data.b64) }];
    } catch (e) { console.warn(e); return []; }
  },
  newWindow(hash) {
    if (N && N.openWindow) { N.openWindow('', (hash || '').replace(/^#/, '')); return; }
    if (this.kind === 'windows') { api('/api/newwindow', { hash: hash || '' }).catch(() => window.open(location.pathname + location.search + (hash || ''), '_blank')); return; }
    if (this.kind === 'web') window.open(location.pathname + location.search + (hash || ''), '_blank');
  },
  // OCR: Android native (Tesseract JNI) works on every WebView; elsewhere Tesseract.js (WASM)
  async ocr(dataUrl, lang, onProgress) {
    lang = lang || 'vie+eng';
    if (N && N.ocr && !supportsWasm()) {
      const b64 = dataUrl.slice(dataUrl.indexOf(',') + 1);
      const r = await nativeCall('ocr', [b64, lang]);
      if (r && r.ok) return { text: r.text, words: r.words || [] };
      throw new Error((r && r.error) || 'OCR lỗi');
    }
    if (PROFILE.asm && (!supportsWasm() || !PROFILE.wasm)) return ocrAsm(dataUrl, lang, onProgress); // old WebViews (Android 6 stock, Chrome < 57) / legacy build: asm.js Tesseract, same models
    if (!supportsWasm()) throw new Error('WebView này không hỗ trợ WebAssembly — hãy cập nhật Android System WebView hoặc dùng bản cài Android Legacy/Universal');
    if (!window.Tesseract) await loadScript('vendor/tesseract/tesseract.min.js');
    if (!Host._tw || Host._twLang !== lang) {
      if (Host._tw) try { await Host._tw.terminate(); } catch (e) {}
      const base = location.href.replace(/[#?].*$/, '').replace(/[^/]*$/, '');
      Host._tw = await window.Tesseract.createWorker(lang.split('+'), 1, {
        workerPath: base + 'vendor/tesseract/worker.min.js', corePath: base + 'vendor/tesseract/core', langPath: base + 'vendor/tesseract/lang', gzip: false,
        workerBlobURL: false, cacheMethod: 'none',
        logger: m => { if (onProgress && m && m.progress != null) onProgress(m.status, m.progress); }
      });
      Host._twLang = lang;
    }
    const r = await Host._tw.recognize(dataUrl);
    return { text: r.data.text, words: (r.data.words || []).map(w => ({ text: w.text, bbox: w.bbox, conf: w.confidence })) };
  },
  // ---------------- v0.0.7: A.I web window, clipboard, HTTP (no CORS), self-update
  async aiWeb(url, title) {
    if (N && N.aiWeb) { N.aiWeb(url, title || 'A.I'); return true; }
    if (this.kind === 'windows') { try { const r = await api('/api/aiweb', { url, title: title || '' }); if (r.ok) return true; } catch (e) {} }
    const w = window.open(url, 'oos_ai', 'width=520,height=820'); if (!w) window.open(url, '_blank'); return true;
  },
  async clipText() {
    if (N && N.clipText) { try { return N.clipText() || ''; } catch (e) {} }
    if (this.kind === 'windows') { try { const r = await api('/api/clipboard', {}); if (r.ok) return r.text || ''; } catch (e) {} }
    try { return await navigator.clipboard.readText(); } catch (e) { return ''; }
  },
  clipSet(t) { if (N && N.clipSet) try { N.clipSet(t); } catch (e) {} else if (this.kind === 'windows') api('/api/clipboard', { set: t }).catch(() => {}); },
  // GET a URL as text (Drive / MediaFire listings for the updater) — native side avoids CORS
  async httpGet(url) {
    if (N && N.httpGet) { const r = await nativeCall('httpGet', [url]); if (r && r.ok) return r.body; throw new Error((r && r.error) || 'HTTP lỗi'); }
    if (this.kind === 'windows') { const r = await api('/api/httpget', { url }); if (r.ok) return r.body; throw new Error(r.error || 'HTTP lỗi'); }
    const r = await fetch(url); if (!r.ok) throw new Error('HTTP ' + r.status); return r.text();
  },
  // download an installer and start it (Windows: silent Setup / portable in-place / appx via App Installer; Android: package installer)
  async downloadInstall(url, name, mode, onProgress) {
    if (N && N.downloadInstall) { const r = await nativeCall('downloadInstall', [url, name]); if (r && r.ok) return r; throw new Error((r && r.error) || 'Không tải được'); }
    if (this.kind === 'windows') {
      const r = await api('/api/update', { url, name, mode: mode || '' }); if (!r.ok) throw new Error(r.error || 'Không tải được');
      for (let i = 0; i < 1800; i++) { await new Promise(ok => setTimeout(ok, 1000)); const s = await api('/api/updatestatus'); if (onProgress) onProgress(s); if (s.state === 'done' || s.state === 'error') { if (s.state === 'error') throw new Error(s.error); return s; } }
      throw new Error('Quá thời gian');
    }
    window.open(url, '_blank'); return { ok: true, manual: true };
  },
  // ---------------- v0.0.8: generic HTTP (cloud APIs; no CORS, no Origin header) and OAuth loopback receiver (rclone style)
  // req: {url, method, headers, body (text) | bytes (Uint8Array), binary} -> {status, text | bytes, type}
  async http(req) {
    const bodyB64 = req.bytes ? b64FromBytes(req.bytes) : req.body != null ? b64FromBytes(utf8Encode(req.body)) : '';
    const done = r => { const out = { status: r.status || 0, type: r.type || '' }; const b = r.b64 ? bytesFromB64(r.b64) : new Uint8Array(0); if (req.binary) out.bytes = b; out.text = req.binary ? null : utf8Decode(b); return out; };
    if (N && N.http) { const r = await nativeCall('http', [JSON.stringify({ url: req.url, method: req.method || 'GET', headers: req.headers || {}, b64: bodyB64 })]); if (!r || r.error) throw new Error((r && r.error) || 'HTTP lỗi'); return done(r); }
    if (this.kind === 'windows') { const r = await api('/api/http', { url: req.url, method: req.method || 'GET', headers: req.headers || {}, b64: bodyB64 }); if (r.error) throw new Error(r.error); return done(r); }
    const rr = await fetch(req.url, { method: req.method || 'GET', headers: req.headers || {}, body: req.bytes || req.body || undefined });
    const ab = new Uint8Array(await rr.arrayBuffer()); return done({ status: rr.status, type: rr.headers.get('content-type'), b64: b64FromBytes(ab) });
  },
  // open the system browser on an OAuth URL and wait for the redirect to http://127.0.0.1:<port>/ ; resolves with the query string
  async oauth(url, port) {
    if (N && N.oauth) { const r = await nativeCall('oauth', [url, String(port)]); if (!r || !r.ok) throw new Error((r && r.error) || 'Hết thời gian chờ đăng nhập'); return r.query; }
    if (this.kind === 'windows') {
      const s = await api('/api/oauth', { url, port }); if (!s.ok) throw new Error(s.error || 'Không mở được cổng ' + port);
      for (let i = 0; i < 600; i++) { await new Promise(ok => setTimeout(ok, 1000)); const p = await api('/api/oauthpoll', {}); if (p.state === 'done') return p.query; if (p.state === 'error') throw new Error(p.error || 'Lỗi'); }
      throw new Error('Hết thời gian chờ đăng nhập (10 phút)');
    }
    throw new Error('Cần bản cài đặt Windows/Android');
  },
  edition() {
    if (N && N.info) try { const i = JSON.parse(N.info()); return { platform: 'android', sdk: i.sdk, abi: i.abi, flavor: (i.version || '').split('-')[1] || PROFILE.profile }; } catch (e) {}
    return { platform: this.kind, profile: PROFILE.profile };
  },
  async loadConfig() {
    if (N && N.loadSettings) {
      try { const s = N.loadSettings(); return s ? JSON.parse(s) : null; } catch (e) {}
    }
    if (this.kind === 'windows') {
      try { const r = await api('/api/settings_get'); return r && r.ok && r.settings ? JSON.parse(r.settings) : null; } catch (e) {}
    }
    return null;
  },
  saveConfig(data) {
    const s = typeof data === 'string' ? data : JSON.stringify(data);
    if (N && N.saveSettings) try { N.saveSettings(s); } catch (e) {}
    else if (this.kind === 'windows') api('/api/settings_save', { settings: s }).catch(() => {});
  },
  async lanInfo() {
    if (this.kind === 'windows') {
      try { return await api('/api/lan_info'); } catch (e) {}
    }
    return { ok: true, port: location.port || 80, ips: [location.hostname || '127.0.0.1'] };
  },
  isSamsungDeX() {
    if (N && N.isSamsungDeX) {
      try { return !!N.isSamsungDeX(); } catch (e) {}
    }
    return false;
  },
  async getDisplayInfo() {
    if (N && N.getDisplayInfo) {
      try { return JSON.parse(N.getDisplayInfo()); } catch (e) {}
    }
    return {
      isDex: false,
      displays: [{
        id: 0,
        name: 'Màn hình chính',
        width: window.screen.width,
        height: window.screen.height,
        densityDpi: Math.round((window.devicePixelRatio || 1) * 160),
        refreshRate: 60
      }]
    };
  },
  async getDeviceNetworkInfo() {
    if (N && N.getDeviceNetworkInfo) {
      try { return JSON.parse(N.getDeviceNetworkInfo()); } catch (e) {}
    }
    if (this.kind === 'windows') {
      try {
        const lan = await this.lanInfo();
        if (lan && lan.ips && lan.ips.length) return { ip: lan.ips[0] };
      } catch (e) {}
    }
    return { ip: location.hostname || '127.0.0.1' };
  },
  setKeepScreenOn(keep) {
    if (N && N.setKeepScreenOn) {
      try { N.setKeepScreenOn(!!keep); } catch (e) {}
    }
    if ('wakeLock' in navigator) {
      if (keep) {
        navigator.wakeLock.request('screen').then(lock => { Host._wakeLock = lock; }).catch(() => {});
      } else if (Host._wakeLock) {
        Host._wakeLock.release().catch(() => {});
        Host._wakeLock = null;
      }
    }
  },
  openSystemCastSettings() {
    if (N && N.openSystemCastSettings) {
      try { N.openSystemCastSettings(); return true; } catch (e) {}
    }
    return false;
  },
  async startWebLink(port) {
    if (N && N.startWebLink) {
      try { return JSON.parse(N.startWebLink(port || 8080)); } catch (e) {}
    }
    return { ok: true, port: port || 8080, running: true };
  },
  async stopWebLink() {
    if (N && N.stopWebLink) {
      try { return JSON.parse(N.stopWebLink()); } catch (e) {}
    }
    return { ok: true, running: false };
  },
  async getWebLinkStatus() {
    if (N && N.getWebLinkStatus) {
      try { return JSON.parse(N.getWebLinkStatus()); } catch (e) {}
    }
    return { running: false, port: 8080 };
  },
  async getInstalledApps() {
    if (N && N.getInstalledApps) {
      try { return JSON.parse(N.getInstalledApps()); } catch (e) {}
    }
    return [
      { packageName: 'com.android.chrome', name: 'Trình duyệt Web (Chrome / Edge)', isSystem: true, icon: 'globe' },
      { packageName: 'com.android.documentsui', name: 'Quản lý Tệp tin (Files)', isSystem: true, icon: 'folder' },
      { packageName: 'com.google.android.calculator', name: 'Máy tính Bỏ túi (Calculator)', isSystem: true, icon: 'calculator' },
      { packageName: 'com.google.android.apps.photos', name: 'Thư viện Ảnh (Gallery)', isSystem: true, icon: 'image' },
      { packageName: 'com.android.settings', name: 'Cài đặt Hệ thống (Settings)', isSystem: true, icon: 'settings' },
      { packageName: 'com.android.camera2', name: 'Máy ảnh (Camera)', isSystem: true, icon: 'camera' },
      { packageName: 'com.officeopensource.office_open_source', name: 'Yomi-Office Open Source v' + VERSION, isSystem: false, icon: 'file-text' }
    ];
  },
  launchApp(pkg) {
    if (N && N.launchApp) {
      try { return N.launchApp(pkg); } catch (e) {}
    }
    return false;
  },
  launchAppWithOptions(pkg, x, y, w, h, displayId) {
    if (N && N.launchAppWithOptions) {
      const n = v => (v == null || isNaN(v) ? -1 : Math.round(v)); // v0.1.6 fix: x = 0 (left half) was sent as -1
      try { return N.launchAppWithOptions(pkg, n(x), n(y), n(w), n(h), displayId != null ? displayId : -1); } catch (e) {}
    }
    return this.launchApp(pkg);
  },
  // v0.1.7: Backstage "Lưu thành" folder browser + "Mở vị trí tệp" (same as the native editor; Windows only)
  async listDir(dir) { if (this.kind !== 'windows') return null; try { return await api('/api/list_dir', { dir: dir || '' }); } catch (e) { return null; } },
  async pickFolder() { if (this.kind !== 'windows') return null; try { const r = await api('/api/pick_folder', {}); return r && r.ok ? r.path : null; } catch (e) { return null; } },
  async fileInfo(path) { if (this.kind !== 'windows' || !path) return null; try { const r = await api('/api/file_info', { path }); return r && r.ok ? r : null; } catch (e) { return null; } },
  reveal(path) { if (this.kind === 'windows' && path) api('/api/reveal', { path }).catch(() => {}); },
  ping() { if (this.kind === 'windows') fetch('/api/ping').catch(() => {}); },
  setTitle(t) { document.title = t; if (N && N.setTitle) try { N.setTitle(t); } catch (e) {} },
  exit() { if (N && N.exit) N.exit(); else if (this.kind === 'windows') { api('/api/exit', {}).catch(() => {}); window.close(); } else window.close(); }
};
if (Host.kind === 'windows') { setInterval(() => Host.ping(), 5000); Host.ping(); window.addEventListener('beforeunload', () => { try { navigator.sendBeacon && navigator.sendBeacon('/api/closing', '1'); } catch (e) {} }); }
window.OOS_Host = Host;
export default Host;
