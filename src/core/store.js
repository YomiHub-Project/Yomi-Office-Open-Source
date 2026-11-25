// Persistent storage: settings (localStorage + disk/native) + recent documents & autosave snapshots (IndexedDB)
import { safeStore, uid } from './util.js';
import Host from './host.js';

const SETTINGS_KEY = 'oos5_settings';
const DEFAULTS = {
  userName: 'Người dùng', initials: 'ND', theme: 'light', lang: 'vi', autosave: true, autosaveSec: 30,
  defaultFont: 'Calibri', defaultSize: 11, wordFont: 'Calibri', wordSize: 11, excelFont: 'Calibri', excelSize: 11,
  showStart: true, ribbonCollapsed: false, glass: true, rulers: true, gridlines: true, cursorBlink: true,
  antiObscuration: true, safePadTop: 0, safePadBottom: 0,
  mouse: { smoothScroll: true, wheelZoomCtrl: true, rightClickMenus: true, dragDrop: true, doubleClickEdit: true, cursorStyle: 'default', scrollLines: 3, miniToolbar: true, hoverPreview: true },
  dex: {
    enabled: false,
    resolution: 'auto',
    uiScale: 100,
    multiWindow: true,
    dualScreenPresenter: true,
    wakeLock: true,
    audioRoute: 'auto',
    keyboardProfile: 'windows',
    trackpadSensitivity: 1.0,
    smoothScroll: true,
    autoLaunchOnHdmi: true,
    webLinkPort: 8080,
    webLinkEnabled: false,
    scrcpyBitrate: 12,
    scrcpyFps: 60,
    scrcpyScreenOff: true
  },
  recentMax: 30
};
let cache = null;
export function settings() {
  if (!cache) {
    const s = safeStore(SETTINGS_KEY) || {};
    cache = Object.assign({}, DEFAULTS, s);
    cache.mouse = Object.assign({}, DEFAULTS.mouse, s.mouse || {});
    cache.dex = Object.assign({}, DEFAULTS.dex, s.dex || {});
    if (s.dexForceDesktop && !cache.dex.enabled) cache.dex.enabled = true;
  }
  return cache;
}
export async function syncNativeSettings() {
  try {
    const native = await Host.loadConfig();
    if (native && typeof native === 'object') {
      Object.assign(settings(), native);
      safeStore(SETTINGS_KEY, cache);
    }
  } catch (e) {}
}
export function saveSettings(patch) {
  Object.assign(settings(), patch || {});
  safeStore(SETTINGS_KEY, cache);
  try { Host.saveConfig(cache); } catch (e) {}
}

// ---- IndexedDB (available in Chrome 24+/WebView 44); falls back to memory
let dbp = null;
function db() {
  if (dbp) return dbp;
  dbp = new Promise(res => {
    try {
      const rq = indexedDB.open('oos5', 1);
      rq.onupgradeneeded = () => { const d = rq.result; if (!d.objectStoreNames.contains('docs')) d.createObjectStore('docs', { keyPath: 'id' }); if (!d.objectStoreNames.contains('meta')) d.createObjectStore('meta', { keyPath: 'id' }); };
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => res(null);
    } catch (e) { res(null); }
  });
  return dbp;
}
const mem = { docs: {}, meta: {} };
async function tx(store, mode, fn) {
  const d = await db();
  if (!d) return fn(null, mem[store]);
  return new Promise((res, rej) => {
    const t = d.transaction(store, mode); const s = t.objectStore(store);
    let out; const r = fn(s); if (r && 'onsuccess' in r) r.onsuccess = () => { out = r.result; };
    t.oncomplete = () => res(out); t.onerror = () => rej(t.error);
  });
}
export async function idbPut(store, obj) { return tx(store, 'readwrite', (s, m) => { if (m) { m[obj.id] = obj; return; } return s.put(obj); }); }
export async function idbGet(store, id) { return tx(store, 'readonly', (s, m) => m ? m[id] : s.get(id)); }
export async function idbDel(store, id) { return tx(store, 'readwrite', (s, m) => { if (m) { delete m[id]; return; } return s.delete(id); }); }
export async function idbAll(store) {
  const d = await db();
  if (!d) return Object.keys(mem[store]).map(k => mem[store][k]);
  return new Promise(res => { const out = []; const rq = d.transaction(store, 'readonly').objectStore(store).openCursor(); rq.onsuccess = () => { const c = rq.result; if (c) { out.push(c.value); c.continue(); } else res(out); }; rq.onerror = () => res(out); });
}

// ---- recent documents  (meta: {id, name, kind, path, ts, pinned, size})
export async function recentList() { const all = await idbAll('meta'); return all.sort((a, b) => ((b.pinned || 0) - (a.pinned || 0)) || (b.ts - a.ts)); }
export async function recentTouch(rec, snapshot) {
  rec.id = rec.id || uid('doc'); rec.ts = Date.now();
  const old = await idbGet('meta', rec.id); rec.pinned = old ? (old.pinned || 0) : (rec.pinned || 0);
  await idbPut('meta', rec);
  if (snapshot) await idbPut('docs', { id: rec.id, data: snapshot, ts: rec.ts });
  // trim
  const list = await recentList(); const max = settings().recentMax || 30;
  for (const r of list.filter(r => !r.pinned).slice(max)) { await idbDel('meta', r.id); await idbDel('docs', r.id); }
  return rec;
}
export async function recentRemove(id) { await idbDel('meta', id); await idbDel('docs', id); }
export async function recentPin(id, v) { const r = await idbGet('meta', id); if (r) { r.pinned = v ? 1 : 0; await idbPut('meta', r); } }
export async function snapshotGet(id) { const r = await idbGet('docs', id); return r ? r.data : null; }
