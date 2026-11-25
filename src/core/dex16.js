// Yomi-Office Open Source v0.1.6 - Android DeX desktop shell (redesigned after the "Android Dex Screen" reference shots)
//  * ChromeOS-style wallpaper + bottom shelf (launcher, pinned apps, status tray)
//  * Quick Settings panel (notifications card, toggle tiles, volume / brightness sliders, power row)
//  * Samsung DeX-style full-screen app drawer (search, paginated 8x4 grid, Contact / Lock / Exit DeX side links)
//  * Real external apps: launchable apps from the device (icons from PackageManager), opened as freeform / split windows
import { h, ic } from './util.js';
import { toast } from './ui.js';
import Host from './host.js';
import { settings, saveSettings } from './store.js';

const CSS = `
#dx16{position:fixed;top:0;left:0;right:0;bottom:0;z-index:9000;font-family:'Segoe UI',Roboto,system-ui,sans-serif;color:#fff;overflow:hidden;user-select:none;
  background:linear-gradient(160deg,#9be7f5 0%,#4fc3f7 22%,#1e88e5 55%,#0d47a1 100%)}
#dx16 .wave{position:absolute;top:0;left:0;right:0;bottom:0;pointer-events:none;background:
  radial-gradient(120% 60% at 30% 38%,rgba(255,255,255,.55) 0,rgba(255,255,255,0) 46%),
  radial-gradient(90% 45% at 70% 30%,rgba(255,255,255,.35) 0,rgba(255,255,255,0) 50%)}
#dx16.dark{background:linear-gradient(160deg,#0b1020,#111827 60%,#0b1020)}
#dx16 .shelf{position:absolute;left:8px;right:8px;bottom:6px;height:52px;border-radius:14px;background:rgba(15,23,42,.72);backdrop-filter:blur(18px);-webkit-backdrop-filter:blur(18px);display:flex;align-items:center;padding:0 10px;gap:6px;box-shadow:0 6px 24px rgba(0,0,0,.35)}
#dx16 .sh-b{width:40px;height:40px;border-radius:50%;display:flex;align-items:center;justify-content:center;cursor:pointer;color:#e5e7eb;background:transparent;border:0}
#dx16 .sh-b:hover{background:rgba(255,255,255,.12)}
#dx16 .pins{flex:1;display:flex;justify-content:center;gap:10px}
#dx16 .pin{width:42px;height:42px;border-radius:12px;display:flex;align-items:center;justify-content:center;cursor:pointer;position:relative}
#dx16 .pin:hover{background:rgba(255,255,255,.12)}
#dx16 .pin img,#dx16 .pin .gl{width:32px;height:32px;border-radius:9px}
#dx16 .pin.run::after{content:'';position:absolute;bottom:2px;width:6px;height:3px;border-radius:2px;background:#93c5fd}
#dx16 .gl{display:flex;align-items:center;justify-content:center;font-weight:700;font-size:15px;color:#fff}
#dx16 .tray{display:flex;align-items:center;gap:6px}
#dx16 .chip{height:34px;border-radius:17px;padding:0 12px;display:flex;align-items:center;gap:8px;background:rgba(255,255,255,.08);cursor:pointer;font-size:13px}
#dx16 .chip:hover{background:rgba(255,255,255,.16)}
#dx16 .qs{position:absolute;right:12px;bottom:68px;width:360px;display:flex;flex-direction:column;gap:10px}
#dx16 .card{border-radius:18px;background:rgba(30,41,59,.88);backdrop-filter:blur(20px);-webkit-backdrop-filter:blur(20px);padding:14px 16px;box-shadow:0 10px 30px rgba(0,0,0,.4)}
#dx16 .qs-top{display:flex;align-items:center;gap:8px;margin-bottom:12px}
#dx16 .qs-top .av{width:30px;height:30px;border-radius:50%;background:linear-gradient(135deg,#38bdf8,#6366f1);display:flex;align-items:center;justify-content:center;font-weight:700}
#dx16 .pill{border-radius:16px;padding:6px 14px;background:rgba(255,255,255,.1);font-size:13px;cursor:pointer;border:0;color:#fff}
#dx16 .round{width:34px;height:34px;border-radius:50%;background:rgba(255,255,255,.1);display:flex;align-items:center;justify-content:center;cursor:pointer;border:0;color:#fff}
#dx16 .tiles{display:flex;flex-wrap:wrap;margin:6px 0 12px}
#dx16 .tiles>.tile{width:33.33%;margin-bottom:12px}
#dx16 .tile{display:flex;flex-direction:column;align-items:center;gap:6px;cursor:pointer;text-align:center;font-size:12.5px;color:#e2e8f0}
#dx16 .tile .ti{width:46px;height:46px;border-radius:50%;background:rgba(255,255,255,.12);display:flex;align-items:center;justify-content:center}
#dx16 .tile.on .ti{background:#8ab4f8;color:#0b1b3a}
#dx16 .tile small{color:#94a3b8;font-size:11px;margin-top:-4px}
#dx16 .sl{display:flex;align-items:center;gap:10px;margin:8px 0}
#dx16 .sl input{flex:1;accent-color:#8ab4f8}
#dx16 .foot{display:flex;justify-content:space-between;color:#cbd5e1;font-size:12px;margin-top:6px}
#dx16 .notif .t{font-size:12px;color:#94a3b8;display:flex;gap:6px;align-items:center}
#dx16 .notif .b{font-size:14px;margin:6px 0 2px}
#dx16 .notif .clr{text-align:right;color:#8ab4f8;font-size:13px;cursor:pointer;margin-top:8px}
#dx16 .drawer{position:absolute;top:0;left:0;right:0;bottom:64px;background:rgba(10,10,12,.92);display:flex;flex-direction:column;align-items:center}
#dx16 .drawer .srch{margin:56px 0 28px;width:min(440px,70vw);height:42px;border-radius:21px;background:rgba(255,255,255,.12);border:0;color:#fff;padding:0 18px;font-size:16px;outline:none}
#dx16 .grid{display:flex;flex-wrap:wrap;align-content:flex-start;justify-content:center;flex:1;width:100%;max-width:1240px}
#dx16 .grid>.app{width:140px;height:150px;margin:3px 7px;box-sizing:border-box}
#dx16 .app{display:flex;flex-direction:column;align-items:center;gap:10px;cursor:pointer;border-radius:16px;padding:10px 4px;font-size:14px;text-align:center;line-height:1.25}
#dx16 .app:hover{background:rgba(255,255,255,.08)}
#dx16 .app img,#dx16 .app .gl{width:64px;height:64px;border-radius:18px;font-size:26px;box-shadow:0 3px 8px rgba(0,0,0,.4)}
#dx16 .app span{max-width:130px;overflow:hidden;text-overflow:ellipsis;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
#dx16 .dots{display:flex;gap:14px;margin:6px 0 18px}
#dx16 .dots i{width:9px;height:9px;border-radius:50%;background:#6b7280;cursor:pointer}
#dx16 .dots i.on{background:#fff}
#dx16 .nav{position:absolute;top:50%;transform:translateY(-50%);width:44px;height:44px;border-radius:50%;display:flex;align-items:center;justify-content:center;cursor:pointer;color:#e5e7eb}
#dx16 .nav:hover{background:rgba(255,255,255,.1)}
#dx16 .side{position:absolute;left:22px;bottom:20px;display:flex;flex-direction:column;gap:22px;font-size:17px;font-weight:600}
#dx16 .side a{display:flex;align-items:center;gap:8px;cursor:pointer;color:#fff}
#dx16 .ctx{position:absolute;z-index:5;min-width:220px;border-radius:12px;background:rgba(30,41,59,.96);padding:6px;box-shadow:0 10px 30px rgba(0,0,0,.5)}
#dx16 .ctx div{padding:9px 12px;border-radius:8px;cursor:pointer;font-size:14px}
#dx16 .ctx div:hover{background:rgba(255,255,255,.1)}
#dx16 .night{position:absolute;top:0;left:0;right:0;bottom:0;pointer-events:none;background:rgba(255,140,0,.18);mix-blend-mode:multiply;display:none}
#dx16 .lock{position:absolute;top:0;left:0;right:0;bottom:0;background:rgba(2,6,23,.9);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;z-index:10;font-size:20px}
#dx16 .lock b{font-size:64px;font-weight:300}
@media (max-width:900px){#dx16 .grid>.app{width:22%}#dx16 .qs{width:calc(100vw - 24px)}}
`;

const INTERNAL = [
  { id: 'word', name: 'Word', color: '#185abd', letter: 'W', open: () => Host.newWindow('#new=word') },
  { id: 'excel', name: 'Excel', color: '#107c41', letter: 'X', open: () => Host.newWindow('#new=excel') },
  { id: 'ppt', name: 'PowerPoint', color: '#c43e1c', letter: 'P', open: () => Host.newWindow('#new=ppt') },
  { id: 'pdf', name: 'PDF Reader', color: '#d32f2f', letter: 'PDF', open: () => Host.newWindow('#new=pdf') },
  { id: 'math', name: 'MathType', color: '#6a1b9a', letter: '∑', open: () => Host.newWindow('#new=math') },
  { id: 'scan', name: 'CamScanner', color: '#00897b', letter: '⧉', open: () => Host.newWindow('#tool=scan') },
  { id: 'ai', name: 'Trợ lý A.I', color: '#4f46e5', letter: 'AI', open: () => Host.newWindow('#tool=ai') },
  { id: 'files', name: 'Tệp gần đây', color: '#f59e0b', letter: '📁', open: () => close() },
];

let root = null, apps = [], page = 0, perPage = 32, query = '', running = new Set();
const state = { wifi: true, bt: true, notif: true, night: false, dark: false, volume: 60, bright: 80 };

function glyph(a, size) {
  return h('div.gl', { style: { background: a.color, fontSize: size ? size + 'px' : null } }, a.letter);
}
function iconFor(a) {
  if (!a.pkg) return glyph(a);
  const img = h('img', { alt: '', loading: 'lazy' });
  const fallback = () => { const g = glyph({ color: colorOf(a.name), letter: (a.name || '?').trim().charAt(0).toUpperCase() }); img.replaceWith(g); };
  try { const src = Host.native && Host.native.getAppIcon ? Host.native.getAppIcon(a.pkg) : ''; if (src) img.src = src; else setTimeout(fallback, 0); } catch (e) { setTimeout(fallback, 0); }
  img.onerror = fallback;
  return img;
}
function colorOf(s) { let x = 0; for (const c of s || '') x = (x * 31 + c.charCodeAt(0)) | 0; const p = ['#1e88e5', '#43a047', '#e53935', '#8e24aa', '#fb8c00', '#00acc1', '#3949ab', '#6d4c41']; return p[Math.abs(x) % p.length]; }

function screenSize() { return { w: window.screen.width * (window.devicePixelRatio || 1), h: window.screen.height * (window.devicePixelRatio || 1) }; }

function launch(a, how) {
  if (!a.pkg) { a.open(); running.add(a.id); renderShelf(); return; }
  const s = screenSize(); let x = -1, y = -1, w = -1, hh = -1;
  if (how === 'left') { x = 0; y = 0; w = Math.round(s.w / 2); hh = s.h; }
  else if (how === 'right') { x = Math.round(s.w / 2); y = 0; w = Math.round(s.w / 2); hh = s.h; }
  else if (how !== 'full') { w = Math.round(s.w * 0.55); hh = Math.round(s.h * 0.7); x = Math.round((s.w - w) / 2); y = Math.round((s.h - hh) / 3); }
  const ok = Host.launchAppWithOptions(a.pkg, x, y, w, hh, 0);
  if (!ok) toast('Không mở được ' + a.name, 'err'); else { running.add(a.pkg); renderShelf(); }
}

function pinned() {
  const s = settings().dexPins;
  const ids = Array.isArray(s) && s.length ? s : ['word', 'excel', 'ppt', 'pdf', 'scan', 'com.android.chrome', 'com.google.android.youtube', 'com.google.android.gm'];
  return ids.map(id => INTERNAL.find(a => a.id === id) || apps.find(a => a.pkg === id)).filter(Boolean);
}
function togglePin(a) {
  const key = a.pkg || a.id; let s = settings().dexPins; s = Array.isArray(s) && s.length ? s.slice() : pinned().map(x => x.pkg || x.id);
  const i = s.indexOf(key); if (i >= 0) s.splice(i, 1); else s.push(key);
  saveSettings({ dexPins: s }); renderShelf();
}

function ctxMenu(ev, a) {
  ev.preventDefault(); ev.stopPropagation();
  const old = root.querySelector('.ctx'); if (old) old.remove();
  const isPinned = pinned().some(p => p === a);
  const items = [['Mở', () => launch(a)], ['Mở cửa sổ tự do', () => launch(a, 'free')], ['Mở toàn màn hình', () => launch(a, 'full')], ['Chia đôi – bên trái', () => launch(a, 'left')], ['Chia đôi – bên phải', () => launch(a, 'right')], [isPinned ? 'Bỏ ghim khỏi thanh tác vụ' : 'Ghim vào thanh tác vụ', () => togglePin(a)]];
  if (!a.pkg) items.splice(1, 4);
  const m = h('div.ctx', { style: { left: Math.min(ev.clientX, innerWidth - 240) + 'px', top: Math.min(ev.clientY, innerHeight - 260) + 'px' } }, items.map(([t, f]) => h('div', { onclick: () => { m.remove(); f(); } }, t)));
  root.appendChild(m);
  setTimeout(() => document.addEventListener('pointerdown', function o(e) { if (!m.contains(e.target)) { m.remove(); document.removeEventListener('pointerdown', o, true); } }, true), 0);
}

function renderShelf() {
  const pins = root.querySelector('.pins'); if (!pins) return;
  pins.innerHTML = '';
  pinned().forEach(a => pins.appendChild(h('div.pin' + (running.has(a.pkg || a.id) ? '.run' : ''), { title: a.name, onclick: () => launch(a), oncontextmenu: e => ctxMenu(e, a) }, iconFor(a))));
}

function clock() {
  const d = new Date();
  const t = d.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' });
  const day = d.toLocaleDateString('vi-VN', { day: 'numeric', month: 'short' });
  const full = d.toLocaleDateString('vi-VN', { weekday: 'short', day: 'numeric', month: 'short' });
  return { t, day, full };
}

async function battery() {
  try { if (navigator.getBattery) { const b = await navigator.getBattery(); return Math.round(b.level * 100) + '%' + (b.charging ? ' ⚡' : ''); } } catch (e) {}
  try { if (Host.native && Host.native.getBattery) return Host.native.getBattery() + '%'; } catch (e) {}
  return '100%';
}

function sysSettings(action) {
  try { if (Host.native && Host.native.openSystemSettings) { Host.native.openSystemSettings(action); return; } } catch (e) {}
  if (action === 'cast') Host.openSystemCastSettings(); else toast('Mở cài đặt hệ thống trên thiết bị');
}

function quickSettings() {
  const old = root.querySelector('.qs'); if (old) { old.remove(); return; }
  const c = clock();
  const tile = (key, icon, label, sub, onTap) => {
    const el = h('div.tile' + (state[key] ? '.on' : ''), { onclick: () => { state[key] = !state[key]; el.classList.toggle('on', !!state[key]); const s = el.querySelector('small'); if (s) s.textContent = sub(); onTap && onTap(state[key]); } },
      h('div.ti', ic(icon, 20)), h('span', label), h('small', sub()));
    return el;
  };
  const qs = h('div.qs',
    h('div.card.notif',
      h('div.t', ic('monitor', 14), 'Yomi-Office Open Source • DeX'),
      h('div.b', 'Không gian làm việc DeX đã sẵn sàng'),
      h('div', { style: { color: '#94a3b8', fontSize: '12px' } }, 'Kết nối màn hình ngoài / TV để mở rộng không gian làm việc'),
      h('div.clr', { onclick: e => e.target.closest('.notif').remove() }, 'Xóa tất cả')),
    h('div.card',
      h('div.qs-top', h('div.av', 'O'), h('button.pill', { onclick: () => exitDex() }, 'Thoát DeX'), h('span', { style: { flex: 1 } }),
        h('button.round', { title: 'Khóa', onclick: () => lockScreen() }, ic('lock', 16)),
        h('button.round', { title: 'Cài đặt DeX', onclick: () => import('./dex.js').then(D => D.openDexControlCenterDialog()) }, ic('settings', 16))),
      h('div.tiles',
        tile('wifi', 'wifi', 'Wi-Fi', () => state.wifi ? 'Đã bật' : 'Tắt', () => sysSettings('wifi')),
        tile('bt', 'bluetooth', 'Bluetooth', () => state.bt ? 'Bật' : 'Tắt', () => sysSettings('bluetooth')),
        tile('notif', 'bell', 'Thông báo', () => state.notif ? 'Bật, mọi ứng dụng' : 'Không làm phiền'),
        h('div.tile', { onclick: () => sysSettings('cast') }, h('div.ti', ic('cast', 20)), h('span', 'Chiếu lên TV'), h('small', 'Smart View')),
        h('div.tile', { onclick: () => import('./dex.js').then(D => D.openVirtualTrackpad()) }, h('div.ti', ic('mouse-pointer', 20)), h('span', 'Bàn rê chuột'), h('small', 'Điện thoại')),
        tile('night', 'moon', 'Ánh sáng đêm', () => state.night ? 'Bật' : 'Tắt', v => { root.querySelector('.night').style.display = v ? 'block' : 'none'; }),
        tile('dark', 'contrast', 'Giao diện tối', () => state.dark ? 'Bật' : 'Tắt', v => root.classList.toggle('dark', v))),
      h('div.sl', ic('volume-2', 18), (() => { const r = h('input', { type: 'range', min: 0, max: 100, value: state.volume }); r.oninput = () => { state.volume = +r.value; try { Host.native && Host.native.setVolume && Host.native.setVolume(state.volume); } catch (e) {} }; return r; })()),
      h('div.sl', ic('sun', 18), (() => { const r = h('input', { type: 'range', min: 5, max: 100, value: state.bright }); r.oninput = () => { state.bright = +r.value; try { Host.native && Host.native.setBrightness && Host.native.setBrightness(state.bright / 100); } catch (e) {} root.style.filter = Host.native && Host.native.setBrightness ? '' : 'brightness(' + (0.4 + state.bright / 166) + ')'; }; return r; })()),
      h('div.foot', h('span', c.full), (() => { const s = h('span', '…'); battery().then(b => s.textContent = b); return s; })())));
  root.appendChild(qs);
}

function lockScreen() {
  const c = clock();
  const l = h('div.lock', h('b', c.t), h('span', c.full), h('small', { style: { color: '#94a3b8' } }, 'Chạm để mở khóa'));
  l.onclick = () => l.remove();
  root.appendChild(l);
}

function exitDex() { close(); import('./dex.js').then(D => D.applyDexMode(false, true)).catch(() => {}); }

function filtered() {
  const all = INTERNAL.concat(apps);
  const q = query.trim().toLowerCase();
  return q ? all.filter(a => (a.name || '').toLowerCase().includes(q) || (a.pkg || '').includes(q)) : all;
}

function renderDrawer() {
  let d = root.querySelector('.drawer');
  if (!d) {
    const srch = h('input.srch', { placeholder: 'Tìm kiếm', oninput: e => { query = e.target.value; page = 0; renderDrawer(); } });
    d = h('div.drawer', srch, h('div.grid'), h('div.dots'),
      h('div.nav', { style: { left: '6%' }, onclick: () => { if (page > 0) { page--; renderDrawer(); } } }, ic('chevron-left', 30)),
      h('div.nav', { style: { right: '6%' }, onclick: () => { if ((page + 1) * perPage < filtered().length) { page++; renderDrawer(); } } }, ic('chevron-right', 30)),
      h('div.side',
        h('a', { onclick: () => Host.openUrl('mailto:?subject=Office%20Open%20Source') }, ic('mail', 20), 'Liên hệ'),
        h('a', { onclick: () => lockScreen() }, ic('lock', 20), 'Khóa DeX'),
        h('a', { onclick: () => exitDex() }, ic('log-out', 20), 'Thoát DeX')));
    d.addEventListener('wheel', e => { if (Math.abs(e.deltaY) < 20) return; const n = Math.ceil(filtered().length / perPage); page = Math.max(0, Math.min(n - 1, page + (e.deltaY > 0 ? 1 : -1))); renderDrawer(); }, { passive: true });
    let sx = null; d.addEventListener('pointerdown', e => { sx = e.clientX; }); d.addEventListener('pointerup', e => { if (sx != null && Math.abs(e.clientX - sx) > 80) { page = Math.max(0, page + (e.clientX < sx ? 1 : -1)); const n = Math.ceil(filtered().length / perPage); page = Math.min(page, n - 1); renderDrawer(); } sx = null; });
    root.insertBefore(d, root.querySelector('.shelf'));
    setTimeout(() => srch.focus(), 50);
  }
  perPage = innerWidth < 900 ? 16 : 32;
  const list = filtered(), pages = Math.max(1, Math.ceil(list.length / perPage));
  const g = d.querySelector('.grid'); g.innerHTML = '';
  list.slice(page * perPage, page * perPage + perPage).forEach(a => g.appendChild(h('div.app', { onclick: () => { launch(a); toggleDrawer(false); }, oncontextmenu: e => ctxMenu(e, a) }, iconFor(a), h('span', a.name))));
  const dots = d.querySelector('.dots'); dots.innerHTML = '';
  for (let i = 0; i < pages; i++) dots.appendChild(h('i' + (i === page ? '.on' : ''), { onclick: () => { page = i; renderDrawer(); } }));
}
function toggleDrawer(force) {
  const d = root.querySelector('.drawer');
  if (d && force !== true) { d.remove(); return; }
  if (!d && force !== false) { query = ''; page = 0; renderDrawer(); }
}

export function close() { if (root) { clearInterval(root._t); root.remove(); root = null; } }

export async function openDesktop(options = {}) {
  close();
  if (!document.getElementById('dx16-css')) document.head.appendChild(h('style#dx16-css', CSS));
  root = h('div#dx16', h('div.wave'), h('div.night'));
  const c = clock();
  const timeChip = h('div.chip', { onclick: () => quickSettings() }, h('span.tm', c.t), ic('wifi', 15), ic('battery', 15));
  const dayChip = h('div.chip', { onclick: () => quickSettings() }, h('span.dy', c.day));
  root.appendChild(h('div.shelf',
    h('button.sh-b', { title: 'Ứng dụng', onclick: () => toggleDrawer() }, ic('grid-3x3', 20)),
    h('button.sh-b', { title: 'Tìm kiếm', onclick: () => toggleDrawer(true) }, ic('search', 19)),
    h('div.pins'),
    h('div.tray',
      h('button.sh-b', { title: 'Bàn rê chuột ảo', onclick: () => import('./dex.js').then(D => D.openVirtualTrackpad()) }, ic('mouse-pointer', 17)),
      h('button.sh-b', { title: 'Chiếu lên TV', onclick: () => sysSettings('cast') }, ic('cast', 17)),
      dayChip, timeChip)));
  root.addEventListener('contextmenu', e => { if (e.target === root || e.target.classList.contains('wave')) { e.preventDefault(); const m = h('div.ctx', { style: { left: e.clientX + 'px', top: e.clientY + 'px' } }, h('div', { onclick: () => { m.remove(); root.classList.toggle('dark'); } }, 'Đổi hình nền sáng / tối'), h('div', { onclick: () => { m.remove(); import('./dex.js').then(D => D.openDexControlCenterDialog()); } }, 'Cài đặt DeX & màn hình')); root.appendChild(m); setTimeout(() => document.addEventListener('pointerdown', function o(ev) { if (!m.contains(ev.target)) { m.remove(); document.removeEventListener('pointerdown', o, true); } }, true), 0); } });
  root.addEventListener('keydown', e => { if (e.key === 'Escape') { const d = root.querySelector('.drawer, .qs'); if (d) d.remove(); } });
  document.body.appendChild(root);
  root._t = setInterval(() => { const k = clock(); const t = root && root.querySelector('.tm'); if (t) t.textContent = k.t; const d = root && root.querySelector('.dy'); if (d) d.textContent = k.day; }, 15000);
  try { apps = (await Host.getInstalledApps()).map(a => ({ pkg: a.packageName, name: a.name, isSystem: a.isSystem })); } catch (e) { apps = []; }
  renderShelf();
  if (options.mode === 'tvapps') toggleDrawer(true);
  if (options.mode === 'multiwin') quickSettings();
}
