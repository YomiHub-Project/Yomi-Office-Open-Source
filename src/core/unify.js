// v0.1.7: ONE ribbon layout for the Web editor and the native (LibreOfficeKit) editor.
// ui_layout.json is generated from both ribbons (tools/ui_layout_gen.py): same tabs, groups, order, labels and icons.
// Every item is bound to the Web implementation when the Web editor has it; otherwise the command runs in the native
// editor (Windows) — and the native editor hands Web-only commands to the Web editor the same way.
import L from './ui_layout.json';
import Host from './host.js';
import { toast } from './ui.js';
import { settings } from './store.js';

const KIND = { word: 'word', excel: 'excel', ppt: 'ppt' };
const wkey = it => (it.label || '') + '\u0001' + (it.title || '');
function leaves(items, out) {
  (items || []).forEach(it => { if (!it) return; if ((it.type === 'col' || it.type === 'row') && Array.isArray(it.items)) leaves(it.items, out); else out.push(it); });
  return out;
}
export function layoutFor(kind) { return L[kind] || null; }

// commands the native editor has and the Web editor implements here
function webExtra(doc, it) {
  const t = (it.t || it.tip || '').toLowerCase();
  if (t === 'toàn màn hình') return () => { const el = document.documentElement; if (document.fullscreenElement) document.exitFullscreen(); else if (el.requestFullscreen) el.requestFullscreen(); };
  if (t === '100%') return () => { if (doc.shell && doc.shell.setZoom) doc.shell.setZoom(100); };
  if (t === 'giới thiệu') return () => import('./help.js').then(H => H.about ? H.about() : toast('Yomi-Office Open Source ' + (window.OOS_VERSION || '0.1.8') + ' — mã nguồn mở (LibreOffice, OnlyOffice, MPL-2.0)'));
  if (t === 'đọc chính tả') return () => {
    const R = window.SpeechRecognition || window.webkitSpeechRecognition; if (!R) return toast('Trình duyệt / WebView này không có nhận dạng giọng nói', 'err');
    const r = new R(); r.lang = 'vi-VN'; r.interimResults = false; r.onresult = e => { const s = e.results[0][0].transcript; document.execCommand('insertText', false, s + ' '); }; r.start(); toast('Đang nghe… hãy nói');
  };
  if (t === 'trường thời gian' || t === 'ngày & giờ') return () => { const s = new Date().toLocaleString('vi-VN'); if (doc.kind === 'excel' && doc.setInput && doc.wb) { const c = doc.wb.cur.sel; doc.setInput(c.r, c.c, s); } else document.execCommand('insertText', false, s); doc.markDirty && doc.markDirty(); };
  return null;
}

// run a command in the native editor (Windows) - the document is handed over with the command id
export async function toNative(doc, id, label) {
  if (Host.kind !== 'windows' || !(await Host.lokAvailable())) return toast('"' + label + '" chạy trong engine Native (bản Windows). Trên thiết bị này hãy dùng lệnh tương đương trong cùng nhóm.', 'warn');
  try {
    const ext = (doc.fileExt && /^(docx|xlsx|xlsm|pptx|docm|pptm|odt|ods|odp)$/.test(doc.fileExt)) ? doc.fileExt : doc.meta.newExt;
    const bytes = doc.dirty || !doc.path ? await doc.save(ext) : null;
    const ok = await Host.lokOpen({ path: doc.dirty ? '' : (doc.path || ''), name: doc.name + '.' + ext, bytes, cmd: id });
    if (!ok) toast('Không mở được engine Native', 'err');
  } catch (e) { toast('Không mở được engine Native: ' + e.message, 'err'); }
}

export function unifyRibbon(doc, def) {
  const lay = L[KIND[doc.kind]]; if (!lay) return def;
  const index = {};
  def.tabs.forEach(t => (typeof t.groups === 'function' ? [] : t.groups).forEach(g => { if (g) leaves(g.items, []).forEach(it => { const k = wkey(it); if (!index[k]) index[k] = { it, g }; }); }));
  const launchers = {};
  def.tabs.forEach(t => (typeof t.groups === 'function' ? [] : t.groups).forEach(g => { if (g && g.launcher) launchers[g.label] = g.launcher; }));
  const byId = {};
  const conv = (e, big) => {
    const w = e.web ? index[e.web] : null;
    let o;
    if (w) {
      o = Object.assign({}, w.it);
      if (o.type === 'big') o.type = 'btn';
      if (e.k === 'large') o.big = true; else if (o.type !== 'gallery' && o.type !== 'custom' && o.type !== 'combo') o.big = false;
      if (e.i && o.type !== 'gallery' && o.type !== 'custom' && o.type !== 'combo') o.icon = e.i;
      if (o.type !== 'combo' && o.type !== 'gallery' && o.type !== 'custom') { if (e.k === 'mini') o.label = ''; else o.label = e.t || ''; }
      if (e.tip) o.title = e.tip;
    } else {
      const run = webExtra(doc, e) || (() => toNative(doc, e.id, e.t || e.tip));
      o = { type: 'btn', big: e.k === 'large', label: e.k === 'mini' ? '' : (e.t || ''), title: e.tip || e.t, icon: e.i, run };
    }
    o.id = e.id; byId[e.id] = o;
    return o;
  };
  const hid = String(settings().ribbonHide || '');   // Tùy chọn > Tùy chỉnh Ribbon (same key as the native editor)
  const tabs = lay.tabs.filter((t, ti) => ti === 0 || hid.indexOf('|' + t.name + '|') < 0).map((t, ti) => {
    const wt = def.tabs.find(x => x.label === t.name);
    return {
      id: (wt && wt.id) || ('u' + ti), label: t.name, en: wt && wt.en,
      groups: t.groups.map(g => ({
        label: g.name, icon: g.icon ? 'n:' + g.icon : null,
        launcher: launchers[g.wlauncher || g.name] || (g.nlauncher != null ? () => toNative(doc, 'launcher:' + g.name, g.name) : null),
        items: g.cols.map(c => {
          if (c.type === 'rows') return { type: 'col', items: c.rows.map(r => ({ type: 'row', items: r.map(e => conv(e)) })) };
          if (c.type === 'stack') return { type: 'col', items: c.items.map(e => conv(e)) };
          return conv(c.items[0]);
        })
      }))
    };
  });
  // contextual tabs (Table / Picture / Chart …) keep their Web definition
  def.tabs.filter(t => t.contextual).forEach(t => tabs.push(t));
  doc._uiById = byId;
  const pend = window.__oosPendingCmd;
  if (pend && pend.name && doc.name && pend.name.indexOf(doc.name) === 0 && pend.id.indexOf('launcher:') === 0) {
    window.__oosPendingCmd = null; const gname = pend.id.slice(9);
    setTimeout(() => { tabs.forEach(t => (t.groups || []).forEach(g => { if (g && g.label === gname && g.launcher) g.launcher(); })); }, 900);
  } else if (pend && pend.name && doc.name && pend.name.indexOf(doc.name) === 0 && pend.id.indexOf('bs:') === 0) {
    window.__oosPendingCmd = null; const sh0 = doc.shell || window.__oosShell;
    setTimeout(() => { try { sh0 && sh0.bsAction(doc, pend.id.slice(3)); } catch (e) { console.warn(e); } }, 900);
  } else if (pend && pend.name && doc.name && pend.name.indexOf(doc.name) === 0 && byId[pend.id]) {
    window.__oosPendingCmd = null;
    setTimeout(() => { const o = byId[pend.id]; try { if (o.run) o.run(doc); else if (o.menu) toast('Mở menu "' + (o.label || o.title) + '" trên Ribbon'); } catch (e) {} }, 900);
  }
  const res = Object.assign({}, def, { tabs, unified: true, fileLabel: 'Tệp', collapseOnReclick: true });
  // the shell's [Đang chỉnh sửa] [Bình luận] [Chia sẻ] like the native ribbon - kept when a document swaps its ribbon definition
  const sh = doc.shell || window.__oosShell;
  if (sh) { res.tail = t => sh.ribbonTail(doc, t); res.onFile = () => sh.backstage(doc); } else delete res.tail;
  return res;
}
