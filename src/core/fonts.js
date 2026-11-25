// Yomi-Office Open Source v0.0.6 — font manager: built-in Vietnamese fonts, Font Pack import (.zip / .ttf / .otf / .ttc),
// fonts installed on the computer (Local Font Access), Windows font folder served by the launcher, lazy loading via FontFace.
// Also: legacy Vietnamese encodings converter (TCVN3/ABC ".Vn…" fonts and VNI fonts → Unicode).
import { h, ic, esc, debounce, JSZipLib } from './util.js';
import { dialog, toast, busy, confirmDlg } from './ui.js';
import Host from './host.js';

// one shared, live list used by every font box (Word / Excel / PowerPoint)
export const FONT_LIST = ['Calibri', 'Calibri Light', 'Times New Roman', 'Arial', 'Tahoma', 'Verdana', 'Cambria', 'Georgia', 'Segoe UI', 'Roboto', 'Courier New', 'Consolas', 'Garamond', 'Trebuchet MS', 'Comic Sans MS', 'Impact', 'Noto Sans', 'Noto Serif', 'Arial Narrow', 'Book Antiqua', 'Palatino Linotype', 'Century Gothic', 'Be Vietnam Pro', '.VnTime', 'VNI-Times'];
// built-in fonts shipped inside the app (free Vietnamese fonts from D:\A.I Project\font)
export const BUILTIN = [
  { family: 'HP001 4 hàng', file: 'HP001_4_hang_normal.ttf', weight: 'normal' }, { family: 'HP001 4 hàng', file: 'HP001_4_hang_bold.ttf', weight: 'bold' },
  { family: 'HP001 4 hàng 1 ô ly', file: 'HP001_4_hang_1_o_ly.ttf', weight: 'normal' }, { family: 'HP001 4 hàng 2 ô ly', file: 'HP001_4_hang_2_o_ly.ttf', weight: 'normal' },
  { family: 'HP001 5 hàng', file: 'HP001_5_hang_normal.ttf', weight: 'normal' }, { family: 'HP001 5 hàng', file: 'HP001_5_hang_bold.ttf', weight: 'bold' },
  { family: 'HP001 5 hàng 1 ô ly', file: 'HP001_5_hang_1_o_ly.ttf', weight: 'normal' }, { family: 'Wisteria Notes', file: 'Wisteria_Notes.ttf', weight: 'normal' }
];
const sources = {};   // family -> 'builtin' | 'pack' | 'system' | 'windows'
function addFamily(f, src) { if (!f) return; if (FONT_LIST.indexOf(f) < 0) FONT_LIST.push(f); if (!sources[f] || src === 'builtin') sources[f] = src; }
export function fontSource(f) { return sources[f] || 'default'; }

let styleEl = null;
function css(rule) { if (!styleEl) { styleEl = document.createElement('style'); styleEl.id = 'oos-fonts'; document.head.appendChild(styleEl); } styleEl.appendChild(document.createTextNode(rule + '\n')); }
const q = s => "'" + String(s).replace(/'/g, "\\'") + "'";

// ---------------------------------------------------------------- font file parsing (name table)
export function fontNames(bytes, index) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength); let off = 0;
  if (dv.getUint32(0) === 0x74746366) { const n = dv.getUint32(8); const out = []; for (let i = 0; i < n; i++) { const r = fontNames(bytes, -1 - dv.getUint32(12 + 4 * i)); if (r) out.push(r); } return index == null ? out : out[index]; }
  if (index != null && index < 0) off = -1 - index;
  try {
    const nt = dv.getUint16(off + 4); let nameOff = -1;
    for (let i = 0; i < nt; i++) { const tag = String.fromCharCode(bytes[off + 12 + 16 * i], bytes[off + 13 + 16 * i], bytes[off + 14 + 16 * i], bytes[off + 15 + 16 * i]); if (tag === 'name') nameOff = dv.getUint32(off + 12 + 16 * i + 8); }
    if (nameOff < 0) return null;
    const count = dv.getUint16(nameOff + 2), so = nameOff + dv.getUint16(nameOff + 4); const res = {};
    for (let j = 0; j < count; j++) { const b = nameOff + 6 + 12 * j; const pid = dv.getUint16(b), nid = dv.getUint16(b + 6), len = dv.getUint16(b + 8), no = dv.getUint16(b + 10);
      if ([1, 2, 4, 16, 17].indexOf(nid) < 0) continue; let s = '';
      if (pid === 3 || pid === 0) { for (let k = 0; k < len; k += 2) s += String.fromCharCode(dv.getUint16(so + no + k)); if (!res[nid] || pid === 3) res[nid] = s; }
      else if (pid === 1 && !res[nid]) { for (let k = 0; k < len; k++) s += String.fromCharCode(bytes[so + no + k]); res[nid] = s; } }
    const family = res[16] || res[1]; const sub = (res[17] || res[2] || 'Regular').toLowerCase();
    return { family, style: sub, full: res[4] || family, weight: /bold|black|heavy|semibold|demibold/.test(sub) ? 'bold' : 'normal', italic: /italic|oblique/.test(sub) };
  } catch (e) { return null; }
}

// ---------------------------------------------------------------- storage (IndexedDB "oos6-fonts")
let dbp = null;
function db() { if (dbp) return dbp; dbp = new Promise(res => { try { const r = indexedDB.open('oos6-fonts', 1); r.onupgradeneeded = () => { const d = r.result; d.createObjectStore('meta', { keyPath: 'id' }); d.createObjectStore('data', { keyPath: 'id' }); }; r.onsuccess = () => res(r.result); r.onerror = () => res(null); } catch (e) { res(null); } }); return dbp; }
async function tx(store, mode, fn) { const d = await db(); if (!d) return null; return new Promise((res, rej) => { const t = d.transaction(store, mode); const s = t.objectStore(store); let out; const r = fn(s); if (r) r.onsuccess = () => { out = r.result; }; t.oncomplete = () => res(out); t.onerror = () => rej(t.error); }); }
const metaAll = () => tx('meta', 'readonly', s => s.getAll ? s.getAll() : null).then(r => r || []);

const loaded = new Set();
async function loadPackFace(m) {
  if (loaded.has(m.id)) return; loaded.add(m.id);
  try { const rec = await tx('data', 'readonly', s => s.get(m.id)); if (!rec) return; const buf = rec.bytes.buffer ? rec.bytes.buffer.slice(rec.bytes.byteOffset, rec.bytes.byteOffset + rec.bytes.byteLength) : rec.bytes;
    if (window.FontFace) { const ff = new FontFace(m.family, buf, { weight: m.weight === 'bold' ? '700' : '400', style: m.italic ? 'italic' : 'normal' }); await ff.load(); document.fonts.add(ff); }
    else { const url = URL.createObjectURL(new Blob([buf], { type: 'font/ttf' })); css('@font-face{font-family:' + q(m.family) + ';src:url(' + url + ');font-weight:' + (m.weight === 'bold' ? 700 : 400) + ';font-style:' + (m.italic ? 'italic' : 'normal') + '}'); }
  } catch (e) { console.warn('font', m.family, e); }
}
let packMeta = [];
// make sure the given family is loaded before it is rendered (fonts from the pack are loaded lazily)
export function ensureFont(family) { if (!family) return; const f = String(family).replace(/['"]/g, '').split(',')[0].trim(); packMeta.filter(m => m.family === f).forEach(loadPackFace); }
export function ensureFontsIn(root) { if (!packMeta.length || !root) return; const seen = new Set(); const walk = el => { if (el.nodeType !== 1) return; const ff = el.style && el.style.fontFamily; if (ff && !seen.has(ff)) { seen.add(ff); ensureFont(ff); } for (let c = el.firstChild; c; c = c.nextSibling) walk(c); }; walk(root); }

// ---------------------------------------------------------------- boot
export async function initFonts() {
  BUILTIN.forEach(b => { css('@font-face{font-family:' + q(b.family) + ';src:url(assets/fonts/' + encodeURI(b.file) + ') format("truetype");font-weight:' + (b.weight === 'bold' ? 700 : 400) + ';font-display:swap}'); addFamily(b.family, 'builtin'); });
  try { packMeta = await metaAll(); packMeta.forEach(m => addFamily(m.family, 'pack')); } catch (e) {}
  if (Host.kind === 'windows') { try { const r = await fetch('/api/fonts', { method: 'POST', body: '{}' }).then(x => x.json()); (r.fonts || []).forEach(f => { css('@font-face{font-family:' + q(f.family) + ';src:url(/fontfile/' + encodeURIComponent(f.file) + ');font-weight:' + (f.weight === 'bold' ? 700 : 400) + ';font-style:' + (f.italic ? 'italic' : 'normal') + ';font-display:swap}'); addFamily(f.family, 'windows'); }); } catch (e) {} }
  if (Host.native && Host.native.systemFonts) { try { JSON.parse(Host.native.systemFonts()).forEach(f => addFamily(f, 'system')); } catch (e) {} }
  if (document.fonts && document.fonts.addEventListener) document.fonts.addEventListener('loadingdone', () => { const d = window.OOS && window.OOS.cur; if (!d) return; if (d.grid) { d.grid.measureCache = {}; d.grid.redraw(); } if (d.renderSlide && d.kind === 'ppt') d.renderThumbs(); });
  // fonts already granted through Local Font Access (Chrome/Edge 103+)
  try { if (navigator.permissions && window.queryLocalFonts) { const st = await navigator.permissions.query({ name: 'local-fonts' }); if (st.state === 'granted') await addSystemFonts(true); } } catch (e) {}
}
export async function addSystemFonts(silent) {
  if (!window.queryLocalFonts) { if (!silent) toast('Trình duyệt/WebView này chưa hỗ trợ đọc danh sách phông của máy (cần Chrome/Edge 103+). Hãy dùng "Nạp phông" để thêm tệp phông.', 'err', 6000); return 0; }
  try { const list = await window.queryLocalFonts(); const fams = new Set(list.map(f => f.family)); fams.forEach(f => addFamily(f, 'system')); if (!silent) toast('Đã thêm ' + fams.size + ' phông chữ đã cài trên máy', 'ok'); return fams.size; }
  catch (e) { if (!silent) toast('Không đọc được phông của máy: ' + (e.message || e), 'err'); return 0; }
}

// ---------------------------------------------------------------- import (files or Font Pack zip)
export async function importFontFiles(files) {
  const b = busy('Đang nạp phông chữ…'); let n = 0; const fams = new Set();
  try {
    const items = [];
    for (const f of files) {
      if (/\.zip$/i.test(f.name)) { const JSZip = await JSZipLib(); const z = await JSZip.loadAsync(f.bytes); const ents = Object.keys(z.files).filter(p => /\.(ttf|otf|ttc)$/i.test(p) && !z.files[p].dir); for (const p of ents) items.push({ name: p.split('/').pop(), bytes: await z.files[p].async('uint8array') }); }
      else if (/\.(ttf|otf|ttc)$/i.test(f.name)) items.push(f);
    }
    for (const it of items) {
      const info = fontNames(it.bytes); const list = Array.isArray(info) ? info : [info];
      for (let i = 0; i < list.length; i++) { const nm = list[i]; if (!nm || !nm.family) continue;
        const id = (nm.full || nm.family) + '|' + nm.style; const meta = { id, family: nm.family, style: nm.style, weight: nm.weight, italic: nm.italic, file: it.name, size: it.bytes.length, ttc: list.length > 1 ? i : null };
        let bytes = it.bytes; if (list.length > 1) bytes = it.bytes; // collections: browsers load the first face only — kept for completeness
        await tx('meta', 'readwrite', s => s.put(meta)); await tx('data', 'readwrite', s => s.put({ id, bytes })); packMeta.push(meta); addFamily(nm.family, 'pack'); fams.add(nm.family); n++; }
    }
  } finally { b.done(); }
  toast('Đã nạp ' + n + ' kiểu phông (' + fams.size + ' họ phông)', 'ok', 4000);
  return n;
}
export async function removeFamily(fam) { const ms = packMeta.filter(m => m.family === fam); for (const m of ms) { await tx('meta', 'readwrite', s => s.delete(m.id)); await tx('data', 'readwrite', s => s.delete(m.id)); } packMeta = packMeta.filter(m => m.family !== fam); const i = FONT_LIST.indexOf(fam); if (i >= 0 && sources[fam] === 'pack') FONT_LIST.splice(i, 1); delete sources[fam]; }

// ---------------------------------------------------------------- Font manager dialog
export function fontManager() {
  const search = h('input.inp', { placeholder: 'Tìm phông chữ…' }); const sample = h('input.inp', { value: 'Tiếng Việt có dấu — Hôm nay trời đẹp quá! 0123456789' });
  const list = h('div.fm-list'); const info = h('div.muted.small');
  const badge = { builtin: 'Tích hợp', pack: 'Font Pack', system: 'Máy tính', windows: 'Windows', default: 'Mặc định' };
  const draw = () => { list.innerHTML = ''; const qv = search.value.trim().toLowerCase(); const fams = FONT_LIST.filter(f => !qv || f.toLowerCase().indexOf(qv) >= 0).slice(0, 400);
    fams.forEach(f => { ensureFont(f); const row = h('div.fm-row', h('div.fm-name', f, h('span.fm-badge.b-' + fontSource(f), badge[fontSource(f)])), h('div.fm-prev', { style: { fontFamily: q(f) + ', sans-serif' } }, sample.value)); if (fontSource(f) === 'pack') { const del = h('button.btn.small', { title: 'Gỡ phông khỏi ứng dụng' }, ic('trash-2', 13)); del.addEventListener('click', async () => { if (await confirmDlg('Gỡ phông "' + f + '" khỏi ứng dụng?', 'Gỡ phông')) { await removeFamily(f); draw(); } }); row.appendChild(del); } list.appendChild(row); });
    info.textContent = FONT_LIST.length + ' họ phông · tích hợp ' + BUILTIN.length + ' tệp · Font Pack ' + packMeta.length + ' kiểu'; };
  search.addEventListener('input', debounce(draw, 150)); sample.addEventListener('input', debounce(draw, 300));
  const body = h('div.fm', h('div.fm-top', search, sample), info, list, h('div.hint', 'Nạp "Font Pack" (OfficeOpenSource-FontPack-*.zip) hoặc từng tệp .ttf/.otf. Phông được lưu trong ứng dụng, dùng được cả khi ngoại tuyến, trên Android và Windows.'));
  draw();
  return dialog({ title: 'Quản lý Phông chữ', width: 760, body, buttons: [
    { label: 'Nạp phông / Font Pack…', primary: true, run: async () => { const fs = await Host.openFiles('.ttf,.otf,.ttc,.zip', true); if (fs.length) { await importFontFiles(fs); draw(); } return false; } },
    { label: 'Phông của máy', run: async () => { await addSystemFonts(); draw(); return false; } },
    { label: 'Đóng' }] });
}

// ---------------------------------------------------------------- legacy Vietnamese encodings → Unicode
const TCVN3 = { 0xA1: 'Ă', 0xA2: 'Â', 0xA3: 'Ê', 0xA4: 'Ô', 0xA5: 'Ơ', 0xA6: 'Ư', 0xA7: 'Đ', 0xA8: 'ă', 0xA9: 'â', 0xAA: 'ê', 0xAB: 'ô', 0xAC: 'ơ', 0xAD: 'ư', 0xAE: 'đ',
  0xB5: 'à', 0xB6: 'ả', 0xB7: 'ã', 0xB8: 'á', 0xB9: 'ạ', 0xBB: 'ằ', 0xBC: 'ẳ', 0xBD: 'ẵ', 0xBE: 'ắ', 0xC6: 'ặ', 0xC7: 'ầ', 0xC8: 'ẩ', 0xC9: 'ẫ', 0xCA: 'ấ', 0xCB: 'ậ',
  0xCC: 'è', 0xCE: 'ẻ', 0xCF: 'ẽ', 0xD0: 'é', 0xD1: 'ẹ', 0xD2: 'ề', 0xD3: 'ể', 0xD4: 'ễ', 0xD5: 'ế', 0xD6: 'ệ', 0xD7: 'ì', 0xD8: 'ỉ', 0xDC: 'ĩ', 0xDD: 'í', 0xDE: 'ị',
  0xDF: 'ò', 0xE1: 'ỏ', 0xE2: 'õ', 0xE3: 'ó', 0xE4: 'ọ', 0xE5: 'ồ', 0xE6: 'ổ', 0xE7: 'ỗ', 0xE8: 'ố', 0xE9: 'ộ', 0xEA: 'ờ', 0xEB: 'ở', 0xEC: 'ỡ', 0xED: 'ớ', 0xEE: 'ợ',
  0xEF: 'ù', 0xF1: 'ủ', 0xF2: 'ũ', 0xF3: 'ú', 0xF4: 'ụ', 0xF5: 'ừ', 0xF6: 'ử', 0xF7: 'ữ', 0xF8: 'ứ', 0xF9: 'ự', 0xFA: 'ỳ', 0xFB: 'ỷ', 0xFC: 'ỹ', 0xFD: 'ý', 0xFE: 'ỵ' };
export function tcvn3ToUnicode(s, upper) { let o = ''; for (const ch of String(s)) { const c = ch.charCodeAt(0); const u = c < 256 ? TCVN3[c] : null; o += u ? (upper ? u.toUpperCase() : u) : ch; } return o; }
// VNI Windows: base letter followed by mark bytes
const VNI_MARK = { 'ù': '\u0301', 'ø': '\u0300', 'û': '\u0309', 'õ': '\u0303', 'ï': '\u0323', 'Ù': '\u0301', 'Ø': '\u0300', 'Û': '\u0309', 'Õ': '\u0303', 'Ï': '\u0323' };
const VNI_CIRC = { 'â': '', 'á': '\u0301', 'à': '\u0300', 'å': '\u0309', 'ã': '\u0303', 'ä': '\u0323', 'Â': '', 'Á': '\u0301', 'À': '\u0300', 'Å': '\u0309', 'Ã': '\u0303', 'Ä': '\u0323' };
const VNI_BREVE = { 'ê': '', 'é': '\u0301', 'è': '\u0300', 'ú': '\u0309', 'ü': '\u0303', 'ë': '\u0323', 'Ê': '', 'É': '\u0301', 'È': '\u0300', 'Ú': '\u0309', 'Ü': '\u0303', 'Ë': '\u0323' };
const VNI_I = { 'í': 'í', 'ì': 'ì', 'æ': 'ỉ', 'ó': 'ĩ', 'ò': 'ị', 'Í': 'Í', 'Ì': 'Ì', 'Æ': 'Ỉ', 'Ó': 'Ĩ', 'Ò': 'Ị' };
export function vniToUnicode(s) {
  s = String(s); const out = [];
  const nfc = x => (x.normalize ? x.normalize('NFC') : x);
  for (let i = 0; i < s.length; i++) {
    const c = s[i]; const last = out.length ? out[out.length - 1] : '';
    if (c === 'ñ') { out.push('đ'); continue; } if (c === 'Ñ') { out.push('Đ'); continue; }
    if (c === 'ô' || c === 'Ô' || c === 'ö' || c === 'Ö') { out.push(({ 'ô': 'ơ', 'Ô': 'Ơ', 'ö': 'ư', 'Ö': 'Ư' })[c]); continue; }
    if (VNI_I[c]) { out.push(VNI_I[c]); continue; }
    if (VNI_CIRC[c] != null && /^[aeoAEO]$/.test(last)) { out[out.length - 1] = nfc(({ a: 'â', A: 'Â', e: 'ê', E: 'Ê', o: 'ô', O: 'Ô' })[last] + VNI_CIRC[c]); continue; }
    if (VNI_BREVE[c] != null && /^[aA]$/.test(last)) { out[out.length - 1] = nfc((last === 'a' ? 'ă' : 'Ă') + VNI_BREVE[c]); continue; }
    if (VNI_MARK[c] && /^[aeiouyAEIOUYơưâêôăƠƯÂÊÔĂ]$/.test(last)) { out[out.length - 1] = nfc(last + VNI_MARK[c]); continue; }
    out.push(c);
  }
  return out.join('');
}
export function detectLegacy(text, font) {
  const f = String(font || '').toLowerCase();
  if (/^\.vn|^vn[a-z]|^vh[a-z]/.test(f.replace(/^\./, '.')) || /\.vn/.test(f)) return /h$|vh/.test(f) ? 'tcvn3-upper' : 'tcvn3';
  if (/^vni/.test(f)) return 'vni';
  const t = String(text); if (/[ñÑ]|[aeo][áàåãäâ]|a[éèúüëê]|[ôö][ùøûõï]/.test(t)) return 'vni';
  if (/[\u00b5-\u00fe]{1}/.test(t) && /[\u00a1-\u00ae]/.test(t)) return 'tcvn3';
  return null;
}
export function convertLegacy(text, kind) { return kind === 'vni' ? vniToUnicode(text) : tcvn3ToUnicode(text, kind === 'tcvn3-upper'); }

// ---------------------------------------------------------------- "Chuyển mã" (Unikey-style converter) for Word / Excel
const FONT_MAP = { times: 'Times New Roman', arial: 'Arial', helv: 'Arial', univ: 'Arial', cour: 'Courier New', avan: 'Arial', centu: 'Century Gothic', book: 'Book Antiqua' };
export function unicodeFontFor(font) { const f = String(font || '').toLowerCase().replace(/^\.vn|^vni-?|^vh/, ''); for (const k in FONT_MAP) if (f.indexOf(k) >= 0) return FONT_MAP[k]; return 'Times New Roman'; }
export async function convertDialog() {
  const r = await dialog({ title: 'Chuyển mã tiếng Việt → Unicode', width: 460, body: h('div.form', h('div.dlg-msg', 'Chuyển văn bản gõ bằng bảng mã cũ (phông .VnTime / .VnArial — TCVN3/ABC, hoặc VNI-Times …) sang Unicode dựng sẵn để hiển thị đúng trên mọi máy.'), h('label.chk', h('input', { type: 'radio', name: 'cv', value: 'auto', checked: true }), h('span', 'Tự nhận diện theo phông chữ (khuyên dùng)')), h('label.chk', h('input', { type: 'radio', name: 'cv', value: 'tcvn3' }), h('span', 'TCVN3 (ABC) — phông .Vn…')), h('label.chk', h('input', { type: 'radio', name: 'cv', value: 'tcvn3-upper' }), h('span', 'TCVN3 chữ hoa — phông .Vn…H')), h('label.chk', h('input', { type: 'radio', name: 'cv', value: 'vni' }), h('span', 'VNI Windows — phông VNI-…'))), buttons: [{ label: 'Chuyển mã', primary: true, run: b => (b.querySelector('input[name=cv]:checked') || {}).value || 'auto' }, { label: 'Hủy', value: null }] });
  return r;
}
// Word: convert the selection (or the whole document) in place, switching legacy fonts to Unicode fonts
export function convertWordLegacy(doc, mode) {
  const sel = window.getSelection(); const range = sel.rangeCount && !sel.isCollapsed && doc.body.contains(sel.anchorNode) ? sel.getRangeAt(0) : null;
  let n = 0; const walker = document.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT, null, false); const nodes = []; let t; while ((t = walker.nextNode())) { if (!range || range.intersectsNode(t)) nodes.push(t); }
  nodes.forEach(node => { const el = node.parentElement; const ff = el ? getComputedStyle(el).fontFamily.replace(/['"]/g, '').split(',')[0].trim() : ''; const kind = mode === 'auto' ? detectLegacy(node.nodeValue, ff) : mode; if (!kind) return; const nv = convertLegacy(node.nodeValue, kind); if (nv !== node.nodeValue) { node.nodeValue = nv; n++; } if (el && /^(\.vn|vni|vh)/i.test(ff)) { let e = el; while (e && e !== doc.body && !(e.style && e.style.fontFamily)) e = e.parentElement; if (e && e !== doc.body) e.style.fontFamily = unicodeFontFor(ff); } });
  return n;
}
