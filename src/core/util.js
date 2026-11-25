// Core helpers (ES2017 source, transpiled to ES5 for Android 6.0 WebView 44)
import ICONS from './icons.gen.js';
import { nativeSvg } from './nicons.js';

export const VERSION = '0.1.8';
export const APP_NAME = 'Yomi-Office Open Source';

export function $(sel, root) { return (root || document).querySelector(sel); }
export function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

// h('div.cls#id', {attrs/props/on*}, children...)
export function h(tag, attrs, ...kids) {
  let m = /^([a-z0-9-]+)?((?:[.#][\w-]+)*)$/i.exec(tag) || [];
  const el = (tag === 'svg' || tag === 'path') ? document.createElementNS('http://www.w3.org/2000/svg', tag) : document.createElement(m[1] || 'div');
  if (m[2]) m[2].replace(/([.#])([\w-]+)/g, (_, t, n) => { if (t === '.') el.classList.add(n); else el.id = n; });
  if (attrs && (typeof attrs !== 'object' || attrs.nodeType || Array.isArray(attrs))) { kids.unshift(attrs); attrs = null; }
  if (attrs) for (const k in attrs) {
    const v = attrs[k];
    if (v == null || v === false) continue;
    if (k === 'class' || k === 'className') { String(v).split(/\s+/).forEach(c => c && el.classList.add(c)); }
    else if (k === 'style' && typeof v === 'object') { for (const s in v) el.style[s] = v[s]; }
    else if (k === 'html') el.innerHTML = v;
    else if (k === 'text') el.textContent = v;
    else if (k.slice(0, 2) === 'on' && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'dataset') { for (const d in v) el.setAttribute('data-' + d, v[d]); }
    else if (k in el && k !== 'list' && k !== 'form' && typeof v !== 'string') el[k] = v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  appendKids(el, kids);
  return el;
}
function appendKids(el, kids) {
  for (const k of kids) {
    if (k == null || k === false) continue;
    if (Array.isArray(k)) appendKids(el, k);
    else if (k.nodeType) el.appendChild(k);
    else el.appendChild(document.createTextNode(String(k)));
  }
}

// Inline Lucide SVG icon
export function ic(name, size, cls) {
  const s = size || 16;
  // v0.1.7: "n:<name>" = native ribbon icon (same picture as the Windows native editor), "l:<name>" = line icon
  if (name && name.indexOf('n:') === 0) { const sp = document.createElement('span'); sp.className = 'ic nic' + (cls ? ' ' + cls : ''); sp.innerHTML = nativeSvg(name.slice(2), s); return sp; }
  if (name && name.indexOf('l:') === 0) name = name.slice(2);
  const body = ICONS[name] || ICONS['file'] || '';
  const span = document.createElement('span');
  span.className = 'ic' + (cls ? ' ' + cls : '');
  span.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="' + s + '" height="' + s + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' + body + '</svg>';
  return span;
}
export function icHtml(name, size) {
  const s = size || 16;
  if (name && name.indexOf('n:') === 0) return '<span class="ic nic">' + nativeSvg(name.slice(2), s) + '</span>';
  if (name && name.indexOf('l:') === 0) name = name.slice(2);
  return '<span class="ic"><svg xmlns="http://www.w3.org/2000/svg" width="' + s + '" height="' + s + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' + (ICONS[name] || '') + '</svg></span>';
}

export function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
export function clamp(v, a, b) { return v < a ? a : v > b ? b : v; }
export function uid(p) { return (p || 'id') + Math.random().toString(36).slice(2, 9) + (uid.n = (uid.n || 0) + 1); }
export function debounce(fn, ms) { let t; return function () { const a = arguments, self = this; clearTimeout(t); t = setTimeout(() => fn.apply(self, a), ms); }; }
export function throttleRaf(fn) { let q = false, a; return function () { a = arguments; if (q) return; q = true; requestAnimationFrame(() => { q = false; fn.apply(null, a); }); }; }
export function deepClone(o) { return o == null ? o : JSON.parse(JSON.stringify(o)); }
export function noop() {}

// ---- keyboard: WebView 44 has no KeyboardEvent.key -> derive from keyCode
const KEYMAP = { 8: 'Backspace', 9: 'Tab', 13: 'Enter', 16: 'Shift', 17: 'Control', 18: 'Alt', 27: 'Escape', 32: ' ', 33: 'PageUp', 34: 'PageDown', 35: 'End', 36: 'Home', 37: 'ArrowLeft', 38: 'ArrowUp', 39: 'ArrowRight', 40: 'ArrowDown', 45: 'Insert', 46: 'Delete', 112: 'F1', 113: 'F2', 114: 'F3', 115: 'F4', 116: 'F5', 117: 'F6', 118: 'F7', 119: 'F8', 120: 'F9', 121: 'F10', 122: 'F11', 123: 'F12', 187: '=', 189: '-', 186: ';', 188: ',', 190: '.', 191: '/', 219: '[', 221: ']', 220: '\\', 222: "'", 192: '`' };
export function keyName(e) {
  if (e.key && e.key !== 'Unidentified' && e.key !== 'Process') return e.key.length === 1 ? e.key : e.key;
  const c = e.keyCode || e.which;
  if (KEYMAP[c]) return KEYMAP[c];
  if (c >= 65 && c <= 90) return String.fromCharCode(e.shiftKey ? c : c + 32);
  if (c >= 48 && c <= 57) return String.fromCharCode(c);
  if (c >= 96 && c <= 105) return String.fromCharCode(c - 48);
  return '';
}
// "ctrl+shift+z" style chord for shortcuts
export function chord(e) {
  let k = keyName(e); if (!k) return '';
  if (k.length === 1) k = k.toLowerCase();
  return (e.ctrlKey || e.metaKey ? 'ctrl+' : '') + (e.shiftKey && k.length > 1 ? 'shift+' : (e.shiftKey && /[a-z0-9]/.test(k) ? 'shift+' : '')) + (e.altKey ? 'alt+' : '') + k.toLowerCase();
}

// ---- pointer abstraction (mouse + touch; WebView 44 has no PointerEvent)
export function pt(e) { const t = e.touches && e.touches[0] || e.changedTouches && e.changedTouches[0] || e; return { x: t.clientX, y: t.clientY }; }
export function onDrag(el, start) {
  function down(e) {
    if (e.type === 'mousedown' && e.button !== 0) return;
    const ctl = start(e, pt(e));
    if (!ctl) return;
    const touch = e.type === 'touchstart';
    const mv = ev => { if (ctl.move) ctl.move(ev, pt(ev)); if (touch && ev.cancelable) ev.preventDefault(); };
    const up = ev => {
      document.removeEventListener(touch ? 'touchmove' : 'mousemove', mv);
      document.removeEventListener(touch ? 'touchend' : 'mouseup', up);
      if (touch) document.removeEventListener('touchcancel', up);
      if (ctl.end) ctl.end(ev, pt(ev));
    };
    document.addEventListener(touch ? 'touchmove' : 'mousemove', mv, touch ? { passive: false } : false);
    document.addEventListener(touch ? 'touchend' : 'mouseup', up);
    if (touch) document.addEventListener('touchcancel', up);
  }
  el.addEventListener('mousedown', down);
  el.addEventListener('touchstart', down, { passive: false });
}

// ---- binary helpers
export function b64FromBytes(bytes) {
  let bin = ''; const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
  return btoa(bin);
}
export function bytesFromB64(b64) { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; }
export function utf8Encode(s) { if (window.TextEncoder) return new TextEncoder().encode(s); return bytesFromBinStr(unescape(encodeURIComponent(s))); }
export function utf8Decode(u) { if (window.TextDecoder) return new TextDecoder('utf-8').decode(u); let s = ''; for (let i = 0; i < u.length; i++) s += String.fromCharCode(u[i]); try { return decodeURIComponent(escape(s)); } catch (e) { return s; } }
function bytesFromBinStr(s) { const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; }
export function dataUrlToBytes(url) { return bytesFromB64(url.slice(url.indexOf(',') + 1)); }
export function bytesToDataUrl(bytes, mime) { return 'data:' + (mime || 'application/octet-stream') + ';base64,' + b64FromBytes(bytes); }
export function readFileAsBytes(file) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(new Uint8Array(r.result)); r.onerror = () => rej(r.error); r.readAsArrayBuffer(file); }); }
export function readFileAsDataUrl(file) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(file); }); }
export function extOf(name) { const m = /\.([a-z0-9]+)$/i.exec(name || ''); return m ? m[1].toLowerCase() : ''; }
export function baseName(name) { return String(name || '').replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, ''); }
export function canvasToBytes(canvas, mime, q) { return dataUrlToBytes(canvas.toDataURL(mime || 'image/png', q)); }
export function loadImage(src) { return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('image')); i.src = src; }); }

export function mimeOf(ext) {
  return ({ docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', pdf: 'application/pdf', csv: 'text/csv', txt: 'text/plain', html: 'text/html', htm: 'text/html', md: 'text/markdown', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', zip: 'application/zip', json: 'application/json', rtf: 'application/rtf', svg: 'image/svg+xml', odt: 'application/vnd.oasis.opendocument.text' })[ext] || 'application/octet-stream';
}

// ---- script loader (vendor libs lazily)
const loaded = {};
export function loadScript(src) {
  if (loaded[src]) return loaded[src];
  loaded[src] = new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = src; s.async = true;
    s.onload = () => res(); s.onerror = () => { delete loaded[src]; rej(new Error('Không tải được ' + src)); };
    document.head.appendChild(s);
  });
  return loaded[src];
}
export function loadCss(href) {
  if (loaded[href]) return loaded[href];
  loaded[href] = new Promise(res => { const l = document.createElement('link'); l.rel = 'stylesheet'; l.href = href; l.onload = res; l.onerror = res; document.head.appendChild(l); });
  return loaded[href];
}
export async function JSZipLib() { if (!window.JSZip) await loadScript('vendor/jszip.min.js'); return window.JSZip; }
export async function PDFLib() { if (!window.PDFLib) await loadScript('vendor/pdf-lib.min.js'); return window.PDFLib; }
export async function PDFJS() {
  if (!window.pdfjsLib) await loadScript('vendor/pdfjs/pdf.min.js');
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.min.js';
  return window.pdfjsLib;
}
export async function KaTeX() { if (!window.katex) { loadCss('vendor/katex/katex.min.css'); await loadScript('vendor/katex/katex.min.js'); } return window.katex; }

// ---- formatting
export function fmtDate(d) { d = d || new Date(); const p = n => (n < 10 ? '0' : '') + n; return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear(); }
export function fmtTime(d) { d = d || new Date(); const p = n => (n < 10 ? '0' : '') + n; return p(d.getHours()) + ':' + p(d.getMinutes()); }
export function fmtSize(n) { return n < 1024 ? n + ' B' : n < 1048576 ? (n / 1024).toFixed(1) + ' KB' : (n / 1048576).toFixed(1) + ' MB'; }
export function relTime(ts) {
  const d = new Date(ts), now = new Date();
  if (d.toDateString() === now.toDateString()) return 'Hôm nay lúc ' + fmtTime(d);
  const y = new Date(now.getTime() - 864e5);
  if (d.toDateString() === y.toDateString()) return 'Hôm qua lúc ' + fmtTime(d);
  return fmtDate(d);
}
export function stripHtml(html) { const d = document.createElement('div'); d.innerHTML = html; return d.textContent || ''; }

export const isTouch = ('ontouchstart' in window) || navigator.maxTouchPoints > 0;
export const isAndroid = /Android/i.test(navigator.userAgent);
export function chromeVersion() { const m = /Chrome\/(\d+)/.exec(navigator.userAgent); return m ? +m[1] : 999; }
export function supportsWasm() { try { return typeof WebAssembly === 'object' && typeof WebAssembly.instantiate === 'function'; } catch (e) { return false; } }

export function safeStore(key, val) {
  try { if (val === undefined) { const v = localStorage.getItem(key); return v == null ? null : JSON.parse(v); } localStorage.setItem(key, JSON.stringify(val)); } catch (e) { return null; }
  return val;
}

export function emitter() {
  const m = {};
  return {
    on(ev, fn) { (m[ev] = m[ev] || []).push(fn); return () => { m[ev] = (m[ev] || []).filter(f => f !== fn); }; },
    emit(ev, a, b) { (m[ev] || []).slice().forEach(f => { try { f(a, b); } catch (e) { console.error(e); } }); }
  };
}
