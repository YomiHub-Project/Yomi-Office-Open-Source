// v0.0.7 — E-book reader (ideas from FBReader, Calibre, Sigil, LNReader/IReader): EPUB 2/3, FB2/FBZ, TXT
// Table of contents, scroll or page mode, font size/family, light/sepia/dark themes, search, read aloud (Web Speech),
// reading position remembered per book, "open in Word" (editable) — fully offline.
import { h, ic, esc, $$, JSZipLib, utf8Decode, utf8Encode, extOf, baseName, bytesToDataUrl, mimeOf, clamp, safeStore } from '../core/util.js';
import { toast, busy } from '../core/ui.js';
import Host from '../core/host.js';

// ---------------------------------------------------------------- parsing
function parseXml(s, html) { const p = new DOMParser(); let d = p.parseFromString(s, html ? 'text/html' : 'application/xml'); if (!html && d.getElementsByTagName('parsererror').length) d = p.parseFromString(s, 'text/html'); return d; }
function byLocal(root, name) { const out = []; const all = root.getElementsByTagName('*'); for (let i = 0; i < all.length; i++) if ((all[i].localName || all[i].nodeName).toLowerCase() === name) out.push(all[i]); return out; }
function first(root, name) { return byLocal(root, name)[0] || null; }
function resolve(base, rel) { if (/^[a-z]+:/i.test(rel) || rel.charAt(0) === '/') return rel.replace(/^\//, ''); const parts = base.split('/'); parts.pop(); rel.split('/').forEach(p => { if (p === '..') parts.pop(); else if (p !== '.' && p !== '') parts.push(p); }); return parts.join('/'); }
function decodeText(bytes) { const s = utf8Decode(bytes); if (s.indexOf('�') < 0) return s; try { return new TextDecoder('windows-1258').decode(bytes); } catch (e) { return s; } }
const SAFE = /^(p|div|span|h[1-6]|em|i|b|strong|u|s|sub|sup|br|hr|img|ul|ol|li|table|thead|tbody|tr|td|th|blockquote|pre|code|a|figure|figcaption|section|article|small|big|dl|dt|dd|center|font|ruby|rt)$/i;
function sanitize(node, doc, img) {
  // copy only safe elements/attributes (no scripts, no external loads)
  const out = doc.createDocumentFragment();
  (function walk(src, dst) {
    for (let c = src.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) { dst.appendChild(doc.createTextNode(c.nodeValue)); continue; }
      if (c.nodeType !== 1) continue;
      const tag = (c.localName || c.nodeName).toLowerCase();
      if (/^(script|style|link|meta|head|title|iframe|object|embed)$/.test(tag)) continue;
      if (tag === 'image' || tag === 'svg') { const im = c.getAttribute('xlink:href') || c.getAttribute('href') || (first(c, 'image') && (first(c, 'image').getAttribute('xlink:href') || first(c, 'image').getAttribute('href'))); if (im) { const e = doc.createElement('img'); const u = img(im); if (u) { e.src = u; dst.appendChild(e); } } continue; }
      if (!SAFE.test(tag)) { walk(c, dst); continue; }
      const e = doc.createElement(tag);
      if (tag === 'img') { const u = img(c.getAttribute('src') || ''); if (!u) continue; e.src = u; e.alt = c.getAttribute('alt') || ''; }
      if (tag === 'a') { const href = c.getAttribute('href') || ''; if (href.charAt(0) === '#' || !/^[a-z]+:/i.test(href)) e.setAttribute('data-href', href); }
      if (c.getAttribute('id')) e.id = 'r-' + c.getAttribute('id');
      if (/^(td|th)$/.test(tag)) { ['colspan', 'rowspan'].forEach(a => { if (c.getAttribute(a)) e.setAttribute(a, c.getAttribute(a)); }); }
      dst.appendChild(e); walk(c, e);
    }
  })(node, out);
  return out;
}

export async function parseEpub(bytes) {
  const JSZip = await JSZipLib(); const zip = await JSZip.loadAsync(bytes);
  const txt = async p => { const f = zip.file(p) || zip.file(decodeURIComponent(p)); if (!f) return null; return decodeText(await f.async('uint8array')); };
  const cont = parseXml(await txt('META-INF/container.xml') || '');
  const rf = first(cont, 'rootfile'); const opfPath = rf ? rf.getAttribute('full-path') : (Object.keys(zip.files).find(n => /\.opf$/i.test(n)));
  const opf = parseXml(await txt(opfPath));
  const meta = { title: (first(opf, 'title') || {}).textContent || baseName(opfPath), author: (first(opf, 'creator') || {}).textContent || '' };
  const items = {}; byLocal(opf, 'item').forEach(it => { items[it.getAttribute('id')] = { href: resolve(opfPath, it.getAttribute('href')), type: it.getAttribute('media-type'), props: it.getAttribute('properties') || '' }; });
  const spine = byLocal(opf, 'itemref').map(r => items[r.getAttribute('idref')]).filter(x => x && /html|xml/.test(x.type || ''));
  const imgCache = {};
  const imgUrl = async p => { if (imgCache[p] !== undefined) return imgCache[p]; const f = zip.file(p) || zip.file(decodeURIComponent(p)); imgCache[p] = f ? bytesToDataUrl(await f.async('uint8array'), mimeOf(extOf(p))) : null; return imgCache[p]; };
  const chapters = [];
  for (const it of spine) {
    const src = await txt(it.href); if (src == null) continue;
    const d = parseXml(src, !/<\?xml|xmlns=/.test(src.slice(0, 400)));
    const body = first(d, 'body') || d.documentElement;
    const imgs = {}; for (const im of byLocal(body, 'img').concat(byLocal(body, 'image'))) { const r = im.getAttribute('src') || im.getAttribute('xlink:href') || im.getAttribute('href'); if (r && !/^data:/.test(r)) imgs[r] = await imgUrl(resolve(it.href, r)); }
    const frag = sanitize(body, document, r => /^data:/.test(r) ? r : imgs[r] || null);
    const holder = document.createElement('div'); holder.appendChild(frag);
    const t = first(d, 'title'); const hd = holder.querySelector('h1,h2,h3');
    chapters.push({ href: it.href, title: (hd && hd.textContent.trim()) || (t && t.textContent.trim()) || ('Chương ' + (chapters.length + 1)), html: holder.innerHTML });
  }
  // TOC: EPUB3 nav or EPUB2 NCX (falls back to chapter headings)
  let toc = [];
  const navItem = Object.keys(items).map(k => items[k]).find(x => /\bnav\b/.test(x.props));
  const ncxItem = Object.keys(items).map(k => items[k]).find(x => /ncx/.test(x.type || ''));
  try {
    if (navItem) { const nd = parseXml(await txt(navItem.href), true); const nav = byLocal(nd, 'nav').find(n => /toc/.test(n.getAttribute('epub:type') || n.getAttribute('type') || 'toc')) || first(nd, 'nav'); if (nav) toc = byLocal(nav, 'a').map(a => ({ title: a.textContent.trim(), href: resolve(navItem.href, a.getAttribute('href') || ''), level: byLocal(nav, 'ol').length ? depth(a, nav) : 0 })); }
    else if (ncxItem) { const nd = parseXml(await txt(ncxItem.href)); toc = byLocal(nd, 'navpoint').map(np => ({ title: ((first(np, 'text') || {}).textContent || '').trim(), href: resolve(ncxItem.href, (first(np, 'content') || { getAttribute: () => '' }).getAttribute('src')), level: depthNcx(np) })); }
  } catch (e) {}
  if (!toc.length) toc = chapters.map(c => ({ title: c.title, href: c.href, level: 0 }));
  return { meta, chapters, toc };
}
function depth(a, root) { let n = 0; for (let p = a.parentNode; p && p !== root; p = p.parentNode) if ((p.localName || '').toLowerCase() === 'ol') n++; return Math.max(0, n - 1); }
function depthNcx(np) { let n = 0; for (let p = np.parentNode; p; p = p.parentNode) if ((p.localName || '').toLowerCase() === 'navpoint') n++; return n; }

export async function parseFb2(bytes, zipped) {
  if (zipped) { const JSZip = await JSZipLib(); const z = await JSZip.loadAsync(bytes); const n = Object.keys(z.files).find(x => /\.fb2$/i.test(x)); bytes = await z.file(n).async('uint8array'); }
  let s = utf8Decode(bytes); const enc = /encoding="([^"]+)"/i.exec(s.slice(0, 200)); if (enc && !/utf-?8/i.test(enc[1])) try { s = new TextDecoder(enc[1]).decode(bytes); } catch (e) {}
  const d = parseXml(s);
  const bins = {}; byLocal(d, 'binary').forEach(b => { bins[b.getAttribute('id')] = 'data:' + (b.getAttribute('content-type') || 'image/jpeg') + ';base64,' + b.textContent.replace(/\s+/g, ''); });
  const ti = first(d, 'title-info') || d; const au = first(ti, 'author');
  const meta = { title: ((first(ti, 'book-title') || {}).textContent || 'Sách FB2').trim(), author: au ? [first(au, 'first-name'), first(au, 'middle-name'), first(au, 'last-name')].filter(Boolean).map(x => x.textContent).join(' ') : '' };
  const conv = (node, lvl) => {
    let out = '';
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) { out += esc(c.nodeValue); continue; } if (c.nodeType !== 1) continue;
      const t = (c.localName || c.nodeName).toLowerCase();
      if (t === 'title') out += '<h' + Math.min(6, lvl + 1) + '>' + conv(c, lvl).replace(/<\/?p>/g, ' ') + '</h' + Math.min(6, lvl + 1) + '>';
      else if (t === 'section') out += '<section>' + conv(c, lvl + 1) + '</section>';
      else if (t === 'p' || t === 'v') out += '<p>' + conv(c, lvl) + '</p>';
      else if (t === 'emphasis') out += '<em>' + conv(c, lvl) + '</em>';
      else if (t === 'strong') out += '<strong>' + conv(c, lvl) + '</strong>';
      else if (t === 'empty-line') out += '<br>';
      else if (t === 'epigraph' || t === 'cite') out += '<blockquote>' + conv(c, lvl) + '</blockquote>';
      else if (t === 'poem' || t === 'stanza') out += '<div class="rd-poem">' + conv(c, lvl) + '</div>';
      else if (t === 'image') { const id = (c.getAttribute('l:href') || c.getAttribute('xlink:href') || c.getAttribute('href') || '').replace(/^#/, ''); if (bins[id]) out += '<img src="' + bins[id] + '">'; }
      else if (t === 'a') out += conv(c, lvl);
      else if (t === 'table' || t === 'tr' || t === 'td' || t === 'th' || t === 'sub' || t === 'sup') out += '<' + t + '>' + conv(c, lvl) + '</' + t + '>';
      else out += conv(c, lvl);
    }
    return out;
  };
  const chapters = []; const toc = [];
  byLocal(d, 'body').forEach(body => {
    if (body.getAttribute('name') === 'notes') { chapters.push({ href: 'notes', title: 'Ghi chú', html: conv(body, 1) }); return; }
    const secs = []; for (let c = body.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && (c.localName || '').toLowerCase() === 'section') secs.push(c);
    if (!secs.length) secs.push(body);
    secs.forEach((sc, i) => { const tt = first(sc, 'title'); const title = tt ? tt.textContent.replace(/\s+/g, ' ').trim() : ('Chương ' + (chapters.length + 1)); const href = 'c' + chapters.length; chapters.push({ href, title, html: conv(sc, 1) }); toc.push({ title, href, level: 0 }); });
  });
  return { meta, chapters, toc };
}
export function parseTxt(bytes, name) {
  const s = decodeText(bytes); const chapters = []; let cur = { href: 'c0', title: baseName(name), html: '' };
  s.split(/\r?\n/).forEach(l => { const t = l.trim(); if (/^(chương|chuong|chapter|hồi|quyển|phần)\s+[\divxlc]+/i.test(t) && t.length < 120) { if (cur.html) chapters.push(cur); cur = { href: 'c' + chapters.length, title: t, html: '<h2>' + esc(t) + '</h2>' }; } else if (t) cur.html += '<p>' + esc(t) + '</p>'; });
  chapters.push(cur);
  return { meta: { title: baseName(name), author: '' }, chapters, toc: chapters.map(c => ({ title: c.title, href: c.href, level: 0 })) };
}
export async function parseBook(f) {
  const ext = extOf(f.name);
  if (ext === 'epub') return parseEpub(f.bytes);
  if (ext === 'fb2') return parseFb2(f.bytes, false);
  if (ext === 'fbz' || (ext === 'zip' && /fb2/i.test(f.name))) return parseFb2(f.bytes, true);
  return parseTxt(f.bytes, f.name);
}
export function bookToHtml(book) { return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + esc(book.meta.title) + '</title></head><body><h1>' + esc(book.meta.title) + '</h1>' + (book.meta.author ? '<p><i>' + esc(book.meta.author) + '</i></p>' : '') + book.chapters.map(c => c.html).join('<hr>') + '</body></html>'; }

// ---------------------------------------------------------------- reader UI
const PREF_KEY = 'oos_reader';
function prefs() { try { return Object.assign({ size: 19, font: 'Georgia, serif', theme: 'sepia', mode: 'scroll', line: 1.6 }, JSON.parse(localStorage.getItem(PREF_KEY) || '{}')); } catch (e) { return { size: 19, font: 'Georgia, serif', theme: 'sepia', mode: 'scroll', line: 1.6 }; } }
export async function openReaderDialog(shell) { const fs = await Host.openFiles('.epub,.fb2,.fbz,.txt', false); if (fs[0]) openBook(shell, fs[0]); }
export async function openBook(shell, f) {
  const b = busy('Đang mở sách ' + f.name + '…'); let book;
  try { book = await parseBook(f); } catch (e) { b.done(); toast('Không đọc được sách: ' + e.message, 'err'); return null; }
  b.done(); return showReader(shell, book, f.name);
}
export function showReader(shell, book, fname) {
  const P = prefs(); const key = 'pos_' + (book.meta.title || fname);
  const content = h('div.rd-content');
  content.innerHTML = book.chapters.map((c, i) => '<div class="rd-ch" data-i="' + i + '" data-href="' + esc(c.href) + '">' + c.html + '</div>').join('');
  const page = h('div.rd-page', content);
  const tocList = h('div.rd-toc', book.toc.map(t => { const a = h('button.rd-toc-i', { style: { paddingLeft: (10 + 14 * (t.level || 0)) + 'px' } }, t.title || '—'); a.addEventListener('click', () => { go(t.href); toc.classList.remove('on'); }); return a; }));
  const toc = h('div.rd-side', h('div.rd-side-h', h('b', 'Mục lục'), h('span', { style: { flex: 1 } })), tocList);
  const prog = h('span.rd-prog', '0%');
  const btn = (icn, title, fn) => { const x = h('button.rd-b', { title }, ic(icn, 18)); x.addEventListener('click', fn); return x; };
  const apply = () => {
    root.className = 'reader rd-' + P.theme + (P.mode === 'page' ? ' rd-paged' : '');
    content.style.fontSize = P.size + 'px'; content.style.fontFamily = P.font; content.style.lineHeight = P.line;
    try { localStorage.setItem(PREF_KEY, JSON.stringify(P)); } catch (e) {}
  };
  const scroller = () => page;
  const onScroll = () => {
    const el = scroller(); const pct = P.mode === 'page' ? el.scrollLeft / Math.max(1, el.scrollWidth - el.clientWidth) : el.scrollTop / Math.max(1, el.scrollHeight - el.clientHeight);
    prog.textContent = Math.round(clamp(pct, 0, 1) * 100) + '%'; safeStore(key, pct);
  };
  page.addEventListener('scroll', onScroll);
  const go = href => {
    const [file, frag] = String(href).split('#');
    let target = frag ? content.querySelector('#r-' + (window.CSS && CSS.escape ? CSS.escape(frag) : frag)) : null;
    if (!target) target = $$('.rd-ch', content).find(c => c.getAttribute('data-href') === file || c.getAttribute('data-href').split('/').pop() === file.split('/').pop());
    if (target) { if (P.mode === 'page') page.scrollLeft = Math.floor(target.offsetLeft / page.clientWidth) * page.clientWidth; else page.scrollTop = target.offsetTop - 10; }
  };
  content.addEventListener('click', e => { const a = e.target.closest && e.target.closest('a[data-href]'); if (a) { e.preventDefault(); go(a.getAttribute('data-href')); } });
  const turn = d => { if (P.mode === 'page') page.scrollLeft += d * page.clientWidth; else page.scrollTop += d * (page.clientHeight - 40); };
  let speaking = false;
  const speak = () => {
    if (!window.speechSynthesis) return toast('Thiết bị không hỗ trợ đọc to', 'err');
    if (speaking) { speechSynthesis.cancel(); speaking = false; return; }
    const sel = String(window.getSelection() || '').trim();
    const text = sel || visibleText(); if (!text) return;
    const parts = text.match(/[^.!?\n]{1,220}[.!?\n]*/g) || [text]; speaking = true;
    const voice = speechSynthesis.getVoices().find(v => /vi/i.test(v.lang));
    parts.forEach((p, i) => { const u = new SpeechSynthesisUtterance(p); u.lang = voice ? voice.lang : 'vi-VN'; if (voice) u.voice = voice; if (i === parts.length - 1) u.onend = () => { speaking = false; }; speechSynthesis.speak(u); });
  };
  const visibleText = () => { const r = page.getBoundingClientRect(); let out = ''; $$('p,h1,h2,h3,h4,li,blockquote', content).forEach(p => { const b = p.getBoundingClientRect(); if (b.bottom > r.top && b.top < r.bottom && b.right > r.left && b.left < r.right) out += p.textContent + '\n'; }); return out; };
  const search = () => import('../core/ui.js').then(U => U.promptDlg('Tìm trong sách:', '', 'Tìm').then(q => {
    if (!q) return; const low = q.toLowerCase(); const els = $$('p,h1,h2,h3,h4,li,td', content); const start = els.findIndex(el => el.getBoundingClientRect().top > 60);
    const order = els.slice(Math.max(0, start)).concat(els.slice(0, Math.max(0, start)));
    const hit = order.find(el => el.textContent.toLowerCase().indexOf(low) >= 0);
    if (!hit) return toast('Không tìm thấy "' + q + '"');
    $$('.rd-hit', content).forEach(x => x.classList.remove('rd-hit')); hit.classList.add('rd-hit');
    if (P.mode === 'page') page.scrollLeft = Math.floor(hit.offsetLeft / page.clientWidth) * page.clientWidth; else page.scrollTop = hit.offsetTop - 80;
  }));
  const close = () => { try { speechSynthesis && speechSynthesis.cancel(); } catch (e) {} document.removeEventListener('keydown', onKey, true); if (root.parentNode) root.parentNode.removeChild(root); };
  const onKey = e => { if (!root.parentNode) return; const k = e.key || ''; if (k === 'Escape') { e.preventDefault(); close(); } else if (k === 'ArrowRight' || k === 'PageDown' || k === ' ') { if (P.mode === 'page' || k !== ' ') { e.preventDefault(); turn(1); } } else if (k === 'ArrowLeft' || k === 'PageUp') { e.preventDefault(); turn(-1); } };
  document.addEventListener('keydown', onKey, true);
  const bar = h('div.rd-bar',
    btn('arrow-left', 'Đóng (Esc)', close), btn('list', 'Mục lục', () => toc.classList.toggle('on')),
    h('div.rd-title', h('b', book.meta.title), book.meta.author ? h('span.muted', ' · ' + book.meta.author) : null),
    btn('search', 'Tìm trong sách', search),
    btn('a-arrow-down', 'Chữ nhỏ hơn', () => { P.size = Math.max(12, P.size - 1); apply(); }), btn('a-arrow-up', 'Chữ lớn hơn', () => { P.size = Math.min(40, P.size + 1); apply(); }),
    btn('type', 'Đổi phông chữ', () => { const F = ['Georgia, serif', '"Times New Roman", serif', 'Arial, sans-serif', '"Segoe UI", Roboto, sans-serif', '"Courier New", monospace']; P.font = F[(F.indexOf(P.font) + 1) % F.length]; apply(); }),
    btn('sun-moon', 'Giao diện: sáng / giấy / tối', () => { const T = ['light', 'sepia', 'dark']; P.theme = T[(T.indexOf(P.theme) + 1) % T.length]; apply(); }),
    btn('book-open', 'Chế độ: cuộn / lật trang', () => { const pct = parseFloat(prog.textContent) / 100; P.mode = P.mode === 'page' ? 'scroll' : 'page'; apply(); setTimeout(() => { if (P.mode === 'page') page.scrollLeft = Math.round(pct * (page.scrollWidth - page.clientWidth) / page.clientWidth) * page.clientWidth; else page.scrollTop = pct * (page.scrollHeight - page.clientHeight); }, 50); }),
    btn('volume-2', 'Đọc to (chọn đoạn hoặc trang đang xem)', speak),
    btn('file-text', 'Mở trong Word để chỉnh sửa', () => { close(); shell.openFile({ name: (book.meta.title || 'sach').replace(/[\\/:*?"<>|]/g, '_') + '.html', bytes: utf8Encode(bookToHtml(book)), here: true }); }),
    prog);
  const nav = h('div.rd-nav', btn('chevron-left', 'Trang trước', () => turn(-1)), btn('chevron-right', 'Trang sau', () => turn(1)));
  const root = h('div.reader', bar, h('div.rd-body', toc, page), nav);
  document.body.appendChild(root); apply();
  // restore position
  setTimeout(() => { let pct = 0; pct = +safeStore(key) || 0; if (pct > 0) { if (P.mode === 'page') page.scrollLeft = Math.round(pct * (page.scrollWidth - page.clientWidth) / page.clientWidth) * page.clientWidth; else page.scrollTop = pct * (page.scrollHeight - page.clientHeight); } onScroll(); }, 80);
  return { root, book, close, go, turn };
}
