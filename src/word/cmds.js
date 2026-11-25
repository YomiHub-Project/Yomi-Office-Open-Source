// Word commands: formatting via execCommand + custom, styles, insert (table/image/link/symbol/eq), find, comments, track changes, print
import { h, ic, esc, $, $$, keyName, chord, clamp, uid, readFileAsDataUrl, bytesToDataUrl, fmtDate, fmtTime } from '../core/util.js';
import { toast, showMenu, contextMenu, dialog, promptDlg, confirmDlg, formDlg, colorPicker } from '../core/ui.js';
import { settings } from '../core/store.js';
import Host from '../core/host.js';

const ec = (cmd, val) => { try { document.execCommand(cmd, false, val); } catch (e) {} };
export function focusBody(doc) { doc.body.focus(); if (!window.getSelection().rangeCount || !doc.body.contains(window.getSelection().anchorNode)) doc.restoreCaret(); }
export function withUndo(doc, label, fn) { focusBody(doc); doc.pushHistory(); fn(); doc.pushHistory(); doc.markDirty(); doc.updatePages(); doc.refreshUi(); }
function run(doc, fn) { focusBody(doc); fn(); doc.markDirty(); doc.scheduleHistory(); doc.refreshUi(); doc.debouncedPages && doc.debouncedPages(); }

// ---------------------------------------------------------------- query current formatting
export function queryState(doc) {
  const st = { bold: false, italic: false, underline: false, strike: false, ul: false, ol: false, alignL: false, alignC: false, alignR: false, alignJ: false, sub: false, sup: false };
  try {
    st.bold = document.queryCommandState('bold'); st.italic = document.queryCommandState('italic'); st.underline = document.queryCommandState('underline'); st.strike = document.queryCommandState('strikeThrough');
    st.ul = document.queryCommandState('insertUnorderedList'); st.ol = document.queryCommandState('insertOrderedList');
    st.alignL = document.queryCommandState('justifyLeft'); st.alignC = document.queryCommandState('justifyCenter'); st.alignR = document.queryCommandState('justifyRight'); st.alignJ = document.queryCommandState('justifyFull');
    st.sub = document.queryCommandState('subscript'); st.sup = document.queryCommandState('superscript');
    st.font = (document.queryCommandValue('fontName') || '').replace(/^['"]|['"]$/g, '').split(',')[0];
    st.fore = document.queryCommandValue('foreColor');
  } catch (e) {}
  const el = currentBlock(doc);
  st.style = el ? blockStyleName(el) : 'Normal';
  st.size = pxToPt(getComputedStyle(blockInline(doc) || doc.body).fontSize);
  return st;
}
function pxToPt(px) { return Math.round(parseFloat(px) * 72 / 96 * 2) / 2; }
export function currentBlock(doc) { let n = window.getSelection().anchorNode; if (!n) return doc.body.firstElementChild; if (n.nodeType === 3) n = n.parentElement; while (n && n !== doc.body && !/^(P|H[1-6]|LI|BLOCKQUOTE|PRE|TD|TH|DIV)$/.test(n.tagName)) n = n.parentElement; return n && n !== doc.body ? n : doc.body.firstElementChild; }
function blockInline(doc) { let n = window.getSelection().anchorNode; if (n && n.nodeType === 3) n = n.parentElement; return n; }
function blockStyleName(el) { if (!el) return 'Normal'; if (/^H[1-6]$/.test(el.tagName)) return 'Heading ' + el.tagName[1]; if (el.tagName === 'BLOCKQUOTE') return 'Quote'; if (el.tagName === 'PRE') return 'Code'; if (el.getAttribute && el.getAttribute('data-style')) return el.getAttribute('data-style'); if (el.classList && el.classList.contains('wd-title')) return 'Title'; if (el.classList && el.classList.contains('wd-subtitle')) return 'Subtitle'; return 'Normal'; }

// ---------------------------------------------------------------- basic formatting
export function toggle(doc, cmd) { run(doc, () => ec(cmd)); }
export function bold(doc) { toggle(doc, 'bold'); } export function italic(doc) { toggle(doc, 'italic'); } export function underline(doc) { toggle(doc, 'underline'); } export function strike(doc) { toggle(doc, 'strikeThrough'); }
export function sub(doc) { run(doc, () => ec('subscript')); } export function sup(doc) { run(doc, () => ec('superscript')); }
export function align(doc, a) { run(doc, () => ec('justify' + a)); }
export function list(doc, kind) { run(doc, () => ec(kind === 'ol' ? 'insertOrderedList' : 'insertUnorderedList')); }
export function indent(doc) { run(doc, () => ec('indent')); } export function outdent(doc) { run(doc, () => ec('outdent')); }
export function setFont(doc, f) { run(doc, () => { ec('fontName', f); }); }
export function setSize(doc, pt) { run(doc, () => applyInline(doc, el => { el.style.fontSize = pt + 'pt'; })); }
export function sizeStep(doc, dir) { const cur = pxToPt(getComputedStyle(blockInline(doc) || doc.body).fontSize); const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72]; let n = dir > 0 ? sizes.find(x => x > cur) : sizes.slice().reverse().find(x => x < cur); setSize(doc, n || (dir > 0 ? cur + 2 : Math.max(1, cur - 1))); }
export function foreColor(doc, c) { run(doc, () => ec('foreColor', c)); }
export function highlight(doc, c) { run(doc, () => { if (!ec('hiliteColor', c)) ec('backColor', c); }); }
export function clearFormat(doc) { run(doc, () => { ec('removeFormat'); const b = currentBlock(doc); if (b && b !== doc.body) { b.removeAttribute('style'); b.removeAttribute('data-style'); b.className = ''; if (/^H[1-6]|BLOCKQUOTE|PRE$/.test(b.tagName)) ec('formatBlock', 'p'); } }); }
// apply a style callback to the selection by wrapping in span
function applyInline(doc, cb) {
  const sel = window.getSelection(); if (!sel.rangeCount) return;
  const r = sel.getRangeAt(0);
  if (r.collapsed) { const block = currentBlock(doc); if (block) cb(block); return; }
  const span = h('span'); try { r.surroundContents(span); } catch (e) { ec('fontName', 'Arial'); const sp = doc.body.querySelector('font[face="Arial"]'); if (sp) { const s2 = h('span'); while (sp.firstChild) s2.appendChild(sp.firstChild); sp.parentNode.replaceChild(s2, sp); cb(s2); } return; }
  cb(span);
  sel.removeAllRanges(); const nr = document.createRange(); nr.selectNodeContents(span); sel.addRange(nr);
}
export function lineSpacing(doc, v) { run(doc, () => { const blocks = selectedBlocks(doc); blocks.forEach(b => { b.style.lineHeight = v; }); }); }
export function spaceBefore(doc, pt) { run(doc, () => selectedBlocks(doc).forEach(b => { b.style.marginTop = pt + 'pt'; })); }
export function spaceAfter(doc, pt) { run(doc, () => selectedBlocks(doc).forEach(b => { b.style.marginBottom = pt + 'pt'; })); }
export function selectedBlocks(doc) {
  const sel = window.getSelection(); if (!sel.rangeCount) { const b = currentBlock(doc); return b ? [b] : []; }
  const r = sel.getRangeAt(0); const out = [];
  $$('p,h1,h2,h3,h4,h5,h6,li,blockquote,pre', doc.body).forEach(b => { if (r.intersectsNode ? r.intersectsNode(b) : true) { const br = document.createRange(); br.selectNode(b); if (r.compareBoundaryPoints(Range.END_TO_START, br) < 0 && r.compareBoundaryPoints(Range.START_TO_END, br) > 0) out.push(b); } });
  if (!out.length) { const b = currentBlock(doc); if (b) out.push(b); }
  return out;
}

// ---------------------------------------------------------------- styles
export const STYLES = [
  ['Normal', 'Bình thường', {}], ['No Spacing', 'Không giãn cách', { ns: 1 }], ['Heading 1', 'Đề mục 1', { tag: 'h1' }], ['Heading 2', 'Đề mục 2', { tag: 'h2' }], ['Heading 3', 'Đề mục 3', { tag: 'h3' }], ['Heading 4', 'Đề mục 4', { tag: 'h4' }],
  ['Title', 'Tiêu đề', { cls: 'wd-title' }], ['Subtitle', 'Phụ đề', { cls: 'wd-subtitle' }], ['Quote', 'Trích dẫn', { tag: 'blockquote' }], ['Intense Quote', 'Trích dẫn nổi bật', { tag: 'blockquote', cls: 'wd-intense' }], ['Code', 'Mã', { tag: 'pre' }], ['List Paragraph', 'Đoạn danh sách', { cls: 'wd-listpara' }]
];
export function applyStyle(doc, name) {
  const def = STYLES.find(s => s[0] === name) || STYLES[0];
  run(doc, () => {
    const blocks = selectedBlocks(doc);
    blocks.forEach(b => {
      b.className = ''; b.removeAttribute('data-style'); if (def[2].ns) b.style.margin = '0';
      if (def[2].tag) { if (b.tagName.toLowerCase() !== def[2].tag) { const nb = h(def[2].tag); nb.innerHTML = b.innerHTML; b.parentNode.replaceChild(nb, b); b = nb; } }
      else if (/^H[1-6]|BLOCKQUOTE|PRE$/.test(b.tagName)) { const nb = h('p'); nb.innerHTML = b.innerHTML; b.parentNode.replaceChild(nb, b); b = nb; }
      if (def[2].cls) b.classList.add(def[2].cls);
      if (name !== 'Normal' && name !== 'No Spacing') b.setAttribute('data-style', name);
    });
  });
}
export function stylesGallery(doc) {
  const cur = queryState(doc).style;
  return STYLES.map(s => ({ title: s[1], label: s[1], active: () => queryState(doc).style === s[0], preview: () => { const t = h('div.style-prev' + (s[2].tag ? '.s-' + s[2].tag : '') + (s[2].cls ? '.' + s[2].cls : ''), s[1]); return t; }, run: () => applyStyle(doc, s[0]) }));
}

// ---------------------------------------------------------------- clipboard
export function onPaste(doc, e) {
  const cd = e.clipboardData; if (!cd) return;
  if (cd.items) for (const it of Array.prototype.slice.call(cd.items)) if (/^image\//.test(it.type)) { e.preventDefault(); const f = it.getAsFile(); readFileAsDataUrl(f).then(u => insertImageUrl(doc, u)); return; }
  const html = cd.getData('text/html');
  e.preventDefault();
  if (html && !/pasteText/.test(doc._pasteMode || '')) { doc.pushHistory(); ec('insertHTML', sanitize(html)); }
  else { const t = cd.getData('text/plain'); doc.pushHistory(); ec('insertText', t); }
  doc.markDirty(); doc.updatePages(); doc.refreshUi();
}
export function onCopy(doc, e) { /* default copy keeps rich text */ }
export function onDrop(doc, e) { if (!e.dataTransfer.files.length) return; e.preventDefault(); const f = e.dataTransfer.files[0]; if (/^image\//.test(f.type)) readFileAsDataUrl(f).then(u => insertImageUrl(doc, u)); }
export function pasteSpecial(doc) { showMenu([{ label: 'Giữ định dạng nguồn', run: () => { doc._pasteMode = ''; ec('paste'); } }, { label: 'Dán văn bản thuần (không định dạng)', run: () => { doc._pasteMode = 'pasteText'; navigator.clipboard && navigator.clipboard.readText ? navigator.clipboard.readText().then(t => { ec('insertText', t); }) : toast('Dùng Ctrl+Shift+V'); } }], { x: 300, y: 150 }); }
export function sanitize(html) {
  const d = document.createElement('div'); d.innerHTML = html;
  d.querySelectorAll('script,style,meta,link,title,o\\:p,xml').forEach(e => e.remove());
  d.querySelectorAll('*').forEach(el => {
    for (const a of Array.prototype.slice.call(el.attributes)) { if (/^on/i.test(a.name) || (a.name === 'href' && /^javascript:/i.test(a.value)) || a.name === 'class' && /Mso/.test(a.value)) el.removeAttribute(a.name); if (a.name === 'class' && /Mso/.test(a.value)) el.removeAttribute('class'); }
    if (el.style) { el.style.removeProperty('position'); el.style.removeProperty('font-family'); if (/Mso/.test(el.getAttribute('style') || '')) el.removeAttribute('style'); }
  });
  return d.innerHTML;
}
export function clearSelectionArtifacts(doc) { $$('.wd-find-hit', doc.body).forEach(e => { const p = e.parentNode; while (e.firstChild) p.insertBefore(e.firstChild, e); p.removeChild(e); p.normalize(); }); }

// ---------------------------------------------------------------- keyboard
export function onKey(doc, e) {
  const c = chord(e), k = keyName(e);
  const map = {
    'alt+f8': () => import('../excel/vba/index.js').then(m => m.macrosDialog(doc)), 'alt+f11': () => import('../excel/vba/index.js').then(m => m.openEditor(doc)),
    'ctrl+b': () => bold(doc), 'ctrl+i': () => italic(doc), 'ctrl+u': () => underline(doc), 'ctrl+5': () => strike(doc),
    'ctrl+l': () => align(doc, 'Left'), 'ctrl+e': () => align(doc, 'Center'), 'ctrl+r': () => align(doc, 'Right'), 'ctrl+j': () => align(doc, 'Full'),
    'ctrl+]': () => sizeStep(doc, 1), 'ctrl+[': () => sizeStep(doc, -1), 'ctrl+shift+.': () => sizeStep(doc, 1), 'ctrl+shift+,': () => sizeStep(doc, -1),
    'ctrl+shift+n': () => applyStyle(doc, 'Normal'), 'ctrl+alt+1': () => applyStyle(doc, 'Heading 1'), 'ctrl+alt+2': () => applyStyle(doc, 'Heading 2'), 'ctrl+alt+3': () => applyStyle(doc, 'Heading 3'),
    'ctrl+k': () => linkDialog(doc), 'ctrl+ ': () => clearFormat(doc), 'ctrl+shift+c': () => { doc._fmtClip = captureFormat(doc); toast('Đã sao chép định dạng'); }, 'ctrl+shift+v': () => applyCapturedFormat(doc),
    'ctrl+shift+l': () => list(doc, 'ul'), 'ctrl+1': () => lineSpacing(doc, '1'), 'ctrl+2': () => lineSpacing(doc, '2'), 'ctrl+shift+8': () => toggleMarks(doc),
    'ctrl+alt+m': () => addComment(doc), 'ctrl+shift+e': () => toggleTrack(doc), 'alt+shift+d': () => insertField(doc, 'date'), 'f7': () => import('../ai/ai.js').then(m => m.proofread(doc)),
    'ctrl+z': () => doc.undo(), 'ctrl+y': () => doc.redo(), 'ctrl+shift+z': () => doc.redo(), 'ctrl+g': () => gotoPageDialog(doc)
  };
  if (map[c]) { e.preventDefault(); map[c](); return; }
  if (c === 'ctrl+enter') { e.preventDefault(); insertPageBreak(doc); return; }
  if (k === 'Tab' && !e.ctrlKey) { const b = currentBlock(doc); if (b && b.tagName === 'LI') return; if (b && b.closest('table')) return; e.preventDefault(); ec('insertText', '\t'); return; }
  if (k === 'Enter' && !e.shiftKey) { const b = currentBlock(doc); if (b && /^H[1-6]$/.test(b.tagName) && atBlockEnd()) { setTimeout(() => { const nb = currentBlock(doc); if (nb && /^H[1-6]$/.test(nb.tagName)) applyStyle(doc, 'Normal'); }, 0); } }
  if (c === 'ctrl+z' || c === 'ctrl+y') return;
}
function atBlockEnd() { const s = window.getSelection(); if (!s.rangeCount) return false; const r = s.getRangeAt(0); return r.endOffset === (r.endContainer.length || r.endContainer.childNodes.length); }
export function caretToStart(doc) { const r = document.createRange(); const f = doc.body.firstChild || doc.body; r.setStart(f, 0); r.collapse(true); const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); }
function captureFormat(doc) { const st = queryState(doc); const el = blockInline(doc); const cs = el ? getComputedStyle(el) : {}; return { b: st.bold, i: st.italic, u: st.underline, s: st.strike, color: cs.color, size: st.size, font: st.font }; }
function applyCapturedFormat(doc) { const f = doc._fmtClip; if (!f) return; run(doc, () => { if (f.b) ec('bold'); if (f.i) ec('italic'); if (f.u) ec('underline'); if (f.color) ec('foreColor', f.color); if (f.size) setSize(doc, f.size); if (f.font) ec('fontName', f.font); }); }
export function formatPainter(doc) { doc._fmtClip = captureFormat(doc); toast('Chọn văn bản đích để áp dụng định dạng'); const once = () => { applyCapturedFormat(doc); doc.body.removeEventListener('mouseup', once); }; setTimeout(() => doc.body.addEventListener('mouseup', once, { once: true }), 50); }
export function autoFormat(doc) {
  // auto bullet/number & smart quotes handled lightly
}
export function normalizeDoc(doc) { $$('div', doc.body).forEach(d => { if (!d.querySelector('p,div,h1,h2,h3,ul,ol,table')) { const p = h('p'); p.innerHTML = d.innerHTML || '<br>'; d.parentNode.replaceChild(p, d); } }); if (!doc.body.firstChild) doc.body.appendChild(h('p', { html: '<br>' })); }

// ---------------------------------------------------------------- insert
// v0.0.8: page break is always a top-level block — the paragraph is split at the caret (like Word), so pagination can
// start the following text on a new page. blank=true inserts a whole empty page (Chèn → Trang trống).
export function insertPageBreak(doc, blank) {
  withUndo(doc, blank ? 'Trang trống' : 'Ngắt trang', () => {
    const body = doc.body; const sel = window.getSelection();
    const mkBreak = () => h('div.wd-pagebreak', { contenteditable: 'false' });
    const emptyP = () => { const p = document.createElement('p'); p.innerHTML = '<br>'; return p; };
    let top = null, rng = sel.rangeCount ? sel.getRangeAt(0) : null;
    if (rng && body.contains(rng.startContainer)) { top = rng.startContainer.nodeType === 3 ? rng.startContainer.parentNode : rng.startContainer; while (top && top.parentNode !== body) top = top.parentNode; if (top === body) top = null; }
    let after = null;
    if (top && /^(P|H[1-6]|BLOCKQUOTE|PRE|DIV)$/.test(top.tagName) && !top.classList.contains('wd-pagebreak') && rng) {
      rng.deleteContents();
      const tail = document.createRange(); tail.setStart(rng.startContainer, rng.startOffset); tail.setEndAfter(top.lastChild || top);
      const frag = tail.extractContents();
      const clone = top.cloneNode(false); clone.appendChild(frag);
      if (!clone.textContent.trim() && !clone.querySelector('img,table,svg')) clone.innerHTML = '<br>';
      if (!top.textContent.trim() && !top.querySelector('img,table,svg')) top.innerHTML = '<br>';
      after = clone;
    }
    const ref = top ? top.nextSibling : null;
    const ins = n => body.insertBefore(n, ref);
    if (!top) { ins(mkBreak()); }
    else { ins(mkBreak()); }
    let caretP = after || emptyP();
    if (blank) { const bp = emptyP(); ins(bp); ins(mkBreak()); }
    ins(caretP);
    const r = document.createRange(); r.setStart(caretP, 0); r.collapse(true); sel.removeAllRanges(); sel.addRange(r);
  });
}
export function insertText(doc, t) { run(doc, () => ec('insertText', t)); }
export function insertHtml(doc, html) { run(doc, () => ec('insertHTML', html)); }
export async function insertImage(doc) { const files = await Host.openFiles('image/*'); if (!files.length) return; const f = files[0]; insertImageUrl(doc, bytesToDataUrl(f.bytes, /\.png$/i.test(f.name) ? 'image/png' : /\.gif$/i.test(f.name) ? 'image/gif' : /\.svg$/i.test(f.name) ? 'image/svg+xml' : 'image/jpeg')); }
export function insertImageUrl(doc, url) { const img = new Image(); img.onload = () => { const w = Math.min(img.naturalWidth, doc.pageW - doc.marg[1] - doc.marg[3]); run(doc, () => ec('insertHTML', '<img src="' + url + '" style="width:' + w + 'px;max-width:100%" alt="">')); }; img.src = url; }
export function selectImage(doc, img) { doc._selImg = img; $$('img.wd-selimg', doc.body).forEach(i => i.classList.remove('wd-selimg')); img.classList.add('wd-selimg'); if (doc.ribbon) { doc._ctxImg = true; doc.updateContextualTabs && doc.updateContextualTabs(); } }
export async function insertTableDialog(doc) {
  const v = await tableGrid(doc); if (!v) return;
  let html = '<table class="wd-table"><tbody>';
  for (let r = 0; r < v.rows; r++) { html += '<tr>'; for (let c = 0; c < v.cols; c++) html += (r === 0 && v.header ? '<th>' : '<td>') + '<br>' + (r === 0 && v.header ? '</th>' : '</td>'); html += '</tr>'; }
  html += '</tbody></table><p><br></p>';
  insertHtml(doc, html);
}
function tableGrid(doc) {
  return new Promise(res => {
    const grid = h('div.tbl-grid'); const lbl = h('div.tbl-lbl', 'Chèn bảng');
    let sel = { r: 0, c: 0 };
    for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) { const cell = h('div.tbl-cell', { dataset: { r, c } }); cell.addEventListener('mouseenter', () => { sel = { r, c }; $$('.tbl-cell', grid).forEach(x => x.classList.toggle('on', +x.dataset.r <= r && +x.dataset.c <= c)); lbl.textContent = (c + 1) + ' × ' + (r + 1) + ' Bảng'; }); cell.addEventListener('click', () => { close({ rows: r + 1, cols: c + 1, header: true }); }); grid.appendChild(cell); }
    const header = h('label.chk'); const hc = h('input', { type: 'checkbox' }); hc.checked = true; header.appendChild(hc); header.appendChild(h('span', 'Hàng đầu là tiêu đề'));
    let close;
    dialog({ title: 'Chèn bảng', width: 320, body: h('div', lbl, grid, header, h('div', h('button.btn', { onclick: () => import('./cmds.js').then(() => drawDlg()) }, 'Tùy chỉnh số hàng/cột…')), h('div.hint', 'Rê chuột để chọn kích thước, hoặc bấm nút tùy chỉnh.')), onOpen: (b, cl) => { close = v => { cl(); res(v ? Object.assign(v, { header: hc.checked }) : null); }; }, buttons: [{ label: 'Hủy', value: null }] }).then(() => res(null));
    async function drawDlg() { const v = await formDlg('Chèn bảng', [{ key: 'rows', type: 'number', label: 'Số hàng', value: 3 }, { key: 'cols', type: 'number', label: 'Số cột', value: 3 }, { key: 'header', type: 'checkbox', label: 'Hàng tiêu đề', value: true }]); if (v) { $$('.dlg-back').forEach(d => d.remove()); res(v); } }
  });
}
export async function linkDialog(doc) {
  const sel = window.getSelection(); const text = sel.toString();
  const existing = sel.anchorNode && sel.anchorNode.parentElement && sel.anchorNode.parentElement.closest('a');
  const v = await formDlg('Chèn liên kết', [{ key: 'text', label: 'Văn bản hiển thị', value: text || (existing ? existing.textContent : '') }, { key: 'url', label: 'Địa chỉ', value: existing ? existing.href : 'https://' }]);
  if (!v) return;
  run(doc, () => { const url = /^[\w.+-]+@[\w-]+\.[\w.]+$/.test(v.url) ? 'mailto:' + v.url : v.url; if (text) ec('createLink', url); else ec('insertHTML', '<a href="' + esc(url) + '">' + esc(v.text || url) + '</a>'); });
  $$('a', doc.body).forEach(a => { a.title = 'Ctrl + bấm để mở'; });
}
export function insertField(doc, type) {
  const map = { date: fmtDate(), time: fmtTime(), datetime: fmtDate() + ' ' + fmtTime(), page: '1', pages: String(doc.pageCount || 1), author: settings().userName, filename: doc.name };
  run(doc, () => ec('insertText', map[type] || ''));
}
export function insertSymbol(doc, ch) { run(doc, () => ec('insertText', ch)); }
export function insertEquation(doc, latex) { import('../math/math.js').then(m => m.renderInline(latex).then(svg => insertHtml(doc, '<span class="wd-eq" contenteditable="false" data-latex="' + esc(latex) + '">' + svg + '</span>&nbsp;'))); }
export function insertHeaderFooter(doc) { toast('Đầu trang/Chân trang: thiết lập ở Bố trí > Thiết lập trang hoặc khi In'); }
export function insertTOC(doc) {
  const heads = $$('h1,h2,h3', doc.body);
  if (!heads.length) return toast('Chưa có đề mục (Heading) nào để tạo mục lục', 'err');
  let html = '<div class="wd-toc" contenteditable="false"><div class="wd-toc-title">Mục lục</div>';
  heads.forEach((hd, i) => { const id = 'toc' + i; hd.id = id; html += '<div class="wd-toc-item lvl' + hd.tagName[1] + '"><span>' + esc(hd.textContent) + '</span><span class="wd-toc-dots"></span><span>' + (i + 1) + '</span></div>'; });
  html += '</div><p><br></p>';
  const first = doc.body.firstChild; doc.pushHistory(); const div = h('div'); div.innerHTML = html; while (div.firstChild) doc.body.insertBefore(div.firstChild, first); doc.pushHistory(); doc.markDirty(); doc.updatePages();
}
export function insertCover(doc) {
  const html = '<div class="wd-cover" contenteditable="false"></div>';
  run(doc, () => {});
  const cover = h('div.wd-cover');
  cover.innerHTML = '<div class="wd-cover-bar"></div><h1 class="wd-cover-title" contenteditable="true">Tiêu đề tài liệu</h1><div class="wd-cover-sub" contenteditable="true">Phụ đề hoặc mô tả</div><div class="wd-cover-meta" contenteditable="true">' + esc(settings().userName) + ' · ' + fmtDate() + '</div><div class="wd-pagebreak"></div>';
  doc.pushHistory(); doc.body.insertBefore(cover, doc.body.firstChild); doc.body.insertBefore(h('p', { html: '<br>' }), cover.nextSibling); doc.pushHistory(); doc.updatePages(); doc.markDirty();
}

// ---------------------------------------------------------------- tables (context)
export function tableOp(doc, op) {
  const sel = window.getSelection(); let cell = sel.anchorNode; if (cell && cell.nodeType === 3) cell = cell.parentElement; cell = cell && cell.closest('td,th'); if (!cell) return;
  const tr = cell.closest('tr'), table = cell.closest('table'); const ci = Array.prototype.indexOf.call(tr.children, cell);
  run(doc, () => {
    if (op === 'rowAbove' || op === 'rowBelow') { const nr = tr.cloneNode(true); Array.prototype.forEach.call(nr.children, c => c.innerHTML = '<br>'); tr.parentNode.insertBefore(nr, op === 'rowBelow' ? tr.nextSibling : tr); }
    else if (op === 'colLeft' || op === 'colRight') { Array.prototype.forEach.call(table.rows, row => { const ref = row.children[ci]; const nc = h(ref.tagName.toLowerCase(), { html: '<br>' }); row.insertBefore(nc, op === 'colRight' ? ref.nextSibling : ref); }); }
    else if (op === 'delRow') { if (table.rows.length > 1) tr.remove(); }
    else if (op === 'delCol') { Array.prototype.forEach.call(table.rows, row => { if (row.children[ci]) row.children[ci].remove(); }); }
    else if (op === 'delTable') { table.remove(); }
    else if (op === 'header') { Array.prototype.forEach.call(table.rows[0].children, c => { const nc = h('th', { html: c.innerHTML }); c.parentNode.replaceChild(nc, c); }); }
    else if (/^align/.test(op)) cell.style.textAlign = op.slice(5).toLowerCase();
    else if (op === 'mergeRight') { const nx = cell.nextElementSibling; if (nx) { cell.colSpan = (cell.colSpan || 1) + (nx.colSpan || 1); cell.innerHTML += ' ' + nx.innerHTML; nx.remove(); } }
  });
}
export function tableStyle(doc, style) { const cell = currentCell(); if (!cell) return; const table = cell.closest('table'); run(doc, () => { table.className = 'wd-table ' + style; }); }
function currentCell() { let n = window.getSelection().anchorNode; if (n && n.nodeType === 3) n = n.parentElement; return n && n.closest ? n.closest('td,th') : null; }

// ---------------------------------------------------------------- find & replace
export function findDialog(doc, replace) {
  const fi = h('input.inp', { value: doc._lastFind || '' }); const ri = h('input.inp');
  const mc = h('input', { type: 'checkbox' }); const whole = h('input', { type: 'checkbox' });
  const info = h('div.muted');
  let hits = [], idx = -1;
  const clearHits = () => { $$('.wd-find-hit', doc.body).forEach(e => { const p = e.parentNode; while (e.firstChild) p.insertBefore(e.firstChild, e); p.removeChild(e); p.normalize(); }); };
  const doFind = () => {
    clearHits(); hits = []; idx = -1; const q = fi.value; if (!q) { info.textContent = ''; return; }
    doc._lastFind = q;
    const re = new RegExp((whole.checked ? '\\b' : '') + q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + (whole.checked ? '\\b' : ''), mc.checked ? 'g' : 'gi');
    const walker = document.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT, null);
    const texts = []; let n; while (n = walker.nextNode()) if (n.parentElement.closest('.wd-toc,.wd-pagebreak') == null) texts.push(n);
    texts.forEach(tn => { let m; const parent = tn.parentNode; const matches = []; while ((m = re.exec(tn.data)) !== null) { matches.push([m.index, m.index + m[0].length]); if (m.index === re.lastIndex) re.lastIndex++; } if (!matches.length) return; let last = 0; const frag = document.createDocumentFragment(); matches.forEach(([s, e]) => { if (s > last) frag.appendChild(document.createTextNode(tn.data.slice(last, s))); const span = h('span.wd-find-hit'); span.textContent = tn.data.slice(s, e); frag.appendChild(span); hits.push(span); last = e; }); if (last < tn.data.length) frag.appendChild(document.createTextNode(tn.data.slice(last))); parent.replaceChild(frag, tn); });
    info.textContent = hits.length + ' kết quả';
    if (hits.length) next();
  };
  const next = () => { if (!hits.length) return; idx = (idx + 1) % hits.length; hits.forEach(hitEl => hitEl.classList.remove('cur')); const el = hits[idx]; el.classList.add('cur'); el.scrollIntoView({ block: 'center' }); info.textContent = (idx + 1) + ' / ' + hits.length; };
  const replaceOne = () => { if (idx < 0 || !hits[idx]) return next(); doc.pushHistory(); hits[idx].outerHTML = esc(ri.value); doc.markDirty(); doFind(); };
  const replaceAll = () => { if (!fi.value) return; doc.pushHistory(); doFind(); const n = hits.length; hits.forEach(hitEl => { hitEl.outerHTML = esc(ri.value); }); clearHits(); doc.markDirty(); doc.updatePages(); toast('Đã thay thế ' + n + ' kết quả', 'ok'); };
  const btns = [{ label: 'Tìm tiếp', primary: true, run: () => { if (!hits.length) doFind(); else next(); return false; } }];
  if (replace) btns.push({ label: 'Thay thế', run: () => { replaceOne(); return false; } }, { label: 'Thay thế tất cả', run: () => { replaceAll(); return false; } });
  btns.push({ label: 'Đóng', run: () => { clearHits(); return true; } });
  fi.addEventListener('keydown', e => { if (keyName(e) === 'Enter') { e.preventDefault(); if (!hits.length) doFind(); else next(); } });
  dialog({ title: replace ? 'Tìm và Thay thế' : 'Tìm kiếm', width: 440, body: h('div.form', h('label.lbl', 'Tìm:'), fi, replace ? h('div', h('label.lbl', 'Thay bằng:'), ri) : null, h('label.chk', mc, h('span', 'Phân biệt hoa/thường')), h('label.chk', whole, h('span', 'Chỉ từ nguyên')), info), buttons: btns, modeless: true });
}
export async function gotoPageDialog(doc) { const v = await promptDlg('Đi tới trang số:', String(doc.curPage || 1), 'Đi tới'); if (v == null) return; doc.updatePages(); doc.gotoPage(clamp(parseInt(v) || 1, 1, doc.pageCount || 1)); }
export function wordCountDialog(doc) { const st = doc.stats(); dialog({ title: 'Thống kê số từ', width: 360, body: h('table.kv', Object.keys(st).map(k => h('tr', h('td', k), h('td', String(st[k]))))), buttons: [{ label: 'Đóng', primary: true }] }); }

// ---------------------------------------------------------------- comments & track changes
export function addComment(doc) {
  const sel = window.getSelection(); const text = sel.toString(); if (!text) return toast('Chọn đoạn văn bản cần chú thích trước', 'err');
  promptDlg('Chú thích cho "' + text.slice(0, 40) + '":', '', 'Chú thích mới', { multiline: true }).then(v => {
    if (!v) return;
    const id = uid('cm'); doc.pushHistory();
    const span = h('span.wd-comment', { dataset: { cid: id } });
    const r = sel.getRangeAt(0); try { r.surroundContents(span); } catch (e) { ec('insertHTML', '<span class="wd-comment" data-cid="' + id + '">' + esc(text) + '</span>'); }
    doc.comments.push({ id, text: v, author: settings().userName, time: new Date().toLocaleString(), anchor: text.slice(0, 60), resolved: false });
    doc.pushHistory(); doc.markDirty(); commentsPane(doc);
  });
}
export function commentsPane(doc) {
  const old = $('.wd-pane.cmt', doc.el); if (old) { old.remove(); return; }
  const pane = h('div.wd-pane.cmt', h('div.pane-head', h('span.t', 'Chú thích'), h('button.pane-x', { onclick: () => pane.remove() }, ic('x', 16))));
  const body = h('div.pane-body'); pane.appendChild(body);
  const draw = () => { body.innerHTML = ''; if (!doc.comments.length) body.appendChild(h('div.muted', 'Chưa có chú thích. Chọn văn bản rồi bấm Chú thích Mới.')); doc.comments.forEach(cm => { const card = h('div.cm-card' + (cm.resolved ? '.resolved' : ''), h('div.cm-head', h('b', cm.author), h('span.muted', ' ' + cm.time)), h('div.muted', '“' + cm.anchor + '”'), h('div', cm.text), h('div.cm-act', h('button.btn', { onclick: () => { const el = doc.body.querySelector('[data-cid="' + cm.id + '"]'); if (el) el.scrollIntoView({ block: 'center' }); } }, 'Đi tới'), h('button.btn', { onclick: () => { cm.resolved = !cm.resolved; draw(); doc.markDirty(); } }, cm.resolved ? 'Mở lại' : 'Giải quyết'), h('button.btn', { onclick: () => { doc.comments = doc.comments.filter(x => x !== cm); const el = doc.body.querySelector('[data-cid="' + cm.id + '"]'); if (el) { const p = el.parentNode; while (el.firstChild) p.insertBefore(el.firstChild, el); p.removeChild(el); } draw(); doc.markDirty(); } }, 'Xóa'))); body.appendChild(card); }); };
  draw(); doc.el.appendChild(pane);
}
export function toggleTrack(doc) { doc.trackChanges = !doc.trackChanges; doc.body.classList.toggle('wd-tracking', doc.trackChanges); toast(doc.trackChanges ? 'Đã bật Theo dõi thay đổi — nội dung mới sẽ được đánh dấu' : 'Đã tắt Theo dõi thay đổi'); doc.updateStatus(); doc.refreshUi(); if (doc.trackChanges) installTrack(doc); }
function installTrack(doc) { if (doc._trackHook) return; doc._trackHook = true; doc.body.addEventListener('beforeinput', e => { if (!doc.trackChanges) return; if (e.inputType === 'insertText' && e.data) { e.preventDefault(); ec('insertHTML', '<ins class="wd-ins" data-author="' + esc(settings().userName) + '">' + esc(e.data) + '</ins>'); } }); }
export function acceptChanges(doc, all) { run(doc, () => { const ins = $$(all ? 'ins.wd-ins' : 'ins.wd-ins', doc.body); ins.forEach(i => { const p = i.parentNode; while (i.firstChild) p.insertBefore(i.firstChild, i); p.removeChild(i); }); $$('del.wd-del', doc.body).forEach(d => d.remove()); }); toast('Đã chấp nhận thay đổi'); }
export function rejectChanges(doc) { run(doc, () => { $$('ins.wd-ins', doc.body).forEach(i => i.remove()); $$('del.wd-del', doc.body).forEach(d => { const p = d.parentNode; while (d.firstChild) p.insertBefore(d.firstChild, d); p.removeChild(d); }); }); toast('Đã từ chối thay đổi'); }
export function toggleMarks(doc) { doc.showMarks = !doc.showMarks; doc.body.classList.toggle('wd-marks', doc.showMarks); }

// ---------------------------------------------------------------- navigation pane
export function navPane(doc) {
  const old = $('.wd-pane.nav', doc.el); if (old) { old.remove(); return; }
  const pane = h('div.wd-pane.nav.left', h('div.pane-head', h('span.t', 'Dẫn hướng'), h('button.pane-x', { onclick: () => pane.remove() }, ic('x', 16))));
  const body = h('div.pane-body'); pane.appendChild(body);
  const draw = () => { body.innerHTML = ''; const heads = $$('h1,h2,h3,h4', doc.body); if (!heads.length) { body.appendChild(h('div.muted', 'Dùng Kiểu đề mục (Heading) để tạo bản đồ tài liệu.')); return; } heads.forEach(hd => { const item = h('div.nav-item.lvl' + hd.tagName[1], hd.textContent || '(trống)'); item.addEventListener('click', () => hd.scrollIntoView({ block: 'start' })); body.appendChild(item); }); };
  draw(); const obs = () => draw(); doc.body.addEventListener('input', obs); doc.el.appendChild(pane);
}

// ---------------------------------------------------------------- context menu
export function contextItems(doc) {
  const cell = currentCell();
  const base = [
    { label: 'Cắt', icon: 'scissors', run: () => ec('cut') },
    { label: 'Sao chép', icon: 'copy', run: () => ec('copy') },
    { label: 'Tùy chọn dán', icon: 'clipboard-paste', sub: [{ label: 'Giữ định dạng', run: () => ec('paste') }, { label: 'Chỉ văn bản', run: () => pasteSpecial(doc) }] },
    { sep: true },
    { label: 'Phông chữ…', icon: 'type', run: () => fontDialog(doc) },
    { label: 'Đoạn văn…', icon: 'pilcrow', run: () => paragraphDialog(doc) },
    { label: 'Kiểu', icon: 'palette', sub: () => STYLES.map(s => ({ label: s[1], run: () => applyStyle(doc, s[0]) })) },
    { sep: true },
    { label: 'Chú thích Mới', icon: 'message-square-plus', run: () => addComment(doc) },
    { label: 'Tra cứu / Hỏi A.I', icon: 'sparkles', run: () => import('../ai/ai.js').then(m => m.openAi(doc.shell, doc, 'Giải thích hoặc cải thiện: ' + (window.getSelection().toString() || '').slice(0, 300))) },
    { label: 'Liên kết…', icon: 'link', run: () => linkDialog(doc) }
  ];
  if (cell) return [
    { label: 'Chèn', icon: 'plus', sub: [{ label: 'Hàng phía trên', run: () => tableOp(doc, 'rowAbove') }, { label: 'Hàng phía dưới', run: () => tableOp(doc, 'rowBelow') }, { label: 'Cột bên trái', run: () => tableOp(doc, 'colLeft') }, { label: 'Cột bên phải', run: () => tableOp(doc, 'colRight') }] },
    { label: 'Xóa', icon: 'trash-2', sub: [{ label: 'Xóa hàng', run: () => tableOp(doc, 'delRow') }, { label: 'Xóa cột', run: () => tableOp(doc, 'delCol') }, { label: 'Xóa bảng', run: () => tableOp(doc, 'delTable') }] },
    { label: 'Hàng đầu làm tiêu đề', icon: 'heading', run: () => tableOp(doc, 'header') },
    { label: 'Căn chỉnh ô', icon: 'align-center', sub: [['alignLeft', 'Trái'], ['alignCenter', 'Giữa'], ['alignRight', 'Phải']].map(x => ({ label: x[1], run: () => tableOp(doc, x[0]) })) },
    { sep: true }
  ].concat(base);
  return base;
}

export function fontDialog(doc) {
  const st = queryState(doc);
  formDlg('Phông chữ', [
    { key: 'font', type: 'select', label: 'Phông chữ', value: st.font || 'Calibri', options: FONTS.map(f => [f, f]) },
    { key: 'size', type: 'number', label: 'Cỡ chữ (pt)', value: st.size || 11 },
    { key: 'b', type: 'checkbox', label: 'Đậm', value: st.bold }, { key: 'i', type: 'checkbox', label: 'Nghiêng', value: st.italic }, { key: 'u', type: 'checkbox', label: 'Gạch chân', value: st.underline }, { key: 's', type: 'checkbox', label: 'Gạch ngang', value: st.strike },
    { key: 'color', type: 'color', label: 'Màu chữ', value: '#000000' }
  ], { width: 420 }).then(v => { if (!v) return; run(doc, () => { ec('fontName', v.font); setSize(doc, v.size); if (v.b !== st.bold) ec('bold'); if (v.i !== st.italic) ec('italic'); if (v.u !== st.underline) ec('underline'); if (v.s !== st.strike) ec('strikeThrough'); if (v.color !== '#000000') ec('foreColor', v.color); }); });
}
export function paragraphDialog(doc) {
  const b = currentBlock(doc); const cs = b ? getComputedStyle(b) : {};
  formDlg('Đoạn văn', [
    { key: 'align', type: 'select', label: 'Căn lề', value: cs.textAlign === 'center' ? 'Center' : cs.textAlign === 'right' ? 'Right' : cs.textAlign === 'justify' ? 'Full' : 'Left', options: [['Left', 'Trái'], ['Center', 'Giữa'], ['Right', 'Phải'], ['Full', 'Đều hai bên']] },
    { key: 'line', type: 'select', label: 'Giãn dòng', value: '1.15', options: [['1', '1,0'], ['1.15', '1,15'], ['1.5', '1,5'], ['2', '2,0']] },
    { key: 'before', type: 'number', label: 'Cách trước (pt)', value: Math.round(parseFloat(cs.marginTop) * 0.75) || 0 },
    { key: 'after', type: 'number', label: 'Cách sau (pt)', value: Math.round(parseFloat(cs.marginBottom) * 0.75) || 8 },
    { key: 'indent', type: 'number', label: 'Thụt lề trái (cm)', value: Math.round(parseFloat(cs.marginLeft) / 37.8 * 10) / 10 || 0 }
  ], { width: 420 }).then(v => { if (!v) return; run(doc, () => { align(doc, v.align); lineSpacing(doc, v.line); selectedBlocks(doc).forEach(bl => { bl.style.marginTop = v.before + 'pt'; bl.style.marginBottom = v.after + 'pt'; bl.style.marginLeft = (v.indent * 37.8) + 'px'; }); }); });
}

// ---------------------------------------------------------------- page setup / print
export function pageSetupDialog(doc) {
  const cm = px => Math.round(px / (96 / 2.54) * 100) / 100; const m = doc.marg || [96, 96, 96, 96];
  formDlg('Thiết lập trang', [
    { key: 'size', type: 'select', label: 'Khổ giấy', value: doc.pageSize, options: ['A4', 'A3', 'A5', 'Letter', 'Legal'].map(x => [x, x]) },
    { key: 'orientation', type: 'radio', label: 'Hướng', value: doc.orientation, options: [['portrait', 'Dọc'], ['landscape', 'Ngang']] },
    { key: 'preset', type: 'select', label: 'Lề có sẵn', value: Array.isArray(doc.margins) ? 'custom' : doc.margins, options: [['custom', 'Tùy chỉnh (nhập bên dưới)'], ['normal', 'Bình thường (2,54 cm)'], ['narrow', 'Hẹp (1,27 cm)'], ['moderate', 'Vừa (2,54 / 1,91 cm)'], ['wide', 'Rộng (2,54 / 5,08 cm)']] },
    { key: 't', type: 'number', label: 'Lề trên (cm)', value: cm(m[0]), step: 0.1 },
    { key: 'b', type: 'number', label: 'Lề dưới (cm)', value: cm(m[2]), step: 0.1 },
    { key: 'l', type: 'number', label: 'Lề trái (cm)', value: cm(m[3]), step: 0.1 },
    { key: 'r', type: 'number', label: 'Lề phải (cm)', value: cm(m[1]), step: 0.1 }
  ], { width: 440 }).then(v => {
    if (!v) return; doc.pageSize = v.size; doc.orientation = v.orientation;
    const px = x => Math.max(0, (parseFloat(String(x).replace(',', '.')) || 0) * 96 / 2.54);
    const custom = [px(v.t), px(v.r), px(v.b), px(v.l)]; const changed = custom.some((x, i) => Math.abs(x - m[i]) > 0.5);
    doc.margins = v.preset === 'custom' || changed ? custom : v.preset;
    doc.applyPageStyle(); doc.updatePages(); doc.updateRuler(); doc.markDirty(); doc.refreshUi();
  });
}
export function setPageQuick(doc, k, val) { doc[k] = val; doc.applyPageStyle(); doc.updatePages(); doc.updateRuler(); doc.markDirty(); doc.refreshUi(); }
export function columns(doc, n) { run(doc, () => { doc.body.style.columnCount = n > 1 ? n : ''; doc.body.style.columnGap = n > 1 ? '36px' : ''; }); }
export function printOptions(doc) {
  const sel = (label, k, opts) => { const i = h('select.inp', opts.map(o => h('option', { value: o[0] }, o[1]))); i.value = doc[k]; i.addEventListener('change', () => { doc[k] = i.value; doc.applyPageStyle(); doc.updatePages(); const pv = $('.print-prev'); if (pv) printPreview(doc, pv); }); return h('div.form-row', h('label.lbl', label), i); };
  return h('div.form', sel('Khổ giấy', 'pageSize', ['A4', 'A3', 'A5', 'Letter', 'Legal'].map(x => [x, x])), sel('Hướng', 'orientation', [['portrait', 'Dọc'], ['landscape', 'Ngang']]), sel('Lề', 'margins', [['normal', 'Bình thường'], ['narrow', 'Hẹp'], ['wide', 'Rộng']]));
}
export function renderPrint(doc, area) {
  clearSelectionArtifacts(doc);
  const [w, hh] = (doc.orientation === 'landscape' ? [doc.pageH, doc.pageW] : [doc.pageW, doc.pageH]);
  const mm = v => (v * 25.4 / 96).toFixed(1) + 'mm';
  const m = doc.marg;
  area.appendChild(h('style', '@page{size:' + mm(w) + ' ' + mm(hh) + ';margin:' + m.map(mm).join(' ') + '} #print-area .wd-body{width:auto;padding:0;min-height:0;box-shadow:none} #print-area .wd-pagebreak{page-break-after:always;height:0} #print-area img{max-width:100%} #print-area table{border-collapse:collapse;width:100%} #print-area td,#print-area th{border:1px solid #000;padding:3px 6px} #print-area .wd-comment{background:none} #print-area .wd-find-hit{background:none}'));
  const clone = doc.body.cloneNode(true);
  clone.querySelectorAll('.wd-frames').forEach(f => f.remove());
  clone.removeAttribute('contenteditable'); clone.removeAttribute('style');
  const wrap = h('div.wd-body'); wrap.innerHTML = clone.innerHTML; area.appendChild(wrap);
}
export function printPreview(doc, box) {
  clearSelectionArtifacts(doc);
  box.innerHTML = '';
  const [w, hh] = (doc.orientation === 'landscape' ? [doc.pageH, doc.pageW] : [doc.pageW, doc.pageH]);
  const sc = Math.min(1, (box.clientWidth - 40) / w);
  const page = h('div.wd-pp', { style: { width: w + 'px', minHeight: hh + 'px', padding: doc.marg.map(x => x + 'px').join(' '), webkitTransform: 'scale(' + sc + ')', transform: 'scale(' + sc + ')', transformOrigin: 'top center', marginBottom: (hh * (sc - 1)) + 'px' } });
  page.innerHTML = doc.body.innerHTML;
  box.appendChild(page);
}

// ---------------------------------------------------------------- export helpers
export function exportHtml(doc) {
  clearSelectionArtifacts(doc);
  return '<!doctype html><html><head><meta charset="utf-8"><title>' + esc(doc.name) + '</title><style>body{font-family:Calibri,Carlito,Arial,sans-serif;max-width:800px;margin:40px auto;line-height:1.5;padding:0 16px} table{border-collapse:collapse} td,th{border:1px solid #ccc;padding:4px 8px} blockquote{border-left:3px solid #ccc;margin-left:0;padding-left:16px;color:#555} .wd-pagebreak{page-break-after:always} img{max-width:100%}</style></head><body>' + doc.body.innerHTML.replace(/ class="wd-[^"]*"/g, '').replace(/ contenteditable="[^"]*"/g, '') + '</body></html>';
}
export function htmlToPlain(el) { return plainWalk(el).replace(/\n{3,}/g, '\n\n').trim(); }
function plainWalk(node) { let out = ''; node.childNodes.forEach(n => { if (n.nodeType === 3) out += n.data; else if (n.nodeType === 1) { const t = n.tagName; if (t === 'BR') out += '\n'; else if (/^(P|DIV|H[1-6]|LI|TR|BLOCKQUOTE|PRE)$/.test(t)) out += plainWalk(n) + '\n'; else if (t === 'TD' || t === 'TH') out += plainWalk(n) + '\t'; else if (n.classList && n.classList.contains('wd-pagebreak')) out += '\n\f\n'; else out += plainWalk(n); } }); return out; }
export { rtfToHtml, htmlToRtf } from '../formats/rtf.js';

export function extraCommands(doc) {
  return [
    { label: 'Tìm (Ctrl+F)', icon: 'search', run: () => findDialog(doc) },
    { label: 'Thay thế (Ctrl+H)', icon: 'replace', run: () => findDialog(doc, true) },
    { label: 'Mục lục', icon: 'list', run: () => insertTOC(doc) },
    { label: 'Dẫn hướng', icon: 'panel-left', run: () => navPane(doc) },
    { label: 'Thống kê số từ', icon: 'calculator', run: () => wordCountDialog(doc) }
  ];
}
import { FONT_LIST as FONTS } from '../core/fonts.js';
export { FONTS };
