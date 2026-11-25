// Yomi-Office Open Source v0.0.6 — OpenDocument Format (ODF 1.2/1.3, ISO/IEC 26300): .odt (Writer), .ods (Calc), .odp (Impress)
// Read + write, fully offline (JSZip + DOMParser). Interoperates with LibreOffice / OpenOffice / OnlyOffice / WPS / Google Docs.
import { JSZipLib, esc, bytesToDataUrl, dataUrlToBytes, uid } from '../core/util.js';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0', text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0', draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink', svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', of: 'urn:oasis:names:tc:opendocument:xmlns:of:1.2', presentation: 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0', number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0', dc: 'http://purl.org/dc/elements/1.1/', manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0', loext: 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0'
};
const ROOT_ATTRS = Object.keys(NS).filter(k => k !== 'manifest').map(k => 'xmlns:' + k + '="' + NS[k] + '"').join(' ');
const MIME = { odt: 'application/vnd.oasis.opendocument.text', ods: 'application/vnd.oasis.opendocument.spreadsheet', odp: 'application/vnd.oasis.opendocument.presentation' };
const x = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
const parse = t => new DOMParser().parseFromString(t, 'application/xml');
const A = (el, ns, n) => (el && el.getAttributeNS ? el.getAttributeNS(NS[ns], n) : null);
const kids = (el, ns, n) => { const out = []; if (!el) return out; for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.namespaceURI === NS[ns] && (!n || c.localName === n)) out.push(c); return out; };
const all = (el, ns, n) => el ? Array.prototype.slice.call(el.getElementsByTagNameNS(NS[ns], n)) : [];
function len2px(v) { if (v == null) return null; const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(String(v).trim()); if (!m) return null; const n = +m[1]; return n * ({ cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 96 / 72, pc: 16, px: 1 }[m[2] || 'cm']); }
const px2cm = p => (Math.round(p / (96 / 2.54) * 1000) / 1000) + 'cm';

async function pack(kind, content, styles, extraFiles) {
  const JSZip = await JSZipLib(); const z = new JSZip();
  z.file('mimetype', MIME[kind], { compression: 'STORE' });   // must be first + stored (ODF 1.2 §3.3)
  z.file('content.xml', '<?xml version="1.0" encoding="UTF-8"?>\n' + content);
  z.file('styles.xml', '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ' + ROOT_ATTRS + ' office:version="1.3">' + (styles || '') + '</office:document-styles>');
  z.file('meta.xml', '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta ' + ROOT_ATTRS + ' office:version="1.3"><office:meta><meta:generator>Yomi-Office Open Source 0.0.6</meta:generator><meta:creation-date>' + new Date().toISOString().replace(/\.\d+Z$/, '') + '</meta:creation-date></office:meta></office:document-meta>');
  const files = ['<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="' + MIME[kind] + '"/>', '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>', '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>', '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>'];
  (extraFiles || []).forEach(f => { z.file(f.path, f.bytes); files.push('<manifest:file-entry manifest:full-path="' + f.path + '" manifest:media-type="' + f.mime + '"/>'); });
  z.file('META-INF/manifest.xml', '<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="' + NS.manifest + '" manifest:version="1.3">' + files.join('') + '</manifest:manifest>');
  return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: MIME[kind] });
}
async function open(bytes) { const JSZip = await JSZipLib(); const z = await JSZip.loadAsync(bytes); const t = async p => { const f = z.file(p); return f ? f.async('string') : null; }; const content = parse(await t('content.xml') || '<x/>'); const styles = parse(await t('styles.xml') || '<x/>'); return { z, content, styles }; }
async function imgUrl(z, href) { const f = z.file(String(href || '').replace(/^\.\//, '')); if (!f) return null; const b = await f.async('uint8array'); const ext = (/\.(\w+)$/.exec(href) || [])[1] || 'png'; return bytesToDataUrl(b, ext === 'svg' ? 'image/svg+xml' : 'image/' + (ext === 'jpg' ? 'jpeg' : ext)); }

// ---------------------------------------------------------------- styles (automatic + common)
function readStyleMap(...docs) {
  const map = {};
  docs.forEach(d => all(d, 'style', 'style').forEach(s => {
    const name = A(s, 'style', 'name'); const tp = kids(s, 'style', 'text-properties')[0]; const pp = kids(s, 'style', 'paragraph-properties')[0]; const cp = kids(s, 'style', 'table-cell-properties')[0]; const colp = kids(s, 'style', 'table-column-properties')[0]; const rowp = kids(s, 'style', 'table-row-properties')[0];
    const o = { parent: A(s, 'style', 'parent-style-name'), family: A(s, 'style', 'family'), dataStyle: A(s, 'style', 'data-style-name') };
    if (tp) { if (/bold|[6-9]00/.test(A(tp, 'fo', 'font-weight') || '')) o.b = true; if (A(tp, 'fo', 'font-style') === 'italic') o.i = true; const u = A(tp, 'style', 'text-underline-style'); if (u && u !== 'none') o.u = true; const lt = A(tp, 'style', 'text-line-through-style'); if (lt && lt !== 'none') o.s = true; const c = A(tp, 'fo', 'color'); if (c) o.color = c; const fs = A(tp, 'fo', 'font-size'); if (fs && /pt$/.test(fs)) o.sz = parseFloat(fs); const fn = A(tp, 'style', 'font-name') || A(tp, 'fo', 'font-family'); if (fn) o.font = fn.replace(/['"]/g, ''); const bg = A(tp, 'fo', 'background-color'); if (bg && bg !== 'transparent') o.hl = bg; const pos = A(tp, 'style', 'text-position'); if (pos) o.pos = /^super|^\d/.test(pos) ? 'sup' : /^sub|^-/.test(pos) ? 'sub' : null; }
    if (pp) { const al = A(pp, 'fo', 'text-align'); if (al) o.align = al === 'end' ? 'right' : al === 'start' ? 'left' : al; const ml = A(pp, 'fo', 'margin-left'); if (ml) o.ml = len2px(ml); const ti = A(pp, 'fo', 'text-indent'); if (ti) o.ti = len2px(ti); }
    if (cp) { const bg = A(cp, 'fo', 'background-color'); if (bg && bg !== 'transparent') o.fill = bg; const va = A(cp, 'style', 'vertical-align'); if (va) o.va = va === 'middle' ? 'middle' : va; const wr = A(cp, 'fo', 'wrap-option'); if (wr === 'wrap') o.wrap = true; const bd = A(cp, 'fo', 'border'); if (bd && bd !== 'none') o.border = true; }
    if (colp) { const w = A(colp, 'style', 'column-width'); if (w) o.colW = len2px(w); }
    if (rowp) { const hgt = A(rowp, 'style', 'row-height'); if (hgt) o.rowH = len2px(hgt); }
    map[name] = o;
  }));
  const resolve = n => { const o = {}; const chain = []; let k = n, guard = 0; while (k && map[k] && guard++ < 20) { chain.unshift(map[k]); k = map[k].parent; } chain.forEach(c => Object.assign(o, c)); return o; };
  return { map, resolve };
}

// ================================================================ ODT (Writer) → HTML
export async function readOdt(bytes) {
  const { z, content, styles } = await open(bytes); const S = readStyleMap(styles, content);
  const body = all(content, 'office', 'text')[0];
  const inline = async (node) => {
    let html = '';
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) { html += esc(n.nodeValue); continue; }
      if (n.nodeType !== 1) continue;
      const ln = n.localName;
      if (n.namespaceURI === NS.text) {
        if (ln === 's') { html += '&nbsp;'.repeat(+(A(n, 'text', 'c') || 1)); continue; }
        if (ln === 'tab') { html += '&emsp;'; continue; }
        if (ln === 'line-break') { html += '<br>'; continue; }
        if (ln === 'span') { const st = S.resolve(A(n, 'text', 'style-name')); let inner = await inline(n); if (st.b) inner = '<b>' + inner + '</b>'; if (st.i) inner = '<i>' + inner + '</i>'; if (st.u) inner = '<u>' + inner + '</u>'; if (st.s) inner = '<s>' + inner + '</s>'; if (st.pos) inner = '<' + st.pos + '>' + inner + '</' + st.pos + '>'; const css = (st.color ? 'color:' + st.color + ';' : '') + (st.sz ? 'font-size:' + st.sz + 'pt;' : '') + (st.font ? 'font-family:' + st.font + ';' : '') + (st.hl ? 'background-color:' + st.hl + ';' : ''); html += css ? '<span style="' + css + '">' + inner + '</span>' : inner; continue; }
        if (ln === 'a') { html += '<a href="' + esc(A(n, 'xlink', 'href') || '#') + '">' + await inline(n) + '</a>'; continue; }
        if (ln === 'note') { const nb = all(n, 'text', 'note-body')[0]; html += '<sup class="wd-fn" title="' + esc(nb ? nb.textContent : '') + '">' + esc((all(n, 'text', 'note-citation')[0] || {}).textContent || '*') + '</sup>'; continue; }
        if (/^(bookmark|bookmark-start|bookmark-end|soft-page-break|sequence-decls)$/.test(ln)) continue;
        html += await inline(n); continue;
      }
      if (n.namespaceURI === NS.draw && ln === 'frame') { const im = kids(n, 'draw', 'image')[0]; if (im) { const url = await imgUrl(z, A(im, 'xlink', 'href')); const w = len2px(A(n, 'svg', 'width')); if (url) html += '<img src="' + url + '"' + (w ? ' style="width:' + Math.round(w) + 'px"' : '') + ' alt="">'; } else { const tb = kids(n, 'draw', 'text-box')[0]; if (tb) html += await blocks(tb); } continue; }
      html += esc(n.textContent);
    }
    return html;
  };
  const pStyle = st => { const css = (st.align && st.align !== 'left' ? 'text-align:' + st.align + ';' : '') + (st.ml ? 'margin-left:' + Math.round(st.ml) + 'px;' : '') + (st.ti ? 'text-indent:' + Math.round(st.ti) + 'px;' : ''); return css ? ' style="' + css + '"' : ''; };
  const blocks = async (node) => {
    let html = '';
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue; const ln = n.localName;
      if (n.namespaceURI === NS.text && ln === 'p') { const st = S.resolve(A(n, 'text', 'style-name')); const nm = (A(n, 'text', 'style-name') || '').toLowerCase(); const inner = await inline(n); const tag = /^title$/.test(nm) ? 'p class="wd-title"' : /^subtitle$/.test(nm) ? 'p class="wd-subtitle"' : 'p'; let c = inner || '<br>'; if (st.b && inner) c = '<b>' + c + '</b>'; html += '<' + tag + pStyle(st) + '>' + c + '</p>'; continue; }
      if (n.namespaceURI === NS.text && ln === 'h') { const lv = Math.min(6, +(A(n, 'text', 'outline-level') || 1)); html += '<h' + lv + pStyle(S.resolve(A(n, 'text', 'style-name'))) + '>' + (await inline(n) || '<br>') + '</h' + lv + '>'; continue; }
      if (n.namespaceURI === NS.text && ln === 'list') { const sn = (A(n, 'text', 'style-name') || ''); const lsEl = all(styles, 'text', 'list-style').concat(all(content, 'text', 'list-style')).find(ls => A(ls, 'style', 'name') === sn); const lvl1 = lsEl ? kids(lsEl, 'text').find(e => A(e, 'text', 'level') === '1') : null; const ordered = !!(lvl1 && lvl1.localName === 'list-level-style-number'); let li = ''; for (const it of kids(n, 'text', 'list-item')) li += '<li>' + (await blocks(it)).replace(/^<p>([\s\S]*)<\/p>$/, '$1') + '</li>'; html += ordered ? '<ol>' + li + '</ol>' : '<ul>' + li + '</ul>'; continue; }
      if (n.namespaceURI === NS.table && ln === 'table') { let t = '<table class="wd-table">'; for (const row of all(n, 'table', 'table-row')) { t += '<tr>'; for (const cell of kids(row, 'table')) { if (cell.localName === 'covered-table-cell') continue; if (cell.localName !== 'table-cell') continue; const cs = A(cell, 'table', 'number-columns-spanned'), rs = A(cell, 'table', 'number-rows-spanned'); const st = S.resolve(A(cell, 'table', 'style-name')); t += '<td' + (cs > 1 ? ' colspan="' + cs + '"' : '') + (rs > 1 ? ' rowspan="' + rs + '"' : '') + (st.fill ? ' style="background:' + st.fill + '"' : '') + '>' + (await blocks(cell)).replace(/^<p>([\s\S]*)<\/p>$/, '$1') + '</td>'; } t += '</tr>'; } html += t + '</table>'; continue; }
      if (n.namespaceURI === NS.text && (ln === 'section' || ln === 'index-body' || ln === 'table-of-content')) { html += await blocks(n); continue; }
      if (n.namespaceURI === NS.draw && ln === 'frame') { html += '<p>' + await inline({ firstChild: n, nextSibling: null }) + '</p>'; continue; }
    }
    return html;
  };
  const html = body ? await blocks(body) : '';
  // page layout (size, margins)
  const pl = all(styles, 'style', 'page-layout-properties')[0]; let sect = null;
  if (pl) { const w = len2px(A(pl, 'fo', 'page-width')), hh = len2px(A(pl, 'fo', 'page-height')); const m = ['margin-top', 'margin-right', 'margin-bottom', 'margin-left'].map(k => len2px(A(pl, 'fo', k)) || 96); sect = { orientation: A(pl, 'style', 'print-orientation') === 'landscape' || (w > hh) ? 'landscape' : 'portrait', size: Math.abs(Math.min(w, hh) - 794) < 20 ? 'A4' : Math.abs(Math.min(w, hh) - 816) < 20 ? 'Letter' : Math.abs(Math.min(w, hh) - 559) < 20 ? 'A5' : 'A4', margins: m }; }
  return { html: html || '<p><br></p>', sect };
}

// HTML (Word body) → ODT
export async function writeOdt(doc) {
  const auto = []; const autoKey = {}; let images = [];
  const tStyle = css => { if (!css) return null; if (autoKey[css]) return autoKey[css]; const n = 'T' + (auto.length + 1); autoKey[css] = n; auto.push('<style:style style:name="' + n + '" style:family="text"><style:text-properties ' + css + '/></style:style>'); return n; };
  const pKey = {}; const pStyleOf = (css, parent) => { const k = parent + '|' + css; if (pKey[k]) return pKey[k]; const n = 'P' + (Object.keys(pKey).length + 1); pKey[k] = n; auto.push('<style:style style:name="' + n + '" style:family="paragraph" style:parent-style-name="' + parent + '"><style:paragraph-properties ' + css + '/></style:style>'); return n; };
  const color = c => { if (!c) return null; const m = /rgb\((\d+),\s*(\d+),\s*(\d+)/.exec(c); if (m) return '#' + [m[1], m[2], m[3]].map(v => (+v < 16 ? '0' : '') + (+v).toString(16)).join(''); return /^#/.test(c) ? c : null; };
  const runCss = (el, inh) => {
    const o = Object.assign({}, inh); const tg = el.tagName;
    if (tg === 'B' || tg === 'STRONG') o.b = 1; if (tg === 'I' || tg === 'EM') o.i = 1; if (tg === 'U') o.u = 1; if (tg === 'S' || tg === 'STRIKE' || tg === 'DEL') o.s = 1; if (tg === 'SUP') o.sup = 1; if (tg === 'SUB') o.sub = 1;
    const st = el.style || {}; if (st.fontWeight === 'bold' || +st.fontWeight >= 600) o.b = 1; if (st.fontStyle === 'italic') o.i = 1; if (/underline/.test(st.textDecoration || st.textDecorationLine || '')) o.u = 1; if (/line-through/.test(st.textDecoration || st.textDecorationLine || '')) o.s = 1;
    if (st.color) o.color = color(st.color); if (st.backgroundColor) o.bg = color(st.backgroundColor); if (st.fontSize && /pt$/.test(st.fontSize)) o.sz = st.fontSize; if (st.fontFamily) o.font = st.fontFamily.split(',')[0].replace(/['"]/g, '');
    if (tg === 'FONT') { if (el.getAttribute('color')) o.color = el.getAttribute('color'); if (el.getAttribute('face')) o.font = el.getAttribute('face'); }
    return o;
  };
  const cssOf = o => [o.b ? 'fo:font-weight="bold"' : '', o.i ? 'fo:font-style="italic"' : '', o.u ? 'style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"' : '', o.s ? 'style:text-line-through-style="solid"' : '', o.color ? 'fo:color="' + o.color + '"' : '', o.bg ? 'fo:background-color="' + o.bg + '"' : '', o.sz ? 'fo:font-size="' + o.sz + '"' : '', o.font ? 'style:font-name="' + x(o.font) + '" fo:font-family="' + x(o.font) + '"' : '', o.sup ? 'style:text-position="super 58%"' : '', o.sub ? 'style:text-position="sub 58%"' : ''].filter(Boolean).join(' ');
  const inline = (node, inh) => {
    let out = '';
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) { const t = n.nodeValue.replace(/ /g, ' '); if (!t) continue; const css = cssOf(inh); const sn = css ? tStyle(css) : null; const body = x(t).replace(/ {2,}/g, m => ' <text:s text:c="' + (m.length - 1) + '"/>').replace(/\t/g, '<text:tab/>'); out += sn ? '<text:span text:style-name="' + sn + '">' + body + '</text:span>' : body; continue; }
      if (n.nodeType !== 1) continue;
      if (n.tagName === 'BR') { out += '<text:line-break/>'; continue; }
      if (n.tagName === 'IMG') { const src = n.getAttribute('src') || ''; if (/^data:/.test(src)) { const m = /^data:image\/(\w+)/.exec(src); const ext = m ? (m[1] === 'jpeg' ? 'jpg' : m[1] === 'svg+xml' ? 'svg' : m[1]) : 'png'; const path = 'Pictures/img' + (images.length + 1) + '.' + ext; images.push({ path, bytes: dataUrlToBytes(src), mime: 'image/' + (ext === 'jpg' ? 'jpeg' : ext === 'svg' ? 'svg+xml' : ext) }); const w = n.width || n.naturalWidth || 300, hh = n.height || n.naturalHeight || 200; out += '<draw:frame draw:name="Image' + images.length + '" text:anchor-type="as-char" svg:width="' + px2cm(w) + '" svg:height="' + px2cm(hh) + '"><draw:image xlink:href="' + path + '" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>'; } continue; }
      if (n.tagName === 'A') { out += '<text:a xlink:type="simple" xlink:href="' + x(n.getAttribute('href') || '') + '">' + inline(n, runCss(n, inh)) + '</text:a>'; continue; }
      out += inline(n, runCss(n, inh));
    }
    return out;
  };
  const paraCss = el => { const st = el.style || {}; const a = []; if (st.textAlign && st.textAlign !== 'left' && st.textAlign !== 'start') a.push('fo:text-align="' + (st.textAlign === 'right' ? 'end' : st.textAlign) + '"'); const ml = parseFloat(st.marginLeft); if (ml) a.push('fo:margin-left="' + px2cm(ml) + '"'); const ti = parseFloat(st.textIndent); if (ti) a.push('fo:text-indent="' + px2cm(ti) + '"'); return a.join(' '); };
  const blocks = (node) => {
    let out = '';
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) { if (n.nodeValue.trim()) out += '<text:p text:style-name="Standard">' + x(n.nodeValue) + '</text:p>'; continue; }
      if (n.nodeType !== 1) continue; const tg = n.tagName;
      if (/^H[1-6]$/.test(tg)) { const lv = +tg[1]; const css = paraCss(n); out += '<text:h text:style-name="' + (css ? pStyleOf(css, 'Heading_20_' + lv) : 'Heading_20_' + lv) + '" text:outline-level="' + lv + '">' + inline(n, {}) + '</text:h>'; continue; }
      if (tg === 'UL' || tg === 'OL') { out += '<text:list text:style-name="' + (tg === 'OL' ? 'LNum' : 'LBul') + '">' + Array.prototype.map.call(n.children, li => '<text:list-item>' + (/^(UL|OL)$/.test((li.lastElementChild || {}).tagName || '') ? '<text:p text:style-name="Standard">' + inline({ firstChild: li.firstChild, nextSibling: null }, {}) + '</text:p>' + blocks({ firstChild: li.lastElementChild }) : '<text:p text:style-name="Standard">' + inline(li, {}) + '</text:p>') + '</text:list-item>').join('') + '</text:list>'; continue; }
      if (tg === 'TABLE') { const rows = Array.prototype.slice.call(n.rows || []); const cols = rows.reduce((m, r) => Math.max(m, Array.prototype.reduce.call(r.cells, (s, c) => s + (c.colSpan || 1), 0)), 0); out += '<table:table table:name="Table' + (out.length % 997) + '"><table:table-column table:number-columns-repeated="' + Math.max(1, cols) + '"/>' + rows.map(r => '<table:table-row>' + Array.prototype.map.call(r.cells, c => '<table:table-cell office:value-type="string"' + (c.colSpan > 1 ? ' table:number-columns-spanned="' + c.colSpan + '"' : '') + (c.rowSpan > 1 ? ' table:number-rows-spanned="' + c.rowSpan + '"' : '') + '><text:p text:style-name="Standard">' + inline(c, c.tagName === 'TH' ? { b: 1 } : {}) + '</text:p></table:table-cell>' + (c.colSpan > 1 ? '<table:covered-table-cell/>'.repeat(c.colSpan - 1) : '')).join('') + '</table:table-row>').join('') + '</table:table>'; continue; }
      if (tg === 'HR' || n.classList.contains('wd-pagebreak')) { out += '<text:p text:style-name="PageBreak"/>'; continue; }
      if (tg === 'DIV' && !/^(P|DIV|H\d|UL|OL|TABLE)$/.test((n.firstElementChild || {}).tagName || '') ) { const css = paraCss(n); out += '<text:p text:style-name="' + (css ? pStyleOf(css, 'Standard') : 'Standard') + '">' + inline(n, {}) + '</text:p>'; continue; }
      if (tg === 'DIV' || tg === 'SECTION' || tg === 'BLOCKQUOTE') { out += blocks(n); continue; }
      const cls = n.classList.contains('wd-title') ? 'Title' : n.classList.contains('wd-subtitle') ? 'Subtitle' : 'Standard'; const css = paraCss(n);
      out += '<text:p text:style-name="' + (css ? pStyleOf(css, cls) : cls) + '">' + inline(n, {}) + '</text:p>';
    }
    return out;
  };
  const body = blocks(doc.body);
  const w = doc.pageW || 794, hh = doc.pageH || 1123, m = doc.marg || [96, 96, 96, 96];
  const content = '<office:document-content ' + ROOT_ATTRS + ' office:version="1.3"><office:font-face-decls/><office:automatic-styles>' + auto.join('') + '<style:style style:name="PageBreak" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:break-before="page"/></style:style></office:automatic-styles><office:body><office:text>' + body + '</office:text></office:body></office:document-content>';
  const font = (doc.docDefaults && doc.docDefaults.font) || 'Calibri';
  const heading = lv => '<style:style style:name="Heading_20_' + lv + '" style:display-name="Heading ' + lv + '" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:default-outline-level="' + lv + '"><style:paragraph-properties fo:margin-top="' + (lv === 1 ? '0.423cm' : '0.212cm') + '" fo:margin-bottom="0.106cm" fo:keep-with-next="always"/><style:text-properties fo:font-size="' + [0, 16, 13, 12, 11, 11, 11][lv] + 'pt" fo:font-weight="' + (lv === 1 ? 'normal' : 'bold') + '" fo:color="#2f5496"/></style:style>';
  const styles = '<office:font-face-decls><style:font-face style:name="' + x(font) + '" svg:font-family="' + x(font) + '"/></office:font-face-decls><office:styles><style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-bottom="0.282cm" fo:line-height="108%"/><style:text-properties style:font-name="' + x(font) + '" fo:font-size="11pt" fo:language="vi" fo:country="VN"/></style:default-style><style:style style:name="Standard" style:family="paragraph" style:class="text"/>' + [1, 2, 3, 4, 5, 6].map(heading).join('') + '<style:style style:name="Title" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-size="28pt" fo:color="#44546a"/></style:style><style:style style:name="Subtitle" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-size="14pt" fo:font-style="italic" fo:color="#595959"/></style:style><text:list-style style:name="LBul"><text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-bullet></text:list-style><text:list-style style:name="LNum"><text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-number></text:list-style></office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="' + px2cm(w) + '" fo:page-height="' + px2cm(hh) + '" style:print-orientation="' + (w > hh ? 'landscape' : 'portrait') + '" fo:margin-top="' + px2cm(m[0]) + '" fo:margin-right="' + px2cm(m[1]) + '" fo:margin-bottom="' + px2cm(m[2]) + '" fo:margin-left="' + px2cm(m[3]) + '"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>';
  return pack('odt', content, styles, images);
}

// ================================================================ ODS (Calc) ⇄ workbook
// OpenFormula ↔ Excel formula
export function ofToExcel(f) {
  let s = String(f || '').replace(/^of:/, '').replace(/^=/, '');
  s = s.replace(/\[\$?'?([^\]']*?)'?\.\$?([A-Z]+)\$?(\d+)(?::\$?'?([^\]'.]*?)'?\.\$?([A-Z]+)\$?(\d+))?\]/g, (m, sh, c1, r1, sh2, c2, r2) => { const pre = sh ? (/[^A-Za-z0-9_]/.test(sh) ? "'" + sh + "'!" : sh + '!') : ''; return pre + c1 + r1 + (c2 ? ':' + c2 + r2 : ''); });
  s = s.replace(/\[\.([A-Z]+)(\d+)(?::\.([A-Z]+)(\d+))?\]/g, (m, c1, r1, c2, r2) => c1 + r1 + (c2 ? ':' + c2 + r2 : ''));
  return s.replace(/;/g, ',');
}
export function excelToOf(f) {
  let out = '', inStr = false; const s = String(f);
  for (let i = 0; i < s.length; i++) { const ch = s[i]; if (ch === '"') inStr = !inStr; out += !inStr && ch === ',' ? ';' : ch; }
  out = out.replace(/(?:('[^']+'|[A-Za-z_][\w.]*)!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?(?![\w(])/g, (m, sh, a, b, off) => { const pre = sh ? '$' + sh.replace(/^'|'$/g, "'") : ''; return '[' + pre + '.' + a + (b ? ':' + pre + '.' + b : '') + ']'; });
  return 'of:=' + out;
}
export async function readOds(wb, bytes, helpers) {
  const { content, styles } = await open(bytes); const S = readStyleMap(styles, content);
  const { internStyle, colIndex } = helpers;
  const numFmtOf = name => { if (!name) return null; const ns = all(content, 'number', 'number-style').concat(all(content, 'number', 'percentage-style'), all(content, 'number', 'currency-style'), all(content, 'number', 'date-style'), all(styles, 'number', 'date-style')).find(e => A(e, 'style', 'name') === name); if (!ns) return null; if (ns.localName === 'percentage-style') return '0%'; if (ns.localName === 'date-style') return 'dd/mm/yyyy'; if (ns.localName === 'currency-style') return '#,##0'; const nn = all(ns, 'number', 'number')[0]; if (!nn) return null; const dp = +(A(nn, 'number', 'decimal-places') || 0); const grp = A(nn, 'number', 'grouping') === 'true'; return (grp ? '#,##0' : '0') + (dp ? '.' + '0'.repeat(dp) : ''); };
  for (const t of all(content, 'table', 'table')) {
    if (t.parentNode && t.parentNode.localName !== 'spreadsheet') continue;
    const s = wb.addSheet(A(t, 'table', 'name') || wb.nextName());
    let c0 = 0; for (const col of all(t, 'table', 'table-column')) { const rep = +(A(col, 'table', 'number-columns-repeated') || 1); const st = S.resolve(A(col, 'table', 'style-name')); if (st.colW && rep < 1000) for (let k = 0; k < rep; k++) s.colW[c0 + k] = Math.round(st.colW); if (A(col, 'table', 'visibility') === 'collapse') for (let k = 0; k < Math.min(rep, 1000); k++) s.hiddenC[c0 + k] = true; c0 += rep; }
    let r = 0;
    for (const row of all(t, 'table', 'table-row')) {
      const rrep = Math.min(+(A(row, 'table', 'number-rows-repeated') || 1), 1048576);
      const rst = S.resolve(A(row, 'table', 'style-name'));
      const cells = kids(row, 'table').filter(e => e.localName === 'table-cell' || e.localName === 'covered-table-cell');
      const hasData = cells.some(e => e.firstChild || A(e, 'office', 'value-type') || A(e, 'table', 'formula'));
      if (!hasData && rrep > 1) { r += rrep; continue; }
      for (let rr = 0; rr < Math.min(rrep, 5000); rr++) {
        if (rst.rowH && Math.abs(rst.rowH - 20) > 2) s.rowH[r] = Math.round(rst.rowH);
        let c = 0;
        for (const cell of cells) {
          const crep = Math.min(+(A(cell, 'table', 'number-columns-repeated') || 1), 16384);
          if (cell.localName === 'covered-table-cell') { c += crep; continue; }
          const vt = A(cell, 'office', 'value-type'); const f = A(cell, 'table', 'formula');
          const text = kids(cell, 'text', 'p').map(p => p.textContent).join('\n');
          const st = S.resolve(A(cell, 'table', 'style-name'));
          const style = {}; if (st.b) style.b = true; if (st.i) style.i = true; if (st.u) style.u = 'single'; if (st.color) style.color = st.color; if (st.fill) style.fill = st.fill; if (st.sz && st.sz !== 10) style.sz = st.sz; if (st.font) style.font = st.font; if (st.align && st.align !== 'start') style.ha = st.align; if (st.wrap) style.wrap = true; if (st.va) style.va = st.va === 'top' ? 'top' : st.va === 'middle' ? 'middle' : null; const nf = numFmtOf(st.dataStyle); if (nf) style.nf = nf;
          if (vt === 'percentage' && !style.nf) style.nf = '0%'; if (vt === 'date' && !style.nf) style.nf = 'dd/mm/yyyy';
          let v = null;
          if (vt === 'float' || vt === 'percentage' || vt === 'currency') v = +A(cell, 'office', 'value');
          else if (vt === 'boolean') v = A(cell, 'office', 'boolean-value') === 'true';
          else if (vt === 'date') { const d = A(cell, 'office', 'date-value'); const m = /^(\d+)-(\d+)-(\d+)(?:T(\d+):(\d+):(\d+))?/.exec(d || ''); if (m) v = helpers.dateToSerial(+m[1], +m[2], +m[3]) + ((+m[4] || 0) * 3600 + (+m[5] || 0) * 60 + (+m[6] || 0)) / 86400; }
          else if (vt === 'time') { const m = /PT(\d+)H(\d+)M([\d.]+)S/.exec(A(cell, 'office', 'time-value') || ''); if (m) { v = (+m[1] * 3600 + +m[2] * 60 + +m[3]) / 86400; style.nf = style.nf || 'hh:mm:ss'; } }
          else if (text) v = text;
          for (let k = 0; k < Math.min(crep, 2000); k++) {
            const patch = {}; if (v != null) patch.v = v; if (f) patch.f = ofToExcel(f); if (Object.keys(style).length) patch.s = internStyle(style);
            if (patch.v != null || patch.f || patch.s) s.setCell(r, c + k, patch);
          }
          const cs = +(A(cell, 'table', 'number-columns-spanned') || 1), rs = +(A(cell, 'table', 'number-rows-spanned') || 1);
          if (cs > 1 || rs > 1) s.merges.push({ r1: r, c1: c, r2: r + rs - 1, c2: c + cs - 1 });
          const ann = kids(cell, 'office', 'annotation')[0]; if (ann) s.setCell(r, c, { note: { text: kids(ann, 'text', 'p').map(p => p.textContent).join('\n'), author: (kids(ann, 'dc', 'creator')[0] || {}).textContent || '' } });
          const link = all(cell, 'text', 'a')[0]; if (link) s.setCell(r, c, { link: { url: A(link, 'xlink', 'href') } });
          c += crep;
        }
        r++;
      }
    }
  }
  return wb;
}
export async function writeOds(wb, helpers) {
  const { colName, isErr, serialToParts } = helpers;
  const auto = []; const sKey = new Map(); let sN = 0; const nfStyles = []; const nfKey = new Map();
  const nfName = nf => { if (!nf || nf === 'General') return null; if (nfKey.has(nf)) return nfKey.get(nf); const n = 'N' + (nfKey.size + 1); nfKey.set(nf, n);
    if (/%/.test(nf)) { const dp = ((/\.(0+)/.exec(nf) || [])[1] || '').length; nfStyles.push('<number:percentage-style style:name="' + n + '"><number:number number:decimal-places="' + dp + '" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>'); }
    else if (/[dmy]/i.test(nf) && !/[#0]/.test(nf)) nfStyles.push('<number:date-style style:name="' + n + '"><number:day number:style="long"/><number:text>/</number:text><number:month number:style="long"/><number:text>/</number:text><number:year number:style="long"/></number:date-style>');
    else { const dp = ((/\.(0+)/.exec(nf) || [])[1] || '').length; nfStyles.push('<number:number-style style:name="' + n + '"><number:number number:decimal-places="' + dp + '" number:min-integer-digits="1"' + (/,/.test(nf) ? ' number:grouping="true"' : '') + '/></number:number-style>'); }
    return n; };
  const cellStyle = st => { if (!st) return null; const k = JSON.stringify(st); if (sKey.has(k)) return sKey.get(k); const n = 'ce' + (++sN); sKey.set(k, n); const nf = nfName(st.nf);
    const tp = [st.b ? 'fo:font-weight="bold"' : '', st.i ? 'fo:font-style="italic"' : '', st.u ? 'style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"' : '', st.s ? 'style:text-line-through-style="solid"' : '', st.color ? 'fo:color="' + st.color + '"' : '', st.sz ? 'fo:font-size="' + st.sz + 'pt"' : '', st.font ? 'style:font-name="' + x(st.font) + '"' : ''].filter(Boolean).join(' ');
    const cp = [st.fill ? 'fo:background-color="' + st.fill + '"' : '', st.wrap ? 'fo:wrap-option="wrap"' : '', st.va ? 'style:vertical-align="' + (st.va === 'middle' ? 'middle' : st.va) + '"' : '', st.bd ? 'fo:border="0.06pt solid #000000"' : ''].filter(Boolean).join(' ');
    auto.push('<style:style style:name="' + n + '" style:family="table-cell" style:parent-style-name="Default"' + (nf ? ' style:data-style-name="' + nf + '"' : '') + '>' + (cp ? '<style:table-cell-properties ' + cp + '/>' : '') + (st.ha ? '<style:paragraph-properties fo:text-align="' + (st.ha === 'right' ? 'end' : st.ha === 'left' ? 'start' : st.ha) + '"/>' : '') + (tp ? '<style:text-properties ' + tp + '/>' : '') + '</style:style>'); return n; };
  const colKey = new Map(); const colStyle = w => { const k = Math.round(w); if (colKey.has(k)) return colKey.get(k); const n = 'co' + (colKey.size + 1); colKey.set(k, n); auto.push('<style:style style:name="' + n + '" style:family="table-column"><style:table-column-properties style:column-width="' + px2cm(k) + '"/></style:style>'); return n; };
  const tables = wb.sheets.map(s => {
    const u = s.used(); const R2 = u.empty ? 0 : u.r2, C2 = u.empty ? 0 : u.c2;
    let cols = ''; for (let c = 0; c <= C2; c++) cols += '<table:table-column table:style-name="' + colStyle(s.colWidth(c) || 64) + '"' + (s.hiddenC[c] ? ' table:visibility="collapse"' : '') + ' table:default-cell-style-name="Default"/>';
    const covered = new Set(); s.merges.forEach(m => { for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(r * 20000 + c); });
    let rows = '';
    for (let r = 0; r <= R2; r++) {
      let row = '';
      for (let c = 0; c <= C2; c++) {
        if (covered.has(r * 20000 + c)) { row += '<table:covered-table-cell/>'; continue; }
        const cell = s.get(r, c); const v = s.value(r, c); const st = cellStyle(s.styleAt(r, c));
        const mg = s.merges.find(m => m.r1 === r && m.c1 === c);
        let a = (st ? ' table:style-name="' + st + '"' : '') + (mg ? ' table:number-columns-spanned="' + (mg.c2 - mg.c1 + 1) + '" table:number-rows-spanned="' + (mg.r2 - mg.r1 + 1) + '"' : '');
        if (cell && cell.f) a += ' table:formula="' + x(excelToOf(cell.f)) + '"';
        let inner = '';
        if (v == null || v === '') {}
        else if (typeof v === 'number') { const nf = (s.styleAt(r, c) || {}).nf || ''; if (/%/.test(nf)) a += ' office:value-type="percentage" office:value="' + v + '"'; else if (/[dmy]/i.test(nf) && !/[#0]/.test(nf)) { const p = serialToParts(v); a += ' office:value-type="date" office:date-value="' + p.y + '-' + (p.m < 10 ? '0' : '') + p.m + '-' + (p.d < 10 ? '0' : '') + p.d + '"'; } else a += ' office:value-type="float" office:value="' + v + '"'; inner = '<text:p>' + x(helpers.display ? helpers.display(s, r, c) : v) + '</text:p>'; }
        else if (typeof v === 'boolean') { a += ' office:value-type="boolean" office:boolean-value="' + v + '"'; inner = '<text:p>' + (v ? 'TRUE' : 'FALSE') + '</text:p>'; }
        else if (isErr(v)) { a += ' office:value-type="string" office:string-value="' + x(v.code) + '"'; inner = '<text:p>' + x(v.code) + '</text:p>'; }
        else { a += ' office:value-type="string"'; inner = String(v).split('\n').map(l => '<text:p>' + (cell && cell.link ? '<text:a xlink:type="simple" xlink:href="' + x(cell.link.url || cell.link) + '">' + x(l) + '</text:a>' : x(l)) + '</text:p>').join(''); }
        if (cell && cell.note) inner = '<office:annotation><dc:creator>' + x(cell.note.author || '') + '</dc:creator><text:p>' + x(cell.note.text || cell.note) + '</text:p></office:annotation>' + inner;
        row += '<table:table-cell' + a + (inner ? '>' + inner + '</table:table-cell>' : '/>');
      }
      rows += '<table:table-row' + (s.hiddenR[r] ? ' table:visibility="collapse"' : '') + '>' + row + '</table:table-row>';
    }
    return '<table:table table:name="' + x(s.name) + '"' + (s.hidden ? ' table:display="false"' : '') + '>' + cols + rows + '</table:table>';
  });
  const content = '<office:document-content ' + ROOT_ATTRS + ' office:version="1.3"><office:automatic-styles>' + nfStyles.join('') + auto.join('') + '</office:automatic-styles><office:body><office:spreadsheet>' + tables.join('') + '</office:spreadsheet></office:body></office:document-content>';
  const styles = '<office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Calibri" fo:font-size="11pt"/></style:default-style><style:style style:name="Default" style:family="table-cell"/></office:styles>';
  return pack('ods', content, styles);
}

// ================================================================ ODP (Impress) ⇄ slides
export async function readOdp(bytes, SW, SH, blankSlide) {
  const { z, content, styles } = await open(bytes); const S = readStyleMap(styles, content);
  const pl = all(styles, 'style', 'page-layout-properties')[0]; const pw = pl ? len2px(A(pl, 'fo', 'page-width')) : 1056, ph = pl ? len2px(A(pl, 'fo', 'page-height')) : 594; const k = SW / (pw || 1056); const ky = SH / (ph || 594);
  const slides = [];
  for (const page of all(content, 'draw', 'page')) {
    const s = blankSlide('blank'); s.objects = [];
    for (const fr of kids(page, 'draw')) {
      const X = len2px(A(fr, 'svg', 'x')) || 0, Y = len2px(A(fr, 'svg', 'y')) || 0, W = len2px(A(fr, 'svg', 'width')) || 200, H = len2px(A(fr, 'svg', 'height')) || 60;
      const geo = { x: X * k, y: Y * ky, w: W * k, h: H * ky };
      const cls = A(fr, 'presentation', 'class');
      const tb = kids(fr, 'draw', 'text-box')[0] || (fr.localName !== 'frame' ? fr : null);
      const img = kids(fr, 'draw', 'image')[0];
      if (img && !tb) { const url = await imgUrl(z, A(img, 'xlink', 'href')); if (url) s.objects.push(Object.assign({ id: uid('o'), type: 'image', src: url, alt: '' }, geo)); continue; }
      if (fr.localName === 'frame' && kids(fr, 'table', 'table').length) { const t = kids(fr, 'table', 'table')[0]; s.objects.push(Object.assign({ id: uid('o'), type: 'table', header: true, rows: all(t, 'table', 'table-row').map(r => kids(r, 'table', 'table-cell').map(c => esc(c.textContent))), fontSize: 18 }, geo)); continue; }
      if (!tb) continue;
      const paras = all(tb, 'text', 'p'); const st0 = paras[0] ? S.resolve(all(tb, 'text', 'span')[0] ? A(all(tb, 'text', 'span')[0], 'text', 'style-name') : null) : {};
      const html = paras.map(p => { const list = p.parentNode && p.parentNode.localName === 'list-item'; return '<div>' + (list ? '• ' : '') + esc(p.textContent) + '</div>'; }).join('');
      const o = Object.assign({ id: uid('o'), type: 'text', html }, geo);
      if (cls === 'title') { o.ph = 'title'; o.fontSize = st0.sz || 40; } else if (cls === 'subtitle') { o.ph = 'subtitle'; o.fontSize = st0.sz || 24; } else { o.fontSize = st0.sz || 18; if (cls === 'outline') o.ph = 'body'; }
      if (st0.color) o.color = st0.color; if (st0.b) o.b = true;
      if (fr.localName === 'custom-shape' || fr.localName === 'rect' || fr.localName === 'ellipse') { const gs = S.resolve(A(fr, 'draw', 'style-name')); s.objects.push(Object.assign({ id: uid('o'), type: 'shape', shape: fr.localName === 'ellipse' ? 'ellipse' : 'rect', fill: gs.fill || '#4472c4', stroke: 'none', text: html }, geo)); continue; }
      s.objects.push(o);
    }
    const notes = kids(page, 'presentation', 'notes')[0]; if (notes) s.notes = esc(all(notes, 'text', 'p').map(p => p.textContent).join('\n'));
    slides.push(s);
  }
  return { slides };
}
export async function writeOdp(doc, SW, SH) {
  const pw = 28 * (96 / 2.54), ph = pw * SH / SW; const k = pw / SW;
  const auto = []; let images = []; const tKey = {};
  const tStyle = (o) => { const css = [o.b || o.ph === 'title' ? 'fo:font-weight="bold"' : '', o.i ? 'fo:font-style="italic"' : '', 'fo:font-size="' + Math.round((o.fontSize || (o.ph === 'title' ? 40 : o.ph === 'subtitle' ? 24 : 18)) * 0.75 * 10) / 10 + 'pt"', 'fo:color="' + (o.color || (o.ph === 'title' ? doc.theme.titleColor : doc.theme.bodyColor) || '#000000') + '"', 'style:font-name="' + x(o.font || doc.theme.font || 'Calibri') + '"'].filter(Boolean).join(' '); if (tKey[css]) return tKey[css]; const n = 'T' + (Object.keys(tKey).length + 1); tKey[css] = n; auto.push('<style:style style:name="' + n + '" style:family="text"><style:text-properties ' + css + '/></style:style>'); return n; };
  const gKey = {}; const gStyle = fill => { const k2 = fill || 'none'; if (gKey[k2]) return gKey[k2]; const n = 'gr' + (Object.keys(gKey).length + 1); gKey[k2] = n; auto.push('<style:style style:name="' + n + '" style:family="graphic"><style:graphic-properties draw:stroke="none"' + (fill && fill !== 'none' ? ' draw:fill="solid" draw:fill-color="' + fill + '"' : ' draw:fill="none"') + '/></style:style>'); return n; };
  const lines = html => { const d = document.createElement('div'); d.innerHTML = html || ''; const t = (d.innerText !== undefined ? d.innerText : d.textContent) || ''; return t.split('\n'); };
  const pos = o => 'svg:x="' + px2cm(o.x * k) + '" svg:y="' + px2cm(o.y * k) + '" svg:width="' + px2cm(o.w * k) + '" svg:height="' + px2cm(o.h * k) + '"';
  const pages = doc.slides.map((s, i) => {
    let body = '';
    s.objects.forEach(o => {
      if (o.type === 'text') { const ts = tStyle(o); const cls = o.ph === 'title' ? ' presentation:class="title"' : o.ph === 'subtitle' ? ' presentation:class="subtitle"' : o.ph ? ' presentation:class="outline"' : ''; body += '<draw:frame draw:style-name="' + gStyle(null) + '"' + cls + ' ' + pos(o) + '><draw:text-box>' + lines(o.html).map(l => '<text:p><text:span text:style-name="' + ts + '">' + x(l) + '</text:span></text:p>').join('') + '</draw:text-box></draw:frame>'; }
      else if (o.type === 'shape') { body += '<draw:custom-shape draw:style-name="' + gStyle(o.fill || doc.theme.colors[0]) + '" ' + pos(o) + '>' + (o.text ? lines(o.text).map(l => '<text:p>' + x(l) + '</text:p>').join('') : '') + '<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="' + (o.shape === 'ellipse' ? 'ellipse' : o.shape === 'triangle' ? 'isosceles-triangle' : o.shape === 'roundRect' ? 'round-rectangle' : 'rectangle') + '"/></draw:custom-shape>'; }
      else if (o.type === 'image' && /^data:/.test(o.src || '')) { const m = /^data:image\/(\w+)/.exec(o.src); const ext = m ? (m[1] === 'jpeg' ? 'jpg' : m[1] === 'svg+xml' ? 'svg' : m[1]) : 'png'; const path = 'Pictures/img' + (images.length + 1) + '.' + ext; images.push({ path, bytes: dataUrlToBytes(o.src), mime: 'image/' + (ext === 'jpg' ? 'jpeg' : ext === 'svg' ? 'svg+xml' : ext) }); body += '<draw:frame ' + pos(o) + '><draw:image xlink:href="' + path + '" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>'; }
      else if (o.type === 'table') { const cols = (o.rows[0] || []).length; body += '<draw:frame ' + pos(o) + '><table:table><table:table-column table:number-columns-repeated="' + cols + '"/>' + o.rows.map(r => '<table:table-row>' + r.map(c => '<table:table-cell><text:p>' + x(lines(c).join(' ')) + '</text:p></table:table-cell>').join('') + '</table:table-row>').join('') + '</table:table></draw:frame>'; }
    });
    const notes = s.notes ? '<presentation:notes><draw:frame presentation:class="notes" svg:x="2cm" svg:y="14cm" svg:width="17cm" svg:height="10cm"><draw:text-box>' + lines(s.notes).map(l => '<text:p>' + x(l) + '</text:p>').join('') + '</draw:text-box></draw:frame></presentation:notes>' : '';
    const bg = s.bg && typeof s.bg === 'string' ? s.bg : (doc.theme.bg && typeof doc.theme.bg === 'string' ? doc.theme.bg : null);
    if (bg) auto.push('<style:style style:name="dp' + (i + 1) + '" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="' + bg + '" presentation:background-visible="true"/></style:style>');
    return '<draw:page draw:name="page' + (i + 1) + '" draw:master-page-name="Default"' + (bg ? ' draw:style-name="dp' + (i + 1) + '"' : '') + '>' + body + notes + '</draw:page>';
  });
  const content = '<office:document-content ' + ROOT_ATTRS + ' office:version="1.3"><office:automatic-styles>' + auto.join('') + '</office:automatic-styles><office:body><office:presentation>' + pages.join('') + '</office:presentation></office:body></office:document-content>';
  const styles = '<office:styles/><office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="' + px2cm(pw) + '" fo:page-height="' + px2cm(ph) + '" style:print-orientation="landscape"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>';
  return pack('odp', content, styles, images);
}
