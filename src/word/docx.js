// DOCX read/write (Office Open XML WordprocessingML) + Markdown/HTML conversion
import { JSZipLib, utf8Decode, utf8Encode, esc, extOf, bytesToDataUrl, dataUrlToBytes, b64FromBytes } from '../core/util.js';

function parseXml(t) { return new DOMParser().parseFromString(t, 'application/xml'); }
function all(el, n) { if (!el) return []; return Array.prototype.slice.call(el.getElementsByTagNameNS('*', n)); }
function kids(el, n) { const o = []; if (!el) return o; for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.localName === n) o.push(c); return o; }
function kid(el, n) { if (!el) return null; for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.localName === n) return c; return null; }
function at(el, n) { if (!el) return null; for (let i = 0; i < el.attributes.length; i++) { const a = el.attributes[i]; if (a.localName === n) return a.value; } return null; }
const EMU = 9525, TW = 1440 / 96; // twips per px

// ================================================================ READ
export async function readDocx(bytes) {
  const JSZip = await JSZipLib(); const zip = await JSZip.loadAsync(bytes);
  const txt = async p => { const f = zip.file(p); return f ? f.async('string') : null; };
  const relsText = await txt('word/_rels/document.xml.rels');
  const rels = {}; if (relsText) all(parseXml(relsText), 'Relationship').forEach(r => { rels[at(r, 'Id')] = { target: at(r, 'Target'), type: at(r, 'Type'), ext: at(r, 'TargetMode') === 'External' }; });
  const stylesText = await txt('word/styles.xml');
  const styleMap = {}; let defaults = { font: 'Calibri', size: 11 };
  if (stylesText) { const sx = parseXml(stylesText); const dr = kid(kid(sx.documentElement, 'docDefaults'), 'rPrDefault'); if (dr) { const rpr = kid(dr, 'rPr'); const rf = kid(rpr, 'rFonts'); if (rf && at(rf, 'ascii')) defaults.font = at(rf, 'ascii'); const sz = kid(rpr, 'sz'); if (sz) defaults.size = +at(sz, 'val') / 2; } all(sx, 'style').forEach(st => { styleMap[at(st, 'styleId')] = { name: (kid(st, 'name') && at(kid(st, 'name'), 'val')) || at(st, 'styleId'), type: at(st, 'type') }; }); }
  const docText = await txt('word/document.xml'); if (!docText) throw new Error('Không tìm thấy nội dung document.xml');
  const doc = parseXml(docText);
  const numbering = await txt('word/numbering.xml');
  const commentsText = await txt('word/comments.xml');
  const comments = [];
  if (commentsText) all(parseXml(commentsText), 'comment').forEach(c => comments.push({ id: at(c, 'id'), author: at(c, 'author') || '', time: (at(c, 'date') || '').slice(0, 16).replace('T', ' '), text: all(c, 't').map(x => x.textContent).join(''), anchor: '' }));
  // images: preload (memory-safe Blob URL to handle massive documents like 277MB docx without base64 string expansion)
  const imgCache = {};
  const canBlob = typeof URL !== 'undefined' && typeof URL.createObjectURL === 'function' && typeof Blob === 'function';
  for (const id in rels) {
    if (/image/.test(rels[id].type) && !rels[id].ext) {
      const p = 'word/' + rels[id].target.replace(/^\.\.\//, '').replace(/^\//, '');
      const f = zip.file(p) || zip.file(rels[id].target.replace(/^\//, ''));
      if (f) {
        const ext = extOf(p);
        const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : ext === 'svg' ? 'image/svg+xml' : ext === 'emf' || ext === 'wmf' ? 'image/png' : 'image/jpeg';
        const b = await f.async('uint8array');
        imgCache[id] = canBlob ? URL.createObjectURL(new Blob([b], { type: mime })) : bytesToDataUrl(b, mime);
      }
    }
  }
  const body = kid(doc.documentElement, 'body');
  let html = ''; let sectPr = null;
  const commentRanges = {};
  for (let node = body.firstChild; node; node = node.nextSibling) {
    if (node.nodeType !== 1) continue;
    if (node.localName === 'p') html += renderP(node, styleMap, rels, imgCache, numbering);
    else if (node.localName === 'tbl') html += renderTbl(node, styleMap, rels, imgCache);
    else if (node.localName === 'sectPr') sectPr = parseSect(node);
  }
  // v0.0.6: Word VBA (.docm/.dotm) — read the module sources of word/vbaProject.bin
  let vba = null;
  const vrel = Object.keys(rels).map(k => rels[k]).find(r => /vbaProject$/.test(r.type));
  const vf = zip.file(vrel ? 'word/' + vrel.target.replace(/^\//, '').replace(/^word\//, '') : 'word/vbaProject.bin') || zip.file('word/vbaProject.bin');
  if (vf) { try { const { readVbaProject } = await import('../excel/vba/ovba.js'); const { vbaToJsa } = await import('../excel/vba/jsa.js'); const pj = readVbaProject(await vf.async('uint8array')); pj.modules.forEach(m => { const js = vbaToJsa(m.code); if (js != null) { m.type = 'jsa'; m.code = js; } }); { const sf = zip.file('oos/vba-source.json'); if (sf) { const { applySidecar } = await import('../excel/vba/ovba.js'); pj.modules = applySidecar(pj.modules, await sf.async('string')); } } pj.modules.forEach(m => { delete m.rawCode; }); vba = { modules: pj.modules, shortcuts: {}, codepage: pj.codepage }; } catch (e) { console.warn('VBA', e); vba = { modules: [], shortcuts: {}, error: String(e.message || e) }; } }
  return { html: html || '<p><br></p>', defaults, sectPr, comments, vba };
}
function parseSect(sp) {
  const sz = kid(sp, 'pgSz'); const mar = kid(sp, 'pgMar'); const out = {};
  if (sz) { const w = +at(sz, 'w'), h = +at(sz, 'h'); out.orientation = at(sz, 'orient') === 'landscape' || w > h ? 'landscape' : 'portrait'; const lw = Math.min(w, h); out.size = Math.abs(lw - 11906) < 200 ? 'A4' : Math.abs(lw - 12240) < 200 ? 'Letter' : Math.abs(lw - 8419) < 200 ? 'A5' : 'A4'; }
  if (mar) { const tw = k => +at(mar, k) || 0; out.margins = [tw('top'), tw('right'), tw('bottom'), tw('left')].map(x => Math.round(x / 15 * 100) / 100); }
  return out;
}
function renderP(p, styleMap, rels, imgCache, numbering) {
  const ppr = kid(p, 'pPr');
  const styleId = ppr && kid(ppr, 'pStyle') && at(kid(ppr, 'pStyle'), 'val');
  const styleName = styleId && styleMap[styleId] ? styleMap[styleId].name : '';
  let tag = 'p', cls = '', dataStyle = '';
  if (/heading\s*1|Heading1/i.test(styleName) || styleId === 'Heading1') { tag = 'h1'; }
  else if (/heading\s*2|Heading2/i.test(styleName) || styleId === 'Heading2') tag = 'h2';
  else if (/heading\s*3|Heading3/i.test(styleName) || styleId === 'Heading3') tag = 'h3';
  else if (/heading\s*4/i.test(styleName)) tag = 'h4';
  else if (/^Title$/i.test(styleName)) { cls = 'wd-title'; dataStyle = 'Title'; }
  else if (/Subtitle/i.test(styleName)) { cls = 'wd-subtitle'; dataStyle = 'Subtitle'; }
  else if (/Quote/i.test(styleName)) tag = 'blockquote';
  const numPr = ppr && kid(ppr, 'numPr');
  let style = '';
  if (ppr) {
    const jc = kid(ppr, 'jc'); if (jc) { const v = at(jc, 'val'); style += 'text-align:' + (v === 'both' ? 'justify' : v) + ';'; }
    const ind = kid(ppr, 'ind'); if (ind) { const l = at(ind, 'left') || at(ind, 'start'); if (l) style += 'margin-left:' + (+l / TW) + 'px;'; const fl = at(ind, 'firstLine'); if (fl) style += 'text-indent:' + (+fl / TW) + 'px;'; }
    const spc = kid(ppr, 'spacing'); if (spc) { const bf = at(spc, 'before'), af = at(spc, 'after'), ln = at(spc, 'line'); if (bf) style += 'margin-top:' + (+bf / 20) + 'pt;'; if (af) style += 'margin-bottom:' + (+af / 20) + 'pt;'; if (ln && at(spc, 'lineRule') !== 'exact') style += 'line-height:' + (+ln / 240) + ';'; }
  }
  let inner = '';
  for (let n = p.firstChild; n; n = n.nextSibling) { if (n.nodeType !== 1) continue; if (n.localName === 'r') inner += renderRun(n, rels, imgCache); else if (n.localName === 'hyperlink') { const rid = at(n, 'id'); const url = rid && rels[rid] ? rels[rid].target : (at(n, 'anchor') ? '#' + at(n, 'anchor') : '#'); inner += '<a href="' + esc(url) + '">' + Array.prototype.map.call(kids(n, 'r'), r => renderRun(r, rels, imgCache)).join('') + '</a>'; } }
  if (!inner) inner = '<br>';
  if (numPr) { const isOl = false; return '<li' + (style ? ' style="' + style + '"' : '') + '>' + inner + '</li>'; }
  return '<' + tag + (cls ? ' class="' + cls + '"' : '') + (dataStyle ? ' data-style="' + dataStyle + '"' : '') + (style ? ' style="' + style + '"' : '') + '>' + inner + '</' + tag + '>';
}
function renderRun(r, rels, imgCache) {
  const rpr = kid(r, 'rPr'); let pre = '', post = '';
  if (rpr) {
    if (kid(rpr, 'b') && at(kid(rpr, 'b'), 'val') !== '0' && at(kid(rpr, 'b'), 'val') !== 'false') { pre = '<b>' + pre; post += '</b>'; }
    if (kid(rpr, 'i') && at(kid(rpr, 'i'), 'val') !== '0') { pre = '<i>' + pre; post += '</i>'; }
    if (kid(rpr, 'u') && at(kid(rpr, 'u'), 'val') !== 'none') { pre = '<u>' + pre; post += '</u>'; }
    if (kid(rpr, 'strike')) { pre = '<s>' + pre; post += '</s>'; }
    if (kid(rpr, 'vertAlign')) { const v = at(kid(rpr, 'vertAlign'), 'val'); if (v === 'superscript') { pre = '<sup>' + pre; post += '</sup>'; } else if (v === 'subscript') { pre = '<sub>' + pre; post += '</sub>'; } }
    const inl = [];
    const rf = kid(rpr, 'rFonts'); if (rf && at(rf, 'ascii')) inl.push('font-family:\'' + at(rf, 'ascii') + '\'');
    const sz = kid(rpr, 'sz'); if (sz) inl.push('font-size:' + (+at(sz, 'val') / 2) + 'pt');
    const col = kid(rpr, 'color'); if (col && at(col, 'val') && at(col, 'val') !== 'auto') inl.push('color:#' + at(col, 'val'));
    const hl = kid(rpr, 'highlight'); if (hl) inl.push('background:' + hlColor(at(hl, 'val')));
    const sh = kid(rpr, 'shd'); if (sh && at(sh, 'fill') && at(sh, 'fill') !== 'auto') inl.push('background:#' + at(sh, 'fill'));
    if (inl.length) { pre = '<span style="' + inl.join(';') + '">' + pre; post += '</span>'; }
  }
  let content = '';
  for (let n = r.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    if (n.localName === 't') content += esc(n.textContent);
    else if (n.localName === 'br' || n.localName === 'cr') content += at(n, 'type') === 'page' ? '</p><div class="wd-pagebreak"></div><p>' : '<br>';
    else if (n.localName === 'tab') content += '&#9;';
    else if (n.localName === 'drawing' || n.localName === 'pict') { const blip = all(n, 'blip')[0] || all(n, 'imagedata')[0]; const rid = blip && (at(blip, 'embed') || at(blip, 'id')); if (rid && imgCache[rid]) { const ext = all(n, 'extent')[0]; let w = ext ? +at(ext, 'cx') / EMU : 200; content += '<img src="' + imgCache[rid] + '" style="width:' + Math.round(w) + 'px;max-width:100%" alt="">'; } }
  }
  return content ? pre + content + post : '';
}
function hlColor(v) { return { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', red: '#ff0000', blue: '#0000ff', darkYellow: '#808000', darkGreen: '#008000', darkCyan: '#008080', lightGray: '#d3d3d3' }[v] || v; }
function renderTbl(tbl, styleMap, rels, imgCache) {
  let html = '<table class="wd-table"><tbody>';
  kids(tbl, 'tr').forEach((tr, ri) => { html += '<tr>'; kids(tr, 'tc').forEach(tc => { const tcpr = kid(tc, 'tcPr'); const gs = tcpr && kid(tcpr, 'gridSpan'); const span = gs ? ' colspan="' + at(gs, 'val') + '"' : ''; const shd = tcpr && kid(tcpr, 'shd'); const bg = shd && at(shd, 'fill') && at(shd, 'fill') !== 'auto' ? ' style="background:#' + at(shd, 'fill') + '"' : ''; let inner = kids(tc, 'p').map(p => { const r = renderP(p, styleMap, rels, imgCache); return r.replace(/^<p[^>]*>|<\/p>$/g, ''); }).join('<br>') || '<br>'; const tag = ri === 0 ? 'th' : 'td'; html += '<' + tag + span + bg + '>' + inner + '</' + tag + '>'; }); html += '</tr>'; });
  return html + '</tbody></table>';
}

// ================================================================ WRITE
function xmlEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ''); }
const PAGE_TW = { A4: [11906, 16838], A3: [16838, 23811], A5: [8419, 11906], Letter: [12240, 15840], Legal: [12240, 20160] };
const MARG_TW = { normal: 1440, narrow: 720, moderate: 1440, wide: 1440 };
export async function writeDocx(doc, opts) {
  const macro = !!(opts && opts.macro);
  const JSZip = await JSZipLib(); const zip = new JSZip();
  const rels = []; const images = []; let relN = 0; const comments = []; const commentsRefs = [];
  const bodyXml = convertHtml(doc.body, { rels, images, nextRel: () => 'rId' + (++relN + 10), doc, comments });
  let [w, hh] = PAGE_TW[doc.pageSize] || PAGE_TW.A4; if (doc.orientation === 'landscape') { const t = w; w = hh; hh = t; }
  const mp = doc.marg || [96, 96, 96, 96]; const tw = x => Math.round(x * 15);
  const sect = '<w:sectPr><w:pgSz w:w="' + w + '" w:h="' + hh + '"' + (doc.orientation === 'landscape' ? ' w:orient="landscape"' : '') + '/><w:pgMar w:top="' + tw(mp[0]) + '" w:right="' + tw(mp[1]) + '" w:bottom="' + tw(mp[2]) + '" w:left="' + tw(mp[3]) + '" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
  const docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>' + bodyXml + sect + '</w:body></w:document>';
  zip.file('word/document.xml', docXml);
  images.forEach(im => zip.file('word/media/' + im.name, im.bytes));
  const relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' + rels.join('') + '</Relationships>';
  if (macro) {
    const { writeVbaProject } = await import('../excel/vba/ovba.js'); const { jsaToVba } = await import('../excel/vba/jsa.js');
    const user = ((doc.vba && doc.vba.modules) || []).filter(m => m.name !== '__Immediate');
    const mods = [{ name: 'ThisDocument', type: 'document', code: (user.find(m => m.type === 'document') || {}).code || '' }]
      .concat(user.filter(m => m.type !== 'document').map(m => ({ name: m.name, type: m.type === 'class' ? 'class' : 'module', code: m.type === 'jsa' ? jsaToVba(m.code) : m.code })));
    zip.file('word/vbaProject.bin', writeVbaProject(mods, { name: 'Project' }));
    { const { vbaSidecar } = await import('../excel/vba/ovba.js'); zip.file('oos/vba-source.json', vbaSidecar([{ name: 'ThisDocument', type: 'document', code: mods[0].code }].concat(user.filter(m => m.type !== 'document').map(m => ({ name: m.name, type: m.type || 'module', code: m.code, wcode: m.type === 'jsa' ? jsaToVba(m.code) : m.code }))))); }
  }
  zip.file('word/_rels/document.xml.rels', macro ? relsXml.replace('</Relationships>', '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>') : relsXml);
  zip.file('word/styles.xml', stylesXml(doc));
  zip.file('word/numbering.xml', numberingXml());
  zip.file('word/comments.xml', commentsXml(doc));
  zip.file('word/settings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/></w:settings>');
  const types = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>', '<Default Extension="png" ContentType="image/png"/>', '<Default Extension="jpeg" ContentType="image/jpeg"/>', '<Default Extension="jpg" ContentType="image/jpeg"/>', '<Default Extension="gif" ContentType="image/gif"/>', '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>', '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'];
  if (macro) { types.push('<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'); types.push('<Default Extension="json" ContentType="application/json"/>'); for (let i = 0; i < types.length; i++) types[i] = types[i].replace('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', 'application/vnd.ms-word.document.macroEnabled.main+xml'); }
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + types.join('') + '</Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + xmlEsc(doc.name) + '</dc:title><dc:creator>Yomi-Office Open Source</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
function cssColor(c) { if (!c) return null; c = c.trim(); if (/^#/.test(c)) return c.slice(1).replace(/^(.)(.)(.)$/, '$1$1$2$2$3$3').toUpperCase(); const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c); if (m) return [m[1], m[2], m[3]].map(x => ('0' + (+x).toString(16)).slice(-2)).join('').toUpperCase(); const named = { black: '000000', red: 'FF0000', blue: '0000FF', green: '008000', white: 'FFFFFF', yellow: 'FFFF00' }; return named[c.toLowerCase()] || null; }
function runProps(st) {
  let p = '';
  if (st.bold) p += '<w:b/>'; if (st.italic) p += '<w:i/>'; if (st.underline) p += '<w:u w:val="single"/>'; if (st.strike) p += '<w:strike/>';
  if (st.sup) p += '<w:vertAlign w:val="superscript"/>'; if (st.sub) p += '<w:vertAlign w:val="subscript"/>';
  if (st.font) p += '<w:rFonts w:ascii="' + xmlEsc(st.font) + '" w:hAnsi="' + xmlEsc(st.font) + '" w:cs="' + xmlEsc(st.font) + '"/>';
  if (st.size) p += '<w:sz w:val="' + Math.round(st.size * 2) + '"/><w:szCs w:val="' + Math.round(st.size * 2) + '"/>';
  if (st.color) { const c = cssColor(st.color); if (c) p += '<w:color w:val="' + c + '"/>'; }
  if (st.bg) { const c = cssColor(st.bg); if (c) p += '<w:highlight w:val="' + nearestHighlight(c) + '"/><w:shd w:val="clear" w:fill="' + c + '"/>'; }
  return p ? '<w:rPr>' + p + '</w:rPr>' : '';
}
function nearestHighlight(hex) { const hl = { yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', red: 'FF0000', blue: '0000FF', darkYellow: '808000', lightGray: 'D3D3D3' }; let best = 'yellow', bd = 1e9; for (const k in hl) { const a = hl[k]; const d = Math.abs(parseInt(a.slice(0, 2), 16) - parseInt(hex.slice(0, 2), 16)) + Math.abs(parseInt(a.slice(2, 4), 16) - parseInt(hex.slice(2, 4), 16)) + Math.abs(parseInt(a.slice(4, 6), 16) - parseInt(hex.slice(4, 6), 16)); if (d < bd) { bd = d; best = k; } } return best; }
// convert contenteditable HTML -> wml
function convertHtml(root, ctx) {
  let out = '';
  for (let node = root.firstChild; node; node = node.nextSibling) out += convertBlock(node, ctx);
  return out || '<w:p/>';
}
function convertBlock(node, ctx) {
  if (node.nodeType === 3) { const t = node.data.trim(); return t ? '<w:p>' + runXml(node.data, {}, ctx) + '</w:p>' : ''; }
  if (node.nodeType !== 1) return '';
  const tag = node.tagName;
  if (node.classList && node.classList.contains('wd-pagebreak')) return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  if (node.classList && (node.classList.contains('wd-frames') || node.classList.contains('wd-find-hit') && false)) return '';
  if (tag === 'TABLE') return convertTable(node, ctx);
  if (tag === 'UL' || tag === 'OL') { let o = ''; kids(node, 'li').forEach(li => { o += paraXml(li, ctx, { numId: tag === 'OL' ? 2 : 1 }); }); return o; }
  if (tag === 'DIV' && (node.classList.contains('wd-toc') || node.classList.contains('wd-cover'))) { let o = ''; node.childNodes.forEach(c => { o += convertBlock(c, ctx); }); return o; }
  if (/^(P|H[1-6]|LI|BLOCKQUOTE|PRE|DIV)$/.test(tag)) return paraXml(node, ctx);
  // inline-only wrapper -> paragraph
  return '<w:p>' + runsFrom(node, {}, ctx) + '</w:p>';
}
function paraXml(el, ctx, opts) {
  opts = opts || {};
  const tag = el.tagName; const cs = el.style || {};
  let ppr = '';
  const styleName = el.getAttribute && el.getAttribute('data-style');
  if (/^H([1-6])$/.test(tag)) ppr += '<w:pStyle w:val="Heading' + tag[1] + '"/>';
  else if (styleName === 'Title') ppr += '<w:pStyle w:val="Title"/>';
  else if (styleName === 'Subtitle') ppr += '<w:pStyle w:val="Subtitle"/>';
  else if (tag === 'BLOCKQUOTE') ppr += '<w:pStyle w:val="Quote"/>';
  if (opts.numId) ppr += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + opts.numId + '"/></w:numPr>';
  const ta = cs.textAlign; if (ta) ppr += '<w:jc w:val="' + (ta === 'justify' ? 'both' : ta) + '"/>';
  let spc = '';
  if (cs.marginTop) spc += ' w:before="' + Math.round(parseFloat(cs.marginTop) * 20) + '"';
  if (cs.marginBottom) spc += ' w:after="' + Math.round(parseFloat(cs.marginBottom) * 20) + '"';
  if (cs.lineHeight && !/px/.test(cs.lineHeight)) spc += ' w:line="' + Math.round(parseFloat(cs.lineHeight) * 240) + '" w:lineRule="auto"';
  if (spc) ppr += '<w:spacing' + spc + '/>';
  if (cs.marginLeft || cs.textIndent) ppr += '<w:ind' + (cs.marginLeft ? ' w:left="' + Math.round(parseFloat(cs.marginLeft) * TW) + '"' : '') + (cs.textIndent ? ' w:firstLine="' + Math.round(parseFloat(cs.textIndent) * TW) + '"' : '') + '/>';
  const runs = runsFrom(el, {}, ctx);
  return '<w:p>' + (ppr ? '<w:pPr>' + ppr + '</w:pPr>' : '') + runs + '</w:p>';
}
function runsFrom(el, inherit, ctx) {
  let out = '';
  el.childNodes.forEach(n => {
    if (n.nodeType === 3) { out += runXml(n.data, inherit, ctx); return; }
    if (n.nodeType !== 1) return;
    const t = n.tagName;
    if (t === 'BR') { out += '<w:r><w:br/></w:r>'; return; }
    if (n.classList && n.classList.contains('wd-pagebreak')) { out += '<w:r><w:br w:type="page"/></w:r>'; return; }
    const st = Object.assign({}, inherit);
    if (t === 'B' || t === 'STRONG') st.bold = true; if (t === 'I' || t === 'EM') st.italic = true; if (t === 'U') st.underline = true; if (t === 'S' || t === 'STRIKE' || t === 'DEL') st.strike = true;
    if (t === 'SUP') st.sup = true; if (t === 'SUB') st.sub = true;
    const sv = n.style || {}; if (sv.fontWeight === 'bold' || +sv.fontWeight >= 600) st.bold = true; if (sv.fontStyle === 'italic') st.italic = true; if (sv.fontFamily) st.font = sv.fontFamily.replace(/['"]/g, '').split(',')[0]; if (sv.fontSize) st.size = parseFloat(sv.fontSize) * (/pt/.test(sv.fontSize) ? 1 : 0.75); if (sv.color) st.color = sv.color; if (sv.backgroundColor || sv.background) st.bg = sv.backgroundColor || sv.background;
    if (t === 'A') { const rid = ctx.nextRel(); ctx.rels.push('<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + xmlEsc(n.href) + '" TargetMode="External"/>'); out += '<w:hyperlink r:id="' + rid + '">' + runsFrom(n, Object.assign({ color: '0563C1', underline: true }, st), ctx) + '</w:hyperlink>'; return; }
    if (t === 'IMG') { out += imgXml(n, ctx); return; }
    if (t === 'SPAN' && n.classList.contains('wd-eq')) { out += '<w:r>' + runProps(st) + '<w:t xml:space="preserve">' + xmlEsc(n.getAttribute('data-latex') || '') + '</w:t></w:r>'; return; }
    out += runsFrom(n, st, ctx);
  });
  return out;
}
function runXml(text, st, ctx) {
  if (text === '') return '';
  // split tabs
  const parts = text.split('\t');
  let o = '';
  parts.forEach((p, i) => { if (i) o += '<w:tab/>'; if (p) o += '<w:t xml:space="preserve">' + xmlEsc(p) + '</w:t>'; });
  return o ? '<w:r>' + runProps(st) + o + '</w:r>' : '';
}
function imgXml(img, ctx) {
  const src = img.src; const m = /^data:image\/(png|jpe?g|gif)/.exec(src); if (!m) return '';
  const ext = m[1] === 'jpg' ? 'jpeg' : m[1]; const name = 'image' + (ctx.images.length + 1) + '.' + ext;
  ctx.images.push({ name, bytes: dataUrlToBytes(src) });
  const rid = ctx.nextRel(); ctx.rels.push('<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + name + '"/>');
  const w = (parseInt(img.style.width) || img.width || 200), hh = Math.round(w * (img.naturalHeight && img.naturalWidth ? img.naturalHeight / img.naturalWidth : 0.6));
  const cx = Math.round(w * EMU), cy = Math.round(hh * EMU); const id = ctx.images.length;
  return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:docPr id="' + id + '" name="Picture ' + id + '" descr="' + xmlEsc(img.alt || '') + '"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="' + id + '" name="' + name + '"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}
function convertTable(tbl, ctx) {
  const rows = kids(tbl.tBodies[0] || tbl, 'tr'); const ncol = rows[0] ? rows[0].children.length : 1;
  let x = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="auto"/>').join('') + '</w:tblBorders></w:tblPr><w:tblGrid>' + Array(ncol).fill('<w:gridCol w:w="' + Math.round(9000 / ncol) + '"/>').join('') + '</w:tblGrid>';
  rows.forEach(tr => { x += '<w:tr>'; Array.prototype.forEach.call(tr.children, td => { const span = td.colSpan > 1 ? '<w:gridSpan w:val="' + td.colSpan + '"/>' : ''; const bg = td.style.background || td.style.backgroundColor; const shd = bg ? '<w:shd w:val="clear" w:fill="' + (cssColor(bg) || 'auto') + '"/>' : ''; const bold = td.tagName === 'TH'; x += '<w:tc><w:tcPr>' + span + shd + '</w:tcPr>' + (td.innerHTML.trim() ? paraXml(wrapInline(td, bold), ctx) : '<w:p/>') + '</w:tc>'; }); x += '</w:tr>'; });
  return x + '</w:tbl><w:p/>';
}
function wrapInline(td, bold) { const p = document.createElement('p'); p.innerHTML = td.innerHTML; if (bold) p.style.fontWeight = 'bold'; if (td.style.textAlign) p.style.textAlign = td.style.textAlign; return p; }
function stylesXml(doc) {
  const head = (n, sz, color) => '<w:style w:type="paragraph" w:styleId="Heading' + n + '"><w:name w:val="heading ' + n + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="' + (n == 1 ? 240 : 160) + '" w:after="80"/><w:outlineLvl w:val="' + (n - 1) + '"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/><w:b/><w:color w:val="' + color + '"/><w:sz w:val="' + sz + '"/></w:rPr></w:style>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="' + (doc.docDefaults && doc.docDefaults.font || 'Calibri') + '" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' + head(1, 32, '2F5496') + head(2, 26, '2F5496') + head(3, 24, '1F3864') + head(4, 22, '2F5496') + '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="80"/></w:pPr><w:rPr><w:sz w:val="56"/><w:color w:val="44546A"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="595959"/><w:sz w:val="28"/><w:i/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>';
}
function numberingXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>';
}
function commentsXml(doc) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + (doc.comments || []).map((c, i) => '<w:comment w:id="' + i + '" w:author="' + xmlEsc(c.author) + '" w:date="' + new Date().toISOString() + '"><w:p><w:r><w:t xml:space="preserve">' + xmlEsc(c.text) + '</w:t></w:r></w:p></w:comment>').join('') + '</w:comments>';
}

// ================================================================ Markdown
export function mdToHtml(md) {
  const lines = md.replace(/\r\n/g, '\n').split('\n'); let html = '', inList = null, inCode = false;
  const inline = s => esc(s).replace(/`([^`]+)`/g, '<code>$1</code>').replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>').replace(/\*([^*]+)\*/g, '<i>$1</i>').replace(/__([^_]+)__/g, '<b>$1</b>').replace(/~~([^~]+)~~/g, '<s>$1</s>').replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
  for (let i = 0; i < lines.length; i++) {
    let l = lines[i];
    if (/^```/.test(l)) { if (inCode) { html += '</pre>'; inCode = false; } else { html += '<pre>'; inCode = true; } continue; }
    if (inCode) { html += esc(l) + '\n'; continue; }
    const h = /^(#{1,6})\s+(.*)/.exec(l);
    if (h) { if (inList) { html += '</' + inList + '>'; inList = null; } html += '<h' + h[1].length + '>' + inline(h[2]) + '</h' + h[1].length + '>'; continue; }
    if (/^\s*[-*+]\s+/.test(l)) { if (inList !== 'ul') { if (inList) html += '</' + inList + '>'; html += '<ul>'; inList = 'ul'; } html += '<li>' + inline(l.replace(/^\s*[-*+]\s+/, '')) + '</li>'; continue; }
    if (/^\s*\d+\.\s+/.test(l)) { if (inList !== 'ol') { if (inList) html += '</' + inList + '>'; html += '<ol>'; inList = 'ol'; } html += '<li>' + inline(l.replace(/^\s*\d+\.\s+/, '')) + '</li>'; continue; }
    if (/^>\s?/.test(l)) { html += '<blockquote>' + inline(l.replace(/^>\s?/, '')) + '</blockquote>'; continue; }
    if (/^(\s*[-*_]){3,}\s*$/.test(l)) { html += '<hr>'; continue; }
    if (inList) { html += '</' + inList + '>'; inList = null; }
    if (l.trim() === '') { html += ''; continue; }
    html += '<p>' + inline(l) + '</p>';
  }
  if (inList) html += '</' + inList + '>'; if (inCode) html += '</pre>';
  return html || '<p><br></p>';
}
export function htmlToMd(root) {
  let out = '';
  const inline = node => { let s = ''; node.childNodes.forEach(n => { if (n.nodeType === 3) s += n.data; else if (n.nodeType === 1) { const t = n.tagName; const inner = inline(n); if (t === 'B' || t === 'STRONG') s += '**' + inner + '**'; else if (t === 'I' || t === 'EM') s += '*' + inner + '*'; else if (t === 'S' || t === 'DEL') s += '~~' + inner + '~~'; else if (t === 'CODE') s += '`' + inner + '`'; else if (t === 'A') s += '[' + inner + '](' + n.href + ')'; else if (t === 'BR') s += '  \n'; else s += inner; } }); return s; };
  root.childNodes.forEach(n => {
    if (n.nodeType !== 1) { if (n.nodeType === 3 && n.data.trim()) out += n.data.trim() + '\n\n'; return; }
    const t = n.tagName;
    if (/^H([1-6])$/.test(t)) out += '#'.repeat(+t[1]) + ' ' + inline(n) + '\n\n';
    else if (t === 'UL') { kids(n, 'li').forEach(li => out += '- ' + inline(li) + '\n'); out += '\n'; }
    else if (t === 'OL') { kids(n, 'li').forEach((li, i) => out += (i + 1) + '. ' + inline(li) + '\n'); out += '\n'; }
    else if (t === 'BLOCKQUOTE') out += '> ' + inline(n) + '\n\n';
    else if (t === 'PRE') out += '```\n' + n.textContent + '\n```\n\n';
    else if (t === 'HR' || (n.classList && n.classList.contains('wd-pagebreak'))) out += '\n---\n\n';
    else if (t === 'TABLE') { const rows = n.querySelectorAll('tr'); rows.forEach((tr, i) => { out += '| ' + Array.prototype.map.call(tr.children, td => inline(td).replace(/\|/g, '\\|')).join(' | ') + ' |\n'; if (i === 0) out += '| ' + Array.prototype.map.call(tr.children, () => '---').join(' | ') + ' |\n'; }); out += '\n'; }
    else { const s = inline(n).trim(); if (s) out += s + '\n\n'; }
  });
  return out.trim() + '\n';
}
