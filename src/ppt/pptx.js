// PPTX read/write (Office Open XML PresentationML)
import { JSZipLib, esc, extOf, bytesToDataUrl, dataUrlToBytes, utf8Decode } from '../core/util.js';
import { THEMES, blankSlide, layoutObjects } from './templates.js';
import { SW, SH } from './ppt.js';

const EMU = 914400 / 96; // emu per px for our 1280x720 logical canvas scaled; we use EMU per px with 96dpi then map
// PPTX default slide 13.333in x 7.5in = 12192000 x 6858000 EMU for 16:9; map to SW x SH px
// v0.0.7 fix: computed from literals — ppt.js ↔ pptx.js import each other, so SW/SH are still undefined while this module
// initialises (EMUX was NaN → every exported .pptx had x/y/cx/cy="NaN" and Microsoft PowerPoint refused to open it).
const EMUX = 12192000 / 1280, EMUY = 6858000 / 720;
function parseXml(t) { return new DOMParser().parseFromString(t, 'application/xml'); }
function all(el, n) { if (!el) return []; return Array.prototype.slice.call(el.getElementsByTagNameNS('*', n)); }
function kid(el, n) { if (!el) return null; for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.localName === n) return c; return null; }
function kids(el, n) { const o = []; if (!el) return o; for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.localName === n) o.push(c); return o; }
function at(el, n) { if (!el) return null; for (let i = 0; i < el.attributes.length; i++) if (el.attributes[i].localName === n) return el.attributes[i].value; return null; }

// ================================================================ READ
export async function readPptx(bytes) {
  const JSZip = await JSZipLib(); const zip = await JSZip.loadAsync(bytes);
  const txt = async p => { const f = zip.file(p); return f ? f.async('string') : null; };
  const presXml = await txt('ppt/presentation.xml');
  let ratio = '16:9';
  if (presXml) { const sz = all(parseXml(presXml), 'sldSz')[0]; if (sz) { const cx = +at(sz, 'cx'), cy = +at(sz, 'cy'); ratio = Math.abs(cx / cy - 4 / 3) < 0.05 ? '4:3' : '16:9'; } }
  const presRels = await relsOf(zip, 'ppt/_rels/presentation.xml.rels');
  // v0.0.7: PowerPoint VBA (.pptm/.ppsm/.potm)
  let vba = null; const vf = zip.file('ppt/vbaProject.bin');
  if (vf) { try { const O = await import('../excel/vba/ovba.js'); const pj = O.readVbaProject(await vf.async('uint8array')); const sf = zip.file('oos/vba-source.json'); if (sf) pj.modules = O.applySidecar(pj.modules, await sf.async('string')); pj.modules.forEach(m => { delete m.rawCode; }); vba = { modules: pj.modules, shortcuts: {}, codepage: pj.codepage }; } catch (e) { console.warn('VBA', e); vba = { modules: [], shortcuts: {}, error: String(e.message || e) }; } }
  // theme (first)
  let theme = Object.assign({}, THEMES[0]);
  const themeXml = await txt('ppt/theme/theme1.xml');
  if (themeXml) { const tx = parseXml(themeXml); const cs = all(tx, 'clrScheme')[0]; if (cs) { const grab = n => { const e = kid(cs, n); if (!e) return null; const c = kid(e, 'srgbClr'); return c ? '#' + at(c, 'val') : null; }; const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map(grab).filter(Boolean); if (accents.length) theme.colors = accents; const fs = all(tx, 'latin')[0]; if (fs && at(fs, 'typeface')) theme.font = at(fs, 'typeface'); theme.name = at(all(tx, 'theme')[0], 'name') || 'Chủ đề'; } }
  // slide order
  const sldIds = all(parseXml(presXml), 'sldId');
  const slides = [];
  for (const sid of sldIds) {
    const relId = relIdOf(sid); const rel = presRels[relId];
    if (!rel) continue;
    let path = rel.target.replace(/^\.\.\//, '').replace(/^\//, '');
    if (!path.startsWith('ppt/')) path = 'ppt/' + path;
    const sx = await txt(path); if (!sx) continue;
    const srels = await relsOf(zip, path.replace(/([^/]+)$/, '_rels/$1.rels'));
    const slide = await parseSlide(zip, parseXml(sx), srels, theme);
    // notes
    for (const id in srels) { if (/notesSlide/.test(srels[id].type)) { let ntPath = srels[id].target.replace(/^\.\.\//, '').replace(/^\//, ''); if (!ntPath.startsWith('ppt/')) ntPath = 'ppt/' + ntPath; const nt = await txt(ntPath); if (nt) slide.notes = esc(all(parseXml(nt), 't').map(x => x.textContent).join(' ')); } }
    slides.push(slide);
  }
  return { slides: slides.length ? slides : [blankSlide('title', theme)], theme, ratio, vba };
}
async function relsOf(zip, path) { const f = zip.file(path); if (!f) return {}; const t = await f.async('string'); const m = {}; all(parseXml(t), 'Relationship').forEach(r => { m[at(r, 'Id')] = { target: at(r, 'Target'), type: at(r, 'Type') }; }); return m; }
function relIdOf(el) {
  if (!el) return null;
  const rid = el.getAttribute('r:id') || (el.getAttributeNS && (el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || el.getAttributeNS('http://schemas.microsoft.com/office/2006/relationships', 'id')));
  if (rid) return rid;
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (a.localName === 'id' && /relationships/.test(a.namespaceURI || '')) return a.value;
  }
  return at(el, 'id');
}
async function parseSlide(zip, doc, rels, theme) {
  const s = { id: 's' + Math.random().toString(36).slice(2), layout: 'blank', objects: [], notes: '', transition: null, bg: null };
  const tr = all(doc, 'transition')[0]; if (tr && tr.firstElementChild) s.transition = mapTransition(tr.firstElementChild.localName);
  const spTree = all(doc, 'spTree')[0]; if (!spTree) return s;
  for (let node = spTree.firstChild; node; node = node.nextSibling) {
    if (node.nodeType !== 1) continue;
    const ln = node.localName;
    if (ln === 'sp') { const o = parseShape(node, theme); if (o) s.objects.push(o); }
    else if (ln === 'pic') { const o = await parsePic(node, rels, zip); if (o) s.objects.push(o); }
    else if (ln === 'graphicFrame') { const o = parseGraphic(node); if (o) s.objects.push(o); }
  }
  { const phs = s.objects.map(o => o.ph).filter(Boolean); if (phs.indexOf('title') >= 0) s.layout = phs.indexOf('subtitle') >= 0 ? 'title' : phs.some(x => /^body/.test(x)) ? (phs.indexOf('body2') >= 0 ? 'two' : 'content') : 'titleonly'; }
  return s;
}
function xfrmOf(sp) { const x = all(sp, 'xfrm')[0]; if (!x) return { x: 60, y: 60, w: 400, h: 200 }; const off = kid(x, 'off'), ext = kid(x, 'ext'); const rot = at(x, 'rot'); return { x: off ? +at(off, 'x') / EMUX : 60, y: off ? +at(off, 'y') / EMUY : 60, w: ext ? +at(ext, 'cx') / EMUX : 400, h: ext ? +at(ext, 'cy') / EMUY : 200, rot: rot ? +rot / 60000 : 0 }; }
function parseShape(sp, theme) {
  const f = xfrmOf(sp);
  const txBody = all(sp, 'txBody')[0];
  let html = '', ph = null; let fontSize = null, align = null, bold = false, color = null;
  const phEl = all(sp, 'ph')[0]; if (phEl) { const t = at(phEl, 'type');
    // v0.0.7: placeholder without its own position inherits it from the layout/master (PowerPoint default 16:9 positions)
    if (!all(sp, 'xfrm')[0]) { const D = { ctrTitle: [120, 130, 1040, 270], subTitle: [160, 410, 960, 190], title: [88, 38, 1104, 139] }[t] || [88, 192, 1104, 457]; f.x = D[0]; f.y = D[1]; f.w = D[2]; f.h = D[3]; } ph = t === 'ctrTitle' || t === 'title' ? 'title' : t === 'subTitle' ? 'subtitle' : t === 'body' ? 'body' : (!t || t === 'obj') && at(phEl, 'idx') ? (at(phEl, 'idx') === '1' ? 'body' : 'body' + at(phEl, 'idx')) : null; }
  if (txBody) {
    const paras = kids(txBody, 'p');
    html = paras.map(p => {
      const runs = kids(p, 'r'); const pPr = kid(p, 'pPr'); if (pPr && at(pPr, 'algn')) align = { ctr: 'center', r: 'right', l: 'left', just: 'justify' }[at(pPr, 'algn')];
      const text = runs.map(r => { const rpr = kid(r, 'rPr'); let t = esc(all(r, 't').map(x => x.textContent).join('')); if (rpr) { if (at(rpr, 'b') === '1') { t = '<b>' + t + '</b>'; bold = true; } if (at(rpr, 'i') === '1') t = '<i>' + t + '</i>'; if (at(rpr, 'sz') && !fontSize) fontSize = +at(rpr, 'sz') / 100; const c = kid(rpr, 'solidFill'); if (c) { const sc = kid(c, 'srgbClr'); if (sc && !color) color = '#' + at(sc, 'val'); } } return t; }).join('');
      const bullet = pPr && kid(pPr, 'buChar');
      return (bullet ? '<div>• ' : '<div>') + (text || '<br>') + '</div>';
    }).join('');
    if (paras.length === 1) html = html.replace(/^<div>|<\/div>$/g, '');
  }
  const fill = all(sp, 'spPr')[0] && kid(kid(all(sp, 'spPr')[0], 'solidFill'), 'srgbClr');
  const geom = all(sp, 'prstGeom')[0]; const shapeType = geom ? at(geom, 'prst') : null;
  if (!txBody && shapeType && shapeType !== 'rect') return { id: 'o' + Math.random().toString(36).slice(2), type: 'shape', shape: mapPrst(shapeType), x: f.x, y: f.y, w: f.w, h: f.h, rot: f.rot, fill: fill ? '#' + at(fill, 'val') : theme.colors[0], stroke: 'none' };
  if (html || ph) return { id: 'o' + Math.random().toString(36).slice(2), type: 'text', ph, x: f.x, y: f.y, w: f.w, h: f.h, rot: f.rot, html, fontSize: fontSize || (ph === 'title' ? 40 : ph === 'subtitle' ? 24 : 18), align, b: bold, color };
  return { id: 'o' + Math.random().toString(36).slice(2), type: 'shape', shape: mapPrst(shapeType || 'rect'), x: f.x, y: f.y, w: f.w, h: f.h, rot: f.rot, fill: fill ? '#' + at(fill, 'val') : theme.colors[0], stroke: 'none', text: html };
}
async function parsePic(pic, rels, zip) {
  const f = xfrmOf(pic); const blip = all(pic, 'blip')[0]; const rid = blip && (at(blip, 'embed') || at(blip, 'link'));
  const rel = rid && rels[rid]; if (!rel) return null;
  const file = zip.file('ppt/' + rel.target.replace(/^\.\.\//, '')); if (!file) return null;
  const b = await file.async('uint8array'); const ext = extOf(rel.target);
  return { id: 'o' + Math.random().toString(36).slice(2), type: 'image', src: bytesToDataUrl(b, ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : ext === 'svg' ? 'image/svg+xml' : 'image/jpeg'), x: f.x, y: f.y, w: f.w, h: f.h, rot: f.rot, alt: (all(pic, 'cNvPr')[0] && at(all(pic, 'cNvPr')[0], 'descr')) || '' };
}
function parseGraphic(gf) {
  const f = xfrmOf(gf); const tbl = all(gf, 'tbl')[0];
  if (tbl) { const rows = kids(tbl, 'tr').map(tr => kids(tr, 'tc').map(tc => esc(all(tc, 't').map(x => x.textContent).join(' ')))); return { id: 'o' + Math.random().toString(36).slice(2), type: 'table', rows, header: true, x: f.x, y: f.y, w: f.w, h: f.h, fontSize: 16 }; }
  const chart = all(gf, 'chart')[0];
  if (chart) return { id: 'o' + Math.random().toString(36).slice(2), type: 'chart', kind: 'column', data: { cats: ['A', 'B', 'C'], series: [{ name: 'Chuỗi', values: [1, 2, 3] }] }, x: f.x, y: f.y, w: f.w, h: f.h };
  return null;
}
function mapTransition(n) { return { fade: 'fade', push: 'push', wipe: 'wipe', split: 'split', cut: 'cut', zoom: 'zoom', morph: 'morph', cover: 'push', pull: 'wipe' }[n] || 'fade'; }
function mapPrst(p) { const m = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', pentagon: 'pentagon', hexagon: 'hexagon', star5: 'star5', rightArrow: 'rightArrow', leftArrow: 'leftArrow', upArrow: 'upArrow', downArrow: 'downArrow', heart: 'heart', cloud: 'cloud', flowChartProcess: 'flowProcess', flowChartDecision: 'flowDecision', chevron: 'chevron', plus: 'plus', line: 'line', parallelogram: 'parallelogram', trapezoid: 'trapezoid', wedgeRectCallout: 'callout', star4: 'star4', smileyFace: 'smiley', bentArrow: 'rightArrow', leftRightArrow: 'leftRightArrow' }; return m[p] || 'rect'; }

// ================================================================ WRITE
function xmlEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function hex(c) { if (!c) return null; c = String(c).replace('#', ''); if (c.length === 3) c = c.split('').map(x => x + x).join(''); return /^[0-9a-f]{6}$/i.test(c) ? c.toUpperCase() : null; }
function plainRuns(html, baseStyle) {
  const div = document.createElement('div'); div.innerHTML = html || '';
  const paras = [];
  const blocks = div.childNodes.length && Array.prototype.some.call(div.childNodes, n => n.nodeType === 1 && /DIV|P/.test(n.tagName)) ? Array.prototype.filter.call(div.childNodes, n => n.nodeType === 1 && /DIV|P/.test(n.tagName)) : [div];
  blocks.forEach(b => { const runs = []; collectRuns(b, {}, runs); const bullet = /^•\s/.test(b.textContent); paras.push({ runs, bullet }); });
  if (!paras.length) paras.push({ runs: [{ text: div.textContent || '', st: {} }], bullet: false });
  return paras;
}
function collectRuns(node, st, out) {
  node.childNodes.forEach(n => {
    if (n.nodeType === 3) { if (n.data) out.push({ text: n.data.replace(/^•\s/, ''), st: Object.assign({}, st) }); return; }
    if (n.nodeType !== 1) return;
    const s2 = Object.assign({}, st); const t = n.tagName;
    if (t === 'B' || t === 'STRONG') s2.b = true; if (t === 'I' || t === 'EM') s2.i = true; if (t === 'U') s2.u = true; if (t === 'BR') { out.push({ text: '\n', st: s2 }); return; }
    if (n.style && n.style.color) s2.color = n.style.color;
    collectRuns(n, s2, out);
  });
}
export async function writePptx(doc, macro) {
  const JSZip = await JSZipLib(); const zip = new JSZip();
  const theme = doc.theme; const images = [];
  const cx = 12192000, cy = 6858000;
  const ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>', '<Default Extension="png" ContentType="image/png"/>', '<Default Extension="jpeg" ContentType="image/jpeg"/>', '<Default Extension="jpg" ContentType="image/jpeg"/>', '<Default Extension="gif" ContentType="image/gif"/>', '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>', '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>', '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>', '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>', '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'];
  const sldIdList = []; const presRels = ['<Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>', '<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'];
  doc.slides.forEach((s, i) => {
    const n = i + 1; const srels = ['<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout' + (LAYOUT_NO[s.layout] || 4) + '.xml"/>'];
    ct.push('<Override PartName="/ppt/slides/slide' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>');
    presRels.push('<Relationship Id="rId' + (10 + n) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + n + '.xml"/>');
    sldIdList.push('<p:sldId id="' + (255 + n) + '" r:id="rId' + (10 + n) + '"/>');
    zip.file('ppt/slides/slide' + n + '.xml', slideXml(s, theme, images, srels, n));
    if (srels.length > 1) zip.file('ppt/slides/_rels/slide' + n + '.xml.rels', relsXml(srels));
    else zip.file('ppt/slides/_rels/slide' + n + '.xml.rels', relsXml(srels));
  });
  images.forEach(im => zip.file('ppt/media/' + im.name, im.bytes));
  zip.file('ppt/presentation.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst><p:sldIdLst>' + sldIdList.join('') + '</p:sldIdLst><p:sldSz cx="' + cx + '" cy="' + cy + '" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>');
  if (macro) {
    const O = await import('../excel/vba/ovba.js'); const { jsaToVba } = await import('../excel/vba/jsa.js');
    const user = ((doc.vba && doc.vba.modules) || []).filter(m => m.name !== '__Immediate' && m.type !== 'document');
    const mods = user.map(m => ({ name: m.name, type: m.type === 'class' ? 'class' : 'module', code: m.type === 'jsa' ? jsaToVba(m.code) : m.code }));
    zip.file('ppt/vbaProject.bin', O.writeVbaProject(mods, { name: 'VBAProject', host: 'ppt' }));
    zip.file('oos/vba-source.json', O.vbaSidecar(user.map(m => ({ name: m.name, type: m.type || 'module', code: m.code, wcode: m.type === 'jsa' ? jsaToVba(m.code) : m.code }))));
    presRels.push('<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>');
    ct.push('<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>', '<Default Extension="json" ContentType="application/json"/>');
    for (let i = 0; i < ct.length; i++) ct[i] = ct[i].replace('application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml', 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml');
  }
  zip.file('ppt/_rels/presentation.xml.rels', relsXml(presRels));
  zip.file('ppt/theme/theme1.xml', themeXml(theme));
  zip.file('ppt/slideMasters/slideMaster1.xml', masterXml());
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', relsXml([1, 2, 3, 4].map(k => '<Relationship Id="rIdL' + k + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout' + k + '.xml"/>').concat(['<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'])));
  [1, 2, 3, 4].forEach(k => {
    zip.file('ppt/slideLayouts/slideLayout' + k + '.xml', layoutXml(k));
    if (k > 1) ct.push('<Override PartName="/ppt/slideLayouts/slideLayout' + k + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>');
    zip.file('ppt/slideLayouts/_rels/slideLayout' + k + '.xml.rels', relsXml(['<Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>']));
  });
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + ct.join('') + '</Types>');
  zip.file('_rels/.rels', relsXml(['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>', '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>']));
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + xmlEsc(doc.name) + '</dc:title><dc:creator>Yomi-Office Open Source</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>');
  zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office PowerPoint</Application><Slides>' + doc.slides.length + '</Slides></Properties>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
function relsXml(list) { return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + list.join('') + '</Relationships>'; }
// v0.0.7: placeholders map onto the real layouts (Title Slide / Title and Content / Title Only / Blank) so PowerPoint
// recognises titles (Shapes.Title, outline view, "New Slide" layouts) in files written by Yomi-Office Open Source.
const LAYOUT_NO = { title: 1, section: 1, content: 2, two: 2, compare: 2, caption: 2, picture: 2, titleonly: 3, blank: 4 };
function phTag(ph, layout) {
  const cover = LAYOUT_NO[layout] === 1;
  if (ph === 'title') return '<p:ph type="' + (cover ? 'ctrTitle' : 'title') + '"/>';
  if (ph === 'subtitle') return cover ? '<p:ph type="subTitle" idx="1"/>' : '<p:ph idx="1"/>';
  if (ph === 'body') return '<p:ph idx="1"/>';
  const m = /(\d+)$/.exec(ph || ''); return '<p:ph idx="' + (m ? +m[1] : 2) + '"/>';
}
function slideXml(s, theme, images, srels, n) {
  let shapes = ''; let idc = 2;
  const E = (px, which) => Math.round(px * (which === 'x' ? EMUX : EMUY));
  s.objects.forEach(o => {
    const xfrm = '<a:xfrm' + (o.rot ? ' rot="' + Math.round(o.rot * 60000) + '"' : '') + '><a:off x="' + E(o.x, 'x') + '" y="' + E(o.y, 'y') + '"/><a:ext cx="' + E(o.w, 'x') + '" cy="' + E(o.h, 'y') + '"/></a:xfrm>';
    if (o.type === 'text' || (o.type === 'shape' && !o.shape)) {
      const paras = plainRuns(o.html);
      const body = paras.map(p => '<a:p>' + (p.bullet ? '<a:pPr' + (o.align ? ' algn="' + algn(o.align) + '"' : '') + '><a:buChar char="•"/></a:pPr>' : (o.align ? '<a:pPr algn="' + algn(o.align) + '"/>' : '')) + p.runs.map(r => runXml(r, o)).join('') + '</a:p>').join('') || '<a:p><a:endParaRPr/></a:p>';
      const phXml = o.ph ? '<p:nvSpPr><p:cNvPr id="' + (idc++) + '" name="' + (o.ph) + '"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' + phTag(o.ph, s.layout) + '</p:nvPr></p:nvSpPr>' : '<p:nvSpPr><p:cNvPr id="' + (idc++) + '" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>';
      shapes += '<p:sp>' + phXml + '<p:spPr>' + xfrm + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap="square"' + (o.valign === 'middle' ? ' anchor="ctr"' : o.valign === 'bottom' ? ' anchor="b"' : '') + '><a:normAutofit/></a:bodyPr><a:lstStyle/>' + body + '</p:txBody></p:sp>';
    } else if (o.type === 'shape') {
      const fill = hex(o.fill) ? '<a:solidFill><a:srgbClr val="' + hex(o.fill) + '"/></a:solidFill>' : '<a:noFill/>';
      const text = o.text ? '<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>' + plainRuns(o.text).map(p => '<a:p><a:pPr algn="ctr"/>' + p.runs.map(r => runXml(r, { color: o.color || '#FFFFFF', b: o.b })).join('') + '</a:p>').join('') + '</p:txBody>' : '<p:txBody><a:bodyPr/><a:p/></p:txBody>';
      shapes += '<p:sp><p:nvSpPr><p:cNvPr id="' + (idc++) + '" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' + xfrm + '<a:prstGeom prst="' + prstOf(o.shape) + '"><a:avLst/></a:prstGeom>' + fill + (hex(o.stroke) ? '<a:ln><a:solidFill><a:srgbClr val="' + hex(o.stroke) + '"/></a:solidFill></a:ln>' : '') + '</p:spPr>' + text + '</p:sp>';
    } else if (o.type === 'image') {
      const m = /^data:image\/(png|jpe?g|gif)/.exec(o.src || ''); if (!m) return; const ext = m[1] === 'jpg' ? 'jpeg' : m[1];
      const name = 'image' + (images.length + 1) + '.' + ext; images.push({ name, bytes: dataUrlToBytes(o.src) });
      const rid = 'rIdImg' + images.length; srels.push('<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' + name + '"/>');
      shapes += '<p:pic><p:nvPicPr><p:cNvPr id="' + (idc++) + '" name="Picture" descr="' + xmlEsc(o.alt || '') + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>' + xfrm + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
    } else if (o.type === 'table') {
      shapes += tableXml(o, E);
    } else if (o.type === 'equation') {
      shapes += '<p:sp><p:nvSpPr><p:cNvPr id="' + (idc++) + '" name="Equation"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>' + xfrm + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="2400"/><a:t>' + xmlEsc(o.latex || '') + '</a:t></a:r></a:p></p:txBody></p:sp>';
    } else if (o.type === 'chart') {
      // represent chart as a titled placeholder box (data table) to keep file valid without a chart part
      shapes += '<p:sp><p:nvSpPr><p:cNvPr id="' + (idc++) + '" name="Chart"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' + xfrm + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1400" b="1"/><a:t>' + xmlEsc('[Biểu đồ] ' + (o.title || '')) + '</a:t></a:r></a:p></p:txBody></p:sp>';
    }
  });
  const bg = s.bg && hex(typeof s.bg === 'string' ? s.bg : (s.bg.color || '')) ? '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="' + hex(typeof s.bg === 'string' ? s.bg : s.bg.color) + '"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>' : '';
  const trans = s.transition ? '<p:transition><p:' + trPrst(s.transition) + '/></p:transition>' : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>' + bg + '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + shapes + '</p:spTree></p:cSld>' + trans + '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>';
}
function runXml(r, o) { const st = r.st || {}; const sz = o && o.fontSize ? ' sz="' + Math.round(o.fontSize * 100) + '"' : ''; const b = (st.b || (o && o.b)) ? ' b="1"' : ''; const it = st.i ? ' i="1"' : ''; const u = st.u ? ' u="sng"' : ''; const col = hex(st.color || (o && o.color)); const fill = col ? '<a:solidFill><a:srgbClr val="' + col + '"/></a:solidFill>' : ''; const font = o && o.font ? '<a:latin typeface="' + xmlEsc(o.font) + '"/>' : ''; return '<a:r><a:rPr lang="vi-VN"' + sz + b + it + u + '>' + fill + font + '</a:rPr><a:t>' + xmlEsc(r.text) + '</a:t></a:r>'; }
function algn(a) { return { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[a] || 'l'; }
function prstOf(shape) { const m = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', pentagon: 'pentagon', hexagon: 'hexagon', octagon: 'octagon', star5: 'star5', star4: 'star4', heart: 'heart', cloud: 'cloud', rightArrow: 'rightArrow', leftArrow: 'leftArrow', upArrow: 'upArrow', downArrow: 'downArrow', leftRightArrow: 'leftRightArrow', chevron: 'chevron', plus: 'plus', parallelogram: 'parallelogram', trapezoid: 'trapezoid', flowProcess: 'flowChartProcess', flowDecision: 'flowChartDecision', flowTerminator: 'flowChartTerminator', callout: 'wedgeRectCallout', smiley: 'smileyFace', line: 'line', lightning: 'lightningBolt', moon: 'moon', donut: 'donut' }; return m[shape] || 'rect'; }
function trPrst(t) { return { fade: 'fade', push: 'push', wipe: 'wipe', split: 'split', cut: 'cut', zoom: 'zoom', morph: 'fade', fall: 'fade', flip: 'fade' }[t] || 'fade'; }
function tableXml(o, E) {
  const rows = o.rows || []; const ncol = rows[0] ? rows[0].length : 1; const cw = Math.round(E(o.w, 'x') / ncol);
  const rh = Math.round(E(o.h, 'y') / Math.max(1, rows.length));
  let grid = '<a:tblGrid>' + Array(ncol).fill('<a:gridCol w="' + cw + '"/>').join('') + '</a:tblGrid>';
  const trs = rows.map((row, ri) => '<a:tr h="' + rh + '">' + row.map(cell => '<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="' + ((o.fontSize || 16) * 100) + '"' + (ri === 0 && o.header ? ' b="1"' : '') + '/><a:t>' + xmlEsc((cell || '').replace(/<[^>]+>/g, '')) + '</a:t></a:r></a:p></a:txBody><a:tcPr' + (ri === 0 && o.header ? '><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:tcPr>' : '/>') + '</a:tc>').join('') + '</a:tr>').join('');
  return '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="50" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="' + E(o.x, 'x') + '" y="' + E(o.y, 'y') + '"/><a:ext cx="' + E(o.w, 'x') + '" cy="' + E(o.h, 'y') + '"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>' + grid + trs + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
}
function themeXml(theme) {
  const acc = (theme.colors || THEMES[0].colors).map(hex);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="' + xmlEsc(theme.name || 'Office') + '"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' + acc.map((c, i) => '<a:accent' + (i + 1) + '><a:srgbClr val="' + (c || '4472C4') + '"/></a:accent' + (i + 1) + '>').join('') + '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="' + xmlEsc(theme.font || 'Calibri') + '"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="' + xmlEsc(theme.font || 'Calibri') + '"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';
}
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const CLR = 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';
function phSp(id, name, ph, x, y, w, hh, text) {
  const xf = x == null ? '<p:spPr/>' : '<p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + w + '" cy="' + hh + '"/></a:xfrm></p:spPr>';
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + name + '"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' + ph + '</p:nvPr></p:nvSpPr>' + xf + '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="vi-VN"/><a:t>' + text + '</a:t></a:r></a:p></p:txBody></p:sp>';
}
function masterXml() {
  const lvl = (n, sz) => '<a:lvl' + n + 'pPr marL="' + (228600 + (n - 1) * 457200) + '" indent="-228600"><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="' + sz + '"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl' + n + 'pPr>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sldMaster ' + NS + '><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>' + GRP +
    phSp(2, 'Title Placeholder 1', '<p:ph type="title"/>', 838200, 365125, 10515600, 1325563, 'Bấm để sửa tiêu đề') +
    phSp(3, 'Text Placeholder 2', '<p:ph type="body" idx="1"/>', 838200, 1825625, 10515600, 4351338, 'Bấm để sửa nội dung') +
    '</p:spTree></p:cSld><p:clrMap ' + CLR + '/><p:sldLayoutIdLst>' + [1, 2, 3, 4].map(k => '<p:sldLayoutId id="' + (2147483648 + k) + '" r:id="rIdL' + k + '"/>').join('') + '</p:sldLayoutIdLst>' +
    '<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle>' + lvl(1, 2800) + lvl(2, 2400) + lvl(3, 2000) + '</p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>';
}
function layoutXml(k) {
  const L = {
    1: ['title', 'Title Slide', phSp(2, 'Title 1', '<p:ph type="ctrTitle"/>', 1524000, 1122363, 9144000, 2387600, 'Bấm để sửa tiêu đề') + phSp(3, 'Subtitle 2', '<p:ph type="subTitle" idx="1"/>', 1524000, 3602038, 9144000, 1655762, 'Bấm để sửa phụ đề')],
    2: ['obj', 'Title and Content', phSp(2, 'Title 1', '<p:ph type="title"/>', null, 0, 0, 0, 'Bấm để sửa tiêu đề') + phSp(3, 'Content Placeholder 2', '<p:ph idx="1"/>', null, 0, 0, 0, 'Bấm để sửa nội dung')],
    3: ['titleOnly', 'Title Only', phSp(2, 'Title 1', '<p:ph type="title"/>', null, 0, 0, 0, 'Bấm để sửa tiêu đề')],
    4: ['blank', 'Blank', '']
  }[k || 4];
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sldLayout ' + NS + ' type="' + L[0] + '" preserve="1"><p:cSld name="' + L[1] + '"><p:spTree>' + GRP + L[2] + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
}
