// Yomi-Office Open Source v0.0.6 — Word VBA object model for the shared VBA engine (../excel/vba/interp.js).
// Integrates Word macros such as those from the Project GitHub sheet (aminya/AcuteMSWord, TLE-11/nsfc-md2docx, linhut/document-ai-assistant).
// Supported: Application, ActiveDocument (Content, Range, Paragraphs, Words, Characters, Tables, Save, ComputeStatistics…),
// Selection (Text, TypeText, TypeParagraph, HomeKey/EndKey, WholeStory, Font, ParagraphFormat, Find/Replace incl. wildcards,
// Copy/Cut/Paste, InsertBefore/After, Collapse…), Range objects, Find/Replacement, Font, ParagraphFormat (+ Shading), Tables/Cells.
// Layout-only settings Word has but a web editor does not (Options, Borders details, widow control…) are accepted and ignored.
import { VbaError, VbArray, Collection, Dictionary, NOT_FOUND, toNum, toStr, truthy } from '../excel/vba/interp.js';
import { CONSTS, builtins } from '../excel/vba/objects.js';

const WD = {
  wdstory: 6, wdline: 5, wdcharacter: 1, wdword: 2, wdsentence: 3, wdparagraph: 4, wdcell: 12, wdmove: 0, wdextend: 1,
  wdreplacenone: 0, wdreplaceone: 1, wdreplaceall: 2, wdfindstop: 0, wdfindcontinue: 1, wdfindask: 2,
  wdalignparagraphleft: 0, wdalignparagraphcenter: 1, wdalignparagraphright: 2, wdalignparagraphjustify: 3,
  wdcollapseend: 0, wdcollapsestart: 1, wdstatisticwords: 0, wdstatisticlines: 1, wdstatisticpages: 2, wdstatisticcharacters: 3, wdstatisticparagraphs: 4, wdstatisticcharacterswithspaces: 5,
  wdlinespacesingle: 0, wdlinespace1pt5: 1, wdlinespacedouble: 2, wdlinespaceatleast: 3, wdlinespaceexactly: 4, wdlinespacemultiple: 5,
  wdundefined: 9999999, wdtoggle: 9999998, wdunderlinenone: 0, wdunderlinesingle: 1, wdunderlinedouble: 3,
  wdnohighlight: 0, wdblack: 1, wdblue: 2, wdturquoise: 3, wdbrightgreen: 4, wdpink: 5, wdred: 6, wdyellow: 7, wdwhite: 8, wddarkblue: 9, wdteal: 10, wdgreen: 11, wdviolet: 12, wddarkred: 13, wddarkyellow: 14, wdgray50: 15, wdgray25: 16,
  wdcolorautomatic: -16777216, wdcolorblack: 0, wdcolorwhite: 16777215, wdcolorred: 255, wdcolorgreen: 65280, wdcolorblue: 16711680, wdcoloryellow: 65535, wdcolorgray25: 12632256,
  wdformatdocumentdefault: 16, wdformatxmldocument: 12, wdformatpdf: 17, wdformattext: 2, wdformatrtf: 6, wdformatfilteredhtml: 10,
  wdheadingnumbering: 0, wdstyleheading1: -2, wdstyleheading2: -3, wdstyleheading3: -4, wdstylenormal: -1, wdstyletitle: -63
};
const HIGHLIGHT = { 1: '#000000', 2: '#0000ff', 3: '#00ffff', 4: '#00ff00', 5: '#ff00ff', 6: '#ff0000', 7: '#ffff00', 8: '#ffffff', 9: '#000080', 10: '#008080', 11: '#008000', 12: '#800080', 13: '#800000', 14: '#808000', 15: '#808080', 16: '#c0c0c0' };
const BLOCK = /^(P|H[1-6]|LI|DIV|BLOCKQUOTE|PRE|TD|TH|TR|TABLE|UL|OL)$/;
const hexOf = n => { n = toNum(n); if (n < 0) return null; const r = n & 255, g = (n >> 8) & 255, b = (n >> 16) & 255; return '#' + [r, g, b].map(x => (x < 16 ? '0' : '') + x.toString(16)).join(''); };
const pt = v => toNum(v) * 96 / 72;   // points → px

class WObj { }
class WIgnore extends WObj { }              // accepts any member get/set (Options, Borders, …)
class WApp extends WObj { constructor(m) { super(); this.m = m; } }
class WDoc extends WObj { constructor(m) { super(); this.m = m; } }
class WSel extends WObj { constructor(m) { super(); this.m = m; } }
class WRange extends WObj { constructor(m, a, b) { super(); this.m = m; this.a = a; this.b = b; } }
class WFont extends WObj { constructor(m, target) { super(); this.m = m; this.t = target; } }
class WPara extends WObj { constructor(m, target) { super(); this.m = m; this.t = target; } }
class WFind extends WObj { constructor(m, target) { super(); this.m = m; this.t = target; this.text = ''; this.repl = { __repl: true, text: '' }; this.matchCase = false; this.whole = false; this.wild = false; this.forward = true; this.wrap = 1; } }
class WColl extends WObj { constructor(items, make) { super(); this.items = items; this.make = make; } each() { return this.items.map(this.make); } }
class WTable extends WObj { constructor(m, el) { super(); this.m = m; this.el = el; } }

// ---------------------------------------------------------------- text ⇄ DOM mapping (Word "character positions")
function textMap(body) {
  const segs = []; let text = ''; let lastBlock = null;
  const walk = node => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) { segs.push({ node: n, start: text.length, end: text.length + n.nodeValue.length }); text += n.nodeValue; continue; }
      if (n.nodeType !== 1) continue;
      if (n.tagName === 'BR') { segs.push({ br: n, start: text.length, end: text.length + 1 }); text += '\v'; continue; }
      walk(n);
      if (BLOCK.test(n.tagName) && !/^(UL|OL|TABLE|TR)$/.test(n.tagName)) { if (n.tagName === 'TD' || n.tagName === 'TH') text += '\u0007'; else { segs.push({ block: n, start: text.length, end: text.length + 1 }); text += '\r'; } }
    }
  };
  walk(body);
  return { text, segs };
}
function domPoint(map, pos, body) {
  for (const s of map.segs) { if (s.node && pos >= s.start && pos <= s.end) return { node: s.node, offset: pos - s.start }; }
  // position at a paragraph mark / br: place after the previous text
  let best = null; for (const s of map.segs) { if (s.start <= pos) best = s; }
  if (best && best.node) return { node: best.node, offset: best.node.nodeValue.length };
  if (best && (best.block || best.br)) { const el = best.block || best.br; return { node: el.parentNode, offset: Array.prototype.indexOf.call(el.parentNode.childNodes, el) + 1 }; }
  return { node: body, offset: 0 };
}
export class WordModel {
  constructor(doc, ui) { this.doc = doc; this.ui = ui; this.app = new WApp(this); this.adoc = new WDoc(this); this.sel = new WSel(this); this.ign = new WIgnore(); this.screenUpdating = true; }
  get body() { return this.doc.body; }
  map() { return textMap(this.body); }
  domRange(a, b) { const mp = this.map(); const r = document.createRange(); const p1 = domPoint(mp, a, this.body), p2 = domPoint(mp, b, this.body); try { r.setStart(p1.node, p1.offset); r.setEnd(p2.node, p2.offset); } catch (e) { r.selectNodeContents(this.body); r.collapse(false); } return r; }
  selOffsets() { const s = window.getSelection(); const mp = this.map(); if (!s.rangeCount || !this.body.contains(s.anchorNode)) { const sv = this.doc._savedRange; if (!sv) return [mp.text.length, mp.text.length]; return this.offsetsOf(sv, mp); } return this.offsetsOf(s.getRangeAt(0), mp); }
  offsetsOf(r, mp) { mp = mp || this.map(); const off = (node, o) => { if (node.nodeType === 3) { const sg = mp.segs.find(x => x.node === node); return sg ? sg.start + o : 0; } const kid = node.childNodes[o]; if (!kid) { let last = 0; mp.segs.forEach(sg => { const n = sg.node || sg.block || sg.br; if (n && node.contains(n)) last = sg.end; }); return last; } for (const sg of mp.segs) { const n = sg.node || sg.block || sg.br; if (n && (n === kid || kid.contains(n))) return sg.start; } return mp.text.length; }; return [off(r.startContainer, r.startOffset), off(r.endContainer, r.endOffset)]; }
  select(a, b) { const r = this.domRange(a, b); const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); this.doc._savedRange = r.cloneRange(); }
  textOf(a, b) { return this.map().text.slice(a, b); }
  setText(a, b, t) {
    let r;
    if (a <= 0 && b >= this.map().text.length) { this.body.innerHTML = '<p></p>'; r = document.createRange(); r.setStart(this.body.firstChild, 0); r.collapse(true); } // whole story replaced
    else { r = this.domRange(a, b); r.deleteContents(); }
    // range collapsed directly in the body (whole content replaced): put the text into a paragraph, never loose in <body>
    if (r.startContainer === this.body) {
      const nx = this.body.childNodes[r.startOffset]; let blk = nx && nx.nodeType === 1 && BLOCK.test(nx.tagName) ? nx : null;
      if (blk && !blk.textContent && !/^(TABLE|UL|OL)$/.test(blk.tagName)) { blk.innerHTML = ''; } else if (!blk || /^(TABLE|UL|OL)$/.test(blk.tagName)) { blk = document.createElement('p'); this.body.insertBefore(blk, nx || null); }
      else { r.setStart(blk, 0); r.collapse(true); blk = null; }
      if (blk) { r.setStart(blk, 0); r.collapse(true); }
    }
    const parts = String(t).split(/\r\n|\r|\n/);
    if (parts.length === 1) { const tn = document.createTextNode(parts[0]); r.insertNode(tn); return a + parts[0].length; }
    // multi-paragraph text: split the current paragraph
    let blk = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement; while (blk && blk !== this.body && !BLOCK.test(blk.tagName)) blk = blk.parentElement;
    const first = document.createTextNode(parts[0]); r.insertNode(first);
    let after = blk && blk !== this.body ? blk : null; let pos = a + parts[0].length + 1;
    const tailRange = document.createRange(); tailRange.setStartAfter(first); if (after) tailRange.setEndAfter(after.lastChild || after); const tail = after ? tailRange.extractContents() : null;
    for (let i = 1; i < parts.length; i++) { const p = document.createElement(after && /^H[1-6]$/.test(after.tagName) ? 'p' : (after ? after.tagName.toLowerCase() === 'li' ? 'li' : 'p' : 'p')); p.appendChild(document.createTextNode(parts[i])); if (i === parts.length - 1 && tail) p.appendChild(tail); if (!p.firstChild || (!p.textContent && !p.querySelector('img'))) p.innerHTML = '<br>'; if (after) after.parentNode.insertBefore(p, after.nextSibling); else this.body.appendChild(p); after = p; pos += parts[i].length + (i < parts.length - 1 ? 1 : 0); }
    return pos;
  }
  blocksIn(a, b) { const mp = this.map(); const out = []; mp.segs.forEach(sg => { if (sg.block && sg.start >= a && sg.start <= Math.max(a, b) && out.indexOf(sg.block) < 0) out.push(sg.block); }); if (!out.length) { const p = domPoint(mp, a, this.body); let e = p.node.nodeType === 1 ? p.node : p.node.parentElement; while (e && e !== this.body && !BLOCK.test(e.tagName)) e = e.parentElement; if (e && e !== this.body) out.push(e); } return out; }
  paragraphs() { const mp = this.map(); const out = []; let start = 0; mp.segs.forEach(sg => { if (sg.block && !/^(TD|TH)$/.test(sg.block.tagName)) { out.push([start, sg.end]); start = sg.end; } }); if (start < mp.text.length) out.push([start, mp.text.length]); return out; }
  // merge adjacent <span>s with identical inline style (left behind by Font.* on word/character ranges)
  tidy() {
    const spans = Array.prototype.slice.call(this.body.querySelectorAll('span[style]'));
    spans.forEach(sp => {
      if (!sp.parentNode || sp.attributes.length !== 1) return;
      let nx = sp.nextSibling;
      while (nx && nx.nodeType === 1 && nx.tagName === 'SPAN' && nx.attributes.length === 1 && nx.getAttribute('style') === sp.getAttribute('style')) { while (nx.firstChild) sp.appendChild(nx.firstChild); const dead = nx; nx = nx.nextSibling; dead.parentNode.removeChild(dead); }
      sp.normalize();
    });
    this._map = null;
  }
  // inline styling on a character range (wrap each text piece in a span)
  styleRange(a, b, css) {
    if (b <= a) { const blk = this.blocksIn(a, a)[0]; if (blk && css.fontWeight == null && css.fontStyle == null) return; }
    const mp = this.map();
    mp.segs.filter(sg => sg.node && sg.end > a && sg.start < b).forEach(sg => {
      let node = sg.node; const s0 = Math.max(0, a - sg.start), s1 = Math.min(node.nodeValue.length, b - sg.start);
      if (s1 < node.nodeValue.length) node.splitText(s1); if (s0 > 0) node = node.splitText(s0);
      const par = node.parentNode;
      // reuse the wrapping span when it holds exactly this text (avoids nested <span><span>… after several Font.* assignments)
      if (par && par.tagName === 'SPAN' && par.childNodes.length === 1) { for (const k in css) { if (css[k] === null) continue; par.style[k] = css[k]; } return; }
      const span = document.createElement('span'); for (const k in css) { if (css[k] === null) continue; span.style[k] = css[k]; }
      par.insertBefore(span, node); span.appendChild(node);
    });
  }
  fontValue(a, b, prop) { const mp = this.map(); const sg = mp.segs.find(x => x.node && x.node.nodeValue && x.end > a && x.start < Math.max(b, a + 1)); let el; if (sg) el = sg.node.parentElement; else { const p = domPoint(mp, a, this.body); el = p.node.nodeType === 1 ? p.node : p.node.parentElement; } const cs = getComputedStyle(el || this.body); switch (prop) { case 'bold': return +cs.fontWeight >= 600 || cs.fontWeight === 'bold'; case 'italic': return cs.fontStyle === 'italic'; case 'underline': return /underline/.test(cs.textDecorationLine || cs.textDecoration) ? 1 : 0; case 'size': return Math.round(parseFloat(cs.fontSize) * 72 / 96 * 2) / 2; case 'name': return cs.fontFamily.split(',')[0].replace(/['"]/g, ''); case 'color': { const m = /(\d+),\s*(\d+),\s*(\d+)/.exec(cs.color); return m ? (+m[1]) + (+m[2]) * 256 + (+m[3]) * 65536 : 0; } } return 0; }
  // ---- Find / Replace on the document text
  findRe(f) {
    let src = f.text; let flags = f.matchCase ? 'g' : 'gi';
    if (f.wild) { // Word wildcards → RegExp
      src = src.replace(/\\([\\\[\]()*?@<>{}!])/g, (m, c) => '\u0000' + c.charCodeAt(0) + '\u0000')
        .replace(/[.+^$|]/g, '\\$&').replace(/\*/g, '[\\s\\S]*?').replace(/\?/g, '[\\s\\S]').replace(/@/g, '+').replace(/</g, '\\b').replace(/>/g, '\\b').replace(/\[!/g, '[^').replace(/\{(\d*),(\d*)\}/g, '{$1,$2}')
        .replace(/\u0000(\d+)\u0000/g, (m, c) => '\\' + String.fromCharCode(+c));
    } else { src = src.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\^p/g, '\\r').replace(/\^t/g, '\\t').replace(/\^\?/g, '.'); if (f.whole) src = '\\b' + src + '\\b'; }
    return new RegExp(src, flags);
  }
  execFind(f, target, replaceMode, replaceWith) {
    if (!f.text) return false;
    const re = this.findRe(f); const mp = this.map(); const scopeA = target instanceof WRange ? target.a : 0, scopeB = target instanceof WRange ? target.b : mp.text.length;
    const fromSel = target instanceof WSel; const [sa, sb] = fromSel ? this.selOffsets() : [scopeA, scopeA];
    const repl = (replaceWith != null ? replaceWith : f.repl.text);
    const conv = s => String(s).replace(/\^p/g, '\n').replace(/\^t/g, '\t').replace(/\\(\d)/g, '$$$1').replace(/\^&/g, '$$&');
    if (replaceMode === 2) { // replace all, from the end so offsets stay valid
      const hits = []; let m; re.lastIndex = 0; const hay = mp.text.slice(scopeA, scopeB); while ((m = re.exec(hay))) { if (!m[0].length) { re.lastIndex++; continue; } hits.push({ i: scopeA + m.index, len: m[0].length, rep: m[0].replace(new RegExp(re.source, re.flags.replace('g', '')), conv(repl)) }); }
      for (let k = hits.length - 1; k >= 0; k--) this.setText(hits[k].i, hits[k].i + hits[k].len, hits[k].rep);
      this.found = hits.length; return hits.length > 0;
    }
    re.lastIndex = fromSel ? (f.forward ? sb : 0) : scopeA; let m = re.exec(mp.text);
    if ((!m || (!fromSel && m.index >= scopeB)) && f.wrap === 1 && fromSel) { re.lastIndex = 0; m = re.exec(mp.text); }
    if (!m || (!fromSel && m.index + m[0].length > scopeB)) return false;
    let a = m.index, b = m.index + m[0].length;
    if (replaceMode === 1) { b = this.setText(a, b, m[0].replace(new RegExp(re.source, re.flags.replace('g', '')), conv(repl))); }
    if (fromSel) this.select(a, b); else if (target instanceof WRange) { target.a = a; target.b = b; }
    return true;
  }
}

// ---------------------------------------------------------------- dispatch
export function installWord(m) {
  const B = builtins(Object.assign(m, { v: v => (v instanceof WRange ? m.textOf(v.a, v.b) : v), typeName: v => typeName(v), create: create, evaluate: () => { throw new VbaError('Evaluate chỉ có trong Excel', 0, 438); } }), null);
  function create(cls) { cls = String(cls).toLowerCase(); if (cls === 'collection') return new Collection(); if (/dictionary$/.test(cls)) return new Dictionary(); if (/regexp$/.test(cls)) return new (B.__regexp || Object)(); throw new VbaError('Không tạo được đối tượng ' + cls + ' (chế độ an toàn)', 0, 429); }
  function typeName(v) { if (v && v.__user) return v.mod.type === 'class' ? v.mod.name : typeName(v.hostObj); if (v === undefined) return 'Empty'; if (v === null) return 'Nothing'; if (typeof v === 'number') return Number.isInteger(v) ? 'Long' : 'Double'; if (typeof v === 'string') return 'String'; if (typeof v === 'boolean') return 'Boolean'; if (v instanceof Date) return 'Date'; if (v instanceof VbArray) return 'Variant()'; if (v instanceof WRange) return 'Range'; if (v instanceof WSel) return 'Selection'; if (v instanceof WDoc) return 'Document'; if (v instanceof WApp) return 'Application'; return 'Object'; }
  const rangeOf = t => t instanceof WSel ? (() => { const o = m.selOffsets(); return new WRange(m, o[0], o[1]); })() : t;
  function member(o, name, a, ip) {
    if (o == null) throw new VbaError('Biến đối tượng chưa được gán (Object variable not set) khi truy cập ".' + name + '"', 0, 91);
    if (o instanceof Collection || o instanceof Dictionary) return o.vcall(name, Array.from(a));
    if (o instanceof WIgnore) return name === 'count' ? 0 : m.ign;
    if (o instanceof WColl) { if (name === 'count') return o.items.length; if (name === '' || name === 'item') return o.make(o.items[toNum(a[0]) - 1]); if (name === 'first') return o.make(o.items[0]); if (name === 'last') return o.make(o.items[o.items.length - 1]); if (name === 'update' || name === 'add') return 0; return m.ign; }
    if (o instanceof WApp) { switch (name) { case 'activedocument': case 'thisdocument': return m.adoc; case 'selection': return m.sel; case 'documents': return new WColl([1], () => m.adoc); case 'screenupdating': return m.screenUpdating; case 'name': return 'Microsoft Word'; case 'version': return '16.0'; case 'username': return m.ui.userName || 'User'; case 'options': return m.ign; case 'statusbar': return ''; case 'run': m.ui.print('Application.Run "' + toStr(a[0] || (a.named && a.named.macroname)) + '" — macro ngoài (MathType/AxMath…) không có trong ứng dụng, đã bỏ qua.'); return; case 'centimeterstopoints': return toNum(a[0]) * 72 / 2.54; case 'inchestopoints': return toNum(a[0]) * 72; case 'quit': return; } return m.ign; }
    if (o instanceof WDoc) {
      switch (name) {
        case 'content': return new WRange(m, 0, m.map().text.length);
        case 'range': { const len = m.map().text.length; return new WRange(m, a[0] == null ? 0 : Math.max(0, toNum(a[0])), a[1] == null ? len : Math.min(len, toNum(a[1]))); }
        case 'paragraphs': { const ps = m.paragraphs(); return a.length ? new WRange(m, ps[toNum(a[0]) - 1][0], ps[toNum(a[0]) - 1][1]) : new WColl(ps, p => new WRange(m, p[0], p[1])); }
        case 'words': { const t = m.map().text; const ws = []; const re = /[^\s\r\v\u0007]+\s*/g; let x; while ((x = re.exec(t))) ws.push([x.index, x.index + x[0].length]); return a.length ? new WRange(m, ws[toNum(a[0]) - 1][0], ws[toNum(a[0]) - 1][1]) : new WColl(ws, p => new WRange(m, p[0], p[1])); }
        case 'characters': { const n = m.map().text.length; return a.length ? new WRange(m, toNum(a[0]) - 1, toNum(a[0])) : { __count: n, __coll: [] }; }
        case 'tables': { const ts = Array.prototype.slice.call(m.body.querySelectorAll('table')); return a.length ? new WTable(m, ts[toNum(a[0]) - 1]) : new WColl(ts, t => new WTable(m, t)); }
        case 'fields': case 'bookmarks': case 'inlineshapes': case 'shapes': case 'comments': case 'revisions': case 'storyranges': case 'omaths': case 'sections': return new WColl([], x => x);
        case 'name': return (m.doc.name || 'Document1') + '.docx'; case 'fullname': return m.doc.path || m.doc.name; case 'path': return m.doc.path ? String(m.doc.path).replace(/[\\/][^\\/]*$/, '') : '';
        case 'saved': return !m.doc.dirty; case 'save': m.saveRequested = 'save'; return; case 'saveas': case 'saveas2': m.saveRequested = 'saveas'; return; case 'close': case 'printout': case 'activate': return;
        case 'computestatistics': { const t = m.map().text; const k = toNum(a[0]); if (k === 0) return (t.match(/[^\s\r\v\u0007]+/g) || []).length; if (k === 3) return t.replace(/[\s\r\v\u0007]/g, '').length; if (k === 5) return t.replace(/[\r\u0007]/g, '').length; if (k === 4) return m.paragraphs().length; if (k === 2) return m.doc.pageCount || 1; if (k === 1) return t.split(/[\r\v]/).length; return 0; }
        case 'undo': m.doc.undo(); return true; case 'redo': m.doc.redo(); return true;
        case 'select': m.select(0, m.map().text.length); return;
      }
      return m.ign;
    }
    if (o instanceof WSel || o instanceof WRange) {
      const R = rangeOf(o);
      switch (name) {
        case '': case 'text': return m.textOf(R.a, R.b);
        case 'range': return new WRange(m, R.a, R.b);
        case 'start': return R.a; case 'end': return R.b;
        case 'font': return new WFont(m, o);
        case 'paragraphformat': return new WPara(m, o);
        case 'paragraphs': { const ps = m.paragraphs().filter(p => p[1] > R.a && p[0] <= Math.max(R.a, R.b)); return a.length ? new WRange(m, ps[toNum(a[0]) - 1][0], ps[toNum(a[0]) - 1][1]) : new WColl(ps, p => new WRange(m, p[0], p[1])); }
        case 'find': { if (!o._find) o._find = new WFind(m, o); return o._find; }
        case 'words': { const t = m.textOf(R.a, R.b); const ws = []; const re = /[^\s\r\v\u0007]+\s*/g; let x; while ((x = re.exec(t))) ws.push([R.a + x.index, R.a + x.index + x[0].length]); if (a.length) { const w = ws[toNum(a[0]) - 1]; if (!w) throw new VbaError(5941, 'Thành viên được yêu cầu của tập hợp không tồn tại.'); return new WRange(m, w[0], w[1]); } return new WColl(ws, w => new WRange(m, w[0], w[1])); }
        case 'characters': { const n = R.b - R.a; if (a.length) { const i = toNum(a[0]); if (i < 1 || i > n) throw new VbaError(5941, 'Thành viên được yêu cầu của tập hợp không tồn tại.'); return new WRange(m, R.a + i - 1, R.a + i); } const cs = []; for (let i = R.a; i < R.b; i++) cs.push(i); return new WColl(cs, i => new WRange(m, i, i + 1)); }
        case 'bold': return m.fontValue(R.a, R.b, 'bold'); case 'italic': return m.fontValue(R.a, R.b, 'italic');
        case 'select': m.select(R.a, R.b); return;
        case 'wholestory': m.select(0, m.map().text.length); return;
        case 'homekey': { const u = a[0] == null ? 5 : toNum(a[0]); if (u === 6) m.select(0, 0); else { const ps = m.paragraphs(); const p = ps.find(x => R.a >= x[0] && R.a < x[1]) || [R.a]; m.select(p[0], p[0]); } return; }
        case 'endkey': { const u = a[0] == null ? 5 : toNum(a[0]); const len = m.map().text.length; if (u === 6) m.select(len, len); else { const ps = m.paragraphs(); const p = ps.find(x => R.a >= x[0] && R.a < x[1]); const e = p ? p[1] - 1 : len; m.select(e, e); } return; }
        case 'moveright': case 'moveleft': { const n = a[1] == null ? 1 : toNum(a[1]); const ext = toNum(a[2] || 0) === 1; const d = name === 'moveright' ? n : -n; const len = m.map().text.length; const p = Math.max(0, Math.min(len, (d > 0 ? R.b : R.a) + d)); if (ext) m.select(R.a, Math.max(R.a, p)); else m.select(p, p); return; }
        case 'collapse': { const d = a[0] == null ? 1 : toNum(a[0]); if (o instanceof WRange) { if (d === 0) o.a = o.b; else o.b = o.a; } else { const p = d === 0 ? R.b : R.a; m.select(p, p); } return; }
        case 'typetext': case 'insertafter': case 'insertbefore': { const t = toStr(a[0] != null ? a[0] : (a.named && a.named.text)); if (name === 'insertbefore') { m.setText(R.a, R.a, t); if (o instanceof WRange) o.b += t.length; } else if (name === 'insertafter') { const len = m.map().text.length; const at = (R.b >= len && /\r$/.test(m.map().text)) ? len - 1 : R.b; /* like Word: nothing goes after the final paragraph mark */ const e = m.setText(at, at, t); if (o instanceof WRange) o.b = e; } else { const e = m.setText(R.a, R.b, t); m.select(e, e); } m.doc.markDirty(); return; }
        case 'typeparagraph': case 'insertparagraphafter': case 'insertparagraph': { const e = m.setText(R.b, R.b, '\n'); if (o instanceof WSel) m.select(e, e); return; }
        case 'typebackspace': { if (R.b > R.a) m.setText(R.a, R.b, ''); else if (R.a > 0) m.setText(R.a - 1, R.a, ''); return; }
        case 'delete': m.setText(R.a, R.b, ''); if (o instanceof WRange) o.b = o.a; return;
        case 'copy': m.select(R.a, R.b); try { document.execCommand('copy'); } catch (e) {} return;
        case 'cut': m.select(R.a, R.b); try { document.execCommand('copy'); } catch (e) {} m.setText(R.a, R.b, ''); return;
        case 'paste': case 'pasteandformat': case 'pastespecial': { m.select(R.a, R.b); const done = m.ui.pasteText ? m.ui.pasteText() : null; return done; }
        case 'information': return 0;
        case 'case': return 0;
        case 'highlightcolorindex': return 0;
        case 'style': { const blk = m.blocksIn(R.a, R.a)[0]; return blk ? (/^H(\d)$/.test(blk.tagName) ? 'Heading ' + blk.tagName[1] : 'Normal') : 'Normal'; }
        case 'tables': { const ts = Array.prototype.slice.call(m.body.querySelectorAll('table')); return new WColl(ts, t => new WTable(m, t)); }
        case 'omaths': case 'fields': case 'inlineshapes': return new WColl([], x => x);
        case 'expand': case 'movestart': case 'moveend': return 0;
      }
      return m.ign;
    }
    if (o instanceof WFont) { const R = rangeOf(o.t); if (['bold', 'italic', 'underline', 'size', 'name', 'color'].indexOf(name) >= 0) return m.fontValue(R.a, R.b, name); return m.ign; }
    if (o instanceof WPara) { const R = rangeOf(o.t); const blk = m.blocksIn(R.a, R.b)[0]; const cs = blk ? getComputedStyle(blk) : {}; switch (name) { case 'alignment': return { left: 0, start: 0, center: 1, right: 2, end: 2, justify: 3 }[cs.textAlign] || 0; case 'leftindent': return (parseFloat(cs.marginLeft) || 0) * 0.75; case 'firstlineindent': return (parseFloat(cs.textIndent) || 0) * 0.75; case 'spacebefore': return (parseFloat(cs.marginTop) || 0) * 0.75; case 'spaceafter': return (parseFloat(cs.marginBottom) || 0) * 0.75; case 'shading': return new WObjShade(m, o.t); } return m.ign; }
    if (o instanceof WObjShade) return m.ign;
    if (o instanceof WFind) { switch (name) { case 'text': return o.text; case 'replacement': return o.repl; case 'execute': { const named = a.named || {}; if (a[0] != null || named.findtext != null) o.text = toStr(a[0] != null ? a[0] : named.findtext); if (a[1] != null || named.matchcase != null) o.matchCase = truthy(a[1] != null ? a[1] : named.matchcase); if (a[2] != null || named.matchwholeword != null) o.whole = truthy(a[2] != null ? a[2] : named.matchwholeword); if (a[3] != null || named.matchwildcards != null) o.wild = truthy(a[3] != null ? a[3] : named.matchwildcards); const rw = a[9] != null ? a[9] : named.replacewith; const mode = a[10] != null ? toNum(a[10]) : (named.replace != null ? toNum(named.replace) : 0); const r = m.execFind(o, o.t, mode, rw == null ? null : toStr(rw)); m.doc.markDirty(); return r; } case 'found': return !!m.found; case 'clearformatting': return; case 'matchcase': return o.matchCase; case 'matchwildcards': return o.wild; case 'forward': return o.forward; case 'wrap': return o.wrap; } return m.ign; }
    if (o && o.__repl) return name === 'text' ? o.text : undefined;
    if (o instanceof WTable) { switch (name) { case 'rows': return { __count: o.el.rows.length, __coll: [] }; case 'columns': return { __count: o.el.rows[0] ? o.el.rows[0].cells.length : 0, __coll: [] }; case 'cell': { const c = o.el.rows[toNum(a[0]) - 1].cells[toNum(a[1]) - 1]; return new WCell(m, c); } case 'range': return new WRange(m, 0, 0); case 'delete': o.el.remove(); return; } return m.ign; }
    if (o instanceof WCell) { if (name === 'range') return o; if (name === 'text') return o.el.textContent; return m.ign; }
    if (o && o.__count != null && name === 'count') return o.__count;
    if (o && o.__vbaErr) return o.__err(name);
    if (o && typeof o === 'object' && 'text' in o && name === 'text') return o.text;
    throw new VbaError('Đối tượng không hỗ trợ "' + name + '" (' + typeName(o) + ')', 0, 438);
  }
  function setMember(o, name, a, val) {
    if (o instanceof Collection || o instanceof Dictionary) { o.vcall(name, Array.from(a), val, true); return; }
    if (o instanceof WIgnore || o instanceof WColl || o instanceof WObjShade && name !== 'backgroundpatterncolor') return;
    if (o instanceof WApp) { if (name === 'screenupdating') m.screenUpdating = truthy(val); return; }
    if (o instanceof WSel || o instanceof WRange) {
      const R = rangeOf(o);
      if (name === 'text' || name === '') { const e = m.setText(R.a, R.b, toStr(val)); if (o instanceof WRange) o.b = e; else m.select(R.a, e); m.doc.markDirty(); return; }
      if (name === 'start') { if (o instanceof WRange) o.a = toNum(val); else m.select(toNum(val), Math.max(toNum(val), R.b)); return; }
      if (name === 'end') { if (o instanceof WRange) o.b = toNum(val); else m.select(R.a, toNum(val)); return; }
      if (name === 'bold') { m.styleRange(R.a, R.b, { fontWeight: truthy(val) ? 'bold' : 'normal' }); return; }
      if (name === 'italic') { m.styleRange(R.a, R.b, { fontStyle: truthy(val) ? 'italic' : 'normal' }); return; }
      if (name === 'highlightcolorindex') { m.styleRange(R.a, R.b, { backgroundColor: HIGHLIGHT[toNum(val)] || 'transparent' }); return; }
      if (name === 'style') { const st = toStr(val).toLowerCase(); const lv = /heading (\d)|đề mục (\d)/.exec(st) || (typeof val === 'number' && val <= -2 && val >= -10 ? [0, String(-val - 1)] : null); m.blocksIn(R.a, R.b).forEach(b => { const tag = lv ? 'h' + (lv[1] || lv[2]) : (/title|tiêu đề/.test(st) ? 'p' : 'p'); if (b.tagName.toLowerCase() === tag && !/title/.test(st)) return; const n = document.createElement(tag); if (/title/.test(st)) n.className = 'wd-title'; while (b.firstChild) n.appendChild(b.firstChild); b.parentNode.replaceChild(n, b); }); return; }
      if (name === 'case') { const t = m.textOf(R.a, R.b); const v = toNum(val); m.setText(R.a, R.b, v === 1 ? t.toUpperCase() : v === 0 ? t.toLowerCase() : t.replace(/(^|\s)(\S)/g, (x, p, c) => p + c.toUpperCase())); return; }
      return;
    }
    if (o instanceof WFont) { const R = rangeOf(o.t); const css = {}; switch (name) { case 'bold': css.fontWeight = truthy(val) ? 'bold' : 'normal'; break; case 'italic': css.fontStyle = truthy(val) ? 'italic' : 'normal'; break; case 'underline': css.textDecoration = toNum(val) ? 'underline' : 'none'; break; case 'strikethrough': css.textDecoration = truthy(val) ? 'line-through' : 'none'; break; case 'size': css.fontSize = toNum(val) + 'pt'; break; case 'name': css.fontFamily = toStr(val); break; case 'color': { const c = hexOf(val); if (c) css.color = c; break; } case 'allcaps': css.textTransform = truthy(val) ? 'uppercase' : 'none'; break; case 'superscript': css.verticalAlign = truthy(val) ? 'super' : 'baseline'; break; case 'subscript': css.verticalAlign = truthy(val) ? 'sub' : 'baseline'; break; default: return; } m.styleRange(R.a, R.b, css); m.doc.markDirty(); return; }
    if (o instanceof WPara) { const R = rangeOf(o.t); const bl = m.blocksIn(R.a, R.b); const apply = (k, v) => bl.forEach(b => { b.style[k] = v; }); switch (name) { case 'alignment': apply('textAlign', ['left', 'center', 'right', 'justify'][toNum(val)] || 'left'); break; case 'leftindent': apply('marginLeft', pt(val) + 'px'); break; case 'rightindent': apply('marginRight', pt(val) + 'px'); break; case 'firstlineindent': apply('textIndent', pt(val) + 'px'); break; case 'spacebefore': apply('marginTop', toNum(val) + 'pt'); break; case 'spaceafter': apply('marginBottom', toNum(val) + 'pt'); break; case 'linespacingrule': apply('lineHeight', ['1.08', '1.5', '2'][toNum(val)] || ''); break; case 'linespacing': apply('lineHeight', (toNum(val) / 12).toFixed(2)); break; } m.doc.markDirty(); return; }
    if (o instanceof WObjShade) { if (name === 'backgroundpatterncolor') { const R = rangeOf(o.t); const c = hexOf(val); m.blocksIn(R.a, R.b).forEach(b => { b.style.backgroundColor = c === '#ffffff' ? '' : c; }); } return; }
    if (o instanceof WFind) { switch (name) { case 'text': o.text = toStr(val); break; case 'matchcase': o.matchCase = truthy(val); break; case 'matchwholeword': o.whole = truthy(val); break; case 'matchwildcards': o.wild = truthy(val); break; case 'forward': o.forward = truthy(val); break; case 'wrap': o.wrap = toNum(val); break; } return; }
    if (o && o.__repl) { if (name === 'text') o.text = toStr(val); return; }
    if (o instanceof WCell && (name === 'text' || name === '')) { o.el.textContent = toStr(val); return; }
    if (o instanceof WDoc) return;
    throw new VbaError('Không thể gán "' + name + '" (' + typeName(o) + ')', 0, 438);
  }
  return {
    value: v => (v instanceof WRange || v instanceof WSel ? (() => { const R = rangeOf(v); return m.textOf(R.a, R.b); })() : v),
    create, typeName, bracket: () => { throw new VbaError('[...] chỉ dùng trong Excel', 0, 438); },
    same: (x, y) => x === y || (x instanceof WRange && y instanceof WRange && x.a === y.a && x.b === y.b), hasDefault: o => o instanceof WRange,
    each: c => { if (c instanceof VbArray) return c.data.slice(); if (c && c.each) return c.each(); if (c && c.__coll) return c.__coll; throw new VbaError('Không thể lặp For Each', 0, 451); },
    async ident(lc, args, ip) {
      if (Object.prototype.hasOwnProperty.call(CONSTS, lc)) return CONSTS[lc];
      if (Object.prototype.hasOwnProperty.call(WD, lc)) return WD[lc];
      if (/^(wd|mso)[a-z0-9]+$/.test(lc)) return 0;          // other Word enums: accepted (layout-only)
      switch (lc) { case 'activedocument': case 'thisdocument': return m.adoc; case 'selection': return m.sel; case 'application': return m.app; case 'documents': return new WColl([1], () => m.adoc); case 'options': case 'wordbasic': case 'system': case 'commandbars': case 'customizationcontext': case 'keybindings': return m.ign;
        case 'centimeterstopoints': return toNum(args[0]) * 72 / 2.54; case 'inchestopoints': return toNum(args[0]) * 72; case 'millimeterstopoints': return toNum(args[0]) * 72 / 25.4; case 'pointstocentimeters': return toNum(args[0]) * 2.54 / 72; case 'linestopoints': return toNum(args[0]) * 12; case 'debug': return { __dbg: true }; }
      const f = B[lc]; if (f && ['range', 'cells', 'rows', 'columns', 'worksheets', 'sheets', 'activesheet', 'activecell', 'activeworkbook', 'thisworkbook', 'workbooks', 'worksheetfunction'].indexOf(lc) < 0) return await f(args || [], ip);
      return NOT_FOUND;
    },
    async setIdent() { return false; },
    async member(o, name, args, ip) { if (o && o.__dbg) { if (name === 'print') m.ui.print(Array.from(args).map(x => toStr(x instanceof WRange ? m.textOf(x.a, x.b) : x)).join(' ')); return; } return member(o, name, args || [], ip); },
    async setMember(o, name, args, val) { return setMember(o, name, args || [], val); }
  };
}
class WObjShade extends WObj { constructor(m, t) { super(); this.m = m; this.t = t; } }
class WCell extends WObj { constructor(m, el) { super(); this.m = m; this.el = el; } }
