// Yomi-Office Open Source v0.0.6 — Rich Text Format (RTF 1.9) reader & writer.
// v0.0.5 only exported plain text and dropped every non-ASCII character on import (Vietnamese text was lost).
// Reader: groups, \uN + \ucN, \'hh with \ansicpg / font charsets, bold/italic/underline/strike, font size, colours, alignment,
// indents, tables (\trowd … \cell … \row), images (\pict \pngblip / \jpegblip), hyperlinks (HYPERLINK fields).
// Writer: headings, paragraphs (alignment/indent), b/i/u/s, colours, sizes, fonts, lists, tables, images (PNG/JPEG).
import { esc, bytesFromB64, b64FromBytes } from '../core/util.js';

const CP1252 = [8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 141, 381, 143, 144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 353, 8250, 339, 157, 382, 376];
const CP1258 = [8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 138, 8249, 338, 141, 142, 143, 144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 154, 8250, 339, 157, 158, 376, 160, 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 258, 196, 197, 198, 199, 200, 201, 202, 203, 768, 205, 206, 207, 272, 209, 777, 211, 212, 416, 214, 215, 216, 217, 218, 219, 220, 431, 771, 223, 224, 225, 226, 259, 228, 229, 230, 231, 232, 233, 234, 235, 769, 237, 238, 239, 273, 241, 803, 243, 244, 417, 246, 247, 248, 249, 250, 251, 252, 432, 8363, 255];
function byteToChar(b, cp) { if (b < 128) return String.fromCharCode(b); if (cp === 1258) return String.fromCharCode(CP1258[b - 128]); if (cp === 1252 || !cp) return String.fromCharCode(b < 160 ? CP1252[b - 128] : b); try { return new TextDecoder('windows-' + cp).decode(new Uint8Array([b])); } catch (e) { return String.fromCharCode(b); } }
const CHARSET_CP = { 0: 1252, 163: 1258, 238: 1250, 204: 1251, 161: 1253, 162: 1254, 177: 1255, 178: 1256, 186: 1257, 222: 874, 128: 932, 134: 936, 129: 949, 136: 950 };

export function rtfToHtml(rtf) {
  const s = String(rtf); let i = 0; const n = s.length;
  let ansiCp = 1252; const fonts = {}; const colors = [];
  let st = { b: 0, i: 0, u: 0, s: 0, fs: 0, cf: 0, f: -1, uc: 1, skip: false, dest: null, sup: 0, sub: 0 };
  const stack = []; let out = ''; let para = { align: '', li: 0, fi: 0 }; let run = ''; let runSt = null; let pending = '';
  let inTable = false, rowCells = [], tableRows = [], cellHtml = ''; let fieldInst = '', fieldUrl = null; let pict = null;
  const runCss = x => { let o = ''; if (x.fs && x.fs !== 22 && x.fs !== 24) o += 'font-size:' + x.fs / 2 + 'pt;'; if (x.cf && colors[x.cf]) o += 'color:' + colors[x.cf] + ';'; if (x.f >= 0 && fonts[x.f] && fonts[x.f].name && !/^(Calibri|Times New Roman)$/.test(fonts[x.f].name)) o += 'font-family:' + fonts[x.f].name + ';'; return o; };
  const flushRun = () => { if (!run) return; let t = esc(run).replace(/\n/g, '<br>'); const x = runSt || st; if (x.b) t = '<b>' + t + '</b>'; if (x.i) t = '<i>' + t + '</i>'; if (x.u) t = '<u>' + t + '</u>'; if (x.s) t = '<s>' + t + '</s>'; if (x.sup) t = '<sup>' + t + '</sup>'; if (x.sub) t = '<sub>' + t + '</sub>'; const c = runCss(x); if (c) t = '<span style="' + c + '">' + t + '</span>'; if (fieldUrl) t = '<a href="' + esc(fieldUrl) + '">' + t + '</a>'; pending += t; run = ''; };
  const addText = t => { const same = runSt && runSt.b === st.b && runSt.i === st.i && runSt.u === st.u && runSt.s === st.s && runSt.fs === st.fs && runSt.cf === st.cf && runSt.f === st.f && runSt.sup === st.sup && runSt.sub === st.sub; if (!same) { flushRun(); runSt = Object.assign({}, st); } run += t; };
  const endPara = () => { flushRun(); const css = (para.align ? 'text-align:' + para.align + ';' : '') + (para.li ? 'margin-left:' + Math.round(para.li / 15) + 'px;' : '') + (para.fi ? 'text-indent:' + Math.round(para.fi / 15) + 'px;' : ''); const p = '<p' + (css ? ' style="' + css + '"' : '') + '>' + (pending || '<br>') + '</p>'; if (inTable) cellHtml += p; else out += p; pending = ''; };
  const endCell = () => { flushRun(); if (pending) { cellHtml += pending; pending = ''; } rowCells.push(cellHtml.replace(/^<p>(.*)<\/p>$/, '$1')); cellHtml = ''; };
  const endRow = () => { tableRows.push(rowCells); rowCells = []; };
  const closeTable = () => { if (!tableRows.length) { inTable = false; return; } out += '<table class="wd-table">' + tableRows.map(r => '<tr>' + r.map(c => '<td>' + (c || '<br>') + '</td>').join('') + '</tr>').join('') + '</table>'; tableRows = []; inTable = false; };
  let ucSkip = 0;
  while (i < n) {
    const ch = s[i];
    if (ch === '{') { stack.push(Object.assign({}, st)); st.dest = /^(skip|fonttbl|colortbl|pict|fldinst)$/.test(st.dest || '') ? st.dest : null; i++; continue; }
    if (ch === '}') {
      if (st.dest === 'fonttblentry') {}
      if (st.dest === 'pict' && pict) { flushRun(); const hex = pict.hex.replace(/[^0-9a-f]/gi, ''); if (hex.length > 10) { const bytes = new Uint8Array(hex.length / 2); for (let k = 0; k < bytes.length; k++) bytes[k] = parseInt(hex.substr(2 * k, 2), 16); const mime = pict.type === 'jpeg' ? 'image/jpeg' : 'image/png'; const w = pict.w ? Math.round(pict.w / 15) : 0; pending += '<img src="data:' + mime + ';base64,' + b64FromBytes(bytes) + '"' + (w ? ' style="width:' + w + 'px"' : '') + ' alt="">'; } pict = null; }
      if (st.dest === 'fldinst') { const m = /HYPERLINK\s+"([^"]+)"/.exec(fieldInst); if (m) fieldUrl = m[1]; fieldInst = ''; }
      if (st.dest === 'fldrslt') { flushRun(); fieldUrl = null; }
      const prev = stack.pop() || st; if (st.dest !== prev.dest || st.b !== prev.b || st.i !== prev.i || st.u !== prev.u || st.fs !== prev.fs || st.cf !== prev.cf) flushRun(); st = prev; i++; continue;
    }
    if (ch === '\\') {
      const m = /^\\([a-zA-Z]+)(-?\d+)? ?|^\\'([0-9a-fA-F]{2})|^\\(.)/.exec(s.slice(i, i + 40));
      if (!m) { i++; continue; } i += m[0].length;
      if (m[3]) { if (ucSkip > 0) { ucSkip--; continue; } if (st.dest === 'skip' || st.dest === 'colortbl') continue; if (st.dest === 'fonttbl') { if (st.curFont != null) fonts[st.curFont].name += byteToChar(parseInt(m[3], 16), ansiCp); continue; } const cp = st.f >= 0 && fonts[st.f] && fonts[st.f].cp ? fonts[st.f].cp : ansiCp; if (st.dest === 'fldinst') { fieldInst += byteToChar(parseInt(m[3], 16), cp); continue; } addText(byteToChar(parseInt(m[3], 16), cp)); continue; }
      if (m[4]) { const c = m[4]; if (c === '*') { st.dest = 'skip'; continue; } if (st.dest === 'skip') continue; if (c === '\\' || c === '{' || c === '}') addText(c); else if (c === '~') addText('\u00a0'); else if (c === '-' || c === '_') addText(c === '_' ? '-' : ''); else if (c === '\n' || c === '\r') endPara(); continue; }
      const w = m[1], p = m[2] == null ? null : +m[2];
      if (st.dest === 'skip') { if (w === 'bin' && p) i += p; continue; }
      if (st.dest === 'pict') { if (w === 'pngblip') pict.type = 'png'; else if (w === 'jpegblip') pict.type = 'jpeg'; else if (w === 'picwgoal') pict.w = p; else if (/^(emfblip|wmetafile|macpict|dibitmap|wbitmap)$/.test(w)) pict.type = 'skip'; continue; }
      switch (w) {
        case 'ansicpg': ansiCp = p; break;
        case 'fonttbl': st.dest = 'fonttbl'; break;
        case 'colortbl': st.dest = 'colortbl'; st.rgb = [null, null, null]; break;
        case 'red': case 'green': case 'blue': if (st.dest === 'colortbl') st.rgb[{ red: 0, green: 1, blue: 2 }[w]] = p; break;
        case 'stylesheet': case 'info': case 'header': case 'footer': case 'headerl': case 'headerr': case 'footerl': case 'footerr': case 'listtable': case 'listoverridetable': case 'rsidtbl': case 'generator': case 'xmlnstbl': case 'themedata': case 'colorschememapping': case 'latentstyles': case 'datastore': case 'mmathPr': case 'pgdsctbl': case 'footnote': case 'annotation': case 'object': case 'nonshppict': case 'fldrslt_skip': st.dest = 'skip'; break;
        case 'pict': flushRun(); st.dest = 'pict'; pict = { hex: '', type: 'png' }; break;
        case 'field': break; case 'fldinst': st.dest = 'fldinst'; break; case 'fldrslt': st.dest = 'fldrslt'; break;
        case 'f': if (st.dest === 'fonttbl') { st.curFont = p; fonts[p] = fonts[p] || { name: '' }; } else st.f = p; break;
        case 'fcharset': if (st.dest === 'fonttbl' && st.curFont != null) fonts[st.curFont].cp = CHARSET_CP[p] || null; break;
        case 'uc': st.uc = p; break;
        case 'u': { if (st.dest === 'fonttbl' || st.dest === 'colortbl') break; let c = p < 0 ? p + 65536 : p; const t = String.fromCharCode(c); if (st.dest === 'fldinst') fieldInst += t; else addText(t); ucSkip = st.uc; break; }
        case 'par': endPara(); break; case 'line': addText('\n'); break; case 'tab': addText('\t'); break; case 'page': flushRun(); endPara(); out += '<p class="wd-pagebreak"><br></p>'; break;
        case 'pard': para = { align: '', li: 0, fi: 0 }; break;
        case 'qc': para.align = 'center'; break; case 'qr': para.align = 'right'; break; case 'qj': para.align = 'justify'; break; case 'ql': para.align = ''; break;
        case 'li': para.li = p || 0; break; case 'fi': para.fi = p || 0; break;
        case 'plain': flushRun(); Object.assign(st, { b: 0, i: 0, u: 0, s: 0, fs: 0, cf: 0, sup: 0, sub: 0 }); break;
        case 'b': st.b = p === 0 ? 0 : 1; break; case 'i': st.i = p === 0 ? 0 : 1; break; case 'ul': st.u = p === 0 ? 0 : 1; break; case 'ulnone': st.u = 0; break; case 'strike': st.s = p === 0 ? 0 : 1; break;
        case 'super': st.sup = 1; break; case 'sub': st.sub = 1; break; case 'nosupersub': st.sup = st.sub = 0; break;
        case 'fs': st.fs = p; break; case 'cf': st.cf = p; break;
        case 'trowd': if (!inTable) { flushRun(); if (pending) endPara(); inTable = true; } break;
        case 'intbl': break;
        case 'cell': endCell(); break; case 'row': endRow(); break;
        case 'emdash': addText('—'); break; case 'endash': addText('–'); break; case 'bullet': addText('•'); break; case 'lquote': addText('‘'); break; case 'rquote': addText('’'); break; case 'ldblquote': addText('“'); break; case 'rdblquote': addText('”'); break;
      }
      continue;
    }
    if (ch === '\r' || ch === '\n') { i++; continue; }
    if (ucSkip > 0) { ucSkip--; i++; continue; }
    if (st.dest === 'skip') { i++; continue; }
    if (st.dest === 'pict') { if (pict) pict.hex += ch; i++; continue; }
    if (st.dest === 'fonttbl') { if (st.curFont != null && ch !== ';') fonts[st.curFont].name = (fonts[st.curFont].name + ch).replace(/^\s+/, ''); i++; continue; }
    if (st.dest === 'colortbl') { if (ch === ';') { colors.push(st.rgb[0] == null ? null : '#' + st.rgb.map(v => ((v || 0) < 16 ? '0' : '') + (v || 0).toString(16)).join('')); st.rgb = [null, null, null]; } i++; continue; }
    if (st.dest === 'fldinst') { fieldInst += ch; i++; continue; }
    if (inTable && tableRows.length && !rowCells.length && /\S/.test(ch) && stack.length <= 1 && false) closeTable();
    addText(ch); i++;
  }
  flushRun(); if (pending) endPara(); if (inTable) closeTable();
  // close a table when non-table paragraphs follow (\pard without \intbl) — handled by end; ensure leftover rows are rendered
  if (tableRows.length) closeTable();
  return out || '<p><br></p>';
}

// ---------------------------------------------------------------- writer
export function htmlToRtf(root) {
  const colors = ['#000000']; const fonts = ['Calibri'];
  const colorIdx = c => { if (!c) return 0; const m = /rgb\((\d+),\s*(\d+),\s*(\d+)/.exec(c); const hex = m ? '#' + [m[1], m[2], m[3]].map(v => (+v < 16 ? '0' : '') + (+v).toString(16)).join('') : (/^#[0-9a-f]{6}$/i.test(c) ? c.toLowerCase() : null); if (!hex) return 0; let i = colors.indexOf(hex); if (i < 0) { colors.push(hex); i = colors.length - 1; } return i + 1; };
  const fontIdx = f => { if (!f) return 0; f = f.split(',')[0].replace(/['"]/g, '').trim(); let i = fonts.indexOf(f); if (i < 0) { fonts.push(f); i = fonts.length - 1; } return i; };
  const txt = t => String(t).replace(/[\\{}]/g, m => '\\' + m).replace(/\t/g, '\\tab ').replace(/\u00a0/g, '\\~').replace(/[^\x00-\x7f]/g, ch => { const c = ch.charCodeAt(0); return '\\u' + (c > 32767 ? c - 65536 : c) + '?'; });
  const inline = (node) => {
    let o = '';
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) { o += txt(n.nodeValue); continue; }
      if (n.nodeType !== 1) continue; const tg = n.tagName; const stl = n.style || {};
      if (tg === 'BR') { o += '\\line '; continue; }
      if (tg === 'IMG') { const m = /^data:image\/(png|jpeg|jpg);base64,(.*)$/.exec(n.getAttribute('src') || ''); if (m) { const bytes = bytesFromB64(m[2]); let hex = ''; for (let k = 0; k < bytes.length; k++) hex += (bytes[k] < 16 ? '0' : '') + bytes[k].toString(16); const w = Math.round((n.width || n.naturalWidth || 300) * 15), hh = Math.round((n.height || n.naturalHeight || 200) * 15); o += '{\\pict\\' + (m[1] === 'png' ? 'pngblip' : 'jpegblip') + '\\picwgoal' + w + '\\pichgoal' + hh + ' ' + hex.replace(/(.{128})/g, '$1\n') + '}'; } continue; }
      let pre = '', post = '';
      if (tg === 'B' || tg === 'STRONG' || stl.fontWeight === 'bold' || +stl.fontWeight >= 600) { pre += '\\b '; post = '\\b0 ' + post; }
      if (tg === 'I' || tg === 'EM' || stl.fontStyle === 'italic') { pre += '\\i '; post = '\\i0 ' + post; }
      if (tg === 'U' || /underline/.test(stl.textDecoration || '')) { pre += '\\ul '; post = '\\ulnone ' + post; }
      if (tg === 'S' || tg === 'STRIKE' || tg === 'DEL' || /line-through/.test(stl.textDecoration || '')) { pre += '\\strike '; post = '\\strike0 ' + post; }
      if (tg === 'SUP') { pre += '\\super '; post = '\\nosupersub ' + post; } if (tg === 'SUB') { pre += '\\sub '; post = '\\nosupersub ' + post; }
      if (stl.color) { pre += '\\cf' + colorIdx(stl.color) + ' '; }
      if (stl.fontSize && /pt$/.test(stl.fontSize)) pre += '\\fs' + Math.round(parseFloat(stl.fontSize) * 2) + ' ';
      if (stl.fontFamily) pre += '\\f' + fontIdx(stl.fontFamily) + ' ';
      if (tg === 'A') { o += '{\\field{\\*\\fldinst HYPERLINK "' + txt(n.getAttribute('href') || '') + '"}{\\fldrslt{\\ul\\cf' + colorIdx('#0563c1') + ' ' + inline(n) + '}}}'; continue; }
      o += '{' + pre + inline(n) + '}';
    }
    return o;
  };
  const pPr = el => { const st = el.style || {}; let o = '\\pard\\plain\\f0\\fs22\\sa160'; const a = st.textAlign; if (a === 'center') o += '\\qc'; else if (a === 'right') o += '\\qr'; else if (a === 'justify') o += '\\qj'; const ml = parseFloat(st.marginLeft); if (ml) o += '\\li' + Math.round(ml * 15); const ti = parseFloat(st.textIndent); if (ti) o += '\\fi' + Math.round(ti * 15); return o + ' '; };
  const blocks = (node) => {
    let o = '';
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) { if (n.nodeValue.trim()) o += '\\pard\\plain\\f0\\fs22 ' + txt(n.nodeValue) + '\\par\n'; continue; }
      if (n.nodeType !== 1) continue; const tg = n.tagName;
      if (/^H[1-6]$/.test(tg)) { const sz = [0, 32, 26, 24, 22, 22, 22][+tg[1]]; o += pPr(n) + '\\outlinelevel' + (+tg[1] - 1) + '\\b\\fs' + sz + '\\cf' + colorIdx('#2f5496') + ' ' + inline(n) + '\\b0\\par\n'; continue; }
      if (tg === 'UL' || tg === 'OL') { let k = 0; Array.prototype.forEach.call(n.children, li => { k++; o += '\\pard\\plain\\f0\\fs22\\fi-360\\li720\\sa80 ' + (tg === 'OL' ? k + '.' : '\\bullet') + '\\tab ' + inline(li) + '\\par\n'; }); continue; }
      if (tg === 'TABLE') { const rows = Array.prototype.slice.call(n.rows || []); rows.forEach(r => { const cells = Array.prototype.slice.call(r.cells); let x = 0; o += '\\trowd\\trgaph108'; cells.forEach(() => { x += Math.round(9000 / Math.max(1, cells.length)); o += '\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs\\cellx' + x; }); o += '\n'; cells.forEach(c => { o += '\\pard\\intbl\\plain\\f0\\fs22 ' + (c.tagName === 'TH' ? '\\b ' : '') + inline(c) + (c.tagName === 'TH' ? '\\b0' : '') + '\\cell '; }); o += '\\row\n'; }); o += '\\pard\n'; continue; }
      if (n.classList && n.classList.contains('wd-pagebreak')) { o += '\\page\n'; continue; }
      if ((tg === 'DIV' || tg === 'SECTION' || tg === 'BLOCKQUOTE') && n.querySelector('p,h1,h2,h3,h4,h5,h6,ul,ol,table,div')) { o += blocks(n); continue; }
      o += pPr(n) + inline(n) + '\\par\n';
    }
    return o;
  };
  const body = blocks(root);
  const fontTbl = '{\\fonttbl' + fonts.map((f, i) => '{\\f' + i + '\\fnil\\fcharset0 ' + txt(f) + ';}').join('') + '}';
  const colTbl = '{\\colortbl;' + colors.map(c => '\\red' + parseInt(c.slice(1, 3), 16) + '\\green' + parseInt(c.slice(3, 5), 16) + '\\blue' + parseInt(c.slice(5, 7), 16) + ';').join('') + '}';
  return '{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1' + fontTbl + colTbl + '{\\*\\generator Yomi-Office Open Source 0.0.6;}\\viewkind4\n' + body + '}';
}
