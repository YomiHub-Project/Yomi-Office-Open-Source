// Yomi-Office Open Source v0.0.6 — extra "Save As" formats seen in the Office "Save as type" lists (Office Screen sheet):
// Word XML / PowerPoint XML (Flat OPC), Single File Web Page (.mht), Excel XML Spreadsheet 2003, SYLK (.slk), DIF, Unicode Text,
// templates & shows (.dotx .xltx .potx .ppsx) and RTF outline for PowerPoint.
import { JSZipLib, esc, b64FromBytes, bytesFromB64, utf8Encode } from '../core/util.js';
import { encodeCp } from '../excel/vba/ovba.js';
// SYLK / DIF have no encoding header: Excel reads them in the Windows ANSI code page (1258 = Vietnamese)
const ansi = t => encodeCp(t, /[\u0102\u0103\u0110\u0111\u01A0\u01A1\u01AF\u01B0\u1EA0-\u1EF9\u0300-\u0323]/.test(t) ? 1258 : 1252);

const xesc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');

// ---------------------------------------------------------------- OPC package helpers
export async function retypePackage(bytes, from, to) {   // e.g. docx → dotx main part content type
  const JSZip = await JSZipLib(); const z = await JSZip.loadAsync(bytes); const ct = await z.file('[Content_Types].xml').async('string');
  z.file('[Content_Types].xml', ct.replace(from, to)); return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
export const CT = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml', potx: 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml', ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml'
};
// Flat OPC (single XML file): Word 2007 XML Document / PowerPoint XML Presentation
export async function zipToFlatOpc(bytes, progId) {
  const JSZip = await JSZipLib(); const z = await JSZip.loadAsync(bytes);
  const ctXml = await z.file('[Content_Types].xml').async('string');
  const defaults = {}, overrides = {}; ctXml.replace(/<Default\s+Extension="([^"]+)"\s+ContentType="([^"]+)"/g, (m, e, c) => { defaults[e.toLowerCase()] = c; return m; }); ctXml.replace(/<Override\s+PartName="([^"]+)"\s+ContentType="([^"]+)"/g, (m, p, c) => { overrides[p] = c; return m; });
  let out = '<?xml version="1.0" standalone="yes"?>\n<?mso-application progid="' + progId + '"?>\n<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">';
  for (const name of Object.keys(z.files)) {
    const f = z.files[name]; if (f.dir || name === '[Content_Types].xml') continue;
    const part = '/' + name; const ext = (/\.([^.]+)$/.exec(name) || [])[1] || ''; const ct = overrides[part] || defaults[ext.toLowerCase()] || 'application/octet-stream';
    if (/xml$/.test(ct) || ext === 'rels' || ext === 'xml') { const t = (await f.async('string')).replace(/^<\?xml[^>]*\?>\s*/, ''); out += '<pkg:part pkg:name="' + part + '" pkg:contentType="' + ct + '"' + (ext === 'rels' ? ' pkg:padding="512"' : '') + '><pkg:xmlData>' + t + '</pkg:xmlData></pkg:part>'; }
    else out += '<pkg:part pkg:name="' + part + '" pkg:contentType="' + ct + '" pkg:compression="store"><pkg:binaryData>' + b64FromBytes(await f.async('uint8array')).replace(/(.{76})/g, '$1\n') + '</pkg:binaryData></pkg:part>';
  }
  return utf8Encode(out + '</pkg:package>');
}
export async function flatOpcToZip(text) {
  const JSZip = await JSZipLib(); const z = new JSZip(); const d = new DOMParser().parseFromString(text, 'application/xml'); const NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
  const types = []; const parts = d.getElementsByTagNameNS(NS, 'part');
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i]; const name = p.getAttributeNS(NS, 'name') || p.getAttribute('pkg:name'); const ct = p.getAttributeNS(NS, 'contentType') || p.getAttribute('pkg:contentType');
    const xd = p.getElementsByTagNameNS(NS, 'xmlData')[0], bd = p.getElementsByTagNameNS(NS, 'binaryData')[0];
    if (xd) { const el = xd.firstElementChild; z.file(name.slice(1), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(el)); }
    else if (bd) z.file(name.slice(1), bytesFromB64(bd.textContent.replace(/\s+/g, '')));
    if (!/\.rels$/.test(name)) types.push('<Override PartName="' + name + '" ContentType="' + ct + '"/>');
  }
  z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + types.join('') + '</Types>');
  return z.generateAsync({ type: 'uint8array' });
}

// ---------------------------------------------------------------- MHTML (Single File Web Page)
export function writeMht(html, title) {
  const B = '----=_NextPart_OOS_' + Date.now().toString(16); const parts = []; let n = 0;
  const body = html.replace(/src="data:([^;"]+);base64,([^"]+)"/g, (m, mime, data) => { const loc = 'file:///C:/oos_files/image' + (++n) + '.' + (mime.split('/')[1] || 'png').replace('jpeg', 'jpg').replace('svg+xml', 'svg'); parts.push('--' + B + '\r\nContent-Location: ' + loc + '\r\nContent-Transfer-Encoding: base64\r\nContent-Type: ' + mime + '\r\n\r\n' + data.replace(/(.{76})/g, '$1\r\n') + '\r\n'); return 'src="' + loc + '"'; });
  const qp = s => utf8Encode(s).reduce((acc, b) => acc + (b === 61 || b > 126 ? '=' + (b < 16 ? '0' : '') + b.toString(16).toUpperCase() : String.fromCharCode(b)), '').replace(/(.{72}[^=]{0,3})/g, '$1=\r\n');
  return 'MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary="' + B + '"; type="text/html"\r\nX-Generator: Yomi-Office Open Source 0.0.6\r\n\r\n--' + B + '\r\nContent-Location: file:///C:/oos.htm\r\nContent-Transfer-Encoding: quoted-printable\r\nContent-Type: text/html; charset="utf-8"\r\n\r\n' + qp('<html><head><meta charset="utf-8"><title>' + esc(title || '') + '</title></head><body>' + body + '</body></html>') + '\r\n\r\n' + parts.join('') + '--' + B + '--\r\n';
}
export function readMht(text) {
  const bm = /boundary="?([^";\r\n]+)"?/i.exec(text); if (!bm) return text;
  const segs = text.split('--' + bm[1]).slice(1); let html = ''; const res = {};
  segs.forEach(sg => { const i = sg.indexOf('\r\n\r\n') >= 0 ? sg.indexOf('\r\n\r\n') : sg.indexOf('\n\n'); if (i < 0) return; const head = sg.slice(0, i), body = sg.slice(i).replace(/^\s+/, ''); const ct = (/Content-Type:\s*([^;\r\n]+)/i.exec(head) || [])[1] || ''; const enc = (/Content-Transfer-Encoding:\s*(\S+)/i.exec(head) || [])[1] || ''; const loc = (/Content-Location:\s*(\S+)/i.exec(head) || [])[1] || '';
    if (/text\/html/i.test(ct) && !html) { let t = body; if (/quoted-printable/i.test(enc)) { const bytes = []; t = t.replace(/=\r?\n/g, ''); for (let k = 0; k < t.length; k++) { if (t[k] === '=' && /[0-9A-F]{2}/i.test(t.substr(k + 1, 2))) { bytes.push(parseInt(t.substr(k + 1, 2), 16)); k += 2; } else bytes.push(t.charCodeAt(k) & 255); } t = new TextDecoder('utf-8').decode(new Uint8Array(bytes)); } else if (/base64/i.test(enc)) t = new TextDecoder('utf-8').decode(bytesFromB64(t.replace(/\s+/g, ''))); html = t; }
    else if (/base64/i.test(enc) && loc) res[loc] = 'data:' + ct + ';base64,' + body.replace(/\s+/g, ''); });
  Object.keys(res).forEach(loc => { html = html.split(loc).join(res[loc]); const short = loc.split('/').pop(); html = html.split('"' + short + '"').join('"' + res[loc] + '"'); });
  return html;
}

// ---------------------------------------------------------------- Excel: XML Spreadsheet 2003, SYLK, DIF, Unicode text
export function writeXml2003(wb, M) {
  const styles = []; const sKey = new Map();
  const sid = st => { if (!st) return null; const k = JSON.stringify(st); if (sKey.has(k)) return sKey.get(k); const id = 's' + (sKey.size + 21); sKey.set(k, id);
    styles.push('<Style ss:ID="' + id + '">' + (st.ha || st.va || st.wrap ? '<Alignment' + (st.ha ? ' ss:Horizontal="' + ({ left: 'Left', center: 'Center', right: 'Right' }[st.ha] || 'General') + '"' : '') + (st.va ? ' ss:Vertical="' + (st.va === 'top' ? 'Top' : st.va === 'middle' ? 'Center' : 'Bottom') + '"' : '') + (st.wrap ? ' ss:WrapText="1"' : '') + '/>' : '') + '<Font ss:FontName="' + xesc(st.font || 'Calibri') + '" ss:Size="' + (st.sz || 11) + '"' + (st.b ? ' ss:Bold="1"' : '') + (st.i ? ' ss:Italic="1"' : '') + (st.u ? ' ss:Underline="Single"' : '') + (st.color ? ' ss:Color="' + st.color + '"' : '') + '/>' + (st.fill ? '<Interior ss:Color="' + st.fill + '" ss:Pattern="Solid"/>' : '') + (st.nf ? '<NumberFormat ss:Format="' + xesc(st.nf) + '"/>' : '') + '</Style>'); return id; };
  const sheets = wb.sheets.map(s => { const u = s.used(); let rows = '';
    for (let r = 0; r <= (u.empty ? -1 : u.r2); r++) { let cells = ''; let idx = 0;
      for (let c = 0; c <= u.c2; c++) { const x = s.get(r, c); const v = s.value(r, c); if ((v == null || v === '') && !(x && x.s)) continue;
        const mg = s.merges.find(m => m.r1 === r && m.c1 === c); const st = sid(x && x.s);
        let a = (c !== idx ? ' ss:Index="' + (c + 1) + '"' : '') + (st ? ' ss:StyleID="' + st + '"' : '') + (mg ? ' ss:MergeAcross="' + (mg.c2 - mg.c1) + '"' + (mg.r2 > mg.r1 ? ' ss:MergeDown="' + (mg.r2 - mg.r1) + '"' : '') : '') + (x && x.f ? ' ss:Formula="=' + xesc(toR1C1(x.f, r, c)) + '"' : '');
        const type = typeof v === 'number' ? 'Number' : typeof v === 'boolean' ? 'Boolean' : M.isErr(v) ? 'Error' : 'String';
        const val = v == null ? '' : type === 'Boolean' ? (v ? '1' : '0') : type === 'Error' ? v.code : String(v);
        cells += '<Cell' + a + '>' + (v == null || v === '' ? '' : '<Data ss:Type="' + type + '">' + xesc(val) + '</Data>') + (x && x.note ? '<Comment><Data>' + xesc(x.note.text || x.note) + '</Data></Comment>' : '') + '</Cell>'; idx = c + 1 + (mg ? mg.c2 - mg.c1 : 0); }
      rows += '<Row ss:Index="' + (r + 1) + '"' + (s.rowH[r] ? ' ss:Height="' + (s.rowH[r] * 0.75) + '"' : '') + '>' + cells + '</Row>'; }
    let cols = ''; for (let c = 0; c <= u.c2; c++) if (s.colW[c]) cols += '<Column ss:Index="' + (c + 1) + '" ss:Width="' + Math.round(s.colW[c] * 0.75) + '"/>';
    return '<Worksheet ss:Name="' + xesc(s.name) + '"><Table>' + cols + rows + '</Table></Worksheet>'; });
  return utf8Encode('<?xml version="1.0" encoding="UTF-8"?>\n<?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40"><DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>Yomi-Office Open Source</Author></DocumentProperties><Styles><Style ss:ID="Default" ss:Name="Normal"><Font ss:FontName="Calibri" ss:Size="11"/></Style>' + styles.join('') + '</Styles>' + sheets.join('') + '</Workbook>');
}
// A1 → R1C1 (relative) for SpreadsheetML 2003 formulas
export function toR1C1(f, r0, c0) {
  return String(f).replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])/g, (m, ac, col, ar, row) => { let c = 0; for (const ch of col) c = c * 26 + ch.charCodeAt(0) - 64; c--; const r = +row - 1; return 'R' + (ar ? (r + 1) : (r - r0 ? '[' + (r - r0) + ']' : '')) + 'C' + (ac ? (c + 1) : (c - c0 ? '[' + (c - c0) + ']' : '')); });
}
export function fromR1C1(f, r0, c0, colName) {
  return String(f).replace(/R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?/g, (m, R, C) => { const rel = x => x == null || /^\[/.test(x); const r = R == null ? r0 : /^\[/.test(R) ? r0 + +R.slice(1, -1) : +R - 1; const c = C == null ? c0 : /^\[/.test(C) ? c0 + +C.slice(1, -1) : +C - 1; return (rel(C) ? '' : '$') + colName(c) + (rel(R) ? '' : '$') + (r + 1); });
}
export function readXml2003(wb, text, M) {
  const d = new DOMParser().parseFromString(text, 'application/xml'); const SS = 'urn:schemas-microsoft-com:office:spreadsheet';
  const at = (el, n) => el.getAttributeNS(SS, n) || el.getAttribute('ss:' + n);
  const styles = {}; Array.prototype.forEach.call(d.getElementsByTagNameNS(SS, 'Style'), st => { const o = {}; const f = st.getElementsByTagNameNS(SS, 'Font')[0]; if (f) { if (at(f, 'Bold') === '1') o.b = true; if (at(f, 'Italic') === '1') o.i = true; if (at(f, 'Color')) o.color = at(f, 'Color'); if (at(f, 'Size') && +at(f, 'Size') !== 11) o.sz = +at(f, 'Size'); if (at(f, 'FontName') && at(f, 'FontName') !== 'Calibri') o.font = at(f, 'FontName'); } const it = st.getElementsByTagNameNS(SS, 'Interior')[0]; if (it && at(it, 'Color')) o.fill = at(it, 'Color'); const nf = st.getElementsByTagNameNS(SS, 'NumberFormat')[0]; if (nf && at(nf, 'Format')) o.nf = at(nf, 'Format'); const al = st.getElementsByTagNameNS(SS, 'Alignment')[0]; if (al) { const hz = at(al, 'Horizontal'); if (hz && hz !== 'General') o.ha = hz.toLowerCase(); if (at(al, 'WrapText') === '1') o.wrap = true; } styles[at(st, 'ID')] = M.internStyle(o); });
  Array.prototype.forEach.call(d.getElementsByTagNameNS(SS, 'Worksheet'), ws => {
    const s = wb.addSheet(at(ws, 'Name')); let r = -1;
    Array.prototype.forEach.call(ws.getElementsByTagNameNS(SS, 'Row'), row => { r = at(row, 'Index') ? +at(row, 'Index') - 1 : r + 1; let c = -1;
      Array.prototype.forEach.call(row.getElementsByTagNameNS(SS, 'Cell'), cell => { c = at(cell, 'Index') ? +at(cell, 'Index') - 1 : c + 1; const data = cell.getElementsByTagNameNS(SS, 'Data')[0]; const patch = {};
        if (data) { const t = at(data, 'Type'); const v = data.textContent; patch.v = t === 'Number' ? +v : t === 'Boolean' ? v === '1' : t === 'DateTime' ? M.jsDateToSerial(new Date(v)) : v; }
        const fm = at(cell, 'Formula'); if (fm) patch.f = fromR1C1(fm.replace(/^=/, ''), r, c, M.colName);
        const st = at(cell, 'StyleID'); if (st && styles[st]) patch.s = styles[st];
        if (patch.v != null || patch.f || patch.s) s.setCell(r, c, patch);
        const ma = +(at(cell, 'MergeAcross') || 0), md = +(at(cell, 'MergeDown') || 0); if (ma || md) s.merges.push({ r1: r, c1: c, r2: r + md, c2: c + ma });
        c += ma; }); });
  });
  return wb;
}
export function writeSylk(s, M) {
  const u = s.used(); const L = ['ID;POfficeOpenSource;N;E', 'B;Y' + (u.r2 + 1) + ';X' + (u.c2 + 1)];
  for (let r = 0; r <= u.r2; r++) for (let c = 0; c <= u.c2; c++) { const x = s.get(r, c); const v = s.value(r, c); if (v == null || v === '') continue; const val = typeof v === 'number' ? String(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : '"' + String(M.isErr(v) ? v.code : v).replace(/;/g, ';;').replace(/\n/g, ' ') + '"'; L.push('C;Y' + (r + 1) + ';X' + (c + 1) + ';K' + val + (x && x.f ? ';E' + toR1C1(x.f, r, c).replace(/;/g, ';;') : '')); }
  L.push('E'); return ansi(L.join('\r\n') + '\r\n');
}
export function writeDif(s, M) {
  const u = s.used(); const q = t => '"' + String(t).replace(/"/g, '""') + '"';
  const L = ['TABLE', '0,1', q(''), 'VECTORS', '0,' + (u.c2 + 1), q(''), 'TUPLES', '0,' + (u.r2 + 1), q(''), 'DATA', '0,0', q('')];
  for (let r = 0; r <= u.r2; r++) { L.push('-1,0', 'BOT'); for (let c = 0; c <= u.c2; c++) { const v = s.value(r, c); if (typeof v === 'number') L.push('0,' + v, 'V'); else if (typeof v === 'boolean') L.push('0,' + (v ? 1 : 0), v ? 'TRUE' : 'FALSE'); else L.push('1,0', q(v == null ? '' : (M.isErr(v) ? v.code : v))); } }
  L.push('-1,0', 'EOD'); return ansi(L.join('\r\n') + '\r\n');
}
export function unicodeText(text) { const out = new Uint8Array(2 + text.length * 2); out[0] = 0xFF; out[1] = 0xFE; for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); out[2 + 2 * i] = c & 255; out[3 + 2 * i] = c >> 8; } return out; }

// ---------------------------------------------------------------- PowerPoint: Outline / RTF
export function outlineRtf(slides, plain) {
  const r = t => String(t).replace(/[\\{}]/g, m => '\\' + m).replace(/[^\x00-\x7f]/g, ch => '\\u' + (ch.charCodeAt(0) > 32767 ? ch.charCodeAt(0) - 65536 : ch.charCodeAt(0)) + '?');
  let out = '{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0 Calibri;}}\\f0 ';
  slides.forEach(s => { const title = s.objects.find(o => o.ph === 'title'); out += '\\pard\\b\\fs40 ' + r(plain(title ? title.html : '')) + '\\b0\\par '; s.objects.filter(o => o.type === 'text' && o.ph !== 'title').forEach(o => plain(o.html).split('\n').filter(Boolean).forEach(l => { out += '\\pard\\li360\\fs28 ' + r(l.replace(/^•\s*/, '')) + '\\par '; })); });
  return utf8Encode(out + '}');
}
