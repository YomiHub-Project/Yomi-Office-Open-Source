// Yomi-Office Open Source v0.0.6 — read/write vbaProject.bin ([MS-CFB] compound file + [MS-OVBA] VBA project) so macros round-trip with
// Microsoft Excel .xlsm files. Writes source-only projects (_VBA_PROJECT version 0xFFFF): Excel recompiles the source when it opens the file.

// ---------------------------------------------------------------- code pages (VBA stores source in the project's ANSI code page)
const CP1258 = [8364, 65533, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 65533, 8249, 338, 65533, 65533, 65533, 65533, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 65533, 8250, 339, 65533, 65533, 376, 160, 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 258, 196, 197, 198, 199, 200, 201, 202, 203, 768, 205, 206, 207, 272, 209, 777, 211, 212, 416, 214, 215, 216, 217, 218, 219, 220, 431, 771, 223, 224, 225, 226, 259, 228, 229, 230, 231, 232, 233, 234, 235, 769, 237, 238, 239, 273, 241, 803, 243, 244, 417, 246, 247, 248, 249, 250, 251, 252, 432, 8363, 255];
const CP1252 = [8364, 65533, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 65533, 381, 65533, 65533, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 353, 8250, 339, 65533, 382, 376];
function table(cp) { if (cp === 1258) return CP1258; const t = CP1252.slice(); for (let i = 160; i < 256; i++) t.push(i); return t; }
export function decodeCp(bytes, cp) {
  if (cp === 65001) return new TextDecoder('utf-8').decode(bytes);
  if (cp !== 1252 && cp !== 1258 && typeof TextDecoder !== 'undefined') { try { return new TextDecoder('windows-' + cp).decode(bytes); } catch (e) {} }
  const t = table(cp); let s = ''; for (let i = 0; i < bytes.length; i++) { const b = bytes[i]; s += String.fromCharCode(b < 128 ? b : t[b - 128]); }
  return s.normalize ? s.normalize('NFC') : s;
}
export function encodeCp(str, cp) {
  const t = table(cp); const rev = {}; t.forEach((u, i) => { if (u !== 65533) rev[u] = 128 + i; });
  // cp1258 has precomposed â ê ô ă ơ ư đ + combining tone marks: decompose everything, then recompose those base letters
  let s = str;
  if (cp === 1258 && s.normalize) s = s.normalize('NFD').replace(/([A-Za-z])([\u0300-\u036f]+)/g, (m0, base, marks) => {
    let b = base, rest = '';
    for (const mk of marks) {
      if (mk === '\u0302' && /[AaEeOo]/.test(b)) b = ({ A: 'Â', a: 'â', E: 'Ê', e: 'ê', O: 'Ô', o: 'ô' })[b];
      else if (mk === '\u0306' && /[Aa]/.test(b)) b = b === 'A' ? 'Ă' : 'ă';
      else if (mk === '\u031B' && /[OoUu]/.test(b)) b = ({ O: 'Ơ', o: 'ơ', U: 'Ư', u: 'ư' })[b];
      else rest += mk;
    }
    return b + rest;
  });
  else if (s.normalize) s = s.normalize('NFC');
  const out = [];
  for (const ch of s) { const u = ch.codePointAt(0); if (u < 128) out.push(u); else if (rev[u] != null) out.push(rev[u]); else out.push(63); }
  return new Uint8Array(out);
}
function needsVietnamese(s) { return /[\u0102\u0103\u0110\u0111\u01A0\u01A1\u01AF\u01B0\u1EA0-\u1EF9]/.test(s); }

// ---------------------------------------------------------------- [MS-OVBA] 2.4.1 compression
export function decompress(buf, off) {
  off = off || 0; if (buf[off] !== 1) throw new Error('VBA: chữ ký nén không hợp lệ');
  const out = []; let p = off + 1;
  while (p < buf.length - 1) {
    const hdr = buf[p] | (buf[p + 1] << 8); const size = (hdr & 0x0FFF) + 3; const compressed = (hdr & 0x8000) !== 0; const end = Math.min(buf.length, p + size); p += 2;
    const start = out.length;
    if (!compressed) { for (let i = 0; i < 4096 && p < end; i++) out.push(buf[p++]); continue; }
    while (p < end) {
      const flags = buf[p++];
      for (let bit = 0; bit < 8 && p < end; bit++) {
        if (!(flags & (1 << bit))) { out.push(buf[p++]); continue; }
        const tok = buf[p] | (buf[p + 1] << 8); p += 2;
        const dpos = out.length - start; let bc = 4; while ((1 << bc) < dpos) bc++; if (bc > 12) bc = 12;
        const lenMask = 0xFFFF >> bc; const offMask = ~lenMask & 0xFFFF;
        const len = (tok & lenMask) + 3; const dist = ((tok & offMask) >> (16 - bc)) + 1;
        const src = out.length - dist; for (let i = 0; i < len; i++) out.push(out[src + i]);
      }
    }
  }
  return new Uint8Array(out);
}
export function compress(data) {
  const out = [1]; let p = 0;
  while (p < data.length) {
    const chunkStart = p; const chunkEnd = Math.min(data.length, p + 4096);
    const body = []; let q = p;
    while (q < chunkEnd) {
      const flagPos = body.length; body.push(0); let flags = 0;
      for (let bit = 0; bit < 8 && q < chunkEnd; bit++) {
        // find longest match in the chunk's already-encoded part
        const dpos = q - chunkStart; let bc = 4; while ((1 << bc) < dpos) bc++; if (bc > 12) bc = 12;
        const lenMask = 0xFFFF >> bc; const maxLen = lenMask + 3;
        let bestLen = 0, bestOff = 0;
        for (let c = q - 1; c >= chunkStart; c--) {
          let l = 0; while (q + l < chunkEnd && data[c + l] === data[q + l] && l < maxLen) l++;
          if (l > bestLen) { bestLen = l; bestOff = q - c; if (l === maxLen) break; }
          if (q - c > 4096) break;
        }
        if (bestLen >= 3) { const tok = ((bestOff - 1) << (16 - bc)) | (bestLen - 3); body.push(tok & 255, tok >> 8); flags |= 1 << bit; q += bestLen; }
        else body.push(data[q++]);
      }
      body[flagPos] = flags;
    }
    if (body.length + 2 > 4098 && chunkEnd - chunkStart === 4096) { const hdr = 0x3000 | 0x0FFF; out.push(hdr & 255, hdr >> 8); for (let i = chunkStart; i < chunkEnd; i++) out.push(data[i]); }
    else { const hdr = 0xB000 | (body.length + 2 - 3); out.push(hdr & 255, hdr >> 8); for (const b of body) out.push(b); }
    p = chunkEnd;
  }
  return new Uint8Array(out);
}

// ---------------------------------------------------------------- [MS-CFB] compound file reader
export function readCfb(buf) {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  if (dv.getUint32(0, true) !== 0xE011CFD0) throw new Error('Không phải tệp OLE/CFB');
  const ss = 1 << dv.getUint16(30, true), mss = 1 << dv.getUint16(32, true);
  const nFat = dv.getUint32(44, true), dirStart = dv.getUint32(48, true), cutoff = dv.getUint32(56, true), miniFatStart = dv.getUint32(60, true), nMiniFat = dv.getUint32(64, true); let difStart = dv.getUint32(68, true); const nDif = dv.getUint32(72, true);
  const sec = n => (n + 1) * ss;
  const fatSecs = []; for (let i = 0; i < 109 && fatSecs.length < nFat; i++) fatSecs.push(dv.getUint32(76 + i * 4, true));
  for (let d = 0; d < nDif && fatSecs.length < nFat; d++) { for (let i = 0; i < ss / 4 - 1 && fatSecs.length < nFat; i++) fatSecs.push(dv.getUint32(sec(difStart) + i * 4, true)); difStart = dv.getUint32(sec(difStart) + ss - 4, true); }
  const fat = []; fatSecs.forEach(s => { for (let i = 0; i < ss / 4; i++) fat.push(dv.getUint32(sec(s) + i * 4, true)); });
  const chain = (start, f) => { const out = []; let s = start; let guard = 0; while (s < 0xFFFFFFFA && guard++ < 1e6) { out.push(s); s = f[s]; } return out; };
  const readBig = (start, size) => { const c = chain(start, fat); const o = new Uint8Array(c.length * ss); c.forEach((s, i) => o.set(buf.subarray(sec(s), sec(s) + ss), i * ss)); return size == null ? o : o.subarray(0, size); };
  const dirBytes = readBig(dirStart);
  const entries = [];
  for (let i = 0; i * 128 < dirBytes.length; i++) {
    const b = dirBytes.subarray(i * 128, i * 128 + 128); const d = new DataView(b.buffer, b.byteOffset, 128);
    const nl = d.getUint16(64, true); let name = ''; for (let k = 0; k < Math.max(0, nl / 2 - 1); k++) name += String.fromCharCode(d.getUint16(k * 2, true));
    entries.push({ name, type: b[66], left: d.getUint32(68, true), right: d.getUint32(72, true), child: d.getUint32(76, true), start: d.getUint32(116, true), size: d.getUint32(120, true) });
  }
  const root = entries[0]; const miniStream = root ? readBig(root.start, root.size) : new Uint8Array(0);
  const miniFat = []; if (nMiniFat) { const mf = readBig(miniFatStart); for (let i = 0; i < mf.length / 4; i++) miniFat.push(new DataView(mf.buffer, mf.byteOffset).getUint32(i * 4, true)); }
  const readStream = e => { if (e.size < cutoff) { const c = chain(e.start, miniFat); const o = new Uint8Array(c.length * mss); c.forEach((s, i) => o.set(miniStream.subarray(s * mss, s * mss + mss), i * mss)); return o.subarray(0, e.size); } return readBig(e.start, e.size); };
  // build path map
  const files = {};
  const walk = (id, path) => { if (id === 0xFFFFFFFF || id >= entries.length) return; const e = entries[id]; walk(e.left, path); const p = path ? path + '/' + e.name : e.name; if (e.type === 2) files[p] = () => readStream(e); if (e.type === 1) walk(e.child, p); walk(e.right, path); };
  walk(root.child, '');
  return files;
}
// ---------------------------------------------------------------- [MS-CFB] writer (v3, 512-byte sectors, mini stream for small streams)
export function writeCfb(tree) { // tree: {name: Uint8Array | {children}} for the root storage
  const SS = 512, MSS = 64, CUT = 4096, FREE = 0xFFFFFFFF, END = 0xFFFFFFFE, FATSECT = 0xFFFFFFFD;
  const entries = [{ name: 'Root Entry', type: 5, kids: [] }];
  const add = (obj, parent) => { Object.keys(obj).forEach(n => { const v = obj[n]; const e = v instanceof Uint8Array ? { name: n, type: 2, data: v } : { name: n, type: 1, kids: [] }; entries.push(e); parent.kids.push(e); if (e.type === 1) add(v, e); }); };
  add(tree, entries[0]);
  entries.forEach((e, i) => { e.id = i; });
  // mini stream
  const mini = []; const miniFat = [];
  entries.filter(e => e.type === 2 && e.data.length < CUT).forEach(e => { const n = Math.max(1, Math.ceil(e.data.length / MSS)); e.start = mini.length / MSS; for (let i = 0; i < n; i++) { miniFat.push(i === n - 1 ? END : e.start + i + 1); } const padded = new Uint8Array(n * MSS); padded.set(e.data); for (const b of padded) mini.push(b); e.size = e.data.length; if (!e.data.length) { e.start = END; } });
  const sectors = []; const fat = [];
  const alloc = (bytes) => { const n = Math.max(1, Math.ceil(bytes.length / SS)); const start = sectors.length; for (let i = 0; i < n; i++) { const s = new Uint8Array(SS); s.set(bytes.subarray(i * SS, i * SS + SS)); sectors.push(s); fat.push(i === n - 1 ? END : start + i + 1); } return start; };
  entries.filter(e => e.type === 2 && e.data.length >= CUT).forEach(e => { e.start = alloc(e.data); e.size = e.data.length; });
  const root = entries[0]; if (mini.length) { root.start = alloc(new Uint8Array(mini)); root.size = mini.length; } else { root.start = END; root.size = 0; }
  let miniFatStart = END, nMiniFat = 0; if (miniFat.length) { const b = new Uint8Array(Math.ceil(miniFat.length * 4 / SS) * SS).fill(255); const dv = new DataView(b.buffer); miniFat.forEach((v, i) => dv.setUint32(i * 4, v, true)); miniFatStart = alloc(b); nMiniFat = b.length / SS; }
  // directory: balanced binary trees per storage (all nodes black — accepted by Office and every reader)
  const cmp = (a, b) => a.name.length !== b.name.length ? a.name.length - b.name.length : (a.name.toUpperCase() < b.name.toUpperCase() ? -1 : a.name.toUpperCase() > b.name.toUpperCase() ? 1 : 0);
  const build = list => { if (!list.length) return FREE; const m = list.length >> 1; const node = list[m]; node.left = build(list.slice(0, m)); node.right = build(list.slice(m + 1)); return node.id; };
  entries.forEach(e => { if (e.kids) e.child = build(e.kids.slice().sort(cmp)); });
  const dirBytes = new Uint8Array(Math.ceil(entries.length / 4) * 512);
  entries.forEach((e, i) => { const o = i * 128; const dv = new DataView(dirBytes.buffer, o, 128); for (let k = 0; k < e.name.length; k++) dv.setUint16(k * 2, e.name.charCodeAt(k), true); dv.setUint16(64, (e.name.length + 1) * 2, true); dirBytes[o + 66] = e.type; dirBytes[o + 67] = 1; dv.setUint32(68, e.left == null ? FREE : e.left, true); dv.setUint32(72, e.right == null ? FREE : e.right, true); dv.setUint32(76, e.child == null ? FREE : e.child, true); dv.setUint32(116, e.start == null ? 0 : e.start, true); dv.setUint32(120, e.size || 0, true); });
  for (let i = entries.length; i < dirBytes.length / 128; i++) { const dv = new DataView(dirBytes.buffer, i * 128, 128); dv.setUint32(68, FREE, true); dv.setUint32(72, FREE, true); dv.setUint32(76, FREE, true); }
  const dirStart = alloc(dirBytes);
  // FAT sectors (they describe themselves too)
  let nFatSec = 1; while ((sectors.length + nFatSec) > nFatSec * (SS / 4)) nFatSec++;
  const fatStart = sectors.length; for (let i = 0; i < nFatSec; i++) { sectors.push(new Uint8Array(SS)); fat.push(FATSECT); }
  if (nFatSec > 109) throw new Error('vbaProject quá lớn');
  const fatAll = new Uint8Array(nFatSec * SS).fill(255); const fdv = new DataView(fatAll.buffer); fat.forEach((v, i) => fdv.setUint32(i * 4, v, true));
  for (let i = 0; i < nFatSec; i++) sectors[fatStart + i] = fatAll.subarray(i * SS, i * SS + SS);
  const hdr = new Uint8Array(512); const h = new DataView(hdr.buffer);
  [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1].forEach((b, i) => { hdr[i] = b; });
  h.setUint16(24, 0x003E, true); h.setUint16(26, 3, true); h.setUint16(28, 0xFFFE, true); h.setUint16(30, 9, true); h.setUint16(32, 6, true);
  h.setUint32(44, nFatSec, true); h.setUint32(48, dirStart, true); h.setUint32(56, CUT, true); h.setUint32(60, miniFatStart, true); h.setUint32(64, nMiniFat, true); h.setUint32(68, END, true); h.setUint32(72, 0, true);
  for (let i = 0; i < 109; i++) h.setUint32(76 + i * 4, i < nFatSec ? fatStart + i : FREE, true);
  const out = new Uint8Array(512 + sectors.length * SS); out.set(hdr, 0); sectors.forEach((s, i) => out.set(s, 512 + i * SS));
  return out;
}

// ---------------------------------------------------------------- read a VBA project → [{name, code, type}]
export function readVbaProject(bin) {
  const files = readCfb(bin);
  const get = p => { const k = Object.keys(files).find(x => x.toLowerCase() === p.toLowerCase()); return k ? files[k]() : null; };
  const dirC = get('VBA/dir'); if (!dirC) throw new Error('vbaProject thiếu luồng dir');
  const dir = decompress(dirC); const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  let p = 0, cp = 1252; const mods = []; let cur = null;
  while (p + 6 <= dir.length) {
    const id = dv.getUint16(p, true); let size = dv.getUint32(p + 2, true); p += 6;
    if (id === 0x0009) { size = 6; }                       // PROJECTVERSION: Reserved(4)=size field, then 6 bytes
    const data = dir.subarray(p, p + size);
    switch (id) {
      case 0x0003: cp = dv.getUint16(p, true); break;
      case 0x0019: cur = { name: decodeCp(data, cp), stream: null, offset: 0, type: 'module' }; mods.push(cur); break;
      case 0x0047: if (cur) cur.uname = String.fromCharCode.apply(null, new Uint16Array(data.slice().buffer)); break;
      case 0x001A: if (cur) cur.stream = decodeCp(data, cp); break;
      case 0x0032: if (cur) cur.ustream = String.fromCharCode.apply(null, new Uint16Array(data.slice().buffer)); break;
      case 0x0031: if (cur) cur.offset = dv.getUint32(p, true); break;
      case 0x0021: if (cur) cur.type = 'module'; break;
      case 0x0022: if (cur) cur.type = 'document'; break;
      case 0x0010: p = dir.length; break;
    }
    p += size;
  }
  const out = [];
  for (const m of mods) {
    const raw = get('VBA/' + (m.ustream || m.stream || m.name)); if (!raw) continue;
    let code = ''; try { code = decodeCp(decompress(raw, m.offset), cp); } catch (e) { continue; }
    const attrs = {}; code.replace(/^Attribute (\w+) = (.*)$/gm, (x, k, v) => { attrs[k] = v; return x; });
    const body = code.replace(/^Attribute VB_.*\r?\n?/gm, '').replace(/\r\n/g, '\n');
    const base = attrs.VB_Base || '';
    out.push({ name: m.uname || m.name, code: unprotectUnicode(body), rawCode: body, type: m.type === 'document' ? (/00020819/.test(base) ? 'workbook' : /00020820/.test(base) ? 'sheet' : /ThisDocument|00020906/i.test(base) ? 'document' : 'class') : 'module' });
  }
  return { modules: out, codepage: cp };
}

// ---------------------------------------------------------------- write a VBA project from modules
function u16(n) { return [n & 255, (n >> 8) & 255]; }
function u32(n) { return [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255]; }
function rec(id, bytes) { return u16(id).concat(u32(bytes.length), Array.from(bytes)); }
function utf16(s) { const o = []; for (let i = 0; i < s.length; i++) o.push(s.charCodeAt(i) & 255, s.charCodeAt(i) >> 8); return o; }
function ascii(s) { return Array.from(s).map(c => c.charCodeAt(0) & 255); }
function hexByte(b) { return (b < 16 ? '0' : '') + b.toString(16).toUpperCase(); }
// [MS-OVBA] 2.4.3.2 data encryption (CMG / DPB / GC values of the PROJECT stream)
function encryptData(projectId, data) {
  const seed = (Math.random() * 256) & 255; const version = 2;
  let projKey = 0; for (let i = 0; i < projectId.length; i++) projKey = (projKey + projectId.charCodeAt(i)) & 255;
  const out = [seed, seed ^ version, seed ^ projKey];
  let unenc1 = projKey, enc1 = out[2], enc2 = out[1];
  const ignored = (seed & 6) >> 1;
  const push = b => { const e = (b ^ ((enc2 + unenc1) & 255)) & 255; out.push(e); enc2 = enc1; enc1 = e; unenc1 = b; };
  for (let i = 0; i < ignored; i++) push(7);
  u32(data.length).forEach(push); data.forEach(push);
  return out.map(hexByte).join('');
}
export const DOC_ATTR = {
  workbook: n => 'Attribute VB_Name = "' + n + '"\r\nAttribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n',
  sheet: n => 'Attribute VB_Name = "' + n + '"\r\nAttribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n',
  module: n => 'Attribute VB_Name = "' + n + '"\r\n',
  // Word ThisDocument (as written by MS Word) and class modules
  document: n => 'Attribute VB_Name = "' + n + '"\r\nAttribute VB_Base = "1Normal.ThisDocument"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = True\r\nAttribute VB_Customizable = True\r\n',
  class: n => 'Attribute VB_Name = "' + n + '"\r\nAttribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = False\r\n'
};
// VBA reads module source with the PC's ANSI code page (not the project code page), so a project written by OOS keeps
// its source pure ASCII: string literals with other characters become "abc" & ChrW(&H1EC7) & "def" (exact at run time on
// any Windows language), comments / identifiers fold Vietnamese to ASCII ("Tiếng Việt" -> "Tieng Viet").
// The exact Unicode source is kept beside vbaProject.bin (oos/vba-source.json) for Yomi-Office Open Source itself.
export function foldAscii(s) { return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\u0111/g, 'd').replace(/\u0110/g, 'D').replace(/[^\x00-\x7f]/g, '?'); }
export function protectUnicode(code) {
  return String(code).split('\n').map(line => {
    if (!/[^\x00-\x7f]/.test(line)) return line;
    if (/^\s*((public|private|global)\s+)?const\s/i.test(line) || /^\s*(#|attribute\s|declare\s)/i.test(line)) return foldAscii(line); // constant expressions cannot call ChrW
    let out = '', i = 0;
    while (i < line.length) {
      const ch = line[i];
      if (ch === "'") { out += foldAscii(line.slice(i)); break; }
      if (/^rem\s/i.test(line.slice(i)) && /^\s*$/.test(out.replace(/.*:/, ''))) { out += foldAscii(line.slice(i)); break; }
      if (ch !== '"') { out += foldAscii(ch); i++; continue; }
      let j = i + 1, lit = '';
      while (j < line.length) { if (line[j] === '"') { if (line[j + 1] === '"') { lit += '"'; j += 2; continue; } break; } lit += line[j++]; }
      if (!/[^\x00-\x7f]/.test(lit)) { out += line.slice(i, j + 1); i = j + 1; continue; }
      const parts = []; let run = '';
      for (let k = 0; k < lit.length; k++) { const c = lit[k]; if (c.charCodeAt(0) < 128) run += c; else { if (run) { parts.push('"' + run.replace(/"/g, '""') + '"'); run = ''; } parts.push('ChrW(&H' + c.charCodeAt(0).toString(16).toUpperCase() + ')'); } }
      if (run) parts.push('"' + run.replace(/"/g, '""') + '"');
      out += '(' + parts.join(' & ') + ')'; i = j + 1;
    }
    return out;
  }).join('\n');
}
// reverse of protectUnicode for literals: ("a" & ChrW(&H1EC7) & "b") -> "aệb"
export function unprotectUnicode(code) {
  const LIT = '"(?:[^"\\n]|"")*"', CW = 'ChrW\\(&H[0-9A-Fa-f]{1,4}\\)';
  const re = new RegExp('\\((' + '(?:' + LIT + '|' + CW + ')(?: & (?:' + LIT + '|' + CW + '))*' + ')\\)', 'g');
  return String(code).replace(re, (m0, inner) => {
    if (inner.indexOf('ChrW(') < 0) return m0;
    let txt = ''; const tre = new RegExp(LIT + '|' + CW, 'g'); let t;
    while ((t = tre.exec(inner))) { if (t[0][0] === '"') txt += t[0].slice(1, -1).replace(/""/g, '"'); else txt += String.fromCharCode(parseInt(t[0].slice(7, -1), 16)); }
    if (/[\x00-\x1f]/.test(txt)) return m0;
    return '"' + txt.replace(/"/g, '""') + '"';
  });
}
export function srcHash(code) { const s = String(code).replace(/\r\n/g, '\n').replace(/^Attribute VB_.*\n/gm, '').replace(/\n+$/, ''); let h = 5381; for (let i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0; return h.toString(16) + ':' + s.length; }
// lossless sidecar: [{name, type, code (Unicode), h: hash of the ASCII text stored in vbaProject.bin}]
export function vbaSidecar(modules) { return JSON.stringify({ v: 1, modules: modules.map(m => ({ name: m.name, type: m.type || 'module', code: m.code, h: srcHash(protectUnicode(String(m.wcode != null ? m.wcode : m.code || '').replace(/\r\n/g, '\n').replace(/^Attribute VB_.*\n/gm, ''))) })) }); }
export function applySidecar(modules, json) {
  let sc; try { sc = JSON.parse(json); } catch (e) { return modules; }
  const map = {}; (sc.modules || []).forEach(x => { map[String(x.name).toLowerCase()] = x; });
  return modules.map(m => { const x = map[String(m.name).toLowerCase()]; if (x && x.h === srcHash(m.rawCode != null ? m.rawCode : m.code)) { const o = { name: m.name, type: m.type, code: x.code }; if (x.type === 'jsa') o.type = 'jsa'; return o; } return m; });
}
// modules: [{name, code, type: 'module'|'workbook'|'sheet'}]; must include ThisWorkbook + one document module per sheet (codeNames)
export function writeVbaProject(modules, opts) {
  const pname = (opts && opts.name) || 'VBAProject';
  const cp = 1252, lcid = 0x0409; // like an English-language Office: an LCID that is not installed makes Excel fail with 80028018
  let dir = [];
  dir = dir.concat(rec(0x0001, u32(1)), rec(0x0002, u32(lcid)), rec(0x0014, u32(lcid)), rec(0x0003, u16(cp)), rec(0x0004, ascii(pname)));
  dir = dir.concat(rec(0x0005, []), rec(0x0040, []), rec(0x0006, []), rec(0x003D, []), rec(0x0007, u32(0)), rec(0x0008, u32(0)));
  dir = dir.concat(u16(0x0009), u32(4), u32(1), u16(0));                       // PROJECTVERSION 1.0
  dir = dir.concat(rec(0x000C, []), rec(0x003C, []));                           // PROJECTCONSTANTS
  // references: stdole (OLE Automation)
  const libid = '*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation';
  dir = dir.concat(rec(0x0016, ascii('stdole')), rec(0x003E, utf16('stdole')));
  dir = dir.concat(u16(0x000D), u32(4 + libid.length + 6), u32(libid.length), ascii(libid), u32(0), u16(0));
  // v0.0.7: Microsoft Office Object Library (MSO.DLL) — defines the mso* constants (msoTextOrientationHorizontal, msoTrue,
  // msoShapeRectangle…). Without it PowerPoint/Word/Excel treat those names as empty variables (= 0) → "value out of range".
  const libmso = '*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.8#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL#Microsoft Office 16.0 Object Library';
  dir = dir.concat(rec(0x0016, ascii('Office')), rec(0x003E, utf16('Office')));
  dir = dir.concat(u16(0x000D), u32(4 + libmso.length + 6), u32(libmso.length), ascii(libmso), u32(0), u16(0));
  dir = dir.concat(u16(0x000F), u32(2), u16(modules.length), u16(0x0013), u32(2), u16(0xFFFF));
  const streams = {};
  modules.forEach(m => {
    dir = dir.concat(rec(0x0019, Array.from(encodeCp(m.name, cp))), rec(0x0047, utf16(m.name)), rec(0x001A, Array.from(encodeCp(m.name, cp))), rec(0x0032, utf16(m.name)));
    dir = dir.concat(rec(0x001C, []), rec(0x0048, []), rec(0x0031, u32(0)), rec(0x001E, u32(0)), rec(0x002C, u16(0xFFFF)));
    dir = dir.concat(u16(m.type === 'module' ? 0x0021 : 0x0022), u32(0));
    dir = dir.concat(u16(0x002B), u32(0));
    const src = ((DOC_ATTR[m.type] || DOC_ATTR.sheet)(m.name) + protectUnicode(String(m.code || '').replace(/\r\n/g, '\n').replace(/^Attribute VB_.*\n/gm, '')).replace(/\n/g, '\r\n')).replace(/(\r\n)*$/, '\r\n');
    streams[m.name] = compress(encodeCp(src, cp));
  });
  dir = dir.concat(u16(0x0010), u32(0));
  streams._VBA_PROJECT = new Uint8Array([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]);
  streams.dir = compress(new Uint8Array(dir));
  // PROJECT stream
  const guid = '{' + [8, 4, 4, 4, 12].map(n => Array.from({ length: n }, () => '0123456789ABCDEF'[(Math.random() * 16) | 0]).join('')).join('-') + '}';
  let txt = 'ID="' + guid + '"\r\n';
  modules.forEach(m => { txt += m.type === 'module' ? 'Module=' + m.name + '\r\n' : m.type === 'class' ? 'Class=' + m.name + '\r\n' : 'Document=' + m.name + '/&H00000000\r\n'; });
  txt += 'Name="' + pname + '"\r\nHelpContextID="0"\r\nVersionCompatible32="393222000"\r\n';
  txt += 'CMG="' + encryptData(guid, [0, 0, 0, 0]) + '"\r\nDPB="' + encryptData(guid, [0]) + '"\r\nGC="' + encryptData(guid, [0xFF]) + '"\r\n\r\n';
  txt += '[Host Extender Info]\r\n&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n\r\n[Workspace]\r\n';
  modules.forEach(m => { txt += m.name + '=0, 0, 0, 0, C\r\n'; });
  // PROJECTwm: module names MBCS + UTF-16
  let wm = []; modules.forEach(m => { wm = wm.concat(Array.from(encodeCp(m.name, cp)), [0], utf16(m.name), [0, 0]); }); wm.push(0, 0);
  const vba = {}; Object.keys(streams).forEach(k => { vba[k] = streams[k]; });
  return writeCfb({ PROJECT: encodeCp(txt, cp), PROJECTwm: new Uint8Array(wm), VBA: vba });
}
