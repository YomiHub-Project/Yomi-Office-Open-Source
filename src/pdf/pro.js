// v0.0.7 — PDF Pro (features from Foxit / PDF24 / iLovePDF / PdfBox-Android listed in the project sheet):
// annotate (highlight, pen, shapes, text, stamps) + signature, fill PDF forms, compress, password protect / remove password
// (Standard Security Handler RC4-128, R3), outline (bookmarks), full-text search, document properties, translate page (OCR + A.I).
import { h, ic, esc, clamp, PDFLib, dataUrlToBytes, fmtSize, utf8Encode } from '../core/util.js';
import { toast, dialog, formDlg, promptDlg, busy } from '../core/ui.js';
import { pageToDataUrl } from './tools.js';
import Host from '../core/host.js';

// ================================================================ crypto (MD5, RC4) for PDF encryption
export function md5(bytes) {
  const K = [], S = [7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21];
  for (let i = 0; i < 64; i++) K[i] = (Math.abs(Math.sin(i + 1)) * 4294967296) >>> 0;
  const n = bytes.length, nb = ((n + 8) >> 6) + 1, M = new Uint32Array(nb * 16);
  for (let i = 0; i < n; i++) M[i >> 2] |= bytes[i] << ((i % 4) * 8);
  M[n >> 2] |= 0x80 << ((n % 4) * 8); M[nb * 16 - 2] = (n * 8) >>> 0; M[nb * 16 - 1] = Math.floor(n / 536870912) >>> 0;
  let a0 = 0x67452301, b0 = 0xefcdab89, c0 = 0x98badcfe, d0 = 0x10325476;
  for (let o = 0; o < M.length; o += 16) {
    let A = a0, B = b0, C = c0, D = d0;
    for (let i = 0; i < 64; i++) {
      let F, g;
      if (i < 16) { F = (B & C) | (~B & D); g = i; } else if (i < 32) { F = (D & B) | (~D & C); g = (5 * i + 1) % 16; } else if (i < 48) { F = B ^ C ^ D; g = (3 * i + 5) % 16; } else { F = C ^ (B | ~D); g = (7 * i) % 16; }
      const tmp = D; D = C; C = B; const x = (A + F + K[i] + M[o + g]) >>> 0; B = (B + ((x << S[i]) | (x >>> (32 - S[i])))) >>> 0; A = tmp;
    }
    a0 = (a0 + A) >>> 0; b0 = (b0 + B) >>> 0; c0 = (c0 + C) >>> 0; d0 = (d0 + D) >>> 0;
  }
  const out = new Uint8Array(16); [a0, b0, c0, d0].forEach((v, i) => { for (let j = 0; j < 4; j++) out[i * 4 + j] = (v >>> (8 * j)) & 255; }); return out;
}
export function rc4(key, data) {
  const s = new Uint8Array(256); for (let i = 0; i < 256; i++) s[i] = i;
  for (let i = 0, j = 0; i < 256; i++) { j = (j + s[i] + key[i % key.length]) & 255; const t = s[i]; s[i] = s[j]; s[j] = t; }
  const out = new Uint8Array(data.length);
  for (let k = 0, i = 0, j = 0; k < data.length; k++) { i = (i + 1) & 255; j = (j + s[i]) & 255; const t = s[i]; s[i] = s[j]; s[j] = t; out[k] = data[k] ^ s[(s[i] + s[j]) & 255]; }
  return out;
}
const PAD = new Uint8Array([0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41, 0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08, 0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80, 0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A]);
function cat() { const arrs = Array.prototype.slice.call(arguments); const n = arrs.reduce((s, a) => s + a.length, 0); const o = new Uint8Array(n); let p = 0; arrs.forEach(a => { o.set(a, p); p += a.length; }); return o; }
function padPw(pw) { const b = latin1(pw || '').slice(0, 32); return cat(b, PAD.slice(0, 32 - b.length)); }
function latin1(s) { const u = utf8Encode(s); return u; } // PDF passwords: UTF-8 bytes (ASCII passwords are identical in every scheme)
function le32(v) { return new Uint8Array([v & 255, (v >> 8) & 255, (v >> 16) & 255, (v >>> 24) & 255]); }
function computeO(owner, user, R, len) { let k = md5(padPw(owner || user)); if (R >= 3) for (let i = 0; i < 50; i++) k = md5(k.slice(0, len)); k = k.slice(0, len); let o = rc4(k, padPw(user)); if (R >= 3) for (let i = 1; i <= 19; i++) o = rc4(k.map(b => b ^ i), o); return o; }
function fileKey(user, O, P, id, R, len) { let k = md5(cat(padPw(user), O, le32(P), id)); if (R >= 3) for (let i = 0; i < 50; i++) k = md5(k.slice(0, len)); return k.slice(0, len); }
function computeU(key, id, R) { if (R === 2) return rc4(key, PAD); let u = rc4(key, md5(cat(PAD, id))); for (let i = 1; i <= 19; i++) u = rc4(key.map(b => b ^ i), u); return cat(u, new Uint8Array(16)); }
function objKey(key, num, gen) { const k = md5(cat(key, new Uint8Array([num & 255, (num >> 8) & 255, (num >> 16) & 255, gen & 255, (gen >> 8) & 255]))); return k.slice(0, Math.min(key.length + 5, 16)); }
function hex(b) { let s = ''; for (let i = 0; i < b.length; i++) s += (b[i] < 16 ? '0' : '') + b[i].toString(16); return s; }

// walk every string / stream of the document and transform it with the per-object RC4 key
function cryptAll(PL, pdf, key, skipRef) {
  const ctx = pdf.context;
  const fixVal = (v, k) => {
    if (v instanceof PL.PDFString || v instanceof PL.PDFHexString) return PL.PDFHexString.of(hex(rc4(k, v.asBytes())));
    if (v instanceof PL.PDFArray) { for (let i = 0; i < v.size(); i++) { const x = v.get(i); const y = fixVal(x, k); if (y !== x) v.set(i, y); } return v; }
    if (v instanceof PL.PDFDict) { v.entries().forEach(([n, x]) => { const y = fixVal(x, k); if (y !== x) v.set(n, y); }); return v; }
    return v;
  };
  ctx.enumerateIndirectObjects().forEach(([ref, obj]) => {
    if (skipRef && ref === skipRef) return;
    const k = objKey(key, ref.objectNumber, ref.generationNumber);
    if (obj instanceof PL.PDFStream) {
      const type = obj.dict.get(PL.PDFName.of('Type'));
      if (type && (type.toString() === '/XRef')) return;
      fixVal(obj.dict, k);
      const raw = obj instanceof PL.PDFRawStream ? obj.contents : obj.getContents();
      ctx.assign(ref, PL.PDFRawStream.of(obj.dict, rc4(k, raw)));
    } else fixVal(obj, k);
  });
}
export async function encryptPdf(bytes, userPw, ownerPw, perms) {
  const PL = await PDFLib(); const pdf = await PL.PDFDocument.load(bytes, { updateMetadata: false });
  const R = 3, len = 16; const P = perms != null ? perms : -3904; // print + copy allowed by default (-3904 = print, modify off… see dialog)
  const id = md5(utf8Encode(Date.now() + ':' + Math.random() + ':' + bytes.length));
  const O = computeO(ownerPw || userPw || id.join(''), userPw, R, len), key = fileKey(userPw, O, P, id, R, len), U = computeU(key, id, R);
  const enc = pdf.context.obj({ Filter: 'Standard', V: 2, R, Length: 128, P, O: PL.PDFHexString.of(hex(O)), U: PL.PDFHexString.of(hex(U)) });
  const encRef = pdf.context.register(enc);
  cryptAll(PL, pdf, key, encRef);
  pdf.context.trailerInfo.Encrypt = encRef;
  pdf.context.trailerInfo.ID = pdf.context.obj([PL.PDFHexString.of(hex(id)), PL.PDFHexString.of(hex(id))]);
  return pdf.save({ useObjectStreams: false, updateFieldAppearances: false });
}
// remove a password (RC4 40/128-bit, R2/R3); user or owner password accepted
export async function decryptPdf(bytes, pw) {
  const PL = await PDFLib(); const pdf = await PL.PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
  const encRef = pdf.context.trailerInfo.Encrypt; if (!encRef) return { bytes, wasEncrypted: false };
  const E = pdf.context.lookup(encRef); const g = n => E.get(PL.PDFName.of(n));
  const V = g('V') ? g('V').asNumber() : 0, R = g('R').asNumber(); if (V > 2 || R > 3) throw new Error('Tệp dùng mã hóa AES (V' + V + '/R' + R + ') — hiện chỉ gỡ được mật khẩu RC4; hãy dùng In → Lưu PDF sau khi mở bằng mật khẩu.');
  const len = (g('Length') ? g('Length').asNumber() : 40) / 8; const O = g('O').asBytes(), Uv = g('U').asBytes(), P = g('P').asNumber() | 0;
  const idArr = pdf.context.trailerInfo.ID; const id = idArr ? pdf.context.lookup(idArr).get(0).asBytes() : new Uint8Array(0);
  const check = upw => { const k = fileKey(upw, O, P, id, R, len); const u = computeU(k, id, R); const n = R === 2 ? 32 : 16; for (let i = 0; i < n; i++) if (u[i] !== Uv[i]) return null; return k; };
  let key = check(pw);
  if (!key) { // owner password → recover user password
    let k = md5(padPw(pw)); if (R >= 3) for (let i = 0; i < 50; i++) k = md5(k.slice(0, len)); k = k.slice(0, len);
    let u = O; if (R === 2) u = rc4(k, u); else for (let i = 19; i >= 0; i--) u = rc4(k.map(b => b ^ i), u);
    const upw = String.fromCharCode.apply(null, Array.prototype.slice.call(u)); const cut = upw.indexOf(String.fromCharCode(0x28, 0xBF)); key = check(cut >= 0 ? upw.slice(0, cut) : upw);
    if (!key) { const k2 = fileKey('', O, P, id, R, len); if (!pw && check('')) key = k2; }
  }
  if (!key) throw new Error('Sai mật khẩu');
  cryptAll(PL, pdf, key, encRef);
  delete pdf.context.trailerInfo.Encrypt; pdf.context.delete(encRef);
  return { bytes: await pdf.save({ useObjectStreams: false, updateFieldAppearances: false }), wasEncrypted: true };
}
export async function isEncrypted(bytes) { const s = new TextDecoder('latin1').decode(bytes.subarray(Math.max(0, bytes.length - 4096))); return /\/Encrypt\s+\d+\s+\d+\s+R/.test(s) || /\/Encrypt\s/.test(new TextDecoder('latin1').decode(bytes.subarray(0, Math.min(bytes.length, 200000)))); }

export async function protectDialog(doc) {
  if (!doc.bytes) return toast('Mở một PDF trước', 'err');
  const v = await formDlg('Đặt mật khẩu cho PDF', [
    { key: 'user', type: 'password', label: 'Mật khẩu mở tệp', value: '' }, { key: 'user2', type: 'password', label: 'Nhập lại mật khẩu', value: '' },
    { key: 'owner', type: 'password', label: 'Mật khẩu chủ sở hữu (quyền chỉnh sửa, để trống = giống trên)', value: '' },
    { key: 'print', type: 'checkbox', label: 'Cho phép in', value: true }, { key: 'copy', type: 'checkbox', label: 'Cho phép sao chép nội dung', value: true }, { key: 'edit', type: 'checkbox', label: 'Cho phép chỉnh sửa / chú thích', value: false }]);
  if (!v) return; if (!v.user) return toast('Chưa nhập mật khẩu', 'err'); if (v.user !== v.user2) return toast('Hai mật khẩu không khớp', 'err');
  let P = 0xFFFFF0C0 | 0; if (v.print) P |= 4 | 2048; if (v.edit) P |= 8 | 32 | 256 | 1024; if (v.copy) P |= 16 | 512;
  const b = busy('Đang mã hóa…');
  try { const out = await encryptPdf(doc.bytes, v.user, v.owner || v.user, P); const r = await Host.saveFile((doc.name || 'tai-lieu') + '-bao-mat.pdf', out, { saveAs: true, mime: 'application/pdf' }); if (r) toast('Đã lưu PDF có mật khẩu (RC4 128-bit)', 'ok'); }
  catch (e) { toast('Lỗi mã hóa: ' + e.message, 'err'); } finally { b.done(); }
}
export async function unlockDialog(doc) {
  if (!doc.bytes) return toast('Mở một PDF trước', 'err');
  const pw = doc._password != null ? doc._password : await promptDlg('Mật khẩu của tệp PDF:', '', 'Gỡ mật khẩu', { type: 'password' }); if (pw == null) return;
  const b = busy('Đang gỡ mật khẩu…');
  try { const r = await decryptPdf(doc.bytes, pw); if (!r.wasEncrypted) toast('Tệp này không có mật khẩu'); else { await doc.commit(r.bytes, 'Đã gỡ mật khẩu — nhớ Lưu tệp'); doc._password = null; } }
  catch (e) { toast(e.message, 'err', 6000); } finally { b.done(); }
}

// ================================================================ compress
export async function compressDialog(doc) {
  if (!doc.bytes) return toast('Mở một PDF trước', 'err');
  const v = await formDlg('Nén PDF (giảm dung lượng)', [{ key: 'level', type: 'radio', value: 'medium', options: [['lossless', 'Không mất chất lượng (tối ưu cấu trúc, giữ chữ chọn được)'], ['medium', 'Cân bằng — ảnh 150 dpi JPEG (khuyên dùng cho bản scan)'], ['strong', 'Mạnh — ảnh 96 dpi JPEG (nhỏ nhất)']] }]);
  if (!v) return; const before = doc.bytes.length; const b = busy('Đang nén…'); let out;
  try {
    const PL = await PDFLib();
    if (v.level === 'lossless') { const p = await PL.PDFDocument.load(doc.bytes, { updateMetadata: false }); out = await p.save({ useObjectStreams: true }); }
    else {
      const dpi = v.level === 'strong' ? 96 : 150, q = v.level === 'strong' ? 0.6 : 0.75; const np = await PL.PDFDocument.create();
      for (let i = 0; i < doc.pdf.numPages; i++) { b.text && b.text('Đang nén trang ' + (i + 1) + '/' + doc.pdf.numPages); const pg = await doc.pdf.getPage(i + 1); const vp = pg.getViewport({ scale: 1 }); const url = await pageToDataUrl(doc, i, dpi / 72, 'image/jpeg', q); const img = await np.embedJpg(dataUrlToBytes(url)); const page = np.addPage([vp.width, vp.height]); page.drawImage(img, { x: 0, y: 0, width: vp.width, height: vp.height }); }
      out = await np.save({ useObjectStreams: true });
    }
  } catch (e) { b.done(); return toast('Lỗi nén: ' + e.message, 'err'); }
  b.done();
  if (out.length >= before) return toast('Tệp đã tối ưu — không giảm thêm được (' + fmtSize(before) + ')');
  await doc.commit(out, 'Đã nén: ' + fmtSize(before) + ' → ' + fmtSize(out.length) + ' (−' + Math.round(100 - 100 * out.length / before) + '%)');
}

// ================================================================ properties
export async function propertiesDialog(doc) {
  if (!doc.bytes) return toast('Mở một PDF trước', 'err');
  const PL = await PDFLib(); const p = await PL.PDFDocument.load(doc.bytes, { updateMetadata: false });
  const v = await formDlg('Thuộc tính tài liệu PDF', [{ key: 'title', label: 'Tiêu đề', value: p.getTitle() || '' }, { key: 'author', label: 'Tác giả', value: p.getAuthor() || '' }, { key: 'subject', label: 'Chủ đề', value: p.getSubject() || '' }, { key: 'keywords', label: 'Từ khóa (cách nhau bằng dấu phẩy)', value: p.getKeywords() || '' }]);
  if (!v) return; p.setTitle(v.title); p.setAuthor(v.author); p.setSubject(v.subject); p.setKeywords(v.keywords ? v.keywords.split(/\s*,\s*/) : []); p.setModificationDate(new Date()); p.setProducer('Yomi-Office Open Source 0.0.7');
  await doc.commit(await p.save(), 'Đã cập nhật thuộc tính');
}

// ================================================================ forms
export async function formsDialog(doc) {
  if (!doc.bytes) return toast('Mở một PDF trước', 'err');
  const PL = await PDFLib(); const p = await PL.PDFDocument.load(doc.bytes, { updateMetadata: false }); let form, fields;
  try { form = p.getForm(); fields = form.getFields(); } catch (e) { fields = []; }
  if (!fields.length) return toast('PDF này không có trường biểu mẫu (AcroForm). Dùng "Chú thích" để điền chữ lên trang.');
  const defs = []; const map = {};
  fields.forEach((f, i) => {
    const name = f.getName(), k = 'f' + i; map[k] = f; const C = f.constructor.name;
    if (f instanceof PL.PDFTextField) defs.push({ key: k, label: name, value: f.getText() || '' });
    else if (f instanceof PL.PDFCheckBox) defs.push({ key: k, type: 'checkbox', label: name, value: f.isChecked() });
    else if (f instanceof PL.PDFDropdown || f instanceof PL.PDFOptionList) defs.push({ key: k, type: 'select', label: name, value: (f.getSelected() || [])[0] || '', options: f.getOptions().map(o => [o, o]) });
    else if (f instanceof PL.PDFRadioGroup) defs.push({ key: k, type: 'select', label: name, value: f.getSelected() || '', options: f.getOptions().map(o => [o, o]) });
    else delete map[k];
  });
  defs.push({ key: '_flat', type: 'checkbox', label: 'Làm phẳng biểu mẫu sau khi điền (không sửa được nữa)', value: false });
  const v = await formDlg('Điền biểu mẫu PDF (' + Object.keys(map).length + ' trường)', defs, { cols: defs.length > 8 });
  if (!v) return;
  let needImg = false;
  Object.keys(map).forEach(k => { const f = map[k]; try {
    if (f instanceof PL.PDFTextField) { f.setText(v[k] || ''); if (/[^\x00-\xff]/.test(v[k] || '')) needImg = true; }
    else if (f instanceof PL.PDFCheckBox) v[k] ? f.check() : f.uncheck();
    else if (v[k]) f.select(v[k]);
  } catch (e) { console.warn(e); } });
  let out;
  try {
    if (needImg) { form.acroForm.dict.set(PL.PDFName.of('NeedAppearances'), PL.PDFBool.True); out = await p.save({ updateFieldAppearances: false }); }
    else { if (v._flat) form.flatten(); out = await p.save(); }
  } catch (e) { form.acroForm.dict.set(PL.PDFName.of('NeedAppearances'), PL.PDFBool.True); out = await p.save({ updateFieldAppearances: false }); }
  await doc.commit(out, 'Đã điền biểu mẫu' + (needImg ? ' (chữ có dấu hiển thị do trình xem PDF dựng)' : ''));
}

// ================================================================ outline + search
export async function outlinePanel(doc) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  const ol = await doc.pdf.getOutline();
  if (!ol || !ol.length) return toast('PDF này không có mục lục (bookmark)');
  const go = async dest => { try { const d = typeof dest === 'string' ? await doc.pdf.getDestination(dest) : dest; if (!d) return; const idx = typeof d[0] === 'object' ? await doc.pdf.getPageIndex(d[0]) : d[0]; doc.goto(idx); } catch (e) {} };
  const list = h('div.pdf-outline');
  (function add(items, lvl) { items.forEach(it => { const b = h('button.pdf-ol-i', { style: { paddingLeft: (8 + lvl * 16) + 'px' } }, it.title); b.addEventListener('click', () => go(it.dest)); list.appendChild(b); if (it.items && it.items.length) add(it.items, lvl + 1); }); })(ol, 0);
  dialog({ title: 'Mục lục (Bookmarks)', width: 420, body: list, buttons: [{ label: 'Đóng' }], modeless: true });
}
export async function searchDialog(doc) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  const q = await promptDlg('Tìm trong PDF:', doc._lastQ || '', 'Tìm'); if (!q) return; doc._lastQ = q;
  const low = q.toLowerCase(); const hits = [];
  for (let i = 0; i < doc.pdf.numPages; i++) { const pg = await doc.pdf.getPage(i + 1); const tc = await pg.getTextContent(); const txt = tc.items.map(x => x.str).join(' '); let k = txt.toLowerCase().indexOf(low); while (k >= 0 && hits.length < 500) { hits.push({ page: i, ctx: txt.slice(Math.max(0, k - 40), k + q.length + 40) }); k = txt.toLowerCase().indexOf(low, k + 1); } }
  if (!hits.length) return toast('Không tìm thấy "' + q + '"' + (Object.keys(doc.ocrText || {}).length ? '' : ' (nếu là bản scan, hãy chạy OCR trước)'));
  const list = h('div.pdf-outline', hits.map(x => { const b = h('button.pdf-ol-i', h('b', 'Trang ' + (x.page + 1) + ': '), x.ctx); b.addEventListener('click', () => { doc.goto(x.page); setTimeout(() => markHits(doc, x.page, low), 400); }); return b; }));
  doc.goto(hits[0].page); setTimeout(() => markHits(doc, hits[0].page, low), 400);
  dialog({ title: hits.length + ' kết quả cho "' + q + '"', width: 520, body: list, buttons: [{ label: 'Đóng' }] });
}
function markHits(doc, page, low) { const p = doc.pageInfo[page]; if (!p || !p.textLayer) return; Array.prototype.forEach.call(p.textLayer.children, s => s.classList.toggle('pdf-hit', s.textContent.toLowerCase().indexOf(low) >= 0)); }

// ================================================================ translate page (OCR → A.I)
export async function translatePage(doc) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  const i = doc.cur || 0; const b = busy('Đang đọc chữ trang ' + (i + 1) + '…'); let text = '';
  try { const pg = await doc.pdf.getPage(i + 1); const tc = await pg.getTextContent(); text = tc.items.map(x => x.str).join(' ').trim(); if (text.length < 20) { const r = await Host.ocr(await pageToDataUrl(doc, i, 2), 'vie+eng'); text = r.text; } }
  catch (e) { b.done(); return toast('Không đọc được chữ: ' + e.message, 'err'); }
  b.done(); if (!text.trim()) return toast('Trang không có chữ');
  const A = await import('../ai/ai.js'); A.openAi(doc.shell, doc, 'Dịch đoạn sau sang tiếng Việt (nếu đã là tiếng Việt thì dịch sang tiếng Anh), giữ nguyên bố cục đoạn:\n' + text.slice(0, 6000));
}

// ================================================================ annotate & sign (drawn on the page, burnt into the PDF with pdf-lib)
function textToPng(text, size, color, bold) { const c = document.createElement('canvas'); const x = c.getContext('2d'); const f = (bold ? 'bold ' : '') + size + 'px Arial, "Segoe UI", sans-serif'; x.font = f; const lines = String(text).split(/\n/); const w = Math.ceil(Math.max.apply(null, lines.map(l => x.measureText(l).width))) + 8, hh = Math.ceil(lines.length * size * 1.3) + 6; c.width = w; c.height = hh; x.font = f; x.fillStyle = color; x.textBaseline = 'top'; lines.forEach((l, i) => x.fillText(l, 4, 3 + i * size * 1.3)); return c; }
export function signaturePad() {
  return new Promise(res => {
    const c = h('canvas', { width: 560, height: 200, style: { width: '100%', border: '1px dashed var(--line,#999)', borderRadius: '6px', background: '#fff', touchAction: 'none' } });
    const x = c.getContext('2d'); x.lineWidth = 2.6; x.lineCap = 'round'; x.lineJoin = 'round'; x.strokeStyle = '#0b2a8a'; let down = false, any = false;
    const P = e => { const r = c.getBoundingClientRect(), t = e.touches ? e.touches[0] : e; return [(t.clientX - r.left) * c.width / r.width, (t.clientY - r.top) * c.height / r.height]; };
    const st = e => { e.preventDefault(); down = true; const p = P(e); x.beginPath(); x.moveTo(p[0], p[1]); }, mv = e => { if (!down) return; e.preventDefault(); const p = P(e); x.lineTo(p[0], p[1]); x.stroke(); any = true; }, en = () => { down = false; };
    c.addEventListener('mousedown', st); c.addEventListener('mousemove', mv); window.addEventListener('mouseup', en); c.addEventListener('touchstart', st); c.addEventListener('touchmove', mv); c.addEventListener('touchend', en);
    const color = h('select.inp.small', [['#0b2a8a', 'Xanh'], ['#000000', 'Đen'], ['#c00000', 'Đỏ']].map(o => h('option', { value: o[0] }, o[1]))); color.addEventListener('change', () => { x.strokeStyle = color.value; });
    dialog({ title: 'Vẽ chữ ký', width: 620, body: h('div', c, h('div.form-row.inline', h('label.lbl', 'Màu mực'), color, h('button.btn', { onclick: () => { x.clearRect(0, 0, c.width, c.height); any = false; } }, 'Xoá'))), buttons: [
      { label: 'Dùng chữ ký', primary: true, run: () => { if (!any) { toast('Hãy ký vào khung', 'err'); return false; } res(trim(c)); } }, { label: 'Hủy', run: () => res(null) }] }).then(v => { if (v === null) res(null); });
  });
}
function trim(c) { const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let x0 = c.width, y0 = c.height, x1 = 0, y1 = 0; for (let y = 0; y < c.height; y++) for (let x = 0; x < c.width; x++) if (d[4 * (y * c.width + x) + 3] > 10) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; } if (x1 <= x0) return c; const o = document.createElement('canvas'); o.width = x1 - x0 + 8; o.height = y1 - y0 + 8; o.getContext('2d').drawImage(c, x0 - 4, y0 - 4, o.width, o.height, 0, 0, o.width, o.height); return o; }

export async function annotate(doc, startTool) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  let pageIdx = doc.cur || 0; const items = []; let tool = startTool || 'highlight'; let color = '#ffd400'; let sig = null;
  const cv = h('canvas.pdf-ann-cv'); const over = h('canvas.pdf-ann-over'); const wrap = h('div.pdf-ann-wrap', cv, over);
  let S = 1, base = null, vpW = 0, vpH = 0;
  const load = async () => { const pg = await doc.pdf.getPage(pageIdx + 1); const vp1 = pg.getViewport({ scale: 1 }); vpW = vp1.width; vpH = vp1.height; S = Math.min((Math.min(window.innerWidth, 1100) - 80) / vpW, (window.innerHeight - 230) / vpH); const vp = pg.getViewport({ scale: S }); cv.width = over.width = Math.round(vp.width); cv.height = over.height = Math.round(vp.height); await pg.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise; base = true; redraw(); pgLbl.textContent = 'Trang ' + (pageIdx + 1) + '/' + doc.pdf.numPages; };
  const redraw = (temp) => { const x = over.getContext('2d'); x.clearRect(0, 0, over.width, over.height); items.filter(it => it.page === pageIdx).concat(temp ? [temp] : []).forEach(it => drawItem(x, it)); };
  const drawItem = (x, it) => {
    x.save();
    if (it.type === 'highlight') { x.globalAlpha = 0.38; x.fillStyle = it.color; x.fillRect(it.x, it.y, it.w, it.h); }
    else if (it.type === 'rect') { x.strokeStyle = it.color; x.lineWidth = 2; x.strokeRect(it.x, it.y, it.w, it.h); }
    else if (it.type === 'ellipse') { x.strokeStyle = it.color; x.lineWidth = 2; x.beginPath(); x.ellipse(it.x + it.w / 2, it.y + it.h / 2, Math.abs(it.w / 2), Math.abs(it.h / 2), 0, 0, 2 * Math.PI); x.stroke(); }
    else if (it.type === 'line' || it.type === 'arrow') { x.strokeStyle = it.color; x.lineWidth = 2.5; x.beginPath(); x.moveTo(it.x, it.y); x.lineTo(it.x + it.w, it.y + it.h); x.stroke(); if (it.type === 'arrow') { const a = Math.atan2(it.h, it.w), L = 12; x.beginPath(); x.moveTo(it.x + it.w, it.y + it.h); x.lineTo(it.x + it.w - L * Math.cos(a - 0.45), it.y + it.h - L * Math.sin(a - 0.45)); x.moveTo(it.x + it.w, it.y + it.h); x.lineTo(it.x + it.w - L * Math.cos(a + 0.45), it.y + it.h - L * Math.sin(a + 0.45)); x.stroke(); } }
    else if (it.type === 'pen') { x.strokeStyle = it.color; x.lineWidth = 2.5; x.lineCap = 'round'; x.beginPath(); it.pts.forEach((p, i) => i ? x.lineTo(p[0], p[1]) : x.moveTo(p[0], p[1])); x.stroke(); }
    else if (it.type === 'image') x.drawImage(it.img, it.x, it.y, it.w, it.h);
    else if (it.type === 'whiteout') { x.fillStyle = '#fff'; x.fillRect(it.x, it.y, it.w, it.h); }
    x.restore();
  };
  const pos = e => { const r = over.getBoundingClientRect(), t = e.touches ? e.touches[0] : e; return [(t.clientX - r.left) * over.width / r.width, (t.clientY - r.top) * over.height / r.height]; };
  let drag = null;
  const down = async e => {
    e.preventDefault(); const p = pos(e);
    if (tool === 'text' || tool === 'stamp') {
      const txt = tool === 'stamp' ? (await formDlg('Đóng dấu', [{ key: 't', type: 'select', label: 'Dấu', value: 'ĐÃ DUYỆT', options: ['ĐÃ DUYỆT', 'BẢN SAO', 'MẬT', 'KHẨN', 'ĐÃ THANH TOÁN', 'NHÁP', 'APPROVED', 'CONFIDENTIAL'].map(s => [s, s]) }]) || {}).t : await promptDlg('Nội dung chữ (có thể nhiều dòng):', '', 'Thêm chữ', { multiline: true });
      if (!txt) return; const img = tool === 'stamp' ? stampPng(txt, color === '#ffd400' ? '#c00000' : color) : textToPng(txt, Math.round(16 * S), color === '#ffd400' ? '#000000' : color);
      items.push({ type: 'image', page: pageIdx, img, x: p[0], y: p[1], w: img.width * (tool === 'stamp' ? S * 0.6 : 1), h: img.height * (tool === 'stamp' ? S * 0.6 : 1) }); redraw(); return;
    }
    if (tool === 'sign') { if (!sig) sig = await signaturePad(); if (!sig) return; const w = 180 * S, hh = w * sig.height / sig.width; items.push({ type: 'image', page: pageIdx, img: sig, x: p[0] - w / 2, y: p[1] - hh / 2, w, h: hh }); redraw(); return; }
    if (tool === 'image') { const fs = await Host.openFiles('image/*', false); if (!fs[0]) return; const { loadImage, b64FromBytes } = await import('../core/util.js'); const img = await loadImage('data:image/' + (/png$/i.test(fs[0].name) ? 'png' : 'jpeg') + ';base64,' + b64FromBytes(fs[0].bytes)); const w = Math.min(over.width / 2, img.width), hh = w * img.height / img.width; items.push({ type: 'image', page: pageIdx, img, x: p[0], y: p[1], w, h: hh }); redraw(); return; }
    drag = tool === 'pen' ? { type: 'pen', page: pageIdx, color, pts: [p] } : { type: tool, page: pageIdx, color, x: p[0], y: p[1], w: 0, h: 0 };
  };
  const move = e => { if (!drag) return; e.preventDefault(); const p = pos(e); if (drag.type === 'pen') drag.pts.push(p); else { drag.w = p[0] - drag.x; drag.h = p[1] - drag.y; } redraw(drag); };
  const up = () => { if (!drag) return; const d = drag; drag = null; if (d.type !== 'pen' && d.type !== 'line' && d.type !== 'arrow') { if (d.w < 0) { d.x += d.w; d.w = -d.w; } if (d.h < 0) { d.y += d.h; d.h = -d.h; } if (d.w < 3 && d.h < 3) return redraw(); } items.push(d); redraw(); };
  over.addEventListener('mousedown', down); over.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
  over.addEventListener('touchstart', down, { passive: false }); over.addEventListener('touchmove', move, { passive: false }); over.addEventListener('touchend', up);
  const TOOLS = [['highlight', 'Tô sáng', 'highlighter'], ['pen', 'Bút vẽ', 'pen-line'], ['rect', 'Hình chữ nhật', 'square'], ['ellipse', 'Hình tròn', 'circle'], ['arrow', 'Mũi tên', 'move-up-right'], ['line', 'Đường thẳng', 'minus'], ['text', 'Chữ', 'type'], ['stamp', 'Đóng dấu', 'stamp'], ['sign', 'Chữ ký', 'signature'], ['image', 'Chèn ảnh', 'image'], ['whiteout', 'Xoá trắng', 'eraser']];
  const tb = h('div.pdf-ann-tools'); const btns = {};
  TOOLS.forEach(t => { const b = h('button.btn' + (t[0] === tool ? '.primary' : ''), { title: t[1] }, ic(t[2], 15), ' ' + t[1]); b.addEventListener('click', () => { tool = t[0]; Object.keys(btns).forEach(k => btns[k].classList.toggle('primary', k === tool)); if (tool === 'sign') sig = null; }); btns[t[0]] = b; tb.appendChild(b); });
  const col = h('input', { type: 'color', value: color, title: 'Màu' }); col.addEventListener('input', () => { color = col.value; }); tb.appendChild(col);
  tb.appendChild(h('button.btn', { title: 'Hoàn tác', onclick: () => { for (let i = items.length - 1; i >= 0; i--) if (items[i].page === pageIdx) { items.splice(i, 1); break; } redraw(); } }, ic('undo-2', 15)));
  const pgLbl = h('span.muted');
  const nav = h('div.pdf-ann-nav', h('button.btn', { onclick: () => { if (pageIdx > 0) { pageIdx--; load(); } } }, ic('chevron-left', 15)), pgLbl, h('button.btn', { onclick: () => { if (pageIdx < doc.pdf.numPages - 1) { pageIdx++; load(); } } }, ic('chevron-right', 15)));
  await load();
  const r = await dialog({ title: 'Chú thích & Ký PDF', width: Math.min(window.innerWidth - 20, 1180), body: h('div', tb, nav, wrap), cls: 'pdf-ann-dlg', enterKey: false, buttons: [{ label: 'Áp dụng vào PDF', primary: true, value: 'apply' }, { label: 'Hủy', value: null }] });
  window.removeEventListener('mouseup', up);
  if (r !== 'apply' || !items.length) return;
  const b = busy('Đang ghi chú thích vào PDF…');
  try { await doc.commit(await burn(doc.bytes, items, S), 'Đã thêm ' + items.length + ' chú thích'); } catch (e) { toast('Lỗi: ' + e.message, 'err'); } finally { b.done(); }
}
function stampPng(text, color) { const size = 34; const t = textToPng(text, size, color, true); const c = document.createElement('canvas'); c.width = t.width + 24; c.height = t.height + 16; const x = c.getContext('2d'); x.strokeStyle = color; x.lineWidth = 4; x.strokeRect(3, 3, c.width - 6, c.height - 6); x.drawImage(t, 12, 8); return c; }
function rgbOf(PL, hexc, a) { const n = parseInt(hexc.slice(1), 16); return PL.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); }
export async function burn(bytes, items, S) {
  const PL = await PDFLib(); const p = await PL.PDFDocument.load(bytes, { updateMetadata: false }); const pages = p.getPages();
  for (const it of items) {
    const pg = pages[it.page]; const { width: W, height: H } = pg.getSize(); const rot = pg.getRotation().angle || 0;
    const X = v => v / S, Y = v => H - v / S; // canvas px → PDF pt (unrotated pages; rotated pages fall back to an image overlay below)
    if (rot % 360 !== 0) continue;
    if (it.type === 'highlight') pg.drawRectangle({ x: X(it.x), y: Y(it.y + it.h), width: it.w / S, height: it.h / S, color: rgbOf(PL, it.color), opacity: 0.38, blendMode: PL.BlendMode ? PL.BlendMode.Multiply : undefined });
    else if (it.type === 'whiteout') pg.drawRectangle({ x: X(it.x), y: Y(it.y + it.h), width: it.w / S, height: it.h / S, color: PL.rgb(1, 1, 1) });
    else if (it.type === 'rect') pg.drawRectangle({ x: X(it.x), y: Y(it.y + it.h), width: it.w / S, height: it.h / S, borderColor: rgbOf(PL, it.color), borderWidth: 2 / S * S * 0.75 });
    else if (it.type === 'ellipse') pg.drawEllipse({ x: X(it.x + it.w / 2), y: Y(it.y + it.h / 2), xScale: it.w / 2 / S, yScale: it.h / 2 / S, borderColor: rgbOf(PL, it.color), borderWidth: 1.5 });
    else if (it.type === 'line' || it.type === 'arrow') { pg.drawLine({ start: { x: X(it.x), y: Y(it.y) }, end: { x: X(it.x + it.w), y: Y(it.y + it.h) }, thickness: 1.8, color: rgbOf(PL, it.color) }); if (it.type === 'arrow') { const a = Math.atan2(-it.h, it.w), L = 9, ex = X(it.x + it.w), ey = Y(it.y + it.h); [a - 0.45, a + 0.45].forEach(b => pg.drawLine({ start: { x: ex, y: ey }, end: { x: ex - L * Math.cos(b), y: ey - L * Math.sin(b) }, thickness: 1.8, color: rgbOf(PL, it.color) })); } }
    else if (it.type === 'pen') { const d = it.pts.map((q, i) => (i ? 'L' : 'M') + (q[0] / S).toFixed(2) + ' ' + (q[1] / S).toFixed(2)).join(' '); pg.drawSvgPath(d, { x: 0, y: H, borderColor: rgbOf(PL, it.color), borderWidth: 1.8 }); }
    else if (it.type === 'image') { const png = await p.embedPng(dataUrlToBytes(it.img.toDataURL('image/png'))); pg.drawImage(png, { x: X(it.x), y: Y(it.y + it.h), width: it.w / S, height: it.h / S }); }
  }
  return p.save();
}
