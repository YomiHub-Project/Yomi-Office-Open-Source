// PDF tools: convert, OCR, merge/split, rotate, extract, watermark, page numbers — all offline
import { h, ic, esc, uid, loadScript, PDFLib, PDFJS, JSZipLib, bytesToDataUrl, dataUrlToBytes, readFileAsDataUrl, utf8Encode, fmtSize } from '../core/util.js';
import { toast, dialog, confirmDlg, promptDlg, formDlg, busy } from '../core/ui.js';
import Host from '../core/host.js';

let _busy=null;
function busyOn(t){ if(_busy)_busy.done(); _busy=busy(t); }
function busyOff(){ if(_busy){_busy.done(); _busy=null;} }

// ---------------------------------------------------------------- render a page to PNG dataURL (for OCR / flatten / export)
export async function pageToDataUrl(doc, i, scale, mime, q) {
  const p = doc.pageInfo[i]; const page = p._page || await doc.measureOne(i);
  const vp = page.getViewport({ scale: scale || 2, rotation: p.rotate || 0 });
  const c = h('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: ctx, viewport: vp }).promise;
  return c.toDataURL(mime || 'image/png', q);
}

// ---------------------------------------------------------------- images -> PDF
export async function imagesToPdf(files) {
  const { PDFDocument } = await PDFLib();
  const pdf = await PDFDocument.create();
  for (const f of files) {
    const bytes = f.bytes instanceof Uint8Array ? f.bytes : new Uint8Array(f.bytes);
    let img, isPng = /\.png$/i.test(f.name) || (bytes[0] === 0x89 && bytes[1] === 0x50);
    // normalize non jpg/png (bmp/gif/webp) via canvas
    if (!isPng && !/\.jpe?g$/i.test(f.name) && !(bytes[0] === 0xff && bytes[1] === 0xd8)) {
      const url = await rasterize(bytesToDataUrl(bytes, 'image/*')); img = await pdf.embedPng(dataUrlToBytes(url)); isPng = true;
    } else img = isPng || (bytes[0] === 0x89) ? await pdf.embedPng(bytes) : await pdf.embedJpg(bytes);
    const page = pdf.addPage([img.width, img.height]);
    page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
  }
  return pdf.save();
}
function rasterize(url) { return new Promise(res => { const im = new Image(); im.onload = () => { const c = h('canvas'); c.width = im.naturalWidth; c.height = im.naturalHeight; c.getContext('2d').drawImage(im, 0, 0); res(c.toDataURL('image/png')); }; im.onerror = () => res(url); im.src = url; }); }
export async function imagesToPdfDialog(doc) {
  const files = await Host.openFiles('image/*', true);
  if (!files || !files.length) return;
  busyOn('Đang tạo PDF từ ' + files.length + ' ảnh…');
  try { const bytes = await imagesToPdf(files); await doc.setBytes(bytes, true); doc.name = 'Ảnh.pdf'; doc.markDirty(); toast('Đã tạo PDF từ ' + files.length + ' ảnh', 'ok'); }
  catch (e) { toast('Lỗi: ' + e.message, 'err'); } busyOff();
}

// ---------------------------------------------------------------- merge
export async function mergeDialog(doc) {
  const files = await Host.openFiles('application/pdf', true);
  if (!files || !files.length) return;
  busyOn('Đang ghép PDF…');
  try {
    const { PDFDocument } = await PDFLib();
    const out = doc.bytes ? await PDFDocument.load(doc.bytes) : await PDFDocument.create();
    for (const f of files) { const src = await PDFDocument.load(f.bytes); const pages = await out.copyPages(src, src.getPageIndices()); pages.forEach(p => out.addPage(p)); }
    await doc.commit(await out.save(), 'Đã ghép ' + files.length + ' tệp');
  } catch (e) { toast('Lỗi ghép: ' + e.message, 'err'); } busyOff();
}

// ---------------------------------------------------------------- split / extract pages
export async function splitDialog(doc) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  const v = await formDlg('Tách / Trích trang', [{ key: 'range', label: 'Phạm vi trang (vd 1-3,5,8-10)', value: '1-' + doc.pdf.numPages }, { key: 'mode', type: 'radio', value: 'extract', options: [['extract', 'Lưu các trang này thành PDF mới'], ['each', 'Mỗi trang một tệp (ZIP)']] }]);
  if (!v) return;
  const idx = parseRange(v.range, doc.pdf.numPages);
  if (!idx.length) return toast('Phạm vi không hợp lệ', 'err');
  busyOn('Đang tách…');
  try {
    const { PDFDocument } = await PDFLib(); const src = await PDFDocument.load(doc.bytes);
    if (v.mode === 'each') {
      const JSZip = await JSZipLib(); const zip = new JSZip();
      for (const i of idx) { const nd = await PDFDocument.create(); const [pg] = await nd.copyPages(src, [i]); nd.addPage(pg); zip.file('trang-' + (i + 1) + '.pdf', await nd.save()); }
      await saveOut(doc, (doc.name || 'pdf').replace(/\.pdf$/i, '') + '-cac-trang.zip', await zip.generateAsync({ type: 'uint8array' }), 'application/zip');
    } else {
      const nd = await PDFDocument.create(); const pages = await nd.copyPages(src, idx); pages.forEach(p => nd.addPage(p));
      await saveOut(doc, (doc.name || 'pdf').replace(/\.pdf$/i, '') + '-trich.pdf', await nd.save(), 'application/pdf');
    }
    toast('Đã tách ' + idx.length + ' trang', 'ok');
  } catch (e) { toast('Lỗi: ' + e.message, 'err'); } busyOff();
}
export function parseRange(s, max) {
  const out = []; (s || '').split(',').forEach(part => { part = part.trim(); if (!part) return; const m = /^(\d+)\s*-\s*(\d+)$/.exec(part); if (m) { for (let i = +m[1]; i <= +m[2]; i++) if (i >= 1 && i <= max) out.push(i - 1); } else { const n = +part; if (n >= 1 && n <= max) out.push(n - 1); } }); return out.filter((v, i, a) => a.indexOf(v) === i);
}

// ---------------------------------------------------------------- rotate / delete / reorder (operate on current bytes)
export async function rotatePages(doc, indices, deg) {
  const { PDFDocument, degrees } = await PDFLib(); const src = await PDFDocument.load(doc.bytes);
  const pages = src.getPages(); indices = indices.filter(i => i >= 0 && i < pages.length); indices.forEach(i => { const cur = pages[i].getRotation().angle || 0; pages[i].setRotation(degrees((cur + deg) % 360)); });
  await doc.commit(await src.save(), 'Đã xoay ' + indices.length + ' trang');
}
export async function deletePages(doc, indices) {
  if (doc.pdf && indices.length >= doc.pdf.numPages) return toast('Không thể xoá tất cả các trang', 'err');
  const { PDFDocument } = await PDFLib(); const src = await PDFDocument.load(doc.bytes);
  indices = indices.filter((v, k, a) => v >= 0 && v < src.getPageCount() && a.indexOf(v) === k); if (!indices.length || indices.length >= src.getPageCount()) return toast('Không thể xoá các trang đã chọn', 'err');
  indices.slice().sort((a, b) => b - a).forEach(i => src.removePage(i));
  await doc.commit(await src.save(), 'Đã xoá ' + indices.length + ' trang');
}
export async function movePage(doc, from, to) {
  const { PDFDocument } = await PDFLib(); const src = await PDFDocument.load(doc.bytes);
  const order = src.getPageIndices(); const n = order.length; // v0.0.8: guard stale indices (page list changed while a command was pending)
  if (!(from >= 0 && from < n) || !(to >= 0 && to < n) || from === to) return toast('Không thể di chuyển trang này', 'err');
  const [m] = order.splice(from, 1); order.splice(to, 0, m);
  const nd = await PDFDocument.create(); const pages = await nd.copyPages(src, order); pages.forEach(p => nd.addPage(p));
  await doc.commit(await nd.save(), 'Đã sắp xếp trang');
}

// ---------------------------------------------------------------- extract images from a PDF
export async function extractImages(doc) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  busyOn('Đang trích ảnh…');
  try {
    const JSZip = await JSZipLib(); const zip = new JSZip(); let count = 0;
    for (let i = 0; i < doc.pdf.numPages; i++) {
      const page = doc.pageInfo[i]._page || await doc.measureOne(i);
      const ops = await page.getOperatorList(); const pdfjs = await PDFJS();
      const OPS = pdfjs.OPS;
      for (let j = 0; j < ops.fnArray.length; j++) {
        if (ops.fnArray[j] === OPS.paintImageXObject || ops.fnArray[j] === OPS.paintJpegXObject) {
          const name = ops.argsArray[j][0];
          try {
            const img = await new Promise((res) => { page.objs.get(name, res); });
            if (img && img.width) { const c = h('canvas'); c.width = img.width; c.height = img.height; const ctx = c.getContext('2d'); const id = ctx.createImageData(img.width, img.height); putImg(id, img); ctx.putImageData(id, 0, 0); zip.file('p' + (i + 1) + '-img' + (++count) + '.png', dataUrlToBytes(c.toDataURL('image/png'))); }
          } catch (e) {}
        }
      }
    }
    if (!count) { toast('Không tìm thấy ảnh nhúng — thử "Xuất trang thành ảnh"', 'err'); busyOff(); return; }
    await saveOut(doc, (doc.name || 'pdf').replace(/\.pdf$/i, '') + '-anh.zip', await zip.generateAsync({ type: 'uint8array' }), 'application/zip');
    toast('Đã trích ' + count + ' ảnh', 'ok');
  } catch (e) { toast('Lỗi: ' + e.message, 'err'); } busyOff();
}
function putImg(id, img) { const d = id.data, s = img.data; if (img.kind === 1 || (s.length === img.width * img.height)) { for (let k = 0, p = 0; k < s.length; k++) { d[p++] = s[k]; d[p++] = s[k]; d[p++] = s[k]; d[p++] = 255; } } else if (s.length === img.width * img.height * 3) { for (let k = 0, p = 0; k < s.length; p += 4) { d[p] = s[k++]; d[p + 1] = s[k++]; d[p + 2] = s[k++]; d[p + 3] = 255; } } else { d.set(s.subarray(0, d.length)); } }

// ---------------------------------------------------------------- export pages as images (PDF -> PNG/JPG)
export async function pdfToImages(doc, fmt) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  fmt = fmt || 'png';
  busyOn('Đang kết xuất ' + doc.pdf.numPages + ' trang…');
  try {
    const JSZip = await JSZipLib(); const zip = new JSZip();
    for (let i = 0; i < doc.pdf.numPages; i++) { const url = await pageToDataUrl(doc, i, 2); const u2 = fmt === 'jpg' ? await toJpg(url) : url; zip.file('trang-' + (i + 1) + '.' + fmt, dataUrlToBytes(u2)); busyOn('Trang ' + (i + 1) + '/' + doc.pdf.numPages + '…'); }
    await saveOut(doc, (doc.name || 'pdf').replace(/\.pdf$/i, '') + '-anh-' + fmt + '.zip', await zip.generateAsync({ type: 'uint8array' }), 'application/zip');
    toast('Đã xuất ' + doc.pdf.numPages + ' ảnh ' + fmt.toUpperCase(), 'ok');
  } catch (e) { toast('Lỗi: ' + e.message, 'err'); } busyOff();
}
function toJpg(url) { return new Promise(res => { const im = new Image(); im.onload = () => { const c = h('canvas'); c.width = im.width; c.height = im.height; const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(im, 0, 0); res(c.toDataURL('image/jpeg', 0.9)); }; im.src = url; }); }

// ---------------------------------------------------------------- OCR
export async function ocrDialog(doc) {
  if (!doc.pdf) return toast('Mở một PDF (hoặc ảnh) trước', 'err');
  const v = await formDlg('Nhận dạng chữ (OCR) — ngoại tuyến', [
    { key: 'range', label: 'Trang (vd 1-3 hoặc để trống = tất cả)', value: '' },
    { key: 'lang', type: 'select', label: 'Ngôn ngữ', value: 'vie+eng', options: [['vie+eng', 'Tiếng Việt + Anh'], ['vie', 'Tiếng Việt'], ['eng', 'Tiếng Anh']] },
    { key: 'out', type: 'radio', value: 'view', options: [['view', 'Hiện văn bản'], ['txt', 'Lưu .txt'], ['docx', 'Lưu .docx'], ['embed', 'Tạo PDF có thể tìm kiếm (ghép lớp chữ)']] }
  ]);
  if (!v) return;
  const idx = v.range.trim() ? parseRange(v.range, doc.pdf.numPages) : doc.pageInfo.map((_, i) => i);
  if (!idx.length) return toast('Phạm vi không hợp lệ', 'err');
  const bar = h('div.ocr-bar'); const label = h('div.muted', 'Đang chuẩn bị bộ nhận dạng…');
  const overlay = h('div.ocr-overlay', h('div.ocr-card', h('b', 'Đang OCR (ngoại tuyến)'), label, h('div.ocr-prog', bar)));
  document.body.appendChild(overlay);
  const closeOverlay = () => { if (overlay.parentNode) overlay.parentNode.removeChild(overlay); };
  const texts = {};
  try {
    for (let k = 0; k < idx.length; k++) {
      const i = idx[k]; label.textContent = 'Trang ' + (i + 1) + ' / ' + idx.length + ' (' + (k + 1) + ')';
      const url = await pageToDataUrl(doc, i, 2);
      const r = await Host.ocr(url, v.lang, (st, p) => { bar.style.width = Math.round(((k + (p || 0)) / idx.length) * 100) + '%'; });
      texts[i] = r.text || ''; doc.ocrText[i] = texts[i];
    }
    closeOverlay();
    const joined = idx.map(i => texts[i]).join('\n\n');
    if (v.out === 'txt') await saveOut(doc, (doc.name || 'pdf').replace(/\.pdf$/i, '') + '.txt', utf8Encode(joined), 'text/plain');
    else if (v.out === 'docx') await saveOut(doc, (doc.name || 'pdf').replace(/\.pdf$/i, '') + '.docx', await buildDocx(joined.split(/\n/)), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    else if (v.out === 'embed') { await embedOcrLayer(doc, texts); toast('Đã tạo PDF có lớp chữ tìm kiếm được', 'ok'); }
    else showText(doc, 'Kết quả OCR', joined);
    toast('OCR hoàn tất: ' + idx.length + ' trang', 'ok');
  } catch (e) { closeOverlay(); toast('Lỗi OCR: ' + e.message, 'err'); }
}
async function embedOcrLayer(doc, texts) {
  const { PDFDocument, rgb } = await PDFLib(); const src = await PDFDocument.load(doc.bytes); const pages = src.getPages();
  const uf = await uniFont(src, false); const font = uf.font; // v0.1.6: Vietnamese text layer keeps diacritics (was ASCII only)
  Object.keys(texts).forEach(ki => { const i = +ki; const page = pages[i]; const t = uf.text(texts[i] || '').replace(/[\u0000-\u0008\u000b-\u001f]/g, ' ').slice(0, 4000); if (page && t.trim()) page.drawText(t, { x: 4, y: 4, size: 1, font, color: rgb(1, 1, 1), opacity: 0, lineHeight: 2, maxWidth: page.getWidth() - 8 }); });
  await doc.commit(await src.save(), null);
}

// ---------------------------------------------------------------- text / html / docx extraction
export async function extractText(doc) {
  if (!doc.pdf) return '';
  let out = [];
  for (let i = 0; i < doc.pdf.numPages; i++) {
    const page = doc.pageInfo[i]._page || await doc.measureOne(i);
    const tc = await page.getTextContent(); let line = '', last = null, s = '';
    tc.items.forEach(it => { if (last && Math.abs(it.transform[5] - last) > 3) { s += line.trimEnd() + '\n'; line = ''; } line += it.str + (it.hasEOL ? '\n' : ' '); last = it.transform[5]; });
    s += line;
    if (!s.trim() && doc.ocrText[i]) s = doc.ocrText[i];
    out.push(s.trim());
  }
  return out.join('\n\n');
}
export async function toHtml(doc) {
  const text = await extractText(doc);
  const paras = text.split(/\n{2,}/).map(p => '<p>' + esc(p).replace(/\n/g, '<br>') + '</p>').join('\n');
  return '<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>' + esc(doc.name || 'PDF') + '</title><style>body{font-family:Calibri,Arial,sans-serif;max-width:820px;margin:32px auto;padding:0 16px;line-height:1.5}p{margin:0 0 10px}</style></head><body>' + paras + '</body></html>';
}
export async function toDocx(doc) { const text = await extractText(doc); return buildDocx(text.split(/\n/)); }
export async function buildDocx(lines) {
  const JSZip = await JSZipLib(); const zip = new JSZip();
  const body = lines.map(l => '<w:p>' + (l.trim() ? '<w:r><w:t xml:space="preserve">' + esc(l) + '</w:t></w:r>' : '') + '</w:p>').join('');
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'uint8array' });
}

// ---------------------------------------------------------------- flatten (rasterize to image PDF)
export async function flatten(doc) {
  const imgs = []; for (let i = 0; i < doc.pdf.numPages; i++) imgs.push({ name: 'p' + i + '.png', bytes: dataUrlToBytes(await pageToDataUrl(doc, i, 2)) });
  return imagesToPdf(imgs);
}

// ---------------------------------------------------------------- watermark / page numbers / header
export async function watermarkDialog(doc) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  const v = await formDlg('Thêm hình mờ (Watermark)', [{ key: 'text', label: 'Nội dung', value: 'MẬT' }, { key: 'color', type: 'color', label: 'Màu', value: '#ff0000' }, { key: 'size', type: 'number', label: 'Cỡ chữ', value: 60 }, { key: 'opacity', type: 'number', label: 'Độ mờ (0-100)', value: 20 }]);
  if (!v) return;
  busyOn('Đang thêm hình mờ…');
  try {
    const { PDFDocument, degrees, rgb, StandardFonts } = await PDFLib(); const src = await PDFDocument.load(doc.bytes);
    const uf = await uniFont(src, true); const font = uf.font, T = uf.text; const c = hexRgb(v.color);
    src.getPages().forEach(pg => { const { width, height } = pg.getSize(); const t = T(v.text); const tw = font.widthOfTextAtSize(t, +v.size); pg.drawText(t, { x: width / 2 - tw * 0.35, y: height / 2 - tw * 0.35, size: +v.size, font, color: rgb(c[0], c[1], c[2]), rotate: degrees(45), opacity: (+v.opacity) / 100 }); });
    await doc.commit(await src.save(), 'Đã thêm hình mờ');
  } catch (e) { toast('Lỗi: ' + e.message, 'err'); } busyOff();
}
export async function pageNumbersDialog(doc) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  const v = await formDlg('Đánh số trang', [{ key: 'pos', type: 'select', label: 'Vị trí', value: 'bc', options: [['bc', 'Giữa dưới'], ['br', 'Phải dưới'], ['bl', 'Trái dưới'], ['tc', 'Giữa trên']] }, { key: 'fmt', label: 'Định dạng ({n}=số, {t}=tổng)', value: 'Trang {n}/{t}' }]);
  if (!v) return;
  busyOn('Đang đánh số…');
  try {
    const { PDFDocument, rgb, StandardFonts } = await PDFLib(); const src = await PDFDocument.load(doc.bytes);
    const uf = await uniFont(src, false); const font = uf.font; const pages = src.getPages(); const tot = pages.length;
    pages.forEach((pg, i) => { const { width, height } = pg.getSize(); const t = uf.text(v.fmt.replace('{n}', i + 1).replace('{t}', tot)); const w = font.widthOfTextAtSize(t, 10); let x = width / 2 - w / 2, y = 18; if (v.pos === 'br') x = width - w - 24; if (v.pos === 'bl') x = 24; if (v.pos === 'tc') y = height - 24; pg.drawText(t, { x, y, size: 10, font, color: rgb(0.2, 0.2, 0.2) }); });
    await doc.commit(await src.save(), 'Đã đánh số trang');
  } catch (e) { toast('Lỗi: ' + e.message, 'err'); } busyOff();
}
function hexRgb(hx) { const m = /^#?([0-9a-f]{6})$/i.exec(hx || '#000000'); if (!m) return [0, 0, 0]; const n = parseInt(m[1], 16); return [(n >> 16 & 255) / 255, (n >> 8 & 255) / 255, (n & 255) / 255]; }
// v0.1.6: Unicode font (DejaVu Sans via fontkit) so Vietnamese watermarks / page numbers keep their diacritics
async function uniFont(src, bold) {
  const L = await PDFLib();
  try {
    if (!window.fontkit) await loadScript('vendor/pdffont/fontkit.umd.min.js');
    const r = await fetch('vendor/pdffont/' + (bold ? 'DejaVuSans-Bold.ttf' : 'DejaVuSans.ttf'));
    if (!r.ok) throw new Error('font ' + r.status);
    src.registerFontkit(window.fontkit);
    const font = await src.embedFont(new Uint8Array(await r.arrayBuffer()), { subset: true });
    return { font, text: s => s || '' };
  } catch (e) {
    console.warn('uniFont fallback', e);
    return { font: await src.embedFont(bold ? L.StandardFonts.HelveticaBold : L.StandardFonts.Helvetica), text: asciiSafe };
  }
}
function asciiSafe(s) { return (s || '').normalize ? s.normalize('NFD').replace(/[̀-ͯ]/g, '') : s; }

// ---------------------------------------------------------------- save helper
export async function saveOut(doc, name, bytes, mime) { const r = await Host.saveFile(name, bytes, { mime, saveAs: true }); if (r) toast('Đã lưu: ' + r.name, 'ok'); return r; }
function showText(doc, title, text) {
  const ta = h('textarea.pdf-textout', { readonly: false }); ta.value = text;
  dialog({ title, width: 720, body: h('div', h('div.muted', 'Có thể chỉnh sửa rồi sao chép hoặc lưu.'), ta), buttons: [{ label: 'Sao chép', run: () => { ta.select(); try { document.execCommand('copy'); toast('Đã sao chép'); } catch (e) {} } }, { label: 'Lưu .txt', run: () => saveOut(doc, (doc.name || 'pdf').replace(/\.pdf$/i, '') + '.txt', utf8Encode(ta.value), 'text/plain') }, { label: 'Đóng', primary: true }] });
}

// ---------------------------------------------------------------- menus / commands
export function pageMenu(doc) {
  const sel = doc.sel.length ? doc.sel : [doc.cur];
  return [
    { label: 'Xoay phải 90°', icon: 'rotate-cw', run: () => rotatePages(doc, sel, 90) },
    { label: 'Xoay trái 90°', icon: 'rotate-ccw', run: () => rotatePages(doc, sel, 270) },
    { label: 'Xoay 180°', icon: 'flip-vertical', run: () => rotatePages(doc, sel, 180) },
    { sep: true },
    { label: 'Trích các trang này…', icon: 'scissors', run: () => { const r = sel.map(i => i + 1).join(','); formDlg('Trích trang', [{ key: 'range', label: 'Phạm vi', value: r }]).then(v => { if (v) { doc._presetRange = v.range; splitRange(doc, v.range, 'extract'); } }); } },
    { label: 'Xuất trang thành ảnh', icon: 'image', run: () => pdfToImages(doc, 'png') },
    { sep: true },
    { label: 'Xoá trang', icon: 'trash-2', run: () => deletePages(doc, sel) }
  ];
}
export async function splitRange(doc, range, mode) {
  const idx = parseRange(range, doc.pdf.numPages); if (!idx.length) return toast('Phạm vi không hợp lệ', 'err');
  const { PDFDocument } = await PDFLib(); const src = await PDFDocument.load(doc.bytes); const nd = await PDFDocument.create();
  const pages = await nd.copyPages(src, idx); pages.forEach(p => nd.addPage(p));
  await saveOut(doc, (doc.name || 'pdf').replace(/\.pdf$/i, '') + '-trich.pdf', await nd.save(), 'application/pdf');
}
export async function exportAs(doc, ext, mime, label) {
  if (!doc.pdf) return toast('Mở một PDF trước', 'err');
  busyOn('Đang chuyển đổi ' + (label || ext) + '…');
  try { const bytes = await doc.save(ext); await saveOut(doc, (doc.name || 'pdf').replace(/\.[^.]+$/i, '') + '.' + (ext === 'pdf-flat' ? 'pdf' : ext), bytes, mime); }
  catch (e) { toast('Lỗi: ' + e.message, 'err'); } busyOff();
}
export function extraCommands(doc) {
  return [
    { label: 'OCR nhận dạng chữ', icon: 'scan-text', run: () => ocrDialog(doc) },
    { label: 'Ghép PDF', icon: 'combine', run: () => mergeDialog(doc) },
    { label: 'Tách PDF', icon: 'scissors', run: () => splitDialog(doc) },
    { label: 'Xuất trang thành ảnh', icon: 'image', run: () => pdfToImages(doc, 'png') },
    { label: 'Chú thích & Ký', icon: 'pen-line', run: () => import('./pro.js').then(P => P.annotate(doc)) },
    { label: 'Đặt mật khẩu PDF', icon: 'lock', run: () => import('./pro.js').then(P => P.protectDialog(doc)) },
    { label: 'Nén PDF', icon: 'archive', run: () => import('./pro.js').then(P => P.compressDialog(doc)) }
  ];
}
