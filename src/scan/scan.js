// v0.0.7 — CamScanner: scan documents with the camera (or from photos), automatic edge detection, perspective correction,
// enhancement filters (Magic color, B&W, shadow removal…), multi-page → PDF / JPG / OCR text — fully offline.
// Algorithms follow the open-source scanners listed in the project sheet (DocumentScanner, Python-Document-Scanner-OpenCV,
// WonderScan, Document-Scanner, PDF_Magic-Color_filter): edge map → document quadrilateral → homography warp → adaptive threshold.
import { h, ic, esc, clamp, loadImage, readFileAsDataUrl, dataUrlToBytes, PDFLib, JSZipLib, fmtDate } from '../core/util.js';
import { toast, busy, dialog } from '../core/ui.js';
import Host from '../core/host.js';

// ---------------------------------------------------------------- image helpers
export function imgToCanvas(img, max) {
  const s = Math.min(1, (max || 2400) / Math.max(img.naturalWidth || img.width, img.naturalHeight || img.height));
  const c = document.createElement('canvas'); c.width = Math.max(1, Math.round((img.naturalWidth || img.width) * s)); c.height = Math.max(1, Math.round((img.naturalHeight || img.height) * s));
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); return c;
}
function gray(id) { const d = id.data, n = id.width * id.height, g = new Float32Array(n); for (let i = 0; i < n; i++) g[i] = 0.299 * d[4 * i] + 0.587 * d[4 * i + 1] + 0.114 * d[4 * i + 2]; return g; }
function boxBlur(src, w, hgt, r) {
  const tmp = new Float32Array(src.length), out = new Float32Array(src.length), k = 2 * r + 1;
  for (let y = 0; y < hgt; y++) { let acc = 0; const o = y * w; for (let x = -r; x <= r; x++) acc += src[o + clamp(x, 0, w - 1)]; for (let x = 0; x < w; x++) { tmp[o + x] = acc / k; acc += src[o + Math.min(w - 1, x + r + 1)] - src[o + Math.max(0, x - r)]; } }
  for (let x = 0; x < w; x++) { let acc = 0; for (let y = -r; y <= r; y++) acc += tmp[clamp(y, 0, hgt - 1) * w + x]; for (let y = 0; y < hgt; y++) { out[y * w + x] = acc / k; acc += tmp[Math.min(hgt - 1, y + r + 1) * w + x] - tmp[Math.max(0, y - r) * w + x]; } }
  return out;
}
function hull(pts) {
  pts = pts.slice().sort((a, b) => a[0] - b[0] || a[1] - b[1]); if (pts.length < 3) return pts;
  const cr = (o, a, b) => (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]); const lo = [], up = [];
  for (const p of pts) { while (lo.length >= 2 && cr(lo[lo.length - 2], lo[lo.length - 1], p) <= 0) lo.pop(); lo.push(p); }
  for (let i = pts.length - 1; i >= 0; i--) { const p = pts[i]; while (up.length >= 2 && cr(up[up.length - 2], up[up.length - 1], p) <= 0) up.pop(); up.push(p); }
  up.pop(); lo.pop(); return lo.concat(up);
}
function polyArea(q) { let a = 0; for (let i = 0; i < q.length; i++) { const p = q[i], n = q[(i + 1) % q.length]; a += p[0] * n[1] - n[0] * p[1]; } return Math.abs(a) / 2; }

// ---------------------------------------------------------------- automatic document detection → [tl, tr, br, bl] in canvas pixels
export function detectQuad(canvas) {
  const W = canvas.width, H = canvas.height, s = Math.min(1, 360 / Math.max(W, H));
  const w = Math.max(8, Math.round(W * s)), hh = Math.max(8, Math.round(H * s));
  const c = document.createElement('canvas'); c.width = w; c.height = hh; const cx = c.getContext('2d'); cx.drawImage(canvas, 0, 0, w, hh);
  const g = boxBlur(gray(cx.getImageData(0, 0, w, hh)), w, hh, 2);
  // Sobel magnitude
  const mag = new Float32Array(w * hh); let mx = 0;
  for (let y = 1; y < hh - 1; y++) for (let x = 1; x < w - 1; x++) {
    const i = y * w + x;
    const gx = -g[i - w - 1] - 2 * g[i - 1] - g[i + w - 1] + g[i - w + 1] + 2 * g[i + 1] + g[i + w + 1];
    const gy = -g[i - w - 1] - 2 * g[i - w] - g[i - w + 1] + g[i + w - 1] + 2 * g[i + w] + g[i + w + 1];
    const m = Math.sqrt(gx * gx + gy * gy); mag[i] = m; if (m > mx) mx = m;
  }
  // Otsu threshold on magnitudes
  const hist = new Array(256).fill(0); for (let i = 0; i < mag.length; i++) hist[Math.min(255, (mag[i] / (mx || 1) * 255) | 0)]++;
  let sum = 0; for (let i = 0; i < 256; i++) sum += i * hist[i]; let sB = 0, wB = 0, best = 0, th = 40;
  for (let i = 0; i < 256; i++) { wB += hist[i]; if (!wB) continue; const wF = mag.length - wB; if (!wF) break; sB += i * hist[i]; const mB = sB / wB, mF = (sum - sB) / wF, v = wB * wF * (mB - mF) * (mB - mF); if (v > best) { best = v; th = i; } }
  const T = Math.max(th, 25) / 255 * mx;
  // connected components of edge pixels; keep the component whose hull is largest (the sheet border)
  const lab = new Int32Array(w * hh); let bestPts = null, bestA = 0, L = 0; const stack = [];
  for (let i = 0; i < mag.length; i++) {
    if (lab[i] || mag[i] < T) continue; L++; lab[i] = L; stack.push(i); const pts = [];
    while (stack.length) { const j = stack.pop(); const x = j % w, y = (j / w) | 0; pts.push([x, y]);
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) { const nx = x + dx, ny = y + dy; if (nx < 0 || ny < 0 || nx >= w || ny >= hh) continue; const k = ny * w + nx; if (!lab[k] && mag[k] >= T) { lab[k] = L; stack.push(k); } } }
    if (pts.length < 30) continue; const hl = hull(pts); const a = polyArea(hl); if (a > bestA) { bestA = a; bestPts = hl; }
  }
  const full = [[0, 0], [W, 0], [W, H], [0, H]];
  if (!bestPts || bestA < w * hh * 0.12) return { quad: full, auto: false };
  // corners = hull points extreme along the 4 diagonals
  const pick = f => bestPts.reduce((b, p) => f(p) > f(b) ? p : b, bestPts[0]);
  let q = [pick(p => -p[0] - p[1]), pick(p => p[0] - p[1]), pick(p => p[0] + p[1]), pick(p => -p[0] + p[1])];
  // the Sobel edge is ~2-3 px wide (blur + kernel): pull the corners 2.5 px inwards so no background strip remains
  const cx0 = (q[0][0] + q[1][0] + q[2][0] + q[3][0]) / 4, cy0 = (q[0][1] + q[1][1] + q[2][1] + q[3][1]) / 4;
  q = q.map(p => { const dx = cx0 - p[0], dy = cy0 - p[1], L = Math.hypot(dx, dy) || 1; return [(p[0] + 2.5 * dx / L) / s, (p[1] + 2.5 * dy / L) / s]; });
  if (polyArea(q) < W * H * 0.12) return { quad: full, auto: false };
  return { quad: q, auto: true };
}

// ---------------------------------------------------------------- perspective transform
function solve(A, b) { const n = b.length; for (let i = 0; i < n; i++) { let p = i; for (let r = i + 1; r < n; r++) if (Math.abs(A[r][i]) > Math.abs(A[p][i])) p = r; [A[i], A[p]] = [A[p], A[i]]; [b[i], b[p]] = [b[p], b[i]]; for (let r = i + 1; r < n; r++) { const f = A[r][i] / A[i][i]; for (let c = i; c < n; c++) A[r][c] -= f * A[i][c]; b[r] -= f * b[i]; } } const x = new Array(n); for (let i = n - 1; i >= 0; i--) { let s = b[i]; for (let c = i + 1; c < n; c++) s -= A[i][c] * x[c]; x[i] = s / A[i][i]; } return x; }
export function homography(src, dst) { // maps src points → dst points (4 pairs)
  const A = [], b = [];
  for (let i = 0; i < 4; i++) { const [x, y] = src[i], [u, v] = dst[i]; A.push([x, y, 1, 0, 0, 0, -u * x, -u * y]); b.push(u); A.push([0, 0, 0, x, y, 1, -v * x, -v * y]); b.push(v); }
  const m = solve(A, b); return [m[0], m[1], m[2], m[3], m[4], m[5], m[6], m[7], 1];
}
const dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
export function warp(canvas, quad, maxSide) {
  let ow = Math.max(dist(quad[0], quad[1]), dist(quad[3], quad[2])), oh = Math.max(dist(quad[0], quad[3]), dist(quad[1], quad[2]));
  const sc = Math.min(1, (maxSide || 2200) / Math.max(ow, oh)); ow = Math.max(1, Math.round(ow * sc)); oh = Math.max(1, Math.round(oh * sc));
  const Hm = homography([[0, 0], [ow, 0], [ow, oh], [0, oh]], quad); // output → source
  const si = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height), sd = si.data, SW = canvas.width, SH = canvas.height;
  const out = document.createElement('canvas'); out.width = ow; out.height = oh; const oc = out.getContext('2d'); const oi = oc.createImageData(ow, oh), od = oi.data;
  for (let y = 0; y < oh; y++) for (let x = 0; x < ow; x++) {
    const z = Hm[6] * x + Hm[7] * y + 1, sx = (Hm[0] * x + Hm[1] * y + Hm[2]) / z, sy = (Hm[3] * x + Hm[4] * y + Hm[5]) / z;
    const x0 = Math.floor(sx), y0 = Math.floor(sy), fx = sx - x0, fy = sy - y0, o = 4 * (y * ow + x);
    if (x0 < 0 || y0 < 0 || x0 >= SW - 1 || y0 >= SH - 1) { od[o] = od[o + 1] = od[o + 2] = 255; od[o + 3] = 255; continue; }
    const i00 = 4 * (y0 * SW + x0), i10 = i00 + 4, i01 = i00 + 4 * SW, i11 = i01 + 4;
    for (let c = 0; c < 3; c++) od[o + c] = (sd[i00 + c] * (1 - fx) + sd[i10 + c] * fx) * (1 - fy) + (sd[i01 + c] * (1 - fx) + sd[i11 + c] * fx) * fy;
    od[o + 3] = 255;
  }
  oc.putImageData(oi, 0, 0); return out;
}

// ---------------------------------------------------------------- enhancement filters
export const FILTERS = [['original', 'Gốc'], ['magic', 'Màu kỳ diệu'], ['shadow', 'Bỏ bóng'], ['gray', 'Thang xám'], ['bw', 'Đen trắng'], ['lighten', 'Làm sáng']];
export function applyFilter(src, kind) {
  const w = src.width, hh = src.height; const out = document.createElement('canvas'); out.width = w; out.height = hh; const ctx = out.getContext('2d'); ctx.drawImage(src, 0, 0);
  if (kind === 'original' || !kind) return out;
  const id = ctx.getImageData(0, 0, w, hh), d = id.data, n = w * hh;
  if (kind === 'gray') { for (let i = 0; i < n; i++) { const v = 0.299 * d[4 * i] + 0.587 * d[4 * i + 1] + 0.114 * d[4 * i + 2]; d[4 * i] = d[4 * i + 1] = d[4 * i + 2] = v; } }
  else if (kind === 'lighten') { for (let i = 0; i < 4 * n; i++) if ((i & 3) !== 3) d[i] = Math.min(255, d[i] * 1.18 + 18); }
  else {
    const g = gray(id); const r = Math.max(8, Math.round(Math.max(w, hh) / 40)); const bg = boxBlur(boxBlur(g, w, hh, r), w, hh, r); // background (paper + shadows)
    if (kind === 'bw') { for (let i = 0; i < n; i++) { const v = g[i] / Math.max(1, bg[i]) > 0.86 ? 255 : 0; d[4 * i] = d[4 * i + 1] = d[4 * i + 2] = v; } } // adaptive threshold (Bradley/Sauvola style)
    else { // shadow removal / magic color: divide every channel by the background, then stretch contrast
      const boost = kind === 'magic' ? 1.25 : 1.0;
      for (let i = 0; i < n; i++) { const k = 255 / Math.max(1, bg[i]); for (let c = 0; c < 3; c++) { let v = d[4 * i + c] * k; if (kind === 'magic') v = (v - 128) * boost + 128 + 6; d[4 * i + c] = v < 0 ? 0 : v > 255 ? 255 : v; } }
      if (kind === 'magic') { for (let i = 0; i < n; i++) { const R = d[4 * i], G = d[4 * i + 1], B = d[4 * i + 2], m = (R + G + B) / 3; for (let c = 0; c < 3; c++) d[4 * i + c] = clamp(m + (d[4 * i + c] - m) * 1.35, 0, 255); } }
    }
  }
  ctx.putImageData(id, 0, 0); return out;
}
export function rotate90(src) { const c = document.createElement('canvas'); c.width = src.height; c.height = src.width; const x = c.getContext('2d'); x.translate(c.width, 0); x.rotate(Math.PI / 2); x.drawImage(src, 0, 0); return c; }

// full pipeline used by the UI and the self test
export function processPage(pg) { const w = warp(pg.src, pg.quad); const f = applyFilter(w, pg.filter); let r = f; for (let i = 0; i < ((pg.rot || 0) / 90 | 0) % 4; i++) r = rotate90(r); return r; }

export async function pagesToPdf(canvases, opts) {
  const PL = await PDFLib(); const pdf = await PL.PDFDocument.create(); opts = opts || {};
  for (const c of canvases) {
    const jpg = await pdf.embedJpg(dataUrlToBytes(c.toDataURL('image/jpeg', opts.q || 0.82)));
    let pw = c.width * 0.75, ph = c.height * 0.75;
    if (opts.a4 !== false) { const A4 = [595.28, 841.89]; const land = c.width > c.height; const [aw, ah] = land ? [A4[1], A4[0]] : A4; const s = Math.min(aw / pw, ah / ph); const page = pdf.addPage([aw, ah]); page.drawImage(jpg, { x: (aw - pw * s) / 2, y: (ah - ph * s) / 2, width: pw * s, height: ph * s }); continue; }
    const page = pdf.addPage([pw, ph]); page.drawImage(jpg, { x: 0, y: 0, width: pw, height: ph });
  }
  pdf.setTitle('Tài liệu quét ' + fmtDate()); pdf.setCreator('Yomi-Office Open Source CamScanner'); pdf.setProducer('Yomi-Office Open Source 0.0.7');
  return pdf.save();
}

// ---------------------------------------------------------------- camera
async function cameraCapture() {
  if (!(navigator.mediaDevices && navigator.mediaDevices.getUserMedia)) return null;
  let stream; try { stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 2560 }, height: { ideal: 1440 } }, audio: false }); } catch (e) { return null; }
  return new Promise(res => {
    const video = h('video', { autoplay: true, playsinline: true, muted: true, style: { width: '100%', maxHeight: '70vh', background: '#000', borderRadius: '8px' } }); video.srcObject = stream; video.setAttribute('playsinline', '');
    const stop = () => stream.getTracks().forEach(t => t.stop());
    dialog({ title: 'Chụp tài liệu', width: 760, body: h('div', video, h('div.muted', 'Đặt tài liệu trên nền tối, giữ máy song song với trang giấy.')), buttons: [
      { label: 'Chụp', primary: true, run: () => { const c = document.createElement('canvas'); c.width = video.videoWidth; c.height = video.videoHeight; c.getContext('2d').drawImage(video, 0, 0); stop(); res(c); } },
      { label: 'Hủy', run: () => { stop(); res(null); } }] }).then(v => { if (v === null) { stop(); res(null); } });
  });
}

// ---------------------------------------------------------------- UI
export async function openScanner(shell) {
  const pages = []; let cur = -1;
  const thumbs = h('div.scan-thumbs'); const stage = h('div.scan-stage'); const tools = h('div.scan-tools');
  const root = h('div.scanner', h('div.scan-bar', h('button.btn', { onclick: () => close() }, ic('arrow-left', 16), ' Đóng'), h('b', ' CamScanner — Quét tài liệu'), h('span', { style: { flex: 1 } }),
    h('button.btn', { onclick: () => addCamera() }, ic('camera', 16), ' Chụp ảnh'), h('button.btn', { onclick: () => addFiles() }, ic('image-plus', 16), ' Thêm ảnh'),
    h('button.btn.primary', { onclick: () => exportPdf() }, ic('file-type', 16), ' Lưu PDF'),
    h('button.btn', { onclick: () => exportJpg() }, ic('image', 16), ' Lưu JPG'),
    h('button.btn', { onclick: () => ocr() }, ic('scan-text', 16), ' OCR → văn bản'),
    h('button.btn', { onclick: () => toPdfTool() }, ic('external-link', 16), ' Mở bằng Công cụ PDF')),
    h('div.scan-main', thumbs, h('div.scan-center', stage, tools)));
  const close = () => { if (root.parentNode) root.parentNode.removeChild(root); };
  const addCanvas = c => { const det = detectQuad(c); pages.push({ src: c, quad: det.quad, auto: det.auto, filter: 'magic', rot: 0 }); cur = pages.length - 1; renderThumbs(); renderStage(); if (!det.auto) toast('Không tự nhận được mép giấy — kéo 4 góc để chọn vùng'); };
  const addFiles = async () => { const fs = await Host.openFiles('image/*', true); const b = busy('Đang xử lý ảnh…'); try { for (const f of fs) { const url = 'data:image/' + (/png$/i.test(f.name) ? 'png' : 'jpeg') + ';base64,' + (await import('../core/util.js')).b64FromBytes(f.bytes); addCanvas(imgToCanvas(await loadImage(url))); } } finally { b.done(); } };
  const addCamera = async () => { const c = await cameraCapture(); if (c) addCanvas(imgToCanvas(c)); else if (c === null && !(navigator.mediaDevices && navigator.mediaDevices.getUserMedia)) { const inp = h('input', { type: 'file', accept: 'image/*', capture: 'environment', style: { display: 'none' } }); inp.addEventListener('change', async () => { if (inp.files[0]) addCanvas(imgToCanvas(await loadImage(await readFileAsDataUrl(inp.files[0])))); }); document.body.appendChild(inp); inp.click(); } };
  const renderThumbs = () => {
    thumbs.innerHTML = '';
    pages.forEach((p, i) => { const c = processPage(Object.assign({}, p, { src: p.src })); const t = document.createElement('canvas'); const s = 110 / Math.max(c.width, c.height); t.width = c.width * s; t.height = c.height * s; t.getContext('2d').drawImage(c, 0, 0, t.width, t.height);
      const el = h('div.scan-th' + (i === cur ? '.on' : ''), t, h('span', String(i + 1))); el.addEventListener('click', () => { cur = i; renderThumbs(); renderStage(); }); thumbs.appendChild(el); });
    if (!pages.length) thumbs.appendChild(h('div.muted', { style: { padding: '12px' } }, 'Chưa có trang nào'));
  };
  let mode = 'crop';
  const renderStage = () => {
    stage.innerHTML = ''; tools.innerHTML = '';
    if (cur < 0) { stage.appendChild(h('div.scan-empty', ic('scan-line', 64), h('h3', 'Quét tài liệu'), h('p.muted', 'Chụp ảnh tài liệu bằng máy ảnh hoặc chọn ảnh có sẵn. Ứng dụng tự nhận mép giấy, nắn thẳng và làm rõ chữ — không cần Internet.'), h('div', h('button.btn.primary', { onclick: addCamera }, ic('camera', 16), ' Chụp ảnh'), ' ', h('button.btn', { onclick: addFiles }, ic('image-plus', 16), ' Chọn ảnh')))); return; }
    const p = pages[cur];
    if (mode === 'crop') {
      const wrap = h('div.scan-crop'); const cv = document.createElement('canvas'); const maxW = Math.min(stage.clientWidth || 800, 900) - 20, maxH = (window.innerHeight - 220);
      const s = Math.min(maxW / p.src.width, maxH / p.src.height, 1); cv.width = p.src.width * s; cv.height = p.src.height * s; wrap.appendChild(cv);
      const draw = () => { const x = cv.getContext('2d'); x.drawImage(p.src, 0, 0, cv.width, cv.height); x.fillStyle = 'rgba(15,108,189,.18)'; x.strokeStyle = '#0f6cbd'; x.lineWidth = 2; x.beginPath(); p.quad.forEach((q, i) => i ? x.lineTo(q[0] * s, q[1] * s) : x.moveTo(q[0] * s, q[1] * s)); x.closePath(); x.fill(); x.stroke(); };
      p.quad.forEach((q, k) => { const hd = h('div.scan-h'); const pos = () => { hd.style.left = (p.quad[k][0] * s) + 'px'; hd.style.top = (p.quad[k][1] * s) + 'px'; }; pos();
        const start = e => { e.preventDefault(); const r = cv.getBoundingClientRect(); const mv = ev => { const t = ev.touches ? ev.touches[0] : ev; p.quad[k] = [clamp((t.clientX - r.left) / s, 0, p.src.width), clamp((t.clientY - r.top) / s, 0, p.src.height)]; pos(); draw(); }; const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); document.removeEventListener('touchmove', mv); document.removeEventListener('touchend', up); renderThumbs(); }; document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up); document.addEventListener('touchmove', mv, { passive: false }); document.addEventListener('touchend', up); };
        hd.addEventListener('mousedown', start); hd.addEventListener('touchstart', start, { passive: false }); wrap.appendChild(hd); });
      draw(); stage.appendChild(wrap);
      tools.appendChild(h('div.scan-row', h('button.btn', { onclick: () => { const d = detectQuad(p.src); p.quad = d.quad; renderStage(); renderThumbs(); toast(d.auto ? 'Đã nhận mép giấy tự động' : 'Không nhận được mép — dùng toàn ảnh'); } }, ic('scan', 15), ' Tự nhận mép'),
        h('button.btn', { onclick: () => { p.quad = [[0, 0], [p.src.width, 0], [p.src.width, p.src.height], [0, p.src.height]]; renderStage(); renderThumbs(); } }, ic('maximize', 15), ' Toàn ảnh'),
        h('button.btn.primary', { onclick: () => { mode = 'filter'; renderStage(); } }, ic('check', 15), ' Cắt & nắn thẳng')));
    } else {
      const out = processPage(p); const img = h('img.scan-result', { src: out.toDataURL('image/jpeg', 0.9) }); stage.appendChild(img);
      tools.appendChild(h('div.scan-row', FILTERS.map(f => { const b = h('button.btn' + (p.filter === f[0] ? '.primary' : ''), f[1]); b.addEventListener('click', () => { p.filter = f[0]; renderStage(); renderThumbs(); }); return b; })));
      tools.appendChild(h('div.scan-row', h('button.btn', { onclick: () => { mode = 'crop'; renderStage(); } }, ic('crop', 15), ' Chỉnh vùng cắt'), h('button.btn', { onclick: () => { p.rot = ((p.rot || 0) + 90) % 360; renderStage(); renderThumbs(); } }, ic('rotate-cw', 15), ' Xoay'),
        h('button.btn', { onclick: () => { if (cur > 0) { pages.splice(cur - 1, 0, pages.splice(cur, 1)[0]); cur--; renderThumbs(); } } }, ic('arrow-up', 15), ' Lên'), h('button.btn', { onclick: () => { if (cur < pages.length - 1) { pages.splice(cur + 1, 0, pages.splice(cur, 1)[0]); cur++; renderThumbs(); } } }, ic('arrow-down', 15), ' Xuống'),
        h('button.btn.danger', { onclick: () => { pages.splice(cur, 1); cur = Math.min(cur, pages.length - 1); mode = 'filter'; renderThumbs(); renderStage(); } }, ic('trash-2', 15), ' Xoá trang')));
    }
  };
  const results = () => pages.map(processPage);
  const need = () => { if (!pages.length) { toast('Chưa có trang nào — hãy chụp hoặc thêm ảnh', 'err'); return false; } return true; };
  const exportPdf = async () => { if (!need()) return; const b = busy('Đang tạo PDF…'); try { const bytes = await pagesToPdf(results()); await Host.saveFile('Scan_' + fmtDate().replace(/\//g, '-') + '.pdf', bytes, { saveAs: true, mime: 'application/pdf' }); } catch (e) { toast('Lỗi: ' + e.message, 'err'); } finally { b.done(); } };
  const exportJpg = async () => { if (!need()) return; const rs = results(); if (rs.length === 1) return Host.saveFile('Scan.jpg', dataUrlToBytes(rs[0].toDataURL('image/jpeg', 0.9)), { saveAs: true, mime: 'image/jpeg' }); const JSZip = await JSZipLib(); const z = new JSZip(); rs.forEach((c, i) => z.file('Scan_' + (i + 1) + '.jpg', dataUrlToBytes(c.toDataURL('image/jpeg', 0.9)))); Host.saveFile('Scan_anh.zip', await z.generateAsync({ type: 'uint8array' }), { saveAs: true, mime: 'application/zip' }); };
  const ocr = async () => { if (!need()) return; const b = busy('Đang nhận dạng chữ (Việt + Anh)…'); let text = ''; try { const rs = results(); for (let i = 0; i < rs.length; i++) { const r = await Host.ocr(applyFilter(rs[i], 'gray').toDataURL('image/png'), 'vie+eng'); text += (rs.length > 1 ? '--- Trang ' + (i + 1) + ' ---\n' : '') + r.text.trim() + '\n\n'; } } catch (e) { toast('OCR lỗi: ' + e.message, 'err'); } finally { b.done(); }
    if (!text) return; const ta = h('textarea.inp', { style: { width: '100%', height: '320px' } }, text);
    dialog({ title: 'Văn bản nhận dạng', width: 720, body: ta, buttons: [{ label: 'Mở trong Word', primary: true, run: () => { close(); shell.openFile({ name: 'Van_ban_quet.txt', bytes: new TextEncoder().encode(ta.value), here: true }); } }, { label: 'Dịch (A.I)', run: () => import('../ai/ai.js').then(A => A.openAi(shell, null, 'Dịch đoạn sau sang tiếng Việt (nếu đã là tiếng Việt thì dịch sang tiếng Anh):\n' + ta.value.slice(0, 6000))) }, { label: 'Sao chép', run: () => { ta.select(); try { document.execCommand('copy'); toast('Đã sao chép'); } catch (e) {} return false; } }, { label: 'Đóng' }] }); };
  const toPdfTool = async () => { if (!need()) return; const bytes = await pagesToPdf(results()); close(); shell.openFile({ name: 'Scan_' + Date.now() + '.pdf', bytes, here: true }); };
  document.body.appendChild(root); renderThumbs(); renderStage();
  return { root, addCanvas, pages, exportPdf, close, _results: results };
}
