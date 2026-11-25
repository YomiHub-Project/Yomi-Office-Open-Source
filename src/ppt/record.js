// Yomi-Office Open Source v0.0.6 — PowerPoint "Record" tab (Ghi): record narration + slide timings, Cameo (camera clip),
// screen recording, clear recordings, export the presentation to a video (.mp4 when the WebView supports it, otherwise .webm).
import { h, ic, uid } from '../core/util.js';
import { toast, confirmDlg, dialog, busy } from '../core/ui.js';
import { SW, SH } from './ppt.js';
import * as P from './cmds.js';

const pickMime = (kind) => { const c = kind === 'audio' ? ['audio/webm;codecs=opus', 'audio/webm', 'audio/mp4', 'audio/ogg'] : ['video/mp4;codecs=avc1', 'video/mp4', 'video/webm;codecs=vp9', 'video/webm;codecs=vp8', 'video/webm']; if (!window.MediaRecorder) return null; for (const m of c) if (MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(m)) return m; return ''; };
const blobToDataUrl = b => new Promise(res => { const r = new FileReader(); r.onload = () => res(r.result); r.readAsDataURL(b); });
export function supported() { return !!(window.MediaRecorder && navigator.mediaDevices); }

// ---------------------------------------------------------------- narration + timings
export async function recordShow(doc, from) {
  if (!supported()) return toast('Thiết bị/WebView này chưa hỗ trợ ghi âm (MediaRecorder).', 'err', 5000);
  let stream; try { stream = await navigator.mediaDevices.getUserMedia({ audio: true }); } catch (e) { return toast('Không truy cập được micro: ' + (e.message || e), 'err', 5000); }
  const mime = pickMime('audio');
  doc.pushHistory();
  let rec = null, chunks = [], t0 = 0, curIdx = -1;
  const startSeg = idx => { curIdx = idx; chunks = []; t0 = Date.now(); rec = new MediaRecorder(stream, mime ? { mimeType: mime } : {}); rec.ondataavailable = e => { if (e.data && e.data.size) chunks.push(e.data); }; rec.start(250); };
  const endSeg = () => new Promise(res => { if (!rec || rec.state === 'inactive') return res(); const idx = curIdx, dur = (Date.now() - t0) / 1000, local = chunks; rec.onstop = async () => { const s = doc.slides[idx]; if (s) { s.timing = Math.round(dur * 10) / 10; if (local.length) s.narration = await blobToDataUrl(new Blob(local, { type: mime || 'audio/webm' })); } res(); }; rec.stop(); });
  P.startShow(doc, from || 0);
  const indicator = h('div.pp-recdot', ic('circle', 12), ' Đang ghi… (Esc để kết thúc)'); doc._show.el.appendChild(indicator);
  startSeg(doc._show.idx);
  const poll = setInterval(async () => {
    if (!doc._show) { clearInterval(poll); await endSeg(); stream.getTracks().forEach(t => t.stop()); doc.useTimings = true; doc.markDirty(); doc.renderThumbs(); toast('Đã ghi lời thuyết minh và thời gian cho các slide', 'ok'); return; }
    if (doc._show.idx !== curIdx) { const next = doc._show.idx; await endSeg(); startSeg(next); }
  }, 120);
}
export function clearRecording(doc, all) {
  doc.change('Xóa bản ghi', () => (all ? doc.slides : [doc.slide]).forEach(s => { delete s.timing; delete s.narration; }));
  toast(all ? 'Đã xóa bản ghi trên mọi slide' : 'Đã xóa bản ghi trên slide hiện tại');
}
// play narration / auto-advance during a slide show (called from startShow render)
export function onShowSlide(doc, s, next) {
  if (doc._showAudio) { try { doc._showAudio.pause(); } catch (e) {} doc._showAudio = null; }
  clearTimeout(doc._showTimer);
  if (s.narration) { const a = new Audio(s.narration); a.play().catch(() => {}); doc._showAudio = a; }
  if (doc.useTimings !== false && s.timing && !doc._recording) doc._showTimer = setTimeout(() => { if (doc._show) next(); }, s.timing * 1000);
}

// ---------------------------------------------------------------- Cameo / screen recording → inserted video objects
async function recordStream(stream, label) {
  const mime = pickMime('video'); const chunks = []; const rec = new MediaRecorder(stream, mime ? { mimeType: mime } : {});
  rec.ondataavailable = e => { if (e.data && e.data.size) chunks.push(e.data); };
  const preview = h('video', { autoplay: true, muted: true, playsinline: true, style: { width: '100%', maxHeight: '50vh', background: '#000' } }); preview.srcObject = stream;
  rec.start(500);
  await dialog({ title: label, body: h('div', preview, h('div.muted', 'Đang ghi… bấm "Dừng" để chèn vào slide.')), buttons: [{ label: 'Dừng', primary: true }] });
  await new Promise(r => { rec.onstop = r; rec.stop(); }); stream.getTracks().forEach(t => t.stop());
  return blobToDataUrl(new Blob(chunks, { type: mime || 'video/webm' }));
}
export async function cameo(doc) { if (!supported()) return toast('Thiết bị chưa hỗ trợ camera/MediaRecorder', 'err'); let st; try { st = await navigator.mediaDevices.getUserMedia({ video: true, audio: true }); } catch (e) { return toast('Không mở được camera: ' + (e.message || e), 'err'); } const url = await recordStream(st, 'Cameo — quay camera'); P.addObj(doc, { type: 'video', src: url, name: 'Cameo', x: SW - 380, y: SH - 250, w: 340, h: 210 }); }
export async function screenRecording(doc) { if (!navigator.mediaDevices || !navigator.mediaDevices.getDisplayMedia) return toast('Quay màn hình cần Edge/Chrome 72+ trên máy tính', 'err', 5000); let st; try { st = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: true }); } catch (e) { return; } const url = await recordStream(st, 'Quay màn hình'); P.addObj(doc, { type: 'video', src: url, name: 'Quay màn hình', x: 140, y: 80, w: 1000, h: 560 }); }
export async function recordAudio(doc) { if (!supported()) return toast('Thiết bị chưa hỗ trợ ghi âm', 'err'); let st; try { st = await navigator.mediaDevices.getUserMedia({ audio: true }); } catch (e) { return toast('Không truy cập được micro', 'err'); } const mime = pickMime('audio'); const chunks = []; const rec = new MediaRecorder(st, mime ? { mimeType: mime } : {}); rec.ondataavailable = e => e.data.size && chunks.push(e.data); rec.start(250); await dialog({ title: 'Ghi âm thanh', body: h('div', ic('mic', 32), h('div', 'Đang ghi âm… bấm Dừng để chèn âm thanh vào slide.')), buttons: [{ label: 'Dừng', primary: true }] }); await new Promise(r => { rec.onstop = r; rec.stop(); }); st.getTracks().forEach(t => t.stop()); const url = await blobToDataUrl(new Blob(chunks, { type: mime || 'audio/webm' })); P.addObj(doc, { type: 'video', audio: true, src: url, name: 'Âm thanh', x: SW - 140, y: SH - 120, w: 100, h: 80 }); }

// ---------------------------------------------------------------- export to video
export async function exportVideo(doc, opts) {
  opts = opts || {};
  if (!window.MediaRecorder || !HTMLCanvasElement.prototype.captureStream) throw new Error('WebView này chưa hỗ trợ xuất video (cần Chrome/Edge 62+ hoặc Android System WebView mới)');
  const secs = opts.seconds || 5; const H = opts.height || 720; const W = Math.round(H * SW / SH);
  const cv = document.createElement('canvas'); cv.width = W; cv.height = H; const ctx = cv.getContext('2d');
  const vstream = cv.captureStream(30); let stream = vstream; let actx = null, dest = null;
  if (doc.slides.some(s => s.narration) && window.AudioContext) { actx = new AudioContext(); dest = actx.createMediaStreamDestination(); stream = new MediaStream(vstream.getVideoTracks().concat(dest.stream.getAudioTracks())); }
  const mime = pickMime('video'); const chunks = []; const rec = new MediaRecorder(stream, mime ? { mimeType: mime, videoBitsPerSecond: 4e6 } : {});
  rec.ondataavailable = e => { if (e.data && e.data.size) chunks.push(e.data); };
  const b = busy('Đang xuất video…');
  try {
    const imgs = []; for (let i = 0; i < doc.slides.length; i++) { const url = await P.slideToPng(doc, i, 'image/png', H / SH); const im = new Image(); await new Promise(r => { im.onload = r; im.onerror = r; im.src = url; }); imgs.push(im); }
    rec.start(200);
    for (let i = 0; i < imgs.length; i++) {
      const s = doc.slides[i]; const dur = (s.timing || secs) * 1000; const t0 = performance.now();
      if (s.narration && actx) { try { const buf = await fetch(s.narration).then(r => r.arrayBuffer()).then(a => actx.decodeAudioData(a)); const src = actx.createBufferSource(); src.buffer = buf; src.connect(dest); src.start(); } catch (e) {} }
      await new Promise(res => { const frame = () => { const t = performance.now() - t0; ctx.fillStyle = '#000'; ctx.fillRect(0, 0, W, H); const fade = s.transition && t < 500 ? t / 500 : 1; ctx.globalAlpha = fade; if (imgs[i].width) ctx.drawImage(imgs[i], 0, 0, W, H); ctx.globalAlpha = 1; if (t >= dur) return res(); requestAnimationFrame(frame); }; frame(); });
      b.update && b.update('Đang xuất video… ' + (i + 1) + '/' + imgs.length);
    }
    await new Promise(r => { rec.onstop = r; rec.stop(); });
  } finally { b.done(); if (actx) actx.close(); }
  const blob = new Blob(chunks, { type: mime || 'video/webm' });
  doc._lastVideoExt = /mp4/.test(mime || '') ? 'mp4' : 'webm';
  return new Uint8Array(await blob.arrayBuffer());
}
