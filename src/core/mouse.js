// Yomi-Office Open Source v0.0.6 — extra mouse behaviours from Word/Excel/PowerPoint (Windows Office):
// middle-button autoscroll (pan), click highlight ring, Word "Click and Type", horizontal Shift+wheel, thicker caret.
import { h } from './util.js';

let installed = false; let opts = {};
export function installMouse(m) {
  opts = m || {};
  if (installed) return; installed = true;
  // ---- middle-click autoscroll (Windows Office behaviour): click the wheel, move the mouse to scroll, click again to stop
  let auto = null;
  const scrollerOf = el => { while (el && el !== document.body) { const cs = getComputedStyle(el); if (/(auto|scroll)/.test(cs.overflowY + cs.overflowX) && (el.scrollHeight > el.clientHeight + 2 || el.scrollWidth > el.clientWidth + 2)) return el; el = el.parentElement; } return null; };
  const stop = () => { if (!auto) return; cancelAnimationFrame(auto.raf); auto.mark.remove(); document.removeEventListener('mousemove', auto.mv, true); auto = null; };
  document.addEventListener('mousedown', e => {
    if (auto) { e.preventDefault(); stop(); return; }
    if (e.button !== 1 || opts.middleAutoscroll === false) return;
    if (e.target.closest && e.target.closest('a,[href]')) return;
    const sc = scrollerOf(e.target); if (!sc) return;
    e.preventDefault();
    const mark = h('div.oos-autoscroll'); mark.style.left = (e.clientX - 14) + 'px'; mark.style.top = (e.clientY - 14) + 'px'; document.body.appendChild(mark);
    const st = { x0: e.clientX, y0: e.clientY, x: e.clientX, y: e.clientY, sc, mark };
    st.mv = ev => { st.x = ev.clientX; st.y = ev.clientY; };
    const step = () => { const dx = st.x - st.x0, dy = st.y - st.y0; const f = v => Math.abs(v) < 12 ? 0 : Math.sign(v) * Math.pow((Math.abs(v) - 12) / 6, 1.4); st.sc.scrollLeft += f(dx); st.sc.scrollTop += f(dy); st.raf = requestAnimationFrame(step); };
    document.addEventListener('mousemove', st.mv, true); auto = st; st.raf = requestAnimationFrame(step);
  }, true);
  document.addEventListener('keydown', e => { if (auto && e.key === 'Escape') stop(); }, true);
  document.addEventListener('auxclick', e => { if (e.button === 1 && auto) e.preventDefault(); }, true);
  // ---- click highlight ring
  document.addEventListener('mousedown', e => {
    if (!opts.clickHighlight) return;
    const r = h('div.oos-click' + (e.button === 2 ? '.right' : '')); r.style.left = (e.clientX - 18) + 'px'; r.style.top = (e.clientY - 18) + 'px'; document.body.appendChild(r);
    setTimeout(() => r.remove(), 450);
  }, true);
}
// Word "Click and Type": double-click below the last paragraph adds empty paragraphs down to that point and puts the caret there
export function clickAndType(doc, e) {
  if (opts.clickAndType === false || doc.readOnly) return false;
  const body = doc.body; const last = body.lastElementChild; if (!last) return false;
  const lr = last.getBoundingClientRect(); if (e.clientY <= lr.bottom + 4) return false;
  const k = doc.zoom / 100; const lineH = 18 * k; const n = Math.min(80, Math.ceil((e.clientY - lr.bottom) / lineH));
  doc.exec('Bấm-và-Gõ', () => {
    let p = null; for (let i = 0; i < n; i++) { p = document.createElement('p'); p.innerHTML = '<br>'; body.appendChild(p); }
    const br = body.getBoundingClientRect(); const relX = (e.clientX - br.left) / k - (doc.marg ? doc.marg[3] : 96);
    const w = (doc.pageW || 794) - (doc.marg ? doc.marg[1] + doc.marg[3] : 192);
    if (p && relX > w * 0.35 && relX < w * 0.65) p.style.textAlign = 'center'; else if (p && relX >= w * 0.65) p.style.textAlign = 'right'; else if (p && relX > 20) p.style.marginLeft = Math.round(relX) + 'px';
    const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
  });
  return true;
}
