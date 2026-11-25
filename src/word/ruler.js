// Yomi-Office Open Source v0.0.6 — Word rulers (horizontal + vertical), rewritten.
// v0.0.5 bug: the ruler was centred independently of the page (ignored the scrollbar, horizontal scroll and zoom origin),
// counted from the page edge instead of the left margin, stopped before the right margin and its indent markers were static.
// Now: the ruler is measured from the real page rectangle on every scroll/zoom/resize, 0 sits on the left margin like Word,
// grey = margins, white = text area, and the markers (first-line, hanging, left, right indent) and margin edges are draggable.
import { h, $$, clamp } from '../core/util.js';
import { settings } from '../core/store.js';

const PX_CM = 96 / 2.54, PX_IN = 96;
function unitPx() { const u = (settings().word && settings().word.unit) || 'cm'; return u === 'in' ? PX_IN : PX_CM; }
function blockOf(node, body) { while (node && node !== body && !(node.nodeType === 1 && /^(P|H[1-6]|LI|DIV|BLOCKQUOTE|PRE)$/.test(node.tagName) && node.parentNode && (node.parentNode === body || /^(UL|OL|BLOCKQUOTE|TD|TH|DIV)$/.test(node.parentNode.tagName)))) node = node.parentNode; return node && node !== body ? node : null; }
export function selectedBlocks(doc) {
  const s = window.getSelection(); if (!s.rangeCount) return [];
  const r = s.getRangeAt(0); if (!doc.body.contains(r.startContainer)) return doc._rulerBlocks || [];
  const a = blockOf(r.startContainer, doc.body), b = blockOf(r.endContainer, doc.body); if (!a) return [];
  const out = [a]; if (b && b !== a) { let n = a; const all = $$('p,h1,h2,h3,h4,h5,h6,li,blockquote,pre', doc.body); const i = all.indexOf(a), j = all.indexOf(b); if (i >= 0 && j > i) return all.slice(i, j + 1); }
  return out;
}

export class Ruler {
  constructor(doc) {
    this.doc = doc;
    this.el = h('div.wd-ruler');
    this.cv = h('canvas.wd-ruler-cv'); this.el.appendChild(this.cv);
    this.vel = h('div.wd-vruler'); this.vcv = h('canvas.wd-ruler-cv'); this.vel.appendChild(this.vcv);
    this.guide = h('div.wd-rguide.hidden');
    this.tip = h('div.wd-rtip.hidden');
    this.el.addEventListener('mousedown', e => this.down(e)); this.el.addEventListener('touchstart', e => this.down(e), { passive: false });
    this.el.addEventListener('mousemove', e => this.hover(e));
    this.el.addEventListener('dblclick', () => import('./cmds.js').then(m => m.pageSetupDialog && m.pageSetupDialog(doc)));
    this.vel.addEventListener('mousedown', e => this.vdown(e)); this.vel.addEventListener('touchstart', e => this.vdown(e), { passive: false });
    this.vel.addEventListener('mousemove', e => { const g = this.vgeom(); if (!g) return; const y = this.lp(e, this.vel).y; this.vel.style.cursor = (Math.abs(y - g.mt) < 4 || Math.abs(y - g.mb) < 4) ? 'ns-resize' : 'default'; });
    if (typeof ResizeObserver !== 'undefined') {
      this._ro = new ResizeObserver(() => this.draw());
      setTimeout(() => {
        try {
          if (doc.sheet) this._ro.observe(doc.sheet);
          if (doc.scroller) this._ro.observe(doc.scroller);
          if (doc.el) this._ro.observe(doc.el);
          if (this.el) this._ro.observe(this.el);
        } catch (e) {}
      }, 0);
    }
    if (typeof document !== 'undefined' && document.fonts && document.fonts.ready) {
      document.fonts.ready.then(() => this.draw());
    }
  }
  get hidden() { return this.el.classList.contains('hidden'); }
  toggle(v) {
    const on = v == null ? this.hidden : v;
    this.el.classList.toggle('hidden', !on);
    this.vel.classList.toggle('hidden', !on);
    this.el.style.paddingLeft = (on && !this.vel.classList.contains('hidden')) ? '20px' : '0px';
    this.draw();
  }
  lp(e, el) { const t = e.touches ? e.touches[0] : e; const r = (el === this.el ? this.cv : el).getBoundingClientRect(); return { x: t.clientX - r.left, y: t.clientY - r.top }; }
  // ---------------------------------------------------------------- geometry
  geom() {
    const d = this.doc; if (!d.sheet || !d.pageW) return null;
    const er = this.el.getBoundingClientRect(), sr = d.sheet.getBoundingClientRect(); if (!er.width || !sr.width) return null;
    const scrollerLeft = d.scroller ? d.scroller.getBoundingClientRect().left : er.left;
    const wantPadL = Math.max(0, Math.round(scrollerLeft - er.left));
    if (this.el.style.paddingLeft !== wantPadL + 'px') this.el.style.paddingLeft = wantPadL + 'px';
    const padL = wantPadL || parseFloat(getComputedStyle(this.el).paddingLeft) || 0;
    const rr = { left: er.left + padL, width: Math.max(10, er.width - padL) };
    const cvr = this.cv.getBoundingClientRect();
    const cvLeft = (cvr.width > 0 && cvr.left > 0) ? cvr.left : rr.left;
    const k = d.zoom / 100, m = d.marg;
    const x0 = sr.left - cvLeft;
    const pw = sr.width || (d.pageW * k);
    const ml = x0 + m[3] * k, mr = x0 + pw - m[1] * k;
    const blk = this.blocks()[0]; let li = 0, ri = 0, fi = 0;
    if (blk) { const cs = getComputedStyle(blk); const extra = blk.tagName === 'LI' ? (blk.parentNode.offsetLeft - d.body.offsetLeft) : 0; li = (parseFloat(cs.marginLeft) || 0) + (parseFloat(cs.paddingLeft) || 0) + Math.max(0, extra - m[3]); ri = (parseFloat(cs.marginRight) || 0) + (parseFloat(cs.paddingRight) || 0); fi = parseFloat(cs.textIndent) || 0; }
    return { k, x0, pw, ml, mr, li: ml + li * k, ri: mr - ri * k, fi: ml + (li + fi) * k, rawLi: li, rawRi: ri, rawFi: fi, W: Math.max(rr.width, Math.round(x0 + pw + 80)), blk };
  }
  vgeom() {
    const d = this.doc; if (!d.sheet || !d.pageH || this.vel.classList.contains('hidden')) return null;
    const vr = this.vel.getBoundingClientRect(), sr = d.sheet.getBoundingClientRect(); if (!vr.height) return null;
    const k = d.zoom / 100, m = d.marg, ph = d.pageH * k, gap = 24 * k;
    // the page that contains the vertical centre of the viewport (or the caret page)
    const rel = vr.height / 2 - (sr.top - vr.top); const pg = clamp(Math.floor(rel / (ph + gap)), 0, Math.max(0, (d.pageCount || 1) - 1));
    const y0 = sr.top - vr.top + pg * (ph + gap);
    return { k, y0, ph, mt: y0 + m[0] * k, mb: y0 + ph - m[2] * k, H: vr.height };
  }
  blocks() { const b = selectedBlocks(this.doc); if (b.length) this.doc._rulerBlocks = b; return b.length ? b : (this.doc._rulerBlocks || []).filter(x => this.doc.body.contains(x)); }
  // ---------------------------------------------------------------- drawing
  draw() {
    if (this.hidden) return;
    const g = this.geom(); if (!g) return;
    const dpr = window.devicePixelRatio || 1, H = 22; const cv = this.cv;
    if (cv.width !== Math.round(g.W * dpr) || cv.height !== H * dpr) { cv.width = Math.round(g.W * dpr); cv.height = H * dpr; cv.style.width = g.W + 'px'; cv.style.height = H + 'px'; }
    const c = cv.getContext('2d'); c.setTransform(dpr, 0, 0, dpr, 0, 0); c.clearRect(0, 0, g.W, H);
    const top = 4, bh = 14, dark = document.body.classList.contains('dark');
    c.fillStyle = dark ? '#3a3a3a' : '#c8cfdb'; c.fillRect(g.x0, top, g.pw, bh);                 // margins (grey)
    c.fillStyle = dark ? '#1f1f1f' : '#fff'; c.fillRect(g.ml, top, g.mr - g.ml, bh);             // text area (white)
    c.strokeStyle = dark ? '#555' : '#a9b1bf'; c.lineWidth = 1; c.strokeRect(Math.round(g.x0) + .5, top + .5, Math.round(g.pw) - 1, bh - 1);
    const u = unitPx() * g.k; const step = u >= 30 ? 0.25 : u >= 14 ? 0.5 : 1;
    c.fillStyle = dark ? '#bbb' : '#444'; c.font = '9px Segoe UI, Roboto, Arial, sans-serif'; c.textAlign = 'center'; c.textBaseline = 'middle';
    const tick = (x, v) => {
      if (x < g.x0 - 0.5 || x > g.x0 + g.pw + 0.5) return;
      const whole = Math.abs(v - Math.round(v)) < 1e-6, half = Math.abs(v * 2 - Math.round(v * 2)) < 1e-6;
      if (whole) {
        if (v !== 0) {
          const lab = String(Math.abs(Math.round(v)));
          if (u >= 10 || Math.round(v) % 2 === 0) c.fillText(lab, x, top + bh / 2 + .5);
        } else {
          c.fillRect(Math.round(x), top + 2, 1, bh - 4);
        }
      } else if (!whole) {
        const len = half ? 6 : 3;
        c.fillRect(Math.round(x), top + (bh - len) / 2, 1, len);
      }
    };
    for (let v = 0; g.ml + v * u <= g.x0 + g.pw + 1; v += step) tick(g.ml + v * u, v);
    for (let v = step; g.ml - v * u >= g.x0 - 1; v += step) tick(g.ml - v * u, v);
    // indent markers
    const mk = (x, y, up, col) => { c.beginPath(); if (up) { c.moveTo(x - 5, y); c.lineTo(x + 5, y); c.lineTo(x, y - 5); } else { c.moveTo(x - 5, y); c.lineTo(x + 5, y); c.lineTo(x, y + 5); } c.closePath(); c.fillStyle = col; c.fill(); c.strokeStyle = '#5b6b8c'; c.stroke(); };
    const fill = dark ? '#ddd' : '#fff';
    mk(g.fi, top - 1, false, fill);                       // first-line indent (top, pointing down)
    mk(g.li, top + bh - 4, true, fill);                   // hanging indent (pointing up)
    c.fillStyle = fill; c.fillRect(g.li - 5, top + bh - 4, 10, 4); c.strokeStyle = '#5b6b8c'; c.strokeRect(g.li - 5 + .5, top + bh - 4 + .5, 9, 4); // left indent box
    mk(g.ri, top + bh, true, fill);                       // right indent
    this._g = g;
    this.drawV();
  }
  drawV() {
    const g = this.vgeom(); if (!g) return; const dpr = window.devicePixelRatio || 1, W = 20, cv = this.vcv;
    if (cv.width !== W * dpr || cv.height !== Math.round(g.H * dpr)) { cv.width = W * dpr; cv.height = Math.round(g.H * dpr); cv.style.width = W + 'px'; cv.style.height = g.H + 'px'; }
    const c = cv.getContext('2d'); c.setTransform(dpr, 0, 0, dpr, 0, 0); c.clearRect(0, 0, W, g.H);
    const dark = document.body.classList.contains('dark');
    c.fillStyle = dark ? '#3a3a3a' : '#c8cfdb'; c.fillRect(3, g.y0, 14, g.ph);
    c.fillStyle = dark ? '#1f1f1f' : '#fff'; c.fillRect(3, g.mt, 14, g.mb - g.mt);
    const u = unitPx() * g.k; const step = u >= 30 ? 0.25 : u >= 14 ? 0.5 : 1;
    c.fillStyle = dark ? '#bbb' : '#444'; c.font = '9px Segoe UI, Roboto, Arial, sans-serif'; c.textAlign = 'center'; c.textBaseline = 'middle';
    const tick = (y, v) => { if (y < g.y0 || y > g.y0 + g.ph) return; const whole = Math.abs(v - Math.round(v)) < 1e-6, half = Math.abs(v * 2 - Math.round(v * 2)) < 1e-6; if (whole && v !== 0) { c.save(); c.translate(10, y); c.rotate(-Math.PI / 2); c.fillText(String(Math.abs(Math.round(v))), 0, 0); c.restore(); } else if (!whole) { const len = half ? 6 : 3; c.fillRect(10 - len / 2, Math.round(y), len, 1); } };
    for (let v = 0; g.mt + v * u <= g.y0 + g.ph; v += step) tick(g.mt + v * u, v);
    for (let v = step; g.mt - v * u >= g.y0; v += step) tick(g.mt - v * u, v);
    this._vg = g;
  }
  // ---------------------------------------------------------------- interaction
  hit(x, y) {
    const g = this._g || this.geom(); if (!g) return null; const near = (a, t) => Math.abs(x - a) <= (t || 5);
    // vertical zones (canvas 22px, bar 4..18): top triangle = first-line, middle = margin edge, lower triangle = hanging, bottom box = left indent
    if (y <= 8 && near(g.fi)) return 'first';
    if (y >= 13 && y < 17 && near(g.li)) return 'hang';
    if (y >= 17 && near(g.li)) return 'left';
    if (y >= 12 && near(g.ri)) return 'right';
    if (y > 4 && y < 18 && near(g.ml, 4)) return 'mleft';
    if (y > 4 && y < 18 && near(g.mr, 4)) return 'mright';
    if (near(g.fi)) return y < 11 ? 'first' : 'hang';
    return null;
  }
  hover(e) { if (this._drag) return; const p = this.lp(e, this.el); const t = this.hit(p.x, p.y); const tips = { first: 'Thụt lề dòng đầu', hang: 'Thụt lề treo', left: 'Thụt lề trái', right: 'Thụt lề phải', mleft: 'Lề trái', mright: 'Lề phải' }; this.el.style.cursor = t ? (t[0] === 'm' ? 'ew-resize' : 'pointer') : 'default'; this.el.title = t ? tips[t] : 'Bấm đúp: Thiết lập trang'; }
  showGuide(x) { const d = this.doc; if (!this.guide.parentNode) d.el.appendChild(this.guide); const rr = this.cv.getBoundingClientRect(), er = d.el.getBoundingClientRect(); this.guide.style.left = (rr.left - er.left + x) + 'px'; this.guide.style.top = (rr.bottom - er.top) + 'px'; this.guide.classList.remove('hidden'); }
  showTip(x, txt) { const d = this.doc; if (!this.tip.parentNode) d.el.appendChild(this.tip); const rr = this.cv.getBoundingClientRect(), er = d.el.getBoundingClientRect(); this.tip.textContent = txt; this.tip.style.left = (rr.left - er.left + x + 8) + 'px'; this.tip.style.top = (rr.bottom - er.top + 4) + 'px'; this.tip.classList.remove('hidden'); }
  fmt(px) { const u = unitPx(); const v = px / u; return (Math.round(v * 100) / 100).toLocaleString('vi-VN') + (u === PX_IN ? ' in' : ' cm'); }
  down(e) {
    if (e.button === 2) return; const p = this.lp(e, this.el); const what = this.hit(p.x, p.y); if (!what) return;
    e.preventDefault(); const d = this.doc; const g = this.geom(); if (!g) return;
    const blocks = what[0] === 'm' ? [] : this.blocks(); if (what[0] !== 'm' && !blocks.length) return;
    const touch = e.type === 'touchstart'; const before = d.body.innerHTML; const m0 = d.marg.slice();
    const snapU = unitPx() / 4; const snap = (v, ev) => ev && ev.altKey ? v : Math.round(v / snapU) * snapU;
    this._drag = true;
    const mv = ev => {
      const q = this.lp(ev, this.el); const dx = (q.x - p.x) / g.k; let x = q.x;
      if (what === 'mleft' || what === 'mright') {
        const m = m0.slice(); if (what === 'mleft') m[3] = clamp(snap(m0[3] + dx, ev), 0, d.pageW - m0[1] - 48); else m[1] = clamp(snap(m0[1] - dx, ev), 0, d.pageW - m0[3] - 48);
        d.setMargins(m, true); this.showTip(x, what === 'mleft' ? 'Lề trái: ' + this.fmt(m[3]) : 'Lề phải: ' + this.fmt(m[1]));
      } else {
        blocks.forEach(b => {
          const cs = getComputedStyle(b); const li = parseFloat(b.style.marginLeft) || 0, fi = parseFloat(cs.textIndent) || 0, ri = parseFloat(b.style.marginRight) || 0;
          if (!b._r0) b._r0 = { li, fi, ri };
          const r0 = b._r0;
          if (what === 'first') { b.style.textIndent = snap(r0.fi + dx, ev) + 'px'; }
          else if (what === 'hang') { const nli = Math.max(0, snap(r0.li + dx, ev)); b.style.marginLeft = nli + 'px'; b.style.textIndent = (r0.fi - (nli - r0.li)) + 'px'; }
          else if (what === 'left') { b.style.marginLeft = Math.max(0, snap(r0.li + dx, ev)) + 'px'; }
          else if (what === 'right') { b.style.marginRight = Math.max(0, snap(r0.ri - dx, ev)) + 'px'; }
        });
        const b = blocks[0]; const cs = getComputedStyle(b);
        this.showTip(x, what === 'first' ? 'Dòng đầu: ' + this.fmt(parseFloat(cs.textIndent) || 0) : what === 'right' ? 'Thụt phải: ' + this.fmt(parseFloat(cs.marginRight) || 0) : 'Thụt trái: ' + this.fmt(parseFloat(cs.marginLeft) || 0));
      }
      this.showGuide(x); this.draw(); if (ev.cancelable) ev.preventDefault();
    };
    const up = () => {
      document.removeEventListener(touch ? 'touchmove' : 'mousemove', mv); document.removeEventListener(touch ? 'touchend' : 'mouseup', up);
      this._drag = false; this.guide.classList.add('hidden'); this.tip.classList.add('hidden');
      blocks.forEach(b => { delete b._r0; });
      if (what[0] === 'm') { d.setMargins(d.marg.slice()); }
      else if (d.body.innerHTML !== before) { d.pushHistory(); d.markDirty(); d.debouncedPages && d.debouncedPages(); }
      this.draw();
    };
    document.addEventListener(touch ? 'touchmove' : 'mousemove', mv, touch ? { passive: false } : false); document.addEventListener(touch ? 'touchend' : 'mouseup', up);
  }
  vdown(e) {
    const g = this.vgeom(); if (!g) return; const p = this.lp(e, this.vel); const d = this.doc;
    const which = Math.abs(p.y - g.mt) < 5 ? 'top' : Math.abs(p.y - g.mb) < 5 ? 'bottom' : null; if (!which) return;
    e.preventDefault(); const touch = e.type === 'touchstart'; const m0 = d.marg.slice(); const snapU = unitPx() / 4;
    const mv = ev => { const q = this.lp(ev, this.vel); const dy = (q.y - p.y) / g.k; const m = m0.slice(); if (which === 'top') m[0] = clamp(Math.round((m0[0] + dy) / snapU) * snapU, 0, d.pageH / 2); else m[2] = clamp(Math.round((m0[2] - dy) / snapU) * snapU, 0, d.pageH / 2); d.setMargins(m, true); this.draw(); if (ev.cancelable) ev.preventDefault(); };
    const up = () => { document.removeEventListener(touch ? 'touchmove' : 'mousemove', mv); document.removeEventListener(touch ? 'touchend' : 'mouseup', up); d.setMargins(d.marg.slice()); };
    document.addEventListener(touch ? 'touchmove' : 'mousemove', mv, touch ? { passive: false } : false); document.addEventListener(touch ? 'touchend' : 'mouseup', up);
  }
}
