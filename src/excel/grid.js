import { drawBreaks, hitPageBreak } from './pages.js';
// Canvas spreadsheet grid: virtualized rendering, frozen panes, selection, mouse/touch/keyboard interactions, in-cell editor
import { h, ic, keyName, chord, clamp, throttleRaf, $ } from '../core/util.js';
import { contextMenu, longPress, showMenu, toast } from '../core/ui.js';
import { settings } from '../core/store.js';
import { MAX_ROWS, MAX_COLS, colName, key, norm, inRange, rangeAddr, Err } from './model.js';
import { format } from './numfmt.js';
import { formulaRefs, cycleAbs } from './formula.js';
import { FN } from './formula.js';
import { Axis } from './axis.js';

const HEAD_H = 20;
const SEL_COLOR = '#107c41';
const REF_COLORS = ['#4472c4', '#c00000', '#7030a0', '#00b050', '#ed7d31', '#00b0f0', '#bf8f00'];
export const PT2PX = 96 / 72;

export class Grid {
  constructor(doc) {
    this.doc = doc;
    this.el = h('div.xl-wrap');
    this.canvas = h('canvas.xl-canvas');
    this.scroller = h('div.xl-scroll', { tabindex: -1 });
    this.spacer = h('div.xl-spacer');
    this.scroller.appendChild(this.spacer);
    this.objLayer = h('div.xl-objs');
    this.editor = h('textarea.xl-editor', { spellcheck: false, autocapitalize: 'off', autocomplete: 'off', autocorrect: 'off', wrap: 'off' });
    this.acBox = h('div.xl-ac.hidden');
    this.tipBox = h('div.xl-fntip.hidden');
    this.el.appendChild(this.canvas); this.el.appendChild(this.scroller); this.el.appendChild(this.objLayer); this.el.appendChild(this.editor); this.el.appendChild(this.acBox); this.el.appendChild(this.tipBox);
    this.ctx = this.canvas.getContext('2d');
    this.editing = null; // {r,c, mode:'enter'|'edit', orig}
    this.copyRange = null; this.dragState = null; this.pointRef = null;
    this.measureCache = {};
    this.redraw = throttleRaf(() => this.draw());
    this.bind();
  }
  get sheet() { return this.doc.wb.cur; }
  get z() { return (this.sheet.zoom || 100) / 100; }
  axes() { const s = this.sheet; if (!s._ax) s._ax = { r: new Axis(s, true), c: new Axis(s, false) }; return s._ax; }
  invalidateAxes() { const a = this.axes(); a.r.dirty = true; a.c.dirty = true; }
  headW() { const s = this.sheet; if (!s.showHead) return 0; const vis = this.visibleRange(); const digits = String((vis ? vis.r2 : 1000) + 1).length; return Math.max(34, Math.round((10 + digits * 7.5) * this.z)); }
  headH() { return this.sheet.showHead ? Math.round(HEAD_H * this.z) : 0; }
  colX(c) { return this.axes().c.pos(c) * this.z; }
  rowY(r) { return this.axes().r.pos(r) * this.z; }
  colW(c) { return this.sheet.colWidth(c) * this.z; }
  rowH(r) { return this.sheet.rowHeight(r) * this.z; }
  frozenW() { const f = this.sheet.freeze; return f.c ? this.colX(f.c) : 0; }
  frozenH() { const f = this.sheet.freeze; return f.r ? this.rowY(f.r) : 0; }

  // ---------------------------------------------------------------- layout
  resize() {
    const r = this.el.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    this.W = Math.max(10, r.width); this.H = Math.max(10, r.height);
    this.canvas.width = Math.round(this.W * dpr); this.canvas.height = Math.round(this.H * dpr);
    this.canvas.style.width = this.W + 'px'; this.canvas.style.height = this.H + 'px';
    this.dpr = dpr;
    this.updateSpacer();
    this.draw();
  }
  updateSpacer() {
    const u = this.sheet.used();
    const sc = this.scroller;
    const needR = Math.min(MAX_ROWS, Math.max(u.r2 + 60, this.rowAtY(sc.scrollTop + this.H * 2) + 40, 200));
    const needC = Math.min(MAX_COLS, Math.max(u.c2 + 20, this.colAtX(sc.scrollLeft + this.W * 2) + 10, 40));
    const w = this.colX(needC) - this.frozenW() + this.headW() + this.frozenW();
    const hgt = this.rowY(needR) - this.frozenH() + this.headH() + this.frozenH();
    this.spacer.style.width = Math.round(w) + 'px'; this.spacer.style.height = Math.round(hgt) + 'px';
  }
  rowAtY(y) { return this.axes().r.at(y / this.z, MAX_ROWS); }
  colAtX(x) { return this.axes().c.at(x / this.z, MAX_COLS); }
  // viewport point -> {r, c, zone}
  hit(vx, vy) {
    const hw = this.headW(), hh = this.headH(), fw = this.frozenW(), fh = this.frozenH();
    const sx = this.scroller.scrollLeft, sy = this.scroller.scrollTop;
    const zone = vx < hw && vy < hh ? 'corner' : vy < hh ? 'colhead' : vx < hw ? 'rowhead' : 'cell';
    const cx = vx - hw < fw ? vx - hw : vx - hw + sx;   // sheet x (zoomed)
    const cy = vy - hh < fh ? vy - hh : vy - hh + sy;
    const c = this.colAtX(Math.max(0, cx)), r = this.rowAtY(Math.max(0, cy));
    return { r, c, zone, cx, cy };
  }
  // sheet (zoomed) coords of a cell -> viewport rect
  cellRect(r, c, r2, c2) {
    const hw = this.headW(), hh = this.headH(), fr = this.sheet.freeze;
    const sx = this.scroller.scrollLeft, sy = this.scroller.scrollTop;
    const x1 = this.colX(c), y1 = this.rowY(r);
    const x2 = this.colX((c2 == null ? c : c2) + 1), y2 = this.rowY((r2 == null ? r : r2) + 1);
    const vx = hw + x1 - (c >= fr.c ? sx : 0), vy = hh + y1 - (r >= fr.r ? sy : 0);
    return { x: vx, y: vy, w: x2 - x1, h: y2 - y1 };
  }
  visibleRange() {
    if (!this.W) return null;
    const sx = this.scroller.scrollLeft, sy = this.scroller.scrollTop, fr = this.sheet.freeze;
    const fw = this.frozenW(), fh = this.frozenH();
    const r1 = this.rowAtY(fh + sy), c1 = this.colAtX(fw + sx);
    const r2 = Math.min(MAX_ROWS - 1, this.rowAtY(fh + sy + this.H)), c2 = Math.min(MAX_COLS - 1, this.colAtX(fw + sx + this.W));
    return { r1: Math.max(r1, fr.r), c1: Math.max(c1, fr.c), r2, c2 };
  }
  scrollTo(r, c) {
    const sc = this.scroller, fr = this.sheet.freeze;
    const hw = this.headW(), hh = this.headH(), fw = this.frozenW(), fh = this.frozenH();
    if (c >= fr.c) {
      const x1 = this.colX(c) - fw, x2 = this.colX(c + 1) - fw; const vw = this.W - hw - fw - 14;
      if (x1 < sc.scrollLeft) sc.scrollLeft = x1; else if (x2 > sc.scrollLeft + vw) sc.scrollLeft = x2 - vw;
    }
    if (r >= fr.r) {
      const y1 = this.rowY(r) - fh, y2 = this.rowY(r + 1) - fh; const vh = this.H - hh - fh - 14;
      if (y1 < sc.scrollTop) sc.scrollTop = y1; else if (y2 > sc.scrollTop + vh) sc.scrollTop = y2 - vh;
    }
    this.updateSpacer();
  }

  // ---------------------------------------------------------------- rendering
  font(st) {
    const z = this.z; const sz = ((st && st.sz) || settings().excelSize || 11) * PT2PX * z;
    const fam = (st && st.font) || settings().excelFont || 'Calibri';
    return (st && st.i ? 'italic ' : '') + (st && st.b ? 'bold ' : '') + sz.toFixed(1) + 'px "' + fam + '", Calibri, Carlito, Arial, sans-serif';
  }
  measure(text, fontStr) {
    const k = fontStr + '\u0001' + text; let w = this.measureCache[k];
    if (w === undefined) { this.ctx.font = fontStr; w = this.ctx.measureText(text).width; if (Object.keys(this.measureCache).length > 4000) this.measureCache = {}; this.measureCache[k] = w; }
    return w;
  }
  display(r, c) {
    // returns {text, color, align, st, v}
    const s = this.sheet; const cell = s.get(r, c);
    let v, st;
    if (cell) { v = cell.v; st = cell.s || s.rowS[r] || s.colS[c] || null; }
    else { const sp = s.spill.get(key(r, c)); v = sp ? sp.v : null; st = s.rowS[r] || s.colS[c] || null; }
    if (s.showFormulas && cell && cell.f) return { text: '=' + cell.f, color: null, align: 'left', st, v };
    if (v == null || v === '') return { text: '', st, v };
    const f = format(v, st && st.nf);
    const cf = this.doc.cfStyle ? this.doc.cfStyle(s, r, c, v) : null;
    let align = st && st.ha && st.ha !== 'general' ? st.ha : (typeof v === 'number' ? 'right' : (typeof v === 'boolean' || v instanceof Err) ? 'center' : 'left');
    return { text: f.t, color: (cf && cf.color) || f.c || (st && st.color) || null, align, st, v, cf };
  }
  draw() {
    if (!this.W) return;
    const ctx = this.ctx, s = this.sheet, z = this.z;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, this.W, this.H);
    const hw = this.headW(), hh = this.headH(), fw = this.frozenW(), fh = this.frozenH(), fr = s.freeze;
    const sx = this.scroller.scrollLeft, sy = this.scroller.scrollTop;
    const vis = this.visibleRange();
    // quadrants: [rowRange, colRange, offsetX, offsetY, clip]
    const quads = [];
    const mainR = { r1: vis.r1, r2: vis.r2 }, mainC = { c1: vis.c1, c2: vis.c2 };
    quads.push([mainR, mainC, hw - sx, hh - sy, [hw + fw, hh + fh, this.W, this.H]]);
    if (fr.r) quads.push([{ r1: 0, r2: fr.r - 1 }, mainC, hw - sx, hh, [hw + fw, hh, this.W, hh + fh]]);
    if (fr.c) quads.push([mainR, { c1: 0, c2: fr.c - 1 }, hw, hh - sy, [hw, hh + fh, hw + fw, this.H]]);
    if (fr.r && fr.c) quads.push([{ r1: 0, r2: fr.r - 1 }, { c1: 0, c2: fr.c - 1 }, hw, hh, [hw, hh, hw + fw, hh + fh]]);
    this._quads = quads;
    for (const q of quads) this.drawQuad(q[0], q[1], q[2], q[3], q[4]);
    // frozen lines
    if (fr.r) { ctx.fillStyle = '#9b9b9b'; ctx.fillRect(hw, hh + fh - 0.5, this.W, 1); }
    if (fr.c) { ctx.fillStyle = '#9b9b9b'; ctx.fillRect(hw + fw - 0.5, hh, 1, this.H); }
    this.drawHeaders(vis, hw, hh, fw, fh, sx, sy);
    this.positionObjects();
    if (this.doc.afterDraw) this.doc.afterDraw();
  }
  drawQuad(R, C, ox, oy, clip) {
    const ctx = this.ctx, s = this.sheet, z = this.z;
    if (R.r2 < R.r1 || C.c2 < C.c1) return;
    ctx.save();
    ctx.beginPath(); ctx.rect(clip[0], clip[1], clip[2] - clip[0], clip[3] - clip[1]); ctx.clip();
    const X = c => ox + this.colX(c), Y = r => oy + this.rowY(r);
    // extend col range left for text overflow
    const c0 = Math.max(0, C.c1 - 8);
    // 1) gridlines
    if (s.showGrid) {
      ctx.fillStyle = '#e1e1e1';
      for (let c = C.c1; c <= C.c2 + 1; c++) { const x = Math.round(X(c)) - 0.5; if (s.colWidth(c - 1) === 0 && c > 0) continue; ctx.fillRect(x, clip[1], 1, clip[3] - clip[1]); }
      for (let r = R.r1; r <= R.r2 + 1; r++) { const y = Math.round(Y(r)) - 0.5; ctx.fillRect(clip[0], y, clip[2] - clip[0], 1); }
    }
    // 2) fills (cell styles, CF), and collect text jobs
    const texts = []; const borders = [];
    const mergesHere = s.merges.filter(m => m.r2 >= R.r1 && m.r1 <= R.r2 && m.c2 >= c0 && m.c1 <= C.c2);
    const inMerge = (r, c) => { for (const m of mergesHere) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m; return null; };
    for (let r = R.r1; r <= R.r2; r++) {
      const rh = this.rowH(r); if (!rh) continue;
      const y = Y(r);
      for (let c = c0; c <= C.c2; c++) {
        const cw = this.colW(c); if (!cw) continue;
        const m = mergesHere.length ? inMerge(r, c) : null;
        if (m && (m.r1 !== r || m.c1 !== c)) { if (!(m.r1 < R.r1 && r === R.r1 && c === m.c1) && !(m.c1 < c0 && c === c0 && r === m.r1)) continue; }
        const rr = m ? m.r1 : r, cc = m ? m.c1 : c;
        const d = this.display(rr, cc);
        const x = X(cc), yy = Y(rr);
        const w = m ? X(m.c2 + 1) - x : cw, hgt = m ? Y(m.r2 + 1) - yy : rh;
        const fill = (d.cf && d.cf.fill) || (d.st && d.st.fill);
        if (c >= C.c1 || m) {
          if (fill && fill !== 'transparent') { ctx.fillStyle = fill; ctx.fillRect(x, yy, w, hgt); }
          else if (m) { ctx.fillStyle = '#fff'; ctx.fillRect(x + 0.5, yy + 0.5, w - 1, hgt - 1); }
          if (d.cf && d.cf.bar) this.drawDataBar(d.cf.bar, x, yy, w, hgt);
          if (d.st && d.st.bd) borders.push([d.st.bd, x, yy, w, hgt]);
        }
        if (d.text !== '') texts.push({ r: rr, c: cc, x, y: yy, w, h: hgt, d, merged: !!m });
        if (m) c = m.c2;
      }
    }
    // 3) text
    for (const t of texts) this.drawText(t, X, R, C);
    // 4) borders
    ctx.lineCap = 'square';
    for (const b of borders) this.drawBorders(b[0], b[1], b[2], b[3], b[4]);
    // 5) indicators: comments, filter buttons, validation arrows, spill outline
    this.drawIndicators(R, C, X, Y);
    // 6) selection
    this.drawSelection(X, Y, R, C);
    ctx.restore();
  }
  drawText(t, X, R, C) {
    const ctx = this.ctx, s = this.sheet, d = t.d, st = d.st || {}, z = this.z;
    const fontStr = this.font(st);
    ctx.font = fontStr;
    ctx.fillStyle = d.color || (d.v instanceof Err ? '#000' : '#000');
    const pad = 3 * z, indent = (st.indent || 0) * 9 * z;
    const lineH = ((st.sz || settings().excelSize || 11) * PT2PX * 1.22) * z;
    let lines;
    const wrap = st.wrap;
    if (wrap) lines = this.wrapLines(d.text, fontStr, t.w - pad * 2 - indent);
    else lines = String(d.text).split('\n').slice(0, 1);
    if (!wrap && !t.merged && typeof d.v === 'number' && this.measure(lines[0], fontStr) > t.w - pad) {
      // numbers that don't fit show ####
      const g = format(d.v, 'General').t;
      lines = [!st.nf || st.nf === 'General' ? (this.measure(g, fontStr) <= t.w - pad ? g : '#'.repeat(Math.max(1, Math.floor((t.w - pad) / this.measure('#', fontStr))))) : '#'.repeat(Math.max(1, Math.floor((t.w - pad) / this.measure('#', fontStr))))];
    }
    const totalH = lines.length * lineH;
    const va = st.va || 'bottom';
    let y = va === 'top' ? t.y + 2 * z : va === 'middle' || va === 'center' ? t.y + (t.h - totalH) / 2 : t.y + t.h - totalH - 2 * z;
    // clipping / overflow
    let clipX1 = t.x, clipX2 = t.x + t.w;
    const tw = this.measure(lines[0], fontStr);
    if (!wrap && !t.merged && typeof d.v === 'string' && tw > t.w - pad * 2) {
      if (d.align === 'left' || d.align === 'general') { let c = t.c + 1; while (c <= C.c2 + 20 && clipX2 - t.x < tw + pad * 2) { if (this.cellHasContent(t.r, c)) break; clipX2 += this.colW(c); c++; } }
      else if (d.align === 'right') { let c = t.c - 1; while (c >= 0 && t.x + t.w - clipX1 < tw + pad * 2) { if (this.cellHasContent(t.r, c)) break; clipX1 -= this.colW(c); c--; } }
      else if (d.align === 'center') { let c1 = t.c - 1, c2 = t.c + 1; while ((clipX2 - clipX1) < tw + pad * 2 && (c1 >= 0 || c2 < C.c2 + 20)) { if (c1 >= 0 && !this.cellHasContent(t.r, c1)) { clipX1 -= this.colW(c1); c1--; } else c1 = -1; if (!this.cellHasContent(t.r, c2)) { clipX2 += this.colW(c2); c2++; } else c2 = 99999999; if (c1 < 0 && c2 > 9999999) break; } }
    }
    ctx.save();
    ctx.beginPath(); ctx.rect(clipX1 + 1, t.y + 1, clipX2 - clipX1 - 2, t.h - 1); ctx.clip();
    ctx.textBaseline = 'alphabetic';
    for (const ln of lines) {
      const w = this.measure(ln, fontStr);
      let x;
      if (d.align === 'right') x = t.x + t.w - pad - w - (st.ha === 'right' ? indent : 0);
      else if (d.align === 'center' || d.align === 'centerAcross') x = t.x + (t.w - w) / 2;
      else if (d.align === 'fill') x = t.x + pad;
      else x = t.x + pad + indent;
      const by = y + lineH * 0.78;
      if (d.align === 'fill' && w > 0) { let xx = x; while (xx < t.x + t.w - pad) { ctx.fillText(ln, xx, by); xx += w; } }
      else ctx.fillText(ln, x, by);
      if (st.u) { ctx.fillRect(x, by + 1.5 * z, w, Math.max(1, z)); if (st.u === 'double') ctx.fillRect(x, by + 3.5 * z, w, Math.max(1, z)); }
      if (st.s) ctx.fillRect(x, by - lineH * 0.28, w, Math.max(1, z));
      y += lineH;
    }
    // hyperlink style
    ctx.restore();
  }
  wrapLines(text, fontStr, maxW) {
    const out = [];
    String(text).split('\n').forEach(par => {
      const words = par.split(/(\s+)/); let cur = '';
      for (const w of words) {
        const test = cur + w;
        if (this.measure(test, fontStr) <= maxW || !cur) { cur = test; }
        else { out.push(cur.replace(/\s+$/, '')); cur = w.replace(/^\s+/, ''); }
        while (this.measure(cur, fontStr) > maxW && cur.length > 1) { // break long word
          let k = cur.length - 1; while (k > 1 && this.measure(cur.slice(0, k), fontStr) > maxW) k--; out.push(cur.slice(0, k)); cur = cur.slice(k);
        }
      }
      out.push(cur);
    });
    return out;
  }
  cellHasContent(r, c) { const x = this.sheet.get(r, c); if (x && x.v != null && x.v !== '') return true; const sp = this.sheet.spill.get(key(r, c)); return !!(sp && sp.v != null); }
  drawDataBar(bar, x, y, w, h) {
    const ctx = this.ctx; const bw = Math.max(0, (w - 4) * bar.ratio);
    const g = ctx.createLinearGradient(x, 0, x + bw, 0); g.addColorStop(0, bar.color); g.addColorStop(1, 'rgba(255,255,255,0.6)');
    ctx.fillStyle = g; ctx.fillRect(x + 2, y + 3, bw, h - 6);
    ctx.strokeStyle = bar.color; ctx.lineWidth = 1; ctx.strokeRect(x + 2.5, y + 3.5, bw, h - 7);
  }
  drawBorders(bd, x, y, w, h) {
    const ctx = this.ctx, z = this.z;
    const line = (b, x1, y1, x2, y2) => {
      if (!b) return;
      const style = b.s || 'thin'; ctx.strokeStyle = b.c || '#000';
      ctx.lineWidth = style === 'thick' ? 3 : style === 'medium' || style === 'mediumDashed' ? 2 : 1;
      ctx.setLineDash(/dash/i.test(style) ? [4 * z, 2 * z] : /dot/i.test(style) || style === 'hair' ? [1, 2] : []);
      if (style === 'double') { ctx.lineWidth = 1; const dx = y1 === y2 ? 0 : 1.5, dy = y1 === y2 ? 1.5 : 0; ctx.beginPath(); ctx.moveTo(x1 - dx, y1 - dy); ctx.lineTo(x2 - dx, y2 - dy); ctx.moveTo(x1 + dx, y1 + dy); ctx.lineTo(x2 + dx, y2 + dy); ctx.stroke(); return; }
      const o = ctx.lineWidth % 2 ? 0.5 : 0;
      ctx.beginPath(); ctx.moveTo(Math.round(x1) - (x1 === x2 ? o : 0), Math.round(y1) - (y1 === y2 ? o : 0)); ctx.lineTo(Math.round(x2) - (x1 === x2 ? o : 0), Math.round(y2) - (y1 === y2 ? o : 0)); ctx.stroke();
    };
    line(bd.t, x, y, x + w, y); line(bd.b, x, y + h, x + w, y + h); line(bd.l, x, y, x, y + h); line(bd.r, x + w, y, x + w, y + h);
    if (bd.d) line(bd.d, x, y + h, x + w, y); if (bd.u) line(bd.u, x, y, x + w, y + h);
    ctx.setLineDash([]);
  }
  drawIndicators(R, C, X, Y) {
    const ctx = this.ctx, s = this.sheet, z = this.z;
    s.cells.forEach((x, k) => {
      if (!x.note) return; const r = Math.floor(k / 20000), c = k % 20000;
      if (r < R.r1 || r > R.r2 || c < C.c1 || c > C.c2) return;
      const xr = X(c + 1), y = Y(r);
      ctx.fillStyle = x.note.thread ? '#8764b8' : '#e81123'; ctx.beginPath(); ctx.moveTo(xr - 6 * z, y + 1); ctx.lineTo(xr - 1, y + 1); ctx.lineTo(xr - 1, y + 6 * z); ctx.closePath(); ctx.fill();
    });
    if (s.filter) {
      const f = s.filter; const r = f.r1;
      if (r >= R.r1 && r <= R.r2) for (let c = Math.max(f.c1, C.c1); c <= Math.min(f.c2, C.c2); c++) {
        const bx = X(c + 1) - 17 * z, by = Y(r) + (this.rowH(r) - 15 * z) / 2;
        const active = f.crit && f.crit[c];
        ctx.fillStyle = '#fff'; ctx.fillRect(bx, by, 15 * z, 15 * z); ctx.strokeStyle = '#a6a6a6'; ctx.lineWidth = 1; ctx.strokeRect(bx + 0.5, by + 0.5, 15 * z - 1, 15 * z - 1);
        ctx.fillStyle = '#444';
        if (active) { ctx.beginPath(); ctx.moveTo(bx + 3 * z, by + 3 * z); ctx.lineTo(bx + 12 * z, by + 3 * z); ctx.lineTo(bx + 8.5 * z, by + 8 * z); ctx.lineTo(bx + 8.5 * z, by + 12 * z); ctx.lineTo(bx + 6.5 * z, by + 11 * z); ctx.lineTo(bx + 6.5 * z, by + 8 * z); ctx.closePath(); ctx.fill(); }
        else { ctx.beginPath(); ctx.moveTo(bx + 4 * z, by + 6 * z); ctx.lineTo(bx + 11 * z, by + 6 * z); ctx.lineTo(bx + 7.5 * z, by + 10 * z); ctx.closePath(); ctx.fill(); }
      }
    }
    // spill range outline for active cell
    const sel = s.sel; const ac = s.get(sel.r, sel.c);
    if (ac && ac.spillTo) { ctx.strokeStyle = '#4472c4'; ctx.setLineDash([3, 2]); ctx.lineWidth = 1; ctx.strokeRect(X(sel.c) + 0.5, Y(sel.r) + 0.5, X(ac.spillTo.c2 + 1) - X(sel.c) - 1, Y(ac.spillTo.r2 + 1) - Y(sel.r) - 1); ctx.setLineDash([]); }
    // page breaks (page break preview / after print preview)
    if (s.showPageBreaks) drawBreaks(this, ctx, X, Y, R, C);
  }
  drawSelection(X, Y, R, C) {
    const ctx = this.ctx, s = this.sheet, sel = s.sel, z = this.z;
    const ranges = sel.ranges;
    const rx = rg => { const x1 = X(rg.c1), y1 = Y(rg.r1); return { x: x1, y: y1, w: X(rg.c2 + 1) - x1, h: Y(rg.r2 + 1) - y1 }; };
    // tint
    ctx.fillStyle = 'rgba(16,124,65,0.10)';
    for (const rg of ranges) {
      if (rg.r2 < R.r1 || rg.r1 > R.r2 || rg.c2 < C.c1 || rg.c1 > C.c2) continue;
      const b = rx(rg); ctx.fillRect(b.x, b.y, b.w, b.h);
    }
    // active cell (white)
    const am = s.mergeAt(sel.r, sel.c) || { r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c };
    const ab = rx(am);
    if (ranges.length === 1 && (ranges[0].r1 !== ranges[0].r2 || ranges[0].c1 !== ranges[0].c2)) { ctx.clearRect(ab.x + 1, ab.y + 1, ab.w - 1, ab.h - 1); ctx.fillStyle = '#fff'; ctx.fillRect(ab.x + 1, ab.y + 1, ab.w - 1, ab.h - 1); this.redrawCellContent(am, X, Y, R, C); }
    // border of last range
    const last = ranges[ranges.length - 1];
    const b = rx(last);
    ctx.strokeStyle = SEL_COLOR; ctx.lineWidth = 2;
    if (!this.editing || this.editing.r !== sel.r || this.editing.c !== sel.c || ranges.length > 1) ctx.strokeRect(Math.round(b.x) + 0.5, Math.round(b.y) + 0.5, Math.round(b.w) - 1, Math.round(b.h) - 1);
    if (ranges.length > 1) { ctx.lineWidth = 1; ctx.strokeRect(Math.round(ab.x) + 0.5, Math.round(ab.y) + 0.5, ab.w - 1, ab.h - 1); }
    // fill handle
    if (!this.editing && ranges.length === 1) {
      const hx = Math.round(b.x + b.w) - 3.5, hy = Math.round(b.y + b.h) - 3.5;
      ctx.fillStyle = '#fff'; ctx.fillRect(hx - 1, hy - 1, 8, 8); ctx.fillStyle = SEL_COLOR; ctx.fillRect(hx, hy, 6, 6);
      this._fillHandle = { x: hx + 3, y: hy + 3 };
    }
    // copy marquee
    if (this.copyRange && this.copyRange.sheet === s) {
      const cb = rx(this.copyRange.range);
      ctx.strokeStyle = SEL_COLOR; ctx.lineWidth = 1.5; ctx.setLineDash([4, 3]); ctx.lineDashOffset = -(Date.now() / 60 % 7);
      ctx.strokeRect(cb.x + 1, cb.y + 1, cb.w - 2, cb.h - 2); ctx.setLineDash([]);
      if (!this._ants) this._ants = setInterval(() => { if (!this.copyRange) { clearInterval(this._ants); this._ants = null; } else this.redraw(); }, 120);
    }
    // formula reference highlights while editing
    if (this.editing && this.editRefs) {
      this.editRefs.forEach((rf, i) => {
        if (rf.sheet && rf.sheet.toLowerCase() !== s.name.toLowerCase()) return;
        const col = REF_COLORS[i % REF_COLORS.length];
        const bb = rx({ r1: Math.min(rf.r1, rf.r2), c1: Math.min(rf.c1, rf.c2), r2: Math.min(MAX_ROWS - 1, Math.max(rf.r1, rf.r2)), c2: Math.min(MAX_COLS - 1, Math.max(rf.c1, rf.c2)) });
        ctx.fillStyle = col.replace(/^#(..)(..)(..)$/, (m, a, b2, c) => 'rgba(' + parseInt(a, 16) + ',' + parseInt(b2, 16) + ',' + parseInt(c, 16) + ',0.12)');
        ctx.fillRect(bb.x, bb.y, bb.w, bb.h);
        ctx.strokeStyle = col; ctx.lineWidth = 1.5; ctx.strokeRect(bb.x + 1, bb.y + 1, bb.w - 2, bb.h - 2);
        [[bb.x, bb.y], [bb.x + bb.w, bb.y], [bb.x, bb.y + bb.h], [bb.x + bb.w, bb.y + bb.h]].forEach(p => { ctx.fillStyle = col; ctx.fillRect(p[0] - 2.5, p[1] - 2.5, 5, 5); });
      });
    }
    // drag previews
    const ds = this.dragState;
    if (ds && ds.preview) { const pb = rx(ds.preview); ctx.strokeStyle = ds.type === 'fill' ? '#7f7f7f' : SEL_COLOR; ctx.setLineDash(ds.type === 'fill' ? [3, 3] : []); ctx.lineWidth = ds.type === 'fill' ? 1 : 2; ctx.strokeRect(pb.x + 1, pb.y + 1, pb.w - 2, pb.h - 2); ctx.setLineDash([]); }
  }
  redrawCellContent(m, X, Y, R, C) {
    const d = this.display(m.r1, m.c1);
    const x = X(m.c1), y = Y(m.r1);
    const fill = (d.cf && d.cf.fill) || (d.st && d.st.fill);
    if (fill && fill !== 'transparent') { this.ctx.fillStyle = fill; this.ctx.fillRect(x + 1, y + 1, X(m.c2 + 1) - x - 1, Y(m.r2 + 1) - y - 1); }
    if (d.text !== '') this.drawText({ r: m.r1, c: m.c1, x, y, w: X(m.c2 + 1) - x, h: Y(m.r2 + 1) - y, d, merged: m.r1 !== m.r2 || m.c1 !== m.c2 }, X, R, C);
  }
  drawHeaders(vis, hw, hh, fw, fh, sx, sy) {
    const ctx = this.ctx, s = this.sheet, z = this.z; if (!s.showHead) return;
    const sel = s.sel; const fr = s.freeze;
    const selCols = {}, selRows = {}, fullCol = {}, fullRow = {};
    sel.ranges.forEach(rg => { for (let c = Math.max(rg.c1, 0); c <= Math.min(rg.c2, vis.c2); c++) { selCols[c] = 1; if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1) fullCol[c] = 1; } for (let r = Math.max(rg.r1, 0); r <= Math.min(rg.r2, vis.r2); r++) { selRows[r] = 1; if (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1) fullRow[r] = 1; } });
    ctx.font = (11 * z).toFixed(1) + 'px "Segoe UI", Roboto, Arial, sans-serif'; ctx.textBaseline = 'middle';
    // column header
    ctx.fillStyle = '#f5f5f5'; ctx.fillRect(hw, 0, this.W, hh);
    const drawCols = (c1, c2, ox, clipL) => {
      ctx.save(); ctx.beginPath(); ctx.rect(clipL, 0, this.W - clipL, hh); ctx.clip();
      for (let c = c1; c <= c2; c++) {
        const w = this.colW(c); if (!w) { const x = ox + this.colX(c); ctx.fillStyle = '#a0a0a0'; ctx.fillRect(x - 1, 2, 2, hh - 4); continue; }
        const x = ox + this.colX(c);
        if (selCols[c]) { ctx.fillStyle = fullCol[c] ? '#107c41' : '#d3f0e0'; ctx.fillRect(x, 0, w, hh); ctx.fillStyle = fullCol[c] ? '#107c41' : SEL_COLOR; ctx.fillRect(x, hh - 2, w, 2); }
        ctx.fillStyle = '#d4d4d4'; ctx.fillRect(Math.round(x + w) - 0.5, 0, 1, hh);
        ctx.fillStyle = fullCol[c] ? '#fff' : selCols[c] ? '#0c5f32' : '#444'; ctx.textAlign = 'center';
        ctx.fillText(colName(c), x + w / 2, hh / 2 + 0.5);
      }
      ctx.restore();
    };
    if (fr.c) drawCols(0, fr.c - 1, hw, hw);
    drawCols(vis.c1, vis.c2, hw - sx, hw + fw);
    ctx.fillStyle = '#d4d4d4'; ctx.fillRect(hw, hh - 1, this.W, 1);
    // row header
    ctx.fillStyle = '#f5f5f5'; ctx.fillRect(0, hh, hw, this.H);
    const drawRows = (r1, r2, oy, clipT) => {
      ctx.save(); ctx.beginPath(); ctx.rect(0, clipT, hw, this.H - clipT); ctx.clip();
      for (let r = r1; r <= r2; r++) {
        const h2 = this.rowH(r); const y = oy + this.rowY(r);
        if (!h2) { ctx.fillStyle = '#a0a0a0'; ctx.fillRect(2, y - 1, hw - 4, 2); continue; }
        if (selRows[r]) { ctx.fillStyle = fullRow[r] ? '#107c41' : '#d3f0e0'; ctx.fillRect(0, y, hw, h2); ctx.fillStyle = SEL_COLOR; ctx.fillRect(hw - 2, y, 2, h2); }
        if (s.hiddenFilterR && s.hiddenFilterR[r + 1]) { /* filtered rows show blue numbers */ }
        ctx.fillStyle = '#d4d4d4'; ctx.fillRect(0, Math.round(y + h2) - 0.5, hw, 1);
        ctx.fillStyle = fullRow[r] ? '#fff' : selRows[r] ? '#0c5f32' : (s.filter && r > s.filter.r1 && r <= s.filter.r2 && s.filter.active ? '#0063b1' : '#444'); ctx.textAlign = 'right';
        ctx.fillText(String(r + 1), hw - 6 * z, y + h2 / 2 + 0.5);
      }
      ctx.restore();
    };
    if (fr.r) drawRows(0, fr.r - 1, hh, hh);
    drawRows(vis.r1, vis.r2, hh - sy, hh + fh);
    ctx.fillStyle = '#d4d4d4'; ctx.fillRect(hw - 1, hh, 1, this.H);
    // corner
    ctx.fillStyle = '#f5f5f5'; ctx.fillRect(0, 0, hw, hh); ctx.fillStyle = '#c0c0c0';
    ctx.beginPath(); ctx.moveTo(hw - 4, 4); ctx.lineTo(hw - 4, hh - 4); ctx.lineTo(hw - 4 - (hh - 8), hh - 4); ctx.closePath(); ctx.fill();
    ctx.fillStyle = '#d4d4d4'; ctx.fillRect(hw - 1, 0, 1, hh); ctx.fillRect(0, hh - 1, hw, 1);
    ctx.textAlign = 'left';
  }
  positionObjects() {
    const hw = this.headW(), hh = this.headH();
    this.objLayer.style.left = hw + 'px'; this.objLayer.style.top = hh + 'px';
    this.objLayer.style.right = '0'; this.objLayer.style.bottom = '0';
    const sx = this.scroller.scrollLeft, sy = this.scroller.scrollTop;
    const inner = this.objLayer.firstChild;
    if (inner) { inner.style.webkitTransform = 'translate(' + (-sx) + 'px,' + (-sy) + 'px)'; inner.style.transform = 'translate(' + (-sx) + 'px,' + (-sy) + 'px)'; }
    if (this.editing) this.placeEditor();
  }

  // ---------------------------------------------------------------- selection helpers
  select(r, c, extend, add) {
    const s = this.sheet; const sel = s.sel;
    r = clamp(r, 0, MAX_ROWS - 1); c = clamp(c, 0, MAX_COLS - 1);
    if (extend) {
      const last = sel.ranges[sel.ranges.length - 1];
      const nr = norm({ r1: sel.r, c1: sel.c, r2: r, c2: c });
      sel.ranges[sel.ranges.length - 1] = this.expandMerges(nr);
      sel.anchorEnd = { r, c };
    } else if (add) { sel.ranges.push(this.expandMerges({ r1: r, c1: c, r2: r, c2: c })); sel.r = r; sel.c = c; sel.anchorEnd = null; }
    else { const m = s.mergeAt(r, c); sel.ranges = [m ? { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 } : { r1: r, c1: c, r2: r, c2: c }]; sel.r = m ? m.r1 : r; sel.c = m ? m.c1 : c; sel.anchorEnd = null; }
    this.doc.onSelection();
    this.redraw();
  }
  expandMerges(rg) {
    const s = this.sheet; let changed = true; rg = Object.assign({}, rg);
    while (changed) { changed = false; for (const m of s.merges) { if (m.r2 >= rg.r1 && m.r1 <= rg.r2 && m.c2 >= rg.c1 && m.c1 <= rg.c2) { const n = { r1: Math.min(rg.r1, m.r1), c1: Math.min(rg.c1, m.c1), r2: Math.max(rg.r2, m.r2), c2: Math.max(rg.c2, m.c2) }; if (n.r1 !== rg.r1 || n.c1 !== rg.c1 || n.r2 !== rg.r2 || n.c2 !== rg.c2) { rg = n; changed = true; } } } }
    return rg;
  }
  selRange() { const s = this.sheet.sel; return s.ranges[s.ranges.length - 1]; }
  // Ctrl+Arrow: jump to edge of data region
  jump(r, c, dr, dc) {
    const s = this.sheet; const has = (rr, cc) => this.cellHasContent(rr, cc);
    const maxR = MAX_ROWS - 1, maxC = MAX_COLS - 1; const u = s.used();
    let nr = r, nc = c;
    const cur = has(r, c), next = has(r + dr, c + dc);
    if (cur && next) { while (nr + dr >= 0 && nr + dr <= maxR && nc + dc >= 0 && nc + dc <= maxC && has(nr + dr, nc + dc)) { nr += dr; nc += dc; } }
    else {
      nr += dr; nc += dc;
      while (nr >= 0 && nr <= maxR && nc >= 0 && nc <= maxC && !has(nr, nc)) { if ((dr > 0 && nr > u.r2) || (dc > 0 && nc > u.c2)) { nr = dr > 0 ? maxR : nr; nc = dc > 0 ? maxC : nc; break; } nr += dr; nc += dc; }
    }
    return { r: clamp(nr, 0, maxR), c: clamp(nc, 0, maxC) };
  }

  // ---------------------------------------------------------------- events
  bind() {
    const sc = this.scroller;
    sc.addEventListener('scroll', () => { this.updateSpacerLazy(); this.redraw(); this.hideAc(); });
    sc.addEventListener('mousedown', e => this.onDown(e));
    sc.addEventListener('mousemove', e => this.onHover(e));
    sc.addEventListener('dblclick', e => this.onDbl(e));
    sc.addEventListener('contextmenu', e => this.onContext(e));
    longPress(sc, e => this.onContext(e));
    // touch: tap to select (native scrolling stays)
    let tStart = null;
    sc.addEventListener('touchstart', e => { if (e.touches.length !== 1) { tStart = null; return; } const t = e.touches[0]; tStart = { x: t.clientX, y: t.clientY, time: Date.now(), sl: sc.scrollLeft, st: sc.scrollTop };
      // drag selection handle on touch
      const p = this.local(t); if (this._fillHandle && Math.abs(p.x - this._fillHandle.x) < 16 && Math.abs(p.y - this._fillHandle.y) < 16) { e.preventDefault(); this.startTouchExtend(); tStart = null; }
    }, { passive: false });
    sc.addEventListener('touchend', e => {
      if (!tStart) return; const t = e.changedTouches[0];
      if (Math.abs(t.clientX - tStart.x) < 8 && Math.abs(t.clientY - tStart.y) < 8 && Date.now() - tStart.time < 500 && sc.scrollLeft === tStart.sl && sc.scrollTop === tStart.st) {
        e.preventDefault();
        const p = this.local(t); const hit = this.hit(p.x, p.y);
        const now = Date.now(); const dbl = this._lastTap && now - this._lastTap.t < 380 && this._lastTap.r === hit.r && this._lastTap.c === hit.c;
        this._lastTap = { t: now, r: hit.r, c: hit.c };
        if (this.editing && this.isPointMode()) { this.insertRef({ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c }); return; }
        if (this.editing) this.commit();
        if (hit.zone === 'colhead') this.selectCols(hit.c, hit.c);
        else if (hit.zone === 'rowhead') this.selectRows(hit.r, hit.r);
        else if (hit.zone === 'corner') this.selectAll();
        else if (this.objectAt && this.doc.objectHitTest && this.doc.objectHitTest(t)) return;
        else { if (this.doc.filterButtonHit && this.doc.filterButtonHit(hit, p)) return; this.select(hit.r, hit.c); if (dbl) this.beginEdit('edit'); }
        this.focusSink();
      }
      tStart = null;
    });
    const ed = this.editor;
    ed.addEventListener('keydown', e => this.onKey(e));
    ed.addEventListener('input', () => this.onEditorInput());
    ed.addEventListener('compositionstart', () => { this.composing = true; if (!this.editing) this.beginEdit('enter', ''); });
    ed.addEventListener('compositionend', () => { this.composing = false; this.onEditorInput(); });
    ed.addEventListener('paste', e => { if (!this.editing) { e.preventDefault(); this.doc.pasteEvent(e); } });
    ed.addEventListener('copy', e => { if (!this.editing || !ed.value) { e.preventDefault(); this.doc.copyEvent(e, false); } });
    ed.addEventListener('cut', e => { if (!this.editing || !ed.value) { e.preventDefault(); this.doc.copyEvent(e, true); } });
    ed.addEventListener('blur', () => { setTimeout(() => { if (this.editing && document.activeElement !== this.editor && !this.doc.formulaInputFocused()) { /* keep editing while formula bar/insert-function used */ } }, 0); });
  }
  updateSpacerLazy() { const sc = this.scroller; const w = parseFloat(this.spacer.style.width) || 0, h2 = parseFloat(this.spacer.style.height) || 0; if (sc.scrollLeft + this.W * 1.5 > w || sc.scrollTop + this.H * 1.5 > h2) this.updateSpacer(); }
  local(e) { const r = this.el.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; }
  focusSink() { if (document.activeElement !== this.editor) { try { this.editor.focus({ preventScroll: true }); } catch (e) { try { this.editor.focus(); } catch (e2) {} } } if (!this.editing) this.placeSink(); }
  placeSink() {
    const s = this.sheet; const rc = this.cellRect(s.sel.r, s.sel.c);
    const ed = this.editor; ed.classList.remove('on'); ed.style.left = Math.max(0, rc.x) + 'px'; ed.style.top = Math.max(0, rc.y) + 'px'; ed.style.width = '2px'; ed.style.height = '2px';
    if (!this.composing) ed.value = '';
  }
  isPointMode() {
    if (!this.editing) return false;
    const v = this.editing.inFormulaBar ? this.doc.fbInput.value : this.editor.value;
    if (v[0] !== '=') return false;
    const pos = this.editing.inFormulaBar ? this.doc.fbInput.selectionStart : this.editor.selectionStart;
    const before = v.slice(0, pos).replace(/\s+$/, '');
    return /[=(,+\-*/^&<>:;%]$/.test(before) || this.pointRef != null;
  }
  onDown(e) {
    if (e.button === 2) return;
    const p = this.local(e); const hit = this.hit(p.x, p.y); const s = this.sheet; const sel = s.sel;
    // scrollbar area?
    if (p.x > this.scroller.clientWidth || p.y > this.scroller.clientHeight) return;
    e.preventDefault();
    if (this.editing && this.isPointMode() && hit.zone === 'cell') {
      const start = { r: hit.r, c: hit.c };
      this.insertRef({ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c });
      this.dragState = { type: 'point', start };
      this.trackDrag(ev => { const q = this.local(ev); const h2 = this.hit(q.x, q.y); this.insertRef(norm({ r1: start.r, c1: start.c, r2: h2.r, c2: h2.c }), true); });
      return;
    }
    if (this.editing) this.commit();
    this.focusSink();
    if (this.doc.objectsDeselect) this.doc.objectsDeselect();
    if (s.showPageBreaks) {
      const pbHit = this.pageBreakHit || hitPageBreak(this, p);
      if (pbHit) return this.startPageBreakDrag(pbHit, e);
    }
    if (hit.zone === 'colhead' || hit.zone === 'rowhead') {
      const rs = this.resizeHit(p, hit);
      if (rs) return this.startResize(rs, e);
      if (hit.zone === 'colhead') {
        const a = e.shiftKey ? sel.c : hit.c;
        if (e.ctrlKey) sel.ranges.push({ r1: 0, c1: hit.c, r2: MAX_ROWS - 1, c2: hit.c }); else this.selectCols(a, hit.c, true);
        this.trackDrag(ev => { const q = this.local(ev); const h2 = this.hit(q.x, q.y); this.selectCols(a, h2.c, true); });
      } else {
        const a = e.shiftKey ? sel.r : hit.r;
        if (e.ctrlKey) sel.ranges.push({ r1: hit.r, c1: 0, r2: hit.r, c2: MAX_COLS - 1 }); else this.selectRows(a, hit.r, true);
        this.trackDrag(ev => { const q = this.local(ev); const h2 = this.hit(q.x, q.y); this.selectRows(a, h2.r, true); });
      }
      return;
    }
    if (hit.zone === 'corner') { this.selectAll(); return; }
    // filter dropdown button
    if (this.doc.filterButtonHit && this.doc.filterButtonHit(hit, p)) return;
    // hyperlink ctrl+click
    const cell = s.get(hit.r, hit.c);
    if (cell && cell.link && (e.ctrlKey || e.metaKey)) { this.doc.followLink(cell.link); return; }
    // fill handle
    if (this._fillHandle && Math.abs(p.x - this._fillHandle.x) <= 5 && Math.abs(p.y - this._fillHandle.y) <= 5) return this.startFill(e);
    // move selection by border drag
    if (settings().mouse.dragDrop && this.onSelBorder(p)) return this.startMove(e);
    if (e.shiftKey) this.select(hit.r, hit.c, true);
    else this.select(hit.r, hit.c, false, e.ctrlKey || e.metaKey);
    const anchor = { r: sel.r, c: sel.c };
    this.trackDrag(ev => {
      const q = this.local(ev); const h2 = this.hit(clamp(q.x, this.headW() + 1, this.W), clamp(q.y, this.headH() + 1, this.H));
      const last = norm({ r1: anchor.r, c1: anchor.c, r2: h2.r, c2: h2.c });
      sel.ranges[sel.ranges.length - 1] = this.expandMerges(last); this.doc.onSelection(); this.redraw();
    });
  }
  selectCols(a, b, keep) { const s = this.sheet; s.sel.ranges = [{ r1: 0, c1: Math.min(a, b), r2: MAX_ROWS - 1, c2: Math.max(a, b) }]; if (!keep || s.sel.c < Math.min(a, b) || s.sel.c > Math.max(a, b)) { s.sel.c = a; } s.sel.r = this.visibleRange().r1; s.sel.c = a; this.doc.onSelection(); this.redraw(); }
  selectRows(a, b, keep) { const s = this.sheet; s.sel.ranges = [{ r1: Math.min(a, b), c1: 0, r2: Math.max(a, b), c2: MAX_COLS - 1 }]; s.sel.r = a; s.sel.c = this.visibleRange().c1; this.doc.onSelection(); this.redraw(); }
  selectAll() { const s = this.sheet; s.sel.ranges = [{ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }]; this.doc.onSelection(); this.redraw(); }
  // autoscroll + mouse tracking
  trackDrag(move, end) {
    let lastEv = null, timer = null;
    const auto = () => {
      if (!lastEv) return; const p = this.local(lastEv); const sc = this.scroller; let moved = false;
      if (p.x > this.W - 20) { sc.scrollLeft += 20; moved = true; } else if (p.x < this.headW() + 5 && sc.scrollLeft > 0) { sc.scrollLeft -= 20; moved = true; }
      if (p.y > this.H - 20) { sc.scrollTop += 20; moved = true; } else if (p.y < this.headH() + 5 && sc.scrollTop > 0) { sc.scrollTop -= 20; moved = true; }
      if (moved) move(lastEv);
    };
    const mv = ev => { lastEv = ev; move(ev); };
    const up = ev => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); clearInterval(timer); if (end) end(ev); this.dragState = null; this.redraw(); };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
    timer = setInterval(auto, 50);
  }
  resizeHit(p, hit) {
    const tol = 4;
    if (hit.zone === 'colhead') {
      for (const c of [hit.c, hit.c - 1]) { if (c < 0) continue; const rc = this.cellRect(0, c); if (Math.abs(p.x - (rc.x + rc.w)) <= tol) return { axis: 'c', i: c, start: p.x, size: this.sheet.colWidth(c) }; }
    } else if (hit.zone === 'rowhead') {
      for (const r of [hit.r, hit.r - 1]) { if (r < 0) continue; const rc = this.cellRect(r, 0); if (Math.abs(p.y - (rc.y + rc.h)) <= tol) return { axis: 'r', i: r, start: p.y, size: this.sheet.rowHeight(r) }; }
    }
    return null;
  }
  startResize(rs, e) {
    const s = this.sheet; const z = this.z;
    // apply to all selected full columns/rows if the resized one is within selection
    const rg = this.selRange();
    const multi = rs.axis === 'c' ? (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1 && rs.i >= rg.c1 && rs.i <= rg.c2) : (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1 && rs.i >= rg.r1 && rs.i <= rg.r2);
    const tipEl = h('div.xl-sizetip'); this.el.appendChild(tipEl);
    let size = rs.size;
    const before = JSON.stringify(rs.axis === 'c' ? s.colW : s.rowH);
    this.trackDrag(ev => {
      const q = this.local(ev); size = Math.max(0, Math.round(rs.size + ((rs.axis === 'c' ? q.x : q.y) - rs.start) / z));
      if (rs.axis === 'c') { if (multi) for (let c = rg.c1; c <= rg.c2; c++) s.colW[c] = size; else s.colW[rs.i] = size; if (!size) s.hiddenC[rs.i] = false; }
      else { if (multi) for (let r = rg.r1; r <= rg.r2; r++) s.rowH[r] = size; else s.rowH[rs.i] = size; }
      this.invalidateAxes(); this.redraw();
      tipEl.textContent = rs.axis === 'c' ? 'Rộng: ' + (size / 7).toFixed(2) + ' (' + size + ' px)' : 'Cao: ' + (size * 0.75).toFixed(2) + ' (' + size + ' px)';
      tipEl.style.left = (q.x + 12) + 'px'; tipEl.style.top = (q.y + 12) + 'px';
    }, () => { tipEl.parentNode && tipEl.parentNode.removeChild(tipEl); this.doc.commitPropsChange(rs.axis === 'c' ? 'colW' : 'rowH', before); });
  }
  onSelBorder(p) {
    const rg = this.selRange(); if (!rg || rg.r2 >= MAX_ROWS - 1 || rg.c2 >= MAX_COLS - 1) return false;
    const rc = this.cellRect(rg.r1, rg.c1, rg.r2, rg.c2); const t = 3;
    const inX = p.x >= rc.x - t && p.x <= rc.x + rc.w + t, inY = p.y >= rc.y - t && p.y <= rc.y + rc.h + t;
    if (!inX || !inY) return false;
    return Math.abs(p.x - rc.x) <= t || Math.abs(p.x - rc.x - rc.w) <= t || Math.abs(p.y - rc.y) <= t || Math.abs(p.y - rc.y - rc.h) <= t;
  }
  startMove(e) {
    const rg = Object.assign({}, this.selRange()); const p0 = this.local(e); const h0 = this.hit(p0.x, p0.y);
    this.dragState = { type: 'move', preview: rg };
    this.trackDrag(ev => { const q = this.local(ev); const h2 = this.hit(q.x, q.y); const dr = h2.r - h0.r, dc = h2.c - h0.c; this.dragState.preview = { r1: Math.max(0, rg.r1 + dr), c1: Math.max(0, rg.c1 + dc), r2: Math.max(0, rg.r2 + dr), c2: Math.max(0, rg.c2 + dc) }; this.dragState.copy = ev.ctrlKey; this.redraw(); },
      ev => { const pv = this.dragState && this.dragState.preview; const cp = ev.ctrlKey; if (pv && (pv.r1 !== rg.r1 || pv.c1 !== rg.c1)) this.doc.moveRange(rg, pv.r1, pv.c1, cp); });
  }
  startFill(e) {
    const rg = Object.assign({}, this.selRange());
    this.dragState = { type: 'fill', preview: rg };
    this.trackDrag(ev => {
      const q = this.local(ev); const h2 = this.hit(q.x, q.y);
      const dr = h2.r > rg.r2 ? h2.r - rg.r2 : h2.r < rg.r1 ? h2.r - rg.r1 : 0, dc = h2.c > rg.c2 ? h2.c - rg.c2 : h2.c < rg.c1 ? h2.c - rg.c1 : 0;
      let pv;
      if (Math.abs(dr) >= Math.abs(dc) * 1) pv = { r1: Math.min(rg.r1, h2.r), r2: Math.max(rg.r2, h2.r), c1: rg.c1, c2: rg.c2 };
      else pv = { r1: rg.r1, r2: rg.r2, c1: Math.min(rg.c1, h2.c), c2: Math.max(rg.c2, h2.c) };
      this.dragState.preview = pv; this.redraw();
      this.showTipAt(q, this.doc.fillPreviewText(rg, pv));
    }, () => { this.hideTip(); const pv = this.dragState && this.dragState.preview; if (pv) this.doc.autoFill(rg, pv); });
  }
  startPageBreakDrag(pbHit, e) {
    const s = this.sheet; const ps = s.pageSetup || (s.pageSetup = {});
    const a = pbHit.area;
    if (pbHit.type === 'rowBreak') {
      const origAt = pbHit.at;
      this.dragState = { type: 'pageBreakRow', original: origAt, targetRow: origAt, targetY: pbHit.y };
      this.trackDrag(ev => {
        const q = this.local(ev);
        const hit = this.hit(q.x, q.y);
        const targetRow = Math.max(0, Math.min(MAX_ROWS - 1, hit.r));
        const isOut = hit.r < a.r1 || hit.r > a.r2;
        this.dragState = { type: 'pageBreakRow', original: origAt, targetRow, targetY: this.cellRect(targetRow, a.c1).y };
        this.showTipAt(q, 'Ngắt trang hàng: ' + (targetRow + 1) + (isOut ? ' (Kéo ra ngoài vùng in để xóa)' : ''));
        this.redraw();
      }, ev => {
        this.hideTip();
        const tr = this.dragState && this.dragState.targetRow;
        this.dragState = null;
        const q = this.local(ev); const hit = this.hit(q.x, q.y);
        const isOut = hit.r < a.r1 || hit.r > a.r2;
        this.doc.change('Điều chỉnh ngắt trang hàng', s, 'props', () => {
          ps.rowBreaks = (ps.rowBreaks || []).filter(r => r !== origAt);
          if (!isOut && tr > a.r1 && tr <= a.r2 && ps.rowBreaks.indexOf(tr) < 0) {
            ps.rowBreaks.push(tr);
            ps.rowBreaks.sort((x, y) => x - y);
          }
          if (ps.fit === 'page') ps.fit = 'none';
        });
        s.showPageBreaks = true;
        this.doc.markDirty();
        this.redraw();
        if (this.doc.refreshUi) this.doc.refreshUi();
        toast(isOut ? 'Đã loại bỏ ngắt trang hàng' : 'Đã điều chỉnh ngắt trang đến hàng ' + (tr + 1));
      });
      return;
    }
    if (pbHit.type === 'colBreak') {
      const origAt = pbHit.at;
      this.dragState = { type: 'pageBreakCol', original: origAt, targetCol: origAt, targetX: pbHit.x };
      this.trackDrag(ev => {
        const q = this.local(ev);
        const hit = this.hit(q.x, q.y);
        const targetCol = Math.max(0, Math.min(MAX_COLS - 1, hit.c));
        const isOut = hit.c < a.c1 || hit.c > a.c2;
        this.dragState = { type: 'pageBreakCol', original: origAt, targetCol, targetX: this.cellRect(a.r1, targetCol).x };
        this.showTipAt(q, 'Ngắt trang cột: ' + colName(targetCol) + (isOut ? ' (Kéo ra ngoài vùng in để xóa)' : ''));
        this.redraw();
      }, ev => {
        this.hideTip();
        const tc = this.dragState && this.dragState.targetCol;
        this.dragState = null;
        const q = this.local(ev); const hit = this.hit(q.x, q.y);
        const isOut = hit.c < a.c1 || hit.c > a.c2;
        this.doc.change('Điều chỉnh ngắt trang cột', s, 'props', () => {
          ps.colBreaks = (ps.colBreaks || []).filter(c => c !== origAt);
          if (!isOut && tc > a.c1 && tc <= a.c2 && ps.colBreaks.indexOf(tc) < 0) {
            ps.colBreaks.push(tc);
            ps.colBreaks.sort((x, y) => x - y);
          }
          if (ps.fit === 'page' || ps.fit === 'width') ps.fit = 'none';
        });
        s.showPageBreaks = true;
        this.doc.markDirty();
        this.redraw();
        if (this.doc.refreshUi) this.doc.refreshUi();
        toast(isOut ? 'Đã loại bỏ ngắt trang cột' : 'Đã điều chỉnh ngắt trang đến trước cột ' + colName(tc));
      });
      return;
    }
    if (pbHit.type === 'printArea') {
      const part = pbHit.part;
      let curArea = { r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 };
      this.dragState = { type: 'printAreaResize', preview: curArea, part };
      this.trackDrag(ev => {
        const q = this.local(ev);
        const hit = this.hit(q.x, q.y);
        let r1 = a.r1, c1 = a.c1, r2 = a.r2, c2 = a.c2;
        if (part === 'edge-top' || part.indexOf('tl') >= 0 || part.indexOf('tr') >= 0) r1 = Math.min(a.r2, Math.max(0, hit.r));
        if (part === 'edge-bottom' || part.indexOf('bl') >= 0 || part.indexOf('br') >= 0) r2 = Math.max(a.r1, Math.min(MAX_ROWS - 1, hit.r));
        if (part === 'edge-left' || part.indexOf('tl') >= 0 || part.indexOf('bl') >= 0) c1 = Math.min(a.c2, Math.max(0, hit.c));
        if (part === 'edge-right' || part.indexOf('tr') >= 0 || part.indexOf('br') >= 0) c2 = Math.max(a.c1, Math.min(MAX_COLS - 1, hit.c));
        curArea = norm({ r1, c1, r2, c2 });
        this.dragState = { type: 'printAreaResize', preview: curArea, part };
        const addr = rangeAddr(curArea);
        const cols = curArea.c2 - curArea.c1 + 1, rows = curArea.r2 - curArea.r1 + 1;
        this.showTipAt(q, 'Vùng in: ' + addr + ' (' + cols + ' cột × ' + rows + ' hàng)');
        this.redraw();
      }, () => {
        this.hideTip();
        const pv = this.dragState && this.dragState.preview;
        this.dragState = null;
        if (pv) {
          const addr = rangeAddr(pv);
          this.doc.change('Điều chỉnh vùng in', s, 'props', () => {
            ps.printArea = addr;
          });
          s.showPageBreaks = true;
          this.doc.markDirty();
          this.redraw();
          if (this.doc.refreshUi) this.doc.refreshUi();
          toast('Đã cập nhật vùng in: ' + addr);
        }
      });
      return;
    }
  }
  startTouchExtend() {
    const s = this.sheet; const anchor = { r: s.sel.r, c: s.sel.c };
    const mv = ev => { const t = ev.touches[0]; const q = this.local(t); const h2 = this.hit(q.x, q.y); s.sel.ranges = [this.expandMerges(norm({ r1: anchor.r, c1: anchor.c, r2: h2.r, c2: h2.c }))]; this.doc.onSelection(); this.redraw(); ev.preventDefault(); };
    const up = () => { document.removeEventListener('touchmove', mv); document.removeEventListener('touchend', up); };
    document.addEventListener('touchmove', mv, { passive: false }); document.addEventListener('touchend', up);
  }
  showTipAt(q, text) { if (!text) return this.hideTip(); this.tipBox.textContent = text; this.tipBox.classList.remove('hidden'); this.tipBox.style.left = (q.x + 14) + 'px'; this.tipBox.style.top = (q.y + 14) + 'px'; }
  hideTip() { this.tipBox.classList.add('hidden'); }
  onHover(e) {
    if (this.dragState) return;
    const p = this.local(e); const hit = this.hit(p.x, p.y); const sc = this.scroller;
    if (this.sheet.showPageBreaks) {
      const pbHit = hitPageBreak(this, p);
      if (pbHit) {
        if (!this.pageBreakHit || this.pageBreakHit.at !== pbHit.at || this.pageBreakHit.part !== pbHit.part) {
          this.pageBreakHit = pbHit;
          this.redraw();
        }
        if (sc.style.cursor !== pbHit.cursor) sc.style.cursor = pbHit.cursor;
        return;
      } else if (this.pageBreakHit) {
        this.pageBreakHit = null;
        this.redraw();
      }
    }
    let cur = 'cell';
    if (hit.zone === 'colhead') cur = this.resizeHit(p, hit) ? 'col-resize' : 's-resize';
    else if (hit.zone === 'rowhead') cur = this.resizeHit(p, hit) ? 'row-resize' : 'e-resize';
    else if (hit.zone === 'corner') cur = 'default';
    else if (this._fillHandle && Math.abs(p.x - this._fillHandle.x) <= 5 && Math.abs(p.y - this._fillHandle.y) <= 5) cur = 'crosshair';
    else if (!this.editing && settings().mouse.dragDrop && this.onSelBorder(p)) cur = 'move';
    else { const cell = this.sheet.get(hit.r, hit.c); if (cell && cell.link) cur = 'pointer'; }
    if (sc.style.cursor !== cur) sc.style.cursor = cur;
    // comment hover preview
    if (hit.zone === 'cell' && settings().mouse.hoverPreview) {
      const cell = this.sheet.get(hit.r, hit.c);
      if (cell && cell.note) { if (!this._noteShown || this._noteShown.r !== hit.r || this._noteShown.c !== hit.c) { this._noteShown = { r: hit.r, c: hit.c }; this.doc.showNotePopup(hit.r, hit.c); } }
      else if (this._noteShown) { this._noteShown = null; this.doc.hideNotePopup(); }
      if (cell && cell.link) this.showTipAt(p, (cell.link.url || cell.link) + '\nCtrl + bấm để mở liên kết'); else if (!this.dragState) this.hideTip();
    }
  }
  onDbl(e) {
    if (!settings().mouse.doubleClickEdit) return;
    const p = this.local(e); const hit = this.hit(p.x, p.y);
    if (hit.zone === 'colhead' || hit.zone === 'rowhead') { const rs = this.resizeHit(p, hit); if (rs) { this.doc.autoFit(rs.axis, rs.i); return; } }
    if (this._fillHandle && Math.abs(p.x - this._fillHandle.x) <= 6 && Math.abs(p.y - this._fillHandle.y) <= 6) { this.doc.autoFillDown(); return; }
    if (hit.zone === 'cell') { this.select(hit.r, hit.c); this.beginEdit('edit'); }
  }
  onContext(e) {
    if (!settings().mouse.rightClickMenus) return;
    const p = this.local(e); const hit = this.hit(p.x, p.y);
    if (this.editing) this.commit();
    const s = this.sheet;
    if (s.showPageBreaks) {
      const pbHit = this.pageBreakHit || hitPageBreak(this, p);
      if (pbHit && (pbHit.type === 'rowBreak' || pbHit.type === 'colBreak' || pbHit.type === 'printArea')) {
        const ps = s.pageSetup || (s.pageSetup = {});
        const items = [
          pbHit.type === 'rowBreak' ? { label: 'Loại bỏ ngắt trang này (Hàng ' + (pbHit.at + 1) + ')', icon: 'minus', run: () => {
            this.doc.change('Loại bỏ ngắt trang', s, 'props', () => { ps.rowBreaks = (ps.rowBreaks || []).filter(x => x !== pbHit.at); });
            this.redraw(); if (this.doc.refreshUi) this.doc.refreshUi(); toast('Đã xóa ngắt trang hàng ' + (pbHit.at + 1));
          } } : null,
          pbHit.type === 'colBreak' ? { label: 'Loại bỏ ngắt trang này (Cột ' + colName(pbHit.at) + ')', icon: 'minus', run: () => {
            this.doc.change('Loại bỏ ngắt trang', s, 'props', () => { ps.colBreaks = (ps.colBreaks || []).filter(x => x !== pbHit.at); });
            this.redraw(); if (this.doc.refreshUi) this.doc.refreshUi(); toast('Đã xóa ngắt trang cột ' + colName(pbHit.at));
          } } : null,
          { label: 'Đặt lại tất cả ngắt trang', icon: 'rotate-ccw', run: () => import('./pages.js').then(m => m.resetBreaks(this.doc)) },
          { sep: true },
          { label: 'Đặt vùng in (vùng chọn)', icon: 'crop', run: () => import('./pages.js').then(m => m.setPrintArea(this.doc)) },
          { label: 'Thêm vào vùng in', icon: 'plus-square', run: () => import('./pages.js').then(m => m.addToPrintArea(this.doc)) },
          { label: 'Xóa vùng in', icon: 'trash-2', run: () => import('./pages.js').then(m => m.clearPrintArea(this.doc)) },
          { sep: true },
          { label: 'Xem trước khi In…', icon: 'printer', run: () => this.doc.shell.backstage(this.doc, 'print') }
        ].filter(Boolean);
        contextMenu(e, items);
        this.focusSink();
        return;
      }
    }
    const inSel = s.sel.ranges.some(rg => inRange(rg, hit.r, hit.c));
    if (hit.zone === 'colhead') { if (!inSel) this.selectCols(hit.c, hit.c); contextMenu(e, this.doc.contextItems('col')); }
    else if (hit.zone === 'rowhead') { if (!inSel) this.selectRows(hit.r, hit.r); contextMenu(e, this.doc.contextItems('row')); }
    else if (hit.zone === 'cell') { if (!inSel) this.select(hit.r, hit.c); contextMenu(e, this.doc.contextItems('cell')); }
    else e.preventDefault();
    this.focusSink();
  }

  // ---------------------------------------------------------------- keyboard
  onKey(e) {
    const k = keyName(e), ch = chord(e); const s = this.sheet; const sel = s.sel;
    if (this.composing) return;
    if (this.editing) return this.onEditKey(e, k, ch);
    const ctrl = e.ctrlKey || e.metaKey;
    const move = (dr, dc) => {
      e.preventDefault();
      if (ctrl) { const t = this.jump(e.shiftKey && sel.anchorEnd ? sel.anchorEnd.r : sel.r, e.shiftKey && sel.anchorEnd ? sel.anchorEnd.c : sel.c, dr, dc); if (e.shiftKey) this.select(t.r, t.c, true); else this.select(t.r, t.c); this.scrollTo(t.r, t.c); return; }
      if (e.shiftKey) { const end = sel.anchorEnd || { r: sel.r, c: sel.c }; let nr = end.r + dr, nc = end.c + dc; while (dr && this.sheet.rowHeight(nr) === 0 && nr > 0 && nr < MAX_ROWS - 1) nr += dr; this.select(nr, nc, true); this.scrollTo(clamp(nr, 0, MAX_ROWS - 1), clamp(nc, 0, MAX_COLS - 1)); return; }
      let nr = sel.r + dr, nc = sel.c + dc;
      const m = s.mergeAt(sel.r, sel.c); if (m) { if (dr > 0) nr = m.r2 + 1; if (dc > 0) nc = m.c2 + 1; }
      while (dr && nr > 0 && nr < MAX_ROWS - 1 && this.sheet.rowHeight(nr) === 0) nr += dr;
      while (dc && nc > 0 && nc < MAX_COLS - 1 && this.sheet.colWidth(nc) === 0) nc += dc;
      this.select(nr, nc); this.scrollTo(clamp(nr, 0, MAX_ROWS - 1), clamp(nc, 0, MAX_COLS - 1));
    };
    switch (k) {
      case 'ArrowUp': return move(-1, 0);
      case 'ArrowDown': return move(1, 0);
      case 'ArrowLeft': return move(0, -1);
      case 'ArrowRight': return move(0, 1);
      case 'Tab': e.preventDefault(); return this.moveInSelection(0, e.shiftKey ? -1 : 1);
      case 'Enter': e.preventDefault(); if (ctrl) return; return this.enterMove(e.shiftKey);
      case 'Home': e.preventDefault(); if (ctrl) { this.select(s.freeze.r, s.freeze.c, e.shiftKey); this.scrollTo(0, 0); } else { this.select(sel.r, 0, e.shiftKey); this.scrollTo(sel.r, 0); } return;
      case 'End': if (ctrl) { e.preventDefault(); const u = s.used(); this.select(u.r2, u.c2, e.shiftKey); this.scrollTo(u.r2, u.c2); } return;
      case 'PageDown': case 'PageUp': {
        e.preventDefault();
        if (ctrl) { this.doc.switchSheet(k === 'PageDown' ? 1 : -1); return; }
        const vr = this.visibleRange(); const n = Math.max(1, vr.r2 - vr.r1 - 1) * (k === 'PageDown' ? 1 : -1);
        if (e.altKey) { const vc = Math.max(1, vr.c2 - vr.c1 - 1) * (k === 'PageDown' ? 1 : -1); this.select(sel.r, clamp(sel.c + vc, 0, MAX_COLS - 1)); this.scrollTo(sel.r, sel.c); return; }
        this.select(clamp(sel.r + n, 0, MAX_ROWS - 1), sel.c, e.shiftKey); this.scroller.scrollTop += n * 20 * this.z; return;
      }
      case 'F2': e.preventDefault(); return this.beginEdit('edit');
      case 'Delete': e.preventDefault(); return this.doc.clearContents();
      case 'Backspace': e.preventDefault(); this.doc.clearContents(true); return this.beginEdit('enter', '');
      case 'Escape': if (this.copyRange) { this.copyRange = null; this.redraw(); } return;
      case 'F4': e.preventDefault(); return this.doc.repeatLast && this.doc.repeatLast();
    }
    if (ch === 'ctrl+a') { e.preventDefault(); return this.selectAll(); }
    if (ch === 'ctrl+ ' || (ctrl && e.keyCode === 32)) { e.preventDefault(); return this.selectCols(sel.c, sel.c); }
    if (ch === 'shift+ ' || (e.shiftKey && !ctrl && e.keyCode === 32 && !this.editing)) { e.preventDefault(); return this.selectRows(sel.r, sel.r); }
    if (this.doc.shortcut(ch, e)) { e.preventDefault(); return; }
    // printable char starts editing (handled by input event to support IME)
  }
  // Excel option "After pressing Enter, move selection" (Tùy chọn > Chuột & con trỏ)
  enterMove(back) { const d = settings().mouse.enterMove || 'down'; if (d === 'none') return; const v = { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1] }[d]; return this.moveInSelection(back ? -v[0] : v[0], back ? -v[1] : v[1]); }
  moveInSelection(dr, dc) {
    const s = this.sheet, sel = s.sel; const rg = this.selRange();
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2 && sel.ranges.length === 1;
    if (single) { let nr = sel.r + dr, nc = sel.c + dc; const m = s.mergeAt(sel.r, sel.c); if (m && dr > 0) nr = m.r2 + 1; if (m && dc > 0) nc = m.c2 + 1; if (dc && this._tabStart == null) this._tabStart = sel.c; if (dr > 0 && this._tabStart != null) { nc = this._tabStart; } if (dr) this._tabStart = dc ? this._tabStart : null; this.select(clamp(nr, 0, MAX_ROWS - 1), clamp(nc, 0, MAX_COLS - 1)); this.scrollTo(s.sel.r, s.sel.c); return; }
    let r = sel.r + dr, c = sel.c + dc;
    if (dr) { if (r > rg.r2) { r = rg.r1; c++; if (c > rg.c2) c = rg.c1; } if (r < rg.r1) { r = rg.r2; c--; if (c < rg.c1) c = rg.c2; } }
    if (dc) { if (c > rg.c2) { c = rg.c1; r++; if (r > rg.r2) r = rg.r1; } if (c < rg.c1) { c = rg.c2; r--; if (r < rg.r1) r = rg.r2; } }
    sel.r = r; sel.c = c; this.doc.onSelection(); this.scrollTo(r, c); this.redraw();
  }

  // ---------------------------------------------------------------- editing
  beginEdit(mode, initial) {
    const s = this.sheet, sel = s.sel;
    if (this.doc.readOnly || (s.protected && this.doc.isLocked(sel.r, sel.c))) { toast('Ô này đang được bảo vệ. Bỏ bảo vệ trang tính (Xem lại > Bỏ bảo vệ) để sửa.', 'err'); return; }
    if (this.doc.dvListFor && mode === 'edit' && false) return;
    const cell = s.get(sel.r, sel.c);
    const text = initial !== undefined ? initial : (cell ? (cell.f ? '=' + cell.f : this.doc.editText(cell)) : '');
    this.editing = { r: sel.r, c: sel.c, mode, orig: cell ? Object.assign({}, cell) : null };
    const ed = this.editor;
    ed.classList.add('on');
    if (initial === undefined || initial !== '' || !this.composing) ed.value = text;
    const st = s.styleAt(sel.r, sel.c) || {};
    ed.style.font = this.font(st); ed.style.color = st.color || '#000'; ed.style.background = st.fill && st.fill !== 'transparent' ? st.fill : '#fff';
    ed.style.textAlign = st.ha === 'right' || st.ha === 'center' ? st.ha : 'left';
    this.placeEditor();
    this.focusSink();
    const L = ed.value.length; try { ed.setSelectionRange(L, L); } catch (e) {}
    this.doc.onEditStart(text);
    this.updateEditRefs();
    this.redraw();
  }
  placeEditor() {
    const ed = this.editor; const g = this.editing; if (!g) return;
    const m = this.sheet.mergeAt(g.r, g.c) || { r1: g.r, c1: g.c, r2: g.r, c2: g.c };
    const rc = this.cellRect(m.r1, m.c1, m.r2, m.c2);
    ed.style.left = rc.x + 'px'; ed.style.top = rc.y + 'px';
    ed.style.minWidth = Math.max(rc.w, 20) + 'px'; ed.style.minHeight = rc.h + 'px'; ed.style.height = rc.h + 'px';
    // auto-grow width to fit text
    const lines = ed.value.split('\n'); const fontStr = ed.style.font;
    const w = Math.max.apply(null, lines.map(l => this.measure(l + 'MM', fontStr)));
    ed.style.width = Math.min(Math.max(rc.w, w + 8), this.W - rc.x - 4) + 'px';
    ed.style.height = Math.max(rc.h, lines.length * parseFloat(fontStr) * 1.25 + 4) + 'px';
  }
  onEditorInput() {
    if (this.composing && !this.editing) return;
    if (!this.editing) { const v = this.editor.value; if (!v) return; if (settings().mouse.typeToEdit === false) { this.editor.value = ''; return; } this.beginEdit('enter', v); return; }
    this.pointRef = null;
    this.placeEditor();
    this.doc.onEditInput(this.editor.value);
    this.updateEditRefs(); this.autoComplete(); this.redraw();
  }
  updateEditRefs() { const v = this.editing ? (this.editing.inFormulaBar ? this.doc.fbInput.value : this.editor.value) : ''; this.editRefs = v[0] === '=' ? formulaRefs(v.slice(1)).map(x => x.ref) : null; }
  onEditKey(e, k, ch) {
    const ed = this.editor; const g = this.editing;
    if (this.acBox && !this.acBox.classList.contains('hidden')) {
      const items = this.acBox.querySelectorAll('.ac-item'); let i = Array.prototype.findIndex.call(items, x => x.classList.contains('on'));
      if (k === 'ArrowDown' || k === 'ArrowUp') { e.preventDefault(); if (items[i]) items[i].classList.remove('on'); i = clamp(i + (k === 'ArrowDown' ? 1 : -1), 0, items.length - 1); items[i].classList.add('on'); items[i].scrollIntoView && items[i].scrollIntoView({ block: 'nearest' }); return; }
      if (k === 'Tab' || (k === 'Enter' && i >= 0 && this._acNav)) { e.preventDefault(); if (items[Math.max(0, i)]) items[Math.max(0, i)].click(); return; }
      if (k === 'Escape') { e.preventDefault(); this.hideAc(); return; }
    }
    if (k === 'Enter' && e.altKey) { e.preventDefault(); const p = ed.selectionStart; ed.value = ed.value.slice(0, p) + '\n' + ed.value.slice(ed.selectionEnd); ed.setSelectionRange(p + 1, p + 1); this.onEditorInput(); return; }
    if (k === 'Enter') { e.preventDefault(); if (e.ctrlKey) { this.commit(true); return; } this.commit(); this.enterMove(e.shiftKey); return; }
    if (k === 'Tab') { e.preventDefault(); this.commit(); this.moveInSelection(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'Escape') { e.preventDefault(); this.cancel(); return; }
    if (k === 'F2') { e.preventDefault(); g.mode = g.mode === 'enter' ? 'edit' : 'enter'; this.doc.onEditStart(ed.value, g.mode); return; }
    if (k === 'F4') { e.preventDefault(); const r = cycleAbs(ed.value, ed.selectionStart); if (r) { ed.value = r.text; ed.setSelectionRange(r.caret, r.caret); this.onEditorInput(); } return; }
    if (g.mode === 'enter' && /^Arrow/.test(k)) {
      if (this.isPointMode()) { e.preventDefault(); this.keyPoint(k, e.shiftKey); return; }
      e.preventDefault(); this.commit();
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
      this.select(this.sheet.sel.r + d[0], this.sheet.sel.c + d[1]); this.scrollTo(this.sheet.sel.r, this.sheet.sel.c);
    }
  }
  keyPoint(k, ext) {
    const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
    const base = this.pointRef || { r1: this.editing.r, c1: this.editing.c, r2: this.editing.r, c2: this.editing.c, anchor: null };
    let nr;
    if (ext && this.pointRef) nr = { r1: base.r1, c1: base.c1, r2: clamp(base.r2 + d[0], 0, MAX_ROWS - 1), c2: clamp(base.c2 + d[1], 0, MAX_COLS - 1) };
    else { const r = clamp(base.r2 + d[0], 0, MAX_ROWS - 1), c = clamp(base.c2 + d[1], 0, MAX_COLS - 1); nr = { r1: r, c1: c, r2: r, c2: c }; }
    this.insertRef(nr, !!this.pointRef);
  }
  // insert / replace the reference being pointed while editing a formula
  insertRef(rg, replace) {
    const fb = this.editing && this.editing.inFormulaBar;
    const ed = fb ? this.doc.fbInput : this.editor;
    const s = this.sheet; const sheetPrefix = this.editing && this.editing.sheet && this.editing.sheet !== s ? (/^[\w.]+$/.test(s.name) ? s.name : "'" + s.name + "'") + '!' : '';
    const n = norm(rg); const txt = sheetPrefix + rangeAddr(n);
    if (replace && this.pointRef && this.pointRef.start != null) {
      ed.value = ed.value.slice(0, this.pointRef.start) + txt + ed.value.slice(this.pointRef.end);
      this.pointRef.end = this.pointRef.start + txt.length;
    } else {
      const p = ed.selectionStart;
      ed.value = ed.value.slice(0, p) + txt + ed.value.slice(ed.selectionEnd);
      this.pointRef = { start: p, end: p + txt.length };
    }
    Object.assign(this.pointRef, { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 });
    try { ed.setSelectionRange(this.pointRef.end, this.pointRef.end); } catch (e) {}
    if (!fb) this.placeEditor();
    this.doc.onEditInput(ed.value, fb);
    this.updateEditRefs(); this.redraw();
    if (fb) ed.focus();
  }
  commit(fillSelection) {
    const g = this.editing; if (!g) return;
    const text = g.inFormulaBar ? this.doc.fbInput.value : this.editor.value;
    this.editing = null; this.pointRef = null; this.editRefs = null; this.hideAc(); this.hideTip();
    this.editor.classList.remove('on'); this.editor.value = '';
    const s = this.sheet;
    if (g.sheet && g.sheet !== s) { this.doc.activateSheet(g.sheet); }
    this.doc.setInput(g.r, g.c, text, fillSelection);
    this.doc.onEditEnd();
    this.placeSink(); this.redraw();
  }
  cancel() {
    if (!this.editing) return;
    const g = this.editing;
    this.editing = null; this.pointRef = null; this.editRefs = null; this.hideAc();
    this.editor.classList.remove('on'); this.editor.value = '';
    if (g.sheet && g.sheet !== this.sheet) this.doc.activateSheet(g.sheet);
    this.doc.onEditEnd(true);
    this.placeSink(); this.redraw();
  }
  // function-name autocomplete while typing a formula
  autoComplete() {
    const g = this.editing; if (!g) return;
    const ed = g.inFormulaBar ? this.doc.fbInput : this.editor;
    const v = ed.value, p = ed.selectionStart;
    if (v[0] !== '=') return this.hideAc();
    const m = /([A-Za-z][A-Za-z0-9.]*)$/.exec(v.slice(0, p));
    this.showFnTip(v.slice(0, p));
    if (!m || /\d$/.test(m[1]) && /^[A-Za-z]{1,3}\d+$/.test(m[1])) return this.hideAc();
    const q = m[1].toUpperCase();
    const names = Object.keys(FN).filter(n => n.indexOf(q) === 0).sort().slice(0, 12);
    if (!names.length) return this.hideAc();
    const meta = this.doc.fnMeta || {};
    this.acBox.innerHTML = '';
    names.forEach((n, i) => {
      const it = h('div.ac-item' + (i === 0 ? '.on' : ''), ic('function-square', 14), h('b', n), h('span', (meta[n] && meta[n][2]) || ''));
      it.addEventListener('mousedown', e2 => { e2.preventDefault(); const start = p - m[1].length; ed.value = v.slice(0, start) + n + '(' + v.slice(p); const np = start + n.length + 1; ed.setSelectionRange(np, np); this.hideAc(); if (g.inFormulaBar) this.doc.onFbInput(); else this.onEditorInput(); ed.focus(); });
      this.acBox.appendChild(it);
    });
    this._acNav = true;
    const rc = g.inFormulaBar ? null : this.cellRect(g.r, g.c);
    this.acBox.classList.remove('hidden');
    if (rc) { this.acBox.style.left = rc.x + 'px'; this.acBox.style.top = (rc.y + rc.h + 2) + 'px'; }
    else { const fr = this.doc.fbInput.getBoundingClientRect(), er = this.el.getBoundingClientRect(); this.acBox.style.left = (fr.left - er.left) + 'px'; this.acBox.style.top = '2px'; }
  }
  hideAc() { if (this.acBox) this.acBox.classList.add('hidden'); }
  showFnTip(before) {
    // find innermost open function
    let depth = 0, name = null, argi = 0;
    for (let i = before.length - 1; i >= 0; i--) {
      const ch = before[i];
      if (ch === ')') depth++;
      else if (ch === '(') { if (depth === 0) { const m = /([A-Za-z][A-Za-z0-9.]*)$/.exec(before.slice(0, i)); name = m ? m[1].toUpperCase() : null; break; } depth--; }
      else if (ch === ',' && depth === 0) argi++;
      else if (ch === '"') { const j = before.lastIndexOf('"', i - 1); if (j >= 0) i = j; }
    }
    const meta = this.doc.fnMeta || {};
    if (!name || !meta[name]) { this.doc.fnTip(null); return; }
    this.doc.fnTip(name, meta[name][1], argi);
  }
}
