// UI kit: menus, context menus, dialogs, toasts, color picker, tooltips, ribbon
import { h, ic, esc, clamp, keyName, $, $$, debounce } from './util.js';

// ---------------------------------------------------------------- toast
let toastBox;
export function toast(msg, kind, ms) {
  if (!toastBox) { toastBox = h('div.toast-box'); document.body.appendChild(toastBox); }
  const t = h('div.toast' + (kind ? '.' + kind : ''), ic(kind === 'err' ? 'circle-alert' : kind === 'ok' ? 'circle-check' : 'info', 16), h('span', msg));
  toastBox.appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove && t.parentNode && t.parentNode.removeChild(t), 300); }, ms || 2600);
}

// ---------------------------------------------------------------- popup menus
let openMenus = [];
export function closeMenus(level) {
  level = level || 0;
  while (openMenus.length > level) { const m = openMenus.pop(); if (m.el.parentNode) m.el.parentNode.removeChild(m.el); if (m.onClose) m.onClose(); }
}
document.addEventListener('mousedown', e => { if (!openMenus.length) return; if (!e.target.closest || !e.target.closest('.menu,.keep-menu')) closeMenus(); }, true);
document.addEventListener('touchstart', e => { if (!openMenus.length) return; if (!e.target.closest || !e.target.closest('.menu,.keep-menu')) closeMenus(); }, true);
document.addEventListener('keydown', e => { if (openMenus.length && keyName(e) === 'Escape') { closeMenus(); e.stopPropagation(); } }, true);
window.addEventListener('blur', () => closeMenus());

function placeAt(el, x, y, anchorRect) {
  document.body.appendChild(el);
  const W = window.innerWidth, H = window.innerHeight;
  const r = el.getBoundingClientRect();
  if (anchorRect) {
    x = anchorRect.left; y = anchorRect.bottom;
    if (y + r.height > H - 4 && anchorRect.top - r.height > 4) y = anchorRect.top - r.height;
  }
  x = clamp(x, 4, Math.max(4, W - r.width - 4));
  y = clamp(y, 4, Math.max(4, H - r.height - 4));
  el.style.left = x + 'px'; el.style.top = y + 'px';
  if (r.height > H - 8) { el.style.maxHeight = (H - 8) + 'px'; el.style.top = '4px'; }
}

// items: [{label, icon, run, checked, disabled, shortcut, sub:[...]|fn, sep:true, header:'..', html:Element}]
export function showMenu(items, opts) {
  opts = opts || {};
  const level = opts.level || 0;
  closeMenus(level);
  const menu = h('div.menu' + (opts.cls ? '.' + opts.cls : ''));
  if (opts.minWidth) menu.style.minWidth = opts.minWidth + 'px';
  (typeof items === 'function' ? items() : items).forEach(it => {
    if (!it) return;
    if (it.sep) { menu.appendChild(h('div.menu-sep')); return; }
    if (it.header) { menu.appendChild(h('div.menu-header', it.header)); return; }
    if (it.el) { menu.appendChild(it.el); return; }
    const row = h('div.menu-item' + (it.disabled ? '.disabled' : '') + (it.checked ? '.checked' : ''),
      h('span.mi-ic', it.checked === true ? ic('check', 15) : it.icon ? ic(it.icon, 15) : it.swatch ? h('span.swatch', { style: { background: it.swatch } }) : ''),
      h('span.mi-label', it.label),
      it.shortcut ? h('span.mi-sc', it.shortcut) : null,
      it.sub ? h('span.mi-arrow', ic('chevron-right', 14)) : null);
    if (it.title) row.title = it.title;
    if (it.style) row.querySelector('.mi-label').setAttribute('style', it.style);
    const openSub = () => {
      const rr = row.getBoundingClientRect();
      showMenu(typeof it.sub === 'function' ? it.sub() : it.sub, { level: level + 1, x: rr.right - 2, y: rr.top - 4 });
    };
    if (it.sub) { row.addEventListener('mouseenter', openSub); row.addEventListener('click', e => { e.stopPropagation(); openSub(); }); }
    else {
      row.addEventListener('mouseenter', () => closeMenus(level + 1));
      row.addEventListener('click', e => {
        e.stopPropagation();
        if (it.disabled) return;
        if (!it.keepOpen) closeMenus();
        if (it.run) setTimeout(() => it.run(e), 0);
      });
    }
    menu.appendChild(row);
  });
  menu.addEventListener('mousedown', e => { if (!e.target.closest('input,textarea,select')) e.preventDefault(); });
  placeAt(menu, opts.x || 0, opts.y || 0, opts.anchor ? opts.anchor.getBoundingClientRect() : null);
  openMenus.push({ el: menu, onClose: opts.onClose });
  return menu;
}
// panel popup (custom content)
export function showPopup(content, opts) {
  opts = opts || {};
  closeMenus(opts.level || 0);
  const p = h('div.menu.popup' + (opts.cls ? '.' + opts.cls : ''), content);
  p.addEventListener('mousedown', e => { if (!e.target.closest('input,textarea,select,[contenteditable]')) e.preventDefault(); });
  placeAt(p, opts.x || 0, opts.y || 0, opts.anchor ? opts.anchor.getBoundingClientRect() : null);
  openMenus.push({ el: p, onClose: opts.onClose });
  return p;
}
export function contextMenu(e, items) {
  e.preventDefault();
  const p = e.touches ? e.touches[0] : e;
  showMenu(items, { x: p.clientX, y: p.clientY, cls: 'ctx' });
}
// long-press -> contextmenu on touch devices
export function longPress(el, fn) {
  let t = null, sx = 0, sy = 0;
  el.addEventListener('touchstart', e => {
    const p = e.touches[0]; sx = p.clientX; sy = p.clientY;
    t = setTimeout(() => { t = null; fn({ clientX: sx, clientY: sy, target: e.target, preventDefault() {}, touches: null }); }, 550);
  }, { passive: true });
  const cancel = e => { if (!t) return; if (e && e.touches && e.touches[0]) { const p = e.touches[0]; if (Math.abs(p.clientX - sx) < 8 && Math.abs(p.clientY - sy) < 8) return; } clearTimeout(t); t = null; };
  el.addEventListener('touchmove', cancel, { passive: true });
  el.addEventListener('touchend', () => { if (t) { clearTimeout(t); t = null; } });
}

// ---------------------------------------------------------------- dialogs
let dlgStack = [];
export function dialog(opts) {
  // opts: {title, body: Element|string, buttons:[{label, primary, value, run}], width, onOpen, cls}
  return new Promise(resolve => {
    const back = h('div.dlg-back');
    const close = v => { if (back.parentNode) back.parentNode.removeChild(back); dlgStack = dlgStack.filter(d => d !== back); document.removeEventListener('keydown', onKey, true); resolve(v); };
    const body = typeof opts.body === 'string' ? h('div', { html: opts.body }) : opts.body;
    const btns = (opts.buttons || [{ label: 'OK', primary: true, value: true }]).map(b => {
      const el = h('button.btn' + (b.primary ? '.primary' : '') + (b.danger ? '.danger' : ''), b.label);
      el.addEventListener('click', async () => {
        if (b.run) { const r = await b.run(body, close); if (r === false) return; if (r !== undefined && r !== true) { close(r); return; } }
        close(b.value !== undefined ? b.value : b.label);
      });
      return el;
    });
    const dlg = h('div.dlg' + (opts.cls ? '.' + opts.cls : ''), { role: 'dialog' },
      h('div.dlg-head', h('div.dlg-title', opts.title || ''), h('button.dlg-x', { title: 'Đóng (Esc)', onclick: () => close(null) }, ic('x', 16))),
      h('div.dlg-body', body),
      btns.length ? h('div.dlg-foot', btns) : null);
    if (opts.width) { dlg.style.width = opts.width + 'px'; dlg.style.maxWidth = '96vw'; }
    back.appendChild(dlg);
    back.addEventListener('mousedown', e => { if (e.target === back && opts.modeless !== true) {/* keep modal */} });
    document.body.appendChild(back);
    dlgStack.push(back);
    // drag dialog by header (mouse)
    const head = dlg.querySelector('.dlg-head');
    head.addEventListener('mousedown', e => {
      if (e.target.closest('button')) return;
      const r = dlg.getBoundingClientRect(); const ox = e.clientX - r.left, oy = e.clientY - r.top;
      dlg.style.position = 'fixed'; dlg.style.margin = '0';
      const mv = ev => { dlg.style.left = clamp(ev.clientX - ox, 0, window.innerWidth - 60) + 'px'; dlg.style.top = clamp(ev.clientY - oy, 0, window.innerHeight - 40) + 'px'; };
      const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); };
      document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
    });
    function onKey(e) {
      if (dlgStack[dlgStack.length - 1] !== back) return;
      const k = keyName(e);
      if (k === 'Escape') { e.preventDefault(); e.stopPropagation(); close(null); }
      else if (k === 'Enter' && !e.target.closest('textarea,[contenteditable]') && opts.enterKey !== false) {
        const p = btns.find(b => b.classList.contains('primary')); if (p) { e.preventDefault(); p.click(); }
      }
    }
    document.addEventListener('keydown', onKey, true);
    if (opts.onOpen) opts.onOpen(body, close);
    setTimeout(() => { const f = body.querySelector('[autofocus],input:not([type=checkbox]):not([type=radio]),select,textarea'); if (f) { try { f.focus(); if (f.select) f.select(); } catch (e) {} } }, 30);
  });
}
export function alertDlg(msg, title) { return dialog({ title: title || 'Yomi-Office Open Source', body: h('div.dlg-msg', msg), buttons: [{ label: 'OK', primary: true, value: true }] }); }
export function confirmDlg(msg, title, yes, no) { return dialog({ title: title || 'Xác nhận', body: h('div.dlg-msg', msg), buttons: [{ label: yes || 'Đồng ý', primary: true, value: true }, { label: no || 'Hủy', value: false }] }).then(v => v === true); }
export function promptDlg(label, value, title, opts) {
  opts = opts || {};
  const inp = h(opts.multiline ? 'textarea.inp' : 'input.inp', { value: value == null ? '' : value, type: opts.type || 'text', rows: opts.rows || 4 });
  if (opts.multiline) inp.value = value || '';
  return dialog({ title: title || 'Yomi-Office Open Source', width: opts.width || 420, body: h('div.form', h('label.lbl', label), inp, opts.hint ? h('div.hint', opts.hint) : null),
    buttons: [{ label: 'OK', primary: true, run: () => inp.value }, { label: 'Hủy', value: null }] });
}
// form builder: fields [{key,label,type:'text|number|select|checkbox|color|textarea|radio',options:[[v,l]],value,min,max,step}]
export function formDlg(title, fields, opts) {
  opts = opts || {};
  const inputs = {};
  const body = h('div.form' + (opts.cols ? '.cols' : ''), fields.map(f => {
    if (f.type === 'html') return f.el;
    if (f.type === 'section') return h('div.form-sec', f.label);
    let el;
    if (f.type === 'select') { el = h('select.inp', (f.options || []).map(o => h('option', { value: o[0] }, o[1]))); el.value = f.value; }
    else if (f.type === 'checkbox') { el = h('input', { type: 'checkbox' }); el.checked = !!f.value; return h('label.chk', el, h('span', f.label), (inputs[f.key] = el) && null); }
    else if (f.type === 'textarea') { el = h('textarea.inp', { rows: f.rows || 4 }); el.value = f.value || ''; }
    else if (f.type === 'radio') {
      const wrap = h('div.radios'); el = wrap;
      (f.options || []).forEach(o => { const r = h('input', { type: 'radio', name: 'r_' + f.key, value: o[0] }); if (String(o[0]) === String(f.value)) r.checked = true; wrap.appendChild(h('label.chk', r, h('span', o[1]))); });
      el.getValue = () => { const c = wrap.querySelector('input:checked'); return c ? c.value : null; };
    }
    else { el = h('input.inp', { type: f.type || 'text', value: f.value == null ? '' : f.value }); if (f.min != null) el.min = f.min; if (f.max != null) el.max = f.max; if (f.step != null) el.step = f.step; if (f.placeholder) el.placeholder = f.placeholder; }
    inputs[f.key] = el;
    return h('div.form-row', h('label.lbl', f.label), el, f.hint ? h('div.hint', f.hint) : null);
  }));
  return dialog({ title, width: opts.width || 460, body, onOpen: opts.onOpen && (() => opts.onOpen(inputs)), buttons: [{ label: opts.ok || 'OK', primary: true, run: () => {
    const out = {};
    for (const k in inputs) { const el = inputs[k]; out[k] = el.getValue ? el.getValue() : el.type === 'checkbox' ? el.checked : el.type === 'number' ? (el.value === '' ? null : +el.value) : el.value; }
    if (opts.validate) { const err = opts.validate(out); if (err) { toast(err, 'err'); return false; } }
    return out;
  } }, { label: 'Hủy', value: null }] });
}

// ---------------------------------------------------------------- color picker
export const THEME_COLORS = [
  ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'],
  ['#F2F2F2', '#7F7F7F', '#D0CECE', '#D6DCE4', '#D9E2F3', '#FBE5D5', '#EDEDED', '#FFF2CC', '#DEEBF6', '#E2EFD9'],
  ['#D8D8D8', '#595959', '#AEABAB', '#ADB9CA', '#B4C6E7', '#F7CBAC', '#DBDBDB', '#FEE599', '#BDD7EE', '#C5E0B3'],
  ['#BFBFBF', '#3F3F3F', '#757070', '#8496B0', '#8EAADB', '#F4B183', '#C9C9C9', '#FFD965', '#9DC3E6', '#A8D08D'],
  ['#A5A5A5', '#262626', '#3A3838', '#323F4F', '#2F5496', '#C55A11', '#7B7B7B', '#BF8F00', '#2E75B5', '#538135'],
  ['#7F7F7F', '#0C0C0C', '#171616', '#222A35', '#1F3864', '#833C0B', '#525252', '#7F6000', '#1E4E79', '#375623']
];
export const STD_COLORS = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];
const recentColors = [];
export function colorPicker(anchor, onPick, opts) {
  opts = opts || {};
  const grid = h('div.cp');
  const sw = c => { const s = h('button.cp-sw', { title: c, style: { background: c } }); s.addEventListener('click', () => { pick(c); }); return s; };
  const pick = c => { closeMenus(); if (c && recentColors.indexOf(c) < 0) { recentColors.unshift(c); recentColors.length = Math.min(recentColors.length, 10); } onPick(c); };
  if (opts.auto) grid.appendChild(h('button.cp-auto', { onclick: () => pick(null) }, h('span.swatch', { style: { background: opts.autoColor || '#000' } }), opts.auto));
  grid.appendChild(h('div.cp-h', 'Màu chủ đề (Theme Colors)'));
  grid.appendChild(h('div.cp-row', THEME_COLORS[0].map(sw)));
  grid.appendChild(h('div.cp-col', THEME_COLORS.slice(1).map(r => h('div.cp-row.tight', r.map(sw)))));
  grid.appendChild(h('div.cp-h', 'Màu chuẩn (Standard Colors)'));
  grid.appendChild(h('div.cp-row', STD_COLORS.map(sw)));
  if (recentColors.length) { grid.appendChild(h('div.cp-h', 'Màu gần đây')); grid.appendChild(h('div.cp-row', recentColors.map(sw))); }
  if (opts.none) grid.appendChild(h('button.cp-auto', { onclick: () => pick('transparent') }, ic('ban', 14), opts.none));
  const more = h('input', { type: 'color', value: '#4472c4' });
  more.addEventListener('change', () => pick(more.value));
  grid.appendChild(h('label.cp-more', ic('palette', 14), h('span', 'Màu khác (More Colors)...'), more));
  showPopup(grid, { anchor });
}

// ---------------------------------------------------------------- tooltip (rich, delayed)
let tipEl = null, tipTimer = null;
export function richTip(el, title, desc, sc) {
  el.addEventListener('mouseenter', () => {
    clearTimeout(tipTimer);
    tipTimer = setTimeout(() => {
      hideTip();
      tipEl = h('div.tip', h('div.tip-t', title + (sc ? ' (' + sc + ')' : '')), desc ? h('div.tip-d', desc) : null);
      document.body.appendChild(tipEl);
      const r = el.getBoundingClientRect(), tr = tipEl.getBoundingClientRect();
      tipEl.style.left = clamp(r.left, 4, window.innerWidth - tr.width - 4) + 'px';
      tipEl.style.top = (r.bottom + 6 + tr.height > window.innerHeight ? r.top - tr.height - 6 : r.bottom + 6) + 'px';
    }, 650);
  });
  el.addEventListener('mouseleave', hideTip);
  el.addEventListener('mousedown', hideTip);
}
function hideTip() { clearTimeout(tipTimer); if (tipEl && tipEl.parentNode) tipEl.parentNode.removeChild(tipEl); tipEl = null; }

// ---------------------------------------------------------------- Ribbon
// def: {tabs:[{id,label,en,groups:[{label, items:[...]}], contextual?:fn}], fileLabel, onFile}
// item: {type:'big'|'btn'|'split'|'drop'|'toggle'|'gallery'|'combo'|'col'|'row'|'color'|'sep'|'custom', id,label,icon,title,desc,sc,run(),menu(),active(),disabled(),items:[], width, options()}
let _mctx = null;
function textWidth(t, font) { try { if (!_mctx) _mctx = document.createElement('canvas').getContext('2d'); _mctx.font = font; return Math.ceil(_mctx.measureText(t || '').width); } catch (e) { return (t || '').length * 6; } }
export class Ribbon {
  constructor(app, def) {
    try { (window.__oosRibbons = window.__oosRibbons || []).push(this); } catch (e) {}   // v0.1.7: lets the build export the Web ribbon layout
    this.app = app; this.def = def; this.el = h('div.ribbon'); this.tabsEl = h('div.rb-tabs'); this.bodyEl = h('div.rb-body');
    this.collapsed = false; this.cur = (def.tabs[0] && def.tabs[0].id) || 'home'; this.updaters = [];
    this.el.appendChild(this.tabsEl); this.el.appendChild(this.bodyEl);
    this.renderTabs(); this.renderBody();
  }
  renderTabs() {
    this.tabsEl.innerHTML = '';
    const file = h('button.rb-tab.rb-file', this.def.fileLabel || 'Tệp');
    file.addEventListener('click', () => this.def.onFile && this.def.onFile());
    this.tabsEl.appendChild(file);
    this.def.tabs.forEach(t => {
      if (t.hidden && t.hidden()) return;
      const b = h('button.rb-tab' + (t.id === this.cur ? '.on' : '') + (t.contextual ? '.ctxl' : ''), { dataset: { tab: t.id } }, t.label);
      if (t.en) b.title = t.en;
      b.addEventListener('click', () => {
        if (this.cur === t.id && !this.collapsed) { if (this.def.collapseOnReclick) this.toggleCollapse(); return; }
        this.select(t.id); if (this.collapsed) this.toggleCollapse(false);
      });
      b.addEventListener('dblclick', () => this.toggleCollapse());
      this.tabsEl.appendChild(b);
    });
    const tail = h('div.rb-tail');
    if (this.def.tail) this.def.tail(tail);
    tail.appendChild(h('button.rb-collapse', { title: 'Thu gọn/mở rộng Ribbon (Ctrl+F1)', onclick: () => this.toggleCollapse() }, ic(this.collapsed ? 'chevron-down' : 'chevron-up', 15)));
    this.tabsEl.appendChild(tail);
  }
  toggleCollapse(v) { this.collapsed = v === undefined ? !this.collapsed : v; this.el.classList.toggle('collapsed', this.collapsed); this.renderTabs(); if (this.app.onLayout) this.app.onLayout(); }
  select(id) { this.cur = id; $$('.rb-tab', this.tabsEl).forEach(b => b.classList.toggle('on', b.getAttribute('data-tab') === id)); this.renderBody(); }
  renderBody() {
    this.bodyEl.innerHTML = ''; this.updaters = [];
    const tab = this.def.tabs.find(t => t.id === this.cur) || this.def.tabs[0];
    const strip = h('div.rb-strip');
    (typeof tab.groups === 'function' ? tab.groups() : tab.groups).forEach(g => {
      if (!g) return;
      const items = h('div.rb-items');
      g.items.forEach(it => { const e = this.item(it); if (e) items.appendChild(e); });
      // v0.1.7: like the native ribbon, a group that does not fit collapses into one button (group icon + ▾) with a drop-down panel
      const gicon = g.icon || (function find(list) { for (const x of list || []) { if (!x) continue; if (x.icon) return x.icon; const r = find(Array.isArray(x.items) ? x.items : null); if (r) return r; } return null; })(g.items) || 'n:more';
      const cbtn = h('button.rb-cbtn', { type: 'button', title: g.label }, ic(gicon, 32), ic('chevron-down', 11, 'dd'));
      cbtn.addEventListener('mousedown', e => e.preventDefault());
      cbtn.addEventListener('click', () => this.popGroup(grp, items, cbtn));
      const grp = h('div.rb-group', cbtn, items, h('div.rb-glabel', g.label, g.launcher ? h('button.rb-launch', { title: g.label + '…', onclick: () => g.launcher() }, ic('square-arrow-out-down-right', 11)) : null));
      strip.appendChild(grp);
    });
    this.bodyEl.appendChild(strip);
    this.fitSoon();
    if (!this._fitBound) { this._fitBound = true; window.addEventListener('resize', () => this.fitSoon()); }
    // horizontal scroll by wheel on narrow screens
    this.bodyEl.onwheel = e => { if (this.bodyEl.scrollWidth > this.bodyEl.clientWidth && Math.abs(e.deltaY) > Math.abs(e.deltaX)) { this.bodyEl.scrollLeft += e.deltaY; e.preventDefault(); } };
    this.refresh();
  }
  item(it) {
    const app = this.app;
    if (it.type === 'sep') return h('div.rb-sep');
    if (it.type === 'col' || it.type === 'row') { const c = h('div.rb-' + it.type); it.items.forEach(x => { const e = this.item(x); if (e) c.appendChild(e); }); return c; }
    if (it.type === 'custom') return it.render(app, this);
    if (it.type === 'combo') return this.combo(it);
    if (it.type === 'gallery') return this.gallery(it);
    if (it.type === 'toggle') return this.item({ type: it.big ? 'big' : 'btn', big: it.big, compact: it.compact, label: it.label, icon: it.icon, title: it.title, sc: it.sc, active: () => it.get(this.app), run: () => { it.set(!it.get(this.app), this.app); this.refresh(); } }); // v0.0.8 fix: setter receives the new value (was the document → stored a circular reference)
    const big = it.type === 'big' || it.big;
    const b = h('button.rb-btn' + (big ? '.big' : '') + (it.label && !big && it.compact ? '.compact' : '') + (!it.label || it.iconOnly ? '.icon-only' : ''), { type: 'button' });
    const icn = it.icon ? ic(it.icon, big ? 32 : 16, it.iconCls) : null;   // v0.1.7: native sizes (32 / 16)
    if (it.colorBar) { const wrap = h('span.ic-color', icn, h('span.cbar', { style: { background: it.colorBar() } })); b.appendChild(wrap); this.updaters.push(() => { const bar = wrap.querySelector('.cbar'); bar.style.background = it.colorBar(); }); }
    else if (icn) b.appendChild(icn);
    if (it.label && !it.iconOnly) b.appendChild(h('span.rb-lbl', it.label));
    const hasMenu = it.menu || it.type === 'drop' || it.type === 'color';
    const isSplit = it.type === 'split' || it.type === 'color';
    richTip(b, it.title || it.label || '', it.desc, it.sc);
    if (isSplit) {
      const wrap = h('div.rb-split' + (big ? '.big' : ''));
      const arrow = h('button.rb-arrow', { type: 'button', title: (it.title || it.label || '') + ' — tùy chọn' }, ic('chevron-down', 12));
      b.addEventListener('click', () => { if (!this.isDisabled(it)) { it.run && it.run(app); app.afterCommand && app.afterCommand(); } });
      arrow.addEventListener('click', () => {
        if (it.type === 'color') colorPicker(wrap, c => { it.pick(app, c); app.afterCommand && app.afterCommand(); this.refresh(); }, it.colorOpts);
        else showMenu(it.menu(app), { anchor: wrap });
      });
      wrap.appendChild(b); wrap.appendChild(arrow);
      b.addEventListener('mousedown', e => e.preventDefault()); arrow.addEventListener('mousedown', e => e.preventDefault());
      this.bindState(it, wrap, b);
      return wrap;
    }
    if (hasMenu) { b.appendChild(ic('chevron-down', 11, 'dd')); b.addEventListener('click', () => { if (!this.isDisabled(it)) showMenu(it.menu(app), { anchor: b, minWidth: it.menuWidth }); }); }
    else b.addEventListener('click', () => { if (!this.isDisabled(it)) { it.run && it.run(app); app.afterCommand && app.afterCommand(); this.refresh(); } });
    b.addEventListener('mousedown', e => e.preventDefault());
    if (it.id) b.setAttribute('data-cmd', it.id);
    this.bindState(it, b, b);
    return b;
  }
  isDisabled(it) { return it.disabled ? it.disabled(this.app) : false; }
  bindState(it, el, b) {
    if (it.active || it.disabled) this.updaters.push(() => {
      if (it.active) b.classList.toggle('on', !!it.active(this.app));
      if (it.disabled) el.classList.toggle('disabled', !!it.disabled(this.app));
    });
  }
  combo(it) {
    const inp = h('input.rb-combo-inp', { type: 'text', spellcheck: false });
    inp.style.width = (it.width || 110) + 'px';
    const arrow = h('button.rb-combo-arr', { type: 'button', tabindex: -1 }, ic('chevron-down', 12));
    const wrap = h('div.rb-combo', inp, arrow);
    if (it.title) wrap.title = it.title;
    const commit = () => { if (inp.value !== inp._last) { if (it.fontBox) import('./fonts.js').then(F => F.ensureFont(inp.value)); it.set(this.app, inp.value); this.app.afterCommand && this.app.afterCommand(); } };
    inp.addEventListener('focus', () => { inp._last = inp.value; setTimeout(() => inp.select(), 0); });
    inp.addEventListener('keydown', e => { const k = keyName(e); if (k === 'Enter') { e.preventDefault(); commit(); this.app.focus && this.app.focus(); } if (k === 'Escape') { inp.value = inp._last; this.app.focus && this.app.focus(); } });
    inp.addEventListener('change', commit);
    arrow.addEventListener('mousedown', e => e.preventDefault());
    arrow.addEventListener('click', () => {
      const opts = it.options(this.app);
      if (it.fontBox) import('./fonts.js').then(F => opts.slice(0, 60).forEach(f => F.ensureFont(f)));
      const items = opts.map(o => { const v = Array.isArray(o) ? o[0] : o, l = Array.isArray(o) ? o[1] : o; return { label: l, style: it.optStyle ? it.optStyle(v) : null, checked: String(v) === inp.value ? true : undefined, run: () => { inp.value = v; if (it.fontBox) import('./fonts.js').then(F => F.ensureFont(v)); it.set(this.app, v); this.app.afterCommand && this.app.afterCommand(); this.refresh(); } }; });
      if (it.fontBox) items.push({ sep: true }, { label: 'Quản lý phông chữ / Nạp Font Pack…', icon: 'type', run: () => import('./fonts.js').then(F => F.fontManager()) });
      showMenu(items, { anchor: wrap, minWidth: Math.max(160, it.width || 110), cls: 'combo-menu' });
    });
    this.updaters.push(() => { if (document.activeElement !== inp) inp.value = it.get(this.app); });
    return wrap;
  }
  gallery(it) {
    const box = h('div.rb-gallery');
    const list = h('div.rb-gal-items');
    const render = () => {
      list.innerHTML = '';
      it.items(this.app).forEach(x => {
        const b = h('button.rb-gal-item', { title: x.title || x.label, type: 'button' }, x.preview ? x.preview() : h('span', x.label));
        b.addEventListener('mousedown', e => e.preventDefault());
        b.addEventListener('click', () => { x.run(this.app); this.app.afterCommand && this.app.afterCommand(); this.refresh(); });
        if (x.active && x.active(this.app)) b.classList.add('on');
        list.appendChild(b);
      });
    };
    render();
    box.appendChild(list);
    if (it.more) { const m = h('button.rb-gal-more', { type: 'button', title: 'Thêm' }, ic('chevron-down', 12)); m.addEventListener('mousedown', e => e.preventDefault()); m.addEventListener('click', () => showMenu(it.more(this.app), { anchor: box })); box.appendChild(m); }
    if (it.width) box.style.width = it.width + 'px';
    this.updaters.push(render);
    return box;
  }
  refresh() { this.updaters.forEach(f => { try { f(); } catch (e) {} }); }
  fitSoon() { clearTimeout(this._fitT); this._fitT = setTimeout(() => this.fit(), 0); }
  // v0.1.7: same rule as Ribbon.LayoutAll() of the native editor - collapse groups from the right until the tab fits
  fit() {
    const strip = this.bodyEl.querySelector('.rb-strip'); if (!strip) return;
    if (!this.bodyEl.clientWidth) { this._fitTry = (this._fitTry || 0) + 1; if (this._fitTry < 50) setTimeout(() => this.fit(), 100); return; }
    this._fitTry = 0;
    const groups = Array.prototype.slice.call(strip.children).filter(x => x.classList.contains('rb-group'));
    groups.forEach(g => { g.classList.remove('collapsed'); g.style.minWidth = ''; });
    const avail = this.bodyEl.clientWidth - 50; let total = 0; const w = groups.map(g => { const x = g.offsetWidth + 1; total += x; return x; });
    for (let k = groups.length - 1; k >= 0 && total > avail; k--) {
      const lbl = groups[k].querySelector('.rb-glabel'); const cw = Math.max(56, textWidth(lbl ? lbl.firstChild && lbl.firstChild.nodeType === 3 ? lbl.firstChild.nodeValue : lbl.textContent : '', '11px Segoe UI, Roboto, Arial, sans-serif') + 16);
      if (cw >= w[k]) continue;
      groups[k].classList.add('collapsed'); groups[k].style.minWidth = cw + 'px'; total -= w[k] - cw;
    }
  }
  popGroup(grp, items, btn) {
    if (this._pop) { this._pop.close(); return; }
    const r = btn.getBoundingClientRect();
    const pop = h('div.rb-pop.ribbon'); const inner = h('div.rb-group'); inner.appendChild(items); pop.appendChild(inner);
    pop.style.left = Math.max(4, Math.min(window.innerWidth - 20, r.left)) + 'px'; pop.style.top = (r.bottom + 2) + 'px';
    document.body.appendChild(pop);
    const close = () => { document.removeEventListener('mousedown', out, true); if (pop.parentNode) pop.parentNode.removeChild(pop); grp.insertBefore(items, grp.querySelector('.rb-glabel')); this._pop = null; };
    const out = e => { if (!pop.contains(e.target) && e.target !== btn && !btn.contains(e.target) && !(e.target.closest && e.target.closest('.menu'))) close(); };
    setTimeout(() => document.addEventListener('mousedown', out, true), 0);
    this._pop = { close };
    const rr = pop.getBoundingClientRect(); if (rr.right > window.innerWidth - 4) pop.style.left = Math.max(4, window.innerWidth - rr.width - 4) + 'px';
  }
}

// ---------------------------------------------------------------- zoom slider (status bar)
export function zoomControl(get, set) {
  const lbl = h('button.zoom-lbl', { title: 'Thu phóng (Zoom)…' });
  const range = h('input.zoom-range', { type: 'range', min: 10, max: 400, step: 5 });
  const minus = h('button.zoom-b', { title: 'Thu nhỏ' }, '−'), plus = h('button.zoom-b', { title: 'Phóng to' }, '+');
  const upd = () => { const z = get(); lbl.textContent = Math.round(z) + '%'; range.value = z; };
  range.addEventListener('input', () => { set(+range.value); upd(); });
  minus.addEventListener('click', () => { set(Math.max(10, Math.round((get() - 10) / 10) * 10)); upd(); });
  plus.addEventListener('click', () => { set(Math.min(400, Math.round((get() + 10) / 10) * 10)); upd(); });
  lbl.addEventListener('click', () => {
    showMenu([200, 150, 125, 100, 75, 50, 25].map(z => ({ label: z + '%', checked: Math.round(get()) === z || undefined, run: () => { set(z); upd(); } })), { anchor: lbl });
  });
  upd();
  return { el: h('div.zoom', minus, range, plus, lbl), update: upd };
}

export function segmented(options, value, onChange) {
  const wrap = h('div.seg');
  options.forEach(o => { const b = h('button.seg-b' + (o[0] === value ? '.on' : ''), o[1]); b.addEventListener('click', () => { $$('.seg-b', wrap).forEach(x => x.classList.remove('on')); b.classList.add('on'); onChange(o[0]); }); wrap.appendChild(b); });
  return wrap;
}

export function spinner(text) { return h('div.spin-wrap', h('div.spinner'), text ? h('div.spin-t', text) : null); }
export function busy(text) {
  const b = h('div.busy', spinner(text));
  document.body.appendChild(b);
  return { set(t) { const s = b.querySelector('.spin-t'); if (s) s.textContent = t; }, done() { if (b.parentNode) b.parentNode.removeChild(b); } };
}
