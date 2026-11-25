// Excel document: workbook view, formula bar, sheet tabs, editing, undo/redo, clipboard, status bar
import { h, ic, $, $$, esc, keyName, chord, clamp, debounce, uid, utf8Encode, utf8Decode, extOf, deepClone } from '../core/util.js';
import { Doc, registerApp } from '../core/shell.js';
import { toast, showMenu, contextMenu, dialog, promptDlg, confirmDlg, formDlg, colorPicker, longPress, busy } from '../core/ui.js';
import { settings } from '../core/store.js';
import { Workbook, Sheet, Err, E, key, colName, addr, rangeAddr, parseRangeAddr, norm, inRange, internStyle, mergeStyle, wbToJSON, wbFromJSON, MAX_ROWS, MAX_COLS, DEF_COL_W, DEF_ROW_H, isErr } from './model.js';
import { Calc, FN, shiftFormula, adjustFormulaForInsert, parse, str as fstr, num as fnum } from './formula.js';
import { FUNCTION_META, functionCatalog } from './functions.js';
import './functions2.js';
import { format, parseInput, isDateFormat } from './numfmt.js';
import { Grid, PT2PX } from './grid.js';
import { ribbonDef as excelRibbon } from './ribbon.js';
import { unifyRibbon } from '../core/unify.js';
import * as X from './cmds.js';
import { readXlsx, writeXlsx, readCsv, writeCsv, sheetToHtml } from './xlsx.js';
import { Objects } from './objects.js';
import { templates as xlTemplates, applyTemplate } from './templates.js';

export class ExcelDoc extends Doc {
  static get kind() { return 'excel'; }
  static templates() { return xlTemplates(); }
  constructor(shell, opts) {
    super(shell, opts);
    this.wb = new Workbook(); this.calc = new Calc(this.wb);
    this.undoStack = []; this.redoStack = []; this.clip = null; this.fnMeta = FUNCTION_META;
    this.statFns = settings().xlStatFns || ['avg', 'count', 'sum'];
    this.buildUi();
  }
  async init() {
    const t = this.opts.template;
    this.wb.addSheet('Sheet1');
    if (t && t.id !== 'blank') await applyTemplate(this, t);
    this.afterLoad();
  }
  async load(bytes, ext, name) {
    this.wb.sheets = []; this.wb.active = 0; // replace, never append to an existing workbook
    if (ext === 'csv' || ext === 'tsv') readCsv(this.wb, bytes, ext === 'tsv' ? '\t' : null, name);
    else if (ext === 'xls' || ext === 'xlt') { if (bytes[0] === 0xD0 && bytes[1] === 0xCF) { const { readXls } = await import('../formats/legacy.js'); const M = await import('./model.js'); readXls(this.wb, bytes, { internStyle: M.internStyle }); } else if (bytes[0] === 0x50 && bytes[1] === 0x4B) await readXlsx(this.wb, bytes); else { const t = new TextDecoder('utf-8').decode(bytes); if (/<Workbook|<\?mso-application/.test(t)) { const { readXml2003 } = await import('../formats/extra.js'); const M = await import('./model.js'); readXml2003(this.wb, t, M); } else readCsv(this.wb, bytes, '\t', name); } }
    else if (ext === 'xml') { const t = new TextDecoder('utf-8').decode(bytes); const { readXml2003 } = await import('../formats/extra.js'); const M = await import('./model.js'); readXml2003(this.wb, t, M); }
    else if (ext === 'ods' || ext === 'ots') { const { readOds } = await import('../formats/odf.js'); const M = await import('./model.js'); await readOds(this.wb, bytes, { internStyle: M.internStyle, colIndex: M.colIndex, dateToSerial: M.dateToSerial }); }
    else await readXlsx(this.wb, bytes);
    if (!this.wb.sheets.length) this.wb.addSheet();
    this.afterLoad();
  }
  async restore(native) { this.wb = wbFromJSON(native); this.calc = new Calc(this.wb); this.afterLoad(); }
  snapshot() { return wbToJSON(this.wb); }
  afterLoad() {
    this.wb.props.name = this.name;
    this.recalc();
    this.zoom = this.wb.cur.zoom || 100;
    this.renderTabs();
  }
  canSaveExt(ext) { return ['xlsx', 'xlsm', 'csv', 'tsv', 'ods', 'xltx', 'xml2003', 'txtu', 'slk', 'dif', 'mht'].indexOf(ext) >= 0; }
  saveFormats() { return [['xlsx', 'Sổ làm việc Excel', 'sheet'], ['xlsm', 'Sổ làm việc Excel có Macro (*.xlsm)', 'square-code'], ['csv', 'CSV UTF-8 (phân cách bằng dấu phẩy)', 'file-spreadsheet'], ['tsv', 'Văn bản phân cách Tab', 'file-text'], ['html', 'Trang web (HTML)', 'globe'], ['json', 'Dữ liệu JSON', 'braces'], ['ods', 'Bảng tính OpenDocument (*.ods)', 'sheet'], ['xltx', 'Mẫu Excel (*.xltx)', 'file-plus'], ['xml2003', 'Bảng tính XML 2003 (*.xml)', 'file-code', 'xml'], ['txtu', 'Văn bản Unicode (*.txt)', 'file-text', 'txt'], ['slk', 'SYLK — Symbolic Link (*.slk)', 'file'], ['dif', 'DIF — Data Interchange Format (*.dif)', 'file'], ['mht', 'Trang web một tệp (*.mht)', 'globe']]; }
  async save(ext) {
    if (this.grid.editing) this.grid.commit();
    for (const s of this.wb.sheets) { if (s._lazyLoad) { try { await s._lazyLoad(); delete s._lazyLoad; } catch (e) {} } }
    if (ext === 'xlsx') { if (this.wb.vba && this.wb.vba.modules && this.wb.vba.modules.length && !this._vbaWarned) { this._vbaWarned = true; toast('Sổ làm việc có macro VBA — định dạng .xlsx không lưu macro. Hãy dùng "Lưu như" > .xlsm để giữ macro.', 'err', 6000); } return writeXlsx(this.wb); }
    if (ext === 'xlsm') return writeXlsx(this.wb, { macro: true });
    if (ext === 'xltx') { const X = await import('../formats/extra.js'); return X.retypePackage(await writeXlsx(this.wb), X.CT.xlsx, X.CT.xltx); }
    if (ext === 'xml2003') { const X = await import('../formats/extra.js'); return X.writeXml2003(this.wb, await import('./model.js')); }
    if (ext === 'txtu') { const X = await import('../formats/extra.js'); return X.unicodeText(writeCsv(this.wb.cur, '\t')); }
    if (ext === 'slk') { const X = await import('../formats/extra.js'); return X.writeSylk(this.wb.cur, await import('./model.js')); }
    if (ext === 'dif') { const X = await import('../formats/extra.js'); return X.writeDif(this.wb.cur, await import('./model.js')); }
    if (ext === 'mht') { const X = await import('../formats/extra.js'); return utf8Encode(X.writeMht(this.wb.sheets.map(s => '<h2>' + esc(s.name) + '</h2>' + sheetToHtml(s, this)).join(''), this.name)); }
    if (ext === 'ods') { const { writeOds } = await import('../formats/odf.js'); const M = await import('./model.js'); return writeOds(this.wb, { colName: M.colName, isErr: M.isErr, serialToParts: M.serialToParts, display: (s, r, c) => this.displayText ? this.displayText(s, r, c) : s.value(r, c) }); }
    if (ext === 'csv') return utf8Encode('﻿' + writeCsv(this.wb.cur, ','));
    if (ext === 'tsv') return utf8Encode(writeCsv(this.wb.cur, '\t'));
    if (ext === 'html') return utf8Encode('<!doctype html><meta charset="utf-8"><title>' + esc(this.name) + '</title><style>body{font-family:Calibri,Arial,sans-serif} table{border-collapse:collapse} td{border:1px solid #ddd;padding:2px 6px}</style>' + this.wb.sheets.map(s => '<h2>' + esc(s.name) + '</h2>' + sheetToHtml(s, this)).join(''));
    if (ext === 'json') { const s = this.wb.cur; const u = s.used(); const rows = []; for (let r = u.r1; r <= u.r2; r++) { const row = []; for (let c = u.c1; c <= u.c2; c++) { const v = s.value(r, c); row.push(v instanceof Err ? v.code : v); } rows.push(row); } return utf8Encode(JSON.stringify({ sheet: s.name, data: rows }, null, 1)); }
    throw new Error('Định dạng không hỗ trợ');
  }

  // ---------------------------------------------------------------- UI
  buildUi() {
    // formula bar
    this.nameBox = h('input.xl-namebox', { spellcheck: false, title: 'Hộp Tên (Name Box): nhập địa chỉ hoặc tên để chuyển tới' });
    this.fbInput = h('textarea.xl-fbinput', { spellcheck: false, rows: 1, wrap: 'off' });
    const fx = h('button.xl-fx', { title: 'Chèn hàm (Shift+F3)' }, h('i', 'f'), h('sub', 'x'));
    this.fbCancel = h('button.xl-fbb', { title: 'Hủy (Esc)' }, ic('x', 15)); this.fbOk = h('button.xl-fbb', { title: 'Nhập (Enter)' }, ic('check', 15));
    const expand = h('button.xl-fbb.xl-fbexp', { title: 'Mở rộng thanh công thức' }, ic('chevron-down', 14));
    const nbDrop = h('button.xl-nbdrop', { title: 'Tên đã đặt' }, ic('chevron-down', 12));
    this.formulaBar = h('div.xl-fbar', h('div.xl-nb', this.nameBox, nbDrop), h('div.xl-fbsep'), this.fbCancel, this.fbOk, fx, h('div.xl-fbwrap', this.fbInput), expand);
    this.fnTipEl = h('div.xl-fnsig.hidden'); this.formulaBar.appendChild(this.fnTipEl);
    expand.addEventListener('click', () => { this.formulaBar.classList.toggle('tall'); expand.innerHTML = ''; expand.appendChild(ic(this.formulaBar.classList.contains('tall') ? 'chevron-up' : 'chevron-down', 14)); });
    fx.addEventListener('click', () => X.insertFunctionDialog(this));
    this.fbCancel.addEventListener('mousedown', e => e.preventDefault()); this.fbOk.addEventListener('mousedown', e => e.preventDefault());
    this.fbCancel.addEventListener('click', () => this.grid.cancel()); this.fbOk.addEventListener('click', () => this.grid.commit());
    nbDrop.addEventListener('click', () => { const names = Object.keys(this.wb.names); showMenu(names.length ? names.map(n => ({ label: n + '  →  ' + (this.wb.names[n].f || this.wb.names[n]), run: () => this.goto(n) })) : [{ label: 'Chưa có tên nào. Dùng Công thức > Xác định tên.', disabled: true }], { anchor: this.nameBox }); });
    this.nameBox.addEventListener('keydown', e => { if (keyName(e) === 'Enter') { e.preventDefault(); this.goto(this.nameBox.value); this.grid.focusSink(); } if (keyName(e) === 'Escape') { this.updateNameBox(); this.grid.focusSink(); } });
    this.nameBox.addEventListener('focus', () => this.nameBox.select());
    this.fbInput.addEventListener('focus', () => { if (!this.grid.editing) { this.grid.beginEdit('edit'); } if (this.grid.editing) { this.grid.editing.inFormulaBar = true; this.grid.editing.sheet = this.wb.cur; } setTimeout(() => this.fbInput.focus(), 0); });
    this.fbInput.addEventListener('input', () => this.onFbInput());
    this.fbInput.addEventListener('keydown', e => {
      const k = keyName(e);
      if (k === 'Enter' && !e.altKey) { e.preventDefault(); this.grid.commit(); this.grid.moveInSelection(e.shiftKey ? -1 : 1, 0); this.grid.focusSink(); }
      else if (k === 'Enter' && e.altKey) { e.preventDefault(); const p = this.fbInput.selectionStart; this.fbInput.value = this.fbInput.value.slice(0, p) + '\n' + this.fbInput.value.slice(p); this.fbInput.setSelectionRange(p + 1, p + 1); this.onFbInput(); }
      else if (k === 'Escape') { e.preventDefault(); this.grid.cancel(); this.grid.focusSink(); }
      else if (k === 'Tab') { e.preventDefault(); this.grid.commit(); this.grid.moveInSelection(0, e.shiftKey ? -1 : 1); this.grid.focusSink(); }
      else if (k === 'F4') { e.preventDefault(); import('./formula.js').then(m => { const r = m.cycleAbs(this.fbInput.value, this.fbInput.selectionStart); if (r) { this.fbInput.value = r.text; this.fbInput.setSelectionRange(r.caret, r.caret); this.onFbInput(); } }); }
    });
    // grid + sheet tabs
    this.grid = new Grid(this);
    this.objects = new Objects(this);
    this.tabsBar = h('div.xl-tabs');
    this.el.classList.add('xl-doc');
    this.el.appendChild(this.grid.el);
    this.el.appendChild(this.tabsBar);
    // status bar
    this.sbMode = h('span.sb-item', 'Sẵn sàng');
    this.sbAcc = h('button.sb-item', ic('n:access', 14), ' Trợ năng: Kiểm tra');   // v0.1.7: native wording
    this.sbAcc.addEventListener('click', () => this.shell.inspect(this));
    this.sbAuto = h('span.sb-item.sb-autosave');
    this.statusLeft.appendChild(this.sbMode); this.statusLeft.appendChild(this.sbAcc); this.statusLeft.appendChild(this.sbAuto);
    this.sbStats = h('button.sb-item.xl-stats', { title: 'Bấm chuột phải để chọn thống kê hiển thị' });
    this.sbStats.addEventListener('contextmenu', e => this.statsMenu(e));
    this.sbStats.addEventListener('click', e => this.statsMenu(e));
    this.statusRight.appendChild(this.sbStats);
    this.noteEl = h('div.xl-note.hidden'); this.el.appendChild(this.noteEl);
  }
  viewButtons() {
    const s = () => this.wb.cur;
    const mk = (icn, title, on, fn) => { const b = h('button.sb-view' + (on ? '.on' : ''), { title }, ic(icn, 14)); b.addEventListener('click', () => { $$('.sb-view', b.parentNode).forEach(x => x.classList.remove('on')); b.classList.add('on'); fn(); }); return b; };
    return h('div.sb-views',
      mk('n:normalview', 'Thường', !s().showPageBreaks, () => { s().showPageBreaks = false; this.grid.redraw(); }),
      mk('n:printlayout', 'Bố cục trang (xem trước khi in)', false, () => this.shell.backstage(this, 'print')),
      mk('n:pagebreak', 'Xem trước ngắt trang', !!s().showPageBreaks, () => { s().showPageBreaks = true; this.grid.redraw(); }));
  }
  ribbonDef() { return unifyRibbon(this, excelRibbon(this)); }   // v0.1.7: same layout as the native editor
  onShow() {
    this.grid.resize();
    this.shell.setZoom(this.wb.cur.zoom || 100);
    this.onSelection();
    this.renderTabs();
    this.objects.render();
    if (!this._ro) { this._ro = true; window.addEventListener('resize', () => { if (this.shell.cur === this) setTimeout(() => this.grid.resize(), 50); }); }
  }
  onLayout() { setTimeout(() => this.grid.resize(), 10); }
  focus() { this.grid.focusSink(); }
  setZoom(z) { this.zoom = z; this.wb.cur.zoom = z; this.grid.invalidateAxes(); this.grid.measureCache = {}; this.grid.updateSpacer(); this.grid.redraw(); this.objects.render(); }
  onBack() { if (this.grid.editing) { this.grid.cancel(); return true; } return false; }
  getText() {
    const out = [];
    this.wb.sheets.forEach(s => { const u = s.used(); if (u.empty) return; out.push('## Trang tính ' + s.name); for (let r = u.r1; r <= Math.min(u.r2, u.r1 + 200); r++) { const row = []; for (let c = u.c1; c <= Math.min(u.c2, u.c1 + 30); c++) { const cell = s.get(r, c); const v = s.value(r, c); row.push(v == null ? '' : (format(v, cell && cell.s && cell.s.nf).t + (cell && cell.f ? ' [=' + cell.f + ']' : ''))); } out.push('| ' + row.join(' | ') + ' |'); } });
    return out.join('\n');
  }
  stats() { let cells = 0, formulas = 0; this.wb.sheets.forEach(s => s.cells.forEach(x => { cells++; if (x.f) formulas++; })); return { 'Số trang tính': this.wb.sheets.length, 'Ô có dữ liệu': cells, 'Công thức': formulas, 'Tên đã đặt': Object.keys(this.wb.names).length, 'Biểu đồ/đối tượng': this.wb.sheets.reduce((n, s) => n + s.objects.length, 0) }; }
  inspect() { const issues = []; this.wb.sheets.forEach(s => { if (/^Sheet\d+$/.test(s.name) && this.wb.sheets.length > 1) issues.push('Trang tính "' + s.name + '" chưa có tên mô tả.'); s.objects.forEach(o => { if (!o.alt && o.type !== 'textbox') issues.push('Đối tượng ' + (o.type === 'chart' ? 'biểu đồ' : 'hình') + ' trên "' + s.name + '" thiếu văn bản thay thế (Alt Text).'); }); if (s.merges.length > 20) issues.push('"' + s.name + '" có nhiều ô gộp, có thể gây khó cho trình đọc màn hình.'); }); return issues; }

  // ---------------------------------------------------------------- sheet tabs
  renderTabs() {
    const bar = this.tabsBar; bar.innerHTML = '';
    const nav = h('div.xl-tabnav',
      h('button', { title: 'Trang trước (Ctrl+PgUp)', onclick: () => this.switchSheet(-1) }, ic('chevron-left', 14)),
      h('button', { title: 'Trang sau (Ctrl+PgDn)', onclick: () => this.switchSheet(1) }, ic('chevron-right', 14)),
      h('button', { title: 'Danh sách trang tính', onclick: e => showMenu(this.wb.sheets.map((s, i) => ({ label: s.name + (s.hidden ? ' (ẩn)' : ''), checked: i === this.wb.active || undefined, run: () => { if (s.hidden) s.hidden = false; this.activateSheet(s); } })), { anchor: e.currentTarget }) }, ic('ellipsis', 14)));
    bar.appendChild(nav);
    const list = h('div.xl-tablist');
    this.wb.sheets.forEach((s, i) => {
      if (s.hidden) return;
      const t = h('div.xl-tab' + (i === this.wb.active ? '.on' : ''), { draggable: 'true' }, h('span', s.name));
      if (s.tabColor) t.style.borderBottom = '3px solid ' + s.tabColor;
      t.addEventListener('click', () => this.activateSheet(s));
      t.addEventListener('dblclick', () => this.renameSheet(s));
      t.addEventListener('contextmenu', e => { this.activateSheet(s); contextMenu(e, this.sheetMenu(s)); });
      longPress(t, e => { this.activateSheet(s); contextMenu(e, this.sheetMenu(s)); });
      t.addEventListener('dragstart', e => { e.dataTransfer.setData('text/plain', 'sheet:' + i); });
      t.addEventListener('dragover', e => e.preventDefault());
      t.addEventListener('drop', e => { e.preventDefault(); e.stopPropagation(); const d = e.dataTransfer.getData('text/plain'); if (/^sheet:/.test(d)) X.moveSheet(this, +d.slice(6), i); });
      list.appendChild(t);
    });
    const add = h('button.xl-tabadd', { title: 'Trang tính mới (Shift+F11)' }, ic('plus', 16));
    add.addEventListener('click', () => X.insertSheet(this));
    list.appendChild(add);
    bar.appendChild(list);
  }
  sheetMenu(s) {
    return [
      { label: 'Chèn trang tính', icon: 'file-plus', run: () => X.insertSheet(this) },
      { label: 'Xóa', icon: 'trash-2', disabled: this.wb.sheets.filter(x => !x.hidden).length < 2, run: () => X.deleteSheet(this, s) },
      { label: 'Đổi tên', icon: 'pencil', run: () => this.renameSheet(s) },
      { label: 'Di chuyển hoặc sao chép…', icon: 'copy', run: () => X.moveCopySheetDialog(this, s) },
      { label: 'Màu tab', icon: 'palette', run: () => colorPicker(this.tabsBar, c => { s.tabColor = c === 'transparent' ? null : c; this.renderTabs(); this.markDirty(); }, { none: 'Không màu' }) },
      { sep: true },
      { label: s.protected ? 'Bỏ bảo vệ trang tính…' : 'Bảo vệ trang tính…', icon: 'lock', run: () => X.toggleProtect(this) },
      { label: 'Ẩn', icon: 'eye-off', disabled: this.wb.sheets.filter(x => !x.hidden).length < 2, run: () => { s.hidden = true; this.switchSheet(1); this.renderTabs(); this.markDirty(); } },
      { label: 'Bỏ ẩn…', icon: 'eye', disabled: !this.wb.sheets.some(x => x.hidden), run: () => X.unhideSheetDialog(this) },
      { label: 'Chọn tất cả trang tính', icon: 'layers', disabled: true }
    ];
  }
  async renameSheet(s) {
    const v = await promptDlg('Tên trang tính:', s.name, 'Đổi tên trang tính');
    if (!v || v === s.name) return;
    if (/[\\/?*[\]:]/.test(v) || v.length > 31) return toast('Tên không hợp lệ (tối đa 31 ký tự, không chứa \\ / ? * [ ] :)', 'err');
    if (this.wb.sheet(v)) return toast('Tên này đã tồn tại', 'err');
    X.renameSheet(this, s, v);
  }
  async activateSheet(s) {
    const i = this.wb.sheets.indexOf(s); if (i < 0) return;
    if (this.grid.editing && !this.grid.isPointMode()) this.grid.commit();
    if (s._lazyLoad) {
      try { await s._lazyLoad(); delete s._lazyLoad; } catch (e) { console.warn(e); }
    }
    this.wb.active = i;
    this.grid.measureCache = {};
    this.zoom = s.zoom || 100;
    this.renderTabs(); this.grid.invalidateAxes(); this.grid.updateSpacer();
    this.grid.scroller.scrollLeft = s.scroll.x || 0; this.grid.scroller.scrollTop = s.scroll.y || 0;
    this.shell._zoomCtl && this.shell._zoomCtl.update();
    this.objects.render();
    this.onSelection(); this.grid.redraw();
  }
  switchSheet(d) {
    const vis = this.wb.sheets.filter(s => !s.hidden); const cur = vis.indexOf(this.wb.cur);
    const n = vis[clamp(cur + d, 0, vis.length - 1)]; if (n && n !== this.wb.cur) { this.wb.cur.scroll = { x: this.grid.scroller.scrollLeft, y: this.grid.scroller.scrollTop }; this.activateSheet(n); }
  }

  // ---------------------------------------------------------------- selection & status
  onSelection() {
    this.updateNameBox();
    if (!this.grid.editing) { const s = this.wb.cur, sel = s.sel; const cell = s.get(sel.r, sel.c); this.fbInput.value = cell ? (cell.f ? '=' + cell.f : this.editText(cell)) : (s.spill.get(key(sel.r, sel.c)) ? '=' + this.spillSource(sel.r, sel.c) : ''); this.fbInput.classList.toggle('ghost', !cell && !!s.spill.get(key(sel.r, sel.c))); }
    this.updateStats();
    if (this.ribbon) this.ribbon.refresh();
    this.objects.onSelectionChanged && this.objects.onSelectionChanged();
    this.shell.updateTitle && this.shell.updateTitle();
    this.updateContextualTabs();
  }
  spillSource(r, c) { const s = this.wb.cur; const sp = s.spill.get(key(r, c)); if (!sp) return ''; const src = s.cells.get(sp.src); return src && src.f ? src.f : ''; }
  updateNameBox() {
    if (document.activeElement === this.nameBox) return;
    const s = this.wb.cur, sel = s.sel; const rg = sel.ranges[sel.ranges.length - 1];
    let t;
    if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1 && rg.c1 === 0 && rg.c2 >= MAX_COLS - 1) t = addr(sel.r, sel.c);
    else if (this.grid.dragState) t = (rg.r2 - rg.r1 + 1) + 'R x ' + (rg.c2 - rg.c1 + 1) + 'C';
    else t = addr(sel.r, sel.c);
    for (const n in this.wb.names) { const d = this.wb.names[n]; const f = String(d.f || d).replace(/^=/, ''); const m = /^(?:'?([^'!]+)'?!)?(.+)$/.exec(f); if (m && (!m[1] || m[1] === s.name)) { const rr = parseRangeAddr(m[2]); if (rr && rr.r1 === rg.r1 && rr.c1 === rg.c1 && rr.r2 === rg.r2 && rr.c2 === rg.c2) { t = n; break; } } }
    this.nameBox.value = t;
  }
  updateStats() {
    const s = this.wb.cur, sel = s.sel;
    const single = sel.ranges.length === 1 && sel.ranges[0].r1 === sel.ranges[0].r2 && sel.ranges[0].c1 === sel.ranges[0].c2;
    if (single) { this.sbStats.textContent = ''; return; }
    let sum = 0, count = 0, numCount = 0, min = Infinity, max = -Infinity;
    sel.ranges.forEach(rg => {
      const R = { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, s.used().r2), c2: Math.min(rg.c2, s.used().c2) };
      this.calc.eachIn({ sheet: s, r1: R.r1, c1: R.c1, r2: R.r2, c2: R.c2, rows: 0, cols: 0 }, (v, r) => { if (s.hiddenR[r]) return; if (v != null && v !== '') count++; if (typeof v === 'number') { sum += v; numCount++; if (v < min) min = v; if (v > max) max = v; } });
    });
    const nf = (s.styleAt(sel.r, sel.c) || {}).nf;
    const f = v => format(+v.toPrecision(12), nf && !/@/.test(nf) ? nf : 'General').t;
    const parts = [];
    const fns = this.statFns;
    if (numCount && fns.indexOf('avg') >= 0) parts.push('Trung bình: ' + f(sum / numCount));
    if (count && fns.indexOf('count') >= 0) parts.push('Số lượng: ' + count);
    if (numCount && fns.indexOf('ncount') >= 0) parts.push('Số đếm: ' + numCount);
    if (numCount && fns.indexOf('min') >= 0) parts.push('Nhỏ nhất: ' + f(min));
    if (numCount && fns.indexOf('max') >= 0) parts.push('Lớn nhất: ' + f(max));
    if (numCount && fns.indexOf('sum') >= 0) parts.push('Tổng: ' + f(sum));
    this.sbStats.textContent = parts.join('    ');
    this._stats = { sum, count, numCount, min, max };
  }
  statsMenu(e) {
    e.preventDefault();
    const opt = [['avg', 'Trung bình (Average)'], ['count', 'Số lượng (Count)'], ['ncount', 'Số đếm (Numerical Count)'], ['min', 'Nhỏ nhất (Minimum)'], ['max', 'Lớn nhất (Maximum)'], ['sum', 'Tổng (Sum)']];
    showMenu([{ header: 'Tùy chỉnh thanh trạng thái' }].concat(opt.map(o => ({ label: o[1], checked: this.statFns.indexOf(o[0]) >= 0 ? true : undefined, keepOpen: false, run: () => { const i = this.statFns.indexOf(o[0]); if (i >= 0) this.statFns.splice(i, 1); else this.statFns.push(o[0]); const st = settings(); st.xlStatFns = this.statFns; import('../core/store.js').then(m => m.saveSettings({ xlStatFns: this.statFns })); this.updateStats(); } }))).concat(this._stats ? [{ sep: true }, { label: 'Sao chép Tổng: ' + this._stats.sum, icon: 'copy', run: () => copyText(String(this._stats.sum)) }] : []), { x: e.clientX, y: e.clientY - 10 });
  }
  updateContextualTabs() {
    const s = this.wb.cur, sel = s.sel;
    const inTable = s.tables.find(t => inRange(t, sel.r, sel.c));
    const objSel = this.objects.selected;
    const key2 = (inTable ? 't' : '') + (objSel ? objSel.type : '');
    if (this._ctxKey !== key2 && this.ribbon) {
      this._ctxKey = key2;
      this.ribbon.def = this.ribbonDef();
      this.ribbon.renderTabs();
      if (!this.ribbon.def.tabs.find(t => t.id === this.ribbon.cur)) this.ribbon.select('home');
    }
  }

  // ---------------------------------------------------------------- editing hooks (from Grid)
  editText(cell) {
    const v = cell.v;
    if (v == null) return '';
    if (v instanceof Err) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') {
      const nf = cell.s && cell.s.nf;
      if (nf && isDateFormat(nf)) { const hasTime = /h|s/i.test(nf) || v % 1; return format(v, Math.floor(v) === 0 ? 'hh:mm:ss' : (hasTime && v % 1 ? 'dd/mm/yyyy hh:mm:ss' : 'dd/mm/yyyy')).t; }
      if (nf && /%/.test(nf)) return +(v * 100).toPrecision(12) + '%';
      return String(+v.toPrecision(15));
    }
    if (cell.s && cell.s.nf === '@') return String(v);
    const p = parseInput(String(v)); if (typeof p.v !== 'string' && typeof v === 'string') return "'" + v;
    return String(v);
  }
  onEditStart(text, mode) { this.fbInput.value = text; this.sbMode.textContent = (mode || (this.grid.editing && this.grid.editing.mode)) === 'edit' ? 'Sửa' : 'Nhập'; this.formulaBar.classList.add('editing'); }
  onEditInput(text, fromFb) { if (!fromFb) this.fbInput.value = text; else { this.grid.editor.value = text; } }
  onFbInput() {
    if (!this.grid.editing) this.grid.beginEdit('edit', this.fbInput.value);
    this.grid.editing.inFormulaBar = true; this.grid.pointRef = null;
    this.grid.editor.value = this.fbInput.value; this.grid.placeEditor(); this.grid.updateEditRefs(); this.grid.autoComplete(); this.grid.redraw();
  }
  onEditEnd() { this.sbMode.textContent = 'Sẵn sàng'; this.formulaBar.classList.remove('editing'); this.fnTip(null); this.onSelection(); }
  formulaInputFocused() { return document.activeElement === this.fbInput; }
  fnTip(name, sig, argi) {
    const el = this.fnTipEl;
    if (!name) { el.classList.add('hidden'); return; }
    const args = String(sig || '').split(/,\s*/);
    el.innerHTML = '<b>' + name + '</b>(' + args.map((a, i) => i === argi ? '<b class="cur">' + esc(a) + '</b>' : esc(a)).join(', ') + ')';
    el.classList.remove('hidden');
  }
  isLocked(r, c) { const st = this.wb.cur.styleAt(r, c); return !(st && st.unlocked); }

  // set cell from user text input
  setInput(r, c, text, fillSelection) {
    const s = this.wb.cur;
    const targets = fillSelection ? this.selectedCells(true) : [[r, c]];
    const dv = X.validate(this, s, r, c, text);
    if (dv === false) return;
    this.change('Nhập dữ liệu', s, targets.length === 1 ? { r1: r, c1: c, r2: r, c2: c } : this.boundingSel(), () => {
      targets.forEach(([rr, cc]) => {
        const cell = s.get(rr, cc) || {};
        let t = text;
        if (t === '' || t == null) { s.setCell(rr, cc, { v: null, f: null }); return; }
        if (/^[=]/.test(t) && t.length > 1) {
          let f = t.slice(1);
          f = X.normalizeFormula(f);
          if (fillSelection && (rr !== r || cc !== c)) f = shiftFormula(f, rr - r, cc - c);
          s.setCell(rr, cc, { f, v: null });
          // auto number format from referenced function types
          if (!cell.s || !cell.s.nf) { const nf = X.inferFormulaFormat(this, f); if (nf) s.setCell(rr, cc, { s: mergeStyle(cell.s, { nf }) }); }
          return;
        }
        if (cell.s && cell.s.nf === '@') { s.setCell(rr, cc, { v: t, f: null }); return; }
        if (t[0] === "'") { s.setCell(rr, cc, { v: t.slice(1), f: null }); return; }
        const p = parseInput(t);
        let v = p.v;
        // percent format on cell: typing 5 means 5%
        if (typeof v === 'number' && cell.s && cell.s.nf && /%/.test(cell.s.nf) && !/%/.test(t)) v = v / 100;
        const patch = { v, f: null };
        if (p.nf && (!cell.s || !cell.s.nf || cell.s.nf === 'General')) patch.s = mergeStyle(cell.s, { nf: p.nf });
        if (typeof v === 'string' && /\n/.test(v)) patch.s = mergeStyle(patch.s || cell.s, { wrap: true });
        // autocomplete hyperlink
        if (typeof v === 'string' && /^(https?:\/\/|www\.)\S+$/i.test(v)) { patch.link = { url: /^www\./i.test(v) ? 'http://' + v : v }; patch.s = mergeStyle(patch.s || cell.s, { color: '#0563C1', u: 'single' }); }
        s.setCell(rr, cc, patch);
      });
      // auto-fit row height for wrapped/multi-line text
      targets.slice(0, 50).forEach(([rr]) => X.autoRowHeight(this, s, rr));
      X.extendTable(this, s, r, c);
    });
  }
  selectedCells(limit) {
    const s = this.wb.cur, out = []; const u = s.used();
    s.sel.ranges.forEach(rg => { const r2 = Math.min(rg.r2, Math.max(u.r2, rg.r1 + 0)), c2 = Math.min(rg.c2, Math.max(u.c2, rg.c1)); for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000) && (r <= r2 || !limit || rg.r2 - rg.r1 < 100000); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 1000) && (c <= c2 || !limit || rg.c2 - rg.c1 < 1000); c++) out.push([r, c]); });
    return out.slice(0, 2000000);
  }
  boundingSel() { const rs = this.wb.cur.sel.ranges; return { r1: Math.min.apply(null, rs.map(r => r.r1)), c1: Math.min.apply(null, rs.map(r => r.c1)), r2: Math.max.apply(null, rs.map(r => r.r2)), c2: Math.max.apply(null, rs.map(r => r.c2)) }; }

  // ---------------------------------------------------------------- undo / redo (snapshot based)
  snapRange(s, rg) {
    if (rg === 'sheet' || !rg) return { whole: true, json: X.sheetState(s) };
    if (rg === 'props') return { props: true, json: X.sheetProps(s) };
    if (rg === 'book') return { book: true, json: wbToJSON(this.wb) };
    const u = s.used();
    const R = { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(u.r2, rg.r1)), c2: Math.min(rg.c2, Math.max(u.c2, rg.c1)) };
    const cells = [];
    if ((R.r2 - R.r1 + 1) * (R.c2 - R.c1 + 1) > s.cells.size * 2) { s.cells.forEach((x, k) => { const r = Math.floor(k / 20000), c = k % 20000; if (inRange(R, r, c)) cells.push([k, Object.assign({}, x)]); }); }
    else for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) { const x = s.cells.get(key(r, c)); if (x) cells.push([key(r, c), Object.assign({}, x)]); }
    return { range: R, cells, merges: s.merges.slice(), rowH: Object.assign({}, s.rowH) };
  }
  restoreSnap(s, snap) {
    if (snap.book) { const wb2 = wbFromJSON(snap.json); this.wb.sheets = wb2.sheets; this.wb.sheets.forEach(x => { x.wb = this.wb; }); this.wb.names = wb2.names; this.wb.active = Math.min(this.wb.active, this.wb.sheets.length - 1); this.calc = new Calc(this.wb); this.renderTabs(); return; }
    if (snap.whole) { X.applySheetState(s, snap.json); return; }
    if (snap.props) { X.applySheetProps(s, snap.json); return; }
    const R = snap.range; const del = [];
    s.cells.forEach((x, k) => { const r = Math.floor(k / 20000), c = k % 20000; if (inRange(R, r, c)) del.push(k); });
    del.forEach(k => s.cells.delete(k));
    snap.cells.forEach(([k, x]) => s.cells.set(k, Object.assign({}, x)));
    s.merges = snap.merges.slice(); s.rowH = Object.assign({}, snap.rowH);
    s._ax && (s._ax.r.dirty = true, s._ax.c.dirty = true);
  }
  // change(label, sheet, rangeOrScope, fn) — records undo, recalcs, redraws
  change(label, s, rg, fn) {
    if (this.readOnly) { toast('Tài liệu đang ở chế độ chỉ đọc', 'err'); return; }
    s = s || this.wb.cur;
    const before = this.snapRange(s, rg);
    const selBefore = deepClone(s.sel);
    let res;
    try { res = fn(); } catch (e) { console.error(e); toast('Lỗi: ' + (e.message || e), 'err'); }
    const after = this.snapRange(s, before.range || rg);
    const selAfter = deepClone(s.sel);
    this.undoStack.push({ label, undo: () => { this.restoreSnap(s, before); s.sel = deepClone(selBefore); }, redo: () => { this.restoreSnap(s, after); s.sel = deepClone(selAfter); }, sheet: s });
    if (this.undoStack.length > 200) this.undoStack.shift();
    this.redoStack = [];
    this.lastChange = { label, fn };
    this.afterChange();
    return res;
  }
  commitPropsChange(prop, beforeJson) {
    const s = this.wb.cur; const after = JSON.stringify(s[prop]);
    if (after === beforeJson) return;
    this.undoStack.push({ label: 'Thay đổi kích thước', undo: () => { s[prop] = JSON.parse(beforeJson); this.grid.invalidateAxes(); }, redo: () => { s[prop] = JSON.parse(after); this.grid.invalidateAxes(); }, sheet: s });
    this.redoStack = [];
    this.afterChange(true);
  }
  afterChange(noRecalc) {
    if (!noRecalc) this.recalc();
    this.grid.invalidateAxes(); this.grid.updateSpacer(); this.grid.redraw();
    this.objects.refreshCharts && this.objects.refreshCharts();
    this.onSelection();
    this.markDirty();
    this.shell.updateTitle();
  }
  recalc(force) {
    if (this.wb.calcMode === 'manual' && !force) return;
    const circ = this.calc.recalc();
    if (circ && !this._circWarned) { this._circWarned = true; toast('Cảnh báo tham chiếu vòng tại ' + circ.sheet + '!' + addr(circ.r, circ.c) + ' — công thức tham chiếu tới chính nó.', 'err', 5000); }
    if (!circ) this._circWarned = false;
  }
  undo() { if (this.grid.editing) { this.grid.cancel(); return; } const u = this.undoStack.pop(); if (!u) return; u.undo(); this.redoStack.push(u); if (u.sheet && u.sheet !== this.wb.cur && this.wb.sheets.indexOf(u.sheet) >= 0) this.activateSheet(u.sheet); this.afterChange(); toast('Hoàn tác: ' + u.label); }
  redo() { const u = this.redoStack.pop(); if (!u) { if (this.lastChange && this.redoStack.length === 0) return; return; } u.redo(); this.undoStack.push(u); this.afterChange(); toast('Làm lại: ' + u.label); }
  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }
  undoHistory() { return this.undoStack.slice().reverse().map(u => u.label); }
  repeatLast() { if (this.lastChange && this.lastRepeatable) this.lastRepeatable(); }

  // ---------------------------------------------------------------- clipboard
  copyEvent(e, cut) { X.copy(this, cut, e); }
  pasteEvent(e) { X.pasteFromEvent(this, e); }

  // keyboard shortcuts (non-editing)
  shortcut(ch, e) {
    if (ch === 'alt+f8') { import('./vba/index.js').then(m => m.macrosDialog(this)); return true; }
    if (ch === 'alt+f11') { import('./vba/index.js').then(m => m.openEditor(this)); return true; }
    if (this.wb.vba && this.wb.vba.shortcuts && e && e.ctrlKey && e.shiftKey && /^[a-z]$/i.test(e.key || '') && this.wb.vba.shortcuts[(e.key || '').toLowerCase()]) { const nm = this.wb.vba.shortcuts[e.key.toLowerCase()]; import('./vba/index.js').then(m => m.runMacro(this, nm)); return true; }
    const map = {
      'ctrl+b': () => X.toggleStyle(this, 'b'), 'ctrl+i': () => X.toggleStyle(this, 'i'), 'ctrl+u': () => X.toggleStyle(this, 'u', 'single'), 'ctrl+5': () => X.toggleStyle(this, 's'),
      'ctrl+z': () => this.undo(), 'ctrl+y': () => this.redo(), 'ctrl+shift+z': () => this.redo(),
      'ctrl+c': null, 'ctrl+x': null, 'ctrl+v': null,
      'ctrl+1': () => X.formatCellsDialog(this), 'ctrl+shift+l': () => X.toggleFilter(this), 'ctrl+f': () => X.findDialog(this), 'ctrl+h': () => X.findDialog(this, true), 'ctrl+g': () => X.gotoDialog(this), 'f5': () => X.gotoDialog(this),
      'ctrl+;': () => this.setInput(this.wb.cur.sel.r, this.wb.cur.sel.c, X.todayText()), 'ctrl+shift+;': () => this.setInput(this.wb.cur.sel.r, this.wb.cur.sel.c, X.nowTimeText()),
      'alt+=': () => X.autoSum(this, 'SUM'), 'ctrl+d': () => X.fillDir(this, 'down'), 'ctrl+r': () => X.fillDir(this, 'right'), 'ctrl+k': () => X.linkDialog(this), 'ctrl+t': () => X.formatAsTable(this), 'ctrl+l': () => X.formatAsTable(this),
      'ctrl+shift+4': () => X.setNumFmt(this, '"$"#,##0.00'), 'ctrl+shift+5': () => X.setNumFmt(this, '0%'), 'ctrl+shift+1': () => X.setNumFmt(this, '#,##0.00'), 'ctrl+shift+3': () => X.setNumFmt(this, 'dd/mm/yyyy'), 'ctrl+shift+`': () => X.setNumFmt(this, 'General'),
      'shift+f11': () => X.insertSheet(this), 'shift+f3': () => X.insertFunctionDialog(this), 'f9': () => { this.recalc(true); this.grid.redraw(); toast('Đã tính toán lại'); }, 'ctrl+`': () => X.toggleShowFormulas(this),
      'ctrl+-': () => X.deleteDialog(this), 'ctrl+shift+=': () => X.insertDialog(this), 'ctrl+9': () => X.hideRows(this, true), 'ctrl+0': () => X.hideCols(this, true), 'shift+f2': () => X.noteDialog(this), 'ctrl+shift+f2': () => X.noteDialog(this),
      'alt+enter': null, 'ctrl+shift+&': () => X.applyBorder(this, 'outside'), 'ctrl+shift+_': () => X.applyBorder(this, 'none'), 'ctrl+e': () => X.flashFill(this), 'alt+f1': () => X.insertChart(this, 'column'), 'f11': () => X.insertChart(this, 'column', true),
      'ctrl+p': null, 'ctrl+n': null, 'ctrl+o': null, 'ctrl+s': null
    };
    const f = map[ch];
    if (f === null) return false;
    if (f) { f(); this.refreshUi(); return true; }
    return false;
  }
  clearContents(noRecord) { X.clear(this, 'contents'); }
  moveRange(rg, r, c, copy) { X.moveRange(this, rg, r, c, copy); }
  autoFill(src, dst) { X.autoFill(this, src, dst); }
  fillPreviewText(src, dst) { return X.fillPreview(this, src, dst); }
  autoFit(axis, i) { X.autoFit(this, axis, i); }
  autoFillDown() { X.autoFillDown(this); }
  contextItems(zone) { return X.contextItems(this, zone); }
  cfStyle(s, r, c, v) { return s.cf.length ? X.cfStyle(this, s, r, c, v) : null; }
  filterButtonHit(hit, p) { return X.filterButtonHit(this, hit, p); }
  followLink(link) { X.followLink(this, link); }
  pageBreaks(s) { return X.pageBreaks(this, s); }
  goto(ref) { X.gotoRef(this, ref); }
  showNotePopup(r, c) {
    const s = this.wb.cur; const cell = s.get(r, c); if (!cell || !cell.note) return;
    const rc = this.grid.cellRect(r, c);
    const n = cell.note; const el = this.noteEl;
    el.innerHTML = '';
    if (n.thread) n.thread.forEach(m => el.appendChild(h('div.xl-note-msg', h('b', m.author), h('span.muted', ' ' + m.time), h('div', m.text))));
    else el.appendChild(h('div', h('b', (n.author || '') + (n.author ? ':' : '')), h('div', { style: { whiteSpace: 'pre-wrap' } }, n.text || '')));
    el.style.left = (rc.x + rc.w + 8) + 'px'; el.style.top = (rc.y) + 'px';
    el.classList.remove('hidden');
  }
  hideNotePopup() { this.noteEl.classList.add('hidden'); }
  objectsDeselect() { this.objects.deselect(); }
  showComments() { X.commentsPane(this); }
  printPreview(box) { X.printPreview(this, box); }
  printOptions() { return X.printOptions(this); }
  renderPrint(area) { return X.renderPrint(this, area); }
  acceptsDrop(e) { return this.objects.acceptsDrop(e); }
  commands() { return X.extraCommands(this); }
  setReadOnly(v) { this.readOnly = v; }
  destroy() { if (this.grid._ants) clearInterval(this.grid._ants); }
}
function copyText(t) { const ta = h('textarea', { style: { position: 'fixed', left: '-1000px' } }); ta.value = t; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); } catch (e) {} ta.parentNode.removeChild(ta); toast('Đã sao chép'); }

registerApp('excel', ExcelDoc);
