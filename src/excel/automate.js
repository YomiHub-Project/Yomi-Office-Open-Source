// Automate tab: Office Scripts-style JavaScript automation with a safe workbook API, built-in script gallery, action recorder
import { h, ic, esc } from '../core/util.js';
import { dialog, toast, promptDlg, alertDlg } from '../core/ui.js';
import { settings, saveSettings } from '../core/store.js';
import { parseRangeAddr, rangeAddr, colName, key, Err, MAX_ROWS, MAX_COLS, mergeStyle } from './model.js';
import { parseInput, format } from './numfmt.js';

export const BUILTIN = [
  { name: 'Bỏ ẩn mọi hàng và cột', code: "const s = workbook.getActiveWorksheet();\ns.unhideAll();\nlog('Đã bỏ ẩn mọi hàng và cột');" },
  { name: 'Cố định vùng chọn', code: "const s = workbook.getActiveWorksheet();\nconst r = workbook.getSelectedRange();\ns.freezeAt(r.row, r.col);\nlog('Đã cố định tại ' + r.address);" },
  { name: 'Xóa liên kết khỏi trang tính', code: "const s = workbook.getActiveWorksheet();\nlet n = 0;\ns.forEachCell((c) => { if (c.link) { c.removeLink(); n++; } });\nlog('Đã xóa ' + n + ' liên kết');" },
  { name: 'Tạo bảng con từ vùng chọn', code: "const r = workbook.getSelectedRange();\nconst ns = workbook.addWorksheet('Bảng con');\nns.getRange('A1').setValues(r.getValues());\nlog('Đã sao chép ' + r.address + ' sang trang tính mới');" },
  { name: 'Đếm hàng trống', code: "const s = workbook.getActiveWorksheet();\nconst u = s.getUsedRange();\nlet n = 0;\nfor (let i = 0; i < u.rowCount; i++) { if (u.getRow(i).every(v => v === null || v === '')) n++; }\nlog('Số hàng trống trong vùng dữ liệu: ' + n);" },
  { name: 'Trả dữ liệu bảng dạng JSON', code: "const s = workbook.getActiveWorksheet();\nconst v = s.getUsedRange().getValues();\nconst head = v[0];\nconst rows = v.slice(1).map(r => { const o = {}; head.forEach((k, i) => o[k] = r[i]); return o; });\nlog(JSON.stringify(rows, null, 2));" },
  { name: 'Tô màu xen kẽ hàng', code: "const s = workbook.getActiveWorksheet();\nconst u = s.getUsedRange();\nfor (let i = 1; i < u.rowCount; i += 2) u.getRowRange(i).setFill('#eef3fb');\nlog('Đã tô ' + Math.floor(u.rowCount / 2) + ' hàng');" },
  { name: 'Viết hoa chữ cái đầu vùng chọn', code: "const r = workbook.getSelectedRange();\nr.setValues(r.getValues().map(row => row.map(v => typeof v === 'string' ? v.toLowerCase().replace(/(^|\\s)(\\S)/g, (m,a,b) => a + b.toUpperCase()) : v)));" },
  { name: 'Xóa khoảng trắng thừa', code: "const r = workbook.getSelectedRange();\nr.setValues(r.getValues().map(row => row.map(v => typeof v === 'string' ? v.replace(/\\s+/g, ' ').trim() : v)));" }
];

function api(doc, out) {
  const wb = doc.wb;
  const rangeObj = (s, rg) => ({
    address: s.name + '!' + rangeAddr(rg), row: rg.r1, col: rg.c1, rowCount: rg.r2 - rg.r1 + 1, columnCount: rg.c2 - rg.c1 + 1,
    getValues() { const o = []; for (let r = rg.r1; r <= rg.r2; r++) { const row = []; for (let c = rg.c1; c <= rg.c2; c++) { const v = s.value(r, c); row.push(v instanceof Err ? v.code : v); } o.push(row); } return o; },
    getValue() { const v = s.value(rg.r1, rg.c1); return v instanceof Err ? v.code : v; },
    getFormulas() { const o = []; for (let r = rg.r1; r <= rg.r2; r++) { const row = []; for (let c = rg.c1; c <= rg.c2; c++) { const x = s.get(r, c); row.push(x && x.f ? '=' + x.f : (x ? x.v : null)); } o.push(row); } return o; },
    setValues(m) { m.forEach((row, i) => row.forEach((v, j) => { if (typeof v === 'string' && /^=/.test(v)) s.setCell(rg.r1 + i, rg.c1 + j, { f: v.slice(1), v: null }); else s.setCell(rg.r1 + i, rg.c1 + j, { v: typeof v === 'string' ? parseInput(v).v : v, f: null }); })); },
    setValue(v) { this.setValues([[v]]); },
    setFormula(f) { s.setCell(rg.r1, rg.c1, { f: String(f).replace(/^=/, ''), v: null }); },
    setFill(c) { for (let r = rg.r1; r <= rg.r2; r++) for (let cc = rg.c1; cc <= Math.min(rg.c2, s.used().c2); cc++) { const x = s.get(r, cc); s.setCell(r, cc, { s: mergeStyle(x && x.s, { fill: c }) }); } },
    setBold(b) { for (let r = rg.r1; r <= rg.r2; r++) for (let cc = rg.c1; cc <= rg.c2; cc++) { const x = s.get(r, cc); s.setCell(r, cc, { s: mergeStyle(x && x.s, { b: b !== false || null }) }); } },
    setNumberFormat(nf) { for (let r = rg.r1; r <= rg.r2; r++) for (let cc = rg.c1; cc <= rg.c2; cc++) { const x = s.get(r, cc); s.setCell(r, cc, { s: mergeStyle(x && x.s, { nf }) }); } },
    clear() { for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) s.cells.delete(key(r, c)); },
    getRow(i) { const o = []; for (let c = rg.c1; c <= rg.c2; c++) o.push(s.value(rg.r1 + i, c)); return o; },
    getRowRange(i) { return rangeObj(s, { r1: rg.r1 + i, c1: rg.c1, r2: rg.r1 + i, c2: rg.c2 }); },
    getCell(i, j) { return rangeObj(s, { r1: rg.r1 + i, c1: rg.c1 + j, r2: rg.r1 + i, c2: rg.c1 + j }); }
  });
  const sheetObj = s => ({
    name: s.name,
    getRange(a) { const rg = parseRangeAddr(a); if (!rg) throw new Error('Địa chỉ không hợp lệ: ' + a); return rangeObj(s, rg); },
    getUsedRange() { return rangeObj(s, s.used()); },
    getCell(r, c) { return rangeObj(s, { r1: r, c1: c, r2: r, c2: c }); },
    unhideAll() { s.hiddenR = {}; s.hiddenC = {}; s.filterHidden = {}; if (s._ax) { s._ax.r.dirty = s._ax.c.dirty = true; } },
    freezeAt(r, c) { s.freeze = { r, c }; },
    forEachCell(fn) { s.cells.forEach((x, k) => fn({ row: Math.floor(k / 20000), col: k % 20000, value: x.v, formula: x.f, link: x.link, removeLink() { delete x.link; if (x.s) x.s = mergeStyle(x.s, { u: null, color: null }); } })); },
    setName(n) { s.name = n; }
  });
  return {
    workbook: {
      getActiveWorksheet: () => sheetObj(wb.cur), getWorksheet: n => { const s = wb.sheet(n); return s ? sheetObj(s) : null; }, getWorksheets: () => wb.sheets.map(sheetObj),
      addWorksheet: n => { let nm = n || wb.nextName(); while (wb.sheet(nm)) nm += '_'; return sheetObj(wb.addSheet(nm)); },
      getSelectedRange: () => rangeObj(wb.cur, doc.grid.selRange()),
      recalc: () => doc.recalc(true)
    },
    log: (...a) => out.push(a.map(x => typeof x === 'object' ? JSON.stringify(x) : String(x)).join(' ')),
    console: { log: (...a) => out.push(a.map(x => typeof x === 'object' ? JSON.stringify(x) : String(x)).join(' ')) }
  };
}
export function runScript(doc, code, name) {
  const out = [];
  try {
    doc.change('Tập lệnh: ' + (name || 'Tập lệnh'), doc.wb.cur, 'book', () => {
      const env = api(doc, out);
      // eslint-disable-next-line no-new-func
      const fn = new Function('workbook', 'log', 'console', '"use strict";\n' + code);
      fn(env.workbook, env.log, env.console);
    });
    doc.renderTabs();
    if (out.length) alertDlg(out.join('\n'), 'Kết quả tập lệnh: ' + (name || ''));
    else toast('Đã chạy tập lệnh "' + (name || '') + '"', 'ok');
  } catch (e) { alertDlg('Lỗi tập lệnh: ' + (e.message || e), 'Office Scripts'); }
}
export function scriptEditor(doc, existing) {
  const ta = h('textarea.inp.code', { rows: 16, spellcheck: false }); ta.value = existing ? existing.code : "// Office Scripts (JavaScript) — chạy ngoại tuyến\nconst s = workbook.getActiveWorksheet();\nconst r = workbook.getSelectedRange();\nlog('Đang chọn: ' + r.address);\n";
  const name = h('input.inp', { value: existing ? existing.name : 'Tập lệnh ' + ((settings().xlScripts || []).length + 1) });
  const body = h('div.form', h('label.lbl', 'Tên tập lệnh'), name, h('label.lbl', 'Mã'), ta, h('div.hint', 'API: workbook.getActiveWorksheet(), getSelectedRange(), addWorksheet(name); range.getValues()/setValues(), setFormula, setFill, setBold, setNumberFormat; sheet.getRange("A1:B5"), getUsedRange(), freezeAt(r,c), unhideAll(); log(...).'));
  return dialog({ title: 'Trình soạn Tập lệnh', width: 720, body, buttons: [
    { label: 'Chạy', primary: true, run: () => { runScript(doc, ta.value, name.value); return false; } },
    { label: 'Lưu', run: () => { const list = (settings().xlScripts || []).filter(x => x.name !== name.value); list.push({ name: name.value, code: ta.value }); saveSettings({ xlScripts: list }); toast('Đã lưu tập lệnh', 'ok'); return false; } },
    { label: 'Đóng', value: null }] });
}
export function scriptsPane(doc) {
  const mine = settings().xlScripts || [];
  const row = (sc, own) => h('div.find-row', h('span', ic('file-code', 14), ' ' + sc.name), h('span'), h('span', h('button.btn', { onclick: () => runScript(doc, sc.code, sc.name) }, 'Chạy'), h('button.btn', { onclick: () => scriptEditor(doc, sc) }, 'Sửa'), own ? h('button.btn', { onclick: () => { saveSettings({ xlScripts: (settings().xlScripts || []).filter(x => x !== sc) }); toast('Đã xóa'); } }, 'Xóa') : null));
  return dialog({ title: 'Tập lệnh (Office Scripts / Macro)', width: 640, body: h('div', h('h3', 'Tập lệnh của tôi'), mine.length ? mine.map(s => row(s, true)) : h('div.muted', 'Chưa có. Bấm "Tập lệnh Mới".'), h('h3', 'Thư viện mẫu'), BUILTIN.map(s => row(s, false))), buttons: [{ label: 'Tập lệnh Mới', run: () => { scriptEditor(doc); return true; } }, { label: 'Đóng', primary: true }] });
}
// simple recorder: wraps doc.setInput & setStyle calls into a script
export function recordActions(doc) {
  if (doc._rec) {
    const steps = doc._rec; doc._rec = null; doc.setInput = doc._origSetInput;
    const code = "const s = workbook.getActiveWorksheet();\n" + steps.map(st => st).join('\n');
    scriptEditor(doc, { name: 'Thao tác đã ghi', code });
    return;
  }
  doc._rec = []; doc._origSetInput = doc.setInput;
  doc.setInput = function (r, c, text, fill) { doc._rec.push("s.getRange('" + colName(c) + (r + 1) + "').setValues([[" + JSON.stringify(text) + "]]);"); return doc._origSetInput.call(doc, r, c, text, fill); };
  toast('Đang ghi thao tác nhập liệu… bấm "Ghi lại Thao tác" lần nữa để dừng', 'ok', 4000);
}
