// Yomi-Office Open Source v0.0.6 — VBA for Excel: project storage, macro runner, Visual Basic Editor (Alt+F11), Macros dialog (Alt+F8),
// macro recorder (records real VBA), keyboard shortcuts for macros (Ctrl+Shift+letter) and .bas import/export.
import { h, ic, esc, $, $$ } from '../../core/util.js';
import { dialog, toast, alertDlg, promptDlg, confirmDlg, formDlg } from '../../core/ui.js';
import { settings } from '../../core/store.js';
import Host from '../../core/host.js';
import { parseModule, Interp, VbaError } from './interp.js';
import { ExcelModel, installDispatch, rgbLong } from './objects.js';
import { runJsa, jsaFunctions, jsaHas } from './jsa.js';
import { colName, addr, wbToJSON, wbFromJSON } from '../model.js';
import { shiftFormula } from '../formula.js';

const SAMPLE = `' Module1 — macro VBA mẫu (Yomi-Office Open Source chạy VBA ngoại tuyến)
Option Explicit

Sub XinChao()
    MsgBox "Xin chào từ VBA! Hôm nay là " & Format(Date, "dd/mm/yyyy"), vbInformation, "Yomi-Office Open Source"
End Sub

' Tạo bảng cửu chương 2..9 bắt đầu tại ô đang chọn
Sub BangCuuChuong()
    Dim i As Integer, j As Integer
    For i = 1 To 9
        For j = 2 To 9
            ActiveCell.Offset(i - 1, j - 2).Value = j & " x " & i & " = " & (i * j)
        Next j
    Next i
    ActiveCell.Resize(9, 8).Columns.AutoFit
End Sub

' Tô màu xen kẽ các hàng trong vùng dữ liệu hiện tại
Sub ToMauXenKe()
    Dim r As Long, vung As Range
    Set vung = ActiveCell.CurrentRegion
    For r = 2 To vung.Rows.Count Step 2
        vung.Rows(r).Interior.Color = RGB(221, 235, 247)
    Next r
    vung.Rows(1).Font.Bold = True
End Sub

' Hàm người dùng tự định nghĩa (UDF) — dùng được trong macro
Function ThueVAT(soTien As Double, Optional tyLe As Double = 0.1) As Double
    ThueVAT = Round(soTien * tyLe, 0)
End Function
`;

// Excel keeps the project on the workbook, Word on the document (both use the same VBA engine)
const WORD_SAMPLE = `' NewMacros — macro VBA cho Word (Yomi-Office Open Source chạy VBA ngoại tuyến)
Option Explicit

Sub ChenNgayGio()
    Selection.TypeText Text:="Hà Nội, ngày " & Format(Date, "dd") & " tháng " & Format(Date, "mm") & " năm " & Format(Date, "yyyy")
End Sub

' Thay thế hàng loạt (Find & Replace) trên toàn văn bản
Sub ChuanHoaKhoangTrang()
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Text = "  "
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
    End With
    MsgBox "Đã chuẩn hóa khoảng trắng. Số từ: " & ActiveDocument.ComputeStatistics(wdStatisticWords), vbInformation
End Sub

' Định dạng tiêu đề văn bản hành chính: in đậm, căn giữa, cỡ 14
Sub DinhDangTieuDe()
    With ActiveDocument.Paragraphs(1).Range
        .Font.Bold = True
        .Font.Size = 14
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
End Sub
`;
export function project(doc) {
  const holder = doc.wb || doc; if (!holder.vba) holder.vba = { modules: [], shortcuts: {} };
  if (!holder.vba.shortcuts) holder.vba.shortcuts = {};
  return holder.vba;
}
function setModules(doc, mods) { (doc.wb || doc).vba.modules = mods; }
export function hasMacros(doc) { const p = (doc.wb || doc).vba; return !!(p && p.modules && p.modules.some(m => /\b(sub|function)\b/i.test(m.code))); }
export function listMacros(doc) {
  const out = [];
  for (const m of project(doc).modules) {
    if (m.type === 'jsa') { jsaFunctions(m.code).forEach(n => out.push({ name: n, module: m.name + ' (JSA)', jsa: true })); continue; }
    let mod; try { mod = parseModule(m.code, m.name); } catch (e) { continue; }
    for (const k of mod.order) { const pr = mod.procs[k]; if (pr.kind === 'sub' && !pr.isPrivate && pr.params.every(p => p.optional || p.paramArray)) out.push({ name: pr.name, module: m.name }); }
  }
  return out;
}

// ---------------------------------------------------------------- UI services used by the engine
function makeUi(doc, out) {
  return {
    userName: settings().userName,
    print: t => { out.push(t); if (doc._vbaImmediate) doc._vbaImmediate(t); },
    status: t => toast(t),
    msgbox: async (prompt, flags, title) => {
      const btnSet = flags & 7; const icon = flags & 0x70;
      const sets = { 0: [['OK', 1]], 1: [['OK', 1], ['Hủy', 2]], 2: [['Hủy bỏ', 3], ['Thử lại', 4], ['Bỏ qua', 5]], 3: [['Có', 6], ['Không', 7], ['Hủy', 2]], 4: [['Có', 6], ['Không', 7]], 5: [['Thử lại', 4], ['Hủy', 2]] };
      const ico = icon === 16 ? 'circle-x' : icon === 32 ? 'circle-help' : icon === 48 ? 'triangle-alert' : icon === 64 ? 'info' : null;
      const defIdx = (flags & 0x300) >> 8;
      const btns = (sets[btnSet] || sets[0]).map((b, i) => ({ label: b[0], value: b[1], primary: i === defIdx }));
      if (out.quiet) { out.push('[MsgBox] ' + prompt); return (btns[defIdx] || btns[0]).value; }
      const body = h('div.vba-msg', ico ? h('span.vba-msg-ic.ic-' + icon, ic(ico, 32)) : null, h('div.vba-msg-t', prompt));
      const r = await dialog({ title, body, buttons: btns, width: 420 });
      return r == null ? (btnSet === 0 ? 1 : 2) : r;
    },
    inputbox: async (prompt, title, def) => {
      const inp = h('input.inp', { value: def || '' });
      const r = await dialog({ title, body: h('div.form', h('div.dlg-msg', prompt), inp), buttons: [{ label: 'OK', primary: true, run: () => inp.value }, { label: 'Hủy', value: null }], width: 420 });
      return r == null ? null : r;
    }
  };
}

// ---------------------------------------------------------------- run
let running = null;
export function isRunning() { return !!running; }
export function stopMacro() { if (running) running.cancelled = true; }
async function buildInterp(doc, out) {
  const mods = [];
  for (const m of project(doc).modules) if (m.type !== 'jsa') { const pm = parseModule(m.code, m.name); pm.type = m.type || 'module'; mods.push(pm); }
  let model, host;
  if (doc.kind === 'word') { const W = await import('../../word/vbaword.js'); model = new W.WordModel(doc, makeUi(doc, out)); host = W.installWord(model); }
  else if (doc.kind === 'ppt') { const PV = await import('../../ppt/vbappt.js'); model = new PV.PptModel(doc, makeUi(doc, out)); model.templates = await import('../../ppt/templates.js'); host = PV.installPpt(model); }
  else { model = new ExcelModel(doc, makeUi(doc, out)); model.fmod = { shiftFormula }; model.mmod = { wbToJSON, wbFromJSON }; host = installDispatch(model); }
  const ip = new Interp(host); ip.load(mods);
  ip.optionExplicit = false;
  ip.deadline = Date.now() + 5 * 60 * 1000;
  await ip.init();
  return { ip, model };
}
export async function runMacro(doc, name, opts) {
  opts = opts || {};
  if (running) { toast('Đang có macro chạy — hãy đợi hoặc bấm Dừng', 'err'); return; }
  if (doc.grid && doc.grid.editing) doc.grid.commit();
  const out = []; if (opts.quiet) out.quiet = true;
  if (doc.kind === 'word') { doc.saveCaret && doc.saveCaret(); doc.pushHistory(); }
  if (doc.kind === 'ppt') doc.pushHistory();
  const s0 = doc.wb ? doc.wb.cur : null; const before = doc.wb ? doc.snapRange(s0, 'book') : null; const selBefore = s0 ? JSON.stringify(s0.sel) : null;
  let model, ip; const t0 = Date.now();
  const stat = doc.shell.statusMsg ? doc.shell.statusMsg('Đang chạy macro ' + name + '… (Esc để dừng)') : null;
  const esc = e => { if (e.key === 'Escape' || (e.ctrlKey && e.key === 'Pause')) stopMacro(); };
  document.addEventListener('keydown', esc, true);
  try {
    const jsaMods = project(doc).modules.filter(m => m.type === 'jsa');
    if (jsaMods.some(m => jsaHas(m.code, name)) && !project(doc).modules.some(m => m.type !== 'jsa' && new RegExp('\\b(sub|function)\\s+' + name + '\\b', 'i').test(m.code))) {
      if (!window.Proxy) throw new Error('Macro JSA (JavaScript) cần WebView/Chrome 49 trở lên — hãy cập nhật "Android System WebView".');
      model = new ExcelModel(doc, makeUi(doc, out)); model.fmod = { shiftFormula }; model.mmod = { wbToJSON, wbFromJSON }; installDispatch(model);
      runJsa(model, jsaMods, name);
    } else {
      ({ ip, model } = await buildInterp(doc, out));
      running = ip;
      await ip.run(name, opts.args || []);
    }
  } catch (e) {
    const line = e && e.vbaLine ? ' (dòng ' + e.vbaLine + ')' : '';
    if (!opts.quiet) await alertDlg(e instanceof VbaError ? 'Lỗi thời gian chạy' + (e.number ? " '" + e.number + "'" : '') + line + ':\n' + e.message : 'Lỗi: ' + (e && e.message || e), 'Microsoft Visual Basic for Applications');
    opts.onError && opts.onError(e);
  } finally {
    running = null; document.removeEventListener('keydown', esc, true); stat && stat.done && stat.done();
    finishRun(doc, model, before, selBefore, 'Macro: ' + name);
  }
  return out;
}
function finishRun(doc, model, before, selBefore, label) {
  if (!model) return;
  if (doc.kind === 'ppt') { if (model.changed) { doc.pushHistory(); doc.markDirty(); } doc.cur = Math.max(0, Math.min(doc.cur, doc.slides.length - 1)); doc.renderThumbs(); doc.renderSlide(); doc.refreshUi && doc.refreshUi(); if (model.runShow) import('../../ppt/cmds.js').then(C => C.startShow(doc, 0)); if (model.saveRequested) doc.shell.save(doc, model.saveRequested === 'saveas'); return; }
  if (doc.kind === 'word') { try { model && model.tidy && model.tidy(); } catch (e) { /* ignore */ } doc.pushHistory(); doc.updatePages(); doc.refreshUi(); doc.updateRuler && doc.updateRuler(); if (model.saveRequested) doc.shell.save(doc, model.saveRequested === 'saveas'); return; }
  const after = doc.snapRange(doc.wb.cur, 'book');
  if (after.json && JSON.stringify(after.json) !== JSON.stringify(before.json)) {
    doc.undoStack.push({ label, undo: () => doc.restoreSnap(doc.wb.cur, before), redo: () => doc.restoreSnap(doc.wb.cur, after), sheet: doc.wb.cur });
    doc.redoStack = [];
  }
  doc.wb.sheets.forEach(s => { if (s._ax) { s._ax.r.dirty = s._ax.c.dirty = true; } s.wb = doc.wb; });
  doc.calc = doc.calc || null; doc.recalc(true);
  doc.renderTabs(); doc.activateSheet(doc.wb.cur);
  const s = doc.wb.cur; if (s.sel) doc.grid.scrollTo(s.sel.r, s.sel.c);
  doc.afterChange(true); doc.markDirty();
  if (model.saveRequested) doc.shell.save(doc, model.saveRequested === 'saveas');
}
// run a single statement / expression typed in the Immediate window
export async function runImmediate(doc, line, print) {
  const t = line.trim(); if (!t) return;
  const code = /^\?/.test(t) ? 'Sub __Immediate()\nDebug.Print ' + t.slice(1) + '\nEnd Sub' : 'Sub __Immediate()\n' + t + '\nEnd Sub';
  const out = [];
  const prev = project(doc).modules; const tmp = prev.concat([{ name: '__Immediate', code }]);
  setModules(doc, tmp);
  const before = doc.wb ? doc.snapRange(doc.wb.cur, 'book') : null; let model;
  try { const b = await buildInterp(doc, out); model = b.model; await b.ip.run('__Immediate'); }
  catch (e) { print('Lỗi: ' + (e.message || e)); }
  finally { setModules(doc, prev); finishRun(doc, model, before, null, 'Cửa sổ Immediate'); }
}
export function checkSyntax(code, name, type) { try { if (type === 'jsa') new Function(code); else parseModule(code, name); return null; } catch (e) { return e; } }
export const JSA_SAMPLE = `// JSA1 — macro JavaScript kiểu WPS JSA (JavaScript for Applications), chạy ngoại tuyến trong Yomi-Office Open Source.
// Mẫu dựa trên thư viện mẫu JSA của dự án okazaki112/Jsa-vibcoding-skill (MIT).

// Tìm hàng cuối có dữ liệu của một cột
function GetLastRow(ws, col) {
    const lastCell = ws.Cells(ws.Rows.Count, col).End(xlUp)
    return lastCell.Row
}

// Tô màu xen kẽ + in đậm tiêu đề cho vùng dữ liệu bắt đầu từ A1
function ToMauBang() {
    const ws = ActiveSheet
    const last = GetLastRow(ws, 1)
    ws.Range("A1").EntireRow.Font.Bold = true
    for (let r = 2; r <= last; r += 2) {
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Interior.Color = RGB(226, 239, 218)
    }
    console.log("Đã tô " + Math.floor((last - 1) / 2) + " hàng")
}

// Tổng hợp nhanh: tổng & trung bình cột B, ghi vào ô kế bên vùng chọn
function TongHop() {
    const ws = ActiveSheet
    const last = GetLastRow(ws, 2)
    const vung = ws.Range("B2:B" + last)
    ws.Range("D1").Value2 = "Tổng"
    ws.Range("E1").Value2 = WorksheetFunction.Sum(vung)
    ws.Range("D2").Value2 = "Trung bình"
    ws.Range("E2").Value2 = WorksheetFunction.Average(vung)
    MsgBox("Tổng = " + ws.Range("E1").Value2)
}
`;

// ---------------------------------------------------------------- Visual Basic Editor (Alt+F11)
export function openEditor(doc, focusProc) {
  const P = project(doc);
  if (!P.modules.length) { if (doc.kind === 'ppt') { import('../../ppt/vbappt.js').then(PV => { if (!P.modules.length) { P.modules.push({ name: 'Module1', code: PV.PPT_SAMPLE }); openEditor(doc, focusProc); } }); return; } P.modules.push({ name: doc.kind === 'word' ? 'NewMacros' : 'Module1', code: doc.kind === 'word' ? WORD_SAMPLE : SAMPLE }); }
  let cur = P.modules[0];
  if (focusProc) { const f = P.modules.find(m => new RegExp('\\b(sub|function)\\s+' + focusProc + '\\b', 'i').test(m.code)); if (f) cur = f; }
  const tree = h('div.vbe-tree');
  const gutter = h('div.vbe-gutter');
  const ta = h('textarea.vbe-code', { spellcheck: false, autocapitalize: 'off', autocomplete: 'off', wrap: 'off' });
  const procSel = h('select.inp.small.vbe-procs');
  const imm = h('div.vbe-imm-out'); const immIn = h('input.inp.vbe-imm-in', { placeholder: 'Cửa sổ Immediate — gõ câu lệnh VBA rồi Enter (vd: ? Range("A1").Value)' });
  const status = h('div.vbe-status');
  doc._vbaImmediate = t => { imm.appendChild(h('div', t)); imm.scrollTop = imm.scrollHeight; };
  const save = () => { cur.code = ta.value; doc.markDirty(); };
  const refreshProcs = () => { procSel.innerHTML = ''; procSel.appendChild(h('option', { value: '' }, '(Chọn thủ tục)'));
    if (cur.type === 'jsa') { const er = checkSyntax(ta.value, cur.name, 'jsa'); status.textContent = er ? 'JS: ' + er.message : 'JSA (JavaScript) — không có lỗi cú pháp'; status.className = 'vbe-status ' + (er ? 'err' : 'ok'); jsaFunctions(ta.value).forEach(n => procSel.appendChild(h('option', { value: n }, 'function ' + n + '()'))); return; }
    let mod = null; try { mod = parseModule(ta.value, cur.name); status.textContent = 'Không có lỗi cú pháp'; status.className = 'vbe-status ok'; } catch (e) { status.textContent = e.message; status.className = 'vbe-status err'; } if (mod) mod.order.forEach(k => procSel.appendChild(h('option', { value: mod.procs[k].name }, mod.procs[k].kind === 'sub' ? 'Sub ' + mod.procs[k].name : 'Function ' + mod.procs[k].name))); };
  const lines = () => { const n = ta.value.split('\n').length; if (gutter.childNodes.length !== n) { gutter.textContent = ''; const f = document.createDocumentFragment(); for (let i = 1; i <= n; i++) f.appendChild(h('div', String(i))); gutter.appendChild(f); } gutter.scrollTop = ta.scrollTop; };
  const drawTree = () => { tree.innerHTML = ''; tree.appendChild(h('div.vbe-proj', ic('folder-open', 14), ' ' + (doc.kind === 'word' ? 'Project' : 'VBAProject') + ' (' + esc(doc.name) + ')')); tree.appendChild(h('div.vbe-grp', ic('folder', 13), doc.kind === 'word' ? ' Microsoft Word Objects' : doc.kind === 'ppt' ? ' Microsoft PowerPoint Objects' : ' Microsoft Excel Objects'));
    if (doc.wb) doc.wb.sheets.forEach((s, i) => tree.appendChild(h('div.vbe-item.muted', ic('sheet', 13), ' Sheet' + (i + 1) + ' (' + s.name + ')'))); else tree.appendChild(h('div.vbe-item.muted', ic('file-text', 13), ' ThisDocument'));
    tree.appendChild(h('div.vbe-grp', ic('folder', 13), ' Modules'));
    P.modules.forEach(m => { const it = h('div.vbe-item' + (m === cur ? '.on' : ''), ic(m.type === 'jsa' ? 'braces' : 'file-code', 13), ' ' + m.name + (m.type === 'jsa' ? ' (JSA)' : '')); it.addEventListener('click', () => { save(); cur = m; load(); drawTree(); }); it.addEventListener('dblclick', () => renameMod(m)); tree.appendChild(it); }); };
  const load = () => { ta.value = cur.code; lines(); refreshProcs(); };
  const renameMod = m => promptDlg('Tên mô-đun:', m.name, 'Đổi tên').then(v => { if (v && /^[A-Za-z_][A-Za-z0-9_]*$/.test(v)) { m.name = v; drawTree(); doc.markDirty(); } else if (v) toast('Tên không hợp lệ', 'err'); });
  const procAtCaret = () => { const before = ta.value.slice(0, ta.selectionStart); if (cur.type === 'jsa') { const re2 = /function\s+([\w$\u00C0-\u1EF9]+)\s*\(\s*\)/g; let m2, l2 = null; while ((m2 = re2.exec(before))) l2 = m2[1]; return procSel.value || l2; } const re = /^\s*(?:public\s+|private\s+)?sub\s+([A-Za-z_À-ỹ][\wÀ-ỹ]*)/gim; let m, last = null; while ((m = re.exec(before))) last = m[1]; return procSel.value || last; };
  const run = async () => { save(); const err = checkSyntax(cur.code, cur.name, cur.type); if (err) { alertDlg(err.message, 'Lỗi biên dịch'); return; } let name = procAtCaret(); if (!name) { const ms = listMacros(doc); if (!ms.length) return toast('Không có Sub nào để chạy', 'err'); name = await pickMacro(doc, ms); if (!name) return; } await runMacro(doc, name); };
  ta.addEventListener('input', () => { lines(); clearTimeout(ta._t); ta._t = setTimeout(() => { save(); refreshProcs(); }, 500); });
  ta.addEventListener('scroll', () => { gutter.scrollTop = ta.scrollTop; });
  ta.addEventListener('keydown', e => {
    if (e.key === 'Tab') { e.preventDefault(); const a = ta.selectionStart, b = ta.selectionEnd; if (ta.setRangeText) ta.setRangeText('    ', a, b, 'end'); else { ta.value = ta.value.slice(0, a) + '    ' + ta.value.slice(b); ta.selectionStart = ta.selectionEnd = a + 4; } lines(); }
    else if (e.key === 'F5') { e.preventDefault(); run(); }
    else if (e.key === 'Enter') { const a = ta.selectionStart; const ln = ta.value.slice(0, a).split('\n').pop(); const ind = /^\s*/.exec(ln)[0]; const extra = /^\s*(if .* then\s*$|for |do\b|while |with |select case|sub |function |else\b|elseif|case )/i.test(ln) ? '    ' : ''; e.preventDefault(); ta.setRangeText('\n' + ind + extra, a, ta.selectionEnd, 'end'); lines(); }
    e.stopPropagation();
  });
  procSel.addEventListener('change', () => { if (!procSel.value) return; const i = ta.value.search(new RegExp('^\\s*(public\\s+|private\\s+)?(sub|function)\\s+' + procSel.value + '\\b', 'im')); if (i >= 0) { ta.focus(); ta.setSelectionRange(i, i); ta.scrollTop = Math.max(0, ta.value.slice(0, i).split('\n').length - 3) * 18; } });
  immIn.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') { const v = immIn.value; imm.appendChild(h('div.muted', v)); immIn.value = ''; save(); runImmediate(doc, v, t => doc._vbaImmediate(t)); } });
  const tb = (icon, title, fn) => { const b = h('button.vbe-tb', { title }, ic(icon, 15)); b.addEventListener('click', fn); return b; };
  const toolbar = h('div.vbe-toolbar',
    tb('play', 'Chạy Sub (F5)', run), tb('square', 'Dừng (Esc)', stopMacro), tb('check-check', 'Kiểm tra cú pháp (Debug > Compile)', () => { save(); const e = checkSyntax(cur.code, cur.name, cur.type); if (e) alertDlg(e.message, 'Lỗi biên dịch'); else toast('Biên dịch thành công — không có lỗi', 'ok'); }),
    h('span.vbe-sep'),
    tb('file-plus', 'Chèn mô-đun mới (Insert > Module)', () => { save(); let i = P.modules.length + 1; while (P.modules.some(m => m.name === 'Module' + i)) i++; const m = { name: 'Module' + i, code: "Option Explicit\n\nSub Macro" + i + "()\n    \nEnd Sub\n" }; P.modules.push(m); cur = m; load(); drawTree(); doc.markDirty(); }),
    tb('braces', 'Chèn mô-đun JSA — macro JavaScript kiểu WPS', () => { save(); let i = 1; while (P.modules.some(m => m.name === 'JSA' + i)) i++; const m = { name: 'JSA' + i, type: 'jsa', code: i === 1 ? JSA_SAMPLE : '// macro JavaScript (JSA)\nfunction Macro' + i + '() {\n    \n}\n' }; P.modules.push(m); cur = m; load(); drawTree(); doc.markDirty(); }),
    tb('library', 'Thư viện macro mẫu từ các dự án GitHub', () => { save(); import('./library.js').then(L => L.libraryDialog(doc, m => { cur = m; load(); drawTree(); })); }),
    tb('pencil', 'Đổi tên mô-đun', () => renameMod(cur)),
    tb('trash-2', 'Xóa mô-đun', async () => { if (!(await confirmDlg('Xóa mô-đun "' + cur.name + '"? (không thể hoàn tác)', 'Xóa mô-đun', 'Xóa', 'Hủy'))) return; P.modules.splice(P.modules.indexOf(cur), 1); if (!P.modules.length) P.modules.push({ name: 'Module1', code: '' }); cur = P.modules[0]; load(); drawTree(); doc.markDirty(); }),
    h('span.vbe-sep'),
    tb('file-down', 'Nhập tệp .bas (Import File)', async () => { const fs = await Host.openFiles('.bas,.txt,.vb,.cls', true); fs.forEach(f => { const txt = new TextDecoder('utf-8').decode(f.bytes); const nm = (/Attribute VB_Name = "([^"]+)"/.exec(txt) || [])[1] || f.name.replace(/\.[^.]+$/, ''); P.modules.push({ name: nm.replace(/\W/g, '_'), code: txt.replace(/^Attribute VB_.*\r?\n/gm, '') }); }); if (fs.length) { cur = P.modules[P.modules.length - 1]; load(); drawTree(); doc.markDirty(); } }),
    tb('file-up', 'Xuất tệp .bas (Export File)', () => { save(); Host.saveFile(cur.name + '.bas', 'Attribute VB_Name = "' + cur.name + '"\r\n' + cur.code.replace(/\r?\n/g, '\r\n')); }),
    h('span.vbe-sep'), procSel,
    h('span.sp'), h('span.muted.small', 'F5: chạy · Esc: dừng · Alt+F8: Macro'));
  const body = h('div.vbe', toolbar, h('div.vbe-main', tree, h('div.vbe-editor', h('div.vbe-codewrap', gutter, ta), status)), h('div.vbe-imm', h('div.vbe-imm-h', 'Immediate'), imm, immIn));
  drawTree(); load();
  return dialog({ title: 'Microsoft Visual Basic for Applications — ' + doc.name, width: 1100, body, cls: 'vbe-dlg', enterKey: false, buttons: [{ label: 'Chạy (F5)', primary: true, run: () => { run(); return false; } }, { label: 'Đóng', run: () => { save(); return true; } }] })
    .then(() => { save(); doc._vbaImmediate = null; });
}
function pickMacro(doc, list) {
  const sel = h('select.inp', { size: 8, style: { width: '100%' } }, list.map(m => h('option', { value: m.name }, m.name + '  (' + m.module + ')')));
  sel.selectedIndex = 0;
  return dialog({ title: 'Chọn macro', body: sel, buttons: [{ label: 'Chạy', primary: true, run: () => sel.value }, { label: 'Hủy', value: null }] });
}

// ---------------------------------------------------------------- Macros dialog (Alt+F8)
export function macrosDialog(doc) {
  const P = project(doc);
  const nameIn = h('input.inp', { placeholder: 'Tên macro' });
  const list = h('div.vba-mlist');
  const desc = h('div.muted.small');
  const draw = () => { list.innerHTML = ''; const ms = listMacros(doc); if (!ms.length) list.appendChild(h('div.muted', 'Chưa có macro. Bấm "Tạo" hoặc "Ghi Macro".')); ms.forEach((m, i) => { const it = h('div.vba-mi' + (i === 0 ? '.on' : ''), m.name, h('span.muted', '  ' + m.module)); it.addEventListener('click', () => { $$('.vba-mi', list).forEach(x => x.classList.toggle('on', x === it)); nameIn.value = m.name; const sc = Object.keys(P.shortcuts).find(k => P.shortcuts[k] === m.name); desc.textContent = sc ? 'Phím tắt: Ctrl+Shift+' + sc.toUpperCase() : ''; }); it.addEventListener('dblclick', () => { close(); runMacro(doc, m.name); }); list.appendChild(it); if (i === 0) nameIn.value = m.name; }); };
  let close = () => {};
  const body = h('div.form', h('label.lbl', 'Tên macro:'), nameIn, list, desc, h('div.hint', 'Macro VBA được lưu cùng sổ làm việc (.xlsm) và chạy hoàn toàn ngoại tuyến. Mở tệp .xlsm từ Excel sẽ tự nạp macro.'));
  draw();
  return dialog({ title: 'Macro', width: 520, body, onOpen: (b, c) => { close = c; }, buttons: [
    { label: 'Chạy', primary: true, run: () => { if (!nameIn.value) return false; setTimeout(() => runMacro(doc, nameIn.value), 0); return true; } },
    { label: 'Sửa', run: () => { setTimeout(() => openEditor(doc, nameIn.value), 0); return true; } },
    { label: 'Tạo', run: () => { const nm = (nameIn.value || 'Macro' + (listMacros(doc).length + 1)).replace(/\W/g, '_'); if (!P.modules.length) P.modules.push({ name: 'Module1', code: 'Option Explicit\n' }); P.modules[0].code += '\n\nSub ' + nm + '()\n    \nEnd Sub\n'; doc.markDirty(); setTimeout(() => openEditor(doc, nm), 0); return true; } },
    { label: 'Xóa', danger: true, run: async () => { const nm = nameIn.value; if (!nm) return false; if (!(await confirmDlg('Xóa macro "' + nm + '"?', 'Xóa macro', 'Xóa', 'Hủy'))) return false; P.modules.forEach(m => { m.code = m.code.replace(new RegExp('(^|\\n)\\s*(public\\s+|private\\s+)?sub\\s+' + nm + '\\b[\\s\\S]*?\\n\\s*end\\s+sub[^\\n]*', 'i'), '$1'); }); doc.markDirty(); draw(); return false; } },
    { label: 'Tùy chọn…', run: async () => { const nm = nameIn.value; if (!nm) return false; const v = await promptDlg('Phím tắt: Ctrl+Shift+ (một chữ cái)', Object.keys(P.shortcuts).find(k => P.shortcuts[k] === nm) || '', 'Tùy chọn Macro'); if (v != null) { Object.keys(P.shortcuts).forEach(k => { if (P.shortcuts[k] === nm) delete P.shortcuts[k]; }); if (/^[a-z]$/i.test(v)) P.shortcuts[v.toLowerCase()] = nm; doc.markDirty(); draw(); } return false; } },
    { label: 'Đóng' }] });
}

// ---------------------------------------------------------------- recorder (records VBA)
export function toggleRecord(doc, relative) {
  if (doc._vbaRec) {
    const rec = doc._vbaRec; doc._vbaRec = null; doc.setInput = rec.origSetInput; if (rec.onSel) doc.grid.el.removeEventListener('mouseup', rec.onSel, true);
    const P = project(doc); if (!P.modules.length) P.modules.push({ name: 'Module1', code: 'Option Explicit\n' });
    const body = rec.lines.map(l => '    ' + l).join('\n');
    P.modules[0].code += '\n\nSub ' + rec.name + '()\n\'\n\' ' + rec.name + ' — macro đã ghi ' + new Date().toLocaleString('vi-VN') + '\n\'\n' + body + '\nEnd Sub\n';
    if (rec.key) P.shortcuts[rec.key] = rec.name;
    doc.markDirty(); doc.refreshUi && doc.refreshUi();
    toast('Đã ghi macro "' + rec.name + '" (' + rec.lines.length + ' thao tác)', 'ok');
    return;
  }
  formDlg('Ghi Macro', [{ key: 'name', label: 'Tên macro', value: 'Macro' + (listMacros(doc).length + 1) }, { key: 'key', label: 'Phím tắt Ctrl+Shift+', value: '' }, { key: 'rel', type: 'checkbox', label: 'Dùng tham chiếu tương đối', value: !!relative }]).then(v => {
    if (!v || !v.name) return;
    const rec = { name: v.name.replace(/\W/g, '_'), key: /^[a-z]$/i.test(v.key || '') ? v.key.toLowerCase() : null, lines: [], rel: !!v.rel, origSetInput: doc.setInput };
    let last = null;
    const selLine = () => { const s = doc.wb.cur; const rg = doc.grid.selRange(); const a = rg.r1 === rg.r2 && rg.c1 === rg.c2 ? addr(rg.r1, rg.c1) : addr(rg.r1, rg.c1) + ':' + addr(rg.r2, rg.c2); if (a === last) return; if (rec.rel && last) { const p = /([A-Z]+)(\d+)/.exec(last); const pr = +p[2] - 1; let pc = 0; for (const ch of p[1]) pc = pc * 26 + ch.charCodeAt(0) - 64; pc--; rec.lines.push('ActiveCell.Offset(' + (rg.r1 - pr) + ', ' + (rg.c1 - pc) + ').Range("A1").Select'); } else rec.lines.push('Range("' + a + '").Select'); last = a; };
    rec.onSel = () => setTimeout(selLine, 0);
    doc.grid.el.addEventListener('mouseup', rec.onSel, true);
    doc.setInput = function (r, c, text, fill) { selLine(); const t = String(text == null ? '' : text); rec.lines.push((/^=/.test(t) ? 'ActiveCell.Formula = ' : 'ActiveCell.Value = ') + '"' + t.replace(/"/g, '""') + '"'); const res = rec.origSetInput.call(doc, r, c, text, fill); setTimeout(() => { last = null; selLine(); }, 0); return res; };
    doc._vbaRec = rec; selLine();
    toast('Đang ghi macro "' + rec.name + '" — bấm "Dừng Ghi" trên tab Nhà phát triển để kết thúc', 'ok', 4000);
    doc.refreshUi && doc.refreshUi();
  });
}
// a recorder hook for other commands (bold, fill …) — called by Excel commands when recording
export function recordLine(doc, line) { if (doc._vbaRec) doc._vbaRec.lines.push(line); }

// Ctrl+Shift+<letter> macro shortcuts
export function shortcutFor(doc, e) {
  const P = (doc.wb || doc).vba; if (!P || !P.shortcuts) return false;
  if (!(e.ctrlKey && e.shiftKey) || e.altKey) return false;
  const k = (e.key || '').toLowerCase(); const nm = P.shortcuts[k]; if (!nm) return false;
  e.preventDefault(); runMacro(doc, nm); return true;
}
export const loadLibrary = () => import('./library.js');
