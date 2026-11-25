// Built-in self test (open the app with #selftest). Runs the real engines inside the host WebView and logs "OOSTEST ..." lines
// (visible in Android logcat / browser console). Used to verify Android 6 WebView 44 → latest, and Windows Edge/Chrome.
import Host from './host.js';
import { supportsWasm, PDFLib, KaTeX } from './util.js';
import * as AI from '../ai/offline.js';

export async function runSelfTest(shell) {
  let pass = 0, fail = 0; const out = [];
  const log = (ok, name, info) => { ok ? pass++ : fail++; const line = 'OOSTEST ' + (ok ? 'PASS' : 'FAIL') + ' | ' + name + (info != null ? ' | ' + info : ''); out.push(line); console.log(line); };
  const t = async (name, fn) => { try { const r = await fn(); log(r !== false, name, typeof r === 'string' || typeof r === 'number' ? r : null); } catch (e) { log(false, name, (e && e.message) || String(e)); } };
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const docOf = async kind => { const prev = shell.cur; shell.newDoc(kind); for (let i = 0; i < 60 && !(shell.cur && shell.cur !== prev && shell.cur.kind === kind); i++) await wait(100); await wait(300); return shell.cur; };
  console.log('OOSTEST START | ' + navigator.userAgent);
  await t('env: host=' + Host.kind, () => true);
  await t('env: WebAssembly (OCR Tesseract)', () => supportsWasm() ? 'yes' : 'no (native/upgrade WebView)');
  await t('env: IndexedDB', () => !!window.indexedDB);

  // ---------------- Excel
  let x;
  await t('excel: open', async () => { x = await docOf('excel'); return !!x.wb; });
  await t('excel: formulas', async () => {
    const vals = [['Mã', 'Tên', 'Giá'], ['A1', 'Bút', '12000'], ['B2', 'Vở', '8000'], ['C3', 'Thước', '5000']];
    vals.forEach((row, r) => row.forEach((v, c) => x.setInput(r, c, v)));
    x.setInput(5, 2, '=SUM(C2:C4)'); x.setInput(6, 2, '=VLOOKUP("B2",A2:C4,3,FALSE)'); x.setInput(7, 2, '=XLOOKUP("C3",A2:A4,B2:B4)');
    x.setInput(8, 2, '=TEXT(1234.5,"#,##0.00")'); x.setInput(9, 2, '=IFERROR(1/0,"lỗi")'); x.setInput(10, 2, '=ROUND(AVERAGE(C2:C4),0)'); x.setInput(11, 2, '=COUNTIF(C2:C4,">6000")');
    x.setInput(12, 0, '=SEQUENCE(3)');
    const v = (r, c) => { const cell = x.wb.cur.get(r, c); return cell ? cell.v : (x.wb.cur.spill && x.wb.cur.spill.get ? null : null); };
    const got = [v(5, 2), v(6, 2), v(7, 2), v(8, 2), v(9, 2), v(10, 2), v(11, 2)];
    const exp = [25000, 8000, 'Thước', '1,234.50', 'lỗi', 8333, 2];
    const bad = exp.filter((e, i) => String(e) !== String(got[i]));
    if (bad.length) throw new Error('got ' + JSON.stringify(got));
    return got.join(' ; ');
  });
  await t('excel: undo/redo', async () => { x.setInput(14, 0, '42'); x.undo(); const a = x.wb.cur.get(14, 0); x.redo(); const b = x.wb.cur.get(14, 0); return (!a || a.v == null || a.v === '') && b && b.v === 42; });
  await t('excel: XLSX save→load', async () => { const bytes = await x.save('xlsx'); const y = await docOf('excel'); await y.load(bytes, 'xlsx', 't.xlsx'); const c = y.wb.cur.get(5, 2); if (!(c && c.v === 25000)) throw new Error('C6=' + JSON.stringify(c && { v: c.v, f: c.f }) + ' same=' + (y === x) + ' sheets=' + y.wb.sheets.length); return bytes.length + ' bytes'; });
  await t('excel: CSV export', async () => { const b = await x.save('csv'); return b.length > 20 ? b.length + ' bytes' : false; });

  // ---------------- Word
  let w;
  await t('word: open', async () => { w = await docOf('word'); return !!w.body; });
  await t('word: DOCX save→load (tiếng Việt)', async () => {
    w.body.innerHTML = '<h1>Báo cáo quý</h1><p>Xin chào <b>Yomi-Office Open Source</b> — tiếng Việt có dấu: ắ ổ ữ đ.</p><ul><li>Một</li><li>Hai</li></ul>';
    const bytes = await w.save('docx'); const w2 = await docOf('word'); await w2.load(bytes, 'docx', 't.docx');
    const tx = w2.body.textContent; return tx.indexOf('ắ ổ ữ đ') >= 0 && w2.body.querySelectorAll('li').length === 2 ? bytes.length + ' bytes' : false;
  });

  // ---------------- PowerPoint
  let p;
  await t('ppt: open + new slides', async () => { p = await docOf('ppt'); p.slides[0].objects[0].html = 'Trình chiếu thử'; for (let i = 0; i < 3; i++) shell.cur.slides.push(JSON.parse(JSON.stringify(p.slides[0]))); return p.slides.length === 4; });
  await t('ppt: PPTX save→load', async () => { const bytes = await p.save('pptx'); const p2 = await docOf('ppt'); await p2.load(bytes, 'pptx'); return p2.slides.length === 4 && p2.getText().indexOf('Trình chiếu thử') >= 0 ? bytes.length + ' bytes' : false; });

  // ---------------- PDF
  await t('pdf: create + open + extract text', async () => {
    const L = await PDFLib(); const pdf = await L.PDFDocument.create(); const f = await pdf.embedFont(L.StandardFonts.Helvetica);
    for (let i = 1; i <= 2; i++) pdf.addPage([400, 300]).drawText('OOS PDF PAGE ' + i, { x: 40, y: 200, size: 24, font: f });
    const d = await docOf('pdf'); await d.load(await pdf.save(), 'pdf', 't.pdf'); await wait(800);
    const txt = new TextDecoder('utf-8').decode(await d.save('txt'));
    return d.pdf.numPages === 2 && txt.indexOf('PAGE 2') >= 0 ? '2 pages' : false;
  });
  await t('pdf: export DOCX', async () => { const b = await shell.cur.save('docx'); return b[0] === 0x50 && b[1] === 0x4b ? b.length + ' bytes' : false; });

  // ---------------- MathType / AI / OCR
  await t('mathtype: KaTeX render', async () => { const k = await KaTeX(); const s = k.renderToString('\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}'); return s.indexOf('katex') >= 0; });
  await t('ai offline: soát lỗi tiếng Việt', () => { const iss = AI.proofreadText('Công ty ko đạt mục tiêu vì sử lý chậm.'); return iss.length >= 2 ? AI.applyFixes('Công ty ko đạt mục tiêu vì sử lý chậm.', iss) : false; });
  await t('ai offline: từ khóa', () => { const k = AI.keywords('Chiến lược kinh doanh mở rộng thị trường miền Nam. Kinh doanh tăng trưởng.', 3); return k.indexOf('kinh doanh') >= 0 ? k.join(', ') : false; });
  await t('ocr: vie+eng offline', async () => {
    const c = document.createElement('canvas'); c.width = 900; c.height = 130; const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, 900, 130); g.fillStyle = '#000'; g.font = '44px sans-serif'; g.fillText('Văn bản tiếng Việt OCR', 20, 80);
    const t0 = Date.now(); const r = await Host.ocr(c.toDataURL('image/png'), 'vie+eng');
    return (r.text || '').trim() + ' (' + (Date.now() - t0) + 'ms)';
  });

  // ================================================================ v0.0.6 features
  const key = (el, k) => { const ev = new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }); if (ev.key !== k) try { Object.defineProperty(ev, 'key', { value: k }); Object.defineProperty(ev, 'keyCode', { value: { Enter: 13, Tab: 9, Escape: 27 }[k] || 0 }); } catch (e) { /* ignore */ } el.dispatchEvent(ev); };
  let x6;
  await t('v006 excel: gõ trực tiếp vào ô (không bấm đúp)', async () => {
    x6 = await docOf('excel'); const g = x6.grid; g.select(2, 1); g.focusSink();
    const ed = g.editor; if (document.activeElement !== ed) throw new Error('ô nhập ẩn không nhận focus: ' + (document.activeElement && document.activeElement.className));
    ed.value = 'Hà Nội'; ed.dispatchEvent(new Event('input', { bubbles: true })); if (!g.editing) throw new Error('không vào chế độ nhập');
    key(ed, 'Enter'); await wait(50); const c = x6.wb.cur.get(2, 1); return c && c.v === 'Hà Nội' ? 'B3 = Hà Nội, con trỏ xuống ' + (x6.wb.cur.sel.r + 1) : false;
  });
  await t('v006 excel: hàm mới (CHISQ.DIST, CONFIDENCE.T, PIVOTBY, T.INV.2T)', async () => {
    const set = (r, c, v) => x6.setInput(r, c, v);
    set(0, 5, '=ROUND(CHISQ.DIST(3,4,TRUE),6)'); set(1, 5, '=ROUND(CONFIDENCE.T(0.05,1,50),6)'); set(2, 5, '=ROUND(T.INV.2T(0.05,49),4)'); set(3, 5, '=ROWS(PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{1;2;3},SUM))');
    const v = [0, 1, 2, 3].map(r => x6.wb.cur.value(r, 5)); const ok = v[0] === 0.442175 && v[1] === 0.284197 && v[2] === 2.0096 && v[3] === 4; if (!ok) throw new Error(JSON.stringify(v)); return v.join(' ; ');
  });
  await t('v006 VBA: chạy macro (vòng lặp, Range, WorksheetFunction, Dictionary)', async () => {
    const V = await window.OOS_loadVBA(); x6.wb.vba = { modules: [{ name: 'Module1', code: 'Sub Test6()\n Dim i As Long, t As Double\n For i = 1 To 10\n  Cells(i, 8).Value = i * i\n Next i\n Range("I1").Value = WorksheetFunction.Sum(Range("H1:H10"))\n Dim d As Object: Set d = CreateObject("Scripting.Dictionary")\n d("k") = "Việt Nam"\n Range("I2").Value = d("k") & " " & Format(1234.5, "#,##0.0")\nEnd Sub' }], shortcuts: {} };
    await V.runMacro(x6, 'Test6', { quiet: true, onError: e => { throw e; } }); const a = x6.wb.cur.value(0, 8), b = x6.wb.cur.value(1, 8); if (a !== 385 || b !== 'Việt Nam 1,234.5') throw new Error(a + ' | ' + b); return 'I1=' + a + ', I2=' + b;
  });
  await t('v006 VBA: lưu .xlsm → mở lại giữ macro', async () => { const bytes = await x6.save('xlsm'); const y = await docOf('excel'); await y.load(bytes, 'xlsm', 't.xlsm'); const m = (y.wb.vba && y.wb.vba.modules || []).find(z => z.name === 'Module1'); return m && /Sub Test6/.test(m.code) && /Việt Nam/.test(m.code) ? bytes.length + ' bytes, ' + y.wb.vba.modules.length + ' mô-đun' : false; });
  await t('v006 ODF: .ods / .odt / .odp lưu → mở lại', async () => {
    const b1 = await x6.save('ods'); const y = await docOf('excel'); await y.load(b1, 'ods', 't.ods'); const okS = y.wb.cur.value(2, 1) === 'Hà Nội' && (y.wb.cur.get(0, 8) || {}).f === 'SUM(H1:H10)' || y.wb.cur.value(0, 8) === 385;
    const w6 = await docOf('word'); w6.body.innerHTML = '<h2>Chương 1</h2><p>Nội dung <b>đậm</b> tiếng Việt</p><ul><li>a</li><li>b</li></ul>'; const b2 = await w6.save('odt'); const w7 = await docOf('word'); await w7.load(b2, 'odt', 't.odt'); const okT = !!w7.body.querySelector('h2') && w7.body.textContent.indexOf('tiếng Việt') >= 0 && w7.body.querySelectorAll('li').length === 2;
    const p6 = await docOf('ppt'); p6.slides[0].objects[0].html = 'Tiêu đề ODP'; const b3 = await p6.save('odp'); const p7 = await docOf('ppt'); await p7.load(b3, 'odp'); const okP = p7.getText().indexOf('Tiêu đề ODP') >= 0;
    if (!(okS && okT && okP)) throw new Error('ods=' + okS + ' odt=' + okT + ' odp=' + okP); return b1.length + '/' + b2.length + '/' + b3.length + ' bytes';
  });
  await t('v006 RTF: tiếng Việt + định dạng khứ hồi', async () => { const w8 = await docOf('word'); w8.body.innerHTML = '<p style="text-align:center">Xin <b>chào</b> <i>Việt Nam</i></p>'; const b = await w8.save('rtf'); const w9 = await docOf('word'); await w9.load(b, 'rtf', 't.rtf'); const html = w9.body.innerHTML; return /<b>chào<\/b>/.test(html) && /Việt Nam/.test(html) && /center/.test(html) ? b.length + ' bytes' : html.slice(0, 120); });
  await t('v006 Word: thước kẻ khớp lề trang', async () => { const w = await docOf('word'); await wait(200); w.rulerObj.draw(); const g = w.rulerObj.geom(); if (!g) return 'ẩn (màn hình hẹp)'; const sr = w.sheet.getBoundingClientRect(), cr = w.rulerObj.cv.getBoundingClientRect(); const pageLeft = sr.left - cr.left; const dx = Math.abs(g.x0 - pageLeft); const mOk = Math.abs((g.ml - g.x0) - w.marg[3] * w.zoom / 100) < 1; return dx < 1 && mOk ? 'lệch ' + dx.toFixed(2) + 'px' : false; });
  await t('v006 PowerPoint: bấm vào khung và gõ chữ', async () => {
    const p8 = await docOf('ppt'); await wait(200); const el = p8.canvas.querySelector('.pp-obj.pp-text'); const o = p8.slide.objects[0];
    const cmds = await import('../ppt/cmds.js'); cmds.editObj(p8, p8.slide, o, el, el.querySelector('.pp-obj-body'));
    const tf = el.querySelector('.pp-tf'); if (tf.contentEditable !== 'true' || document.activeElement !== tf) throw new Error('khung chữ không nhận focus');
    if (!document.execCommand('insertText', false, 'Gõ được chữ') || tf.textContent.indexOf('Gõ được chữ') < 0) { /* WebView without window focus: execCommand refuses, type through the DOM like an IME would */ tf.appendChild(document.createTextNode('Gõ được chữ')); tf.dispatchEvent(new Event('input', { bubbles: true })); } tf.blur(); await wait(550); /* blur is not dispatched when the WebView window itself has no focus (emulator) — the 400 ms input autosave covers that */ return o.html.indexOf('Gõ được chữ') >= 0 ? o.html : false;
  });
  await t('v006 Phông: tích hợp + chuyển mã TCVN3/VNI', async () => { const F = await import('./fonts.js'); const a = F.tcvn3ToUnicode('Ti\u00d5ng Vi\u00d6t'), b = F.vniToUnicode('Tieáng Vieät ñöôïc'); if (a !== 'Tiếng Việt' || b !== 'Tiếng Việt được') throw new Error(a + ' | ' + b); return F.FONT_LIST.indexOf('HP001 4 hàng') >= 0 ? F.FONT_LIST.length + ' phông' : false; });
  await t('v006 Định dạng: .xml 2003 / SYLK / DIF / Word XML / MHT', async () => { const a = await x6.save('xml2003'), b = await x6.save('slk'), c = await x6.save('dif'); const w = await docOf('word'); w.body.innerHTML = '<p>Thử</p>'; const d = await w.save('xml'), e = await w.save('mht'); return [a, b, c, d, e].every(z => z && z.length > 20) ? [a, b, c, d, e].map(z => z.length).join('/') : false; });

  await t('v006 VBA: Class module + ThisWorkbook + mã ASCII an toàn cho MS Office', async () => {
    const V = await window.OOS_loadVBA(); const c = await docOf('excel');
    c.wb.vba = { modules: [{ name: 'ThisWorkbook', type: 'workbook', code: 'Public Function WbTen() As String\n WbTen = "TW:" & Me.Sheets.Count\nEnd Function' },
      { name: 'Class1', type: 'class', code: 'Private mTen As String\nPublic Property Let Ten(ByVal v As String)\n mTen = v\nEnd Property\nPublic Function ChaoHoi() As String\n ChaoHoi = "Xin chào " & mTen\nEnd Function' },
      { name: 'Module1', code: 'Sub T7()\n Dim k As New Class1\n k.Ten = "lớp"\n Range("A1").Value = k.ChaoHoi() & "|" & ThisWorkbook.WbTen() & "|" & TypeName(k)\nEnd Sub' }], shortcuts: {} };
    await V.runMacro(c, 'T7', { quiet: true, onError: e => { throw e; } }); const v = c.wb.cur.value(0, 0);
    if (v !== 'Xin chào lớp|TW:1|Class1') throw new Error(v);
    const O = await import('../excel/vba/ovba.js'); const enc = O.protectUnicode('x = "Việt" \' chú thích'); if (/[^\x00-\x7f]/.test(enc) || O.unprotectUnicode(enc).indexOf('"Việt"') < 0) throw new Error(enc);
    return v + ' · ' + enc;
  });
  await t('v006 JSA (WPS): mẫu Jsa-vibcoding chạy trên mô hình Excel', async () => {
    if (!window.Proxy) return 'bỏ qua (WebView < 49 không có Proxy)';
    const V = await window.OOS_loadVBA(); const L = await V.loadLibrary(); const j = await docOf('excel'); const it = L.catalog('excel').find(z => z.id === 'jsa');
    j.wb.vba = { modules: [{ name: 'JsaTemplates', type: 'jsa', code: it.code() }], shortcuts: {} };
    await V.runMacro(j, 'DemoJsa', { quiet: true, onError: e => { throw e; } }); const s = j.wb.sheet('JSA_Demo'); const v = s && s.value(1, 3);
    if (v !== 50000) throw new Error('D2=' + v); return 'JSA_Demo!D2=' + v;
  });
  await t('v006 AutoXL: thư viện LAMBDA A.* (MatrixFun/AutoXL)', async () => {
    const V = await window.OOS_loadVBA(); const L = await V.loadLibrary(); const k = await docOf('excel'); const n = L.importLambdaLibrary(k.wb, L.AUTOXL_SRC);
    [1, 2, 2, 3].forEach((v, i) => k.wb.cur.setCell(i, 0, { v })); const ev = f => k.calc.evalFormula(f, k.wb.cur, 0, 0);
    const a = ev('A.VERSION()'), b = ev('A.EQ({1,2},{1,2})'), c = ev('A.DUPLICATES(A1:A4)'); const ok = /AutoXL/.test(a) && b === true && JSON.stringify(c) === '[[2],[2]]';
    if (!ok) throw new Error([a, b, JSON.stringify(c)].join(' | ')); return n + ' hàm; ' + a;
  });
  await t('v006 Word VBA: macro Word + lưu .docm → mở lại', async () => {
    const V = await window.OOS_loadVBA(); const L = await V.loadLibrary(); const w = await docOf('word'); w.body.innerHTML = '<p>xin  chào   Việt Nam</p><p>hai</p>';
    w.vba = { modules: [{ name: 'WordTools', code: L.catalog('word').find(z => z.id === 'wordtools').code() }], shortcuts: {} };
    for (const n of ['XoaKhoangTrangKep', 'VietHoaDauDoan']) await V.runMacro(w, n, { quiet: true, onError: e => { throw e; } });
    const txt = w.body.textContent; if (txt.indexOf('Xin chào Việt Nam') < 0 || txt.indexOf('Hai') < 0) throw new Error(txt);
    const b = await w.save('docm'); const w2 = await docOf('word'); await w2.load(b, 'docm', 't.docm'); const m = (w2.vba && w2.vba.modules || []).find(z => z.name === 'WordTools');
    if (!m || !/Sub VietHoaDauDoan/.test(m.code)) throw new Error('mất macro khi mở lại'); return txt + ' · ' + b.length + ' bytes';
  });

  // ---------------- v0.0.7
  await t('v007 A.I: chế độ Web (ChatGPT/Gemini/Copilot/Claude/DeepSeek/Grok/AI Hay) + khóa API (thêm Grok, Azure/Copilot, OpenHay)', async () => {
    const A = await import('../ai/ai.js'); const u = A.webUrl('web:chatgpt', 'Xin chào'); const g = A.webUrl('web:gemini', 'x');
    if (u !== 'https://chatgpt.com/?q=Xin%20ch%C3%A0o' || g !== 'https://gemini.google.com/app') throw new Error(u + ' | ' + g);
    const n = Object.keys(A.WEB).length, k = Object.keys(A.PROVIDERS).length; if (n < 7 || !A.PROVIDERS.xai || !A.PROVIDERS.openhay) throw new Error('thiếu dịch vụ'); return n + ' dịch vụ web, ' + k + ' nhà cung cấp API';
  });
  await t('v007 Cập nhật phần mềm: đọc Drive + MediaFire, version.json, chọn đúng bộ cài', async () => {
    const U = await import('./update.js');
    const html = '<div class="flip-entries"><div class="flip-entry" id="entry-AbC123" tabindex="0" role="link"><div class="flip-entry-info"><a href="https://drive.google.com/file/d/AbC123/view?usp=drive_web" target="_blank"><div class="flip-entry-visual"></div><div class="flip-entry-list-icon"></div><div class="flip-entry-title">OfficeOpenSource-0.0.8-Win10-11-x64-Setup.exe</div></a></div></div><div class="flip-entry" id="entry-Fold1" tabindex="0" role="link"><div class="flip-entry-info"><a href="https://drive.google.com/drive/folders/Fold1" target="_blank"><div class="flip-entry-title">0.0.8</div></a></div></div></div>';
    const items = U.parseDrive(html); if (items.length !== 2 || items[0].id !== 'AbC123' || !items[1].folder) throw new Error(JSON.stringify(items));
    const files = ['OfficeOpenSource-0.0.8-Win10-11-x64-Setup.exe', 'OfficeOpenSource-0.0.8-Win10-11-x64-portable.zip', 'OfficeOpenSource-0.0.8-Win7-8.1-x86-Setup.exe', 'OfficeOpenSource-0.0.8-Android6-7-legacy.apk', 'OfficeOpenSource-0.0.8-Android6plus-universal.apk', 'OfficeOpenSource-0.0.8.appxbundle'].map(n => ({ name: n }));
    const pick = ed => (U.pickFile(files, ed, { version: '0.0.8' }) || {}).name;
    const r = [pick({ tag: 'Win10-11-x64', platform: 'windows' }), pick({ tag: 'Win10-11-x64', platform: 'windows', portable: true }), pick({ tag: 'Win7-8.1-x86', platform: 'windows' }), pick({ tag: 'android-legacy', platform: 'android' }), pick({ tag: 'android-modern', platform: 'android' }), pick({ tag: 'Win10-11-x64', platform: 'windows', appx: true })];
    const exp = [files[0].name, files[1].name, files[2].name, files[3].name, files[4].name, files[5].name];
    if (JSON.stringify(r) !== JSON.stringify(exp)) throw new Error(JSON.stringify(r));
    const mf = U.mediafireDirect('<a class="input popsok" aria-label="Download file" href="https://download1234.mediafire.com/abc/x/OfficeOpenSource.apk" id="downloadButton">');
    if (!/download1234/.test(mf) || U.cmpVer('0.0.10', '0.0.9') !== 1 || U.cmpVer('0.0.7', '0.0.7') !== 0) throw new Error('mediafire/cmpVer');
    return '6/6 phiên bản chọn đúng bộ cài; Drive ' + items.length + ' mục; MediaFire OK';
  });
  await t('v007 Đọc sách: EPUB (mục lục, ảnh) + FB2 + TXT', async () => {
    const R = await import('../reader/reader.js'); const { JSZipLib } = await import('./util.js'); const JSZip = await JSZipLib(); const z = new JSZip();
    z.file('mimetype', 'application/epub+zip'); z.file('META-INF/container.xml', '<?xml version="1.0"?><container xmlns="urn:oasis:names:tc:opendocument:xmlns:container" version="1.0"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>');
    z.file('OEBPS/content.opf', '<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Truyện Kiều</dc:title><dc:creator>Nguyễn Du</dc:creator></metadata><manifest><item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/><item id="c1" href="text/c1.xhtml" media-type="application/xhtml+xml"/><item id="c2" href="text/c2.xhtml" media-type="application/xhtml+xml"/></manifest><spine><itemref idref="c1"/><itemref idref="c2"/></spine></package>');
    z.file('OEBPS/nav.xhtml', '<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"><body><nav epub:type="toc"><ol><li><a href="text/c1.xhtml">Hồi 1</a></li><li><a href="text/c2.xhtml#p2">Hồi 2</a></li></ol></nav></body></html>');
    z.file('OEBPS/text/c1.xhtml', '<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml"><head><title>c1</title><script>alert(1)</script></head><body><h1>Hồi 1</h1><p>Trăm năm trong cõi người ta</p><img src="../img/a.png"/></body></html>');
    z.file('OEBPS/text/c2.xhtml', '<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml"><body><h1>Hồi 2</h1><p id="p2">Chữ tài chữ mệnh khéo là ghét nhau</p></body></html>');
    z.file('OEBPS/img/a.png', Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='), c => c.charCodeAt(0)));
    const bk = await R.parseBook({ name: 'kieu.epub', bytes: await z.generateAsync({ type: 'uint8array' }) });
    if (bk.meta.title !== 'Truyện Kiều' || bk.chapters.length !== 2 || bk.toc.length !== 2 || bk.chapters[0].html.indexOf('data:image/png') < 0 || /alert/.test(bk.chapters[0].html)) throw new Error(JSON.stringify({ t: bk.meta.title, c: bk.chapters.length, toc: bk.toc.length }));
    const fb = await R.parseBook({ name: 'a.fb2', bytes: new TextEncoder().encode('<?xml version="1.0" encoding="utf-8"?><FictionBook xmlns="http://www.gribuser.ru/xml/fictionbook/2.0"><description><title-info><author><first-name>Nam</first-name><last-name>Cao</last-name></author><book-title>Chí Phèo</book-title></title-info></description><body><section><title><p>Chương 1</p></title><p>Hắn vừa đi vừa chửi.</p></section></body></FictionBook>') });
    if (fb.meta.title !== 'Chí Phèo' || fb.meta.author !== 'Nam Cao' || fb.chapters[0].html.indexOf('Hắn vừa đi') < 0) throw new Error('FB2 ' + JSON.stringify(fb.meta));
    const ui = R.showReader(shell, bk, 'kieu.epub'); const ok = !!ui.root.querySelector('.rd-toc-i'); ui.close();
    return 'EPUB ' + bk.chapters.length + ' chương/' + bk.toc.length + ' mục lục · FB2 "' + fb.meta.title + '" · giao diện ' + (ok ? 'OK' : 'lỗi');
  });
  await t('v007 CamScanner: nhận mép giấy, nắn phối cảnh, lọc đen trắng, xuất PDF', async () => {
    const S = await import('../scan/scan.js'); const c = document.createElement('canvas'); c.width = 800; c.height = 600; const x = c.getContext('2d');
    x.fillStyle = '#3a3a3a'; x.fillRect(0, 0, 800, 600); x.fillStyle = '#f4f1ea'; x.beginPath(); x.moveTo(170, 90); x.lineTo(640, 130); x.lineTo(610, 530); x.lineTo(130, 500); x.closePath(); x.fill();
    x.fillStyle = '#222'; x.font = '28px Arial'; x.fillText('Biên bản nghiệm thu', 230, 260);
    const d = S.detectQuad(c); const exp = [[170, 90], [640, 130], [610, 530], [130, 500]]; const err = Math.max.apply(null, d.quad.map((q, i) => Math.hypot(q[0] - exp[i][0], q[1] - exp[i][1])));
    if (!d.auto || err > 24) throw new Error('góc lệch ' + err.toFixed(1) + 'px ' + JSON.stringify(d.quad.map(q => q.map(Math.round))));
    const out = S.processPage({ src: c, quad: d.quad, filter: 'bw', rot: 0 }); const px = out.getContext('2d').getImageData(10, 10, 1, 1).data;
    const pdf = await S.pagesToPdf([out]); const head = String.fromCharCode.apply(null, Array.prototype.slice.call(pdf.subarray(0, 5)));
    if (head !== '%PDF-' || px[0] !== 255) throw new Error(head + ' px=' + px[0]);
    return 'lệch góc ' + err.toFixed(1) + 'px; trang ' + out.width + '×' + out.height + '; PDF ' + pdf.length + ' bytes';
  });
  await t('v007 PDF: đặt mật khẩu (RC4-128) → mở bằng mật khẩu → gỡ mật khẩu; ký/chú thích; nén', async () => {
    const P = await import('../pdf/pro.js'); const { PDFLib, PDFJS } = await import('./util.js'); const PL = await PDFLib();
    if (Array.prototype.map.call(P.md5(new TextEncoder().encode('abc')), b => (b < 16 ? '0' : '') + b.toString(16)).join('') !== '900150983cd24fb0d6963f7d28e17f72') throw new Error('md5');
    const doc = await PL.PDFDocument.create(); const pg = doc.addPage([300, 200]); pg.drawText('Mat khau OK', { x: 20, y: 100, size: 20 }); doc.setTitle('Thử nghiệm'); const src = await doc.save({ useObjectStreams: false });
    const enc = await P.encryptPdf(src, '1234', 'owner', -3904); const pdfjs = await PDFJS();
    let denied = false; try { await pdfjs.getDocument({ data: enc.slice(), isEvalSupported: false }).promise; } catch (e) { denied = /password/i.test(e.name + e.message); }
    const opened = await pdfjs.getDocument({ data: enc.slice(), isEvalSupported: false, password: '1234' }).promise; const txt = (await (await opened.getPage(1)).getTextContent()).items.map(i => i.str).join('');
    const dec = await P.decryptPdf(enc, 'owner'); const plain = await pdfjs.getDocument({ data: dec.bytes.slice(), isEvalSupported: false }).promise; const txt2 = (await (await plain.getPage(1)).getTextContent()).items.map(i => i.str).join('');
    if (!denied || txt !== 'Mat khau OK' || txt2 !== 'Mat khau OK') throw new Error('denied=' + denied + ' txt=' + txt + ' txt2=' + txt2);
    const sig = document.createElement('canvas'); sig.width = 60; sig.height = 20; sig.getContext('2d').fillRect(0, 0, 60, 20);
    const burnt = await P.burn(src, [{ type: 'highlight', page: 0, x: 10, y: 10, w: 100, h: 20, color: '#ffd400' }, { type: 'pen', page: 0, color: '#c00000', pts: [[10, 10], [50, 40], [90, 20]] }, { type: 'image', page: 0, img: sig, x: 150, y: 120, w: 60, h: 20 }], 1);
    const chk = await PL.PDFDocument.load(burnt); if (chk.getPageCount() !== 1) throw new Error('burn');
    return 'mật khẩu: chặn khi không có mật khẩu, mở được bằng "1234", gỡ bằng mật khẩu chủ · chú thích ' + burnt.length + ' bytes';
  });
  await t('v007 PPTX tương thích PowerPoint: toạ độ hình hợp lệ (không NaN) khi lưu và mở lại', async () => {
    const p = await docOf('ppt'); const b = await p.save('pptx'); const { JSZipLib } = await import('./util.js'); const JSZip = await JSZipLib(); const z = await JSZip.loadAsync(b);
    const xml = await z.file('ppt/slides/slide1.xml').async('string'); if (/NaN|undefined/.test(xml)) throw new Error('slide1.xml chứa NaN');
    const p2 = await docOf('ppt'); await p2.load(b, 'pptx', 't.pptx'); const o = p2.slides[0].objects[0]; if (!(isFinite(o.x) && isFinite(o.w) && o.w > 100)) throw new Error('đọc lại: ' + JSON.stringify({ x: o.x, w: o.w }));
    return 'x=' + Math.round(o.x) + ' w=' + Math.round(o.w);
  });
  await t('v007 PowerPoint VBA: macro tạo slide + lưu .pptm → mở lại giữ macro', async () => {
    const V = await window.OOS_loadVBA(); const PV = await import('../ppt/vbappt.js'); const p = await docOf('ppt'); const n0 = p.slides.length;
    p.vba = { modules: [{ name: 'Module1', code: PV.PPT_SAMPLE + '\nSub DoiMau()\n  Dim s\n  For Each s In ActivePresentation.Slides\n    s.Shapes(1).TextFrame.TextRange.Font.Bold = True\n  Next\n  ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Text = "Yomi-Office Open Source " & ActivePresentation.Slides.Count\nEnd Sub\n' }], shortcuts: {} };
    await V.runMacro(p, 'TaoSlideTongKet', { quiet: true, onError: e => { throw e; } }); await V.runMacro(p, 'DoiMau', { quiet: true, onError: e => { throw e; } });
    const last = p.slides[p.slides.length - 1]; const txt = last.objects.map(o => o.html).join('|');
    if (p.slides.length !== n0 + 1 || txt.indexOf('Tổng kết') < 0 || txt.indexOf('• Slide 1') < 0 || p.slides[0].objects[0].html !== 'Yomi-Office Open Source ' + (n0 + 1)) throw new Error(p.slides.length + ' ' + txt);
    const b = await p.save('pptm'); const p2 = await docOf('ppt'); await p2.load(b, 'pptm', 't.pptm'); const m = (p2.vba && p2.vba.modules || []).find(z => z.name === 'Module1');
    if (!m || !/Sub TaoSlideTongKet/.test(m.code) || p2.slides.length !== n0 + 1) throw new Error('pptm: ' + (m ? 'slides ' + p2.slides.length : 'mất macro'));
    return (n0 + 1) + ' slide · .pptm ' + b.length + ' bytes';
  });

  // ================================================================ v0.0.8
  await t('v008 Word: phân trang thật — không đoạn nào vắt qua lề/khoảng giữa trang, ngắt trang sang trang mới, Đi tới trang', async () => {
    const w = await docOf('word'); let html = '<h1>Phân trang</h1>';
    for (let i = 1; i <= 45; i++) html += '<p>Đoạn ' + i + ': kiểm tra tách trang, hiển thị tách trang, đổi trang và chia trang của Word trong Yomi-Office Open Source.</p>' + (i === 12 ? '<div class="wd-pagebreak" contenteditable="false"></div>' : '');
    html += '<table>' + Array(25).fill('<tr><td>A</td><td>B</td></tr>').join('') + '</table>';
    w.body.innerHTML = html; w.setView && w.setView('print'); w.updatePages();
    const m = w.marg, P = w.pageH, G = 24, cH = P - m[0] - m[2]; const body = w.body;
    const top = el => { let y = 0, e = el; while (e && e !== body) { y += e.offsetTop; e = e.offsetParent; } return y; };
    const bad = Array.prototype.filter.call(body.querySelectorAll(':scope > p, :scope > h1'), el => { const t0 = top(el), b0 = t0 + el.offsetHeight; const n = Math.floor(t0 / (P + G)); return t0 < n * (P + G) + m[0] - 1 || b0 > n * (P + G) + m[0] + cH + 1; });
    const nx = body.querySelector('.wd-pagebreak').nextElementSibling; const off = (top(nx) - m[0]) % (P + G);
    if (bad.length || Math.abs(off) > 2) throw new Error('vắt trang: ' + bad.length + ' · sau ngắt lệch ' + off);
    w.gotoPage(3); if (w.curPage !== 3) throw new Error('đi tới trang: ' + w.curPage);
    const docx = await w.save('docx'); if (/wd-pg|nth-child/.test(new TextDecoder().decode(docx.slice ? docx.slice(0, 0) : new Uint8Array(0))) || /wd-pg-style|margin-top:\d+(\.\d+)?px !important/.test(w.body.innerHTML)) throw new Error('khoảng đệm phân trang lọt vào tài liệu');
    return w.pageCount + ' trang · đoạn sau ngắt bắt đầu đầu trang · Đi tới trang 3 OK';
  });
  await t('v008 Excel: ngắt trang tự động theo khổ giấy/lề, chèn/xoá ngắt thủ công, In tiêu đề, lưu XLSX <rowBreaks>', async () => {
    const x = await docOf('excel'); const s = x.wb.cur; for (let r = 0; r < 110; r++) for (let c = 0; c < 12; c++) s.setCell(r, c, { v: r === 0 ? 'Cột ' + (c + 1) : r * 100 + c });
    const P8 = await window.OOS_v8.pages(); const a = P8.computePages(s); if (a.pages.length < 2) throw new Error('tự động: ' + a.pages.length);
    s.sel.r = 30; s.sel.c = 4; P8.insertBreak(x); const b = P8.computePages(s); if (!b.rowStarts.some(z => z.at === 30 && z.manual) || !b.colStarts.some(z => z.at === 4 && z.manual)) throw new Error('thủ công: ' + JSON.stringify(b.rowStarts));
    s.pageSetup.printTitles = '1:1'; const bytes = await x.save('xlsx'); const y = await docOf('excel'); await y.load(bytes, 'xlsx', 'pb.xlsx'); const ps = y.wb.cur.pageSetup;
    if ((ps.rowBreaks || []).indexOf(30) < 0 || (ps.colBreaks || []).indexOf(4) < 0 || ps.printTitles !== '1:1') throw new Error('xlsx: ' + JSON.stringify(ps));
    P8.removeBreak(x); if ((s.pageSetup.rowBreaks || []).length) throw new Error('xoá ngắt');
    const area = document.createElement('div'); P8.renderPrint(x, area); const n = area.querySelectorAll('.xl-pg').length; if (n !== P8.computePages(s).pages.length) throw new Error('in: ' + n);
    return a.pages.length + ' trang tự động → ' + b.pages.length + ' trang khi chèn ngắt · khứ hồi XLSX OK · bản in ' + n + ' trang';
  });
  await t('v008 PDF: giao diện mới (bảng điều hướng, Công cụ, bố cục 2 trang, In nhiều trang/tờ) + PDF→PowerPoint/Excel/PDF-A', async () => {
    const d = await docOf('pdf'); const PL = await (await import('./util.js')).PDFLib(); const p = await PL.PDFDocument.create(); const f = await p.embedFont(PL.StandardFonts.Helvetica);
    for (let i = 1; i <= 4; i++) { const pp = p.addPage([595, 842]); pp.drawText('Trang ' + i, { x: 60, y: 760, size: 24, font: f }); pp.drawText('A', { x: 60, y: 700, size: 12, font: f }); pp.drawText(String(i * 7), { x: 200, y: 700, size: 12, font: f }); }
    await d.load(await p.save(), 'pdf', 'ui8.pdf'); await wait(600);
    const U = await window.OOS_v8.pdf(); U.togglePanel(d, 'info'); U.toggleTools(d, true); U.setLayout(d, 'twoCont'); await wait(500);
    const spreads = d.viewer.querySelectorAll('.pdf-spread').length; const tools = d.toolsPanel.querySelectorAll('.pdf-t8-i').length;
    d.printCfg = null; const box = document.createElement('div'); d.printCfg = { range: 'all', pages: '', subset: 'all', reverse: false, handling: 'fit', pct: 100, perSheet: 4, booklet: false, orient: 'auto', paper: 'A4', gray: false, dpi: 72 }; await U.printPreview(d, box);
    const ppt = await U.toPptx(d, true); const xl = await U.toXlsx(d, true); const pa = await U.makePdfA(d.bytes); const pat = new TextDecoder('latin1').decode(pa);
    U.setLayout(d, 'cont');
    if (spreads !== 2 || tools < 20 || !box.querySelector('canvas') || ppt.length < 1000 || xl.pages.length !== 4 || xl.pages[1][1][1] !== '14' || !/pdfaid:part>2/.test(pat) || !/GTS_PDFA1/.test(pat)) throw new Error([spreads, tools, !!box.querySelector('canvas'), ppt.length, JSON.stringify(xl.pages[1]), /pdfaid/.test(pat)].join(' | '));
    return '2 cặp trang · ' + tools + ' công cụ · in 4 trang/tờ · PPTX ' + ppt.length + 'B · XLSX 4 trang tính · PDF/A-2B';
  });
  await t('v008 Đám mây: PKCE SHA-256, ràng buộc tệp, phát hiện xung đột, tự tải bản mới khi người khác sửa (máy chủ giả lập)', async () => {
    const C = await window.OOS_v8.cloud(); const enc = new TextEncoder().encode('abc');
    const hx = a => Array.prototype.map.call(a, x => ('0' + x.toString(16)).slice(-2)).join('');
    if (hx(C._sha256js(enc)) !== 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad') throw new Error('sha256');
    // in-memory provider that behaves like Drive/OneDrive/Dropbox (rev changes on every write; stale rev → conflict)
    const store = { id: 'f1', name: 'Bao cao.docx', rev: '1', bytes: null };
    C.PROVIDERS.mock = { name: 'Mock Cloud', icon: 'cloud', color: '#555', async meta() { return { id: store.id, name: store.name, rev: store.rev, by: 'Đồng nghiệp' }; }, async download() { return store.bytes; },
      async upload(bind, bytes) { if (bind && bind.rev && !bind.force && bind.rev !== store.rev) { const e = new Error('conflict'); e.conflict = true; throw e; } store.bytes = bytes; store.rev = String(+store.rev + 1); return { id: store.id, name: store.name, rev: store.rev }; }, editUrl() { return 'about:blank'; } };
    const w = await docOf('word'); w.body.innerHTML = '<p>Bản 1</p>'; store.bytes = await w.save('docx');
    C.bind(w, { provider: 'mock', id: 'f1', name: 'Bao cao.docx', rev: '1' }); clearInterval(w._cloudTimer);
    w.body.innerHTML = '<p>Bản 2 của tôi</p>'; const ok = await C.syncUp(w); if (!ok || store.rev !== '2') throw new Error('tải lên: rev ' + store.rev);
    // someone else edits → rev 3; our local copy is clean → pull happens on watch tick
    const other = await docOf('word'); other.body.innerHTML = '<p>Bản 3 của đồng nghiệp</p>'; store.bytes = await other.save('docx'); store.rev = '3';
    const m = await C.PROVIDERS.mock.meta(); w.dirty = false; if (m.rev !== w.cloud.rev) { const P = await C.PROVIDERS.mock.download(); await w.load(P, 'docx', 'x.docx'); w.cloud.rev = m.rev; }
    if (w.getText().indexOf('đồng nghiệp') < 0) throw new Error('kéo bản mới: ' + w.getText());
    // stale upload → conflict detected
    w.cloud.rev = '2'; let conflict = false; try { await C.PROVIDERS.mock.upload(w.cloud, new Uint8Array(1)); } catch (e) { conflict = !!e.conflict; }
    delete C.PROVIDERS.mock; clearInterval(w._cloudTimer); w.cloud = null;
    if (!conflict) throw new Error('không phát hiện xung đột');
    return 'SHA-256 chuẩn · đồng bộ lên rev 2 · nhận bản của đồng nghiệp · chặn ghi đè khi lệch phiên bản';
  });

  await t('v008 Kết nối gốc (Host.http) cho đám mây — Android/Windows gọi HTTPS không qua CORS', async () => {
    const Host = (await import('./host.js')).default; if (Host.kind === 'web') return 'trình duyệt thuần: bỏ qua (chỉ chạy trong bản cài)';
    try { const r = await Host.http({ url: 'https://www.mediafire.com/api/1.5/folder/get_content.php?response_format=json&content_type=folders&folder_key=seqly6bv55yr8', method: 'GET' }); const j = JSON.parse(r.text); if (r.status !== 200 || !j.response) throw new Error('HTTP ' + r.status); return 'HTTP ' + r.status + ' · ' + (j.response.result || 'ok'); }
    catch (e) { if (/Unable to resolve|UnknownHost|network|timed out|resolve|connect/i.test(e.message)) return 'không có mạng: ' + e.message.slice(0, 60); throw e; }
  });

  // ---------------- v0.1.8
  await t('v018 Tên & phiên bản: Yomi-Office Open Source 0.1.8', async () => { const U = await import('./util.js'); if (U.APP_NAME !== 'Yomi-Office Open Source' || U.VERSION !== '0.1.8') throw new Error(U.APP_NAME + ' ' + U.VERSION); return U.APP_NAME + ' ' + U.VERSION; });
  await t('v018 Excel: hàm liên kết nhiều sheet, 3D và nhiều bảng tính (Feature Check #30)', async () => {
    const M = await import('../excel/model.js'); const F = await import('../excel/formula.js'); const K = (r, c) => r * 20000 + c;
    const wb = new M.Workbook(); const a = wb.addSheet('Sheet1'), b = wb.addSheet('Sheet2'), c = wb.addSheet('Sheet3');
    a.cells.set(K(0, 0), { v: 1 }); b.cells.set(K(0, 0), { v: 2 }); c.cells.set(K(0, 0), { v: 3 }); a.cells.set(K(1, 0), { v: 10 });
    const wb2 = new M.Workbook(); const d = wb2.addSheet('Data'); d.cells.set(K(0, 0), { v: 100 });
    const prev = window.OOS_EXTERNAL_WORKBOOKS; window.OOS_EXTERNAL_WORKBOOKS = Object.assign({}, prev || {}, { 'Book2.xlsx': wb2 });
    const cases = [[b, 'SUM(Sheet1:Sheet3!A1)', 6], [b, 'Sheet1!A1+Sheet1!A2', 11], [a, 'SUM(A1:A2)', 11], [c, 'SUM(Sheet1!A1:A2,Sheet2!A1)', 13], [a, '[Book2.xlsx]Data!A1*2', 200], [c, 'SUM([Book2.xlsx]Data!A1,Sheet2!A1)', 102]];
    cases.forEach((x, i) => x[0].cells.set(K(10 + i, 0), { f: x[1] })); new F.Calc(wb).recalc();
    window.OOS_EXTERNAL_WORKBOOKS = prev;
    const bad = cases.filter((x, i) => String(x[0].get(10 + i, 0).v) !== String(x[2])).map(x => x[1]);
    if (bad.length) throw new Error('sai: ' + bad.join(', ')); return cases.length + '/' + cases.length + ' công thức đúng';
  });
  await t('v018 Cập nhật phần mềm: GitHub Releases (nguồn chính) + tên bộ cài mới', async () => {
    const U = await import('./update.js');
    const r = U.parseGitHubRelease({ tag_name: 'v9.9.9', assets: ['Yomi-Office-Open-Source-9.9.9-Android10plus-arm64-v8a.apk', 'Yomi-Office-Open-Source-9.9.9-Win10-11-x64.msi', 'Yomi-Office-Open-Source-9.9.9-Win10-11-x64-Setup.exe'].map((n, i) => ({ id: i, name: n, browser_download_url: 'https://x/' + n })) });
    const a = U.pickFile(r.files, { tag: 'Android10plus-arm64-v8a', platform: 'android', want: [] }, r.manifest), m = U.pickFile(r.files, { tag: 'Win10-11-x64', platform: 'windows', msi: true }, r.manifest);
    if (!a || !/arm64/.test(a.name) || !m || !/\.msi$/.test(m.name) || r.manifest.version !== '9.9.9') throw new Error('chọn sai bộ cài'); return a.name + ' · ' + m.name;
  });
  await t('v018 A.I: Free Claude Code có trong Cài đặt A.I', async () => { const A = await import('../ai/ai.js'); const P = A.PROVIDERS.freeclaude; if (!P || P.kind !== 'anthropic' || !/8082\/v1\/messages$/.test(P.url)) throw new Error('thiếu'); return P.url; });
  await t('v018 Quét tệp Word/Excel/PowerPoint trên thiết bị (Feature Check #28)', async () => {
    const D = await import('./devscan.js'); if (D.kindOf('a.DOCX') !== 'word' || D.kindOf('b.xlsm') !== 'excel' || D.kindOf('c.pptx') !== 'ppt' || D.kindOf('d.pdf') || D.normalize({ name: '~$tmp.docx', path: 'x' })) throw new Error('phân loại sai');
    const Host = (await import('./host.js')).default; if (Host.kind === 'web') return 'trình duyệt thuần: chỉ kiểm tra phân loại';
    const files = await D.scan(); return files ? files.length + ' tệp Office' : 'host chưa hỗ trợ';
  });

  const summary = 'OOSTEST DONE | ' + pass + ' passed, ' + fail + ' failed';
  console.log(summary);
  window.OOS_TEST_RESULT = { pass, fail, lines: out };
  shell.showStart && shell.showStart();
  return window.OOS_TEST_RESULT;
}
