// Yomi-Office Open Source v0.0.6 — macro library: ready-made VBA / JSA / LAMBDA modules gathered from the VBA-related
// projects of the "GitHub Suite" sheet (all MIT licensed; attribution kept in each module header):
//   • aminya/AcuteMSWord            — Word VBA (AminMacros.bas)
//   • TLE-11/nsfc-md2docx           — Word VBA (Md2DocxTools.bas: field update / LaTeX helpers)
//   • okazaki112/Jsa-vibcoding-skill — WPS JSA templates (adapted to run offline against our Excel object model)
//   • MatrixFun/AutoXL              — Excel LAMBDA function library (imported as workbook names A.*)
// linhut/document-ai-assistant (Word VBA that calls an online AI HTTP API) is listed but not bundled: it cannot work offline.
import { h, ic, esc } from '../../core/util.js';
import { dialog, toast, alertDlg } from '../../core/ui.js';
import AMIN from './lib/AminMacros.bas';
import MD2DOCX from './lib/Md2DocxTools.bas';
import AUTOXL from './lib/AutoXL.txt';

const stripAttr = s => String(s).replace(/\r\n/g, '\n').replace(/^Attribute VB_.*\n/gm, '');

const XL_TOOLS = `' ExcelTools — tiện ích VBA mẫu cho Excel (Yomi-Office Open Source, chạy ngoại tuyến)
Option Explicit

' Tô đỏ các số âm trong vùng đang chọn
Sub ToDoSoAm()
    Dim c As Range
    For Each c In Selection
        If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
            If c.Value < 0 Then c.Font.Color = RGB(192, 0, 0) Else c.Font.Color = RGB(0, 0, 0)
        End If
    Next c
End Sub

' Xóa khoảng trắng thừa trong vùng chọn (giống hàm TRIM)
Sub XoaKhoangTrang()
    Dim c As Range
    For Each c In Selection
        If VarType(c.Value) = vbString And Not c.HasFormula Then c.Value = Application.WorksheetFunction.Trim(c.Value)
    Next c
End Sub

' Tạo trang "MucLuc" liệt kê mọi trang tính
Sub TaoMucLuc()
    Dim ws As Worksheet, toc As Worksheet, i As Long
    On Error Resume Next
    Set toc = Worksheets("MucLuc")
    On Error GoTo 0
    If toc Is Nothing Then
        Set toc = Worksheets.Add(Before:=Worksheets(1))
        toc.Name = "MucLuc"
    End If
    toc.Cells.Clear
    toc.Range("A1").Value = "Mục lục"
    toc.Range("A1").Font.Bold = True
    i = 2
    For Each ws In Worksheets
        If ws.Name <> toc.Name Then
            toc.Cells(i, 1).Value = ws.Name
            toc.Cells(i, 2).Value = ws.UsedRange.Address
            i = i + 1
        End If
    Next ws
    toc.Columns("A:B").AutoFit
End Sub

' Tổng hợp nhanh cột A: số dòng, tổng, trung bình
Sub ThongKeCotA()
    Dim last As Long
    last = Cells(Rows.Count, 1).End(xlUp).Row
    If last < 1 Then Exit Sub
    MsgBox "Số dòng: " & last & vbCrLf & _
           "Tổng: " & Application.WorksheetFunction.Sum(Range("A1:A" & last)) & vbCrLf & _
           "Trung bình: " & Format(Application.WorksheetFunction.Average(Range("A1:A" & last)), "0.00"), vbInformation, "Thống kê cột A"
End Sub

' Xóa các dòng trùng lặp theo cột A (giữ lần xuất hiện đầu tiên)
Sub XoaDongTrung()
    Dim d As Object, r As Long, last As Long, k As String, del As Collection, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    Set del = New Collection
    last = Cells(Rows.Count, 1).End(xlUp).Row
    For r = 1 To last
        k = CStr(Cells(r, 1).Value)
        If d.Exists(k) Then del.Add r Else d.Add k, r
    Next r
    For i = del.Count To 1 Step -1
        Rows(del(i)).Delete
    Next i
    MsgBox "Đã xóa " & del.Count & " dòng trùng", vbInformation
End Sub
`;

// JSA templates adapted from okazaki112/Jsa-vibcoding-skill (MIT, (c) 2026 okazaki112) — code_templates.md
const JSA_TPL = `// JsaTemplates — mẫu JSA (JavaScript for Applications, kiểu WPS) phỏng theo okazaki112/Jsa-vibcoding-skill (MIT)
// Hàm không tham số (ví dụ DemoJsa) xuất hiện trong danh sách Macro (Alt+F8).

function GetLastRow(ws, col) {           // dòng cuối có dữ liệu trong cột
    ws = ws || ActiveSheet
    var r = ws.Cells(ws.Rows.Count, col || 1).End(xlUp).Row
    return (r === 1 && ws.Cells(1, col || 1).Value2 == null) ? 0 : r
}
function GetLastCol(ws, row) {           // cột cuối có dữ liệu trong hàng
    ws = ws || ActiveSheet
    return ws.Cells(row || 1, ws.Columns.Count).End(xlToLeft).Column
}
function GetColumnLetter(n) {            // 28 -> "AB"
    if (n <= 0 || n > 16384) throw new Error('1 <= cột <= 16384')
    var s = ''
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) }
    return s
}
function SheetExists(name) {
    for (var i = 1; i <= ThisWorkbook.Worksheets.Count; i++) if (ThisWorkbook.Worksheets(i).Name === name) return true
    return false
}
function IsRangeEmpty(rng) { return Application.WorksheetFunction.CountA(rng) === 0 }

// Đọc A2:D(n) vào mảng một lần, xử lý trong bộ nhớ, ghi trả một lần: Thành tiền = Số lượng × Đơn giá
function ProcessWithArray(ws) {
    ws = ws || ActiveSheet
    var last = GetLastRow(ws, 1)
    if (last < 2) return
    var data = ws.Range("A2:D" + last).Value2
    for (var i = 0; i < data.length; i++) data[i][3] = data[i][1] * data[i][2]
    ws.Range("A2:D" + last).Value2 = data
}
// Viết hoa mọi chuỗi trong A1:D10
function Process2DArray(ws) {
    ws = ws || ActiveSheet
    var data = ws.Range("A1:D10").Value2
    for (var r = 0; r < data.length; r++)
        for (var c = 0; c < data[r].length; c++)
            if (typeof data[r][c] === 'string') data[r][c] = data[r][c].toUpperCase()
    ws.Range("A1:D10").Value2 = data
}
// Tạo (hoặc làm trống) một trang tính theo tên
function CreateWorksheet(sheetName) {
    var ws
    if (SheetExists(sheetName)) ws = ThisWorkbook.Worksheets(sheetName)
    else { ws = ThisWorkbook.Worksheets.Add(); ws.Name = sheetName }
    ws.Cells.Clear()
    return ws
}
// Tìm mọi ô trong cột có giá trị bằng searchValue, trả về mảng địa chỉ
function FindInColumn(ws, col, searchValue) {
    ws = ws || ActiveSheet
    var out = [], last = GetLastRow(ws, col)
    for (var r = 1; r <= last; r++) if (ws.Cells(r, col).Value2 == searchValue) out.push(ws.Cells(r, col).Address())
    return out
}
// Định dạng tiêu đề + kẻ khung bảng dữ liệu bắt đầu ở A1
function FormatTable(ws) {
    ws = ws || ActiveSheet
    var lastR = GetLastRow(ws, 1), lastC = GetLastCol(ws, 1)
    if (!lastR) return
    var head = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastC))
    head.Font.Bold = true
    head.Interior.Color = RGB(221, 235, 247)
    ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Columns.AutoFit()
}

// Ví dụ chạy được ngay: tạo trang "JSA_Demo" với dữ liệu, tính thành tiền và định dạng
function DemoJsa() {
    var ws = CreateWorksheet("JSA_Demo")
    ws.Range("A1:D1").Value2 = [["Mặt hàng", "Số lượng", "Đơn giá", "Thành tiền"]]
    ws.Range("A2:C4").Value2 = [["Bút", 10, 5000], ["Vở", 4, 12000], ["Thước", 2, 8000]]
    ProcessWithArray(ws)
    FormatTable(ws)
    ws.Activate()
    console.log("Cột cuối: " + GetColumnLetter(GetLastCol(ws, 1)) + ", dòng cuối: " + GetLastRow(ws, 1))
}
`;

const WORD_TOOLS = `' WordTools — tiện ích VBA mẫu cho Word (Yomi-Office Open Source, chạy ngoại tuyến)
Option Explicit

' Thay hai dấu cách liên tiếp bằng một dấu cách trong toàn văn bản
Sub XoaKhoangTrangKep()
    Dim lan As Integer
    Do While InStr(ActiveDocument.Content.Text, "  ") > 0 And lan < 20
        lan = lan + 1
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Text = "  "
            .Replacement.Text = " "
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Loop
End Sub

' Đếm từ / đoạn trong văn bản
Sub ThongKeVanBan()
    MsgBox "Số đoạn: " & ActiveDocument.Paragraphs.Count & vbCrLf & _
           "Số từ: " & ActiveDocument.Words.Count & vbCrLf & _
           "Số ký tự: " & ActiveDocument.Characters.Count, vbInformation, "Thống kê"
End Sub

' Định dạng mọi đoạn: Times New Roman 13, giãn dòng 1.5, căn đều (theo Nghị định 30/2020)
Sub DinhDangVanBanHanhChinh()
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        p.Range.Font.Name = "Times New Roman"
        p.Range.Font.Size = 13
        p.LineSpacingRule = wdLineSpace1pt5
        p.Alignment = wdAlignParagraphJustify
    Next p
End Sub

' Viết hoa chữ đầu mỗi đoạn
Sub VietHoaDauDoan()
    Dim p As Paragraph, t As String
    For Each p In ActiveDocument.Paragraphs
        t = p.Range.Text
        If Len(t) > 1 Then p.Range.Characters(1).Text = UCase(Left(t, 1))
    Next p
End Sub
`;

// ---------------------------------------------------------------- AutoXL (LAMBDA library) → workbook names
// Strips // comments (outside strings), splits top-level `;`, turns `[param]` into `param` (optional parameters work via ISOMITTED)
export function parseLambdaLibrary(src) {
  let txt = '', i = 0, inStr = false;
  src = String(src).replace(/\r\n/g, '\n');
  while (i < src.length) {
    const ch = src[i];
    if (inStr) { txt += ch; if (ch === '"') { if (src[i + 1] === '"') { txt += '"'; i++; } else inStr = false; } i++; continue; }
    if (ch === '"') { inStr = true; txt += ch; i++; continue; }
    if (ch === '/' && src[i + 1] === '/') { while (i < src.length && src[i] !== '\n') i++; continue; }
    txt += ch; i++;
  }
  const stmts = []; let depth = 0, cur = ''; inStr = false;
  for (const ch of txt) {
    if (inStr) { cur += ch; if (ch === '"') inStr = false; continue; }
    if (ch === '"') inStr = true;
    else if (ch === '(' || ch === '{') depth++;
    else if (ch === ')' || ch === '}') depth--;
    else if (ch === ';' && depth === 0) { stmts.push(cur); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim()) stmts.push(cur);
  const out = [];
  stmts.forEach(s => {
    const m = /^\s*([A-Za-z_\\][\w.]*)\s*=\s*([\s\S]+?)\s*$/.exec(s);
    if (!m) return;
    const f = m[2].replace(/\[([A-Za-z_][\w.]*)\]/g, '$1').replace(/\s*\n\s*/g, ' ');
    out.push({ name: m[1], f });
  });
  return out;
}
export function importLambdaLibrary(wb, src) {
  const defs = parseLambdaLibrary(src);
  defs.forEach(d => { wb.names[d.name.toUpperCase()] = { f: d.f, display: d.name, comment: 'AutoXL (MIT)' }; });
  return defs.length;
}

// ---------------------------------------------------------------- catalog
export function catalog(kind) {
  if (kind === 'word') return [
    { id: 'wordtools', name: 'WordTools', title: 'Tiện ích Word (mẫu Yomi-Office Open Source)', src: 'Yomi-Office Open Source', desc: 'Xóa khoảng trắng kép, thống kê, định dạng văn bản hành chính, viết hoa đầu đoạn.', code: () => WORD_TOOLS },
    { id: 'amin', name: 'AminMacros', title: 'AminMacros — AcuteMSWord', src: 'aminya/AcuteMSWord (MIT)', desc: 'Dán mã nguồn có định dạng (codePaste, pasteSelected), MathType2Word. Lệnh gọi MathType/Application.Run bị bỏ qua khi chạy ngoại tuyến.', code: () => "' Nguồn: https://github.com/aminya/AcuteMSWord — MIT License (c) Amin Yahyaabadi\n" + stripAttr(AMIN) },
    { id: 'md2docx', name: 'Md2DocxTools', title: 'Md2DocxTools — nsfc-md2docx', src: 'TLE-11/nsfc-md2docx (MIT)', desc: 'Cập nhật trường, đếm/chọn công thức LaTeX, dọn dẹp sau chuyển đổi. Chuyển LaTeX→MathType/AxMath cần add-in Windows nên chỉ chạy được trên MS Word.', code: () => "' Nguồn: https://github.com/TLE-11/nsfc-md2docx — MIT License\n" + stripAttr(MD2DOCX) },
    { id: 'docai', name: null, title: 'document-ai-assistant (chỉ tham khảo)', src: 'linhut/document-ai-assistant', desc: 'Plugin VBA gọi API AI trực tuyến qua HTTP — không chạy được ngoại tuyến nên không tích hợp. Dùng "Trợ lý AI" ngoại tuyến có sẵn của ứng dụng thay thế.', disabled: true }
  ];
  return [
    { id: 'xltools', name: 'ExcelTools', title: 'Tiện ích Excel VBA (mẫu Yomi-Office Open Source)', src: 'Yomi-Office Open Source', desc: 'Tô đỏ số âm, xóa khoảng trắng, tạo mục lục, thống kê cột, xóa dòng trùng.', code: () => XL_TOOLS },
    { id: 'jsa', name: 'JsaTemplates', type: 'jsa', title: 'Mẫu JSA (WPS) — Jsa-vibcoding-skill', src: 'okazaki112/Jsa-vibcoding-skill (MIT)', desc: 'Hàm JSA: GetLastRow, GetLastCol, GetColumnLetter, SheetExists, ProcessWithArray, CreateWorksheet, FindInColumn, FormatTable, DemoJsa.', code: () => JSA_TPL },
    { id: 'autoxl', name: null, lambda: true, title: 'AutoXL — thư viện hàm LAMBDA', src: 'MatrixFun/AutoXL (MIT)', desc: '29 hàm mảng A.* (A.EQ, A.XLOOKUP.ROWS, A.UNION.ROWS, A.DUPLICATES, …) được thêm vào Trình quản lý Tên; dùng trực tiếp trong công thức, ví dụ =A.VERSION().' }
  ];
}

export function libraryDialog(doc, onAdd) {
  const P = (doc.wb || doc).vba;
  const list = h('div.vba-lib');
  catalog(doc.kind).forEach(it => {
    const btn = h('button.btn' + (it.disabled ? '' : '.primary'), it.disabled ? 'Không khả dụng ngoại tuyến' : it.lambda ? 'Thêm vào Tên' : 'Thêm mô-đun');
    if (it.disabled) btn.disabled = true;
    btn.addEventListener('click', () => {
      if (it.lambda) {
        const n = doc.change ? (() => { let k = 0; doc.change('Nhập thư viện LAMBDA', doc.wb.cur, 'book', () => { k = importLambdaLibrary(doc.wb, AUTOXL); }); return k; })() : importLambdaLibrary(doc.wb, AUTOXL);
        doc.calc && doc.calc.invalidate && doc.calc.invalidate();
        doc.markDirty && doc.markDirty();
        toast('Đã thêm ' + n + ' hàm LAMBDA của AutoXL (A.*) vào Trình quản lý Tên', 'ok');
        return;
      }
      let name = it.name, i = 2; while (P.modules.some(m => m.name === name)) name = it.name + (i++);
      const m = { name, code: it.code() }; if (it.type) m.type = it.type;
      P.modules.push(m); doc.markDirty && doc.markDirty();
      toast('Đã thêm mô-đun ' + name, 'ok'); onAdd && onAdd(m);
    });
    list.appendChild(h('div.vba-lib-it' + (it.disabled ? '.off' : ''), h('div.vba-lib-t', ic(it.type === 'jsa' ? 'braces' : it.lambda ? 'function-square' : 'file-code', 16), ' ', h('b', it.title)), h('div.small', it.desc), h('div.muted.small', 'Nguồn: ' + it.src), btn));
  });
  return dialog({ title: 'Thư viện macro (GitHub Suite)', width: 640, body: h('div', list, h('div.hint', 'Các mô-đun được tích hợp từ những dự án VBA trong bảng GitHub Suite (giấy phép MIT). Tất cả chạy ngoại tuyến bằng bộ thông dịch VBA/JSA của Yomi-Office Open Source.')), buttons: [{ label: 'Đóng', primary: true }] });
}
export const AUTOXL_SRC = AUTOXL;
