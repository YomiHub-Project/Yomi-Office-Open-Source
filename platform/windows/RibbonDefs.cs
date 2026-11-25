// Office Open Source v0.1.6 - Ribbon content for the native Word / Excel / PowerPoint editors (tabs per the reference screenshots)
// Every command maps to a real LibreOfficeKit .uno: command or to an implemented feature of this app.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OfficeOpenSource
{
    partial class NativeDocForm
    {
        static RItem I(string text, string icon, string tip, Action a) { return Ribbon.I(text, icon, tip, a); }
        RItem St(string text, string icon, string tip, string uno) { RItem i = I(text, icon, tip, U(uno)); i.State = uno; return i; }
        static RItem M(string text, string icon, string tip, params RMenu[] menu) { RItem i = I(text, icon, tip, null); i.Menu = new List<RMenu>(menu); return i; }
        static RItem Sp(string text, string icon, string tip, Action a, params RMenu[] menu) { RItem i = I(text, icon, tip, a); i.Menu = new List<RMenu>(menu); i.Split = true; return i; }
        static RMenu Mi(string t, string ic, Action a) { return new RMenu(t, ic, a); }
        bool Hidden(string tab) { return NativeSettings.Get("ribbon.hide", "").Contains("|" + tab + "|"); }
        public static string[] TabNames(int t)
        {
            if (t == LokDoc.SPREADSHEET) return new[] { "Trang đầu", "Chèn", "Vẽ", "Bố trí Trang", "Công thức", "Dữ liệu", "Xem lại", "Xem", "Trợ giúp" };
            if (t == LokDoc.PRESENTATION) return new[] { "Trang đầu", "Chèn", "Vẽ", "Thiết kế", "Chuyển tiếp", "Hoạt hình", "Trình chiếu", "Xem lại", "Xem", "Trợ giúp" };
            return new[] { "Trang đầu", "Chèn", "Vẽ", "Thiết kế", "Bố cục", "Tham chiếu", "Thư", "Xem lại", "Xem", "Trợ giúp" };
        }

        void BuildRibbon()
        {
            ribbon.Tabs.Clear(); ribbon.Tab("Tệp");
            if (type == LokDoc.TEXT) { ribbon.EditMode = "Đang chỉnh sửa"; ribbon.EditModeMenu = new List<RMenu> { Mi("Chỉnh sửa", "editor", () => { ribbon.EditMode = "Đang chỉnh sửa"; ribbon.LayoutAll(); ribbon.Invalidate(); UnoView(".uno:TrackChanges"); UnoView(".uno:TrackChanges"); }), Mi("Xem lại (theo dõi thay đổi)", "track", () => { ribbon.EditMode = "Đang xem lại"; ribbon.LayoutAll(); ribbon.Invalidate(); Uno(".uno:TrackChanges"); }) }; }
            foreach (string t in TabNames(type))
            {
                if (Hidden(t) && t != "Trang đầu") continue;
                RTab tab = ribbon.Tab(t);
                switch (t)
                {
                    case "Trang đầu": TabHome(tab); break;
                    case "Chèn": TabInsert(tab); break;
                    case "Vẽ": TabDraw(tab); break;
                    case "Thiết kế": if (type == LokDoc.PRESENTATION) TabDesignSlides(tab); else TabDesignWord(tab); break;
                    case "Bố cục": case "Bố trí Trang": TabLayout(tab); break;
                    case "Tham chiếu": TabReferences(tab); break;
                    case "Công thức": TabFormulas(tab); break;
                    case "Dữ liệu": TabData(tab); break;
                    case "Chuyển tiếp": TabTransitions(tab); break;
                    case "Hoạt hình": TabAnimations(tab); break;
                    case "Thư": TabMailings(tab); break;
                    case "Trình chiếu": TabSlideShow(tab); break;
                    case "Xem lại": TabReview(tab); break;
                    case "Xem": TabView(tab); break;
                    case "Trợ giúp": TabHelp(tab); break;
                }
            }
            if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OOS_DUMP_UI")) || !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OOS_DUMP_UNIFIED"))) ApplyUnifiedLayout();   // v0.1.7: same layout as the Web editor
            ribbon.Selected = 1;
            if (NativeSettings.Bool("ribbon.min", false)) ribbon.SetMinimized(true);
            ribbon.LayoutAll(); ribbon.Invalidate();
        }

        // ---------------- shared groups ----------------
        void GroupClipboard(RTab tab)
        {
            RGroup g = tab.Add("Bảng tạm"); g.Icon = "paste";
            g.Large(Sp("Dán", "paste", "Dán (Ctrl+V)", () => { canvas.PasteClipboard(); MarkDirty(); },
                Mi("Dán", "paste", () => { canvas.PasteClipboard(); MarkDirty(); }),
                Mi("Chỉ giữ văn bản", "textbox", () => { try { if (Clipboard.ContainsText()) { if (!doc.Paste("text/plain;charset=utf-8", Encoding.UTF8.GetBytes(Clipboard.GetText()))) canvas.TypeText(Clipboard.GetText()); MarkDirty(); } } catch { } })));
            g.Stack(I("Cắt", "cut", "Cắt (Ctrl+X)", () => { canvas.CopySelection(true); MarkDirty(); }),
                    I("Sao chép", "copy", "Sao chép (Ctrl+C)", () => canvas.CopySelection(false)),
                    St("Chổi định dạng", "painter", "Sao chép định dạng từ vùng chọn rồi áp dụng cho nội dung khác", ".uno:FormatPaintbrush"));
        }

        RItem FontCombo()
        {
            RItem f = new RItem { Kind = RK.Combo, Width = 150, Tip = "Phông chữ", State = ".uno:CharFontName" };
            f.Combo = new ComboBox { Font = Ui.F(9f), FlatStyle = FlatStyle.Flat, DropDownHeight = 420 };
            foreach (string fn in new[] { "Times New Roman", "Arial", "Calibri", "Cambria", "Tahoma", "Segoe UI", "Verdana", "Courier New", "Georgia", "Roboto", "Liberation Serif", "Liberation Sans", "DejaVu Sans" }) f.Combo.Items.Add(fn);
            Action apply = () => { if (f.Combo.Text.Trim().Length > 0) Uno(".uno:CharFontName", Args(JS("CharFontName.FamilyName", f.Combo.Text.Trim()))); };
            f.Combo.SelectionChangeCommitted += (s, e) => { f.Combo.Text = (string)f.Combo.SelectedItem; apply(); };
            f.Combo.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; apply(); } };
            f.Combo.DropDown += (s, e) => LoadFontList(f.Combo);
            ribbon.Controls.Add(f.Combo); return f;
        }
        bool fontsLoaded;
        void LoadFontList(ComboBox c)
        {
            if (fontsLoaded) return; fontsLoaded = true;
            try
            {
                string js = doc.CommandValues(".uno:CharFontName"); if (string.IsNullOrEmpty(js)) return;
                Dictionary<string, object> o = Json.Deserialize<Dictionary<string, object>>(js);
                Dictionary<string, object> vals = o["commandValues"] as Dictionary<string, object>; if (vals == null) return;
                string cur = c.Text; List<string> names = new List<string>(vals.Keys); names.Sort(StringComparer.CurrentCultureIgnoreCase);
                c.BeginUpdate(); c.Items.Clear(); foreach (string n in names) c.Items.Add(n); c.EndUpdate(); c.Text = cur;
            }
            catch { }
        }
        RItem SizeCombo()
        {
            RItem f = new RItem { Kind = RK.Combo, Width = 52, Tip = "Cỡ phông", State = ".uno:FontHeight" };
            f.Combo = new ComboBox { Font = Ui.F(9f), FlatStyle = FlatStyle.Flat };
            foreach (string n in new[] { "8", "9", "10", "10.5", "11", "12", "13", "14", "16", "18", "20", "24", "28", "36", "48", "72" }) f.Combo.Items.Add(n);
            Action apply = () => { float v; if (float.TryParse(f.Combo.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out v)) Uno(".uno:FontHeight", "{\"FontHeight.Height\":{\"type\":\"float\",\"value\":\"" + v.ToString(CultureInfo.InvariantCulture) + "\"}}"); };
            f.Combo.SelectionChangeCommitted += (s, e) => { f.Combo.Text = (string)f.Combo.SelectedItem; apply(); };
            f.Combo.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; apply(); } };
            ribbon.Controls.Add(f.Combo); return f;
        }
        RItem FontColorItem()
        {
            RItem c = I(null, "fontcolor", "Màu phông chữ", null); c.Kind = RK.Mini; c.Swatch = lastFontColor;
            c.Click = () => Uno(".uno:Color", Args(JL("Color", Rgb(lastFontColor))));
            c.PickColor = col => { lastFontColor = col; Uno(".uno:Color", Args(JL("Color", col == Color.Transparent ? -1 : Rgb(col)))); };
            return c;
        }
        RItem HighlightItem()
        {
            RItem c = I(null, "highlight", "Màu tô sáng văn bản", null); c.Kind = RK.Mini; c.Swatch = lastHighlight; c.Palette = ColorPopup.Highlight;
            c.Click = () => Uno(".uno:CharBackColor", Args(JL("CharBackColor", Rgb(lastHighlight))));
            c.PickColor = col => { lastHighlight = col; Uno(".uno:CharBackColor", Args(JL("CharBackColor", col == Color.Transparent ? -1 : Rgb(col)))); };
            return c;
        }
        RItem FillItem(string cmd, string arg, string tip)
        {
            RItem c = I(null, "pagecolor", tip, null); c.Kind = RK.Mini; c.Swatch = lastFill;
            c.Click = () => Uno(cmd, Args(JL(arg, Rgb(lastFill))));
            c.PickColor = col => { lastFill = col; Uno(cmd, Args(JL(arg, col == Color.Transparent ? -1 : Rgb(col)))); };
            return c;
        }
        RItem Mini(string icon, string tip, Action a) { RItem i = I(null, icon, tip, a); i.Kind = RK.Mini; return i; }
        RItem MiniSt(string icon, string tip, string uno) { RItem i = St(null, icon, tip, uno); i.Kind = RK.Mini; return i; }
        RItem CaseMenu()
        {
            RItem i = M(null, "case", "Đổi chữ hoa/thường",
                Mi("Viết hoa đầu câu.", null, U(".uno:ChangeCaseToSentenceCase")), Mi("chữ thường", null, U(".uno:ChangeCaseToLower")),
                Mi("CHỮ HOA", null, U(".uno:ChangeCaseToUpper")), Mi("Viết Hoa Mỗi Từ", null, U(".uno:ChangeCaseToTitleCase")), Mi("đẢO cHỮ hOA/tHƯỜNG", null, U(".uno:ChangeCaseToToggleCase")));
            i.Kind = RK.Mini; return i;
        }

        void GroupFont(RTab tab, bool cellFill)
        {
            RGroup g = tab.Add("Phông chữ"); g.Icon = "fontcolor"; g.Launcher = FontDialogShow; g.LauncherTip = "Phông chữ… (Ctrl+D)";
            List<RItem> r1 = new List<RItem> { FontCombo(), SizeCombo(), Mini("grow", "Tăng cỡ phông (Ctrl+])", U(".uno:Grow")), Mini("shrink", "Giảm cỡ phông (Ctrl+[)", U(".uno:Shrink")) };
            if (!cellFill) r1.Add(CaseMenu());
            r1.Add(Mini("clearfmt", "Xóa mọi định dạng", U(".uno:ResetAttributes")));
            List<RItem> r2 = new List<RItem> { MiniSt("bold", "Đậm (Ctrl+B)", ".uno:Bold"), MiniSt("italic", "Nghiêng (Ctrl+I)", ".uno:Italic") };
            RItem u = MiniSt("underline", "Gạch dưới (Ctrl+U)", ".uno:Underline"); u.Menu = new List<RMenu> { Mi("Gạch dưới đơn", "underline", U(".uno:Underline")), Mi("Gạch dưới kép", "underline", U(".uno:UnderlineDouble")) }; u.Split = true; r2.Add(u);
            r2.Add(MiniSt("strike", "Gạch ngang chữ", ".uno:Strikeout"));
            if (!cellFill) { r2.Add(MiniSt("sub", "Chỉ số dưới (Ctrl+=)", ".uno:SubScript")); r2.Add(MiniSt("sup", "Chỉ số trên (Ctrl+Shift++)", ".uno:SuperScript")); }
            if (type == LokDoc.PRESENTATION) r2.Add(MiniSt("wordart", "Đổ bóng văn bản", ".uno:Shadowed"));
            if (cellFill) r2.Add(FillItem(".uno:BackgroundColor", "BackgroundColor", "Màu tô ô"));
            else r2.Add(HighlightItem());
            r2.Add(FontColorItem());
            g.Rows(r1.ToArray(), r2.ToArray());
        }

        RItem LineSpacingMenu()
        {
            RItem i = M(null, "linespacing", "Giãn cách dòng và đoạn",
                Mi("1,0", null, U(".uno:SpacePara1")), Mi("1,5", null, U(".uno:SpacePara15")), Mi("2,0", null, U(".uno:SpacePara2")), RMenu.Sep(),
                Mi("Tùy chọn giãn cách dòng…", "linespacing", ParagraphDialog), RMenu.Sep(),
                Mi("Tăng khoảng cách giữa các đoạn", null, U(".uno:ParaspaceIncrease")), Mi("Giảm khoảng cách giữa các đoạn", null, U(".uno:ParaspaceDecrease")));
            i.Kind = RK.Mini; return i;
        }
        void GroupParagraph(RTab tab)
        {
            RGroup g = tab.Add("Đoạn văn"); g.Icon = "alignleft"; g.Launcher = ParagraphDialog; g.LauncherTip = "Cài đặt đoạn văn";
            List<RItem> r1 = new List<RItem> { MiniSt("bullets", "Dấu đầu dòng", ".uno:DefaultBullet"), MiniSt("numbering", "Đánh số", ".uno:DefaultNumbering"),
                Mini("indentdec", "Giảm mức thụt lề", U(".uno:DecrementIndent")), Mini("indentinc", "Tăng mức thụt lề", U(".uno:IncrementIndent")) };
            if (type == LokDoc.TEXT) r1.Add(MiniSt("pilcrow", "Hiện/Ẩn dấu đoạn văn và ký hiệu định dạng ẩn (Ctrl+*)", ".uno:ControlCodes"));
            List<RItem> r2 = new List<RItem> { MiniSt("alignleft", "Căn trái (Ctrl+L)", ".uno:LeftPara"), MiniSt("aligncenter", "Căn giữa (Ctrl+E)", ".uno:CenterPara"),
                MiniSt("alignright", "Căn phải (Ctrl+R)", ".uno:RightPara"), MiniSt("justify", "Căn đều (Ctrl+J)", ".uno:JustifyPara"), LineSpacingMenu() };
            g.Rows(r1.ToArray(), r2.ToArray());
        }

        void GroupEditing(RTab tab)
        {
            RGroup g = tab.Add("Chỉnh sửa"); g.Icon = "find";
            g.Stack(Sp("Tìm", "find", "Tìm văn bản (Ctrl+F)", () => FindReplace(false), Mi("Tìm", "find", () => FindReplace(false)), Mi("Đi tới…", "select", GoToDialog)),
                    I("Thay thế", "replace", "Thay thế văn bản (Ctrl+H)", () => FindReplace(true)),
                    M("Chọn", "select", "Chọn văn bản hoặc đối tượng", Mi("Chọn tất cả", "select", () => UnoView(".uno:SelectAll")), Mi("Chọn đối tượng", "select", () => UnoView(".uno:SelectObject"))));
        }

        void GroupVoiceEtc(RTab tab)
        {
            RGroup v = tab.Add("Giọng nói"); v.Large(I("Đọc chính tả", "dictate", "Nhập văn bản bằng giọng nói (Windows Voice Typing, Win+H)", Dictate));
            RGroup e = tab.Add("Trình soạn thảo"); RItem sp = I("Trình soạn thảo", "editor", "Bật/tắt tự động kiểm tra chính tả", U(".uno:SpellOnline")); sp.State = ".uno:SpellOnline"; e.Large(sp);
            RGroup ai = tab.Add("Trợ lý A.I"); ai.Large(M("Trợ lý A.I", "ai", "Hỏi A.I về nội dung đang chọn (ChatGPT, Gemini, Copilot, Claude…)", AiMenu().ToArray()));
            RGroup a = tab.Add("Yomi-Office Open Source"); a.Large(I("Giao diện Web", "addins", "Mở tệp này bằng trình soạn thảo Web đầy đủ của Yomi-Office Open Source (AI, đám mây, mẫu…)", OpenInWeb));
        }

        void GroupMacros(RTab tab)
        {
            RGroup m = tab.Add("Macro"); m.Icon = "macros";
            m.Large(Sp("Macro", "macros", "Chạy macro VBA/Basic trong tài liệu", RunMacro, Mi("Chạy macro…", "macros", RunMacro), Mi(Lok.MacrosAllowed ? "Tắt macro" : "Bật macro…", "protect", ToggleMacros)));
        }

        void GroupShapes(RGroup g, bool large)
        {
            RItem shapes = M("Hình dạng", "shapes", "Chèn hình có sẵn: chọn rồi kéo trên tài liệu",
                Mi("Hình chữ nhật", null, U(".uno:BasicShapes.rectangle")), Mi("Hình chữ nhật bo góc", null, U(".uno:BasicShapes.round-rectangle")),
                Mi("Hình elip", null, U(".uno:BasicShapes.ellipse")), Mi("Tam giác", null, U(".uno:BasicShapes.isosceles-triangle")),
                Mi("Mũi tên phải", null, U(".uno:ArrowShapes.right-arrow")), Mi("Ngôi sao 5 cánh", null, U(".uno:StarShapes.star5")),
                Mi("Chú thích hình chữ nhật", null, U(".uno:CalloutShapes.round-rectangular-callout")), RMenu.Sep(),
                Mi("Đường thẳng", null, U(".uno:Line")), Mi("Mũi tên", null, U(".uno:LineArrowEnd")));
            if (large) g.Large(shapes); else g.Stack(shapes);
        }

        // ---------------- Home ----------------
        void TabHome(RTab tab)
        {
            GroupClipboard(tab);
            if (type == LokDoc.SPREADSHEET)
            {
                GroupFont(tab, true);
                RGroup al = tab.Add("Căn chỉnh"); al.Icon = "alignleft";
                RItem wrap = St(null, "wrap", "Ngắt dòng văn bản", ".uno:WrapText"); wrap.Kind = RK.Mini;
                RItem merge = Sp(null, "merge", "Gộp và căn giữa", U(".uno:ToggleMergeCells"), Mi("Gộp & căn giữa", "merge", () => { Uno(".uno:ToggleMergeCells"); Uno(".uno:AlignHorizontalCenter"); }), Mi("Gộp ô / bỏ gộp ô", "merge", U(".uno:ToggleMergeCells"))); merge.Kind = RK.Mini; merge.State = ".uno:ToggleMergeCells";
                al.Rows(new[] { MiniSt("aligntop", "Căn trên", ".uno:AlignTop"), MiniSt("alignmiddle", "Căn giữa theo chiều dọc", ".uno:AlignVCenter"), MiniSt("alignbottom", "Căn dưới", ".uno:AlignBottom"), wrap },
                        new[] { MiniSt("alignleft", "Căn trái", ".uno:AlignLeft"), MiniSt("aligncenter", "Căn giữa", ".uno:AlignHorizontalCenter"), MiniSt("alignright", "Căn phải", ".uno:AlignRight"), Mini("indentdec", "Giảm thụt lề", U(".uno:DecrementIndent")), Mini("indentinc", "Tăng thụt lề", U(".uno:IncrementIndent")), merge });
                RGroup num = tab.Add("Số"); num.Icon = "currency";
                RItem fmt = new RItem { Kind = RK.Combo, Width = 128, Tip = "Định dạng số" }; fmt.Combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = Ui.F(9f), FlatStyle = FlatStyle.Flat };
                string[][] fm = { new[] { "Chung", ".uno:NumberFormatStandard" }, new[] { "Số", ".uno:NumberFormatDecimal" }, new[] { "Tiền tệ", ".uno:NumberFormatCurrency" }, new[] { "Ngày", ".uno:NumberFormatDate" }, new[] { "Giờ", ".uno:NumberFormatTime" }, new[] { "Phần trăm", ".uno:NumberFormatPercent" }, new[] { "Khoa học", ".uno:NumberFormatScientific" } };
                foreach (string[] f in fm) fmt.Combo.Items.Add(f[0]); fmt.Combo.SelectedIndex = 0;
                fmt.Combo.SelectionChangeCommitted += (s, e) => Uno(fm[fmt.Combo.SelectedIndex][1]); ribbon.Controls.Add(fmt.Combo);
                num.Rows(new[] { fmt }, new[] { Mini("currency", "Định dạng tiền tệ", U(".uno:NumberFormatCurrency")), Mini("percent", "Kiểu phần trăm (Ctrl+Shift+%)", U(".uno:NumberFormatPercent")), Mini("thousands", "Dấu phân cách hàng nghìn", U(".uno:NumberFormatThousands")), Mini("incdec", "Tăng số thập phân", U(".uno:NumberFormatIncDecimals")), Mini("decdec", "Giảm số thập phân", U(".uno:NumberFormatDecDecimals")) });
                RGroup sty = tab.Add("Kiểu"); sty.Icon = "cellstyles";
                List<RMenu> cs = new List<RMenu>();
                foreach (string[] c in new[] { new[] { "Tốt", "Good" }, new[] { "Trung tính", "Neutral" }, new[] { "Xấu", "Bad" }, new[] { "Cảnh báo", "Warning" }, new[] { "Lỗi", "Error" }, new[] { "Ghi chú", "Note" }, new[] { "Đầu đề", "Heading" }, new[] { "Đầu đề 1", "Heading 1" }, new[] { "Đầu đề 2", "Heading 2" }, new[] { "Nhấn mạnh", "Accent" }, new[] { "Nhấn mạnh 1", "Accent 1" }, new[] { "Nhấn mạnh 2", "Accent 2" }, new[] { "Nhấn mạnh 3", "Accent 3" }, new[] { "Kết quả", "Result" }, new[] { "Mặc định", "Default" } })
                { string sn = c[1]; cs.Add(Mi(c[0], null, U(".uno:StyleApply", Args(JS("Style", sn), JS("FamilyName", "CellStyles"))))); }
                RItem cfm = I("Định dạng có điều kiện", "condfmt", "Tô sáng ô theo quy tắc, thanh dữ liệu, thang màu", null); cfm.Menu = CondMenu(); sty.Large(cfm);
                RItem cst = I("Kiểu ô", "cellstyles", "Áp dụng kiểu ô có sẵn", null); cst.Menu = cs; sty.Large(cst);
                RGroup cells = tab.Add("Ô"); cells.Icon = "insertcells";
                cells.Large(M("Chèn", "insertcells", "Chèn ô, hàng, cột hoặc trang tính", Mi("Chèn hàng trang tính", "insertrow", U(".uno:InsertRowsBefore")), Mi("Chèn cột trang tính", "insertcol", U(".uno:InsertColumnsBefore")), Mi("Chèn trang tính", "sheet", InsertSheet)));
                cells.Large(M("Xóa", "deletecells", "Xóa ô, hàng, cột hoặc trang tính", Mi("Xóa hàng trang tính", "deleterow", U(".uno:DeleteRows")), Mi("Xóa cột trang tính", "deletecol", U(".uno:DeleteColumns")), Mi("Xóa trang tính", "sheet", () => { if (doc.Parts > 1) { Uno(".uno:Remove"); RebuildParts(); } })));
                cells.Large(M("Định dạng", "formatcells", "Kích thước, ẩn/hiện, sắp xếp trang tính",
                    Mi("Tự động khớp độ rộng cột", "columns", U(".uno:SetOptimalColumnWidthDirect")), RMenu.Sep(),
                    Mi("Ẩn hàng", null, U(".uno:HideRow")), Mi("Bỏ ẩn hàng", null, U(".uno:ShowRow")), Mi("Ẩn cột", null, U(".uno:HideColumn")), Mi("Bỏ ẩn cột", null, U(".uno:ShowColumn")), RMenu.Sep(),
                    Mi("Đổi tên trang tính", "textbox", () => SheetTabMenuRename())));
                RGroup ed = tab.Add("Chỉnh sửa"); ed.Icon = "autosum";
                ed.Stack(Sp("Tự động tính tổng", "autosum", "Tổng (Alt+=)", U(".uno:AutoSum"), Mi("Tổng", "autosum", U(".uno:AutoSum")), Mi("Trung bình", null, () => InsertFunction("AVERAGE")), Mi("Đếm số", null, () => InsertFunction("COUNT")), Mi("Lớn nhất", null, () => InsertFunction("MAX")), Mi("Nhỏ nhất", null, () => InsertFunction("MIN")), RMenu.Sep(), Mi("Hàm khác…", "function", () => InsertFunction(null))),
                         M("Điền", "fill", "Điền tiếp dữ liệu", Mi("Xuống", null, U(".uno:FillDown")), Mi("Sang phải", null, U(".uno:FillRight")), Mi("Lên", null, U(".uno:FillUp")), Mi("Sang trái", null, U(".uno:FillLeft"))),
                         M("Xóa", "clear", "Xóa nội dung hoặc định dạng", Mi("Xóa nội dung", null, () => { doc.Key(0, 0, 1286); doc.Key(1, 0, 1286); MarkDirty(); }), Mi("Xóa định dạng", null, U(".uno:ResetAttributes"))));
                ed.Large(M("Sắp xếp & Lọc", "sortaz", "Sắp xếp và lọc dữ liệu", Mi("Sắp xếp từ A đến Z", "sortaz", U(".uno:SortAscending")), Mi("Sắp xếp từ Z đến A", "sortaz", U(".uno:SortDescending")), RMenu.Sep(), Mi("Bộ lọc", "filter", U(".uno:DataFilterAutoFilter"))));
                ed.Large(M("Tìm & Chọn", "find", "Tìm, thay thế, đi tới", Mi("Tìm…", "find", () => FindReplace(false)), Mi("Thay thế…", "replace", () => FindReplace(true)), Mi("Đi tới…", "select", () => { formula.NameBox.Focus(); formula.NameBox.SelectAll(); })));
                RGroup ai = tab.Add("Trợ lý A.I"); ai.Large(M("Trợ lý A.I", "ai", "Hỏi A.I về các ô đang chọn (ChatGPT, Gemini, Copilot, Claude…)", AiMenu().ToArray()));
                RGroup a = tab.Add("Yomi-Office Open Source"); a.Large(I("Giao diện Web", "addins", "Mở bằng trình soạn thảo Web đầy đủ (AI, đám mây, mẫu…)", OpenInWeb));
                return;
            }
            if (type == LokDoc.PRESENTATION)
            {
                RGroup sl = tab.Add("Trang chiếu"); sl.Icon = "newslide";
                sl.Large(Sp("Trang chiếu mới", "newslide", "Thêm trang chiếu (Ctrl+M)", NewSlide, Mi("Trang chiếu mới", "newslide", NewSlide), Mi("Nhân bản trang chiếu", "duplicate", () => { Uno(".uno:DuplicatePage"); RebuildParts(); })));
                sl.Stack(LayoutMenu(), I("Xóa", "deleteslide", "Xóa trang chiếu", DeleteSlide), St("Ẩn trang chiếu", "hideslide", "Ẩn/hiện trang chiếu khi trình chiếu", ".uno:HideSlide"));
            }
            GroupFont(tab, false);
            GroupParagraph(tab);
            if (type == LokDoc.TEXT) { GroupStyles(tab); GroupEditing(tab); GroupVoiceEtc(tab); return; }
            RGroup dr = tab.Add("Vẽ"); dr.Icon = "shapes";
            GroupShapes(dr, true);
            dr.Large(M("Sắp xếp", "arrange", "Thứ tự đối tượng", Mi("Đưa lên trước cùng", null, U(".uno:BringToFront")), Mi("Đưa lên một lớp", null, U(".uno:Forward")), Mi("Đưa ra sau một lớp", null, U(".uno:Backward")), Mi("Đưa ra sau cùng", null, U(".uno:SendToBack"))));
            RItem sf = I("Tô hình", "shapefill", "Màu tô của hình", null); sf.Kind = RK.Small; sf.PickColor = c => Uno(".uno:FillColor", Args(JL("FillColor", Rgb(c))));
            RItem so = I("Đường viền hình", "shapeoutline", "Màu đường viền của hình", null); so.Kind = RK.Small; so.PickColor = c => Uno(".uno:XLineColor", Args(JL("XLineColor", Rgb(c))));
            dr.Stack(sf, so);
            GroupEditing(tab);
            GroupVoiceEtc(tab);
        }

        RItem LayoutMenu()
        {
            Func<int, Action> L = n => () => Uno(".uno:AssignLayout", Args("\"WhatPage\":{\"type\":\"unsigned short\",\"value\":\"" + (doc.Part + 1) + "\"}", "\"WhatLayout\":{\"type\":\"unsigned short\",\"value\":\"" + n + "\"}"));
            return M("Bố trí", "layout", "Thay đổi bố trí trang chiếu", Mi("Trang chiếu tiêu đề", null, L(0)), Mi("Tiêu đề và nội dung", null, L(1)), Mi("Hai nội dung", null, L(3)), Mi("Chỉ tiêu đề", null, L(19)), Mi("Văn bản căn giữa", null, L(32)), Mi("Trống", null, L(20)));
        }

        // Word styles gallery (built from the document's paragraph styles)
        RItem styleGallery;
        void GroupStyles(RTab tab)
        {
            RGroup g = tab.Add("Kiểu"); g.Icon = "quickstyles"; g.Launcher = () => ribbon.ShowMenu(StyleMenu(), new Point(g.Bounds.X, g.Bounds.Bottom));
            styleGallery = new RItem { Kind = RK.Gallery, Width = 4 * 78 + 22, Gallery = new List<RGalleryItem>() };
            FillStyleGallery(new List<string> { "Standard", "No Spacing", "Heading 1", "Heading 2", "Title", "Subtitle", "Quotations", "List Paragraph" });
            g.Gallery(styleGallery);
        }
        static string StyleLabel(string s)
        {
            switch (s) { case "Standard": return "Bình thường"; case "No Spacing": return "Không giãn cách"; case "Heading 1": return "Đầu đề 1"; case "Heading 2": return "Đầu đề 2"; case "Heading 3": return "Đầu đề 3"; case "Title": return "Tiêu đề"; case "Subtitle": return "Phụ đề"; case "Quotations": return "Trích dẫn"; case "List Paragraph": return "Đoạn danh sách"; case "Text Body": case "Body Text": return "Thân văn bản"; }
            return s;
        }
        void FillStyleGallery(List<string> names)
        {
            styleGallery.Gallery.Clear();
            foreach (string n in names)
            {
                string sn = n; Font f = Ui.F(10f); Color c = Ui.Text; string pv = "AaBbCc";
                if (n.StartsWith("Heading 1")) { f = new Font("Calibri Light", 13f); c = Color.FromArgb(0x2F, 0x54, 0x96); pv = "AaBbCc"; }
                else if (n.StartsWith("Heading")) { f = new Font("Calibri Light", 11f); c = Color.FromArgb(0x2F, 0x54, 0x96); }
                else if (n == "Title") { f = new Font("Calibri Light", 16f); pv = "AaB"; }
                else if (n == "Subtitle") { f = new Font("Calibri", 10f, FontStyle.Italic); c = Color.FromArgb(0x59, 0x59, 0x59); }
                else if (n == "Quotations") { f = new Font("Calibri", 10f, FontStyle.Italic); c = Color.FromArgb(0x40, 0x40, 0x40); }
                styleGallery.Gallery.Add(new RGalleryItem { Name = StyleLabel(n), Font = f, Color = c, Preview = pv, Click = () => Uno(".uno:StyleApply", Args(JS("Style", sn), JS("FamilyName", "ParagraphStyles"))) });
            }
        }
        List<string> paraStyles = new List<string>();
        void LoadStyles()
        {
            if (styleGallery == null) return;
            try
            {
                string js = doc.CommandValues(".uno:StyleApply"); if (string.IsNullOrEmpty(js)) return;
                Dictionary<string, object> o = Json.Deserialize<Dictionary<string, object>>(js);
                Dictionary<string, object> v = o["commandValues"] as Dictionary<string, object>;
                System.Collections.ArrayList ps = v != null && v.ContainsKey("ParagraphStyles") ? v["ParagraphStyles"] as System.Collections.ArrayList : null; if (ps == null) return;
                paraStyles.Clear(); foreach (object x in ps) paraStyles.Add(Convert.ToString(x));
                List<string> pick = new List<string>();
                foreach (string want in new[] { "Standard", "No Spacing", "Heading 1", "Heading 2", "Heading 3", "Title", "Subtitle", "Quotations", "List Paragraph", "Text Body" }) if (paraStyles.Contains(want)) pick.Add(want);
                if (pick.Count >= 3) { FillStyleGallery(pick); ribbon.Invalidate(); }
            }
            catch { }
        }
        List<RMenu> StyleMenu()
        {
            List<RMenu> m = new List<RMenu>();
            List<string> src = paraStyles.Count > 0 ? paraStyles : new List<string> { "Standard", "Heading 1", "Heading 2", "Heading 3", "Title", "Subtitle", "Quotations" };
            foreach (string s in src) { string ss = s; m.Add(Mi(StyleLabel(s), null, U(".uno:StyleApply", Args(JS("Style", ss), JS("FamilyName", "ParagraphStyles"))))); if (m.Count > 30) break; }
            return m;
        }

        // ---------------- Insert ----------------
        void TabInsert(RTab tab)
        {
            if (type == LokDoc.TEXT)
            {
                RGroup pg = tab.Add("Trang"); pg.Icon = "blankpage";
                pg.Large(I("Trang trống", "blankpage", "Chèn một trang trống tại vị trí con trỏ", () => { Uno(".uno:InsertPagebreak"); Uno(".uno:InsertPagebreak"); doc.Key(0, 0, 1025); doc.Key(1, 0, 1025); }));
                pg.Large(I("Ngắt trang", "pagebreak", "Bắt đầu trang tiếp theo tại vị trí hiện tại (Ctrl+Enter)", U(".uno:InsertPagebreak")));
            }
            if (type == LokDoc.PRESENTATION) { RGroup sl = tab.Add("Trang chiếu"); sl.Large(Sp("Trang chiếu mới", "newslide", "Thêm trang chiếu", NewSlide, Mi("Trang chiếu mới", "newslide", NewSlide), Mi("Nhân bản trang chiếu", "duplicate", () => { Uno(".uno:DuplicatePage"); RebuildParts(); }))); }
            if (type == LokDoc.SPREADSHEET) { RGroup pv = tab.Add("Bảng"); pv.Large(I("PivotTable", "pivot", "Tóm tắt dữ liệu bằng PivotTable (trên trang tính mới)", InsertPivot)); }
            if (type != LokDoc.SPREADSHEET) { RGroup tb = tab.Add("Bảng"); RItem t = I("Bảng", "table", "Chèn bảng", null); t.Click = () => TablePicker.Show(ribbon, new Point(t.Bounds.X, t.Bounds.Bottom + 2), (r, c) => Uno(".uno:InsertTable", Args(JL("Columns", c), JL("Rows", r)))); tb.Large(t); }
            RGroup il = tab.Add("Hình minh họa"); il.Icon = "picture";
            il.Large(I("Ảnh", "picture", "Chèn ảnh từ máy tính", InsertImage));
            GroupShapes(il, true);
            il.Large(I("Biểu đồ", "chart", "Chèn biểu đồ (từ vùng dữ liệu đang chọn)", U(".uno:InsertObjectChart")));
            il.Large(I("Ảnh chụp màn hình", "screenshot", "Chụp màn hình và chèn vào tài liệu", InsertScreenshot));
            RGroup ln = tab.Add("Liên kết"); ln.Large(I("Liên kết", "link", "Chèn siêu liên kết (Ctrl+K)", InsertLink));
            if (type == LokDoc.TEXT) ln.Stack(I("Dấu trang", "bookmark", "Chèn dấu trang", () => { string n = Ask(this, "Dấu trang", "Tên dấu trang:", "DauTrang1"); if (!string.IsNullOrEmpty(n)) Uno(".uno:InsertBookmark", Args(JS("Bookmark", n.Replace(" ", "_")))); }));
            RGroup cm = tab.Add(type == LokDoc.SPREADSHEET ? "Ghi chú" : "Bình luận"); cm.Large(I(type == LokDoc.SPREADSHEET ? "Ghi chú" : "Bình luận", "comment", "Thêm ghi chú về phần này", InsertComment));
            if (type == LokDoc.TEXT)
            {
                RGroup hf = tab.Add("Đầu trang & Chân trang"); hf.Icon = "header";
                hf.Large(I("Đầu trang", "header", "Bật vùng đầu trang", U(".uno:InsertPageHeader", Args(JB("On", true)))));
                hf.Large(I("Chân trang", "footer", "Bật vùng chân trang", U(".uno:InsertPageFooter", Args(JB("On", true)))));
                hf.Large(M("Số trang", "pagenum", "Chèn số trang", Mi("Số trang hiện tại", "pagenum", U(".uno:InsertPageNumberField")), Mi("Tổng số trang", "pagenum", U(".uno:InsertPageCountField"))));
            }
            RGroup tx = tab.Add("Văn bản"); tx.Icon = "textbox";
            tx.Large(I("Hộp văn bản", "textbox", "Vẽ hộp văn bản: chọn rồi kéo trên tài liệu", U(type == LokDoc.PRESENTATION ? ".uno:Text" : ".uno:DrawText")));
            if (type == LokDoc.TEXT) tx.Stack(I("Ngày & Giờ", "datetime", "Chèn ngày hiện tại", U(".uno:InsertDateField")), I("Trường thời gian", "datetime", "Chèn giờ hiện tại", U(".uno:InsertTimeField")));
            if (type == LokDoc.PRESENTATION) tx.Stack(I("Số trang chiếu", "pagenum", "Chèn số trang chiếu", U(".uno:InsertPageField")), I("Ngày & Giờ", "datetime", "Chèn ngày", U(".uno:InsertDateFieldFix")));
            RGroup sy = tab.Add("Ký hiệu"); sy.Icon = "symbol";
            RItem sym = I("Ký hiệu", "symbol", "Chèn ký hiệu đặc biệt", null); sym.Click = () => SymbolPopup.Show(ribbon, new Point(sym.Bounds.X, sym.Bounds.Bottom + 2), s => { canvas.TypeText(s); MarkDirty(); }); sy.Large(sym);
            sy.Large(M("Phương trình", "equation", "Chèn phương trình toán học (LibreOffice Math, kiểu MathType)", EquationMenu().ToArray()));
            if (type == LokDoc.SPREADSHEET) sy.Large(I("Hàm", "function", "Chèn hàm", () => InsertFunction(null)));
        }

        // ---------------- Draw ----------------
        void TabDraw(RTab tab)
        {
            RGroup t = tab.Add("Công cụ"); t.Icon = "select";
            t.Large(I("Chọn", "select", "Chọn đối tượng", () => UnoView(".uno:SelectObject")));
            RGroup d = tab.Add("Công cụ Vẽ"); d.Icon = "editor";
            d.Large(I("Bút vẽ tự do", "editor", "Vẽ tự do: kéo chuột trên tài liệu", U(".uno:Freeline_Unfilled")));
            d.Large(I("Đường thẳng", "shapeoutline", "Vẽ đường thẳng", U(".uno:Line")));
            d.Large(I("Mũi tên", "export", "Vẽ mũi tên", U(".uno:LineArrowEnd")));
            d.Large(I("Hình chữ nhật", "arrange", "Vẽ hình chữ nhật", U(".uno:Rect")));
            d.Large(I("Hình elip", "shapes", "Vẽ hình elip", U(".uno:Ellipse")));
            RGroup c = tab.Add("Màu"); RItem lc = I("Màu nét", "shapeoutline", "Màu nét của đối tượng đang chọn", null); lc.PickColor = col => Uno(".uno:XLineColor", Args(JL("XLineColor", Rgb(col)))); c.Large(lc);
            RItem fc = I("Màu tô", "shapefill", "Màu tô của đối tượng đang chọn", null); fc.PickColor = col => Uno(".uno:FillColor", Args(JL("FillColor", Rgb(col)))); c.Large(fc);
            RGroup ar = tab.Add("Sắp xếp"); ar.Large(M("Thứ tự", "arrange", "Đưa đối tượng lên trước / ra sau", Mi("Đưa lên trước cùng", null, U(".uno:BringToFront")), Mi("Đưa ra sau cùng", null, U(".uno:SendToBack"))));
        }

        // ---------------- Design ----------------
        void TabDesignWord(RTab tab)
        {
            RGroup bg = tab.Add("Nền trang"); bg.Icon = "watermark";
            Func<string, Action> wm = s => U(".uno:Watermark", Args(JS("Text", s)));
            bg.Large(M("Hình mờ", "watermark", "Thêm chữ chìm phía sau nội dung trang", Mi("MẬT", null, wm("MẬT")), Mi("KHẨN", null, wm("KHẨN")), Mi("BẢN NHÁP", null, wm("BẢN NHÁP")), Mi("BẢN SAO", null, wm("BẢN SAO")), RMenu.Sep(),
                Mi("Hình mờ tùy chỉnh…", null, () => { string s = Ask(this, "Hình mờ", "Nội dung chữ chìm:", "MẪU"); if (!string.IsNullOrEmpty(s)) wm(s)(); }), Mi("Loại bỏ hình mờ", null, wm(""))));
            RItem pc = I("Màu trang", "pagecolor", "Màu nền trang", null); pc.PickColor = c => Uno(".uno:FillPageColor", Args(JL("FillPageColor", Rgb(c)))); bg.Large(pc);
            RGroup ps = tab.Add("Định dạng tài liệu"); ps.Icon = "quickstyles";
            ps.Large(M("Khoảng cách đoạn", "linespacing", "Đổi khoảng cách giữa các đoạn", Mi("Tăng khoảng cách", null, U(".uno:ParaspaceIncrease")), Mi("Giảm khoảng cách", null, U(".uno:ParaspaceDecrease"))));
        }
        void TabDesignSlides(RTab tab)
        {
            RGroup cu = tab.Add("Tùy chỉnh"); cu.Icon = "layout";
            Func<int, int, Action> sz = (w, h) => U(".uno:AttributePageSize", Args(JL("AttributePageSize.Width", w), JL("AttributePageSize.Height", h)));
            cu.Large(M("Kích thước trang chiếu", "layout", "Tỷ lệ khung hình", Mi("Màn hình rộng (16:9)", null, () => { sz(28000, 15750)(); AfterSlideSize(); }), Mi("Chuẩn (4:3)", null, () => { sz(28000, 21000)(); AfterSlideSize(); }), Mi("16:10", null, () => { sz(28000, 17500)(); AfterSlideSize(); })));
            RItem bgc = I("Định dạng nền", "pagecolor", "Màu nền trang chiếu", null); bgc.PickColor = c => Uno(".uno:FillPageColor", Args(JL("FillPageColor", Rgb(c)))); cu.Large(bgc);
        }
        void AfterSlideSize() { canvas.RefreshSize(); if (slides != null) { slides.SetAspect(doc.SizeTwips); slides.Reset(); } canvas.SetZoom(FitSlideZoom()); }

        // ---------------- Layout / Page Layout ----------------
        void TabLayout(RTab tab)
        {
            RGroup g = tab.Add("Thiết lập trang"); g.Icon = "margins"; g.Launcher = ShowPageSetup; g.LauncherTip = "Thiết lập trang…";
            Func<int, int, int, int, Action> mg = (t, b, l, r) => () => SetMargins(t, b, l, r);
            g.Large(M("Lề", "margins", "Chọn kích thước lề",
                Mi("Bình thường  (T/D/T/P 2,54 cm)", null, mg(254, 254, 254, 254)), Mi("Hẹp  (1,27 cm)", null, mg(127, 127, 127, 127)), Mi("Vừa  (2,54 / 1,91 cm)", null, mg(254, 254, 191, 191)),
                Mi("Rộng  (2,54 / 5,08 cm)", null, mg(254, 254, 508, 508)), Mi("Văn bản hành chính VN  (2 / 2 / 3 / 1,5 cm)", null, mg(200, 200, 300, 150)), RMenu.Sep(), Mi("Lề tùy chỉnh…", null, ShowPageSetup)));
            g.Large(M("Hướng", "orientation", "Hướng trang", Mi("Dọc", null, () => SetLandscape(false)), Mi("Ngang", null, () => SetLandscape(true))));
            Func<int, int, Action> ps = (w, h) => () => SetPaper(w, h);
            g.Large(M("Khổ giấy", "size", "Chọn khổ giấy", Mi("A4  (21 × 29,7 cm)", null, ps(2100, 2970)), Mi("A3  (29,7 × 42 cm)", null, ps(2970, 4200)), Mi("A5  (14,8 × 21 cm)", null, ps(1480, 2100)), Mi("Letter  (21,59 × 27,94 cm)", null, ps(2159, 2794)), Mi("Legal  (21,59 × 35,56 cm)", null, ps(2159, 3556)), Mi("B5  (17,6 × 25 cm)", null, ps(1760, 2500)), RMenu.Sep(), Mi("Khổ giấy khác…", null, ShowPageSetup)));
            if (type == LokDoc.TEXT)
            {
                g.Large(M("Cột", "columns", "Chia văn bản thành cột", Mi("Một", null, () => Uno(".uno:PageColumnType", Args(JL("PageColumnType", 1)))), Mi("Hai", null, () => Uno(".uno:PageColumnType", Args(JL("PageColumnType", 2)))), Mi("Ba", null, () => Uno(".uno:PageColumnType", Args(JL("PageColumnType", 3))))));
                g.Stack(M("Ngắt", "breaks", "Ngắt trang, ngắt cột", Mi("Ngắt trang", "pagebreak", U(".uno:InsertPagebreak")), Mi("Ngắt cột", "columns", U(".uno:InsertColumnBreak"))));
                RGroup p = tab.Add("Đoạn văn"); p.Icon = "indentinc"; p.Launcher = ParagraphDialog;
                p.Stack(I("Tăng thụt lề", "indentinc", null, U(".uno:IncrementIndent")), I("Giảm thụt lề", "indentdec", null, U(".uno:DecrementIndent")));
                p.Stack(I("Tăng khoảng cách", "linespacing", "Tăng khoảng cách trước/sau đoạn", U(".uno:ParaspaceIncrease")), I("Giảm khoảng cách", "linespacing", "Giảm khoảng cách trước/sau đoạn", U(".uno:ParaspaceDecrease")));
                RGroup ar = tab.Add("Sắp xếp"); ar.Icon = "arrange";
                ar.Large(M("Ngắt dòng văn bản", "wrap", "Cách văn bản bao quanh đối tượng đang chọn", Mi("Trên cùng dòng văn bản", null, U(".uno:WrapOff")), Mi("Vuông", null, U(".uno:WrapOn")), Mi("Tối ưu", null, U(".uno:WrapIdeal")), Mi("Sau văn bản", null, U(".uno:WrapThru")), Mi("Trái", null, U(".uno:WrapLeft")), Mi("Phải", null, U(".uno:WrapRight"))));
                ar.Stack(I("Đưa lên trước", "arrange", null, U(".uno:BringToFront")), I("Đưa ra sau", "arrange", null, U(".uno:SendToBack")));
            }
            else
            {
                g.Large(M("Vùng in", "printlayout", "Đặt / xóa vùng in", Mi("Đặt vùng in", null, U(".uno:DefinePrintArea")), Mi("Xóa vùng in", null, U(".uno:DeletePrintArea"))));
                g.Large(M("Ngắt trang", "pagebreak", "Chèn ngắt trang", Mi("Chèn ngắt hàng", null, U(".uno:InsertRowBreak")), Mi("Chèn ngắt cột", null, U(".uno:InsertColumnBreak")), Mi("Xóa ngắt hàng", null, U(".uno:DeleteRowbreak")), Mi("Xóa ngắt cột", null, U(".uno:DeleteColumnbreak"))));
                RGroup o = tab.Add("Tùy chọn trang tính"); o.Icon = "gridlines";
                RItem grid = I("Đường lưới", "gridlines", "Hiện/ẩn đường lưới", () => UnoView(".uno:ToggleSheetGrid")); o.Stack(grid, I("Đầu đề", "freeze", "Hiện/ẩn đầu đề hàng & cột", ToggleHeaders));
            }
        }
        void SetMargins(int t, int b, int l, int r)
        {   // values in 1/100 mm per 1/10 mm units given (cm*100) -> 1/100 mm = cm*1000
            Uno(".uno:PageULMargin", Args(JL("PageULMargin.Upper", t * 10), JL("PageULMargin.Lower", b * 10)));
            Uno(".uno:PageLRMargin", Args(JL("PageLRMargin.Left", l * 10), JL("PageLRMargin.Right", r * 10)));
            if (type == LokDoc.TEXT) SetVMargins(t * 10 * 1440 / 2540, b * 10 * 1440 / 2540);
            canvas.RefreshSize();
        }
        void SetLandscape(bool on)
        {
            Size s = PageTw(); bool isLand = s.Width > s.Height;
            if (isLand != on) Uno(".uno:AttributePage", Args(JB("AttributePage.Landscape", on)));
            canvas.RefreshSize(); SyncBars();
        }
        void SetPaper(int w, int h)
        {   // w/h in 1/10 mm -> 1/100 mm
            Size s = PageTw(); bool land = s.Width > s.Height;
            Uno(".uno:AttributePageSize", Args(JL("AttributePageSize.Width", (land ? h : w) * 10), JL("AttributePageSize.Height", (land ? w : h) * 10)));
            canvas.RefreshSize(); SyncBars();
        }
        void ToggleHeaders() { if (headers == null) return; bool show = canvas.Location.X == 0; canvas.Location = show ? new Point(SheetHeaders.RowW, SheetHeaders.ColH) : Point.Empty; canvas.Size = show ? new Size(headers.Width - SheetHeaders.RowW, headers.Height - SheetHeaders.ColH) : headers.Size; headers.Invalidate(); }

        // ---------------- References (Word) ----------------
        void TabReferences(RTab tab)
        {
            RGroup toc = tab.Add("Mục lục"); toc.Icon = "toc";
            toc.Large(M("Mục lục", "toc", "Chèn mục lục tự động từ các đầu đề (Heading 1–n)",
                Mi("Mục lục tự động (3 cấp)", "toc", () => InsertToc(3)), Mi("Mục lục tự động (2 cấp)", "toc", () => InsertToc(2)), Mi("Mục lục tự động (4 cấp)", "toc", () => InsertToc(4)), RMenu.Sep(),
                Mi("Cập nhật mục lục", "refresh", () => { string r = Tool("UpdateTOC"); if (ToolOk(r, "Mục lục")) MarkDirty(); }),
                Mi("Loại bỏ mục lục", "deletecells", () => { string r = Tool("RemoveTOC"); if (ToolOk(r, "Mục lục")) MarkDirty(); })));
            toc.Large(I("Cập nhật bảng", "refresh", "Cập nhật mục lục và các chỉ mục trong tài liệu", U(".uno:UpdateAllIndexes")));
            toc.Stack(I("Cập nhật trường", "refresh", "Cập nhật mọi trường (số trang, ngày…)", U(".uno:UpdateFields")));
            RGroup fn = tab.Add("Chú thích cuối trang"); fn.Icon = "footnote";
            fn.Large(I("Chèn chú thích cuối trang", "footnote", "Chèn chú thích cuối trang (Alt+Ctrl+F)", U(".uno:InsertFootnote")));
            fn.Stack(I("Chèn chú thích cuối tài liệu", "footnote", "Chèn chú thích cuối tài liệu", U(".uno:InsertEndnote")));
            RGroup cp = tab.Add("Chú thích"); cp.Icon = "picture";
            cp.Large(M("Chèn chú thích", "picture", "Chèn chú thích đánh số tự động (Hình 1, Bảng 1…) dưới đối tượng",
                Mi("Hình…", null, () => InsertCaption("Hình")), Mi("Bảng…", null, () => InsertCaption("Bảng")), Mi("Phương trình…", null, () => InsertCaption("Phương trình"))));
        }

        // ---------------- Formulas / Data (Excel) ----------------
        static readonly string[][] FuncCats = {
            new[] { "Tài chính", "PMT", "FV", "PV", "NPV", "IRR", "RATE", "NPER" },
            new[] { "Lô-gic", "IF", "IFS", "AND", "OR", "NOT", "IFERROR", "SWITCH" },
            new[] { "Văn bản", "LEFT", "RIGHT", "MID", "LEN", "TRIM", "CONCAT", "TEXT", "UPPER", "LOWER", "PROPER", "SUBSTITUTE", "FIND" },
            new[] { "Ngày & Giờ", "TODAY", "NOW", "DATE", "DAY", "MONTH", "YEAR", "WEEKDAY", "DATEDIF", "EOMONTH", "NETWORKDAYS" },
            new[] { "Tra cứu & Tham chiếu", "VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "CHOOSE", "OFFSET", "INDIRECT" },
            new[] { "Toán & Lượng giác", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "MOD", "ABS", "PRODUCT", "SQRT" },
            new[] { "Thống kê", "AVERAGE", "AVERAGEIF", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "COUNTBLANK", "MAX", "MIN", "MEDIAN", "RANK" } };
        List<RMenu> FunctionMenu(string cat)
        {
            List<RMenu> m = new List<RMenu>();
            foreach (string[] c in FuncCats)
            {
                if (cat != null && c[0] != cat) continue;
                if (cat == null) m.Add(new RMenu("— " + c[0] + " —", null, null));
                for (int i = 1; i < c.Length; i++) { string f = c[i]; m.Add(Mi(f, "function", () => InsertFunction(f))); }
            }
            return m;
        }
        void InsertFunction(string name)
        {
            if (formula == null) return;
            if (name == null) { ribbon.ShowMenu(FunctionMenu(null), ribbon.PointToClient(formula.PointToScreen(formula.FxPoint))); return; }
            formula.Formula.Focus(); formula.Formula.Text = "=" + name + "()"; formula.Formula.SelectionStart = name.Length + 2;
        }
        void SheetTabMenuRename() { string nm = Ask(this, "Đổi tên trang tính", "Tên mới:", doc.PartName(doc.Part)); if (!string.IsNullOrEmpty(nm)) { Uno(".uno:RenameTable", Args(JS("NewName", nm))); RebuildParts(); } }
        void TabFormulas(RTab tab)
        {
            RGroup lib = tab.Add("Thư viện hàm"); lib.Icon = "function";
            lib.Large(I("Chèn hàm", "function", "Chèn hàm vào ô (Shift+F3)", () => InsertFunction(null)));
            lib.Large(Sp("Tự động tính tổng", "autosum", "Tổng các ô phía trên/bên trái (Alt+=)", U(".uno:AutoSum"), Mi("Tổng", null, U(".uno:AutoSum")), Mi("Trung bình", null, () => InsertFunction("AVERAGE")), Mi("Đếm số", null, () => InsertFunction("COUNT")), Mi("Lớn nhất", null, () => InsertFunction("MAX")), Mi("Nhỏ nhất", null, () => InsertFunction("MIN"))));
            string[] icons = { "currency", "validation", "textbox", "datetime", "find", "calc", "chart" };
            for (int i = 0; i < FuncCats.Length; i++) { RItem c = I(FuncCats[i][0], icons[i], "Hàm " + FuncCats[i][0].ToLower(), null); c.Menu = FunctionMenu(FuncCats[i][0]); lib.Large(c); }
            RGroup au = tab.Add("Kiểm tra công thức"); au.Icon = "function";
            RItem sf = St("Hiện công thức", "function", "Hiện công thức thay cho kết quả (Ctrl+`)", ".uno:ToggleFormula"); au.Stack(sf);
            RGroup ca = tab.Add("Tính toán"); ca.Icon = "calc";
            RItem auto = St("Tính toán tự động", "calc", "Bật/tắt tự động tính lại công thức", ".uno:AutomaticCalculation"); ca.Large(auto);
            ca.Stack(I("Tính ngay", "calc", "Tính lại (F9)", U(".uno:Calculate")), I("Tính lại toàn bộ", "refresh", "Tính lại mọi công thức (Ctrl+Shift+F9)", U(".uno:CalculateHard")));
        }
        void TabData(RTab tab)
        {
            RGroup s = tab.Add("Sắp xếp & Lọc"); s.Icon = "filter";
            s.Stack(I("Sắp xếp A → Z", "sortaz", "Sắp xếp tăng dần", U(".uno:SortAscending")), I("Sắp xếp Z → A", "sortaz", "Sắp xếp giảm dần", U(".uno:SortDescending")));
            s.Large(St("Bộ lọc", "filter", "Bật lọc cho các ô đang chọn (Ctrl+Shift+L)", ".uno:DataFilterAutoFilter"));
            RGroup o = tab.Add("Dàn bài"); o.Icon = "section";
            o.Large(I("Nhóm", "section", "Nhóm các hàng/cột đang chọn (F12)", U(".uno:Group"))); o.Large(I("Rã nhóm", "section", "Bỏ nhóm", U(".uno:Ungroup")));
            o.Stack(I("Hiện chi tiết", null, null, U(".uno:ShowDetail")), I("Ẩn chi tiết", null, null, U(".uno:HideDetail")));
            RGroup t = tab.Add("Công cụ Dữ liệu"); t.Icon = "columns";
            t.Large(M("Văn bản thành cột", "columns", "Tách nội dung cột đầu của vùng chọn thành nhiều cột",
                Mi("Tách theo dấu phẩy ( , )", null, () => SplitColumns(",")), Mi("Tách theo dấu chấm phẩy ( ; )", null, () => SplitColumns(";")), Mi("Tách theo Tab", null, () => SplitColumns("TAB")), Mi("Tách theo dấu cách", null, () => SplitColumns("SPACE")), Mi("Ký tự khác…", null, () => { string d = Ask(this, "Văn bản thành cột", "Ký tự phân tách:", "-"); if (!string.IsNullOrEmpty(d)) SplitColumns(d); })));
            t.Large(I("Loại bỏ trùng lặp", "removedup", "Xóa các hàng trùng lặp trong vùng dữ liệu", RemoveDuplicates));
            t.Large(M("Xác thực dữ liệu", "validation", "Giới hạn dữ liệu được nhập vào ô",
                Mi("Danh sách thả xuống…", null, () => ValidationRule("list")), Mi("Số nguyên trong khoảng…", null, () => ValidationRule("whole")), Mi("Số thập phân trong khoảng…", null, () => ValidationRule("decimal")), RMenu.Sep(), Mi("Xóa xác thực", "deletecells", () => ValidationRule("clear"))));
            t.Stack(I("Làm mới PivotTable", "refresh", "Cập nhật mọi PivotTable theo dữ liệu nguồn", () => { string r = Tool("RefreshPivots"); if (ToolOk(r, "PivotTable")) MarkDirty(); }));
            t.Large(I("Độ rộng cột tối ưu", "columns", "Tự động khớp độ rộng các cột đang chọn", U(".uno:SetOptimalColumnWidthDirect")));
        }

        // ---------------- Transitions / Slide show (PowerPoint) ----------------
        string transition = "Không"; int transitionMs = 600;
        void TabTransitions(RTab tab)
        {
            transition = NativeSettings.Get("ppt.transition", "Mờ dần");
            RGroup g = tab.Add("Chuyển tiếp sang trang chiếu này"); g.Icon = "transition";
            RItem gal = new RItem { Kind = RK.Gallery, Width = 5 * 78 + 22, Gallery = new List<RGalleryItem>() };
            foreach (string n in new[] { "Không", "Mờ dần", "Đẩy", "Lau", "Thu phóng", "Tách" })
            { string nn = n; gal.Gallery.Add(new RGalleryItem { Name = n, Font = Ui.F(15f), Color = accent, Preview = n == "Không" ? "▢" : n == "Mờ dần" ? "◐" : n == "Đẩy" ? "⇨" : n == "Lau" ? "▤" : n == "Thu phóng" ? "⊕" : "⇔", Click = () => { transition = nn; NativeSettings.Set("ppt.transition", nn); ApplyTransition(false); status.Segs[0].Text = "Chuyển tiếp: " + nn; status.Invalidate(); } }); }
            g.Gallery(gal);
            RGroup tm = tab.Add("Thời gian"); tm.Icon = "datetime";
            tm.Large(M("Thời lượng", "datetime", "Thời lượng hiệu ứng chuyển tiếp", Mi("Nhanh (0,4 giây)", null, () => { transitionMs = 400; ApplyTransition(false); }), Mi("Vừa (0,7 giây)", null, () => { transitionMs = 700; ApplyTransition(false); }), Mi("Chậm (1,2 giây)", null, () => { transitionMs = 1200; ApplyTransition(false); })));
            tm.Stack(I("Áp dụng cho tất cả", "transition", "Dùng hiệu ứng chuyển tiếp này cho mọi trang chiếu", () => ApplyTransition(true)));
        }
        // v0.1.7: the transition is stored in the presentation itself (PowerPoint / LibreOffice play it too), not only in the app's settings
        void ApplyTransition(bool all)
        {
            string k = transition == "Mờ dần" ? "fade" : transition == "Đẩy" ? "push" : transition == "Lau" ? "wipe" : transition == "Thu phóng" ? "zoom" : transition == "Tách" ? "split" : "none";
            string r = Tool("SlideTransition", "type", k, "dur", (transitionMs / 1000.0).ToString(CultureInfo.InvariantCulture), "all", all ? "1" : "0");
            if (ToolOk(r, "Chuyển tiếp")) { MarkDirty(); status.Segs[0].Text = "Chuyển tiếp: " + transition + (all ? " (mọi trang chiếu)" : ""); status.Invalidate(); }
        }

        // ---------------- Animations (PowerPoint) - v0.1.7 ----------------
        string animSpeed = "medium";
        static readonly string[][] AnimList = {
            new[] { "Xuất hiện", "appear", "▣" }, new[] { "Hòa tan", "dissolve", "░" }, new[] { "Bay vào (trái)", "flyleft", "⇨" }, new[] { "Bay vào (dưới)", "flybottom", "⇧" },
            new[] { "Bay vào (phải)", "flyright", "⇦" }, new[] { "Bay vào (trên)", "flytop", "⇩" }, new[] { "Lau", "wipe", "▤" }, new[] { "Thu phóng", "zoom", "⊕" },
            new[] { "Xoay", "spin", "↻" }, new[] { "Sọc dọc", "stripes", "▥" }, new[] { "Bàn cờ", "checker", "▦" }, new[] { "Mở từ tâm", "fadecenter", "◎" } };
        void TabAnimations(RTab tab)
        {
            RGroup pv = tab.Add("Xem trước"); pv.Icon = "play";
            pv.Large(I("Xem trước", "play", "Trình chiếu trang hiện tại", () => SlideShow(doc.Part, true)));
            RGroup g = tab.Add("Hoạt hình"); g.Icon = "animation";
            RItem gal = new RItem { Kind = RK.Gallery, Width = 6 * 78 + 22, Gallery = new List<RGalleryItem>() };
            gal.Gallery.Add(new RGalleryItem { Name = "Không", Font = Ui.F(15f), Color = accent, Preview = "▢", Click = () => Animate("none") });
            foreach (string[] a in AnimList) { string[] aa = a; gal.Gallery.Add(new RGalleryItem { Name = aa[0], Font = Ui.F(15f), Color = Color.FromArgb(0x2E, 0x9E, 0x44), Preview = aa[2], Click = () => Animate(aa[1]) }); }
            g.Gallery(gal);
            g.Large(M("Tùy chọn hiệu ứng", "animation", "Tốc độ của hiệu ứng",
                Mi("Nhanh", null, () => animSpeed = "fast"), Mi("Vừa", null, () => animSpeed = "medium"), Mi("Chậm", null, () => animSpeed = "slow")));
            RGroup adv = tab.Add("Hoạt hình Nâng cao"); adv.Icon = "animpane";
            adv.Large(I("Ngăn Hoạt hình", "animpane", "Danh sách hiệu ứng trên trang chiếu này", AnimationPane));
            adv.Stack(I("Xóa mọi hiệu ứng", "reject", "Bỏ mọi hoạt hình trên trang chiếu này", () => { string r = Tool("ClearAnimations"); if (ToolOk(r, "Hoạt hình")) { MarkDirty(); status.Segs[0].Text = "Đã xóa " + r.Substring(2).Trim() + " hiệu ứng"; status.Invalidate(); } }));
        }
        void Animate(string effect)
        {
            string r = Tool("Animate", "effect", effect, "speed", animSpeed);
            if (r != null && r.StartsWith("ERR:select")) { MessageBox.Show(this, "Hãy bấm chọn hình, ảnh hoặc hộp văn bản trên trang chiếu trước.", "Hoạt hình", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (ToolOk(r, "Hoạt hình")) { MarkDirty(); status.Segs[0].Text = effect == "none" ? "Đã bỏ hiệu ứng" : "Đã thêm hiệu ứng cho " + r.Substring(2).Trim() + " đối tượng"; status.Invalidate(); }
        }
        void AnimationPane()
        {
            string r = Tool("ListAnimations"); if (!ToolOk(r, "Ngăn Hoạt hình")) return;
            StringBuilder sb = new StringBuilder(); int n = 0;
            foreach (string line in r.Split('\n')) { string[] f = line.Split('|'); if (f.Length < 3) continue; n++; sb.Append(n).Append(".  ").Append(f[2]).Append("\r\n"); }
            MessageBox.Show(this, n == 0 ? "Trang chiếu này chưa có hiệu ứng hoạt hình.\r\nChọn một đối tượng rồi bấm một hiệu ứng trong nhóm Hoạt hình." : "Thứ tự hoạt hình trên trang chiếu " + (doc.Part + 1) + ":\r\n\r\n" + sb, "Ngăn Hoạt hình", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ---------------- Mailings (Word) - v0.1.7 ----------------
        string mergeSource;
        void TabMailings(RTab tab)
        {
            RGroup st = tab.Add("Bắt đầu Phối Thư"); st.Icon = "mail";
            st.Large(I("Chọn người nhận", "mail", "Chọn danh sách người nhận (Excel .xlsx/.xls/.ods/.csv, dòng 1 là tiêu đề cột)", PickMergeSource));
            RGroup wf = tab.Add("Ghi & Chèn Trường"); wf.Icon = "mergefield";
            wf.Large(I("Chèn trường phối", "mergefield", "Chèn «Tên cột» tại vị trí con trỏ", InsertMergeField));
            RGroup fin = tab.Add("Hoàn tất"); fin.Icon = "mailmerge";
            fin.Large(M("Hoàn tất & Phối", "mailmerge", "Tạo một tệp cho mỗi người nhận",
                Mi("Phối ra các tệp Word (.docx)…", "mailmerge", () => RunMerge("docx")), Mi("Phối ra các tệp PDF…", "pdf", () => RunMerge("pdf"))));
        }
        void PickMergeSource()
        {
            using (OpenFileDialog o = new OpenFileDialog { Filter = "Danh sách (*.xlsx;*.xls;*.xlsm;*.ods;*.csv)|*.xlsx;*.xls;*.xlsm;*.ods;*.csv", Title = "Chọn danh sách người nhận" })
                if (o.ShowDialog(this) == DialogResult.OK) { mergeSource = o.FileName; mergeRows = null; status.Segs[0].Text = "Người nhận: " + Path.GetFileName(mergeSource); status.Invalidate(); }
        }
        // recipient list -> rows of strings (row 0 = column names); .xlsx/.xls/.ods are converted to CSV by the engine first
        List<string[]> mergeRows;
        List<string[]> MergeRows()
        {
            if (mergeRows != null) return mergeRows;
            string csv = mergeSource;
            try
            {
                string ext0 = Path.GetExtension(mergeSource).ToLowerInvariant();
                if (ext0 == ".xlsx" || ext0 == ".xlsm") { List<string[]> xr = XlsxRows(mergeSource); if (xr != null && xr.Count > 0) { mergeRows = xr; return xr; } }
                if (!mergeSource.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    csv = Path.Combine(Path.GetTempPath(), "oos_merge_" + Guid.NewGuid().ToString("N").Substring(0, 6) + ".csv");
                    Cursor = Cursors.WaitCursor;
                    try { if (!Lok.Convert(mergeSource, csv, "44,34,76,1,,0,false,true,false,false")) { MessageBox.Show(this, "Không đọc được danh sách người nhận.", "Phối thư"); return null; } }
                    finally { Cursor = Cursors.Default; Lok.Activate(doc); }
                }
                List<string[]> rows = new List<string[]>();
                byte[] raw = File.ReadAllBytes(csv); string txt;
                try { txt = new UTF8Encoding(false, true).GetString(raw); } catch { txt = Encoding.Default.GetString(raw); }   // CSV saved by Excel in the ANSI code page
                foreach (string line in txt.Replace("\r\n", "\n").Split('\n'))
                {
                    if (line.Trim().Length == 0) continue;
                    List<string> f = new List<string>(); StringBuilder cur = new StringBuilder(); bool q = false;
                    for (int i = 0; i < line.Length; i++)
                    {
                        char c = line[i];
                        if (q) { if (c == '"') { if (i + 1 < line.Length && line[i + 1] == '"') { cur.Append('"'); i++; } else q = false; } else cur.Append(c); }
                        else if (c == '"') q = true; else if (c == ',') { f.Add(cur.ToString()); cur.Length = 0; } else cur.Append(c);
                    }
                    f.Add(cur.ToString()); rows.Add(f.ToArray());
                }
                if (rows.Count == 0) { MessageBox.Show(this, "Danh sách người nhận trống.", "Phối thư"); return null; }
                mergeRows = rows; return rows;
            }
            catch (Exception ex) { MessageBox.Show(this, "Không đọc được danh sách người nhận: " + ex.Message, "Phối thư"); return null; }
        }
        // first worksheet of an .xlsx read directly (keeps Vietnamese text exactly; no engine round trip)
        static List<string[]> XlsxRows(string file)
        {
            try
            {
                string wb = DocMargins.Entry(file, "xl/workbook.xml"), rels = DocMargins.Entry(file, "xl/_rels/workbook.xml.rels"), sheetPath = "xl/worksheets/sheet1.xml";
                System.Text.RegularExpressions.Match ms = System.Text.RegularExpressions.Regex.Match(wb ?? "", "<sheet [^>]*r:id=\"([^\"]+)\"");
                if (ms.Success && rels != null)
                {
                    System.Text.RegularExpressions.Match mr = System.Text.RegularExpressions.Regex.Match(rels, "<Relationship [^>]*Id=\"" + ms.Groups[1].Value + "\"[^>]*Target=\"([^\"]+)\"");
                    if (!mr.Success) mr = System.Text.RegularExpressions.Regex.Match(rels, "<Relationship [^>]*Target=\"([^\"]+)\"[^>]*Id=\"" + ms.Groups[1].Value + "\"");
                    if (mr.Success) { string t = mr.Groups[1].Value.TrimStart('/'); sheetPath = t.StartsWith("xl/") ? t : "xl/" + t; }
                }
                string sh = DocMargins.Entry(file, sheetPath); if (sh == null) return null;
                List<string> shared = new List<string>(); string ss = DocMargins.Entry(file, "xl/sharedStrings.xml");
                if (ss != null) foreach (System.Text.RegularExpressions.Match si in System.Text.RegularExpressions.Regex.Matches(ss, "<si>(.*?)</si>", System.Text.RegularExpressions.RegexOptions.Singleline))
                    { StringBuilder sb = new StringBuilder(); foreach (System.Text.RegularExpressions.Match t in System.Text.RegularExpressions.Regex.Matches(si.Groups[1].Value, "<t[^>]*>(.*?)</t>", System.Text.RegularExpressions.RegexOptions.Singleline)) sb.Append(t.Groups[1].Value); shared.Add(System.Net.WebUtility.HtmlDecode(sb.ToString())); }
                List<string[]> rows = new List<string[]>();
                foreach (System.Text.RegularExpressions.Match row in System.Text.RegularExpressions.Regex.Matches(sh, "<row[^>]*>(.*?)</row>", System.Text.RegularExpressions.RegexOptions.Singleline))
                {
                    SortedDictionary<int, string> cells = new SortedDictionary<int, string>(); int maxc = -1;
                    foreach (System.Text.RegularExpressions.Match c in System.Text.RegularExpressions.Regex.Matches(row.Groups[1].Value, "<c ([^>]*?)(?:/>|>(.*?)</c>)", System.Text.RegularExpressions.RegexOptions.Singleline))
                    {
                        string at = c.Groups[1].Value, body = c.Groups[2].Value;
                        System.Text.RegularExpressions.Match rr = System.Text.RegularExpressions.Regex.Match(at, "r=\"([A-Z]+)"); if (!rr.Success) continue;
                        int col = 0; foreach (char ch in rr.Groups[1].Value) col = col * 26 + (ch - 'A' + 1); col--;
                        System.Text.RegularExpressions.Match tt = System.Text.RegularExpressions.Regex.Match(at, "t=\"(\\w+)\""), v = System.Text.RegularExpressions.Regex.Match(body, "<v>(.*?)</v>", System.Text.RegularExpressions.RegexOptions.Singleline);
                        string val = "";
                        if (tt.Success && tt.Groups[1].Value == "s" && v.Success) { int k; if (int.TryParse(v.Groups[1].Value, out k) && k < shared.Count) val = shared[k]; }
                        else if (tt.Success && tt.Groups[1].Value == "inlineStr") { System.Text.RegularExpressions.Match it = System.Text.RegularExpressions.Regex.Match(body, "<t[^>]*>(.*?)</t>", System.Text.RegularExpressions.RegexOptions.Singleline); if (it.Success) val = System.Net.WebUtility.HtmlDecode(it.Groups[1].Value); }
                        else if (v.Success) val = System.Net.WebUtility.HtmlDecode(v.Groups[1].Value);
                        cells[col] = val; maxc = Math.Max(maxc, col);
                    }
                    if (maxc < 0) continue;
                    string[] arr = new string[maxc + 1]; for (int i = 0; i <= maxc; i++) arr[i] = cells.ContainsKey(i) ? cells[i] : "";
                    rows.Add(arr);
                }
                return rows;
            }
            catch (Exception ex) { Lok.Trace("xlsx rows " + ex.Message); return null; }
        }
        void InsertMergeField()
        {
            if (mergeSource == null) PickMergeSource(); if (mergeSource == null) return;
            List<string[]> rows = MergeRows(); if (rows == null) return;
            List<RMenu> menu = new List<RMenu>();
            foreach (string f in rows[0]) { string ff = f.Trim(); if (ff.Length == 0) continue; menu.Add(Mi(ff, "mergefield", () => { if (ToolOk(Tool("InsertMergeField", "field", ff), "Phối thư")) MarkDirty(); })); }
            if (menu.Count == 0) { MessageBox.Show(this, "Dòng 1 của danh sách không có tên cột.", "Phối thư"); return; }
            ribbon.ShowMenu(menu, ribbon.PointToClient(Cursor.Position));
        }
        void RunMerge(string fmt)
        {
            if (mergeSource == null) PickMergeSource(); if (mergeSource == null) return;
            List<string[]> rows = MergeRows(); if (rows == null) return;
            string defDir = Path.Combine(Path.GetDirectoryName(path ?? mergeSource), "Phối thư - " + Path.GetFileNameWithoutExtension(path ?? "tai lieu"));
            using (FolderBrowserDialog fb = new FolderBrowserDialog { Description = "Chọn thư mục lưu các tệp đã phối (mặc định: thư mục \"Phối thư\" cạnh tài liệu)", ShowNewFolderButton = true })
            {
                try { Directory.CreateDirectory(defDir); fb.SelectedPath = defDir; } catch { }
                if (fb.ShowDialog(this) != DialogResult.OK) return;
                int made = 0; string err = null; string[] head = rows[0];
                for (int r = 1; r < rows.Count; r++)
                {
                    List<string> kv = new List<string>();
                    for (int c = 0; c < head.Length; c++) { kv.Add("f" + c); kv.Add(head[c]); kv.Add("v" + c); kv.Add(c < rows[r].Length ? rows[r][c] : ""); }
                    string key = rows[r].Length > 0 ? rows[r][0] : "";
                    foreach (char ch in Path.GetInvalidFileNameChars()) key = key.Replace(ch, '_');
                    if (key.Length > 60) key = key.Substring(0, 60);
                    kv.Add("format"); kv.Add(fmt); kv.Add("out"); kv.Add(Path.Combine(fb.SelectedPath, r.ToString("000") + " - " + key + (fmt == "pdf" ? ".pdf" : ".docx")));
                    status.Segs[0].Text = "Đang phối " + r + " / " + (rows.Count - 1) + "…"; status.Invalidate(); status.Update();
                    string res = Tool("MergeOne", kv.ToArray());
                    if (res == null || !res.StartsWith("OK")) { err = res ?? "không có kết quả"; break; }
                    made++;
                }
                status.Segs[0].Text = "Đã phối " + made + " tệp"; status.Invalidate();
                if (err != null) MessageBox.Show(this, "Dừng ở bản ghi " + (made + 1) + ": " + err, "Phối thư", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else MessageBox.Show(this, "Đã tạo " + made + " tệp trong:\r\n" + fb.SelectedPath, "Phối thư", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        void TabSlideShow(RTab tab)
        {
            RGroup st = tab.Add("Bắt đầu Trình chiếu"); st.Icon = "play";
            st.Large(I("Từ đầu", "play", "Trình chiếu từ trang đầu (F5)", () => SlideShow(0, true)));
            st.Large(I("Từ trang chiếu hiện tại", "playcurrent", "Trình chiếu từ trang hiện tại (Shift+F5)", () => SlideShow(doc.Part, true)));
            RGroup su = tab.Add("Thiết lập"); su.Icon = "hideslide";
            su.Large(St("Ẩn trang chiếu", "hideslide", "Không chiếu trang này khi trình chiếu", ".uno:HideSlide"));
            List<RMenu> mons = new List<RMenu>();
            for (int i = 0; i < Screen.AllScreens.Length; i++) { int ii = i; Screen sc = Screen.AllScreens[i]; mons.Add(new RMenu("Màn hình " + (i + 1) + "  (" + sc.Bounds.Width + "×" + sc.Bounds.Height + (sc.Primary ? ", chính" : "") + ")", "pc", () => NativeSettings.Set("ppt.monitor", ii.ToString())) { Checked = () => NativeSettings.Int("ppt.monitor", Screen.AllScreens.Length > 1 ? 1 : 0) == ii }); }
            RGroup mo = tab.Add("Màn hình"); mo.Large(I("Màn hình trình chiếu", "pc", "Chọn màn hình/TV để trình chiếu", null)); mo.Cols[0].Large.Menu = mons;
        }

        // ---------------- Review ----------------
        void TabReview(RTab tab)
        {
            RGroup p = tab.Add("Soát lỗi"); p.Icon = "spelling";
            p.Large(St("Chính tả", "spelling", "Tự động gạch chân lỗi chính tả", ".uno:SpellOnline"));
            if (type == LokDoc.TEXT) p.Stack(I("Thống kê từ", "wordcount", "Số từ, ký tự, đoạn văn", WordCount));
            RGroup c = tab.Add(type == LokDoc.SPREADSHEET ? "Ghi chú" : "Bình luận"); c.Icon = "comment";
            c.Large(I(type == LokDoc.SPREADSHEET ? "Ghi chú mới" : "Bình luận mới", "comment", "Thêm ghi chú về phần đang chọn", InsertComment));
            if (type == LokDoc.TEXT) c.Stack(I("Xóa", "reject", "Xóa bình luận hiện tại", U(".uno:DeleteComment")), I("Xóa tất cả", "reject", "Xóa mọi bình luận", U(".uno:DeleteAllNotes")));
            if (type == LokDoc.SPREADSHEET) c.Stack(I("Xóa ghi chú", "reject", null, U(".uno:DeleteNote")), I("Hiện/ẩn ghi chú", "comment", null, U(".uno:ShowNote")));
            if (type == LokDoc.PRESENTATION) c.Stack(I("Xóa tất cả", "reject", "Xóa mọi bình luận", U(".uno:DeleteAllAnnotation")));
            if (type == LokDoc.TEXT)
            {
                RGroup tr = tab.Add("Theo dõi"); tr.Icon = "track";
                tr.Large(St("Theo dõi thay đổi", "track", "Ghi lại mọi thay đổi trong tài liệu (Ctrl+Shift+E)", ".uno:TrackChanges"));
                tr.Stack(St("Hiện đánh dấu", "readmode", "Hiện các thay đổi được theo dõi", ".uno:ShowTrackedChanges"));
                RGroup ch = tab.Add("Thay đổi"); ch.Icon = "accept";
                ch.Large(Sp("Chấp nhận", "accept", "Chấp nhận thay đổi này", U(".uno:AcceptTrackedChange"), Mi("Chấp nhận thay đổi này", "accept", U(".uno:AcceptTrackedChange")), Mi("Chấp nhận mọi thay đổi", "accept", U(".uno:AcceptAllTrackedChanges"))));
                ch.Large(Sp("Từ chối", "reject", "Từ chối thay đổi này", U(".uno:RejectTrackedChange"), Mi("Từ chối thay đổi này", "reject", U(".uno:RejectTrackedChange")), Mi("Từ chối mọi thay đổi", "reject", U(".uno:RejectAllTrackedChanges"))));
                ch.Stack(I("Trước", "indentdec", "Thay đổi trước", () => UnoView(".uno:PreviousTrackedChange")), I("Tiếp", "indentinc", "Thay đổi tiếp theo", () => UnoView(".uno:NextTrackedChange")));
            }
            GroupMacros(tab);
        }

        // ---------------- View ----------------
        void TabView(RTab tab)
        {
            RGroup v = tab.Add("Dạng xem"); v.Icon = "printlayout";
            if (type == LokDoc.TEXT)
            {
                v.Large(I("Chế độ đọc", "readmode", "Thu gọn Ribbon, vừa trang để đọc", () => ReadMode(true)));
                v.Large(I("Bố cục in", "printlayout", "Xem tài liệu như khi in", () => { ReadMode(false); SetWebLayout(false); }));
                v.Large(I("Bố cục web", "weblayout", "Xem tài liệu như trang web", () => { ReadMode(false); SetWebLayout(true); }));
                RGroup sh = tab.Add("Hiển thị"); sh.Icon = "ruler";
                sh.Stack(I("Thước", "ruler", "Hiện/ẩn thước", () => { bool on = !rulerTop.Visible; rulerTop.Visible = rulerLeft.Visible = on; NativeSettings.Set("word.ruler", on); }), MiniStToSmall(St("Ký tự định dạng", "pilcrow", "Hiện ký tự định dạng ẩn", ".uno:ControlCodes")));
            }
            else if (type == LokDoc.SPREADSHEET)
            {
                v.Large(I("Thường", "normalview", null, () => UnoView(".uno:NormalViewMode")));
                v.Large(I("Xem trước ngắt trang", "pagebreak", null, () => UnoView(".uno:PagebreakMode")));
                RGroup sh = tab.Add("Hiển thị"); sh.Icon = "gridlines";
                sh.Stack(I("Đường lưới", "gridlines", null, () => UnoView(".uno:ToggleSheetGrid")), I("Thanh công thức", "function", null, () => formula.Visible = !formula.Visible), I("Đầu đề", "freeze", null, ToggleHeaders));
            }
            else
            {
                v.Large(I("Thường", "normalview", null, () => { if (notesMode) ToggleNotes(); ShowSorter(false); }));
                v.Large(I("Sắp xếp trang chiếu", "sorter", null, () => ShowSorter(true)));
                v.Large(I("Trang ghi chú", "notes", null, ToggleNotes));
                v.Large(I("Dạng xem đọc", "readmode", null, () => SlideShow(doc.Part, false)));
            }
            RGroup z = tab.Add("Thu phóng"); z.Icon = "zoom";
            z.Large(I("Thu phóng", "zoom", "Chọn mức thu phóng", ZoomDialog));
            z.Large(I("100%", "zoom100", "Thu phóng 100%", () => canvas.SetZoom(1f)));
            if (type == LokDoc.TEXT) z.Stack(I("Một trang", "onepage", null, () => { Size p = PageTw(); canvas.SetZoom(Math.Min((canvas.Width - 40) * 15f / p.Width, (canvas.Height - 30) * 15f / p.Height)); }), I("Nhiều trang", "multipage", null, () => { Size p = PageTw(); canvas.SetZoom((canvas.Height - 30) * 15f / p.Height); }), I("Chiều rộng trang", "pagewidth", null, () => canvas.SetZoom(FitWidthZoom(4f))));
            if (type == LokDoc.PRESENTATION) z.Large(I("Vừa cửa sổ", "pagewidth", null, () => canvas.SetZoom(FitSlideZoom())));
            if (type == LokDoc.SPREADSHEET)
            {
                RGroup w = tab.Add("Cửa sổ"); w.Icon = "freeze";
                w.Large(M("Cố định ngăn", "freeze", "Giữ một phần trang tính hiển thị khi cuộn", Mi("Cố định ngăn (theo ô đang chọn)", null, U(".uno:FreezePanes")), Mi("Cố định hàng trên cùng", null, U(".uno:FreezePanesFirstRow")), Mi("Cố định cột đầu tiên", null, U(".uno:FreezePanesFirstColumn"))));
                w.Stack(St("Tách", "freeze", "Tách cửa sổ thành nhiều ngăn", ".uno:SplitWindow"));
            }
            RGroup fs = tab.Add("Cửa sổ "); fs.Large(I("Toàn màn hình", "fullscreen", "Toàn màn hình (F11)", ToggleFullScreen));
            GroupMacros(tab);
        }
        static RItem MiniStToSmall(RItem i) { i.Kind = RK.Small; return i; }

        void TabHelp(RTab tab)
        {
            RGroup h = tab.Add("Trợ giúp"); h.Icon = "info";
            h.Large(I("Trợ giúp", "info", "Phím tắt và hướng dẫn nhanh", HelpDialog));
            h.Large(I("Giới thiệu", "account", "Phiên bản và thông tin engine", () => OpenBackstage("account")));
        }

        // ---------------- commands ----------------
        void OpenInWeb() { if (dirty) Save(true); App.AddPending(path, true); App.ActivateExistingWindow(); }
        void InsertComment()
        {
            string t = Ask(this, type == LokDoc.SPREADSHEET ? "Ghi chú mới" : "Bình luận mới", "Nội dung:", ""); if (string.IsNullOrEmpty(t)) return;
            Uno(".uno:InsertAnnotation", Args(JS("Text", t), JS("Author", Ui.UserName())));
        }
        void InsertLink()
        {
            string sel = null; try { sel = doc.GetTextSelection("text/plain;charset=utf-8"); } catch { }
            using (LinkDialog d = new LinkDialog(sel)) if (d.ShowDialog(this) == DialogResult.OK)
                    Uno(".uno:SetHyperlink", Args(JS("Hyperlink.Text", d.TextValue), JS("Hyperlink.URL", d.Url), JL("Hyperlink.Type", 1)));
        }
        void InsertImage()
        {
            using (OpenFileDialog d = new OpenFileDialog { Filter = "Hình ảnh|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.svg;*.webp;*.tif;*.tiff;*.emf;*.wmf" })
                if (d.ShowDialog(this) == DialogResult.OK) Uno(".uno:InsertGraphic", Args(JS("FileName", new Uri(d.FileName).AbsoluteUri)));
        }
        void InsertScreenshot()
        {
            WindowState = FormWindowState.Minimized; Application.DoEvents(); System.Threading.Thread.Sleep(450);
            Rectangle b = Screen.PrimaryScreen.Bounds; string f = Path.Combine(Path.GetTempPath(), "oos-screenshot-" + DateTime.Now.ToString("HHmmss") + ".png");
            using (Bitmap bmp = new Bitmap(b.Width, b.Height)) { using (Graphics g = Graphics.FromImage(bmp)) g.CopyFromScreen(b.Location, Point.Empty, b.Size); bmp.Save(f, ImageFormat.Png); }
            WindowState = FormWindowState.Normal; Activate();
            Uno(".uno:InsertGraphic", Args(JS("FileName", new Uri(f).AbsoluteUri)));
        }
        [DllImport("user32")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
        void Dictate()
        {
            if (Environment.OSVersion.Version < new Version(6, 2)) { MessageBox.Show("Nhập liệu bằng giọng nói cần Windows 10/11 (Win+H).", "Đọc chính tả"); return; }
            canvas.Focus(); Application.DoEvents();
            keybd_event(0x5B, 0, 0, UIntPtr.Zero); keybd_event(0x48, 0, 0, UIntPtr.Zero); keybd_event(0x48, 0, 2, UIntPtr.Zero); keybd_event(0x5B, 0, 2, UIntPtr.Zero);
        }
        void WordCount()
        {
            string all = null; Lok.Activate(doc);
            doc.Uno(".uno:SelectAll", null);
            try { all = doc.GetTextSelection("text/plain;charset=utf-8"); } catch { }
            doc.ResetSelection(); all = all ?? "";
            int words = all.Split(new[] { ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
            int paras = all.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;
            string noSpace = all.Replace(" ", "").Replace("\r", "").Replace("\n", "").Replace("\t", "");
            MessageBox.Show("Trang:\t\t\t" + Math.Max(1, doc.PageRects().Count) + "\nTừ:\t\t\t" + words + "\nKý tự (không có khoảng trắng):\t" + noSpace.Length + "\nKý tự (có khoảng trắng):\t" + all.Replace("\r", "").Replace("\n", "").Length + "\nĐoạn văn:\t\t\t" + paras, "Thống kê từ");
        }
        void GoToPage(int n)
        {
            List<Rectangle> pr = doc.PageRects(); if (n < 1 || n > pr.Count) return;
            canvas.ScrollTo(canvas.Scroll.X, (int)(pr[n - 1].Y / (15f / canvas.Zoom)));
        }
        void GoToDialog()
        {
            if (type == LokDoc.TEXT) { string p = Ask(this, "Đi tới", "Số trang:", ""); int n; if (int.TryParse(p, out n)) GoToPage(n); }
            else if (type == LokDoc.SPREADSHEET) { formula.NameBox.Focus(); formula.NameBox.SelectAll(); }
            else { string p = Ask(this, "Đi tới", "Số trang chiếu:", ""); int n; if (int.TryParse(p, out n)) SelectPart(n - 1); }
        }
        bool readMode;
        void ReadMode(bool on)
        {
            readMode = on; ribbon.SetMinimized(on);
            if (on) { Size p = PageTw(); canvas.SetZoom(Math.Min((canvas.Width - 40) * 15f / p.Width, (canvas.Height - 30) * 15f / p.Height) * 1.6f); }
            status.ActiveView = on ? 0 : 1; status.Invalidate();
        }
        bool webLayout;
        void SetWebLayout(bool on) { if (on != webLayout) { webLayout = on; UnoView(".uno:BrowseView"); canvas.RefreshSize(); SyncBars(); } status.ActiveView = on ? 2 : 1; status.Invalidate(); }
        void ZoomDialog()
        {
            using (ZoomForm z = new ZoomForm(canvas.Zoom)) if (z.ShowDialog(this) == DialogResult.OK)
                {
                    if (z.Fit == 1) canvas.SetZoom(FitWidthZoom(4f));
                    else if (z.Fit == 2) { Size p = type == LokDoc.TEXT ? PageTw() : doc.SizeTwips; canvas.SetZoom(Math.Min((canvas.Width - 40) * 15f / p.Width, (canvas.Height - 30) * 15f / p.Height)); }
                    else canvas.SetZoom(z.Value);
                }
        }
        void ToggleMacros()
        {
            if (!Lok.MacrosAllowed && MessageBox.Show("Cho phép chạy macro VBA/Basic trong tài liệu?\n\nChỉ bật với tệp từ nguồn tin cậy. Thay đổi có hiệu lực khi mở lại tài liệu.", "Macro", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            Lok.MacrosAllowed = !Lok.MacrosAllowed;
            MessageBox.Show("Đã " + (Lok.MacrosAllowed ? "bật" : "tắt") + " macro. Hãy đóng và mở lại tài liệu để áp dụng.", "Macro");
            BuildRibbon();
        }
        void RunMacro()
        {
            string name = Ask(this, "Chạy macro", "Tên macro (VBA: Module1.TenMacro  hoặc  Basic: Standard.Module1.Main):", "Module1.");
            if (string.IsNullOrEmpty(name)) return;
            if (!Lok.MacrosAllowed) { MessageBox.Show("Macro đang bị tắt. Bật ở Xem lại → Macro → Bật macro, rồi mở lại tài liệu.", "Macro"); return; }
            // LibreOfficeKit runs macros only through runMacro with macro:// URLs ("macro://./" = the active document)
            string url = name.StartsWith("macro:") ? name
                : name.StartsWith("vnd.sun.star.script:") ? "macro://./" + name.Substring(20).Split('?')[0]
                : name.Split('.').Length == 2 ? "macro://./VBAProject." + name : "macro://./" + name;
            Lok.Activate(doc); if (!Lok.RunMacro(url)) MessageBox.Show(this, "Không chạy được macro \"" + name + "\".\n" + Lok.LastError(), "Macro", MessageBoxButtons.OK, MessageBoxIcon.Warning); else MarkDirty();
        }
        void HelpDialog()
        {
            MessageBox.Show(
                "Phím tắt thường dùng\n\n" +
                "Ctrl+S  Lưu            F12  Lưu thành            Ctrl+P  In\n" +
                "Ctrl+Z / Ctrl+Y  Hoàn tác / Làm lại\n" +
                "Ctrl+B / I / U  Đậm / Nghiêng / Gạch dưới\n" +
                "Ctrl+L / E / R / J  Căn trái / giữa / phải / đều\n" +
                "Ctrl+F  Tìm        Ctrl+H  Thay thế        Ctrl+K  Liên kết\n" +
                "Alt+Q  Ô tìm kiếm lệnh        Ctrl+F1  Thu gọn / ghim Ribbon\n" +
                "Ctrl + cuộn chuột  Thu phóng        F11  Toàn màn hình\n" +
                (type == LokDoc.PRESENTATION ? "F5  Trình chiếu từ đầu        Shift+F5  Từ trang hiện tại\n" : "") +
                (type == LokDoc.SPREADSHEET ? "Alt+=  Tự động tính tổng        F9  Tính lại\n" : "") +
                "\nYomi-Office Open Source " + App.Version + " — engine LibreOffice (Native)", "Trợ giúp");
        }

        // ---------------- find / replace ----------------
        FindForm findForm;
        void FindReplace(bool replace)
        {
            string sel = null; try { sel = doc.GetTextSelection("text/plain;charset=utf-8"); } catch { }
            if (findForm == null || findForm.IsDisposed) { findForm = new FindForm(); findForm.FindNext += (q, mc, back) => Find(q, back, mc); findForm.Replace += (q, r, mc, all) => { Uno(".uno:ExecuteSearch", Args(JS("SearchItem.SearchString", q), JS("SearchItem.ReplaceString", r), JB("SearchItem.MatchCase", mc), JL("SearchItem.Command", all ? 3 : 2))); }; }
            findForm.Show(this, replace, sel != null && sel.Length < 80 && !sel.Contains("\n") ? sel : null);
        }
        void Find(string q, bool back) { Find(q, back, false); }
        void Find(string q, bool back, bool mc) { Lok.Activate(doc); doc.Uno(".uno:ExecuteSearch", Args(JS("SearchItem.SearchString", q), JB("SearchItem.Backward", back), JB("SearchItem.MatchCase", mc), JL("SearchItem.Command", 0))); }

        // ---------------- font / paragraph / page setup dialogs ----------------
        void FontDialogShow()
        {
            using (FontDialog d = new FontDialog { ShowColor = true, ShowEffects = true, AllowVerticalFonts = false, FontMustExist = false, Color = lastFontColor })
            {
                try { d.Font = new Font(string.IsNullOrEmpty(CurrentState(".uno:CharFontName")) ? "Times New Roman" : CurrentState(".uno:CharFontName"), Math.Max(6f, ParseF(CurrentState(".uno:FontHeight"), 12f))); } catch { }
                if (d.ShowDialog(this) != DialogResult.OK) return;
                Font f = d.Font;
                Uno(".uno:CharFontName", Args(JS("CharFontName.FamilyName", f.FontFamily.Name)));
                Uno(".uno:FontHeight", "{\"FontHeight.Height\":{\"type\":\"float\",\"value\":\"" + f.SizeInPoints.ToString(CultureInfo.InvariantCulture) + "\"}}");
                SetToggle(".uno:Bold", f.Bold); SetToggle(".uno:Italic", f.Italic); SetToggle(".uno:Underline", f.Underline); SetToggle(".uno:Strikeout", f.Strikeout);
                lastFontColor = d.Color; Uno(".uno:Color", Args(JL("Color", Rgb(d.Color))));
            }
        }
        readonly Dictionary<string, string> states = new Dictionary<string, string>();
        string CurrentState(string k) { string v; return states.TryGetValue(k, out v) ? v : ""; }
        static float ParseF(string s, float def) { float v; return float.TryParse((s ?? "").Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? v : def; }
        void SetToggle(string uno, bool want) { if ((CurrentState(uno) == "true") != want) Uno(uno); }
        void ParagraphDialog()
        {
            using (ParaForm p = new ParaForm(type == LokDoc.TEXT)) if (p.ShowDialog(this) == DialogResult.OK)
                {
                    string[] al = { ".uno:LeftPara", ".uno:CenterPara", ".uno:RightPara", ".uno:JustifyPara" }; if (p.Align >= 0) Uno(al[p.Align]);
                    string[] ls = { ".uno:SpacePara1", ".uno:SpacePara15", ".uno:SpacePara2" }; if (p.LineSpacing >= 0) Uno(ls[p.LineSpacing]);
                    for (int i = 0; i < Math.Abs(p.IndentSteps); i++) Uno(p.IndentSteps > 0 ? ".uno:IncrementIndent" : ".uno:DecrementIndent");
                    for (int i = 0; i < Math.Abs(p.SpaceSteps); i++) Uno(p.SpaceSteps > 0 ? ".uno:ParaspaceIncrease" : ".uno:ParaspaceDecrease");
                }
        }
        public void ShowPageSetup()
        {
            Size s = PageTw(); if (type != LokDoc.TEXT) s = new Size(11906, 16838);
            using (PageSetupForm f = new PageSetupForm(s)) if (f.ShowDialog(this) == DialogResult.OK)
                {
                    SetLandscape(f.Landscape);
                    Uno(".uno:AttributePageSize", Args(JL("AttributePageSize.Width", (f.Landscape ? f.PaperH : f.PaperW)), JL("AttributePageSize.Height", (f.Landscape ? f.PaperW : f.PaperH))));
                    Uno(".uno:PageULMargin", Args(JL("PageULMargin.Upper", f.MTop), JL("PageULMargin.Lower", f.MBottom)));
                    Uno(".uno:PageLRMargin", Args(JL("PageLRMargin.Left", f.Left_), JL("PageLRMargin.Right", f.Right_)));
                    if (type == LokDoc.TEXT) SetVMargins(f.MTop * 1440 / 2540, f.MBottom * 1440 / 2540);
                    canvas.RefreshSize(); SyncBars(); canvas.Invalidate();
                }
        }

        // ---------------- slide show (own renderer, with transitions) ----------------
        void SlideShow(int start, bool fullScreen)
        {
            if (type != LokDoc.PRESENTATION) return;
            Lok.Activate(doc);
            int mi = Math.Min(Screen.AllScreens.Length - 1, NativeSettings.Int("ppt.monitor", Screen.AllScreens.Length > 1 ? 1 : 0));
            Screen scr = Screen.AllScreens[Math.Max(0, mi)];
            Form show = new Form { FormBorderStyle = fullScreen ? FormBorderStyle.None : FormBorderStyle.Sizable, BackColor = Color.Black, KeyPreview = true, StartPosition = FormStartPosition.Manual, Text = Path.GetFileName(path) + " - Trình chiếu" };
            if (fullScreen) { show.Bounds = scr.Bounds; show.TopMost = true; } else { show.Bounds = new Rectangle(Bounds.X + 60, Bounds.Y + 60, Width * 2 / 3, Height * 2 / 3); }
            DoubleBufferedPanel pb = new DoubleBufferedPanel { Dock = DockStyle.Fill, BackColor = Color.Black };
            show.Controls.Add(pb);
            int cur = Math.Max(0, Math.Min(start, doc.Parts - 1)), orig = doc.Part;
            Bitmap shown = null, prev = null; DateTime tStart = DateTime.MinValue; string fx = transition; int ms = transitionMs;
            Timer anim = new Timer { Interval = 15 };
            Func<int, Bitmap> render = i =>
            {
                canvas.Suspend = true; try { doc.Part = i; Size tw = doc.SizeTwips; if (tw.Width <= 0) return null; int w = Math.Max(640, Math.Min(2560, show.ClientSize.Width)), h = (int)((long)w * tw.Height / tw.Width); return RenderTw(new Rectangle(0, 0, tw.Width, tw.Height), w, h); }
                finally { canvas.Suspend = false; }
            };
            Action<int> go = i =>
            {
                if (i < 0) return; if (i >= doc.Parts) { show.Close(); return; }
                if (prev != null) prev.Dispose(); prev = shown; cur = i; shown = render(i); tStart = DateTime.Now; anim.Start(); pb.Invalidate();
            };
            anim.Tick += (s, e) => { pb.Invalidate(); if ((DateTime.Now - tStart).TotalMilliseconds > ms) anim.Stop(); };
            pb.Paint += (s, e) =>
            {
                Graphics g = e.Graphics; g.Clear(Color.Black); g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                if (shown == null) return;
                float sc = Math.Min(pb.Width / (float)shown.Width, pb.Height / (float)shown.Height); int w = (int)(shown.Width * sc), h = (int)(shown.Height * sc);
                Rectangle dst = new Rectangle((pb.Width - w) / 2, (pb.Height - h) / 2, w, h);
                double t = Math.Min(1, (DateTime.Now - tStart).TotalMilliseconds / Math.Max(1, ms)); if (prev == null || fx == "Không") t = 1;
                if (t < 1 && prev != null)
                {
                    if (fx == "Mờ dần") { g.DrawImage(prev, dst); using (ImageAttributes ia = new ImageAttributes()) { ColorMatrix cm = new ColorMatrix { Matrix33 = (float)t }; ia.SetColorMatrix(cm); g.DrawImage(shown, dst, 0, 0, shown.Width, shown.Height, GraphicsUnit.Pixel, ia); } return; }
                    if (fx == "Đẩy") { int off = (int)(w * t); g.DrawImage(prev, new Rectangle(dst.X - off, dst.Y, w, h)); g.DrawImage(shown, new Rectangle(dst.X + w - off, dst.Y, w, h)); return; }
                    if (fx == "Lau") { g.DrawImage(prev, dst); g.SetClip(new Rectangle(dst.X, dst.Y, (int)(w * t), h)); g.DrawImage(shown, dst); g.ResetClip(); return; }
                    if (fx == "Thu phóng") { g.DrawImage(prev, dst); int zw = (int)(w * t), zh = (int)(h * t); g.DrawImage(shown, new Rectangle(dst.X + (w - zw) / 2, dst.Y + (h - zh) / 2, zw, zh)); return; }
                    if (fx == "Tách") { g.DrawImage(prev, dst); int half = (int)(w * t / 2); g.SetClip(new Rectangle(dst.X + w / 2 - half, dst.Y, half * 2, h)); g.DrawImage(shown, dst); g.ResetClip(); return; }
                }
                g.DrawImage(shown, dst);
            };
            show.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape) show.Close();
                else if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Down || e.KeyCode == Keys.PageDown || e.KeyCode == Keys.Space || e.KeyCode == Keys.Enter || e.KeyCode == Keys.N) go(cur + 1);
                else if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Up || e.KeyCode == Keys.PageUp || e.KeyCode == Keys.Back || e.KeyCode == Keys.P) { if (cur > 0) go(cur - 1); }
                else if (e.KeyCode == Keys.Home) go(0); else if (e.KeyCode == Keys.End) go(doc.Parts - 1);
            };
            pb.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) go(cur + 1); else if (cur > 0) go(cur - 1); };
            show.FormClosed += (s, e) => { anim.Dispose(); doc.Part = orig; canvas.RefreshSize(); canvas.Invalidate(); if (shown != null) shown.Dispose(); if (prev != null) prev.Dispose(); };
            show.Shown += (s, e) => go(cur);
            show.Show(this);
        }

        // ---------------- built-in tools (Basic library "OOS" in the engine profile) ----------------
        string Tool(string macro, params string[] kv)
        {
            string tmp = Path.GetTempPath(), a = Path.Combine(tmp, "oos_macro.txt"), o = Path.Combine(tmp, "oos_macro_out.txt");
            try { if (File.Exists(o)) File.Delete(o); } catch { }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i + 1 < kv.Length; i += 2) sb.Append(kv[i]).Append('=').Append((kv[i + 1] ?? "").Replace("\r", " ").Replace("\n", " ")).Append('\n');
            try { File.WriteAllText(a, sb.ToString(), new UTF8Encoding(false)); } catch { }
            Cursor = Cursors.WaitCursor;
            try
            {
                Lok.Activate(doc); Lok.RunMacro(string.Format(BasicTools.Url, macro));
                int waitN = Array.IndexOf(kv, "wait") >= 0 ? 6000 : 60;   // v0.1.7: mail merge runs up to 5 minutes
                for (int i = 0; i < waitN && !File.Exists(o); i++) { Application.DoEvents(); System.Threading.Thread.Sleep(50); }
                string r = null; try { if (File.Exists(o)) r = File.ReadAllText(o, Encoding.UTF8); } catch { }
                Lok.Trace("tool " + macro + ": " + (r ?? "(no result)"));
                return r;
            }
            finally { Cursor = Cursors.Default; canvas.RepaintAll(); if (headers != null) headers.Invalidate(); canvas.Focus(); }
        }
        bool ToolOk(string r, string what)
        {
            if (r == null) { MessageBox.Show(this, "Không chạy được công cụ \"" + what + "\".", what, MessageBoxButtons.OK, MessageBoxIcon.Warning); return false; }
            if (r.StartsWith("ERR:select")) { MessageBox.Show(this, "Hãy chọn vùng ô cần áp dụng trước.", what, MessageBoxButtons.OK, MessageBoxIcon.Information); return false; }
            if (r.StartsWith("ERR")) { MessageBox.Show(this, "Không áp dụng được (" + r.Substring(4) + ").", what, MessageBoxButtons.OK, MessageBoxIcon.Warning); return false; }
            return true;
        }
        static string FormulaValue(string v)
        {
            v = (v ?? "").Trim(); double d;
            if (v.StartsWith("=")) return v.Substring(1);
            if (double.TryParse(v.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d.ToString(CultureInfo.InvariantCulture);
            return "\"" + v.Replace("\"", "\"\"") + "\"";
        }
        List<RMenu> CondMenu()
        {
            return new List<RMenu> {
                new RMenu("— Quy tắc tô sáng ô —", null, null),
                Mi("Lớn hơn…", null, () => CondRule("gt", "Lớn hơn", "Định dạng các ô LỚN HƠN:", true, false)),
                Mi("Nhỏ hơn…", null, () => CondRule("lt", "Nhỏ hơn", "Định dạng các ô NHỎ HƠN:", true, false)),
                Mi("Nằm giữa…", null, () => CondRule("between", "Nằm giữa", "Định dạng các ô NẰM GIỮA:", true, true)),
                Mi("Bằng với…", null, () => CondRule("eq", "Bằng với", "Định dạng các ô BẰNG:", true, false)),
                Mi("Văn bản có chứa…", null, () => CondRule("text", "Văn bản có chứa", "Định dạng các ô có chứa văn bản:", true, false)),
                Mi("Giá trị trùng lặp…", null, () => CondRule("dup", "Giá trị trùng lặp", null, false, false)),
                Mi("Giá trị duy nhất…", null, () => CondRule("unique", "Giá trị duy nhất", null, false, false)),
                new RMenu("— Quy tắc trên/dưới —", null, null),
                Mi("10 mục đầu…", null, () => CondRule("top", "10 mục đầu", "Định dạng các ô xếp hạng TRÊN CÙNG (số mục):", true, false)),
                Mi("10 mục cuối…", null, () => CondRule("bottom", "10 mục cuối", "Định dạng các ô xếp hạng DƯỚI CÙNG (số mục):", true, false)),
                Mi("Trên trung bình…", null, () => CondRule("above", "Trên trung bình", null, false, false)),
                Mi("Dưới trung bình…", null, () => CondRule("below", "Dưới trung bình", null, false, false)),
                new RMenu("— Thanh dữ liệu —", null, null),
                Mi("Thanh dữ liệu xanh lam", null, () => CondScale("databar", 0x638EC6)), Mi("Thanh dữ liệu xanh lục", null, () => CondScale("databar", 0x63BE7B)),
                Mi("Thanh dữ liệu cam", null, () => CondScale("databar", 0xFFB628)), Mi("Thanh dữ liệu đỏ", null, () => CondScale("databar", 0xF8696B)),
                new RMenu("— Thang màu —", null, null),
                Mi("Xanh lục – Vàng – Đỏ", null, () => CondScale("scale3", 0)), Mi("Trắng – Xanh lục", null, () => CondScale("scale2", 0x63BE7B)), Mi("Trắng – Đỏ", null, () => CondScale("scale2", 0xF8696B)),
                RMenu.Sep(),
                Mi("Xóa quy tắc khỏi ô đã chọn", "deletecells", () => { if (ToolOk(Tool("CondFormat", "kind", "clear"), "Định dạng có điều kiện")) MarkDirty(); }) };
        }
        void CondRule(string kind, string title, string label, bool value, bool two)
        {
            bool rank = kind == "top" || kind == "bottom";
            using (CondRuleForm f = new CondRuleForm(title, label, value, two, rank ? "10" : ""))
            {
                if (f.ShowDialog(this) != DialogResult.OK) return;
                string v1 = f.V1, v2 = f.V2;
                if (rank) { int n; if (!int.TryParse(v1, out n) || n < 1) n = 10; v1 = n.ToString(); }
                else if (value) { if (string.IsNullOrEmpty(v1)) return; v1 = FormulaValue(v1); if (two) v2 = FormulaValue(v2); }
                if (ToolOk(Tool("CondFormat", "kind", kind, "v1", v1, "v2", v2, "style", f.StyleKey), "Định dạng có điều kiện")) MarkDirty();
            }
        }
        void CondScale(string kind, int rgb) { if (ToolOk(Tool("CondFormat", "kind", kind, "color", rgb.ToString()), "Định dạng có điều kiện")) MarkDirty(); }
        void InsertPivot()
        {
            string r = Tool("ListHeaders"); if (!ToolOk(r, "PivotTable")) return;
            string[] lines = r.Replace("\r", "").Split('\n'); string[] a = lines[0].Split('\t');
            if (a.Length < 5 || lines.Length < 2) { MessageBox.Show(this, "Hãy chọn một ô trong vùng dữ liệu có hàng tiêu đề.", "PivotTable"); return; }
            if (int.Parse(a[4]) <= int.Parse(a[2])) { MessageBox.Show(this, "Vùng dữ liệu cần có hàng tiêu đề và ít nhất một hàng dữ liệu.", "PivotTable"); return; }
            string[] heads = new string[lines.Length - 1]; int lastNum = -1;
            for (int i = 0; i < heads.Length; i++) { string[] hp = lines[i + 1].Split('\t'); heads[i] = hp[0]; if (hp.Length > 1 && hp[1] == "1") lastNum = i; }
            using (PivotForm f = new PivotForm(heads, lastNum))
            {
                if (f.ShowDialog(this) != DialogResult.OK) return;
                string title = f.FuncName + " của " + heads[f.Data] + " theo " + heads[f.Row] + (f.Col >= 0 ? " và " + heads[f.Col] : "");
                string res = Tool("PivotTable", "sheet", a[0], "c1", a[1], "r1", a[2], "c2", a[3], "r2", a[4], "row", f.Row.ToString(), "col", f.Col.ToString(), "data", f.Data.ToString(), "func", f.Func, "name", "PivotTable", "title", title);
                if (!ToolOk(res, "PivotTable")) return;
                MarkDirty(); RebuildParts(); SelectPart(int.Parse(a[0])); canvas.RepaintAll();
            }
        }
        void RemoveDuplicates()
        {
            DialogResult h = MessageBox.Show(this, "Dữ liệu có hàng tiêu đề không?\n\nCó: giữ nguyên hàng đầu tiên.  Không: xét cả hàng đầu.", "Loại bỏ trùng lặp", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (h == DialogResult.Cancel) return;
            string r = Tool("RemoveDuplicates", "header", h == DialogResult.Yes ? "1" : "0"); if (!ToolOk(r, "Loại bỏ trùng lặp")) return;
            string[] p = r.Split('\t'); MarkDirty();
            MessageBox.Show(this, p.Length >= 3 ? "Đã xóa " + p[1] + " hàng trùng lặp; còn lại " + p[2] + " hàng." : "Xong.", "Loại bỏ trùng lặp", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        void SplitColumns(string delim)
        {
            string r = Tool("TextToColumns", "delim", delim); if (!ToolOk(r, "Văn bản thành cột")) return;
            string[] p = r.Split('\t'); if (p.Length > 1 && p[1] == "0") MessageBox.Show(this, "Không có ô nào chứa ký tự phân tách trong cột đầu của vùng chọn.", "Văn bản thành cột"); else MarkDirty();
        }
        void ValidationRule(string kind)
        {
            string v1 = null, v2 = null, msg = "";
            if (kind == "list")
            {
                string items = Ask(this, "Xác thực dữ liệu", "Các giá trị của danh sách (cách nhau bằng dấu phẩy):", ""); if (string.IsNullOrEmpty(items)) return;
                List<string> q = new List<string>(); foreach (string it in items.Split(',')) if (it.Trim().Length > 0) q.Add("\"" + it.Trim().Replace("\"", "\"\"") + "\"");
                v1 = string.Join(";", q.ToArray()); msg = "Chỉ được chọn: " + items;
            }
            else if (kind != "clear")
            {
                string rg = Ask(this, "Xác thực dữ liệu", "Khoảng cho phép (nhỏ nhất – lớn nhất), ví dụ: 0 - 100", "0 - 100"); if (string.IsNullOrEmpty(rg)) return;
                string[] ab = rg.Split(new[] { '-', '–', ';' }, StringSplitOptions.RemoveEmptyEntries); if (ab.Length < 2) return;
                v1 = FormulaValue(ab[0]); v2 = FormulaValue(ab[1]); msg = "Giá trị phải nằm trong khoảng " + ab[0].Trim() + " – " + ab[1].Trim() + ".";
            }
            if (ToolOk(Tool("Validation", "kind", kind, "v1", v1, "v2", v2, "msg", msg), "Xác thực dữ liệu")) MarkDirty();
        }
        void InsertToc(int levels) { if (ToolOk(Tool("InsertTOC", "levels", levels.ToString(), "title", "Mục lục"), "Mục lục")) MarkDirty(); }
        void InsertCaption(string cat)
        {
            string t = Ask(this, "Chèn chú thích", "Nội dung chú thích (" + cat + " n: …):", ""); if (t == null) return;
            if (ToolOk(Tool("InsertCaption", "cat", cat, "text", t), "Chèn chú thích")) MarkDirty();
        }

        // ---------------- equations ----------------
        static readonly string[][] Equations = {
            new[] { "Định lý Pythagore", "a^2 + b^2 = c^2" },
            new[] { "Công thức nghiệm bậc hai", "x = {-b +- sqrt{b^2 - 4ac}} over {2a}" },
            new[] { "Diện tích hình tròn", "A = %pi r^2" },
            new[] { "Khai triển nhị thức", "(x + a)^n = sum from{k=0} to{n} binom{n}{k} x^k a^{n-k}" },
            new[] { "Tích phân xác định", "int from{a} to{b} f(x) dx = F(b) - F(a)" },
            new[] { "Chuỗi Taylor của e^x", "e^x = sum from{n=0} to{infinity} {x^n} over {n!}" },
            new[] { "Giới hạn", "lim from{x -> 0} {sin x} over x = 1" },
            new[] { "Ma trận 2×2", "left( matrix{ a # b ## c # d } right)" } };
        List<RMenu> EquationMenu()
        {
            List<RMenu> m = new List<RMenu> { Mi("Phương trình mới…", "equation", () => { string f = Ask(this, "Phương trình", "Cú pháp LibreOffice Math, ví dụ:  a over b   sqrt{x}   x^2   sum from{i=1} to{n} i   %alpha", "x = {-b +- sqrt{b^2 - 4ac}} over {2a}"); if (!string.IsNullOrEmpty(f)) InsertEquation(f); }), RMenu.Sep() };
            foreach (string[] e in Equations) { string f = e[1]; m.Add(Mi(e[0], null, () => InsertEquation(f))); }
            return m;
        }
        void InsertEquation(string f) { if (ToolOk(Tool("InsertFormula", "formula", f), "Phương trình")) MarkDirty(); }

        // ---------------- AI assistant (web services, no API key needed; API keys: web editor) ----------------
        static readonly string[][] AiServices = {
            new[] { "ChatGPT", "https://chatgpt.com/" }, new[] { "Gemini", "https://gemini.google.com/app" }, new[] { "Microsoft Copilot", "https://copilot.microsoft.com/" },
            new[] { "Claude", "https://claude.ai/new" }, new[] { "DeepSeek", "https://chat.deepseek.com/" }, new[] { "Grok", "https://grok.com/" } };
        List<RMenu> AiMenu()
        {
            List<RMenu> m = new List<RMenu> { new RMenu("— Hỏi về phần đang chọn —", null, null) };
            foreach (string[] p in new[] {
                new[] { "Hỏi A.I…", "" }, new[] { "Tóm tắt", "Tóm tắt ngắn gọn nội dung sau bằng tiếng Việt:" },
                new[] { "Dịch sang tiếng Anh", "Dịch nội dung sau sang tiếng Anh:" }, new[] { "Dịch sang tiếng Việt", "Dịch nội dung sau sang tiếng Việt:" },
                new[] { "Sửa chính tả & ngữ pháp", "Sửa lỗi chính tả và ngữ pháp, giữ nguyên ý, trả về văn bản đã sửa:" },
                new[] { "Viết lại trang trọng hơn", "Viết lại nội dung sau theo văn phong trang trọng, chuyên nghiệp:" },
                new[] { type == LokDoc.SPREADSHEET ? "Phân tích dữ liệu" : "Gợi ý ý tưởng", type == LokDoc.SPREADSHEET ? "Phân tích bảng dữ liệu sau (cột phân tách bằng Tab), nêu xu hướng và gợi ý công thức Excel phù hợp:" : "Gợi ý thêm ý tưởng và cách phát triển nội dung sau:" } })
            { string pr = p[1]; m.Add(Mi(p[0], null, () => AskAi(pr))); }
            m.Add(new RMenu("— Dịch vụ A.I —", null, null));
            foreach (string[] s in AiServices) { string n = s[0]; m.Add(new RMenu(n, null, () => NativeSettings.Set("ai.service", n)) { Checked = () => NativeSettings.Get("ai.service", "ChatGPT") == n }); }
            m.Add(RMenu.Sep());
            m.Add(Mi("Mở khung Trợ lý A.I (web không đăng nhập / đăng nhập / khóa API)…", "ai", () => AskAi(null)));
            m.Add(Mi("Trợ lý A.I của Giao diện Web", "addins", OpenInWeb));
            return m;
        }
        void AskAi(string prompt)
        {
            string sel = null; try { Lok.Activate(doc); sel = doc.GetTextSelection("text/plain;charset=utf-8"); } catch { }
            sel = (sel ?? "").Trim();
            // v0.1.7: native A.I window - the question goes to the website (no login / signed in) or the API and the answer comes back here
            if (prompt == null || prompt == "" || sel.Length > 0)
            {
                AiPane pane = new AiPane(sel, string.IsNullOrEmpty(prompt) ? null : prompt, InsertAiText);
                pane.Show(this); return;
            }
            if (sel.Length == 0) { MessageBox.Show(this, "Hãy chọn đoạn văn bản" + (type == LokDoc.SPREADSHEET ? " hoặc vùng ô" : "") + " cần xử lý trước.", "Trợ lý A.I", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            string text = prompt + (sel.Length > 0 ? "\n\n" + (sel.Length > 30000 ? sel.Substring(0, 30000) : sel) : "");
            try { Clipboard.SetText(text); } catch { }
            string svc = NativeSettings.Get("ai.service", "ChatGPT"), url = AiServices[0][1];
            foreach (string[] s in AiServices) if (s[0] == svc) url = s[1];
            Api.AiWeb(url, "Trợ lý A.I");
            status.Segs[0].Text = "Đã chép yêu cầu — bấm Ctrl+V trong cửa sổ " + svc + " rồi Enter"; status.Invalidate();
        }

        void InsertAiText(string text)
        {
            try
            {
                if (IsDisposed || doc == null) return;
                Lok.Activate(doc);
                if (!doc.Paste("text/plain;charset=utf-8", Encoding.UTF8.GetBytes(text.Replace("\r\n", "\n")))) canvas.TypeText(text);
                MarkDirty(); Activate(); canvas.Focus();
            }
            catch (Exception ex) { Lok.Trace("ai insert " + ex.Message); }
        }

        // ---------------- small dialogs ----------------
        static string Ask(IWin32Window owner, string title, string label, string def)
        {
            using (Form f = new Form { Text = title, Width = 460, Height = 158, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false, Font = Ui.F(9f), BackColor = Color.White })
            {
                Label l = new Label { Text = label, Left = 14, Top = 14, Width = 420 };
                TextBox t = new TextBox { Left = 14, Top = 38, Width = 416, Text = def ?? "" };
                Button ok = new Button { Text = "OK", Left = 268, Top = 76, Width = 80, DialogResult = DialogResult.OK };
                Button no = new Button { Text = "Hủy", Left = 352, Top = 76, Width = 80, DialogResult = DialogResult.Cancel };
                f.Controls.AddRange(new Control[] { l, t, ok, no }); f.AcceptButton = ok; f.CancelButton = no;
                return f.ShowDialog(owner) == DialogResult.OK ? t.Text : null;
            }
        }
    }

    class DoubleBufferedPanel : Panel { public DoubleBufferedPanel() { DoubleBuffered = true; SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true); } }
}
