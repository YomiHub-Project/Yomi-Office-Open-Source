// Office Open Source v0.1.6 - dialogs of the native editors: Options, Find & Replace, Hyperlink, Zoom, Paragraph, Page Setup,
// table-size picker and symbol picker (layouts follow the Office reference screenshots)
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace OfficeOpenSource
{
    class OfficeDialog : Form
    {
        public OfficeDialog(string title, int w, int h)
        {
            Text = title; ClientSize = new Size(w, h); FormBorderStyle = FormBorderStyle.FixedDialog; StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false; MaximizeBox = false; ShowInTaskbar = false; Font = Ui.F(9f); BackColor = Color.White;
        }
        protected Label L(string t, int x, int y, int w) { Label l = new Label { Text = t, Location = new Point(x, y + 3), Size = new Size(w, 18), BackColor = Color.Transparent }; Controls.Add(l); return l; }
        protected Label H(string t, int x, int y, int w) { Label l = new Label { Text = t, Location = new Point(x, y), Size = new Size(w, 20), Font = Ui.F(9f, FontStyle.Bold), BackColor = Color.Transparent }; Controls.Add(l); return l; }
        protected Button OkCancel(int y)
        {
            Button ok = new Button { Text = "OK", Location = new Point(ClientSize.Width - 186, y), Size = new Size(84, 26), DialogResult = DialogResult.OK };
            Button no = new Button { Text = "Hủy", Location = new Point(ClientSize.Width - 96, y), Size = new Size(84, 26), DialogResult = DialogResult.Cancel };
            Controls.Add(ok); Controls.Add(no); AcceptButton = ok; CancelButton = no; return ok;
        }
    }

    // v0.1.7: OptionsDialog moved to UnifiedSettings.cs (rendered from ui_settings.json, shared with the Web editor)

    /// <summary>Tìm và Thay thế (modeless, tabs Tìm / Thay thế).</summary>
    class FindForm : OfficeDialog
    {
        public event Action<string, bool, bool> FindNext; public event Action<string, string, bool, bool> Replace;
        readonly TabControl tabs = new TabControl { Location = new Point(10, 10), Size = new Size(470, 190) };
        readonly TextBox find1 = new TextBox(), find2 = new TextBox(), repl = new TextBox(); readonly CheckBox mc = new CheckBox { Text = "Khớp chữ hoa/thường", AutoSize = true };
        public FindForm() : base("Tìm và Thay thế", 490, 240)
        {
            TabPage p1 = new TabPage("Tìm"), p2 = new TabPage("Thay thế"); tabs.TabPages.Add(p1); tabs.TabPages.Add(p2);
            p1.Controls.Add(new Label { Text = "Tìm gì:", Location = new Point(10, 18), AutoSize = true }); find1.SetBounds(110, 14, 330, 24); p1.Controls.Add(find1);
            p2.Controls.Add(new Label { Text = "Tìm gì:", Location = new Point(10, 18), AutoSize = true }); find2.SetBounds(110, 14, 330, 24); p2.Controls.Add(find2);
            p2.Controls.Add(new Label { Text = "Thay bằng:", Location = new Point(10, 52), AutoSize = true }); repl.SetBounds(110, 48, 330, 24); p2.Controls.Add(repl);
            mc.Location = new Point(14, 210); Controls.Add(tabs); Controls.Add(mc);
            Button next = new Button { Text = "Tìm tiếp", Location = new Point(180, 204), Size = new Size(90, 26) };
            Button rep = new Button { Text = "Thay thế", Location = new Point(276, 204), Size = new Size(90, 26) };
            Button all = new Button { Text = "Thay tất cả", Location = new Point(372, 204), Size = new Size(100, 26) };
            Controls.Add(next); Controls.Add(rep); Controls.Add(all); AcceptButton = next;
            find1.TextChanged += (s, e) => { if (find2.Text != find1.Text) find2.Text = find1.Text; }; find2.TextChanged += (s, e) => { if (find1.Text != find2.Text) find1.Text = find2.Text; };
            next.Click += (s, e) => { if (find1.Text.Length > 0 && FindNext != null) FindNext(find1.Text, mc.Checked, false); };
            rep.Click += (s, e) => { if (find1.Text.Length > 0 && Replace != null) Replace(find1.Text, repl.Text, mc.Checked, false); };
            all.Click += (s, e) => { if (find1.Text.Length > 0 && Replace != null) Replace(find1.Text, repl.Text, mc.Checked, true); };
            tabs.SelectedIndexChanged += (s, e) => { rep.Visible = all.Visible = tabs.SelectedIndex == 1; };
            FormClosing += (s, e) => { if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); } };
            ShowInTaskbar = false; TopMost = false; rep.Visible = all.Visible = false;
        }
        public void Show(IWin32Window owner, bool replace, string text)
        {
            if (text != null) find1.Text = text;
            tabs.SelectedIndex = replace ? 1 : 0;
            if (!Visible) { Owner = owner as Form; StartPosition = FormStartPosition.Manual; Form o = owner as Form; if (o != null) Location = new Point(o.Right - Width - 40, o.Top + 180); Show(); }
            Activate(); (replace ? find2 : find1).Focus();
        }
    }

    class LinkDialog : OfficeDialog
    {
        readonly TextBox text = new TextBox(), url = new TextBox();
        public string TextValue { get { return text.Text.Length > 0 ? text.Text : url.Text; } }
        public string Url { get { string u = url.Text.Trim(); if (u.Length > 0 && !u.Contains(":") && !u.StartsWith("#")) u = (u.Contains("@") ? "mailto:" : "https://") + u; return u; } }
        public LinkDialog(string sel) : base("Chèn siêu liên kết", 480, 150)
        {
            L("Văn bản hiển thị:", 14, 16, 120); text.SetBounds(140, 14, 320, 24); text.Text = sel != null && !sel.Contains("\n") ? sel : ""; Controls.Add(text);
            L("Địa chỉ:", 14, 52, 120); url.SetBounds(140, 50, 320, 24); url.Text = "https://"; Controls.Add(url);
            OkCancel(108);
        }
    }

    class ZoomForm : OfficeDialog
    {
        public float Value; public int Fit;
        public ZoomForm(float cur) : base("Thu phóng", 330, 250)
        {
            Value = cur; H("Thu phóng tới", 14, 10, 200);
            int[] p = { 200, 100, 75 }; int y = 36;
            foreach (int v in p) { int vv = v; RadioButton r = new RadioButton { Text = v + "%", Location = new Point(20, y), AutoSize = true, Checked = Math.Abs(cur * 100 - v) < 1 }; r.CheckedChanged += (s, e) => { if (r.Checked) { Value = vv / 100f; Fit = 0; } }; Controls.Add(r); y += 26; }
            RadioButton w = new RadioButton { Text = "Chiều rộng trang", Location = new Point(160, 36), AutoSize = true }; w.CheckedChanged += (s, e) => { if (w.Checked) Fit = 1; }; Controls.Add(w);
            RadioButton f = new RadioButton { Text = "Toàn trang", Location = new Point(160, 62), AutoSize = true }; f.CheckedChanged += (s, e) => { if (f.Checked) Fit = 2; }; Controls.Add(f);
            L("Phần trăm:", 20, y + 10, 80); NumericUpDown n = new NumericUpDown { Location = new Point(110, y + 10), Width = 70, Minimum = 10, Maximum = 500, Value = (decimal)Math.Max(10, Math.Min(500, Math.Round(cur * 100))) };
            n.ValueChanged += (s, e) => { Value = (float)n.Value / 100f; Fit = 0; }; Controls.Add(n);
            OkCancel(210);
        }
    }

    class ParaForm : OfficeDialog
    {
        public int Align = -1, LineSpacing = -1, IndentSteps, SpaceSteps;
        public ParaForm(bool word) : base("Đoạn văn", 420, 300)
        {
            H("Chung", 14, 10, 300);
            L("Căn chỉnh:", 20, 38, 110); ComboBox a = new ComboBox { Location = new Point(140, 38), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList }; a.Items.AddRange(new object[] { "(Giữ nguyên)", "Trái", "Giữa", "Phải", "Đều hai bên" }); a.SelectedIndex = 0; Controls.Add(a);
            H("Thụt lề", 14, 76, 300);
            L("Thụt lề (bước):", 20, 104, 110); NumericUpDown ind = new NumericUpDown { Location = new Point(140, 104), Width = 70, Minimum = -10, Maximum = 10 }; Controls.Add(ind);
            L("(+ tăng, − giảm)", 220, 104, 150).ForeColor = Ui.Sub;
            H("Giãn cách", 14, 142, 300);
            L("Giãn dòng:", 20, 170, 110); ComboBox ls = new ComboBox { Location = new Point(140, 170), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList }; ls.Items.AddRange(new object[] { "(Giữ nguyên)", "Đơn (1,0)", "1,5 dòng", "Kép (2,0)" }); ls.SelectedIndex = 0; Controls.Add(ls);
            L("Khoảng cách đoạn:", 20, 206, 120); NumericUpDown sp = new NumericUpDown { Location = new Point(140, 206), Width = 70, Minimum = -10, Maximum = 10 }; Controls.Add(sp);
            L("(+ tăng, − giảm)", 220, 206, 150).ForeColor = Ui.Sub;
            Button ok = OkCancel(258);
            ok.Click += (s, e) => { Align = a.SelectedIndex - 1; LineSpacing = ls.SelectedIndex - 1; IndentSteps = (int)ind.Value; SpaceSteps = (int)sp.Value; };
        }
    }

    /// <summary>Thiết lập trang: lề, hướng, khổ giấy (values returned in 1/100 mm).</summary>
    class PageSetupForm : OfficeDialog
    {
        public int MTop, MBottom, Left_, Right_, PaperW = 21000, PaperH = 29700; public bool Landscape;
        static readonly string[] Papers = { "A4|21000|29700", "A3|29700|42000", "A5|14800|21000", "Letter|21590|27940", "Legal|21590|35560", "B5|17600|25000" };
        public PageSetupForm(Size pageTw) : base("Thiết lập trang", 430, 380)
        {
            TabControl tabs = new TabControl { Location = new Point(10, 10), Size = new Size(410, 320) }; Controls.Add(tabs);
            TabPage m = new TabPage("Lề"), p = new TabPage("Giấy"); tabs.TabPages.Add(m); tabs.TabPages.Add(p);
            Landscape = pageTw.Width > pageTw.Height;
            Func<string, int, int, NumericUpDown> num = (label, x, yy) => { m.Controls.Add(new Label { Text = label, Location = new Point(x, yy + 3), AutoSize = true }); NumericUpDown n = new NumericUpDown { Location = new Point(x + 60, yy), Width = 80, DecimalPlaces = 2, Increment = 0.1m, Minimum = 0, Maximum = 20, Value = 2.0m }; m.Controls.Add(n); return n; };
            m.Controls.Add(new Label { Text = "Lề (cm)", Location = new Point(10, 10), AutoSize = true, Font = Ui.F(9f, FontStyle.Bold) });
            NumericUpDown t = num("Trên:", 14, 36), b = num("Dưới:", 200, 36), l = num("Trái:", 14, 70), r = num("Phải:", 200, 70);
            t.Value = 2.0m; b.Value = 2.0m; l.Value = 3.0m; r.Value = 1.5m;
            m.Controls.Add(new Label { Text = "Hướng", Location = new Point(10, 112), AutoSize = true, Font = Ui.F(9f, FontStyle.Bold) });
            RadioButton por = new RadioButton { Text = "Dọc", Location = new Point(20, 138), AutoSize = true, Checked = !Landscape }, lan = new RadioButton { Text = "Ngang", Location = new Point(120, 138), AutoSize = true, Checked = Landscape };
            m.Controls.Add(por); m.Controls.Add(lan);
            PictureBox prev = new PictureBox { Location = new Point(250, 120), Size = new Size(120, 150) }; m.Controls.Add(prev);
            prev.Paint += (s, e) => { bool land = lan.Checked; int w = land ? 100 : 72, h = land ? 72 : 100; Rectangle rc = new Rectangle((prev.Width - w) / 2, (prev.Height - h) / 2, w, h); e.Graphics.FillRectangle(Brushes.White, rc); e.Graphics.DrawRectangle(Pens.Gray, rc); for (int k = 0; k < 7; k++) e.Graphics.DrawLine(Pens.Silver, rc.X + 10, rc.Y + 12 + k * (h - 24) / 7, rc.Right - 10, rc.Y + 12 + k * (h - 24) / 7); };
            por.CheckedChanged += (s, e) => prev.Invalidate();
            p.Controls.Add(new Label { Text = "Khổ giấy:", Location = new Point(14, 18), AutoSize = true });
            ComboBox ps = new ComboBox { Location = new Point(100, 14), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList }; foreach (string x in Papers) ps.Items.Add(x.Split('|')[0]); p.Controls.Add(ps);
            int pw = Math.Min(pageTw.Width, pageTw.Height) * 1764 / 1000, ph = Math.Max(pageTw.Width, pageTw.Height) * 1764 / 1000; // twips -> 1/100 mm
            int best = 0; for (int i = 0; i < Papers.Length; i++) { string[] a = Papers[i].Split('|'); if (Math.Abs(int.Parse(a[1]) - pw) < 300 && Math.Abs(int.Parse(a[2]) - ph) < 300) best = i; }
            ps.SelectedIndex = best;
            Button ok = OkCancel(342);
            ok.Click += (s, e) =>
            {
                MTop = (int)(t.Value * 1000); MBottom = (int)(b.Value * 1000); Left_ = (int)(l.Value * 1000); Right_ = (int)(r.Value * 1000); Landscape = lan.Checked;
                string[] a = Papers[ps.SelectedIndex].Split('|'); PaperW = int.Parse(a[1]); PaperH = int.Parse(a[2]);
            };
        }
    }

    /// <summary>Table size picker (hover grid) like Office's Insert Table menu.</summary>
    class TablePicker : Control
    {
        int hr, hc; Action<int, int> pick; ToolStripDropDown host; const int C = 10, R = 8, S = 20;
        public static void Show(Control owner, Point at, Action<int, int> onPick)
        {
            TablePicker t = new TablePicker { pick = onPick };
            ToolStripControlHost h = new ToolStripControlHost(t) { AutoSize = false, Size = t.Size, Padding = Padding.Empty, Margin = Padding.Empty };
            ToolStripDropDown dd = new ToolStripDropDown { Padding = new Padding(1), Renderer = OfficeMenu.Renderer }; dd.Items.Add(h); t.host = dd; dd.Closed += (s, e) => dd.BeginInvoke(new Action(dd.Dispose)); dd.Show(owner, at);
        }
        TablePicker() { SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true); Size = new Size(C * S + 20, R * S + 66); BackColor = Color.White; Font = Ui.F(9f); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Color.White);
            Ui.Str(g, hr > 0 ? "Bảng " + hc + "×" + hr : "Chèn bảng", Ui.F(9f, FontStyle.Bold), Ui.Text, new Rectangle(10, 6, Width - 20, 20), TextFormatFlags.VerticalCenter);
            for (int r = 0; r < R; r++) for (int c = 0; c < C; c++)
                {
                    Rectangle rc = new Rectangle(10 + c * S, 32 + r * S, S - 3, S - 3); bool on = r < hr && c < hc;
                    using (SolidBrush b = new SolidBrush(on ? Color.FromArgb(0xFD, 0xE4, 0xCF) : Color.White)) g.FillRectangle(b, rc);
                    using (Pen p = new Pen(on ? Color.FromArgb(0xE8, 0x6B, 0x1E) : Color.FromArgb(0xC0, 0xC0, 0xC0))) g.DrawRectangle(p, rc);
                }
            Ui.Str(g, "Chèn bảng tùy chỉnh…", Font, Ui.Text, new Rectangle(10, Height - 28, Width - 20, 22), TextFormatFlags.VerticalCenter);
        }
        protected override void OnMouseMove(MouseEventArgs e) { int c = (e.X - 10) / S + 1, r = (e.Y - 32) / S + 1; if (e.Y < 32 || e.Y > 32 + R * S || c > C) { c = 0; r = 0; } if (c != hc || r != hr) { hc = c; hr = r; Invalidate(); } }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (e.Y > Height - 30)
            {
                host.Close();
                using (OfficeDialog d = new OfficeDialog("Chèn bảng", 260, 130))
                {
                    d.Controls.Add(new Label { Text = "Số cột:", Location = new Point(14, 18), AutoSize = true }); NumericUpDown c = new NumericUpDown { Location = new Point(110, 14), Width = 70, Minimum = 1, Maximum = 63, Value = 5 }; d.Controls.Add(c);
                    d.Controls.Add(new Label { Text = "Số hàng:", Location = new Point(14, 52), AutoSize = true }); NumericUpDown r = new NumericUpDown { Location = new Point(110, 48), Width = 70, Minimum = 1, Maximum = 500, Value = 2 }; d.Controls.Add(r);
                    Button ok = new Button { Text = "OK", Location = new Point(90, 90), Size = new Size(76, 26), DialogResult = DialogResult.OK }; Button no = new Button { Text = "Hủy", Location = new Point(172, 90), Size = new Size(76, 26), DialogResult = DialogResult.Cancel };
                    d.Controls.Add(ok); d.Controls.Add(no); d.AcceptButton = ok; d.CancelButton = no;
                    if (d.ShowDialog() == DialogResult.OK) pick((int)r.Value, (int)c.Value);
                }
                return;
            }
            if (hr > 0 && hc > 0) { host.Close(); pick(hr, hc); }
        }
    }

    /// <summary>Symbol picker: common Vietnamese/office symbols.</summary>
    class SymbolPopup : Control
    {
        const string Chars = "©®™§¶†‡•…–—‰€£¥₫¢°±×÷≈≠≤≥∞√∑∏∫πΩµαβγδθλσφψω←↑→↓↔⇒⇔∀∃∈∉∩∪⊂⊃∧∨¬★☆♠♣♥♦✓✔✗✘☐☑☒①②③④⑤⑥⑦⑧⑨⑩½¼¾²³¹";
        int hot = -1; Action<string> pick; ToolStripDropDown host; const int N = 12, S = 28;
        public static void Show(Control owner, Point at, Action<string> onPick)
        {
            SymbolPopup t = new SymbolPopup { pick = onPick };
            ToolStripControlHost h = new ToolStripControlHost(t) { AutoSize = false, Size = t.Size, Padding = Padding.Empty, Margin = Padding.Empty };
            ToolStripDropDown dd = new ToolStripDropDown { Padding = new Padding(1), Renderer = OfficeMenu.Renderer }; dd.Items.Add(h); t.host = dd; dd.Closed += (s, e) => dd.BeginInvoke(new Action(dd.Dispose)); dd.Show(owner, at);
        }
        SymbolPopup() { SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true); Size = new Size(N * S + 12, ((Chars.Length + N - 1) / N) * S + 12); BackColor = Color.White; Font = new Font("Segoe UI Symbol", 12f); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Color.White);
            for (int i = 0; i < Chars.Length; i++)
            {
                Rectangle r = new Rectangle(6 + (i % N) * S, 6 + (i / N) * S, S - 2, S - 2);
                if (i == hot) Ui.FillRound(g, Ui.Hover, r, 3);
                TextRenderer.DrawText(g, Chars[i].ToString(), Font, r, Ui.Text, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
        }
        int Hit(Point p) { int c = (p.X - 6) / S, r = (p.Y - 6) / S, i = r * N + c; return c >= 0 && c < N && i >= 0 && i < Chars.Length ? i : -1; }
        protected override void OnMouseMove(MouseEventArgs e) { int h = Hit(e.Location); if (h != hot) { hot = h; Invalidate(); } }
        protected override void OnMouseUp(MouseEventArgs e) { int h = Hit(e.Location); if (h < 0) return; host.Close(); pick(Chars[h].ToString()); }
    }

    /// <summary>Định dạng có điều kiện – một quy tắc (giá trị + kiểu tô sáng).</summary>
    class CondRuleForm : OfficeDialog
    {
        readonly TextBox t1 = new TextBox(), t2 = new TextBox(); readonly ComboBox st = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        static readonly string[][] Styles = { new[] { "red", "Tô đỏ nhạt, chữ đỏ sẫm" }, new[] { "yellow", "Tô vàng, chữ vàng sẫm" }, new[] { "green", "Tô xanh lục, chữ xanh lục sẫm" }, new[] { "lightred", "Tô đỏ nhạt" }, new[] { "redtext", "Chữ đỏ" }, new[] { "bold", "Chữ đậm" } };
        public string V1 { get { return t1.Text; } }
        public string V2 { get { return t2.Text; } }
        public string StyleKey { get { return Styles[Math.Max(0, st.SelectedIndex)][0]; } }
        public CondRuleForm(string title, string label, bool value, bool two, string def) : base(title, 460, 150)
        {
            int y = 14;
            if (value)
            {
                L(label, 14, y, 420); y += 26;
                t1.SetBounds(14, y, two ? 190 : 420, 24); t1.Text = def; Controls.Add(t1);
                if (two) { L("và", 212, y, 30); t2.SetBounds(244, y, 190, 24); Controls.Add(t2); }
                y += 34;
            }
            else { L("Định dạng các ô của vùng chọn với:", 14, y, 420); y += 30; }
            L("Kiểu:", 14, y, 50); st.SetBounds(70, y, 364, 24); foreach (string[] s in Styles) st.Items.Add(s[1]); st.SelectedIndex = 0; Controls.Add(st);
            y += 40; ClientSize = new Size(460, y + 36); OkCancel(y);
        }
    }

    /// <summary>Tạo PivotTable: chọn trường hàng / cột / giá trị và hàm tóm tắt.</summary>
    class PivotForm : OfficeDialog
    {
        readonly ComboBox row = Cb(), col = Cb(), val = Cb(), fn = Cb();
        static ComboBox Cb() { return new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList }; }
        static readonly string[][] Funcs = { new[] { "SUM", "Tổng" }, new[] { "COUNT", "Đếm" }, new[] { "AVERAGE", "Trung bình" }, new[] { "MAX", "Lớn nhất" }, new[] { "MIN", "Nhỏ nhất" } };
        public int Row { get { return row.SelectedIndex; } }
        public int Col { get { return col.SelectedIndex - 1; } }
        public int Data { get { return val.SelectedIndex; } }
        public string Func { get { return Funcs[fn.SelectedIndex][0]; } }
        public string FuncName { get { return Funcs[fn.SelectedIndex][1]; } }
        public PivotForm(string[] heads, int valueCol) : base("Tạo PivotTable", 420, 230)
        {
            H("Chọn các trường cho báo cáo (đặt trên trang tính mới)", 14, 12, 400);
            L("Hàng:", 14, 44, 110); row.SetBounds(130, 44, 270, 24);
            L("Cột (tùy chọn):", 14, 78, 110); col.SetBounds(130, 78, 270, 24);
            L("Giá trị:", 14, 112, 110); val.SetBounds(130, 112, 270, 24);
            L("Tóm tắt theo:", 14, 146, 110); fn.SetBounds(130, 146, 270, 24);
            col.Items.Add("(không)");
            foreach (string h in heads) { row.Items.Add(h); col.Items.Add(h); val.Items.Add(h); }
            foreach (string[] f in Funcs) fn.Items.Add(f[1]);
            row.SelectedIndex = 0; col.SelectedIndex = 0; val.SelectedIndex = valueCol >= 0 ? valueCol : heads.Length - 1; fn.SelectedIndex = 0;
            Controls.AddRange(new Control[] { row, col, val, fn });
            OkCancel(190);
        }
    }

    /// <summary>Tùy chọn → Cập nhật phần mềm for the native editor (port of src/core/update.js): lists GitHub Releases,
    /// Google Drive and MediaFire release folders, finds version.json or Yomi-Office-Open-Source-X.Y.Z-&lt;edition&gt;.* files,
    /// offers the installer that matches this copy (Setup / MSI / portable / appx, edition from app\edition.txt) and installs it through Updater.</summary>
    static class NativeUpdate
    {
        const string GitHubRepo = "YomiHub-Project/Yomi-Office-Open-Source";
        const string DriveId = "1kkqNrHFGQXLbf7PPYYc6HWgD1mFOdIyr", MediaFireKey = "seqly6bv55yr8";
        static readonly System.Text.RegularExpressions.Regex FileRe = new System.Text.RegularExpressions.Regex(@"^(?:OfficeOpenSource|Yomi-Office-Open-Source)-(\d+\.\d+\.\d+)(?:-(.+?))?\.(exe|msi|zip|apk|appxbundle|appx|msix|msixbundle)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        public class Item { public string Name, Url, Src, Ver, Ed, Ext; }
        public class Result { public string Latest, Notes, Error; public Item File; public List<Item> Alt = new List<Item>(); public bool Newer; }
        static bool busy, reached;

        static string Get(string url)
        {
            try { System.Net.ServicePointManager.SecurityProtocol |= (System.Net.SecurityProtocolType)3072; } catch { }
            System.Net.HttpWebRequest r = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
            r.Timeout = 20000; r.AllowAutoRedirect = true; r.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) YomiOfficeOpenSource/" + App.Version;
            using (System.Net.WebResponse resp = r.GetResponse()) using (System.IO.StreamReader sr = new System.IO.StreamReader(resp.GetResponseStream(), System.Text.Encoding.UTF8)) return sr.ReadToEnd();
        }
        static List<Item> ListGitHub(string repo, out string ghVer, out string ghNotes)
        {
            ghVer = null; ghNotes = null;
            List<Item> o = new List<Item>();
            try
            {
                System.Web.Script.Serialization.JavaScriptSerializer js = new System.Web.Script.Serialization.JavaScriptSerializer();
                Dictionary<string, object> j = js.Deserialize<Dictionary<string, object>>(Get("https://api.github.com/repos/" + repo + "/releases/latest"));
                reached = true;
                if (j == null || !j.ContainsKey("tag_name")) return o;
                string tag = Convert.ToString(j["tag_name"] ?? "").Trim();
                if (tag.StartsWith("v", StringComparison.OrdinalIgnoreCase)) tag = tag.Substring(1);
                ghVer = tag;
                if (j.ContainsKey("body")) ghNotes = Convert.ToString(j["body"] ?? "");
                object assetsObj;
                if (j.TryGetValue("assets", out assetsObj) && assetsObj is System.Collections.ArrayList)
                {
                    foreach (object x in (System.Collections.ArrayList)assetsObj)
                    {
                        Dictionary<string, object> a = x as Dictionary<string, object>;
                        if (a == null) continue;
                        string name = Convert.ToString(a.ContainsKey("name") ? a["name"] : "");
                        string url = Convert.ToString(a.ContainsKey("browser_download_url") ? a["browser_download_url"] : "");
                        if (name.Length > 0 && url.Length > 0) o.Add(new Item { Name = name, Src = "github", Url = url });
                    }
                }
            }
            catch { }
            return o;
        }
        static string Dec(string s) { return System.Net.WebUtility.HtmlDecode(s ?? ""); }
        static List<Item> ListDrive(string id, int depth)
        {
            List<Item> o = new List<Item>();
            try
            {
                string html = Get("https://drive.google.com/embeddedfolderview?id=" + Uri.EscapeDataString(id)); reached = true;
                foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(html, "<div class=\"flip-entry\" id=\"entry-([\\w-]+)\"[\\s\\S]*?<a href=\"([^\"]+)\"[\\s\\S]*?<div class=\"flip-entry-title\">([^<]*)</div>"))
                {
                    string fid = m.Groups[1].Value, name = Dec(m.Groups[3].Value).Trim();
                    if (m.Groups[2].Value.Contains("/folders/")) { if (depth < 2) o.AddRange(ListDrive(fid, depth + 1)); }
                    else o.Add(new Item { Name = name, Src = "drive", Url = "https://drive.usercontent.google.com/download?id=" + fid + "&export=download&confirm=t" });
                }
            }
            catch { }
            return o;
        }
        static List<Item> ListMediaFire(string key, int depth)
        {
            List<Item> o = new List<Item>();
            try
            {
                System.Web.Script.Serialization.JavaScriptSerializer js = new System.Web.Script.Serialization.JavaScriptSerializer();
                string b = "https://www.mediafire.com/api/1.5/folder/get_content.php?response_format=json&folder_key=" + Uri.EscapeDataString(key);
                Dictionary<string, object> fc = Content(js.Deserialize<Dictionary<string, object>>(Get(b + "&content_type=files&chunk_size=1000"))); reached = true;
                if (fc != null && fc.ContainsKey("files"))
                    foreach (object x in (System.Collections.ArrayList)fc["files"])
                    {
                        Dictionary<string, object> f = (Dictionary<string, object>)x; string page = "https://www.mediafire.com/file/" + f["quickkey"];
                        if (f.ContainsKey("links") && f["links"] is Dictionary<string, object> && ((Dictionary<string, object>)f["links"]).ContainsKey("normal_download")) page = (string)((Dictionary<string, object>)f["links"])["normal_download"];
                        o.Add(new Item { Name = (string)f["filename"], Src = "mediafire", Url = page });
                    }
                if (depth < 2)
                {
                    Dictionary<string, object> dc = Content(js.Deserialize<Dictionary<string, object>>(Get(b + "&content_type=folders")));
                    if (dc != null && dc.ContainsKey("folders")) foreach (object x in (System.Collections.ArrayList)dc["folders"]) o.AddRange(ListMediaFire((string)((Dictionary<string, object>)x)["folderkey"], depth + 1));
                }
            }
            catch { }
            return o;
        }
        static Dictionary<string, object> Content(Dictionary<string, object> j)
        {
            object r, c; if (j != null && j.TryGetValue("response", out r) && r is Dictionary<string, object> && ((Dictionary<string, object>)r).TryGetValue("folder_content", out c)) return c as Dictionary<string, object>; return null;
        }
        public static int Cmp(string a, string b)
        {
            string[] x = (a ?? "0").Split('.'), y = (b ?? "0").Split('.');
            for (int i = 0; i < Math.Max(x.Length, y.Length); i++) { int p = 0, q = 0; if (i < x.Length) int.TryParse(x[i], out p); if (i < y.Length) int.TryParse(y[i], out q); if (p != q) return p > q ? 1 : -1; }
            return 0;
        }
        static string Tag() { try { return System.IO.File.ReadAllText(System.IO.Path.Combine(App.Dir, "app\\edition.txt")).Trim(); } catch { return ""; } }
        static bool Portable() { return System.IO.File.Exists(System.IO.Path.Combine(App.Dir, "portable.flag")); }
        static bool Appx() { return App.Dir.IndexOf("\\WindowsApps\\", StringComparison.OrdinalIgnoreCase) >= 0; }
        static bool Msi() { return System.IO.File.Exists(System.IO.Path.Combine(App.Dir, "installed_by_msi.txt")); }
        static string Download(Item it) { return (it.Src == "drive" || it.Src == "github") ? Get(it.Url) : Get(MediaFireDirect(Get(it.Url))); }
        static string MediaFireDirect(string html)
        {
            System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(html, "href=\"(https?://download[^\"]+mediafire\\.com/[^\"]+)\"");
            if (!m.Success) m = System.Text.RegularExpressions.Regex.Match(html, "\"(https?://download\\d*\\.mediafire\\.com/[^\"]+)\"");
            if (!m.Success) throw new Exception("MediaFire: không tìm thấy liên kết tải"); return Dec(m.Groups[1].Value);
        }

        public static Result Check()
        {
            Result res = new Result(); List<Item> files = new List<Item>(); reached = false;
            string ghVer, ghNotes;
            files.AddRange(ListGitHub(GitHubRepo, out ghVer, out ghNotes));
            files.AddRange(ListDrive(DriveId, 0)); files.AddRange(ListMediaFire(MediaFireKey, 0));
            NativeSettings.Set("update.last", DateTime.Now.ToString("o"));
            if (files.Count == 0 && string.IsNullOrEmpty(ghVer)) { if (!reached) res.Error = "Không đọc được nguồn cập nhật trên GitHub / Google Drive / MediaFire (kiểm tra kết nối mạng)."; return res; }
            Dictionary<string, object> manifest = null; System.Web.Script.Serialization.JavaScriptSerializer js = new System.Web.Script.Serialization.JavaScriptSerializer();
            foreach (Item vf in files) if (vf.Name.Equals("version.json", StringComparison.OrdinalIgnoreCase))
                    try { Dictionary<string, object> m = js.Deserialize<Dictionary<string, object>>(Download(vf)); if (m != null && m.ContainsKey("version") && (manifest == null || Cmp((string)m["version"], (string)manifest["version"]) > 0)) manifest = m; } catch { }
            foreach (Item f in files) { System.Text.RegularExpressions.Match m = FileRe.Match(f.Name); if (m.Success) { f.Ver = m.Groups[1].Value; f.Ed = m.Groups[2].Value; f.Ext = m.Groups[3].Value.ToLowerInvariant(); } }
            if (manifest != null) { res.Latest = (string)manifest["version"]; object n; if (manifest.TryGetValue("notes", out n)) res.Notes = n as string; }
            else foreach (Item f in files) if (f.Ver != null && (res.Latest == null || Cmp(f.Ver, res.Latest) > 0)) res.Latest = f.Ver;
            if (!string.IsNullOrEmpty(ghVer) && (res.Latest == null || Cmp(ghVer, res.Latest) >= 0))
            {
                res.Latest = ghVer;
                if (!string.IsNullOrEmpty(ghNotes)) res.Notes = ghNotes;
            }
            res.Newer = res.Latest != null && Cmp(res.Latest, App.Version) > 0;
            if (!res.Newer) return res;
            string tag = Tag(); bool port = Portable(), appx = Appx(), msi = Msi();
            // 1) version.json "files" mapping
            object mf; if (manifest != null && manifest.TryGetValue("files", out mf) && mf is Dictionary<string, object> && (string)manifest["version"] == res.Latest)
                foreach (string k in new[] { tag + (port ? "-portable" : appx ? "-appx" : msi ? "-msi" : "-setup"), tag + (port ? "-portable" : appx ? "-appx" : "-setup"), tag, appx ? "appxbundle" : null, port ? "portable" : msi ? "msi" : "setup" })
                {
                    object fn; if (k == null || !((Dictionary<string, object>)mf).TryGetValue(k, out fn)) continue;
                    foreach (Item f in files) if (f.Name == (string)fn) { if (res.File == null) res.File = f; else if (f.Src != res.File.Src) res.Alt.Add(f); }
                    if (res.File != null) return res;
                }
            // 2) by file name
            List<Item> ok = files.FindAll(f => f.Ver == res.Latest && (appx ? (f.Ext.StartsWith("appx") || f.Ext.StartsWith("msix")) : port ? f.Ext == "zip" : msi ? f.Ext == "msi" : f.Ext == "exe"));
            if (msi && ok.Count == 0) ok = files.FindAll(f => f.Ver == res.Latest && f.Ext == "exe");
            res.File = ok.Find(f => tag.Length > 0 && f.Ed.Equals(appx || f.Ext == "msi" ? tag : tag + (port ? "-portable" : "-setup"), StringComparison.OrdinalIgnoreCase))
                ?? ok.Find(f => tag.Length > 0 && f.Ed.StartsWith(tag, StringComparison.OrdinalIgnoreCase))
                ?? ok.Find(f => f.Ed == "" || f.Ed.Equals("Setup", StringComparison.OrdinalIgnoreCase) || f.Ed.Equals("portable", StringComparison.OrdinalIgnoreCase))
                ?? (ok.Count > 0 ? ok[0] : null);
            if (res.File != null) foreach (Item f in ok) if (f != res.File && f.Name == res.File.Name && f.Src != res.File.Src) res.Alt.Add(f);
            return res;
        }

        public static string LastText() { DateTime t; return DateTime.TryParse(NativeSettings.Get("update.last", ""), out t) ? t.ToString("dd/MM/yyyy HH:mm") : "chưa kiểm tra"; }
        public static void AutoCheck(Form owner)
        {
            if (!NativeSettings.Bool("update.auto", true)) return;
            DateTime t; if (DateTime.TryParse(NativeSettings.Get("update.last", ""), out t) && (DateTime.Now - t).TotalHours < 24) return;
            CheckAsync(owner, false);
        }
        public static void CheckAsync(Form owner, bool manual)
        {
            if (busy || owner == null) return; busy = true;
            System.Threading.Thread th = new System.Threading.Thread(() =>
            {
                Result r; try { r = Check(); } catch (Exception ex) { r = new Result { Error = ex.Message }; }
                busy = false;
                try { if (!owner.IsDisposed) owner.BeginInvoke((Action)(() => Show(owner, r, manual))); } catch { }
            }) { IsBackground = true };
            th.Start();
        }
        static void Show(Form owner, Result r, bool manual)
        {
            if (!r.Newer)
            {
                if (manual) MessageBox.Show(owner, r.Error ?? ("Bạn đang dùng phiên bản mới nhất: " + App.Version + "." + (r.Latest != null ? "\nPhiên bản trên máy chủ: " + r.Latest : "\nChưa có bản phát hành nào trong thư mục cập nhật.")), "Cập nhật phần mềm", MessageBoxButtons.OK, r.Error != null ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                return;
            }
            if (!manual && NativeSettings.Get("update.skip", "") == r.Latest) return;
            if (r.File == null) { MessageBox.Show(owner, "Có phiên bản mới " + r.Latest + " nhưng chưa có bộ cài phù hợp với bản này (" + Tag() + ").", "Cập nhật phần mềm", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            DialogResult a = MessageBox.Show(owner, "Có phiên bản mới: Yomi-Office Open Source " + r.Latest + " (đang dùng " + App.Version + ").\n" + (string.IsNullOrEmpty(r.Notes) ? "" : "\n" + r.Notes + "\n") +
                "\nBộ cài: " + r.File.Name + " (" + (r.File.Src == "github" ? "GitHub" : r.File.Src == "drive" ? "Google Drive" : "MediaFire") + ")\n\nCó: tải và cài đặt ngay (ứng dụng sẽ khởi động lại)\nKhông: để sau\nHủy: bỏ qua phiên bản này",
                "Cập nhật phần mềm", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
            if (a == DialogResult.Cancel) { NativeSettings.Set("update.skip", r.Latest); return; }
            if (a != DialogResult.Yes) return;
            Install(owner, r);
        }
        static void Install(Form owner, Result r)
        {
            string mode = Portable() ? "portable" : Appx() ? "appx" : "";
            List<Item> tries = new List<Item> { r.File }; tries.AddRange(r.Alt);
            using (Form f = new Form { Text = "Cập nhật phần mềm", ClientSize = new Size(440, 96), FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false, Font = Ui.F(9f), BackColor = Color.White })
            {
                Label l = new Label { Location = new Point(14, 14), Size = new Size(412, 20), Text = "Đang tải " + r.File.Name + "…" };
                ProgressBar pb = new ProgressBar { Location = new Point(14, 44), Size = new Size(412, 18) };
                f.Controls.Add(l); f.Controls.Add(pb);
                int idx = 0; bool started = false;
                Timer t = new Timer { Interval = 400 };
                Action start = () => { Item it = tries[idx]; string url = it.Url; try { if (it.Src == "mediafire") url = MediaFireDirect(Get(it.Url)); } catch { } Updater.Start(url, it.Name, mode); started = true; };
                t.Tick += (s, e) =>
                {
                    if (!started) { start(); return; }
                    Dictionary<string, object> st = App.Json.Deserialize<Dictionary<string, object>>(Updater.Status());
                    string state = Convert.ToString(st["state"]); long got = Convert.ToInt64(st["got"]), total = Convert.ToInt64(st["total"]);
                    if (state == "downloading") { if (total > 0) { pb.Value = (int)Math.Min(100, 100 * got / total); l.Text = "Đang tải " + tries[idx].Name + ": " + pb.Value + "%"; } }
                    else if (state == "installing") { pb.Value = 100; l.Text = "Đang cài đặt — ứng dụng sẽ khởi động lại…"; }
                    else if (state == "done") { t.Stop(); l.Text = mode == "appx" ? "Đã mở gói cài đặt." : "Đang cài đặt — ứng dụng sẽ khởi động lại."; }
                    else if (state == "error") { if (++idx < tries.Count) { started = false; return; } t.Stop(); l.Text = "Lỗi: " + Convert.ToString(st["error"]); }
                };
                f.Shown += (s, e) => t.Start(); f.FormClosed += (s, e) => t.Dispose();
                f.ShowDialog(owner);
            }
        }
    }
}
