// Office Open Source v0.1.6 - File (Backstage) view: Home, New, Open, Info, Save As, Print, Share, Export, Account, Options
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.IO;
using System.Windows.Forms;

namespace OfficeOpenSource
{
    interface IBackstageHost
    {
        int DocType { get; }
        string DocPath { get; }
        string DocKindName { get; }
        Color Accent { get; }
        void CloseBackstage();
        void SaveDoc();
        bool SaveDocAs(string file, string ext);
        void NewDoc(int type);
        void OpenDoc(string file);
        void CloseDoc();
        void ExportPdf();
        void ShowOptions();
        void ShareLan();
        int PageCount { get; }
        Bitmap RenderPage(int index, int widthPx);
        Size PageSizeTwips(int index);
        void PrintDoc(PrinterSettings ps, string range, bool landscape);
        Dictionary<string, string> Properties();
        List<string> AutoSaveVersions();
    }

    class Backstage : Control
    {
        readonly IBackstageHost host;
        class Nav { public string Key, Text, Icon; public Rectangle R; public bool Bottom; }
        readonly List<Nav> nav = new List<Nav>();
        string page = "home"; object hot; Rectangle rBack;
        readonly Panel content = new Panel { AutoScroll = true, BackColor = Color.White };

        public Backstage(IBackstageHost h)
        {
            host = h;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Dock = DockStyle.Fill; BackColor = Color.White; Font = Ui.F(10f);
            // v0.1.7: navigation from ui_settings.json - identical in the Web editor
            bool bottom = false; int ni = 0;
            foreach (string[] it in UiSpec.Nav()) { if (it == null) { if (ni > 4) bottom = true; else nav.Add(new Nav { Key = "-" }); ni++; continue; } nav.Add(new Nav { Key = it[0], Text = it[1], Icon = it[2], Bottom = bottom }); ni++; }
            Controls.Add(content);
        }

        public void ShowPage(string p) { page = p; BuildPage(); Invalidate(); }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            content.Bounds = new Rectangle(214, 0, Math.Max(100, Width - 214), Height);
            int y = 56; rBack = new Rectangle(14, 12, 32, 32);
            foreach (Nav n in nav) { if (n.Bottom) continue; if (n.Key == "-") { y += 12; continue; } n.R = new Rectangle(8, y, 198, 34); y += 36; }
            int by = Height - 12 - 36 * 2; foreach (Nav n in nav) if (n.Bottom) { n.R = new Rectangle(8, by, 198, 34); by += 36; }
            if (content.Controls.Count == 0) BuildPage();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Color.White);
            using (SolidBrush b = new SolidBrush(Ui.Chrome)) g.FillRectangle(b, 0, 0, 214, Height);
            if (hot == (object)"back") Ui.FillRound(g, Ui.Hover, Rectangle.Inflate(rBack, 2, 2), 18);
            Icons.Draw(g, "back", rBack, true);
            foreach (Nav n in nav)
            {
                if (n.Key == "-") { int yy = nav[nav.IndexOf(n) - 1].R.Bottom + 6; using (Pen p = new Pen(Ui.Line)) g.DrawLine(p, 18, yy, 196, yy); continue; }
                bool sel = n.Key == page;
                if (sel) { Ui.FillRound(g, Color.FromArgb(0xE0, 0xE0, 0xE0), n.R, 4); Ui.FillRound(g, host.Accent, new Rectangle(n.R.X, n.R.Y + 8, 3, n.R.Height - 16), 1); }
                else if (hot == (object)n) Ui.FillRound(g, Ui.Hover, n.R, 4);
                Icons.Draw(g, n.Icon, new Rectangle(n.R.X + 14, n.R.Y + 9, 16, 16), true);
                Ui.Str(g, n.Text, sel ? Ui.F(10f, FontStyle.Bold) : Font, Ui.Text, new Rectangle(n.R.X + 40, n.R.Y, n.R.Width - 44, n.R.Height), TextFormatFlags.VerticalCenter);
            }
        }
        object Hit(Point p) { if (rBack.Contains(p)) return "back"; foreach (Nav n in nav) if (n.Key != "-" && n.R.Contains(p)) return n; return null; }
        protected override void OnMouseMove(MouseEventArgs e) { object h = Hit(e.Location); if (h != hot) { hot = h; Invalidate(); Cursor = h != null ? Cursors.Hand : Cursors.Default; } }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            object h = Hit(e.Location); if (h == null) return;
            if ((h as string) == "back") { host.CloseBackstage(); return; }
            Nav n = (Nav)h;
            switch (n.Key)
            {
                case "save": host.CloseBackstage(); host.SaveDoc(); return;
                case "close": host.CloseDoc(); return;
                case "options": host.ShowOptions(); return;
            }
            ShowPage(n.Key);
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys k) { if (k == Keys.Escape) { host.CloseBackstage(); return true; } return base.ProcessCmdKey(ref msg, k); }

        // ---------------- pages ----------------
        Label H1(string t, int y) { Label l = new Label { Text = t, Font = Ui.F(22f), ForeColor = Ui.Text, Location = new Point(40, y), AutoSize = true, BackColor = Color.White }; content.Controls.Add(l); return l; }
        Label H2(string t, int x, int y) { Label l = new Label { Text = t, Font = Ui.F(12f, FontStyle.Bold), ForeColor = Ui.Text, Location = new Point(x, y), AutoSize = true, BackColor = Color.White }; content.Controls.Add(l); return l; }
        Label P(string t, int x, int y, int w) { Label l = new Label { Text = t, Font = Ui.F(9.5f), ForeColor = Ui.Sub, Location = new Point(x, y), Size = new Size(w, 20), AutoEllipsis = true, BackColor = Color.White }; content.Controls.Add(l); return l; }
        Tile Big(string icon, string title, string sub, int x, int y, int w, Action a) { Tile t = new Tile { Icon = icon, Title = title, Sub = sub, Bounds = new Rectangle(x, y, w, sub == null ? 44 : 60), Click = a }; content.Controls.Add(t); return t; }

        // v0.1.7: tiles of a page come from ui_settings.json (same order, icons and text as the Web editor)
        void SpecTiles(string pg, int x, ref int y, int w, int step, Func<string, Action> act)
        {
            foreach (string[] t in UiSpec.Tiles(pg))
            {
                string title = t[2], sub = t[3];
                if (t[0] == "macros") title = "Macro: " + (Lok.MacrosAllowed ? "đang bật" : "đang tắt");
                Tile tl = Big(t[1], title, sub, x, y, w, act(t[0]) ?? (() => Web(t[0])));
                y += sub == null ? 56 : step;
            }
        }
        void Web(string id) { NativeDocForm f = host as NativeDocForm; if (f != null) f.WebHandoff(id); }
        static void WebAct(string id) { NativeDocForm.WebAct(id); }

        void BuildPage()
        {
            content.SuspendLayout(); foreach (Control c in content.Controls) c.Dispose(); content.Controls.Clear(); content.AutoScrollPosition = Point.Empty;
            try
            {
                switch (page)
                {
                    case "home": PageHome(true); break;
                    case "new": PageHome(false); break;
                    case "open": PageOpen(); break;
                    case "info": PageInfo(); break;
                    case "saveas": PageSaveAs(); break;
                    case "print": PagePrint(); break;
                    case "share": PageShare(); break;
                    case "export": PageExport(); break;
                    case "account": PageAccount(); break;
                }
            }
            catch (Exception ex) { P("Lỗi: " + ex.Message, 40, 100, 600); }
            content.ResumeLayout();
        }

        void PageHome(bool withRecent)
        {
            H1(withRecent ? Ui.Greeting() : "Mới", 26);
            int y = 82;
            H2("Mới", 40, y); y += 34;
            int[] types = { host.DocType, host.DocType == LokDoc.TEXT ? LokDoc.SPREADSHEET : LokDoc.TEXT, host.DocType == LokDoc.PRESENTATION ? LokDoc.SPREADSHEET : LokDoc.PRESENTATION };
            int x = 40;
            foreach (int t in types)
            {
                int tt = t; string name = t == LokDoc.TEXT ? "Tài liệu trống" : t == LokDoc.SPREADSHEET ? "Sổ làm việc trống" : "Bản trình bày trống";
                Template tp = new Template { Kind = t, Title = name, Bounds = new Rectangle(x, y, 150, 190), Click = () => host.NewDoc(tt) }; content.Controls.Add(tp); x += 166;
            }
            LinkLabel more = new LinkLabel { Text = "Thêm mẫu  →", Location = new Point(x + 10, y + 80), AutoSize = true, Font = Ui.F(9.5f), LinkColor = host.Accent, BackColor = Color.White };
            more.LinkClicked += (s, e) => { App.ActivateExistingWindow(); };
            content.Controls.Add(more);
            y += 214;
            if (!withRecent) return;
            RecentList(y, true);
        }

        void RecentList(int y, bool tabs)
        {
            H2("Gần đây", 40, y); y += 34;
            TextBox search = new TextBox { Location = new Point(40, y), Width = 320, Font = Ui.F(9.5f) };
            SetCue(search, "Tìm kiếm tệp");
            content.Controls.Add(search); y += 34;
            RecentGrid grid = new RecentGrid { Location = new Point(40, y), Size = new Size(Math.Max(500, content.Width - 90), 420), Accent = host.Accent };
            grid.Open += f => host.OpenDoc(f);
            grid.Load(NativeSettings.Pinned(), NativeSettings.Recent(), "");
            search.TextChanged += (s, e) => grid.Load(NativeSettings.Pinned(), NativeSettings.Recent(), search.Text);
            content.Controls.Add(grid);
        }

        [System.Runtime.InteropServices.DllImport("user32", CharSet = System.Runtime.InteropServices.CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, string l);
        static void SetCue(TextBox t, string cue) { t.HandleCreated += (s, e) => SendMessage(t.Handle, 0x1501, (IntPtr)1, cue); }

        void PageOpen()
        {
            H1("Mở", 26);
            int y = 90;
            Action browse = () =>
            {
                using (OpenFileDialog d = new OpenFileDialog { Filter = "Tài liệu Office|*.docx;*.docm;*.doc;*.dotx;*.odt;*.rtf;*.xlsx;*.xlsm;*.xls;*.xltx;*.ods;*.csv;*.pptx;*.pptm;*.ppt;*.potx;*.odp|Tất cả tệp|*.*" })
                    if (d.ShowDialog(FindForm()) == DialogResult.OK) host.OpenDoc(d.FileName);
            };
            Action pc = () =>
            {
                using (OpenFileDialog d = new OpenFileDialog { InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Filter = "Tài liệu Office|*.docx;*.docm;*.doc;*.odt;*.rtf;*.xlsx;*.xlsm;*.xls;*.ods;*.csv;*.pptx;*.pptm;*.ppt;*.odp|Tất cả tệp|*.*" })
                    if (d.ShowDialog(FindForm()) == DialogResult.OK) host.OpenDoc(d.FileName);
            };
            SpecTiles("open", 40, ref y, 300, 70, id => id == "browse" ? browse : id == "pc" ? pc : id == "cloudOpen" ? (Action)(() => WebAct("cloudOpen")) : null);
            y -= 70;
            RecentList(y + 90, false);
        }

        void PageInfo()
        {
            H1("Thông tin", 26);
            P(Path.GetFileName(host.DocPath), 40, 74, 700).Font = Ui.F(12f);
            P(Path.GetDirectoryName(host.DocPath), 40, 98, 700);
            int y = 140;
            List<string> vers = host.AutoSaveVersions();
            SpecTiles("info", 40, ref y, 420, 70, id =>
                id == "reveal" ? (Action)(() => { try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + host.DocPath + "\""); } catch { } }) :
                id == "protect" ? (Action)(() => ShowPage("saveas")) :
                id == "macros" ? (Action)(() => { Lok.MacrosAllowed = !Lok.MacrosAllowed; SharedSettings.Set(new Dictionary<string, object> { { "macrosAllowed", Lok.MacrosAllowed } }); MessageBox.Show("Đã " + (Lok.MacrosAllowed ? "bật" : "tắt") + " macro. Mở lại tài liệu để áp dụng.", "Macro"); ShowPage("info"); }) :
                id == "versions" ? (Action)(() => { if (vers.Count > 0) host.OpenDoc(vers[0]); else MessageBox.Show("Chưa có bản tự động lưu.", "Phiên bản"); }) : null);
            int px = 500; int py = 140;
            H2("Thuộc tính", px, py - 4); py += 34;
            foreach (KeyValuePair<string, string> kv in host.Properties())
            {
                P(kv.Key, px, py, 150).ForeColor = Ui.Sub; Label v = P(kv.Value, px + 150, py, 360); v.ForeColor = Ui.Text; py += 24;
            }
        }

        void PageSaveAs()
        {
            H1("Lưu thành", 26);
            string dir = Path.GetDirectoryName(host.DocPath); if (!Directory.Exists(dir)) dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            int y = 90;
            bool first = true;
            foreach (string[] t in UiSpec.Tiles("saveas"))
            {
                string id = t[0]; Action a = id == "recent" ? (Action)(() => Browse(dir)) : id == "pc" ? (Action)(() => Browse(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)))
                    : id == "addplace" ? (Action)(() => { using (FolderBrowserDialog f = new FolderBrowserDialog()) if (f.ShowDialog(FindForm()) == DialogResult.OK) Browse(f.SelectedPath); })
                    : id == "browse" ? (Action)(() => ClassicSaveAs(dir)) : (() => Web(id));
                Tile tl = Big(t[1], t[2], null, 40, y, 260, a); if (first) tl.Selected = true; first = false; y += 56;
            }
            BuildSaveAsRight(dir);
        }
        TextBox saveName; ComboBox saveType; Label savePath; ListView folderList; string saveDir;
        void Browse(string d) { if (Directory.Exists(d)) { saveDir = d; savePath.Text = "📁  " + d; FillFolder(); } }
        void BuildSaveAsRight(string dir)
        {
            int x = 330, y = 92, w = Math.Max(420, content.Width - x - 50);
            saveDir = dir;
            savePath = P("📁  " + dir, x, y, w); savePath.ForeColor = Ui.Text; savePath.Cursor = Cursors.Hand; savePath.Click += (s, e) => { DirectoryInfo up = Directory.GetParent(saveDir); if (up != null) Browse(up.FullName); };
            y += 30;
            saveName = new TextBox { Location = new Point(x, y), Width = w - 120, Font = Ui.F(10f), Text = Path.GetFileNameWithoutExtension(host.DocPath) }; content.Controls.Add(saveName);
            Button save = new Button { Text = "Lưu", Location = new Point(x + w - 110, y - 1), Size = new Size(110, 28), FlatStyle = FlatStyle.Flat, BackColor = host.Accent, ForeColor = Color.White, Font = Ui.F(9.5f, FontStyle.Bold) };
            save.FlatAppearance.BorderSize = 0; content.Controls.Add(save);
            y += 34;
            saveType = new ComboBox { Location = new Point(x, y), Width = w - 120, DropDownStyle = ComboBoxStyle.DropDownList, Font = Ui.F(9.5f) };
            foreach (KeyValuePair<string, string> f in Formats(host.DocType)) saveType.Items.Add(f.Value + " (*." + UiSpec.FileExt(f.Key) + ")");
            string cur = Path.GetExtension(host.DocPath).TrimStart('.').ToLowerInvariant(); int idx = 0; int k = 0;
            foreach (KeyValuePair<string, string> f in Formats(host.DocType)) { if (f.Key == cur) idx = k; k++; }
            saveType.SelectedIndex = idx; content.Controls.Add(saveType);
            y += 40;
            folderList = new ListView { Location = new Point(x, y), Size = new Size(w, Math.Max(200, content.Height - y - 30)), View = View.Details, FullRowSelect = true, BorderStyle = BorderStyle.None, Font = Ui.F(9.5f), HeaderStyle = ColumnHeaderStyle.Nonclickable };
            folderList.Columns.Add("Tên", w - 200); folderList.Columns.Add("Ngày sửa đổi", 180);
            folderList.ItemActivate += (s, e) => { if (folderList.SelectedItems.Count == 0) return; string p = (string)folderList.SelectedItems[0].Tag; if (Directory.Exists(p)) Browse(p); else saveName.Text = Path.GetFileNameWithoutExtension(p); };
            content.Controls.Add(folderList); FillFolder();
            save.Click += (s, e) =>
            {
                List<KeyValuePair<string, string>> fm = Formats(host.DocType); string ext = fm[Math.Max(0, saveType.SelectedIndex)].Key;
                if (WebOnly(host.DocType, ext)) { Web("fmt:" + ext); return; }
                string file = Path.Combine(saveDir, saveName.Text.Trim() + "." + ext);
                if (File.Exists(file) && !string.Equals(file, host.DocPath, StringComparison.OrdinalIgnoreCase) && MessageBox.Show("\"" + Path.GetFileName(file) + "\" đã tồn tại. Thay thế?", "Lưu thành", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                if (host.SaveDocAs(file, ext)) host.CloseBackstage();
            };
        }
        void FillFolder()
        {
            folderList.Items.Clear();
            try
            {
                foreach (string d in Directory.GetDirectories(saveDir)) { if ((File.GetAttributes(d) & FileAttributes.Hidden) != 0) continue; ListViewItem it = new ListViewItem(new[] { "📁  " + Path.GetFileName(d), Directory.GetLastWriteTime(d).ToString("dd/MM/yyyy HH:mm") }) { Tag = d }; folderList.Items.Add(it); }
                HashSet<string> exts = new HashSet<string>(); foreach (KeyValuePair<string, string> f in Formats(host.DocType)) exts.Add("." + f.Key);
                foreach (string f in Directory.GetFiles(saveDir)) { if (!exts.Contains(Path.GetExtension(f).ToLowerInvariant()) || Path.GetFileName(f).StartsWith("~$") || Path.GetFileName(f).StartsWith(".~lock")) continue; ListViewItem it = new ListViewItem(new[] { "📄  " + Path.GetFileName(f), File.GetLastWriteTime(f).ToString("dd/MM/yyyy HH:mm") }) { Tag = f }; folderList.Items.Add(it); }
            }
            catch { }
        }
        void ClassicSaveAs(string dir)
        {
            List<KeyValuePair<string, string>> fm = Formats(host.DocType); string filter = "";
            foreach (KeyValuePair<string, string> f in fm) filter += (filter.Length > 0 ? "|" : "") + f.Value + " (*." + f.Key + ")|*." + f.Key;
            using (SaveFileDialog d = new SaveFileDialog { Filter = filter, InitialDirectory = dir, FileName = Path.GetFileNameWithoutExtension(host.DocPath) })
                if (d.ShowDialog(FindForm()) == DialogResult.OK) { string ext = Path.GetExtension(d.FileName).TrimStart('.').ToLowerInvariant(); if (host.SaveDocAs(d.FileName, ext)) host.CloseBackstage(); }
        }

        public static List<KeyValuePair<string, string>> Formats(int type)
        {
            // v0.1.7: same file-type list as the Web editor (ui_settings.json); types only the Web editor writes are handed to it
            List<KeyValuePair<string, string>> r = new List<KeyValuePair<string, string>>();
            foreach (string[] f in UiSpec.Formats(type)) r.Add(new KeyValuePair<string, string>(f[0], f[1]));
            return r;
        }
        public static bool WebOnly(int type, string ext) { foreach (string[] f in UiSpec.Formats(type)) if (f[0] == ext) return f[2] == "w"; return false; }

        // ---------------- Print ----------------
        int previewPage; PictureBox preview; Label pageLbl; NumericUpDown copies; ComboBox printerBox, rangeBox, orientBox, sidesBox, collateBox; TextBox pagesBox;
        void PagePrint()
        {
            H1("In", 26);
            int x = 40, y = 84, w = 300;
            Button print = new Button { Location = new Point(x, y), Size = new Size(84, 84), FlatStyle = FlatStyle.Flat, Text = "In", TextAlign = ContentAlignment.BottomCenter, Font = Ui.F(10f), BackColor = Color.White, ImageAlign = ContentAlignment.TopCenter };
            print.FlatAppearance.BorderColor = Color.FromArgb(0xC8, 0xC8, 0xC8);
            Bitmap pi = new Bitmap(32, 32); using (Graphics g = Graphics.FromImage(pi)) Icons.Draw(g, "print", new Rectangle(0, 0, 32, 32), true); print.Image = pi; print.Padding = new Padding(0, 10, 0, 6);
            content.Controls.Add(print);
            P("Bản sao:", x + 100, y + 30, 70).ForeColor = Ui.Text;
            copies = new NumericUpDown { Location = new Point(x + 170, y + 28), Width = 60, Minimum = 1, Maximum = 999, Value = 1 }; content.Controls.Add(copies);
            y += 100;
            H2("Máy in", x, y); y += 30;
            printerBox = new ComboBox { Location = new Point(x, y), Width = w, DropDownStyle = ComboBoxStyle.DropDownList, Font = Ui.F(9.5f) };
            string def = new PrinterSettings().PrinterName; int di = 0;
            foreach (string pn in PrinterSettings.InstalledPrinters) { if (pn == def) di = printerBox.Items.Count; printerBox.Items.Add(pn); }
            if (printerBox.Items.Count == 0) printerBox.Items.Add("(Không có máy in)");
            printerBox.SelectedIndex = di; content.Controls.Add(printerBox); y += 30;
            LinkLabel props = new LinkLabel { Text = "Thuộc tính máy in", Location = new Point(x + w - 110, y), AutoSize = true, LinkColor = host.Accent, BackColor = Color.White };
            props.LinkClicked += (s, e) => { using (PrintDialog d = new PrintDialog { PrinterSettings = new PrinterSettings { PrinterName = (string)printerBox.SelectedItem }, UseEXDialog = true }) if (d.ShowDialog(FindForm()) == DialogResult.OK) { } };
            content.Controls.Add(props); y += 30;
            H2("Cài đặt", x, y); y += 30;
            string unit = host.DocType == LokDoc.PRESENTATION ? "trang chiếu" : "trang";
            rangeBox = Combo(x, ref y, w, "In tất cả " + unit, "In " + unit + " hiện tại", "Phạm vi tùy chỉnh");
            P("Trang:", x, y + 3, 50).ForeColor = Ui.Text; pagesBox = new TextBox { Location = new Point(x + 54, y), Width = w - 54, Font = Ui.F(9.5f) }; SetCue(pagesBox, "vd: 1-3, 5"); content.Controls.Add(pagesBox); y += 32;
            pagesBox.TextChanged += (s, e) => { if (pagesBox.Text.Length > 0) rangeBox.SelectedIndex = 2; };
            sidesBox = Combo(x, ref y, w, "In một mặt", "In hai mặt (lật theo cạnh dài)", "In hai mặt (lật theo cạnh ngắn)");
            collateBox = Combo(x, ref y, w, "Đối chiếu  (1,2,3  1,2,3  1,2,3)", "Không đối chiếu  (1,1,1  2,2,2  3,3,3)");
            orientBox = Combo(x, ref y, w, "Hướng dọc", "Hướng ngang"); orientBox.SelectedIndex = host.DocType == LokDoc.PRESENTATION ? 1 : (PageLandscape() ? 1 : 0);
            LinkLabel setup = new LinkLabel { Text = "Thiết lập trang", Location = new Point(x + w - 90, y), AutoSize = true, LinkColor = host.Accent, BackColor = Color.White };
            setup.LinkClicked += (s, e) => { host.CloseBackstage(); NativeDocForm f = FindForm() as NativeDocForm; if (f != null) f.ShowPageSetup(); };
            if (host.DocType == LokDoc.TEXT) content.Controls.Add(setup);
            print.Click += (s, e) =>
            {
                PrinterSettings ps = new PrinterSettings { PrinterName = (string)printerBox.SelectedItem, Copies = (short)copies.Value, Collate = collateBox.SelectedIndex == 0 };
                ps.Duplex = sidesBox.SelectedIndex == 0 ? Duplex.Simplex : sidesBox.SelectedIndex == 1 ? Duplex.Vertical : Duplex.Horizontal;
                string range = rangeBox.SelectedIndex == 0 ? "" : rangeBox.SelectedIndex == 1 ? "current" : pagesBox.Text;
                host.PrintDoc(ps, range, orientBox.SelectedIndex == 1); host.CloseBackstage();
            };
            // preview
            int px = 380, pw = Math.Max(300, content.Width - px - 40), ph = Math.Max(300, content.Height - 120);
            preview = new PictureBox { Location = new Point(px, 84), Size = new Size(pw, ph), SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(0xEE, 0xEE, 0xEE) };
            content.Controls.Add(preview);
            Button prev = new Button { Text = "◀", Location = new Point(px, 84 + ph + 8), Size = new Size(30, 24), FlatStyle = FlatStyle.Flat }; prev.FlatAppearance.BorderSize = 0;
            Button next = new Button { Text = "▶", Location = new Point(px + 120, 84 + ph + 8), Size = new Size(30, 24), FlatStyle = FlatStyle.Flat }; next.FlatAppearance.BorderSize = 0;
            pageLbl = new Label { Location = new Point(px + 34, 84 + ph + 12), Size = new Size(84, 20), TextAlign = ContentAlignment.TopCenter, BackColor = Color.White };
            prev.Click += (s, e) => { if (previewPage > 0) { previewPage--; RenderPreview(); } };
            next.Click += (s, e) => { if (previewPage < host.PageCount - 1) { previewPage++; RenderPreview(); } };
            content.Controls.Add(prev); content.Controls.Add(next); content.Controls.Add(pageLbl);
            previewPage = 0; RenderPreview();
        }
        bool PageLandscape() { Size s = host.PageSizeTwips(0); return s.Width > s.Height; }
        ComboBox Combo(int x, ref int y, int w, params string[] items)
        {
            ComboBox c = new ComboBox { Location = new Point(x, y), Width = w, DropDownStyle = ComboBoxStyle.DropDownList, Font = Ui.F(9.5f) };
            c.Items.AddRange(items); c.SelectedIndex = 0; content.Controls.Add(c); y += 32; return c;
        }
        void RenderPreview()
        {
            int n = Math.Max(1, host.PageCount);
            pageLbl.Text = (previewPage + 1) + " / " + n;
            Bitmap b = null; try { b = host.RenderPage(previewPage, 900); } catch { }
            Image old = preview.Image; preview.Image = b; if (old != null) old.Dispose();
        }

        void PageShare()
        {
            H1("Chia sẻ", 26);
            int y = 90;
            SpecTiles("share", 40, ref y, 520, 72, id =>
                id == "lanshare" ? (Action)host.ShareLan :
                id == "email" ? (Action)(() => { try { System.Diagnostics.Process.Start("mailto:?subject=" + Uri.EscapeDataString(Path.GetFileName(host.DocPath))); System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + host.DocPath + "\""); } catch { } }) :
                id == "sendpdf" ? (Action)host.ExportPdf : null);
        }

        void PageExport()
        {
            H1("Xuất", 26);
            int y = 90;
            SpecTiles("export", 40, ref y, 520, 80, id => id == "pdf" ? (Action)host.ExportPdf : null);
            H2("Thay đổi loại tệp", 40, y); y += 34;
            foreach (KeyValuePair<string, string> f in Formats(host.DocType))
            {
                KeyValuePair<string, string> ff = f;
                Big(f.Key == "pdf" ? "pdf" : host.DocType == LokDoc.SPREADSHEET ? "excel" : host.DocType == LokDoc.PRESENTATION ? "powerpoint" : "word", f.Value, "*." + UiSpec.FileExt(f.Key), 40, y, 520, () =>
                {
                    if (WebOnly(host.DocType, ff.Key)) { Web("fmt:" + ff.Key); return; }
                    using (SaveFileDialog d = new SaveFileDialog { Filter = ff.Value + "|*." + ff.Key, FileName = Path.GetFileNameWithoutExtension(host.DocPath) + "." + ff.Key, InitialDirectory = Path.GetDirectoryName(host.DocPath) })
                        if (d.ShowDialog(FindForm()) == DialogResult.OK) host.SaveDocAs(d.FileName, ff.Key);
                });
                y += 64;
            }
        }

        void PageAccount()
        {
            H1("Tài khoản", 26);
            int y = 90;
            H2("Thông tin người dùng", 40, y); y += 34;
            P("Tên của bạn (dùng cho ghi chú, theo dõi thay đổi):", 40, y, 420); y += 24;
            TextBox name = new TextBox { Location = new Point(40, y), Width = 300, Text = Ui.UserName(), Font = Ui.F(10f) }; content.Controls.Add(name);
            Button ok = new Button { Text = "Lưu tên", Location = new Point(350, y - 1), Size = new Size(90, 26), FlatStyle = FlatStyle.Flat }; content.Controls.Add(ok);
            ok.Click += (s, e) => { NativeSettings.Set("user", name.Text.Trim()); MessageBox.Show("Đã lưu tên người dùng.", "Tài khoản"); };
            y += 50;
            H2("Thông tin sản phẩm", 40, y); y += 34;
            P("Yomi-Office Open Source " + App.Version, 40, y, 600).ForeColor = Ui.Text; y += 24;
            P("Engine Native: LibreOffice " + ShortVer() + " (LibreOfficeKit, MPL-2.0)", 40, y, 600); y += 24;
            P("Ứng dụng mã nguồn mở, miễn phí. Giao diện tham khảo bố cục Microsoft 365.", 40, y, 700); y += 40;
            H2("Cập nhật phần mềm", 40, y); y += 34;
            P("Nguồn: GitHub Releases, Google Drive và MediaFire của Yomi-Office Open Source. Lần kiểm tra gần nhất: " + NativeUpdate.LastText(), 40, y, 700); y += 28;
            CheckBox auto = new CheckBox { Text = "Tự động kiểm tra khi mở tài liệu (mỗi ngày một lần)", Checked = NativeSettings.Bool("update.auto", true), Location = new Point(40, y), AutoSize = true, Font = Ui.F(9.5f), BackColor = Color.White };
            auto.CheckedChanged += (s, e) => NativeSettings.Set("update.auto", auto.Checked); content.Controls.Add(auto); y += 34;
            Big("update", "Cập nhật ngay", "Kiểm tra phiên bản mới và tự động cài đặt", 40, y, 420, () => NativeUpdate.CheckAsync(FindForm(), true));
        }
        static string ShortVer() { string v = Lok.VersionInfo ?? ""; int i = v.IndexOf("\"ProductVersion\": \""); int o = 19; if (i < 0) { i = v.IndexOf("\"ProductVersion\":\""); o = 18; } if (i < 0) return ""; i += o; int j = v.IndexOf('"', i); string e = ""; int k = v.IndexOf("\"ProductExtension\": \""); if (k > 0) { k += 21; int l = v.IndexOf('"', k); if (l > k) e = v.Substring(k, l - k); } return (j > i ? v.Substring(i, j - i) : "") + e; }

        // ---------------- small controls ----------------
        class Tile : Control
        {
            public string Icon, Title, Sub; public Action Click; public bool Selected; bool h;
            public Tile() { SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true); BackColor = Color.White; Cursor = Cursors.Hand; }
            protected override void OnPaint(PaintEventArgs e)
            {
                Graphics g = e.Graphics; g.Clear(Color.White);
                if (Selected) Ui.FillRound(g, Ui.Checked, new Rectangle(0, 0, Width - 1, Height - 1), 4); else if (h && Click != null) Ui.FillRound(g, Ui.Hover, new Rectangle(0, 0, Width - 1, Height - 1), 4);
                int isz = Sub == null ? 24 : 32; Icons.Draw(g, Icon, new Rectangle(10, (Height - isz) / 2, isz, isz), Click != null || Selected);
                Ui.Str(g, Title, Ui.F(10.5f), Click != null || Selected ? Ui.Text : Ui.Sub, new Rectangle(isz + 22, Sub == null ? 0 : 10, Width - isz - 26, Sub == null ? Height : 20), Sub == null ? TextFormatFlags.VerticalCenter : TextFormatFlags.Top);
                if (Sub != null) Ui.Str(g, Sub, Ui.F(9f), Ui.Sub, new Rectangle(isz + 22, 32, Width - isz - 26, 18), TextFormatFlags.Top);
            }
            protected override void OnMouseEnter(EventArgs e) { h = true; Invalidate(); }
            protected override void OnMouseLeave(EventArgs e) { h = false; Invalidate(); }
            protected override void OnMouseUp(MouseEventArgs e) { if (e.Button == MouseButtons.Left && Click != null) Click(); }
        }

        class Template : Control
        {
            public int Kind; public string Title; public Action Click; bool h;
            public Template() { SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true); BackColor = Color.White; Cursor = Cursors.Hand; }
            protected override void OnPaint(PaintEventArgs e)
            {
                Graphics g = e.Graphics; g.Clear(Color.White);
                Rectangle card = Kind == LokDoc.TEXT ? new Rectangle(22, 4, 106, 146) : new Rectangle(4, 30, 142, 92);
                using (SolidBrush sh = new SolidBrush(Color.FromArgb(24, 0, 0, 0))) g.FillRectangle(sh, card.X + 2, card.Y + 2, card.Width, card.Height);
                g.FillRectangle(Brushes.White, card);
                using (Pen p = new Pen(h ? Ui.AccentFor(Kind) : Color.FromArgb(0xD0, 0xD0, 0xD0), h ? 2f : 1f)) g.DrawRectangle(p, card);
                if (Kind == LokDoc.SPREADSHEET) using (Pen p = new Pen(Color.FromArgb(0xE2, 0xE2, 0xE2))) { for (int x = card.X; x < card.Right; x += 24) g.DrawLine(p, x, card.Y, x, card.Bottom); for (int y = card.Y; y < card.Bottom; y += 12) g.DrawLine(p, card.X, y, card.Right, y); using (Pen a = new Pen(Ui.AccentFor(Kind), 2f)) g.DrawRectangle(a, card.X + 1, card.Y + 1, 23, 11); }
                if (Kind == LokDoc.PRESENTATION) { Ui.Str(g, "Bấm để thêm tiêu đề", Ui.F(8f), Ui.Sub, new Rectangle(card.X, card.Y + 30, card.Width, 20), TextFormatFlags.HorizontalCenter); }
                Icons.Draw(g, Kind == LokDoc.TEXT ? "word" : Kind == LokDoc.SPREADSHEET ? "excel" : "powerpoint", new Rectangle(card.X + 4, card.Bottom - 22, 16, 16), true);
                Ui.Str(g, Title, Ui.F(9f), Ui.Text, new Rectangle(0, Height - 34, Width, 30), TextFormatFlags.HorizontalCenter | TextFormatFlags.WordBreak);
            }
            protected override void OnMouseEnter(EventArgs e) { h = true; Invalidate(); }
            protected override void OnMouseLeave(EventArgs e) { h = false; Invalidate(); }
            protected override void OnMouseUp(MouseEventArgs e) { if (e.Button == MouseButtons.Left && Click != null) Click(); }
        }

        class RecentGrid : Control
        {
            public Color Accent; public event Action<string> Open;
            readonly List<string> files = new List<string>(); readonly HashSet<string> pinned = new HashSet<string>(); int hot = -1, scroll;
            public RecentGrid() { SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true); BackColor = Color.White; Font = Ui.F(9.5f); }
            public void Load(List<string> pin, List<string> recent, string filter)
            {
                files.Clear(); pinned.Clear(); foreach (string p in pin) pinned.Add(p);
                foreach (string p in pin) if (Match(p, filter)) files.Add(p);
                foreach (string p in recent) if (!pinned.Contains(p) && Match(p, filter)) files.Add(p);
                scroll = 0; Invalidate();
            }
            static bool Match(string p, string f) { return string.IsNullOrEmpty(f) || Path.GetFileName(p).IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0; }
            protected override void OnPaint(PaintEventArgs e)
            {
                Graphics g = e.Graphics; g.Clear(Color.White);
                Ui.Str(g, "Tên", Ui.F(9f), Ui.Sub, new Rectangle(44, 0, 300, 24), TextFormatFlags.VerticalCenter);
                Ui.Str(g, "Ngày sửa đổi", Ui.F(9f), Ui.Sub, new Rectangle(Width - 170, 0, 160, 24), TextFormatFlags.VerticalCenter);
                using (Pen p = new Pen(Ui.Line)) g.DrawLine(p, 0, 25, Width, 25);
                if (files.Count == 0) { Ui.Str(g, "Chưa có tệp gần đây. Các tệp bạn mở sẽ xuất hiện ở đây.", Font, Ui.Sub, new Rectangle(44, 40, Width - 60, 24), TextFormatFlags.Left); return; }
                for (int i = 0; i < files.Count; i++)
                {
                    int y = 28 + i * 46 - scroll; if (y + 46 < 26) continue; if (y > Height) break;
                    Rectangle r = new Rectangle(0, y, Width, 44); if (i == hot) Ui.FillRound(g, Ui.Hover, r, 4);
                    string f = files[i], ext = Path.GetExtension(f).ToLowerInvariant();
                    Icons.Draw(g, ext.StartsWith(".x") || ext == ".ods" || ext == ".csv" ? "excel" : ext.StartsWith(".p") || ext == ".odp" ? "powerpoint" : "word", new Rectangle(10, y + 10, 24, 24), true);
                    Ui.Str(g, Path.GetFileName(f), Font, Ui.Text, new Rectangle(44, y + 5, Width - 260, 18), TextFormatFlags.Left);
                    Ui.Str(g, Path.GetDirectoryName(f), Ui.F(8.5f), Ui.Sub, new Rectangle(44, y + 23, Width - 260, 16), TextFormatFlags.Left | TextFormatFlags.PathEllipsis);
                    if (pinned.Contains(f) || i == hot) Icons.Draw(g, "pin", new Rectangle(Width - 200, y + 14, 16, 16), pinned.Contains(f));
                    Ui.Str(g, FriendlyDate(File.GetLastWriteTime(f)), Ui.F(9f), Ui.Sub, new Rectangle(Width - 170, y, 160, 44), TextFormatFlags.VerticalCenter);
                }
            }
            static string FriendlyDate(DateTime d) { if (d.Date == DateTime.Today) return "Hôm nay lúc " + d.ToString("HH:mm"); if (d.Date == DateTime.Today.AddDays(-1)) return "Hôm qua lúc " + d.ToString("HH:mm"); return d.ToString("dd/MM/yyyy"); }
            int Hit(Point p) { if (p.Y < 28) return -1; int i = (p.Y - 28 + scroll) / 46; return i >= 0 && i < files.Count ? i : -1; }
            protected override void OnMouseMove(MouseEventArgs e) { int h = Hit(e.Location); if (h != hot) { hot = h; Cursor = h >= 0 ? Cursors.Hand : Cursors.Default; Invalidate(); } }
            protected override void OnMouseLeave(EventArgs e) { hot = -1; Invalidate(); }
            protected override void OnMouseWheel(MouseEventArgs e) { scroll = Math.Max(0, Math.Min(Math.Max(0, files.Count * 46 + 30 - Height), scroll - e.Delta / 2)); Invalidate(); }
            protected override void OnMouseUp(MouseEventArgs e)
            {
                int h = Hit(e.Location); if (h < 0) return;
                if (e.X >= Width - 204 && e.X <= Width - 180) { NativeSettings.TogglePin(files[h]); Load(NativeSettings.Pinned(), NativeSettings.Recent(), ""); return; }
                if (e.Button == MouseButtons.Right)
                {
                    string f = files[h];
                    OfficeMenu.Build(new List<RMenu> { new RMenu("Mở", "open", () => { if (Open != null) Open(f); }), new RMenu(pinned.Contains(f) ? "Bỏ ghim" : "Ghim vào danh sách", "pin", () => { NativeSettings.TogglePin(f); Load(NativeSettings.Pinned(), NativeSettings.Recent(), ""); }), new RMenu("Mở vị trí tệp", "folder", () => { try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + f + "\""); } catch { } }), new RMenu("Sao chép đường dẫn", "copy", () => { try { Clipboard.SetText(f); } catch { } }) }).Show(this, e.Location);
                    return;
                }
                if (Open != null) Open(files[h]);
            }
        }
    }
}
