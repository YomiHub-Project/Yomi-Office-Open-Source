// Office Open Source v0.1.6 - native Word / Excel / PowerPoint editor window (LibreOfficeKit canvas + Office-style chrome)
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace OfficeOpenSource
{
    partial class NativeDocForm : Form, IBackstageHost
    {
        readonly LokCanvas canvas = new LokCanvas { Dock = DockStyle.Fill };
        readonly VScrollBar vs = new VScrollBar { Dock = DockStyle.Right };
        readonly HScrollBar hs = new HScrollBar { Dock = DockStyle.Bottom };
        readonly TitleBar title = new TitleBar();
        readonly Ribbon ribbon = new Ribbon();
        readonly StatusBar status = new StatusBar();
        FormulaBar formula; SheetHeaders headers; SheetTabs sheetTabs; SlidePane slides; Panel notesBar; LokRuler rulerTop; LokRuler rulerLeft; Backstage backstage;
        readonly Panel body = new Panel { Dock = DockStyle.Fill, BackColor = Ui.Workspace };
        LokDoc doc; string path; bool dirty, readOnlyCopy; int type; Color accent;
        StatusBar.Seg stPage, stWords, stLang, stMain, stAccess, stSum;
        readonly Timer autosave = new Timer { Interval = 120000 };
        readonly Timer liveSave = new Timer { Interval = 4000 };
        readonly Timer headerTimer = new Timer { Interval = 60 };
        Color lastFontColor = Color.FromArgb(0xC0, 0, 0), lastHighlight = Color.Yellow, lastFill = Color.Yellow;
        static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        public static void OpenAsync(string file, string cmd) { PendingUiCmd = string.IsNullOrEmpty(cmd) ? null : cmd; OpenAsync(file); }
        public static void OpenAsync(string file) { if (!Lok.Init()) { MessageBox.Show("Engine Native (LibreOffice) chưa sẵn sàng:\n" + Lok.Error, "Yomi-Office Open Source"); return; } Lok.InLoop(() => Open(file)); }
        /// <summary>Office-style splash shown while LibreOffice loads the file (large workbooks can take a while).</summary>
        static Form Splash(string file)
        {
            string ext = Path.GetExtension(file).ToLowerInvariant();
            int t = ext.StartsWith(".x") || ext == ".ods" || ext == ".csv" ? LokDoc.SPREADSHEET : ext.StartsWith(".p") || ext == ".odp" ? LokDoc.PRESENTATION : LokDoc.TEXT;
            Color ac = Ui.AccentFor(t);
            Form f = new Form { FormBorderStyle = FormBorderStyle.None, StartPosition = FormStartPosition.CenterScreen, Size = new Size(460, 230), ShowInTaskbar = true, BackColor = Color.White, Text = "Đang mở " + Path.GetFileName(file), TopMost = false };
            f.Paint += (s, e) =>
            {
                Graphics g = e.Graphics; g.Clear(ac);
                Icons.Draw(g, t == LokDoc.SPREADSHEET ? "excel" : t == LokDoc.PRESENTATION ? "powerpoint" : "word", new Rectangle(28, 28, 32, 32), true);
                Ui.Str(g, t == LokDoc.SPREADSHEET ? "Excel" : t == LokDoc.PRESENTATION ? "PowerPoint" : "Word", Ui.F(20f), Color.White, new Rectangle(70, 24, 360, 40), TextFormatFlags.VerticalCenter);
                Ui.Str(g, "Yomi-Office Open Source", Ui.F(9f), Color.FromArgb(230, 255, 255, 255), new Rectangle(72, 62, 360, 20), TextFormatFlags.Left);
                Ui.Str(g, "Đang mở…", Ui.F(10f), Color.White, new Rectangle(28, 140, 400, 22), TextFormatFlags.Left);
                Ui.Str(g, Path.GetFileName(file), Ui.F(9f), Color.FromArgb(230, 255, 255, 255), new Rectangle(28, 164, 400, 20), TextFormatFlags.Left | TextFormatFlags.PathEllipsis);
                using (SolidBrush b = new SolidBrush(Color.FromArgb(120, 255, 255, 255))) for (int i = 0; i < 5; i++) g.FillEllipse(b, 28 + i * 14, 196, 6, 6);
            };
            f.Show(); f.Refresh(); Application.DoEvents();
            return f;
        }
        public static NativeDocForm Open(string file)
        {
            Form splash = null; try { splash = Splash(file); } catch { }
            LokDoc d;
            try { d = Lok.Load(file, Lok.MacrosAllowed); } finally { if (splash != null) { splash.Close(); splash.Dispose(); } }
            if (d == null)
            {
                Lok.Trace("native open failed, fallback to web: " + file + " " + Lok.Error);
                App.AddPending(file, true); App.ActivateExistingWindow();
                return null;
            }
            NativeDocForm f = new NativeDocForm(d, file);
            if (Lok.LastReadOnly) f.readOnlyCopy = true;
            f.UpdateTitle();
            f.Show(); f.Activate();
            NativeUpdate.AutoCheck(f); // Tùy chọn → Cập nhật phần mềm: once a day, GitHub + Google Drive + MediaFire
            return f;
        }

        NativeDocForm(LokDoc d, string file)
        {
            doc = d; path = file; type = d.Type; accent = Ui.AccentFor(type);
            if (!file.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase)) NativeSettings.AddRecent(file);
            Text = Path.GetFileName(file) + " - Yomi-Office Open Source";
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            FormBorderStyle = FormBorderStyle.None; DoubleBuffered = true; KeyPreview = true;
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            Size = new Size(Math.Min(1440, wa.Width - 40), Math.Min(920, wa.Height - 40)); StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(760, 480);
            Font = Ui.F(9f); BackColor = Ui.Chrome; Padding = new Padding(1);
            title.Accent = accent; title.AppIcon = type == LokDoc.SPREADSHEET ? "excel" : type == LokDoc.PRESENTATION ? "powerpoint" : "word";
            title.AutoSave = NativeSettings.Bool("autosave.live", false);
            ribbon.Accent = accent; status.Accent = accent;

            Lok.Activate(doc);
            BuildRibbon();
            BuildBody();
            BuildStatus();
            Controls.Add(body);
            if (formula != null) Controls.Add(formula);
            Controls.Add(ribbon); Controls.Add(title); Controls.Add(status);
            body.BringToFront();

            d.InitRender("{\".uno:ShowBorderShadow\":{\"type\":\"boolean\",\"value\":false},\".uno:HideWhitespace\":{\"type\":\"boolean\",\"value\":false},\".uno:Author\":{\"type\":\"string\",\"value\":\"" + J(Ui.UserName()) + "\"}}");
            canvas.Attach(d);
            canvas.Message += OnMessage;
            canvas.Scrolled += () => { SyncBars(); headerTimer.Stop(); headerTimer.Start(); };
            canvas.MouseUp += (s, e) => { if (e.Button == MouseButtons.Right) ShowContextMenu(e.Location); };
            vs.Scroll += (s, e) => canvas.ScrollTo(canvas.Scroll.X, vs.Value);
            hs.Scroll += (s, e) => canvas.ScrollTo(hs.Value, canvas.Scroll.Y);
            headerTimer.Tick += (s, e) => { headerTimer.Stop(); if (freezeCol > 0 || freezeRow > 0) UpdateFreeze(); RefreshHeaders(); };
            Shown += (s, e) =>
            {
                { string tc = Environment.GetEnvironmentVariable("OOS_TEST_CMD"); if (!string.IsNullOrEmpty(tc) && PendingUiCmd == null) PendingUiCmd = tc; }   // v0.1.7 UI tests
                if (PendingUiCmd != null) { string pc = PendingUiCmd; PendingUiCmd = null; Timer ct = new Timer { Interval = 1200 }; ct.Tick += (s3, e3) => { ct.Stop(); ct.Dispose(); RunUiCommand(pc); }; ct.Start(); }
                string dumpUi = Environment.GetEnvironmentVariable("OOS_DUMP_UI");   // v0.1.7: export ribbon / status bar / icons for the Web UI
                if (!string.IsNullOrEmpty(dumpUi)) { try { DumpUi(dumpUi); } catch (Exception dx) { File.WriteAllText(dumpUi + ".err", dx.ToString()); } dirty = false; Lok.HardExit(0); return; }
                canvas.SetZoom(type == LokDoc.PRESENTATION ? FitSlideZoom() : type == LokDoc.SPREADSHEET ? 1f : FitWidthZoom(1f));
                SyncBars(); canvas.Focus(); UpdateInfo(); RefreshHeaders(); LoadStyles();
                string ut = Environment.GetEnvironmentVariable("OOS_UI_TEST");   // automated screenshots: "tab=2" or "bs=print"
                if (!string.IsNullOrEmpty(ut)) { Timer tt = new Timer { Interval = 2500 }; tt.Tick += (s2, e2) => { tt.Stop(); foreach (string kv in ut.Split(';')) { if (kv.StartsWith("tab=")) { ribbon.Selected = int.Parse(kv.Substring(4)); ribbon.LayoutAll(); ribbon.Invalidate(); } if (kv.StartsWith("bs=")) OpenBackstage(kv.Substring(3)); if (kv == "max") WindowState = FormWindowState.Maximized;
                    if (kv.StartsWith("uno=")) Uno(kv.Substring(4));
                    if (kv.StartsWith("item=")) { string want = kv.Substring(5); foreach (RTab t2 in ribbon.Tabs) foreach (RGroup g2 in t2.Groups) foreach (RItem i2 in g2.All()) if ((i2.Tip ?? "").StartsWith(want) || i2.Text == want) { if (i2.Click != null) i2.Click(); goto done; } done:; }
                    if (kv.StartsWith("saveas=")) { string f2 = kv.Substring(7); SaveAsFile(f2, Path.GetExtension(f2).TrimStart('.').ToLowerInvariant()); }
                    if (kv == "close") { dirty = false; Close(); } } }; tt.Start(); }
            };
            Activated += (s, e) => Lok.Activate(doc);
            Resize += (s, e) => { SyncBars(); Padding = WindowState == FormWindowState.Maximized ? Padding.Empty : new Padding(1); title.Invalidate(); };
            FormClosing += OnClosing;
            autosave.Interval = Math.Max(1, NativeSettings.Int("autosave.minutes", 2)) * 60000;
            autosave.Tick += (s, e) => { if (dirty) AutoSaveCopy(); }; autosave.Start();
            liveSave.Tick += (s, e) => { liveSave.Stop(); if (title.AutoSave && dirty && !readOnlyCopy) Save(true); };
            WireTitle();
        }

        void DumpUi(string file)
        {
            Func<RItem, Dictionary<string, object>> item = null;
            Func<List<RMenu>, List<object>> menu = m => { List<object> r = new List<object>(); if (m != null) foreach (RMenu x in m) r.Add(x.Separator ? (object)"-" : new Dictionary<string, object> { { "t", x.Text }, { "i", x.Icon } }); return r; };
            item = it => { Dictionary<string, object> o = new Dictionary<string, object> { { "k", it.Kind.ToString() }, { "t", it.Text }, { "i", it.Icon }, { "tip", it.Tip }, { "split", it.Split }, { "color", it.PickColor != null }, { "state", it.State } };
                if (it.Menu != null) o["menu"] = menu(it.Menu);
                if (it.Gallery != null) { List<object> g = new List<object>(); foreach (RGalleryItem gi in it.Gallery) g.Add(new Dictionary<string, object> { { "t", gi.Name }, { "p", gi.Preview } }); o["gallery"] = g; }
                if (it.Combo != null) o["combo"] = true; return o; };
            List<object> tabs = new List<object>();
            foreach (RTab t in ribbon.Tabs)
            {
                List<object> groups = new List<object>();
                foreach (RGroup g in t.Groups)
                {
                    List<object> cols = new List<object>();
                    foreach (RColumn c in g.Cols)
                    {
                        Dictionary<string, object> co = new Dictionary<string, object>();
                        if (c.Large != null) { co["type"] = "large"; co["items"] = new List<object> { item(c.Large) }; }
                        else if (c.GalleryItem != null) { co["type"] = "gallery"; co["items"] = new List<object> { item(c.GalleryItem) }; }
                        else if (c.Stack != null) { co["type"] = "stack"; List<object> l = new List<object>(); foreach (RItem i in c.Stack) l.Add(item(i)); co["items"] = l; }
                        else if (c.Rows != null) { co["type"] = "rows"; List<object> rows = new List<object>(); foreach (List<RItem> r in c.Rows) { List<object> l = new List<object>(); foreach (RItem i in r) l.Add(item(i)); rows.Add(l); } co["rows"] = rows; }
                        cols.Add(co);
                    }
                    groups.Add(new Dictionary<string, object> { { "name", g.Name }, { "icon", g.Icon }, { "launcher", g.Launcher != null ? (g.LauncherTip ?? "") : null }, { "cols", cols } });
                }
                tabs.Add(new Dictionary<string, object> { { "name", t.Name }, { "groups", groups } });
            }
            List<object> segs = new List<object>(); foreach (StatusBar.Seg sg in status.Segs) segs.Add(new Dictionary<string, object> { { "t", sg.Text }, { "i", sg.Icon }, { "tip", sg.Tip } });
            List<object> views = new List<object>(); foreach (StatusBar.Seg sg in status.Views) views.Add(new Dictionary<string, object> { { "i", sg.Icon }, { "tip", sg.Tip } });
            Dictionary<string, string> pal = new Dictionary<string, string>(); foreach (var kv in Icons.Palette) pal[kv.Key.ToString()] = "#" + kv.Value.R.ToString("x2") + kv.Value.G.ToString("x2") + kv.Value.B.ToString("x2");
            Dictionary<string, object> root = new Dictionary<string, object> { { "type", type }, { "accent", "#" + accent.R.ToString("x2") + accent.G.ToString("x2") + accent.B.ToString("x2") }, { "tabs", tabs }, { "status", segs }, { "views", views }, { "icons", Icons.Specs }, { "palette", pal } };
            File.WriteAllText(file, new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Serialize(root), new UTF8Encoding(false));
        }

        // ---------------- window frame (borderless with resize, snap, shadow) ----------------
        protected override CreateParams CreateParams
        {
            get { CreateParams cp = base.CreateParams; cp.Style |= 0x00020000 | 0x00010000 | 0x00080000 | 0x00040000; cp.ClassStyle |= 0x00020000; return cp; } // MINIMIZEBOX | MAXIMIZEBOX | SYSMENU | THICKFRAME, CS_DROPSHADOW
        }
        [StructLayout(LayoutKind.Sequential)] struct NCCALCSIZE_PARAMS { public RECT r0, r1, r2; public IntPtr pos; }
        [StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }
        [StructLayout(LayoutKind.Sequential)] struct MINMAXINFO { public Point res, maxSize, maxPos, minTrack, maxTrack; }
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr h, int i);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr h, int i, int v);
        protected override void WndProc(ref Message m)
        {
            const int WM_NCCALCSIZE = 0x83, WM_NCHITTEST = 0x84, WM_GETMINMAXINFO = 0x24;
            // borderless window keeps WS_CAPTION for snap/animations: stop Windows from painting the classic caption over our
            // title bar when the window is deactivated (modal dialogs) or its text/icon changes
            if (m.Msg == 0x86) { m.LParam = new IntPtr(-1); base.WndProc(ref m); return; } // WM_NCACTIVATE
            if ((m.Msg == 0x0C || m.Msg == 0x80) && IsHandleCreated) // WM_SETTEXT / WM_SETICON
            {
                int st = GetWindowLong(Handle, -16); SetWindowLong(Handle, -16, st & ~0x10000000);
                base.WndProc(ref m); SetWindowLong(Handle, -16, st); return;
            }
            if (m.Msg == WM_NCCALCSIZE && m.WParam != IntPtr.Zero)
            {
                // whole window is client area (THICKFRAME kept for Aero snap / resize); when maximized keep it inside the work area
                if (WindowState == FormWindowState.Maximized)
                {
                    NCCALCSIZE_PARAMS p = (NCCALCSIZE_PARAMS)Marshal.PtrToStructure(m.LParam, typeof(NCCALCSIZE_PARAMS));
                    Rectangle wa = Screen.FromHandle(Handle).WorkingArea; p.r0.L = wa.Left; p.r0.T = wa.Top; p.r0.R = wa.Right; p.r0.B = wa.Bottom;
                    Marshal.StructureToPtr(p, m.LParam, false);
                }
                m.Result = IntPtr.Zero; return;
            }
            if (m.Msg == WM_GETMINMAXINFO)
            {
                MINMAXINFO mi = (MINMAXINFO)Marshal.PtrToStructure(m.LParam, typeof(MINMAXINFO));
                Screen sc = Screen.FromHandle(Handle); Rectangle wa = sc.WorkingArea, b = sc.Bounds;
                mi.maxPos = new Point(wa.Left - b.Left, wa.Top - b.Top); mi.maxSize = new Point(wa.Width, wa.Height); mi.minTrack = new Point(MinimumSize.Width, MinimumSize.Height);
                Marshal.StructureToPtr(mi, m.LParam, false); m.Result = IntPtr.Zero; return;
            }
            if (m.Msg == WM_NCHITTEST)
            {
                Point p = PointToClient(new Point((short)((long)m.LParam & 0xFFFF), (short)(((long)m.LParam >> 16) & 0xFFFF)));
                int e = 6;
                if (WindowState != FormWindowState.Maximized)
                {
                    bool l = p.X < e, r = p.X >= Width - e, t = p.Y < e, b = p.Y >= Height - e;
                    if (t && l) { m.Result = (IntPtr)13; return; } if (t && r) { m.Result = (IntPtr)14; return; }
                    if (b && l) { m.Result = (IntPtr)16; return; } if (b && r) { m.Result = (IntPtr)17; return; }
                    if (l) { m.Result = (IntPtr)10; return; } if (r) { m.Result = (IntPtr)11; return; }
                    if (t) { m.Result = (IntPtr)12; return; } if (b) { m.Result = (IntPtr)15; return; }
                }
                if (p.Y < title.Bottom && title.IsDragArea(title.PointToClient(PointToScreen(p)))) { m.Result = (IntPtr)2; return; }
                m.Result = (IntPtr)1; return;
            }
            base.WndProc(ref m);
        }
        protected override void OnPaint(PaintEventArgs e) { base.OnPaint(e); if (WindowState != FormWindowState.Maximized) using (Pen p = new Pen(Color.FromArgb(0xB4, 0xB4, 0xB4))) e.Graphics.DrawRectangle(p, 0, 0, Width - 1, Height - 1); }

        // ---------------- layout ----------------
        void BuildBody()
        {
            body.Controls.Clear();
            if (type == LokDoc.SPREADSHEET)
            {
                canvas.SheetMode = true; canvas.BackColor = Color.White;
                headers = new SheetHeaders { Accent = accent };
                headers.Select += a => { doc.Uno(".uno:GoToCell", "{\"ToPoint\":{\"type\":\"string\",\"value\":\"" + J(a) + "\"}}"); canvas.Focus(); };
                headers.SelectAll += () => { doc.Uno(".uno:SelectAll", null); canvas.Focus(); };
                canvas.Dock = DockStyle.None; canvas.Location = new Point(SheetHeaders.RowW, SheetHeaders.ColH);
                headers.Controls.Add(canvas);
                headers.Resize += (s, e) => canvas.Size = new Size(Math.Max(10, headers.Width - SheetHeaders.RowW), Math.Max(10, headers.Height - SheetHeaders.ColH));
                Panel bottom = new Panel { Dock = DockStyle.Bottom, Height = 28, BackColor = Ui.Chrome };
                sheetTabs = new SheetTabs { Dock = DockStyle.Left, Width = 560, Accent = accent };
                sheetTabs.TabClicked += i => SelectPart(i);
                sheetTabs.AddClicked += InsertSheet;
                sheetTabs.TabMenu += (i, pt) => SheetTabMenu(i, sheetTabs.PointToScreen(pt));
                hs.Dock = DockStyle.Fill; bottom.Controls.Add(hs); bottom.Controls.Add(sheetTabs);
                bottom.Resize += (s, e) => sheetTabs.Width = Math.Max(200, bottom.Width * 6 / 10);
                body.Controls.Add(headers); body.Controls.Add(vs); body.Controls.Add(bottom);
                BuildFormulaBar();
                RebuildParts();
                return;
            }
            if (type == LokDoc.PRESENTATION)
            {
                canvas.BackColor = Ui.Workspace;
                slides = new SlidePane { Accent = accent };
                slides.SetAspect(doc.SizeTwips);
                slides.Render = RenderSlideThumb;
                slides.SlideClicked += i => SelectPart(i);
                slides.SlideMenu += (i, pt) => SlideMenu(i, slides.PointToScreen(pt));
                notesBar = new Panel { Dock = DockStyle.Bottom, Height = 34, BackColor = Ui.Chrome, Cursor = Cursors.IBeam };
                notesBar.Paint += (s, e) => { Rectangle r = new Rectangle(8, 4, notesBar.Width - 16, notesBar.Height - 8); Ui.FillRound(e.Graphics, Color.White, r, 4); Ui.Str(e.Graphics, notesMode ? "Đang ở chế độ Trang ghi chú — bấm để quay lại dạng xem Thường" : "Bấm để thêm ghi chú", Ui.F(9.5f), Ui.Sub, new Rectangle(r.X + 10, r.Y, r.Width - 20, r.Height), TextFormatFlags.VerticalCenter); };
                notesBar.Click += (s, e) => ToggleNotes();
                Panel center = new Panel { Dock = DockStyle.Fill };
                center.Controls.Add(canvas); center.Controls.Add(vs); center.Controls.Add(hs); center.Controls.Add(notesBar);
                Splitter sp = new Splitter { Dock = DockStyle.Left, Width = 4, BackColor = Ui.Chrome };
                body.Controls.Add(center); body.Controls.Add(sp); body.Controls.Add(slides);
                RebuildParts();
                return;
            }
            canvas.BackColor = Ui.Workspace;
            rulerTop = new LokRuler(canvas, false); rulerLeft = new LokRuler(canvas, true);
            Panel corner = new Panel { Dock = DockStyle.Top, Height = 22, BackColor = Ui.Workspace };
            body.Controls.Add(canvas); body.Controls.Add(rulerLeft); body.Controls.Add(rulerTop); body.Controls.Add(vs); body.Controls.Add(hs);
            rulerTop.Visible = rulerLeft.Visible = NativeSettings.Bool("word.ruler", true);
            // vertical ruler: LibreOfficeKit only reports horizontal ruler data, so read the real top/bottom page margins from the file
            int[] vm = DocMargins.Read(path); Lok.Trace("ruler v-margins: " + (vm == null ? "n/a" : vm[0] + "/" + vm[1])); if (vm != null) SetVMargins(vm[0], vm[1]);
        }
        /// <summary>Top/bottom page margins for the vertical ruler, in twips.</summary>
        public void SetVMargins(int topTw, int bottomTw)
        {
            if (rulerLeft == null || topTw < 0 || bottomTw < 0) return;
            topMarginKnown = true; rulerLeft.Margin1 = topTw; rulerLeft.Margin2 = bottomTw; rulerLeft.Invalidate();
        }

        void BuildFormulaBar()
        {
            formula = new FormulaBar();
            formula.NameBox.KeyDown += (s, e) => { if (e.KeyCode != Keys.Enter) return; e.SuppressKeyPress = true; doc.Uno(".uno:GoToCell", "{\"ToPoint\":{\"type\":\"string\",\"value\":\"" + J(formula.NameBox.Text) + "\"}}"); canvas.Focus(); };
            formula.NameBox.SelectedIndexChanged += (s, e) => { if (formula.NameBox.SelectedItem != null) { doc.Uno(".uno:GoToCell", "{\"ToPoint\":{\"type\":\"string\",\"value\":\"" + J(formula.NameBox.SelectedItem.ToString()) + "\"}}"); canvas.Focus(); } };
            formula.Formula.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape) { canvas.Focus(); e.SuppressKeyPress = true; return; }
                if (e.KeyCode == Keys.Enter && !e.Alt) { e.SuppressKeyPress = true; CommitFormula(); }
            };
            formula.EnterClicked += CommitFormula;
            formula.CancelClicked += () => canvas.Focus();
            formula.FxClicked += () => ribbon.ShowMenu(FunctionMenu(null), ribbon.PointToClient(formula.PointToScreen(formula.FxPoint)));
        }
        void CommitFormula()
        {
            string t = NormalizeArgSep(formula.Formula.Text.Replace("\r", "").Replace("\n", " "));
            doc.Key(0, 0, 1286); doc.Key(1, 0, 1286);    // clear the cell (Delete)
            canvas.TypeText(t); doc.Key(0, 13, 1280); doc.Key(1, 13, 1280);
            MarkDirty(); canvas.Focus();
        }

        // v0.1.7: the engine runs with Vietnamese number formats (argument separator ';'). Formulas typed the English-Excel way
        // (=SUM(A1,B1), =IF(A1>0,"x","y")) used to fail with Err:501 - commas outside quotes / braces become ';' when the formula has none
        static string NormalizeArgSep(string t)
        {
            if (string.IsNullOrEmpty(t) || t[0] != '=' || t.IndexOf(';') >= 0 || t.IndexOf('(') < 0 || t.IndexOf(',') < 0) return t;
            System.Text.StringBuilder b = new System.Text.StringBuilder(t.Length); bool q = false; int brace = 0, depth = 0;
            foreach (char c in t)
            {
                if (c == '"') q = !q;
                if (!q) { if (c == '{') brace++; else if (c == '}') brace--; else if (c == '(') depth++; else if (c == ')') depth--; }
                b.Append(!q && brace == 0 && depth > 0 && c == ',' ? ';' : c);
            }
            return b.ToString();
        }

        void BuildStatus()
        {
            status.Font = Ui.F(8.5f);
            if (type == LokDoc.TEXT)
            {
                stPage = status.Add("Trang 1 / 1", null, "Số trang trong tài liệu", () => { string p = Ask(this, "Đi tới", "Đi tới trang số:", ""); int n; if (int.TryParse(p, out n)) GoToPage(n); });
                stWords = status.Add(InitialWords(path), null, "Số từ — bấm để xem thống kê", WordCount);
                stLang = status.Add("Tiếng Việt", "language", "Ngôn ngữ", null);
                stAccess = status.Add("Trợ năng: Tốt", "access", "Kiểm tra trợ năng", null);
                status.View("readmode", "Chế độ đọc", () => ReadMode(true));
                status.View("printlayout", "Bố cục in", () => { ReadMode(false); SetWebLayout(false); });
                status.View("weblayout", "Bố cục web", () => { ReadMode(false); SetWebLayout(true); });
                status.ActiveView = 1;
            }
            else if (type == LokDoc.SPREADSHEET)
            {
                stMain = status.Add("Sẵn sàng", null, null, null);
                stAccess = status.Add("Trợ năng: Kiểm tra", "access", null, null);
                stSum = status.Add("", null, "Tổng của vùng chọn", null);
                status.View("normalview", "Thường", () => doc.Uno(".uno:NormalViewMode", null));
                status.View("printlayout", "Bố cục trang (xem trước khi in)", () => OpenBackstage("print"));
                status.View("pagebreak", "Xem trước ngắt trang", () => doc.Uno(".uno:PagebreakMode", null));
            }
            else
            {
                stPage = status.Add("Trang chiếu 1 / 1", null, null, null);
                stLang = status.Add("Tiếng Việt", "language", null, null);
                stAccess = status.Add("Trợ năng: Tốt", "access", null, null);
                status.Add("Ghi chú", "notes", "Trang ghi chú", ToggleNotes);
                status.View("normalview", "Thường", () => { if (notesMode) ToggleNotes(); ShowSorter(false); });
                status.View("sorter", "Sắp xếp trang chiếu", () => ShowSorter(true));
                status.View("readmode", "Dạng xem đọc", () => SlideShow(doc.Part, false));
                status.View("play", "Trình chiếu từ trang hiện tại", () => SlideShow(doc.Part, true));
            }
            status.ZoomChanged += z => canvas.SetZoom(z);
            status.ZoomClicked += ZoomDialog;
        }

        void WireTitle()
        {
            title.SaveClicked += () => Save(false);
            title.UndoClicked += () => Uno(".uno:Undo");
            title.RedoClicked += () => Uno(".uno:Redo");
            title.AutoSaveToggled += () =>
            {
                NativeSettings.Set("autosave.live", title.AutoSave);
                if (title.AutoSave && readOnlyCopy) { MessageBox.Show("Tệp đang mở ở chế độ chỉ đọc. Hãy Lưu thành một tệp khác để bật Tự động lưu.", "Tự động lưu"); title.AutoSave = false; title.Invalidate(); return; }
                if (title.AutoSave && dirty) Save(true);
                UpdateTitle();
            };
            title.DocNameClicked += () => ribbon.ShowMenu(new List<RMenu> {
                new RMenu("Mở vị trí tệp", "folder", () => { try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + path + "\""); } catch { } }),
                new RMenu("Sao chép đường dẫn", "copy", () => { try { Clipboard.SetText(path); } catch { } }),
                new RMenu("Lưu thành…", "saveas", () => OpenBackstage("saveas")),
                new RMenu("Phiên bản (bản tự động lưu)", "refresh", () => OpenBackstage("info")) }, ribbon.PointToClient(title.PointToScreen(title.Below("name"))));
            title.QatClicked += () => ribbon.ShowMenu(new List<RMenu> {
                new RMenu("Mới", "blankpage", () => NewDoc(type)), new RMenu("Mở…", "open", () => OpenBackstage("open")),
                new RMenu("In nhanh", "print", () => PrintDoc(new PrinterSettings(), "", type == LokDoc.PRESENTATION)), new RMenu("Xem trước khi in", "printlayout", () => OpenBackstage("print")),
                new RMenu("Xuất PDF", "pdf", ExportPdf), RMenu.Sep(),
                new RMenu(ribbon.Minimized ? "Ghim Ribbon" : "Thu gọn Ribbon", "fullscreen", () => ribbon.SetMinimized(!ribbon.Minimized)) }, ribbon.PointToClient(title.PointToScreen(title.Below("qat"))));
            title.AccountClicked += () => OpenBackstage("account");
            title.SearchSubmitted += SearchCommand;
            ribbon.FileClicked += () => OpenBackstage("home");
            ribbon.ShareClicked += () => OpenBackstage("share");
            ribbon.CommentsClicked += InsertComment;
        }

        // ---------------- helpers ----------------
        static string J(string s) { return (s ?? "").Replace("\\", "\\\\").Replace("\"", "\\\""); }
        static string JS(string name, string v) { return "\"" + name + "\":{\"type\":\"string\",\"value\":\"" + J(v) + "\"}"; }
        static string JL(string name, long v) { return "\"" + name + "\":{\"type\":\"long\",\"value\":" + v + "}"; }
        static string JB(string name, bool v) { return "\"" + name + "\":{\"type\":\"boolean\",\"value\":" + (v ? "true" : "false") + "}"; }
        static string Args(params string[] parts) { return "{" + string.Join(",", parts) + "}"; }
        static int Rgb(Color c) { return c == Color.Transparent ? -1 : (c.R << 16) | (c.G << 8) | c.B; }
        void MarkDirty() { if (!dirty) { dirty = true; UpdateTitle(); } if (title.AutoSave) { liveSave.Stop(); liveSave.Start(); } UpdateInfo(); }
        void Uno(string cmd) { Uno(cmd, null); }
        void Uno(string cmd, string args) { Lok.Activate(doc); doc.Uno(cmd, args); MarkDirty(); canvas.Focus(); }
        void UnoView(string cmd) { Lok.Activate(doc); doc.Uno(cmd, null); canvas.Focus(); }
        Action U(string cmd) { return () => Uno(cmd); }
        Action U(string cmd, string args) { return () => Uno(cmd, args); }
        string DocKind { get { return type == LokDoc.TEXT ? "Word" : type == LokDoc.SPREADSHEET ? "Excel" : "PowerPoint"; } }

        void UpdateTitle()
        {
            title.DocName = Path.GetFileName(path);
            title.DocState = readOnlyCopy ? "Chỉ đọc" : dirty ? (title.AutoSave ? "Đang lưu…" : "Chưa lưu") : "Đã lưu vào PC này";
            Text = Path.GetFileName(path) + (readOnlyCopy ? " [Chỉ đọc]" : "") + " - Yomi-Office Open Source";
            title.Invalidate();
        }

        float FitSlideZoom() { Size tw = doc.SizeTwips; if (tw.Width <= 0) return 1f; int w = canvas.Width > 50 ? canvas.Width : body.Width - 220, h = canvas.Height > 50 ? canvas.Height : body.Height - 60; return Math.Max(0.1f, Math.Min(Math.Min((w - 60) * 15f / tw.Width, (h - 40) * 15f / tw.Height), 3f)); }
        float FitWidthZoom(float max) { Size tw = PageTw(); if (tw.Width <= 0) return 1f; return Math.Max(0.25f, Math.Min(max, (canvas.Width - 60) * 15f / tw.Width)); }
        Size PageTw() { List<Rectangle> pr = doc.PageRects(); if (pr.Count > 0) return pr[0].Size; return doc.SizeTwips; }

        void SyncBars()
        {
            Size d = canvas.DocPx;
            vs.Maximum = Math.Max(0, d.Height + 40); vs.LargeChange = Math.Max(1, canvas.Height); vs.Value = Math.Min(Math.Max(0, canvas.Scroll.Y), Math.Max(0, vs.Maximum));
            bool hv = canvas.SheetMode || d.Width > canvas.Width; hs.Visible = hv; hs.Maximum = Math.Max(0, d.Width + 16); hs.LargeChange = Math.Max(1, canvas.Width); hs.Value = Math.Min(Math.Max(0, canvas.Scroll.X), Math.Max(0, hs.Maximum));
            status.Zoom = canvas.Zoom; status.Invalidate();
            if (rulerTop != null) { rulerTop.Invalidate(); rulerLeft.Invalidate(); }
        }

        void UpdateInfo()
        {
            if (type == LokDoc.PRESENTATION && stPage != null) { stPage.Text = "Trang chiếu " + (doc.Part + 1) + " / " + doc.Parts; status.Invalidate(); }
            if (type == LokDoc.TEXT && stPage != null)
            {
                List<Rectangle> pr = doc.PageRects(); int cur = 1;
                if (pr.Count > 0) { Point mid = canvas.ToTw(new Point(canvas.Width / 2, canvas.Height / 3)); for (int i = 0; i < pr.Count; i++) if (mid.Y >= pr[i].Y) cur = i + 1; }
                stPage.Text = "Trang " + cur + " / " + Math.Max(1, pr.Count); status.Invalidate();
            }
        }

        // ---------------- parts (sheets / slides) ----------------
        bool notesMode, sorterMode;
        void SelectPart(int i)
        {
            if (i < 0 || i >= doc.Parts) return;
            if (i != doc.Part) { Lok.Activate(doc); doc.Part = i; canvas.RefreshSize(); canvas.ScrollTo(0, 0); }
            if (sheetTabs != null) { sheetTabs.Active = i; sheetTabs.EnsureVisible(i); }
            if (slides != null) { slides.Active = i; slides.EnsureVisible(i); slides.Invalidate(); if (type == LokDoc.PRESENTATION) canvas.SetZoom(FitSlideZoom()); }
            if (sorterMode) ShowSorter(false);
            UpdateInfo(); RefreshHeaders(); canvas.Focus();
        }
        void RebuildParts()
        {
            int n = doc.Parts, cur = doc.Part;
            if (sheetTabs != null) { sheetTabs.Names.Clear(); for (int i = 0; i < n; i++) { string nm = doc.PartName(i); sheetTabs.Names.Add(string.IsNullOrEmpty(nm) ? "Sheet" + (i + 1) : nm); } sheetTabs.Active = cur; sheetTabs.EnsureVisible(cur); }
            if (slides != null) { slides.Count = n; slides.Active = cur; slides.Reset(); }
            UpdateInfo();
        }
        void InsertSheet() { Lok.Activate(doc); doc.Uno(".uno:Insert", Args(JS("Name", "Sheet" + (doc.Parts + 1)), JL("Index", doc.Parts + 1))); MarkDirty(); RebuildParts(); SelectPart(doc.Parts - 1); }
        void SheetTabMenu(int i, Point screen)
        {
            SelectPart(i);
            OfficeMenu.Build(new List<RMenu> {
                new RMenu("Chèn trang tính", "insertcells", InsertSheet),
                new RMenu("Xóa", "deletecells", () => { if (doc.Parts > 1 && MessageBox.Show("Xóa trang tính \"" + doc.PartName(doc.Part) + "\"?", DocKind, MessageBoxButtons.YesNo) == DialogResult.Yes) { Uno(".uno:Remove"); RebuildParts(); } }),
                new RMenu("Đổi tên", "textbox", () => { string nm = Ask(this, "Đổi tên trang tính", "Tên mới:", doc.PartName(doc.Part)); if (!string.IsNullOrEmpty(nm)) { Uno(".uno:RenameTable", Args(JS("NewName", nm))); RebuildParts(); } }),
                new RMenu("Ẩn", "hideslide", () => { Uno(".uno:Hide"); RebuildParts(); }),
                new RMenu("Bỏ ẩn…", "sheet", () => { Uno(".uno:Show"); RebuildParts(); }) }).Show(screen);
        }
        void SlideMenu(int i, Point screen)
        {
            SelectPart(i);
            OfficeMenu.Build(new List<RMenu> {
                new RMenu("Trang chiếu mới", "newslide", NewSlide), new RMenu("Nhân bản trang chiếu", "duplicate", () => { Uno(".uno:DuplicatePage"); RebuildParts(); }),
                new RMenu("Xóa trang chiếu", "deleteslide", DeleteSlide), RMenu.Sep(),
                new RMenu("Di chuyển lên", "sortaz", () => { Uno(".uno:MovePageUp"); RebuildParts(); }), new RMenu("Di chuyển xuống", "sortaz", () => { Uno(".uno:MovePageDown"); RebuildParts(); }), RMenu.Sep(),
                new RMenu("Ẩn / hiện trang chiếu", "hideslide", () => Uno(".uno:HideSlide")) }).Show(screen);
        }
        void NewSlide() { Uno(".uno:InsertPage"); RebuildParts(); SelectPart(Math.Min(doc.Parts - 1, doc.Part)); }
        void DeleteSlide() { if (doc.Parts <= 1) return; Uno(".uno:DeletePage"); RebuildParts(); SelectPart(Math.Min(doc.Part, doc.Parts - 1)); }

        Bitmap RenderSlideThumb(int i, int w)
        {
            if (i >= doc.Parts) return null;
            Lok.Activate(doc);
            int cur = doc.Part; canvas.Suspend = true;
            try
            {
                if (i != cur) doc.Part = i;
                Size tw = doc.SizeTwips; if (tw.Width <= 0) return null; int h = (int)((long)w * tw.Height / tw.Width);
                return RenderTw(new Rectangle(0, 0, tw.Width, tw.Height), w, h);
            }
            finally { if (doc.Part != cur) doc.Part = cur; canvas.Suspend = false; }
        }
        Bitmap RenderTw(Rectangle tw, int w, int h)
        {
            Bitmap b = new Bitmap(Math.Max(1, w), Math.Max(1, h), PixelFormat.Format32bppPArgb);
            BitmapData bd = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppPArgb);
            try { doc.Paint(bd.Scan0, b.Width, b.Height, tw.X, tw.Y, tw.Width, tw.Height); if (doc.TileMode == 0) LokCanvas.SwapRBPublic(bd.Scan0, b.Width * b.Height); }
            finally { b.UnlockBits(bd); }
            return b;
        }

        void ToggleNotes()
        {
            if (type != LokDoc.PRESENTATION) return;
            notesMode = !notesMode; Lok.Activate(doc); doc.PartMode = notesMode ? 1 : 0;
            canvas.RefreshSize(); canvas.SetZoom(FitSlideZoom()); notesBar.Invalidate();
        }

        Panel sorter;
        void ShowSorter(bool on)
        {
            if (type != LokDoc.PRESENTATION) return;
            sorterMode = on;
            if (on)
            {
                if (sorter == null)
                {
                    sorter = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Ui.Workspace };
                    body.Controls.Add(sorter); sorter.BringToFront();
                }
                foreach (Control c in sorter.Controls) c.Dispose(); sorter.Controls.Clear(); sorter.Visible = true;
                int tw = 240, x = 24, y = 20;
                for (int i = 0; i < doc.Parts; i++)
                {
                    int ii = i; Bitmap b = RenderSlideThumb(i, tw);
                    PictureBox pb = new PictureBox { Image = b, Size = new Size(tw, b == null ? 135 : b.Height), Location = new Point(x, y), BorderStyle = BorderStyle.FixedSingle, Cursor = Cursors.Hand };
                    pb.Click += (s, e) => { status.ActiveView = 0; status.Invalidate(); SelectPart(ii); };
                    Label l = new Label { Text = (i + 1).ToString(), Location = new Point(x - 18, y), AutoSize = true, ForeColor = Ui.Sub };
                    sorter.Controls.Add(pb); sorter.Controls.Add(l);
                    x += tw + 36; if (x + tw > sorter.Width - 20) { x = 24; y += pb.Height + 30; }
                }
            }
            else if (sorter != null) sorter.Visible = false;
        }

        // ---------------- LOK messages ----------------
        public int RulerMargin1 = -1, RulerMargin2 = -1; bool topMarginKnown;
        void OnMessage(int t, string p)
        {
            if (t >= 35 && t != 8 && p != null && Environment.GetEnvironmentVariable("OOS_UI_TEST") != null) Lok.Trace("CB " + t + " " + (p.Length > 300 ? p.Substring(0, 300) : p));
            if (p != null && p.StartsWith("{") && p.Contains("\"margin1\"")) ParseRuler(p);
            switch (t)
            {
                case 0: if (slides != null) slides.MarkDirty(doc.Part); break;
                case 8: // STATE_CHANGED ".uno:Bold=true"
                    if (p == null) break;
                    if (p.IndexOf("Freeze", StringComparison.Ordinal) >= 0 || p.IndexOf("Split", StringComparison.Ordinal) >= 0) Lok.Trace("STATE " + p);
                    if (p.StartsWith("{")) { HandleJsonState(p); break; }
                    int eq = p.IndexOf('='); if (eq <= 0) break;
                    string k = p.Substring(0, eq), v = p.Substring(eq + 1);
                    states[k] = v; ribbon.SetState(k, v); ApplyFreezeState(k, v);
                    if (k == ".uno:Undo") { title.UndoEnabled = v != "disabled"; title.Invalidate(); }
                    if (k == ".uno:Redo") { title.RedoEnabled = v != "disabled"; title.Invalidate(); }
                    if (k == ".uno:ModifiedStatus") { bool d = v == "true"; if (d != dirty) { dirty = d; UpdateTitle(); if (d && title.AutoSave) { liveSave.Stop(); liveSave.Start(); } } }
                    if (k == ".uno:StateWordCount" && stWords != null) { stWords.Text = WordsFrom(v); status.Invalidate(); }
                    if (k == ".uno:StatePageNumber" && stPage != null) { stPage.Text = PageFrom(v); status.Invalidate(); }
                    if (k == ".uno:LanguageStatus" && stLang != null && v.Length > 0) { stLang.Text = LangName(v); status.Invalidate(); }
                    if (k == ".uno:StateTableCell" && stSum != null) { stSum.Text = SumFrom(v); status.Invalidate(); }
                    if (k == ".uno:RowColSelCount" && stMain != null) { stMain.Text = v.Length > 0 ? v.Replace("Selected:", "Đã chọn:").Replace("rows", "hàng").Replace("columns", "cột") : "Sẵn sàng"; status.Invalidate(); }
                    if (k == ".uno:PageStatus" && stPage != null && type == LokDoc.PRESENTATION) UpdateInfo();
                    break;
                case 13: SyncBars(); if (type == LokDoc.TEXT) UpdateInfo(); RefreshHeaders(); break;
                case 14: if (!canvas.Suspend) { if (sheetTabs != null && sheetTabs.Names.Count != doc.Parts) RebuildParts(); if (sheetTabs != null) { sheetTabs.Active = doc.Part; sheetTabs.Invalidate(); } if (slides != null) { if (slides.Count != doc.Parts) RebuildParts(); slides.Active = doc.Part; slides.Invalidate(); } UpdateInfo(); } break;
                case 17: RefreshHeaders(); break;
                case 19: if (formula != null && !formula.Formula.Focused) formula.Formula.Text = p ?? ""; break; // CELL_FORMULA
                case 34: if (formula != null && !formula.NameBox.Focused) formula.NameBox.Text = p ?? ""; break; // CELL_ADDRESS
                case 7: if (!string.IsNullOrEmpty(p) && (p.StartsWith("http") || p.StartsWith("mailto:"))) try { System.Diagnostics.Process.Start(p); } catch { } break;
                case 1:
                    if (type == LokDoc.TEXT)
                    {
                        UpdateInfo();
                        // vertical ruler: LibreOffice sends no vertical ruler data, so take the top margin from the first caret
                        // position after loading (cursor at the start of the body text) and mirror it for the bottom
                        if (!topMarginKnown && rulerLeft != null) { List<Rectangle> pr = doc.PageRects(); Rectangle c = canvas.CaretTw; if (pr.Count > 0 && !c.IsEmpty) { int top = c.Y - pr[0].Y; if (top > 0 && top < 6000) { topMarginKnown = true; SetVMargins(top, top); } } }
                    }
                    break;
            }
        }
        void HandleJsonState(string p)
        {
            try
            {
                Dictionary<string, object> o = Json.Deserialize<Dictionary<string, object>>(p);
                string cmd = o.ContainsKey("commandName") ? Convert.ToString(o["commandName"]) : null; object st = o.ContainsKey("state") ? o["state"] : null;
                if (cmd == null || st == null) return;
                ApplyFreezeState(cmd, Convert.ToString(st, CultureInfo.InvariantCulture));
            }
            catch { }
        }
        void ParseRuler(string p)
        {   // RULER_UPDATE (callback 35): {"margin1":..,"margin2":..,"leftOffset":..,"pageOffset":..,"pageWidth":..,"tabs":{..}} in 1/100 mm.
            // left page margin = leftOffset, right page margin = pageWidth - leftOffset - margin2 (same maths as Collabora's ruler).
            // Parsed with regexes: the "tabs" object repeats the key "" which JavaScriptSerializer rejects.
            try
            {
                Func<string, int> num = k => { System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(p, "\"" + k + "\"\\s*:\\s*\"?(-?\\d+)"); return m.Success ? int.Parse(m.Groups[1].Value) : int.MinValue; };
                int left = num("leftOffset"), m2 = num("margin2"), width = num("pageWidth");
                if (left == int.MinValue || m2 == int.MinValue || width <= 0) return;
                RulerMargin1 = left * 1440 / 2540; RulerMargin2 = Math.Max(0, width - left - m2) * 1440 / 2540;
                if (rulerTop != null) { rulerTop.Margin1 = RulerMargin1; rulerTop.Margin2 = RulerMargin2; rulerTop.Invalidate(); }
            }
            catch { }
        }
        // v0.1.7: the engine only reports the word count after the first edit/selection - show the count stored in the file meanwhile
        static string InitialWords(string file)
        {
            try
            {
                string ext = Path.GetExtension(file ?? "").ToLowerInvariant(); string x = null;
                if (ext == ".docx" || ext == ".docm" || ext == ".dotx") { x = DocMargins.Entry(file, "docProps/app.xml"); Match m = Regex.Match(x ?? "", "<Words>(\\d+)</Words>"); if (m.Success) return int.Parse(m.Groups[1].Value).ToString("N0", CultureInfo.GetCultureInfo("vi-VN")) + " từ"; }
                if (ext == ".odt") { x = DocMargins.Entry(file, "meta.xml"); Match m = Regex.Match(x ?? "", "meta:word-count=\"(\\d+)\""); if (m.Success) return int.Parse(m.Groups[1].Value).ToString("N0", CultureInfo.GetCultureInfo("vi-VN")) + " từ"; }
            }
            catch { }
            return "Đang đếm từ…";
        }
        static string WordsFrom(string v) { // "123 words, 456 characters" / "Selected: 2 words..."
            string n = ""; foreach (char c in v) { if (char.IsDigit(c) || c == '.' || c == ',') n += c; else if (n.Length > 0) break; }
            return n.Length > 0 ? n + " từ" : v;
        }
        static string PageFrom(string v) { string[] nums = System.Text.RegularExpressions.Regex.Split(v, "[^0-9]+"); List<string> l = new List<string>(); foreach (string s in nums) if (s.Length > 0) l.Add(s); return l.Count >= 2 ? "Trang " + l[0] + " / " + l[1] : v; }
        static string LangName(string v) { string s = v.Split(';')[0]; if (s.StartsWith("Vietnam") || v.Contains("vi-VN") || s.Contains("Việt") || s.Contains("Viet")) return "Tiếng Việt"; if (s.StartsWith("English")) return "Tiếng Anh" + (s.Contains("USA") ? " (Mỹ)" : s.Contains("UK") ? " (Anh)" : ""); return s; }
        static string SumFrom(string v) { string r = v.Replace("Sum=", "Tổng: ").Replace("Average=", "Trung bình: ").Replace("Count=", "Số đếm: ").Replace("Sum:", "Tổng:").Replace("Average:", "Trung bình:").Replace("Count:", "Số đếm:").Replace("CountA:", "Số đếm:"); return System.Text.RegularExpressions.Regex.IsMatch(r, "[1-9]") ? r : ""; }

        // Excel frozen panes: ".uno:FreezePanesColumn" / ".uno:FreezePanesRow" state = "tab, index" (LibreOfficeKit, as used by Collabora Online)
        int freezeCol, freezeRow;
        void ApplyFreezeState(string cmd, string v)
        {
            if (headers == null || (cmd != ".uno:FreezePanesColumn" && cmd != ".uno:FreezePanesRow")) return;
            // LibreOffice 24.8 sends "count / other" (e.g. ".uno:FreezePanesRow=10 / 10"); Collabora builds send "tab, count"
            int tab = doc.Part, idx;
            if (v.Contains("/")) { if (!int.TryParse(v.Split('/')[0].Trim(), out idx)) return; }
            else { string[] a = v.Split(','); if (a.Length < 2 || !int.TryParse(a[0].Trim(), out tab) || !int.TryParse(a[1].Trim(), out idx)) return; }
            if (tab != doc.Part) return;
            if (CurrentState(".uno:FreezePanes") == "false") idx = 0;
            if (cmd == ".uno:FreezePanesColumn") freezeCol = idx; else freezeRow = idx;
            UpdateFreeze();
        }
        void UpdateFreeze()
        {
            // pixel extent of the frozen rows / columns from the header data at scroll 0
            try
            {
                int fx = 0, fy = 0;
                if (freezeCol > 0 || freezeRow > 0)
                {
                    float tw = 15f / canvas.Zoom;
                    string js = doc.CommandValues(".uno:ViewRowColumnHeaders?x=0&y=0&width=" + (int)(4000 * tw) + "&height=" + (int)(4000 * tw));
                    Dictionary<string, object> o = Json.Deserialize<Dictionary<string, object>>(js); int s0;
                    List<KeyValuePair<string, int>> cols = Parse(o, "columns", out s0, 0), rows = Parse(o, "rows", out s0, 0);
                    if (freezeCol > 0 && freezeCol <= cols.Count) fx = cols[freezeCol - 1].Value;
                    if (freezeRow > 0 && freezeRow <= rows.Count) fy = rows[freezeRow - 1].Value;
                }
                if (fx != canvas.FreezeX || fy != canvas.FreezeY) { canvas.FreezeX = fx; canvas.FreezeY = fy; canvas.Invalidate(); RefreshHeaders(); }
            }
            catch (Exception ex) { Lok.Trace("freeze " + ex.Message); }
        }

        // Excel headers
        void RefreshHeaders()
        {
            if (headers == null || doc == null) return;
            try
            {
                float tpp = 15f / canvas.Zoom;
                int ax = canvas.FreezeX > 0 ? 0 : (int)(canvas.Scroll.X * tpp), ay = canvas.FreezeY > 0 ? 0 : (int)(canvas.Scroll.Y * tpp);
                int w = (int)((canvas.Width + (canvas.FreezeX > 0 ? canvas.Scroll.X : 0)) * tpp), h = (int)((canvas.Height + (canvas.FreezeY > 0 ? canvas.Scroll.Y : 0)) * tpp);
                string js = doc.CommandValues(".uno:ViewRowColumnHeaders?x=" + ax + "&y=" + ay + "&width=" + w + "&height=" + h);
                if (string.IsNullOrEmpty(js)) return;
                Dictionary<string, object> o = Json.Deserialize<Dictionary<string, object>>(js);
                int cs, rs;
                headers.Cols = Parse(o, "columns", out cs, canvas.Scroll.X, canvas.FreezeX); headers.Rows = Parse(o, "rows", out rs, canvas.Scroll.Y, canvas.FreezeY);
                headers.ColStart = cs; headers.RowStart = rs;
                headers.Current = canvas.CellPx; headers.Invalidate();
                if (formula != null && !formula.NameBox.Focused && !headers.Current.IsEmpty)
                {   // Name Box: column letter + row number of the cell cursor (CELL_ADDRESS is only sent after the cursor moves)
                    string col = null, row = null; int prev = headers.ColStart;
                    foreach (KeyValuePair<string, int> c in headers.Cols) { if (headers.Current.X + 2 >= prev && headers.Current.X + 2 < c.Value) col = c.Key; prev = c.Value; }
                    prev = headers.RowStart; foreach (KeyValuePair<string, int> r in headers.Rows) { if (headers.Current.Y + 2 >= prev && headers.Current.Y + 2 < r.Value) row = r.Key; prev = r.Value; }
                    if (col != null && row != null) formula.NameBox.Text = col + row;
                }
            }
            catch (Exception ex) { Lok.Trace("headers " + ex.Message); }
        }
        List<KeyValuePair<string, int>> Parse(Dictionary<string, object> o, string key, out int start, int scroll) { return Parse(o, key, out start, scroll, 0); }
        List<KeyValuePair<string, int>> Parse(Dictionary<string, object> o, string key, out int start, int scroll, int freeze)
        {
            List<KeyValuePair<string, int>> raw = ParseRaw(o, key, out start, scroll);
            if (freeze <= 0) return raw;
            // frozen panes: entries inside the frozen band keep their position, the rest shift by the scroll offset
            List<KeyValuePair<string, int>> r = new List<KeyValuePair<string, int>>(); start += scroll;
            foreach (KeyValuePair<string, int> e in raw)
            {
                int end = e.Value + scroll;   // back to unscrolled position
                if (end <= freeze) r.Add(new KeyValuePair<string, int>(e.Key, end));
                else if (end - scroll > freeze) r.Add(new KeyValuePair<string, int>(e.Key, end - scroll));
            }
            return r;
        }
        List<KeyValuePair<string, int>> ParseRaw(Dictionary<string, object> o, string key, out int start, int scroll)
        {
            List<KeyValuePair<string, int>> r = new List<KeyValuePair<string, int>>(); start = 0;
            if (!o.ContainsKey(key)) return r;
            System.Collections.ArrayList arr = o[key] as System.Collections.ArrayList; if (arr == null) return r;
            // positions are pixels at the LOK client zoom (tile px per tile twips) = our zoom; relative to the document origin
            bool first = true;
            foreach (object e in arr)
            {
                Dictionary<string, object> d = e as Dictionary<string, object>; if (d == null) continue;
                string text = Convert.ToString(d["text"], CultureInfo.InvariantCulture); int size = Convert.ToInt32(d["size"], CultureInfo.InvariantCulture);
                if (first) { first = false; if (text == "0") { start = size - scroll; continue; } }
                r.Add(new KeyValuePair<string, int>(text, size - scroll));
            }
            return r;
        }

        // ---------------- context menu ----------------
        void ShowContextMenu(Point p)
        {
            List<RMenu> m = new List<RMenu> {
                new RMenu("Cắt", "cut", () => { canvas.CopySelection(true); MarkDirty(); }), new RMenu("Sao chép", "copy", () => canvas.CopySelection(false)),
                new RMenu("Dán", "paste", () => { canvas.PasteClipboard(); MarkDirty(); }), new RMenu("Dán chỉ văn bản", "paste", () => { try { if (Clipboard.ContainsText()) { if (!doc.Paste("text/plain;charset=utf-8", Encoding.UTF8.GetBytes(Clipboard.GetText()))) canvas.TypeText(Clipboard.GetText()); MarkDirty(); } } catch { } }), RMenu.Sep() };
            if (type == LokDoc.SPREADSHEET)
            {
                m.Add(new RMenu("Chèn hàng phía trên", "insertrow", U(".uno:InsertRowsBefore"))); m.Add(new RMenu("Chèn cột bên trái", "insertcol", U(".uno:InsertColumnsBefore")));
                m.Add(new RMenu("Xóa hàng", "deleterow", U(".uno:DeleteRows"))); m.Add(new RMenu("Xóa cột", "deletecol", U(".uno:DeleteColumns")));
                m.Add(new RMenu("Xóa nội dung", "clear", () => { doc.Key(0, 0, 1286); doc.Key(1, 0, 1286); MarkDirty(); })); m.Add(RMenu.Sep());
                m.Add(new RMenu("Sắp xếp A → Z", "sortaz", U(".uno:SortAscending"))); m.Add(new RMenu("Bộ lọc", "filter", U(".uno:DataFilterAutoFilter")));
                m.Add(new RMenu("Chèn ghi chú", "comment", InsertComment)); m.Add(new RMenu("Độ rộng cột tối ưu", "columns", U(".uno:SetOptimalColumnWidthDirect")));
            }
            else
            {
                m.Add(new RMenu("Phông chữ…", "fontcolor", FontDialogShow)); m.Add(new RMenu("Đoạn văn…", "linespacing", ParagraphDialog));
                m.Add(RMenu.Sep()); m.Add(new RMenu("Liên kết…", "link", InsertLink)); m.Add(new RMenu("Bình luận mới", "comment", InsertComment));
                if (type == LokDoc.PRESENTATION) { m.Add(RMenu.Sep()); m.Add(new RMenu("Trang chiếu mới", "newslide", NewSlide)); }
            }
            OfficeMenu.Build(m).Show(canvas, p);
        }

        // ---------------- Backstage ----------------
        public void OpenBackstage(string page)
        {
            if (backstage == null) { backstage = new Backstage(this); Controls.Add(backstage); }
            backstage.Bounds = new Rectangle(Padding.Left, title.Bottom, ClientSize.Width - Padding.Horizontal, ClientSize.Height - title.Bottom - Padding.Bottom);
            backstage.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom; backstage.Dock = DockStyle.None;
            backstage.Visible = true; backstage.BringToFront(); backstage.ShowPage(page); backstage.Focus();
        }
        public void CloseBackstage() { if (backstage != null) backstage.Visible = false; canvas.Focus(); }
        int IBackstageHost.DocType { get { return type; } }
        string IBackstageHost.DocPath { get { return path; } }
        string IBackstageHost.DocKindName { get { return DocKind; } }
        Color IBackstageHost.Accent { get { return accent; } }
        void IBackstageHost.SaveDoc() { Save(false); }
        bool IBackstageHost.SaveDocAs(string file, string ext) { return SaveAsFile(file, ext); }
        public void NewDoc(int t)
        {
            try
            {
                string dir = Path.Combine(Path.GetTempPath(), "OfficeOpenSource-new"); Directory.CreateDirectory(dir);
                string ext = t == LokDoc.SPREADSHEET ? "xlsx" : t == LokDoc.PRESENTATION ? "pptx" : "docx";
                string baseName = t == LokDoc.SPREADSHEET ? "Sổ làm việc" : t == LokDoc.PRESENTATION ? "Bản trình bày" : "Tài liệu";
                int n = 1; string f; do { f = Path.Combine(dir, baseName + n + "." + ext); n++; } while (File.Exists(f));
                File.WriteAllBytes(f, BlankDocs.Get(ext));
                NativeDocForm nf = Open(f); if (nf != null) nf.isNewDoc = true;
                CloseBackstage();
            }
            catch (Exception e) { MessageBox.Show("Không tạo được tài liệu mới: " + e.Message, "Yomi-Office Open Source"); }
        }
        bool isNewDoc;
        public void OpenDoc(string file) { CloseBackstage(); if (File.Exists(file)) Open(file); }
        public void CloseDoc() { Close(); }
        void IBackstageHost.ExportPdf() { ExportPdf(); }
        public void ShowOptions() { using (OptionsDialog d = new OptionsDialog(this)) d.ShowDialog(this); ApplyOptions(); }
        void IBackstageHost.ShareLan() { ShareLan(); }
        int IBackstageHost.PageCount { get { if (type == LokDoc.TEXT) return Math.Max(1, doc.PageRects().Count); if (type == LokDoc.PRESENTATION) return doc.Parts; return Math.Max(1, CalcPages().Count); } }
        Size IBackstageHost.PageSizeTwips(int i) { if (type == LokDoc.TEXT) { List<Rectangle> pr = doc.PageRects(); if (i < pr.Count) return pr[i].Size; } return type == LokDoc.PRESENTATION ? doc.SizeTwips : new Size(11906, 16838); }
        Bitmap IBackstageHost.RenderPage(int i, int w) { return RenderPage(i, w); }
        void IBackstageHost.PrintDoc(PrinterSettings ps, string range, bool landscape) { PrintDoc(ps, range, landscape); }
        Dictionary<string, string> IBackstageHost.Properties()
        {
            Dictionary<string, string> r = new Dictionary<string, string>();
            try { FileInfo fi = new FileInfo(path); r["Kích thước"] = fi.Length >= 1048576 ? (fi.Length / 1048576.0).ToString("0.0") + " MB" : Math.Max(1, fi.Length / 1024) + " KB"; r["Ngày tạo"] = fi.CreationTime.ToString("dd/MM/yyyy HH:mm"); r["Lần sửa đổi cuối"] = fi.LastWriteTime.ToString("dd/MM/yyyy HH:mm"); } catch { }
            if (type == LokDoc.TEXT) { r["Số trang"] = Math.Max(1, doc.PageRects().Count).ToString(); if (stWords != null) r["Số từ"] = stWords.Text; }
            else r[type == LokDoc.SPREADSHEET ? "Số trang tính" : "Số trang chiếu"] = doc.Parts.ToString();
            r["Tác giả"] = Ui.UserName(); r["Định dạng"] = Path.GetExtension(path).TrimStart('.').ToUpperInvariant(); r["Engine"] = "LibreOffice (Native)";
            if (readOnlyCopy) r["Trạng thái"] = "Chỉ đọc (tệp đang mở ở nơi khác)";
            return r;
        }
        List<string> IBackstageHost.AutoSaveVersions()
        {
            List<string> r = new List<string>();
            try { string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOpenSource\\autosave"); foreach (string f in Directory.GetFiles(dir, Path.GetFileNameWithoutExtension(path) + ".autosave.*")) r.Add(f); r.Sort((a, b) => File.GetLastWriteTime(b).CompareTo(File.GetLastWriteTime(a))); } catch { }
            return r;
        }

        // ---------------- pages / printing ----------------
        List<Rectangle> CalcPages()
        {
            // approximate print pages for a sheet: the used area split into A4-sized blocks (portrait, 1 cm margins)
            List<Rectangle> r = new List<Rectangle>(); Size s = doc.SizeTwips; int pw = 11906 - 1134, ph = 16838 - 1134;
            if (s.Width <= 0 || s.Height <= 0) { r.Add(new Rectangle(0, 0, pw, ph)); return r; }
            for (int x = 0; x < s.Width; x += pw) for (int y = 0; y < s.Height; y += ph) r.Add(new Rectangle(x, y, Math.Min(pw, s.Width - x), Math.Min(ph, s.Height - y)));
            r.Sort((a, b) => a.X != b.X ? a.X.CompareTo(b.X) : a.Y.CompareTo(b.Y));
            return r;
        }
        Bitmap RenderPage(int i, int w)
        {
            Lok.Activate(doc);
            if (type == LokDoc.PRESENTATION) return RenderSlideThumb(Math.Min(i, doc.Parts - 1), w);
            Rectangle tw;
            if (type == LokDoc.TEXT) { List<Rectangle> pr = doc.PageRects(); if (pr.Count == 0) tw = new Rectangle(Point.Empty, doc.SizeTwips); else tw = pr[Math.Min(i, pr.Count - 1)]; }
            else { List<Rectangle> cp = CalcPages(); tw = cp[Math.Min(i, cp.Count - 1)]; int pw = 11906, ph = 16838; Bitmap page = new Bitmap(w, (int)((long)w * ph / pw)); using (Graphics g = Graphics.FromImage(page)) { g.Clear(Color.White); float sc = w / (float)pw; using (Bitmap c = RenderTw(tw, Math.Max(1, (int)(tw.Width * sc)), Math.Max(1, (int)(tw.Height * sc)))) g.DrawImageUnscaled(c, (int)(567 * sc), (int)(567 * sc)); } return page; }
            return RenderTw(tw, w, (int)((long)w * tw.Height / Math.Max(1, tw.Width)));
        }
        static List<int> ParseRange(string range, int count, int current)
        {
            List<int> r = new List<int>();
            if (string.IsNullOrEmpty(range)) { for (int i = 0; i < count; i++) r.Add(i); return r; }
            if (range == "current") { r.Add(current); return r; }
            foreach (string part in range.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] ab = part.Split('-'); int a, b;
                if (ab.Length == 2 && int.TryParse(ab[0].Trim(), out a) && int.TryParse(ab[1].Trim(), out b)) { for (int i = a; i <= b; i++) if (i >= 1 && i <= count && !r.Contains(i - 1)) r.Add(i - 1); }
                else if (int.TryParse(part.Trim(), out a) && a >= 1 && a <= count && !r.Contains(a - 1)) r.Add(a - 1);
            }
            return r;
        }
        public void PrintDoc(PrinterSettings ps, string range, bool landscape)
        {
            int count = ((IBackstageHost)this).PageCount; int current = type == LokDoc.PRESENTATION ? doc.Part : 0;
            if (type == LokDoc.TEXT && range == "current") { List<Rectangle> pr = doc.PageRects(); Point mid = canvas.ToTw(new Point(canvas.Width / 2, canvas.Height / 3)); for (int i = 0; i < pr.Count; i++) if (mid.Y >= pr[i].Y) current = i; }
            List<int> pages = ParseRange(range, count, current); if (pages.Count == 0) { MessageBox.Show("Phạm vi trang không hợp lệ.", "In"); return; }
            if (!ps.IsValid) { MessageBox.Show("Không tìm thấy máy in \"" + ps.PrinterName + "\".", "In"); return; }
            PrintDocument pd = new PrintDocument { PrinterSettings = ps, DocumentName = Path.GetFileName(path) };
            pd.DefaultPageSettings.Landscape = landscape; pd.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0); pd.OriginAtMargins = false;
            int k = 0;
            pd.PrintPage += (s, e) =>
            {
                Rectangle area = e.PageBounds; // 1/100 inch
                int dpi = 200; int w = area.Width * dpi / 100;
                using (Bitmap b = RenderPage(pages[k], w))
                {
                    if (b != null)
                    {
                        float sc = Math.Min(area.Width / (float)b.Width, area.Height / (float)b.Height);
                        float bw = b.Width * sc, bh = b.Height * sc;
                        e.Graphics.DrawImage(b, (area.Width - bw) / 2 - e.PageSettings.HardMarginX, (area.Height - bh) / 2 - e.PageSettings.HardMarginY, bw, bh);
                    }
                }
                k++; e.HasMorePages = k < pages.Count;
            };
            try { Cursor = Cursors.WaitCursor; pd.Print(); status.Segs[0].Text = "Đã gửi " + pages.Count + " trang tới " + ps.PrinterName; status.Invalidate(); }
            catch (Exception ex) { MessageBox.Show("In lỗi: " + ex.Message, "In"); }
            finally { Cursor = Cursors.Default; }
        }

        // ---------------- file ops ----------------
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (backstage != null && backstage.Visible) { if (keyData == Keys.Escape) { CloseBackstage(); return true; } return base.ProcessCmdKey(ref msg, keyData); }
            switch (keyData)
            {
                case Keys.Control | Keys.S: Save(false); return true;
                case Keys.F12: case Keys.Control | Keys.Shift | Keys.S: OpenBackstage("saveas"); return true;
                case Keys.Control | Keys.P: OpenBackstage("print"); return true;
                case Keys.Control | Keys.O: OpenBackstage("open"); return true;
                case Keys.Control | Keys.N: NewDoc(type); return true;
                case Keys.Control | Keys.H: FindReplace(true); return true;
                case Keys.Control | Keys.F: FindReplace(false); return true;
                case Keys.Control | Keys.W: Close(); return true;
                case Keys.Alt | Keys.Q: title.Search.Focus(); return true;
                case Keys.Control | Keys.F1: ribbon.SetMinimized(!ribbon.Minimized); return true;
                case Keys.F11: ToggleFullScreen(); return true;
            }
            if (type == LokDoc.PRESENTATION && keyData == Keys.F5) { SlideShow(0, true); return true; }
            if (type == LokDoc.PRESENTATION && keyData == (Keys.Shift | Keys.F5)) { SlideShow(doc.Part, true); return true; }
            if (type == LokDoc.PRESENTATION && slides != null && slides.Focused && keyData == Keys.Delete) { DeleteSlide(); return true; }
            if (keyData == (Keys.Control | Keys.K)) { InsertLink(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        protected override void OnKeyPress(KeyPressEventArgs e) { base.OnKeyPress(e); if (canvas.Focused && !char.IsControl(e.KeyChar)) MarkDirty(); }

        void Save(bool silent)
        {
            if (readOnlyCopy || isNewDoc) { if (!silent) OpenBackstage("saveas"); return; }
            string ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
            try
            {
                Lok.Activate(doc);
                string tmp = path + ".oos-saving";
                bool ok = doc.SaveAs(tmp, ext, ext == "csv" ? "44,34,76,1" : "");
                if (ok && File.Exists(tmp)) { File.Copy(tmp, path, true); File.Delete(tmp); dirty = false; UpdateTitle(); }
                else if (!silent) MessageBox.Show("Không lưu được: " + (Lok.LastError() ?? "lỗi không xác định"), "Lưu");
            }
            catch (Exception e) { if (!silent) MessageBox.Show("Không lưu được: " + e.Message, "Lưu"); }
        }
        bool SaveAsFile(string file, string ext)
        {
            Lok.Activate(doc);
            if (doc.SaveAs(file, ext, ext == "csv" ? "44,34,76,1" : ""))
            {
                if (ext != "pdf" && ext != "png" && ext != "svg" && ext != "html" && ext != "txt" && ext != "epub") { path = file; readOnlyCopy = false; isNewDoc = false; dirty = false; NativeSettings.AddRecent(file); UpdateTitle(); }
                return true;
            }
            MessageBox.Show("Không lưu được: " + Lok.LastError(), "Lưu thành"); return false;
        }
        public void ExportPdf()
        {
            using (SaveFileDialog d = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", FileName = Path.GetFileNameWithoutExtension(path) + ".pdf", InitialDirectory = Path.GetDirectoryName(path) })
                if (d.ShowDialog(this) == DialogResult.OK) { Lok.Activate(doc); if (doc.SaveAs(d.FileName, "pdf", "")) { try { System.Diagnostics.Process.Start(d.FileName); } catch { } } else MessageBox.Show("Xuất PDF lỗi: " + Lok.LastError()); }
        }
        void ShareLan()
        {
            try
            {
                if (dirty) Save(true);
                Dictionary<string, object> r = LanShare.Share(Path.GetFileName(path), File.ReadAllBytes(path));
                string[] urls = r["urls"] as string[];
                string link = urls != null && urls.Length > 0 ? urls[0] : "(không có mạng)";
                try { Clipboard.SetText(link); } catch { }
                MessageBox.Show("Mở liên kết sau trên điện thoại / máy tính cùng mạng Wi-Fi để tải tệp:\n\n" + string.Join("\n", urls ?? new string[0]) + "\n\n(Đã sao chép liên kết. Nhận tệp từ máy khác: mở " + LanShareReceiveUrl(r) + ")", "Gửi qua Wi-Fi");
            }
            catch (Exception e) { MessageBox.Show("Không chia sẻ được: " + e.Message, "Chia sẻ"); }
        }
        static string LanShareReceiveUrl(Dictionary<string, object> r) { string[] a = r["receive"] as string[]; return a != null && a.Length > 0 ? a[0] : ""; }
        void AutoSaveCopy()
        {
            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOpenSource\\autosave"); Directory.CreateDirectory(dir);
                string ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
                Lok.Activate(doc); doc.SaveAs(Path.Combine(dir, Path.GetFileNameWithoutExtension(path) + ".autosave." + ext), ext, "");
            }
            catch { }
        }
        void OnClosing(object s, FormClosingEventArgs e)
        {
            if (dirty)
            {
                DialogResult r = MessageBox.Show("Bạn có muốn lưu các thay đổi đối với \"" + Path.GetFileName(path) + "\"?", DocKind, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (r == DialogResult.Cancel) { e.Cancel = true; return; }
                if (r == DialogResult.Yes) { if (readOnlyCopy || isNewDoc) { OpenBackstage("saveas"); e.Cancel = true; return; } Save(false); }
            }
            autosave.Stop(); liveSave.Stop();
            if (Lok.LoopRunning && App.OpenWindows.Count == 0) { int open = 0; foreach (Form f in Application.OpenForms) if (f is NativeDocForm && f != this) open++; if (open == 0) BeginInvoke(new Action(() => Lok.HardExit(0))); }
            canvas.Dispose();
            try { Lok.Forget(doc); doc.Destroy(); } catch { }
        }
        void ToggleFullScreen()
        {
            bool full = !title.Visible;
            title.Visible = full; status.Visible = full; ribbon.SetMinimized(!full);
            WindowState = full ? FormWindowState.Normal : FormWindowState.Maximized;
        }
        void ApplyOptions()
        {
            autosave.Interval = Math.Max(1, NativeSettings.Int("autosave.minutes", 2)) * 60000;
            if (rulerTop != null) rulerTop.Visible = rulerLeft.Visible = NativeSettings.Bool("word.ruler", true);
            title.Invalidate(); status.Invalidate();
        }

        // ---------------- search (Tell me) ----------------
        void SearchCommand(string q)
        {
            List<RMenu> m = new List<RMenu>();
            foreach (RTab t in ribbon.Tabs) foreach (RGroup g in t.Groups) foreach (RItem i in g.All())
                    {
                        string name = string.IsNullOrEmpty(i.Text) ? i.Tip : i.Text; if (string.IsNullOrEmpty(name)) continue;
                        if ((name + " " + i.Tip).IndexOf(q, StringComparison.CurrentCultureIgnoreCase) < 0) continue;
                        RItem ii = i;
                        if (i.Click != null) m.Add(new RMenu(name + "   (" + t.Name + ")", i.Icon, () => ii.Click()));
                        else if (i.Menu != null) foreach (RMenu sub in i.Menu) if (!sub.Separator) m.Add(new RMenu(name + " › " + sub.Text, i.Icon ?? sub.Icon, sub.Click));
                        if (m.Count > 12) break;
                    }
            if (m.Count > 0) m.Add(RMenu.Sep());
            m.Add(new RMenu("Tìm \"" + q + "\" trong tài liệu", "find", () => Find(q, false)));
            ribbon.ShowMenu(m, ribbon.PointToClient(title.PointToScreen(title.Below("search"))));
        }
    }

    /// <summary>Minimal empty OOXML documents used by "New" (Tệp → Mới).</summary>
    static class BlankDocs
    {
        public static byte[] Get(string ext)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (ZipWriter z = new ZipWriter(ms))
                {
                    if (ext == "docx")
                    {
                        z.Add("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/></Types>");
                        z.Add("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
                        z.Add("word/_rels/document.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
                        z.Add("word/styles.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\" w:eastAsia=\"Times New Roman\"/><w:sz w:val=\"26\"/><w:szCs w:val=\"26\"/><w:lang w:val=\"vi-VN\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=\"160\" w:line=\"259\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults>"
                            + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style>"
                            + "<w:style w:type=\"paragraph\" w:styleId=\"NoSpacing\"><w:name w:val=\"No Spacing\"/><w:qFormat/><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr></w:style>"
                            + "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before=\"360\" w:after=\"80\"/><w:outlineLvl w:val=\"0\"/></w:pPr><w:rPr><w:color w:val=\"2F5496\"/><w:sz w:val=\"32\"/></w:rPr></w:style>"
                            + "<w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before=\"160\" w:after=\"80\"/><w:outlineLvl w:val=\"1\"/></w:pPr><w:rPr><w:color w:val=\"2F5496\"/><w:sz w:val=\"28\"/></w:rPr></w:style>"
                            + "<w:style w:type=\"paragraph\" w:styleId=\"Title\"><w:name w:val=\"Title\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:qFormat/><w:pPr><w:spacing w:after=\"80\"/></w:pPr><w:rPr><w:sz w:val=\"56\"/></w:rPr></w:style>"
                            + "<w:style w:type=\"paragraph\" w:styleId=\"Subtitle\"><w:name w:val=\"Subtitle\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:qFormat/><w:rPr><w:color w:val=\"595959\"/><w:sz w:val=\"28\"/></w:rPr></w:style></w:styles>");
                        z.Add("word/document.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p/><w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:right=\"1134\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>");
                    }
                    else if (ext == "xlsx")
                    {
                        z.Add("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
                        z.Add("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                        z.Add("xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
                        z.Add("xl/workbook.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                        z.Add("xl/worksheets/sheet1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>");
                    }
                    else
                    {
                        // presentation: one 16:9 title slide
                        const string P = "http://schemas.openxmlformats.org/presentationml/2006/main", A = "http://schemas.openxmlformats.org/drawingml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", REL = "http://schemas.openxmlformats.org/package/2006/relationships", RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
                        z.Add("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/><Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/><Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/><Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/><Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>");
                        z.Add("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "officeDocument\" Target=\"ppt/presentation.xml\"/></Relationships>");
                        z.Add("ppt/_rels/presentation.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/><Relationship Id=\"rId2\" Type=\"" + RT + "slide\" Target=\"slides/slide1.xml\"/><Relationship Id=\"rId3\" Type=\"" + RT + "theme\" Target=\"theme/theme1.xml\"/></Relationships>");
                        z.Add("ppt/presentation.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:presentation xmlns:a=\"" + A + "\" xmlns:r=\"" + R + "\" xmlns:p=\"" + P + "\"><p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id=\"256\" r:id=\"rId2\"/></p:sldIdLst><p:sldSz cx=\"12192000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>");
                        string sp = "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Title 1\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"ctrTitle\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"1524000\" y=\"1122363\"/><a:ext cx=\"9144000\" cy=\"2387600\"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor=\"b\"/><a:lstStyle/><a:p><a:pPr algn=\"ctr\"/><a:endParaRPr lang=\"vi-VN\" sz=\"6000\"/></a:p></p:txBody></p:sp>"
                                  + "<p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Subtitle 2\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"subTitle\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"1524000\" y=\"3602038\"/><a:ext cx=\"9144000\" cy=\"1655762\"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn=\"ctr\"/><a:endParaRPr lang=\"vi-VN\" sz=\"2400\"/></a:p></p:txBody></p:sp>";
                        string tree = "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>";
                        z.Add("ppt/slides/slide1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"" + A + "\" xmlns:r=\"" + R + "\" xmlns:p=\"" + P + "\">" + tree + sp + "</p:spTree></p:cSld></p:sld>");
                        z.Add("ppt/slides/_rels/slide1.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/></Relationships>");
                        z.Add("ppt/slideLayouts/slideLayout1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sldLayout xmlns:a=\"" + A + "\" xmlns:r=\"" + R + "\" xmlns:p=\"" + P + "\" type=\"title\">" + tree + sp + "</p:spTree></p:cSld></p:sldLayout>");
                        z.Add("ppt/slideLayouts/_rels/slideLayout1.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/></Relationships>");
                        z.Add("ppt/slideMasters/slideMaster1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sldMaster xmlns:a=\"" + A + "\" xmlns:r=\"" + R + "\" xmlns:p=\"" + P + "\">" + tree + "</p:spTree></p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/><p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst></p:sldMaster>");
                        z.Add("ppt/slideMasters/_rels/slideMaster1.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/><Relationship Id=\"rId2\" Type=\"" + RT + "theme\" Target=\"../theme/theme1.xml\"/></Relationships>");
                        z.Add("ppt/theme/theme1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><a:theme xmlns:a=\"" + A + "\" name=\"Office\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2><a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4><a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"12700\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>");
                    }
                }
                return ms.ToArray();
            }
        }
    }

    /// <summary>Built-in Basic tools (application library "OOS") installed into the LibreOffice user profile; used for
    /// features LibreOfficeKit only offers through dialogs (conditional formatting, PivotTable, TOC, data tools…).</summary>
    static class BasicTools
    {
        public const string Url = "macro:///OOS.Tools.{0}";
        const string Hdr = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
        /// <summary>LibreOffice re-copies presets/basic into the profile, so the library goes into the engine's presets too
        /// (when writable: per-user Setup and portable builds; the build script pre-installs it for read-only installs).</summary>
        public static void Install(string userDir, string presetsBasic)
        {
            try { if (!string.IsNullOrEmpty(presetsBasic) && Directory.Exists(presetsBasic)) InstallInto(presetsBasic, false); } catch (Exception ex) { Lok.Trace("basic tools presets: " + ex.Message); }
            InstallInto(Path.Combine(userDir, "basic"), true);
        }
        static void InstallInto(string b, bool standard)
        {
            try
            {
                string oos = Path.Combine(b, "OOS"), std = Path.Combine(b, "Standard");
                Directory.CreateDirectory(oos); if (standard) Directory.CreateDirectory(std);
                string xba = Path.Combine(oos, "Tools.xba"), body = Hdr + "<!DOCTYPE script:module PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"module.dtd\">\n<script:module xmlns:script=\"http://openoffice.org/2000/script\" script:name=\"Tools\" script:language=\"StarBasic\">" + Module.Replace("\r\n", "\n") + "</script:module>";
                if (File.Exists(xba) && File.ReadAllText(xba) == body && File.Exists(Path.Combine(b, "script.xlc")) && File.ReadAllText(Path.Combine(b, "script.xlc")).Contains("\"OOS\"")) return;
                UTF8Encoding u = new UTF8Encoding(false);
                File.WriteAllText(xba, body, u);
                File.WriteAllText(Path.Combine(oos, "script.xlb"), Hdr + "<!DOCTYPE library:library PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"library.dtd\">\n<library:library xmlns:library=\"http://openoffice.org/2000/library\" library:name=\"OOS\" library:readonly=\"false\" library:passwordprotected=\"false\">\n <library:element library:name=\"Tools\"/>\n</library:library>", u);
                File.WriteAllText(Path.Combine(oos, "dialog.xlb"), Hdr + "<!DOCTYPE library:library PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"library.dtd\">\n<library:library xmlns:library=\"http://openoffice.org/2000/library\" library:name=\"OOS\" library:readonly=\"false\" library:passwordprotected=\"false\"/>", u);
                if (standard && !File.Exists(Path.Combine(std, "script.xlb")))
                {
                    File.WriteAllText(Path.Combine(std, "script.xlb"), Hdr + "<!DOCTYPE library:library PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"library.dtd\">\n<library:library xmlns:library=\"http://openoffice.org/2000/library\" library:name=\"Standard\" library:readonly=\"false\" library:passwordprotected=\"false\">\n <library:element library:name=\"Module1\"/>\n</library:library>", u);
                    File.WriteAllText(Path.Combine(std, "Module1.xba"), Hdr + "<!DOCTYPE script:module PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"module.dtd\">\n<script:module xmlns:script=\"http://openoffice.org/2000/script\" script:name=\"Module1\" script:language=\"StarBasic\">Sub Main\nEnd Sub\n</script:module>", u);
                }
                if (standard && !File.Exists(Path.Combine(std, "dialog.xlb")))
                    File.WriteAllText(Path.Combine(std, "dialog.xlb"), Hdr + "<!DOCTYPE library:library PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"library.dtd\">\n<library:library xmlns:library=\"http://openoffice.org/2000/library\" library:name=\"Standard\" library:readonly=\"false\" library:passwordprotected=\"false\"/>", u);
                Register(Path.Combine(b, "script.xlc"), "script.xlb", u);
                Register(Path.Combine(b, "dialog.xlc"), "dialog.xlb", u);
            }
            catch (Exception ex) { Lok.Trace("basic tools: " + ex.Message); }
        }
        static void Register(string xlc, string xlb, Encoding u)
        {
            string entry = " <library:library library:name=\"{0}\" library:link=\"false\"/>\n";
            string s = File.Exists(xlc) ? File.ReadAllText(xlc) : Hdr + "<!DOCTYPE library:libraries PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"libraries.dtd\">\n<library:libraries xmlns:library=\"http://openoffice.org/2000/library\" xmlns:xlink=\"http://www.w3.org/1999/xlink\">\n</library:libraries>";
            bool changed = false;
            foreach (string lib in new[] { "Standard", "OOS" })
                if (!s.Contains("library:name=\"" + lib + "\"")) { int i = s.LastIndexOf("</library:libraries>", StringComparison.Ordinal); if (i < 0) return; s = s.Insert(i, string.Format(entry, lib)); changed = true; }
            if (changed || !File.Exists(xlc)) File.WriteAllText(xlc, s, u);
        }
        public const string Module = @"REM Yomi-Office Open Source - built-in tools (application Basic library ""OOS"", module ""Tools"")
REM Arguments: %TEMP%\oos_macro.txt (UTF-8, key=value per line). Result: %TEMP%\oos_macro_out.txt
Option Explicit
Dim gFound As Boolean
Dim gTr As String
Dim pg1 As Object

Function Doc() As Object
	Doc = ThisComponent
	On Error Resume Next
	Dim c As Object
	c = StarDesktop.getCurrentComponent()
	If Not IsNull(c) Then
		If HasUnoInterfaces(c, ""com.sun.star.frame.XModel"") Then Doc = c
	End If
End Function

Function Arg(k As String) As String
	Dim sfa As Object, tis As Object, s As String, p As Integer, url As String
	Arg = """"
	url = ConvertToURL(Environ(""TEMP"") &amp; ""\oos_macro.txt"")
	sfa = CreateUnoService(""com.sun.star.ucb.SimpleFileAccess"")
	If Not sfa.exists(url) Then Exit Function
	tis = CreateUnoService(""com.sun.star.io.TextInputStream"")
	tis.setInputStream(sfa.openFileRead(url))
	tis.setEncoding(""UTF-8"")
	REM v0.1.7: read everything at once - isEOF() turns true while the last line is still buffered, so the last argument was lost
	Dim sAll As String, lines() As String, i As Integer
	sAll = tis.readString(Array(), False)
	tis.closeInput()
	lines = Split(sAll, Chr(10))
	For i = LBound(lines) To UBound(lines)
		s = lines(i)
		If Right(s, 1) = Chr(13) Then s = Left(s, Len(s) - 1)
		p = InStr(s, ""="")
		If p &gt; 1 Then
			If Left(s, p - 1) = k Then Arg = Mid(s, p + 1)
		End If
	Next i
End Function

Sub WriteOut(s As String)
	Dim sfa As Object, tos As Object, url As String
	url = ConvertToURL(Environ(""TEMP"") &amp; ""\oos_macro_out.txt"")
	sfa = CreateUnoService(""com.sun.star.ucb.SimpleFileAccess"")
	If sfa.exists(url) Then sfa.kill(url)
	tos = CreateUnoService(""com.sun.star.io.TextOutputStream"")
	tos.setOutputStream(sfa.openFileWrite(url))
	tos.setEncoding(""UTF-8"")
	tos.writeString(s)
	tos.closeOutput()
End Sub

Function ColName(ByVal c As Long) As String
	Dim s As String
	c = c + 1
	Do While c &gt; 0
		s = Chr(65 + ((c - 1) Mod 26)) &amp; s
		c = (c - 1) \ 26
	Loop
	ColName = s
End Function

REM current selection as one cell range; a single cell expands to its data region
Function SelRange(Optional expand As Boolean) As Object
	Dim d As Object, sel As Object, cur As Object
	d = Doc()
	sel = d.getCurrentSelection()
	SelRange = Nothing
	If IsNull(sel) Then Exit Function
	If Not sel.supportsService(""com.sun.star.sheet.SheetCellRange"") Then Exit Function
	If Not IsMissing(expand) Then
		If expand And sel.RangeAddress.StartColumn = sel.RangeAddress.EndColumn And sel.RangeAddress.StartRow = sel.RangeAddress.EndRow Then
			cur = sel.getSpreadsheet().createCursorByRange(sel)
			cur.collapseToCurrentRegion()
			sel = sel.getSpreadsheet().getCellRangeByPosition(cur.RangeAddress.StartColumn, cur.RangeAddress.StartRow, cur.RangeAddress.EndColumn, cur.RangeAddress.EndRow)
		End If
	End If
	SelRange = sel
End Function

Function AbsRef(sel As Object) As String
	Dim a As Object
	a = sel.RangeAddress
	AbsRef = ""$"" &amp; ColName(a.StartColumn) &amp; ""$"" &amp; (a.StartRow + 1) &amp; "":$"" &amp; ColName(a.EndColumn) &amp; ""$"" &amp; (a.EndRow + 1)
End Function

Function EnsureStyle(d As Object, nm As String, back As Long, fore As Long, bold As Boolean) As String
	Dim fam As Object, st As Object
	fam = d.getStyleFamilies().getByName(""CellStyles"")
	If Not fam.hasByName(nm) Then
		st = d.createInstance(""com.sun.star.style.CellStyle"")
		fam.insertByName(nm, st)
		If back &gt;= 0 Then
			st.CellBackColor = back
			st.IsCellBackgroundTransparent = False
		End If
		If fore &gt;= 0 Then st.CharColor = fore
		If bold Then st.CharWeight = 150
	End If
	EnsureStyle = nm
End Function

Function StyleFor(d As Object, k As String) As String
	Select Case k
		Case ""green"": StyleFor = EnsureStyle(d, ""OOS Green Fill"", RGB(198, 239, 206), RGB(0, 97, 0), False)
		Case ""yellow"": StyleFor = EnsureStyle(d, ""OOS Yellow Fill"", RGB(255, 235, 156), RGB(156, 87, 0), False)
		Case ""redtext"": StyleFor = EnsureStyle(d, ""OOS Red Text"", -1, RGB(156, 0, 6), False)
		Case ""lightred"": StyleFor = EnsureStyle(d, ""OOS Light Red Fill"", RGB(255, 199, 206), -1, False)
		Case ""bold"": StyleFor = EnsureStyle(d, ""OOS Bold"", -1, -1, True)
		Case Else: StyleFor = EnsureStyle(d, ""OOS Red Fill"", RGB(255, 199, 206), RGB(156, 0, 6), False)
	End Select
End Function

REM ---------- Calc: conditional formatting ----------
REM Rules are added as entries of one format per range (Excel-like stacking). A format created only through
REM XConditionalFormats is not painted until the document is reloaded, so a new format is created the classic way
REM (a never-true dummy condition through the ConditionalFormat property / ScDocFunc) and entries are added to it.
Function FindExact(sel As Object) As Object
	Dim all As Variant, i As Integer, r As Variant, a As Variant, s As Variant
	gFound = False
	s = sel.RangeAddress
	all = sel.getSpreadsheet().ConditionalFormats.getConditionalFormats()
	For i = LBound(all) To UBound(all)
		r = all(i).Range.getRangeAddresses()
		If UBound(r) = 0 Then
			a = r(0)
			If a.StartColumn = s.StartColumn And a.StartRow = s.StartRow And a.EndColumn = s.EndColumn And a.EndRow = s.EndRow Then
				FindExact = all(i)
				gFound = True
			End If
		End If
	Next i
End Function

Function EnsureCF(sel As Object) As Object
	Dim tc As Object, cf As Object
	Dim p(3) As New com.sun.star.beans.PropertyValue
	cf = FindExact(sel)
	If gFound Then
		EnsureCF = cf
		Exit Function
	End If
	tc = sel.ConditionalFormat
	p(0).Name = ""Operator""
	p(0).Value = com.sun.star.sheet.ConditionOperator.FORMULA
	p(1).Name = ""Formula1""
	p(1).Value = ""0""
	p(2).Name = ""StyleName""
	p(2).Value = ""Default""
	p(3).Name = ""SourcePosition""
	p(3).Value = sel.getCellByPosition(0, 0).getCellAddress()
	tc.addNew(p())
	sel.ConditionalFormat = tc
	EnsureCF = FindExact(sel)
End Function

Sub CondFormat
	On Error GoTo Fail
	Dim d As Object, sel As Object, cf As Object, e As Object, kind As String, op As Long, f1 As String, f2 As String, n As Integer
	d = Doc()
	sel = SelRange(False)
	If IsNull(sel) Then
		WriteOut(""ERR:select"")
		Exit Sub
	End If
	kind = Arg(""kind"")
	If kind = ""clear"" Then
		cf = sel.ConditionalFormat
		cf.clear()
		sel.ConditionalFormat = cf
		ClearScales(d, sel)
		WriteOut(""OK"")
		Exit Sub
	End If
	If kind = ""databar"" Or kind = ""scale2"" Or kind = ""scale3"" Then
		ScaleFormat(d, sel, kind)
		Exit Sub
	End If
	f1 = Arg(""v1"")
	f2 = Arg(""v2"")
	Select Case kind
		Case ""gt"": op = com.sun.star.sheet.ConditionFormatOperator.GREATER
		Case ""lt"": op = com.sun.star.sheet.ConditionFormatOperator.LESS
		Case ""eq"": op = com.sun.star.sheet.ConditionFormatOperator.EQUAL
		Case ""between"": op = com.sun.star.sheet.ConditionFormatOperator.BETWEEN
		Case ""text"": op = com.sun.star.sheet.ConditionFormatOperator.CONTAINS
		Case ""dup"": op = com.sun.star.sheet.ConditionFormatOperator.DUPLICATE
		Case ""unique"": op = com.sun.star.sheet.ConditionFormatOperator.UNIQUE
		Case ""top"": op = com.sun.star.sheet.ConditionFormatOperator.TOP_N_ELEMENTS
		Case ""bottom"": op = com.sun.star.sheet.ConditionFormatOperator.BOTTOM_N_ELEMENTS
		Case ""above"": op = com.sun.star.sheet.ConditionFormatOperator.ABOVE_AVERAGE
		Case ""below"": op = com.sun.star.sheet.ConditionFormatOperator.BELOW_AVERAGE
		Case Else
			WriteOut(""ERR:kind"")
			Exit Sub
	End Select
	cf = EnsureCF(sel)
	If Not gFound Then
		WriteOut(""ERR:cf"")
		Exit Sub
	End If
	n = cf.getCount()
	cf.createEntry(com.sun.star.sheet.ConditionEntryType.CONDITION, n)
	e = cf.getByIndex(n)
	e.Operator = op
	If f1 &lt;&gt; """" Then e.Formula1 = f1
	If f2 &lt;&gt; """" Then e.Formula2 = f2
	e.StyleName = StyleFor(d, Arg(""style""))
	WriteOut(""OK"")
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub ScaleFormat(d As Object, sel As Object, kind As String)
	On Error GoTo Fail
	Dim cf As Object, e As Object, ent As Variant, n As Integer
	cf = EnsureCF(sel)
	If Not gFound Then
		WriteOut(""ERR:cf"")
		Exit Sub
	End If
	n = cf.getCount()
	If kind = ""databar"" Then
		cf.createEntry(com.sun.star.sheet.ConditionEntryType.DATABAR, n)
		e = cf.getByIndex(n)
		ent = e.DataBarEntries
		ent(0).setType(com.sun.star.sheet.DataBarEntryType.DATABAR_MIN)
		ent(1).setType(com.sun.star.sheet.DataBarEntryType.DATABAR_MAX)
		e.Color = Val(Arg(""color""))
	ElseIf kind = ""scale2"" Then
		cf.createEntry(com.sun.star.sheet.ConditionEntryType.COLORSCALE2, n)
		e = cf.getByIndex(n)
		ent = e.ColorScaleEntries
		ent(0).setType(com.sun.star.sheet.ColorScaleEntryType.COLORSCALE_MIN)
		ent(0).setColor(RGB(255, 255, 255))
		ent(1).setType(com.sun.star.sheet.ColorScaleEntryType.COLORSCALE_MAX)
		ent(1).setColor(Val(Arg(""color"")))
	Else
		cf.createEntry(com.sun.star.sheet.ConditionEntryType.COLORSCALE3, n)
		e = cf.getByIndex(n)
		ent = e.ColorScaleEntries
		ent(0).setType(com.sun.star.sheet.ColorScaleEntryType.COLORSCALE_MIN)
		ent(0).setColor(RGB(248, 105, 107))
		ent(1).setType(com.sun.star.sheet.ColorScaleEntryType.COLORSCALE_PERCENTILE)
		ent(1).setFormula(""50"")
		ent(1).setColor(RGB(255, 235, 132))
		ent(2).setType(com.sun.star.sheet.ColorScaleEntryType.COLORSCALE_MAX)
		ent(2).setColor(RGB(99, 190, 123))
	End If
	WriteOut(""OK"")
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub ClearScales(d As Object, sel As Object)
	On Error Resume Next
	Dim cfs As Object, all As Variant, i As Integer, r As Variant, a As Variant, s As Variant, k As Integer, hit As Boolean
	s = sel.RangeAddress
	cfs = sel.getSpreadsheet().ConditionalFormats
	all = cfs.getConditionalFormats()
	For i = UBound(all) To LBound(all) Step -1
		r = all(i).Range.getRangeAddresses()
		hit = False
		For k = LBound(r) To UBound(r)
			a = r(k)
			If a.StartColumn &lt;= s.EndColumn And a.EndColumn &gt;= s.StartColumn And a.StartRow &lt;= s.EndRow And a.EndRow &gt;= s.StartRow Then hit = True
		Next k
		If hit Then cfs.removeByID(all(i).ID)
	Next i
End Sub

REM ---------- Calc: PivotTable ----------
Sub ListHeaders
	On Error GoTo Fail
	Dim sel As Object, a As Object, s As String, c As Long, t As String
	sel = SelRange(True)
	If IsNull(sel) Then
		WriteOut(""ERR:select"")
		Exit Sub
	End If
	a = sel.RangeAddress
	s = a.Sheet &amp; Chr(9) &amp; a.StartColumn &amp; Chr(9) &amp; a.StartRow &amp; Chr(9) &amp; a.EndColumn &amp; Chr(9) &amp; a.EndRow
	For c = 0 To a.EndColumn - a.StartColumn
		t = sel.getCellByPosition(c, 0).getString()
		If t = """" Then t = ""Cột "" &amp; ColName(a.StartColumn + c)
		If a.EndRow &gt; a.StartRow And sel.getCellByPosition(c, 1).getType() = com.sun.star.table.CellContentType.VALUE Then t = t &amp; Chr(9) &amp; ""1""
		s = s &amp; Chr(10) &amp; t
	Next c
	WriteOut(s)
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub PivotTable
	On Error GoTo Fail
	Dim d As Object, sheets As Object, src As New com.sun.star.table.CellRangeAddress, ns As Object, dp As Object, desc As Object, f As Object
	Dim nm As String, i As Integer, rowI As Integer, colI As Integer, dataI As Integer, fn As String
	d = Doc()
	sheets = d.getSheets()
	src.Sheet = Val(Arg(""sheet""))
	src.StartColumn = Val(Arg(""c1""))
	src.StartRow = Val(Arg(""r1""))
	src.EndColumn = Val(Arg(""c2""))
	src.EndRow = Val(Arg(""r2""))
	rowI = Val(Arg(""row""))
	colI = Val(Arg(""col""))
	dataI = Val(Arg(""data""))
	fn = Arg(""func"")
	i = 1
	nm = Arg(""name"")
	If nm = """" Then nm = ""PivotTable""
	Do While sheets.hasByName(nm &amp; i)
		i = i + 1
	Loop
	nm = nm &amp; i
	sheets.insertNewByName(nm, src.Sheet)
	src.Sheet = src.Sheet + 1
	ns = sheets.getByName(nm)
	dp = ns.getDataPilotTables()
	desc = dp.createDataPilotDescriptor()
	desc.setSourceRange(src)
	f = desc.getDataPilotFields().getByIndex(rowI)
	f.Orientation = com.sun.star.sheet.DataPilotFieldOrientation.ROW
	If colI &gt;= 0 And colI &lt;&gt; rowI And colI &lt;&gt; dataI Then
		f = desc.getDataPilotFields().getByIndex(colI)
		f.Orientation = com.sun.star.sheet.DataPilotFieldOrientation.COLUMN
	End If
	f = desc.getDataPilotFields().getByIndex(dataI)
	f.Orientation = com.sun.star.sheet.DataPilotFieldOrientation.DATA
	Select Case fn
		Case ""COUNT"": f.Function = com.sun.star.sheet.GeneralFunction.COUNT
		Case ""AVERAGE"": f.Function = com.sun.star.sheet.GeneralFunction.AVERAGE
		Case ""MAX"": f.Function = com.sun.star.sheet.GeneralFunction.MAX
		Case ""MIN"": f.Function = com.sun.star.sheet.GeneralFunction.MIN
		Case Else: f.Function = com.sun.star.sheet.GeneralFunction.SUM
	End Select
	dp.insertNewByName(nm, ns.getCellByPosition(0, 2).getCellAddress(), desc)
	ns.getCellByPosition(0, 0).setString(Arg(""title""))
	d.getCurrentController().setActiveSheet(ns)
	WriteOut(""OK"" &amp; Chr(9) &amp; (src.Sheet - 1))
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub RefreshPivots
	On Error GoTo Fail
	Dim d As Object, sh As Object, dp As Object, i As Integer, k As Integer, n As Integer
	d = Doc()
	For i = 0 To d.getSheets().getCount() - 1
		dp = d.getSheets().getByIndex(i).getDataPilotTables()
		For k = 0 To dp.getCount() - 1
			dp.getByIndex(k).refresh()
			n = n + 1
		Next k
	Next i
	WriteOut(""OK"" &amp; Chr(9) &amp; n)
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

REM ---------- Calc: data tools ----------
Sub RemoveDuplicates
	On Error GoTo Fail
	Dim sel As Object, data As Variant, outRows() As Variant, keys As Object, r As Long, c As Long, k As String, n As Long, first As Long, w As Long, blank() As Variant, removed As Long
	sel = SelRange(True)
	If IsNull(sel) Then
		WriteOut(""ERR:select"")
		Exit Sub
	End If
	data = sel.getDataArray()
	keys = com.sun.star.container.EnumerableMap.create(""string"", ""boolean"")
	first = 0
	If Arg(""header"") = ""1"" Then first = 1
	ReDim outRows(UBound(data))
	n = 0
	For r = 0 To UBound(data)
		k = """"
		For c = 0 To UBound(data(r))
			k = k &amp; Chr(1) &amp; CStr(data(r)(c))
		Next c
		If r &lt; first Then
			outRows(n) = data(r)
			n = n + 1
		ElseIf Not keys.containsKey(k) Then
			keys.put(k, True)
			outRows(n) = data(r)
			n = n + 1
		End If
	Next r
	removed = UBound(data) + 1 - n
	w = UBound(data(0))
	ReDim blank(w)
	For c = 0 To w
		blank(c) = """"
	Next c
	For r = n To UBound(data)
		outRows(r) = blank
	Next r
	If removed &gt; 0 Then
		sel.getCellRangeByPosition(0, n, w, UBound(data)).clearContents(com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.FORMULA + com.sun.star.sheet.CellFlags.DATETIME)
		ReDim Preserve outRows(n - 1)
		sel.getCellRangeByPosition(0, 0, w, n - 1).setDataArray(outRows)
	End If
	WriteOut(""OK"" &amp; Chr(9) &amp; removed &amp; Chr(9) &amp; n)
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub TextToColumns
	On Error GoTo Fail
	Dim sel As Object, sh As Object, a As Object, r As Long, parts As Variant, k As Integer, dl As String, s As String, n As Long, cell As Object
	sel = SelRange(False)
	If IsNull(sel) Then
		WriteOut(""ERR:select"")
		Exit Sub
	End If
	dl = Arg(""delim"")
	If dl = ""TAB"" Then dl = Chr(9)
	If dl = ""SPACE"" Then dl = "" ""
	sh = sel.getSpreadsheet()
	a = sel.RangeAddress
	For r = a.StartRow To a.EndRow
		s = sh.getCellByPosition(a.StartColumn, r).getString()
		If InStr(s, dl) &gt; 0 Then
			parts = Split(s, dl)
			For k = 0 To UBound(parts)
				cell = sh.getCellByPosition(a.StartColumn + k, r)
				If IsNumeric(parts(k)) And Trim(parts(k)) &lt;&gt; """" Then
					cell.setFormula(Trim(parts(k)))
				Else
					cell.setString(parts(k))
				End If
			Next k
			n = n + 1
		End If
	Next r
	WriteOut(""OK"" &amp; Chr(9) &amp; n)
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub Validation
	On Error GoTo Fail
	Dim sel As Object, v As Object, kind As String
	sel = SelRange(False)
	If IsNull(sel) Then
		WriteOut(""ERR:select"")
		Exit Sub
	End If
	kind = Arg(""kind"")
	v = sel.Validation
	If kind = ""clear"" Then
		v.Type = com.sun.star.sheet.ValidationType.ANY
		v.ShowErrorMessage = False
	ElseIf kind = ""list"" Then
		v.Type = com.sun.star.sheet.ValidationType.LIST
		v.ShowList = 1
		v.setFormula1(Arg(""v1""))
		v.ShowErrorMessage = True
		v.ErrorAlertStyle = com.sun.star.sheet.ValidationAlertStyle.STOP
	Else
		If kind = ""decimal"" Then
			v.Type = com.sun.star.sheet.ValidationType.DECIMAL
		ElseIf kind = ""date"" Then
			v.Type = com.sun.star.sheet.ValidationType.DATE
		Else
			v.Type = com.sun.star.sheet.ValidationType.WHOLE
		End If
		v.setOperator(com.sun.star.sheet.ConditionOperator.BETWEEN)
		v.setFormula1(Arg(""v1""))
		v.setFormula2(Arg(""v2""))
		v.ShowErrorMessage = True
		v.ErrorAlertStyle = com.sun.star.sheet.ValidationAlertStyle.STOP
	End If
	v.ErrorTitle = ""Yomi-Office Open Source""
	v.ErrorMessage = Arg(""msg"")
	sel.Validation = v
	WriteOut(""OK"")
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

REM ---------- Writer: table of contents / captions ----------
Sub InsertTOC
	On Error GoTo Fail
	Dim d As Object, vc As Object, idx As Object, lv As Integer, t As String
	d = Doc()
	vc = d.getCurrentController().getViewCursor()
	idx = d.createInstance(""com.sun.star.text.ContentIndex"")
	lv = Val(Arg(""levels""))
	If lv &lt; 1 Then lv = 3
	idx.CreateFromOutline = True
	idx.Level = lv
	t = Arg(""title"")
	idx.Title = t
	vc.getText().insertTextContent(vc.getStart(), idx, False)
	idx.update()
	WriteOut(""OK"")
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub UpdateTOC
	On Error GoTo Fail
	Dim d As Object, i As Integer, n As Integer
	d = Doc()
	For i = 0 To d.getDocumentIndexes().getCount() - 1
		d.getDocumentIndexes().getByIndex(i).update()
		n = n + 1
	Next i
	WriteOut(""OK"" &amp; Chr(9) &amp; n)
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub RemoveTOC
	On Error GoTo Fail
	Dim d As Object, i As Integer, n As Integer, x As Object
	d = Doc()
	For i = d.getDocumentIndexes().getCount() - 1 To 0 Step -1
		x = d.getDocumentIndexes().getByIndex(i)
		If x.supportsService(""com.sun.star.text.ContentIndex"") Then
			x.dispose()
			n = n + 1
		End If
	Next i
	WriteOut(""OK"" &amp; Chr(9) &amp; n)
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub InsertCaption
	On Error GoTo Fail
	Dim d As Object, vc As Object, txt As Object, cat As String, fm As Object, num As Object
	d = Doc()
	vc = d.getCurrentController().getViewCursor()
	txt = vc.getText()
	cat = Arg(""cat"")
	If cat = """" Then cat = ""Hình""
	vc.gotoEndOfLine(False)
	txt.insertControlCharacter(vc.getEnd(), com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
	txt.insertString(vc.getEnd(), cat &amp; "" "", False)
	num = d.createInstance(""com.sun.star.text.TextField.SetExpression"")
	fm = d.getTextFieldMasters()
	Dim mName As String
	mName = ""com.sun.star.text.fieldmaster.SetExpression."" &amp; cat
	Dim master As Object
	If fm.hasByName(mName) Then
		master = fm.getByName(mName)
	Else
		master = d.createInstance(""com.sun.star.text.fieldmaster.SetExpression"")
		master.Name = cat
		master.SubType = com.sun.star.text.SetVariableType.SEQUENCE
	End If
	num.attachTextFieldMaster(master)
	num.NumberingType = com.sun.star.style.NumberingType.ARABIC
	num.Content = cat &amp; ""+1""
	txt.insertTextContent(vc.getEnd(), num, False)
	txt.insertString(vc.getEnd(), "": "" &amp; Arg(""text""), False)
	On Error Resume Next
	vc.ParaStyleName = ""Caption""
	d.getTextFields().refresh()
	WriteOut(""OK"")
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

REM ---------- Equations (LibreOffice Math object, StarMath syntax) ----------
Sub InsertFormula
	On Error GoTo Fail
	Dim d As Object, f As String, obj As Object, vc As Object, page As Object, sh As Object, sel As Object
	Dim pos As New com.sun.star.awt.Point, sz As New com.sun.star.awt.Size
	d = Doc()
	f = Arg(""formula"")
	If d.supportsService(""com.sun.star.text.TextDocument"") Then
		obj = d.createInstance(""com.sun.star.text.TextEmbeddedObject"")
		obj.CLSID = ""078B7ABA-54FC-457F-8551-6147e776a997""
		obj.AnchorType = com.sun.star.text.TextContentAnchorType.AS_CHARACTER
		vc = d.getCurrentController().getViewCursor()
		vc.getText().insertTextContent(vc.getEnd(), obj, False)
		obj.getEmbeddedObject().Formula = f
		On Error Resume Next
		obj.ExtendedControlOverEmbeddedObject.update()
	Else
		If d.supportsService(""com.sun.star.sheet.SpreadsheetDocument"") Then
			page = d.getCurrentController().getActiveSheet().getDrawPage()
			sel = d.getCurrentSelection()
			On Error Resume Next
			pos = sel.Position
			On Error GoTo Fail
		Else
			page = d.getCurrentController().getCurrentPage()
			pos.X = 2000
			pos.Y = 2000
		End If
		sh = d.createInstance(""com.sun.star.drawing.OLE2Shape"")
		page.add(sh)
		sh.CLSID = ""078B7ABA-54FC-457F-8551-6147e776a997""
		sh.Model.Formula = f
		sh.setPosition(pos)
		REM natural size of the formula (VisibleArea of the Math object), else an estimate from its length
		Dim va As Variant
		sz.Width = Len(f) * 170
		If sz.Width &lt; 2000 Then sz.Width = 2000
		If sz.Width &gt; 14000 Then sz.Width = 14000
		sz.Height = 1100
		On Error Resume Next
		va = sh.VisibleArea
		If Not IsNull(va) Then
			If va.Width &gt; 200 And va.Height &gt; 200 Then
				sz.Width = va.Width
				sz.Height = va.Height
			End If
		End If
		On Error GoTo Fail
		sh.setSize(sz)
	End If
	WriteOut(""OK"")
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

REM ===== v0.1.7: PowerPoint animations (written into the file as real animation nodes via EffectMigration) =====
Function AnimEffect(n As String) As Long
	Select Case n
		Case ""appear"": AnimEffect = com.sun.star.presentation.AnimationEffect.APPEAR
		Case ""dissolve"": AnimEffect = com.sun.star.presentation.AnimationEffect.DISSOLVE
		Case ""flyleft"": AnimEffect = com.sun.star.presentation.AnimationEffect.MOVE_FROM_LEFT
		Case ""flyright"": AnimEffect = com.sun.star.presentation.AnimationEffect.MOVE_FROM_RIGHT
		Case ""flytop"": AnimEffect = com.sun.star.presentation.AnimationEffect.MOVE_FROM_TOP
		Case ""flybottom"": AnimEffect = com.sun.star.presentation.AnimationEffect.MOVE_FROM_BOTTOM
		Case ""wipe"": AnimEffect = com.sun.star.presentation.AnimationEffect.FADE_FROM_LEFT
		Case ""zoom"": AnimEffect = com.sun.star.presentation.AnimationEffect.ZOOM_IN_FROM_CENTER
		Case ""spin"": AnimEffect = com.sun.star.presentation.AnimationEffect.CLOCKWISE
		Case ""stripes"": AnimEffect = com.sun.star.presentation.AnimationEffect.VERTICAL_STRIPES
		Case ""checker"": AnimEffect = com.sun.star.presentation.AnimationEffect.HORIZONTAL_CHECKERBOARD
		Case ""fadecenter"": AnimEffect = com.sun.star.presentation.AnimationEffect.FADE_FROM_CENTER
		Case Else: AnimEffect = com.sun.star.presentation.AnimationEffect.NONE
	End Select
End Function

Function AnimSpeed(s As String) As Long
	If s = ""slow"" Then
		AnimSpeed = com.sun.star.presentation.AnimationSpeed.SLOW
	ElseIf s = ""fast"" Then
		AnimSpeed = com.sun.star.presentation.AnimationSpeed.FAST
	Else
		AnimSpeed = com.sun.star.presentation.AnimationSpeed.MEDIUM
	End If
End Function

Sub Animate
	On Error GoTo Fail
	Dim d As Object, sel As Object, i As Long, n As Long, sh As Object, eff As String
	d = Doc()
	eff = Arg(""effect"")
	sel = d.getCurrentController().getSelection()
	If IsNull(sel) Then
		WriteOut(""ERR:select"")
		Exit Sub
	End If
	If Not HasUnoInterfaces(sel, ""com.sun.star.container.XIndexAccess"") Then
		WriteOut(""ERR:select"")
		Exit Sub
	End If
	For i = 0 To sel.getCount() - 1
		sh = sel.getByIndex(i)
		sh.Effect = AnimEffect(eff)
		sh.Speed = AnimSpeed(Arg(""speed""))
		n = n + 1
	Next i
	If n = 0 Then
		WriteOut(""ERR:select"")
	Else
		WriteOut(""OK "" &amp; n)
	End If
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub ListAnimations
	On Error GoTo Fail
	Dim d As Object, pg As Object, i As Long, sh As Object, s As String, nm As String, k As Long
	d = Doc()
	pg = d.getCurrentController().getCurrentPage()
	For i = 0 To pg.getCount() - 1
		sh = pg.getByIndex(i)
		On Error Resume Next
		k = 0
		k = sh.Effect
		On Error GoTo Fail
		If k &lt;&gt; com.sun.star.presentation.AnimationEffect.NONE Then
			nm = sh.Name
			If nm = """" Then nm = ""Đối tượng "" &amp; (i + 1)
			If HasUnoInterfaces(sh, ""com.sun.star.text.XText"") Then
				If Len(sh.getString()) &gt; 0 Then nm = nm &amp; "": "" &amp; Left(Replace(sh.getString(), Chr(10), "" ""), 40)
			End If
			s = s &amp; i &amp; ""|"" &amp; k &amp; ""|"" &amp; nm &amp; Chr(10)
		End If
	Next i
	WriteOut(""OK"" &amp; Chr(10) &amp; s)
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub ClearAnimations
	On Error GoTo Fail
	Dim d As Object, pg As Object, i As Long, sh As Object, n As Long
	d = Doc()
	pg = d.getCurrentController().getCurrentPage()
	For i = 0 To pg.getCount() - 1
		sh = pg.getByIndex(i)
		On Error Resume Next
		If sh.Effect &lt;&gt; com.sun.star.presentation.AnimationEffect.NONE Then
			sh.Effect = com.sun.star.presentation.AnimationEffect.NONE
			n = n + 1
		End If
		On Error GoTo Fail
	Next i
	WriteOut(""OK "" &amp; n)
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

REM ===== v0.1.7: slide transitions stored in the file (PowerPoint keeps them) =====
Sub SetTransitionOn(pg As Object, t As String, dur As Double)
	Dim tt As Integer, st As Integer
	tt = 0
	st = 0
	Select Case t
		Case ""fade"": tt = com.sun.star.animations.TransitionType.FADE: st = com.sun.star.animations.TransitionSubType.CROSSFADE
		Case ""push"": tt = com.sun.star.animations.TransitionType.PUSHWIPE: st = com.sun.star.animations.TransitionSubType.FROMRIGHT
		Case ""wipe"": tt = com.sun.star.animations.TransitionType.BARWIPE: st = com.sun.star.animations.TransitionSubType.LEFTTORIGHT
		Case ""zoom"": tt = com.sun.star.animations.TransitionType.IRISWIPE: st = com.sun.star.animations.TransitionSubType.RECTANGLE
		Case ""split"": tt = com.sun.star.animations.TransitionType.BARNDOORWIPE: st = com.sun.star.animations.TransitionSubType.VERTICAL
	End Select
	On Error Resume Next
	pg.setPropertyValue(""TransitionType"", tt)
	pg.setPropertyValue(""TransitionSubtype"", st)
	pg.setPropertyValue(""TransitionDirection"", True)
	If dur &lt;= 0.5 Then
		pg.setPropertyValue(""Speed"", com.sun.star.presentation.AnimationSpeed.FAST)
	ElseIf dur &gt;= 1 Then
		pg.setPropertyValue(""Speed"", com.sun.star.presentation.AnimationSpeed.SLOW)
	Else
		pg.setPropertyValue(""Speed"", com.sun.star.presentation.AnimationSpeed.MEDIUM)
	End If
	pg.setPropertyValue(""TransitionDuration"", dur)
	gTr = gTr &amp; "" "" &amp; pg.getPropertyValue(""TransitionType"") &amp; ""/"" &amp; pg.getPropertyValue(""TransitionSubtype"") &amp; ""/"" &amp; pg.getPropertyValue(""TransitionDuration"")
End Sub

Sub SlideTransition
	On Error GoTo Fail
	gTr = """"
	Dim d As Object, pages As Object, i As Long, t As String, dur As Double
	d = Doc()
	t = Arg(""type"")
	dur = Val(Arg(""dur""))
	If dur &lt;= 0 Then dur = 0.7
	If Arg(""all"") = ""1"" Then
		Dim dp As Object, cnt As Long
		dp = d.getDrawPages()
		cnt = dp.getCount()
		For i = 0 To cnt - 1
			pg1 = dp.getByIndex(i)
			SetTransitionOn(pg1, t, dur)
		Next i
	Else
		SetTransitionOn(d.getCurrentController().getCurrentPage(), t, dur)
	End If
	WriteOut(""OK"" &amp; gTr)
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

REM ===== v0.1.7: Mailings (Word) - the recipient list is read by the app; one call per record (no second document is opened in the engine)
Sub InsertMergeField
	On Error GoTo Fail
	Dim d As Object, vc As Object
	d = Doc()
	vc = d.getCurrentController().getViewCursor()
	vc.getText().insertString(vc, Chr(171) &amp; Arg(""field"") &amp; Chr(187), False)
	WriteOut(""OK"")
	Exit Sub
Fail:
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub

Sub MergeOne
	On Error GoTo Fail
	Dim d As Object, um As Object, rd As Object, i As Integer, nm As String, n As Long
	d = Doc()
	um = d.getUndoManager()
	um.enterUndoContext(""OOS merge"")
	For i = 0 To 63
		nm = Arg(""f"" &amp; i)
		If nm = """" Then Exit For
		rd = d.createReplaceDescriptor()
		rd.setSearchString(Chr(171) &amp; nm &amp; Chr(187))
		rd.setReplaceString(Arg(""v"" &amp; i))
		n = n + d.replaceAll(rd)
	Next i
	um.leaveUndoContext()
	Dim sp(0) As New com.sun.star.beans.PropertyValue
	sp(0).Name = ""FilterName""
	If Arg(""format"") = ""pdf"" Then
		sp(0).Value = ""writer_pdf_Export""
	Else
		sp(0).Value = ""MS Word 2007 XML""
	End If
	d.storeToURL(ConvertToURL(Arg(""out"")), sp())
	If n &gt; 0 Then um.undo()
	WriteOut(""OK "" &amp; n)
	Exit Sub
Fail:
	On Error Resume Next
	um.leaveUndoContext()
	WriteOut(""ERR:basic "" &amp; Erl &amp; "" "" &amp; Error$)
End Sub
";
    }

    /// <summary>Reads the first section's top/bottom page margins (twips) from .docx / .odt files (no ZipArchive on .NET 4.0).</summary>
    static class DocMargins
    {
        public static int[] Read(string file)
        {
            try
            {
                if (string.IsNullOrEmpty(file) || !File.Exists(file)) return null;
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".docx" || ext == ".docm" || ext == ".dotx" || ext == ".dotm")
                {
                    string x = Entry(file, "word/document.xml"); if (x == null) return null;
                    Match m = Regex.Match(x, "<w:pgMar\\b[^>]*>"); if (!m.Success) return null;
                    int t = Attr(m.Value, "w:top", 1440), b = Attr(m.Value, "w:bottom", 1440);
                    return new int[] { Math.Abs(t), Math.Abs(b) };
                }
                if (ext == ".odt" || ext == ".ott")
                {
                    string x = Entry(file, "styles.xml"); if (x == null) return null;
                    Match mp = Regex.Match(x, "<style:master-page\\b[^>]*style:page-layout-name=\"([^\"]+)\"");
                    Match pl = mp.Success ? Regex.Match(x, "<style:page-layout\\b[^>]*style:name=\"" + Regex.Escape(mp.Groups[1].Value) + "\"[\\s\\S]*?<style:page-layout-properties\\b[^>]*>") : Match.Empty;
                    if (!pl.Success) pl = Regex.Match(x, "<style:page-layout-properties\\b[^>]*>");
                    if (!pl.Success) return null;
                    Match mt = Regex.Match(pl.Value, "<style:page-layout-properties\\b[^>]*>");
                    string props = mt.Success ? mt.Groups[0].Value : pl.Value;
                    props = props.Substring(props.LastIndexOf("<style:page-layout-properties", StringComparison.Ordinal));
                    return new int[] { Len(props, "fo:margin-top"), Len(props, "fo:margin-bottom") };
                }
            }
            catch (Exception ex) { Lok.Trace("ruler margins: " + ex.Message); }
            return null;
        }
        static int Attr(string tag, string name, int def)
        {
            Match m = Regex.Match(tag, name + "=\"(-?\\d+)\""); int v; return m.Success && int.TryParse(m.Groups[1].Value, out v) ? v : def;
        }
        static int Len(string tag, string name)
        {
            Match m = Regex.Match(tag, name + "=\"([\\d.]+)(cm|mm|in|pt|pc|px)?\""); if (!m.Success) return 0;
            double v = double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            switch (m.Groups[2].Value) { case "mm": v *= 1440 / 25.4; break; case "in": v *= 1440; break; case "pt": v *= 20; break; case "pc": v *= 240; break; case "px": v *= 15; break; default: v *= 1440 / 2.54; break; }
            return (int)Math.Round(v);
        }
        /// <summary>Extracts one entry (stored or deflated) as UTF-8 text using the central directory.</summary>
        internal static string Entry(string zip, string name)
        {
            byte[] z; using (FileStream fs = new FileStream(zip, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) { z = new byte[fs.Length]; int o = 0, r; while (o < z.Length && (r = fs.Read(z, o, z.Length - o)) > 0) o += r; }
            int eocd = -1; for (int i = z.Length - 22; i >= Math.Max(0, z.Length - 65558); i--) if (BitConverter.ToUInt32(z, i) == 0x06054b50) { eocd = i; break; }
            if (eocd < 0) return null;
            int n = BitConverter.ToUInt16(z, eocd + 10), p = (int)BitConverter.ToUInt32(z, eocd + 16);
            for (int k = 0; k < n && p + 46 <= z.Length; k++)
            {
                if (BitConverter.ToUInt32(z, p) != 0x02014b50) return null;
                int method = BitConverter.ToUInt16(z, p + 10), csize = (int)BitConverter.ToUInt32(z, p + 20), usize = (int)BitConverter.ToUInt32(z, p + 24);
                int fl = BitConverter.ToUInt16(z, p + 28), el = BitConverter.ToUInt16(z, p + 30), cl = BitConverter.ToUInt16(z, p + 32), lo = (int)BitConverter.ToUInt32(z, p + 42);
                string en = Encoding.UTF8.GetString(z, p + 46, fl);
                if (en == name)
                {
                    int d = lo + 30 + BitConverter.ToUInt16(z, lo + 26) + BitConverter.ToUInt16(z, lo + 28);
                    if (method == 0) return Encoding.UTF8.GetString(z, d, csize);
                    using (MemoryStream ms = new MemoryStream(z, d, csize))
                    using (System.IO.Compression.DeflateStream ds = new System.IO.Compression.DeflateStream(ms, System.IO.Compression.CompressionMode.Decompress))
                    using (MemoryStream o = new MemoryStream(Math.Max(usize, 16))) { byte[] buf = new byte[65536]; int r; while ((r = ds.Read(buf, 0, buf.Length)) > 0) o.Write(buf, 0, r); return Encoding.UTF8.GetString(o.ToArray()); }
                }
                p += 46 + fl + el + cl;
            }
            return null;
        }
    }

    /// <summary>Tiny ZIP (stored, no compression) writer for the blank templates; .NET 4.0 has no ZipArchive.</summary>
    class ZipWriter : IDisposable
    {
        readonly Stream s; readonly List<byte[]> central = new List<byte[]>(); int count;
        public ZipWriter(Stream st) { s = st; }
        static uint Crc(byte[] d) { uint c = 0xFFFFFFFF; foreach (byte b in d) { c ^= b; for (int k = 0; k < 8; k++) c = (c & 1) != 0 ? (c >> 1) ^ 0xEDB88320 : c >> 1; } return ~c; }
        public void Add(string name, string text)
        {
            byte[] data = new UTF8Encoding(false).GetBytes(text), nm = Encoding.UTF8.GetBytes(name); uint crc = Crc(data); long off = s.Position;
            BinaryWriter w = new BinaryWriter(s);
            w.Write(0x04034b50); w.Write((short)20); w.Write((short)0x0800); w.Write((short)0); w.Write((short)0); w.Write((short)0x21); w.Write(crc); w.Write(data.Length); w.Write(data.Length); w.Write((short)nm.Length); w.Write((short)0); w.Write(nm); w.Write(data); w.Flush();
            using (MemoryStream c = new MemoryStream()) using (BinaryWriter cw = new BinaryWriter(c))
            { cw.Write(0x02014b50); cw.Write((short)20); cw.Write((short)20); cw.Write((short)0x0800); cw.Write((short)0); cw.Write((short)0); cw.Write((short)0x21); cw.Write(crc); cw.Write(data.Length); cw.Write(data.Length); cw.Write((short)nm.Length); cw.Write((short)0); cw.Write((short)0); cw.Write((short)0); cw.Write((short)0); cw.Write(0); cw.Write((int)off); cw.Write(nm); cw.Flush(); central.Add(c.ToArray()); }
            count++;
        }
        public void Dispose()
        {
            long start = s.Position; int size = 0; foreach (byte[] c in central) { s.Write(c, 0, c.Length); size += c.Length; }
            BinaryWriter w = new BinaryWriter(s); w.Write(0x06054b50); w.Write((short)0); w.Write((short)0); w.Write((short)count); w.Write((short)count); w.Write(size); w.Write((int)start); w.Write((short)0); w.Flush();
        }
    }
}
