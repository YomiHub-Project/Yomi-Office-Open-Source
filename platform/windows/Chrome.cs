// Office Open Source v0.1.6 - title bar, status bar, Excel formula bar / headers / sheet tabs, PowerPoint slide pane
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace OfficeOpenSource
{
    /// <summary>Custom title bar: app badge, AutoSave switch, Quick Access Toolbar, document name, search, account, caption buttons.</summary>
    class TitleBar : Control
    {
        public Color Accent; public string AppIcon = "word";
        public string DocName = "", DocState = "";
        public bool AutoSave; public event Action AutoSaveToggled, SaveClicked, UndoClicked, RedoClicked, DocNameClicked, AccountClicked, QatClicked;
        public bool UndoEnabled = true, RedoEnabled = true;
        public readonly TextBox Search = new TextBox { BorderStyle = BorderStyle.None, Font = Ui.F(9.5f) };
        public event Action<string> SearchSubmitted;
        Rectangle rIcon, rAuto, rSave, rUndo, rUndoArrow, rRedo, rQat, rName, rSearch, rAccount, rMin, rMax, rClose;
        string hot; const string Placeholder = "Tìm kiếm (Alt + Q)";

        public TitleBar()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Dock = DockStyle.Top; Height = 40; BackColor = Ui.Chrome; Font = Ui.F(9f);
            Search.Text = Placeholder; Search.ForeColor = Ui.Sub; Search.BackColor = Color.White;
            Search.GotFocus += (s, e) => { if (Search.Text == Placeholder) { Search.Text = ""; Search.ForeColor = Ui.Text; } Invalidate(); };
            Search.LostFocus += (s, e) => { if (Search.Text.Length == 0) { Search.Text = Placeholder; Search.ForeColor = Ui.Sub; } Invalidate(); };
            Search.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter && Search.Text.Trim().Length > 0) { e.SuppressKeyPress = true; if (SearchSubmitted != null) SearchSubmitted(Search.Text.Trim()); } if (e.KeyCode == Keys.Escape) { Search.Text = ""; Parent.SelectNextControl(this, true, true, true, true); } };
            Controls.Add(Search);
        }

        Form Owner { get { return FindForm(); } }

        protected override void OnResize(EventArgs e) { base.OnResize(e); LayoutParts(); }
        void LayoutParts()
        {
            int x = 10;
            rIcon = new Rectangle(x, 10, 20, 20); x += 30;
            int aw = Ui.Measure("Tự động lưu", Font).Width + 66; rAuto = new Rectangle(x, 8, aw, 24); x += aw + 6;
            rSave = new Rectangle(x, 6, 28, 28); x += 30;
            rUndo = new Rectangle(x, 6, 28, 28); rUndoArrow = new Rectangle(x + 28, 6, 12, 28); x += 42;
            rRedo = new Rectangle(x, 6, 28, 28); x += 30;
            rQat = new Rectangle(x, 6, 18, 28); x += 26;
            int sw = Math.Min(520, Math.Max(200, Width / 3)); rSearch = new Rectangle((Width - sw) / 2, 6, sw, 28);
            if (rSearch.X < x + 180) rSearch.X = x + 180;
            rClose = new Rectangle(Width - 46, 0, 46, Height); rMax = new Rectangle(rClose.X - 46, 0, 46, Height); rMin = new Rectangle(rMax.X - 46, 0, 46, Height);
            rAccount = new Rectangle(rMin.X - 44, 6, 28, 28);
            if (rSearch.Right > rAccount.X - 12) rSearch.Width = Math.Max(120, rAccount.X - 12 - rSearch.X);
            rName = new Rectangle(x, 6, Math.Max(40, rSearch.X - x - 12), 28);
            Search.Bounds = new Rectangle(rSearch.X + 32, rSearch.Y + 6, rSearch.Width - 40, 18);
        }

        public bool IsDragArea(Point p)
        {
            foreach (Rectangle r in new[] { rIcon, rAuto, rSave, rUndo, rUndoArrow, rRedo, rQat, rSearch, rAccount, rMin, rMax, rClose }) if (r.Contains(p)) return false;
            if (rName.Contains(p) && p.X < rName.X + NameWidth() + 20) return false;
            return true;
        }
        int NameWidth() { return Ui.Measure(DocName + "  •  " + DocState, Font).Width; }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x84) // WM_NCHITTEST: let the form treat empty title-bar space as caption (drag, double-click maximize, snap)
            {
                Point p = PointToClient(new Point((short)((long)m.LParam & 0xFFFF), (short)(((long)m.LParam >> 16) & 0xFFFF)));
                if (IsDragArea(p) && p.Y >= 4) { m.Result = (IntPtr)(-1); return; } // HTTRANSPARENT
            }
            base.WndProc(ref m);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Ui.Chrome);
            Icons.Draw(g, AppIcon, rIcon, true);
            // AutoSave
            if (hot == "auto") Ui.FillRound(g, Ui.Hover, rAuto, 4);
            Ui.Str(g, "Tự động lưu", Font, Ui.Text, new Rectangle(rAuto.X + 6, rAuto.Y, Ui.Measure("Tự động lưu", Font).Width + 4, rAuto.Height), TextFormatFlags.VerticalCenter);
            int tw0 = Ui.Measure("Tự động lưu", Font).Width; Rectangle sw = new Rectangle(rAuto.X + 6 + tw0 + 8, rAuto.Y + 5, 34, 14);
            Ui.FillRound(g, AutoSave ? Accent : Color.White, sw, 7); Ui.DrawRound(g, AutoSave ? Accent : Ui.Sub, sw, 7, 1f);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (SolidBrush b = new SolidBrush(AutoSave ? Color.White : Ui.Sub)) g.FillEllipse(b, AutoSave ? sw.Right - 12 : sw.X + 3, sw.Y + 3, 8, 8);
            Ui.Str(g, AutoSave ? "Bật" : "Tắt", Ui.F(8f), Ui.Sub, new Rectangle(sw.Right + 4, sw.Y - 1, 24, 16), TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            Hot(g, "save", rSave); Icons.Draw(g, "save", new Rectangle(rSave.X + 6, rSave.Y + 6, 16, 16), true);
            Hot(g, "undo", rUndo); Icons.Draw(g, "undo", new Rectangle(rUndo.X + 6, rUndo.Y + 6, 16, 16), UndoEnabled);
            Hot(g, "undoarrow", rUndoArrow); Ui.Arrow(g, Ui.Sub, rUndoArrow.X + 5, rUndoArrow.Y + 14, true);
            Hot(g, "redo", rRedo); Icons.Draw(g, "redo", new Rectangle(rRedo.X + 6, rRedo.Y + 6, 16, 16), RedoEnabled);
            Hot(g, "qat", rQat); using (Pen p = new Pen(Ui.Sub)) g.DrawLine(p, rQat.X + 5, rQat.Y + 10, rQat.X + 12, rQat.Y + 10); Ui.Arrow(g, Ui.Sub, rQat.X + 9, rQat.Y + 15, true);
            // document name
            int nw = Math.Min(rName.Width - 14, NameWidth());
            Rectangle nr = new Rectangle(rName.X, rName.Y, nw + 16, rName.Height); if (hot == "name") Ui.FillRound(g, Ui.Hover, nr, 4);
            Ui.Str(g, DocName + (DocState.Length > 0 ? "  •  " + DocState : ""), Font, Ui.Text, new Rectangle(nr.X + 4, nr.Y, nw, nr.Height), TextFormatFlags.VerticalCenter);
            if (nw > 40) Ui.Arrow(g, Ui.Sub, nr.Right - 6, nr.Y + 14, true);
            // search
            Ui.FillRound(g, Color.White, rSearch, 5); Ui.DrawRound(g, Search.Focused ? Accent : Color.FromArgb(0xC8, 0xC8, 0xC8), rSearch, 5, 1f);
            Icons.Draw(g, "find", new Rectangle(rSearch.X + 10, rSearch.Y + 6, 16, 16), true);
            // account
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (SolidBrush b = new SolidBrush(Ui.Light(Accent, 0.25f))) g.FillEllipse(b, rAccount);
            Ui.Str(g, Ui.Initials(Ui.UserName()), Ui.F(8f, FontStyle.Bold), Color.White, rAccount, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            // caption buttons
            if (hot == "min") g.FillRectangle(new SolidBrush(Ui.Hover), rMin);
            if (hot == "max") g.FillRectangle(new SolidBrush(Ui.Hover), rMax);
            if (hot == "close") g.FillRectangle(new SolidBrush(Color.FromArgb(0xC4, 0x2B, 0x1C)), rClose);
            g.SmoothingMode = SmoothingMode.None;
            using (Pen p = new Pen(Ui.Text, 1f))
            {
                g.DrawLine(p, rMin.X + 18, 20, rMin.X + 28, 20);
                Form f = Owner;
                if (f != null && f.WindowState == FormWindowState.Maximized) { g.DrawRectangle(p, rMax.X + 18, 17, 8, 8); g.DrawLine(p, rMax.X + 20, 15, rMax.X + 28, 15); g.DrawLine(p, rMax.X + 28, 15, rMax.X + 28, 23); }
                else g.DrawRectangle(p, rMax.X + 18, 15, 10, 10);
            }
            using (Pen p = new Pen(hot == "close" ? Color.White : Ui.Text, 1.1f)) { g.SmoothingMode = SmoothingMode.AntiAlias; g.DrawLine(p, rClose.X + 18, 15, rClose.X + 28, 25); g.DrawLine(p, rClose.X + 28, 15, rClose.X + 18, 25); }
        }
        void Hot(Graphics g, string k, Rectangle r) { if (hot == k) Ui.FillRound(g, Ui.Hover, r, 4); }

        string Hit(Point p)
        {
            if (rAuto.Contains(p)) return "auto"; if (rSave.Contains(p)) return "save"; if (rUndo.Contains(p)) return "undo"; if (rUndoArrow.Contains(p)) return "undoarrow";
            if (rRedo.Contains(p)) return "redo"; if (rQat.Contains(p)) return "qat"; if (rAccount.Contains(p)) return "account";
            if (rMin.Contains(p)) return "min"; if (rMax.Contains(p)) return "max"; if (rClose.Contains(p)) return "close";
            if (rName.Contains(p) && p.X < rName.X + NameWidth() + 20) return "name";
            return null;
        }
        readonly ToolTip tip = new ToolTip();
        protected override void OnMouseMove(MouseEventArgs e)
        {
            string h = Hit(e.Location); if (h == hot) return; hot = h; Invalidate();
            string t = h == "save" ? "Lưu (Ctrl+S)" : h == "undo" ? "Hoàn tác (Ctrl+Z)" : h == "redo" ? "Làm lại (Ctrl+Y)" : h == "auto" ? "Tự động lưu: lưu tệp ngay sau mỗi thay đổi" : h == "qat" ? "Tùy chỉnh Thanh công cụ Truy nhập Nhanh" : h == "min" ? "Thu nhỏ" : h == "max" ? "Phóng to / Khôi phục" : h == "close" ? "Đóng" : h == "account" ? Ui.UserName() : null;
            if (t != null) tip.Show(t, this, e.X, Height + 2, 3000); else tip.Hide(this);
        }
        protected override void OnMouseLeave(EventArgs e) { hot = null; tip.Hide(this); Invalidate(); }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return; string h = Hit(e.Location); Form f = Owner;
            switch (h)
            {
                case "auto": AutoSave = !AutoSave; Invalidate(); if (AutoSaveToggled != null) AutoSaveToggled(); break;
                case "save": if (SaveClicked != null) SaveClicked(); break;
                case "undo": case "undoarrow": if (UndoClicked != null) UndoClicked(); break;
                case "redo": if (RedoClicked != null) RedoClicked(); break;
                case "qat": if (QatClicked != null) QatClicked(); break;
                case "name": if (DocNameClicked != null) DocNameClicked(); break;
                case "account": if (AccountClicked != null) AccountClicked(); break;
                case "min": f.WindowState = FormWindowState.Minimized; break;
                case "max": f.WindowState = f.WindowState == FormWindowState.Maximized ? FormWindowState.Normal : FormWindowState.Maximized; break;
                case "close": f.Close(); break;
            }
        }
        public Point Below(string part) { Rectangle r = part == "name" ? rName : part == "qat" ? rQat : part == "account" ? rAccount : rSearch; return new Point(r.X, r.Bottom + 2); }
    }

    /// <summary>Status bar: info segments on the left, view buttons + zoom slider on the right.</summary>
    class StatusBar : Control
    {
        public class Seg { public string Text, Icon, Tip; public Action Click; public Rectangle R; }
        public readonly List<Seg> Segs = new List<Seg>(), Views = new List<Seg>();
        public int ActiveView; public float Zoom = 1f; public event Action<float> ZoomChanged; public event Action ZoomClicked;
        public Color Accent;
        Rectangle rMinus, rPlus, rTrack, rPct; object hot; bool drag;
        public StatusBar() { SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true); Dock = DockStyle.Bottom; Height = 26; BackColor = Ui.Chrome; Font = Ui.F(8.5f); }
        public Seg Add(string text, string icon, string tip, Action click) { Seg s = new Seg { Text = text, Icon = icon, Tip = tip, Click = click }; Segs.Add(s); return s; }
        public Seg View(string icon, string tip, Action click) { Seg s = new Seg { Icon = icon, Tip = tip, Click = click }; Views.Add(s); return s; }
        static float ZoomToPos(float z) { return z <= 1f ? (z - 0.1f) / 0.9f * 0.5f : 0.5f + (Math.Min(5f, z) - 1f) / 4f * 0.5f; }
        static float PosToZoom(float p) { p = Math.Max(0, Math.Min(1, p)); return p <= 0.5f ? 0.1f + p / 0.5f * 0.9f : 1f + (p - 0.5f) / 0.5f * 4f; }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Ui.Chrome);
            int x = 10;
            foreach (Seg s in Segs)
            {
                int w = (s.Icon != null ? 20 : 0) + Ui.Measure(s.Text, Font).Width + 12; s.R = new Rectangle(x, 2, w, Height - 4);
                if (hot == (object)s && s.Click != null) Ui.FillRound(g, Ui.Hover, s.R, 3);
                if (s.Icon != null) Icons.Draw(g, s.Icon, new Rectangle(x + 4, (Height - 14) / 2, 14, 14), true);
                Ui.Str(g, s.Text, Font, Ui.Sub, new Rectangle(x + (s.Icon != null ? 22 : 6), 0, w, Height), TextFormatFlags.VerticalCenter);
                x += w + 6;
            }
            int r = Width - 10;
            rPct = new Rectangle(r - 42, 3, 42, Height - 6); r = rPct.X - 4;
            rPlus = new Rectangle(r - 18, 4, 18, Height - 8); r = rPlus.X - 2;
            rTrack = new Rectangle(r - 104, Height / 2 - 1, 100, 2); r = rTrack.X - 4;
            rMinus = new Rectangle(r - 18, 4, 18, Height - 8); r = rMinus.X - 12;
            for (int i = Views.Count - 1; i >= 0; i--) { Views[i].R = new Rectangle(r - 26, 2, 26, Height - 4); r -= 28; }
            for (int i = 0; i < Views.Count; i++)
            {
                Seg v = Views[i]; if (i == ActiveView) Ui.FillRound(g, Ui.Checked, v.R, 3); else if (hot == (object)v) Ui.FillRound(g, Ui.Hover, v.R, 3);
                Icons.Draw(g, v.Icon, new Rectangle(v.R.X + 5, v.R.Y + (v.R.Height - 16) / 2, 16, 16), true);
            }
            using (Pen p = new Pen(Ui.Sub, 1.2f))
            {
                g.DrawLine(p, rMinus.X + 5, Height / 2, rMinus.X + 13, Height / 2);
                g.DrawLine(p, rPlus.X + 5, Height / 2, rPlus.X + 13, Height / 2); g.DrawLine(p, rPlus.X + 9, Height / 2 - 4, rPlus.X + 9, Height / 2 + 4);
            }
            using (SolidBrush b = new SolidBrush(Color.FromArgb(0xA0, 0xA0, 0xA0))) g.FillRectangle(b, rTrack);
            g.DrawLine(Pens.Gray, rTrack.X + rTrack.Width / 2, rTrack.Y - 3, rTrack.X + rTrack.Width / 2, rTrack.Y + 4);
            int tx = rTrack.X + (int)(ZoomToPos(Zoom) * rTrack.Width);
            using (SolidBrush b = new SolidBrush(Ui.Text)) g.FillRectangle(b, tx - 2, rTrack.Y - 6, 4, 14);
            if (hot == (object)"pct") Ui.FillRound(g, Ui.Hover, rPct, 3);
            Ui.Str(g, (int)Math.Round(Zoom * 100) + "%", Font, Ui.Sub, rPct, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
        object Hit(Point p)
        {
            foreach (Seg s in Segs) if (s.R.Contains(p)) return s; foreach (Seg s in Views) if (s.R.Contains(p)) return s;
            if (rMinus.Contains(p)) return "minus"; if (rPlus.Contains(p)) return "plus"; if (rPct.Contains(p)) return "pct";
            if (Rectangle.Inflate(rTrack, 4, 8).Contains(p)) return "track"; return null;
        }
        readonly ToolTip tip = new ToolTip();
        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (drag) { SetZoom(PosToZoom((e.X - rTrack.X) / (float)rTrack.Width)); return; }
            object h = Hit(e.Location); if (h == hot) return; hot = h; Invalidate();
            Seg s = h as Seg; string t = s != null ? s.Tip : (h as string) == "minus" ? "Thu nhỏ" : (h as string) == "plus" ? "Phóng to" : (h as string) == "pct" ? "Mức thu phóng" : null;
            if (t != null) tip.Show(t, this, e.X, -22, 2500); else tip.Hide(this);
        }
        protected override void OnMouseLeave(EventArgs e) { hot = null; tip.Hide(this); Invalidate(); }
        protected override void OnMouseDown(MouseEventArgs e) { if ((Hit(e.Location) as string) == "track") { drag = true; SetZoom(PosToZoom((e.X - rTrack.X) / (float)rTrack.Width)); } }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (drag) { drag = false; return; }
            object h = Hit(e.Location); Seg s = h as Seg;
            if (s != null) { if (s.Click != null) { int vi = Views.IndexOf(s); if (vi >= 0) ActiveView = vi; s.Click(); Invalidate(); } return; }
            if ((h as string) == "minus") SetZoom((float)Math.Max(0.1, Math.Round(Zoom * 10 - 1) / 10));
            if ((h as string) == "plus") SetZoom((float)Math.Min(5, Math.Round(Zoom * 10 + 1) / 10));
            if ((h as string) == "pct" && ZoomClicked != null) ZoomClicked();
        }
        void SetZoom(float z) { if (Math.Abs(z - 1f) < 0.04f) z = 1f; Zoom = z; Invalidate(); if (ZoomChanged != null) ZoomChanged(z); }
    }

    /// <summary>Excel formula bar: Name Box, cancel / enter / fx buttons, formula text, expand.</summary>
    class FormulaBar : Control
    {
        public readonly ComboBox NameBox = new ComboBox { FlatStyle = FlatStyle.Flat, Font = Ui.F(9f), DropDownStyle = ComboBoxStyle.DropDown };
        public readonly TextBox Formula = new TextBox { BorderStyle = BorderStyle.None, Font = new Font("Segoe UI", 9.5f), Multiline = true, WordWrap = true };
        public event Action CancelClicked, EnterClicked, FxClicked;
        Rectangle rCancel, rEnter, rFx, rExpand, rBox; string hot; bool expanded;
        public FormulaBar()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Dock = DockStyle.Top; Height = 32; BackColor = Ui.Chrome;
            Controls.Add(NameBox); Controls.Add(Formula);
            Formula.AcceptsReturn = false;
        }
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            NameBox.Bounds = new Rectangle(8, 5, 110, 22);
            rCancel = new Rectangle(140, 5, 22, 22); rEnter = new Rectangle(164, 5, 22, 22); rFx = new Rectangle(188, 5, 24, 22);
            rBox = new Rectangle(218, 4, Width - 218 - 8, Height - 8); rExpand = new Rectangle(rBox.Right - 20, rBox.Y + 2, 18, 18);
            Formula.Bounds = new Rectangle(rBox.X + 6, rBox.Y + 4, rBox.Width - 30, rBox.Height - 6);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Ui.Chrome);
            using (Pen p = new Pen(Ui.Sub)) { for (int k = 0; k < 3; k++) g.FillEllipse(Brushes.Gray, 127, 9 + k * 5, 2, 2); }
            if (hot == "cancel") Ui.FillRound(g, Ui.Hover, rCancel, 3); if (hot == "enter") Ui.FillRound(g, Ui.Hover, rEnter, 3); if (hot == "fx") Ui.FillRound(g, Ui.Hover, rFx, 3);
            Color c = Formula.Focused ? Ui.Text : Color.FromArgb(0xB0, 0xB0, 0xB0);
            using (Pen p = new Pen(c, 1.4f)) { g.SmoothingMode = SmoothingMode.AntiAlias; g.DrawLine(p, rCancel.X + 7, rCancel.Y + 7, rCancel.X + 15, rCancel.Y + 15); g.DrawLine(p, rCancel.X + 15, rCancel.Y + 7, rCancel.X + 7, rCancel.Y + 15); g.DrawLines(p, new[] { new Point(rEnter.X + 5, rEnter.Y + 11), new Point(rEnter.X + 9, rEnter.Y + 15), new Point(rEnter.X + 17, rEnter.Y + 7) }); }
            Ui.Str(g, "fx", new Font("Cambria", 10f, FontStyle.Italic), Ui.Text, rFx, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            Ui.FillRound(g, Color.White, rBox, 3); Ui.DrawRound(g, Color.FromArgb(0xD0, 0xD0, 0xD0), rBox, 3, 1f);
            Ui.Arrow(g, Ui.Sub, rExpand.X + 9, rExpand.Y + 9, !expanded);
        }
        string Hit(Point p) { return rCancel.Contains(p) ? "cancel" : rEnter.Contains(p) ? "enter" : rFx.Contains(p) ? "fx" : rExpand.Contains(p) ? "expand" : null; }
        protected override void OnMouseMove(MouseEventArgs e) { string h = Hit(e.Location); if (h != hot) { hot = h; Invalidate(); } }
        protected override void OnMouseLeave(EventArgs e) { hot = null; Invalidate(); }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            switch (Hit(e.Location))
            {
                case "cancel": if (CancelClicked != null) CancelClicked(); break;
                case "enter": if (EnterClicked != null) EnterClicked(); break;
                case "fx": if (FxClicked != null) FxClicked(); break;
                case "expand": expanded = !expanded; Height = expanded ? 76 : 32; break;
            }
        }
        public Point FxPoint { get { return new Point(rFx.X, rFx.Bottom + 2); } }
    }

    /// <summary>Excel sheet tabs + add button (left of the horizontal scroll bar).</summary>
    class SheetTabs : Control
    {
        public readonly List<string> Names = new List<string>(); public int Active; public Color Accent;
        public event Action<int> TabClicked; public event Action AddClicked; public event Action<int, Point> TabMenu;
        readonly List<Rectangle> rects = new List<Rectangle>(); Rectangle rPrev, rNext, rAdd; object hot; int first;
        public void EnsureVisible(int i) { if (i < first || i > first + 5) first = Math.Max(0, i - 2); Invalidate(); }
        public SheetTabs() { SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true); Height = 28; BackColor = Ui.Chrome; Font = Ui.F(9f); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Ui.Chrome); rects.Clear();
            rPrev = new Rectangle(6, 4, 20, 20); rNext = new Rectangle(28, 4, 20, 20);
            using (SolidBrush b = new SolidBrush(first > 0 ? Ui.Text : Color.Silver)) g.FillPolygon(b, new[] { new Point(rPrev.X + 12, rPrev.Y + 5), new Point(rPrev.X + 12, rPrev.Y + 15), new Point(rPrev.X + 7, rPrev.Y + 10) });
            using (SolidBrush b = new SolidBrush(Ui.Text)) g.FillPolygon(b, new[] { new Point(rNext.X + 8, rNext.Y + 5), new Point(rNext.X + 8, rNext.Y + 15), new Point(rNext.X + 13, rNext.Y + 10) });
            int x = 56;
            for (int i = 0; i < Names.Count; i++)
            {
                if (i < first) { rects.Add(Rectangle.Empty); continue; }
                int w = Ui.Measure(Names[i], i == Active ? Ui.F(9f, FontStyle.Bold) : Font).Width + 24; Rectangle r = new Rectangle(x, 2, w, Height - 4);
                if (r.Right > Width - 40) { rects.Add(Rectangle.Empty); continue; }
                rects.Add(r);
                if (i == Active) { Ui.FillRound(g, Color.White, r, 3); Ui.FillRound(g, Accent, new Rectangle(r.X + 8, r.Bottom - 3, r.Width - 16, 3), 1); }
                else if (hot is int && (int)hot == i) Ui.FillRound(g, Ui.Hover, r, 3);
                Ui.Str(g, Names[i], i == Active ? Ui.F(9f, FontStyle.Bold) : Font, i == Active ? Accent : Ui.Text, r, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                x = r.Right + 2;
            }
            rAdd = new Rectangle(x + 4, 4, 20, 20); if (hot is string) Ui.FillRound(g, Ui.Hover, rAdd, 10);
            using (Pen p = new Pen(Ui.Sub, 1.3f)) { g.DrawLine(p, rAdd.X + 5, rAdd.Y + 10, rAdd.X + 15, rAdd.Y + 10); g.DrawLine(p, rAdd.X + 10, rAdd.Y + 5, rAdd.X + 10, rAdd.Y + 15); }
        }
        object Hit(Point p) { for (int i = 0; i < rects.Count; i++) if (rects[i].Contains(p)) return i; if (rAdd.Contains(p)) return "add"; if (rPrev.Contains(p)) return "prev"; if (rNext.Contains(p)) return "next"; return null; }
        protected override void OnMouseMove(MouseEventArgs e) { object h = Hit(e.Location); if (!Equals(h, hot)) { hot = h; Invalidate(); } }
        protected override void OnMouseLeave(EventArgs e) { hot = null; Invalidate(); }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            object h = Hit(e.Location);
            if (h is int) { if (e.Button == MouseButtons.Right) { if (TabMenu != null) TabMenu((int)h, e.Location); } else if (TabClicked != null) TabClicked((int)h); }
            else if ((h as string) == "add") { if (AddClicked != null) AddClicked(); }
            else if ((h as string) == "prev") { first = Math.Max(0, first - 1); Invalidate(); }
            else if ((h as string) == "next") { first = Math.Min(Math.Max(0, Names.Count - 1), first + 1); Invalidate(); }
        }
    }

    /// <summary>PowerPoint slide thumbnails pane.</summary>
    class SlidePane : Control
    {
        public int Count, Active; public Color Accent; public Func<int, int, Bitmap> Render; public event Action<int> SlideClicked; public event Action<int, Point> SlideMenu;
        readonly Dictionary<int, Bitmap> thumbs = new Dictionary<int, Bitmap>(); readonly HashSet<int> dirty = new HashSet<int>();
        int scroll; float aspect = 9f / 16f; int hot = -1;
        readonly Timer idle = new Timer { Interval = 120 };
        public SlidePane()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            Dock = DockStyle.Left; Width = 200; BackColor = Ui.Chrome; Font = Ui.F(9f);
            idle.Tick += (s, e) => RenderNext(); idle.Start();
        }
        public void SetAspect(Size tw) { if (tw.Width > 0) aspect = tw.Height / (float)tw.Width; }
        public void MarkDirty(int i) { dirty.Add(i); }
        public void Reset() { foreach (Bitmap b in thumbs.Values) b.Dispose(); thumbs.Clear(); dirty.Clear(); Invalidate(); }
        int ThumbW { get { return Width - 52; } }
        int ItemH { get { return (int)(ThumbW * aspect) + 16; } }
        Rectangle ThumbRect(int i) { return new Rectangle(34, 10 + i * ItemH - scroll, ThumbW, (int)(ThumbW * aspect)); }
        void RenderNext()
        {
            if (Render == null || !Visible) return;
            for (int i = 0; i < Count; i++)
            {
                Rectangle r = ThumbRect(i); if (r.Bottom < -200 || r.Y > Height + 200) continue;
                Bitmap b; if (thumbs.TryGetValue(i, out b) && !dirty.Contains(i) && b.Width == r.Width) continue;
                Bitmap nb = null; try { nb = Render(i, r.Width); } catch { }
                if (nb != null) { if (b != null) b.Dispose(); thumbs[i] = nb; } dirty.Remove(i); Invalidate(); return;
            }
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Ui.Chrome);
            for (int i = 0; i < Count; i++)
            {
                Rectangle r = ThumbRect(i); if (r.Bottom < 0 || r.Y > Height) continue;
                Ui.Str(g, (i + 1).ToString(), Font, i == Active ? Accent : Ui.Sub, new Rectangle(4, r.Y, 24, 18), TextFormatFlags.Right);
                Bitmap b; if (thumbs.TryGetValue(i, out b)) g.DrawImage(b, r); else g.FillRectangle(Brushes.White, r);
                if (i == Active) Ui.DrawRound(g, Accent, Rectangle.Inflate(r, 3, 3), 4, 2.5f); else if (i == hot) Ui.DrawRound(g, Color.FromArgb(0xB0, 0xB0, 0xB0), Rectangle.Inflate(r, 2, 2), 3, 1.5f);
                else using (Pen p = new Pen(Color.FromArgb(0xD0, 0xD0, 0xD0))) g.DrawRectangle(p, r);
            }
        }
        int Hit(Point p) { for (int i = 0; i < Count; i++) if (Rectangle.Inflate(ThumbRect(i), 4, 4).Contains(p)) return i; return -1; }
        protected override void OnMouseMove(MouseEventArgs e) { int h = Hit(e.Location); if (h != hot) { hot = h; Invalidate(); } }
        protected override void OnMouseLeave(EventArgs e) { hot = -1; Invalidate(); }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            Focus(); int h = Hit(e.Location); if (h < 0) return;
            if (e.Button == MouseButtons.Right) { if (SlideMenu != null) SlideMenu(h, e.Location); }
            else if (SlideClicked != null) SlideClicked(h);
        }
        protected override void OnMouseWheel(MouseEventArgs e) { ScrollBy(-e.Delta); }
        void ScrollBy(int d) { int max = Math.Max(0, Count * ItemH + 20 - Height); scroll = Math.Max(0, Math.Min(max, scroll + d)); Invalidate(); }
        public void EnsureVisible(int i) { Rectangle r = ThumbRect(i); if (r.Y < 0) ScrollBy(r.Y - 10); else if (r.Bottom > Height) ScrollBy(r.Bottom - Height + 10); }
        protected override bool IsInputKey(Keys k) { return k == Keys.Up || k == Keys.Down || base.IsInputKey(k); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Down || e.KeyCode == Keys.PageDown) && Active < Count - 1 && SlideClicked != null) SlideClicked(Active + 1);
            if ((e.KeyCode == Keys.Up || e.KeyCode == Keys.PageUp) && Active > 0 && SlideClicked != null) SlideClicked(Active - 1);
        }
        protected override void Dispose(bool d) { idle.Dispose(); foreach (Bitmap b in thumbs.Values) b.Dispose(); base.Dispose(d); }
    }

    /// <summary>Excel column letters / row numbers around the sheet canvas (data from .uno:ViewRowColumnHeaders).</summary>
    class SheetHeaders : Control
    {
        public const int RowW = 44, ColH = 22;
        public List<KeyValuePair<string, int>> Cols = new List<KeyValuePair<string, int>>(), Rows = new List<KeyValuePair<string, int>>();
        public int ColStart, RowStart; // pixel position (canvas) of the first entry's start
        public Rectangle Current = Rectangle.Empty; // in canvas pixels
        public Color Accent; public event Action<string> Select; public event Action SelectAll;
        public SheetHeaders() { SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true); Dock = DockStyle.Fill; BackColor = Color.White; Font = Ui.F(8.5f); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            using (SolidBrush hb = new SolidBrush(Color.FromArgb(0xF5, 0xF5, 0xF5))) { g.FillRectangle(hb, 0, 0, Width, ColH); g.FillRectangle(hb, 0, 0, RowW, Height); }
            using (Pen ln = new Pen(Color.FromArgb(0xD8, 0xD8, 0xD8)))
            using (SolidBrush curB = new SolidBrush(Color.FromArgb(0xD2, 0xD2, 0xD2)))
            using (Pen curP = new Pen(Accent, 2f))
            {
                int prev = ColStart;
                foreach (KeyValuePair<string, int> c in Cols)
                {
                    int x0 = RowW + prev, x1 = RowW + c.Value; prev = c.Value; if (x1 < RowW) continue; if (x0 > Width) break;
                    bool cur = !Current.IsEmpty && Current.Right > x0 - RowW && Current.X < x1 - RowW;
                    if (cur) { g.FillRectangle(curB, x0, 0, x1 - x0, ColH); g.DrawLine(curP, x0, ColH - 1, x1, ColH - 1); }
                    g.DrawLine(ln, x1, 0, x1, ColH);
                    if (x1 - x0 > 8) Ui.Str(g, c.Key, Font, cur ? Ui.Text : Ui.Sub, new Rectangle(Math.Max(RowW, x0), 0, x1 - Math.Max(RowW, x0), ColH), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }
                prev = RowStart;
                foreach (KeyValuePair<string, int> r in Rows)
                {
                    int y0 = ColH + prev, y1 = ColH + r.Value; prev = r.Value; if (y1 < ColH) continue; if (y0 > Height) break;
                    bool cur = !Current.IsEmpty && Current.Bottom > y0 - ColH && Current.Y < y1 - ColH;
                    if (cur) { g.FillRectangle(curB, 0, y0, RowW, y1 - y0); g.DrawLine(curP, RowW - 1, y0, RowW - 1, y1); }
                    g.DrawLine(ln, 0, y1, RowW, y1);
                    if (y1 - y0 > 6) Ui.Str(g, r.Key, Font, cur ? Ui.Text : Ui.Sub, new Rectangle(0, Math.Max(ColH, y0), RowW - 4, y1 - Math.Max(ColH, y0)), TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
                }
                g.DrawLine(ln, 0, ColH - 1, Width, ColH - 1); g.DrawLine(ln, RowW - 1, 0, RowW - 1, Height);
            }
            using (SolidBrush tri = new SolidBrush(Color.FromArgb(0xB8, 0xB8, 0xB8))) g.FillPolygon(tri, new[] { new Point(RowW - 4, 6), new Point(RowW - 4, ColH - 4), new Point(RowW - 4 - (ColH - 10), ColH - 4) });
        }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (e.X < RowW && e.Y < ColH) { if (SelectAll != null) SelectAll(); return; }
            if (e.Y < ColH) { int prev = ColStart; foreach (KeyValuePair<string, int> c in Cols) { if (e.X - RowW >= prev && e.X - RowW < c.Value) { if (Select != null) Select(c.Key + ":" + c.Key); return; } prev = c.Value; } }
            if (e.X < RowW) { int prev = RowStart; foreach (KeyValuePair<string, int> r in Rows) { if (e.Y - ColH >= prev && e.Y - ColH < r.Value) { if (Select != null) Select(r.Key + ":" + r.Key); return; } prev = r.Value; } }
        }
    }
}
