// Office Open Source v0.1.6 - owner-drawn Ribbon (tabs, groups, large / small / icon commands, combos, galleries)
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace OfficeOpenSource
{
    enum RK { Large, Small, Mini, Combo, Gallery }

    class RMenu
    {
        public string Text, Icon; public Action Click; public bool Separator; public Func<bool> Checked;
        public RMenu(string text, string icon, Action click) { Text = text; Icon = icon; Click = click; }
        public static RMenu Sep() { return new RMenu(null, null, null) { Separator = true }; }
    }

    class RGalleryItem { public string Name; public Font Font; public Color Color; public Action Click; public string Preview; }

    class RItem
    {
        public RK Kind; public string Text, Icon, Tip, State; public Action Click; public List<RMenu> Menu; public bool Split;
        public bool Checked, Disabled; public Color Swatch = Color.Empty; public Action<Color> PickColor; public Color[] Palette;
        public int Width = 120; public ComboBox Combo; public string[] ComboItems; public Action<string> ComboApply;
        public List<RGalleryItem> Gallery; public int GalleryFirst;
        public Rectangle Bounds, Arrow; public bool HasArrow { get { return Menu != null || PickColor != null; } }
    }

    class RColumn { public RItem Large; public List<RItem> Stack; public List<List<RItem>> Rows; public RItem GalleryItem; public int W; }

    class RGroup
    {
        public string Name; public List<RColumn> Cols = new List<RColumn>(); public Action Launcher; public string LauncherTip;
        public Rectangle Bounds, LauncherBounds; public bool Collapsed; public string Icon;
        public RColumn Large(RItem i) { RColumn c = new RColumn { Large = i }; i.Kind = RK.Large; Cols.Add(c); return c; }
        public RColumn Stack(params RItem[] items) { RColumn c = new RColumn { Stack = new List<RItem>(items) }; foreach (RItem i in items) i.Kind = RK.Small; Cols.Add(c); return c; }
        public RColumn Rows(params RItem[][] rows) { RColumn c = new RColumn { Rows = new List<List<RItem>>() }; foreach (RItem[] r in rows) c.Rows.Add(new List<RItem>(r)); Cols.Add(c); return c; }
        public RColumn Gallery(RItem g) { g.Kind = RK.Gallery; RColumn c = new RColumn { GalleryItem = g }; Cols.Add(c); return c; }
        public IEnumerable<RItem> All()
        {
            foreach (RColumn c in Cols)
            {
                if (c.Large != null) yield return c.Large;
                if (c.GalleryItem != null) yield return c.GalleryItem;
                if (c.Stack != null) foreach (RItem i in c.Stack) yield return i;
                if (c.Rows != null) foreach (List<RItem> r in c.Rows) foreach (RItem i in r) yield return i;
            }
        }
    }

    class RTab { public string Name; public List<RGroup> Groups = new List<RGroup>(); public Rectangle Bounds; public RGroup Add(string name) { RGroup g = new RGroup { Name = name }; Groups.Add(g); return g; } }

    class Ribbon : Control
    {
        public readonly List<RTab> Tabs = new List<RTab>();
        public int Selected = 1;
        public Color Accent = Color.FromArgb(0x18, 0x5A, 0xBD);
        public event Action FileClicked;
        public event Action CommentsClicked, ShareClicked;
        public bool Minimized;
        const int TabH = 30, BodyH = 98, RowH = 24;
        object hot, down; Rectangle hotArrow;
        readonly ToolTip tip = new ToolTip { InitialDelay = 600, ReshowDelay = 200, AutoPopDelay = 6000 };
        Rectangle commentsBtn, shareBtn, collapseBtn, editModeBtn;
        public string EditMode = null; public List<RMenu> EditModeMenu;
        bool expandedTemp;

        public Ribbon()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Dock = DockStyle.Top; Height = TabH + BodyH + 6; Font = Ui.F(9f); BackColor = Ui.Chrome;
        }

        public RTab Tab(string name) { RTab t = new RTab { Name = name }; Tabs.Add(t); return t; }

        public static RItem I(string text, string icon, string tip, Action click) { return new RItem { Text = text, Icon = icon, Tip = tip, Click = click }; }

        public void SetMinimized(bool m) { Minimized = m; expandedTemp = false; Height = m ? TabH + 4 : TabH + BodyH + 6; LayoutAll(); Invalidate(); }

        public void SetState(string key, string value)
        {
            bool changed = false;
            foreach (RTab t in Tabs) foreach (RGroup g in t.Groups) foreach (RItem i in g.All())
                    {
                        if (i.State != key) continue;
                        if (key == ".uno:Undo" || key == ".uno:Redo") { bool d = value == "disabled"; if (i.Disabled != d) { i.Disabled = d; changed = true; } }
                        else { bool c = value == "true"; if (i.Checked != c) { i.Checked = c; changed = true; } }
                        if (i.Combo != null && !i.Combo.Focused && i.Combo.Text != value && value != "disabled" && value.Length > 0 && value != "0") i.Combo.Text = value;
                    }
            if (changed) Invalidate();
        }

        // ---------------- layout ----------------
        Size Text(string s) { return Ui.Measure(s, Font); }
        static string[] SplitLabel(string s)
        {
            if (string.IsNullOrEmpty(s)) return new[] { "" };
            if (s.Length <= 9 || !s.Contains(" ")) return new[] { s };
            int mid = s.Length / 2, best = -1;
            for (int k = 0; k < s.Length; k++) if (s[k] == ' ' && (best < 0 || Math.Abs(k - mid) < Math.Abs(best - mid))) best = k;
            return best < 0 ? new[] { s } : new[] { s.Substring(0, best), s.Substring(best + 1) };
        }

        public void LayoutAll()
        {
            using (Graphics g = CreateGraphics()) { }
            int x = 8;
            for (int i = 0; i < Tabs.Count; i++) { int w = Text(Tabs[i].Name).Width + 20; Tabs[i].Bounds = new Rectangle(x, 2, w, TabH - 4); x += w + 2; }
            int rx = Width - 8;
            shareBtn = new Rectangle(rx - 88, 4, 88, 24); rx = shareBtn.X - 8;
            commentsBtn = new Rectangle(rx - 104, 4, 104, 24); rx = commentsBtn.X - 8;
            editModeBtn = EditMode != null ? new Rectangle(rx - 136, 4, 136, 24) : Rectangle.Empty;
            foreach (RTab t in Tabs) foreach (RGroup gg in t.Groups) foreach (RItem it in gg.All()) if (it.Combo != null) it.Combo.Visible = false;
            if (Selected < 0 || Selected >= Tabs.Count) return;
            if (Minimized && !expandedTemp) return;
            RTab tab = Tabs[Selected];
            int top = TabH + 6, bodyBottom = TabH + BodyH - 2, labelH = 18;
            int avail = Width - 50; int gx = 14;
            // measure natural widths; collapse groups from the right while they do not fit
            List<int> widths = new List<int>();
            foreach (RGroup g in tab.Groups) { g.Collapsed = false; widths.Add(MeasureGroup(g)); }
            int total = 0; foreach (int w in widths) total += w + 1;
            if (total > avail)   // first make galleries narrower (Office shrinks galleries before collapsing groups)
                for (int k = 0; k < tab.Groups.Count; k++) foreach (RColumn c in tab.Groups[k].Cols) if (c.GalleryItem != null && c.GalleryItem.Width > 2 * 78 + 22) { c.GalleryItem.Width = 2 * 78 + 22; total -= widths[k]; widths[k] = MeasureGroup(tab.Groups[k]); total += widths[k]; }
            for (int k = tab.Groups.Count - 1; k >= 0 && total > avail; k--)
            {
                int cw = Math.Max(56, Text(tab.Groups[k].Name).Width + 16); if (cw >= widths[k]) continue;
                total -= widths[k]; tab.Groups[k].Collapsed = true; widths[k] = cw; total += widths[k];
            }
            for (int gi = 0; gi < tab.Groups.Count; gi++)
            {
                RGroup g = tab.Groups[gi]; int gw = widths[gi];
                g.Bounds = new Rectangle(gx, top, gw, bodyBottom - top);
                g.LauncherBounds = g.Launcher != null && !g.Collapsed ? new Rectangle(g.Bounds.Right - 16, g.Bounds.Bottom - 15, 12, 12) : Rectangle.Empty;
                if (!g.Collapsed)
                {
                    int cx = gx + 4;
                    foreach (RColumn c in g.Cols)
                    {
                        int cy = top + 2;
                        if (c.Large != null) c.Large.Bounds = new Rectangle(cx, cy, c.W, bodyBottom - labelH - cy - 2);
                        else if (c.GalleryItem != null) c.GalleryItem.Bounds = new Rectangle(cx, cy + 2, c.W, bodyBottom - labelH - cy - 8);
                        else if (c.Stack != null) { int yy = cy; foreach (RItem i in c.Stack) { i.Bounds = new Rectangle(cx, yy, c.W, RowH - 2); yy += RowH; } }
                        else if (c.Rows != null)
                        {
                            int yy = cy + (c.Rows.Count < 3 ? (3 - c.Rows.Count) * RowH / 2 : 0);
                            foreach (List<RItem> row in c.Rows)
                            {
                                int xx = cx;
                                foreach (RItem i in row) { int w = ItemWidth(i); i.Bounds = new Rectangle(xx, yy, w, RowH - 2); xx += w + (i.Kind == RK.Combo ? 4 : 0); }
                                yy += RowH;
                            }
                        }
                        cx += c.W + 4;
                    }
                    foreach (RItem i in g.All())
                    {
                        i.Arrow = i.HasArrow ? (i.Kind == RK.Large ? new Rectangle(i.Bounds.X, i.Bounds.Bottom - 14, i.Bounds.Width, 14) : new Rectangle(i.Bounds.Right - 12, i.Bounds.Y, 12, i.Bounds.Height)) : Rectangle.Empty;
                        if (i.Combo != null) { i.Combo.Bounds = new Rectangle(i.Bounds.X, i.Bounds.Y, i.Bounds.Width, i.Bounds.Height); i.Combo.Visible = true; }
                    }
                }
                gx += gw + 1;
            }
            collapseBtn = new Rectangle(Width - 30, bodyBottom - 22, 20, 18);
        }

        int ItemWidth(RItem i)
        {
            if (i.Kind == RK.Combo) return i.Width;
            int w = 24; if (!string.IsNullOrEmpty(i.Text) && i.Kind == RK.Mini && i.Icon == null) w = Text(i.Text).Width + 12;
            return w + (i.HasArrow ? 10 : 0);
        }

        int MeasureGroup(RGroup g)
        {
            int w = 8;
            foreach (RColumn c in g.Cols)
            {
                int cw = 0;
                if (c.Large != null) { foreach (string l in SplitLabel(c.Large.Text)) cw = Math.Max(cw, Text(l).Width + 10); cw = Math.Max(cw, 44); if (c.Large.HasArrow && SplitLabel(c.Large.Text).Length == 1) cw = Math.Max(cw, Text(c.Large.Text).Width + 12); }
                else if (c.GalleryItem != null) cw = c.GalleryItem.Width;
                else if (c.Stack != null) foreach (RItem i in c.Stack) cw = Math.Max(cw, 22 + Text(i.Text).Width + (i.HasArrow ? 14 : 6));
                else if (c.Rows != null) foreach (List<RItem> row in c.Rows) { int rw = 0; foreach (RItem i in row) rw += ItemWidth(i) + (i.Kind == RK.Combo ? 4 : 0); cw = Math.Max(cw, rw); }
                c.W = cw; w += cw + 4;
            }
            return Math.Max(w + 4, Text(g.Name).Width + 20);
        }

        protected override void OnResize(EventArgs e) { base.OnResize(e); LayoutAll(); }
        protected override void OnCreateControl() { base.OnCreateControl(); LayoutAll(); }

        // ---------------- painting ----------------
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Ui.Chrome);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            for (int i = 0; i < Tabs.Count; i++)
            {
                RTab t = Tabs[i]; bool sel = i == Selected && !(Minimized && !expandedTemp), h = hot == (object)t;
                if (h && !sel) Ui.FillRound(g, Ui.Hover, t.Bounds, 4);
                Ui.Str(g, t.Name, sel ? Ui.F(9f, FontStyle.Bold) : Font, sel ? Accent : Ui.Text, t.Bounds, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                if (sel) { int w = Math.Min(t.Bounds.Width - 12, Text(t.Name).Width); Ui.FillRound(g, Accent, new Rectangle(t.Bounds.X + (t.Bounds.Width - w) / 2, t.Bounds.Bottom - 2, w, 3), 1); }
            }
            // Editing mode / Comments / Share
            if (!editModeBtn.IsEmpty) { Btn(g, editModeBtn, "editor", EditMode, false, hot == (object)"edit"); }
            Btn(g, commentsBtn, "comment", "Bình luận", false, hot == (object)"comments");
            Ui.FillRound(g, hot == (object)"share" ? Ui.Light(Accent, -0.1f) : Accent, shareBtn, 4);
            Icons.Draw(g, "share", new Rectangle(shareBtn.X + 8, shareBtn.Y + 4, 16, 16), true);
            using (Bitmap b = new Bitmap(16, 16)) { }
            Ui.Str(g, "Chia sẻ", Ui.F(9f, FontStyle.Bold), Color.White, new Rectangle(shareBtn.X + 28, shareBtn.Y, shareBtn.Width - 34, shareBtn.Height), TextFormatFlags.VerticalCenter);
            if (Minimized && !expandedTemp) return;
            Rectangle body = new Rectangle(4, TabH + 2, Width - 8, BodyH);
            using (GraphicsPath p = Ui.Round(new Rectangle(body.X + 1, body.Y + 2, body.Width - 2, body.Height - 2), 7)) using (SolidBrush sh = new SolidBrush(Color.FromArgb(18, 0, 0, 0))) { g.SmoothingMode = SmoothingMode.AntiAlias; g.FillPath(sh, p); }
            Ui.FillRound(g, Ui.Card, new Rectangle(body.X, body.Y, body.Width, body.Height - 2), 7);
            Ui.DrawRound(g, Ui.Line, new Rectangle(body.X, body.Y, body.Width, body.Height - 2), 7, 1f);
            if (Selected < 0 || Selected >= Tabs.Count) return;
            foreach (RGroup gr in Tabs[Selected].Groups) PaintGroup(g, gr);
            // collapse ribbon chevron
            if (hot == (object)"collapse") Ui.FillRound(g, Ui.Hover, collapseBtn, 3);
            using (Pen pn = new Pen(Ui.Sub, 1.4f)) { int cx = collapseBtn.X + 10, cy = collapseBtn.Y + 9; g.DrawLines(pn, new[] { new Point(cx - 4, cy + 2), new Point(cx, cy - 2), new Point(cx + 4, cy + 2) }); }
        }

        void Btn(Graphics g, Rectangle r, string icon, string text, bool filled, bool h)
        {
            Ui.FillRound(g, h ? Ui.Hover : Color.White, r, 4); Ui.DrawRound(g, Color.FromArgb(0xC8, 0xC8, 0xC8), r, 4, 1f);
            Icons.Draw(g, icon, new Rectangle(r.X + 6, r.Y + 4, 16, 16), true);
            Ui.Str(g, text, Font, Ui.Text, new Rectangle(r.X + 26, r.Y, r.Width - 30, r.Height), TextFormatFlags.VerticalCenter);
        }

        void PaintGroup(Graphics g, RGroup gr)
        {
            Rectangle b = gr.Bounds;
            using (Pen pn = new Pen(Ui.Line)) g.DrawLine(pn, b.Right, b.Y + 6, b.Right, b.Bottom - 6);
            Ui.Str(g, gr.Name, Ui.F(8.25f), Ui.Sub, new Rectangle(b.X, b.Bottom - 18, b.Width - (gr.LauncherBounds.IsEmpty ? 0 : 14), 16), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            if (!gr.LauncherBounds.IsEmpty)
            {
                Rectangle l = gr.LauncherBounds; if (hot == (object)gr) Ui.FillRound(g, Ui.Hover, Rectangle.Inflate(l, 2, 2), 2);
                using (Pen pn = new Pen(Ui.Sub, 1f)) { g.DrawLine(pn, l.X + 2, l.Y + 2, l.X + 9, l.Y + 2); g.DrawLine(pn, l.X + 2, l.Y + 2, l.X + 2, l.Y + 9); g.DrawLine(pn, l.X + 4, l.Y + 4, l.X + 10, l.Y + 10); g.DrawLine(pn, l.X + 10, l.Y + 6, l.X + 10, l.Y + 10); g.DrawLine(pn, l.X + 6, l.Y + 10, l.X + 10, l.Y + 10); }
            }
            if (gr.Collapsed)
            {
                Rectangle r = new Rectangle(b.X + 4, b.Y + 2, b.Width - 8, b.Height - 22); if (hot == (object)gr) Ui.FillRound(g, Ui.Hover, r, 4);
                string ic = gr.Icon; if (ic == null) foreach (RItem i in gr.All()) if (i.Icon != null) { ic = i.Icon; break; }
                Icons.Draw(g, ic, new Rectangle(r.X + (r.Width - 32) / 2, r.Y + 4, 32, 32), true);
                Ui.Arrow(g, Ui.Sub, r.X + r.Width / 2, r.Bottom - 10, true);
                return;
            }
            foreach (RItem i in gr.All()) PaintItem(g, i);
        }

        void PaintItem(Graphics g, RItem i)
        {
            if (i.Kind == RK.Combo) return;
            Rectangle r = i.Bounds; bool h = hot == (object)i && !i.Disabled, d = down == (object)i;
            if (i.Kind == RK.Gallery) { PaintGallery(g, i); return; }
            if (i.Checked) { Ui.FillRound(g, Ui.Checked, r, 4); Ui.DrawRound(g, Color.FromArgb(0xB8, 0xB8, 0xB8), r, 4, 1f); }
            if (h)
            {
                Ui.FillRound(g, d ? Ui.Press : Ui.Hover, r, 4);
                if (i.Split && !i.Arrow.IsEmpty) using (Pen pn = new Pen(Color.FromArgb(0xC8, 0xC8, 0xC8))) { if (i.Kind == RK.Large) g.DrawLine(pn, r.X + 3, i.Arrow.Y, r.Right - 3, i.Arrow.Y); else g.DrawLine(pn, i.Arrow.X, r.Y + 3, i.Arrow.X, r.Bottom - 3); }
            }
            Color tc = i.Disabled ? Color.FromArgb(0xA6, 0xA6, 0xA6) : Ui.Text;
            if (i.Kind == RK.Large)
            {
                Icons.Draw(g, i.Icon, new Rectangle(r.X + (r.Width - 32) / 2, r.Y + 3, 32, 32), !i.Disabled);
                if (!i.Swatch.IsEmpty) using (SolidBrush sb = new SolidBrush(i.Swatch)) g.FillRectangle(sb, r.X + (r.Width - 24) / 2, r.Y + 31, 24, 4);
                string[] lines = SplitLabel(i.Text); int y = r.Y + 38;
                for (int k = 0; k < lines.Length; k++)
                {
                    string s = lines[k]; bool last = k == lines.Length - 1;
                    Rectangle lr = new Rectangle(r.X, y, r.Width, 15);
                    if (last && i.HasArrow) { Size sz = Text(s); int tx = r.X + (r.Width - sz.Width - (lines.Length == 1 ? 0 : 10)) / 2; if (lines.Length == 1) { Ui.Str(g, s, Font, tc, lr, TextFormatFlags.HorizontalCenter); Ui.Arrow(g, Ui.Sub, r.X + r.Width / 2, y + 20, true); } else { Ui.Str(g, s, Font, tc, new Rectangle(tx, y, sz.Width + 2, 15), TextFormatFlags.Left); Ui.Arrow(g, Ui.Sub, tx + sz.Width + 6, y + 7, true); } }
                    else Ui.Str(g, s, Font, tc, lr, TextFormatFlags.HorizontalCenter);
                    y += 14;
                }
                return;
            }
            if (i.Kind == RK.Small)
            {
                Icons.Draw(g, i.Icon, new Rectangle(r.X + 3, r.Y + (r.Height - 16) / 2, 16, 16), !i.Disabled);
                Ui.Str(g, i.Text, Font, tc, new Rectangle(r.X + 22, r.Y, r.Width - 22 - (i.HasArrow ? 12 : 0), r.Height), TextFormatFlags.VerticalCenter);
                if (i.HasArrow) Ui.Arrow(g, Ui.Sub, i.Arrow.X + 6, r.Y + r.Height / 2, true);
                return;
            }
            // Mini
            if (i.Icon != null) Icons.Draw(g, i.Icon, new Rectangle(r.X + 4, r.Y + (r.Height - 16) / 2, 16, 16), !i.Disabled);
            else Ui.Str(g, i.Text, Font, tc, new Rectangle(r.X, r.Y, r.Width - (i.HasArrow ? 10 : 0), r.Height), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            if (!i.Swatch.IsEmpty) using (SolidBrush sb = new SolidBrush(i.Swatch)) g.FillRectangle(sb, r.X + 4, r.Bottom - 6, 16, 4);
            if (i.HasArrow) Ui.Arrow(g, Ui.Sub, i.Arrow.X + 5, r.Y + r.Height / 2, true);
        }

        void PaintGallery(Graphics g, RItem i)
        {
            Rectangle r = i.Bounds;
            using (SolidBrush wb = new SolidBrush(Color.White)) g.FillRectangle(wb, r);
            using (Pen pn = new Pen(Color.FromArgb(0xD0, 0xD0, 0xD0))) g.DrawRectangle(pn, r);
            int cell = 78, n = Math.Max(1, (r.Width - 18) / cell);
            for (int k = 0; k < n && i.GalleryFirst + k < i.Gallery.Count; k++)
            {
                RGalleryItem gi = i.Gallery[i.GalleryFirst + k]; Rectangle c = new Rectangle(r.X + 2 + k * cell, r.Y + 2, cell - 2, r.Height - 4);
                if (hot == (object)gi) { using (SolidBrush hb = new SolidBrush(Ui.Hover)) g.FillRectangle(hb, c); using (Pen pn = new Pen(Color.FromArgb(0xB0, 0xB0, 0xB0))) g.DrawRectangle(pn, c); }
                Ui.Str(g, gi.Preview ?? "AaBbCcĐđ", gi.Font, gi.Color, new Rectangle(c.X + 2, c.Y + 4, c.Width - 4, c.Height - 22), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                Ui.Str(g, gi.Name, Ui.F(8f), Ui.Text, new Rectangle(c.X, c.Bottom - 17, c.Width, 15), TextFormatFlags.HorizontalCenter);
            }
            Rectangle sb2 = new Rectangle(r.Right - 16, r.Y, 16, r.Height);
            using (SolidBrush bb = new SolidBrush(Color.FromArgb(0xF5, 0xF5, 0xF5))) g.FillRectangle(bb, sb2);
            int third = r.Height / 3;
            Ui.Arrow(g, Ui.Sub, sb2.X + 8, r.Y + third / 2, false); Ui.Arrow(g, Ui.Sub, sb2.X + 8, r.Y + third + third / 2, true);
            using (Pen pn = new Pen(Ui.Sub)) g.DrawLine(pn, sb2.X + 4, r.Y + 2 * third + 4, sb2.X + 12, r.Y + 2 * third + 4);
            Ui.Arrow(g, Ui.Sub, sb2.X + 8, r.Y + 2 * third + third / 2 + 2, true);
        }

        // ---------------- interaction ----------------
        object HitTest(Point p, out Rectangle arrow)
        {
            arrow = Rectangle.Empty;
            foreach (RTab t in Tabs) if (t.Bounds.Contains(p)) return t;
            if (shareBtn.Contains(p)) return "share"; if (commentsBtn.Contains(p)) return "comments"; if (!editModeBtn.IsEmpty && editModeBtn.Contains(p)) return "edit";
            if (Minimized && !expandedTemp) return null;
            if (collapseBtn.Contains(p)) return "collapse";
            if (Selected < 0 || Selected >= Tabs.Count) return null;
            foreach (RGroup g in Tabs[Selected].Groups)
            {
                if (!g.Bounds.Contains(p)) continue;
                if (g.Collapsed) return g;
                if (!g.LauncherBounds.IsEmpty && Rectangle.Inflate(g.LauncherBounds, 2, 2).Contains(p)) return g;
                foreach (RItem i in g.All())
                {
                    if (i.Kind == RK.Combo || !i.Bounds.Contains(p)) continue;
                    if (i.Kind == RK.Gallery)
                    {
                        Rectangle r = i.Bounds; if (p.X >= r.Right - 16) { arrow = new Rectangle(r.Right - 16, r.Y, 16, r.Height); return i; }
                        int k = (p.X - r.X - 2) / 78; if (i.GalleryFirst + k < i.Gallery.Count) return i.Gallery[i.GalleryFirst + k];
                        return null;
                    }
                    if (!i.Arrow.IsEmpty && i.Arrow.Contains(p)) arrow = i.Arrow;
                    return i;
                }
            }
            return null;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            Rectangle a; object h = HitTest(e.Location, out a);
            if (h != hot || a != hotArrow)
            {
                hot = h; hotArrow = a; Invalidate();
                RItem it = h as RItem; string t = null;
                if (it != null) t = string.IsNullOrEmpty(it.Tip) ? it.Text : (string.IsNullOrEmpty(it.Text) || it.Tip.StartsWith(it.Text) ? it.Tip : it.Text + "\n" + it.Tip);
                else if (h is RGroup && !((RGroup)h).Collapsed) t = ((RGroup)h).LauncherTip ?? ("Tùy chọn " + ((RGroup)h).Name.ToLower());
                if (t != null) tip.Show(t, this, e.X + 4, Height + 2, 5000); else tip.Hide(this);
            }
        }
        protected override void OnMouseLeave(EventArgs e) { hot = null; tip.Hide(this); Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { Rectangle a; down = HitTest(e.Location, out a); Invalidate(); }
        protected override void OnMouseDoubleClick(MouseEventArgs e) { Rectangle a; if (HitTest(e.Location, out a) is RTab && Tabs.IndexOf((RTab)HitTest(e.Location, out a)) > 0) SetMinimized(!Minimized); }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            Rectangle a; object h = HitTest(e.Location, out a); object d = down; down = null; Invalidate();
            if (h == null || h != d || e.Button != MouseButtons.Left) return;
            RTab t = h as RTab;
            if (t != null)
            {
                int idx = Tabs.IndexOf(t);
                if (idx == 0) { if (FileClicked != null) FileClicked(); return; }
                if (Minimized) { expandedTemp = !(expandedTemp && Selected == idx); Height = expandedTemp ? TabH + BodyH + 6 : TabH + 4; }
                Selected = idx; LayoutAll(); Invalidate(); return;
            }
            if ((h as string) == "share") { if (ShareClicked != null) ShareClicked(); return; }
            if ((h as string) == "comments") { if (CommentsClicked != null) CommentsClicked(); return; }
            if ((h as string) == "collapse") { SetMinimized(true); return; }
            if ((h as string) == "edit" && EditModeMenu != null) { ShowMenu(EditModeMenu, new Point(editModeBtn.X, editModeBtn.Bottom + 2)); return; }
            RGroup g = h as RGroup;
            if (g != null)
            {
                if (g.Collapsed) { List<RMenu> m = new List<RMenu>(); foreach (RItem i in g.All()) { if (i.Kind == RK.Combo || i.Kind == RK.Gallery) continue; RItem ii = i; m.Add(new RMenu(string.IsNullOrEmpty(i.Text) ? i.Tip : i.Text, i.Icon, () => Run(ii, false, ii.Bounds))); } ShowMenu(m, new Point(g.Bounds.X, g.Bounds.Bottom)); }
                else if (g.Launcher != null) g.Launcher();
                return;
            }
            RGalleryItem gi = h as RGalleryItem; if (gi != null) { if (gi.Click != null) gi.Click(); return; }
            RItem it = h as RItem;
            if (it != null)
            {
                if (it.Kind == RK.Gallery)
                {
                    int third = it.Bounds.Height / 3, n = Math.Max(1, (it.Bounds.Width - 18) / 78);
                    if (e.Y < it.Bounds.Y + third) it.GalleryFirst = Math.Max(0, it.GalleryFirst - n);
                    else if (e.Y < it.Bounds.Y + 2 * third) it.GalleryFirst = Math.Min(Math.Max(0, it.Gallery.Count - n), it.GalleryFirst + n);
                    else { List<RMenu> m = new List<RMenu>(); foreach (RGalleryItem x in it.Gallery) { RGalleryItem xx = x; m.Add(new RMenu(x.Name, null, x.Click)); } ShowMenu(m, new Point(it.Bounds.X, it.Bounds.Bottom)); }
                    Invalidate(); return;
                }
                Run(it, !a.IsEmpty, it.Bounds);
            }
        }

        void Run(RItem it, bool onArrow, Rectangle r)
        {
            if (it.Disabled) return;
            Point below = new Point(r.X, r.Bottom + 2);
            if (it.PickColor != null && (onArrow || it.Click == null)) { ColorPopup.Show(this, below, it.Palette, c => { it.Swatch = c; Invalidate(); it.PickColor(c); }); return; }
            if (it.Menu != null && (onArrow || !it.Split || it.Click == null)) { ShowMenu(it.Menu, below); return; }
            if (it.Click != null) it.Click();
        }

        public void ShowMenu(List<RMenu> items, Point at)
        {
            ContextMenuStrip m = OfficeMenu.Build(items);
            m.Show(this, at);
        }
    }

    /// <summary>Light, Office-like drop-down menus.</summary>
    static class OfficeMenu
    {
        class R : ToolStripProfessionalRenderer
        {
            public R() : base(new C()) { RoundedEdges = false; }
            protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e) { e.TextColor = e.Item.Enabled ? Ui.Text : Color.Gray; base.OnRenderItemText(e); }
        }
        class C : ProfessionalColorTable
        {
            public override Color MenuItemSelected { get { return Ui.Hover; } }
            public override Color MenuItemBorder { get { return Ui.Hover; } }
            public override Color MenuBorder { get { return Color.FromArgb(0xC8, 0xC8, 0xC8); } }
            public override Color ToolStripDropDownBackground { get { return Color.White; } }
            public override Color ImageMarginGradientBegin { get { return Color.White; } }
            public override Color ImageMarginGradientMiddle { get { return Color.White; } }
            public override Color ImageMarginGradientEnd { get { return Color.White; } }
            public override Color SeparatorDark { get { return Ui.Line; } }
            public override Color MenuItemSelectedGradientBegin { get { return Ui.Hover; } }
            public override Color MenuItemSelectedGradientEnd { get { return Ui.Hover; } }
            public override Color CheckBackground { get { return Ui.Checked; } }
            public override Color CheckSelectedBackground { get { return Ui.Press; } }
        }
        public static readonly ToolStripRenderer Renderer = new R();
        public static Image Img(string icon) { if (!Icons.Has(icon)) return null; Bitmap b = new Bitmap(16, 16); using (Graphics g = Graphics.FromImage(b)) Icons.Draw(g, icon, new Rectangle(0, 0, 16, 16), true); return b; }
        public static ContextMenuStrip Build(List<RMenu> items)
        {
            ContextMenuStrip m = new ContextMenuStrip { Renderer = Renderer, Font = Ui.F(9f), ShowImageMargin = true };
            foreach (RMenu i in items)
            {
                if (i.Separator) { m.Items.Add(new ToolStripSeparator()); continue; }
                RMenu ii = i;
                ToolStripMenuItem t = new ToolStripMenuItem(i.Text, Img(i.Icon)) { Padding = new Padding(2, 3, 2, 3) };
                if (i.Checked != null) t.Checked = i.Checked();
                if (i.Click == null) t.Enabled = false; else t.Click += (s, e) => ii.Click();
                m.Items.Add(t);
            }
            m.Closed += (s, e) => m.BeginInvoke(new Action(m.Dispose));
            return m;
        }
    }

    /// <summary>Theme colour palette pop-up (10 theme colours x 6 tints + standard colours), like Office's font colour picker.</summary>
    class ColorPopup : Control
    {
        static readonly Color[] Theme = { Color.White, Color.Black, Color.FromArgb(0xE7, 0xE6, 0xE6), Color.FromArgb(0x44, 0x54, 0x6A), Color.FromArgb(0x44, 0x72, 0xC4), Color.FromArgb(0xED, 0x7D, 0x31), Color.FromArgb(0xA5, 0xA5, 0xA5), Color.FromArgb(0xFF, 0xC0, 0x00), Color.FromArgb(0x5B, 0x9B, 0xD5), Color.FromArgb(0x70, 0xAD, 0x47) };
        public static readonly Color[] Standard = { Color.FromArgb(0xC0, 0, 0), Color.Red, Color.FromArgb(0xFF, 0xC0, 0), Color.Yellow, Color.FromArgb(0x92, 0xD0, 0x50), Color.FromArgb(0, 0xB0, 0x50), Color.FromArgb(0, 0xB0, 0xF0), Color.FromArgb(0, 0x70, 0xC0), Color.FromArgb(0, 0x20, 0x60), Color.FromArgb(0x70, 0x30, 0xA0) };
        public static readonly Color[] Highlight = { Color.Yellow, Color.Lime, Color.Cyan, Color.Magenta, Color.Blue, Color.Red, Color.Navy, Color.Teal, Color.Green, Color.Purple, Color.Maroon, Color.Olive, Color.Gray, Color.Silver, Color.Black };
        readonly List<KeyValuePair<Rectangle, Color>> cells = new List<KeyValuePair<Rectangle, Color>>();
        Rectangle noneRect; int hotIdx = -1; Action<Color> pick; ToolStripDropDown host;

        public static void Show(Control owner, Point at, Color[] palette, Action<Color> onPick)
        {
            ColorPopup p = new ColorPopup(palette) { pick = onPick };
            ToolStripControlHost h = new ToolStripControlHost(p) { Padding = Padding.Empty, Margin = Padding.Empty, AutoSize = false, Size = p.Size };
            ToolStripDropDown dd = new ToolStripDropDown { Padding = new Padding(1), Renderer = OfficeMenu.Renderer }; dd.Items.Add(h); p.host = dd;
            dd.Closed += (s, e) => dd.BeginInvoke(new Action(dd.Dispose));
            dd.Show(owner, at);
        }

        ColorPopup(Color[] palette)
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
            BackColor = Color.White; Font = Ui.F(9f);
            int x0 = 10, y = 30, s = 16, gap = 4;
            if (palette != null)
            {
                for (int i = 0; i < palette.Length; i++) cells.Add(new KeyValuePair<Rectangle, Color>(new Rectangle(x0 + (i % 5) * (s + gap + 6), y + (i / 5) * (s + gap + 4), s + 6, s + 4), palette[i]));
                Size = new Size(10 + 5 * (s + gap + 6) + 6, y + ((palette.Length + 4) / 5) * (s + gap + 4) + 10);
            }
            else
            {
                for (int c = 0; c < 10; c++)
                {
                    cells.Add(new KeyValuePair<Rectangle, Color>(new Rectangle(x0 + c * (s + gap), y, s, s), Theme[c]));
                    for (int t = 1; t <= 5; t++)
                    {
                        Color b = Theme[c]; bool dark = b.GetBrightness() < 0.2f, white = b.GetBrightness() > 0.95f;
                        Color v = white ? Ui.Light(Color.Black, 0.05f + (5 - t) * 0.15f + 0.1f) : dark ? Ui.Light(b, 0.15f * (6 - t)) : t <= 3 ? Ui.Light(b, 0.8f - (t - 1) * 0.2f) : Color.FromArgb((int)(b.R * (t == 4 ? 0.75 : 0.5)), (int)(b.G * (t == 4 ? 0.75 : 0.5)), (int)(b.B * (t == 4 ? 0.75 : 0.5)));
                        cells.Add(new KeyValuePair<Rectangle, Color>(new Rectangle(x0 + c * (s + gap), y + 22 + (t - 1) * (s + 1), s, s), v));
                    }
                }
                int sy = y + 22 + 5 * (s + 1) + 28;
                for (int c = 0; c < 10; c++) cells.Add(new KeyValuePair<Rectangle, Color>(new Rectangle(x0 + c * (s + gap), sy, s, s), Standard[c]));
                Size = new Size(x0 * 2 + 10 * (s + gap) - gap, sy + s + 40);
            }
            noneRect = new Rectangle(8, Height - 30, Width - 16, 24);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(Color.White);
            Ui.Str(g, cells.Count > 20 ? "Màu chủ đề" : "Màu tô sáng", Ui.F(9f, FontStyle.Bold), Ui.Text, new Rectangle(10, 6, Width - 20, 20), TextFormatFlags.VerticalCenter);
            if (cells.Count > 20) Ui.Str(g, "Màu chuẩn", Ui.F(9f, FontStyle.Bold), Ui.Text, new Rectangle(10, cells[cells.Count - 1].Key.Y - 24, Width - 20, 20), TextFormatFlags.VerticalCenter);
            for (int i = 0; i < cells.Count; i++)
            {
                using (SolidBrush b = new SolidBrush(cells[i].Value)) g.FillRectangle(b, cells[i].Key);
                using (Pen p = new Pen(i == hotIdx ? Color.FromArgb(0xE8, 0x6B, 0x1E) : Color.FromArgb(0xC8, 0xC8, 0xC8), i == hotIdx ? 2f : 1f)) g.DrawRectangle(p, cells[i].Key);
            }
            if (hotIdx == -2) Ui.FillRound(g, Ui.Hover, noneRect, 3);
            using (Pen p = new Pen(Color.Gray)) g.DrawRectangle(p, noneRect.X + 6, noneRect.Y + 5, 14, 14);
            using (Pen p = new Pen(Color.Red, 1.5f)) g.DrawLine(p, noneRect.X + 6, noneRect.Y + 19, noneRect.X + 20, noneRect.Y + 5);
            Ui.Str(g, "Không màu / Tự động", Font, Ui.Text, new Rectangle(noneRect.X + 28, noneRect.Y, noneRect.Width - 30, noneRect.Height), TextFormatFlags.VerticalCenter);
        }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            int h = -1; for (int i = 0; i < cells.Count; i++) if (cells[i].Key.Contains(e.Location)) h = i;
            if (noneRect.Contains(e.Location)) h = -2; if (h != hotIdx) { hotIdx = h; Invalidate(); }
        }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            Color c = Color.Empty; bool ok = false;
            for (int i = 0; i < cells.Count; i++) if (cells[i].Key.Contains(e.Location)) { c = cells[i].Value; ok = true; }
            if (noneRect.Contains(e.Location)) { c = Color.Transparent; ok = true; }
            if (!ok) return; host.Close(); if (pick != null) pick(c);
        }
    }
}
