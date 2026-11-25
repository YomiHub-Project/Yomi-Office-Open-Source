// Office Open Source v0.1.6 - Office-style chrome for the native Word / Excel / PowerPoint editors.
// Layout follows the reference screenshots in "Office Screen.docx": title bar (AutoSave, Quick Access, search, account,
// caption buttons), tabbed Ribbon with large/small commands and group labels, Backstage (File) view, status bar with
// view buttons and zoom slider. Everything is owner-drawn GDI+ (works on Windows 7 .NET 4 without extra fonts);
// icons are small vector drawings defined below (own artwork, no third-party logos).
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficeOpenSource
{
    static class Ui
    {
        public static readonly Color Chrome = Color.FromArgb(0xF3, 0xF3, 0xF3);
        public static readonly Color Card = Color.FromArgb(0xFB, 0xFB, 0xFB);
        public static readonly Color Line = Color.FromArgb(0xE0, 0xE0, 0xE0);
        public static readonly Color Text = Color.FromArgb(0x24, 0x24, 0x24);
        public static readonly Color Sub = Color.FromArgb(0x61, 0x61, 0x61);
        public static readonly Color Hover = Color.FromArgb(0xE6, 0xE6, 0xE6);
        public static readonly Color Press = Color.FromArgb(0xD4, 0xD4, 0xD4);
        public static readonly Color Checked = Color.FromArgb(0xDA, 0xDA, 0xDA);
        public static readonly Color Workspace = Color.FromArgb(0xE9, 0xE9, 0xE9);
        public static Color AccentFor(int docType)
        {
            if (docType == LokDoc.SPREADSHEET) return Color.FromArgb(0x10, 0x7C, 0x41);
            if (docType == LokDoc.PRESENTATION) return Color.FromArgb(0xC4, 0x3E, 0x1C);
            return Color.FromArgb(0x18, 0x5A, 0xBD);
        }
        public static Color Light(Color c, float k)
        {
            if (k < 0) return Color.FromArgb(Clamp(c.R * (1 + k)), Clamp(c.G * (1 + k)), Clamp(c.B * (1 + k)));
            return Color.FromArgb(Clamp(c.R + (255 - c.R) * k), Clamp(c.G + (255 - c.G) * k), Clamp(c.B + (255 - c.B) * k));
        }
        static int Clamp(float v) { return Math.Max(0, Math.Min(255, (int)v)); }

        static readonly Dictionary<string, Font> fonts = new Dictionary<string, Font>();
        public static Font F(float size) { return F(size, FontStyle.Regular); }
        public static Font F(float size, FontStyle st)
        {
            string k = size.ToString(CultureInfo.InvariantCulture) + "|" + (int)st; Font f;
            if (!fonts.TryGetValue(k, out f)) { f = new Font("Segoe UI", size, st, GraphicsUnit.Point); fonts[k] = f; }
            return f;
        }
        public static GraphicsPath Round(Rectangle r, int rad)
        {
            GraphicsPath p = new GraphicsPath(); int d = Math.Max(1, rad * 2);
            if (rad <= 0) { p.AddRectangle(r); return p; }
            p.AddArc(r.X, r.Y, d, d, 180, 90); p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90); p.AddArc(r.X, r.Bottom - d, d, d, 90, 90); p.CloseFigure(); return p;
        }
        public static void FillRound(Graphics g, Color c, Rectangle r, int rad)
        {
            SmoothingMode sm = g.SmoothingMode; g.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath p = Round(r, rad)) using (SolidBrush b = new SolidBrush(c)) g.FillPath(b, p);
            g.SmoothingMode = sm;
        }
        public static void DrawRound(Graphics g, Color c, Rectangle r, int rad, float w)
        {
            SmoothingMode sm = g.SmoothingMode; g.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath p = Round(r, rad)) using (Pen pn = new Pen(c, w)) g.DrawPath(pn, p);
            g.SmoothingMode = sm;
        }
        public static Size Measure(string s, Font f) { return TextRenderer.MeasureText(s ?? "", f, new Size(9999, 9999), TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix); }
        public static void Str(Graphics g, string s, Font f, Color c, Rectangle r, TextFormatFlags fl) { TextRenderer.DrawText(g, s ?? "", f, r, c, fl | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix); }
        public static void Arrow(Graphics g, Color c, int cx, int cy, bool down)
        {
            using (SolidBrush b = new SolidBrush(c))
                g.FillPolygon(b, down ? new[] { new Point(cx - 3, cy - 1), new Point(cx + 4, cy - 1), new Point(cx, cy + 3) } : new[] { new Point(cx - 3, cy + 2), new Point(cx + 4, cy + 2), new Point(cx, cy - 2) });
        }
        public static string Greeting() { int h = DateTime.Now.Hour; return h < 11 ? "Chào buổi sáng" : h < 13 ? "Chào buổi trưa" : h < 18 ? "Chào buổi chiều" : "Chào buổi tối"; }
        public static string UserName()
        {
            string n = NativeSettings.Get("user", ""); if (n.Length > 0) return n;
            try { return Environment.UserName; } catch { return "Người dùng"; }
        }
        public static string Initials(string n)
        {
            string[] p = (n ?? "").Split(new[] { ' ', '.', '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (p.Length == 0) return "?"; if (p.Length == 1) return p[0].Substring(0, Math.Min(2, p[0].Length)).ToUpperInvariant();
            return (p[0].Substring(0, 1) + p[p.Length - 1].Substring(0, 1)).ToUpperInvariant();
        }
    }

    /// <summary>Small persisted settings for the native editors (Options dialog, AutoSave switch, recent files).</summary>
    static class NativeSettings
    {
        static Dictionary<string, string> map;
        static string FilePath { get { return App.DataDir != null ? Path.Combine(App.DataDir, "native-settings.ini") : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOpenSource\\native-settings.ini"); } }   // v0.1.7: portable keeps its settings in .\\data
        static void Load()
        {
            if (map != null) return; map = new Dictionary<string, string>();
            try { foreach (string l in File.ReadAllLines(FilePath)) { int i = l.IndexOf('='); if (i > 0) map[l.Substring(0, i)] = l.Substring(i + 1); } } catch { }
            SharedSettings.Overlay(map);   // v0.1.7: options changed in the Web editor (settings.json)
        }
        public static bool Has(string k) { Load(); return map.ContainsKey(k); }
        public static string Get(string k, string def) { Load(); string v; return map.TryGetValue(k, out v) ? v : def; }
        public static bool Bool(string k, bool def) { string v = Get(k, def ? "1" : "0"); return v == "1" || v == "true"; }
        public static int Int(string k, int def) { int v; return int.TryParse(Get(k, ""), out v) ? v : def; }
        public static void Set(string k, string v) { SetLocal(k, v); SharedSettings.Mirror(k, (v ?? "").Replace("\r", " ").Replace("\n", " ")); }
        public static void SetLocal(string k, string v)
        {
            Load(); map[k] = (v ?? "").Replace("\r", " ").Replace("\n", " ");
            try { Directory.CreateDirectory(Path.GetDirectoryName(FilePath)); List<string> o = new List<string>(); foreach (KeyValuePair<string, string> kv in map) o.Add(kv.Key + "=" + kv.Value); File.WriteAllLines(FilePath, o.ToArray()); } catch { }
        }
        public static void Set(string k, bool v) { Set(k, v ? "1" : "0"); }
        public static List<string> Recent()
        {
            List<string> r = new List<string>();
            foreach (string s in Get("recent", "").Split('|')) if (s.Length > 0 && File.Exists(s) && !r.Contains(s)) r.Add(s);
            return r;
        }
        public static void AddRecent(string path)
        {
            List<string> r = Recent(); r.Remove(path); r.Insert(0, path); if (r.Count > 30) r.RemoveRange(30, r.Count - 30);
            Set("recent", string.Join("|", r.ToArray()));
        }
        public static List<string> Pinned() { List<string> r = new List<string>(); foreach (string s in Get("pinned", "").Split('|')) if (s.Length > 0 && File.Exists(s)) r.Add(s); return r; }
        public static void TogglePin(string path) { List<string> r = Pinned(); if (r.Contains(path)) r.Remove(path); else r.Insert(0, path); Set("pinned", string.Join("|", r.ToArray())); }
    }

    /// <summary>Vector icons drawn from tiny scripts in a 32x32 design grid (scaled for 16/20/32 px).</summary>
    static class Icons
    {
        // ops: r x y w h fill stroke [w] | q x y w h rad fill stroke [w] | o x y w h fill stroke [w] | l x1 y1 x2 y2 color w
        //      p fill stroke w x,y x,y ... | t cx cy size color style text | a x y w h start sweep color w
        // colors: #rrggbb, '-' none, or palette letters below
        static readonly Dictionary<char, Color> pal = new Dictionary<char, Color> {
            { 'D', Color.FromArgb(0x5C,0x5C,0x5C) }, { 'K', Color.FromArgb(0x24,0x24,0x24) }, { 'W', Color.White },
            { 'B', Color.FromArgb(0x1E,0x6F,0xD0) }, { 'b', Color.FromArgb(0xCF,0xE3,0xFA) }, { 'G', Color.FromArgb(0x1F,0x9D,0x55) }, { 'g', Color.FromArgb(0xCD,0xEB,0xD7) },
            { 'O', Color.FromArgb(0xE8,0x6B,0x1E) }, { 'o', Color.FromArgb(0xFD,0xE4,0xCF) }, { 'R', Color.FromArgb(0xD1,0x34,0x38) }, { 'r', Color.FromArgb(0xF7,0xD3,0xD5) },
            { 'Y', Color.FromArgb(0xF2,0xB2,0x00) }, { 'y', Color.FromArgb(0xFF,0xF1,0xC1) }, { 'P', Color.FromArgb(0x7A,0x50,0xC1) }, { 'p', Color.FromArgb(0xE6,0xDC,0xF7) },
            { 'N', Color.FromArgb(0x8B,0x5E,0x3C) }, { 'n', Color.FromArgb(0xE8,0xD2,0xB5) }, { 'S', Color.FromArgb(0x9E,0x9E,0x9E) }, { 's', Color.FromArgb(0xEC,0xEC,0xEC) },
        };
        public static Dictionary<string, string> Specs { get { return spec; } }   // v0.1.7: exported to the Web UI (same icons)
        public static Dictionary<char, Color> Palette { get { return pal; } }
        static readonly Dictionary<string, string> spec = new Dictionary<string, string>
        {
            { "animpane", "r 4 4 24 24 W D 1.2; r 7 8 4 4 O -; l 13 10 25 10 D 1.5; r 7 15 4 4 B -; l 13 17 25 17 D 1.5; r 7 22 4 4 G -; l 13 24 22 24 D 1.5" },
            { "mail", "r 3 7 26 18 W D 1.5; p - B 1.5 3,7 16,18 29,7" },
            { "mergefield", "q 3 9 26 14 2 b B 1.2; t 16 16 11 B b «»" },
            { "mailmerge", "r 3 4 17 21 W D 1.2; r 12 9 17 21 W B 1.2; l 15 14 26 14 B 1.2; l 15 18 26 18 B 1.2; l 15 22 23 22 B 1.2" },
            { "paste", "q 6 4 20 25 2 n N 1.5; r 12 2 8 5 S D 1.2; r 11 9 15 19 W D 1.2; l 14 14 23 14 S 1.5; l 14 18 23 18 S 1.5; l 14 22 20 22 S 1.5" },
            { "cut", "o 4 19 9 9 - D 2; o 19 19 9 9 - D 2; l 11 20 22 4 D 2; l 21 20 10 4 D 2" },
            { "copy", "r 5 4 15 18 W D 1.5; r 12 10 15 18 W B 1.5; l 15 15 24 15 B 1.2; l 15 19 24 19 B 1.2; l 15 23 21 23 B 1.2" },
            { "painter", "q 7 4 18 8 2 Y N 1.5; l 16 12 16 17 D 2; r 13 17 6 11 N D 1.2" },
            { "bold", "t 16 16 22 K b B" }, { "italic", "t 16 16 22 K bi I" }, { "underline", "t 16 14 20 K n U; l 8 27 24 27 K 2" },
            { "strike", "t 16 15 18 K n ab; l 5 16 27 16 K 2" }, { "sub", "t 12 14 19 K n x; t 23 23 11 B b 2" }, { "sup", "t 12 18 19 K n x; t 23 8 11 B b 2" },
            { "grow", "t 12 17 21 K n A; p B - 0 21,11 31,11 26,4" }, { "shrink", "t 12 18 16 K n A; p B - 0 21,5 31,5 26,12" },
            { "fontcolor", "t 16 13 19 K n A; r 5 24 22 5 R -" }, { "highlight", "p y D 1.2 9,19 19,6 25,11 15,24; p D - 0 9,19 15,24 8,25; r 5 26 22 4 Y -" },
            { "clearfmt", "t 12 14 19 K n A; q 17 16 12 9 2 r R 1.2" }, { "case", "t 16 16 17 K n Aa" },
            { "alignleft", "l 5 8 27 8 K 2; l 5 13 19 13 K 2; l 5 18 27 18 K 2; l 5 23 19 23 K 2" },
            { "aligncenter", "l 5 8 27 8 K 2; l 9 13 23 13 K 2; l 5 18 27 18 K 2; l 9 23 23 23 K 2" },
            { "alignright", "l 5 8 27 8 K 2; l 13 13 27 13 K 2; l 5 18 27 18 K 2; l 13 23 27 23 K 2" },
            { "justify", "l 5 8 27 8 K 2; l 5 13 27 13 K 2; l 5 18 27 18 K 2; l 5 23 27 23 K 2" },
            { "aligntop", "l 4 5 28 5 K 2; l 10 10 22 10 B 2; l 10 15 22 15 B 2" }, { "alignmiddle", "l 10 11 22 11 B 2; l 10 16 22 16 B 2; l 10 21 22 21 B 2; l 4 5 28 5 S 1; l 4 27 28 27 S 1" },
            { "alignbottom", "l 4 27 28 27 K 2; l 10 17 22 17 B 2; l 10 22 22 22 B 2" },
            { "bullets", "o 4 6 4 4 K -; l 11 8 27 8 K 2; o 4 14 4 4 K -; l 11 16 27 16 K 2; o 4 22 4 4 K -; l 11 24 27 24 K 2" },
            { "numbering", "t 6 8 8 B b 1; l 12 8 27 8 K 2; t 6 16 8 B b 2; l 12 16 27 16 K 2; t 6 24 8 B b 3; l 12 24 27 24 K 2" },
            { "indentdec", "l 14 8 27 8 K 2; l 14 14 27 14 K 2; l 14 20 27 20 K 2; l 4 26 27 26 K 2; p B - 0 10,9 10,19 4,14" },
            { "indentinc", "l 14 8 27 8 K 2; l 14 14 27 14 K 2; l 14 20 27 20 K 2; l 4 26 27 26 K 2; p B - 0 4,9 4,19 10,14" },
            { "linespacing", "l 14 8 28 8 K 2; l 14 14 28 14 K 2; l 14 20 28 20 K 2; l 14 26 28 26 K 2; l 6 7 6 27 B 1.5; p B - 0 2,10 10,10 6,4; p B - 0 2,24 10,24 6,30" },
            { "pilcrow", "t 16 16 21 K n ¶" },
            { "sortaz", "t 9 9 10 B b A; t 9 22 10 B b Z; l 22 5 22 26 K 2; p K - 0 17,21 27,21 22,28" },
            { "find", "o 4 4 17 17 - D 2.5; l 18 18 27 27 D 3" },
            { "replace", "t 10 10 11 K n ab; t 22 24 11 B n ac; l 9 18 9 24 B 1.5; l 9 24 14 24 B 1.5; p B - 0 13,21 17,24 13,27" },
            { "select", "p W K 1.5 9,4 9,26 15,20 19,28 22,26 18,19 26,19" },
            { "dictate", "q 12 3 9 16 4 B -; a 7 8 19 16 0 180 D 2; l 16.5 24 16.5 29 D 2" },
            { "editor", "p B - 0 6,26 8,18 22,4 28,10 14,24; l 6 28 26 28 D 1.5" },
            { "addins", "r 5 5 10 10 O -; r 17 5 10 10 R -; r 5 17 10 10 R -; r 17 17 10 10 O -" },
            { "designer", "r 4 6 24 20 b B 1.5; r 7 9 8 6 B -; l 18 10 25 10 B 1.5; l 18 14 25 14 B 1.5; l 7 19 25 19 B 1.5" },
            { "coverpage", "r 7 3 18 26 W D 1.5; r 7 3 18 9 B -; l 10 17 22 17 S 1.5; l 10 21 22 21 S 1.5" },
            { "blankpage", "p W D 1.5 7,3 19,3 25,9 25,29 7,29; p s D 1 19,3 19,9 25,9" },
            { "pagebreak", "r 7 2 18 11 W D 1.5; r 7 19 18 11 W D 1.5; l 3 16 29 16 B 1.5" },
            { "table", "r 4 5 24 22 W D 1.5; r 4 5 24 6 B -; l 4 16 28 16 D 1; l 4 21 28 21 D 1; l 12 5 12 27 D 1; l 20 5 20 27 D 1" },
            { "picture", "r 3 6 26 20 b B 1.5; p G - 0 4,25 12,15 18,21 22,17 28,25; o 20 9 5 5 Y -" },
            { "shapes", "o 3 3 13 13 b B 1.5; r 15 14 13 13 o O 1.5; p g G 1.5 4,28 10,18 16,28" },
            { "icons", "o 11 4 10 10 B -; p B - 0 6,28 8,18 24,18 26,28" },
            { "smartart", "r 11 3 10 7 G -; r 3 21 10 7 G -; r 19 21 10 7 G -; l 16 10 16 16 D 1.5; l 8 16 24 16 D 1.5; l 8 16 8 21 D 1.5; l 24 16 24 21 D 1.5" },
            { "chart", "r 5 16 5 11 B -; r 13 9 5 18 O -; r 21 4 5 23 G -; l 3 28 29 28 D 1.5" },
            { "link", "q 3 11 15 10 5 - B 2.5; q 14 11 15 10 5 - D 2.5" },
            { "bookmark", "p B - 0 9,3 23,3 23,29 16,23 9,29" },
            { "comment", "q 3 5 26 18 3 y Y 1.5; p y Y 1.5 9,22 9,29 16,22; l 8 11 24 11 D 1.2; l 8 16 20 16 D 1.2" },
            { "header", "r 6 3 20 26 W D 1.5; r 9 6 14 5 B -; l 9 15 23 15 S 1.2; l 9 19 23 19 S 1.2" },
            { "footer", "r 6 3 20 26 W D 1.5; l 9 8 23 8 S 1.2; l 9 12 23 12 S 1.2; r 9 21 14 5 B -" },
            { "pagenum", "r 6 3 20 26 W D 1.5; t 16 17 12 B b #" },
            { "textbox", "r 4 5 24 22 W D 1.5; t 16 16 14 K n A" },
            { "wordart", "t 16 16 23 B bi A" },
            { "datetime", "r 4 6 24 22 W D 1.5; r 4 6 24 6 R -; l 10 3 10 9 D 2; l 22 3 22 9 D 2; r 8 16 4 3 D -; r 14 16 4 3 D -; r 20 16 4 3 D -; r 8 21 4 3 D -" },
            { "equation", "t 16 16 21 K n π" }, { "ai", "t 16 16 22 B n ✦" }, { "update", "t 16 16 22 B n ⟳" }, { "symbol", "t 16 16 21 K n Ω" },
            { "object", "r 4 4 24 24 W D 1.5; r 8 8 8 8 b B 1; o 17 17 8 8 o O 1" },
            { "footnote", "t 11 12 15 K n ab; t 23 8 9 B b 1; l 4 23 16 23 D 1.5; l 4 27 26 27 S 1.5" },
            { "toc", "l 4 7 28 7 B 2; l 8 13 28 13 D 1.5; l 8 18 28 18 D 1.5; l 4 24 28 24 B 2" },
            { "refresh", "a 5 5 22 22 20 300 G 2.5; p G - 0 22,1 29,7 21,10" },
            { "watermark", "r 7 3 18 26 W D 1.5; t 16 16 17 S b A" },
            { "pagecolor", "r 7 3 18 26 y D 1.5; r 4 25 24 4 Y -" },
            { "margins", "r 6 3 20 26 W D 1.5; r 9 7 14 18 b B 1" },
            { "orientation", "r 4 3 14 20 W D 1.5; r 12 13 17 13 W B 1.5" },
            { "size", "r 7 3 18 26 W D 1.5; t 16 16 9 B b A4" },
            { "columns", "r 4 4 24 24 W D 1.5; l 8 9 14 9 D 1.2; l 8 13 14 13 D 1.2; l 8 17 14 17 D 1.2; l 8 21 14 21 D 1.2; l 18 9 24 9 D 1.2; l 18 13 24 13 D 1.2; l 18 17 24 17 D 1.2; l 18 21 24 21 D 1.2" },
            { "spelling", "t 14 11 10 K b ABC; p - G 2.5 8,21 13,26 25,14" },
            { "thesaurus", "q 4 4 24 23 2 b B 1.5; l 16 4 16 27 B 1.5; l 7 10 13 10 B 1.2; l 19 10 25 10 B 1.2" },
            { "wordcount", "t 16 16 12 K b 123" },
            { "track", "r 6 3 20 26 W D 1.5; l 9 9 23 9 S 1.2; p - R 2 10,21 16,15 20,19; l 9 25 23 25 R 1.5" },
            { "accept", "r 6 3 20 26 W D 1.5; p - G 3 9,17 14,22 24,10" },
            { "reject", "r 6 3 20 26 W D 1.5; l 10 11 22 23 R 3; l 22 11 10 23 R 3" },
            { "protect", "q 7 14 18 14 2 Y N 1.5; a 10 5 12 18 180 180 D 2.5" },
            { "readmode", "q 3 6 12 20 1 W D 1.5; q 17 6 12 20 1 W D 1.5" },
            { "printlayout", "r 8 3 16 26 W D 1.5; l 11 9 21 9 S 1.5; l 11 13 21 13 S 1.5; l 11 17 21 17 S 1.5" },
            { "weblayout", "o 4 4 24 24 b B 1.5; l 4 16 28 16 B 1.2; o 10 4 12 24 - B 1.2" },
            { "ruler", "r 3 11 26 10 y Y 1.5; l 8 11 8 15 D 1.2; l 13 11 13 17 D 1.2; l 18 11 18 15 D 1.2; l 23 11 23 17 D 1.2" },
            { "zoom", "o 4 4 17 17 W D 2.5; l 18 18 27 27 D 3; l 8 12.5 17 12.5 B 2; l 12.5 8 12.5 17 B 2" },
            { "zoom100", "t 16 16 12 B b 100" },
            { "onepage", "r 9 4 14 24 W D 1.5" }, { "multipage", "r 3 6 11 20 W D 1.5; r 18 6 11 20 W D 1.5" },
            { "pagewidth", "r 4 4 24 24 W D 1.5; l 6 16 26 16 B 2; p B - 0 5,16 10,12 10,20; p B - 0 27,16 22,12 22,20" },
            { "newwindow", "r 3 5 20 16 W D 1.5; r 9 11 20 16 W B 1.5" },
            { "macros", "r 4 6 24 20 W D 1.5; p G - 0 12,11 12,21 21,16" },
            { "fullscreen", "r 3 6 26 20 W D 1.5; l 7 10 12 10 B 2; l 7 10 7 15 B 2; l 25 22 20 22 B 2; l 25 22 25 17 B 2" },
            { "condfmt", "r 4 4 24 24 W D 1.2; r 7 7 7 5 R -; r 7 14 7 5 Y -; r 7 21 7 5 G -; l 17 9 25 9 S 2; l 17 16 25 16 S 2; l 17 23 25 23 S 2" },
            { "formattable", "r 4 5 24 22 W D 1.5; r 4 5 24 5 G -; r 4 15 24 4 g -; r 4 23 24 3 g -" },
            { "cellstyles", "r 4 7 11 8 g G 1.2; r 17 7 11 8 r R 1.2; r 4 18 11 8 y Y 1.2; r 17 18 11 8 b B 1.2" },
            { "insertcells", "r 4 8 20 18 W D 1.5; l 4 14 24 14 D 1; l 4 20 24 20 D 1; l 14 8 14 26 D 1; o 20 1 11 11 G -; l 25.5 3.5 25.5 9.5 W 1.8; l 22.5 6.5 28.5 6.5 W 1.8" },
            { "deletecells", "r 4 8 20 18 W D 1.5; l 4 14 24 14 D 1; l 4 20 24 20 D 1; l 14 8 14 26 D 1; o 20 1 11 11 R -; l 23 4 28 9 W 1.8; l 28 4 23 9 W 1.8" },
            { "formatcells", "r 4 8 20 18 W D 1.5; l 4 14 24 14 D 1; l 4 20 24 20 D 1; l 14 8 14 26 D 1; o 20 1 11 11 B -; o 23 4 5 5 - W 1.5" },
            { "autosum", "t 16 16 23 K n Σ" },
            { "fill", "r 9 3 14 13 W D 1.2; p B - 0 10,18 22,18 16,28" },
            { "clear", "q 5 14 22 11 2 r R 1.5; l 14 14 22 25 R 1.2" },
            { "filter", "p y D 1.5 4,5 28,5 19,16 19,26 13,29 13,16" },
            { "currency", "t 16 16 17 G b ₫" }, { "percent", "t 16 16 18 K b %" }, { "thousands", "t 16 16 12 K b 000" },
            { "incdec", "t 16 13 10 K n .0; t 16 24 10 B b .00" }, { "decdec", "t 16 13 10 K n .00; t 16 24 10 B b .0" },
            { "merge", "r 3 8 26 16 W D 1.5; p B - 0 7,16 12,12 12,20; p B - 0 25,16 20,12 20,20; l 12 16 20 16 B 1.5" },
            { "wrap", "l 4 8 27 8 K 2; l 4 15 22 15 K 2; a 16 11 11 9 270 180 B 2; l 21 20 15 20 B 2; p B - 0 16,17 12,20 16,23" },
            { "function", "t 16 16 18 K i fx" },
            { "pivot", "r 4 4 24 24 W D 1.5; r 4 4 24 6 G -; r 4 4 7 24 G -" },
            { "recchart", "r 4 3 24 20 W D 1.2; r 8 13 4 8 B -; r 14 9 4 12 O -; r 20 6 4 15 G -; l 8 27 24 27 B 2" },
            { "freeze", "r 4 4 24 24 W D 1.5; r 4 4 24 6 b -; r 4 4 7 24 b -; l 4 10 28 10 B 1.5; l 11 4 11 28 B 1.5" },
            { "gridlines", "r 4 4 24 24 W D 1.5; l 4 12 28 12 S 1.2; l 4 20 28 20 S 1.2; l 12 4 12 28 S 1.2; l 20 4 20 28 S 1.2" },
            { "calc", "r 6 3 20 26 s D 1.5; r 9 6 14 6 G -; r 9 15 4 4 D -; r 14 15 4 4 D -; r 19 15 4 4 D -; r 9 21 4 4 D -; r 14 21 4 4 D -; r 19 21 4 4 O -" },
            { "validation", "r 4 5 24 22 W D 1.5; p - G 2.5 9,16 14,21 23,10" },
            { "removedup", "r 4 4 15 15 W D 1.2; r 12 12 15 15 W D 1.2; l 20 20 26 26 R 2; l 26 20 20 26 R 2" },
            { "insertrow", "r 4 4 24 24 W D 1.5; r 4 12 24 8 g G 1.2" }, { "insertcol", "r 4 4 24 24 W D 1.5; r 12 4 8 24 g G 1.2" },
            { "deleterow", "r 4 4 24 24 W D 1.5; r 4 12 24 8 r R 1.2" }, { "deletecol", "r 4 4 24 24 W D 1.5; r 12 4 8 24 r R 1.2" },
            { "sheet", "r 4 4 24 24 W G 1.5; l 4 12 28 12 g 1.2; l 4 20 28 20 g 1.2; l 12 4 12 28 g 1.2; l 20 4 20 28 g 1.2" },
            { "newslide", "r 3 6 22 17 W D 1.5; r 6 9 16 4 O -; l 6 17 18 17 S 1.2; o 19 18 11 11 G -; l 24.5 20 24.5 27 W 2; l 21 23.5 28 23.5 W 2" },
            { "duplicate", "r 3 3 18 14 W D 1.5; r 11 13 18 14 W O 1.5" },
            { "deleteslide", "r 3 6 22 17 W D 1.5; o 19 18 11 11 R -; l 22 21 27 26 W 1.8; l 27 21 22 26 W 1.8" },
            { "layout", "r 3 6 26 20 W D 1.5; r 6 9 20 4 O -; r 6 15 9 8 s S 1; r 17 15 9 8 s S 1" },
            { "reset", "a 5 5 22 22 40 290 B 2.5; p B - 0 21,1 28,8 19,10" },
            { "section", "r 3 4 26 7 o O 1.2; r 3 14 26 5 W D 1.2; r 3 22 26 5 W D 1.2" },
            { "arrange", "r 4 4 14 14 b B 1.2; r 14 14 14 14 o O 1.2" },
            { "quickstyles", "q 4 4 11 11 2 B -; q 17 4 11 11 2 O -; q 4 17 11 11 2 G -; q 17 17 11 11 2 P -" },
            { "shapefill", "p W D 1.5 6,14 16,4 26,14 16,24; r 4 25 24 4 O -" },
            { "shapeoutline", "r 6 5 20 15 - D 2; r 4 25 24 4 B -" },
            { "transition", "r 3 8 13 10 o O 1.2; r 16 14 13 10 b B 1.2; p D - 0 14,3 20,7 14,11" },
            { "animation", "p Y N 1 16,3 19.5,12 29,12 21.5,18 24.5,28 16,22 7.5,28 10.5,18 3,12 12.5,12" },
            { "play", "p G - 0 9,5 9,27 26,16" },
            { "playcurrent", "r 3 6 18 14 W D 1.5; p G - 0 17,14 17,29 29,21.5" },
            { "record", "o 6 6 20 20 - D 2; o 11 11 10 10 R -" },
            { "notes", "r 4 4 24 24 W D 1.5; l 8 10 24 10 D 1.2; l 8 15 24 15 D 1.2; l 8 20 18 20 D 1.2" },
            { "sorter", "r 3 5 11 9 W D 1.2; r 18 5 11 9 W D 1.2; r 3 18 11 9 W D 1.2; r 18 18 11 9 W D 1.2" },
            { "normalview", "r 3 5 26 22 W D 1.2; r 3 5 8 22 s D 1.2" },
            { "hideslide", "r 3 6 26 18 W D 1.5; l 4 26 28 4 R 2" },
            { "home", "p - D 2 4,16 16,5 28,16; p - D 2 8,14 8,28 24,28 24,14" },
            { "open", "p y N 1.5 3,8 13,8 16,11 29,11 29,27 3,27" },
            { "info", "o 4 4 24 24 - D 2; l 16 14 16 23 D 2.5; o 14.5 8 3 3 D -" },
            { "save", "r 5 5 22 22 - D 2; r 10 5 12 8 - D 1.8; r 9 17 14 10 - D 1.8" },
            { "saveas", "r 5 5 22 22 - D 2; r 10 5 12 8 - D 1.8; p B - 0 16,28 18,22 27,13 30,16 21,25" },
            { "print", "r 8 4 16 8 W D 1.5; r 3 12 26 11 s D 1.5; r 8 19 16 9 W D 1.5; o 23 15 3 3 G -" },
            { "share", "o 3 13 7 7 - D 2; o 21 4 7 7 - D 2; o 21 21 7 7 - D 2; l 10 15 21 9 D 1.5; l 10 18 21 23 D 1.5" },
            { "export", "r 4 4 17 24 - D 2; l 13 16 28 16 B 2; p B - 0 29,16 23,11 23,21" },
            { "close", "l 8 8 24 24 D 2.2; l 24 8 8 24 D 2.2" },
            { "account", "o 10 4 12 12 - D 2; a 5 18 22 22 180 180 D 2" },
            { "options", "o 9 9 14 14 - D 2.5; l 16 3 16 8 D 3; l 16 24 16 29 D 3; l 3 16 8 16 D 3; l 24 16 29 16 D 3; l 7 7 10 10 D 3; l 22 22 25 25 D 3; l 25 7 22 10 D 3; l 7 25 10 22 D 3" },
            { "back", "o 2 2 28 28 - D 1.5; l 9 16 23 16 D 2; p - D 2 15,10 9,16 15,22" },
            { "undo", "l 9 12 19 12 D 2.2; a 13 12 12 12 270 180 D 2.2; l 19 24 11 24 D 2.2; p D - 0 3,12 10,6 10,18" },
            { "redo", "l 13 12 23 12 D 2.2; a 7 12 12 12 90 180 D 2.2; l 13 24 21 24 D 2.2; p D - 0 29,12 22,6 22,18" },
            { "focus", "o 6 6 20 20 - D 1.5; o 12 12 8 8 D -" },
            { "access", "o 13 3 6 6 D -; l 6 12 26 12 D 2; l 16 12 16 21 D 2; l 16 21 11 29 D 2; l 16 21 21 29 D 2" },
            { "pdf", "p W R 1.5 7,3 19,3 25,9 25,29 7,29; r 4 14 18 9 R -; t 13 18 7 W b PDF" },
            { "word", "r 5 3 22 26 B -; l 9 9 23 9 W 1.8; l 9 14 23 14 W 1.8; l 9 19 23 19 W 1.8; l 9 24 18 24 W 1.8" },
            { "excel", "r 4 4 24 24 G -; l 4 12 28 12 W 1.5; l 4 20 28 20 W 1.5; l 12 4 12 28 W 1.5; l 20 4 20 28 W 1.5" },
            { "powerpoint", "r 3 6 26 20 O -; r 7 10 18 9 W -; l 10 23 22 23 W 1.8" },
            { "folder", "p y N 1.5 3,8 13,8 16,11 29,11 29,27 3,27" },
            { "pc", "r 3 5 26 17 b D 1.5; l 11 27 21 27 D 2; l 16 22 16 27 D 1.5" },
            { "addplace", "o 4 4 24 24 - D 1.8; l 16 10 16 22 D 2; l 10 16 22 16 D 2" },
            { "pin", "p - D 1.8 12,4 20,4 19,13 23,17 9,17 13,13; l 16 17 16 28 D 1.8" },
            { "lanshare", "o 11 20 10 10 B -; a 5 14 22 22 225 90 B 2.2; a 0 9 32 32 225 90 B 2.2" },
            { "email", "r 3 7 26 18 W D 1.5; p - D 1.5 3,7 16,18 29,7" },
            { "language", "t 10 12 13 K b A; t 21 22 12 B b 文" },
            { "hyphen", "t 16 16 15 K n a-b" },
            { "breaks", "r 7 2 18 11 W D 1.5; r 7 19 18 11 W D 1.5; l 3 16 29 16 B 1.5" },
        };
        static readonly Dictionary<string, Bitmap> cache = new Dictionary<string, Bitmap>();

        public static bool Has(string name) { if (name == null) return false; if (name.StartsWith("n:")) name = name.Substring(2); else if (name.StartsWith("l:")) return UnifiedUi.HasLucide(name.Substring(2)); return spec.ContainsKey(name); }
        public static void Draw(Graphics g, string name, Rectangle r, bool enabled)
        {
            if (name != null && name.StartsWith("l:")) { UnifiedUi.DrawLucide(g, name.Substring(2), r, enabled); return; }   // v0.1.7: shared icon set with the Web UI
            if (name != null && name.StartsWith("n:")) name = name.Substring(2);
            if (!Has(name) || r.Width <= 0) return;
            string key = name + "|" + r.Width + "|" + enabled; Bitmap b;
            if (!cache.TryGetValue(key, out b))
            {
                b = new Bitmap(r.Width, r.Height);
                using (Graphics bg = Graphics.FromImage(b)) Render(bg, spec[name], r.Width, enabled);
                cache[key] = b;
            }
            g.DrawImageUnscaled(b, r.X, r.Y);
        }

        static Color Col(string t, bool en)
        {
            if (t == "-") return Color.Empty; Color c;
            if (t.StartsWith("#")) c = ColorTranslator.FromHtml(t); else if (!pal.TryGetValue(t[0], out c)) c = Color.Gray;
            if (!en) { int v = (int)(c.R * 0.3 + c.G * 0.59 + c.B * 0.11); v = 170 + v / 4; c = Color.FromArgb(c.A, Math.Min(255, v), Math.Min(255, v), Math.Min(255, v)); }
            return c;
        }
        static float N(string s) { return float.Parse(s, CultureInfo.InvariantCulture); }
        static void Render(Graphics g, string script, int size, bool en)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias; g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit; g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            float s = size / 32f; g.ScaleTransform(s, s);
            foreach (string raw in script.Split(';'))
            {
                string op = raw.Trim(); if (op.Length == 0) continue;
                string[] a = op.Split(' ');
                try
                {
                    switch (a[0])
                    {
                        case "r": case "o": case "q":
                            {
                                int o = a[0] == "q" ? 1 : 0;
                                RectangleF rc = new RectangleF(N(a[1]), N(a[2]), N(a[3]), N(a[4]));
                                Color fill = Col(a[5 + o], en), stroke = Col(a[6 + o], en); float w = a.Length > 7 + o ? N(a[7 + o]) : 1.5f;
                                using (GraphicsPath p = new GraphicsPath())
                                {
                                    if (a[0] == "o") p.AddEllipse(rc);
                                    else if (a[0] == "q") { float d = N(a[5]) * 2; p.AddArc(rc.X, rc.Y, d, d, 180, 90); p.AddArc(rc.Right - d, rc.Y, d, d, 270, 90); p.AddArc(rc.Right - d, rc.Bottom - d, d, d, 0, 90); p.AddArc(rc.X, rc.Bottom - d, d, d, 90, 90); p.CloseFigure(); }
                                    else p.AddRectangle(rc);
                                    if (!fill.IsEmpty) using (SolidBrush br = new SolidBrush(fill)) g.FillPath(br, p);
                                    if (!stroke.IsEmpty) using (Pen pn = new Pen(stroke, Math.Max(w, 1f / s))) g.DrawPath(pn, p);
                                }
                                break;
                            }
                        case "l":
                            using (Pen pn = new Pen(Col(a[5], en), Math.Max(N(a[6]), 1f / s)) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                                g.DrawLine(pn, N(a[1]), N(a[2]), N(a[3]), N(a[4]));
                            break;
                        case "a":
                            using (Pen pn = new Pen(Col(a[7], en), Math.Max(N(a[8]), 1f / s)) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                                g.DrawArc(pn, N(a[1]), N(a[2]), N(a[3]), N(a[4]), N(a[5]), N(a[6]));
                            break;
                        case "p":
                            {
                                Color fill = Col(a[1], en), stroke = Col(a[2], en); float w = N(a[3]);
                                List<PointF> pts = new List<PointF>();
                                for (int i = 4; i < a.Length; i++) { string[] xy = a[i].Split(','); if (xy.Length == 2) pts.Add(new PointF(N(xy[0]), N(xy[1]))); }
                                if (pts.Count < 2) break;
                                if (!fill.IsEmpty) using (SolidBrush br = new SolidBrush(fill)) g.FillPolygon(br, pts.ToArray());
                                if (!stroke.IsEmpty && w > 0) using (Pen pn = new Pen(stroke, Math.Max(w, 1f / s)) { LineJoin = LineJoin.Round, StartCap = LineCap.Round, EndCap = LineCap.Round }) { if (!fill.IsEmpty) g.DrawPolygon(pn, pts.ToArray()); else g.DrawLines(pn, pts.ToArray()); }
                                break;
                            }
                        case "t":
                            {
                                float cx = N(a[1]), cy = N(a[2]), sz = N(a[3]); Color c = Col(a[4], en); string st = a[5];
                                string txt = string.Join(" ", a, 6, a.Length - 6);
                                FontStyle fs = FontStyle.Regular; if (st.Contains("b")) fs |= FontStyle.Bold; if (st.Contains("i")) fs |= FontStyle.Italic;
                                using (Font f = new Font(txt == "Ω" || txt == "π" || txt == "Σ" || txt == "¶" ? "Cambria" : "Segoe UI", sz, fs, GraphicsUnit.Pixel))
                                using (SolidBrush br = new SolidBrush(c))
                                using (StringFormat sf = new StringFormat(StringFormat.GenericTypographic) { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                                    g.DrawString(txt, f, br, new RectangleF(cx - 20, cy - 16, 40, 32), sf);
                                break;
                            }
                    }
                }
                catch { }
            }
        }
    }
}
