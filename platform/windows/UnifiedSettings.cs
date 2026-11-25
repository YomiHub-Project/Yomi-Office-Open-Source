// Office Open Source v0.1.7 - Backstage navigation/tiles/file types and the Tùy chọn (Options) dialog shared with the Web
// editor. Everything comes from ui_settings.json (embedded resource; the Web bundle imports the same file), and the
// values live in <data>\settings.json - the Web editor's settings file - so an option changed in one editor is the
// same option in the other. Keys the native code already reads through NativeSettings are mirrored both ways ("nat").
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace OfficeOpenSource
{
    static class UiSpec
    {
        static Dictionary<string, object> root; static bool loaded;
        public static Dictionary<string, object> Root
        {
            get
            {
                if (loaded) return root; loaded = true;
                try
                {
                    using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("ui_settings.json"))
                    using (StreamReader r = new StreamReader(s, Encoding.UTF8))
                        root = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Deserialize<Dictionary<string, object>>(r.ReadToEnd());
                }
                catch (Exception ex) { Lok.Trace("ui_settings: " + ex.Message); root = null; }
                return root;
            }
        }
        public static IEnumerable<Dictionary<string, object>> List(object o) { IEnumerable e = o as IEnumerable; if (e == null) yield break; foreach (object x in e) yield return x as Dictionary<string, object>; }
        public static string S(Dictionary<string, object> d, string k) { object o; return d != null && d.TryGetValue(k, out o) && o != null ? Convert.ToString(o) : null; }
        public static string KindName(int docType) { return docType == LokDoc.SPREADSHEET ? "excel" : docType == LokDoc.PRESENTATION ? "ppt" : "word"; }
        public static IEnumerable<Dictionary<string, object>> Options() { return Root == null ? new Dictionary<string, object>[0] : List(Root["options"]); }
        public static IEnumerable<Dictionary<string, object>> AllItems()
        {
            foreach (Dictionary<string, object> c in Options()) foreach (Dictionary<string, object> s in List(c["secs"])) foreach (Dictionary<string, object> i in List(s["items"])) yield return i;
        }
        /// <summary>Backstage navigation: {key, text, icon} or null for a separator.</summary>
        public static List<string[]> Nav()
        {
            List<string[]> r = new List<string[]>();
            if (Root == null) return r;
            foreach (object o in (IEnumerable)Root["nav"]) { if (o == null) { r.Add(null); continue; } List<string> a = new List<string>(); foreach (object x in (IEnumerable)o) a.Add(Convert.ToString(x)); r.Add(a.ToArray()); }
            return r;
        }
        /// <summary>Backstage tiles of a page: {id, icon, title, sub}.</summary>
        public static List<string[]> Tiles(string page)
        {
            List<string[]> r = new List<string[]>(); object o;
            if (Root == null || !((Dictionary<string, object>)Root["tiles"]).TryGetValue(page, out o)) return r;
            foreach (object t in (IEnumerable)o) { List<string> a = new List<string>(); foreach (object x in (IEnumerable)t) a.Add(x == null ? null : Convert.ToString(x)); r.Add(a.ToArray()); }
            return r;
        }
        /// <summary>File types {ext, label, where}: where b = both editors, n = native only, w = Web editor only.</summary>
        public static List<string[]> Formats(int docType)
        {
            List<string[]> r = new List<string[]>(); object o;
            if (Root == null || !((Dictionary<string, object>)Root["formats"]).TryGetValue(KindName(docType), out o)) return r;
            foreach (object t in (IEnumerable)o) { List<string> a = new List<string>(); foreach (object x in (IEnumerable)t) a.Add(Convert.ToString(x)); r.Add(a.ToArray()); }
            return r;
        }
        public static string FileExt(string ext) { return ext == "xml2003" ? "xml" : ext == "video" ? "webm" : ext; }
    }

    /// <summary>The Web editor's settings.json (same folder as native-settings.ini).</summary>
    static class SharedSettings
    {
        static string FilePath { get { return Path.Combine(App.DataDir ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOpenSource"), "settings.json"); } }
        public static Dictionary<string, object> Load()
        {
            try { if (File.Exists(FilePath)) return App.Json.Deserialize<Dictionary<string, object>>(File.ReadAllText(FilePath, Encoding.UTF8)) ?? new Dictionary<string, object>(); } catch { }
            return new Dictionary<string, object>();
        }
        public static object Get(Dictionary<string, object> s, string key)
        {
            object v = s; foreach (string p in key.Split('.')) { Dictionary<string, object> d = v as Dictionary<string, object>; object n; if (d == null || !d.TryGetValue(p, out n)) return null; v = n; }
            return v;
        }
        public static void Put(Dictionary<string, object> s, string key, object val)
        {
            string[] ps = key.Split('.'); Dictionary<string, object> d = s;
            for (int i = 0; i < ps.Length - 1; i++) { object n; if (!d.TryGetValue(ps[i], out n) || !(n is Dictionary<string, object>)) { n = new Dictionary<string, object>(); d[ps[i]] = n; } d = (Dictionary<string, object>)n; }
            d[ps[ps.Length - 1]] = val;
        }
        public static void Save(Dictionary<string, object> s)
        {
            try { Directory.CreateDirectory(Path.GetDirectoryName(FilePath)); File.WriteAllText(FilePath, App.Json.Serialize(s), new UTF8Encoding(false)); } catch (Exception ex) { Lok.Trace("settings.json: " + ex.Message); }
        }
        public static void Set(Dictionary<string, object> patch)
        {
            Dictionary<string, object> s = Load();
            foreach (KeyValuePair<string, object> kv in patch) Put(s, kv.Key, kv.Value);
            if (patch.ContainsKey("autosaveMin")) s["autosaveSec"] = Math.Max(1, Convert.ToInt32(patch["autosaveMin"])) * 60;
            if (patch.ContainsKey("dex.enabled")) s["dexForceDesktop"] = patch["dex.enabled"];
            Save(s);
        }
        /// <summary>Current value (settings.json, then native alias, then the spec default) typed for the item kind.</summary>
        public static object Value(Dictionary<string, object> s, Dictionary<string, object> it)
        {
            string key = UiSpec.S(it, "key"), nat = UiSpec.S(it, "nat"), k = UiSpec.S(it, "k");
            object v = Get(s, key);
            if (key == "macrosAllowed") return Lok.MacrosAllowed;
            if (v == null && key == "autosaveMin") { object sec = Get(s, "autosaveSec"); if (sec != null) v = Math.Max(1, Convert.ToInt32(sec) / 60); }
            if ((v == null || (key == "userName" && Convert.ToString(v) == "Người dùng") || (key == "initials" && Convert.ToString(v) == "ND")) && nat != null && NativeSettings.Has(nat)) v = NativeSettings.Get(nat, "");
            if (v == null) { object d; it.TryGetValue("d", out d); v = d; }
            if (key == "userName" && (string.IsNullOrEmpty(Convert.ToString(v)) || Convert.ToString(v) == "Người dùng")) v = Ui.UserName();
            if (key == "initials" && (string.IsNullOrEmpty(Convert.ToString(v)) || Convert.ToString(v) == "ND")) v = Ui.Initials(Ui.UserName());
            if (key == "saveDir" && string.IsNullOrEmpty(Convert.ToString(v))) v = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (k == "chk") { string sv = Convert.ToString(v).ToLowerInvariant(); return sv == "true" || sv == "1"; }
            if (k == "num") { int n; return int.TryParse(Convert.ToString(v), out n) ? n : 0; }
            return Convert.ToString(v);
        }
        /// <summary>Called by NativeSettings.Load: values changed in the Web editor win over native-settings.ini.</summary>
        public static void Overlay(Dictionary<string, string> map)
        {
            try
            {
                if (UiSpec.Root == null || !File.Exists(FilePath)) return;
                Dictionary<string, object> s = Load();
                // the Web editor's default name "Người dùng" -> the Windows user, so both editors show the same author and avatar
                object un = Get(s, "userName");
                if (un == null || Convert.ToString(un) == "Người dùng") { string nm = map.ContainsKey("user") && map["user"].Length > 0 ? map["user"] : Ui.UserName(); s["userName"] = nm; s["initials"] = Ui.Initials(nm); Save(s); }
                foreach (Dictionary<string, object> it in UiSpec.AllItems())
                {
                    string key = UiSpec.S(it, "key"), nat = UiSpec.S(it, "nat"); if (key == null || nat == null || nat.StartsWith("__")) continue;
                    object v = Get(s, key); if (v == null) continue;
                    if ((key == "userName" && Convert.ToString(v) == "Người dùng") || (key == "initials" && Convert.ToString(v) == "ND")) continue;
                    if (key == "autosaveMin" && Get(s, "autosaveMin") == null) continue;
                    map[nat] = v is bool ? ((bool)v ? "1" : "0") : Convert.ToString(v);
                }
            }
            catch { }
        }
        /// <summary>Called by NativeSettings.Set: native writes of an aliased key go to settings.json too.</summary>
        public static void Mirror(string nat, string val)
        {
            try
            {
                if (UiSpec.Root == null) return;
                foreach (Dictionary<string, object> it in UiSpec.AllItems())
                {
                    if (UiSpec.S(it, "nat") != nat) continue;
                    string k = UiSpec.S(it, "k"); object v = val;
                    if (k == "chk") v = val == "1" || val == "true"; else if (k == "num") { int n; v = int.TryParse(val, out n) ? n : 0; }
                    Set(new Dictionary<string, object> { { UiSpec.S(it, "key"), v } }); return;
                }
            }
            catch { }
        }
    }

    /// <summary>Tùy chọn - rendered from ui_settings.json exactly like the Web editor's dialog.</summary>
    class OptionsDialog : OfficeDialog
    {
        readonly ListBox cats = new ListBox { Location = new Point(10, 10), Size = new Size(170, 420), BorderStyle = BorderStyle.None, BackColor = Ui.Chrome, IntegralHeight = false, ItemHeight = 26, DrawMode = DrawMode.OwnerDrawFixed };
        readonly Panel page = new Panel { Location = new Point(190, 10), Size = new Size(540, 420), BackColor = Color.White, AutoScroll = true };
        readonly Dictionary<string, object> pend = new Dictionary<string, object>();
        readonly List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
        readonly NativeDocForm owner; readonly int docType; readonly string kind; readonly Dictionary<string, object> cur;
        public OptionsDialog(NativeDocForm o) : base("Tùy chọn", 740, 480)
        {
            owner = o; docType = ((IBackstageHost)o).DocType; kind = UiSpec.KindName(docType); cur = SharedSettings.Load();
            foreach (Dictionary<string, object> c in UiSpec.Options()) { list.Add(c); string tx = UiSpec.S(c, "tx"); cats.Items.Add(kind == "excel" && tx != null ? tx : UiSpec.S(c, "t")); }
            cats.DrawItem += (s, e) => { if (e.Index < 0) return; bool sel = (e.State & DrawItemState.Selected) != 0; e.Graphics.FillRectangle(new SolidBrush(sel ? Color.White : Ui.Chrome), e.Bounds); if (sel) e.Graphics.FillRectangle(new SolidBrush(Ui.AccentFor(docType)), e.Bounds.X, e.Bounds.Y + 4, 3, e.Bounds.Height - 8); TextRenderer.DrawText(e.Graphics, (string)cats.Items[e.Index], Font, new Rectangle(e.Bounds.X + 12, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), Ui.Text, TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix); };
            cats.SelectedIndexChanged += (s, e) => ShowCat(cats.SelectedIndex);
            Controls.Add(cats); Controls.Add(page);
            Button ok = OkCancel(442); ok.Click += (s, e) => SaveAll();
            if (cats.Items.Count > 0) cats.SelectedIndex = 0;
            string shot = Environment.GetEnvironmentVariable("OOS_TEST_SHOT");   // v0.1.7 UI tests: save every category as an image
            if (!string.IsNullOrEmpty(shot)) Shown += (s, e) =>
            {
                Timer t = new Timer { Interval = 400 }; int ci = 0;
                t.Tick += (s2, e2) => { try { using (Bitmap bm = new Bitmap(Width, Height)) { DrawToBitmap(bm, new Rectangle(0, 0, Width, Height)); bm.Save(shot + "_" + ci + ".png"); } } catch { } ci++; if (ci >= cats.Items.Count) { t.Stop(); Close(); } else cats.SelectedIndex = ci; };
                t.Start();
            };
        }
        int y;
        bool Vis(Dictionary<string, object> d) { string only = UiSpec.S(d, "only"); return only == null || only == kind; }
        object Val(Dictionary<string, object> it) { string key = UiSpec.S(it, "key"); return key != null && pend.ContainsKey(key) ? pend[key] : SharedSettings.Value(cur, it); }
        void ShowCat(int i)
        {
            page.SuspendLayout(); foreach (Control c in page.Controls) c.Dispose(); page.Controls.Clear(); page.AutoScrollPosition = Point.Empty; y = 10;
            foreach (Dictionary<string, object> s in UiSpec.List(list[i]["secs"]))
            {
                if (!Vis(s)) continue;
                page.Controls.Add(new Label { UseMnemonic = false, Text = UiSpec.S(s, "h"), Location = new Point(10, y), AutoSize = true, Font = Ui.F(10.5f, FontStyle.Bold) }); y += 26;
                string sub = UiSpec.S(s, "sub"); if (sub != null) { page.Controls.Add(new Label { UseMnemonic = false, Text = sub, Location = new Point(10, y), Size = new Size(510, 34), ForeColor = Ui.Sub }); y += 36; }
                foreach (Dictionary<string, object> it in UiSpec.List(s["items"])) if (Vis(it)) Item(it);
                y += 4;
            }
            page.ResumeLayout();
        }
        void Item(Dictionary<string, object> it)
        {
            string k = UiSpec.S(it, "k"), key = UiSpec.S(it, "key"), t = UiSpec.S(it, "t");
            switch (k)
            {
                case "chk":
                    { CheckBox c = new CheckBox { UseMnemonic = false, Text = t, Location = new Point(16, y), AutoSize = true, Checked = (bool)Val(it) }; c.CheckedChanged += (s, e) => pend[key] = c.Checked; page.Controls.Add(c); y += 26; break; }
                case "num":
                    {
                        page.Controls.Add(new Label { UseMnemonic = false, Text = t, Location = new Point(16, y + 3), Size = new Size(260, 20) });
                        int mn = Convert.ToInt32(it["min"]), mx = Convert.ToInt32(it["max"]);
                        NumericUpDown n = new NumericUpDown { Location = new Point(280, y), Width = 70, Minimum = mn, Maximum = mx, Value = Math.Max(mn, Math.Min(mx, Convert.ToInt32(Val(it)))) };
                        n.ValueChanged += (s, e) => pend[key] = (int)n.Value; page.Controls.Add(n); y += 30; break;
                    }
                case "sel":
                    {
                        page.Controls.Add(new Label { UseMnemonic = false, Text = t, Location = new Point(16, y + 3), Size = new Size(200, 20) });
                        ComboBox cb = new ComboBox { Location = new Point(220, y), Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
                        List<string> vals = new List<string>(); string v = Convert.ToString(Val(it));
                        foreach (object o in (IEnumerable)it["opts"]) { List<string> a = new List<string>(); foreach (object x in (IEnumerable)o) a.Add(Convert.ToString(x)); vals.Add(a[0]); cb.Items.Add(a[1]); }
                        cb.SelectedIndex = Math.Max(0, vals.IndexOf(v)); cb.SelectedIndexChanged += (s, e) => pend[key] = vals[cb.SelectedIndex];
                        page.Controls.Add(cb); y += 30; break;
                    }
                case "text":
                    {
                        page.Controls.Add(new Label { UseMnemonic = false, Text = t, Location = new Point(16, y + 3), Size = new Size(200, 20) });
                        TextBox tb = new TextBox { Location = new Point(220, y), Width = 250, Text = Convert.ToString(Val(it)) }; tb.TextChanged += (s, e) => pend[key] = tb.Text; page.Controls.Add(tb); y += 30; break;
                    }
                case "note":
                    { Label l = new Label { UseMnemonic = false, Text = t, Location = new Point(16, y), Size = new Size(500, 20), ForeColor = Ui.Text }; page.Controls.Add(l); y += 26; break; }
                case "btn":
                    {
                        string act = UiSpec.S(it, "act"), ic = UiSpec.S(it, "i");
                        Button b = new Button { UseMnemonic = false, Text = "      " + t, Location = new Point(16, y), AutoSize = true, Height = 26, TextAlign = ContentAlignment.MiddleLeft };
                        b.Paint += (s, e) => Icons.Draw(e.Graphics, ic, new Rectangle(8, (b.Height - 16) / 2, 16, 16), true);
                        b.Click += (s, e) => owner.OptionAction(act, this); page.Controls.Add(b); y += 36; break;
                    }
                case "tabs":
                    {
                        string hid = Convert.ToString(Val(it)) ?? ""; Dictionary<string, object> lay = UnifiedUi.Layout(kind); int ti = 0;
                        List<CheckBox> boxes = new List<CheckBox>();
                        if (lay != null) foreach (Dictionary<string, object> tab in UiSpec.List(lay["tabs"]))
                            {
                                string name = UiSpec.S(tab, "name");
                                CheckBox c = new CheckBox { UseMnemonic = false, Text = name, Location = new Point(16, y), AutoSize = true, Checked = !hid.Contains("|" + name + "|"), Enabled = ti++ != 0 }; page.Controls.Add(c); boxes.Add(c); y += 24;
                                c.CheckedChanged += (s, e) => { string h = Convert.ToString(Val(it)) ?? ""; h = h.Replace("|" + c.Text + "|", ""); if (!c.Checked) h += "|" + c.Text + "|"; pend[key] = h; };
                            }
                        break;
                    }
            }
        }
        void SaveAll()
        {
            if (pend.Count == 0) return;
            foreach (Dictionary<string, object> it in UiSpec.AllItems())
            {
                string key = UiSpec.S(it, "key"), nat = UiSpec.S(it, "nat"); if (key == null || nat == null || !pend.ContainsKey(key)) continue;
                object v = pend[key];
                if (nat == "__macros") { Lok.MacrosAllowed = (bool)v; continue; }
                NativeSettings.SetLocal(nat, v is bool ? ((bool)v ? "1" : "0") : Convert.ToString(v));
            }
            SharedSettings.Set(pend);
        }
    }

    partial class NativeDocForm
    {
        /// <summary>Tùy chọn buttons (Bổ trợ / Cập nhật / DeX): native when the native editor has it, otherwise the Web window.</summary>
        internal void OptionAction(string act, Form dlg)
        {
            switch (act)
            {
                case "ai": if (dlg != null) dlg.Close(); AskAi(null); return;
                case "updateNow": NativeUpdate.CheckAsync(dlg ?? (Form)this, true); return;
            }
            WebAct(act);
        }
        /// <summary>Runs a Web-only action that needs no document (cloud accounts, fonts, DeX…) in the main Web window.</summary>
        internal static void WebAct(string act)
        {
            lock (App.Pending) App.Pending.Add(new Dictionary<string, object> { { "act", act } });
            App.ActivateExistingWindow();
        }
        /// <summary>Hands the document to the Web editor and runs Backstage action "id" there.</summary>
        internal void WebHandoff(string id) { CloseBackstage(); OpenInWebCmd("bs:" + id); }
        /// <summary>bs:* commands sent by the Web editor (file types only the native engine writes).</summary>
        void RunBsCommand(string id)
        {
            if (id.StartsWith("page:")) { OpenBackstage(id.Substring(5)); return; }
            if (id == "options") { ShowOptions(); return; }
            if (id.StartsWith("fmt:"))
            {
                string ext = id.Substring(4); string label = ext;
                foreach (string[] f in UiSpec.Formats(type)) if (f[0] == ext) label = f[1];
                using (SaveFileDialog d = new SaveFileDialog { Filter = label + " (*." + ext + ")|*." + ext, FileName = Path.GetFileNameWithoutExtension(path) + "." + ext, InitialDirectory = Path.GetDirectoryName(path) })
                    if (d.ShowDialog(this) == DialogResult.OK) SaveAsFile(d.FileName, ext);
            }
        }
    }
}
