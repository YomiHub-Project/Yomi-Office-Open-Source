// Office Open Source v0.1.6 - Windows launcher (+ native LibreOfficeKit document engine, see LokView.cs)
// (was v0.1.1) (Native Hybrid Runtime, Windows 7+, x86/x64/ARM64 via AnyCPU, .NET Framework 4.6.2+)
// Hosts embedded WebView2 inside a native Form window (OfficeMainWindow), with zero-latency local server,
// native DPI awareness, single instance activation, and native drag-and-drop file opening.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Win32;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

[assembly: AssemblyTitle("Yomi-Office Open Source")]
[assembly: AssemblyProduct("Yomi-Office Open Source")]
[assembly: AssemblyVersion("0.1.8.0")]
[assembly: AssemblyFileVersion("0.1.8.0")]
[assembly: AssemblyCompany("Yomi-Office Open Source Project")]
[assembly: AssemblyCopyright("Open source")]

namespace OfficeOpenSource
{
    public class OfficeMainWindow : Form
    {
        public int Port { get; private set; }
        private WebView2 webView;
        private string initialHash;
        private bool isFullScreen = false;
        private FormWindowState prevWindowState;
        private FormBorderStyle prevBorderStyle;
        private Rectangle prevBounds;

        public OfficeMainWindow(int port, string hash)
        {
            this.Port = port;
            this.initialHash = hash;
            this.Text = "Yomi-Office Open Source v" + App.Version;
            this.Size = new Size(1440, 920);   // v0.1.7: same default size as the native editor window
            this.MinimumSize = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AllowDrop = true;
            // v0.1.7: borderless like the native editor - the Web title bar draws the caption buttons (same as TitleBar.cs)
            this.FormBorderStyle = FormBorderStyle.None; this.Padding = new Padding(1); this.BackColor = Color.FromArgb(0xF3, 0xF3, 0xF3); this.DoubleBuffered = true;
            this.Resize += (s, e) => { try { if (webView != null && webView.CoreWebView2 != null) webView.CoreWebView2.PostWebMessageAsString(this.WindowState == FormWindowState.Maximized ? "state:max" : "state:normal"); } catch { } Invalidate(); };

            try
            {
                string iconPath = Path.Combine(App.Dir, "app.ico");
                if (!File.Exists(iconPath)) iconPath = Path.Combine(App.Dir, "app", "favicon.ico");
                if (File.Exists(iconPath)) this.Icon = new Icon(iconPath);
            }
            catch { }

            this.DragEnter += (s, e) => {
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
            };
            this.DragDrop += (s, e) => {
                try
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    if (files != null && files.Length > 0)
                    {
                        foreach (string f in files) App.AddPending(f);
                        if (webView != null && webView.CoreWebView2 != null)
                        {
                            webView.CoreWebView2.ExecuteScriptAsync("window.OOS_Host && OOS_Host.pending().then(function(f){if(f&&f.length){f.forEach(function(x){window.OOS.openFile(x)})}})");
                        }
                    }
                }
                catch { }
            };

            this.KeyPreview = true;
            this.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.F11)
                {
                    ToggleFullScreen();
                }
            };

            this.FormClosing += (s, e) => {
                App.OnWindowClosing(this);
            };

            InitWebView();
        }

        private async void InitWebView()
        {
            try
            {
                webView = new WebView2();
                webView.Dock = DockStyle.Fill;
                this.Controls.Add(webView);

                string prof = Path.Combine(App.DataDir, "webview2_profile");
                Directory.CreateDirectory(prof);
                var opts = new CoreWebView2EnvironmentOptions(
                    "--allow-file-access-from-files --disable-features=Translate,TranslateUI,msImplicitSignin,msEdgeSyncConsent,msUndersideButton --autoplay-policy=no-user-gesture-required --lang=vi"
                );
                var env = await CoreWebView2Environment.CreateAsync(null, prof, opts);
                await webView.EnsureCoreWebView2Async(env);

                webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
                webView.CoreWebView2.WebMessageReceived += (s, e) =>
                {
                    string m = null; try { m = e.TryGetWebMessageAsString(); } catch { }
                    if (m == "drag") { ReleaseCapture(); SendMessage(Handle, 0xA1, (IntPtr)2, IntPtr.Zero); }
                    else if (m == "min") WindowState = FormWindowState.Minimized;
                    else if (m == "max") WindowState = WindowState == FormWindowState.Maximized ? FormWindowState.Normal : FormWindowState.Maximized;
                    else if (m == "close") Close();
                    else if (m == "state") { try { webView.CoreWebView2.PostWebMessageAsString(WindowState == FormWindowState.Maximized ? "state:max" : "state:normal"); } catch { } }
                };
                webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
                webView.CoreWebView2.Settings.IsZoomControlEnabled = false;
                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;

                webView.CoreWebView2.DocumentTitleChanged += (s, e) => {
                    try
                    {
                        if (!string.IsNullOrEmpty(webView.CoreWebView2.DocumentTitle))
                        {
                            this.Text = webView.CoreWebView2.DocumentTitle + " - Yomi-Office Open Source";
                        }
                    }
                    catch { }
                };

                webView.CoreWebView2.NewWindowRequested += (s, e) => {
                    e.Handled = true;
                    if (!string.IsNullOrEmpty(e.Uri))
                    {
                        if (e.Uri.StartsWith("http://127.0.0.1:" + Port) || e.Uri.StartsWith("http://localhost:" + Port))
                        {
                            Uri u = new Uri(e.Uri);
                            App.OpenNativeWindow(Port, u.Fragment);
                        }
                        else
                        {
                            try { Process.Start(e.Uri); } catch { }
                        }
                    }
                };

                string url = "http://127.0.0.1:" + Port + "/index.html?host=win" + 
                    (string.IsNullOrEmpty(initialHash) ? "" : (initialHash.StartsWith("#") ? initialHash : "#" + initialHash));
                webView.CoreWebView2.Navigate(url);
            }
            catch (Exception ex)
            {
                InitFallback(ex);
            }
        }

        private void InitFallback(Exception ex)
        {
            string browser = App.FindBrowser();
            if (browser != null)
            {
                this.Hide();
                App.OpenFallbackBrowser(Port, initialHash);
                return;
            }

            Panel panel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(243, 244, 246) };
            Label lbl = new Label {
                Text = "Yomi-Office Open Source Native Hybrid yêu cầu Microsoft Edge WebView2 Runtime.\n\nVui lòng cài đặt WebView2 Runtime để chạy ứng dụng độc lập:",
                Dock = DockStyle.Top, Height = 100, Font = new Font("Segoe UI", 11, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter
            };
            Button btn = new Button {
                Text = "Tải WebView2 Runtime (Microsoft)",
                Size = new Size(260, 45),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(24, 90, 189),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btn.Location = new Point((this.ClientSize.Width - btn.Width) / 2, 120);
            btn.Click += (s, e) => {
                try { Process.Start("https://go.microsoft.com/fwlink/p/?LinkId=2124703"); } catch { }
            };
            panel.Controls.Add(btn);
            panel.Controls.Add(lbl);
            this.Controls.Add(panel);
        }

        public void ActivateWindow()
        {
            if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;
            this.Activate();
            this.BringToFront();
            if (webView != null && webView.CoreWebView2 != null)
            {
                webView.CoreWebView2.ExecuteScriptAsync("window.OOS_Host && OOS_Host.pending().then(function(f){if(f&&f.length){f.forEach(function(x){window.OOS.openFile(x)})}})");
            }
        }

        [DllImport("user32.dll")] static extern bool ReleaseCapture();
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr h, int i);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr h, int i, int v);
        [StructLayout(LayoutKind.Sequential)] struct NCRECT { public int L, T, R, B; }
        [StructLayout(LayoutKind.Sequential)] struct NCCALC { public NCRECT r0, r1, r2; public IntPtr pos; }
        [StructLayout(LayoutKind.Sequential)] struct MINMAX { public Point res, maxSize, maxPos, minTrack, maxTrack; }
        protected override CreateParams CreateParams
        {
            get { CreateParams cp = base.CreateParams; cp.Style |= 0x00020000 | 0x00010000 | 0x00080000 | 0x00040000; cp.ClassStyle |= 0x00020000; return cp; }
        }
        protected override void WndProc(ref Message m)
        {
            // same frame handling as NativeDocForm: whole window is client area, Aero snap / resize kept
            if (m.Msg == 0x86) { m.LParam = new IntPtr(-1); base.WndProc(ref m); return; }
            if ((m.Msg == 0x0C || m.Msg == 0x80) && IsHandleCreated) { int st = GetWindowLong(Handle, -16); SetWindowLong(Handle, -16, st & ~0x10000000); base.WndProc(ref m); SetWindowLong(Handle, -16, st); return; }
            if (m.Msg == 0x83 && m.WParam != IntPtr.Zero)
            {
                if (WindowState == FormWindowState.Maximized && !isFullScreen) { NCCALC p = (NCCALC)Marshal.PtrToStructure(m.LParam, typeof(NCCALC)); Rectangle wa = Screen.FromHandle(Handle).WorkingArea; p.r0.L = wa.Left; p.r0.T = wa.Top; p.r0.R = wa.Right; p.r0.B = wa.Bottom; Marshal.StructureToPtr(p, m.LParam, false); }
                m.Result = IntPtr.Zero; return;
            }
            if (m.Msg == 0x24)
            {
                MINMAX mi = (MINMAX)Marshal.PtrToStructure(m.LParam, typeof(MINMAX)); Screen sc = Screen.FromHandle(Handle); Rectangle wa = sc.WorkingArea, bd = sc.Bounds;
                mi.maxPos = new Point(wa.Left - bd.Left, wa.Top - bd.Top); mi.maxSize = new Point(wa.Width, wa.Height); mi.minTrack = new Point(MinimumSize.Width, MinimumSize.Height);
                Marshal.StructureToPtr(mi, m.LParam, false); m.Result = IntPtr.Zero; return;
            }
            if (m.Msg == 0x84)
            {
                Point p = PointToClient(new Point((short)((long)m.LParam & 0xFFFF), (short)(((long)m.LParam >> 16) & 0xFFFF))); int e = 6;
                if (WindowState != FormWindowState.Maximized)
                {
                    bool l = p.X < e, r = p.X >= Width - e, t = p.Y < e, bt = p.Y >= Height - e;
                    if (t && l) { m.Result = (IntPtr)13; return; } if (t && r) { m.Result = (IntPtr)14; return; }
                    if (bt && l) { m.Result = (IntPtr)16; return; } if (bt && r) { m.Result = (IntPtr)17; return; }
                    if (l) { m.Result = (IntPtr)10; return; } if (r) { m.Result = (IntPtr)11; return; } if (t) { m.Result = (IntPtr)12; return; } if (bt) { m.Result = (IntPtr)15; return; }
                }
                m.Result = (IntPtr)1; return;
            }
            base.WndProc(ref m);
        }
        protected override void OnPaint(PaintEventArgs e) { base.OnPaint(e); if (WindowState != FormWindowState.Maximized) using (Pen p = new Pen(Color.FromArgb(0xB4, 0xB4, 0xB4))) e.Graphics.DrawRectangle(p, 0, 0, Width - 1, Height - 1); }

        private void ToggleFullScreen()
        {
            if (!isFullScreen)
            {
                prevWindowState = this.WindowState;
                prevBorderStyle = this.FormBorderStyle;
                prevBounds = this.Bounds;

                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Normal;
                this.Bounds = Screen.FromControl(this).Bounds;
                isFullScreen = true;
            }
            else
            {
                this.FormBorderStyle = prevBorderStyle;
                this.WindowState = prevWindowState;
                this.Bounds = prevBounds;
                isFullScreen = false;
            }
        }
    }

    static class App
    {
        public const string Version = "0.1.8";
        public static string Dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        public static string WebRoot;
        public static string DataDir;
        public static int Port;
        public static readonly List<Dictionary<string, object>> Pending = new List<Dictionary<string, object>>();
        public static readonly Dictionary<string, Dictionary<string, object>> Stash = new Dictionary<string, Dictionary<string, object>>();
        public static DateTime LastPing = DateTime.UtcNow;
        public static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue, RecursionLimit = 64 };
        static Mutex mutex;

        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();

        public static readonly List<OfficeMainWindow> OpenWindows = new List<OfficeMainWindow>();

        public static void OpenNativeWindow(int port, string hash)
        {
            if (OpenWindows.Count > 0 && OpenWindows[0].InvokeRequired)
            {
                OpenWindows[0].BeginInvoke(new Action(() => OpenNativeWindow(port, hash)));
                return;
            }
            var win = new OfficeMainWindow(port, hash);
            OpenWindows.Add(win);
            win.Show();
        }

        public static void OnWindowClosing(OfficeMainWindow win)
        {
            OpenWindows.Remove(win);
            if (OpenWindows.Count == 0)
            {
                // when LibreOffice's loop owns the UI thread, Application.Exit cannot end it: leave the process once no editor is open
                if (Lok.LoopRunning) { bool editors = false; foreach (Form f in Application.OpenForms) if (f is NativeDocForm && f.Visible) editors = true; if (!editors) Lok.HardExit(0); }
                Application.Exit();
            }
        }

        public static void ActivateExistingWindow()
        {
            if (OpenWindows.Count > 0)
            {
                var win = OpenWindows[0];
                if (win.InvokeRequired) win.BeginInvoke(new Action(() => win.ActivateWindow()));
                else win.ActivateWindow();
            }
        }

        public static bool OpenFallbackBrowser(int port, string hash)
        {
            string url = "http://127.0.0.1:" + port + "/index.html?host=win" + (string.IsNullOrEmpty(hash) ? "" : (hash.StartsWith("#") ? hash : "#" + hash));
            string browser = FindChromium();
            if (browser != null)
            {
                string prof = Path.Combine(DataDir, "profile");
                ProcessStartInfo psi = new ProcessStartInfo(browser,
                    "--app=\"" + url + "\" --user-data-dir=\"" + prof + "\" --window-size=1360,880 --no-first-run --no-default-browser-check --disable-sync --disable-extensions --disable-features=Translate,TranslateUI,msImplicitSignin,msEdgeSyncConsent,msUndersideButton --lang=vi --autoplay-policy=no-user-gesture-required");
                psi.UseShellExecute = false;
                try { Process.Start(psi); return true; } catch { }
            }
            return false;
        }

        public static bool OpenWindow(int port, string hash)
        {
            OpenNativeWindow(port, hash);
            return true;
        }

        [STAThread]
        static int Main(string[] args)
        {
            try { if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware(); } catch { }
            try { ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | (SecurityProtocolType)768 | SecurityProtocolType.Tls; } catch { }
            if (args.Length > 0 && args[0] == "--uninstall") { Installer.Uninstall(); return 0; }
            if (args.Length > 0 && args[0] == "--register") { Installer.RegisterAssociations(Assembly.GetExecutingAssembly().Location); return 0; }

            Application.EnableVisualStyles();
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (s0, e0) => { Lok.Trace("UI exception: " + e0.Exception); try { MessageBox.Show("Đã xảy ra lỗi: " + e0.Exception.Message + "\n\nỨng dụng sẽ tiếp tục chạy.", "Yomi-Office Open Source", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { } };
            AppDomain.CurrentDomain.UnhandledException += (s0, e0) => Lok.Trace("Unhandled: " + e0.ExceptionObject);
            Application.SetCompatibleTextRenderingDefault(false);

            // v0.1.6: native engine entry points (also used by the automated tests)
            //   OfficeOpenSource.exe --convert <in> <out>   headless conversion through LibreOfficeKit (exit 0 = ok)
            //   OfficeOpenSource.exe --native <file>        open the file directly in the native editor window
            if (args.Length >= 3 && args[0] == "--convert")
            {
                bool ok = false;
                // watchdog: a headless conversion must never hang its caller (engine start-up can rarely block)
                Thread wd = new Thread(() => { Thread.Sleep(300000); Lok.Trace("convert watchdog: no result after 5 min, exiting"); Lok.HardExit(5); }) { IsBackground = true }; wd.Start();
                // run inside a Win32 message loop: LibreOffice's windows/timers live on this thread
                if (!Lok.Init()) return 3;
                string cin = Path.GetFullPath(args[1]), cout = Path.GetFullPath(args[2]);
                Lok.InLoop(() => { try { ok = Lok.Convert(cin, cout, ""); } catch { } Lok.Trace("exit " + ok); Lok.HardExit(ok ? 0 : 3); });
                return ok ? 0 : 3;
                try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "oos_convert.log"), DateTime.Now + " " + args[1] + " -> " + args[2] + " ok=" + ok + " " + (ok ? "" : (Lok.LastError() ?? Lok.Error)) + "\r\n"); } catch { }
                return ok ? 0 : 3;
            }
            // --unotest <in> <out> <script>: runs .uno commands (one "cmd<TAB>json" per line, "type<TAB>text", "part<TAB>n") then saves; used to verify ribbon commands
            if (args.Length >= 4 && args[0] == "--unotest")
            {
                if (!Lok.Init()) return 3;
                string uin = Path.GetFullPath(args[1]), uout = Path.GetFullPath(args[2]); string[] lines = File.ReadAllLines(args[3]);
                Lok.InLoop(() =>
                {
                    bool ok = false;
                    try
                    {
                        LokDoc d = Lok.Load(uin, false);
                        if (d != null)
                        {
                            d.InitRender(""); Lok.Activate(d);
                            foreach (string ln in lines)
                            {
                                if (ln.Trim().Length == 0 || ln.StartsWith("#")) continue;
                                string[] kv = ln.Split(new[] { '\t' }, 2); string a2 = kv.Length > 1 ? kv[1] : "";
                                if (kv[0] == "type") { foreach (char ch in a2) { d.Key(0, ch, 0); d.Key(1, ch, 0); } }
                                else if (kv[0] == "part") d.Part = int.Parse(a2);
                                else if (kv[0] == "wins") { foreach (KeyValuePair<IntPtr, string> w in LokWin.SalWindows()) Lok.Trace("WIN " + w.Key + " " + w.Value); Lok.Trace("VIEW " + d.GetView() + " count " + d.ViewsCount()); }
                                else if (kv[0] == "activate") LokWin.ActivateFrames();
                                else if (kv[0] == "newview") { int ov = d.GetView(); int nv = d.CreateView(); Lok.Trace("CREATEVIEW " + ov + " -> " + nv); if (a2.StartsWith("back")) d.SetView(ov); if (a2 == "backdestroy") { d.DestroyView(nv); Lok.Trace("destroyed view " + nv + " count " + d.ViewsCount()); } }
                                else if (kv[0] == "click") { string[] xy = a2.Split(','); int mx = int.Parse(xy[0]), my = int.Parse(xy[1]); d.Mouse(0, mx, my, 1, 1, 0); d.Mouse(1, mx, my, 1, 1, 0); }
                                else if (kv[0] == "wait") { for (int wi = 0; wi < int.Parse(a2) / 50; wi++) { Application.DoEvents(); Thread.Sleep(50); } }
                                else if (kv[0] == "key") { int kc = int.Parse(a2); d.Key(0, kc == 1280 ? 13 : 0, kc); d.Key(1, kc == 1280 ? 13 : 0, kc); }
                                else if (kv[0] == "margs") File.WriteAllText(Path.Combine(Path.GetTempPath(), "oos_macro.txt"), a2.Replace("|", "\n"), new UTF8Encoding(false));
                                else if (kv[0] == "tool") { string mo = Path.Combine(Path.GetTempPath(), "oos_macro_out.txt"); try { File.Delete(mo); } catch { } Lok.RunMacro(a2.StartsWith("macro:") ? a2 : string.Format(BasicTools.Url, a2)); for (int wi = 0; wi < 40 && !File.Exists(mo); wi++) { Application.DoEvents(); Thread.Sleep(50); } Lok.Trace("TOOL " + a2 + " => " + (File.Exists(mo) ? File.ReadAllText(mo) : "(none)") + " err=" + Lok.LastError()); }
                                else if (kv[0] == "tile") { System.Drawing.Bitmap tb = new System.Drawing.Bitmap(512, 256, System.Drawing.Imaging.PixelFormat.Format32bppArgb); System.Drawing.Imaging.BitmapData bd = tb.LockBits(new System.Drawing.Rectangle(0, 0, 512, 256), System.Drawing.Imaging.ImageLockMode.WriteOnly, tb.PixelFormat); d.Paint(bd.Scan0, 512, 256, 0, 0, 7680, 3840); tb.UnlockBits(bd); tb.Save(a2); Lok.Trace("TILE " + a2); }
                                else if (kv[0] == "values") Lok.Trace("VALUES " + a2 + " => " + d.CommandValues(a2));
                                else { d.Uno(kv[0], a2.Length > 0 ? a2 : null); Lok.Trace("UNO " + kv[0] + " " + a2 + " err=" + Lok.LastError()); }
                                Application.DoEvents(); Thread.Sleep(150); Application.DoEvents();
                            }
                            ok = d.SaveAs(uout, Path.GetExtension(uout).TrimStart('.').ToLowerInvariant(), "");
                        }
                    }
                    catch (Exception ex) { Lok.Trace("unotest " + ex); }
                    Lok.HardExit(ok ? 0 : 3);
                });
                return 0;
            }
            string uiCmd = null; int ci = Array.IndexOf(args, "--cmd");   // v0.1.7: command handed over by the native editor
            if (ci >= 0 && ci + 1 < args.Length) { uiCmd = args[ci + 1]; List<string> la = new List<string>(args); la.RemoveRange(ci, 2); args = la.ToArray(); }
            bool webArgs = args.Length >= 2 && args[0] == "--web";   // native engine could not open the file -> full web editor
            if (webArgs) { string[] rest = new string[args.Length - 1]; Array.Copy(args, 1, rest, 0, rest.Length); args = rest; }
            if (args.Length >= 2 && args[0] == "--native")
            {
                bool portableN = File.Exists(Path.Combine(Dir, "portable.flag"));
                DataDir = portableN ? Path.Combine(Dir, "data") : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOpenSource");
                Directory.CreateDirectory(DataDir);
                if (!Lok.Init()) { MessageBox.Show(Lok.Error, "Yomi-Office Open Source"); return 4; }
                string nf = Path.GetFullPath(args[1]);
                Lok.InLoop(() => { NativeDocForm f = NativeDocForm.Open(nf); if (f == null) { try { System.Diagnostics.Process.Start(Application.ExecutablePath, "--web \"" + nf + "\""); } catch { } Lok.HardExit(4); } f.FormClosed += (s2, e2) => Lok.HardExit(0); });
                if (!Lok.Unipoll) Application.Run();
                return 0;
            }

            WebRoot = Path.Combine(Dir, "app");
            if (!File.Exists(Path.Combine(WebRoot, "index.html"))) {
                MessageBox.Show("Không tìm thấy thư mục ứng dụng:\n" + WebRoot, "Yomi-Office Open Source", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2;
            }
            bool portable = File.Exists(Path.Combine(Dir, "portable.flag"));
            DataDir = portable ? Path.Combine(Dir, "data") : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOpenSource");
            Directory.CreateDirectory(DataDir);
            string portFile = Path.Combine(DataDir, "server.port");

            // single instance: if a server is already running, hand over files and activate window
            bool created;
            mutex = new Mutex(true, "OfficeOpenSource_v011_Hybrid_" + DataDir.GetHashCode(), out created);
            if (!created)
            {
                int p;
                string[] eargs = args; if (uiCmd != null) { List<string> l2 = new List<string> { "cmd:" + uiCmd }; foreach (string x in args) if (x != "--web") l2.Add(x); eargs = l2.ToArray(); }
                if (File.Exists(portFile) && int.TryParse(File.ReadAllText(portFile).Trim(), out p) && Enqueue(p, eargs))
                {
                    try
                    {
                        using (WebClient wc = new WebClient()) wc.DownloadString("http://127.0.0.1:" + p + "/api/activate");
                    }
                    catch { }
                    return 0;
                }
            }
            foreach (string a in args) AddPending(a, webArgs, uiCmd);

            TcpListener l = new TcpListener(IPAddress.Loopback, 0); l.Start();
            Port = ((IPEndPoint)l.LocalEndpoint).Port;
            File.WriteAllText(portFile, Port.ToString());
            Thread t = new Thread(() => Serve(l)); t.IsBackground = true; t.Start();

            var mainWin = new OfficeMainWindow(Port, "");
            OpenWindows.Add(mainWin);
            // v0.1.6: warm up the native engine in the background so the first Word/Excel/PowerPoint file opens fast
            if (Lok.Available) { Thread warm = new Thread(() => { try { Lok.Init(); } catch { } }) { IsBackground = true, Name = "LOK warm-up" }; warm.Start(); }
            Application.Run(mainWin);

            try { File.Delete(portFile); } catch { }
            return 0;
        }

        static bool Enqueue(int port, string[] args)
        {
            try
            {
                using (WebClient wc = new WebClient())
                {
                    wc.Encoding = Encoding.UTF8; wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                    wc.UploadString("http://127.0.0.1:" + port + "/api/enqueue", Json.Serialize(new Dictionary<string, object> { { "paths", args } }));
                }
                return true;
            }
            catch { return false; }
        }

        public static void AddPending(string path) { AddPending(path, false); }
        public static void AddPending(string path, bool forceWeb) { AddPending(path, forceWeb, null); }
        public static void AddPending(string path, bool forceWeb, string cmd)
        {
            try { if (File.Exists(path)) { Dictionary<string, object> o = FileObj(path); if (forceWeb) o["forceWeb"] = true; if (!string.IsNullOrEmpty(cmd)) o["cmd"] = cmd; lock (Pending) Pending.Add(o); } } catch { }
        }
        public static Dictionary<string, object> FileObj(string path)
        {
            return new Dictionary<string, object> { { "name", Path.GetFileName(path) }, { "path", path }, { "b64", Convert.ToBase64String(File.ReadAllBytes(path)) } };
        }

        public static string FindBrowser() { return FindChromium(); }
        static string FindChromium()
        {
            string pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string pf86 = Environment.GetEnvironmentVariable("ProgramFiles(x86)") ?? pf;
            string lad = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string[] cands = {
                Path.Combine(pf86, @"Microsoft\Edge\Application\msedge.exe"), Path.Combine(pf, @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(pf, @"Google\Chrome\Application\chrome.exe"), Path.Combine(pf86, @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(lad, @"Google\Chrome\Application\chrome.exe"), Path.Combine(lad, @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(lad, @"Chromium\Application\chrome.exe"), Path.Combine(pf, @"BraveSoftware\Brave-Browser\Application\brave.exe")
            };
            foreach (string c in cands) if (File.Exists(c)) return c;
            foreach (string exe in new[] { "msedge.exe", "chrome.exe" })
                foreach (RegistryKey root in new[] { Registry.LocalMachine, Registry.CurrentUser })
                {
                    try { using (RegistryKey k = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + exe)) { if (k != null) { string v = k.GetValue("") as string; if (v != null && File.Exists(v.Trim('"'))) return v.Trim('"'); } } } catch { }
                }
            return null;
        }

        // ------------------------------------------------------------ minimal HTTP/1.1 server (no URL ACL / admin rights needed)
        static void Serve(TcpListener l)
        {
            while (true)
            {
                TcpClient c; try { c = l.AcceptTcpClient(); } catch { return; }
                ThreadPool.QueueUserWorkItem(o => { try { Handle((TcpClient)o); } catch { } }, c);
            }
        }
        static void Handle(TcpClient c)
        {
            using (c)
            using (NetworkStream ns = c.GetStream())
            {
                c.ReceiveTimeout = 30000;
                // read headers
                MemoryStream hb = new MemoryStream(); int last4 = 0;
                while (true) { int b = ns.ReadByte(); if (b < 0) return; hb.WriteByte((byte)b); last4 = (last4 << 8) | b; if (last4 == 0x0D0A0D0A) break; if (hb.Length > 65536) return; }
                string head = Encoding.ASCII.GetString(hb.ToArray());
                string[] lines = head.Split(new[] { "\r\n" }, StringSplitOptions.None);
                string[] rl = lines[0].Split(' '); if (rl.Length < 2) return;
                string method = rl[0], target = rl[1];
                int len = 0;
                foreach (string ln in lines) { int i = ln.IndexOf(':'); if (i > 0 && ln.Substring(0, i).Trim().Equals("Content-Length", StringComparison.OrdinalIgnoreCase)) int.TryParse(ln.Substring(i + 1).Trim(), out len); }
                byte[] body = new byte[len]; int got = 0; while (got < len) { int n = ns.Read(body, got, len - got); if (n <= 0) break; got += n; }
                string path = Uri.UnescapeDataString(target.Split('?')[0]);
                if (path.StartsWith("/api/")) { byte[] r = Encoding.UTF8.GetBytes(Api.Handle(path.Substring(5), Encoding.UTF8.GetString(body))); Write(ns, 200, "application/json; charset=utf-8", r); return; }
                if (path == "/") path = "/index.html";
                string full = Path.GetFullPath(Path.Combine(WebRoot, path.TrimStart('/').Replace('/', '\\')));
                if (!full.StartsWith(WebRoot, StringComparison.OrdinalIgnoreCase) || !File.Exists(full)) { Write(ns, 404, "text/plain", Encoding.UTF8.GetBytes("Not found")); return; }
                Write(ns, 200, Mime(full), File.ReadAllBytes(full));
            }
        }
        static void Write(Stream s, int code, string type, byte[] data)
        {
            string h = "HTTP/1.1 " + code + (code == 200 ? " OK" : " Not Found") + "\r\nContent-Type: " + type + "\r\nContent-Length: " + data.Length + "\r\nCache-Control: no-cache\r\nX-Content-Type-Options: nosniff\r\nConnection: close\r\n\r\n";
            byte[] hb = Encoding.ASCII.GetBytes(h); s.Write(hb, 0, hb.Length); s.Write(data, 0, data.Length); s.Flush();
        }
        static string Mime(string f)
        {
            switch (Path.GetExtension(f).ToLowerInvariant())
            {
                case ".html": return "text/html; charset=utf-8"; case ".js": return "application/javascript; charset=utf-8"; case ".css": return "text/css; charset=utf-8";
                case ".json": return "application/json"; case ".png": return "image/png"; case ".jpg": return "image/jpeg"; case ".svg": return "image/svg+xml";
                case ".ico": return "image/x-icon"; case ".woff2": return "font/woff2"; case ".woff": return "font/woff"; case ".ttf": return "font/ttf"; case ".wasm": return "application/wasm";
                default: return "application/octet-stream";
            }
        }
    }

    // ---------------------------------------------------------------- /api bridge (mirrors src/core/host.js)
    static class Api
    {
        public static string Handle(string name, string body)
        {
            Dictionary<string, object> q = null;
            try { q = string.IsNullOrEmpty(body) ? new Dictionary<string, object>() : App.Json.Deserialize<Dictionary<string, object>>(body); } catch { q = new Dictionary<string, object>(); }
            try
            {
                switch (name)
                {
                    case "activate": App.ActivateExistingWindow(); return "{\"ok\":true}";
                    case "ping": App.LastPing = DateTime.UtcNow; return "{\"ok\":true}";
                    case "closing": App.LastPing = DateTime.UtcNow.AddSeconds(-15); return "{\"ok\":true}";
                    case "info": return App.Json.Serialize(new Dictionary<string, object> { { "platform", "windows" }, { "version", App.Version }, { "os", Environment.OSVersion.VersionString }, { "x64", Environment.Is64BitOperatingSystem }, { "clr", Environment.Version.ToString() }, { "portable", File.Exists(Path.Combine(App.Dir, "portable.flag")) }, { "msi", IsMsiInstall() }, { "appx", App.Dir.IndexOf("\\WindowsApps\\", StringComparison.OrdinalIgnoreCase) >= 0 }, { "dir", App.Dir } });
                    case "aiweb": return AiWeb(Str(q, "url"), Str(q, "title"));
                    case "clipboard": { if (q.ContainsKey("set")) { string t = Str(q, "set"); OnSta(() => { try { Clipboard.SetText(t); } catch { } return 0; }); return "{\"ok\":true}"; } string txt = OnSta(() => { try { return Clipboard.ContainsText() ? Clipboard.GetText() : ""; } catch { return ""; } }); return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "text", txt } }); }
                    case "scanfiles": return ScanFiles(Str(q, "folder"));
                    case "scan_office": return ScanOffice();
                    case "httpget": return HttpGet(Str(q, "url"));
                    case "update": return Updater.Start(Str(q, "url"), Str(q, "name"), Str(q, "mode"));
                    case "updatestatus": return Updater.Status();
                    case "pending": lock (App.Pending) { string r = App.Json.Serialize(new Dictionary<string, object> { { "files", App.Pending.ToArray() } }); App.Pending.Clear(); return r; }
                    case "enqueue": { object[] ps = q.ContainsKey("paths") ? q["paths"] as object[] : null; if (ps == null && q.ContainsKey("paths")) { System.Collections.ArrayList al = q["paths"] as System.Collections.ArrayList; if (al != null) ps = al.ToArray(); } if (ps != null) { string ecmd = null; foreach (object p in ps) { string sp = Convert.ToString(p); if (sp.StartsWith("cmd:")) ecmd = sp.Substring(4); else App.AddPending(sp, ecmd != null, ecmd); } } return "{\"ok\":true}"; }
                    case "open": return Open(Str(q, "filter"), Bool(q, "multiple"));
                    case "save": return Save(Str(q, "name"), Str(q, "path"), Str(q, "b64"), Bool(q, "saveAs"));
                    case "read": { string p = Str(q, "path"); if (!File.Exists(p)) return "{\"ok\":false}"; Dictionary<string, object> f = App.FileObj(p); f["ok"] = true; return App.Json.Serialize(f); }
                    case "openurl": { string u = Str(q, "url"); if (u.StartsWith("http://") || u.StartsWith("https://") || u.StartsWith("mailto:")) Process.Start(u); return "{\"ok\":true}"; }
                    case "newwindow": App.OpenNativeWindow(App.Port, Str(q, "hash")); return "{\"ok\":true}";
                    case "stash": { string tok = Guid.NewGuid().ToString("N"); Dictionary<string, object> f = new Dictionary<string, object> { { "name", Str(q, "name") }, { "path", Str(q, "path") }, { "b64", Str(q, "b64") } };
                        if (f["b64"] as string == "" && File.Exists(Str(q, "path"))) f = App.FileObj(Str(q, "path"));
                        lock (App.Stash) App.Stash[tok] = f; App.OpenNativeWindow(App.Port, "open=" + tok); return "{\"ok\":true,\"token\":\"" + tok + "\"}"; }
                    case "unstash": { string tok = Str(q, "token"); Dictionary<string, object> f = null; lock (App.Stash) { if (App.Stash.TryGetValue(tok, out f)) App.Stash.Remove(tok); } return App.Json.Serialize(new Dictionary<string, object> { { "files", f == null ? new object[0] : new object[] { f } } }); }
                    case "lan_share": return App.Json.Serialize(LanShare.Share(Str(q, "name"), Convert.FromBase64String(Str(q, "b64"))));
                    case "lan_received": { string[] got; lock (LanShare.Received) { got = LanShare.Received.ToArray(); LanShare.Received.Clear(); } return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "files", got } }); }
                    case "lok_info": return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "available", Lok.Available }, { "dir", Lok.FindProgramDir() ?? "" }, { "version", Lok.VersionInfo ?? "" } });
                    case "lok_open": return LokOpen(Str(q, "path"), Str(q, "name"), Str(q, "b64"), Str(q, "cmd"));
                    case "lok_convert": return LokConvert(Str(q, "path"), Str(q, "name"), Str(q, "b64"), Str(q, "fmt"));
                    case "exit": App.LastPing = DateTime.UtcNow.AddSeconds(-60); return "{\"ok\":true}";
                    case "proxy": return Proxy(Str(q, "url"), q.ContainsKey("headers") ? q["headers"] as Dictionary<string, object> : null, Str(q, "body"));
                    case "esbuild_transform": return EsbuildTransform(Str(q, "code"), Str(q, "loader"), Str(q, "jsxFactory"), Str(q, "jsxFragment"));
                    case "http": return Http(Str(q, "url"), Str(q, "method"), q.ContainsKey("headers") ? q["headers"] as Dictionary<string, object> : null, Str(q, "b64"));
                    case "oauth": return OAuth.Start(Str(q, "url"), q.ContainsKey("port") ? Convert.ToInt32(q["port"]) : 53682, Bool(q, "nobrowser"));
                    case "oauthpoll": return OAuth.Poll();
                    case "settings_get": {
                        string sf = Path.Combine(App.DataDir, "settings.json");
                        {   // v0.1.7: same author name / avatar as the native editor (Windows user) until the user sets one
                            Dictionary<string, object> js = null; try { if (File.Exists(sf)) js = App.Json.Deserialize<Dictionary<string, object>>(File.ReadAllText(sf, Encoding.UTF8)); } catch { }
                            if (js == null) js = new Dictionary<string, object>();
                            object un; js.TryGetValue("userName", out un);
                            if (un == null || Convert.ToString(un) == "Người dùng") { string nm = NativeSettings.Get("user", ""); if (nm.Length == 0) nm = Ui.UserName(); js["userName"] = nm; js["initials"] = Ui.Initials(nm); try { File.WriteAllText(sf, App.Json.Serialize(js), new UTF8Encoding(false)); } catch { } }
                        }
                        if (File.Exists(sf)) {
                            try { return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "settings", File.ReadAllText(sf, Encoding.UTF8) } }); }
                            catch (Exception ex) { return App.Json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", ex.Message } }); }
                        }
                        return "{\"ok\":true,\"settings\":null}";
                    }
                    case "settings_save": {
                        string stxt = Str(q, "settings");
                        if (!string.IsNullOrEmpty(stxt)) {
                            try {
                                string sf = Path.Combine(App.DataDir, "settings.json");
                                File.WriteAllText(sf, stxt, Encoding.UTF8);
                                return "{\"ok\":true}";
                            } catch (Exception ex) { return App.Json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", ex.Message } }); }
                        }
                        return "{\"ok\":false,\"error\":\"empty\"}";
                    }
                    case "list_dir": {   // v0.1.7: Backstage "Lưu thành" folder list in the web editor (same as the native page)
                        string d = Str(q, "dir"); if (string.IsNullOrEmpty(d) || !Directory.Exists(d)) d = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        List<object> its = new List<object>();
                        try {
                            foreach (string x in Directory.GetDirectories(d)) { try { if ((File.GetAttributes(x) & FileAttributes.Hidden) != 0) continue; } catch { continue; } its.Add(new Dictionary<string, object> { { "name", Path.GetFileName(x) }, { "path", x }, { "dir", true }, { "mtime", (long)(Directory.GetLastWriteTimeUtc(x) - new DateTime(1970, 1, 1)).TotalMilliseconds } }); }
                            foreach (string x in Directory.GetFiles(d)) { string fn = Path.GetFileName(x); if (fn.StartsWith("~$") || fn.StartsWith(".~lock")) continue; its.Add(new Dictionary<string, object> { { "name", fn }, { "path", x }, { "dir", false }, { "mtime", (long)(File.GetLastWriteTimeUtc(x) - new DateTime(1970, 1, 1)).TotalMilliseconds } }); }
                        } catch { }
                        DirectoryInfo up = null; try { up = Directory.GetParent(d); } catch { }
                        return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "dir", d }, { "parent", up != null ? up.FullName : "" }, { "items", its } });
                    }
                    case "pick_folder": {
                        string pf = OnSta(() => { using (FolderBrowserDialog f = new FolderBrowserDialog()) using (Form o = TopForm()) return f.ShowDialog(o) == DialogResult.OK ? f.SelectedPath : null; });
                        return App.Json.Serialize(new Dictionary<string, object> { { "ok", pf != null }, { "path", pf ?? "" } });
                    }
                    case "file_info": { string fp = Str(q, "path"); if (!File.Exists(fp)) return "{\"ok\":false}"; FileInfo fi = new FileInfo(fp); DateTime ep = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc); return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "size", fi.Length }, { "ctime", (long)(fi.CreationTimeUtc - ep).TotalMilliseconds }, { "mtime", (long)(fi.LastWriteTimeUtc - ep).TotalMilliseconds } }); }
                    case "reveal": { string rp = Str(q, "path"); try { if (File.Exists(rp)) Process.Start("explorer.exe", "/select,\"" + rp + "\""); } catch { } return "{\"ok\":true}"; }
                    case "lan_info": {
                        var ips = new List<string>();
                        try {
                            var host = Dns.GetHostEntry(Dns.GetHostName());
                            foreach (var ip in host.AddressList) {
                                if (ip.AddressFamily == AddressFamily.InterNetwork && !IPAddress.IsLoopback(ip)) {
                                    ips.Add(ip.ToString());
                                }
                            }
                        } catch { }
                        return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "port", App.Port }, { "ips", ips.ToArray() } });
                    }
                }
                return "{\"ok\":false,\"error\":\"unknown api\"}";
            }
            catch (Exception e) { return App.Json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", e.Message } }); }
        }

        static bool IsMsiInstall()
        {
            return File.Exists(Path.Combine(App.Dir, "installed_by_msi.txt"));
        }

        static string FindTestOfficeDir()
        {
            try
            {
                string cur = App.Dir;
                for (int i = 0; i < 6 && !string.IsNullOrEmpty(cur); i++)
                {
                    string cand = Path.Combine(cur, "Test Office");
                    if (Directory.Exists(cand)) return cand;
                    DirectoryInfo p = Directory.GetParent(cur);
                    cur = p != null ? p.FullName : null;
                }
            }
            catch { }
            return null;
        }

        static string ScanFiles(string customFolder)
        {
            var list = new List<Dictionary<string, object>>();
            var dirs = new List<string>();
            try { dirs.Add(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)); } catch { }
            try { dirs.Add(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)); } catch { }
            try { dirs.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")); } catch { }
            string testOffice = FindTestOfficeDir();
            if (!string.IsNullOrEmpty(testOffice)) dirs.Add(testOffice);
            if (!string.IsNullOrEmpty(customFolder) && Directory.Exists(customFolder)) dirs.Add(customFolder);

            var validExts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                ".docx", ".docm", ".dotx", ".doc",
                ".xlsx", ".xlsm", ".xltx", ".xls", ".csv",
                ".pptx", ".pptm", ".ppsx", ".ppt",
                ".pdf", ".odt", ".ods", ".odp", ".txt", ".md",
                ".tex", ".latex", ".mml", ".eqx", ".mathtype"
            };

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string dir in dirs)
            {
                if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) continue;
                try
                {
                    ScanDirRecursive(new DirectoryInfo(dir), 0, 2, validExts, seen, list);
                }
                catch { }
            }
            return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "files", list.ToArray() } });
        }

        static void ScanDirRecursive(DirectoryInfo di, int depth, int maxDepth, HashSet<string> validExts, HashSet<string> seen, List<Dictionary<string, object>> list)
        {
            try
            {
                foreach (FileInfo fi in di.GetFiles())
                {
                    if ((fi.Attributes & FileAttributes.Hidden) != 0) continue;
                    string ext = fi.Extension.ToLowerInvariant();
                    if (validExts.Contains(ext) && seen.Add(fi.FullName))
                    {
                        string kind = "word";
                        if (ext.StartsWith(".xl") || ext == ".csv" || ext == ".ods") kind = "excel";
                        else if (ext.StartsWith(".pp") || ext == ".odp") kind = "ppt";
                        else if (ext == ".pdf") kind = "pdf";
                        else if (ext == ".tex" || ext == ".latex" || ext == ".mml" || ext == ".eqx" || ext == ".mathtype") kind = "mathtype";

                        list.Add(new Dictionary<string, object> {
                            { "name", fi.Name },
                            { "path", fi.FullName },
                            { "ext", ext.TrimStart('.') },
                            { "size", fi.Length },
                            { "modified", fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") },
                            { "kind", kind }
                        });
                    }
                }
                if (depth < maxDepth)
                {
                    foreach (DirectoryInfo sub in di.GetDirectories())
                    {
                        if ((sub.Attributes & FileAttributes.Hidden) != 0 || sub.Name.StartsWith(".") || sub.Name.Equals("node_modules", StringComparison.OrdinalIgnoreCase)) continue;
                        ScanDirRecursive(sub, depth + 1, maxDepth, validExts, seen, list);
                    }
                }
            }
            catch { }
        }

        static string ScanOffice()
        {
            var list = new List<Dictionary<string, object>>();
            var dirs = new List<string>();
            try { string up = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile); if (!string.IsNullOrEmpty(up) && Directory.Exists(up)) dirs.Add(up); } catch { }
            try { string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); if (!string.IsNullOrEmpty(docs) && Directory.Exists(docs)) dirs.Add(docs); } catch { }
            try { string desk = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); if (!string.IsNullOrEmpty(desk) && Directory.Exists(desk)) dirs.Add(desk); } catch { }
            try { string dl = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"); if (Directory.Exists(dl)) dirs.Add(dl); } catch { }
            foreach (string odVar in new[] { "OneDrive", "OneDriveConsumer", "OneDriveCommercial" })
                try { string od = Environment.GetEnvironmentVariable(odVar); if (!string.IsNullOrEmpty(od) && Directory.Exists(od)) dirs.Add(od); } catch { }
            string testOffice = FindTestOfficeDir();
            if (!string.IsNullOrEmpty(testOffice)) dirs.Add(testOffice);

            var officeExts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                ".docx", ".docm", ".doc", ".dotx", ".dotm", ".dot", ".odt", ".ott", ".rtf",
                ".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx", ".xltm", ".ods", ".ots", ".csv",
                ".pptx", ".pptm", ".ppt", ".ppsx", ".pps", ".potx", ".potm", ".odp", ".otp"
            };
            var skipDirs = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                "node_modules", "AppData", "Application Data", "Windows", "Program Files", "Program Files (x86)",
                "ProgramData", "System Volume Information", "Recovery", "PerfLogs", "__pycache__",
                ".git", "build", "dist", "obj", "bin", "Temp", "tmp", "Cache", "cache"
            };
            var seenDirs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var sw = Stopwatch.StartNew();
            var queue = new Queue<KeyValuePair<string, int>>();
            foreach (string d in dirs) if (!string.IsNullOrEmpty(d) && Directory.Exists(d)) queue.Enqueue(new KeyValuePair<string, int>(d, 0));

            while (queue.Count > 0 && seenDirs.Count < 6000 && list.Count < 5000 && sw.ElapsedMilliseconds < 15000)
            {
                var item = queue.Dequeue();
                string dir = item.Key; int depth = item.Value;
                if (!seenDirs.Add(dir)) continue;
                try
                {
                    DirectoryInfo di = new DirectoryInfo(dir);
                    foreach (FileInfo fi in di.GetFiles())
                    {
                        if (list.Count >= 5000) break;
                        if ((fi.Attributes & FileAttributes.Hidden) != 0) continue;
                        if (fi.Name.StartsWith("~$", StringComparison.Ordinal)) continue;
                        string ext = fi.Extension.ToLowerInvariant();
                        if (!officeExts.Contains(ext) || !seenFiles.Add(fi.FullName)) continue;
                        string kind = ext.StartsWith(".x") || ext == ".ods" || ext == ".ots" || ext == ".csv" ? "excel"
                            : ext.StartsWith(".p") || ext == ".odp" || ext == ".otp" ? "ppt" : "word";
                        list.Add(new Dictionary<string, object> {
                            { "name", fi.Name },
                            { "path", fi.FullName },
                            { "kind", kind },
                            { "mtime", (long)(fi.LastWriteTimeUtc - epoch).TotalMilliseconds },
                            { "size", fi.Length },
                            { "dir", dir }
                        });
                    }
                    if (depth < 7)
                    {
                        foreach (DirectoryInfo sub in di.GetDirectories())
                        {
                            if ((sub.Attributes & (FileAttributes.Hidden | FileAttributes.System | FileAttributes.ReparsePoint)) != 0) continue;
                            if (sub.Name.StartsWith(".") || sub.Name.StartsWith("$") || sub.Name.StartsWith("~") || skipDirs.Contains(sub.Name)) continue;
                            queue.Enqueue(new KeyValuePair<string, int>(sub.FullName, depth + 1));
                        }
                    }
                }
                catch { }
            }
            list.Sort((a, b) => ((long)b["mtime"]).CompareTo((long)a["mtime"]));
            return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "files", list.ToArray() } });
        }

        internal static string AiWeb(string url, string title)
        {
            if (!(url.StartsWith("https://"))) return "{\"ok\":false,\"error\":\"blocked\"}";
            string browser = App.FindBrowser();
            if (browser == null) { Process.Start(url); return "{\"ok\":true}"; }
            string prof = Path.Combine(App.DataDir, "profile-ai");
            int w = 560, h = 900, x = 0;
            try { Rectangle wa = Screen.PrimaryScreen.WorkingArea; h = Math.Min(h, wa.Height); x = wa.Right - w; } catch { }
            ProcessStartInfo psi = new ProcessStartInfo(browser, "--app=\"" + url + "\" --user-data-dir=\"" + prof + "\" --window-size=" + w + "," + h + " --window-position=" + x + ",0 --no-first-run --no-default-browser-check --lang=vi");
            psi.UseShellExecute = false; Process.Start(psi);
            return "{\"ok\":true}";
        }
        static string HttpGet(string url)
        {
            if (!url.StartsWith("https://") && !url.StartsWith("http://")) return "{\"ok\":false,\"error\":\"blocked\"}";
            try
            {
                HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url); r.Method = "GET"; r.Timeout = 25000; r.AllowAutoRedirect = true; r.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) YomiOfficeOpenSource/" + App.Version;
                using (WebResponse resp = r.GetResponse()) using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                    return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "body", sr.ReadToEnd() } });
            }
            catch (Exception ex)
            {
                return App.Json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", ex.Message } });
            }
        }
        // v0.1.6: native LibreOfficeKit editor window (no WebView) and headless conversion
        static string MaterializeTemp(string path, string name, string b64)
        {
            if (!string.IsNullOrEmpty(path) && File.Exists(path)) return path;
            if (string.IsNullOrEmpty(b64)) return null;
            string dir = Path.Combine(App.DataDir, "lok-temp", Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(dir);
            string f = Path.Combine(dir, string.IsNullOrEmpty(name) ? "document.docx" : Path.GetFileName(name));
            File.WriteAllBytes(f, Convert.FromBase64String(b64));
            return f;
        }
        static string LokOpen(string path, string name, string b64) { return LokOpen(path, name, b64, null); }
        static string LokOpen(string path, string name, string b64, string cmd)
        {
            string f = MaterializeTemp(path, name, b64);
            if (f == null) return "{\"ok\":false,\"error\":\"no file\"}";
            if (!Lok.Available) return "{\"ok\":false,\"error\":\"engine missing\"}";
            Form host = App.OpenWindows.Count > 0 ? (Form)App.OpenWindows[0] : null;
            if (host == null) return "{\"ok\":false,\"error\":\"no window\"}";
            // the native editor runs inside LibreOffice's main loop on the UI thread; answer the web UI right away
            host.BeginInvoke(new Action(() => NativeDocForm.OpenAsync(f, cmd)));
            return "{\"ok\":true}";
        }
        static string LokConvert(string path, string name, string b64, string fmt)
        {
            string f = MaterializeTemp(path, name, b64);
            if (f == null || !Lok.Available) return "{\"ok\":false}";
            string outp = Path.Combine(Path.GetTempPath(), "oos-conv-" + Guid.NewGuid().ToString("N") + "." + fmt);
            bool ok = false;
            Form host = App.OpenWindows.Count > 0 ? (Form)App.OpenWindows[0] : null;
            if (host == null || !Lok.Init()) return "{\"ok\":false}";
            ManualResetEvent done = new ManualResetEvent(false);
            host.BeginInvoke(new Action(() => Lok.InLoop(() => { try { ok = Lok.Convert(f, outp, ""); } finally { done.Set(); } })));
            done.WaitOne(300000);
            if (!ok || !File.Exists(outp)) return App.Json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", Lok.LastError() ?? "" } });
            string r = App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "name", Path.GetFileNameWithoutExtension(string.IsNullOrEmpty(name) ? f : name) + "." + fmt }, { "b64", Convert.ToBase64String(File.ReadAllBytes(outp)) } });
            try { File.Delete(outp); } catch { }
            return r;
        }
        static string Str(Dictionary<string, object> q, string k) { object v; return q != null && q.TryGetValue(k, out v) && v != null ? Convert.ToString(v) : ""; }
        static bool Bool(Dictionary<string, object> q, string k) { object v; return q != null && q.TryGetValue(k, out v) && v is bool && (bool)v; }

        static T OnSta<T>(Func<T> f)
        {
            // v0.1.6: run shell dialogs on the UI thread (a throw-away STA thread left explorerframe.dll callbacks behind
            // that crashed the process once the LibreOffice engine was loaded)
            Form ui = App.OpenWindows.Count > 0 ? (Form)App.OpenWindows[0] : null;
            if (ui != null && ui.IsHandleCreated && !ui.IsDisposed) { T r = default(T); ui.Invoke(new Action(() => { r = f(); })); return r; }
            T res = default(T); Exception err = null;
            Thread t = new Thread(() => { try { res = f(); } catch (Exception e) { err = e; } });
            t.SetApartmentState(ApartmentState.STA); t.Start(); t.Join();
            if (err != null) throw err; return res;
        }
        static string Filter(string accept)
        {
            string a = (accept ?? "").ToLowerInvariant();
            // v0.1.6 fix: the generic "Open" list also contains .pdf, which used to select the PDF-only filter
            bool office = a.Contains("doc") || a.Contains("xls") || a.Contains("ppt");
            if (a.Contains("image") && !office) return "Ảnh|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp|Mọi tệp|*.*";
            if (a.Contains("pdf") && !office) return "PDF|*.pdf|Mọi tệp|*.*";
            return "Tài liệu Office|*.docx;*.docm;*.dotx;*.xlsx;*.xlsm;*.xltx;*.pptx;*.pptm;*.ppsx;*.potx;*.pdf;*.epub;*.fb2;*.csv;*.txt;*.md;*.html;*.htm;*.mht;*.xml;*.rtf;*.doc;*.xls;*.ppt;*.odt;*.ods;*.odp;*.slk;*.dif|Word|*.docx;*.docm;*.dotx;*.doc;*.rtf;*.odt;*.txt;*.md;*.mht|Excel|*.xlsx;*.xlsm;*.xltx;*.xls;*.csv;*.ods;*.slk;*.dif|PowerPoint|*.pptx;*.pptm;*.ppsx;*.potx;*.ppt;*.odp|PDF|*.pdf|Ảnh|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp|Mọi tệp|*.*";
        }
        static string Open(string accept, bool multi)
        {
            string[] files = OnSta(() =>
            {
                using (OpenFileDialog d = new OpenFileDialog { Filter = Filter(accept), Multiselect = multi, Title = "Mở tệp - Yomi-Office Open Source" })
                    using (Form o = TopForm()) return d.ShowDialog(o) == DialogResult.OK ? d.FileNames : new string[0];
            });
            List<object> list = new List<object>(); foreach (string f in files) list.Add(App.FileObj(f));
            return App.Json.Serialize(new Dictionary<string, object> { { "files", list } });
        }
        static string Save(string name, string path, string b64, bool saveAs)
        {
            byte[] data = Convert.FromBase64String(b64 ?? "");
            if (!saveAs && !string.IsNullOrEmpty(path) && Directory.Exists(Path.GetDirectoryName(path)))
            { File.WriteAllBytes(path, data); return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "name", Path.GetFileName(path) }, { "path", path } }); }
            string target = OnSta(() =>
            {
                string ext = Path.GetExtension(name ?? "").TrimStart('.');
                using (SaveFileDialog d = new SaveFileDialog { FileName = name, Title = "Lưu tệp - Yomi-Office Open Source", Filter = (ext.Length > 0 ? ext.ToUpperInvariant() + "|*." + ext + "|" : "") + "Mọi tệp|*.*", OverwritePrompt = true })
                {
                    if (!string.IsNullOrEmpty(path)) try { d.InitialDirectory = Path.GetDirectoryName(path); } catch { }
                    using (Form o = TopForm()) return d.ShowDialog(o) == DialogResult.OK ? d.FileName : null;
                }
            });
            if (target == null) return "{\"ok\":false}";
            File.WriteAllBytes(target, data);
            return App.Json.Serialize(new Dictionary<string, object> { { "ok", true }, { "name", Path.GetFileName(target) }, { "path", target } });
        }
        static Form TopForm() { Form f = new Form { TopMost = true, ShowInTaskbar = false, Opacity = 0, StartPosition = FormStartPosition.CenterScreen, Width = 1, Height = 1 }; f.Show(); f.Activate(); return f; }

        static string EsbuildTransform(string code, string loader, string jsxFactory, string jsxFragment)
        {
            try
            {
                string nm = Path.Combine(App.Dir, "node_modules");
                if (!Directory.Exists(nm))
                {
                    string parent = Path.GetFullPath(Path.Combine(App.Dir, "..\\..\\..\\v005\\deps\\node_modules"));
                    if (Directory.Exists(parent)) nm = parent;
                }
                string tempScript = Path.Combine(Path.GetTempPath(), "esbuild_tr_" + Guid.NewGuid().ToString("N") + ".js");
                string tempCode = Path.Combine(Path.GetTempPath(), "esbuild_cd_" + Guid.NewGuid().ToString("N") + ".txt");
                File.WriteAllText(tempCode, code ?? "", Encoding.UTF8);

                string js = "const fs = require('fs');\n" +
                            "const esbuild = require(" + App.Json.Serialize(Path.Combine(nm, "esbuild").Replace('\\', '/')) + ");\n" +
                            "const c = fs.readFileSync(" + App.Json.Serialize(tempCode.Replace('\\', '/')) + ", 'utf8');\n" +
                            "const res = esbuild.transformSync(c, { loader: " + App.Json.Serialize(string.IsNullOrEmpty(loader) ? "ts" : loader) + ", jsxFactory: " + App.Json.Serialize(string.IsNullOrEmpty(jsxFactory) ? "h" : jsxFactory) + ", jsxFragment: " + App.Json.Serialize(string.IsNullOrEmpty(jsxFragment) ? "Fragment" : jsxFragment) + " });\n" +
                            "process.stdout.write(JSON.stringify({ ok: true, code: res.code }));\n";
                File.WriteAllText(tempScript, js, Encoding.UTF8);

                ProcessStartInfo psi = new ProcessStartInfo("node.exe", "\"" + tempScript + "\"");
                psi.UseShellExecute = false;
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
                psi.CreateNoWindow = true;
                using (Process p = Process.Start(psi))
                {
                    string stdout = p.StandardOutput.ReadToEnd();
                    string stderr = p.StandardError.ReadToEnd();
                    p.WaitForExit(5000);
                    try { File.Delete(tempScript); File.Delete(tempCode); } catch { }
                    if (!string.IsNullOrEmpty(stdout) && stdout.StartsWith("{")) return stdout;
                    return App.Json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", stderr } });
                }
            }
            catch (Exception ex)
            {
                return App.Json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", ex.Message } });
            }
        }

        static string Http(string url, string method, Dictionary<string, object> headers, string b64)
        {
            if (!url.StartsWith("https://")) return "{\"error\":\"blocked\"}";
            try
            {
                HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url); r.Method = string.IsNullOrEmpty(method) ? "GET" : method.ToUpperInvariant(); r.Timeout = 300000; r.ReadWriteTimeout = 300000; r.AllowAutoRedirect = true;
                r.UserAgent = "YomiOfficeOpenSource/" + App.Version + " (Windows)";
                if (headers != null) foreach (KeyValuePair<string, object> kv in headers)
                {
                    string v = Convert.ToString(kv.Value);
                    if (kv.Key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase)) r.ContentType = v;
                    else if (kv.Key.Equals("Accept", StringComparison.OrdinalIgnoreCase)) r.Accept = v;
                    else if (kv.Key.Equals("If-Match", StringComparison.OrdinalIgnoreCase)) r.Headers.Add("If-Match", v);
                    else r.Headers[kv.Key] = v;
                }
                byte[] body = string.IsNullOrEmpty(b64) ? null : Convert.FromBase64String(b64);
                if (body != null || r.Method == "POST" || r.Method == "PUT" || r.Method == "PATCH") { if (body == null) body = new byte[0]; r.ContentLength = body.Length; using (Stream st = r.GetRequestStream()) st.Write(body, 0, body.Length); }
                HttpWebResponse resp;
                try { resp = (HttpWebResponse)r.GetResponse(); }
                catch (WebException we) { resp = we.Response as HttpWebResponse; if (resp == null) return App.Json.Serialize(new Dictionary<string, object> { { "error", we.Message } }); }
                using (resp)
                using (Stream rs = resp.GetResponseStream())
                using (MemoryStream ms = new MemoryStream())
                {
                    byte[] buf = new byte[65536]; int n; while ((n = rs.Read(buf, 0, buf.Length)) > 0) ms.Write(buf, 0, n);
                    return App.Json.Serialize(new Dictionary<string, object> { { "status", (int)resp.StatusCode }, { "type", resp.ContentType ?? "" }, { "b64", Convert.ToBase64String(ms.ToArray()) } });
                }
            }
            catch (Exception e) { return App.Json.Serialize(new Dictionary<string, object> { { "error", e.Message } }); }
        }
        static string Proxy(string url, Dictionary<string, object> headers, string body)
        {
            if (!(url.StartsWith("https://") || url.StartsWith("http://127.0.0.1") || url.StartsWith("http://localhost"))) return "{\"error\":\"blocked\"}";
            HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url); r.Method = "POST"; r.Timeout = 120000;
            if (headers != null) foreach (KeyValuePair<string, object> kv in headers) { if (kv.Key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase)) r.ContentType = Convert.ToString(kv.Value); else r.Headers[kv.Key] = Convert.ToString(kv.Value); }
            byte[] b = Encoding.UTF8.GetBytes(body ?? ""); r.ContentLength = b.Length; using (Stream s = r.GetRequestStream()) s.Write(b, 0, b.Length);
            string txt;
            try { using (WebResponse resp = r.GetResponse()) using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) txt = sr.ReadToEnd(); }
            catch (WebException we) { if (we.Response == null) throw; using (StreamReader sr = new StreamReader(we.Response.GetResponseStream(), Encoding.UTF8)) txt = sr.ReadToEnd(); }
            return App.Json.Serialize(new Dictionary<string, object> { { "body", txt } });
        }
    }

    // ---------------------------------------------------------------- OAuth loopback receiver
    static class OAuth
    {
        static string state = "idle", query = "", error = ""; static TcpListener l4, l6; static int gen = 0;
        public static string Poll() { return App.Json.Serialize(new Dictionary<string, object> { { "state", state }, { "query", query }, { "error", error } }); }
        public static string Start(string url, int port, bool noBrowser)
        {
            if (!url.StartsWith("https://")) return "{\"ok\":false,\"error\":\"blocked\"}";
            Stop(); int my = ++gen; state = "waiting"; query = ""; error = "";
            try { l4 = new TcpListener(IPAddress.Loopback, port); l4.Start(); } catch (Exception e) { state = "error"; error = "Cổng " + port + " đang bận: " + e.Message; return App.Json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", error } }); }
            try { l6 = new TcpListener(IPAddress.IPv6Loopback, port); l6.Start(); } catch { l6 = null; }
            foreach (TcpListener l in new TcpListener[] { l4, l6 }) { if (l == null) continue; TcpListener ll = l; Thread t = new Thread(() => Accept(ll, my)); t.IsBackground = true; t.Start(); }
            Thread to = new Thread(() => { Thread.Sleep(600000); if (gen == my && state == "waiting") { state = "error"; error = "Hết thời gian chờ"; Stop(); } }); to.IsBackground = true; to.Start();
            try { if (!noBrowser) Process.Start(url); } catch (Exception e) { state = "error"; error = e.Message; Stop(); return App.Json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", error } }); }
            return "{\"ok\":true}";
        }
        static void Stop() { try { if (l4 != null) l4.Stop(); } catch { } try { if (l6 != null) l6.Stop(); } catch { } l4 = null; l6 = null; }
        static void Accept(TcpListener l, int my)
        {
            while (gen == my && state == "waiting")
            {
                TcpClient c; try { c = l.AcceptTcpClient(); } catch { return; }
                try
                {
                    using (c) using (NetworkStream ns = c.GetStream())
                    {
                        ns.ReadTimeout = 10000; byte[] buf = new byte[16384]; int n = ns.Read(buf, 0, buf.Length); string req = Encoding.ASCII.GetString(buf, 0, Math.Max(0, n));
                        string line = req.Split('\n')[0]; string[] parts = line.Split(' '); string path = parts.Length > 1 ? parts[1] : "/";
                        bool hit = path.IndexOf("code=") >= 0 || path.IndexOf("error=") >= 0;
                        string html = hit ? "<!doctype html><meta charset=utf-8><title>Yomi-Office Open Source</title><body style='font-family:Segoe UI,Arial;text-align:center;padding:60px'><h2>" + (path.IndexOf("code=") >= 0 ? "Đăng nhập thành công" : "Đăng nhập bị hủy") + "</h2><p>Bạn có thể đóng thẻ này và quay lại Yomi-Office Open Source.</p><script>setTimeout(function(){window.close()},1500)</script>" : "";
                        byte[] body = Encoding.UTF8.GetBytes(html);
                        string head = (hit ? "HTTP/1.1 200 OK" : "HTTP/1.1 404 Not Found") + "\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: " + body.Length + "\r\nConnection: close\r\n\r\n";
                        byte[] hb = Encoding.ASCII.GetBytes(head); ns.Write(hb, 0, hb.Length); ns.Write(body, 0, body.Length); ns.Flush();
                        if (hit) { query = path; state = "done"; Stop(); return; }
                    }
                }
                catch { }
            }
        }
    }

    // ---------------------------------------------------------------- Self-update
    static class Updater
    {
        static string state = "idle", error = "", file = ""; static long got = 0, total = 0;
        public static string Status() { return App.Json.Serialize(new Dictionary<string, object> { { "state", state }, { "error", error }, { "got", got }, { "total", total }, { "file", file } }); }
        public static string Start(string url, string name, string mode)
        {
            if (!url.StartsWith("https://")) return "{\"ok\":false,\"error\":\"blocked\"}";
            if (state == "downloading" || state == "installing") return "{\"ok\":true}";
            name = Path.GetFileName(string.IsNullOrEmpty(name) ? "OfficeOpenSource-update.exe" : name);
            state = "downloading"; error = ""; got = 0; total = 0;
            Thread t = new Thread(() => Run(url, name, mode)); t.IsBackground = true; t.Start();
            return "{\"ok\":true}";
        }
        static void Run(string url, string name, string mode)
        {
            try
            {
                string dir = Path.Combine(Path.GetTempPath(), "OfficeOpenSourceUpdate"); Directory.CreateDirectory(dir);
                file = Path.Combine(dir, name);
                CookieContainer cookies = new CookieContainer();
                HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url);
                r.CookieContainer = cookies;
                r.Timeout = 60000; r.ReadWriteTimeout = 120000; r.AllowAutoRedirect = true;
                r.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) YomiOfficeOpenSource/" + App.Version;
                WebResponse resp = r.GetResponse();
                if ((resp.ContentType ?? "").StartsWith("text/html"))
                {
                    string html;
                    using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) html = sr.ReadToEnd();
                    resp.Close();
                    var m = System.Text.RegularExpressions.Regex.Match(html, "href=\"([^\"]*confirm=[^\"]*)\"");
                    if (!m.Success) m = System.Text.RegularExpressions.Regex.Match(html, "action=\"([^\"]+)\"");
                    if (!m.Success) m = System.Text.RegularExpressions.Regex.Match(html, "href=\"(https?://download[^\"]+mediafire\\.com/[^\"]+)\"");
                    if (m.Success)
                    {
                        string newUrl = System.Net.WebUtility.HtmlDecode(m.Groups[1].Value);
                        if (newUrl.StartsWith("/")) newUrl = new Uri(new Uri(url), newUrl).ToString();
                        r = (HttpWebRequest)WebRequest.Create(newUrl);
                        r.CookieContainer = cookies;
                        r.Timeout = 60000; r.ReadWriteTimeout = 120000; r.AllowAutoRedirect = true;
                        r.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) YomiOfficeOpenSource/" + App.Version;
                        resp = r.GetResponse();
                    }
                }
                using (resp)
                {
                    total = resp.ContentLength;
                    using (Stream s = resp.GetResponseStream()) using (FileStream fs = File.Create(file)) { byte[] buf = new byte[65536]; int n; while ((n = s.Read(buf, 0, buf.Length)) > 0) { fs.Write(buf, 0, n); got += n; } }
                }
                state = "installing";
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".exe")
                {
                    Process.Start(new ProcessStartInfo(file, "/S /UPDATE") { UseShellExecute = true });
                }
                else if (ext == ".msi")
                {
                    string ps = Path.Combine(Path.GetDirectoryName(file), "apply_msi_update.ps1");
                    string exeCur = Path.Combine(App.Dir, "OfficeOpenSource.exe");
                    string exeAlt = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\YomiOfficeOpenSource\OfficeOpenSource.exe");
                    string script = "Start-Sleep 2; Get-Process OfficeOpenSource -EA SilentlyContinue | Stop-Process -Force; Start-Sleep 1;\r\n" +
                        "$p = Start-Process 'msiexec.exe' -ArgumentList '/i \"" + file.Replace("\"", "\\\"") + "\" /passive' -Wait -PassThru;\r\n" +
                        "$exe = if (Test-Path '" + exeCur.Replace("'", "''") + "') { '" + exeCur.Replace("'", "''") + "' } elseif (Test-Path '" + exeAlt.Replace("'", "''") + "') { '" + exeAlt.Replace("'", "''") + "' } else { $null };\r\n" +
                        "if ($exe) { Start-Process $exe }";
                    File.WriteAllText(ps, script, Encoding.UTF8);
                    Process.Start(new ProcessStartInfo("powershell.exe", "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + ps + "\"") { UseShellExecute = false, CreateNoWindow = true });
                }
                else if (ext == ".zip")
                {
                    string ps = Path.Combine(Path.GetDirectoryName(file), "apply_update.ps1");
                    string script = "Start-Sleep 3; Get-Process OfficeOpenSource -EA SilentlyContinue | Stop-Process -Force; Start-Sleep 1;" +
                        "$sh = New-Object -ComObject Shell.Application; $src = $sh.NameSpace('" + file.Replace("'", "''") + "'); $dst = $sh.NameSpace('" + App.Dir.Replace("'", "''") + "');" +
                        "$dst.CopyHere($src.Items(), 0x614); Start-Sleep 2; Start-Process '" + Path.Combine(App.Dir, "OfficeOpenSource.exe").Replace("'", "''") + "'";
                    File.WriteAllText(ps, script, Encoding.UTF8);
                    Process.Start(new ProcessStartInfo("powershell.exe", "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + ps + "\"") { UseShellExecute = false, CreateNoWindow = true });
                }
                else Process.Start(new ProcessStartInfo(file) { UseShellExecute = true });
                state = "done";
                if (ext != ".appxbundle" && ext != ".appx" && ext != ".msix") { Thread.Sleep(1500); App.LastPing = DateTime.UtcNow.AddSeconds(-60); }
            }
            catch (Exception e) { state = "error"; error = e.Message; }
        }
    }

    // ---------------------------------------------------------------- Per-user registration / uninstall (no admin rights)
    static class Installer
    {
        public static readonly string[] Exts = { ".docx", ".docm", ".dotx", ".xlsx", ".xlsm", ".xltx", ".pptx", ".pptm", ".ppsx", ".potx", ".pdf", ".csv", ".md", ".odt", ".ods", ".odp", ".rtf", ".docm", ".xlsm", ".pptm", ".dotx", ".xltx", ".potx", ".ppsx", ".doc", ".xls", ".ppt", ".tex", ".latex", ".mml", ".eqx", ".mathtype" };
        public static void RegisterAssociations(string exe)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(@"Software\Classes\OfficeOpenSource.Document"))
            {
                k.SetValue("", "Tài liệu Yomi-Office Open Source");
                using (RegistryKey i = k.CreateSubKey("DefaultIcon")) i.SetValue("", "\"" + exe + "\",0");
                using (RegistryKey c = k.CreateSubKey(@"shell\open\command")) c.SetValue("", "\"" + exe + "\" \"%1\"");
            }
            foreach (string e in Exts) using (RegistryKey k = Registry.CurrentUser.CreateSubKey(@"Software\Classes\" + e + @"\OpenWithProgids")) k.SetValue("OfficeOpenSource.Document", new byte[0], RegistryValueKind.None);
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\OfficeOpenSource.exe")) k.SetValue("", exe);
        }
        public static void Uninstall()
        {
            if (MessageBox.Show("Gỡ cài đặt Yomi-Office Open Source?\n(Tài liệu của bạn không bị xóa.)", "Gỡ cài đặt", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\OfficeOpenSource.Document"); } catch { }
            foreach (string e in Exts) try { using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Classes\" + e + @"\OpenWithProgids", true)) if (k != null) k.DeleteValue("OfficeOpenSource.Document", false); } catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\App Paths\OfficeOpenSource.exe"); } catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeOpenSource"); } catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\YomiOfficeOpenSource"); } catch { }
            foreach (string f in new[] {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Office Open Source.lnk"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "Office Open Source.lnk"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Yomi-Office Open Source.lnk"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "Yomi-Office Open Source.lnk")
            })
                try { File.Delete(f); } catch { }
            string cmd = "/c timeout /t 2 /nobreak >nul & rmdir /s /q \"" + App.Dir + "\"";
            Process.Start(new ProcessStartInfo("cmd.exe", cmd) { WindowStyle = ProcessWindowStyle.Hidden, CreateNoWindow = true, WorkingDirectory = Path.GetTempPath() });
            MessageBox.Show("Đã gỡ cài đặt Yomi-Office Open Source.", "Gỡ cài đặt", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
