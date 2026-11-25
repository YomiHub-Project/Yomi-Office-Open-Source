// Office Open Source v0.0.8 - Setup (Windows 7+, .NET Framework 4.0+, x86/x64/ARM64)
// Per-user install (no admin), Vietnamese UI, DPI-aware fonts, silent mode: Setup.exe /S [/D=C:\path]
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Yomi-Office Open Source 0.1.8 Setup")]
[assembly: AssemblyProduct("Yomi-Office Open Source")]
[assembly: AssemblyVersion("0.1.8.0")]
[assembly: AssemblyFileVersion("0.1.8.0")]
[assembly: AssemblyCompany("Yomi-Office Open Source Project")]

namespace OfficeOpenSourceSetup
{
    static class Program
    {
        public const string Ver = "0.1.8";
        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
        public static string DefaultDir { get { return Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs"), "YomiOfficeOpenSource"); } }

        [STAThread]
        static int Main(string[] args)
        {
            try { if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware(); } catch { }
            bool silent = false, relaunch = false; string dir = null;
            foreach (string a in args)
            {
                if (a.Equals("/S", StringComparison.OrdinalIgnoreCase) || a.Equals("/silent", StringComparison.OrdinalIgnoreCase)) silent = true;
                else if (a.Equals("/UPDATE", StringComparison.OrdinalIgnoreCase)) relaunch = true; // v0.0.7 self-update: start the app again afterwards
                else if (a.StartsWith("/D=", StringComparison.OrdinalIgnoreCase)) dir = a.Substring(3).Trim('"');
            }
            if (silent)
            {
                try { if (relaunch) System.Threading.Thread.Sleep(2500); string d = dir ?? ExistingDir() ?? DefaultDir; Install(d, true, true, true, null); if (relaunch) try { Process.Start(Path.Combine(Path.GetFullPath(d), "OfficeOpenSource.exe")); } catch { } return 0; }
                catch (Exception ex) { try { File.WriteAllText(Path.Combine(Path.GetTempPath(), "OfficeOpenSource_setup_error.txt"), ex.ToString(), Encoding.UTF8); } catch { } return 1; }
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SetupForm());
            return 0;
        }

        public static string ExistingDir()
        {
            foreach (string sub in new[] { @"Software\Microsoft\Windows\CurrentVersion\Uninstall\YomiOfficeOpenSource", @"Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeOpenSource" })
                try { using (RegistryKey k = Registry.CurrentUser.OpenSubKey(sub)) { if (k != null) { string d = k.GetValue("InstallLocation") as string; if (!string.IsNullOrEmpty(d)) return d; } } } catch { }
            return null;
        }
        public static string ExistingVersion()
        {
            foreach (string sub in new[] { @"Software\Microsoft\Windows\CurrentVersion\Uninstall\YomiOfficeOpenSource", @"Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeOpenSource" })
                try { using (RegistryKey k = Registry.CurrentUser.OpenSubKey(sub)) { if (k != null) return k.GetValue("DisplayVersion") as string; } } catch { }
            return null;
        }

        // payload: gzip( repeat { int32 pathLen, utf8 path, int64 size, bytes } )
        public static void Install(string dir, bool desktop, bool startMenu, bool assoc, Action<int, string> progress)
        {
            dir = Path.GetFullPath(dir);
            // close a running copy from that folder (upgrade)
            foreach (Process p in Process.GetProcessesByName("OfficeOpenSource")) { try { if (p.MainModule.FileName.StartsWith(dir, StringComparison.OrdinalIgnoreCase)) { p.Kill(); p.WaitForExit(3000); } } catch { } }
            string appDir = Path.Combine(dir, "app");
            if (Directory.Exists(appDir)) try { Directory.Delete(appDir, true); } catch { }
            Directory.CreateDirectory(dir);
            Stream res = Assembly.GetExecutingAssembly().GetManifestResourceStream("Payload");
            long total = res.Length; int count = 0;
            using (GZipStream gz = new GZipStream(res, CompressionMode.Decompress))
            using (BinaryReader br = new BinaryReader(gz))
            {
                byte[] buf = new byte[81920];
                while (true)
                {
                    int plen; try { plen = br.ReadInt32(); } catch (EndOfStreamException) { break; }
                    if (plen <= 0 || plen > 4096) break;
                    string rel = Encoding.UTF8.GetString(br.ReadBytes(plen));
                    long size = br.ReadInt64();
                    string full = Path.GetFullPath(Path.Combine(dir, rel));
                    if (!full.StartsWith(dir, StringComparison.OrdinalIgnoreCase)) throw new Exception("Đường dẫn không hợp lệ trong gói: " + rel);
                    Directory.CreateDirectory(Path.GetDirectoryName(full));
                    using (FileStream fs = File.Create(full)) { long left = size; while (left > 0) { int n = br.Read(buf, 0, (int)Math.Min(buf.Length, left)); if (n <= 0) throw new Exception("Gói cài đặt bị hỏng"); fs.Write(buf, 0, n); left -= n; } }
                    count++;
                    if (progress != null) progress((int)Math.Min(95, 5 + res.Position * 90 / Math.Max(1, total)), rel);
                }
            }
            if (count == 0) throw new Exception("Gói cài đặt trống");
            try { File.Delete(Path.Combine(dir, "installed_by_msi.txt")); } catch { }
            string exe = Path.Combine(dir, "OfficeOpenSource.exe");
            foreach (string oldLnk in new[] { Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Office Open Source.lnk"), Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "Office Open Source.lnk") })
                try { if (File.Exists(oldLnk)) File.Delete(oldLnk); } catch { }
            if (desktop) Shortcut(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Yomi-Office Open Source.lnk"), exe, dir);
            if (startMenu) Shortcut(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "Yomi-Office Open Source.lnk"), exe, dir);
            if (assoc) try { Process.Start(new ProcessStartInfo(exe, "--register") { UseShellExecute = false }).WaitForExit(10000); } catch { }
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeOpenSource"))
            {
                k.SetValue("DisplayName", "Yomi-Office Open Source");
                k.SetValue("DisplayVersion", Ver);
                k.SetValue("Publisher", "Yomi-Office Open Source Project");
                k.SetValue("DisplayIcon", exe + ",0");
                k.SetValue("InstallLocation", dir);
                k.SetValue("UninstallString", "\"" + exe + "\" --uninstall");
                k.SetValue("NoModify", 1, RegistryValueKind.DWord); k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                k.SetValue("EstimatedSize", (int)(DirSize(new DirectoryInfo(dir)) / 1024), RegistryValueKind.DWord);
            }
            if (progress != null) progress(100, "Hoàn tất");
        }
        static long DirSize(DirectoryInfo d) { long s = 0; foreach (FileInfo f in d.GetFiles("*", SearchOption.AllDirectories)) s += f.Length; return s; }
        static void Shortcut(string lnk, string target, string workDir)
        {
            try
            {
                Type t = Type.GetTypeFromProgID("WScript.Shell"); object sh = Activator.CreateInstance(t);
                object sc = t.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, sh, new object[] { lnk });
                Type st = sc.GetType();
                st.InvokeMember("TargetPath", BindingFlags.SetProperty, null, sc, new object[] { target });
                st.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, sc, new object[] { workDir });
                st.InvokeMember("IconLocation", BindingFlags.SetProperty, null, sc, new object[] { target + ",0" });
                st.InvokeMember("Description", BindingFlags.SetProperty, null, sc, new object[] { "Yomi-Office Open Source - Word, Excel, PowerPoint, PDF" });
                st.InvokeMember("Save", BindingFlags.InvokeMethod, null, sc, null);
            }
            catch { }
        }
    }

    class SetupForm : Form
    {
        static readonly Color Brand = Color.FromArgb(24, 90, 189);
        TextBox path; CheckBox cDesk, cStart, cAssoc, cRun; Button bInstall, bBrowse, bCancel; ProgressBar bar; Label status, title, sub;
        Panel page1, page2;

        static Font UiFont(float size, FontStyle st)
        {
            // Segoe UI ships with Windows 7+; fall back to Tahoma (full Vietnamese coverage) if missing
            foreach (string n in new[] { "Segoe UI", "Tahoma", "Arial" }) { using (Font f = new Font(n, size, st, GraphicsUnit.Point)) { if (f.Name == n) return new Font(n, size, st, GraphicsUnit.Point); } }
            return new Font(FontFamily.GenericSansSerif, size, st);
        }

        public SetupForm()
        {
            AutoScaleMode = AutoScaleMode.Dpi; AutoScaleDimensions = new SizeF(96f, 96f);
            Font = UiFont(9.5f, FontStyle.Regular);
            Text = "Cài đặt Yomi-Office Open Source " + Program.Ver;
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(560, 380); BackColor = Color.White;
            try { Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location); } catch { }

            Panel head = new Panel { Dock = DockStyle.Top, Height = 92, BackColor = Brand };
            title = new Label { Text = "Yomi-Office Open Source", ForeColor = Color.White, Font = UiFont(18f, FontStyle.Bold), AutoSize = true, Location = new Point(24, 16), BackColor = Color.Transparent };
            string old = Program.ExistingVersion();
            sub = new Label { Text = "Phiên bản " + Program.Ver + " — Word · Excel · PowerPoint · PDF · A.I" + (old != null ? "   (nâng cấp từ " + old + ")" : ""), ForeColor = Color.FromArgb(225, 235, 255), AutoSize = true, Location = new Point(26, 56), BackColor = Color.Transparent };
            head.Controls.Add(title); head.Controls.Add(sub);

            page1 = new Panel { Location = new Point(0, 92), Size = new Size(560, 230) };
            Label l1 = new Label { Text = "Thư mục cài đặt (không cần quyền quản trị):", AutoSize = true, Location = new Point(24, 18) };
            path = new TextBox { Location = new Point(26, 42), Width = 410, Text = Program.ExistingDir() ?? Program.DefaultDir };
            bBrowse = new Button { Text = "Duyệt…", Location = new Point(444, 40), Width = 90, Height = 27 };
            bBrowse.Click += (s, e) => { using (FolderBrowserDialog d = new FolderBrowserDialog { Description = "Chọn thư mục cài đặt", SelectedPath = path.Text }) if (d.ShowDialog(this) == DialogResult.OK) path.Text = Path.Combine(d.SelectedPath, "YomiOfficeOpenSource"); };
            cDesk = new CheckBox { Text = "Tạo lối tắt trên màn hình Desktop", Checked = true, AutoSize = true, Location = new Point(26, 84) };
            cStart = new CheckBox { Text = "Thêm vào menu Start", Checked = true, AutoSize = true, Location = new Point(26, 110) };
            cAssoc = new CheckBox { Text = "Thêm vào \"Mở bằng\" cho .docx .xlsx .pptx .pdf .odt .docm .xlsm …", Checked = true, AutoSize = true, Location = new Point(26, 136) };
            cRun = new CheckBox { Text = "Mở Yomi-Office Open Source sau khi cài", Checked = true, AutoSize = true, Location = new Point(26, 162) };
            Label req = new Label { Text = "Yêu cầu: Windows 7 trở lên (x86, x64, ARM64) · Microsoft Edge hoặc Google Chrome · hoạt động ngoại tuyến", ForeColor = Color.Gray, AutoSize = true, Location = new Point(24, 198) };
            page1.Controls.AddRange(new Control[] { l1, path, bBrowse, cDesk, cStart, cAssoc, cRun, req });

            page2 = new Panel { Location = new Point(0, 92), Size = new Size(560, 230), Visible = false };
            status = new Label { Text = "Đang chuẩn bị…", AutoSize = false, Location = new Point(24, 60), Size = new Size(510, 40) };
            bar = new ProgressBar { Location = new Point(26, 104), Size = new Size(508, 22), Minimum = 0, Maximum = 100 };
            page2.Controls.Add(status); page2.Controls.Add(bar);

            bInstall = new Button { Text = Program.ExistingVersion() != null ? "Nâng cấp" : "Cài đặt", Location = new Point(344, 334), Size = new Size(96, 30), BackColor = Brand, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            bInstall.FlatAppearance.BorderSize = 0;
            bCancel = new Button { Text = "Hủy", Location = new Point(448, 334), Size = new Size(88, 30) };
            bCancel.Click += (s, e) => Close();
            bInstall.Click += (s, e) => DoInstall();
            Controls.AddRange(new Control[] { head, page1, page2, bInstall, bCancel });
            AcceptButton = bInstall; CancelButton = bCancel;
        }

        void DoInstall()
        {
            if (bInstall.Text == "Hoàn tất") { if (cRun.Checked) try { Process.Start(Path.Combine(path.Text, "OfficeOpenSource.exe")); } catch { } Close(); return; }
            string dir = path.Text.Trim(); bool d = cDesk.Checked, sm = cStart.Checked, a = cAssoc.Checked;
            page1.Visible = false; page2.Visible = true; bInstall.Enabled = false; bCancel.Enabled = false;
            Thread t = new Thread(() =>
            {
                try
                {
                    Program.Install(dir, d, sm, a, (p, f) => BeginInvoke((Action)(() => { bar.Value = p; status.Text = "Đang cài đặt: " + f; })));
                    BeginInvoke((Action)(() => { status.Text = "✔ Đã cài đặt Yomi-Office Open Source " + Program.Ver + " vào:\n" + dir; bar.Value = 100; bInstall.Text = "Hoàn tất"; bInstall.Enabled = true; }));
                }
                catch (Exception ex)
                {
                    BeginInvoke((Action)(() => { status.ForeColor = Color.Firebrick; status.Text = "Lỗi: " + ex.Message; bCancel.Enabled = true; bCancel.Text = "Đóng"; }));
                }
            });
            t.IsBackground = true; t.Start();
        }
    }
}
