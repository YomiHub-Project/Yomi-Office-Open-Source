// Office Open Source v0.1.6 - Native document engine for Windows (LibreOfficeKit, no WebView)
// Loads the bundled LibreOffice core (lo\program\sofficeapp.dll) in-process through its C ABI
// (LibreOfficeKit vtables, see include/LibreOfficeKit/LibreOfficeKit.h, MPL-2.0) and renders
// documents tile-by-tile into a double-buffered WinForms control. Editing goes through LOK key/mouse
// events and .uno: commands, so Word/Excel/PowerPoint files keep full Microsoft Office fidelity
// (layout, formulas, charts, VBA macros via LibreOffice's VBA compatibility layer).
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace OfficeOpenSource
{
    static class Utf8
    {
        public static byte[] Z(string s) { if (s == null) return null; byte[] b = Encoding.UTF8.GetBytes(s); byte[] r = new byte[b.Length + 1]; Buffer.BlockCopy(b, 0, r, 0, b.Length); return r; }
        public static string S(IntPtr p)
        {
            if (p == IntPtr.Zero) return null;
            int n = 0; while (Marshal.ReadByte(p, n) != 0) n++;
            byte[] b = new byte[n]; Marshal.Copy(p, b, 0, n); return Encoding.UTF8.GetString(b);
        }
    }

    /// <summary>LibreOfficeKit process-wide instance (one per process, like soffice).</summary>
    static class Lok
    {
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)] static extern IntPtr LoadLibraryW(string path);
        [DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true)] static extern IntPtr GetProcAddress(IntPtr h, string name);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)] static extern bool SetDllDirectoryW(string path);
        [DllImport("kernel32")] static extern IntPtr GetCommandLineW();
        [DllImport("kernel32")] static extern IntPtr GetCurrentProcess();
        [DllImport("kernel32")] static extern bool TerminateProcess(IntPtr h, uint code);
        /// <summary>LibreOffice's DLL teardown can block process exit; end the process directly (files are already saved).</summary>
        public static void HardExit(int code) { try { TerminateProcess(GetCurrentProcess(), (uint)code); } catch { Environment.Exit(code); } }
        [DllImport("kernel32")] static extern IntPtr GetCommandLineA();

        // LibreOffice's soffice_main() parses the *process* command line (GetCommandLineW). Our own switches (--convert, --native,
        // a file path) make it print help / open the file itself and leave its main loop, which deadlocks Calc (threaded xlsx
        // import yields on that thread). .NET has already parsed our arguments, so blank the process command line in place.
        static void HideProcessArguments()
        {
            try
            {
                string exe = "\"" + System.Windows.Forms.Application.ExecutablePath + "\"";
                IntPtr w = GetCommandLineW(); string cur = Marshal.PtrToStringUni(w) ?? "";
                if (cur.Length >= exe.Length) { char[] z = (exe + "\0").ToCharArray(); Marshal.Copy(z, 0, w, z.Length); }
                IntPtr a = GetCommandLineA(); string curA = Marshal.PtrToStringAnsi(a) ?? "";
                byte[] ab = Encoding.Default.GetBytes(exe + "\0"); if (curA.Length >= ab.Length - 1) Marshal.Copy(ab, 0, a, ab.Length);
                Trace("cmdline now: " + Marshal.PtrToStringUni(GetCommandLineW()));
            }
            catch (Exception e) { Trace("HideProcessArguments " + e.Message); }
        }

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate IntPtr Hook2(byte[] installPath, byte[] profileUrl);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate IntPtr F_LoadOpts(IntPtr self, byte[] url, byte[] opts);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate IntPtr F_GetError(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void F_FreeError(IntPtr p);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate IntPtr F_Str(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] public delegate void Callback(int type, IntPtr payload, IntPtr data);

        public static IntPtr Office = IntPtr.Zero;
        public static bool Unipoll = Environment.GetEnvironmentVariable("OOS_LOK_UNIPOLL") == "1";
        public static string ProgramDir, Error, VersionInfo;
        static readonly object Gate = new object();

        internal static T Slot<T>(IntPtr obj, int slot) where T : class
        {
            IntPtr cls = Marshal.ReadIntPtr(obj);
            long size = IntPtr.Size == 8 ? Marshal.ReadInt64(cls) : Marshal.ReadInt32(cls);
            if ((slot + 1) * IntPtr.Size > size) return null; // older LibreOffice without this entry
            IntPtr fn = Marshal.ReadIntPtr(cls, slot * IntPtr.Size);
            return fn == IntPtr.Zero ? null : (T)(object)Marshal.GetDelegateForFunctionPointer(fn, typeof(T));
        }

        public static string FindProgramDir()
        {
            string[] cands = {
                Path.Combine(App.Dir, "lo\\program"),
                Path.Combine(App.Dir, "LibreOffice\\program"),
                Path.Combine(App.Dir, "..\\lo\\program"),
            };
            foreach (string c in cands) if (File.Exists(Path.Combine(c, "sofficeapp.dll")) || File.Exists(Path.Combine(c, "mergedlo.dll"))) return Path.GetFullPath(c);
            string env = Environment.GetEnvironmentVariable("OOS_LO_PROGRAM");
            if (!string.IsNullOrEmpty(env) && Directory.Exists(env)) return env;
            return null;
        }

        public static bool Available { get { return FindProgramDir() != null; } }

        public static bool Init()
        {
            lock (Gate)
            {
                if (Office != IntPtr.Zero) return true;
                ProgramDir = FindProgramDir();
                if (ProgramDir == null) { Error = "Không tìm thấy engine LibreOffice (thư mục lo\\program)."; return false; }
                try
                {
                    HideProcessArguments();
                    SetDllDirectoryW(ProgramDir);
                    // unipoll: the calling (UI) thread becomes LibreOffice's VCL main thread, so Calc's threaded import /
                    // Application::Yield never waits on a thread that is not pumping messages (fixes Excel files hanging)
                    if (Unipoll) Environment.SetEnvironmentVariable("SAL_LOK_OPTIONS", "unipoll");
                    // several app windows/processes share one private profile; a stale .lock (crash, kill) must not stop the engine's main loop
                    Environment.SetEnvironmentVariable("SAL_NOLOCK_PROFILE", "1");
                    Environment.SetEnvironmentVariable("SAL_DISABLE_JAVA", "1");
                    // no Skia/Vulkan: tiles are rendered off-screen anyway, and creating a Vulkan device can block start-up for
                    // minutes on some GPU drivers (seen right after another engine process was terminated)
                    Environment.SetEnvironmentVariable("SAL_DISABLESKIA", "1");
                    Environment.SetEnvironmentVariable("SAL_SKIA", "raster");
                    Environment.SetEnvironmentVariable("PATH", ProgramDir + ";" + Path.GetFullPath(Path.Combine(ProgramDir, "..\\URE\\bin")) + ";" + Environment.GetEnvironmentVariable("PATH"));
                    EnsureClipboardResponsive();
                    Trace("init: loading engine");
                    IntPtr lib = LoadLibraryW(Path.Combine(ProgramDir, "sofficeapp.dll"));
                    if (lib == IntPtr.Zero) lib = LoadLibraryW(Path.Combine(ProgramDir, "mergedlo.dll"));
                    if (lib == IntPtr.Zero) { Error = "LoadLibrary lỗi " + Marshal.GetLastWin32Error(); return false; }
                    IntPtr sym = GetProcAddress(lib, "libreofficekit_hook_2");
                    if (sym == IntPtr.Zero) { Error = "Thiếu libreofficekit_hook_2"; return false; }
                    Hook2 hook = (Hook2)Marshal.GetDelegateForFunctionPointer(sym, typeof(Hook2));
                    // separate user profile for the embedded engine (portable builds keep it next to the exe)
                    string prof = File.Exists(Path.Combine(App.Dir, "portable.flag")) ? Path.Combine(App.Dir, "data\\lo-profile")
                        : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOpenSource\\lo-profile");
                    // one profile per engine version (24.8 for Windows 7/8.1 builds, 26.8 for Windows 10/11): profiles are not downgrade-safe
                    try { string vi = File.ReadAllText(Path.Combine(ProgramDir, "version.ini")); int k = vi.IndexOf("MsiProductVersion="); prof += "-" + (k >= 0 ? vi.Substring(k + 18, 4).Trim() : "x"); } catch { }
                    prof = ClaimProfile(prof);
                    Directory.CreateDirectory(prof);
                    // LibreOffice's own online updater (MAR) must never run inside Office Open Source: it downloads LibreOffice
                    // updates into the profile and, on the next start, launches updater.exe and waits for it - start-up then hangs
                    // (and it would try to patch the embedded engine). Drop any pending update before starting the engine.
                    try { string upd = Path.Combine(prof, "updates"); if (Directory.Exists(upd)) { Directory.Delete(upd, true); Trace("removed pending LibreOffice update in profile"); } } catch (Exception ue) { Trace("updates cleanup " + ue.Message); }
                    // v0.1.7: a crash in an earlier session leaves crash\dump.ini behind; on the next start LibreOffice then opens its
                    // (invisible) "Crash Report" dialog as soon as Writer lays out a picture and every .docx with images hangs forever
                    try { string cr = Path.Combine(prof, "crash"); if (Directory.Exists(cr)) { foreach (string cf in Directory.GetFiles(cr)) try { File.Delete(cf); } catch { } Trace("cleared crash reports in profile"); } } catch (Exception ce) { Trace("crash cleanup " + ce.Message); }
                    WriteConfig(Path.Combine(prof, "user")); BasicTools.Install(Path.Combine(prof, "user"), Path.Combine(Path.GetDirectoryName(ProgramDir.TrimEnd('\\')), "presets\\basic"));
                    Trace("init: engine loaded, starting LibreOfficeKit with profile " + prof);
                    Office = hook(Utf8.Z(ProgramDir), Utf8.Z(new Uri(prof).AbsoluteUri));
                    Trace("init: LibreOfficeKit ready");
                    if (Office == IntPtr.Zero) { Error = "LibreOfficeKit không khởi tạo được"; return false; }
                    F_Str ver = Slot<F_Str>(Office, 10);
                    if (ver != null) VersionInfo = Utf8.S(ver(Office));
                    return true;
                }
                catch (Exception e) { Error = e.Message; return false; }
            }
        }

        // One LibreOffice user profile must never be used by two live processes (or by a new process while a killed one is
        // still tearing down): LibreOffice can then block forever during start-up. Each process claims a profile with a
        // named mutex (released automatically when the owner dies) and falls back to a sibling profile when it is busy.
        [DllImport("ole32")] static extern int OleGetClipboard(out IntPtr dataObj);
        [DllImport("ole32")] static extern int OleInitialize(IntPtr r);
        [DllImport("user32")] static extern bool OpenClipboard(IntPtr h);
        [DllImport("user32")] static extern bool EmptyClipboard();
        [DllImport("user32")] static extern bool CloseClipboard();
        /// <summary>LibreOffice queries the OLE clipboard while starting. If the clipboard owner was a process that died
        /// (e.g. Word/Excel killed), that call can block for minutes. Probe it first and reset a dead clipboard.</summary>
        static void EnsureClipboardResponsive()
        {
            try
            {
                bool done = false;
                Thread t = new Thread(() => { try { OleInitialize(IntPtr.Zero); IntPtr o; if (OleGetClipboard(out o) == 0 && o != IntPtr.Zero) Marshal.Release(o); } catch { } done = true; }) { IsBackground = true };
                t.SetApartmentState(ApartmentState.STA); t.Start();
                if (!t.Join(3000) && !done)
                {
                    Trace("clipboard owner not responding - clearing the clipboard before starting the engine");
                    for (int i = 0; i < 10; i++) { if (OpenClipboard(IntPtr.Zero)) { EmptyClipboard(); CloseClipboard(); break; } Thread.Sleep(100); }
                }
            }
            catch (Exception e) { Trace("clipboard probe " + e.Message); }
        }

        static Mutex profileMutex;
        static string ClaimProfile(string prof)
        {
            for (int i = 0; i < 6; i++)
            {
                string cand = i == 0 ? prof : prof + "-" + (i + 1);
                string name = "Local\\OOS_LOPROFILE_" + (uint)cand.ToLowerInvariant().GetHashCode();
                Mutex m = new Mutex(false, name); bool got;
                try { got = m.WaitOne(0); } catch (AbandonedMutexException) { got = true; }
                if (got) { profileMutex = m; if (i > 0) Trace("profile busy, using " + cand); return cand; }
                m.Dispose();
            }
            string tmp = Path.Combine(Path.GetTempPath(), "OfficeOpenSource-lo-" + System.Diagnostics.Process.GetCurrentProcess().Id); Trace("all profiles busy, using " + tmp); return tmp;
        }

        /// <summary>true when the user allowed document macros (Xem lại &amp; Macro → Cho phép macro); stored next to the profile.</summary>
        public static bool MacrosAllowed
        {
            get { try { return File.Exists(MacroFlag); } catch { return false; } }
            set { try { if (value) File.WriteAllText(MacroFlag, DateTime.Now.ToString("o")); else if (File.Exists(MacroFlag)) File.Delete(MacroFlag); } catch { } }
        }
        static string MacroFlag { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOpenSource\\lok-macros.flag"); } }

        // Embedded engine must never block on a modal question nobody can see:
        //  * external links in Excel files are not refreshed on load (cached values are shown, like Excel's "Enable Content" bar)
        //  * macro security: unsigned document macros only run when the user explicitly allowed macros (security level 3 +
        //    load option); macro execution itself stays enabled so the built-in "OOS" application library can run
        static void WriteConfig(string userDir)
        {
            try
            {
                Directory.CreateDirectory(userDir);
                string f = Path.Combine(userDir, "registrymodifications.xcu");
                string items =
                    "<item oor:path=\"/org.openoffice.Office.Calc/Content/Update\"><prop oor:name=\"Link\" oor:op=\"fuse\"><value>1</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Writer/Content/Update\"><prop oor:name=\"Link\" oor:op=\"fuse\"><value>1</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Common/Security/Scripting\"><prop oor:name=\"MacroSecurityLevel\" oor:op=\"fuse\"><value>" + (MacrosAllowed ? "0" : "3") + "</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Common/Security/Scripting\"><prop oor:name=\"DisableMacrosExecution\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Jobs/Jobs/org.openoffice.Office.Jobs:Job['UpdateCheck']/Arguments\"><prop oor:name=\"AutoCheckEnabled\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Jobs/Jobs/org.openoffice.Office.Jobs:Job['UpdateCheck']/Arguments\"><prop oor:name=\"AutoDownloadEnabled\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Update/Update\"><prop oor:name=\"Enabled\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Calc/Formula/Load\"><prop oor:name=\"OOXMLRecalcMode\" oor:op=\"fuse\"><value>1</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Calc/Formula/Load\"><prop oor:name=\"ODFRecalcMode\" oor:op=\"fuse\"><value>1</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Calc/Layout/Line\"><prop oor:name=\"PageBreak\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Common/Misc\"><prop oor:name=\"FirstRun\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Setup/Office\"><prop oor:name=\"ooSetupInstCompleted\" oor:op=\"fuse\"><value>true</value></prop></item>\n" +
                    // no Document Recovery / crash / tip dialogs on LibreOffice's (invisible) main thread: they block it forever
                    "<item oor:path=\"/org.openoffice.Office.Recovery/RecoveryInfo\"><prop oor:name=\"Enabled\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Recovery/AutoSave\"><prop oor:name=\"Enabled\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Recovery/RecoveryInfo\"><prop oor:name=\"Crashed\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Common/Misc\"><prop oor:name=\"ShowTipOfTheDay\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    "<item oor:path=\"/org.openoffice.Office.Common/Misc\"><prop oor:name=\"CrashReport\" oor:op=\"fuse\"><value>false</value></prop></item>\n" +
                    // area around the pages: light grey like Word instead of the black the tiles get on Windows
                    "<item oor:path=\"/org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice.Office.UI:ColorScheme['LibreOffice']/AppBackground\"><prop oor:name=\"Color\" oor:op=\"fuse\"><value>15132390</value></prop></item>\n";
                string body = File.Exists(f) ? File.ReadAllText(f) : "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<oor:items xmlns:oor=\"http://openoffice.org/2001/registry\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n</oor:items>\n";
                foreach (string key in new[] { "AppBackground\"><prop oor:name=\"Color", "RecoveryInfo\"><prop oor:name=\"Enabled", "AutoSave\"><prop oor:name=\"Enabled", "RecoveryInfo\"><prop oor:name=\"Crashed", "Misc\"><prop oor:name=\"ShowTipOfTheDay", "Misc\"><prop oor:name=\"CrashReport", "Calc/Content/Update\"><prop oor:name=\"Link", "Writer/Content/Update\"><prop oor:name=\"Link", "Scripting\"><prop oor:name=\"MacroSecurityLevel", "Scripting\"><prop oor:name=\"DisableMacrosExecution", "Misc\"><prop oor:name=\"FirstRun", "Line\"><prop oor:name=\"PageBreak", "Load\"><prop oor:name=\"OOXMLRecalcMode", "Arguments\"><prop oor:name=\"AutoCheckEnabled", "Arguments\"><prop oor:name=\"AutoDownloadEnabled", "Update\"><prop oor:name=\"Enabled", "Load\"><prop oor:name=\"ODFRecalcMode", "Office\"><prop oor:name=\"ooSetupInstCompleted" })
                {
                    int i = body.IndexOf(key); if (i < 0) continue;
                    int a = body.LastIndexOf("<item", i), b = body.IndexOf("</item>", i); if (a >= 0 && b > a) body = body.Remove(a, b + 7 - a);
                }
                body = System.Text.RegularExpressions.Regex.Replace(body, "<item oor:path=\"/org.openoffice.Office.Recovery/RecoveryList[^\"]*\">.*?</item>\\s*", "", System.Text.RegularExpressions.RegexOptions.Singleline);
                try { string bk = Path.Combine(userDir, "backup"); if (Directory.Exists(bk)) Directory.Delete(bk, true); } catch { }
                body = body.Replace("</oor:items>", items + "</oor:items>");
                File.WriteAllText(f, body, new UTF8Encoding(false));
            }
            catch { }
        }

        // ---- unipoll main loop: LibreOffice's Application::Execute runs on the UI thread and pumps every Win32 message
        //      (WinForms windows included). Required on Windows: Calc's threaded xlsx import yields on the VCL main thread.
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate int PollCb(IntPtr data, int timeoutUs);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void WakeCb(IntPtr data);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void RunLoopF(IntPtr self, PollCb poll, WakeCb wake, IntPtr data);
        static PollCb keepPoll = (d, t) => { System.Threading.Thread.Sleep(Math.Max(0, Math.Min(t / 1000, 50))); return 0; };
        static WakeCb keepWake = d => { };
        public static bool LoopRunning;
        static IntPtr activeDoc = IntPtr.Zero;
        /// <summary>.uno: commands go to the desktop's *active frame*; LibreOffice's (hidden) frame windows never get Win32 activation
        /// here, so make the document's frame active by creating a throw-away view and switching back (MakeActive_Impl).</summary>
        public static void Activate(LokDoc d)
        {
            if (d == null || d.H == activeDoc) return;
            try { int v = d.GetView(); int nv = d.CreateView(); if (nv >= 0) { d.SetView(v); d.DestroyView(nv); } activeDoc = d.H; }
            catch (Exception e) { Trace("Activate " + e.Message); }
        }
        public static void Forget(LokDoc d) { if (d != null && d.H == activeDoc) activeDoc = IntPtr.Zero; }
        static readonly List<Action> pendingInLoop = new List<Action>();
        /// <summary>Runs <paramref name="then"/> on the UI thread inside LibreOffice's main loop (starting the loop on first use).</summary>
        public static void InLoop(Action then)
        {
            if (!Unipoll || LoopRunning) { then(); return; }
            pendingInLoop.Add(then);
            if (pendingInLoop.Count > 1) return;
            System.Windows.Forms.Timer t = new System.Windows.Forms.Timer { Interval = 20 };
            t.Tick += (s, e) => { t.Stop(); t.Dispose(); List<Action> a = new List<Action>(pendingInLoop); pendingInLoop.Clear(); foreach (Action x in a) { try { x(); } catch (Exception ex) { Trace("InLoop " + ex); } } };
            t.Start();
            LoopRunning = true;
            RunLoopF run = Slot<RunLoopF>(Office, 13);
            Trace("runLoop enter");
            run(Office, keepPoll, keepWake, IntPtr.Zero);   // returns only when LibreOffice quits
            Trace("runLoop exit");
            Environment.Exit(0);
        }

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate int F_RunMacro(IntPtr self, byte[] url);
        /// <summary>Runs a Basic macro (macro:///Library.Module.Sub) through LibreOfficeKit's runMacro (postUnoCommand only takes .uno: commands).</summary>
        public static bool RunMacro(string url)
        {
            F_RunMacro f = Slot<F_RunMacro>(Office, 11); if (f == null) { Trace("runMacro not available"); return false; }
            bool ok = f(Office, Utf8.Z(url)) != 0; if (!ok) Trace("runMacro " + url + " failed: " + LastError()); return ok;
        }
        public static string LastError()
        {
            if (Office == IntPtr.Zero) return Error;
            F_GetError g = Slot<F_GetError>(Office, 3); if (g == null) return null;
            IntPtr p = g(Office); string s = Utf8.S(p);
            F_FreeError f = Slot<F_FreeError>(Office, 5); if (f != null && p != IntPtr.Zero) f(p);
            return s;
        }

        public static LokDoc Load(string path, bool macros)
        {
            if (!Init()) return null;
            F_LoadOpts load = Slot<F_LoadOpts>(Office, 4);
            // Batch (LOKSilent): load-time questions are cancelled instead of waiting forever on an invisible dialog. A few Excel
            // workbooks then refuse to load ("empty reference"); those fall back to the built-in web Excel editor.
            string opts = "Language=vi-VN" + (macros ? ",EnableMacrosExecution=true" : "") + (Environment.GetEnvironmentVariable("OOS_LOK_INTERACTIVE") == "1" ? "" : ",Batch=true");
            LastReadOnly = false;
            ClearStaleLock(path);
            IntPtr d = load(Office, Utf8.Z(new Uri(path).AbsoluteUri), Utf8.Z(opts));
            if (d == IntPtr.Zero)
            {
                // The file is locked (Microsoft Office has it open, or a lock left by a crashed app): in Batch mode LibreOffice
                // cancels its "open read-only?" question and returns nothing. Open a private copy instead, like Office's read-only mode.
                Trace("load failed (" + LastError() + "), retry from a private copy: " + path);
                try
                {
                    string dir = Path.Combine(Path.Combine(Path.GetTempPath(), "OfficeOpenSource-ro"), Guid.NewGuid().ToString("N").Substring(0, 8));
                    Directory.CreateDirectory(dir); string copy = Path.Combine(dir, Path.GetFileName(path));
                    using (FileStream src = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) using (FileStream dst = File.Create(copy)) src.CopyTo(dst);
                    d = load(Office, Utf8.Z(new Uri(copy).AbsoluteUri), Utf8.Z(opts));
                    if (d != IntPtr.Zero) LastReadOnly = true;
                }
                catch (Exception e) { Trace("copy retry " + e.Message); }
            }
            if (d == IntPtr.Zero) { Error = LastError() ?? "Không mở được tài liệu"; return null; }
            return new LokDoc(d);
        }

        /// <summary>True when the last Load had to use a private copy (original locked) - saving then asks for a new name.</summary>
        public static bool LastReadOnly;

        /// <summary>Removes a LibreOffice ".~lock.name#" file left behind by a crashed/killed process (no LibreOffice-based process running).</summary>
        static void ClearStaleLock(string path)
        {
            try
            {
                string lk = Path.Combine(Path.GetDirectoryName(path), ".~lock." + Path.GetFileName(path) + "#");
                if (!File.Exists(lk)) return;
                int me = System.Diagnostics.Process.GetCurrentProcess().Id, others = 0;
                foreach (string n in new[] { "OfficeOpenSource", "soffice", "soffice.bin" })
                    foreach (System.Diagnostics.Process pr in System.Diagnostics.Process.GetProcessesByName(n)) if (pr.Id != me) others++;
                if (others == 0 || File.GetLastWriteTimeUtc(lk) < DateTime.UtcNow.AddHours(-12)) { File.Delete(lk); Trace("stale lock removed " + lk); }
            }
            catch (Exception e) { Trace("ClearStaleLock " + e.Message); }
        }

        /// <summary>Headless conversion (docx->pdf, xlsx->ods, ...): used by /api/lok/convert and the PDF tools.</summary>
        public static void Trace(string m) { try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "oos_lok_trace.log"), DateTime.Now.ToString("HH:mm:ss.fff") + " [" + Thread.CurrentThread.ManagedThreadId + "] " + m + "\r\n"); } catch { } }
        public static bool Convert(string input, string output, string filterOpts)
        {
            Trace("convert " + input);
            if (!Init()) { Trace("init failed " + Error); return false; }
            Trace("init ok " + VersionInfo);
            LokDoc d = Load(input, false); Trace("loaded " + (d != null) + (d == null ? " error=" + Error : "")); if (d == null) return false;
            try { Trace("type=" + d.Type + " parts=" + d.Parts); bool ok = d.SaveAs(output, Path.GetExtension(output).TrimStart('.').ToLowerInvariant(), filterOpts); Trace("saved " + ok); return ok; }
            finally { d.Destroy(); Trace("destroyed"); }
        }
    }


    /// <summary>Win32 helpers for LibreOffice's own (hidden) frame windows on the UI thread.</summary>
    static class LokWin
    {
        delegate bool EnumProc(IntPtr h, IntPtr l);
        [DllImport("user32")] static extern bool EnumThreadWindows(uint tid, EnumProc p, IntPtr l);
        [DllImport("user32")] static extern bool EnumChildWindows(IntPtr parent, EnumProc p, IntPtr l);
        [DllImport("user32", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
        [DllImport("user32")] static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32")] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
        [DllImport("kernel32")] static extern uint GetCurrentThreadId();
        public static List<KeyValuePair<IntPtr, string>> SalWindows()
        {
            List<KeyValuePair<IntPtr, string>> r = new List<KeyValuePair<IntPtr, string>>();
            EnumThreadWindows(GetCurrentThreadId(), (h, l) =>
            {
                StringBuilder sb = new StringBuilder(128); GetClassName(h, sb, 128); string c = sb.ToString();
                if (c.StartsWith("SAL", StringComparison.OrdinalIgnoreCase)) r.Add(new KeyValuePair<IntPtr, string>(h, c + (IsWindowVisible(h) ? " visible" : " hidden")));
                return true;
            }, IntPtr.Zero);
            return r;
        }
        /// <summary>Tell LibreOffice its document frame window is the active one, so .uno: commands find an active frame.</summary>
        public static void ActivateFrames()
        {
            foreach (KeyValuePair<IntPtr, string> w in SalWindows())
                if (w.Value.StartsWith("SALFRAME", StringComparison.OrdinalIgnoreCase))
                {
                    SendMessage(w.Key, 0x0086, (IntPtr)1, IntPtr.Zero);   // WM_NCACTIVATE
                    SendMessage(w.Key, 0x0006, (IntPtr)1, IntPtr.Zero);   // WM_ACTIVATE(WA_ACTIVE)
                    SendMessage(w.Key, 0x0007, IntPtr.Zero, IntPtr.Zero); // WM_SETFOCUS
                }
        }
    }

    /// <summary>A loaded LibreOfficeKit document (thin typed wrapper over the vtable).</summary>
    class LokDoc
    {
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void V(IntPtr s);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate int I(IntPtr s);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void VI(IntPtr s, int a);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate IntPtr SI(IntPtr s, int a);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate int SaveAsF(IntPtr s, byte[] url, byte[] fmt, byte[] opts);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void PaintF(IntPtr s, IntPtr buf, int cw, int ch, int x, int y, int w, int h);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void SizeF32(IntPtr s, out int w, out int h);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void SizeF64(IntPtr s, out long w, out long h);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void StrF(IntPtr s, byte[] a);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void RegF(IntPtr s, Lok.Callback cb, IntPtr data);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void KeyF(IntPtr s, int type, int ch, int key);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void MouseF(IntPtr s, int type, int x, int y, int count, int buttons, int mods);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void UnoF(IntPtr s, byte[] cmd, byte[] args, [MarshalAs(UnmanagedType.I1)] bool notify);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void SelF(IntPtr s, int type, int x, int y);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate IntPtr GetSelF(IntPtr s, byte[] mime, out IntPtr usedMime);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] [return: MarshalAs(UnmanagedType.I1)] delegate bool PasteF(IntPtr s, byte[] mime, byte[] data, UIntPtr size);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate IntPtr CmdValF(IntPtr s, byte[] cmd);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void ZoomF(IntPtr s, int tpw, int tph, int ttw, int tth);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void AreaF(IntPtr s, int x, int y, int w, int h);

        public readonly IntPtr H;
        Lok.Callback keep; // keep the delegate alive for native code
        public const int TEXT = 0, SPREADSHEET = 1, PRESENTATION = 2, DRAWING = 3;
        public LokDoc(IntPtr h) { H = h; }

        public void Destroy() { V f = Lok.Slot<V>(H, 1); if (f != null) f(H); }
        public bool SaveAs(string path, string fmt, string opts)
        {
            SaveAsF f = Lok.Slot<SaveAsF>(H, 2);
            return f(H, Utf8.Z(new Uri(path).AbsoluteUri), Utf8.Z(fmt), Utf8.Z(opts)) != 0;
        }
        public int Type { get { I f = Lok.Slot<I>(H, 3); return f == null ? TEXT : f(H); } }
        public int Parts { get { I f = Lok.Slot<I>(H, 4); return f == null ? 1 : f(H); } }
        public int Part { get { I f = Lok.Slot<I>(H, 6); return f == null ? 0 : f(H); } set { VI f = Lok.Slot<VI>(H, 7); if (f != null) f(H, value); } }
        public string PartName(int i) { SI f = Lok.Slot<SI>(H, 8); return f == null ? ("Trang " + (i + 1)) : Utf8.S(f(H, i)); }
        public int TileMode { get { I f = Lok.Slot<I>(H, 11); return f == null ? 1 : f(H); } }
        public int PartMode { set { VI f = Lok.Slot<VI>(H, 9); if (f != null) f(H, value); } }
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate IntPtr RectsF(IntPtr s);
        /// <summary>Writer page rectangles in twips (getPartPageRectangles), cached until the document size changes.</summary>
        public List<Rectangle> PageRects()
        {
            if (Type != TEXT) return new List<Rectangle>();
            Size sz = SizeTwips; if (rectsFor == sz && rects != null) return rects;
            List<Rectangle> r = new List<Rectangle>();
            try { RectsF f = Lok.Slot<RectsF>(H, 5); string s = f == null ? null : Utf8.S(f(H)); if (!string.IsNullOrEmpty(s)) foreach (string part in s.Split(';')) { string[] a = part.Split(','); int x, y, w, h; if (a.Length >= 4 && int.TryParse(a[0].Trim(), out x) && int.TryParse(a[1].Trim(), out y) && int.TryParse(a[2].Trim(), out w) && int.TryParse(a[3].Trim(), out h)) r.Add(new Rectangle(x, y, w, h)); } }
            catch { }
            rects = r; rectsFor = sz; return r;
        }
        List<Rectangle> rects; Size rectsFor;
        public void Paint(IntPtr buf, int cw, int ch, int x, int y, int w, int h) { Lok.Slot<PaintF>(H, 10)(H, buf, cw, ch, x, y, w, h); }
        public Size SizeTwips
        {
            get
            {
                // getDocumentSize(long*, long*): long is 32-bit on Windows (LLP64)
                int w, h; Lok.Slot<SizeF32>(H, 12)(H, out w, out h); return new Size(w, h);
            }
        }
        public void InitRender(string args) { StrF f = Lok.Slot<StrF>(H, 13); if (f != null) f(H, Utf8.Z(args ?? "")); }
        public void Register(Lok.Callback cb) { keep = cb; RegF f = Lok.Slot<RegF>(H, 14); if (f != null) f(H, cb, IntPtr.Zero); }
        public void Key(int type, int ch, int key) { Lok.Slot<KeyF>(H, 15)(H, type, ch, key); }
        public void Mouse(int type, int x, int y, int count, int buttons, int mods) { Lok.Slot<MouseF>(H, 16)(H, type, x, y, count, buttons, mods); }
        public void Uno(string cmd, string args) { Lok.Slot<UnoF>(H, 17)(H, Utf8.Z(cmd), Utf8.Z(args ?? ""), false); }
        public void SetTextSelection(int type, int x, int y) { SelF f = Lok.Slot<SelF>(H, 18); if (f != null) f(H, type, x, y); }
        public string GetTextSelection(string mime)
        {
            GetSelF f = Lok.Slot<GetSelF>(H, 19); if (f == null) return null;
            IntPtr used; return Utf8.S(f(H, Utf8.Z(mime), out used));
        }
        public bool Paste(string mime, byte[] data)
        {
            PasteF f = Lok.Slot<PasteF>(H, 20); if (f == null) return false;
            return f(H, Utf8.Z(mime), data, new UIntPtr((uint)data.Length));
        }
        public void ResetSelection() { V f = Lok.Slot<V>(H, 22); if (f != null) f(H); }
        public string CommandValues(string cmd) { CmdValF f = Lok.Slot<CmdValF>(H, 23); return f == null ? null : Utf8.S(f(H, Utf8.Z(cmd))); }
        public void ClientZoom(int tpw, int tph, int ttw, int tth) { ZoomF f = Lok.Slot<ZoomF>(H, 24); if (f != null) f(H, tpw, tph, ttw, tth); }
        public void VisibleArea(int x, int y, int w, int h) { AreaF f = Lok.Slot<AreaF>(H, 25); if (f != null) f(H, x, y, w, h); }
        public int CreateView() { I f = Lok.Slot<I>(H, 26); return f == null ? -1 : f(H); }
        public void SetView(int id) { VI f = Lok.Slot<VI>(H, 28); if (f != null) f(H, id); }
        public void DestroyView(int id) { VI f = Lok.Slot<VI>(H, 27); if (f != null) f(H, id); }
        public int GetView() { I f = Lok.Slot<I>(H, 29); return f == null ? -1 : f(H); }
        public int ViewsCount() { I f = Lok.Slot<I>(H, 30); return f == null ? -1 : f(H); }
    }

    /// <summary>Scrollable tiled canvas showing a LokDoc. All coordinates in the document are twips.</summary>
    class LokCanvas : Control
    {
        public LokDoc Doc;
        public float Zoom = 1f;                       // 1 = 100%
        public Point Scroll;                          // scroll offset in device pixels
        public event Action<int, string> Message;     // LOK callbacks re-posted on the UI thread
        const int TILE = 256;
        readonly Dictionary<long, Bitmap> tiles = new Dictionary<long, Bitmap>();
        readonly Queue<long> want = new Queue<long>();
        readonly HashSet<long> queued = new HashSet<long>();
        readonly object tl = new object();
        volatile bool alive = true; readonly AutoResetEvent kick = new AutoResetEvent(false);
        readonly System.Windows.Forms.Timer renderTimer = new System.Windows.Forms.Timer { Interval = 15 };
        Size docTw;
        Rectangle caretTw = Rectangle.Empty; bool caretOn = true, caretVisible = true;
        List<Rectangle> selTw = new List<Rectangle>();
        Rectangle cellTw = Rectangle.Empty;
        readonly System.Windows.Forms.Timer blink = new System.Windows.Forms.Timer { Interval = 530 };
        int partForTiles = -1;
        public bool PageGaps = true;
        public bool SheetMode;          // Excel: sheet starts at the top-left, white everywhere, scrolls past the used range
        public bool Suspend;            // true while thumbnails / slide-show frames are rendered from other parts
        public bool ParentDiffers;
        public static void SwapRBPublic(IntPtr p, int n) { SwapRB(p, n); }
        public Rectangle CaretTw { get { return caretTw; } }
        public Rectangle CellPx { get { if (cellTw.IsEmpty) return Rectangle.Empty; Rectangle r = ToPx(cellTw); return r; } }

        public LokCanvas()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.Selectable | ControlStyles.ResizeRedraw, true);
            BackColor = Color.FromArgb(0xE6, 0xE6, 0xE6);
            Cursor = Cursors.IBeam;
            dragScroll.Tick += (s, e) => DragAutoScroll();
            blink.Tick += (s, e) => { caretOn = !caretOn; Invalidate(CaretPx()); };
            blink.Start();
        }

        float TwPerPx { get { return 15f / Zoom; } }         // 1440 twips/inch at 96 dpi
        int TileTw { get { return (int)Math.Round(TILE * TwPerPx); } }
        public Size DocPx { get { Size d = new Size((int)(docTw.Width / TwPerPx), (int)(docTw.Height / TwPerPx)); if (SheetMode) d = new Size(Math.Max(d.Width + Width, Scroll.X + Width * 2), Math.Max(d.Height + Height, Scroll.Y + Height * 2)); return d; } }
        int OffsetX { get { if (SheetMode) return -Scroll.X; int dw = DocPx.Width; return dw < Width ? (Width - dw) / 2 : -Scroll.X; } }
        int OffsetY { get { return (SheetMode ? 0 : 8) - Scroll.Y; } }
        public int PageLeftPx { get { return OffsetX; } }
        public float PxPerCm { get { return 567f / TwPerPx; } } // 1 cm = 567 twips
        public int FreezeX, FreezeY;    // frozen panes extent in pixels (Excel), from the top-left of the sheet
        public Point ToTw(Point px)
        {
            int dx = SheetMode && px.X < FreezeX ? 0 : OffsetX, dy = SheetMode && px.Y < FreezeY ? 0 : OffsetY;
            return new Point((int)((px.X - dx) * TwPerPx), (int)((px.Y - dy) * TwPerPx));
        }
        Rectangle ToPx(Rectangle tw)
        {
            int x = (int)(tw.X / TwPerPx), y = (int)(tw.Y / TwPerPx);
            int dx = SheetMode && FreezeX > 0 && x < FreezeX ? 0 : OffsetX, dy = SheetMode && FreezeY > 0 && y < FreezeY ? 0 : OffsetY;
            return new Rectangle(x + dx, y + dy, Math.Max(1, (int)(tw.Width / TwPerPx)), Math.Max(1, (int)(tw.Height / TwPerPx)));
        }
        /// <summary>Pixel rectangles (with their pane clip) of a document rectangle: one part per frozen / scrolled pane it touches.</summary>
        List<Rectangle[]> Panes(Rectangle tw)
        {
            List<Rectangle[]> res = new List<Rectangle[]>();
            if (!SheetMode || (FreezeX <= 0 && FreezeY <= 0)) { res.Add(new[] { ToPx(tw), ClientRectangle }); return res; }
            int fxTw = (int)(FreezeX * TwPerPx), fyTw = (int)(FreezeY * TwPerPx);
            List<int[]> xs = new List<int[]>(), ys = new List<int[]>();
            if (FreezeX > 0 && tw.Left < fxTw && tw.Right > fxTw) { xs.Add(new[] { tw.Left, fxTw, 0 }); xs.Add(new[] { fxTw, tw.Right, 1 }); } else xs.Add(new[] { tw.Left, tw.Right, FreezeX > 0 && tw.Left < fxTw ? 0 : 1 });
            if (FreezeY > 0 && tw.Top < fyTw && tw.Bottom > fyTw) { ys.Add(new[] { tw.Top, fyTw, 0 }); ys.Add(new[] { fyTw, tw.Bottom, 1 }); } else ys.Add(new[] { tw.Top, tw.Bottom, FreezeY > 0 && tw.Top < fyTw ? 0 : 1 });
            foreach (int[] xr in xs) foreach (int[] yr in ys)
                {
                    int ox = xr[2] == 0 ? 0 : OffsetX, oy = yr[2] == 0 ? 0 : OffsetY;
                    Rectangle px = new Rectangle((int)(xr[0] / TwPerPx) + ox, (int)(yr[0] / TwPerPx) + oy, Math.Max(1, (int)((xr[1] - xr[0]) / TwPerPx)), Math.Max(1, (int)((yr[1] - yr[0]) / TwPerPx)));
                    int cx = xr[2] == 0 || FreezeX <= 0 ? 0 : FreezeX, cy = yr[2] == 0 || FreezeY <= 0 ? 0 : FreezeY;
                    int cw = xr[2] == 0 && FreezeX > 0 ? FreezeX : Width - cx, ch = yr[2] == 0 && FreezeY > 0 ? FreezeY : Height - cy;
                    res.Add(new[] { px, new Rectangle(cx, cy, cw, ch) });
                }
            return res;
        }
        Rectangle CaretPx() { Rectangle r = ToPx(caretTw); r.Width = Math.Max(2, r.Width); r.Inflate(2, 2); return r; }

        public void Attach(LokDoc d)
        {
            Doc = d;
            if (d.Type == LokDoc.SPREADSHEET) Cursor = CellCursor;   // v0.1.7: Excel pointer from the first frame
            d.Register(OnLok);
            RefreshSize();
            renderTimer.Tick += (s, e) => RenderSome();
            renderTimer.Start();
        }

        public void RefreshSize() { if (Doc != null) { docTw = Doc.SizeTwips; } }

        void OnLok(int type, IntPtr payload, IntPtr data)
        {
            string p = Utf8.S(payload);
            try { if (IsHandleCreated) BeginInvoke(new Action(() => Handle_(type, p))); } catch { }
        }

        static Rectangle ParseRect(string s)
        {
            string[] a = s.Split(',');
            if (a.Length < 4) return Rectangle.Empty;
            int x, y, w, h;
            int.TryParse(a[0].Trim(), out x); int.TryParse(a[1].Trim(), out y); int.TryParse(a[2].Trim(), out w); int.TryParse(a[3].Trim(), out h);
            return new Rectangle(x, y, w, h);
        }

        void Handle_(int type, string p)
        {
            switch (type)
            {
                case 0: // INVALIDATE_TILES  ("x, y, w, h[, part]" or "EMPTY[, part]")
                    if (Suspend) return;
                    if (p != null) { string[] f = p.Split(','); int pp; if (Doc != null && Doc.Type != LokDoc.TEXT && f.Length >= (p.StartsWith("EMPTY") ? 2 : 5) && int.TryParse(f[p.StartsWith("EMPTY") ? 1 : 4].Trim(), out pp) && pp != Doc.Part) return; }
                    if (p == null || p.StartsWith("EMPTY")) DropAll(); else DropTiles(ParseRect(p));
                    Invalidate(); break;
                case 1: caretTw = ParseRect(p ?? ""); caretOn = true; Invalidate(); EnsureVisible(caretTw); break; // INVALIDATE_VISIBLE_CURSOR
                case 2: // TEXT_SELECTION
                    selTw.Clear();
                    if (!string.IsNullOrEmpty(p)) foreach (string r in p.Split(';')) { Rectangle rr = ParseRect(r); if (!rr.IsEmpty) selTw.Add(rr); }
                    Invalidate(); break;
                case 5: caretVisible = p == "true"; Invalidate(); break; // CURSOR_VISIBLE
                case 13: RefreshSize(); DropAll(); Invalidate(); break; // DOCUMENT_SIZE_CHANGED
                case 14: if (Suspend) return; RefreshSize(); DropAll(); Invalidate(); break; // SET_PART
                case 17: cellTw = p == "EMPTY" ? Rectangle.Empty : ParseRect(p ?? ""); Invalidate(); if (!cellTw.IsEmpty) EnsureVisible(cellTw); break; // CELL_CURSOR
                case 18: SetPointer(p); break; // MOUSE_POINTER
            }
            if (Message != null) Message(type, p);
        }

        // v0.1.7: Excel-style white "fat cross" pointer for worksheets (was the thin system cross / I-beam)
        static Cursor cellCursor;
        public static Cursor CellCursor
        {
            get
            {
                if (cellCursor != null) return cellCursor;
                try
                {
                    using (Bitmap b = new Bitmap(32, 32, PixelFormat.Format32bppArgb))
                    {
                        using (Graphics g = Graphics.FromImage(b))
                        {
                            g.Clear(Color.Transparent);
                            Rectangle[] arms = { new Rectangle(13, 6, 7, 21), new Rectangle(6, 13, 21, 7) };
                            foreach (Rectangle r in arms) g.FillRectangle(Brushes.Black, r);
                            g.FillRectangle(Brushes.White, 14, 7, 5, 19); g.FillRectangle(Brushes.White, 7, 14, 19, 5);
                        }
                        cellCursor = new Cursor(b.GetHicon());
                    }
                }
                catch { cellCursor = Cursors.Cross; }
                return cellCursor;
            }
        }
        bool IsSheet { get { return Doc != null && Doc.Type == LokDoc.SPREADSHEET; } }
        void SetPointer(string p)
        {
            switch (p)
            {
                case "text": Cursor = IsSheet && caretTw.IsEmpty ? CellCursor : Cursors.IBeam; break;
                case "pointer": Cursor = Cursors.Hand; break;
                case "col-resize": Cursor = Cursors.VSplit; break;
                case "row-resize": Cursor = Cursors.HSplit; break;
                case "move": Cursor = Cursors.SizeAll; break;
                case "crosshair": case "cell": Cursor = IsSheet ? CellCursor : Cursors.Cross; break;
                case "nwse-resize": Cursor = Cursors.SizeNWSE; break;
                case "nesw-resize": Cursor = Cursors.SizeNESW; break;
                case "ns-resize": Cursor = Cursors.SizeNS; break;
                case "ew-resize": Cursor = Cursors.SizeWE; break;
                default: Cursor = IsSheet ? CellCursor : Cursors.Default; break;
            }
        }

        public void EnsureVisible(Rectangle tw)
        {
            if (tw.IsEmpty) return;
            Rectangle r = ToPx(tw);
            int y = Scroll.Y, x = Scroll.X;
            // v0.1.7: with frozen panes the scrolled body starts below / right of the pinned rows and columns
            int top = 0, left = 0;
            if (SheetMode)
            {
                int dy = (int)(tw.Y / TwPerPx), dx = (int)(tw.X / TwPerPx);
                if (FreezeY > 0 && dy >= FreezeY) top = FreezeY;
                if (FreezeX > 0 && dx >= FreezeX) left = FreezeX;
                if (FreezeY > 0 && dy < FreezeY) { r.Y = 0; r.Height = 0; }   // pinned row: never scroll vertically for it
                if (FreezeX > 0 && dx < FreezeX) { r.X = 0; r.Width = 0; }
            }
            int pad = SheetMode ? 4 : 40;
            if (r.Top < top) y += r.Top - top - pad; else if (r.Bottom > Height) y += r.Bottom - Height + pad;
            if (DocPx.Width > Width || SheetMode) { if (r.Left < left) x += r.Left - left - pad; else if (r.Right > Width) x += r.Right - Width + pad; }
            if (y != Scroll.Y || x != Scroll.X) { ScrollTo(x, y); }
        }

        public event Action Scrolled;
        public void ScrollTo(int x, int y)
        {
            Size d = DocPx;
            if (SheetMode) { Scroll = new Point(Math.Max(0, x), Math.Max(0, y)); d = DocPx; }
            x = Math.Max(0, Math.Min(x, Math.Max(0, d.Width - Width + 16)));
            y = Math.Max(0, Math.Min(y, Math.Max(0, d.Height - Height + 40)));
            Scroll = new Point(x, y);
            if (Doc != null) { Point a = ToTw(Point.Empty); try { Doc.VisibleArea(Math.Max(0, a.X), Math.Max(0, a.Y), (int)(Width * TwPerPx), (int)(Height * TwPerPx)); } catch { } }
            Invalidate(); if (Scrolled != null) Scrolled();
        }

        public void SetZoom(float z)
        {
            z = Math.Max(0.25f, Math.Min(4f, z));
            float ratio = z / Zoom; Zoom = z;
            DropAll();
            if (Doc != null) try { Doc.ClientZoom(TILE, TILE, TileTw, TileTw); } catch { }
            ScrollTo((int)(Scroll.X * ratio), (int)(Scroll.Y * ratio));
        }

        /// <summary>Forces every tile to be re-rendered (API changes such as new conditional formats send no invalidation).</summary>
        public void RepaintAll() { DropAll(); Invalidate(); }
        void DropAll() { lock (tl) { foreach (Bitmap b in tiles.Values) b.Dispose(); tiles.Clear(); queued.Clear(); want.Clear(); } }
        void DropTiles(Rectangle tw)
        {
            int t = TileTw;
            lock (tl)
            {
                List<long> kill = new List<long>();
                foreach (long k in tiles.Keys)
                {
                    int tx = (int)(k >> 32), ty = (int)(k & 0xffffffff);
                    if (new Rectangle(tx * t, ty * t, t, t).IntersectsWith(tw)) kill.Add(k);
                }
                foreach (long k in kill) { tiles[k].Dispose(); tiles.Remove(k); }
            }
        }

        // UI-thread tile rendering: a few tiles per tick keeps the window responsive (LibreOffice's VCL main thread == UI thread)
        void RenderSome()
        {
            if (!alive || Doc == null) return;
            DateTime until = DateTime.UtcNow.AddMilliseconds(40); bool any = false;
            while (DateTime.UtcNow < until)
            {
                long k = -1; float zoom = Zoom;
                lock (tl) { if (want.Count > 0) k = want.Dequeue(); }
                if (k < 0) break;
                any = true;
                int tx = (int)(k >> 32), ty = (int)(k & 0xffffffff), t = (int)Math.Round(TILE * 15f / zoom);
                Bitmap bmp = new Bitmap(TILE, TILE, PixelFormat.Format32bppPArgb);
                BitmapData bd = bmp.LockBits(new Rectangle(0, 0, TILE, TILE), ImageLockMode.ReadWrite, PixelFormat.Format32bppPArgb);
                try
                {
                    Doc.Paint(bd.Scan0, TILE, TILE, tx * t, ty * t, t, t);
                    if (Doc.TileMode == 0) SwapRB(bd.Scan0, TILE * TILE); // RGBA -> BGRA
                }
                catch { }
                bmp.UnlockBits(bd);
                bool stale = zoom != Zoom;
                lock (tl) { queued.Remove(k); if (!stale && !tiles.ContainsKey(k)) tiles[k] = bmp; else bmp.Dispose(); }
            }
            if (any) Invalidate();
        }

        static unsafe void SwapRB(IntPtr p, int n) { byte* b = (byte*)p; for (int i = 0; i < n; i++, b += 4) { byte t = b[0]; b[0] = b[2]; b[2] = t; } }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(BackColor);
            if (Doc == null) return;
            if (partForTiles != Doc.Part) { partForTiles = Doc.Part; }
            Size d = DocPx; int ox = OffsetX, oy = OffsetY;
            Rectangle page = SheetMode ? new Rectangle(0, 0, Width, Height) : new Rectangle(ox, oy, d.Width, d.Height);
            Region clip = new Region(page);
            List<Rectangle> pages = !SheetMode && Doc.Type == LokDoc.TEXT ? Doc.PageRects() : null;
            if (pages != null && pages.Count > 0)
            {
                // Writer: only the pages are white; the gaps between them stay workspace grey (LibreOffice paints them black on Windows)
                clip.MakeEmpty();
                using (SolidBrush sh = new SolidBrush(Color.FromArgb(34, 0, 0, 0)))
                    foreach (Rectangle pr in pages)
                    {
                        Rectangle px = ToPx(pr); if (px.Bottom < 0 || px.Top > Height) continue;
                        g.FillRectangle(sh, px.X + 2, px.Y + 2, px.Width, px.Height); g.FillRectangle(Brushes.White, px); clip.Union(px);
                    }
            }
            else
            {
                if (!SheetMode) using (SolidBrush sh = new SolidBrush(Color.FromArgb(40, 0, 0, 0))) g.FillRectangle(sh, page.X + 3, page.Y + 3, page.Width, page.Height);
                g.FillRectangle(Brushes.White, page);
            }
            bool need = false;
            if (SheetMode && (FreezeX > 0 || FreezeY > 0))
            {
                // Excel frozen panes: scrolled body + pinned top rows / left columns / corner
                int fx = Math.Min(FreezeX, Width), fy = Math.Min(FreezeY, Height);
                need |= DrawTiles(g, -Scroll.X, -Scroll.Y, d, new Rectangle(fx, fy, Width - fx, Height - fy));
                if (fy > 0) need |= DrawTiles(g, -Scroll.X, 0, d, new Rectangle(fx, 0, Width - fx, fy));
                if (fx > 0) need |= DrawTiles(g, 0, -Scroll.Y, d, new Rectangle(0, fy, fx, Height - fy));
                if (fx > 0 && fy > 0) need |= DrawTiles(g, 0, 0, d, new Rectangle(0, 0, fx, fy));
                using (Pen fp = new Pen(Color.FromArgb(0x9E, 0x9E, 0x9E))) { if (fy > 0) g.DrawLine(fp, 0, fy, Width, fy); if (fx > 0) g.DrawLine(fp, fx, 0, fx, Height); }
                clip.Dispose();
            }
            else
            {
                g.SetClip(clip, System.Drawing.Drawing2D.CombineMode.Replace); // tiles past the page edge come back opaque black from LibreOffice
                need |= DrawTiles(g, ox, oy, d, Rectangle.Empty);
                g.ResetClip(); clip.Dispose();
            }
            if (need) kick.Set();
            // v0.1.7: selection / cell cursor are split at the frozen-pane lines and clipped per pane (they used to be drawn
            // over the pinned rows and columns while the body was scrolled)
            using (SolidBrush sb = new SolidBrush(Color.FromArgb(70, 0x21, 0x73, 0xD9))) foreach (Rectangle r in selTw) foreach (Rectangle[] pc in Panes(r)) { g.SetClip(pc[1]); g.FillRectangle(sb, pc[0]); g.ResetClip(); }
            if (!cellTw.IsEmpty)
                using (Pen pn = new Pen(Color.FromArgb(0x10, 0x7C, 0x41), 2f)) using (SolidBrush hb = new SolidBrush(Color.FromArgb(0x10, 0x7C, 0x41)))
                    foreach (Rectangle[] pc in Panes(cellTw)) { g.SetClip(pc[1]); Rectangle cr = pc[0]; g.DrawRectangle(pn, cr); g.FillRectangle(hb, cr.Right - 3, cr.Bottom - 3, 6, 6); g.ResetClip(); }
            if (caretVisible && caretOn && !caretTw.IsEmpty && Focused) { Rectangle c = ToPx(caretTw); g.FillRectangle(Brushes.Black, c.X, c.Y, Math.Max(2, c.Width), c.Height); }
        }

        bool DrawTiles(Graphics g, int ox, int oy, Size d, Rectangle clipRect)
        {
            bool need = false;
            Rectangle area = clipRect.IsEmpty ? new Rectangle(0, 0, Width, Height) : clipRect;
            if (!clipRect.IsEmpty) g.SetClip(clipRect);
            int x0 = Math.Max(0, (area.X - ox) / TILE), y0 = Math.Max(0, (area.Y - oy) / TILE);
            int x1 = Math.Min((d.Width - 1) / TILE, (area.Right - ox) / TILE), y1 = Math.Min((d.Height - 1) / TILE, (area.Bottom - oy) / TILE);
            lock (tl)
            {
                for (int ty = y0; ty <= y1; ty++) for (int tx = x0; tx <= x1; tx++)
                    {
                        long k = ((long)tx << 32) | (uint)ty; Bitmap b;
                        if (tiles.TryGetValue(k, out b)) g.DrawImageUnscaled(b, ox + tx * TILE, oy + ty * TILE);
                        else if (!queued.Contains(k)) { queued.Add(k); want.Enqueue(k); need = true; }
                    }
                if (tiles.Count > 1800) { foreach (Bitmap b in tiles.Values) b.Dispose(); tiles.Clear(); }
            }
            if (!clipRect.IsEmpty) g.ResetClip();
            return need;
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            if ((ModifierKeys & Keys.Control) != 0) { SetZoom(Zoom * (e.Delta > 0 ? 1.1f : 1 / 1.1f)); return; }
            if ((ModifierKeys & Keys.Shift) != 0) ScrollTo(Scroll.X - e.Delta, Scroll.Y); else ScrollTo(Scroll.X, Scroll.Y - e.Delta);
        }

        static int Btn(MouseButtons b) { return (b & MouseButtons.Left) != 0 ? 1 : (b & MouseButtons.Right) != 0 ? 4 : (b & MouseButtons.Middle) != 0 ? 2 : 1; }
        static int Mods() { int m = 0; if ((ModifierKeys & Keys.Shift) != 0) m |= 0x1000; if ((ModifierKeys & Keys.Control) != 0) m |= 0x2000; if ((ModifierKeys & Keys.Alt) != 0) m |= 0x4000; return m; }
        int clicks = 1; DateTime lastDown; Point lastPt;
        protected override void OnMouseDown(MouseEventArgs e)
        {
            Focus(); if (Doc == null) return;
            if ((DateTime.Now - lastDown).TotalMilliseconds < SystemInformation.DoubleClickTime && Math.Abs(e.X - lastPt.X) < 4 && Math.Abs(e.Y - lastPt.Y) < 4) clicks = Math.Min(3, clicks + 1); else clicks = 1;
            lastDown = DateTime.Now; lastPt = e.Location;
            Point t = ToTw(e.Location); Doc.Mouse(0, t.X, t.Y, clicks, Btn(e.Button), Mods());
        }
        protected override void OnMouseUp(MouseEventArgs e) { dragScroll.Stop(); if (Doc == null) return; Point t = ToTw(e.Location); Doc.Mouse(1, t.X, t.Y, clicks, Btn(e.Button), Mods()); }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (Doc == null) return; Point t = ToTw(e.Location); Doc.Mouse(2, t.X, t.Y, 1, e.Button == MouseButtons.None ? 0 : Btn(e.Button), Mods());
            // v0.1.7: drag-select past the window edge scrolls the sheet / document like Office
            if (e.Button == MouseButtons.Left && !ClientRectangle.Contains(e.Location)) { dragPt = e.Location; if (!dragScroll.Enabled) dragScroll.Start(); } else dragScroll.Stop();
        }
        readonly System.Windows.Forms.Timer dragScroll = new System.Windows.Forms.Timer { Interval = 60 };
        Point dragPt;
        void DragAutoScroll()
        {
            if (Doc == null || (MouseButtons & MouseButtons.Left) == 0) { dragScroll.Stop(); return; }
            Point p = PointToClient(MousePosition); dragPt = p;
            int sx = p.X < 0 ? Math.Max(-60, p.X) : p.X > Width ? Math.Min(60, p.X - Width) : 0;
            int sy = p.Y < 0 ? Math.Max(-60, p.Y) : p.Y > Height ? Math.Min(60, p.Y - Height) : 0;
            if (sx == 0 && sy == 0) { dragScroll.Stop(); return; }
            ScrollTo(Scroll.X + sx, Scroll.Y + sy);
            Point c = new Point(Math.Max(0, Math.Min(Width - 1, p.X)), Math.Max(0, Math.Min(Height - 1, p.Y)));
            Point t = ToTw(c); Doc.Mouse(2, t.X, t.Y, 1, 1, Mods());
        }

        protected override bool IsInputKey(Keys k) { return true; }
        // com.sun.star.awt.Key codes
        static int AwtKey(Keys k)
        {
            Keys c = k & Keys.KeyCode;
            if (c >= Keys.D0 && c <= Keys.D9) return 256 + (c - Keys.D0);
            if (c >= Keys.NumPad0 && c <= Keys.NumPad9) return 256 + (c - Keys.NumPad0);
            if (c >= Keys.A && c <= Keys.Z) return 512 + (c - Keys.A);
            if (c >= Keys.F1 && c <= Keys.F24) return 768 + (c - Keys.F1);
            switch (c)
            {
                case Keys.Down: return 1024; case Keys.Up: return 1025; case Keys.Left: return 1026; case Keys.Right: return 1027;
                case Keys.Home: return 1028; case Keys.End: return 1029; case Keys.PageUp: return 1030; case Keys.PageDown: return 1031;
                case Keys.Enter: return 1280; case Keys.Escape: return 1281; case Keys.Tab: return 1282; case Keys.Back: return 1283;
                case Keys.Space: return 1284; case Keys.Insert: return 1285; case Keys.Delete: return 1286;
                case Keys.Add: case Keys.Oemplus: return 1287; case Keys.Subtract: case Keys.OemMinus: return 1288;
                case Keys.Multiply: return 1289; case Keys.Divide: return 1290; case Keys.OemPeriod: case Keys.Decimal: return 1291; case Keys.Oemcomma: return 1292;
            }
            return 0;
        }
        bool handledByKeyDown;
        protected override void OnKeyDown(KeyEventArgs e)
        {
            handledByKeyDown = false; if (Doc == null) return;
            if (SheetMode && e.Control && e.KeyCode == Keys.Home) ScrollTo(0, 0);   // v0.1.7: Ctrl+Home also brings the scrolled body back to the top (Excel)
            int awt = AwtKey(e.KeyCode);
            bool ctrlCombo = (e.Control || e.Alt) && awt != 0;
            bool special = awt >= 768 && !(awt == 1284 && !e.Control) || (awt >= 1024 && awt <= 1031);
            if (e.KeyCode == Keys.V && e.Control && !e.Alt) { PasteClipboard(); e.Handled = true; handledByKeyDown = true; return; }
            if (e.KeyCode == Keys.C && e.Control && !e.Alt) { CopySelection(false); e.Handled = true; handledByKeyDown = true; return; }
            if (e.KeyCode == Keys.X && e.Control && !e.Alt) { CopySelection(true); e.Handled = true; handledByKeyDown = true; return; }
            if (ctrlCombo || (special && awt != 1284))
            {
                int code = awt | Mods();
                Doc.Key(0, 0, code); Doc.Key(1, 0, code);
                e.Handled = true; e.SuppressKeyPress = true; handledByKeyDown = true; caretOn = true;
            }
        }
        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            if (Doc == null || handledByKeyDown) return;
            char ch = e.KeyChar; if (ch < 32 && ch != '\t') return;
            Doc.Key(0, ch, 0); Doc.Key(1, ch, 0); e.Handled = true; caretOn = true;
        }
        public void TypeText(string s) { if (Doc == null || s == null) return; foreach (char ch in s) { if (ch == '\n') { Doc.Key(0, 13, 1280); Doc.Key(1, 13, 1280); } else if (ch != '\r') { Doc.Key(0, ch, 0); Doc.Key(1, ch, 0); } } }

        public void CopySelection(bool cut)
        {
            if (Doc == null) return;
            string html = null, text = null;
            try { text = Doc.GetTextSelection("text/plain;charset=utf-8"); } catch { }
            try { html = Doc.GetTextSelection("text/html"); } catch { }
            DataObject o = new DataObject();
            if (!string.IsNullOrEmpty(text)) o.SetText(text, TextDataFormat.UnicodeText);
            if (!string.IsNullOrEmpty(html)) o.SetData(DataFormats.Html, ClipHtml(html));
            try { Clipboard.SetDataObject(o, true); } catch { }
            if (cut) Doc.Uno(".uno:Cut", null);
        }
        static string ClipHtml(string html)
        {
            const string head = "Version:0.9\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";
            string pre = "<html><body><!--StartFragment-->", post = "<!--EndFragment--></body></html>";
            int hl = string.Format(head, 0, 0, 0, 0).Length;
            int sh = hl, sf = sh + Encoding.UTF8.GetByteCount(pre), ef = sf + Encoding.UTF8.GetByteCount(html), eh = ef + Encoding.UTF8.GetByteCount(post);
            return string.Format(head, sh, eh, sf, ef) + pre + html + post;
        }
        public void PasteClipboard()
        {
            if (Doc == null) return;
            try
            {
                if (Clipboard.ContainsImage())
                {
                    using (MemoryStream ms = new MemoryStream()) { Clipboard.GetImage().Save(ms, ImageFormat.Png); Doc.Paste("image/png", ms.ToArray()); }
                    return;
                }
                if (Clipboard.ContainsData(DataFormats.Html))
                {
                    string h = Clipboard.GetData(DataFormats.Html) as string;
                    if (h != null) { int i = h.IndexOf("<html", StringComparison.OrdinalIgnoreCase); if (i < 0) i = h.IndexOf("<!--StartFragment", StringComparison.OrdinalIgnoreCase); if (i >= 0 && Doc.Paste("text/html", Encoding.UTF8.GetBytes(h.Substring(i)))) return; }
                }
                if (Clipboard.ContainsText()) { string t = Clipboard.GetText(); if (!Doc.Paste("text/plain;charset=utf-8", Encoding.UTF8.GetBytes(t))) TypeText(t); }
            }
            catch { }
        }

        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        protected override void OnResize(EventArgs e) { base.OnResize(e); if (Doc != null) ScrollTo(Scroll.X, Scroll.Y); }
        protected override void Dispose(bool disposing) { alive = false; kick.Set(); renderTimer.Stop(); blink.Dispose(); DropAll(); base.Dispose(disposing); }
    }


    /// <summary>Word-style rulers (cm): horizontal along the top, vertical on the left; margins shaded grey.</summary>
    class LokRuler : Control
    {
        readonly LokCanvas c; readonly bool vertical;
        public int Margin1 = -1, Margin2 = -1;   // left/right page margins in twips (from LibreOffice RULER_UPDATE)
        public LokRuler(LokCanvas canvas, bool vert)
        {
            c = canvas; vertical = vert; DoubleBuffered = true; BackColor = Color.FromArgb(0xE9, 0xE9, 0xE9);
            if (vert) { Width = 22; Dock = DockStyle.Left; } else { Height = 22; Dock = DockStyle.Top; }
            c.Scrolled += Invalidate; c.Resize += (s, e) => Invalidate();
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.Clear(BackColor); float cm = c.PxPerCm; if (cm <= 1 || c.Doc == null) return;
            List<Rectangle> pages = c.Doc.PageRects(); Rectangle pg = pages.Count > 0 ? pages[0] : new Rectangle(Point.Empty, c.Doc.SizeTwips);
            float tw2px = c.Zoom / 15f;
            if (vertical && pages.Count > 0) { Point mid = c.ToTw(new Point(0, c.Height / 3)); foreach (Rectangle r in pages) if (mid.Y >= r.Y) pg = r; }
            float origin = vertical ? (pg.Y * tw2px - c.Scroll.Y + 8) : c.PageLeftPx + pg.X * tw2px, len = (vertical ? pg.Height : pg.Width) * tw2px;
            origin += vertical ? (c.Top - Top) : (c.Left - Left);   // rulers and canvas share the same parent panel
            float mA = (Margin1 >= 0 ? Margin1 : (vertical ? 1134 : 1701)) * tw2px, mB = (Margin2 >= 0 ? Margin2 : 1134) * tw2px;
            RectangleF band = vertical ? new RectangleF(4, origin, Width - 8, len) : new RectangleF(origin, 4, len, Height - 8);
            using (SolidBrush w = new SolidBrush(Color.White)) g.FillRectangle(w, band);
            using (SolidBrush sb = new SolidBrush(Color.FromArgb(0xC8, 0xC8, 0xC8)))
            {
                if (vertical) { g.FillRectangle(sb, band.X, band.Y, band.Width, mA); g.FillRectangle(sb, band.X, band.Bottom - mB, band.Width, mB); }
                else { g.FillRectangle(sb, band.X, band.Y, mA, band.Height); g.FillRectangle(sb, band.Right - mB, band.Y, mB, band.Height); }
            }
            using (Font f = new Font("Segoe UI", 7f))
            using (Pen p = new Pen(Color.FromArgb(0x70, 0x70, 0x70)))
            {
                float zero = vertical ? band.Y + mA : band.X + mA;
                for (int i = -(int)(mA / cm); ; i++)
                {
                    float x = zero + i * cm; if (x > (vertical ? band.Bottom : band.Right)) break; if (x < (vertical ? band.Y : band.X)) continue;
                    if (i != 0 && i % 1 == 0) { string t = Math.Abs(i).ToString(); if (vertical) TextRenderer.DrawText(g, t, f, new Rectangle(2, (int)x - 7, Width - 4, 14), Color.FromArgb(0x50, 0x50, 0x50), TextFormatFlags.HorizontalCenter); else TextRenderer.DrawText(g, t, f, new Rectangle((int)x - 8, 4, 16, 14), Color.FromArgb(0x50, 0x50, 0x50), TextFormatFlags.HorizontalCenter); }
                    float h = x + cm / 2;
                    if (vertical) g.DrawLine(p, Width / 2 - 2, h, Width / 2 + 2, h); else g.DrawLine(p, h, Height / 2 - 2, h, Height / 2 + 2);
                }
            }
        }
    }

}
