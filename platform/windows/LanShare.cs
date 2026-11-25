// Office Open Source v0.1.6 - AirDrop / QuickShare over Wi-Fi (LAN): real two-way file transfer.
// GET  /s/<token>      -> download page for a shared document (QR code points here)
// GET  /s/<token>/dl   -> the file itself
// GET  /               -> "send a file to this computer" page,  POST /up?name=... (raw body) -> saved to Downloads\Office Open Source
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace OfficeOpenSource
{
    static class LanShare
    {
        public const int PortNo = 8765;
        static TcpListener listener;
        static readonly Dictionary<string, KeyValuePair<string, byte[]>> files = new Dictionary<string, KeyValuePair<string, byte[]>>();
        public static readonly List<string> Received = new List<string>();

        public static string InboxDir
        {
            get
            {
                string d = Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"), "Yomi-Office Open Source");
                Directory.CreateDirectory(d); return d;
            }
        }

        public static List<string> Ips()
        {
            List<string> r = new List<string>();
            try
            {
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (ni.OperationalStatus != OperationalStatus.Up || ni.NetworkInterfaceType == NetworkInterfaceType.Loopback) continue;
                    foreach (UnicastIPAddressInformation a in ni.GetIPProperties().UnicastAddresses)
                        if (a.Address.AddressFamily == AddressFamily.InterNetwork && !IPAddress.IsLoopback(a.Address)) r.Add(a.Address.ToString());
                }
            }
            catch { }
            return r;
        }

        public static bool Start()
        {
            if (listener != null) return true;
            try
            {
                listener = new TcpListener(IPAddress.Any, PortNo); listener.Start();
                Thread t = new Thread(() => { while (true) { TcpClient c; try { c = listener.AcceptTcpClient(); } catch { return; } ThreadPool.QueueUserWorkItem(o => { try { Serve((TcpClient)o); } catch { } }, c); } });
                t.IsBackground = true; t.Start();
                return true;
            }
            catch { listener = null; return false; }
        }

        public static Dictionary<string, object> Share(string name, byte[] data)
        {
            bool ok = Start();
            string tok = Guid.NewGuid().ToString("N").Substring(0, 10);
            lock (files) files[tok] = new KeyValuePair<string, byte[]>(name, data);
            List<string> urls = new List<string>(); foreach (string ip in Ips()) urls.Add("http://" + ip + ":" + PortNo + "/s/" + tok);
            List<string> recv = new List<string>(); foreach (string ip in Ips()) recv.Add("http://" + ip + ":" + PortNo + "/");
            return new Dictionary<string, object> { { "ok", ok }, { "token", tok }, { "urls", urls.ToArray() }, { "receive", recv.ToArray() }, { "inbox", InboxDir } };
        }

        static string Esc(string s) { return WebUtility.HtmlEncode(s ?? ""); }
        static string Page(string title, string body)
        {
            return "<!doctype html><html lang=vi><head><meta charset=utf-8><meta name=viewport content='width=device-width,initial-scale=1'><title>" + Esc(title) + "</title>" +
                "<style>body{font-family:system-ui,Segoe UI,Roboto,sans-serif;background:#f3f6fb;margin:0;display:flex;min-height:100vh;align-items:center;justify-content:center}" +
                ".c{background:#fff;border-radius:16px;box-shadow:0 8px 30px rgba(0,0,0,.12);padding:28px;max-width:420px;width:90%;text-align:center}" +
                "a.b,button{display:inline-block;background:#185abd;color:#fff;border:0;border-radius:10px;padding:12px 22px;font-size:16px;text-decoration:none;margin-top:14px;cursor:pointer}" +
                "small{color:#64748b}</style></head><body><div class=c>" + body + "</div></body></html>";
        }

        static void Serve(TcpClient c)
        {
            using (c)
            using (NetworkStream ns = c.GetStream())
            {
                ns.ReadTimeout = 30000;
                // headers
                MemoryStream hb = new MemoryStream(); int last4 = 0, b;
                while ((b = ns.ReadByte()) >= 0) { hb.WriteByte((byte)b); last4 = (last4 << 8) | b; if (last4 == 0x0D0A0D0A || hb.Length > 65536) break; }
                string head = Encoding.UTF8.GetString(hb.ToArray());
                string[] lines = head.Split(new[] { "\r\n" }, StringSplitOptions.None);
                string[] rq = lines[0].Split(' '); if (rq.Length < 2) return;
                string method = rq[0], url = rq[1];
                long clen = 0; foreach (string l in lines) if (l.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)) long.TryParse(l.Substring(15).Trim(), out clen);
                string path = url.Split('?')[0];
                if (method == "GET" && path.StartsWith("/s/"))
                {
                    string[] parts = path.Substring(3).Split('/');
                    KeyValuePair<string, byte[]> f; bool has; lock (files) has = files.TryGetValue(parts[0], out f);
                    if (!has) { Write(ns, 404, "text/html; charset=utf-8", Encoding.UTF8.GetBytes(Page("Không tìm thấy", "<h2>Liên kết đã hết hạn</h2>")), null); return; }
                    if (parts.Length > 1 && parts[1] == "dl") { Write(ns, 200, "application/octet-stream", f.Value, "attachment; filename*=UTF-8''" + Uri.EscapeDataString(f.Key)); return; }
                    Write(ns, 200, "text/html; charset=utf-8", Encoding.UTF8.GetBytes(Page(f.Key, "<div style='font-size:48px'>📄</div><h2>" + Esc(f.Key) + "</h2><small>" + (f.Value.Length / 1024) + " KB • gửi từ Yomi-Office Open Source (" + Esc(Environment.MachineName) + ")</small><br><a class=b href='/s/" + parts[0] + "/dl'>Tải xuống</a>")), null);
                    return;
                }
                if (method == "GET" && (path == "/" || path == "/index.html"))
                {
                    string js = "<script>function up(){var f=document.getElementById('f').files;if(!f.length)return;var n=0,m=document.getElementById('m');m.textContent='Đang gửi…';[].forEach.call(f,function(x){var r=new XMLHttpRequest();r.open('POST','/up?name='+encodeURIComponent(x.name));r.onload=function(){n++;m.textContent='Đã gửi '+n+'/'+f.length+' tệp';};r.send(x);});}</script>";
                    Write(ns, 200, "text/html; charset=utf-8", Encoding.UTF8.GetBytes(Page("Gửi tệp", "<h2>Gửi tệp tới " + Esc(Environment.MachineName) + "</h2><small>Tệp được lưu vào Downloads\\Yomi-Office Open Source</small><br><input id=f type=file multiple style='margin-top:16px'><br><button onclick=up()>Gửi</button><p id=m></p>" + js)), null);
                    return;
                }
                if (method == "POST" && path == "/up")
                {
                    string q = url.Contains("?") ? url.Substring(url.IndexOf('?') + 1) : "";
                    string name = "tep-nhan-" + DateTime.Now.ToString("HHmmss");
                    foreach (string kv in q.Split('&')) if (kv.StartsWith("name=")) name = Uri.UnescapeDataString(kv.Substring(5).Replace('+', ' '));
                    name = string.Join("_", Path.GetFileName(name).Split(Path.GetInvalidFileNameChars()));
                    if (clen <= 0 || clen > 1024L * 1024 * 1024) { Write(ns, 413, "text/plain", Encoding.UTF8.GetBytes("too large"), null); return; }
                    string dest = Path.Combine(InboxDir, name); int k = 1;
                    while (File.Exists(dest)) dest = Path.Combine(InboxDir, Path.GetFileNameWithoutExtension(name) + " (" + (k++) + ")" + Path.GetExtension(name));
                    using (FileStream fs = File.Create(dest)) { byte[] buf = new byte[81920]; long left = clen; while (left > 0) { int n = ns.Read(buf, 0, (int)Math.Min(buf.Length, left)); if (n <= 0) break; fs.Write(buf, 0, n); left -= n; } }
                    lock (Received) Received.Add(dest);
                    Write(ns, 200, "application/json", Encoding.UTF8.GetBytes("{\"ok\":true}"), null);
                    return;
                }
                Write(ns, 404, "text/plain", Encoding.UTF8.GetBytes("not found"), null);
            }
        }

        static void Write(Stream s, int code, string type, byte[] body, string disp)
        {
            string h = "HTTP/1.1 " + code + " OK\r\nContent-Type: " + type + "\r\nContent-Length: " + body.Length + "\r\n" + (disp != null ? "Content-Disposition: " + disp + "\r\n" : "") + "Connection: close\r\nCache-Control: no-store\r\n\r\n";
            byte[] hb = Encoding.UTF8.GetBytes(h); s.Write(hb, 0, hb.Length); s.Write(body, 0, body.Length); s.Flush();
        }
    }
}
