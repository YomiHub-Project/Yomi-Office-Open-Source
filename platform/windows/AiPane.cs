// Office Open Source v0.1.7 - native A.I assistant window (Windows)
//  * Web mode, no login: the question is typed into the A.I website inside an embedded WebView2 and the answer is read back into the chat
//  * Web mode, signed in: same window, the user signs in once on the website (persistent profile)
//  * API key mode: OpenAI, Claude, Gemini, DeepSeek, Groq, OpenRouter, xAI - the key is saved encrypted (DPAPI) on this computer
// Without the WebView2 runtime (Windows 7/8.1 older builds) web mode falls back to an external browser window + clipboard.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Web.Script.Serialization;

namespace OfficeOpenSource
{
    static class AiKeys
    {
        static string FilePath { get { return Path.Combine(App.DataDir ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOpenSource"), "ai_keys.dat"); } }
        static Dictionary<string, string> Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new Dictionary<string, string>();
                byte[] b = ProtectedData.Unprotect(File.ReadAllBytes(FilePath), Encoding.UTF8.GetBytes("OOS-AI"), DataProtectionScope.CurrentUser);
                return new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(Encoding.UTF8.GetString(b)) ?? new Dictionary<string, string>();
            }
            catch { return new Dictionary<string, string>(); }
        }
        public static string Get(string provider) { string v; return Load().TryGetValue(provider, out v) ? v : ""; }
        public static void Set(string provider, string key)
        {
            Dictionary<string, string> d = Load();
            if (string.IsNullOrEmpty(key)) d.Remove(provider); else d[provider] = key.Trim();
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                byte[] b = ProtectedData.Protect(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(d)), Encoding.UTF8.GetBytes("OOS-AI"), DataProtectionScope.CurrentUser);
                File.WriteAllBytes(FilePath, b);
            }
            catch { }
        }
    }

    class AiProvider
    {
        public string Id, Name, Url, Model, Kind; public bool Web, NoLogin, NoKey, CustomUrl;
        public override string ToString() { return Name; }
    }

    class AiPane : Form
    {
        static readonly AiProvider[] Providers = {
            new AiProvider { Id = "web:chatgpt", Name = "ChatGPT (web, không cần đăng nhập)", Url = "https://chatgpt.com/", Web = true, NoLogin = true },
            new AiProvider { Id = "web:copilot", Name = "Microsoft Copilot (web, không cần đăng nhập)", Url = "https://copilot.microsoft.com/", Web = true, NoLogin = true },
            new AiProvider { Id = "web:grok", Name = "Grok (web, không cần đăng nhập)", Url = "https://grok.com/", Web = true, NoLogin = true },
            new AiProvider { Id = "web:gemini", Name = "Google Gemini (web, đăng nhập)", Url = "https://gemini.google.com/app", Web = true },
            new AiProvider { Id = "web:claude", Name = "Claude (web, đăng nhập)", Url = "https://claude.ai/new", Web = true },
            new AiProvider { Id = "web:deepseek", Name = "DeepSeek (web, đăng nhập)", Url = "https://chat.deepseek.com/", Web = true },
            new AiProvider { Id = "freeclaude", Name = "Free Claude Code (local / proxy)", Url = "http://127.0.0.1:8082/v1/messages", Model = "claude-sonnet-4-5", Kind = "anthropic", NoKey = true, CustomUrl = true },
            new AiProvider { Id = "openai", Name = "OpenAI ChatGPT (khóa API)", Url = "https://api.openai.com/v1/chat/completions", Model = "gpt-4o-mini", Kind = "openai" },
            new AiProvider { Id = "anthropic", Name = "Anthropic Claude (khóa API)", Url = "https://api.anthropic.com/v1/messages", Model = "claude-haiku-4-5", Kind = "anthropic" },
            new AiProvider { Id = "gemini", Name = "Google Gemini (khóa API)", Url = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent", Model = "gemini-2.0-flash", Kind = "gemini" },
            new AiProvider { Id = "deepseek", Name = "DeepSeek (khóa API)", Url = "https://api.deepseek.com/chat/completions", Model = "deepseek-chat", Kind = "openai" },
            new AiProvider { Id = "groq", Name = "Groq (khóa API)", Url = "https://api.groq.com/openai/v1/chat/completions", Model = "llama-3.3-70b-versatile", Kind = "openai" },
            new AiProvider { Id = "openrouter", Name = "OpenRouter (khóa API)", Url = "https://openrouter.ai/api/v1/chat/completions", Model = "openrouter/auto", Kind = "openai" },
            new AiProvider { Id = "xai", Name = "xAI Grok (khóa API)", Url = "https://api.x.ai/v1/chat/completions", Model = "grok-4", Kind = "openai" },
        };
        const string SysPrompt = "Bạn là trợ lý A.I trong bộ ứng dụng văn phòng Yomi-Office Open Source (Word, Excel, PowerPoint). Trả lời ngắn gọn, chính xác, mặc định bằng tiếng Việt.";

        readonly ComboBox prov = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 330 };
        readonly Button keyBtn = new Button { Text = "Khóa API…", AutoSize = true };
        readonly Button webBtn = new Button { Text = "Hiện trang web", AutoSize = true };
        readonly RichTextBox log = new RichTextBox { ReadOnly = true, Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.White };
        readonly TextBox input = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical };
        readonly Button send = new Button { Text = "Gửi", Width = 76, Dock = DockStyle.Right };
        readonly Button insert = new Button { Text = "Chèn vào tài liệu", AutoSize = true, Enabled = false };
        readonly Label state = new Label { AutoSize = true, ForeColor = Color.DimGray };
        readonly SplitContainer split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical };
        readonly Action<string> insertInto;
        readonly string context;
        Microsoft.Web.WebView2.WinForms.WebView2 web; bool webReady, webFailed; string webFor;
        string lastAnswer = "";
        readonly List<string[]> history = new List<string[]>();

        public AiPane(string context, string preset, Action<string> insertInto)
        {
            this.context = context ?? ""; this.insertInto = insertInto;
            Text = "Trợ lý A.I — Yomi-Office Open Source"; Width = 1100; Height = 720; StartPosition = FormStartPosition.CenterScreen; Font = Ui.F(9.5f); BackColor = Color.White;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; // TLS 1.2 on Windows 7 / .NET 4.x

            FlowLayoutPanel top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(8, 6, 8, 0), WrapContents = false };
            top.Controls.Add(new Label { Text = "Dịch vụ:", AutoSize = true, Padding = new Padding(0, 6, 0, 0) });
            foreach (AiProvider p in Providers) prov.Items.Add(p);
            string sel = NativeSettings.Get("ai.provider", "web:chatgpt");
            for (int i = 0; i < Providers.Length; i++) if (Providers[i].Id == sel) prov.SelectedIndex = i;
            if (prov.SelectedIndex < 0) prov.SelectedIndex = 0;
            top.Controls.Add(prov); top.Controls.Add(keyBtn); top.Controls.Add(webBtn); top.Controls.Add(state);

            Panel bottom = new Panel { Dock = DockStyle.Bottom, Height = 96, Padding = new Padding(8) };
            Panel row = new Panel { Dock = DockStyle.Bottom, Height = 32 };
            FlowLayoutPanel acts = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            acts.Controls.Add(insert);
            Button copy = new Button { Text = "Sao chép", AutoSize = true }; acts.Controls.Add(copy);
            Button clear = new Button { Text = "Cuộc trò chuyện mới", AutoSize = true }; acts.Controls.Add(clear);
            row.Controls.Add(acts);
            Panel inRow = new Panel { Dock = DockStyle.Fill }; inRow.Controls.Add(input); inRow.Controls.Add(send);
            bottom.Controls.Add(inRow); bottom.Controls.Add(row);

            split.Panel1.Controls.Add(log); split.Panel1.Padding = new Padding(8);
            Controls.Add(split); Controls.Add(bottom); Controls.Add(top);
            Shown += (s, e) => { try { split.SplitterDistance = Width / 2; } catch { } UpdateMode(); if (!string.IsNullOrEmpty(preset)) { input.Text = preset; Send(); } };

            prov.SelectedIndexChanged += (s, e) => { NativeSettings.Set("ai.provider", Cur.Id); UpdateMode(); };
            keyBtn.Click += (s, e) => AskKey(Cur);
            webBtn.Click += (s, e) => { split.Panel2Collapsed = !split.Panel2Collapsed; webBtn.Text = split.Panel2Collapsed ? "Hiện trang web" : "Ẩn trang web"; };
            send.Click += (s, e) => Send();
            input.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter && !e.Shift) { e.SuppressKeyPress = true; Send(); } };
            insert.Click += (s, e) => { if (lastAnswer.Length > 0 && insertInto != null) insertInto(lastAnswer); };
            copy.Click += (s, e) => { if (lastAnswer.Length > 0) try { Clipboard.SetText(lastAnswer); } catch { } };
            clear.Click += (s, e) => { history.Clear(); log.Clear(); lastAnswer = ""; insert.Enabled = false; if (Cur.Web && web != null && webReady) try { web.CoreWebView2.Navigate(Cur.Url); } catch { } };
            if (this.context.Length > 0) Say("Ngữ cảnh", "Đã kèm " + this.context.Length + " ký tự từ phần đang chọn trong tài liệu.", Color.DimGray);
        }
        AiProvider Cur { get { return (AiProvider)prov.SelectedItem; } }

        void UpdateMode()
        {
            AiProvider p = Cur;
            keyBtn.Visible = !p.Web; webBtn.Visible = p.Web;
            keyBtn.Text = p.CustomUrl ? "Cấu hình proxy…" : "Khóa API…";
            if (p.Web)
            {
                split.Panel2Collapsed = p.NoLogin; webBtn.Text = split.Panel2Collapsed ? "Hiện trang web" : "Ẩn trang web";
                state.Text = p.NoLogin ? "Câu hỏi được gửi tự động lên trang web, câu trả lời hiện ở khung chat." : "Đăng nhập một lần ở khung bên phải, sau đó hỏi ở khung chat.";
                EnsureWeb(p);
            }
            else if (p.NoKey)
            {
                split.Panel2Collapsed = true;
                string curUrl = NativeSettings.Get("ai.url." + p.Id, p.Url);
                state.Text = "Proxy: " + curUrl + (AiKeys.Get(p.Id).Length > 0 ? " (đã lưu token)" : " (không cần khóa)");
            }
            else
            {
                split.Panel2Collapsed = true;
                state.Text = AiKeys.Get(p.Id).Length > 0 ? "Đã lưu khóa API trên máy này." : "Chưa có khóa API — bấm \"Khóa API…\".";
            }
        }

        void EnsureWeb(AiProvider p)
        {
            if (webFailed) return;
            if (web == null)
            {
                try
                {
                    web = new Microsoft.Web.WebView2.WinForms.WebView2 { Dock = DockStyle.Fill };
                    split.Panel2.Controls.Add(web);
                    string prof = Path.Combine(App.DataDir ?? Path.GetTempPath(), "profile-ai-wv2");
                    var envTask = Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, prof, null);
                    envTask.ContinueWith(t =>
                    {
                        if (t.IsFaulted) { BeginInvoke((Action)(() => WebFailed(t.Exception))); return; }
                        BeginInvoke((Action)(() =>
                        {
                            web.EnsureCoreWebView2Async(t.Result).ContinueWith(t2 => BeginInvoke((Action)(() =>
                            {
                                if (t2.IsFaulted) { WebFailed(t2.Exception); return; }
                                webReady = true; webFor = null; Navigate(Cur);
                                web.CoreWebView2.NavigationCompleted += (s3, e3) => { if (queued != null) { string qq = queued; queued = null; System.Windows.Forms.Timer w = new System.Windows.Forms.Timer { Interval = 2500 }; w.Tick += (s4, e4) => { w.Stop(); w.Dispose(); AskWeb(Cur, qq); }; w.Start(); } };
                            })));
                        }));
                    });
                }
                catch (Exception ex) { WebFailed(ex); }
                return;
            }
            if (webReady) Navigate(p);
        }
        void Navigate(AiProvider p) { if (!p.Web || webFor == p.Id) return; webFor = p.Id; try { web.CoreWebView2.Navigate(p.Url); } catch { } }
        void WebFailed(Exception ex)
        {
            webFailed = true; split.Panel2Collapsed = true; webBtn.Enabled = false;
            state.Text = "Máy chưa có WebView2 Runtime: câu hỏi sẽ được chép và mở trang web ở cửa sổ trình duyệt.";
            Lok.Trace("ai webview2: " + ex);
        }

        void AskKey(AiProvider p)
        {
            if (p.CustomUrl)
            {
                using (Form f = new Form { Text = "Cấu hình — " + p.Name, Width = 540, Height = 265, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false, Font = Font })
                {
                    Label lu = new Label { Left = 12, Top = 12, Width = 500, Text = "Địa chỉ Endpoint (mặc định http://127.0.0.1:8082/v1/messages):" };
                    TextBox tu = new TextBox { Left = 12, Top = 34, Width = 500, Text = NativeSettings.Get("ai.url." + p.Id, p.Url) };
                    Label lm = new Label { Left = 12, Top = 64, Width = 500, Text = "Model:" };
                    TextBox tm = new TextBox { Left = 12, Top = 86, Width = 500, Text = NativeSettings.Get("ai.model." + p.Id, p.Model) };
                    Label lk = new Label { Left = 12, Top = 116, Width = 500, Text = "Token proxy (không bắt buộc — chỉ nhập nếu bật Proxy Authentication):" };
                    TextBox tk = new TextBox { Left = 12, Top = 138, Width = 500, UseSystemPasswordChar = true, Text = AiKeys.Get(p.Id) };
                    Button rst = new Button { Text = "Mặc định", Left = 12, Top = 180, Width = 90 };
                    Button ok = new Button { Text = "Lưu", DialogResult = DialogResult.OK, Left = 350, Top = 180, Width = 76 };
                    Button cancel = new Button { Text = "Hủy", DialogResult = DialogResult.Cancel, Left = 436, Top = 180, Width = 76 };
                    rst.Click += (s, e) => { tu.Text = p.Url; tm.Text = p.Model; tk.Text = ""; };
                    f.Controls.AddRange(new Control[] { lu, tu, lm, tm, lk, tk, rst, ok, cancel }); f.AcceptButton = ok; f.CancelButton = cancel;
                    if (f.ShowDialog(this) == DialogResult.OK)
                    {
                        string u = tu.Text.Trim(); NativeSettings.Set("ai.url." + p.Id, u.Length > 0 ? u : p.Url);
                        string m = tm.Text.Trim(); NativeSettings.Set("ai.model." + p.Id, m.Length > 0 ? m : p.Model);
                        AiKeys.Set(p.Id, tk.Text); UpdateMode();
                    }
                }
                return;
            }
            using (Form f = new Form { Text = "Khóa API — " + p.Name, Width = 520, Height = 170, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false, Font = Font })
            {
                TextBox tb = new TextBox { Left = 12, Top = 34, Width = 480, UseSystemPasswordChar = true, Text = AiKeys.Get(p.Id) };
                Label l = new Label { Left = 12, Top = 10, Width = 480, Text = "Khóa được mã hóa và lưu trên máy này, lần sau không phải nhập lại:" };
                Button ok = new Button { Text = "Lưu", DialogResult = DialogResult.OK, Left = 330, Top = 80, Width = 76 };
                Button del = new Button { Text = "Xóa khóa", Left = 12, Top = 80, Width = 90 };
                Button cancel = new Button { Text = "Hủy", DialogResult = DialogResult.Cancel, Left = 414, Top = 80, Width = 76 };
                del.Click += (s, e) => { tb.Text = ""; };
                f.Controls.AddRange(new Control[] { l, tb, ok, del, cancel }); f.AcceptButton = ok; f.CancelButton = cancel;
                if (f.ShowDialog(this) == DialogResult.OK) { AiKeys.Set(p.Id, tb.Text); UpdateMode(); }
            }
        }

        void Say(string who, string text, Color c)
        {
            log.SelectionStart = log.TextLength; log.SelectionColor = c; log.SelectionFont = new Font(Font, FontStyle.Bold); log.AppendText(who + "\n");
            log.SelectionColor = Color.Black; log.SelectionFont = Font; log.AppendText(text.Trim() + "\n\n"); log.ScrollToCaret();
        }

        bool busy;
        void Send()
        {
            string q = input.Text.Trim(); if (q.Length == 0 || busy) return;
            input.Clear(); Say("Bạn", q, Color.FromArgb(0x1E, 0x6F, 0xD0));
            string full = history.Count == 0 && context.Length > 0 ? q + "\n\n" + (context.Length > 30000 ? context.Substring(0, 30000) : context) : q;
            AiProvider p = Cur; busy = true; send.Enabled = false; state.Text = "Đang chờ " + p.Name + "…";
            if (p.Web) { AskWeb(p, full); return; }
            string key = AiKeys.Get(p.Id);
            if (key.Length == 0 && !p.NoKey) { busy = false; send.Enabled = true; AskKey(p); if (AiKeys.Get(p.Id).Length == 0) return; key = AiKeys.Get(p.Id); busy = true; send.Enabled = false; }
            history.Add(new[] { "user", full });
            List<string[]> h = new List<string[]>(history);
            new Thread(() =>
            {
                string ans, err = null;
                try { ans = CallApi(p, key, h); } catch (Exception ex) { ans = null; err = ex.Message; }
                BeginInvoke((Action)(() => Done(p, ans, err)));
            }) { IsBackground = true }.Start();
        }
        void Done(AiProvider p, string ans, string err)
        {
            busy = false; send.Enabled = true;
            if (ans == null) { Say(p.Name, "Lỗi: " + err, Color.Firebrick); state.Text = "Lỗi khi gọi " + p.Name; if (!p.Web && history.Count > 0) history.RemoveAt(history.Count - 1); return; }
            if (!p.Web) history.Add(new[] { "assistant", ans });
            lastAnswer = ans; insert.Enabled = insertInto != null; Say(p.Name, ans, Color.FromArgb(0x2E, 0x7D, 0x32)); state.Text = "Xong.";
        }

        // ---------------- API mode ----------------
        static string CallApi(AiProvider p, string key, List<string[]> h)
        {
            JavaScriptSerializer js = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            string model = NativeSettings.Get("ai.model." + p.Id, p.Model), url = NativeSettings.Get("ai.url." + p.Id, p.Url); object body; Dictionary<string, string> hd = new Dictionary<string, string>();
            if (p.Kind == "anthropic")
            {
                List<object> msgs = new List<object>(); foreach (string[] m in h) msgs.Add(new Dictionary<string, object> { { "role", m[0] }, { "content", m[1] } });
                body = new Dictionary<string, object> { { "model", model }, { "max_tokens", 2048 }, { "system", SysPrompt }, { "messages", msgs } };
                hd["x-api-key"] = string.IsNullOrEmpty(key) ? (p.NoKey ? "free-claude-code" : "") : key;
                if (p.NoKey && !string.IsNullOrEmpty(key)) hd["Authorization"] = "Bearer " + key;
                hd["anthropic-version"] = "2023-06-01";
            }
            else if (p.Kind == "gemini")
            {
                List<object> c = new List<object>(); foreach (string[] m in h) c.Add(new Dictionary<string, object> { { "role", m[0] == "assistant" ? "model" : "user" }, { "parts", new object[] { new Dictionary<string, object> { { "text", m[1] } } } } });
                body = new Dictionary<string, object> { { "systemInstruction", new Dictionary<string, object> { { "parts", new object[] { new Dictionary<string, object> { { "text", SysPrompt } } } } } }, { "contents", c } };
                url = url.Replace("{model}", model) + "?key=" + Uri.EscapeDataString(key);
            }
            else
            {
                List<object> msgs = new List<object> { new Dictionary<string, object> { { "role", "system" }, { "content", SysPrompt } } };
                foreach (string[] m in h) msgs.Add(new Dictionary<string, object> { { "role", m[0] }, { "content", m[1] } });
                body = new Dictionary<string, object> { { "model", model }, { "messages", msgs }, { "temperature", 0.4 } };
                hd["Authorization"] = "Bearer " + key;
            }
            HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url); r.Method = "POST"; r.ContentType = "application/json"; r.Timeout = 120000;
            foreach (var kv in hd) r.Headers[kv.Key] = kv.Value;
            byte[] data = Encoding.UTF8.GetBytes(js.Serialize(body)); using (Stream s = r.GetRequestStream()) s.Write(data, 0, data.Length);
            string txt;
            try { using (WebResponse resp = r.GetResponse()) using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) txt = sr.ReadToEnd(); }
            catch (WebException we)
            {
                string e2 = we.Message; try { using (StreamReader sr = new StreamReader(we.Response.GetResponseStream(), Encoding.UTF8)) e2 = sr.ReadToEnd(); } catch { }
                if (e2.Length > 400) e2 = e2.Substring(0, 400); throw new Exception(e2);
            }
            dynamic j = js.Deserialize<object>(txt);
            if (p.Kind == "anthropic") { StringBuilder sb = new StringBuilder(); foreach (dynamic part in j["content"]) if (part["type"] == "text") sb.Append(part["text"]); return sb.ToString(); }
            if (p.Kind == "gemini") { StringBuilder sb = new StringBuilder(); foreach (dynamic part in j["candidates"][0]["content"]["parts"]) sb.Append(part["text"]); return sb.ToString(); }
            return j["choices"][0]["message"]["content"];
        }

        // ---------------- web mode: type into the website, read the last answer back ----------------
        // returns "<count>\u0001<last answer>\u0001<page text>" - known answer containers first, the page text as a fallback
        const string ReadJs = @"(function(){var s=['[data-message-author-role=""assistant""]','[data-content=""ai-message""]','.ds-markdown','message-content','.font-claude-response','[class*=""response-content""]','[data-testid*=""bot-message""]','.markdown.prose','.markdown'];var c=0,t='';for(var i=0;i<s.length;i++){var n=document.querySelectorAll(s[i]);if(n.length){c=n.length;t=n[n.length-1].innerText||'';break;}}var b=(document.body&&document.body.innerText)||'';if(b.length>60000)b=b.slice(-60000);return c+'\u0001'+t+'\u0001'+b;})()";
        static string SendJs(string q)
        {
            return "(function(q){var el=document.querySelector('#prompt-textarea,textarea#userInput,textarea[name=\"query\"],div[contenteditable=\"true\"],textarea');if(!el)return 'noinput';el.focus();" +
                "if(el.tagName==='TEXTAREA'){var st=Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype,'value').set;st.call(el,q);el.dispatchEvent(new Event('input',{bubbles:true}));}" +
                "else{document.execCommand('selectAll',false,null);document.execCommand('insertText',false,q);}" +
                "setTimeout(function(){var b=document.querySelector('[data-testid=\"send-button\"],button[aria-label*=\"Send\" i],button[aria-label*=\"Gửi\" i],button[aria-label*=\"Submit\" i],button[type=\"submit\"]');" +
                "if(b&&!b.disabled)b.click();else{['keydown','keypress','keyup'].forEach(function(t){el.dispatchEvent(new KeyboardEvent(t,{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));});}},600);return 'ok';})(" +
                new JavaScriptSerializer().Serialize(q) + ")";
        }
        static string Unjson(string r) { try { return new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Deserialize<string>(r) ?? ""; } catch { return r ?? ""; } }

        void AskWeb(AiProvider p, string q)
        {
            if (webFailed || web == null || !webReady)
            {
                if (!webFailed && web != null && !webReady) { queued = q; state.Text = "Trang web đang tải — câu hỏi sẽ tự gửi khi trang sẵn sàng…"; return; }
                try { Clipboard.SetText(q); } catch { }
                Api.AiWeb(p.Url, "Trợ lý A.I");
                Done(p, "(Đã mở " + p.Name + " ở cửa sổ trình duyệt và chép câu hỏi — bấm Ctrl+V rồi Enter.)", null); return;
            }
            Navigate(p); pendingQ = q;
            web.CoreWebView2.ExecuteScriptAsync(ReadJs).ContinueWith(t0 => BeginInvoke((Action)(() =>
            {
                string r0 = t0.IsFaulted ? "" : Unjson(t0.Result); int before = 0; int.TryParse(r0.Split('\u0001')[0], out before);
                web.CoreWebView2.ExecuteScriptAsync(SendJs(q)).ContinueWith(t1 => BeginInvoke((Action)(() =>
                {
                    string ok = t1.IsFaulted ? "err" : Unjson(t1.Result);
                    if (ok != "ok") { split.Panel2Collapsed = false; webBtn.Text = "Ẩn trang web"; Done(p, null, "trang web chưa sẵn sàng hoặc cần đăng nhập — hãy xem khung bên phải rồi thử lại."); return; }
                    Poll(p, before, "", 0, DateTime.Now);
                })));
            })));
        }
        string pendingQ = "", queued;
        static readonly string[] Noise = { "can make mistakes", "có thể mắc lỗi", "Message ChatGPT", "Ask anything", "Ask ChatGPT", "Hỏi bất kỳ điều gì", "Log in to get answers", "Đăng nhập để", "We use cookies", "Accept all", "Reject non-essential", "\nImages\n", "\nSearch\n" };
        string Extract(string[] r)
        {
            string text = r.Length > 1 ? r[1].Trim() : "";
            if (text.Length > 0) return text;
            // fallback: whatever the page shows after the last copy of the question
            string body = r.Length > 2 ? r[2] : "", key = pendingQ.Length > 40 ? pendingQ.Substring(0, 40) : pendingQ;
            int i = key.Length > 0 ? body.LastIndexOf(key, StringComparison.Ordinal) : -1; if (i < 0) return "";
            int j = body.IndexOf('\n', i + key.Length); string a = j < 0 ? "" : body.Substring(j + 1);
            if (pendingQ.Length > 40) { int k = a.IndexOf(pendingQ.Substring(pendingQ.Length - Math.Min(20, pendingQ.Length))); if (k >= 0 && k < 200) a = a.Substring(a.IndexOf('\n', k) + 1); }
            foreach (string n in Noise) { int k = a.IndexOf(n, StringComparison.OrdinalIgnoreCase); if (k >= 0) { int ls = a.LastIndexOf('\n', Math.Max(0, k - 1)); a = a.Substring(0, ls > 0 && n[0] != '\n' ? ls : k); } }
            a = System.Text.RegularExpressions.Regex.Replace(a.Trim(), @"^(ChatGPT|Copilot|Grok|Gemini|Claude|DeepSeek)( said| nói)?:?\s*\n", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            return a.Trim();
        }
        void Poll(AiProvider p, int before, string last, int stable, DateTime start)
        {
            System.Windows.Forms.Timer t = new System.Windows.Forms.Timer { Interval = 1000 };
            t.Tick += (s, e) =>
            {
                t.Stop(); t.Dispose();
                web.CoreWebView2.ExecuteScriptAsync(ReadJs).ContinueWith(tr => BeginInvoke((Action)(() =>
                {
                    string[] r = (tr.IsFaulted ? "0\u0001\u0001" : Unjson(tr.Result)).Split(new[] { '\u0001' }, 3);
                    int n = 0; int.TryParse(r[0], out n); string text = Extract(r);
                    string r1 = r.Length > 1 ? r[1].Trim() : ""; bool fresh = (n > before || r1.Length == 0) && text.Length > 0;
                    int st = fresh && text == last ? stable + 1 : 0;
                    if (fresh && st >= 3) { Done(p, text, null); return; }
                    if ((DateTime.Now - start).TotalSeconds > 150) { split.Panel2Collapsed = false; Done(p, fresh ? text : null, "quá thời gian chờ câu trả lời — xem khung trang web bên phải."); return; }
                    state.Text = fresh ? "Đang nhận câu trả lời…" : "Đang chờ " + p.Name + "…";
                    Poll(p, before, text, st, start);
                })));
            };
            t.Start();
        }
    }
}
