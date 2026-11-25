// A.I Copilot pane (v0.0.7): offline engine + Web mode (ChatGPT, Gemini, Copilot, Claude, DeepSeek, Grok, AI Hay — no login or signed in on the web)
// + API-key mode (OpenAI, Claude, Gemini, DeepSeek, Grok/xAI, Groq, OpenRouter, OpenHay self-host, Ollama, LM Studio)
import { h, ic, esc } from '../core/util.js';
import { toast, formDlg, dialog } from '../core/ui.js';
import { settings, saveSettings } from '../core/store.js';
import Host from '../core/host.js';
import * as O from './offline.js';

export const PROVIDERS = {
  offline: { name: 'Ngoại tuyến (tích hợp)', local: true },
  ollama: { name: 'Ollama (mô hình trên máy)', local: true, url: 'http://localhost:11434/v1/chat/completions', model: 'llama3.2', kind: 'openai', nokey: true },
  lmstudio: { name: 'LM Studio (mô hình trên máy)', local: true, url: 'http://localhost:1234/v1/chat/completions', model: 'local-model', kind: 'openai', nokey: true },
  openai: { name: 'OpenAI (ChatGPT)', url: 'https://api.openai.com/v1/chat/completions', model: 'gpt-4o-mini', kind: 'openai' },
  anthropic: { name: 'Anthropic (Claude)', url: 'https://api.anthropic.com/v1/messages', model: 'claude-haiku-4-5', kind: 'anthropic' },
  gemini: { name: 'Google Gemini', url: 'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent', model: 'gemini-2.0-flash', kind: 'gemini' },
  deepseek: { name: 'DeepSeek', url: 'https://api.deepseek.com/chat/completions', model: 'deepseek-chat', kind: 'openai' },
  groq: { name: 'Groq', url: 'https://api.groq.com/openai/v1/chat/completions', model: 'llama-3.3-70b-versatile', kind: 'openai' },
  openrouter: { name: 'OpenRouter', url: 'https://openrouter.ai/api/v1/chat/completions', model: 'openrouter/auto', kind: 'openai' },
  xai: { name: 'xAI (Grok)', url: 'https://api.x.ai/v1/chat/completions', model: 'grok-4', kind: 'openai' },
  azure: { name: 'Microsoft Copilot / Azure OpenAI', url: 'https://<tên-tài-nguyên>.openai.azure.com/openai/deployments/<deployment>/chat/completions?api-version=2024-10-21', model: 'gpt-4o-mini', kind: 'azure', custom: true },
  // v0.1.8: Free Claude Code (github.com/Alishahryar1/free-claude-code, AGPL-3.0) — proxy cục bộ dạng Anthropic Messages API,
  // chuyển tiếp tới ~61 nhà cung cấp (NVIDIA NIM, OpenRouter, Groq, Ollama…). Ứng dụng chỉ KẾT NỐI tới proxy, không nhúng mã của dự án.
  // Mặc định 127.0.0.1:8082; trên Android nhập địa chỉ IP của máy tính đang chạy fcc-server. Khóa = token "Proxy Authentication" (nếu bật).
  freeclaude: { name: 'Free Claude Code (proxy miễn phí)', local: true, url: 'http://127.0.0.1:8082/v1/messages', model: 'claude-sonnet-4-5', kind: 'anthropic', nokey: true, optkey: true, custom: true,
    help: 'Cài Free Claude Code trên máy tính (github.com/Alishahryar1/free-claude-code), chạy fcc-server rồi chọn nhà cung cấp miễn phí (vd NVIDIA NIM, OpenRouter) trong Admin UI. Mặc định proxy nghe ở http://127.0.0.1:8082. Nếu bật Proxy Authentication, dán token vào ô Khóa.' },
  openhay: { name: 'OpenHay (tự host)', url: 'http://localhost:8000', model: 'mặc định', kind: 'openhay', nokey: true, local: true }
};
// v0.0.7: Web mode — the official website of each A.I service opens in its own window (Windows: Edge/Chrome app window with a
// separate, persistent profile; Android: in-app WebView screen). The question is pre-filled (?q= where the site supports it) and
// copied to the clipboard; the answer comes back with "Lấy câu trả lời" (clipboard). Works without login (where the site allows) or signed in.
export const WEB = {
  'web:chatgpt': { name: 'ChatGPT (web)', url: 'https://chatgpt.com/', q: 'https://chatgpt.com/?q={q}', nologin: true },
  'web:gemini': { name: 'Google Gemini (web)', url: 'https://gemini.google.com/app', q: null, nologin: false },
  'web:copilot': { name: 'Microsoft Copilot (web)', url: 'https://copilot.microsoft.com/', q: 'https://copilot.microsoft.com/?q={q}', nologin: true },
  'web:claude': { name: 'Claude (web)', url: 'https://claude.ai/new', q: 'https://claude.ai/new?q={q}', nologin: false },
  'web:deepseek': { name: 'DeepSeek (web)', url: 'https://chat.deepseek.com/', q: null, nologin: false },
  'web:grok': { name: 'Grok (web)', url: 'https://grok.com/', q: 'https://grok.com/?q={q}', nologin: true },
  'web:aihay': { name: 'AI Hay / OpenHay (web)', url: 'https://ai-hay.vn/', q: null, nologin: true }
};
export function webUrl(id, prompt) { const W = WEB[id]; if (!W) return null; return W.q && prompt && prompt.length < 1800 ? W.q.replace('{q}', encodeURIComponent(prompt)) : W.url; }
export async function askWeb(id, prompt) {
  try { if (navigator.clipboard && navigator.clipboard.writeText) await navigator.clipboard.writeText(prompt); } catch (e) {}
  Host.clipSet && Host.clipSet(prompt);
  return Host.aiWeb(webUrl(id, prompt), WEB[id].name);
}
export const PERSISTENT_KEYS_STORE = 'oos_ai_persistent_keys';

export function loadPersistentKeys() {
  try {
    const raw = localStorage.getItem(PERSISTENT_KEYS_STORE);
    return raw ? JSON.parse(raw) : {};
  } catch (e) { return {}; }
}

export function persistKey(k, val) {
  try {
    const all = loadPersistentKeys();
    if (val && String(val).trim()) all[k] = String(val).trim(); else delete all[k];
    localStorage.setItem(PERSISTENT_KEYS_STORE, JSON.stringify(all));
  } catch (e) {}
}

export function saveAiKey(k, val) {
  const c = cfg();
  c.keys[k] = (val || '').trim();
  persistKey(k, c.keys[k]);
  saveSettings({ ai: c });
  try { Host.saveConfig({ ai_keys: c.keys }); } catch (e) {}
}

function cfg() {
  const s = settings();
  s.ai = Object.assign({ provider: 'offline', keys: {}, models: {}, urls: {} }, s.ai || {});
  const diskKeys = loadPersistentKeys();
  s.ai.keys = Object.assign({}, diskKeys, s.ai.keys);
  return s.ai;
}
function provName(k) { return (PROVIDERS[k] || WEB[k] || PROVIDERS.offline).name; }
async function openhay(url, prompt) {
  const base = url.replace(/\/+$/, ''); const st = cfg(); st._ohTok = st._ohTok || {};
  let tok = st._ohTok[base];
  if (!tok) { const r = await fetch(base + '/api/auth/token/guest', { method: 'POST' }); const j = await r.json(); tok = j.access_token; st._ohTok[base] = tok; }
  const r = await fetch(base + '/api/chat', { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + tok }, body: JSON.stringify({ message: prompt, media: [] }) });
  if (r.status === 401) { delete st._ohTok[base]; throw new Error('OpenHay: phiên hết hạn, thử lại'); }
  const txt = await r.text(); let out = ''; let ev = '';
  txt.split(/\r?\n/).forEach(l => { if (l.indexOf('event:') === 0) ev = l.slice(6).trim(); else if (l.indexOf('data:') === 0 && ev === 'ai_message') { try { const d = JSON.parse(l.slice(5)); out += (d.chunk && d.chunk.content) || ''; } catch (e) {} } });
  if (!out) throw new Error('OpenHay không trả lời (HTTP ' + r.status + ')'); return out;
}

// ---------------------------------------------------------------- document context
function ctxOf(doc) {
  const sel = (window.getSelection && String(window.getSelection())) || '';
  let text = ''; try { text = doc && doc.getText ? doc.getText() : ''; } catch (e) {}
  let range = null; try { if (doc && doc.kind === 'excel') { const s = doc.wb.cur.sel; const rg = s.ranges[s.ranges.length - 1]; range = rg ? colName(rg.c1) + (rg.r1 + 1) + ':' + colName(rg.c2) + (rg.r2 + 1) : null; } } catch (e) {}
  return { selection: sel.trim(), text, range, kind: doc ? doc.kind : null };
}
function colName(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }

// ---------------------------------------------------------------- network (direct; Windows host can proxy to avoid CORS)
async function post(url, headers, body) {
  const opts = { method: 'POST', headers: Object.assign({ 'Content-Type': 'application/json' }, headers), body: JSON.stringify(body) };
  try { const r = await fetch(url, opts); const j = await r.json(); if (!r.ok) throw new Error((j.error && (j.error.message || j.error)) || ('HTTP ' + r.status)); return j; }
  catch (e) {
    if (Host.kind === 'windows' && /fetch|network|cors/i.test(String(e.message))) { const r = await fetch('/api/proxy', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ url, headers: opts.headers, body: opts.body }) }); const j = await r.json(); if (j.error) throw new Error(j.error); return JSON.parse(j.body); }
    throw e;
  }
}
export async function complete(messages, system) {
  const c = cfg(); if (WEB[c.provider]) throw new Error('Chế độ Web: dùng khung Copilot'); const P = PROVIDERS[c.provider] || PROVIDERS.offline; if (c.provider === 'offline') throw new Error('offline');
  const key = c.keys[c.provider] || ''; const model = c.models[c.provider] || P.model; const url = c.urls[c.provider] || P.url;
  if (!key && !P.nokey) throw new Error('Chưa nhập khóa API cho ' + P.name + ' (bấm ⚙️).');
  if (P.kind === 'openai') { const j = await post(url, key ? { Authorization: 'Bearer ' + key } : {}, { model, messages: (system ? [{ role: 'system', content: system }] : []).concat(messages), temperature: 0.4 }); return j.choices[0].message.content; }
  if (P.kind === 'anthropic') { const j = await post(url, Object.assign({ 'x-api-key': key || (P.nokey ? 'free-claude-code' : '') }, key && P.nokey ? { Authorization: 'Bearer ' + key } : {}, { 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' }), { model, max_tokens: 2048, system: system || undefined, messages }); return j.content.map(x => x.text || '').join(''); }
  if (P.kind === 'azure') { const j = await post(url, { 'api-key': key }, { messages: (system ? [{ role: 'system', content: system }] : []).concat(messages), temperature: 0.4 }); return j.choices[0].message.content; }
  if (P.kind === 'openhay') return openhay(url, (system ? system + '\n\n' : '') + messages.filter(m => m.role === 'user').slice(-1).map(m => m.content).join(''));
  if (P.kind === 'gemini') { const j = await post(url.replace('{model}', model) + '?key=' + encodeURIComponent(key), {}, { systemInstruction: system ? { parts: [{ text: system }] } : undefined, contents: messages.map(m => ({ role: m.role === 'assistant' ? 'model' : 'user', parts: [{ text: m.content }] })) }); return j.candidates[0].content.parts.map(p => p.text || '').join(''); }
  throw new Error('Nhà cung cấp không hợp lệ');
}
const SYSTEM = 'Bạn là trợ lý A.I trong bộ ứng dụng văn phòng Yomi-Office Open Source (Word, Excel, PowerPoint, PDF). Trả lời ngắn gọn, chính xác, mặc định bằng tiếng Việt. Khi được yêu cầu viết công thức Excel, chỉ trả về công thức bắt đầu bằng dấu =. Khi tạo nội dung để chèn vào tài liệu, chỉ trả về nội dung, không giải thích thêm.';

// ---------------------------------------------------------------- insert into the active document
export function insertInto(doc, text, asOutline) {
  if (!doc || (text == null && !asOutline)) return;
  try {
    if (doc.kind === 'word') { doc.restoreCaret && doc.restoreCaret(); doc.body.focus(); document.execCommand('insertHTML', false, text.split(/\n/).map(l => '<p>' + (esc(l) || '<br>') + '</p>').join('')); doc.markDirty(); return toast('Đã chèn vào tài liệu'); }
    if (doc.kind === 'excel') { const s = doc.wb.cur.sel; const lines = String(text).trim().split(/\n/); if (lines.length === 1) doc.setInput(s.r, s.c, lines[0]); else lines.forEach((l, i) => doc.setInput(s.r + i, s.c, l)); doc.markDirty(); return toast('Đã chèn vào ô'); }
    if (doc.kind === 'ppt') {
      import('../ppt/cmds.js').then(P => {
        const ol = asOutline || O.outlineFromText(text);
        if (asOutline || ol.length > 1) { import('../ppt/templates.js').then(Tm => { doc.change('Tạo slide A.I', () => { ol.forEach(sl => { const s = Tm.blankSlide('content', doc.theme); s.objects[0].html = esc(sl.title); s.objects[1].html = sl.bullets.map(b => '<div>• ' + esc(b) + '</div>').join(''); doc.slides.splice(++doc.cur, 0, s); }); }); toast('Đã tạo ' + ol.length + ' slide', 'ok'); }); }
        else P.addObj(doc, { type: 'text', x: 140, y: 200, w: 1000, h: 300, html: esc(text).replace(/\n/g, '<br>'), fontSize: 22 });
      });
      return;
    }
    navigator.clipboard && navigator.clipboard.writeText(text); toast('Đã sao chép kết quả');
  } catch (e) { toast('Không chèn được: ' + e.message, 'err'); }
}

// ---------------------------------------------------------------- Copilot pane
let pane = null, history = [];
export function openAi(shell, doc, preset) {
  if (pane && pane.parentNode) { if (preset) { pane._input.value = preset; pane._send(); } else close(); return; }
  history = [];
  const msgs = h('div.ai-msgs'); const input = h('textarea.ai-input', { rows: 3, placeholder: 'Hỏi Copilot… (Enter để gửi, Shift+Enter xuống dòng)' });
  const provSel = h('select.ai-prov',
    h('optgroup', { label: 'Ngoại tuyến / trên máy' }, Object.keys(PROVIDERS).filter(k => PROVIDERS[k].local).map(k => h('option', { value: k }, PROVIDERS[k].name))),
    h('optgroup', { label: 'Web — không cần đăng nhập hoặc đăng nhập trên web' }, Object.keys(WEB).map(k => h('option', { value: k }, WEB[k].name + (WEB[k].nologin ? '' : ' · cần đăng nhập')))),
    h('optgroup', { label: 'Khóa API (tích hợp vào ứng dụng)' }, Object.keys(PROVIDERS).filter(k => !PROVIDERS[k].local).map(k => h('option', { value: k }, PROVIDERS[k].name))));
  provSel.value = cfg().provider; provSel.addEventListener('change', () => { cfg().provider = provSel.value; saveSettings({ ai: cfg() }); note(); });
  const status = h('div.ai-status');
  const note = () => { const k = cfg().provider; const W = WEB[k]; const P = PROVIDERS[k] || PROVIDERS.offline;
    if (W) { status.textContent = '● Web — mở ' + W.name.replace(' (web)', '') + (W.nologin ? ' (dùng được không cần đăng nhập)' : ' (đăng nhập trên trang web)'); status.className = 'ai-status net'; webBar.style.display = ''; return; }
    webBar.style.display = 'none'; if (P.help) status.title = P.help; status.textContent = k === 'freeclaude' ? '● Free Claude Code — proxy ' + ((cfg().urls || {})[k] || P.url).replace(/\/v1\/messages$/, '') : P.local ? '● Không cần Internet' : '● Trực tuyến — dữ liệu gửi tới ' + P.name; status.className = 'ai-status ' + (P.local ? 'ok' : 'net'); };
  const webBar = h('div.ai-webbar', { style: { display: 'none' } });
  const fetchBtn = h('button.btn', ic('clipboard-paste', 13), ' Lấy câu trả lời');
  const openBtn = h('button.btn', ic('external-link', 13), ' Mở trang / đăng nhập');
  webBar.appendChild(openBtn); webBar.appendChild(fetchBtn);
  openBtn.addEventListener('click', () => Host.aiWeb(WEB[cfg().provider].url, WEB[cfg().provider].name));
  const takeAnswer = async txt => {
    txt = txt != null ? txt : await Host.clipText();
    if (!txt || !txt.trim()) return toast('Bộ nhớ tạm trống — trên trang A.I, bấm nút "Sao chép" dưới câu trả lời rồi thử lại', 'err', 5000);
    if (txt.trim() === (pane && pane._lastPrompt || '').trim()) return toast('Bộ nhớ tạm vẫn là câu hỏi — hãy sao chép câu trả lời trên trang A.I', 'err', 5000);
    history.push({ role: 'assistant', content: txt }); add('bot', txt, { insert: txt.trim(), raw: txt, outline: O.outlineFromText(txt).length > 1 ? O.outlineFromText(txt) : null });
  };
  fetchBtn.addEventListener('click', () => takeAnswer());
  window.OOS_onAiAnswer = t => { if (pane) takeAnswer(String(t || '')); };
  note();
  const chips = h('div.ai-chips', [['Tóm tắt', 'Tóm tắt nội dung'], ['Soát lỗi', 'Soát lỗi chính tả'], ['Từ khóa', 'Liệt kê từ khóa chính'], ['Dàn ý slide', 'Tạo dàn ý trình chiếu'], ['Công thức', 'Gợi ý công thức tính tổng ' + ((ctxOf(doc).range) || 'A1:A10')], ['Thống kê', 'Thống kê văn bản'], ['Dịch EN', 'Dịch sang tiếng Anh']].map(c => { const b = h('button.ai-chip', c[0]); b.addEventListener('click', () => { input.value = c[1]; send(); }); return b; }));
  const add = (role, text, extra) => {
    const m = h('div.ai-msg.' + role, h('div.ai-bubble', text));
    if (role === 'bot' && extra) {
      const act = h('div.ai-act');
      if (extra.insert != null) { const b = h('button.btn', ic('corner-down-left', 13), ' ' + (extra.fixLabel || 'Chèn vào tài liệu')); b.addEventListener('click', () => insertInto(shell.cur || doc, extra.insert)); act.appendChild(b); }
      if (extra.outline) { const b = h('button.btn', ic('presentation', 13), ' Tạo slide'); b.addEventListener('click', () => { const d = shell.cur || doc; if (d && d.kind === 'ppt') insertInto(d, null, extra.outline); else toast('Mở PowerPoint để tạo slide'); }); act.appendChild(b); }
      const cp = h('button.btn', ic('copy', 13), ' Sao chép'); cp.addEventListener('click', () => { try { navigator.clipboard.writeText(extra.raw || text); toast('Đã sao chép'); } catch (e) {} }); act.appendChild(cp);
      m.appendChild(act);
    }
    msgs.appendChild(m); msgs.scrollTop = msgs.scrollHeight; return m;
  };
  const send = async () => {
    const q = input.value.trim(); if (!q) return; input.value = ''; add('me', q);
    const d = shell.cur || doc; const ctx = ctxOf(d);
    if (cfg().provider === 'offline') { const r = O.offlineAnswer(q, ctx); add('bot', r.text, Object.assign({ raw: r.insert || r.text }, r)); return; }
    if (WEB[cfg().provider]) {
      const W = WEB[cfg().provider];
      const prompt = (ctx.selection ? 'Đoạn đang chọn:\n + ctx.selection.slice(0, 6000) + \n\n' : '') + q;
      pane._lastPrompt = prompt; history.push({ role: 'user', content: prompt });
      try { await askWeb(cfg().provider, prompt); } catch (e) { add('bot', '⚠️ Không mở được ' + W.name + ': ' + e.message); return; }
      add('bot', 'Đã gửi câu hỏi tới ' + W.name + (W.q ? ' (đã điền sẵn).' : ' — câu hỏi đã được sao chép, hãy dán (Ctrl+V) vào ô chat của trang.') + '\nKhi có câu trả lời: bấm "Sao chép" dưới câu trả lời trên trang, rồi bấm "Lấy câu trả lời" ở đây' + (Host.kind === 'android' ? ' (hoặc nút "Chèn câu trả lời" trên màn hình web).' : '.'));
      return;
    }
    const wait = add('bot', '…');
    try {
      const content = (ctx.selection ? 'Đoạn đang chọn:\n"""' + ctx.selection.slice(0, 8000) + '"""\n\n' : (ctx.text ? 'Nội dung tài liệu (' + ctx.kind + '):\n"""' + ctx.text.slice(0, 12000) + '"""\n\n' : '')) + q;
      history.push({ role: 'user', content }); const ans = await complete(history.slice(-10), SYSTEM); history.push({ role: 'assistant', content: ans });
      wait.remove(); add('bot', ans, { insert: ans.trim(), raw: ans, outline: /slide|trình chiếu|dàn ý/i.test(q) ? O.outlineFromText(ans) : null });
    } catch (e) { wait.remove(); const r = O.offlineAnswer(q, ctx); add('bot', '⚠️ ' + e.message + '\n\n↪ Kết quả ngoại tuyến:\n' + r.text, Object.assign({ raw: r.insert || r.text }, r)); }
  };
  input.addEventListener('keydown', e => { if ((e.key === 'Enter' || e.keyCode === 13) && !e.shiftKey) { e.preventDefault(); send(); } e.stopPropagation(); });
  const sendBtn = h('button.btn.primary', ic('send', 14)); sendBtn.addEventListener('click', send);
  const gear = h('button.pane-x', { title: 'Cài đặt A.I' }, ic('settings', 15)); gear.addEventListener('click', () => aiSettings().then(() => { provSel.value = cfg().provider; note(); }));
  const x = h('button.pane-x', { title: 'Đóng' }, ic('x', 16)); x.addEventListener('click', close);
  pane = h('div.ai-pane', h('div.pane-head', ic('sparkles', 16), h('span.t', ' Copilot'), h('span.sp'), gear, x), h('div.ai-top', provSel, status, webBar), msgs, chips, h('div.ai-compose', input, sendBtn));
  pane._input = input; pane._send = send;
  document.body.appendChild(pane);
  add('bot', 'Xin chào! Tôi có thể tóm tắt, soát lỗi, gợi ý công thức, tạo dàn ý slide… ngay cả khi không có Internet. Bôi đen văn bản để làm việc với đoạn chọn.');
  if (preset) { input.value = preset; send(); } else setTimeout(() => input.focus(), 50);
}
function close() { if (pane && pane.parentNode) pane.parentNode.removeChild(pane); pane = null; }

export async function aiSettings() {
  const c = cfg();
  const fields = [{ key: 'provider', type: 'select', label: 'Dịch vụ mặc định', value: c.provider, options: Object.keys(PROVIDERS).map(k => [k, PROVIDERS[k].name]).concat(Object.keys(WEB).map(k => [k, WEB[k].name])) }];
  Object.keys(PROVIDERS).filter(k => k !== 'offline').forEach(k => {
    const P = PROVIDERS[k];
    if (!P.nokey || P.optkey) fields.push({ key: 'key_' + k, label: (P.optkey ? 'Token proxy (không bắt buộc) — ' : 'Khóa API — ') + P.name, value: c.keys[k] || '', type: 'password' });
    fields.push({ key: 'model_' + k, label: 'Mô hình — ' + P.name, value: c.models[k] || P.model });
    if (P.local || P.custom) fields.push({ key: 'url_' + k, label: 'Địa chỉ máy chủ — ' + P.name, value: c.urls[k] || P.url });
  });
  const v = await formDlg('Cài đặt A.I (khóa chỉ lưu trên máy này)', fields, { cols: true });
  if (!v) return;
  c.provider = v.provider;
  Object.keys(PROVIDERS).forEach(k => {
    if (v['key_' + k] != null) {
      c.keys[k] = v['key_' + k].trim();
      persistKey(k, c.keys[k]);
    }
    if (v['model_' + k] != null) c.models[k] = v['model_' + k].trim();
    if (v['url_' + k] != null) c.urls[k] = v['url_' + k].trim();
  });
  saveSettings({ ai: c });
  try { Host.saveConfig({ ai_keys: c.keys }); } catch (e) {}
  toast('Đã lưu khóa API và cài đặt A.I vĩnh viễn trên máy này!', 'ok');
}

// ---------------------------------------------------------------- proofreading (Word/PPT/any doc) — offline, with A.I upgrade if configured
export function proofread(doc) {
  const ctx = ctxOf(doc); const text = ctx.selection || ctx.text;
  if (!text || !text.trim()) return toast('Không có văn bản để soát', 'err');
  const issues = O.proofreadText(text);
  const list = h('div.pf-list');
  if (!issues.length) list.appendChild(h('div.muted', '✅ Không phát hiện lỗi chính tả/định dạng phổ biến.'));
  issues.slice(0, 200).forEach(x => list.appendChild(h('div.pf-item', h('span.pf-bad', x.found), ' → ', h('b.pf-good', x.suggestion), h('span.muted', '  ' + x.reason))));
  const canFix = doc && doc.kind === 'word' && !ctx.selection;
  dialog({ title: 'Soát lỗi chính tả & ngữ pháp (ngoại tuyến) — ' + issues.length + ' vấn đề', width: 620, body: h('div', list), buttons: [
    issues.length && canFix ? { label: 'Sửa tất cả', primary: true, run: () => { fixWord(doc, issues); } } : null,
    issues.length && ctx.selection ? { label: 'Thay đoạn chọn', primary: true, run: () => { document.execCommand('insertText', false, O.applyFixes(text, issues)); doc && doc.markDirty && doc.markDirty(); } } : null,
    cfg().provider !== 'offline' && !WEB[cfg().provider] ? { label: 'Soát sâu bằng A.I', run: () => { openAi(doc.shell, doc, 'Soát lỗi chính tả và ngữ pháp, trả về bản đã sửa:\n' + text.slice(0, 6000)); } } : null,
    { label: 'Đóng' }
  ].filter(Boolean) });
}
function fixWord(doc, issues) {
  // apply per text node so formatting is preserved
  const walker = document.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT, null, false); let n, count = 0;
  doc.pushHistory && doc.pushHistory();
  while ((n = walker.nextNode())) { const iss = O.proofreadText(n.nodeValue); if (iss.length) { n.nodeValue = O.applyFixes(n.nodeValue, iss); count += iss.length; } }
  doc.markDirty(); toast('Đã sửa ' + count + ' lỗi', 'ok');
}
