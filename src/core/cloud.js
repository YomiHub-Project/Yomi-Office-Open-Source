// v0.0.8 — Cloud sync (mục XVI: Rclone, Google Workspace/Drive, OneDrive, Dropbox, MediaFire)
// * Đăng nhập OAuth 2.0 + PKCE kiểu rclone: trình duyệt hệ thống → chuyển về http://127.0.0.1:53682/ (máy chủ nhận mã chạy
//   tạm trong ứng dụng: Windows = Launcher, Android = Bridge). Client ID do chủ ứng dụng đăng ký, nhập trong Tùy chọn → Đám mây.
// * Mở / Lưu / Lưu thành lên Google Drive (gồm Google Docs/Sheets/Slides: xuất ↔ nhập .docx/.xlsx/.pptx), OneDrive, Dropbox.
// * Tự đồng bộ: Ctrl+S tải lên; kiểm tra phiên bản trên mây (etag / rev / version) để phát hiện xung đột;
//   cứ 20 giây xem tệp trên mây có ai sửa không — nếu tài liệu chưa sửa thì tự tải bản mới, nếu đang sửa thì hỏi.
// * "Sửa cùng người khác": mở đúng tệp trong Google Docs / Office cho web (đồng tác giả thời gian thực của dịch vụ đó)
//   và tạo liên kết chia sẻ có quyền sửa.
// * MediaFire: đọc thư mục công khai (API 1.5 không cần đăng nhập) và tải tệp về để mở.
// Mọi lệnh HTTP đi qua Host.http (Windows/Android gọi trực tiếp, không bị CORS; token không rời khỏi máy).
import { h, ic, $, $$, esc, fmtSize, extOf, baseName, b64FromBytes, bytesFromB64, utf8Encode, utf8Decode, relTime } from './util.js';
import { toast, dialog, formDlg, confirmDlg, promptDlg, busy } from './ui.js';
import { settings, saveSettings } from './store.js';
import { qrSvg } from './insert.js';
import Host from './host.js';

export const PORT = 53682;
const MIME = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', pdf: 'application/pdf', csv: 'text/csv', txt: 'text/plain', odt: 'application/vnd.oasis.opendocument.text', ods: 'application/vnd.oasis.opendocument.spreadsheet', odp: 'application/vnd.oasis.opendocument.presentation', md: 'text/markdown', rtf: 'application/rtf', docm: 'application/vnd.ms-word.document.macroEnabled.12', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12', pptm: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12' };
const GNATIVE = { 'application/vnd.google-apps.document': 'docx', 'application/vnd.google-apps.spreadsheet': 'xlsx', 'application/vnd.google-apps.presentation': 'pptx' };
const OPENABLE = /\.(docx?|docm|dotx|xlsx?|xlsm|xltx|csv|tsv|pptx?|pptm|ppsx|potx|pdf|odt|ods|odp|rtf|txt|md|html?|epub|fb2|png|jpe?g)$/i;

// ---------------------------------------------------------------- settings
export function cfg() { const s = settings(); s.cloud = Object.assign({ gdrive: {}, onedrive: {}, dropbox: {}, mediafire: { folders: [] }, autoSync: true, poll: 20 }, s.cloud || {}); return s.cloud; }
function saveCfg() { saveSettings({ cloud: cfg() }); }

// ---------------------------------------------------------------- PKCE + OAuth (loopback redirect, rclone style)
function rnd(n) { const a = new Uint8Array(n); (window.crypto || window.msCrypto).getRandomValues(a); return a; }
function b64url(bytes) { return b64FromBytes(bytes).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); }
async function sha256(str) {
  const data = utf8Encode(str);
  if (window.crypto && crypto.subtle && crypto.subtle.digest) { try { return new Uint8Array(await crypto.subtle.digest('SHA-256', data)); } catch (e) {} }
  return sha256js(data); // http://127.0.0.1 is a secure context, but old WebViews may lack SubtleCrypto
}
function sha256js(m) {
  const K = [0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
  const H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  const l = m.length, nb = ((l + 9 + 63) >> 6) << 6, b = new Uint8Array(nb); b.set(m); b[l] = 0x80; const bits = l * 8; b[nb - 4] = bits >>> 24; b[nb - 3] = bits >>> 16 & 255; b[nb - 2] = bits >>> 8 & 255; b[nb - 1] = bits & 255; b[nb - 5] = Math.floor(bits / 4294967296) & 255;
  const w = new Array(64); const r = (x, n) => (x >>> n) | (x << (32 - n));
  for (let o = 0; o < nb; o += 64) {
    for (let i = 0; i < 16; i++) w[i] = (b[o + i * 4] << 24) | (b[o + i * 4 + 1] << 16) | (b[o + i * 4 + 2] << 8) | b[o + i * 4 + 3];
    for (let i = 16; i < 64; i++) { const s0 = r(w[i - 15], 7) ^ r(w[i - 15], 18) ^ (w[i - 15] >>> 3), s1 = r(w[i - 2], 17) ^ r(w[i - 2], 19) ^ (w[i - 2] >>> 10); w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0; }
    let [a, bb, c, d, e, f, g, hh] = H;
    for (let i = 0; i < 64; i++) { const S1 = r(e, 6) ^ r(e, 11) ^ r(e, 25), ch = (e & f) ^ (~e & g), t1 = (hh + S1 + ch + K[i] + w[i]) | 0, S0 = r(a, 2) ^ r(a, 13) ^ r(a, 22), mj = (a & bb) ^ (a & c) ^ (bb & c), t2 = (S0 + mj) | 0; hh = g; g = f; f = e; e = (d + t1) | 0; d = c; c = bb; bb = a; a = (t1 + t2) | 0; }
    H[0] = (H[0] + a) | 0; H[1] = (H[1] + bb) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0; H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + hh) | 0;
  }
  const out = new Uint8Array(32); H.forEach((x, i) => { out[i * 4] = x >>> 24; out[i * 4 + 1] = x >>> 16 & 255; out[i * 4 + 2] = x >>> 8 & 255; out[i * 4 + 3] = x & 255; }); return out;
}
export const _sha256js = sha256js;
function form(o) { return Object.keys(o).filter(k => o[k] != null && o[k] !== '').map(k => encodeURIComponent(k) + '=' + encodeURIComponent(o[k])).join('&'); }

// ---------------------------------------------------------------- providers
export const PROVIDERS = {
  gdrive: {
    name: 'Google Drive / Workspace', icon: 'hard-drive', color: '#1a73e8', secret: true,
    redirect: 'http://127.0.0.1:' + PORT + '/',
    help: 'console.cloud.google.com → API và Dịch vụ → bật "Google Drive API" → Thông tin xác thực → Tạo mã ứng dụng OAuth loại "Ứng dụng dành cho máy tính" (Desktop app). Dán Client ID và Client secret vào đây.',
    auth: (c, ch, st) => 'https://accounts.google.com/o/oauth2/v2/auth?' + form({ client_id: c.clientId, redirect_uri: PROVIDERS.gdrive.redirect, response_type: 'code', scope: 'https://www.googleapis.com/auth/drive openid email', code_challenge: ch, code_challenge_method: 'S256', access_type: 'offline', prompt: 'consent', state: st }),
    token: 'https://oauth2.googleapis.com/token',
    async me(t) { const j = await api('gdrive', 'GET', 'https://www.googleapis.com/drive/v3/about?fields=user(displayName,emailAddress)'); return j.user ? (j.user.emailAddress || j.user.displayName) : ''; },
    async list(folder) {
      const q = folder === 'shared' ? 'sharedWithMe=true and trashed=false' : "'" + (folder || 'root') + "' in parents and trashed=false";
      const j = await api('gdrive', 'GET', 'https://www.googleapis.com/drive/v3/files?' + form({ q, fields: 'files(id,name,mimeType,modifiedTime,size,version,webViewLink,lastModifyingUser(displayName)),nextPageToken', pageSize: 300, orderBy: 'folder,name', supportsAllDrives: 'true', includeItemsFromAllDrives: 'true' }));
      const out = (j.files || []).map(f => ({ id: f.id, name: f.name + (GNATIVE[f.mimeType] && !/\.\w{3,4}$/.test(f.name) ? '.' + GNATIVE[f.mimeType] : ''), folder: f.mimeType === 'application/vnd.google-apps.folder', native: GNATIVE[f.mimeType] || null, mime: f.mimeType, size: +f.size || 0, modified: f.modifiedTime, rev: String(f.version || ''), web: f.webViewLink, by: f.lastModifyingUser && f.lastModifyingUser.displayName }));
      if (!folder || folder === 'root') out.unshift({ id: 'shared', name: 'Được chia sẻ với tôi', folder: true, virtual: true });
      return out;
    },
    async meta(id) { const f = await api('gdrive', 'GET', 'https://www.googleapis.com/drive/v3/files/' + id + '?fields=id,name,mimeType,version,modifiedTime,webViewLink,lastModifyingUser(displayName)&supportsAllDrives=true'); return { id: f.id, name: f.name, rev: String(f.version || ''), modified: f.modifiedTime, web: f.webViewLink, by: f.lastModifyingUser && f.lastModifyingUser.displayName, native: GNATIVE[f.mimeType] || null }; },
    async download(it) { return it.native ? raw('gdrive', 'GET', 'https://www.googleapis.com/drive/v3/files/' + it.id + '/export?mimeType=' + encodeURIComponent(MIME[it.native])) : raw('gdrive', 'GET', 'https://www.googleapis.com/drive/v3/files/' + it.id + '?alt=media&supportsAllDrives=true'); },
    async upload(bind, bytes, name, parent) {
      const ext = extOf(name); const ct = MIME[ext] || 'application/octet-stream';
      let id = bind && bind.id;
      if (!id) { const j = await api('gdrive', 'POST', 'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&fields=id', JSON.stringify({ name, parents: parent && parent !== 'root' && parent !== 'shared' ? [parent] : undefined }), { 'Content-Type': 'application/json' }); id = j.id; }
      // PATCH media (Google-native Docs/Sheets/Slides are converted back from the Office file by Drive)
      await raw('gdrive', 'PATCH', 'https://www.googleapis.com/upload/drive/v3/files/' + id + '?uploadType=media&supportsAllDrives=true', bytes, { 'Content-Type': ct });
      return this.meta(id);
    },
    async share(it) { await api('gdrive', 'POST', 'https://www.googleapis.com/drive/v3/files/' + it.id + '/permissions?supportsAllDrives=true', JSON.stringify({ role: 'writer', type: 'anyone' }), { 'Content-Type': 'application/json' }); const m = await this.meta(it.id); return m.web; },
    editUrl(it) { return it.web || 'https://drive.google.com/open?id=' + it.id; }
  },
  onedrive: {
    name: 'OneDrive / Microsoft 365', icon: 'cloud', color: '#0364b8', secret: false,
    redirect: 'http://localhost:' + PORT + '/',
    help: 'portal.azure.com → Microsoft Entra ID → Đăng ký ứng dụng → Mới: chọn "Tài khoản ở mọi tổ chức và tài khoản Microsoft cá nhân"; Nền tảng "Ứng dụng di động và máy tính" với URI chuyển hướng http://localhost; Quyền API: Files.ReadWrite.All, User.Read, offline_access. Dán Application (client) ID vào đây.',
    auth: (c, ch, st) => 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?' + form({ client_id: c.clientId, redirect_uri: PROVIDERS.onedrive.redirect, response_type: 'code', response_mode: 'query', scope: 'Files.ReadWrite.All User.Read offline_access', code_challenge: ch, code_challenge_method: 'S256', state: st, prompt: 'select_account' }),
    token: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    async me() { const j = await api('onedrive', 'GET', 'https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName,mail'); return j.mail || j.userPrincipalName || j.displayName || ''; },
    async list(folder) {
      const url = folder === 'shared' ? 'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe' : 'https://graph.microsoft.com/v1.0/me/drive/' + (folder && folder !== 'root' ? 'items/' + folder : 'root') + '/children?$top=500&$select=id,name,folder,file,size,lastModifiedDateTime,eTag,webUrl,parentReference,remoteItem,lastModifiedBy';
      const j = await api('onedrive', 'GET', url);
      const out = (j.value || []).map(f => { const r = f.remoteItem || f; return { id: (r.parentReference && r.parentReference.driveId && f.remoteItem ? r.parentReference.driveId + '!' : '') + r.id, name: f.name, folder: !!(r.folder || f.folder), size: r.size || f.size || 0, modified: f.lastModifiedDateTime, rev: f.eTag || '', web: f.webUrl, by: f.lastModifiedBy && f.lastModifiedBy.user && f.lastModifiedBy.user.displayName }; });
      if (!folder || folder === 'root') out.unshift({ id: 'shared', name: 'Được chia sẻ với tôi', folder: true, virtual: true });
      return out;
    },
    item(id) { const k = String(id).indexOf('!'); return k > 0 && /^[a-zA-Z0-9]+$/.test(id.slice(0, k)) && id.length - k > 10 && id.split('!').length === 2 && id.slice(k + 1).indexOf('!') < 0 && !/^[0-9A-F]+$/.test(id.slice(0, k)) ? 'https://graph.microsoft.com/v1.0/drives/' + id.slice(0, k) + '/items/' + id.slice(k + 1) : 'https://graph.microsoft.com/v1.0/me/drive/items/' + id; },
    async meta(id) { const f = await api('onedrive', 'GET', this.item(id) + '?$select=id,name,eTag,lastModifiedDateTime,webUrl,lastModifiedBy'); return { id, name: f.name, rev: f.eTag, modified: f.lastModifiedDateTime, web: f.webUrl, by: f.lastModifiedBy && f.lastModifiedBy.user && f.lastModifiedBy.user.displayName }; },
    async download(it) { return raw('onedrive', 'GET', this.item(it.id) + '/content'); },
    async upload(bind, bytes, name, parent) {
      const ct = MIME[extOf(name)] || 'application/octet-stream';
      const url = bind && bind.id ? this.item(bind.id) + '/content' : 'https://graph.microsoft.com/v1.0/me/drive/' + (parent && parent !== 'root' && parent !== 'shared' ? 'items/' + parent : 'root') + ':/' + encodeURIComponent(name) + ':/content?@microsoft.graph.conflictBehavior=rename';
      const hd = { 'Content-Type': ct }; if (bind && bind.rev && !bind.force) hd['If-Match'] = bind.rev;
      const r = await raw('onedrive', 'PUT', url, bytes, hd, true); const j = JSON.parse(utf8Decode(r));
      return { id: bind && bind.id ? bind.id : j.id, name: j.name, rev: j.eTag, modified: j.lastModifiedDateTime, web: j.webUrl };
    },
    async share(it) { const j = await api('onedrive', 'POST', this.item(it.id) + '/createLink', JSON.stringify({ type: 'edit', scope: 'anonymous' }), { 'Content-Type': 'application/json' }); return j.link && j.link.webUrl; },
    editUrl(it) { return it.web; }
  },
  dropbox: {
    name: 'Dropbox', icon: 'package', color: '#0061fe', secret: false,
    redirect: 'http://127.0.0.1:' + PORT + '/',
    help: 'dropbox.com/developers/apps → Create app → Scoped access, Full Dropbox → tab Permissions: bật files.metadata.read/write, files.content.read/write, sharing.write → tab Settings: thêm Redirect URI http://127.0.0.1:' + PORT + '/ và bật "Allow public clients (PKCE)". Dán App key vào đây.',
    auth: (c, ch, st) => 'https://www.dropbox.com/oauth2/authorize?' + form({ client_id: c.clientId, redirect_uri: PROVIDERS.dropbox.redirect, response_type: 'code', code_challenge: ch, code_challenge_method: 'S256', token_access_type: 'offline', state: st }),
    token: 'https://api.dropboxapi.com/oauth2/token',
    async me() { const j = await api('dropbox', 'POST', 'https://api.dropboxapi.com/2/users/get_current_account'); return j.email || (j.name && j.name.display_name) || ''; },
    async list(folder) {
      let j = await api('dropbox', 'POST', 'https://api.dropboxapi.com/2/files/list_folder', JSON.stringify({ path: folder && folder !== 'root' ? folder : '', limit: 2000 }), { 'Content-Type': 'application/json' });
      let all = j.entries || []; while (j.has_more) { j = await api('dropbox', 'POST', 'https://api.dropboxapi.com/2/files/list_folder/continue', JSON.stringify({ cursor: j.cursor }), { 'Content-Type': 'application/json' }); all = all.concat(j.entries || []); }
      return all.map(f => ({ id: f.id, path: f.path_lower, name: f.name, folder: f['.tag'] === 'folder', size: f.size || 0, modified: f.server_modified, rev: f.rev || '' })).sort((a, b) => (b.folder - a.folder) || a.name.localeCompare(b.name));
    },
    async meta(id) { const f = await api('dropbox', 'POST', 'https://api.dropboxapi.com/2/files/get_metadata', JSON.stringify({ path: id }), { 'Content-Type': 'application/json' }); return { id: f.id, name: f.name, rev: f.rev, modified: f.server_modified, path: f.path_lower }; },
    async download(it) { return raw('dropbox', 'POST', 'https://content.dropboxapi.com/2/files/download', null, { 'Dropbox-API-Arg': asciiJson({ path: it.id }) }); },
    async upload(bind, bytes, name, parent) {
      let path, mode;
      if (bind && bind.id) { const m = await this.meta(bind.id); path = m.path; mode = bind.rev && !bind.force ? { '.tag': 'update', update: bind.rev } : 'overwrite'; }
      else { path = (parent && parent !== 'root' ? (await this.meta(parent)).path : '') + '/' + name; mode = 'add'; }
      const r = await raw('dropbox', 'POST', 'https://content.dropboxapi.com/2/files/upload', bytes, { 'Content-Type': 'application/octet-stream', 'Dropbox-API-Arg': asciiJson({ path, mode, autorename: !(bind && bind.id), strict_conflict: !!(bind && bind.rev && !bind.force) }) }, true);
      const j = JSON.parse(utf8Decode(r)); return { id: j.id, name: j.name, rev: j.rev, modified: j.server_modified, path: j.path_lower };
    },
    async share(it) { try { const j = await api('dropbox', 'POST', 'https://api.dropboxapi.com/2/sharing/create_shared_link_with_settings', JSON.stringify({ path: it.id, settings: { requested_visibility: 'public', access: 'editor' } }), { 'Content-Type': 'application/json' }); return j.url; } catch (e) { const j = await api('dropbox', 'POST', 'https://api.dropboxapi.com/2/sharing/list_shared_links', JSON.stringify({ path: it.id, direct_only: true }), { 'Content-Type': 'application/json' }); return j.links && j.links[0] && j.links[0].url; } },
    editUrl(it) { return 'https://www.dropbox.com/home' + (it.path ? it.path.replace(/\/[^/]*$/, '') : '') + '?preview=' + encodeURIComponent(it.name || ''); }
  }
};
function asciiJson(o) { return JSON.stringify(o).replace(/[\u007f-￿]/g, c => '\\u' + ('0000' + c.charCodeAt(0).toString(16)).slice(-4)); }

export function isConnected(p) { const c = cfg()[p]; return !!(c && c.refresh || c && c.access); }
export async function connect(p) {
  const P = PROVIDERS[p]; const c = cfg()[p];
  if (!c.clientId) { await accountDialog(p); if (!cfg()[p].clientId) return false; }
  if (Host.kind === 'web') { toast('Đăng nhập đám mây cần bản cài đặt Windows hoặc Android (trình duyệt chặn máy chủ nhận mã 127.0.0.1)', 'err', 6000); return false; }
  const verifier = b64url(rnd(48)); const challenge = b64url(await sha256(verifier)); const state = b64url(rnd(12));
  const b = busy('Đang mở trình duyệt để đăng nhập ' + P.name + '… Hoàn tất đăng nhập rồi quay lại ứng dụng.');
  let q;
  try { q = await Host.oauth(P.auth(c, challenge, state), PORT); } catch (e) { b.done(); toast('Đăng nhập thất bại: ' + e.message, 'err', 6000); return false; }
  b.done();
  const qs = parseQs(q); if (qs.error) { toast('Đăng nhập bị từ chối: ' + (qs.error_description || qs.error), 'err', 6000); return false; }
  if (!qs.code || qs.state !== state) { toast('Phản hồi đăng nhập không hợp lệ', 'err'); return false; }
  const tok = await tokenReq(p, { grant_type: 'authorization_code', code: qs.code, redirect_uri: P.redirect, code_verifier: verifier });
  setTok(p, tok); try { c.user = await P.me(); } catch (e) {} saveCfg();
  toast('Đã kết nối ' + P.name + (c.user ? ' — ' + c.user : ''), 'ok'); return true;
}
export function disconnect(p) { const c = cfg()[p]; delete c.access; delete c.refresh; delete c.exp; delete c.user; saveCfg(); toast('Đã đăng xuất ' + PROVIDERS[p].name); }
function parseQs(q) { const o = {}; String(q || '').replace(/^.*?\?/, '').split('&').forEach(kv => { if (!kv) return; const i = kv.indexOf('='); o[decodeURIComponent(kv.slice(0, i < 0 ? kv.length : i))] = i < 0 ? '' : decodeURIComponent(kv.slice(i + 1).replace(/\+/g, ' ')); }); return o; }
async function tokenReq(p, body) {
  const P = PROVIDERS[p], c = cfg()[p];
  const r = await Host.http({ url: P.token, method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: form(Object.assign({ client_id: c.clientId, client_secret: P.secret ? c.secret : null }, body)) });
  const j = JSON.parse(r.text || '{}'); if (r.status >= 400 || !j.access_token) throw new Error(j.error_description || j.error || ('HTTP ' + r.status)); return j;
}
function setTok(p, j) { const c = cfg()[p]; c.access = j.access_token; if (j.refresh_token) c.refresh = j.refresh_token; c.exp = Date.now() + ((+j.expires_in || 3600) - 60) * 1000; saveCfg(); }
async function access(p) {
  const c = cfg()[p]; if (c.access && c.exp > Date.now()) return c.access;
  if (!c.refresh) throw new Error('Chưa đăng nhập ' + PROVIDERS[p].name);
  setTok(p, await tokenReq(p, { grant_type: 'refresh_token', refresh_token: c.refresh })); return c.access;
}
async function raw(p, method, url, body, headers, wantConflict) {
  for (let attempt = 0; attempt < 2; attempt++) {
    const hd = Object.assign({ Authorization: 'Bearer ' + await access(p) }, headers || {});
    const r = await Host.http({ url, method, headers: hd, bytes: body instanceof Uint8Array ? body : null, body: typeof body === 'string' ? body : null, binary: true });
    if (r.status === 401 && attempt === 0) { cfg()[p].exp = 0; continue; }
    if (r.status === 412 || (r.status === 409 && /conflict/i.test(utf8Decode(r.bytes || new Uint8Array(0))))) { const e = new Error('Tệp trên đám mây đã bị người khác thay đổi'); e.conflict = true; throw e; }
    if (r.status >= 400) { let msg = ''; try { const t = utf8Decode(r.bytes); const j = JSON.parse(t); msg = (j.error && (j.error.message || j.error.code || j.error)) || j.error_summary || t.slice(0, 200); } catch (e) {} throw new Error('HTTP ' + r.status + (msg ? ': ' + (typeof msg === 'string' ? msg : JSON.stringify(msg)) : '')); }
    return r.bytes || new Uint8Array(0);
  }
}
async function api(p, method, url, body, headers) { const b = await raw(p, method, url, body, headers); const t = utf8Decode(b); return t ? JSON.parse(t) : {}; }

// ---------------------------------------------------------------- MediaFire (public folders, no login)
export async function mediafireList(key) {
  const base = 'https://www.mediafire.com/api/1.5/folder/get_content.php?response_format=json&folder_key=' + encodeURIComponent(key);
  const fj = JSON.parse(await Host.httpGet(base + '&content_type=files&chunk_size=1000')); let dj = { response: {} };
  try { dj = JSON.parse(await Host.httpGet(base + '&content_type=folders')); } catch (e) {}
  return (((dj.response.folder_content || {}).folders) || []).map(d => ({ id: d.folderkey, name: d.name, folder: true }))
    .concat((((fj.response.folder_content || {}).files) || []).map(f => ({ id: f.quickkey, name: f.filename, size: +f.size || 0, modified: f.created, page: (f.links && f.links.normal_download) || ('https://www.mediafire.com/file/' + f.quickkey) })));
}
export function mediafireKey(s) { const m = /mediafire\.com\/folder\/([a-z0-9]+)/i.exec(s || '') || /^([a-z0-9]{10,16})$/i.exec((s || '').trim()); return m ? m[1] : null; }
async function mediafireDownload(it) {
  const U = await import('./update.js'); const html = await Host.httpGet(it.page); const d = U.mediafireDirect(html); if (!d) throw new Error('MediaFire: không tìm thấy liên kết tải');
  const r = await Host.http({ url: d, method: 'GET', binary: true }); if (r.status >= 400) throw new Error('HTTP ' + r.status); return r.bytes;
}

// ---------------------------------------------------------------- browser dialog (open / pick folder)
export async function cloudOpenDialog(shell) {
  const avail = Object.keys(PROVIDERS).filter(isConnected);
  const tiles = h('div.cl-provs');
  const body = h('div.cl-wrap', tiles, h('div.cl-browser'));
  const open = (p) => browse(shell, body.lastChild, p, 'open');
  Object.keys(PROVIDERS).forEach(p => { const P = PROVIDERS[p]; const on = isConnected(p); const b = h('button.cl-prov' + (on ? '.on' : ''), h('span.cl-ic', { style: { background: P.color } }, ic(P.icon, 16)), h('div', h('b', P.name), h('div.muted', on ? (cfg()[p].user || 'Đã kết nối') : 'Chưa kết nối — bấm để đăng nhập'))); b.addEventListener('click', async () => { if (!isConnected(p) && !(await connect(p))) return; $$('.cl-prov', tiles).forEach(x => x.classList.remove('sel')); b.classList.add('sel'); open(p); }); tiles.appendChild(b); });
  const mf = h('button.cl-prov', h('span.cl-ic', { style: { background: '#1d6ff2' } }, ic('flame', 16)), h('div', h('b', 'MediaFire (thư mục công khai)'), h('div.muted', 'Dán liên kết thư mục để duyệt và tải tệp')));
  mf.addEventListener('click', async () => { const v = await promptDlg('Liên kết thư mục MediaFire (vd: https://www.mediafire.com/folder/abc123/Ten):', (cfg().mediafire.folders || [])[0] || 'https://www.mediafire.com/folder/seqly6bv55yr8/Office_Open_Source', 'MediaFire'); const k = mediafireKey(v); if (!k) return; const c = cfg(); c.mediafire.folders = [v].concat((c.mediafire.folders || []).filter(x => x !== v)).slice(0, 8); saveCfg(); $$('.cl-prov', tiles).forEach(x => x.classList.remove('sel')); mf.classList.add('sel'); browse(shell, body.lastChild, 'mediafire', 'open', k); });
  tiles.appendChild(mf);
  const set = h('button.btn', { style: { marginTop: '8px' } }, ic('settings', 13), ' Tài khoản & Client ID…'); set.addEventListener('click', () => accountsDialog()); tiles.appendChild(set);
  dialog({ title: 'Mở từ đám mây', width: 860, body, buttons: [{ label: 'Đóng' }] });
  if (avail.length) setTimeout(() => { const b = $$('.cl-prov.on', tiles)[0]; if (b) b.click(); }, 50);
}
async function browse(shell, box, p, mode, start, onPick) {
  const stack = [{ id: start || 'root', name: p === 'mediafire' ? 'MediaFire' : PROVIDERS[p].name }];
  const render = async () => {
    box.innerHTML = ''; const cur = stack[stack.length - 1];
    const crumbs = h('div.cl-crumbs'); stack.forEach((s, i) => { const c = h('button.cl-crumb', s.name); c.addEventListener('click', () => { stack.length = i + 1; render(); }); crumbs.appendChild(c); if (i < stack.length - 1) crumbs.appendChild(ic('chevron-right', 12)); });
    box.appendChild(crumbs);
    const list = h('div.cl-list', h('div.muted', 'Đang tải…')); box.appendChild(list);
    let items; try { items = p === 'mediafire' ? await mediafireList(cur.id) : await PROVIDERS[p].list(cur.id); } catch (e) { list.innerHTML = ''; list.appendChild(h('div.cl-err', ic('circle-alert', 14), ' ', e.message)); return; }
    list.innerHTML = '';
    if (mode === 'folder') { const here = h('button.btn.primary', ic('folder-check', 13), ' Lưu vào thư mục này'); here.addEventListener('click', () => onPick({ provider: p, parent: cur.id === 'shared' ? 'root' : cur.id, folderName: cur.name })); box.insertBefore(here, list); }
    if (!items.length) list.appendChild(h('div.muted', 'Thư mục trống'));
    items.forEach(it => {
      if (mode === 'folder' && !it.folder) return;
      const ok = it.folder || OPENABLE.test(it.name);
      const row = h('button.cl-item' + (ok ? '' : '.dim'), ic(it.folder ? 'folder' : fileIcon(it.name), 16), h('span.cl-name', it.name), h('span.muted.cl-meta', it.folder ? '' : (it.size ? fmtSize(it.size) : '') + (it.modified ? ' · ' + relTime(Date.parse(it.modified)) : '') + (it.by ? ' · ' + it.by : '')));
      row.addEventListener('click', async () => {
        if (it.folder) { stack.push({ id: it.id, name: it.name }); return render(); }
        if (!ok) return toast('Định dạng này chưa mở được');
        const b = busy('Đang tải ' + it.name + '…');
        try {
          const bytes = p === 'mediafire' ? await mediafireDownload(it) : await PROVIDERS[p].download(it);
          b.done(); $$('.dlg-back').forEach(x => { const c = x.querySelector('.dlg-x'); if (c) c.click(); });
          const f = { name: it.name, bytes, here: true }; const doc = await shell.openFile(f);
          if (doc && p !== 'mediafire') bind(doc, { provider: p, id: it.id, name: it.name, rev: it.rev, web: it.web, path: it.path, native: it.native, parent: cur.id });
        } catch (e) { b.done(); toast('Không tải được: ' + e.message, 'err', 6000); }
      });
      list.appendChild(row);
    });
  };
  render();
}
function fileIcon(n) { return /\.(docx?|docm|odt|rtf)$/i.test(n) ? 'file-type-2' : /\.(xlsx?|xlsm|csv|ods)$/i.test(n) ? 'sheet' : /\.(pptx?|pptm|odp)$/i.test(n) ? 'presentation' : /\.pdf$/i.test(n) ? 'file-type' : 'file'; }

// ---------------------------------------------------------------- save to cloud (Lưu thành → Đám mây)
export async function cloudSaveDialog(shell, doc) {
  const provs = Object.keys(PROVIDERS);
  const tiles = h('div.cl-provs'); const body = h('div.cl-wrap', tiles, h('div.cl-browser', h('div.muted', 'Chọn dịch vụ đám mây ở bên trái.')));
  let close;
  const done = async pick => {
    const ext = doc.fileExt && (!doc.canSaveExt || doc.canSaveExt(doc.fileExt)) ? doc.fileExt : doc.meta.newExt;
    const name = await promptDlg('Tên tệp trên ' + PROVIDERS[pick.provider].name + ':', (doc.name || 'Tài liệu') + '.' + ext, 'Lưu lên đám mây'); if (!name) return;
    if (close) close();
    const b = busy('Đang tải lên…');
    try { const bytes = await doc.save(extOf(name) || ext); const m = await PROVIDERS[pick.provider].upload(null, bytes, name, pick.parent); b.done(); bind(doc, Object.assign({ provider: pick.provider, parent: pick.parent }, m)); doc.dirty = false; shell.updateTitle(); toast('Đã lưu lên ' + PROVIDERS[pick.provider].name + ': ' + m.name, 'ok'); }
    catch (e) { b.done(); toast('Không tải lên được: ' + e.message, 'err', 6000); }
  };
  provs.forEach(p => { const P = PROVIDERS[p]; const b = h('button.cl-prov' + (isConnected(p) ? '.on' : ''), h('span.cl-ic', { style: { background: P.color } }, ic(P.icon, 16)), h('div', h('b', P.name), h('div.muted', isConnected(p) ? (cfg()[p].user || 'Đã kết nối') : 'Chưa kết nối'))); b.addEventListener('click', async () => { if (!isConnected(p) && !(await connect(p))) return; browse(shell, body.lastChild, p, 'folder', 'root', done); }); tiles.appendChild(b); });
  const pr = dialog({ title: 'Lưu lên đám mây', width: 820, body, buttons: [{ label: 'Hủy' }] }); close = () => { $$('.dlg-back').forEach(x => { if (x.contains(body)) { const c = x.querySelector('.dlg-x'); if (c) c.click(); } }); };
  return pr;
}

// ---------------------------------------------------------------- binding a document to a cloud file: sync, conflicts, live pull
export function bind(doc, info) {
  doc.cloud = Object.assign({}, info); doc.cloud.since = Date.now();
  if (doc.statusLeft && !doc._cloudSb) { doc._cloudSb = h('button.sb-item.cl-sb'); doc._cloudSb.addEventListener('click', () => cloudMenu(doc)); doc.statusLeft.appendChild(doc._cloudSb); }
  status(doc, 'ok'); watch(doc);
}
function status(doc, st, msg) {
  if (!doc._cloudSb) return; const P = PROVIDERS[doc.cloud.provider];
  const icn = { ok: 'cloud-check', sync: 'refresh-cw', dirty: 'cloud-upload', err: 'cloud-off', conflict: 'triangle-alert', remote: 'cloud-download' }[st] || 'cloud';
  doc._cloudSb.innerHTML = ''; doc._cloudSb.appendChild(ic(icn, 13)); doc._cloudSb.appendChild(document.createTextNode(' ' + (msg || ({ ok: 'Đã đồng bộ', sync: 'Đang đồng bộ…', dirty: 'Chưa tải lên', err: 'Lỗi đồng bộ', conflict: 'Xung đột', remote: 'Có bản mới trên mây' }[st])) + ' · ' + P.name.split(' ')[0]));
  doc._cloudSb.className = 'sb-item cl-sb cl-' + st;
}
export async function syncUp(doc, force) {
  const c = doc.cloud; if (!c) return false; const P = PROVIDERS[c.provider];
  status(doc, 'sync');
  try {
    const ext = extOf(c.name) || doc.fileExt || doc.meta.newExt;
    const bytes = await doc.save(c.native || ext);
    const m = await P.upload(Object.assign({}, c, { force: !!force }), bytes, c.name, c.parent);
    Object.assign(c, { rev: m.rev, modified: m.modified, web: m.web || c.web, id: m.id || c.id }); c.since = Date.now();
    status(doc, 'ok'); return true;
  } catch (e) {
    if (e.conflict) { status(doc, 'conflict'); return resolveConflict(doc); }
    status(doc, 'err', 'Lỗi: ' + e.message.slice(0, 40)); toast('Không đồng bộ được: ' + e.message, 'err', 6000); return false;
  }
}
async function resolveConflict(doc) {
  const c = doc.cloud; let m = {}; try { m = await PROVIDERS[c.provider].meta(c.id); } catch (e) {}
  const v = await dialog({ title: 'Xung đột phiên bản', width: 520, body: h('div', h('p', 'Tệp "' + c.name + '" trên ' + PROVIDERS[c.provider].name + ' đã được ' + (m.by || 'người khác') + ' sửa' + (m.modified ? ' lúc ' + new Date(m.modified).toLocaleString('vi-VN') : '') + ' sau lần bạn mở.'), h('p.muted', 'Chọn cách xử lý:')),
    buttons: [{ label: 'Lấy bản trên mây', value: 'pull' }, { label: 'Lưu bản của tôi thành bản sao', value: 'copy' }, { label: 'Ghi đè bằng bản của tôi', primary: true, value: 'mine' }, { label: 'Để sau', value: null }] });
  if (v === 'mine') return syncUp(doc, true);
  if (v === 'pull') return pull(doc, true);
  if (v === 'copy') { const P = PROVIDERS[c.provider]; const ext = extOf(c.name); const name = baseName(c.name) + ' (bản của ' + (settings().userName || 'tôi') + ').' + ext; const bytes = await doc.save(c.native || ext); const nm = await P.upload(null, bytes, name, c.parent); Object.assign(c, nm, { name: nm.name || name }); status(doc, 'ok'); toast('Đã lưu bản sao: ' + (nm.name || name), 'ok'); return true; }
  return false;
}
async function pull(doc, forced) {
  const c = doc.cloud; const P = PROVIDERS[c.provider];
  try {
    const m = await P.meta(c.id); const bytes = await P.download(Object.assign({}, c, m));
    const snap = doc.saveScroll ? doc.saveScroll() : null;
    await doc.load(bytes, c.native || extOf(c.name), c.name); Object.assign(c, { rev: m.rev, modified: m.modified }); c.since = Date.now();
    doc.dirty = false; doc.shell.updateTitle(); doc.onShow && doc.onShow(); doc.refreshUi && doc.refreshUi();
    status(doc, 'ok', 'Đã cập nhật bản mới' + (m.by ? ' của ' + m.by : '')); if (forced) toast('Đã tải bản mới nhất từ ' + P.name, 'ok');
    return true;
  } catch (e) { status(doc, 'err'); toast('Không tải được bản mới: ' + e.message, 'err'); return false; }
}
function watch(doc) {
  if (doc._cloudTimer) clearInterval(doc._cloudTimer);
  const sec = Math.max(10, +cfg().poll || 20);
  doc._cloudTimer = setInterval(async () => {
    if (!doc.cloud || doc._cloudBusy) return; if (doc.shell.docs.indexOf(doc) < 0) { clearInterval(doc._cloudTimer); return; }
    doc._cloudBusy = true;
    try {
      const m = await PROVIDERS[doc.cloud.provider].meta(doc.cloud.id);
      if (m.rev && doc.cloud.rev && m.rev !== doc.cloud.rev) {
        if (!doc.dirty && cfg().autoSync) await pull(doc);
        else if (doc._remoteNotified !== m.rev) { doc._remoteNotified = m.rev; status(doc, 'remote', (m.by || 'Người khác') + ' vừa sửa'); toast((m.by || 'Một người dùng khác') + ' vừa cập nhật "' + doc.cloud.name + '" trên ' + PROVIDERS[doc.cloud.provider].name + '. Bấm biểu tượng đám mây ở thanh trạng thái để xử lý.', 'info', 7000); }
      } else if (doc.dirty) status(doc, 'dirty');
    } catch (e) {} finally { doc._cloudBusy = false; }
  }, sec * 1000);
}
export function cloudMenu(doc) {
  const c = doc.cloud; const P = PROVIDERS[c.provider];
  import('./ui.js').then(U => U.showMenu([
    { label: 'Đồng bộ ngay (tải lên)', icon: 'cloud-upload', run: () => syncUp(doc) },
    { label: 'Lấy bản mới nhất từ đám mây', icon: 'cloud-download', run: () => pull(doc, true) },
    { sep: true },
    { label: 'Sửa cùng người khác (mở trên web)', icon: 'users', run: () => coEdit(doc) },
    { label: 'Chia sẻ — tạo liên kết có quyền sửa', icon: 'link', run: () => shareLink(doc) },
    { sep: true },
    { label: 'Ngắt liên kết với đám mây', icon: 'unlink', run: () => { clearInterval(doc._cloudTimer); doc.cloud = null; if (doc._cloudSb) doc._cloudSb.parentNode.removeChild(doc._cloudSb); doc._cloudSb = null; toast('Tài liệu không còn đồng bộ với ' + P.name); } }
  ], { anchor: doc._cloudSb }));
}
export async function coEdit(doc) {
  const c = doc.cloud; if (!c) return toast('Tài liệu chưa lưu trên đám mây. Dùng Lưu → Đám mây trước.', 'err');
  if (doc.dirty) await syncUp(doc);
  const P = PROVIDERS[c.provider]; let url = P.editUrl(c); if (!url) { try { const m = await P.meta(c.id); c.web = m.web; url = P.editUrl(c); } catch (e) {} }
  if (!url) return toast('Không lấy được địa chỉ chỉnh sửa trên web', 'err');
  Host.openUrl(url); toast('Đã mở trên ' + P.name + ' — mọi người có quyền sửa đều cùng soạn thảo thời gian thực. Khi họ lưu, ứng dụng tự tải bản mới.', 'ok', 7000);
}
export async function shareLink(doc) {
  const c = doc.cloud; if (!c) return toast('Tài liệu chưa lưu trên đám mây', 'err');
  const b = busy('Đang tạo liên kết chia sẻ…'); try { const url = await PROVIDERS[c.provider].share(c); b.done(); Host.clipSet ? Host.clipSet(url) : null; try { await navigator.clipboard.writeText(url); } catch (e) {} dialog({ title: 'Liên kết chia sẻ (quyền chỉnh sửa)', width: 520, body: h('div', h('input.inp', { value: url, readonly: true, style: { width: '100%' } }), h('p.muted', 'Đã sao chép vào bộ nhớ tạm. Ai có liên kết đều mở và cùng sửa được.')), buttons: [{ label: 'Đóng', primary: true }] }); }
  catch (e) { b.done(); toast('Không tạo được liên kết: ' + e.message, 'err', 6000); }
}

export async function airDropDialog(doc) {
  // v0.1.6: real AirDrop / QuickShare over Wi-Fi — the document is served on port 8765 of this device; other phones / PCs on the
  // same network scan the QR code to download it, and can send files back ("Nhận tệp"). Android also offers the system share
  // sheet (Quick Share / Nearby Share, Bluetooth, Zalo…).
  if (!Host.lanShare) return toast('Thiết bị này chưa hỗ trợ chia sẻ LAN', 'err');
  const ext = doc && doc.saveFormats ? doc.saveFormats()[0][0] : (doc && doc.fileExt) || 'bin';
  const b = busy('Đang chuẩn bị chia sẻ…');
  let name = 'tai-lieu', bytes = null, r = null;
  try {
    if (doc) { bytes = doc.kind === 'pdf' ? doc.bytes : await doc.save(ext); name = (doc.name || 'tai-lieu') + (doc.kind === 'pdf' && /\.pdf$/i.test(doc.name || '') ? '' : '.' + (doc.kind === 'pdf' ? 'pdf' : ext)); }
    r = await Host.lanShare(name, bytes || new Uint8Array(0));
  } catch (e) { toast('Không chia sẻ được: ' + e.message, 'err'); return; } finally { b.done(); }
  if (!r || !r.ok) return toast('Không mở được cổng chia sẻ 8765 (tường lửa?)', 'err', 6000);
  const urls = r.urls || [], recv = r.receive || [];
  if (!urls.length) return toast('Thiết bị chưa kết nối Wi-Fi / LAN', 'err');
  let cur = 0;
  const qrBox = h('div.adrop-qr', { style: { width: '170px', height: '170px', border: '1px solid #ddd', borderRadius: '8px', padding: '6px', background: '#fff' } });
  const urlInp = h('input.inp', { readonly: true, style: { width: '100%', marginBottom: '8px' } });
  const show = () => { const u = mode === 'send' ? urls[cur] : recv[cur]; urlInp.value = u; qrBox.innerHTML = qrSvg(u, { fg: '#000', bg: '#fff' }); };
  let mode = doc ? 'send' : 'recv';
  const tabs = h('div', { style: { display: 'flex', gap: '6px', marginBottom: '10px' } },
    h('button.btn' + (mode === 'send' ? '.primary' : ''), { onclick: e => { mode = 'send'; [...tabs.children].forEach(x => x.classList.remove('primary')); e.currentTarget.classList.add('primary'); show(); } }, ic('upload', 14), ' Gửi tài liệu này'),
    h('button.btn' + (mode === 'recv' ? '.primary' : ''), { onclick: e => { mode = 'recv'; [...tabs.children].forEach(x => x.classList.remove('primary')); e.currentTarget.classList.add('primary'); show(); } }, ic('download', 14), ' Nhận tệp từ máy khác'));
  const status = h('div.hint', 'Tệp nhận được lưu tại: ' + (r.inbox || ''));
  const box = h('div.airdrop-box', tabs,
    h('div', { style: { display: 'flex', gap: '16px', alignItems: 'center' } }, qrBox,
      h('div', { style: { flex: '1' } },
        h('p.muted', 'Thiết bị nhận (cùng Wi-Fi) quét mã QR bằng Camera, hoặc mở liên kết trong trình duyệt:'), urlInp,
        h('div', { style: { display: 'flex', gap: '6px', flexWrap: 'wrap' } },
          h('button.btn', { onclick: () => { Host.clipSet ? Host.clipSet(urlInp.value) : null; try { navigator.clipboard.writeText(urlInp.value); } catch (e) {} toast('Đã sao chép liên kết'); } }, ic('copy', 14), ' Sao chép'),
          urls.length > 1 ? h('button.btn', { onclick: () => { cur = (cur + 1) % urls.length; show(); } }, ic('wifi', 14), ' Mạng khác') : null,
          doc && Host.kind === 'android' ? h('button.btn', { onclick: () => Host.share(name, bytes) }, ic('share-2', 14), ' Quick Share / Bluetooth…') : null))),
    status);
  show();
  const poll = setInterval(async () => { try { const got = await Host.lanReceived(); if (got && got.length) { status.textContent = 'Đã nhận: ' + got.map(p => p.split(/[\\/]/).pop()).join(', '); toast('Đã nhận ' + got.length + ' tệp qua Wi-Fi', 'ok'); got.forEach(p => Host.readPath(p).then(f => f && window.OOS && window.OOS.openFile(f))); } } catch (e) {} }, 2000);
  dialog({ title: 'AirDrop & QuickShare — Chia sẻ qua Wi-Fi / LAN', width: 600, body: box, buttons: [{ label: 'Đóng', primary: true }] }).then(() => clearInterval(poll));
  setTimeout(() => { if (!document.body.contains(box)) clearInterval(poll); }, 600000);
}

// ---------------------------------------------------------------- accounts (Tùy chọn → Đám mây)
export async function accountDialog(p) {
  const P = PROVIDERS[p]; const c = cfg()[p];
  const f = [{ key: 'clientId', label: (p === 'dropbox' ? 'App key' : 'Client ID') + ':', value: c.clientId || '' }];
  if (P.secret) f.push({ key: 'secret', label: 'Client secret (Google yêu cầu cho ứng dụng máy tính):', value: c.secret || '' });
  const v = await formDlg(P.name + ' — mã ứng dụng OAuth', f.concat([{ type: 'html', el: h('div.hint.cl-help', h('b', 'Cách lấy: '), P.help, h('div', h('b', 'URI chuyển hướng: '), h('code', P.redirect))) }]), { width: 560 }); if (!v) return;
  c.clientId = (v.clientId || '').trim(); if (P.secret) c.secret = (v.secret || '').trim(); saveCfg();
}
export function accountsDialog() {
  const list = h('div.cl-accts');
  const render = () => { list.innerHTML = ''; Object.keys(PROVIDERS).forEach(p => { const P = PROVIDERS[p], c = cfg()[p]; const on = isConnected(p);
    list.appendChild(h('div.cl-acct', h('span.cl-ic', { style: { background: P.color } }, ic(P.icon, 16)), h('div.cl-acct-t', h('b', P.name), h('div.muted', on ? 'Đã kết nối: ' + (c.user || '') : c.clientId ? 'Chưa đăng nhập' : 'Chưa nhập Client ID')),
      h('button.btn', { onclick: () => accountDialog(p).then(render) }, 'Client ID…'),
      on ? h('button.btn', { onclick: () => { disconnect(p); render(); } }, 'Đăng xuất') : h('button.btn.primary', { onclick: () => connect(p).then(render) }, 'Đăng nhập'))); }); };
  render();
  const auto = h('input', { type: 'checkbox' }); auto.checked = cfg().autoSync !== false; auto.addEventListener('change', () => { cfg().autoSync = auto.checked; saveCfg(); });
  dialog({ title: 'Tài khoản đám mây', width: 640, body: h('div', list, h('label.chk', auto, h('span', 'Tự tải bản mới khi người khác sửa (nếu tài liệu chưa có thay đổi chưa lưu)')), h('p.muted', 'Thông tin đăng nhập (token) chỉ lưu trong máy này. Đồng tác giả thời gian thực dùng trình soạn thảo web của chính Google / Microsoft; MediaFire chỉ hỗ trợ đọc thư mục công khai.')), buttons: [{ label: 'Đóng', primary: true }] });
}
export function optionsSection(sec) {
  const st = Object.keys(PROVIDERS).map(p => PROVIDERS[p].name.split(' ')[0] + (isConnected(p) ? ' ✓' : '')).join(' · ');
  return sec('Đám mây (Google Drive, OneDrive, Dropbox, MediaFire)',
    h('div.form-row.inline', h('label.lbl', st), h('button.btn', { onclick: () => accountsDialog() }, ic('cloud', 14), ' Tài khoản & đồng bộ…')));
}
