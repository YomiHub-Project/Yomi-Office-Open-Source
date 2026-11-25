// v0.1.8 — Office Feature Check #28: khi mở ứng dụng, quét toàn bộ tệp trên máy và lọc ra các tệp Word, Excel, PowerPoint.
// Đường nhanh: Host.scanOffice() (Android: MediaStore; Windows: /api/scan_office nếu host có). Dự phòng trên Windows: duyệt thư mục
// người dùng bằng /api/list_dir (giới hạn độ sâu, số thư mục và thời gian để không làm chậm máy).
import Host from './host.js';
import { settings, saveSettings } from './store.js';

export const SCAN_EXTS = {
  word: ['docx', 'docm', 'doc', 'dotx', 'dotm', 'dot', 'odt', 'ott', 'rtf'],
  excel: ['xlsx', 'xlsm', 'xlsb', 'xls', 'xltx', 'xltm', 'ods', 'ots', 'csv'],
  ppt: ['pptx', 'pptm', 'ppt', 'ppsx', 'pps', 'potx', 'potm', 'odp', 'otp']
};
const EXT2KIND = {}; Object.keys(SCAN_EXTS).forEach(k => SCAN_EXTS[k].forEach(e => { EXT2KIND[e] = k; }));
// Thư mục hệ thống / bộ nhớ đệm / mã nguồn: bỏ qua cho nhanh
const SKIP = /^(\.|\$|~|node_modules$|AppData$|Application Data$|Windows$|Program Files|ProgramData$|System Volume Information$|Recovery$|PerfLogs$|__pycache__$|\.git$|build$|dist$|obj$|bin$|Temp$|tmp$|Cache|cache$)/i;
const MAX_DIRS = 6000, MAX_DEPTH = 7, MAX_MS = 45000, MAX_FILES = 5000;

export function kindOf(name) { const m = /\.([a-z0-9]+)$/i.exec(name || ''); return m ? EXT2KIND[m[1].toLowerCase()] || null : null; }
export function cfg() { const s = settings(); s.devscan = Object.assign({ onStart: true, last: 0, files: [] }, s.devscan || {}); return s.devscan; }
export function normalize(f) {
  if (!f || !f.name || /^~\$/.test(f.name)) return null; // bỏ tệp khóa tạm của Office (~$...)
  const kind = kindOf(f.name); if (!kind) return null; // chỉ Word/Excel/PowerPoint (host Android trả cả .pdf/.txt)
  return { name: f.name, kind, path: f.uri || f.path, mtime: +f.mtime || 0, size: +f.size || 0, dir: f.dir || '' };
}

async function walkWindows(onProgress) {
  const start = await Host.listDir(''); if (!start || !start.ok) return null;
  // listDir('') mở thư mục Tài liệu → thư mục người dùng là thư mục cha
  const home = /[\\/]Documents[\\/]?$/i.test(start.dir) && start.parent ? start.parent : start.dir;
  const queue = [{ dir: home, depth: 0 }], seen = {}, out = []; const t0 = Date.now(); let dirs = 0;
  while (queue.length && dirs < MAX_DIRS && Date.now() - t0 < MAX_MS && out.length < MAX_FILES) {
    const it = queue.shift(); if (seen[it.dir.toLowerCase()]) continue; seen[it.dir.toLowerCase()] = 1; dirs++;
    let r; try { r = await Host.listDir(it.dir); } catch (e) { r = null; } if (!r || !r.ok) continue;
    for (const x of r.items || []) {
      if (x.dir) { if (it.depth < MAX_DEPTH && !SKIP.test(x.name)) queue.push({ dir: x.path, depth: it.depth + 1 }); }
      else if (kindOf(x.name)) out.push({ name: x.name, path: x.path, mtime: x.mtime, size: x.size, dir: it.dir });
    }
    if (onProgress && dirs % 25 === 0) onProgress({ dirs, files: out.length });
  }
  return out;
}

let running = null;
export function scan(onProgress) { if (!running) running = doScan(onProgress).then(r => { running = null; return r; }, e => { running = null; throw e; }); return running; }
async function doScan(onProgress) {
  let files = null;
  try { files = Host.scanOffice ? await Host.scanOffice() : null; } catch (e) { files = null; }
  if (!files && Host.kind === 'windows') files = await walkWindows(onProgress);
  if (!files) return null; // trình duyệt web thuần: không có quyền đọc ổ đĩa
  const seen = {};
  files = files.map(normalize).filter(f => f && !seen[f.path] && (seen[f.path] = 1)).sort((a, b) => b.mtime - a.mtime).slice(0, MAX_FILES);
  const c = cfg(); c.files = files; c.last = Date.now(); saveSettings({ devscan: c });
  const n = k => files.filter(f => f.kind === k).length;
  try { console.log('OOSDEVSCAN | ' + files.length + ' files | word=' + n('word') + ' excel=' + n('excel') + ' ppt=' + n('ppt')); } catch (e) {}
  return files;
}
export function startupScan() { const c = cfg(); if (c.onStart && Host.kind !== 'web') setTimeout(() => scan().catch(() => {}), 4000); }
