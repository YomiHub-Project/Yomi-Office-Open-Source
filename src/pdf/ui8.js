// v0.0.8 — PDF reader interface rebuilt after the reference screens (Foxit PDF Reader ribbon/backstage/print dialog,
// PDF24 compact toolbar & save options, iLovePDF tools panel & document info panel):
//   left navigation rail (Trang / Dấu trang / Chú thích / Tệp đính kèm / Thông tin), right "Công cụ" panel,
//   page navigation + layout + zoom in the status bar, tool modes (Tay, Chọn chữ, Ảnh chụp nhanh, Thu phóng vùng, Đo),
//   page layouts (một trang / liên tục / hai trang / hai trang có trang bìa), xoay khung nhìn, đảo thứ tự, chế độ đêm,
//   dàn lại chữ (Reflow), tự cuộn, đọc to, đếm từ, hộp thoại In kiểu Foxit (phạm vi, chẵn/lẻ, co giãn, nhiều trang/tờ,
//   sách gấp, hướng, xem trước), Thuộc tính (Mô tả / Bảo mật / Phông chữ / Nâng cao), PDF/A, PDF → PowerPoint/Excel,
//   cắt lề, sửa tệp hỏng, biểu mẫu (đặt lại / xuất / nhập dữ liệu), đóng dấu thời gian. All offline.
import { h, ic, $, $$, clamp, esc, PDFLib, JSZipLib, dataUrlToBytes, fmtSize, utf8Encode, b64FromBytes, bytesFromB64 } from '../core/util.js';
import { toast, dialog, formDlg, promptDlg, busy, showMenu, contextMenu } from '../core/ui.js';
import Host from '../core/host.js';
import * as T from './tools.js';

const PRO = () => import('./pro.js');
const MM = 25.4 / 72;

// ================================================================ layout scaffolding
export function buildShell(d) {
  d.layout = d.layout || 'cont'; d.viewRot = 0; d.reverse = false; d.toolMode = 'select';
  d.rail = h('div.pdf-rail');
  d.navPanel = h('div.pdf-nav');
  d.navHead = h('div.pdf-nav-head'); d.navBody = h('div.pdf-nav-body');
  d.navPanel.appendChild(d.navHead); d.navPanel.appendChild(d.navBody);
  d.navBody.appendChild(d.thumbs);
  d.toolsPanel = h('div.pdf-tools8');
  const tabs = [['pages', 'layout-grid', 'Trang (hình thu nhỏ)'], ['bookmarks', 'bookmark', 'Dấu trang / Mục lục'], ['comments', 'message-square', 'Chú thích'], ['attach', 'paperclip', 'Tệp đính kèm'], ['info', 'info', 'Thông tin tài liệu'], ['search', 'search', 'Tìm kiếm']];
  tabs.forEach(t => { const b = h('button.pdf-rail-b', { title: t[2], dataset: { k: t[0] } }, ic(t[1], 18)); b.addEventListener('click', () => togglePanel(d, t[0])); d.rail.appendChild(b); });
  const tb = h('button.pdf-rail-b.pdf-rail-tools', { title: 'Bảng Công cụ', dataset: { k: 'tools' } }, ic('wrench', 18)); tb.addEventListener('click', () => toggleTools(d)); d.rail.appendChild(tb);
  d.el.insertBefore(d.navPanel, d.el.firstChild); d.el.insertBefore(d.rail, d.el.firstChild);
  d.el.appendChild(d.toolsPanel);
  d.navKey = 'pages'; if ((window.innerWidth || 1000) < 760) d.el.classList.add('pdf-nothumbs'); markRail(d);
  // status bar: page navigation (Foxit / PDF24 style)
  d.sbNav = h('div.pdf-sbnav',
    navBtn('chevrons-left', 'Trang đầu (Home)', () => d.goto(0)), navBtn('chevron-left', 'Trang trước (PgUp)', () => d.goto(d.cur - 1)),
    d.sbPageInp = h('input.pdf-pg-inp', { type: 'text', value: '1', title: 'Số trang — gõ rồi nhấn Enter' }), d.sbPageTot = h('span.pdf-pg-tot', '/ 0'),
    navBtn('chevron-right', 'Trang sau (PgDn)', () => d.goto(d.cur + 1)), navBtn('chevrons-right', 'Trang cuối (End)', () => d.goto(d.pageInfo.length - 1)));
  d.sbPageInp.addEventListener('keydown', e => { if (e.key === 'Enter') { const n = parseInt(d.sbPageInp.value); if (n) d.goto(n - 1); d.viewer.focus(); } });
  d.statusLeft.insertBefore(d.sbNav, d.statusLeft.firstChild);
  d.viewer.setAttribute('tabindex', '0');
  d.viewer.addEventListener('keydown', e => onKey(d, e));
  bindToolModes(d);
}
function navBtn(icn, title, fn) { const b = h('button.sb-item.pdf-nb', { title }, ic(icn, 14)); b.addEventListener('click', fn); return b; }
function markRail(d) { $$('.pdf-rail-b', d.rail).forEach(b => b.classList.toggle('on', b.dataset.k === d.navKey && !d.el.classList.contains('pdf-nothumbs') || (b.dataset.k === 'tools' && d.el.classList.contains('pdf-tools-on')))); }
export function togglePanel(d, k) {
  const open = !d.el.classList.contains('pdf-nothumbs');
  if (open && d.navKey === k) { d.el.classList.add('pdf-nothumbs'); markRail(d); d.relayout(); return; }
  d.navKey = k; d.el.classList.remove('pdf-nothumbs'); if ((window.innerWidth || 1000) < 760) d.el.classList.remove('pdf-tools-on'); markRail(d); renderPanel(d); d.relayout();
}
export function toggleTools(d, on) { d.el.classList.toggle('pdf-tools-on', on == null ? !d.el.classList.contains('pdf-tools-on') : on); if (d.el.classList.contains('pdf-tools-on') && (window.innerWidth || 1000) < 760) d.el.classList.add('pdf-nothumbs'); if (d.el.classList.contains('pdf-tools-on')) renderTools(d); markRail(d); setTimeout(() => d.relayout(), 30); }
export function updateNav(d) {
  if (!d.sbPageInp) return; const n = d.pageInfo.length;
  if (document.activeElement !== d.sbPageInp) d.sbPageInp.value = n ? (d.cur || 0) + 1 : 0;
  d.sbPageTot.textContent = '/ ' + n;
}
export function refreshPanels(d) { if (!d.el.classList.contains('pdf-nothumbs') && d.navKey !== 'pages') renderPanel(d); if (d.el.classList.contains('pdf-tools-on')) renderTools(d); updateNav(d); }

async function renderPanel(d) {
  const k = d.navKey; const titles = { pages: 'Trang', bookmarks: 'Dấu trang', comments: 'Chú thích', attach: 'Tệp đính kèm', info: 'Thông tin', search: 'Tìm kiếm' };
  d.navHead.innerHTML = ''; d.navHead.appendChild(h('b', titles[k]));
  const x = h('button.pdf-nav-x', { title: 'Đóng bảng' }, ic('x', 14)); x.addEventListener('click', () => togglePanel(d, k)); d.navHead.appendChild(x);
  Array.prototype.slice.call(d.navBody.children).forEach(c => { if (c !== d.thumbs) d.navBody.removeChild(c); });
  d.thumbs.style.display = k === 'pages' ? '' : 'none';
  if (k === 'pages') return;
  const box = h('div.pdf-nav-list'); d.navBody.appendChild(box);
  if (!d.pdf) { box.appendChild(h('div.muted.pdf-nav-empty', 'Chưa mở tệp PDF')); return; }
  try {
    if (k === 'bookmarks') await panelOutline(d, box);
    else if (k === 'comments') await panelComments(d, box);
    else if (k === 'attach') await panelAttach(d, box);
    else if (k === 'info') await panelInfo(d, box);
    else if (k === 'search') panelSearch(d, box);
  } catch (e) { box.appendChild(h('div.muted.pdf-nav-empty', 'Không đọc được: ' + e.message)); }
}
async function gotoDest(d, dest) { try { const x = typeof dest === 'string' ? await d.pdf.getDestination(dest) : dest; if (!x) return; const idx = typeof x[0] === 'object' ? await d.pdf.getPageIndex(x[0]) : x[0]; d.goto(idx); } catch (e) {} }
async function panelOutline(d, box) {
  const ol = await d.pdf.getOutline();
  if (!ol || !ol.length) { box.appendChild(h('div.muted.pdf-nav-empty', 'Tài liệu không có dấu trang. Bấm "Tạo từ tiêu đề" để tạo mục lục tạm từ cỡ chữ lớn.')); const b = h('button.btn', 'Tạo từ tiêu đề'); b.addEventListener('click', () => autoOutline(d, box)); box.appendChild(b); return; }
  (function add(items, lvl) { items.forEach(it => { const b = h('button.pdf-ol-i', { style: { paddingLeft: (8 + lvl * 14) + 'px' } }, it.items && it.items.length ? ic('chevron-right', 12) : ic('bookmark', 12), ' ', it.title); b.addEventListener('click', () => gotoDest(d, it.dest)); box.appendChild(b); if (it.items && it.items.length) add(it.items, lvl + 1); }); })(ol, 0);
}
async function autoOutline(d, box) {
  box.innerHTML = ''; const hits = [];
  for (let i = 0; i < d.pdf.numPages && hits.length < 200; i++) { const pg = await d.measureOne(i); const tc = await pg.getTextContent(); const sizes = tc.items.map(t => t.height).filter(Boolean).sort((a, b) => a - b); const med = sizes[Math.floor(sizes.length / 2)] || 10; tc.items.forEach(t => { if (t.str.trim().length > 2 && t.height >= med * 1.35) hits.push({ i, t: t.str.trim(), s: t.height }); }); }
  if (!hits.length) { box.appendChild(h('div.muted.pdf-nav-empty', 'Không tìm thấy tiêu đề (bản scan cần OCR trước).')); return; }
  hits.forEach(x => { const b = h('button.pdf-ol-i', ic('hash', 12), ' ', x.t.slice(0, 80), h('span.muted', '  tr.' + (x.i + 1))); b.addEventListener('click', () => d.goto(x.i)); box.appendChild(b); });
}
const ANN = { Text: 'Ghi chú', Highlight: 'Tô sáng', Underline: 'Gạch chân', StrikeOut: 'Gạch ngang', Squiggly: 'Gạch sóng', FreeText: 'Hộp chữ', Ink: 'Bút vẽ', Square: 'Hình chữ nhật', Circle: 'Hình tròn', Line: 'Đường thẳng', Polygon: 'Đa giác', Stamp: 'Con dấu', Link: 'Liên kết', Widget: 'Trường biểu mẫu', FileAttachment: 'Tệp đính kèm', Popup: 'Popup' };
async function panelComments(d, box) {
  let n = 0;
  for (let i = 0; i < d.pdf.numPages; i++) {
    const pg = await d.measureOne(i); const an = (await pg.getAnnotations()).filter(a => a.subtype !== 'Popup' && a.subtype !== 'Link' && a.subtype !== 'Widget');
    if (!an.length) continue; box.appendChild(h('div.pdf-nav-grp', 'Trang ' + (i + 1) + ' (' + an.length + ')'));
    an.forEach(a => { n++; const txt = (a.contentsObj && a.contentsObj.str) || a.contents || ''; const b = h('button.pdf-ol-i', ic('message-square', 12), ' ', h('b', ANN[a.subtype] || a.subtype), txt ? ' — ' + txt.slice(0, 80) : '', a.titleObj && a.titleObj.str ? h('div.muted', a.titleObj.str) : null); b.addEventListener('click', () => d.goto(i)); box.appendChild(b); });
  }
  if (!n) box.appendChild(h('div.muted.pdf-nav-empty', 'Không có chú thích. Dùng tab "Chú thích" để thêm.'));
}
async function panelAttach(d, box) {
  const at = await d.pdf.getAttachments();
  const keys = at ? Object.keys(at) : [];
  if (!keys.length) { box.appendChild(h('div.muted.pdf-nav-empty', 'Không có tệp đính kèm.')); }
  keys.forEach(k => { const a = at[k]; const b = h('button.pdf-ol-i', ic('paperclip', 12), ' ', a.filename || k, h('span.muted', '  ' + fmtSize((a.content || []).length))); b.addEventListener('click', () => T.saveOut(d, a.filename || k, a.content, 'application/octet-stream')); box.appendChild(b); });
  const add = h('button.btn', { style: { marginTop: '8px' } }, ic('plus', 13), ' Đính kèm tệp…'); add.addEventListener('click', () => attachFile(d)); box.appendChild(add);
}
async function attachFile(d) {
  const f = await pickFiles('*/*', false); if (!f || !f.length) return;
  const PL = await PDFLib(); const p = await PL.PDFDocument.load(d.bytes);
  await p.attach(f[0].bytes, f[0].name, { mimeType: 'application/octet-stream', description: f[0].name, creationDate: new Date(), modificationDate: new Date() });
  await d.commit(await p.save(), 'Đã đính kèm ' + f[0].name); if (d.navKey === 'attach') renderPanel(d);
}
export async function docInfo(d) {
  const meta = await d.pdf.getMetadata().catch(() => ({ info: {} })); const info = meta.info || {};
  const head = new TextDecoder('latin1').decode(d.bytes.subarray(0, 1024));
  const ver = (/%PDF-(\d\.\d)/.exec(head) || [])[1] || info.PDFFormatVersion || '?';
  const tail = new TextDecoder('latin1').decode(d.bytes.subarray(0, Math.min(d.bytes.length, 4000000)));
  const pdfa = /pdfaid:part\s*=\s*["']?(\d)/.exec(tail) || /<pdfaid:part>(\d)</.exec(tail); const conf = /pdfaid:conformance\s*=\s*["']?([ABU])/i.exec(tail) || /<pdfaid:conformance>([ABU])</i.exec(tail);
  const pg = await d.measureOne(0); const vp = pg.getViewport({ scale: 1 });
  const dt = s => { if (!s) return '—'; const m = /D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?/.exec(s); return m ? (m[3] || '01') + '/' + (m[2] || '01') + '/' + m[1] + ' ' + (m[4] || '00') + ':' + (m[5] || '00') + ':' + (m[6] || '00') : s; };
  return {
    file: d.path || d.name, version: ver, size: d.bytes.length, pages: d.pdf.numPages,
    pageSize: (vp.width / 72).toFixed(2) + ' × ' + (vp.height / 72).toFixed(2) + ' in (' + Math.round(vp.width * MM) + ' × ' + Math.round(vp.height * MM) + ' mm)',
    title: info.Title || '', author: info.Author || '', subject: info.Subject || '', keywords: info.Keywords || '', creator: info.Creator || '', producer: info.Producer || '',
    created: dt(info.CreationDate), modified: dt(info.ModDate), tagged: /\/StructTreeRoot/.test(tail) || /\/MarkInfo\s*<<[^>]*\/Marked\s+true/.test(tail), linear: /\/Linearized/.test(head),
    encrypted: !!d._password || /\/Encrypt\s+\d+\s+\d+\s+R/.test(tail.slice(-8000)), pdfa: pdfa ? 'PDF/A-' + pdfa[1] + (conf ? conf[1].toUpperCase() : '') : ''
  };
}
async function panelInfo(d, box) {
  const i = await docInfo(d);
  const row = (k, v) => h('div.pdf-info-r', h('div.muted', k), h('div', v || '—'));
  const ed = h('button.pdf-nav-x', { title: 'Sửa thuộc tính' }, ic('pencil', 13)); ed.addEventListener('click', () => propertiesDialog(d)); d.navHead.insertBefore(ed, d.navHead.lastChild);
  [['Loại tệp', 'PDF'], ['Phiên bản PDF', i.version], ['Kích thước tệp', fmtSize(i.size)], ['Số trang', String(i.pages)], ['Khổ trang', i.pageSize], ['Ngày tạo', i.created], ['Ngày sửa đổi', i.modified], ['Tác giả', i.author], ['Ứng dụng tạo', i.creator], ['Trình tạo PDF', i.producer], ['Tiêu đề', i.title], ['Chủ đề', i.subject], ['Từ khóa', i.keywords], ['Gắn thẻ (Tagged)', i.tagged ? 'Có' : 'Không'], ['Xem nhanh trên web', i.linear ? 'Có' : 'Không'], ['Mã hóa', i.encrypted ? 'Có mật khẩu' : 'Không']].forEach(r => box.appendChild(row(r[0], r[1])));
  box.appendChild(h('div.pdf-pdfa' + (i.pdfa ? '.ok' : ''), ic(i.pdfa ? 'badge-check' : 'circle-alert', 14), ' ', i.pdfa ? 'Tệp khai báo tuân thủ ' + i.pdfa + '.' : 'Tệp PDF này KHÔNG tuân thủ PDF/A.'));
}
function panelSearch(d, box) {
  const inp = h('input.inp', { type: 'search', placeholder: 'Tìm trong tài liệu…', value: d._lastQ || '' });
  const opts = h('label.pdf-chk', h('input', { type: 'checkbox' }), ' Khớp chữ hoa/thường');
  const res = h('div');
  const go = async () => {
    const q = inp.value.trim(); res.innerHTML = ''; if (!q) return; d._lastQ = q; const cs = opts.firstChild.checked;
    const norm = s => cs ? s : s.toLowerCase(); const nq = norm(q); let n = 0;
    for (let i = 0; i < d.pdf.numPages && n < 500; i++) { const pg = await d.measureOne(i); const tc = await pg.getTextContent(); let txt = tc.items.map(x => x.str).join(' '); if (!txt.trim() && d.ocrText[i]) txt = d.ocrText[i]; let k = norm(txt).indexOf(nq); while (k >= 0 && n < 500) { n++; const b = h('button.pdf-ol-i', h('b', 'Tr.' + (i + 1) + ' '), txt.slice(Math.max(0, k - 30), k), h('mark', txt.substr(k, q.length)), txt.slice(k + q.length, k + q.length + 40)); const pi = i; b.addEventListener('click', () => { d.goto(pi); setTimeout(() => markHits(d, pi, q, cs), 350); }); res.appendChild(b); k = norm(txt).indexOf(nq, k + 1); } }
    res.insertBefore(h('div.pdf-nav-grp', n ? n + ' kết quả' : 'Không tìm thấy' + (Object.keys(d.ocrText || {}).length ? '' : ' (bản scan: chạy OCR trước)')), res.firstChild);
  };
  inp.addEventListener('keydown', e => { if (e.key === 'Enter') go(); });
  box.appendChild(inp); box.appendChild(opts); box.appendChild(res); setTimeout(() => inp.focus(), 50); if (inp.value) go();
}
function markHits(d, page, q, cs) { const p = d.pageInfo[page]; if (!p || !p.textLayer) return; const nq = cs ? q : q.toLowerCase(); Array.prototype.forEach.call(p.textLayer.children, s => s.classList.toggle('pdf-hit', (cs ? s.textContent : s.textContent.toLowerCase()).indexOf(nq) >= 0)); }

// ================================================================ right "Công cụ" panel (iLovePDF style)
const TOOLGROUPS = d => [
  ['Chuyển đổi từ PDF', [['PDF sang JPG', 'image', '#f2b705', () => T.pdfToImages(d, 'jpg')], ['PDF sang Word', 'file-type-2', '#2b579a', () => T.exportAs(d, 'docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Word')], ['PDF sang PowerPoint', 'presentation', '#c43e1c', () => toPptx(d)], ['PDF sang Excel', 'sheet', '#107c41', () => toXlsx(d)], ['PDF sang PDF/A', 'badge-check', '#2b6cb0', () => toPdfA(d)], ['PDF sang văn bản', 'file-text', '#666', () => T.exportAs(d, 'txt', 'text/plain', 'TXT')]]],
  ['Chuyển đổi sang PDF', [['Ảnh sang PDF', 'image-plus', '#f2b705', () => T.imagesToPdfDialog(d)], ['Quét tài liệu', 'scan-line', '#d83b01', () => import('../scan/scan.js').then(S => S.openScanner(d.shell))]]],
  ['Chỉnh sửa PDF', [['Xoay tệp PDF', 'rotate-cw', '#8e44ad', () => T.rotatePages(d, d.pageInfo.map((_, i) => i), 90)], ['Số trang', 'list-ordered', '#8e44ad', () => T.pageNumbersDialog(d)], ['Dấu bản quyền', 'droplets', '#8e44ad', () => T.watermarkDialog(d)], ['Chỉnh sửa PDF', 'pen-line', '#8e44ad', () => PRO().then(P => P.annotate(d, 'text'))], ['Cắt cúp PDF', 'crop', '#8e44ad', () => cropDialog(d)]]],
  ['Tối ưu hóa PDF', [['Nén PDF', 'archive', '#27ae60', () => PRO().then(P => P.compressDialog(d))], ['Sửa tệp PDF', 'wrench', '#27ae60', () => repairPdf(d)], ['PDF sang OCR', 'scan-text', '#27ae60', () => T.ocrDialog(d)]]],
  ['Sắp xếp PDF', [['Nối PDF', 'combine', '#e74c3c', () => T.mergeDialog(d)], ['Tách PDF', 'scissors', '#e74c3c', () => T.splitDialog(d)], ['Xoá trang', 'trash-2', '#e74c3c', () => T.deletePages(d, d.sel.length ? d.sel : [d.cur || 0])], ['Sắp xếp trang', 'layout-grid', '#e74c3c', () => d.setFit('grid')], ['Đảo thứ tự trang', 'arrow-down-up', '#e74c3c', () => reverseOrder(d)], ['Chèn trang trống', 'file-plus', '#e74c3c', () => insertBlank(d)]]],
  ['Bảo mật PDF', [['Khóa PDF (mật khẩu)', 'lock', '#34495e', () => PRO().then(P => P.protectDialog(d))], ['Mở khóa PDF', 'lock-open', '#34495e', () => PRO().then(P => P.unlockDialog(d))], ['Ký PDF', 'signature', '#34495e', () => PRO().then(P => P.annotate(d, 'sign'))], ['Đóng dấu thời gian', 'clock', '#34495e', () => timeStamp(d)], ['Xoá trắng (che thông tin)', 'eraser', '#34495e', () => PRO().then(P => P.annotate(d, 'whiteout'))]]]
];
function renderTools(d) {
  const P = d.toolsPanel; P.innerHTML = '';
  const hd = h('div.pdf-t8-head', ic('wrench', 16), h('b', 'Công cụ')); const x = h('button.pdf-nav-x', { title: 'Đóng' }, ic('chevron-right', 14)); x.addEventListener('click', () => toggleTools(d, false)); hd.appendChild(x); P.appendChild(hd);
  const body = h('div.pdf-t8-body'); P.appendChild(body);
  d._toolFold = d._toolFold || {};
  TOOLGROUPS(d).forEach(g => {
    const sec = h('div.pdf-t8-sec' + (d._toolFold[g[0]] ? '.fold' : ''));
    const head = h('button.pdf-t8-sh', g[0], ic('chevron-down', 13)); head.addEventListener('click', () => { d._toolFold[g[0]] = !d._toolFold[g[0]]; sec.classList.toggle('fold'); });
    sec.appendChild(head);
    g[1].forEach(t => { const b = h('button.pdf-t8-i', h('span.pdf-t8-ic', { style: { background: t[2] } }, ic(t[1], 13)), t[0]); b.addEventListener('click', () => { if (!d.pdf && !/Ảnh sang|Quét|Nối/.test(t[0])) return toast('Mở một PDF trước', 'err'); t[3](); }); sec.appendChild(b); });
    body.appendChild(sec);
  });
}

// ================================================================ tool modes (Hand / Select / Snapshot / Marquee zoom / Measure)
export function setTool(d, m) { d.toolMode = m; d.el.classList.remove('pdf-m-hand', 'pdf-m-snap', 'pdf-m-zoom', 'pdf-m-measure', 'pdf-m-select'); d.el.classList.add('pdf-m-' + m); d.refreshUi(); }
function bindToolModes(d) {
  let st = null, box = null;
  const pageAt = (x, y) => { const el = document.elementFromPoint(x, y); return el && el.closest ? el.closest('.pdf-page') : null; };
  d.viewer.addEventListener('mousedown', e => {
    if (e.button !== 0 || d.toolMode === 'select' || !d.pdf) return;
    e.preventDefault();
    if (d.toolMode === 'hand') { st = { x: e.clientX, y: e.clientY, sl: d.viewer.scrollLeft, stp: d.viewer.scrollTop }; d.el.classList.add('pdf-dragging'); return; }
    const pe = pageAt(e.clientX, e.clientY); if (!pe) return;
    st = { x: e.clientX, y: e.clientY, pe };
    box = h('div.pdf-rubber' + (d.toolMode === 'measure' ? '.line' : '')); document.body.appendChild(box);
  });
  window.addEventListener('mousemove', e => {
    if (!st) return;
    if (d.toolMode === 'hand') { d.viewer.scrollLeft = st.sl - (e.clientX - st.x); d.viewer.scrollTop = st.stp - (e.clientY - st.y); return; }
    if (!box) return;
    if (d.toolMode === 'measure') {
      const dx = e.clientX - st.x, dy = e.clientY - st.y, len = Math.sqrt(dx * dx + dy * dy), ang = Math.atan2(dy, dx) * 180 / Math.PI;
      Object.assign(box.style, { left: st.x + 'px', top: st.y + 'px', width: len + 'px', transform: 'rotate(' + ang + 'deg)', WebkitTransform: 'rotate(' + ang + 'deg)' });
      const pt = len / d.scale; box.setAttribute('data-l', (pt * MM).toFixed(1) + ' mm · ' + (pt / 72).toFixed(2) + ' in');
    } else Object.assign(box.style, { left: Math.min(st.x, e.clientX) + 'px', top: Math.min(st.y, e.clientY) + 'px', width: Math.abs(e.clientX - st.x) + 'px', height: Math.abs(e.clientY - st.y) + 'px' });
  });
  window.addEventListener('mouseup', async e => {
    if (!st) return; const s0 = st; st = null; d.el.classList.remove('pdf-dragging');
    if (d.toolMode === 'hand') return;
    const b = box; box = null; if (!b) return;
    if (d.toolMode === 'measure') { b.classList.add('done'); setTimeout(() => b.parentNode && b.parentNode.removeChild(b), 6000); toast('Khoảng cách: ' + b.getAttribute('data-l')); return; }
    const r = b.getBoundingClientRect(); b.parentNode.removeChild(b);
    if (r.width < 6 || r.height < 6) return;
    if (d.toolMode === 'zoom') { const k = Math.min(d.viewer.clientWidth / r.width, d.viewer.clientHeight / r.height); const pr = s0.pe.getBoundingClientRect(); const fx = (r.left - pr.left) / pr.width, fy = (r.top - pr.top) / pr.height; d.shell.setZoom(clamp(Math.round(d.zoom * k * 0.95), 10, 500)); setTimeout(() => { d.viewer.scrollLeft = s0.pe.offsetLeft + fx * s0.pe.offsetWidth - 10; d.viewer.scrollTop = s0.pe.offsetTop + fy * s0.pe.offsetHeight - 10; }, 120); return; }
    if (d.toolMode === 'snap') await snapshot(d, s0.pe, r);
  });
}
async function snapshot(d, pe, r) {
  const i = +pe.dataset.i; const pr = pe.getBoundingClientRect();
  const k = 2; const url = await T.pageToDataUrl(d, i, d.scale * k * (1 / ((d.pageInfo[i].rotate || 0) % 180 ? 1 : 1)));
  const im = await new Promise(res => { const x = new Image(); x.onload = () => res(x); x.src = url; });
  const fx = im.naturalWidth / pr.width; const c = h('canvas'); c.width = Math.round(r.width * fx); c.height = Math.round(r.height * fx);
  c.getContext('2d').drawImage(im, (r.left - pr.left) * fx, (r.top - pr.top) * fx, c.width, c.height, 0, 0, c.width, c.height);
  const png = c.toDataURL('image/png');
  let copied = false;
  try { if (navigator.clipboard && window.ClipboardItem) { const blob = await (await fetch(png)).blob(); await navigator.clipboard.write([new window.ClipboardItem({ 'image/png': blob })]); copied = true; } } catch (e) {}
  dialog({ title: 'Ảnh chụp nhanh' + (copied ? ' — đã sao chép vào bộ nhớ tạm' : ''), width: 560, body: h('div', { style: { textAlign: 'center' } }, h('img', { src: png, style: { maxWidth: '100%', maxHeight: '50vh', boxShadow: '0 1px 6px rgba(0,0,0,.3)' } })),
    buttons: [{ label: 'Lưu PNG', run: () => T.saveOut(d, 'anh-chup-trang' + (i + 1) + '.png', dataUrlToBytes(png), 'image/png') }, { label: 'OCR vùng này', run: async () => { const b2 = busy('Đang nhận dạng chữ…'); try { const o = await Host.ocr(png, 'vie+eng'); b2.done(); dialog({ title: 'Chữ trong vùng chọn', width: 520, body: h('textarea.pdf-textout', o.text || ''), buttons: [{ label: 'Đóng', primary: true }] }); } catch (e) { b2.done(); toast(e.message, 'err'); } } }, { label: 'Đóng', primary: true }] });
}

// ================================================================ keyboard
function onKey(d, e) {
  if (e.target && /INPUT|TEXTAREA|SELECT/.test(e.target.tagName)) return;
  const k = e.key; const page = () => { e.preventDefault(); };
  if (k === 'PageDown' || (k === 'ArrowRight' && d.layout !== 'cont' && !e.ctrlKey)) { page(); d.goto(d.cur + (d.layout.indexOf('two') === 0 ? 2 : 1)); }
  else if (k === 'PageUp' || (k === 'ArrowLeft' && d.layout !== 'cont' && !e.ctrlKey)) { page(); d.goto(d.cur - (d.layout.indexOf('two') === 0 ? 2 : 1)); }
  else if (k === 'Home' && e.ctrlKey || k === 'Home' && d.layout === 'single') { page(); d.goto(0); }
  else if (k === 'End' && e.ctrlKey || k === 'End' && d.layout === 'single') { page(); d.goto(d.pageInfo.length - 1); }
  else if (k === 'Escape' && d._auto) stopAuto(d);
  else if (k === 'Escape' && d.el.classList.contains('pdf-readmode')) { d.el.classList.remove('pdf-readmode'); d.relayout(); d.refreshUi(); }
  else if (k === ' ' && d.layout === 'single') { page(); d.goto(d.cur + (e.shiftKey ? -1 : 1)); }
}

// ================================================================ view: layouts, rotate view, reverse, reflow, auto-scroll, read aloud, word count
export function setLayout(d, l) { d.layout = l; d.el.classList.toggle('pdf-single', l === 'single'); d.renderViewer(); d.refreshUi(); }
export function rotateView(d, deg) { d.viewRot = ((d.viewRot || 0) + deg + 360) % 360; d.pageInfo.forEach(p => { p.rendered = false; }); d.relayout(); d.renderThumbs(); toast('Xoay khung nhìn ' + d.viewRot + '° (không đổi tệp)'); }
export function toggleReverse(d) { d.reverse = !d.reverse; d.renderViewer(); d.refreshUi(); }
export async function reflowView(d, on) {
  d.reflow = on == null ? !d.reflow : on; d.el.classList.toggle('pdf-reflow', d.reflow);
  let rv = d.el.querySelector('.pdf-reflow-view');
  if (!d.reflow) { if (rv) rv.parentNode.removeChild(rv); d.refreshUi(); return; }
  if (!rv) { rv = h('div.pdf-reflow-view'); d.viewer.parentNode.insertBefore(rv, d.viewer.nextSibling); }
  rv.innerHTML = ''; const fs = d._reflowFs || 17;
  const bar = h('div.pdf-reflow-bar', h('b', 'Dàn lại chữ (Reflow) — đọc thoải mái trên màn hình nhỏ'), h('span', { style: { flex: 1 } }));
  [['a-arrow-down', -1, 'Chữ nhỏ hơn'], ['a-arrow-up', 1, 'Chữ to hơn']].forEach(x => { const b = h('button.btn', { title: x[2] }, ic(x[0], 14)); b.addEventListener('click', () => { d._reflowFs = clamp(fs + x[1] * 2, 11, 36); reflowView(d, true); }); bar.appendChild(b); });
  const cl = h('button.btn', 'Đóng'); cl.addEventListener('click', () => reflowView(d, false)); bar.appendChild(cl);
  rv.appendChild(bar);
  const col = h('div.pdf-reflow-col', { style: { fontSize: fs + 'px' } }); rv.appendChild(col);
  for (let i = 0; i < d.pdf.numPages; i++) {
    const pg = await d.measureOne(i); const tc = await pg.getTextContent(); let paras = [], line = '', last = null, lastH = 0;
    tc.items.forEach(it => { const y = it.transform[5]; if (last != null && Math.abs(y - last) > 3) { if (Math.abs(y - last) > lastH * 1.9) { paras.push(line.trim()); line = ''; } else line += ' '; } line += it.str; last = y; lastH = it.height || lastH || 10; });
    paras.push(line.trim()); if (!paras.join('').trim() && d.ocrText[i]) paras = d.ocrText[i].split(/\n{2,}/);
    col.appendChild(h('div.pdf-reflow-pg', 'Trang ' + (i + 1)));
    paras.filter(Boolean).forEach(p => col.appendChild(h('p', p.replace(/-\s+(?=[a-zà-ỹ])/g, ''))));
  }
  d.refreshUi();
}
export function autoScroll(d) {
  if (d._auto) { d._autoSpeed = Math.min(8, d._autoSpeed + 1); toast('Tốc độ tự cuộn: ' + d._autoSpeed + ' (Esc để dừng)'); return; }
  d._autoSpeed = 1; let last = 0;
  const step = t => { if (!d._auto) return; if (last) d.viewer.scrollTop += d._autoSpeed * (t - last) / 16 * 0.8; last = t; d._auto = requestAnimationFrame(step); };
  d._auto = requestAnimationFrame(step); toast('Đang tự cuộn — bấm lại để tăng tốc, Esc để dừng'); d.viewer.focus(); d.refreshUi();
}
export function stopAuto(d) { if (d._auto) cancelAnimationFrame(d._auto); d._auto = 0; d.refreshUi(); }
export async function readAloud(d) {
  if (!window.speechSynthesis) return toast('Thiết bị không hỗ trợ đọc to (cài giọng đọc tiếng Việt trong cài đặt hệ thống)', 'err');
  if (d._speaking) { speechSynthesis.cancel(); d._speaking = false; d.refreshUi(); return; }
  const pg = await d.measureOne(d.cur || 0); const tc = await pg.getTextContent(); let text = tc.items.map(x => x.str).join(' ').trim(); if (!text && d.ocrText[d.cur]) text = d.ocrText[d.cur];
  if (!text) return toast('Trang không có chữ (bản scan: chạy OCR trước)');
  const voice = speechSynthesis.getVoices().find(v => /vi/i.test(v.lang)); const parts = text.match(/[^.!?\n]{1,220}[.!?\n]*/g) || [text]; d._speaking = true;
  parts.forEach((p, i) => { const u = new SpeechSynthesisUtterance(p); u.lang = voice ? voice.lang : 'vi-VN'; if (voice) u.voice = voice; if (i === parts.length - 1) u.onend = () => { d._speaking = false; d.refreshUi(); }; speechSynthesis.speak(u); });
  toast('Đang đọc trang ' + ((d.cur || 0) + 1) + ' — bấm lại để dừng'); d.refreshUi();
}
export async function wordCount(d) {
  const b = busy('Đang đếm…'); let words = 0, chars = 0, charsNs = 0, lines = 0; const per = [];
  for (let i = 0; i < d.pdf.numPages; i++) { const pg = await d.measureOne(i); const tc = await pg.getTextContent(); let t = tc.items.map(x => x.str + (x.hasEOL ? '\n' : ' ')).join(''); if (!t.trim() && d.ocrText[i]) t = d.ocrText[i]; const w = (t.match(/\S+/g) || []).length; words += w; chars += t.replace(/\n/g, '').length; charsNs += t.replace(/\s/g, '').length; lines += (t.match(/\n/g) || []).length; per.push(w); }
  b.done();
  dialog({ title: 'Đếm từ', width: 380, body: h('table.kv', [['Số trang', d.pdf.numPages], ['Số từ', words], ['Ký tự (không khoảng trắng)', charsNs], ['Ký tự (có khoảng trắng)', chars], ['Số dòng', lines], ['Từ ở trang hiện tại', per[d.cur || 0]]].map(r => h('tr', h('td', r[0]), h('td', String(r[1]))))), buttons: [{ label: 'Đóng', primary: true }] });
}
export function zoomMenu(d) {
  const z = v => () => d.shell.setZoom(v);
  return [10, 25, 50, 75, 100, 125, 150, 200, 300, 400].map(v => ({ label: v + '%', checked: d.fitMode === 'custom' && d.zoom === v || undefined, run: z(v) })).concat([{ sep: true },
    { label: 'Kích thước thực', icon: 'scan', run: z(100) }, { label: 'Vừa trang', icon: 'file', run: () => d.setFit('page') }, { label: 'Vừa chiều rộng', icon: 'move-horizontal', run: () => d.setFit('width') }, { label: 'Vừa vùng có nội dung', icon: 'maximize', run: () => d.setFit('visible') }]);
}

// ================================================================ print (Foxit-style dialog; also used by Tệp → In)
const PAPERS = { auto: null, A4: [595.3, 841.9], A3: [841.9, 1190.6], A5: [419.5, 595.3], Letter: [612, 792], Legal: [612, 1008] };
function pcfg(d) { return d.printCfg || (d.printCfg = { range: 'all', pages: '', subset: 'all', reverse: false, handling: 'fit', pct: 100, perSheet: 1, booklet: false, orient: 'auto', paper: 'A4', gray: false, dpi: 150, border: false }); }
// sheets = [[pageIndex|null,...], ...] in print order
function planSheets(d) {
  const c = pcfg(d); const n = d.pdf.numPages; let idx;
  if (c.range === 'current') idx = [d.cur || 0];
  else if (c.range === 'pages') idx = T.parseRange(c.pages, n);
  else idx = d.pageInfo.map((_, i) => i);
  if (c.subset === 'odd') idx = idx.filter(i => i % 2 === 0); else if (c.subset === 'even') idx = idx.filter(i => i % 2 === 1);
  if (c.reverse) idx = idx.slice().reverse();
  if (!idx.length) idx = [d.cur || 0];
  if (c.booklet) { const pad = idx.slice(); while (pad.length % 4) pad.push(null); const out = []; let lo = 0, hi = pad.length - 1; while (lo < hi) { out.push([pad[hi--], pad[lo++]]); out.push([pad[lo++], pad[hi--]]); } return { sheets: out, cols: 2, rows: 1 }; }
  const per = c.perSheet || 1; const grid = { 1: [1, 1], 2: [2, 1], 4: [2, 2], 6: [3, 2], 9: [3, 3], 16: [4, 4] }[per] || [1, 1];
  const out = []; for (let i = 0; i < idx.length; i += per) out.push(idx.slice(i, i + per));
  return { sheets: out, cols: grid[0], rows: grid[1] };
}
function sheetSize(d, plan) {
  const c = pcfg(d); const p0 = d.pageInfo[0]; let w = p0._baseW || 595, hh = p0._baseH || 842;
  if (c.paper !== 'auto' && PAPERS[c.paper]) { w = PAPERS[c.paper][0]; hh = PAPERS[c.paper][1]; }
  let land = c.orient === 'landscape' || (c.orient === 'auto' && ((plan.cols > plan.rows) !== (p0._baseW > p0._baseH)) && plan.cols !== plan.rows);
  if (c.orient === 'auto' && plan.cols === 1 && plan.rows === 1) land = (p0._baseW || 1) > (p0._baseH || 1);
  if (c.booklet) land = true;
  return land ? [Math.max(w, hh), Math.min(w, hh)] : [Math.min(w, hh), Math.max(w, hh)];
}
async function drawSheet(d, plan, si, pxPerPt) {
  const c = pcfg(d); const [sw, sh] = sheetSize(d, plan); const cv = h('canvas'); cv.width = Math.round(sw * pxPerPt); cv.height = Math.round(sh * pxPerPt);
  const x = cv.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, cv.width, cv.height);
  const cellW = cv.width / plan.cols, cellH = cv.height / plan.rows; const pad = (plan.cols * plan.rows > 1 ? 10 : 0) * pxPerPt;
  const list = plan.sheets[si] || [];
  for (let k = 0; k < list.length; k++) {
    const pi = list[k]; if (pi == null) continue;
    const col = k % plan.cols, row = Math.floor(k / plan.cols); const page = await d.measureOne(pi); const rot = ((d.pageInfo[pi].rotate || 0)) % 360;
    const vp1 = page.getViewport({ scale: 1, rotation: rot }); const avW = cellW - pad * 2, avH = cellH - pad * 2;
    let s = Math.min(avW / vp1.width, avH / vp1.height);
    if (plan.cols * plan.rows === 1 && !c.booklet) { const actual = pxPerPt; if (c.handling === 'actual') s = actual; else if (c.handling === 'shrink') s = Math.min(s, actual); else if (c.handling === 'custom') s = actual * (c.pct || 100) / 100; }
    const vp = page.getViewport({ scale: s, rotation: rot }); const pc = h('canvas'); pc.width = Math.round(vp.width); pc.height = Math.round(vp.height);
    const px = pc.getContext('2d'); px.fillStyle = '#fff'; px.fillRect(0, 0, pc.width, pc.height); await page.render({ canvasContext: px, viewport: vp }).promise;
    const ox = col * cellW + (cellW - pc.width) / 2, oy = row * cellH + (cellH - pc.height) / 2;
    x.drawImage(pc, ox, oy); if (c.border || plan.cols * plan.rows > 1) { x.strokeStyle = '#999'; x.lineWidth = Math.max(1, pxPerPt * 0.5); x.strokeRect(ox, oy, pc.width, pc.height); }
  }
  if (c.gray) { const id = x.getImageData(0, 0, cv.width, cv.height); const a = id.data; for (let i = 0; i < a.length; i += 4) { const g = a[i] * 0.299 + a[i + 1] * 0.587 + a[i + 2] * 0.114; a[i] = a[i + 1] = a[i + 2] = g; } x.putImageData(id, 0, 0); }
  return cv;
}
export function printOptions(d) {
  if (!d.pdf) return h('div.muted', 'Mở một PDF để in.');
  const c = pcfg(d); const upd = () => { const pv = $('.print-prev'); if (pv) printPreview(d, pv); };
  const sel = (label, key, opts) => { const s = h('select.inp', opts.map(o => h('option', { value: o[0] }, o[1]))); s.value = String(c[key]); s.addEventListener('change', () => { c[key] = /^\d+$/.test(s.value) ? +s.value : s.value; upd(); }); return h('div.form-row', h('label.lbl', label), s); };
  const chk = (label, key) => { const i = h('input', { type: 'checkbox' }); i.checked = !!c[key]; i.addEventListener('change', () => { c[key] = i.checked; upd(); }); return h('label.pdf-chk', i, ' ', label); };
  const pages = h('input.inp', { placeholder: 'Ví dụ: 1,3-5,9', value: c.pages }); pages.addEventListener('input', () => { c.pages = pages.value; c.range = 'pages'; rangeSel.value = 'pages'; upd(); });
  const rangeSel = h('select.inp', [['all', 'Tất cả trang'], ['current', 'Trang hiện tại'], ['pages', 'Các trang…']].map(o => h('option', { value: o[0] }, o[1]))); rangeSel.value = c.range; rangeSel.addEventListener('change', () => { c.range = rangeSel.value; upd(); });
  return h('div.form.pdf-print-opts',
    h('div.pdf-print-h', 'Phạm vi in'), h('div.form-row', h('label.lbl', 'Trang'), rangeSel), pages,
    sel('Tập con', 'subset', [['all', 'Tất cả trang trong phạm vi'], ['odd', 'Chỉ trang lẻ'], ['even', 'Chỉ trang chẵn']]), chk('In ngược thứ tự trang', 'reverse'),
    h('div.pdf-print-h', 'Xử lý khi in'),
    sel('Co giãn', 'handling', [['fit', 'Vừa lề máy in'], ['shrink', 'Thu nhỏ trang quá khổ'], ['actual', 'Kích thước thực'], ['custom', 'Tỷ lệ tùy chỉnh']]), sel('Tỷ lệ (%)', 'pct', [[50, '50%'], [75, '75%'], [90, '90%'], [100, '100%'], [125, '125%'], [150, '150%']]),
    sel('Nhiều trang trên một tờ', 'perSheet', [[1, '1 trang/tờ'], [2, '2 trang/tờ'], [4, '4 trang/tờ'], [6, '6 trang/tờ'], [9, '9 trang/tờ'], [16, '16 trang/tờ']]),
    chk('In kiểu sách gấp (Booklet)', 'booklet'), chk('Viền quanh trang', 'border'),
    h('div.pdf-print-h', 'Giấy & chất lượng'),
    sel('Hướng', 'orient', [['auto', 'Tự động dọc/ngang'], ['portrait', 'Dọc'], ['landscape', 'Ngang']]), sel('Khổ giấy', 'paper', [['A4', 'A4'], ['A3', 'A3'], ['A5', 'A5'], ['Letter', 'Letter'], ['Legal', 'Legal'], ['auto', 'Theo khổ trang PDF']]),
    sel('Chất lượng', 'dpi', [[110, 'Nháp (110 dpi)'], [150, 'Thường (150 dpi)'], [220, 'Cao (220 dpi)'], [300, 'Rất cao (300 dpi)']]), chk('In thang xám (đen trắng)', 'gray'));
}
export async function printPreview(d, box) {
  if (!d.pdf) return; const plan = planSheets(d); d._pvSheet = clamp(d._pvSheet || 0, 0, plan.sheets.length - 1);
  const token = (d._pvTok = (d._pvTok || 0) + 1);
  const cv = await drawSheet(d, plan, d._pvSheet, 0.9); if (token !== d._pvTok) return;
  box.innerHTML = ''; cv.className = 'pdf-pv-sheet'; box.appendChild(cv);
  const [sw, sh] = sheetSize(d, plan);
  const nav = h('div.xl-pp-nav', navBtn('chevrons-left', 'Tờ đầu', () => { d._pvSheet = 0; printPreview(d, box); }), navBtn('chevron-left', 'Tờ trước', () => { d._pvSheet--; printPreview(d, box); }),
    h('span.xl-pp-lbl', 'Tờ ' + (d._pvSheet + 1) + ' / ' + plan.sheets.length + ' · ' + Math.round(sw * MM) + '×' + Math.round(sh * MM) + ' mm'),
    navBtn('chevron-right', 'Tờ sau', () => { d._pvSheet++; printPreview(d, box); }), navBtn('chevrons-right', 'Tờ cuối', () => { d._pvSheet = plan.sheets.length - 1; printPreview(d, box); }));
  box.appendChild(nav);
}
export async function renderPrint(d, area) {
  if (!d.pdf) return; const c = pcfg(d); const plan = planSheets(d); const [sw, sh] = sheetSize(d, plan);
  const b = busy('Đang chuẩn bị bản in…');
  area.appendChild(h('style', '@page{size:' + (sw * MM).toFixed(1) + 'mm ' + (sh * MM).toFixed(1) + 'mm;margin:0} #print-area .pdf-print-sheet{display:block;width:' + (sw * MM).toFixed(1) + 'mm;height:' + (sh * MM).toFixed(1) + 'mm;page-break-after:always;break-after:page} #print-area .pdf-print-sheet:last-child{page-break-after:auto}'));
  try { for (let i = 0; i < plan.sheets.length; i++) { const cv = await drawSheet(d, plan, i, (c.dpi || 150) / 72); area.appendChild(h('img.pdf-print-sheet', { src: cv.toDataURL('image/jpeg', 0.92) })); } } finally { b.done(); }
}

// ================================================================ properties (Foxit: Mô tả / Bảo mật / Phông chữ / Nâng cao)
export async function listFonts(d) {
  const PL = await PDFLib(); const out = [];
  try {
    const p = await PL.PDFDocument.load(d.bytes, { ignoreEncryption: true, updateMetadata: false });
    const N = PL.PDFName; p.context.enumerateIndirectObjects().forEach(([ref, obj]) => {
      if (!(obj instanceof PL.PDFDict)) return; const t = obj.get(N.of('Type')); if (!t || t.toString() !== '/Font') return;
      const st = obj.get(N.of('Subtype')); if (st && st.toString() === '/Type0') return;
      const bf = obj.get(N.of('BaseFont')); let fd = obj.lookup(N.of('FontDescriptor'));
      if (!fd && st && /Type0/.test(st.toString())) return;
      const emb = !!(fd && (fd.get(N.of('FontFile')) || fd.get(N.of('FontFile2')) || fd.get(N.of('FontFile3'))));
      const name = bf ? bf.toString().replace(/^\//, '') : '(không tên)';
      out.push({ name: name.replace(/^[A-Z]{6}\+/, ''), type: (st ? st.toString().replace('/', '') : '?'), embedded: emb, subset: /^[A-Z]{6}\+/.test(name) });
    });
  } catch (e) {}
  const seen = {}; return out.filter(f => { const k = f.name + f.type; if (seen[k]) return false; seen[k] = 1; return true; });
}
export async function propertiesDialog(d) {
  if (!d.pdf) return toast('Mở một PDF trước', 'err');
  const i = await docInfo(d); const fields = {};
  const inp = (k, v) => (fields[k] = h('input.inp', { value: v || '' }));
  const row = (k, v) => h('tr', h('td.muted', k), h('td', v));
  const pane = { desc: h('table.kv.pdf-prop', row('Tệp', i.file), row('Phiên bản PDF', 'PDF-' + i.version), row('Xem nhanh trên web', i.linear ? 'Có' : 'Không'), row('PDF gắn thẻ', i.tagged ? 'Có' : 'Không'), row('Khổ trang', i.pageSize), row('Số trang', String(i.pages)), row('Tiêu đề', inp('title', i.title)), row('Chủ đề', inp('subject', i.subject)), row('Tác giả', inp('author', i.author)), row('Từ khóa', inp('keywords', i.keywords)), row('Ứng dụng tạo', i.creator || '—'), row('Trình tạo PDF', i.producer || '—'), row('Ngày tạo', i.created), row('Sửa lần cuối', i.modified)),
    sec: h('div', h('table.kv.pdf-prop', row('Phương thức bảo mật', i.encrypted ? 'Mật khẩu (Standard Security)' : 'Không bảo mật'), row('Mở tài liệu', i.encrypted ? 'Cần mật khẩu' : 'Tự do'), row('In', 'Được phép'), row('Sao chép nội dung', 'Được phép'), row('Chỉnh sửa / chú thích', 'Được phép')), h('div', { style: { marginTop: '10px' } }, h('button.btn', { onclick: () => PRO().then(P => P.protectDialog(d)) }, ic('lock', 13), ' Đặt mật khẩu…'), ' ', h('button.btn', { onclick: () => PRO().then(P => P.unlockDialog(d)) }, ic('lock-open', 13), ' Gỡ mật khẩu…'))),
    fonts: h('div.pdf-fontlist', h('div.muted', 'Đang đọc danh sách phông…')),
    adv: h('table.kv.pdf-prop', row('PDF/A', i.pdfa || 'Không tuân thủ PDF/A'), row('Kích thước tệp', fmtSize(i.size) + ' (' + i.size.toLocaleString('vi-VN') + ' byte)'), row('Mã hóa', i.encrypted ? 'Có' : 'Không'), row('Hướng đọc', 'Trái sang phải'), row('Bố cục mở', d.layout)) };
  listFonts(d).then(fs => { pane.fonts.innerHTML = ''; if (!fs.length) pane.fonts.appendChild(h('div.muted', 'Không có phông (tài liệu chỉ chứa ảnh) hoặc không đọc được.')); fs.forEach(f => pane.fonts.appendChild(h('div.pdf-font', ic('type', 13), ' ', h('b', f.name), h('div.muted', 'Loại: ' + f.type + ' · ' + (f.embedded ? (f.subset ? 'Nhúng một phần (subset)' : 'Nhúng đầy đủ') : 'Không nhúng'))))); });
  const tabs = [['desc', 'Mô tả', 'file-text'], ['sec', 'Bảo mật', 'shield'], ['fonts', 'Phông chữ', 'type'], ['adv', 'Nâng cao', 'settings']];
  const left = h('div.pdf-prop-tabs'); const right = h('div.pdf-prop-body');
  const show = k => { right.innerHTML = ''; right.appendChild(pane[k]); $$('.pdf-prop-tab', left).forEach(b => b.classList.toggle('on', b.dataset.k === k)); };
  tabs.forEach(t => { const b = h('button.pdf-prop-tab', { dataset: { k: t[0] } }, ic(t[2], 14), ' ', t[1]); b.addEventListener('click', () => show(t[0])); left.appendChild(b); }); show('desc');
  const ok = await dialog({ title: 'Thuộc tính tài liệu', width: 720, body: h('div.pdf-prop-wrap', left, right), buttons: [{ label: 'Lưu thay đổi', primary: true, value: true }, { label: 'Đóng', value: false }] });
  if (!ok) return;
  const PL = await PDFLib(); const p = await PL.PDFDocument.load(d.bytes, { updateMetadata: false });
  p.setTitle(fields.title.value); p.setAuthor(fields.author.value); p.setSubject(fields.subject.value); p.setKeywords(fields.keywords.value ? fields.keywords.value.split(/\s*[,;]\s*/) : []); p.setModificationDate(new Date()); p.setProducer('Yomi-Office Open Source 0.0.8');
  await d.commit(await p.save(), 'Đã cập nhật thuộc tính'); refreshPanels(d);
}

// ================================================================ save options (PDF24)
export function saveMenu(d) {
  return [
    { label: 'Lưu', icon: 'save', sc: 'Ctrl+S', run: () => d.shell.save(d) },
    { label: 'Lưu thành…', icon: 'save-all', run: () => d.shell.save(d, true) },
    { label: 'Lưu dạng nén', icon: 'archive', run: () => PRO().then(P => P.compressDialog(d)) },
    { label: 'Lưu dạng phẳng (ảnh)', icon: 'layers', run: () => T.exportAs(d, 'pdf-flat', 'application/pdf', 'PDF phẳng') },
    { label: 'Lưu được bảo vệ bằng mật khẩu', icon: 'lock', run: () => PRO().then(P => P.protectDialog(d)) },
    { label: 'Lưu dạng PDF/A (lưu trữ)', icon: 'badge-check', run: () => toPdfA(d) }
  ];
}

// ================================================================ conversions
async function pageImages(d, scale, mime, q) { const out = []; for (let i = 0; i < d.pdf.numPages; i++) out.push(await T.pageToDataUrl(d, i, scale, mime, q)); return out; }
export async function buildPptx(images, sizes) {
  const JSZip = await JSZipLib(); const z = new JSZip(); const EMU = 12700;
  const W = Math.round(sizes[0][0] * EMU), H = Math.round(sizes[0][1] * EMU);
  const ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'];
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const sld = [], prs = [];
  images.forEach((img, i) => {
    const n = i + 1; const [pw, ph] = sizes[i]; const k = Math.min(W / (pw * EMU), H / (ph * EMU)); const w = Math.round(pw * EMU * k), hh = Math.round(ph * EMU * k);
    z.file('ppt/media/image' + n + '.jpeg', dataUrlToBytes(img));
    z.file('ppt/slides/slide' + n + '.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ' + NS + '><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="2" name="Trang ' + n + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="' + Math.round((W - w) / 2) + '" y="' + Math.round((H - hh) / 2) + '"/><a:ext cx="' + w + '" cy="' + hh + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>');
    z.file('ppt/slides/_rels/slide' + n + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image' + n + '.jpeg"/></Relationships>');
    ct.push('<Override PartName="/ppt/slides/slide' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>');
    sld.push('<p:sldId id="' + (255 + n) + '" r:id="rId' + (n + 2) + '"/>'); prs.push('<Relationship Id="rId' + (n + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + n + '.xml"/>');
  });
  ct.push('</Types>'); z.file('[Content_Types].xml', ct.join(''));
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  z.file('ppt/presentation.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation ' + NS + '><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>' + sld.join('') + '</p:sldIdLst><p:sldSz cx="' + W + '" cy="' + H + '"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>');
  z.file('ppt/_rels/presentation.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' + prs.join('') + '</Relationships>');
  const sp = '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>';
  z.file('ppt/slideMasters/slideMaster1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster ' + NS + '>' + sp + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>');
  z.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>');
  z.file('ppt/slideLayouts/slideLayout1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout ' + NS + ' type="blank" preserve="1">' + sp.replace('<p:cSld>', '<p:cSld name="Blank">') + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>');
  z.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
  const clr = (n, v) => '<a:' + n + '><a:srgbClr val="' + v + '"/></a:' + n + '>';
  z.file('ppt/theme/theme1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' + clr('dk2', '44546A') + clr('lt2', 'E7E6E6') + clr('accent1', '4472C4') + clr('accent2', 'ED7D31') + clr('accent3', 'A5A5A5') + clr('accent4', 'FFC000') + clr('accent5', '5B9BD5') + clr('accent6', '70AD47') + clr('hlink', '0563C1') + clr('folHlink', '954F72') + '</a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>');
  return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
export async function toPptx(d, noSave) {
  if (!d.pdf) return toast('Mở một PDF trước', 'err'); const b = busy('Đang chuyển PDF → PowerPoint…');
  try { const imgs = await pageImages(d, 1.6, 'image/jpeg', 0.88); const sizes = []; for (let i = 0; i < d.pdf.numPages; i++) { const pg = await d.measureOne(i); const vp = pg.getViewport({ scale: 1, rotation: d.pageInfo[i].rotate || 0 }); sizes.push([vp.width, vp.height]); }
    const bytes = await buildPptx(imgs, sizes); b.done(); if (noSave) return bytes; await T.saveOut(d, (d.name || 'pdf').replace(/\.pdf$/i, '') + '.pptx', bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'); return bytes; }
  catch (e) { b.done(); toast('Lỗi: ' + e.message, 'err'); }
}
// PDF → Excel: words grouped into rows (by baseline) and columns (by x gaps) per page, one sheet per page
export async function pdfTables(d) {
  const pages = [];
  for (let i = 0; i < d.pdf.numPages; i++) {
    const pg = await d.measureOne(i); const tc = await pg.getTextContent(); const items = tc.items.filter(t => t.str.trim()).map(t => ({ s: t.str.trim(), x: t.transform[4], y: t.transform[5], w: t.width || t.str.length * 5, hh: t.height || 10 }));
    items.sort((a, b) => b.y - a.y || a.x - b.x);
    const rows = []; items.forEach(it => { const r = rows.find(r => Math.abs(r.y - it.y) <= Math.max(2, it.hh * 0.45)); if (r) r.items.push(it); else rows.push({ y: it.y, items: [it] }); });
    rows.forEach(r => r.items.sort((a, b) => a.x - b.x));
    // merge adjacent words that belong to one cell (gap < ~1.2 char)
    rows.forEach(r => { const m = []; r.items.forEach(it => { const l = m[m.length - 1]; if (l && it.x - (l.x + l.w) < Math.max(4, it.hh * 0.9)) { l.s += ' ' + it.s; l.w = it.x + it.w - l.x; } else m.push(Object.assign({}, it)); }); r.cells = m; });
    // column anchors from all cell x positions
    const xs = []; rows.forEach(r => r.cells.forEach(c => xs.push(c.x))); xs.sort((a, b) => a - b);
    const anchors = []; xs.forEach(x => { const a = anchors.find(a => Math.abs(a.x - x) < 14); if (a) { a.n++; a.x = (a.x * (a.n - 1) + x) / a.n; } else anchors.push({ x, n: 1 }); }); anchors.sort((a, b) => a.x - b.x);
    const grid = rows.map(r => { const line = []; r.cells.forEach(c => { let k = 0, best = 1e9; anchors.forEach((a, j) => { const dd = Math.abs(a.x - c.x); if (dd < best) { best = dd; k = j; } }); while (line[k] != null) k++; line[k] = c.s; }); return line; });
    pages.push(grid);
  }
  return pages;
}
export async function toXlsx(d, noSave) {
  if (!d.pdf) return toast('Mở một PDF trước', 'err'); const b = busy('Đang chuyển PDF → Excel…');
  try {
    const pages = await pdfTables(d); const JSZip = await JSZipLib(); const z = new JSZip();
    const xe = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const col = c => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; };
    const sheets = pages.map((g, pi) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + g.map((row, r) => '<row r="' + (r + 1) + '">' + row.map((v, c) => { if (v == null) return ''; const num = /^-?[\d.,]+$/.test(v) && !/^0\d/.test(v) ? Number(v.replace(/\./g, '').replace(',', '.')) : NaN; const ref = col(c) + (r + 1); return isFinite(num) && /\d/.test(v) && v.length < 16 ? '<c r="' + ref + '"><v>' + num + '</v></c>' : '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + xe(v) + '</t></is></c>'; }).join('') + '</row>').join('') + '</sheetData></worksheet>');
    z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + sheets.map((_, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>');
    z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    z.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets.map((_, i) => '<sheet name="Trang ' + (i + 1) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets></workbook>');
    z.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') + '</Relationships>');
    sheets.forEach((s, i) => z.file('xl/worksheets/sheet' + (i + 1) + '.xml', s));
    const bytes = await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }); b.done();
    if (noSave) return { bytes, pages };
    if (!pages.some(g => g.length)) toast('PDF không có lớp chữ (bản scan) — hãy chạy OCR trước để có dữ liệu', 'err');
    await T.saveOut(d, (d.name || 'pdf').replace(/\.pdf$/i, '') + '.xlsx', bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); return { bytes, pages };
  } catch (e) { b.done(); toast('Lỗi: ' + e.message, 'err'); }
}
// PDF/A-2B: XMP (pdfaid), sRGB OutputIntent, document ID, no encryption; optional rasterize when fonts are not embedded
export async function makePdfA(bytes, opts) {
  opts = opts || {}; const PL = await PDFLib(); const ICC = (await import('./srgb_icc.js')).default;
  const p = await PL.PDFDocument.load(bytes, { updateMetadata: false }); const N = PL.PDFName, ctx = p.context;
  const title = p.getTitle() || opts.title || 'Tài liệu', author = p.getAuthor() || '', now = new Date(); const iso = now.toISOString().replace(/\.\d+Z$/, 'Z');
  p.setProducer('Yomi-Office Open Source 0.0.8'); p.setCreator(p.getCreator() || 'Yomi-Office Open Source'); p.setModificationDate(now); if (!p.getCreationDate()) p.setCreationDate(now); p.setTitle(title);
  const xe = s => String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const xmp = '<?xpacket begin="﻿" id="W5M0MpCehiHzreSzNTczkc9d"?><x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">' +
    '<rdf:Description rdf:about="" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"><pdfaid:part>2</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance></rdf:Description>' +
    '<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:format>application/pdf</dc:format><dc:title><rdf:Alt><rdf:li xml:lang="x-default">' + xe(title) + '</rdf:li></rdf:Alt></dc:title>' + (author ? '<dc:creator><rdf:Seq><rdf:li>' + xe(author) + '</rdf:li></rdf:Seq></dc:creator>' : '') + '</rdf:Description>' +
    '<rdf:Description rdf:about="" xmlns:xmp="http://ns.adobe.com/xap/1.0/"><xmp:CreatorTool>Yomi-Office Open Source</xmp:CreatorTool><xmp:CreateDate>' + iso + '</xmp:CreateDate><xmp:ModifyDate>' + iso + '</xmp:ModifyDate><xmp:MetadataDate>' + iso + '</xmp:MetadataDate></rdf:Description>' +
    '<rdf:Description rdf:about="" xmlns:pdf="http://ns.adobe.com/pdf/1.3/"><pdf:Producer>Yomi-Office Open Source 0.0.8</pdf:Producer></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end="w"?>';
  const meta = ctx.stream(new TextEncoder().encode(xmp), { Type: 'Metadata', Subtype: 'XML' }); p.catalog.set(N.of('Metadata'), ctx.register(meta));
  const icc = ctx.flateStream(bytesFromB64(ICC), { N: 3 }); const iccRef = ctx.register(icc);
  const oi = ctx.obj({ Type: 'OutputIntent', S: 'GTS_PDFA1', OutputConditionIdentifier: PL.PDFString.of('sRGB IEC61966-2.1'), Info: PL.PDFString.of('sRGB IEC61966-2.1'), DestOutputProfile: iccRef });
  p.catalog.set(N.of('OutputIntents'), ctx.obj([ctx.register(oi)]));
  const idb = new Uint8Array(16); for (let i = 0; i < 16; i++) idb[i] = Math.floor(Math.random() * 256); const hex = Array.prototype.map.call(idb, x => ('0' + x.toString(16)).slice(-2)).join('');
  ctx.trailerInfo.ID = ctx.obj([PL.PDFHexString.of(hex), PL.PDFHexString.of(hex)]);
  p.catalog.delete(N.of('OpenAction')); p.catalog.delete(N.of('JavaScript'));
  return p.save({ useObjectStreams: false });
}
export async function toPdfA(d) {
  if (!d.pdf) return toast('Mở một PDF trước', 'err');
  const fonts = await listFonts(d); const notEmb = fonts.filter(f => !f.embedded);
  const v = await formDlg('Chuyển sang PDF/A-2B (lưu trữ lâu dài)', [{ key: 'mode', type: 'radio', label: notEmb.length ? 'Có ' + notEmb.length + ' phông chưa nhúng (' + notEmb.slice(0, 3).map(f => f.name).join(', ') + ') — PDF/A yêu cầu nhúng phông:' : 'Cách chuyển:', value: notEmb.length ? 'raster' : 'keep', options: [['keep', 'Giữ nguyên nội dung (giữ chữ chọn được)'], ['raster', 'Chuyển trang thành ảnh (chắc chắn hợp lệ, không chọn được chữ)']] }]); if (!v) return;
  const b = busy('Đang tạo PDF/A…');
  try { let src = d.bytes; if (d._password) { const P = await PRO(); src = await P.decryptPdf(d.bytes, d._password); } if (v.mode === 'raster') src = await T.flatten(d);
    const out = await makePdfA(src, { title: (d.name || '').replace(/\.pdf$/i, '') }); b.done(); await T.saveOut(d, (d.name || 'pdf').replace(/\.pdf$/i, '') + '-PDFA.pdf', out, 'application/pdf'); }
  catch (e) { b.done(); toast('Lỗi PDF/A: ' + e.message, 'err'); }
}

// ================================================================ page tools: crop, repair, reverse, blank page, time stamp
export async function cropDialog(d) {
  if (!d.pdf) return toast('Mở một PDF trước', 'err');
  const v = await formDlg('Cắt cúp PDF (cắt lề)', [{ key: 't', label: 'Cắt trên (mm)', value: '10' }, { key: 'b', label: 'Cắt dưới (mm)', value: '10' }, { key: 'l', label: 'Cắt trái (mm)', value: '10' }, { key: 'r', label: 'Cắt phải (mm)', value: '10' }, { key: 'scope', type: 'radio', label: 'Áp dụng cho', value: 'all', options: [['all', 'Tất cả trang'], ['cur', 'Trang hiện tại / trang đã chọn']] }]); if (!v) return;
  const PL = await PDFLib(); const p = await PL.PDFDocument.load(d.bytes); const pt = x => (parseFloat(x) || 0) / MM;
  const idx = v.scope === 'all' ? p.getPages().map((_, i) => i) : (d.sel.length ? d.sel : [d.cur || 0]);
  idx.forEach(i => { const pg = p.getPage(i); const mb = pg.getCropBox ? pg.getCropBox() : pg.getMediaBox(); const w = mb.width - pt(v.l) - pt(v.r), hh = mb.height - pt(v.t) - pt(v.b); if (w < 20 || hh < 20) return; pg.setCropBox(mb.x + pt(v.l), mb.y + pt(v.b), w, hh); });
  await d.commit(await p.save(), 'Đã cắt lề ' + idx.length + ' trang');
}
export async function repairPdf(d) {
  if (!d.bytes) return toast('Mở một PDF trước', 'err'); const b = busy('Đang sửa tệp PDF…'); const PL = await PDFLib();
  try { const p = await PL.PDFDocument.load(d.bytes, { ignoreEncryption: true, throwOnInvalidObject: false }); const nd = await PL.PDFDocument.create(); const pages = await nd.copyPages(p, p.getPageIndices()); pages.forEach(x => nd.addPage(x)); try { nd.setTitle(p.getTitle() || ''); } catch (e) {} const out = await nd.save(); b.done(); await d.commit(out, 'Đã dựng lại cấu trúc PDF (' + pages.length + ' trang)'); }
  catch (e) { // last resort: rebuild from what pdf.js can still render
    try { const out = await T.flatten(d); b.done(); await d.commit(out, 'Cấu trúc hỏng nặng — đã dựng lại từ ảnh trang'); } catch (e2) { b.done(); toast('Không sửa được: ' + e.message, 'err'); } }
}
export async function reverseOrder(d) { const PL = await PDFLib(); const src = await PL.PDFDocument.load(d.bytes); const nd = await PL.PDFDocument.create(); const idx = src.getPageIndices().reverse(); (await nd.copyPages(src, idx)).forEach(p => nd.addPage(p)); await d.commit(await nd.save(), 'Đã đảo thứ tự trang'); }
export async function insertBlank(d) { const PL = await PDFLib(); const p = await PL.PDFDocument.load(d.bytes); const at = (d.cur || 0) + 1; const ref = p.getPage(d.cur || 0); p.insertPage(at, [ref.getWidth(), ref.getHeight()]); await d.commit(await p.save(), 'Đã chèn trang trống sau trang ' + at); setTimeout(() => d.goto(at), 300); }
function stampPng(text, size, color) { const c = document.createElement('canvas'); const x = c.getContext('2d'); const f = size + 'px Arial, "Segoe UI", sans-serif'; x.font = f; c.width = Math.ceil(x.measureText(text).width) + 8; c.height = Math.ceil(size * 1.35); x.font = f; x.fillStyle = color; x.textBaseline = 'top'; x.fillText(text, 4, 2); return c.toDataURL('image/png'); }
export async function timeStamp(d) {
  const now = new Date(); const pad = n => (n < 10 ? '0' : '') + n; const def = 'Đã xem lúc ' + pad(now.getHours()) + ':' + pad(now.getMinutes()) + ' ngày ' + pad(now.getDate()) + '/' + pad(now.getMonth() + 1) + '/' + now.getFullYear();
  const v = await formDlg('Đóng dấu thời gian', [{ key: 't', label: 'Nội dung', value: def }, { key: 'pos', type: 'select', label: 'Vị trí', value: 'br', options: [['br', 'Dưới phải'], ['bl', 'Dưới trái'], ['tr', 'Trên phải'], ['tl', 'Trên trái']] }, { key: 'all', type: 'checkbox', label: 'Mọi trang', value: true }]); if (!v) return;
  const PL = await PDFLib(); const p = await PL.PDFDocument.load(d.bytes); const img = await p.embedPng(dataUrlToBytes(stampPng(v.t, 40, '#c00000')));
  const idx = v.all ? p.getPageIndices() : [d.cur || 0];
  idx.forEach(i => { const pg = p.getPage(i); const w = img.width / 4, hh = img.height / 4; const W = pg.getWidth(), H = pg.getHeight(); const x = /r/.test(v.pos) ? W - w - 24 : 24, y = /b/.test(v.pos) ? 18 : H - hh - 18; pg.drawImage(img, { x, y, width: w, height: hh }); });
  await d.commit(await p.save(), 'Đã đóng dấu thời gian');
}

// ================================================================ forms: reset / export / import data
async function formOf(d) { const PL = await PDFLib(); const p = await PL.PDFDocument.load(d.bytes); const f = p.getForm(); return { PL, p, f, fields: f.getFields() }; }
function fieldVal(PL, fl) { try { if (fl instanceof PL.PDFTextField) return fl.getText() || ''; if (fl instanceof PL.PDFCheckBox) return fl.isChecked(); if (fl instanceof PL.PDFDropdown || fl instanceof PL.PDFOptionList) return fl.getSelected().join(', '); if (fl instanceof PL.PDFRadioGroup) return fl.getSelected() || ''; } catch (e) {} return ''; }
export async function resetForm(d) {
  const { PL, p, fields } = await formOf(d); if (!fields.length) return toast('PDF không có biểu mẫu');
  fields.forEach(fl => { try { if (fl instanceof PL.PDFTextField) fl.setText(''); else if (fl instanceof PL.PDFCheckBox) fl.uncheck(); else if (fl instanceof PL.PDFDropdown || fl instanceof PL.PDFOptionList) fl.clear(); else if (fl instanceof PL.PDFRadioGroup) fl.clear(); } catch (e) {} });
  p.getForm().acroForm.dict.set(PL.PDFName.of('NeedAppearances'), PL.PDFBool.True);
  await d.commit(await p.save(), 'Đã đặt lại ' + fields.length + ' trường biểu mẫu');
}
export async function exportFormData(d, fmt) {
  const { PL, fields } = await formOf(d); if (!fields.length) return toast('PDF không có biểu mẫu');
  const rows = fields.map(fl => [fl.getName(), fieldVal(PL, fl)]); const base = (d.name || 'bieu-mau').replace(/\.pdf$/i, '');
  if (fmt === 'csv') { const csv = '﻿Trường,Giá trị\n' + rows.map(r => r.map(x => '"' + String(x).replace(/"/g, '""') + '"').join(',')).join('\n'); return T.saveOut(d, base + '-du-lieu.csv', utf8Encode(csv), 'text/csv'); }
  if (fmt === 'sheet') { const csv = rows.map(r => r.map(x => '"' + String(x).replace(/"/g, '""') + '"').join(',')).join('\n'); return d.shell.openFile({ name: base + '-du-lieu.csv', bytes: utf8Encode('Trường,Giá trị\n' + csv) }); }
  const xfdf = '<?xml version="1.0" encoding="UTF-8"?>\n<xfdf xmlns="http://ns.adobe.com/xfdf/" xml:space="preserve"><fields>' + rows.map(r => '<field name="' + esc(r[0]) + '"><value>' + esc(typeof r[1] === 'boolean' ? (r[1] ? 'Yes' : 'Off') : r[1]) + '</value></field>').join('') + '</fields></xfdf>';
  return T.saveOut(d, base + '.xfdf', utf8Encode(xfdf), 'application/vnd.adobe.xfdf');
}
export async function importFormData(d) {
  const f = await pickFiles('.xfdf,.xml,.json,.csv', false); if (!f || !f.length) return; const txt = new TextDecoder().decode(f[0].bytes);
  const map = {};
  if (/^\s*[{[]/.test(txt)) Object.assign(map, JSON.parse(txt));
  else if (/<xfdf|<field/.test(txt)) { const doc = new DOMParser().parseFromString(txt, 'application/xml'); Array.prototype.forEach.call(doc.getElementsByTagName('field'), fe => { const v = fe.getElementsByTagName('value')[0]; map[fe.getAttribute('name')] = v ? v.textContent : ''; }); }
  else txt.split(/\r?\n/).slice(1).forEach(l => { const m = /^"?((?:[^"]|"")*)"?,"?((?:[^"]|"")*)"?$/.exec(l); if (m) map[m[1].replace(/""/g, '"')] = m[2].replace(/""/g, '"'); });
  const { PL, p, fields } = await formOf(d); let n = 0;
  fields.forEach(fl => { const k = fl.getName(); if (!(k in map)) return; const v = map[k]; try { if (fl instanceof PL.PDFTextField) fl.setText(String(v)); else if (fl instanceof PL.PDFCheckBox) (v === true || /^(yes|on|true|1|có)$/i.test(v) ? fl.check() : fl.uncheck()); else if (fl instanceof PL.PDFDropdown || fl instanceof PL.PDFRadioGroup) fl.select(String(v)); n++; } catch (e) {} });
  p.getForm().acroForm.dict.set(PL.PDFName.of('NeedAppearances'), PL.PDFBool.True);
  await d.commit(await p.save(), 'Đã nhập ' + n + ' trường');
}
function pickFiles(accept, multi) { return Host.openFiles(accept, !!multi); }

// ================================================================ share
export function shareMenu(d) {
  return [
    { label: 'Gửi qua Email / ứng dụng khác', icon: 'mail', run: () => { if (!Host.share(d.name || 'tai-lieu.pdf', d.bytes, 'application/pdf')) { T.saveOut(d, d.name || 'tai-lieu.pdf', d.bytes, 'application/pdf').then(() => Host.openUrl('mailto:?subject=' + encodeURIComponent(d.name || 'Tài liệu PDF') + '&body=' + encodeURIComponent('Tệp PDF đính kèm: ' + (d.name || '')))); toast('Đã lưu tệp — hãy đính kèm vào email vừa mở'); } } },
    { label: 'Lưu lên đám mây (Drive / OneDrive / Dropbox)…', icon: 'cloud-upload', run: () => import('../core/cloud.js').then(C => C.cloudSaveDialog(d.shell, d)) },
    { label: 'Sao chép toàn bộ văn bản', icon: 'copy', run: async () => { const t = await T.extractText(d); try { await navigator.clipboard.writeText(t); toast('Đã sao chép ' + t.length + ' ký tự'); } catch (e) { toast('Không truy cập được bộ nhớ tạm', 'err'); } } }
  ];
}
