// PowerPoint ribbon: Home, Insert, Draw, Design, Transitions, Animations, Slide Show, Review, View, Help (+ contextual Shape/Picture/Chart/Table)
import { h, ic } from '../core/util.js';
import { toast, showMenu, colorPicker, dialog, promptDlg, formDlg } from '../core/ui.js';
import { settings } from '../core/store.js';
import * as P from './cmds.js';
import { helpTab } from '../core/help.js';
import { drawTab } from '../core/ink.js';
import { symbolDialog, qrDialog, iconsDialog } from '../core/insert.js';
import { SHAPES, shapeSvg } from '../core/shapes.js';
import { THEMES, LAYOUTS } from './templates.js';

import { FONT_LIST as FONTS } from '../core/fonts.js';
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 54, 60, 66, 72, 88, 96];
const sel = d => d.selectedObj;
const isText = d => sel(d) && (sel(d).type === 'text' || sel(d).type === 'shape');

export function ribbonDef(doc) {
  const tabs = [homeTab(doc), insertTab(doc), drawTab(doc, doc.getInk()), designTab(doc), transitionsTab(doc), animationsTab(doc), showTab(doc), recordTab(doc), reviewTab(doc), viewTab(doc), developerTab(doc), helpTab(doc, 'ppt')];
  const o = doc.selectedObj;
  if (o) {
    if (o.type === 'shape' || o.type === 'text') tabs.push(shapeFormatTab(doc));
    else if (o.type === 'image') tabs.push(pictureFormatTab(doc));
    else if (o.type === 'chart') tabs.push(chartFormatTab(doc));
    else if (o.type === 'table') tabs.push(tableFormatTab(doc));
  }
  return { tabs };
}

function homeTab(d) {
  return { id: 'home', label: 'Trang đầu', en: 'Home', groups: [
    { label: 'Bảng tạm', items: [
      { type: 'split', big: true, label: 'Dán', icon: 'clipboard-paste', sc: 'Ctrl+V', run: () => P.pasteObj(d), menu: () => [{ label: 'Dán đối tượng', icon: 'clipboard-paste', disabled: !d.clip, run: () => P.pasteObj(d) }] },
      { type: 'col', items: [
        { type: 'btn', label: 'Cắt', icon: 'scissors', sc: 'Ctrl+X', run: () => { if (sel(d)) { d.clip = JSON.parse(JSON.stringify(sel(d))); P.removeObj(d, sel(d)); } } },
        { type: 'btn', label: 'Sao chép', icon: 'copy', sc: 'Ctrl+C', run: () => { if (sel(d)) { d.clip = JSON.parse(JSON.stringify(sel(d))); toast('Đã sao chép'); } } },
        { type: 'btn', label: 'Sao định dạng', icon: 'paintbrush', run: () => toast('Chọn đối tượng nguồn rồi áp dụng — đang phát triển') }
      ] }
    ] },
    { label: 'Trang chiếu', items: [
      { type: 'drop', big: true, label: 'Trang chiếu Mới', icon: 'plus-square', sc: 'Ctrl+M', run: () => P.newSlide(d), menu: () => P.newSlideMenu(d) },
      { type: 'col', items: [
        { type: 'drop', label: 'Bố cục', icon: 'layout-template', menu: () => P.layoutMenu(d) },
        { type: 'btn', label: 'Đặt lại', icon: 'rotate-ccw', run: () => P.resetSlide(d) },
        { type: 'drop', label: 'Phần', icon: 'folder', menu: () => [{ label: 'Thêm phần', run: () => toast('Phần: nhóm các trang chiếu') }] }
      ] }
    ] },
    { label: 'Phông chữ', launcher: () => fontDialog(d), items: [
      { type: 'col', items: [
        { type: 'row', items: [
          { type: 'combo', width: 128, title: 'Phông chữ', fontBox: true, get: () => (sel(d) && sel(d).font) || d.theme.font || 'Calibri', set: (a, v) => P.setFont(d, v), options: () => FONTS, optStyle: v => 'font-family:' + v },
          { type: 'combo', width: 46, title: 'Cỡ chữ', get: () => String((sel(d) && sel(d).fontSize) || 18), set: (a, v) => P.setFontSize(d, parseFloat(v)), options: () => SIZES.map(String) },
          { type: 'btn', icon: 'a-arrow-up', title: 'Tăng cỡ chữ', run: () => P.sizeStep(d, 1) },
          { type: 'btn', icon: 'a-arrow-down', title: 'Giảm cỡ chữ', run: () => P.sizeStep(d, -1) }
        ] },
        { type: 'row', items: [
          { type: 'btn', icon: 'bold', title: 'Đậm', sc: 'Ctrl+B', active: () => sel(d) && sel(d).b, run: () => P.fmtText(d, 'bold') },
          { type: 'btn', icon: 'italic', title: 'Nghiêng', sc: 'Ctrl+I', active: () => sel(d) && sel(d).i, run: () => P.fmtText(d, 'italic') },
          { type: 'btn', icon: 'underline', title: 'Gạch chân', sc: 'Ctrl+U', run: () => P.fmtText(d, 'underline') },
          { type: 'btn', icon: 'strikethrough', title: 'Gạch ngang', run: () => P.fmtText(d, 'strikeThrough') },
          { type: 'btn', icon: 'subscript', title: 'Chỉ số dưới', run: () => P.fmtText(d, 'subscript') },
          { type: 'btn', icon: 'superscript', title: 'Chỉ số trên', run: () => P.fmtText(d, 'superscript') },
          { type: 'color', icon: 'baseline', title: 'Màu chữ', colorBar: () => d._fc || '#c00000', run: () => P.textColor(d, d._fc || '#c00000'), pick: (a, c) => { d._fc = c; P.textColor(d, c); }, colorOpts: { auto: 'Tự động' } }
        ] }
      ] }
    ] },
    { label: 'Đoạn văn', items: [
      { type: 'col', items: [
        { type: 'row', items: [
          { type: 'btn', icon: 'list', title: 'Dấu đầu dòng', run: () => P.fmtText(d, 'insertUnorderedList') },
          { type: 'btn', icon: 'list-ordered', title: 'Đánh số', run: () => P.fmtText(d, 'insertOrderedList') },
          { type: 'btn', icon: 'indent-decrease', title: 'Giảm thụt lề', run: () => P.fmtText(d, 'outdent') },
          { type: 'btn', icon: 'indent-increase', title: 'Tăng thụt lề', run: () => P.fmtText(d, 'indent') }
        ] },
        { type: 'row', items: [
          { type: 'btn', icon: 'align-left', title: 'Căn trái', active: () => sel(d) && sel(d).align === 'left', run: () => P.setAlign(d, 'left') },
          { type: 'btn', icon: 'align-center', title: 'Căn giữa', active: () => sel(d) && sel(d).align === 'center', run: () => P.setAlign(d, 'center') },
          { type: 'btn', icon: 'align-right', title: 'Căn phải', active: () => sel(d) && sel(d).align === 'right', run: () => P.setAlign(d, 'right') },
          { type: 'btn', icon: 'align-justify', title: 'Căn đều', run: () => P.setAlign(d, 'justify') }
        ] }
      ] }
    ] },
    { label: 'Vẽ', items: [
      { type: 'gallery', width: 220, items: () => SHAPES.slice(0, 12).map(s => ({ title: s[1], preview: () => shapePrev(s[0], d.theme.colors[0]), run: () => P.addShape(d, s[0]) })), more: () => SHAPES.map(s => ({ label: s[1], icon: 'shapes', run: () => P.addShape(d, s[0]) })) },
      { type: 'col', items: [
        { type: 'color', label: 'Tô hình', icon: 'paint-bucket', colorBar: () => d._shfill || d.theme.colors[0], run: () => fillSel(d, d._shfill || d.theme.colors[0]), pick: (a, c) => { d._shfill = c; fillSel(d, c); } },
        { type: 'color', label: 'Viền hình', icon: 'square', colorBar: () => d._shstroke || '#2f528f', run: () => strokeSel(d, d._shstroke || '#2f528f'), pick: (a, c) => { d._shstroke = c; strokeSel(d, c); } },
        { type: 'drop', label: 'Sắp xếp', icon: 'layout-template', menu: () => arrangeMenu(d) }
      ] }
    ] },
    { label: 'Chỉnh sửa', items: [
      { type: 'col', items: [
        { type: 'btn', label: 'Tìm', icon: 'search', sc: 'Ctrl+F', run: () => findDialog(d) },
        { type: 'btn', label: 'Thay thế', icon: 'replace', run: () => replaceDialog(d) },
        { type: 'btn', label: 'Chọn', icon: 'box-select', run: () => selectionPane(d) }
      ] }
    ] },
    { label: 'A.I', items: [
      { type: 'big', label: 'Nhà thiết kế (A.I)', icon: 'wand-sparkles', run: () => designerDialog(d) },
      { type: 'big', label: 'Copilot', icon: 'sparkles', run: () => import('../ai/ai.js').then(m => m.openAi(d.shell, d)) }
    ] }
  ] };
}

function insertTab(d) {
  return { id: 'insert', label: 'Chèn', en: 'Insert', groups: [
    { label: 'Trang chiếu', items: [{ type: 'drop', big: true, label: 'Trang chiếu Mới', icon: 'plus-square', run: () => P.newSlide(d), menu: () => P.newSlideMenu(d) }] },
    { label: 'Bảng', items: [{ type: 'drop', big: true, label: 'Bảng', icon: 'table', menu: () => [{ label: 'Chèn bảng…', icon: 'table', run: () => P.addTable(d) }, { label: 'Bảng 2×2', run: () => quickTable(d, 2, 2) }, { label: 'Bảng 3×3', run: () => quickTable(d, 3, 3) }, { label: 'Bảng 4×5', run: () => quickTable(d, 4, 5) }] }] },
    { label: 'Hình ảnh', items: [
      { type: 'big', label: 'Ảnh', icon: 'image', run: () => P.addImage(d) },
      { type: 'drop', big: true, label: 'Hình dạng', icon: 'shapes', menu: () => shapeMenu(d) },
      { type: 'big', label: 'Biểu tượng', icon: 'smile', run: () => iconsDialog(url => P.addIcon(d, url)) },
      { type: 'drop', big: true, label: 'Biểu đồ', icon: 'chart-column', menu: () => [['column', 'Cột'], ['bar', 'Thanh'], ['line', 'Đường'], ['pie', 'Tròn'], ['area', 'Vùng'], ['scatter', 'Phân tán']].map(k => ({ label: k[1], icon: 'chart-' + (k[0] === 'bar' ? 'bar' : k[0] === 'line' ? 'line' : k[0] === 'pie' ? 'pie' : 'column'), run: () => P.addChart(d, k[0]) })) }
    ] },
    { label: 'Văn bản', items: [
      { type: 'big', label: 'Hộp văn bản', icon: 'type', run: () => P.addTextbox(d) },
      { type: 'drop', big: true, label: 'Đầu & Chân trang', icon: 'panel-top', menu: () => [{ label: 'Số trang chiếu', run: () => headerFooter(d, 'number') }, { label: 'Ngày và giờ', run: () => headerFooter(d, 'date') }, { label: 'Chân trang…', run: () => headerFooter(d, 'footer') }] },
      { type: 'big', label: 'WordArt', icon: 'whole-word', run: () => wordArt(d) }
    ] },
    { label: 'Ký hiệu', items: [
      { type: 'big', label: 'Phương trình (MathType)', icon: 'sigma-square', run: () => P.addEquation(d) },
      { type: 'drop', big: true, label: 'Ký hiệu', icon: 'omega', menu: () => [{ label: 'Thêm ký hiệu…', run: () => symbolDialog(ch => insertChar(d, ch)) }].concat(['©', '®', '™', '°', '→', '★', '•'].map(c => ({ label: c, run: () => insertChar(d, c) }))) }
    ] },
    { label: 'Phương tiện', items: [
      { type: 'big', label: 'Video', icon: 'video', run: () => P.addVideo(d) },
      { type: 'big', label: 'Âm thanh', icon: 'volume-2', run: () => toast('Âm thanh: sẽ phát khi trình chiếu — đang phát triển') },
      { type: 'big', label: 'Mã QR', icon: 'qr-code', run: () => qrDialog(url => P.addObj(d, { type: 'image', src: url, x: 540, y: 220, w: 280, h: 280, alt: 'QR' })) }
    ] },
    { label: 'Liên kết', items: [{ type: 'big', label: 'Liên kết', icon: 'link', sc: 'Ctrl+K', run: () => linkObj(d) }, { type: 'big', label: 'Hành động', icon: 'mouse-pointer-click', run: () => actionDialog(d) }] }
  ] };
}

function designTab(d) {
  return { id: 'design', label: 'Thiết kế', en: 'Design', groups: [
    { label: 'Chủ đề', items: [{ type: 'gallery', width: 420, items: () => THEMES.map(t => ({ title: t.name, preview: () => themePrev(t), run: () => applyTheme(d, t) })), more: () => THEMES.map(t => ({ label: t.name, swatch: t.colors[0], run: () => applyTheme(d, t) })) }] },
    { label: 'Biến thể', items: [
      { type: 'drop', big: true, label: 'Màu', icon: 'palette', menu: () => THEMES.map(t => ({ label: t.name, swatch: t.colors[0], run: () => applyTheme(d, t) })) },
      { type: 'drop', big: true, label: 'Phông chữ', icon: 'type', menu: () => FONTS.map(f => ({ label: f, style: 'font-family:' + f, run: () => { d.change('Phông chủ đề', () => d.theme = Object.assign({}, d.theme, { font: f })); } })) },
      { type: 'drop', big: true, label: 'Nền', icon: 'image', menu: () => [{ label: 'Định dạng nền…', run: () => P.bgDialog(d) }, { label: 'Nền trắng', run: () => d.change('Nền', () => d.slides.forEach(s => s.bg = '#ffffff')) }, { label: 'Nền theo chủ đề', run: () => d.change('Nền', () => d.slides.forEach(s => s.bg = null)) }] }
    ] },
    { label: 'Tùy chỉnh', items: [
      { type: 'drop', big: true, label: 'Cỡ Trang chiếu', icon: 'ratio', menu: () => [['16:9', 'Màn hình rộng (16:9)'], ['4:3', 'Chuẩn (4:3)']].map(r => ({ label: r[1], checked: d.ratio === r[0] || undefined, run: () => { d.ratio = r[0]; toast('Tỉ lệ: ' + r[0] + ' (áp dụng khi xuất PPTX)'); } })) },
      { type: 'big', label: 'Định dạng Nền', icon: 'paint-bucket', run: () => P.bgDialog(d) }
    ] },
    { label: 'Nhà thiết kế', items: [{ type: 'big', label: 'Ý tưởng Thiết kế', icon: 'wand-sparkles', run: () => designerDialog(d) }] }
  ] };
}

function transitionsTab(d) {
  return { id: 'transitions', label: 'Chuyển tiếp', en: 'Transitions', groups: [
    { label: 'Xem trước', items: [{ type: 'big', label: 'Xem trước', icon: 'play', run: () => previewTransition(d) }] },
    { label: 'Chuyển tiếp sang Trang chiếu này', items: [
      { type: 'gallery', width: 420, items: () => P.TRANSITIONS.map(t => ({ title: t[1], preview: () => h('div.tr-prev', ic(t[2], 22)), run: () => P.setTransition(d, t[0]) })), more: () => P.TRANSITIONS.map(t => ({ label: t[1], icon: t[2], checked: (d.slide.transition || 'none') === t[0] || undefined, run: () => P.setTransition(d, t[0]) })) }
    ] },
    { label: 'Thời gian', items: [
      { type: 'col', items: [
        { type: 'drop', label: 'Âm thanh', icon: 'volume-2', menu: () => ['Không', 'Click', 'Chime', 'Whoosh'].map(s => ({ label: s, run: () => toast('Âm thanh chuyển tiếp: ' + s) })) },
        { type: 'drop', label: 'Thời lượng', icon: 'timer', menu: () => [0.5, 0.8, 1, 1.5, 2].map(s => ({ label: s + ' giây', run: () => d.change('Thời lượng', () => d.slide.trDur = s) })) }
      ] },
      { type: 'col', items: [
        { type: 'btn', label: 'Áp dụng cho Tất cả', icon: 'copy-plus', run: () => P.setTransition(d, d.slide.transition || 'none', true) },
        { type: 'toggle', label: 'Khi bấm chuột', icon: 'mouse-pointer-click', get: () => d.slide.advClick !== false, set: v => d.change('Nâng cao', () => d.slide.advClick = v) },
        { type: 'toggle', label: 'Tự động sau', icon: 'timer', get: () => !!d.slide.advTime, set: v => { if (v) promptDlg('Tự chuyển sau (giây):', '3', 'Thời gian').then(x => { if (x) d.change('Tự chuyển', () => d.slide.advTime = parseFloat(x)); }); else d.change('Tự chuyển', () => d.slide.advTime = 0); } }
      ] }
    ] }
  ] };
}

function animationsTab(d) {
  return { id: 'animations', label: 'Hoạt ảnh', en: 'Animations', groups: [
    { label: 'Xem trước', items: [{ type: 'big', label: 'Xem trước', icon: 'play', run: () => previewAnim(d) }] },
    { label: 'Hoạt ảnh', items: [{ type: 'gallery', width: 360, items: () => P.ANIMS.map(a => ({ title: a[1], preview: () => h('div.tr-prev', ic('sparkle', 18)), run: () => P.setAnimation(d, a[0]) })), more: () => P.animationMenu(d) }] },
    { label: 'Hoạt ảnh nâng cao', items: [
      { type: 'big', label: 'Thêm Hoạt ảnh', icon: 'plus', run: () => showMenu(P.animationMenu(d), menuPos()) },
      { type: 'col', items: [
        { type: 'btn', label: 'Ngăn Hoạt ảnh', icon: 'panel-right', run: () => animPane(d) },
        { type: 'btn', label: 'Kích hoạt', icon: 'mouse-pointer-click', run: () => toast('Hoạt ảnh chạy khi bấm chuột lúc trình chiếu') }
      ] }
    ] },
    { label: 'Thời gian', items: [
      { type: 'col', items: [
        { type: 'drop', label: 'Bắt đầu', icon: 'play', menu: () => ['Khi bấm chuột', 'Cùng lúc trước', 'Sau mục trước'].map(s => ({ label: s, run: () => { const o = sel(d); if (o) d.change('Bắt đầu', () => o.animStart = s); } })) },
        { type: 'drop', label: 'Thời lượng', icon: 'timer', menu: () => [0.3, 0.5, 1, 2].map(s => ({ label: s + 's', run: () => { const o = sel(d); if (o) d.change('Thời lượng', () => o.animDur = s); } })) }
      ] }
    ] }
  ] };
}

function showTab(d) {
  return { id: 'show', label: 'Trình chiếu', en: 'Slide Show', groups: [
    { label: 'Bắt đầu Trình chiếu', items: [
      { type: 'big', label: 'Từ Đầu', icon: 'play', sc: 'F5', run: () => P.startShow(d, 0) },
      { type: 'big', label: 'Từ Trang Hiện tại', icon: 'play-circle', sc: 'Shift+F5', run: () => P.startShow(d, d.cur) },
      { type: 'big', label: 'Trình chiếu Tùy chỉnh', icon: 'list-video', run: () => toast('Trình chiếu tùy chỉnh: chọn tập hợp trang — đang phát triển') }
    ] },
    { label: 'Thiết lập', items: [
      { type: 'big', label: 'Thiết lập Trình chiếu', icon: 'settings', run: () => showSetup(d) },
      { type: 'col', items: [
        { type: 'btn', label: 'Ẩn Trang chiếu', icon: 'eye-off', active: () => d.slide.hidden, run: () => d.change('Ẩn trang', () => d.slide.hidden = !d.slide.hidden) },
        { type: 'btn', label: 'Ghi lại', icon: 'circle-dot', run: () => toast('Ghi hình trình chiếu — đang phát triển') },
        { type: 'btn', label: 'Diễn tập Thời gian', icon: 'timer', run: () => toast('Diễn tập: đo thời gian mỗi trang') }
      ] }
    ] },
    { label: 'Màn hình & Android DeX', items: [
      { type: 'big', label: 'Diễn giả 2 Màn hình', icon: 'presentation', title: 'Màn hình ngoài chiếu slide, điện thoại hiện ghi chú & điều khiển', run: () => import('../core/dex.js').then(D => D.startDualScreenPresenter(d)) },
      { type: 'col', items: [
        { type: 'btn', label: 'Bàn rê chuột ảo', icon: 'mouse-pointer', run: () => import('../core/dex.js').then(D => D.openVirtualTrackpad()) },
        { type: 'btn', label: 'Cài đặt DeX & Chiếu TV', icon: 'monitor', run: () => import('../core/dex.js').then(D => D.openDexControlCenterDialog()) }
      ] }
    ] }
  ] };
}

// v0.0.6: Record tab (Ghi) like PowerPoint 365
function recordTab(d) {
  const R = f => () => import('./record.js').then(m => f(m));
  return { id: 'record', label: 'Ghi', en: 'Record', groups: [
    { label: 'Camera', items: [{ type: 'big', label: 'Cameo', icon: 'video', title: 'Quay video từ camera và chèn vào slide', run: R(m => m.cameo(d)) }] },
    { label: 'Ghi', items: [
      { type: 'big', label: 'Từ Đầu', icon: 'circle-dot', title: 'Ghi lời thuyết minh + thời gian từ slide đầu', run: R(m => m.recordShow(d, 0)) },
      { type: 'big', label: 'Từ Slide Hiện tại', icon: 'circle-play', run: R(m => m.recordShow(d, d.cur)) },
      { type: 'big', label: 'Quay Màn hình', icon: 'monitor', run: R(m => m.screenRecording(d)) },
      { type: 'big', label: 'Âm thanh', icon: 'mic', run: R(m => m.recordAudio(d)) }
    ] },
    { label: 'Chỉnh sửa', items: [{ type: 'drop', big: true, label: 'Xóa Bản ghi', icon: 'eraser', menu: () => [{ label: 'Xóa bản ghi trên slide hiện tại', run: R(m => m.clearRecording(d, false)) }, { label: 'Xóa bản ghi trên mọi slide', run: R(m => m.clearRecording(d, true)) }] }, { type: 'toggle', label: 'Dùng thời gian đã ghi', icon: 'timer', get: () => d.useTimings !== false, set: v => { d.useTimings = v; } }] },
    { label: 'Xuất', items: [
      { type: 'big', label: 'Lưu thành Trình chiếu', icon: 'play', title: 'Lưu .ppsx (mở là trình chiếu ngay)', run: () => d.shell.save(d, true, 'ppsx') },
      { type: 'big', label: 'Xuất thành Video', icon: 'film', title: 'Xuất .mp4 / .webm', run: () => d.shell.save(d, true, 'video') }
    ] }
  ] };
}
function reviewTab(d) {
  return { id: 'review', label: 'Xem lại', en: 'Review', groups: [
    { label: 'Soát lỗi', items: [
      { type: 'big', label: 'Chính tả', icon: 'spell-check', sc: 'F7', run: () => import('../ai/ai.js').then(m => m.proofread(d)) },
      { type: 'big', label: 'Từ đồng nghĩa', icon: 'book-a', run: () => import('../ai/ai.js').then(m => m.openAi(d.shell, d, 'Gợi ý từ đồng nghĩa')) }
    ] },
    { label: 'Ngôn ngữ', items: [{ type: 'big', label: 'Dịch', icon: 'languages', run: () => import('../ai/ai.js').then(m => m.openAi(d.shell, d, 'Dịch nội dung trang chiếu sang tiếng Anh')) }] },
    { label: 'Chú thích', items: [
      { type: 'big', label: 'Chú thích Mới', icon: 'message-square-plus', run: () => addComment(d) },
      { type: 'col', items: [{ type: 'btn', label: 'Hiển thị', icon: 'messages-square', run: () => commentsPane(d) }, { type: 'btn', label: 'Xóa', icon: 'message-square-x', run: () => { d.change('Xóa chú thích', () => { (d.slide.comments || []).length = 0; }); } }] }
    ] },
    { label: 'So sánh', items: [{ type: 'big', label: 'Kiểm tra Trợ năng', icon: 'accessibility', run: () => d.shell.inspect(d) }] },
    { label: 'Bảo vệ', items: [{ type: 'big', label: 'Hạn chế Chỉnh sửa', icon: 'lock', active: () => d.readOnly, run: () => { d.setReadOnly(!d.readOnly); d.refreshUi(); toast(d.readOnly ? 'Đã khóa (chỉ đọc)' : 'Đã mở khóa'); } }] }
  ] };
}

function viewTab(d) {
  return { id: 'view', label: 'Xem', en: 'View', groups: [
    { label: 'Dạng xem Bản trình bày', items: [
      { type: 'big', label: 'Thường', icon: 'square', run: () => { const s = document.querySelector('.pp-sorter'); if (s) s.remove(); } },
      { type: 'big', label: 'Sắp xếp Trang chiếu', icon: 'layout-grid', run: () => P.slideSorter(d) },
      { type: 'big', label: 'Trang Ghi chú', icon: 'sticky-note', run: () => { d.notesEl.focus(); } },
      { type: 'big', label: 'Đọc', icon: 'book-open', run: () => P.startShow(d, d.cur) },
      { type: 'big', label: 'Android DeX', icon: 'monitor', title: 'Chế độ Android DeX Desktop Mode & Chiếu màn hình ngoài / PC Link', run: () => import('../core/dex.js').then(m => m.openDexGuideDialog()) },
      { type: 'drop', big: true, label: 'Bảo vệ mắt', icon: 'sun', title: 'Chế độ bảo vệ mắt (WPS Eye Protection)', menu: () => import('../core/suites.js').then(m => m.eyeProtectionMenu(d)) }
    ] },
    { label: 'Hiển thị', items: [
      { type: 'col', items: [
        { type: 'toggle', label: 'Thước', icon: 'ruler', get: () => !!d._ruler, set: v => { d._ruler = v; d.el.classList.toggle('pp-ruler-on', v); } },
        { type: 'toggle', label: 'Đường lưới', icon: 'grid-3x3', get: () => !!d._grid, set: v => { d._grid = v; d.canvas.classList.toggle('pp-grid-on', v); } },
        { type: 'toggle', label: 'Ghi chú', icon: 'panel-bottom', get: () => !d.el.classList.contains('pp-notes-off'), set: v => d.el.classList.toggle('pp-notes-off', !v) }
      ] },
      { type: 'big', label: 'Bộ định hướng', icon: 'compass', title: 'Bộ định hướng đối tượng (LibreOffice Navigator)', run: () => import('../core/suites.js').then(m => m.openNavigator(d)) }
    ] },
    { label: 'Thu phóng', items: [
      { type: 'big', label: 'Thu phóng', icon: 'zoom-in', run: () => showMenu([200, 150, 125, 100, 75, 50, 33].map(z => ({ label: z + '%', run: () => { d._userZoom = z !== 100; d.setZoom(z); } })), menuPos()) },
      { type: 'big', label: 'Vừa Cửa sổ', icon: 'maximize', run: () => { d._userZoom = false; d.fit(); } }
    ] },
    { label: 'Màu/Thang độ xám', items: [{ type: 'drop', big: true, label: 'Màu', icon: 'palette', menu: () => [['', 'Màu'], ['grayscale', 'Thang độ xám'], ['bw', 'Đen trắng']].map(m => ({ label: m[1], run: () => { d.canvas.style.filter = m[0] === 'grayscale' ? 'grayscale(1)' : m[0] === 'bw' ? 'grayscale(1) contrast(3)' : ''; } })) }] },
    { label: 'Cửa sổ', items: [{ type: 'big', label: 'Cửa sổ Mới', icon: 'app-window', run: () => import('../core/host.js').then(m => m.default.newWindow()) }] }
  ] };
}

// ---------------------------------------------------------------- contextual tabs
function shapeFormatTab(d) {
  return { id: 'fmt-shape', label: 'Định dạng Hình', en: 'Shape Format', contextual: true, accent: '#c43e1c', groups: [
    { label: 'Chèn Hình', items: [{ type: 'drop', big: true, label: 'Hình dạng', icon: 'shapes', menu: () => shapeMenu(d) }, { type: 'drop', big: true, label: 'Sửa Hình', icon: 'pencil-ruler', menu: () => shapeMenu(d, true) }] },
    { label: 'Kiểu Hình', items: [
      { type: 'gallery', width: 260, items: () => d.theme.colors.map(c => ({ title: '', preview: () => h('div', { style: { width: '38px', height: '26px', background: c, borderRadius: '4px' } }), run: () => fillSel(d, c) })) },
      { type: 'col', items: [
        { type: 'color', label: 'Tô Hình', icon: 'paint-bucket', colorBar: () => d._shfill || d.theme.colors[0], run: () => fillSel(d, d._shfill || d.theme.colors[0]), pick: (a, c) => { d._shfill = c; fillSel(d, c); } },
        { type: 'color', label: 'Đường viền', icon: 'square', colorBar: () => d._shstroke || '#2f528f', run: () => strokeSel(d, d._shstroke || '#2f528f'), pick: (a, c) => { d._shstroke = c; strokeSel(d, c); } },
        { type: 'drop', label: 'Hiệu ứng', icon: 'sparkles', menu: () => [['none', 'Không'], ['shadow', 'Đổ bóng'], ['glow', 'Phát sáng']].map(e => ({ label: e[1], run: () => effectSel(d, e[0]) })) }
      ] }
    ] },
    { label: 'WordArt', items: [{ type: 'color', big: true, label: 'Màu chữ', icon: 'baseline', colorBar: () => d._fc || '#000', run: () => P.textColor(d, d._fc || '#000'), pick: (a, c) => { d._fc = c; P.textColor(d, c); } }] },
    { label: 'Sắp xếp', items: [
      { type: 'col', items: [{ type: 'btn', label: 'Đưa lên trước', icon: 'bring-to-front', run: () => sel(d) && P.order(d, sel(d), 1) }, { type: 'btn', label: 'Đưa ra sau', icon: 'send-to-back', run: () => sel(d) && P.order(d, sel(d), -1) }] },
      { type: 'drop', big: true, label: 'Căn chỉnh', icon: 'align-center-horizontal', menu: () => alignMenu(d) },
      { type: 'col', items: [{ type: 'btn', label: 'Xoay 90°', icon: 'rotate-cw', run: () => { const o = sel(d); if (o) d.change('Xoay', () => o.rot = ((o.rot || 0) + 90) % 360); } }, { type: 'btn', label: 'Nhóm', icon: 'group', run: () => toast('Nhóm đối tượng — đang phát triển') }] }
    ] },
    { label: 'Kích thước', items: [{ type: 'col', items: [
      { type: 'combo', width: 64, title: 'Cao (px)', get: () => String(Math.round(sel(d) ? sel(d).h : 0)), set: (a, v) => { const o = sel(d); if (o) d.change('Chiều cao', () => o.h = parseFloat(v) || o.h); }, options: () => ['100', '150', '200', '300', '400'] },
      { type: 'combo', width: 64, title: 'Rộng (px)', get: () => String(Math.round(sel(d) ? sel(d).w : 0)), set: (a, v) => { const o = sel(d); if (o) d.change('Chiều rộng', () => o.w = parseFloat(v) || o.w); }, options: () => ['100', '200', '300', '400', '600'] }
    ] }] }
  ] };
}
function pictureFormatTab(d) {
  return { id: 'fmt-pic', label: 'Định dạng Ảnh', en: 'Picture Format', contextual: true, accent: '#c43e1c', groups: [
    { label: 'Điều chỉnh', items: [
      { type: 'drop', big: true, label: 'Hiệu chỉnh', icon: 'sun', menu: () => [['', 'Gốc'], ['brightness(1.2)', 'Sáng hơn'], ['brightness(.8)', 'Tối hơn'], ['contrast(1.4)', 'Tương phản cao']].map(f => ({ label: f[1], run: () => { const o = sel(d); if (o) d.change('Hiệu chỉnh', () => o.filter = f[0]); } })) },
      { type: 'drop', big: true, label: 'Màu', icon: 'palette', menu: () => [['', 'Gốc'], ['grayscale(1)', 'Thang xám'], ['sepia(1)', 'Nâu đỏ'], ['saturate(1.8)', 'Rực rỡ']].map(f => ({ label: f[1], run: () => { const o = sel(d); if (o) d.change('Màu ảnh', () => o.filter = f[0]); } })) },
      { type: 'drop', big: true, label: 'Hiệu ứng Nghệ thuật', icon: 'wand', menu: () => [['', 'Không'], ['blur(2px)', 'Mờ'], ['invert(1)', 'Đảo màu']].map(f => ({ label: f[1], run: () => { const o = sel(d); if (o) d.change('Hiệu ứng', () => o.filter = f[0]); } })) }
    ] },
    { label: 'Kiểu Ảnh', items: [
      { type: 'gallery', width: 220, items: () => [['none', 'Vuông'], ['1px solid #000', 'Viền mảnh'], ['6px solid #fff', 'Viền trắng']].map(b => ({ title: b[1], preview: () => h('div', { style: { width: '34px', height: '24px', background: '#bcd', border: b[0] === 'none' ? '' : b[0] } }), run: () => { const o = sel(d); if (o) d.change('Kiểu ảnh', () => o.border = b[0] === 'none' ? '' : b[0]); } })) },
      { type: 'col', items: [{ type: 'btn', label: 'Văn bản thay thế', icon: 'accessibility', run: () => sel(d) && P.altText(d, sel(d)) }, { type: 'drop', label: 'Cách vừa', icon: 'maximize', menu: () => [['contain', 'Vừa khít'], ['cover', 'Lấp đầy'], ['fill', 'Kéo giãn']].map(f => ({ label: f[1], run: () => { const o = sel(d); if (o) d.change('Vừa ảnh', () => o.fit = f[0]); } })) }] }
    ] },
    { label: 'Sắp xếp', items: [
      { type: 'col', items: [{ type: 'btn', label: 'Đưa lên trước', icon: 'bring-to-front', run: () => sel(d) && P.order(d, sel(d), 1) }, { type: 'btn', label: 'Đưa ra sau', icon: 'send-to-back', run: () => sel(d) && P.order(d, sel(d), -1) }] },
      { type: 'drop', big: true, label: 'Căn chỉnh', icon: 'align-center-horizontal', menu: () => alignMenu(d) }
    ] }
  ] };
}
function chartFormatTab(d) {
  return { id: 'fmt-chart', label: 'Thiết kế Biểu đồ', en: 'Chart Design', contextual: true, accent: '#217346', groups: [
    { label: 'Dữ liệu', items: [{ type: 'big', label: 'Sửa Dữ liệu', icon: 'table', run: () => sel(d) && P.chartDataDialog(d, sel(d)) }, { type: 'big', label: 'Đổi Hàng/Cột', icon: 'arrow-left-right', run: () => toast('Đổi hàng/cột — đang phát triển') }] },
    { label: 'Loại Biểu đồ', items: [{ type: 'gallery', width: 260, items: () => ['column', 'bar', 'line', 'pie', 'area'].map(k => ({ title: k, preview: () => h('div.tr-prev', ic('chart-' + (k === 'bar' ? 'bar' : k === 'line' ? 'line' : k === 'pie' ? 'pie' : 'column'), 20)), run: () => { const o = sel(d); if (o) d.change('Loại biểu đồ', () => o.kind = k); } })) }] },
    { label: 'Bố cục', items: [
      { type: 'toggle', label: 'Chú giải', icon: 'list', get: () => sel(d) && sel(d).legend !== false, set: v => { const o = sel(d); if (o) d.change('Chú giải', () => o.legend = v); } },
      { type: 'btn', label: 'Tiêu đề', icon: 'type', run: () => { const o = sel(d); if (o) promptDlg('Tiêu đề biểu đồ:', o.title || '', 'Tiêu đề').then(t => { if (t != null) d.change('Tiêu đề', () => o.title = t); }); } }
    ] }
  ] };
}
function tableFormatTab(d) {
  return { id: 'fmt-table', label: 'Thiết kế Bảng', en: 'Table Design', contextual: true, accent: '#c43e1c', groups: [
    { label: 'Tùy chọn', items: [{ type: 'toggle', label: 'Hàng tiêu đề', icon: 'table', get: () => sel(d) && sel(d).header, set: v => { const o = sel(d); if (o) d.change('Hàng tiêu đề', () => o.header = v); } }] },
    { label: 'Kiểu Bảng', items: [{ type: 'gallery', width: 260, items: () => d.theme.colors.slice(0, 5).map(c => ({ title: '', preview: () => h('div', { style: { width: '40px', height: '26px', background: c, borderRadius: '3px' } }), run: () => { const o = sel(d); if (o) d.change('Kiểu bảng', () => { o.style = { head: c }; }); } })) }] },
    { label: 'Hàng & Cột', items: [{ type: 'col', items: [{ type: 'btn', label: 'Thêm hàng', icon: 'plus', run: () => tableRowCol(d, 'row', 1) }, { type: 'btn', label: 'Thêm cột', icon: 'plus', run: () => tableRowCol(d, 'col', 1) }] }] }
  ] };
}

// ---------------------------------------------------------------- helpers
function menuPos() { return { x: 320, y: 150 }; }
function shapePrev(shape, color) { const d = h('div.pp-shape-prev'); d.innerHTML = shapeSvg(shape, 36, 26, color, 'none', 0); return d; }
function shapeMenu(d, edit) { return SHAPES.map(s => ({ label: s[1], icon: 'shapes', run: () => { if (edit && sel(d) && sel(d).type === 'shape') d.change('Đổi hình', () => sel(d).shape = s[0]); else P.addShape(d, s[0]); } })); }
function themePrev(t) { const bg = typeof t.bg === 'string' ? t.bg : 'linear-gradient(135deg,' + (t.bg.stops || ['#fff']).join(',') + ')'; const d = h('div.pp-theme-prev', { style: { background: bg } }); d.appendChild(h('div', { style: { color: t.titleColor, fontFamily: t.font, fontWeight: 'bold', fontSize: '12px' } }, 'Aa')); const bar = h('div', { style: { display: 'flex', gap: '2px', marginTop: '3px' } }); t.colors.slice(0, 5).forEach(c => bar.appendChild(h('span', { style: { width: '7px', height: '7px', background: c, borderRadius: '50%' } }))); d.appendChild(bar); return d; }
function applyTheme(d, t) { d.change('Chủ đề', () => d.theme = Object.assign({}, t)); d.sbTheme.textContent = t.name; toast('Đã áp dụng chủ đề ' + t.name); }
function fillSel(d, c) { const o = sel(d); if (!o) return toast('Chọn một hình', 'err'); d.change('Màu tô', () => o.fill = c); }
function strokeSel(d, c) { const o = sel(d); if (!o) return toast('Chọn một hình', 'err'); d.change('Màu viền', () => { o.stroke = c; o.strokeW = o.strokeW || 2; }); }
function effectSel(d, e) { const o = sel(d); if (!o) return; d.change('Hiệu ứng', () => o.effect = e === 'none' ? null : e); }
function alignMenu(d) { return [['left', 'Căn trái'], ['hcenter', 'Căn giữa ngang'], ['right', 'Căn phải'], ['top', 'Căn trên'], ['vcenter', 'Căn giữa dọc'], ['bottom', 'Căn dưới']].map(a => ({ label: a[1], run: () => P.align(d, a[0]) })); }
function arrangeMenu(d) { return [{ label: 'Đưa lên trước', icon: 'bring-to-front', run: () => sel(d) && P.order(d, sel(d), 1) }, { label: 'Đưa ra sau', icon: 'send-to-back', run: () => sel(d) && P.order(d, sel(d), -1) }, { sep: true }].concat(alignMenu(d)); }
function insertChar(d, ch) { const o = sel(d); if (o && (o.type === 'text' || o.type === 'shape')) { const el = d.canvas.querySelector('[data-id="' + o.id + '"] .pp-tf, [data-id="' + o.id + '"] .pp-shape-text'); if (el && el.isContentEditable) { document.execCommand('insertText', false, ch); return; } d.change('Chèn ký hiệu', () => { if (o.type === 'text') o.html = (o.html || '') + ch; else o.text = (o.text || '') + ch; }); } else P.addObj(d, { type: 'text', x: 500, y: 320, w: 200, h: 80, html: ch, fontSize: 40 }); }
function quickTable(d, r, c) { const rows = []; for (let i = 0; i < r; i++) { const row = []; for (let j = 0; j < c; j++) row.push(i === 0 ? 'Cột ' + (j + 1) : ''); rows.push(row); } d.change('Chèn bảng', () => d.slide.objects.push({ id: 'o' + Date.now(), type: 'table', rows, header: true, x: 180, y: 220, w: 920, h: 300, fontSize: 18 })); }
function tableRowCol(d, kind, n) { const o = sel(d); if (!o || o.type !== 'table') return toast('Chọn một bảng', 'err'); d.change('Bảng', () => { if (kind === 'row') o.rows.push(o.rows[0].map(() => '')); else o.rows.forEach(r => r.push('')); }); }
function wordArt(d) { promptDlg('Văn bản WordArt:', 'Tiêu đề', 'WordArt').then(v => { if (v) P.addObj(d, { type: 'text', x: 300, y: 280, w: 680, h: 140, html: '<b>' + v + '</b>', fontSize: 60, color: d.theme.colors[0], align: 'center', wordart: true }); }); }
function headerFooter(d, which) { if (which === 'number') { d.change('Số trang', () => d.slides.forEach((s, i) => { if (!s.objects.some(o => o.ph === 'slideno')) s.objects.push({ id: 'n' + i, type: 'text', ph: 'slideno', x: 1180, y: 680, w: 80, h: 30, html: String(i + 1), fontSize: 12, align: 'right' }); })); toast('Đã thêm số trang'); } else promptDlg('Nội dung chân trang:', '', 'Chân trang').then(v => { if (v == null) return; d.change('Chân trang', () => d.slides.forEach(s => { s.objects.push({ id: 'f' + Date.now() + Math.random(), type: 'text', ph: 'footer', x: 60, y: 680, w: 600, h: 30, html: v, fontSize: 12 }); })); }); }
function linkObj(d) { const o = sel(d); if (!o) return toast('Chọn đối tượng để thêm liên kết', 'err'); promptDlg('Địa chỉ liên kết (URL hoặc "slide:N"):', o.link || '', 'Liên kết').then(v => { if (v != null) d.change('Liên kết', () => o.link = v); }); }
function actionDialog(d) { const o = sel(d); if (!o) return toast('Chọn đối tượng', 'err'); showMenu([{ label: 'Chuyển tới trang tiếp theo', run: () => d.change('Hành động', () => o.action = 'next') }, { label: 'Chuyển tới trang trước', run: () => d.change('Hành động', () => o.action = 'prev') }, { label: 'Trang đầu', run: () => d.change('Hành động', () => o.action = 'first') }, { label: 'Kết thúc trình chiếu', run: () => d.change('Hành động', () => o.action = 'end') }], menuPos()); }
function findDialog(d) { promptDlg('Tìm văn bản trong các trang chiếu:', '', 'Tìm').then(q => { if (!q) return; const hits = []; d.slides.forEach((s, i) => s.objects.forEach(o => { const t = P.objText(o); if (t && t.toLowerCase().indexOf(q.toLowerCase()) >= 0) hits.push(i); })); if (hits.length) { d.goto(hits[0]); toast('Tìm thấy ở ' + hits.length + ' trang (tới trang ' + (hits[0] + 1) + ')'); } else toast('Không tìm thấy', 'err'); }); }
function replaceDialog(d) { formDlg('Tìm và Thay thế', [{ key: 'find', label: 'Tìm' }, { key: 'repl', label: 'Thay bằng' }]).then(v => { if (!v || !v.find) return; let n = 0; d.change('Thay thế', () => d.slides.forEach(s => s.objects.forEach(o => { if (o.type === 'text' && o.html && o.html.indexOf(v.find) >= 0) { o.html = o.html.split(v.find).join(v.repl || ''); n++; } }))); toast('Đã thay thế trong ' + n + ' đối tượng'); }); }
function selectionPane(d) { showMenu(d.slide.objects.map((o, i) => ({ label: (o.ph || o.type) + ' ' + (i + 1), icon: o.type === 'text' ? 'type' : o.type === 'image' ? 'image' : o.type === 'chart' ? 'chart-column' : 'square', run: () => { d.selectedObj = o; d.renderSlide(); d.onSelection(); } })), menuPos()); }
function selectionPane2() {}
function addComment(d) { promptDlg('Chú thích cho trang ' + (d.cur + 1) + ':', '', 'Chú thích mới').then(v => { if (v) { d.slide.comments = d.slide.comments || []; d.slide.comments.push({ text: v, at: Date.now() }); d.markDirty(); toast('Đã thêm chú thích'); } }); }
function commentsPane(d) { const all = []; d.slides.forEach((s, i) => (s.comments || []).forEach(c => all.push('Trang ' + (i + 1) + ': ' + c.text))); dialog({ title: 'Chú thích', width: 480, body: h('div', all.length ? all.map(t => h('p', t)) : [h('p.muted', 'Chưa có chú thích.')]), buttons: [{ label: 'Đóng', primary: true }] }); }
function showSetup(d) { formDlg('Thiết lập Trình chiếu', [{ key: 'loop', type: 'checkbox', label: 'Lặp lại liên tục đến khi nhấn Esc' }, { key: 'noanim', type: 'checkbox', label: 'Trình chiếu không có hoạt ảnh' }, { key: 'pen', type: 'color', label: 'Màu bút', value: '#e81123' }]).then(v => { if (v) { d._showOpts = v; toast('Đã lưu thiết lập trình chiếu'); } }); }
function previewTransition(d) { const el = d.canvas; const t = d.slide.transition; if (!t) return toast('Trang này chưa có chuyển tiếp'); el.classList.remove('tr-anim'); void el.offsetWidth; el.classList.add('tr-anim', 'tr-' + t); setTimeout(() => el.classList.remove('tr-anim', 'tr-' + t), 800); }
function previewAnim(d) { const o = sel(d); if (!o || !o.anim) return toast('Chọn đối tượng có hoạt ảnh'); const el = d.canvas.querySelector('[data-id="' + o.id + '"]'); if (el) { el.classList.remove('anim-' + o.anim); void el.offsetWidth; el.classList.add('anim-' + o.anim); } }
function animPane(d) { const items = d.slide.objects.filter(o => o.anim).map((o, i) => (i + 1) + '. ' + (o.ph || o.type) + ' — ' + (P.ANIMS.find(a => a[0] === o.anim) || ['', o.anim])[1]); dialog({ title: 'Ngăn Hoạt ảnh', width: 420, body: h('div', items.length ? items.map(t => h('p', t)) : [h('p.muted', 'Chưa có hoạt ảnh trên trang này. Chọn đối tượng rồi thêm hoạt ảnh.')]), buttons: [{ label: 'Đóng', primary: true }] }); }
function designerDialog(d) { dialog({ title: 'Nhà thiết kế (A.I)', width: 520, body: h('div', h('p', 'Chọn một bố cục gợi ý cho trang hiện tại:'), h('div.pp-designer', LAYOUTS.slice(0, 6).map(l => { const b = h('button.pp-design-idea', ic(l[2], 20), h('div', l[1])); b.addEventListener('click', () => { P.changeLayout(d, l[0]); document.querySelector('.dlg-back') && document.querySelector('.dlg-back').remove(); toast('Đã áp dụng bố cục ' + l[1]); }); return b; }))), buttons: [{ label: 'Đóng', primary: true }] }); }
function fontDialog(d) { const o = sel(d); if (!o) return toast('Chọn một hộp văn bản', 'err'); formDlg('Phông chữ', [{ key: 'font', label: 'Phông', value: o.font || d.theme.font }, { key: 'size', type: 'number', label: 'Cỡ', value: o.fontSize || 18 }, { key: 'color', type: 'color', label: 'Màu', value: o.color || '#000000' }]).then(v => { if (v) d.change('Phông chữ', () => { o.font = v.font; o.fontSize = parseFloat(v.size) || o.fontSize; o.color = v.color; }); }); }

// v0.0.7: Developer tab — VBA for PowerPoint (Alt+F11 / Alt+F8, .pptm)
function developerTab(d) {
  const V = f => () => import('../excel/vba/index.js').then(m => f(m));
  return { id: 'developer', label: 'Nhà phát triển', en: 'Developer', groups: [
    { label: 'Mã', items: [
      { type: 'big', label: 'Visual Basic', icon: 'square-code', title: 'Mở trình soạn VBA (Alt+F11)', run: V(m => m.openEditor(d)) },
      { type: 'big', label: 'Macro', icon: 'list-video', title: 'Danh sách macro (Alt+F8)', run: V(m => m.macrosDialog(d)) }
    ] },
    { label: 'esbuild (TS / TSX)', items: [
      { type: 'big', label: 'TypeScript / TSX', icon: 'file-code', title: 'Trình biên dịch & Thực thi TypeScript / TSX Macro bằng esbuild (ts vs tsx)', run: () => import('../core/suites.js').then(m => m.openEsbuildEditor(d, 'ts')) }
    ] },
    { label: 'Lưu', items: [{ type: 'big', label: 'Lưu .pptm', icon: 'save', title: 'Lưu bản trình bày kèm macro', run: () => d.shell.save(d, true, 'pptm') }] }
  ] };
}
