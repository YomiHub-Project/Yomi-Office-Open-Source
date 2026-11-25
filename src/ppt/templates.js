// PowerPoint themes, layouts, templates
import { h, uid } from '../core/util.js';
import { SW, SH } from './ppt.js';

export const THEMES = [
  { name: 'Office', font: 'Calibri', bg: '#ffffff', titleColor: '#44546a', bodyColor: '#404040', colors: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'], accentBar: false },
  { name: 'Liquid Glass', font: 'Segoe UI', bg: { type: 'gradient', angle: 135, stops: ['#eef4ff', '#e3ecff'] }, titleColor: '#0f3d6e', bodyColor: '#333', colors: ['#0F6CBD', '#8764B8', '#00B7C3', '#E3008C', '#FFB900', '#107C10'], accentBar: true },
  { name: 'VISICONS', font: 'Times New Roman', bg: '#ffffff', titleColor: '#1f3864', bodyColor: '#222', colors: ['#1F3864', '#C00000', '#2E75B5', '#BF8F00', '#548235', '#7030A0'], accentBar: true },
  { name: 'Tối giản', font: 'Arial', bg: '#ffffff', titleColor: '#111', bodyColor: '#333', colors: ['#111111', '#e63946', '#457b9d', '#a8dadc', '#f1faee', '#1d3557'], accentBar: false },
  { name: 'Đậm chất', font: 'Century Gothic', bg: { type: 'gradient', angle: 135, stops: ['#1a2a6c', '#2a4858'] }, titleColor: '#ffffff', bodyColor: '#e0e0e0', colors: ['#fdbb2d', '#22c1c3', '#ff6a88', '#ff99ac', '#a1c4fd', '#c2e9fb'], accentBar: false },
  { name: 'Tự nhiên', font: 'Georgia', bg: '#f8f5ee', titleColor: '#33691e', bodyColor: '#3e2723', colors: ['#558B2F', '#8D6E63', '#F9A825', '#4E342E', '#AED581', '#FF8A65'], accentBar: true },
  { name: 'Hoàng hôn', font: 'Segoe UI', bg: { type: 'gradient', angle: 160, stops: ['#ff9a9e', '#fecfef'] }, titleColor: '#6a1b4d', bodyColor: '#4a2c3a', colors: ['#e91e63', '#ff6f00', '#ffc107', '#9c27b0', '#673ab7', '#f06292'], accentBar: false },
  { name: 'Công nghệ', font: 'Consolas', bg: '#0d1117', titleColor: '#58a6ff', bodyColor: '#c9d1d9', colors: ['#58a6ff', '#3fb950', '#f85149', '#d29922', '#a371f7', '#39c5cf'], accentBar: true }
];
export const LAYOUTS = [
  ['title', 'Trang bìa', 'layout-template'], ['content', 'Tiêu đề và Nội dung', 'layout-list'], ['section', 'Đầu mục', 'heading'], ['two', 'Hai nội dung', 'columns-2'], ['compare', 'So sánh', 'arrow-left-right'], ['titleonly', 'Chỉ tiêu đề', 'heading-1'], ['blank', 'Trống', 'square'], ['caption', 'Nội dung có chú thích', 'captions'], ['picture', 'Ảnh có chú thích', 'image']
];
export function layoutObjects(layout, theme, existing) {
  existing = existing || {};
  const txt = (ph, x, y, w, hh, fs, extra) => Object.assign({ id: uid('o'), type: 'text', ph, x, y, w, h: hh, fontSize: fs, html: existing[ph] ? existing[ph].html : '' }, extra || {});
  switch (layout) {
    case 'title': return [txt('title', 120, 260, SW - 240, 160, 54, { align: 'center', valign: 'middle', b: true }), txt('subtitle', 160, 440, SW - 320, 80, 26, { align: 'center' })];
    case 'section': return [txt('title', 120, 300, SW - 240, 120, 44, { align: 'left', b: true }), txt('subtitle', 120, 430, SW - 240, 60, 22)];
    case 'content': return [txt('title', 60, 40, SW - 120, 90, 36, { b: true }), txt('body', 60, 150, SW - 120, SH - 200, 22, { valign: 'top' })];
    case 'titleonly': return [txt('title', 60, 40, SW - 120, 90, 36, { b: true })];
    case 'two': return [txt('title', 60, 40, SW - 120, 90, 36, { b: true }), txt('body', 60, 150, (SW - 160) / 2, SH - 200, 20), txt('body2', SW / 2 + 20, 150, (SW - 160) / 2, SH - 200, 20)];
    case 'compare': return [txt('title', 60, 40, SW - 120, 80, 32, { b: true }), txt('h1', 60, 130, (SW - 160) / 2, 50, 22, { b: true, color: theme.colors[0] }), txt('body', 60, 190, (SW - 160) / 2, SH - 240, 20), txt('h2', SW / 2 + 20, 130, (SW - 160) / 2, 50, 22, { b: true, color: theme.colors[1] }), txt('body2', SW / 2 + 20, 190, (SW - 160) / 2, SH - 240, 20)];
    case 'caption': return [txt('title', 60, 40, SW / 2 - 90, 80, 30, { b: true }), txt('body', 60, 140, SW / 2 - 90, SH - 200, 18), { id: uid('o'), type: 'shape', shape: 'rect', x: SW / 2 + 20, y: 60, w: SW / 2 - 80, h: SH - 120, fill: theme.colors[2], stroke: 'none' }];
    case 'picture': return [txt('title', 60, SH - 140, SW - 120, 60, 28, { b: true }), txt('body', 60, SH - 80, SW - 120, 50, 18), { id: uid('o'), type: 'shape', shape: 'rect', x: 60, y: 40, w: SW - 120, h: SH - 200, fill: '#e8e8e8', stroke: 'none' }];
    case 'blank': default: return [];
  }
}
export function blankSlide(layout, theme) { theme = theme || THEMES[0]; return { id: uid('s'), layout: layout || 'content', objects: layoutObjects(layout || 'content', theme), notes: '', transition: null, bg: null }; }
function th(theme) { return () => { const d = h('div.th-slide', { style: { background: typeof theme.bg === 'string' ? theme.bg : 'linear-gradient(135deg,' + (theme.bg.stops || ['#fff', '#fff']).join(',') + ')' } }); d.appendChild(h('div.t', { style: { color: theme.titleColor, fontFamily: theme.font } }, 'Tiêu đề')); d.appendChild(h('div.s', { style: { color: theme.bodyColor } }, 'Phụ đề')); const bar = h('div', { style: { display: 'flex', marginTop: '6px' } }); theme.colors.slice(0, 5).forEach(c => bar.appendChild(h('span', { style: { width: '8px', height: '8px', background: c, margin: '0 1px', borderRadius: '50%' } }))); d.appendChild(bar); return d; }; }
export function templates() {
  const custom = [
    { id: 'annual_report', name: 'Báo cáo tổng kết kinh doanh năm', theme: THEMES[2], thumb: th(THEMES[2]) },
    { id: 'company_profile', name: 'Hồ sơ năng lực công ty (Profile)', theme: THEMES[1], thumb: th(THEMES[1]) },
    { id: 'pitch_deck', name: 'Pitch Deck gọi vốn khởi nghiệp', theme: THEMES[4], thumb: th(THEMES[4]) },
    { id: 'thesis', name: 'Bản trình bày học thuật / Luận văn', theme: THEMES[5], thumb: th(THEMES[5]) }
  ];
  const list = THEMES.map((t, i) => ({ id: i === 0 ? 'blank' : 'theme' + i, name: i === 0 ? 'Bản trình bày trống' : t.name, theme: t, thumb: th(t) }));
  return [list[0]].concat(custom).concat(list.slice(1));
}
export async function applyTemplate(doc, t) {
  doc.theme = t.theme || THEMES[0]; doc.name = t.name;
  if (t.id === 'annual_report') {
    doc.slides = [
      blankSlide('title', doc.theme),
      blankSlide('content', doc.theme),
      blankSlide('two', doc.theme),
      blankSlide('section', doc.theme)
    ];
    let o = doc.slides[0].objects.find(x => x.ph === 'title'); if (o) o.html = 'BÁO CÁO TỔNG KẾT KINH DOANH NĂM';
    o = doc.slides[0].objects.find(x => x.ph === 'subtitle'); if (o) o.html = 'TẬP ĐOÀN VISICONS · NĂM TÀI CHÍNH 2026';
    o = doc.slides[1].objects.find(x => x.ph === 'title'); if (o) o.html = '1. Kết quả thực hiện các chỉ tiêu cốt lõi';
    o = doc.slides[1].objects.find(x => x.ph === 'body'); if (o) o.html = '<div>• Doanh thu thuần đạt 148 tỷ đồng (vượt 18% kế hoạch đề ra)</div><div>• Lợi nhuận gộp đạt 45 tỷ đồng, biên lợi nhuận duy trì mức 30.4%</div><div>• Hoàn thành bàn giao 03 gói thầu trọng điểm đúng tiến độ cam kết</div>';
    o = doc.slides[2].objects.find(x => x.ph === 'title'); if (o) o.html = '2. So sánh hiệu quả theo từng khối hoạt động';
    o = doc.slides[2].objects.find(x => x.ph === 'body'); if (o) o.html = '<div><b>Khối Thi công xây lắp:</b></div><div>• Doanh số: 110 tỷ</div><div>• Tỷ trọng: 74%</div><div>• Đạt chuẩn an toàn 100%</div>';
    o = doc.slides[2].objects.find(x => x.ph === 'body2'); if (o) o.html = '<div><b>Khối Thiết kế & Tư vấn:</b></div><div>• Doanh số: 38 tỷ</div><div>• Tỷ trọng: 26%</div><div>• Triển khai giải pháp số hóa toàn diện</div>';
  } else if (t.id === 'company_profile') {
    doc.slides = [
      blankSlide('title', doc.theme),
      blankSlide('content', doc.theme),
      blankSlide('two', doc.theme)
    ];
    let o = doc.slides[0].objects.find(x => x.ph === 'title'); if (o) o.html = 'HỒ SƠ NĂNG LỰC DOANH NGHIỆP';
    o = doc.slides[0].objects.find(x => x.ph === 'subtitle'); if (o) o.html = 'VISICONS - GIẢI PHÁP THIẾT KẾ & XÂY DỰNG CHẤT LƯỢNG CAO';
    o = doc.slides[1].objects.find(x => x.ph === 'title'); if (o) o.html = 'Tầm nhìn & Sứ mệnh phát triển';
    o = doc.slides[1].objects.find(x => x.ph === 'body'); if (o) o.html = '<div>• <b>Tầm nhìn:</b> Trở thành tổng thầu hàng đầu trong lĩnh vực xây dựng dân dụng và công nghiệp.</div><div>• <b>Sứ mệnh:</b> Kiến tạo những công trình bền vững, đạt chuẩn chất lượng quốc tế và an toàn tuyệt đối.</div>';
    o = doc.slides[2].objects.find(x => x.ph === 'title'); if (o) o.html = 'Năng lực nhân sự & Thiết bị thi công';
    o = doc.slides[2].objects.find(x => x.ph === 'body'); if (o) o.html = '<div><b>Đội ngũ chuyên gia:</b></div><div>• Hơn 150 kỹ sư trình độ cao</div><div>• 100% có chứng chỉ hành nghề chuẩn</div>';
    o = doc.slides[2].objects.find(x => x.ph === 'body2'); if (o) o.html = '<div><b>Hệ thống thiết bị:</b></div><div>• Dàn máy móc cơ giới hiện đại</div><div>• Ứng dụng công nghệ phần mềm Yomi-Office Open Source</div>';
  } else if (t.id === 'pitch_deck') {
    doc.slides = [
      blankSlide('title', doc.theme),
      blankSlide('content', doc.theme),
      blankSlide('two', doc.theme)
    ];
    let o = doc.slides[0].objects.find(x => x.ph === 'title'); if (o) o.html = 'OFFICE OPEN SOURCE - PITCH DECK';
    o = doc.slides[0].objects.find(x => x.ph === 'subtitle'); if (o) o.html = 'Bộ giải pháp văn phòng mã nguồn mở hoạt động ngoại tuyến và đa nền tảng';
    o = doc.slides[1].objects.find(x => x.ph === 'title'); if (o) o.html = 'Vấn đề thị trường & Giải pháp';
    o = doc.slides[1].objects.find(x => x.ph === 'body'); if (o) o.html = '<div>• <b>Vấn đề:</b> Chi phí bản quyền phần mềm văn phòng định kỳ đắt đỏ; nguy cơ lộ dữ liệu khi bắt buộc kết nối máy chủ ngoại cảnh.</div><div>• <b>Giải pháp:</b> Yomi-Office Open Source chạy offline 100%, bảo mật tuyệt đối, đầy đủ Word, Excel, PowerPoint, PDF và A.I.</div>';
  } else {
    doc.slides = [blankSlide('title', doc.theme), blankSlide('content', doc.theme)];
    const title = doc.slides[0].objects.find(o => o.ph === 'title'); if (title) title.html = doc.theme.name === t.name ? 'Tiêu đề bản trình bày' : t.name;
    const sub = doc.slides[0].objects.find(o => o.ph === 'subtitle'); if (sub) sub.html = 'Phụ đề hoặc tên của bạn';
    const t2 = doc.slides[1].objects.find(o => o.ph === 'title'); if (t2) t2.html = 'Nội dung chính';
    const b2 = doc.slides[1].objects.find(o => o.ph === 'body'); if (b2) b2.html = '<div>• Điểm thứ nhất</div><div>• Điểm thứ hai</div><div>• Điểm thứ ba</div>';
  }
}
