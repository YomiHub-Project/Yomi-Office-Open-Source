// Basic shape library (SVG) shared by Word, Excel, PowerPoint
export const SHAPES = [
  ['rect', 'Hình chữ nhật'], ['roundRect', 'Chữ nhật bo góc'], ['ellipse', 'Hình elip'], ['triangle', 'Tam giác'], ['rtTriangle', 'Tam giác vuông'], ['diamond', 'Hình thoi'],
  ['parallelogram', 'Hình bình hành'], ['trapezoid', 'Hình thang'], ['pentagon', 'Ngũ giác'], ['hexagon', 'Lục giác'], ['octagon', 'Bát giác'], ['plus', 'Dấu cộng'],
  ['star5', 'Ngôi sao 5 cánh'], ['star4', 'Ngôi sao 4 cánh'], ['heart', 'Trái tim'], ['cloud', 'Đám mây'], ['lightning', 'Tia sét'], ['moon', 'Mặt trăng'],
  ['rightArrow', 'Mũi tên phải'], ['leftArrow', 'Mũi tên trái'], ['upArrow', 'Mũi tên lên'], ['downArrow', 'Mũi tên xuống'], ['leftRightArrow', 'Mũi tên hai chiều'], ['chevron', 'Chữ V (Chevron)'],
  ['line', 'Đường thẳng'], ['arrowLine', 'Đường mũi tên'], ['callout', 'Chú thích hình chữ nhật'], ['wedgeEllipse', 'Chú thích hình elip'], ['frame', 'Khung'], ['donut', 'Vành khuyên'],
  ['flowProcess', 'Lưu đồ: Xử lý'], ['flowDecision', 'Lưu đồ: Quyết định'], ['flowTerminator', 'Lưu đồ: Bắt đầu/Kết thúc'], ['flowData', 'Lưu đồ: Dữ liệu'], ['flowDocument', 'Lưu đồ: Tài liệu'], ['smiley', 'Mặt cười']
];
function poly(pts, w, h) { return '<polygon points="' + pts.map(p => (p[0] * w).toFixed(1) + ',' + (p[1] * h).toFixed(1)).join(' ') + '"/>'; }
function star(n, inner, w, h) { const pts = []; for (let i = 0; i < n * 2; i++) { const r = i % 2 ? inner : 0.5; const a = -Math.PI / 2 + i * Math.PI / n; pts.push([0.5 + r * Math.cos(a), 0.5 + r * Math.sin(a)]); } return poly(pts, w, h); }
function reg(n, w, h, rot) { const pts = []; for (let i = 0; i < n; i++) { const a = (rot || -Math.PI / 2) + i * 2 * Math.PI / n; pts.push([0.5 + 0.5 * Math.cos(a), 0.5 + 0.5 * Math.sin(a)]); } return poly(pts, w, h); }
export function shapeBody(shape, w, h) {
  switch (shape) {
    case 'rect': case 'flowProcess': return '<rect x="0" y="0" width="' + w + '" height="' + h + '"/>';
    case 'roundRect': return '<rect x="0" y="0" width="' + w + '" height="' + h + '" rx="' + Math.min(w, h) * 0.18 + '"/>';
    case 'flowTerminator': return '<rect x="0" y="0" width="' + w + '" height="' + h + '" rx="' + h / 2 + '"/>';
    case 'ellipse': return '<ellipse cx="' + w / 2 + '" cy="' + h / 2 + '" rx="' + w / 2 + '" ry="' + h / 2 + '"/>';
    case 'donut': return '<path fill-rule="evenodd" d="M0,' + h / 2 + 'a' + w / 2 + ',' + h / 2 + ' 0 1,0 ' + w + ',0a' + w / 2 + ',' + h / 2 + ' 0 1,0 ' + (-w) + ',0ZM' + w * 0.25 + ',' + h / 2 + 'a' + w / 4 + ',' + h / 4 + ' 0 1,0 ' + w / 2 + ',0a' + w / 4 + ',' + h / 4 + ' 0 1,0 ' + (-w / 2) + ',0Z"/>';
    case 'triangle': return poly([[0.5, 0], [1, 1], [0, 1]], w, h);
    case 'rtTriangle': return poly([[0, 0], [1, 1], [0, 1]], w, h);
    case 'diamond': case 'flowDecision': return poly([[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]], w, h);
    case 'parallelogram': case 'flowData': return poly([[0.22, 0], [1, 0], [0.78, 1], [0, 1]], w, h);
    case 'trapezoid': return poly([[0.2, 0], [0.8, 0], [1, 1], [0, 1]], w, h);
    case 'pentagon': return reg(5, w, h); case 'hexagon': return poly([[0.25, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0.25, 1], [0, 0.5]], w, h); case 'octagon': return reg(8, w, h, -Math.PI / 8);
    case 'plus': return poly([[0.35, 0], [0.65, 0], [0.65, 0.35], [1, 0.35], [1, 0.65], [0.65, 0.65], [0.65, 1], [0.35, 1], [0.35, 0.65], [0, 0.65], [0, 0.35], [0.35, 0.35]], w, h);
    case 'star5': return star(5, 0.2, w, h); case 'star4': return star(4, 0.16, w, h);
    case 'heart': return '<path d="M' + w / 2 + ',' + h * 0.95 + 'C' + w * -0.15 + ',' + h * 0.55 + ' ' + w * 0.1 + ',' + h * -0.12 + ' ' + w / 2 + ',' + h * 0.22 + 'C' + w * 0.9 + ',' + h * -0.12 + ' ' + w * 1.15 + ',' + h * 0.55 + ' ' + w / 2 + ',' + h * 0.95 + 'Z"/>';
    case 'cloud': return '<path d="M' + w * 0.25 + ',' + h * 0.85 + 'a' + w * 0.2 + ',' + h * 0.22 + ' 0 0,1 ' + w * 0.02 + ',' + (-h * 0.42) + 'a' + w * 0.22 + ',' + h * 0.25 + ' 0 0,1 ' + w * 0.38 + ',' + (-h * 0.1) + 'a' + w * 0.18 + ',' + h * 0.2 + ' 0 0,1 ' + w * 0.28 + ',' + h * 0.22 + 'a' + w * 0.15 + ',' + h * 0.2 + ' 0 0,1 ' + (-w * 0.05) + ',' + h * 0.3 + 'Z"/>';
    case 'lightning': return poly([[0.4, 0], [0.75, 0.38], [0.58, 0.45], [0.9, 1], [0.3, 0.55], [0.48, 0.48], [0.15, 0.12]], w, h);
    case 'moon': return '<path d="M' + w * 0.75 + ',0A' + w / 2 + ',' + h / 2 + ' 0 1,0 ' + w * 0.75 + ',' + h + 'A' + w * 0.38 + ',' + h * 0.5 + ' 0 1,1 ' + w * 0.75 + ',0Z"/>';
    case 'rightArrow': return poly([[0, 0.3], [0.6, 0.3], [0.6, 0], [1, 0.5], [0.6, 1], [0.6, 0.7], [0, 0.7]], w, h);
    case 'leftArrow': return poly([[1, 0.3], [0.4, 0.3], [0.4, 0], [0, 0.5], [0.4, 1], [0.4, 0.7], [1, 0.7]], w, h);
    case 'upArrow': return poly([[0.3, 1], [0.3, 0.4], [0, 0.4], [0.5, 0], [1, 0.4], [0.7, 0.4], [0.7, 1]], w, h);
    case 'downArrow': return poly([[0.3, 0], [0.3, 0.6], [0, 0.6], [0.5, 1], [1, 0.6], [0.7, 0.6], [0.7, 0]], w, h);
    case 'leftRightArrow': return poly([[0, 0.5], [0.25, 0], [0.25, 0.3], [0.75, 0.3], [0.75, 0], [1, 0.5], [0.75, 1], [0.75, 0.7], [0.25, 0.7], [0.25, 1]], w, h);
    case 'chevron': return poly([[0, 0], [0.7, 0], [1, 0.5], [0.7, 1], [0, 1], [0.3, 0.5]], w, h);
    case 'line': return '<line x1="0" y1="' + h / 2 + '" x2="' + w + '" y2="' + h / 2 + '" stroke-width="3"/>';
    case 'arrowLine': return '<line x1="0" y1="' + h / 2 + '" x2="' + (w - 10) + '" y2="' + h / 2 + '" stroke-width="3"/>' + poly([[1 - 14 / w, 0.5 - 8 / h], [1, 0.5], [1 - 14 / w, 0.5 + 8 / h]], w, h);
    case 'callout': return poly([[0, 0], [1, 0], [1, 0.75], [0.45, 0.75], [0.25, 1], [0.28, 0.75], [0, 0.75]], w, h);
    case 'wedgeEllipse': return '<ellipse cx="' + w / 2 + '" cy="' + h * 0.42 + '" rx="' + w / 2 + '" ry="' + h * 0.42 + '"/>' + poly([[0.3, 0.7], [0.18, 1], [0.45, 0.8]], w, h);
    case 'frame': return '<path fill-rule="evenodd" d="M0,0H' + w + 'V' + h + 'H0ZM' + w * 0.1 + ',' + h * 0.1 + 'V' + h * 0.9 + 'H' + w * 0.9 + 'V' + h * 0.1 + 'Z"/>';
    case 'flowDocument': return '<path d="M0,0H' + w + 'V' + h * 0.82 + 'C' + w * 0.75 + ',' + h * 0.68 + ' ' + w * 0.25 + ',' + h * 1.05 + ' 0,' + h * 0.85 + 'Z"/>';
    case 'smiley': return '<ellipse cx="' + w / 2 + '" cy="' + h / 2 + '" rx="' + w / 2 + '" ry="' + h / 2 + '"/><circle cx="' + w * 0.35 + '" cy="' + h * 0.38 + '" r="' + Math.min(w, h) * 0.05 + '" fill="#000"/><circle cx="' + w * 0.65 + '" cy="' + h * 0.38 + '" r="' + Math.min(w, h) * 0.05 + '" fill="#000"/><path d="M' + w * 0.3 + ',' + h * 0.62 + 'Q' + w / 2 + ',' + h * 0.82 + ' ' + w * 0.7 + ',' + h * 0.62 + '" fill="none" stroke="#000"/>';
  }
  return '<rect x="0" y="0" width="' + w + '" height="' + h + '"/>';
}
export function shapeSvg(shape, w, h, fill, stroke, sw) {
  const pad = (sw || 1.5);
  return '<svg xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + h + '" viewBox="' + (-pad) + ' ' + (-pad) + ' ' + (w + pad * 2) + ' ' + (h + pad * 2) + '" preserveAspectRatio="none" style="display:block;overflow:visible"><g fill="' + (fill === 'transparent' ? 'none' : fill) + '" stroke="' + (stroke || 'none') + '" stroke-width="' + (sw || 1.5) + '">' + shapeBody(shape, w, h) + '</g></svg>';
}
export function shapeIcon(shape) { return shapeSvg(shape, 22, 16, '#dfe7f5', '#2f528f', 1); }
