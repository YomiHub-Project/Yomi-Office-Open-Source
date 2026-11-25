// Help tab (shared): Help, Contact support, Feedback, Training, What's new, Keyboard shortcuts
import { h, ic, VERSION, APP_NAME } from './util.js';
import { dialog } from './ui.js';
import Host from './host.js';

const SHORTCUTS = {
  common: [['Ctrl+N', 'Tạo mới'], ['Ctrl+O', 'Mở'], ['Ctrl+S', 'Lưu'], ['F12', 'Lưu Như'], ['Ctrl+P', 'In'], ['Ctrl+Z / Ctrl+Y', 'Hoàn tác / Làm lại'], ['Ctrl+C / X / V', 'Sao chép / Cắt / Dán'], ['Ctrl+F / Ctrl+H', 'Tìm / Thay thế'], ['Alt+Q', 'Tìm kiếm lệnh'], ['Ctrl+F1', 'Thu gọn Ribbon'], ['Ctrl + lăn chuột', 'Phóng to/thu nhỏ'], ['F11', 'Toàn màn hình'], ['Chuột phải', 'Menu ngữ cảnh']],
  excel: [['F2', 'Sửa ô'], ['Alt+Enter', 'Xuống dòng trong ô'], ['Ctrl+Enter', 'Điền cùng giá trị cho vùng chọn'], ['F4', 'Đổi tham chiếu tuyệt đối/tương đối'], ['Ctrl+Mũi tên', 'Nhảy tới biên dữ liệu'], ['Shift+Mũi tên', 'Mở rộng vùng chọn'], ['Ctrl+Space / Shift+Space', 'Chọn cột / hàng'], ['Ctrl+1', 'Định dạng ô'], ['Alt+=', 'Tự động tính tổng'], ['Ctrl+D / Ctrl+R', 'Điền xuống / sang phải'], ['Ctrl+E', 'Điền nhanh'], ['Ctrl+;', 'Ngày hiện tại'], ['Ctrl+Shift+L', 'Bật/tắt bộ lọc'], ['Ctrl+T', 'Tạo bảng'], ['Ctrl+`', 'Hiện công thức'], ['F9', 'Tính lại'], ['Shift+F3', 'Chèn hàm'], ['Shift+F11', 'Trang tính mới'], ['Ctrl+PgUp/PgDn', 'Chuyển trang tính'], ['Bấm đúp viền tiêu đề cột', 'Tự khớp độ rộng'], ['Kéo ô vuông góc dưới phải', 'Tự động điền (Fill Handle)'], ['Kéo viền vùng chọn', 'Di chuyển ô (giữ Ctrl để sao chép)']],
  word: [['Ctrl+B / I / U', 'Đậm / Nghiêng / Gạch chân'], ['Ctrl+L / E / R / J', 'Căn trái / giữa / phải / đều'], ['Ctrl+1 / 2 / 5', 'Giãn dòng 1 / 2 / 1,5'], ['Ctrl+Alt+1..3', 'Đề mục 1..3'], ['Ctrl+Shift+N', 'Kiểu Bình thường'], ['Ctrl+] / [', 'Tăng / giảm cỡ chữ'], ['Ctrl+K', 'Chèn liên kết'], ['Ctrl+Enter', 'Ngắt trang'], ['Ctrl+Shift+C / V', 'Sao chép / dán định dạng'], ['Ctrl+Shift+8', 'Hiện ký hiệu định dạng'], ['Ctrl+Alt+M', 'Chú thích mới'], ['Ctrl+Shift+E', 'Theo dõi thay đổi'], ['F7', 'Chính tả'], ['Bấm 3 lần', 'Chọn đoạn văn']],
  ppt: [['F5', 'Trình chiếu từ đầu'], ['Shift+F5', 'Trình chiếu từ trang hiện tại'], ['Ctrl+M', 'Trang chiếu mới'], ['Ctrl+D', 'Nhân bản đối tượng/trang'], ['Delete', 'Xóa đối tượng/trang'], ['Mũi tên', 'Dịch chuyển đối tượng'], ['Shift + kéo góc', 'Giữ tỉ lệ khi đổi kích thước'], ['B / W', 'Màn hình đen / trắng khi trình chiếu'], ['Esc', 'Kết thúc trình chiếu'], ['Ctrl+G', 'Nhóm đối tượng']],
  pdf: [['Ctrl+F', 'Tìm trong PDF'], ['+ / -', 'Phóng to / thu nhỏ'], ['PgUp / PgDn', 'Trang trước / sau']]
};
export function shortcutsDialog(kind) {
  const list = SHORTCUTS.common.concat(SHORTCUTS[kind] || []);
  return dialog({ title: 'Phím tắt & thao tác chuột', width: 620, body: h('table.kv', list.map(x => h('tr', h('td', h('b', x[0])), h('td', x[1])))), buttons: [{ label: 'Đóng', primary: true }] });
}
export function whatsNew() {
  return dialog({ title: 'Có gì mới trong ' + APP_NAME + ' ' + VERSION, width: 640, body: h('div.whatsnew',
    h('p', 'Phiên bản 0.0.6:'),
    h('ul', [
      'Excel: gõ trực tiếp vào ô đang chọn (không cần bấm đúp); VBA — trình soạn Visual Basic (Alt+F11), hộp thoại Macro (Alt+F8), ghi macro, chạy macro ngoại tuyến; đọc/ghi .xlsm giữ nguyên macro với Microsoft Excel.',
      'Excel: đủ 525 hàm của Excel 365 (thêm CHISQ.DIST/INV, CONFIDENCE.NORM/T, BINOM.DIST.RANGE, FORECAST.ETS…, PIVOTBY, hàm rút gọn SUM/AVERAGE trong GROUPBY/MAP); sửa T.INV.2T.',
      'Word: Thước ngang + thước dọc mới, căn đúng theo trang, kéo thụt lề / lề trang trực tiếp; Bấm-và-Gõ; Chuyển mã TCVN3/VNI → Unicode.',
      'PowerPoint: sửa lỗi không gõ được chữ — bấm vào khung là gõ ngay; tab Ghi (thuyết minh, Cameo, quay màn hình, xuất video).',
      'Mỗi tệp mở trong một cửa sổ riêng (Windows: cửa sổ mới; Android: tác vụ riêng trong Gần đây).',
      'Định dạng mới: OpenDocument (.odt .ods .odp), Office 97-2003 (.doc .xls .ppt — đọc), RTF đầy đủ, Word/PowerPoint XML, MHT, XML Spreadsheet 2003, SYLK, DIF, .dotx .xltx .potx .ppsx, xuất video.',
      'Phông chữ: tích hợp phông tiếng Việt (HP001 Tiểu học, Wisteria Notes), Quản lý phông chữ, nạp Font Pack, dùng phông của máy.',
      'Chuột: cuộn tự động bằng nút giữa, vệt sáng khi bấm, tùy chọn hướng di chuyển sau Enter, lăn để zoom (Excel), chế độ cảm ứng.'
    ].map(t => h('li', t)))), buttons: [{ label: 'Đóng', primary: true }] });
}
export function helpTab(app, kind) {
  return { id: 'help', label: 'Trợ giúp', en: 'Help', groups: [
    { label: 'Trợ giúp', items: [
      { type: 'big', label: 'Trợ giúp', icon: 'circle-help', sc: 'F1', run: () => shortcutsDialog(kind) },
      { type: 'big', label: 'Liên hệ Hỗ trợ', icon: 'headset', run: () => dialog({ title: 'Liên hệ hỗ trợ', width: 480, body: h('div', h('p', APP_NAME + ' ' + VERSION + ' là dự án mã nguồn mở.'), h('p', 'Báo lỗi hoặc góp ý: mở mục Phản hồi, hoặc gửi thông tin bên dưới cho nhóm phát triển.'), h('pre.muted', JSON.stringify(Host.info(), null, 1))), buttons: [{ label: 'Đóng', primary: true }] }) },
      { type: 'big', label: 'Phản hồi', icon: 'message-circle-heart', run: () => Host.openUrl('https://github.com/') },
      { type: 'big', label: 'Hiện Đào tạo', icon: 'graduation-cap', run: () => shortcutsDialog(kind) },
      { type: 'big', label: 'Có gì Mới', icon: 'gift', run: () => whatsNew() }
    ] }
  ] };
}
