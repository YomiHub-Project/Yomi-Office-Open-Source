// Excel templates (Start screen "New" and Template suggestions)
import { h } from '../core/util.js';
import { internStyle, key } from './model.js';
import { parseInput } from './numfmt.js';

function thumb(rows, head) {
  return () => { const t = h('div.th-sheet'); const tb = h('table'); rows.forEach((r, i) => tb.appendChild(h('tr', r.map(c => h('td', { style: i === 0 && head ? { background: head, color: '#fff', fontWeight: 'bold' } : null }, c))))); t.appendChild(tb); return t; };
}

export function templates() {
  return [
    { id: 'blank', name: 'Sổ làm việc trống', thumb: thumb([['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', '']]) },
    { id: 'payroll', name: 'Bảng lương & Chấm công', thumb: thumb([['Họ tên', 'Lương CB', 'Phụ cấp', 'Thực lĩnh'], ['Nguyễn A', '15,000,000', '2,000,000', '17,000,000'], ['Trần B', '12,000,000', '1,500,000', '13,500,000']], '#107c41') },
    { id: 'pnl', name: 'Báo cáo P&L & Lợi nhuận', thumb: thumb([['Chỉ tiêu', 'Năm 2025', 'Năm 2026'], ['Doanh thu', '120 tỷ', '150 tỷ'], ['Lợi nhuận gộp', '35 tỷ', '45 tỷ']], '#1f3864') },
    { id: 'gantt', name: 'Tiến độ dự án (Gantt Chart)', thumb: thumb([['Công việc', 'Bắt đầu', 'Thời hạn', 'Tiến độ'], ['Khảo sát', '01/03', '10 ngày', '100%'], ['Thi công', '11/03', '60 ngày', '45%']], '#0f6cbd') },
    { id: 'stock_pro', name: 'Quản lý Xuất Nhập Tồn kho', thumb: thumb([['Mã VT', 'Tồn đầu', 'Nhập', 'Xuất', 'Tồn cuối'], ['VT01', '100', '50', '30', '120'], ['VT02', '200', '80', '120', '160']], '#c43e1c') },
    { id: 'kpi', name: 'Bảng theo dõi KPI nhân sự', thumb: thumb([['Nhân sự', 'Mục tiêu', 'Thực hiện', 'Tỷ lệ'], ['Nguyễn A', '100 tr', '115 tr', '115%'], ['Trần B', '80 tr', '76 tr', '95%']], '#8764b8') },
    { id: 'welcome', name: 'Chào mừng đến với Excel', thumb: thumb([['Tháng', 'Doanh thu'], ['T1', '12,000'], ['T2', '15,500'], ['Tổng', '27,500']], '#107c41') },
    { id: 'formulas', name: 'Hướng dẫn công thức', thumb: thumb([['Hàm', 'Kết quả'], ['SUM', '270'], ['IF', 'Đạt'], ['VLOOKUP', '65']], '#0f6cbd') },
    { id: 'budget', name: 'Ngân sách sự kiện', thumb: thumb([['Hạng mục', 'Dự toán', 'Thực tế'], ['Địa điểm', '20 tr', '18 tr'], ['Ăn uống', '15 tr', '16 tr']], '#8764b8') },
    { id: 'invoice', name: 'Hóa đơn bán hàng', thumb: thumb([['Mặt hàng', 'SL', 'Thành tiền'], ['Giấy A4', '10', '650,000'], ['', 'Tổng', '715,000']], '#d83b01') }
  ];
}

function put(s, r, c, v, st) {
  const cell = {};
  if (typeof v === 'string' && v[0] === '=') cell.f = v.slice(1);
  else if (typeof v === 'string') { const p = parseInput(v); cell.v = p.v; if (p.nf && !(st && st.nf)) st = Object.assign({}, st || {}, { nf: p.nf }); }
  else if (v != null) cell.v = v;
  if (st) cell.s = internStyle(st);
  s.cells.set(key(r, c), cell);
}
const B = { s: 'thin', c: '#7f7f7f' }; const BD = { t: B, b: B, l: B, r: B };
function table(s, r0, head, rows, fill, widths, nfs) {
  head.forEach((v, j) => put(s, r0, j, v, { b: true, fill, color: '#fff', ha: 'center', bd: BD }));
  rows.forEach((row, i) => row.forEach((v, j) => put(s, r0 + 1 + i, j, v, { bd: BD, nf: nfs && nfs[j] })));
  (widths || []).forEach((w, i) => { s.colW[i] = w; });
}

export async function applyTemplate(doc, t) {
  const s = doc.wb.sheets[0]; doc.name = t.name;
  if (t.id === 'payroll') {
    s.name = 'Bảng lương';
    put(s, 0, 0, 'BẢNG THANH TOÁN TIỀN LƯƠNG & PHỤ CẤP 2026', { b: true, sz: 16, color: '#107c41' });
    const rows = [
      [1, 'Nguyễn Văn A', 'Kỹ sư trưởng', 18000000, 3000000, '=D3+E3', '=F3*10.5%', '=F3-G3'],
      [2, 'Trần Thị B', 'Kế toán tổng hợp', 15000000, 2000000, '=D4+E4', '=F4*10.5%', '=F4-G4'],
      [3, 'Lê Văn C', 'Chỉ huy phó', 16000000, 2500000, '=D5+E5', '=F5*10.5%', '=F5-G5'],
      [4, 'Phạm Văn D', 'Giám sát kỹ thuật', 13000000, 1500000, '=D6+E6', '=F6*10.5%', '=F6-G6'],
      [5, 'Hoàng Thị E', 'Hành chính nhân sự', 11000000, 1200000, '=D7+E7', '=F7*10.5%', '=F7-G7']
    ];
    table(s, 1, ['STT', 'Họ và tên', 'Chức danh', 'Lương cơ bản', 'Phụ cấp', 'Tổng thu nhập', 'Trừ BHXH (10.5%)', 'Thực lĩnh'], rows, '#107c41', [45, 160, 140, 120, 100, 130, 130, 130], [null, null, null, '#,##0', '#,##0', '#,##0', '#,##0', '#,##0']);
    const totRow = 2 + rows.length;
    put(s, totRow, 1, 'TỔNG CỘNG', { b: true, bd: BD });
    for (let c = 3; c <= 7; c++) {
      const colLetter = String.fromCharCode(65 + c);
      put(s, totRow, c, '=SUM(' + colLetter + '3:' + colLetter + totRow + ')', { b: true, nf: '#,##0', bd: BD, color: '#c00000' });
    }
  } else if (t.id === 'pnl') {
    s.name = 'P&L';
    put(s, 0, 0, 'BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH (P&L)', { b: true, sz: 16, color: '#1f3864' });
    const pnlRows = [
      ['1. Doanh thu bán hàng và cung cấp dịch vụ', 125000000000, 148000000000, '=(C3-B3)/B3'],
      ['2. Các khoản giảm trừ doanh thu', 1200000000, 1500000000, '=(C4-B4)/B4'],
      ['3. Doanh thu thuần về bán hàng (1 - 2)', '=B3-B4', '=C3-C4', '=(C5-B5)/B5'],
      ['4. Giá vốn hàng bán (COGS)', 85000000000, 98000000000, '=(C6-B6)/B6'],
      ['5. Lợi nhuận gộp về bán hàng (3 - 4)', '=B5-B6', '=C5-C6', '=(C7-B7)/B7'],
      ['6. Chi phí tài chính', 4500000000, 4200000000, '=(C8-B8)/B8'],
      ['7. Chi phí bán hàng & Quản lý doanh nghiệp', 12000000000, 13500000000, '=(C9-B9)/B9'],
      ['8. Lợi nhuận thuần từ HĐKD (5 - 6 - 7)', '=B7-B8-B9', '=C7-C8-C9', '=(C10-B10)/B10'],
      ['9. Lợi nhuận trước thuế (EBT)', '=B10', '=C10', '=(C11-B11)/B11'],
      ['10. Thuế TNDN hiện hành (20%)', '=B11*20%', '=C11*20%', '=(C12-B12)/B12'],
      ['11. Lợi nhuận sau thuế TNDN (NPAT)', '=B11-B12', '=C11-C12', '=(C13-B13)/B13']
    ];
    table(s, 1, ['CHỈ TIÊU TÀI CHÍNH', 'NĂM 2025', 'NĂM 2026', 'TĂNG TRƯỞNG'], pnlRows, '#1f3864', [320, 150, 150, 120], [null, '#,##0', '#,##0', '0.0%']);
  } else if (t.id === 'gantt') {
    s.name = 'Tiến độ';
    put(s, 0, 0, 'KẾ HOẠCH TIẾN ĐỘ DỰ ÁN (GANTT CHART)', { b: true, sz: 16, color: '#0f6cbd' });
    const gRows = [
      [1, 'Khảo sát hiện trường & Lập phương án', '01/03/2026', '10/03/2026', 10, '100%'],
      [2, 'Thiết kế bản vẽ thi công & Dự toán', '11/03/2026', '25/03/2026', 15, '90%'],
      [3, 'San lấp mặt bằng & Thi công cọc móng', '26/03/2026', '25/04/2026', 30, '40%'],
      [4, 'Thi công kết cấu phần thân', '26/04/2026', '25/07/2026', 90, '0%'],
      [5, 'Thi công cơ điện & Hoàn thiện', '26/07/2026', '30/09/2026', 65, '0%'],
      [6, 'Nghiệm thu bàn giao đưa vào sử dụng', '01/10/2026', '15/10/2026', 15, '0%']
    ];
    table(s, 1, ['STT', 'Hạng mục công việc', 'Bắt đầu', 'Kết thúc', 'Số ngày', 'Tiến độ thực hiện'], gRows, '#0f6cbd', [45, 260, 110, 110, 80, 130], [null, null, 'dd/mm/yyyy', 'dd/mm/yyyy', '#,##0', '0%']);
    s.cf.push({ range: { r1: 2, c1: 5, r2: 7, c2: 5 }, type: 'databar', color: '#0f6cbd' });
  } else if (t.id === 'stock_pro') {
    s.name = 'Kho';
    put(s, 0, 0, 'BẢNG TỔNG HỢP XUẤT NHẬP TỒN KHO HÀNG HÓA', { b: true, sz: 16, color: '#c43e1c' });
    const skRows = [
      ['VT001', 'Xi măng Hoàng Thạch PCB40', 'Tấn', 50, 120, 95, '=D3+E3-F3'],
      ['VT002', 'Thép Hòa Phát D16', 'Tấn', 30, 80, 60, '=D4+E4-F4'],
      ['VT003', 'Gạch tuynel đặc 6.5x10.5x22', 'Viên', 50000, 100000, 85000, '=D5+E5-F5'],
      ['VT004', 'Cát vàng đổ bê tông', 'm3', 120, 250, 210, '=D6+E6-F6'],
      ['VT005', 'Đá 1x2 xanh Ninh Bình', 'm3', 150, 300, 260, '=D7+E7-F7']
    ];
    table(s, 1, ['Mã vật tư', 'Tên vật tư hàng hóa', 'ĐVT', 'Tồn đầu kỳ', 'Nhập trong kỳ', 'Xuất trong kỳ', 'Tồn cuối kỳ'], skRows, '#c43e1c', [80, 230, 60, 100, 100, 100, 110], [null, null, null, '#,##0', '#,##0', '#,##0', '#,##0']);
    const tR = 2 + skRows.length;
    put(s, tR, 1, 'TỔNG KHỐI LƯỢNG', { b: true, bd: BD });
    for (let c = 3; c <= 6; c++) {
      const colL = String.fromCharCode(65 + c);
      put(s, tR, c, '=SUM(' + colL + '3:' + colL + tR + ')', { b: true, nf: '#,##0', bd: BD });
    }
  } else if (t.id === 'welcome') {
    s.name = 'Bắt đầu';
    put(s, 0, 0, 'Chào mừng đến với Excel', { b: true, sz: 18, color: '#107c41' });
    table(s, 2, ['Tháng', 'Doanh thu'], [['Tháng 1', 12000000], ['Tháng 2', 15500000], ['Tháng 3', 13800000]], '#107c41', [120, 130], [null, '#,##0']);
    put(s, 6, 0, 'Tổng', { b: true, bd: BD }); put(s, 6, 1, '=SUM(B4:B6)', { b: true, nf: '#,##0', bd: BD });
  } else if (t.id === 'formulas') {
    s.name = 'Công thức';
    table(s, 0, ['Mặt hàng', 'Số lượng'], [['Bút', 40], ['Vở', 65], ['Thước', 30], ['Tẩy', 80], ['Kéo', 55]], '#0f6cbd', [110, 90, 30, 120, 260, 120]);
    const ex = [['Tổng', '=SUM(B2:B6)'], ['Trung bình', '=AVERAGE(B2:B6)'], ['Lớn nhất', '=MAX(B2:B6)'], ['Điều kiện', '=IF(B2>=50,"Đạt","Chưa đạt")'], ['Đếm > 50', '=COUNTIF(B2:B6,">50")'], ['Tra cứu Vở', '=VLOOKUP("Vở",A2:B6,2,FALSE)'], ['XLOOKUP Thước', '=XLOOKUP("Thước",A2:A6,B2:B6)'], ['Hôm nay', '=TEXT(TODAY(),"dd/mm/yyyy")']];
    put(s, 0, 3, 'Ví dụ', { b: true, fill: '#d9e1f2' }); put(s, 0, 4, 'Công thức', { b: true, fill: '#d9e1f2' }); put(s, 0, 5, 'Kết quả', { b: true, fill: '#d9e1f2' });
    ex.forEach((e, i) => { put(s, i + 1, 3, e[0]); s.cells.set(key(i + 1, 4), { v: e[1], s: internStyle({ color: '#555', nf: '@' }) }); put(s, i + 1, 5, e[1], { b: true }); });
  } else if (t.id === 'budget') {
    s.name = 'Ngân sách';
    table(s, 0, ['Hạng mục', 'Dự toán', 'Thực tế', 'Chênh lệch'], [['Địa điểm', 20000000, 18500000, '=B2-C2'], ['Ăn uống', 15000000, 16000000, '=B3-C3'], ['Trang trí', 5000000, 4200000, '=B4-C4'], ['Âm thanh', 7000000, 7000000, '=B5-C5']], '#8764b8', [140, 120, 120, 120], [null, '#,##0', '#,##0', '#,##0;[Red]-#,##0']);
    put(s, 5, 0, 'Tổng', { b: true, bd: BD }); ['B', 'C', 'D'].forEach((c, j) => put(s, 5, j + 1, '=SUM(' + c + '2:' + c + '5)', { b: true, nf: '#,##0', bd: BD }));
    s.cf.push({ range: { r1: 1, c1: 3, r2: 4, c2: 3 }, type: 'databar', color: '#8764b8' });
  } else if (t.id === 'invoice') {
    s.name = 'Hóa đơn';
    put(s, 0, 0, 'HÓA ĐƠN BÁN HÀNG', { b: true, sz: 16, color: '#d83b01' });
    table(s, 2, ['Mặt hàng', 'Số lượng', 'Đơn giá', 'Thành tiền'], [['Giấy A4', 10, 65000, '=B4*C4'], ['Bút bi', 20, 5000, '=B5*C5'], ['Kẹp tài liệu', 5, 12000, '=B6*C6']], '#d83b01', [160, 80, 100, 120], [null, null, '#,##0', '#,##0']);
    put(s, 6, 2, 'Cộng', { b: true }); put(s, 6, 3, '=SUM(D4:D6)', { nf: '#,##0' });
    put(s, 7, 2, 'VAT 10%', { b: true }); put(s, 7, 3, '=D7*10%', { nf: '#,##0' });
    put(s, 8, 2, 'Tổng thanh toán', { b: true }); put(s, 8, 3, '=D7+D8', { b: true, nf: '#,##0 "₫"', color: '#d83b01' });
  }
}
