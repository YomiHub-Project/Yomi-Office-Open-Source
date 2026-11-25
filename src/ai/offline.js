// Offline A.I engine — runs 100% on-device, no network: summarize, keywords, proofread (vi/en), formula help, outline, stats, translate (phrase level)

import { segment, VI_SET } from './viwords.js';
// ---------------------------------------------------------------- text utils
// NOTE: no \p{L}/u regex — unsupported on Android 6 WebView 44; use explicit Latin + Vietnamese ranges
const LT = 'A-Za-z0-9\\u00C0-\\u024F\\u1E00-\\u1EFF';
const WORD_RE = new RegExp('['+LT+']+', 'g');
export function sentences(t) { return (t || '').replace(/\s+/g, ' ').match(/[^.!?…]+[.!?…]+|[^.!?…]+$/g) || []; }
export function words(t) { return ((t || '').toLowerCase().match(WORD_RE) || []); }
const STOP = new Set(('và của là có được cho với các những một này đó trong khi thì mà để từ theo về đã sẽ đang rất không cũng như nhưng hay hoặc nếu vì do bởi tại trên dưới sau trước ra vào lại nên còn đến tới bị người năm ngày việc nhiều ít hơn nhất mỗi mọi chỉ vẫn đều cùng nữa thế nào gì ai đâu bao sao vậy ở nó họ chúng ta tôi bạn anh chị em ông bà '
  + 'the a an and or but of to in on at for with by from is are was were be been being this that these those it its as not no yes do does did have has had will would can could should may might must than then so such into over under about up down out if when while which who whom what where why how all any each few more most other some own same too very just also only').split(' '));
export function keywords(t, n) {
  // segment each phrase into Vietnamese words (dictionary max-match), then rank; unknown adjacent syllables that repeat together are treated as a word too
  const f = {}, pairs = {};
  const phrases = (t || '').toLowerCase().split(/[\n.,;:!?()\[\]"“”'…\-–—\/]+/);
  phrases.forEach(ph => { const sy = ph.match(WORD_RE) || []; for (let i = 0; i < sy.length - 1; i++) pairs[sy[i] + ' ' + sy[i + 1]] = (pairs[sy[i] + ' ' + sy[i + 1]] || 0) + 1; });
  phrases.forEach(ph => {
    const sy = ph.match(WORD_RE) || []; const seg = segment(sy);
    for (let i = 0; i < seg.length; i++) {
      let w = seg[i];
      // merge two unknown single syllables when the pair repeats in the text (learned compound)
      if (w.indexOf(' ') < 0 && i + 1 < seg.length && seg[i + 1].indexOf(' ') < 0 && pairs[w + ' ' + seg[i + 1]] >= 2 && !STOP.has(w) && !STOP.has(seg[i + 1])) { w = w + ' ' + seg[i + 1]; i++; }
      if (STOP.has(w) || /^\d+$/.test(w) || (w.indexOf(' ') < 0 && w.length < 4)) continue;
      f[w] = (f[w] || 0) + (w.indexOf(' ') >= 0 ? 1.6 : 0.8);
    }
  });
  return Object.keys(f).sort((a, b) => f[b] - f[a]).slice(0, n || 10);
}
// extractive summary (TextRank-lite: sentence score = sum of keyword weights, position bonus)
export function summarize(t, ratio) {
  const ss = sentences(t).map(s => s.trim()).filter(s => s.length > 12);
  if (ss.length <= 3) return ss.join(' ');
  const kw = keywords(t, 25); const w = {}; kw.forEach((k, i) => w[k] = 25 - i);
  const scored = ss.map((s, i) => { const lw = s.toLowerCase(); let sc = 0; kw.forEach(k => { if (lw.indexOf(k) >= 0) sc += w[k]; }); sc = sc / Math.sqrt(words(s).length + 1); if (i === 0) sc *= 1.4; return { s, i, sc }; });
  const n = Math.max(2, Math.round(ss.length * (ratio || 0.3)));
  return scored.slice().sort((a, b) => b.sc - a.sc).slice(0, n).sort((a, b) => a.i - b.i).map(x => x.s).join(' ');
}
export function stats(t) {
  const ws = words(t); const ss = sentences(t);
  const chars = (t || '').length, noSpace = (t || '').replace(/\s/g, '').length;
  return { 'Ký tự': chars, 'Ký tự (không khoảng trắng)': noSpace, 'Từ': ws.length, 'Câu': ss.length, 'Đoạn': (t || '').split(/\n\s*\n|\n/).filter(x => x.trim()).length, 'Thời gian đọc': Math.max(1, Math.round(ws.length / 200)) + ' phút', 'Độ dài câu TB': ss.length ? Math.round(ws.length / ss.length) + ' từ' : '—' };
}

// ---------------------------------------------------------------- proofreading
const VI_FIX = [
  ['ko', 'không'], ['k0', 'không'], ['hok', 'không'], ['khg', 'không'], ['dc', 'được'], ['đc', 'được'], ['j', 'gì'], ['bít', 'biết'], ['vs', 'với'], ['ntn', 'như thế nào'], ['bh', 'bây giờ'], ['mn', 'mọi người'], ['tks', 'cảm ơn'], ['nhiu', 'nhiêu'],
  ['sử lý', 'xử lý'], ['xử lí', 'xử lý'], ['chuẩn đoán', 'chẩn đoán'], ['sát nhập', 'sáp nhập'], ['dành dật', 'giành giật'], ['tranh dành', 'tranh giành'], ['dấu diếm', 'giấu giếm'], ['xắp xếp', 'sắp xếp'], ['sắp sếp', 'sắp xếp'], ['xuất xắc', 'xuất sắc'], ['suất sắc', 'xuất sắc'],
  ['nghành', 'ngành'], ['ngiêm', 'nghiêm'], ['nghỉ ngơi', 'nghỉ ngơi'], ['giáo giục', 'giáo dục'], ['trao chuốt', 'trau chuốt'], ['chau chuốt', 'trau chuốt'], ['bàng hoàn', 'bàng hoàng'], ['yếu điểm', 'điểm yếu'], ['bổ xung', 'bổ sung'], ['xem sét', 'xem xét'], ['sơ xuất', 'sơ suất'],
  ['kết quả tốt đẹp', 'kết quả tốt'], ['tham quan', 'tham quan'], ['thăm quan', 'tham quan'], ['đường xá', 'đường sá'], ['xinh xắn', 'xinh xắn'], ['dữ dội', 'dữ dội'], ['vô hình chung', 'vô hình trung'], ['cọ sát', 'cọ xát'], ['sứ mệnh', 'sứ mệnh'], ['chính sát', 'chính xác'], ['trính xác', 'chính xác'], ['trương trình', 'chương trình'], ['chương chình', 'chương trình'],
  ['nhửng', 'những'], ['những gì mà', 'những gì'], ['cám ơn', 'cảm ơn'], ['lãng phí', 'lãng phí'], ['lảng phí', 'lãng phí'], ['giành cho', 'dành cho'], ['dàng buộc', 'ràng buộc'], ['rành buộc', 'ràng buộc'], ['dang dở', 'dang dở'], ['day dứt', 'day dứt'], ['xâu chuỗi', 'xâu chuỗi'], ['sâu chuỗi', 'xâu chuỗi'],
  ['teh', 'the'], ['recieve', 'receive'], ['seperate', 'separate'], ['definately', 'definitely'], ['occured', 'occurred'], ['untill', 'until'], ['wich', 'which'], ['accomodate', 'accommodate'], ['adress', 'address'], ['enviroment', 'environment'], ['goverment', 'government'], ['beleive', 'believe'], ['calender', 'calendar'], ['tommorow', 'tomorrow'], ['alot', 'a lot']
];
const VI_MAP = VI_FIX.filter(x => x[0] !== x[1]);
function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
// returns [{index, length, found, suggestion, reason}]
export function proofreadText(t) {
  const out = []; const add = (index, length, found, suggestion, reason) => out.push({ index, length, found, suggestion, reason });
  let m;
  VI_MAP.forEach(([bad, good]) => { const re = new RegExp('(^|[^' + LT + '])(' + escRe(bad) + ')(?=$|[^' + LT + '])', 'gi'); while ((m = re.exec(t))) { const idx = m.index + m[1].length; const f = m[2]; const sug = /^[A-ZĐ]/.test(f) ? good.charAt(0).toUpperCase() + good.slice(1) : good; add(idx, f.length, f, sug, /^[a-z ]+$/i.test(bad) && /[a-z]/.test(bad) && bad.length > 3 ? 'Lỗi chính tả' : 'Lỗi chính tả / viết tắt'); } });
  const re2 = new RegExp('([' + LT + ']+)\\s+\\1(?=[^' + LT + ']|$)', 'gi'); while ((m = re2.exec(t))) add(m.index, m[0].length, m[0], m[1], 'Lặp từ');
  const re3 = / {2,}/g; while ((m = re3.exec(t))) add(m.index, m[0].length, m[0], ' ', 'Thừa khoảng trắng');
  const re4 = /\s+([,.;:!?])/g; while ((m = re4.exec(t))) add(m.index, m[0].length, m[0], m[1], 'Khoảng trắng trước dấu câu');
  const re5 = new RegExp('([,;:])(?=[' + LT + '])', 'g'); while ((m = re5.exec(t))) add(m.index, 1, m[1], m[1] + ' ', 'Thiếu khoảng trắng sau dấu câu');
  const re6 = /([.!?])\s+([a-zđàáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵ])/g; while ((m = re6.exec(t))) add(m.index + m[0].length - 1, 1, m[2], m[2].toUpperCase(), 'Viết hoa đầu câu');
  return out.sort((a, b) => a.index - b.index).filter((x, i, a) => !i || x.index >= a[i - 1].index + a[i - 1].length);
}
export function applyFixes(t, issues) { let r = t; issues.slice().sort((a, b) => b.index - a.index).forEach(x => { r = r.slice(0, x.index) + x.suggestion + r.slice(x.index + x.length); }); return r; }

// ---------------------------------------------------------------- Excel formula helper (natural language → formula)
export function formulaFromText(q, ref) {
  const s = (q || '').toLowerCase(); const r = ref || 'A1:A10';
  const rng = (/([a-z]{1,3}\d+\s*:\s*[a-z]{1,3}\d+)/i.exec(q) || [])[1]; const R = rng ? rng.toUpperCase().replace(/\s/g, '') : r;
  const rules = [
    [/tổng|sum|cộng/, `=SUM(${R})`, 'Tính tổng các ô'], [/trung bình|average|bình quân/, `=AVERAGE(${R})`, 'Giá trị trung bình'],
    [/đếm.*(rỗng|trống)|countblank/, `=COUNTBLANK(${R})`, 'Đếm ô trống'], [/đếm.*điều kiện|countif/, `=COUNTIF(${R},">0")`, 'Đếm theo điều kiện'], [/đếm|count/, `=COUNTA(${R})`, 'Đếm ô có dữ liệu'],
    [/lớn nhất|max|cao nhất/, `=MAX(${R})`, 'Giá trị lớn nhất'], [/nhỏ nhất|min|thấp nhất/, `=MIN(${R})`, 'Giá trị nhỏ nhất'],
    [/tìm|tra cứu|vlookup|lookup/, `=XLOOKUP(E2,A2:A100,B2:B100,"Không thấy")`, 'Tra cứu giá trị (XLOOKUP)'], [/nếu|if|điều kiện/, `=IF(A2>=5,"Đạt","Không đạt")`, 'Rẽ nhánh theo điều kiện'],
    [/xếp hạng|rank/, `=RANK.EQ(A2,${R})`, 'Xếp hạng'], [/làm tròn|round/, `=ROUND(A2,0)`, 'Làm tròn'], [/phần trăm|percent|%/, `=A2/SUM(${R})`, 'Tỷ lệ phần trăm (định dạng %)'],
    [/ngày|date|hôm nay|today/, `=TODAY()`, 'Ngày hiện tại'], [/tuổi|age/, `=DATEDIF(A2,TODAY(),"Y")`, 'Tính tuổi'], [/ghép|nối|concat|join/, `=TEXTJOIN(" ",TRUE,A2:C2)`, 'Nối chuỗi'],
    [/duy nhất|unique|loại trùng/, `=UNIQUE(${R})`, 'Danh sách không trùng'], [/sắp xếp|sort/, `=SORT(${R})`, 'Sắp xếp (mảng động)'], [/lọc|filter/, `=FILTER(A2:C100,B2:B100>0)`, 'Lọc dữ liệu (mảng động)'],
    [/khoản vay|trả góp|pmt|lãi/, `=PMT(lãi_suất/12,số_tháng,-tiền_vay)`, 'Tiền trả góp hàng tháng'], [/độ lệch|stdev/, `=STDEV.S(${R})`, 'Độ lệch chuẩn']
  ];
  for (const x of rules) if (x[0].test(s)) return { formula: x[1], explain: x[2] };
  return null;
}

// ---------------------------------------------------------------- outline / slides
function headline(sent, maxWords) {
  let s = (sent || '').replace(/^[\s•\-–*\d.)]+/, '').trim(); s = s.split(/[,;:]/)[0].replace(/[.!?…]+$/, '');
  const raw = s.split(/\s+/); if (raw.length <= (maxWords || 6) + 2) return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
  const seg = segment(raw.map(x => x.toLowerCase())); let n = 0, out = [];
  for (const w of seg) { const k = w.split(' ').length; if (n + k > (maxWords || 6) && out.length) break; out.push(w); n += k; }
  while (out.length > 1 && STOP.has(out[out.length - 1])) out.pop();
  const txt = raw.slice(0, out.join(' ').split(' ').length).join(' ');
  return txt.charAt(0).toUpperCase() + txt.slice(1);
}
export function cleanDocText(t) { return (t || '').split('\n').filter(l => !/^---\s*Slide \d+\s*---$/.test(l.trim())).map(l => l.replace(/^\[Ghi chú\]\s*/, '')).join('\n'); }
export function outlineFromText(t, maxSlides) {
  t = cleanDocText(t);
  const paras = t.split(/\n\s*\n|\n/).map(x => x.trim()).filter(Boolean);
  const slides = []; maxSlides = maxSlides || 8;
  if (!paras.length) return slides;
  const first = paras[0]; const title = first.length <= 90 ? first.replace(/[.!?]+$/, '') : headline(first, 8);
  const body = paras.slice(paras[0].length <= 90 ? 1 : 0);
  const allSent = sentences(body.join(' ')).map(x => x.trim()).filter(x => x.length > 3);
  slides.push({ title, bullets: [keywords(t, 3).map(k => k.charAt(0).toUpperCase() + k.slice(1)).join(' · ')] });
  // group sentences into content slides of up to 3 bullets
  const per = Math.max(2, Math.min(4, Math.ceil(allSent.length / Math.max(1, maxSlides - 2))));
  for (let i = 0; i < allSent.length && slides.length < maxSlides - 1; i += per) {
    const grp = allSent.slice(i, i + per);
    slides.push({ title: headline(grp[0], 5), bullets: grp.map(s => s.replace(/[.]+$/, '')).map(s => s.length > 120 ? s.slice(0, 117) + '…' : s) });
  }
  if (slides.length > 2) slides.push({ title: 'Tổng kết', bullets: slides.slice(1).map(s => s.title) });
  return slides;
}

// ---------------------------------------------------------------- phrase-level translation (offline glossary for common office text)
const GLOSS = { 'xin chào': 'hello', 'cảm ơn': 'thank you', 'báo cáo': 'report', 'doanh thu': 'revenue', 'lợi nhuận': 'profit', 'chi phí': 'cost', 'kế hoạch': 'plan', 'dự án': 'project', 'cuộc họp': 'meeting', 'hợp đồng': 'contract', 'khách hàng': 'customer', 'nhân viên': 'employee', 'công ty': 'company', 'ngân sách': 'budget', 'tổng cộng': 'total', 'tháng': 'month', 'năm': 'year', 'ngày': 'day', 'tuần': 'week', 'quý': 'quarter', 'mục tiêu': 'goal', 'kết quả': 'result', 'giải pháp': 'solution', 'vấn đề': 'issue', 'tài liệu': 'document', 'bảng tính': 'spreadsheet', 'trình chiếu': 'presentation', 'sản phẩm': 'product', 'dịch vụ': 'service', 'thị trường': 'market', 'tăng trưởng': 'growth', 'giảm': 'decrease', 'tăng': 'increase', 'trân trọng': 'sincerely', 'kính gửi': 'dear', 'đính kèm': 'attached', 'hạn chót': 'deadline', 'phê duyệt': 'approve', 'đề xuất': 'proposal', 'tóm tắt': 'summary', 'mục lục': 'table of contents', 'phụ lục': 'appendix', 'và': 'and', 'của': 'of', 'cho': 'for', 'với': 'with', 'trong': 'in', 'không': 'not', 'có': 'have' };
export function translateGlossary(t, toEn) {
  const pairs = Object.keys(GLOSS).map(k => toEn ? [k, GLOSS[k]] : [GLOSS[k], k]).sort((a, b) => b[0].length - a[0].length);
  let r = t; pairs.forEach(([a, b]) => { r = r.replace(new RegExp('(^|[^' + LT + '])' + escRe(a) + '(?=$|[^' + LT + '])', 'gi'), (m, p) => p + b); });
  return r;
}

// ---------------------------------------------------------------- intent router for the offline chat
export function offlineAnswer(prompt, ctx) {
  const p = (prompt || '').toLowerCase(); const text = cleanDocText(ctx.selection || ctx.text || '');
  const has = re => re.test(p);
  if (has(/tóm tắt|summar/)) return text ? { text: '📝 Tóm tắt:\n' + summarize(text, 0.3), insert: summarize(text, 0.3) } : { text: 'Hãy mở tài liệu có nội dung hoặc bôi đen đoạn cần tóm tắt.' };
  if (has(/từ khóa|keyword|ý chính/)) return { text: '🔑 Từ khóa chính:\n• ' + keywords(text, 12).join('\n• ') };
  if (has(/soát|chính tả|lỗi|proofread|spell/)) { const iss = proofreadText(text); return { text: iss.length ? '✍️ Tìm thấy ' + iss.length + ' lỗi:\n' + iss.slice(0, 30).map(x => '• "' + x.found + '" → "' + x.suggestion + '" (' + x.reason + ')').join('\n') : '✅ Không phát hiện lỗi chính tả phổ biến.', insert: iss.length ? applyFixes(text, iss) : null, fixLabel: 'Thay bằng bản đã sửa' }; }
  if (has(/công thức|formula|hàm/)) { const f = formulaFromText(prompt, ctx.range); return f ? { text: '🧮 ' + f.explain + ':\n' + f.formula, insert: f.formula, fixLabel: 'Chèn vào ô' } : { text: 'Hãy mô tả, ví dụ: "công thức tính tổng B2:B20", "tra cứu giá theo mã", "đếm ô trống".' }; }
  if (has(/slide|trình chiếu|dàn ý|outline/)) { const ol = outlineFromText(text); return { text: '🎞️ Dàn ý ' + ol.length + ' trang chiếu:\n' + ol.map((s, i) => (i + 1) + '. ' + s.title + '\n' + s.bullets.map(b => '   – ' + b).join('\n')).join('\n'), outline: ol }; }
  if (has(/thống kê|đếm từ|word count|stat/)) { const s = stats(text); return { text: '📊 Thống kê:\n' + Object.keys(s).map(k => '• ' + k + ': ' + s[k]).join('\n') }; }
  if (has(/dịch|translate/)) { const toEn = !has(/sang tiếng việt|to vietnamese|tiếng việt$/); const src = text || prompt.replace(/.*?:/, ''); const r = translateGlossary(src, toEn); return { text: '🌐 Dịch thuật ngữ (ngoại tuyến, mức cụm từ):\n' + r + '\n\nℹ️ Để dịch trọn câu tự nhiên, chọn một dịch vụ A.I (hoặc Ollama chạy cục bộ) trong ⚙️.', insert: r }; }
  if (has(/viết hoa|uppercase/)) return { text: text.toUpperCase(), insert: text.toUpperCase() };
  if (has(/chữ thường|lowercase/)) return { text: text.toLowerCase(), insert: text.toLowerCase() };
  return { text: 'Chế độ ngoại tuyến hỗ trợ: tóm tắt, từ khóa, soát chính tả, gợi ý công thức Excel, tạo dàn ý trình chiếu, thống kê văn bản, dịch thuật ngữ, đổi chữ hoa/thường.\n\nĐể trò chuyện tự do, bấm ⚙️ và chọn OpenAI / Claude / Gemini / DeepSeek / Groq / OpenRouter (cần khóa API) hoặc Ollama / LM Studio (chạy mô hình trên máy, không cần Internet).' };
}
