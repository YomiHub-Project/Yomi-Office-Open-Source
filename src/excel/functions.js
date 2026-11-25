// Excel function library (~480 functions): native implementations + Formula.js adapter for the long tail
import { defFn, FN, Ref, Lambda, num, str, bool, cmp, isArr, parse } from './formula.js';
import { Err, E, isErr, colName, key, dateToSerial, serialToParts, jsDateToSerial, serialToJsDate, MAX_ROWS, MAX_COLS, parseRangeAddr, addr } from './model.js';
import { format, parseInput, generalNumber } from './numfmt.js';
import * as FJS from '@formulajs/formulajs';

// ---------------------------------------------------------------- helpers
const chk = v => { if (v instanceof Err) throw v; return v; };
const N = v => chk(num(v)), S = v => chk(str(v)), B = v => chk(bool(v));
function scalar(calc, v) { if (v instanceof Ref) return calc.deref(v, true); if (isArr(v)) return v[0][0]; return v; }
function arr2(calc, v) { if (v instanceof Ref) return calc.deref(v); if (isArr(v)) return v; return [[v]]; }
function dims(calc, v) { if (v instanceof Ref) return [v.rows, v.cols]; if (isArr(v)) return [v.length, v[0].length]; return [1, 1]; }
function flatVals(calc, v, out) { out = out || []; if (v instanceof Ref) calc.eachIn(v, x => out.push(x)); else if (isArr(v)) v.forEach(r => r.forEach(x => out.push(x))); else out.push(v); return out; }
// numbers for aggregation: direct args coerce text/bool, range/array members only numbers
function nums(calc, args, opts) {
  opts = opts || {}; const out = [];
  for (const a of args) {
    if (a instanceof Ref || isArr(a)) {
      const list = flatVals(calc, a);
      for (const x of list) { if (typeof x === 'number') out.push(x); else if (x instanceof Err) { if (!opts.ignoreErr) throw x; } else if (opts.a) out.push(typeof x === 'boolean' ? (x ? 1 : 0) : (x == null ? null : 0)); }
    } else if (a == null) { if (opts.countMissing) out.push(0); }
    else { const n = num(a); if (n instanceof Err) throw n; out.push(n); }
  }
  return opts.a ? out.filter(x => x !== null) : out;
}
function lift(f) { return f; }
const sum = a => { let s = 0; for (const x of a) s += x; return s; };
function ser(y, m, d) { return dateToSerial(y, m, d); }
function toInt(v) { return Math.trunc(N(v)); }
function wild(p) { return new RegExp('^' + p.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '[\\s\\S]').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'i'); }
// criteria predicate for *IF(S)
export function criteria(c) {
  if (c instanceof Err) return x => x instanceof Err && x.code === c.code;
  if (typeof c === 'number') return x => typeof x === 'number' ? x === c : (typeof x === 'string' && x.trim() !== '' && +x === c);
  if (typeof c === 'boolean') return x => x === c;
  if (c == null) c = '';
  const s = String(c);
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s);
  const op = m[1] || '=', rhs = m[2];
  const p = parseInput(rhs); const isNum = typeof p.v === 'number' && rhs.trim() !== '';
  if (isNum) {
    const n = p.v;
    return x => { if (typeof x !== 'number') { if (op === '<>') return true; return false; } switch (op) { case '=': return x === n; case '<>': return x !== n; case '<': return x < n; case '>': return x > n; case '<=': return x <= n; case '>=': return x >= n; } };
  }
  if (rhs === '' && op === '=') return x => x == null || x === '';
  if (rhs === '' && op === '<>') return x => !(x == null || x === '');
  if (/^(TRUE|FALSE)$/i.test(rhs)) { const b = /^TRUE$/i.test(rhs); return x => op === '<>' ? x !== b : x === b; }
  if (op === '=' || op === '<>') { const re = wild(rhs); return x => { const ok = x != null && !(x instanceof Err) && typeof x !== 'number' && typeof x !== 'boolean' && re.test(String(x)); return op === '=' ? ok : !ok; }; }
  return x => { if (typeof x !== 'string') return false; const c2 = cmp(x, rhs); return op === '<' ? c2 < 0 : op === '>' ? c2 > 0 : op === '<=' ? c2 <= 0 : c2 >= 0; };
}
function ifsCollect(calc, pairs, target) {
  // pairs: [range, crit]...; returns list of [row, col] offsets matching all; uses dims of first range
  const [rows, cols] = dims(calc, pairs[0]);
  const grids = []; const preds = [];
  for (let i = 0; i < pairs.length; i += 2) { const d = dims(calc, pairs[i]); if (d[0] !== rows || d[1] !== cols) throw E.VALUE; grids.push(arr2(calc, pairs[i])); preds.push(criteria(scalar(calc, pairs[i + 1]))); }
  const hits = [];
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) { let ok = true; for (let k = 0; k < grids.length && ok; k++) ok = preds[k]((grids[k][r] || [])[c]); if (ok) hits.push([r, c]); }
  return hits;
}
function sortedNums(a) { return a.slice().sort((x, y) => x - y); }
function variance(a, pop) { const n = a.length; if (n - (pop ? 0 : 1) <= 0) throw E.DIV0; const m = sum(a) / n; let s = 0; for (const x of a) s += (x - m) * (x - m); return s / (n - (pop ? 0 : 1)); }
function percentileInc(a, k) { if (!a.length || k < 0 || k > 1) throw E.NUM; const s = sortedNums(a); const p = (s.length - 1) * k; const lo = Math.floor(p), hi = Math.ceil(p); return s[lo] + (s[hi] - s[lo]) * (p - lo); }
function percentileExc(a, k) { const s = sortedNums(a); const n = s.length; const p = k * (n + 1) - 1; if (p < 0 || p > n - 1) throw E.NUM; const lo = Math.floor(p), hi = Math.ceil(p); return s[lo] + (s[hi] - s[lo]) * (p - lo); }
function toArr(v) { return isArr(v) ? v : [[v]]; }

// ---------------------------------------------------------------- definitions helper
const meta = {}; // name -> [category, argsSignature, vi description]
export function registerMeta(name, cat, sig, desc) { meta[name] = [cat, sig, desc]; }
function def(name, cat, sig, desc, f, opts) { defFn(name, f, opts); meta[name] = [cat, sig, desc]; }
const MATH = 'Toán học & Lượng giác', STAT = 'Thống kê', LOGIC = 'Lô-gic', TEXT = 'Văn bản', DATE = 'Ngày & Giờ', LOOK = 'Tra cứu & Tham chiếu', INFO = 'Thông tin', FIN = 'Tài chính', ENG = 'Kỹ thuật', DB = 'Cơ sở dữ liệu', COMPAT = 'Tương thích', WEB = 'Web', CUBE = 'Khối (Cube)';
const SC = { scalar: true };

// ================================================================ MATH & aggregates
def('SUM', MATH, 'number1, [number2], ...', 'Cộng tất cả các số trong một phạm vi ô.', (a, x, c) => fixf(sum(nums(c, a))));
def('SUMSQ', MATH, 'number1, [number2], ...', 'Tổng bình phương các đối số.', (a, x, c) => sum(nums(c, a).map(v => v * v)));
def('PRODUCT', MATH, 'number1, [number2], ...', 'Nhân tất cả các số.', (a, x, c) => { const n = nums(c, a); return n.length ? n.reduce((p, v) => p * v, 1) : 0; });
def('AVERAGE', STAT, 'number1, [number2], ...', 'Trung bình cộng của các đối số.', (a, x, c) => { const n = nums(c, a); if (!n.length) throw E.DIV0; return sum(n) / n.length; });
def('AVERAGEA', STAT, 'value1, [value2], ...', 'Trung bình, tính cả văn bản (0) và giá trị logic.', (a, x, c) => { const n = nums(c, a, { a: true }); if (!n.length) throw E.DIV0; return sum(n) / n.length; });
def('COUNT', STAT, 'value1, [value2], ...', 'Đếm số ô chứa số.', (a, x, c) => { let k = 0; for (const v of a) { if (v instanceof Ref || isArr(v)) flatVals(c, v).forEach(z => { if (typeof z === 'number') k++; }); else if (typeof v === 'number' || (typeof v === 'string' && typeof parseInput(v).v === 'number') || typeof v === 'boolean') k++; } return k; });
def('COUNTA', STAT, 'value1, [value2], ...', 'Đếm số ô không trống.', (a, x, c) => { let k = 0; for (const v of a) { if (v instanceof Ref || isArr(v)) flatVals(c, v).forEach(z => { if (z != null && z !== '') k++; }); else if (v !== null) k++; } return k; });
def('COUNTBLANK', STAT, 'range', 'Đếm số ô trống trong phạm vi.', (a, x, c) => { let k = 0; const r = a[0]; if (r instanceof Ref) { const tot = r.rows * r.cols; let filled = 0; c.eachIn(r, z => { if (z != null && z !== '') filled++; }); return tot - filled; } flatVals(c, r).forEach(z => { if (z == null || z === '') k++; }); return k; });
def('MAX', STAT, 'number1, [number2], ...', 'Giá trị lớn nhất.', (a, x, c) => { const n = nums(c, a); return n.length ? Math.max.apply(null, n) : 0; });
def('MIN', STAT, 'number1, [number2], ...', 'Giá trị nhỏ nhất.', (a, x, c) => { const n = nums(c, a); return n.length ? Math.min.apply(null, n) : 0; });
def('MAXA', STAT, 'value1, [value2], ...', 'Giá trị lớn nhất, tính cả văn bản và logic.', (a, x, c) => { const n = nums(c, a, { a: true }); return n.length ? Math.max.apply(null, n) : 0; });
def('MINA', STAT, 'value1, [value2], ...', 'Giá trị nhỏ nhất, tính cả văn bản và logic.', (a, x, c) => { const n = nums(c, a, { a: true }); return n.length ? Math.min.apply(null, n) : 0; });
def('MEDIAN', STAT, 'number1, [number2], ...', 'Trung vị của các số.', (a, x, c) => { const n = sortedNums(nums(c, a)); if (!n.length) throw E.NUM; const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; });
const modeS = (a, x, c) => { const n = nums(c, a); const cnt = new Map(); let best = null, bc = 1; for (const v of n) { const k = (cnt.get(v) || 0) + 1; cnt.set(v, k); if (k > bc) { bc = k; best = v; } } if (best === null) throw E.NA; return best; };
def('MODE', COMPAT, 'number1, [number2], ...', 'Giá trị xuất hiện nhiều nhất.', modeS); def('MODE.SNGL', STAT, 'number1, [number2], ...', 'Giá trị xuất hiện nhiều nhất.', modeS);
def('MODE.MULT', STAT, 'number1, [number2], ...', 'Mảng dọc các giá trị xuất hiện nhiều nhất.', (a, x, c) => { const n = nums(c, a); const cnt = new Map(); n.forEach(v => cnt.set(v, (cnt.get(v) || 0) + 1)); let m = 1; cnt.forEach(v => { if (v > m) m = v; }); const out = []; cnt.forEach((v, k) => { if (v === m && m > 1) out.push([k]); }); if (!out.length) throw E.NA; return out; });
def('ABS', MATH, 'number', 'Giá trị tuyệt đối.', a => Math.abs(N(a[0])), SC);
def('SIGN', MATH, 'number', 'Dấu của số (1, 0, -1).', a => Math.sign ? Math.sign(N(a[0])) : (N(a[0]) > 0 ? 1 : N(a[0]) < 0 ? -1 : 0), SC);
def('INT', MATH, 'number', 'Làm tròn xuống số nguyên gần nhất.', a => Math.floor(N(a[0])), SC);
def('TRUNC', MATH, 'number, [num_digits]', 'Cắt bỏ phần thập phân.', a => { const n = N(a[0]), d = a[1] == null ? 0 : toInt(a[1]); const f = Math.pow(10, d); return Math.trunc(n * f) / f; }, SC);
function roundX(n, d, mode) { const f = Math.pow(10, d); const x = Math.abs(n) * f; const r = mode === 0 ? Math.round(+(x.toPrecision(15))) : mode > 0 ? Math.ceil(+(x.toPrecision(15))) : Math.floor(+(x.toPrecision(15))); return (n < 0 ? -r : r) / f; }
def('ROUND', MATH, 'number, num_digits', 'Làm tròn số tới số chữ số chỉ định.', a => roundX(N(a[0]), toInt(a[1]), 0), SC);
def('ROUNDUP', MATH, 'number, num_digits', 'Làm tròn lên (xa số 0).', a => roundX(N(a[0]), toInt(a[1]), 1), SC);
def('ROUNDDOWN', MATH, 'number, num_digits', 'Làm tròn xuống (về phía số 0).', a => roundX(N(a[0]), toInt(a[1]), -1), SC);
def('MROUND', MATH, 'number, multiple', 'Làm tròn tới bội số gần nhất.', a => { const n = N(a[0]), m = N(a[1]); if (m === 0) return 0; if (n * m < 0) throw E.NUM; return Math.round(n / m) * m; }, SC);
const ceil = (n, s) => { if (s === 0) return 0; if (n > 0 && s < 0) throw E.NUM; return Math.ceil(+(n / s).toPrecision(15)) * s; };
const floorF = (n, s) => { if (s === 0) { if (n === 0) return 0; throw E.DIV0; } if (n > 0 && s < 0) throw E.NUM; return Math.floor(+(n / s).toPrecision(15)) * s; };
def('CEILING', MATH, 'number, significance', 'Làm tròn lên tới bội số của significance.', a => ceil(N(a[0]), N(a[1])), SC);
def('CEILING.MATH', MATH, 'number, [significance], [mode]', 'Làm tròn lên tới số nguyên hoặc bội số gần nhất.', a => { const n = N(a[0]); let s = a[1] == null ? 1 : Math.abs(N(a[1])); const mode = a[2] == null ? 0 : N(a[2]); if (!s) return 0; if (n < 0 && mode) return -Math.ceil(Math.abs(n) / s) * s; return Math.ceil(+(n / s).toPrecision(15)) * s; }, SC);
def('CEILING.PRECISE', MATH, 'number, [significance]', 'Làm tròn lên bất kể dấu.', a => { const n = N(a[0]); const s = a[1] == null ? 1 : Math.abs(N(a[1])); return s ? Math.ceil(+(n / s).toPrecision(15)) * s : 0; }, SC);
FN['ISO.CEILING'] = FN['CEILING.PRECISE']; meta['ISO.CEILING'] = [MATH, 'number, [significance]', 'Làm tròn lên (chuẩn ISO).'];
def('FLOOR', MATH, 'number, significance', 'Làm tròn xuống tới bội số.', a => floorF(N(a[0]), N(a[1])), SC);
def('FLOOR.MATH', MATH, 'number, [significance], [mode]', 'Làm tròn xuống tới số nguyên/bội số gần nhất.', a => { const n = N(a[0]); const s = a[1] == null ? 1 : Math.abs(N(a[1])); const mode = a[2] == null ? 0 : N(a[2]); if (!s) return 0; if (n < 0 && mode) return -Math.floor(Math.abs(n) / s) * s; return Math.floor(+(n / s).toPrecision(15)) * s; }, SC);
def('FLOOR.PRECISE', MATH, 'number, [significance]', 'Làm tròn xuống bất kể dấu.', a => { const n = N(a[0]); const s = a[1] == null ? 1 : Math.abs(N(a[1])); return s ? Math.floor(+(n / s).toPrecision(15)) * s : 0; }, SC);
def('EVEN', MATH, 'number', 'Làm tròn lên số nguyên chẵn gần nhất.', a => { const n = N(a[0]); const v = Math.ceil(Math.abs(n) / 2) * 2; return n < 0 ? -v : v; }, SC);
def('ODD', MATH, 'number', 'Làm tròn lên số nguyên lẻ gần nhất.', a => { const n = N(a[0]); let v = Math.ceil(Math.abs(n)); if (v % 2 === 0) v++; return n < 0 ? -v : v; }, SC);
def('MOD', MATH, 'number, divisor', 'Phần dư của phép chia.', a => { const n = N(a[0]), d = N(a[1]); if (d === 0) throw E.DIV0; const r = n - d * Math.floor(n / d); return +r.toPrecision(15); }, SC);
def('QUOTIENT', MATH, 'numerator, denominator', 'Phần nguyên của phép chia.', a => { const d = N(a[1]); if (d === 0) throw E.DIV0; return Math.trunc(N(a[0]) / d); }, SC);
def('POWER', MATH, 'number, power', 'Lũy thừa.', a => { const r = Math.pow(N(a[0]), N(a[1])); if (!isFinite(r) || isNaN(r)) throw E.NUM; return r; }, SC);
def('SQRT', MATH, 'number', 'Căn bậc hai.', a => { const n = N(a[0]); if (n < 0) throw E.NUM; return Math.sqrt(n); }, SC);
def('SQRTPI', MATH, 'number', 'Căn bậc hai của (số * Pi).', a => { const n = N(a[0]); if (n < 0) throw E.NUM; return Math.sqrt(n * Math.PI); }, SC);
def('EXP', MATH, 'number', 'e mũ n.', a => Math.exp(N(a[0])), SC);
def('LN', MATH, 'number', 'Logarit tự nhiên.', a => { const n = N(a[0]); if (n <= 0) throw E.NUM; return Math.log(n); }, SC);
def('LOG', MATH, 'number, [base]', 'Logarit theo cơ số.', a => { const n = N(a[0]), b = a[1] == null ? 10 : N(a[1]); if (n <= 0 || b <= 0 || b === 1) throw E.NUM; return Math.log(n) / Math.log(b); }, SC);
def('LOG10', MATH, 'number', 'Logarit cơ số 10.', a => { const n = N(a[0]); if (n <= 0) throw E.NUM; return Math.log(n) / Math.LN10; }, SC);
def('PI', MATH, '', 'Số Pi (3.14159…).', () => Math.PI);
def('RAND', MATH, '', 'Số ngẫu nhiên từ 0 đến 1.', () => Math.random(), { volatile: true });
def('RANDBETWEEN', MATH, 'bottom, top', 'Số nguyên ngẫu nhiên trong khoảng.', a => { const lo = Math.ceil(N(a[0])), hi = Math.floor(N(a[1])); if (lo > hi) throw E.NUM; return lo + Math.floor(Math.random() * (hi - lo + 1)); }, { volatile: true });
def('RANDARRAY', MATH, '[rows], [columns], [min], [max], [integer]', 'Mảng số ngẫu nhiên.', a => { const r = a[0] == null ? 1 : toInt(a[0]), c = a[1] == null ? 1 : toInt(a[1]), lo = a[2] == null ? 0 : N(a[2]), hi = a[3] == null ? 1 : N(a[3]), int = a[4] != null && B(a[4]); const out = []; for (let i = 0; i < r; i++) { const row = []; for (let j = 0; j < c; j++) row.push(int ? Math.floor(lo + Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo)); out.push(row); } return out; }, { volatile: true });
def('SEQUENCE', MATH, 'rows, [columns], [start], [step]', 'Tạo dãy số liên tiếp.', a => { const r = toInt(a[0]), c = a[1] == null ? 1 : toInt(a[1]), st = a[2] == null ? 1 : N(a[2]), sp = a[3] == null ? 1 : N(a[3]); if (r < 1 || c < 1) throw E.CALC; const out = []; let v = st; for (let i = 0; i < r; i++) { const row = []; for (let j = 0; j < c; j++) { row.push(v); v += sp; } out.push(row); } return out; });
const fact = n => { n = Math.floor(n); if (n < 0) throw E.NUM; let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; };
def('FACT', MATH, 'number', 'Giai thừa.', a => fact(N(a[0])), SC);
def('COMBIN', MATH, 'number, number_chosen', 'Tổ hợp chập k của n.', a => { const n = Math.floor(N(a[0])), k = Math.floor(N(a[1])); if (n < 0 || k < 0 || k > n) throw E.NUM; let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); }, SC);
def('PERMUT', STAT, 'number, number_chosen', 'Chỉnh hợp chập k của n.', a => { const n = Math.floor(N(a[0])), k = Math.floor(N(a[1])); if (n < 0 || k < 0 || k > n) throw E.NUM; let r = 1; for (let i = 0; i < k; i++) r *= n - i; return r; }, SC);
const gcd2 = (x, y) => { while (y) { const t = y; y = x % y; x = t; } return x; };
def('GCD', MATH, 'number1, [number2], ...', 'Ước số chung lớn nhất.', (a, x, c) => nums(c, a).map(v => { if (v < 0) throw E.NUM; return Math.floor(v); }).reduce(gcd2, 0));
def('LCM', MATH, 'number1, [number2], ...', 'Bội số chung nhỏ nhất.', (a, x, c) => nums(c, a).map(v => { if (v < 0) throw E.NUM; return Math.floor(v); }).reduce((p, v) => (p && v) ? p * v / gcd2(p, v) : 0, 1));
def('SUMPRODUCT', MATH, 'array1, [array2], ...', 'Tổng các tích của các phần tử tương ứng.', (a, x, c) => {
  const arrs = a.map(v => arr2(c, v)); const r = arrs[0].length, cc = arrs[0][0].length;
  for (const m of arrs) if (m.length !== r || m[0].length !== cc) throw E.VALUE;
  let s = 0; for (let i = 0; i < r; i++) for (let j = 0; j < cc; j++) { let p = 1; for (const m of arrs) { const v = m[i][j]; if (v instanceof Err) throw v; p *= typeof v === 'number' ? v : typeof v === 'boolean' && arrs.length === 1 ? 0 : (typeof v === 'boolean' ? (v ? 1 : 0) : 0); } s += p; }
  return s;
});
def('SUMIF', MATH, 'range, criteria, [sum_range]', 'Cộng các ô thỏa mãn điều kiện.', (a, x, c) => {
  const hits = ifsCollect(c, [a[0], a[1]]); const sr = a[2] instanceof Ref ? new Ref(a[2].sheet, a[2].r1, a[2].c1, a[2].r1 + dims(c, a[0])[0] - 1, a[2].c1 + dims(c, a[0])[1] - 1) : (a[2] || a[0]);
  const g = arr2(c, sr); let s = 0; for (const [r, cc] of hits) { const v = (g[r] || [])[cc]; if (typeof v === 'number') s += v; else if (v instanceof Err) throw v; } return fixf(s);
});
def('SUMIFS', MATH, 'sum_range, criteria_range1, criteria1, ...', 'Cộng các ô thỏa mãn nhiều điều kiện.', (a, x, c) => { const hits = ifsCollect(c, a.slice(1)); const g = arr2(c, a[0]); let s = 0; for (const [r, cc] of hits) { const v = (g[r] || [])[cc]; if (typeof v === 'number') s += v; } return fixf(s); });
def('COUNTIF', STAT, 'range, criteria', 'Đếm số ô thỏa mãn điều kiện.', (a, x, c) => ifsCollect(c, [a[0], a[1]]).length);
def('COUNTIFS', STAT, 'criteria_range1, criteria1, ...', 'Đếm số ô thỏa mãn nhiều điều kiện.', (a, x, c) => ifsCollect(c, a).length);
def('AVERAGEIF', STAT, 'range, criteria, [average_range]', 'Trung bình các ô thỏa mãn điều kiện.', (a, x, c) => { const hits = ifsCollect(c, [a[0], a[1]]); const g = arr2(c, a[2] || a[0]); const v = hits.map(([r, cc]) => (g[r] || [])[cc]).filter(z => typeof z === 'number'); if (!v.length) throw E.DIV0; return sum(v) / v.length; });
def('AVERAGEIFS', STAT, 'average_range, criteria_range1, criteria1, ...', 'Trung bình theo nhiều điều kiện.', (a, x, c) => { const hits = ifsCollect(c, a.slice(1)); const g = arr2(c, a[0]); const v = hits.map(([r, cc]) => (g[r] || [])[cc]).filter(z => typeof z === 'number'); if (!v.length) throw E.DIV0; return sum(v) / v.length; });
def('MAXIFS', STAT, 'max_range, criteria_range1, criteria1, ...', 'Giá trị lớn nhất theo điều kiện.', (a, x, c) => { const hits = ifsCollect(c, a.slice(1)); const g = arr2(c, a[0]); const v = hits.map(([r, cc]) => (g[r] || [])[cc]).filter(z => typeof z === 'number'); return v.length ? Math.max.apply(null, v) : 0; });
def('MINIFS', STAT, 'min_range, criteria_range1, criteria1, ...', 'Giá trị nhỏ nhất theo điều kiện.', (a, x, c) => { const hits = ifsCollect(c, a.slice(1)); const g = arr2(c, a[0]); const v = hits.map(([r, cc]) => (g[r] || [])[cc]).filter(z => typeof z === 'number'); return v.length ? Math.min.apply(null, v) : 0; });
function subtotalFn(fnum, vals, calc) {
  const k = fnum % 100; const n = vals;
  switch (k) {
    case 1: if (!n.length) throw E.DIV0; return sum(n) / n.length;
    case 2: return n.length; case 3: return n.length;
    case 4: return n.length ? Math.max.apply(null, n) : 0; case 5: return n.length ? Math.min.apply(null, n) : 0;
    case 6: return n.reduce((p, v) => p * v, 1); case 7: return Math.sqrt(variance(n, false)); case 8: return Math.sqrt(variance(n, true));
    case 9: return fixf(sum(n)); case 10: return variance(n, false); case 11: return variance(n, true);
    case 12: { const s = sortedNums(n); const m = s.length >> 1; if (!s.length) throw E.NUM; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; }
    case 13: { const cnt = new Map(); let best = null, bc = 1; n.forEach(v => { const q = (cnt.get(v) || 0) + 1; cnt.set(v, q); if (q > bc) { bc = q; best = v; } }); if (best === null) throw E.NA; return best; }
  }
  throw E.VALUE;
}
def('SUBTOTAL', MATH, 'function_num, ref1, [ref2], ...', 'Tổng phụ (bỏ qua hàng ẩn khi mã 101-111).', (a, x, c) => {
  const f = toInt(a[0]); const vals = [];
  for (const r of a.slice(1)) {
    if (r instanceof Ref) c.eachIn(r, (v, rr, cc) => {
      if (f > 100 && r.sheet.hiddenR[rr]) return;
      const cell = r.sheet.get(rr, cc); if (cell && cell.f && /SUBTOTAL\(|AGGREGATE\(/i.test(cell.f)) return;
      if (f % 100 === 3) { if (v != null && v !== '') vals.push(1); } else if (f % 100 === 2) { if (typeof v === 'number') vals.push(v); } else if (typeof v === 'number') vals.push(v); else if (v instanceof Err) throw v;
    }); else nums(c, [r]).forEach(v => vals.push(v));
  }
  return subtotalFn(f, vals, c);
});
def('AGGREGATE', MATH, 'function_num, options, ref1, [ref2], ...', 'Tổng hợp với tùy chọn bỏ qua lỗi / hàng ẩn.', (a, x, c) => {
  const f = toInt(a[0]), opt = toInt(a[1]); const ignoreErr = [2, 3, 6, 7].indexOf(opt) >= 0, ignoreHidden = [1, 3, 5, 7].indexOf(opt) >= 0;
  const vals = [];
  for (const r of (f >= 14 ? [a[2]] : a.slice(2))) {
    if (r instanceof Ref) c.eachIn(r, (v, rr) => { if (ignoreHidden && r.sheet.hiddenR[rr]) return; if (v instanceof Err) { if (ignoreErr) return; throw v; } if (typeof v === 'number') vals.push(v); });
    else flatVals(c, r).forEach(v => { if (v instanceof Err) { if (ignoreErr) return; throw v; } if (typeof v === 'number') vals.push(v); });
  }
  if (f <= 13) return subtotalFn(f === 2 || f === 3 ? (f === 3 ? 3 : 2) : f, vals, c);
  const k = N(scalar(c, a[3]));
  switch (f) { case 14: { const s = sortedNums(vals).reverse(); if (k < 1 || k > s.length) throw E.NUM; return s[k - 1]; } case 15: { const s = sortedNums(vals); if (k < 1 || k > s.length) throw E.NUM; return s[k - 1]; } case 16: return percentileInc(vals, k); case 17: return percentileInc(vals, k / 4); case 18: return percentileExc(vals, k); case 19: return percentileExc(vals, k / 4); }
  throw E.VALUE;
});
def('MMULT', MATH, 'array1, array2', 'Tích ma trận.', (a, x, c) => { const A = arr2(c, a[0]), Bm = arr2(c, a[1]); if (A[0].length !== Bm.length) throw E.VALUE; return A.map(row => Bm[0].map((_, j) => row.reduce((s, v, k) => { const p = N(v), q = N(Bm[k][j]); return s + p * q; }, 0))); });
def('MDETERM', MATH, 'array', 'Định thức ma trận.', (a, x, c) => { const M = arr2(c, a[0]).map(r => r.map(N)); const n = M.length; if (n !== M[0].length) throw E.VALUE; let det = 1; for (let i = 0; i < n; i++) { let p = i; for (let r = i + 1; r < n; r++) if (Math.abs(M[r][i]) > Math.abs(M[p][i])) p = r; if (!M[p][i]) return 0; if (p !== i) { const t = M[p]; M[p] = M[i]; M[i] = t; det = -det; } det *= M[i][i]; for (let r = i + 1; r < n; r++) { const f = M[r][i] / M[i][i]; for (let k = i; k < n; k++) M[r][k] -= f * M[i][k]; } } return +det.toPrecision(15); });
def('MINVERSE', MATH, 'array', 'Ma trận nghịch đảo.', (a, x, c) => { const M = arr2(c, a[0]).map(r => r.map(N)); const n = M.length; if (n !== M[0].length) throw E.VALUE; const A = M.map((r, i) => r.concat(Array.from({ length: n }, (_, j) => i === j ? 1 : 0))); for (let i = 0; i < n; i++) { let p = i; for (let r = i + 1; r < n; r++) if (Math.abs(A[r][i]) > Math.abs(A[p][i])) p = r; if (Math.abs(A[p][i]) < 1e-14) throw E.NUM; const t = A[p]; A[p] = A[i]; A[i] = t; const d = A[i][i]; for (let k = 0; k < 2 * n; k++) A[i][k] /= d; for (let r = 0; r < n; r++) if (r !== i) { const f = A[r][i]; for (let k = 0; k < 2 * n; k++) A[r][k] -= f * A[i][k]; } } return A.map(r => r.slice(n).map(v => +v.toPrecision(15))); });
def('MUNIT', MATH, 'dimension', 'Ma trận đơn vị.', a => { const n = toInt(a[0]); if (n < 1) throw E.VALUE; return Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => i === j ? 1 : 0)); });
def('RADIANS', MATH, 'angle', 'Đổi độ sang radian.', a => N(a[0]) * Math.PI / 180, SC);
def('DEGREES', MATH, 'angle', 'Đổi radian sang độ.', a => N(a[0]) * 180 / Math.PI, SC);
['SIN', 'COS', 'TAN', 'ASIN', 'ACOS', 'ATAN'].forEach(fn => def(fn, MATH, 'number', { SIN: 'Sin của một góc (radian).', COS: 'Cosin của một góc.', TAN: 'Tang của một góc.', ASIN: 'Arcsin.', ACOS: 'Arccos.', ATAN: 'Arctang.' }[fn], a => { const r = Math[fn.toLowerCase()](N(a[0])); if (isNaN(r)) throw E.NUM; return r; }, SC));
def('ATAN2', MATH, 'x_num, y_num', 'Arctang từ tọa độ x, y.', a => { const x = N(a[0]), y = N(a[1]); if (!x && !y) throw E.DIV0; return Math.atan2(y, x); }, SC);

// ================================================================ STATISTICS (native core)
const stdev = pop => (a, x, c) => Math.sqrt(variance(nums(c, a), pop));
def('STDEV', COMPAT, 'number1, [number2], ...', 'Độ lệch chuẩn mẫu.', stdev(false)); def('STDEV.S', STAT, 'number1, [number2], ...', 'Độ lệch chuẩn dựa trên mẫu.', stdev(false));
def('STDEVP', COMPAT, 'number1, [number2], ...', 'Độ lệch chuẩn tổng thể.', stdev(true)); def('STDEV.P', STAT, 'number1, [number2], ...', 'Độ lệch chuẩn tổng thể.', stdev(true));
def('STDEVA', STAT, 'value1, [value2], ...', 'Độ lệch chuẩn mẫu (tính cả văn bản/logic).', (a, x, c) => Math.sqrt(variance(nums(c, a, { a: true }), false)));
def('STDEVPA', STAT, 'value1, [value2], ...', 'Độ lệch chuẩn tổng thể (tính cả văn bản/logic).', (a, x, c) => Math.sqrt(variance(nums(c, a, { a: true }), true)));
const vr = pop => (a, x, c) => variance(nums(c, a), pop);
def('VAR', COMPAT, 'number1, [number2], ...', 'Phương sai mẫu.', vr(false)); def('VAR.S', STAT, 'number1, [number2], ...', 'Phương sai mẫu.', vr(false));
def('VARP', COMPAT, 'number1, [number2], ...', 'Phương sai tổng thể.', vr(true)); def('VAR.P', STAT, 'number1, [number2], ...', 'Phương sai tổng thể.', vr(true));
def('VARA', STAT, 'value1, [value2], ...', 'Phương sai mẫu (tính cả văn bản/logic).', (a, x, c) => variance(nums(c, a, { a: true }), false));
def('VARPA', STAT, 'value1, [value2], ...', 'Phương sai tổng thể (tính cả văn bản/logic).', (a, x, c) => variance(nums(c, a, { a: true }), true));
def('LARGE', STAT, 'array, k', 'Giá trị lớn thứ k.', (a, x, c) => { const s = sortedNums(nums(c, [a[0]])).reverse(); const k = Math.ceil(N(scalar(c, a[1]))); if (k < 1 || k > s.length) throw E.NUM; return s[k - 1]; });
def('SMALL', STAT, 'array, k', 'Giá trị nhỏ thứ k.', (a, x, c) => { const s = sortedNums(nums(c, [a[0]])); const k = Math.ceil(N(scalar(c, a[1]))); if (k < 1 || k > s.length) throw E.NUM; return s[k - 1]; });
const rank = avg => (a, x, c) => { const v = N(scalar(c, a[0])); const list = nums(c, [a[1]]); const asc = a[2] != null && N(scalar(c, a[2])) !== 0; if (list.indexOf(v) < 0) throw E.NA; const better = list.filter(z => asc ? z < v : z > v).length; if (!avg) return better + 1; const eq = list.filter(z => z === v).length; return better + (eq + 1) / 2; };
def('RANK', COMPAT, 'number, ref, [order]', 'Thứ hạng của số trong danh sách.', rank(false)); def('RANK.EQ', STAT, 'number, ref, [order]', 'Thứ hạng (bằng nhau cùng hạng cao nhất).', rank(false)); def('RANK.AVG', STAT, 'number, ref, [order]', 'Thứ hạng (bằng nhau lấy trung bình).', rank(true));
def('PERCENTILE', COMPAT, 'array, k', 'Phân vị thứ k.', (a, x, c) => percentileInc(nums(c, [a[0]]), N(scalar(c, a[1]))));
def('PERCENTILE.INC', STAT, 'array, k', 'Phân vị thứ k (bao gồm 0, 1).', (a, x, c) => percentileInc(nums(c, [a[0]]), N(scalar(c, a[1]))));
def('PERCENTILE.EXC', STAT, 'array, k', 'Phân vị thứ k (loại trừ 0, 1).', (a, x, c) => percentileExc(nums(c, [a[0]]), N(scalar(c, a[1]))));
def('QUARTILE', COMPAT, 'array, quart', 'Tứ phân vị.', (a, x, c) => { const q = toInt(scalar(c, a[1])); if (q < 0 || q > 4) throw E.NUM; return percentileInc(nums(c, [a[0]]), q / 4); });
def('QUARTILE.INC', STAT, 'array, quart', 'Tứ phân vị (bao gồm).', (a, x, c) => { const q = toInt(scalar(c, a[1])); if (q < 0 || q > 4) throw E.NUM; return percentileInc(nums(c, [a[0]]), q / 4); });
def('QUARTILE.EXC', STAT, 'array, quart', 'Tứ phân vị (loại trừ).', (a, x, c) => { const q = toInt(scalar(c, a[1])); if (q < 1 || q > 3) throw E.NUM; return percentileExc(nums(c, [a[0]]), q / 4); });
def('COUNTUNIQUE', STAT, 'value1, ...', 'Đếm số giá trị duy nhất (mở rộng).', (a, x, c) => { const s = new Set(); a.forEach(v => flatVals(c, v).forEach(z => { if (z != null && z !== '') s.add(typeof z === 'string' ? z.toLowerCase() : z); })); return s.size; });

// ================================================================ LOGICAL
def('IF', LOGIC, 'logical_test, [value_if_true], [value_if_false]', 'Trả về giá trị tùy theo điều kiện đúng/sai.', (args, ctx, calc) => {
  let t = calc.eval(args[0], ctx);
  if (t instanceof Ref) t = calc.deref(t, true);
  if (isArr(t)) return t.map(row => row.map(v => { const b = bool(v); if (b instanceof Err) return b; const br = b ? args[1] : args[2]; if (!br || br.t === 'missing') return b ? (args[1] ? 0 : true) : false; const r = calc.eval(br, ctx); return r instanceof Ref ? calc.deref(r, true) : r; }));
  const b = bool(t); if (b instanceof Err) throw b;
  if (b) return args.length > 1 ? (args[1].t === 'missing' ? 0 : calc.eval(args[1], ctx)) : true;
  return args.length > 2 ? (args[2].t === 'missing' ? 0 : calc.eval(args[2], ctx)) : false;
}, { special: true });
def('IFS', LOGIC, 'logical_test1, value_if_true1, ...', 'Kiểm tra nhiều điều kiện, trả về giá trị của điều kiện đúng đầu tiên.', (args, ctx, calc) => { for (let i = 0; i + 1 < args.length; i += 2) { const b = bool(scalar(calc, calc.eval(args[i], ctx))); if (b instanceof Err) throw b; if (b) return calc.eval(args[i + 1], ctx); } throw E.NA; }, { special: true });
def('IFERROR', LOGIC, 'value, value_if_error', 'Trả về giá trị khác nếu biểu thức lỗi.', (args, ctx, calc) => { let v; try { v = calc.eval(args[0], ctx); if (v instanceof Ref) v = calc.deref(v, true); } catch (e) { if (e instanceof Err) return calc.eval(args[1], ctx); throw e; } if (isArr(v)) { return v.map(r => r.map(z => z instanceof Err ? scalar(calc, calc.eval(args[1], ctx)) : z)); } return v instanceof Err ? calc.eval(args[1], ctx) : v; }, { special: true });
def('IFNA', LOGIC, 'value, value_if_na', 'Trả về giá trị khác nếu biểu thức là #N/A.', (args, ctx, calc) => { let v; try { v = calc.eval(args[0], ctx); if (v instanceof Ref) v = calc.deref(v, true); } catch (e) { if (e instanceof Err && e.code === '#N/A') return calc.eval(args[1], ctx); throw e; } return v instanceof Err && v.code === '#N/A' ? calc.eval(args[1], ctx) : v; }, { special: true });
function logicVals(c, a) { const out = []; for (const v of a) { if (v instanceof Ref || isArr(v)) flatVals(c, v).forEach(z => { if (z instanceof Err) throw z; if (typeof z === 'boolean') out.push(z); else if (typeof z === 'number') out.push(z !== 0); }); else { const b = bool(v); if (b instanceof Err) throw b; out.push(b); } } if (!out.length) throw E.VALUE; return out; }
def('AND', LOGIC, 'logical1, [logical2], ...', 'TRUE nếu tất cả đối số đúng.', (a, x, c) => logicVals(c, a).every(Boolean));
def('OR', LOGIC, 'logical1, [logical2], ...', 'TRUE nếu bất kỳ đối số nào đúng.', (a, x, c) => logicVals(c, a).some(Boolean));
def('XOR', LOGIC, 'logical1, [logical2], ...', 'Phép XOR logic.', (a, x, c) => logicVals(c, a).filter(Boolean).length % 2 === 1);
def('NOT', LOGIC, 'logical', 'Đảo giá trị logic.', a => !B(a[0]), SC);
def('TRUE', LOGIC, '', 'Giá trị TRUE.', () => true); def('FALSE', LOGIC, '', 'Giá trị FALSE.', () => false);
def('SWITCH', LOGIC, 'expression, value1, result1, ..., [default]', 'So sánh biểu thức với danh sách giá trị.', (args, ctx, calc) => { const v = scalar(calc, calc.eval(args[0], ctx)); for (let i = 1; i + 1 < args.length; i += 2) { const k = scalar(calc, calc.eval(args[i], ctx)); if (cmp(v, k) === 0 && typeof v === typeof k) return calc.eval(args[i + 1], ctx); } if (args.length % 2 === 0) return calc.eval(args[args.length - 1], ctx); throw E.NA; }, { special: true });
def('CHOOSE', LOOK, 'index_num, value1, [value2], ...', 'Chọn một giá trị theo chỉ số.', (args, ctx, calc) => { const i = Math.floor(N(scalar(calc, calc.eval(args[0], ctx)))); if (i < 1 || i >= args.length) throw E.VALUE; return calc.eval(args[i], ctx); }, { special: true });
def('LET', LOGIC, 'name1, value1, [name2, value2], ..., calculation', 'Gán tên cho kết quả tính toán.', (args, ctx, calc) => { const vars = {}; const env = { vars, parent: ctx.env }; const c2 = Object.assign({}, ctx, { env }); for (let i = 0; i + 1 < args.length; i += 2) { const nm = args[i]; if (nm.t !== 'name') throw E.VALUE; vars[nm.n.toUpperCase()] = calc.eval(args[i + 1], c2); } return calc.eval(args[args.length - 1], c2); }, { special: true });
def('LAMBDA', LOGIC, '[parameter1, ...], calculation', 'Tạo hàm tùy chỉnh.', (args, ctx) => { const params = args.slice(0, -1).map(a => { if (a.t !== 'name') throw E.VALUE; return a.n.toUpperCase(); }); return new Lambda(params, args[args.length - 1], ctx.env); }, { special: true });
def('ISOMITTED', INFO, 'argument', 'Kiểm tra tham số LAMBDA có bị bỏ qua.', a => a[0] == null);
function needLam(v) { if (!(v instanceof Lambda)) throw E.VALUE; return v; }
def('MAP', LOGIC, 'array1, ..., lambda', 'Áp dụng LAMBDA cho từng phần tử.', (a, ctx, c) => { const lam = needLam(a[a.length - 1]); const arrs = a.slice(0, -1).map(v => arr2(c, v)); return arrs[0].map((row, i) => row.map((_, j) => scalar(c, c.applyLambda(lam, arrs.map(m => m[i][j]), ctx)))); });
def('REDUCE', LOGIC, '[initial_value], array, lambda', 'Thu gọn mảng thành một giá trị.', (a, ctx, c) => { const lam = needLam(a[2]); let acc = a[0]; flatVals(c, a[1]).forEach(v => { acc = scalar(c, c.applyLambda(lam, [acc, v], ctx)); }); return acc; });
def('SCAN', LOGIC, '[initial_value], array, lambda', 'Quét mảng, trả về các giá trị trung gian.', (a, ctx, c) => { const lam = needLam(a[2]); let acc = a[0]; const m = arr2(c, a[1]); return m.map(row => row.map(v => (acc = scalar(c, c.applyLambda(lam, [acc, v], ctx))))); });
def('BYROW', LOGIC, 'array, lambda', 'Áp dụng LAMBDA cho từng hàng.', (a, ctx, c) => { const lam = needLam(a[1]); return arr2(c, a[0]).map(row => [scalar(c, c.applyLambda(lam, [[row]], ctx))]); });
def('BYCOL', LOGIC, 'array, lambda', 'Áp dụng LAMBDA cho từng cột.', (a, ctx, c) => { const lam = needLam(a[1]); const m = arr2(c, a[0]); return [m[0].map((_, j) => scalar(c, c.applyLambda(lam, [m.map(r => [r[j]])], ctx)))]; });
def('MAKEARRAY', LOGIC, 'rows, cols, lambda', 'Tạo mảng bằng LAMBDA.', (a, ctx, c) => { const lam = needLam(a[2]); const R = toInt(a[0]), C = toInt(a[1]); const out = []; for (let i = 1; i <= R; i++) { const row = []; for (let j = 1; j <= C; j++) row.push(scalar(c, c.applyLambda(lam, [i, j], ctx))); out.push(row); } return out; });

// ================================================================ TEXT
const T1 = v => S(v);
def('LEN', TEXT, 'text', 'Số ký tự của chuỗi.', a => T1(a[0]).length, SC); FN.LENB = FN.LEN; meta.LENB = [TEXT, 'text', 'Số byte của chuỗi.'];
def('LEFT', TEXT, 'text, [num_chars]', 'Lấy ký tự bên trái.', a => { const n = a[1] == null ? 1 : toInt(a[1]); if (n < 0) throw E.VALUE; return T1(a[0]).slice(0, n); }, SC); FN.LEFTB = FN.LEFT; meta.LEFTB = [TEXT, 'text, [num_bytes]', 'Lấy byte bên trái.'];
def('RIGHT', TEXT, 'text, [num_chars]', 'Lấy ký tự bên phải.', a => { const s = T1(a[0]); const n = a[1] == null ? 1 : toInt(a[1]); if (n < 0) throw E.VALUE; return n ? s.slice(-n) : ''; }, SC); FN.RIGHTB = FN.RIGHT; meta.RIGHTB = [TEXT, 'text, [num_bytes]', 'Lấy byte bên phải.'];
def('MID', TEXT, 'text, start_num, num_chars', 'Lấy ký tự ở giữa chuỗi.', a => { const st = toInt(a[1]), n = toInt(a[2]); if (st < 1 || n < 0) throw E.VALUE; return T1(a[0]).substr(st - 1, n); }, SC); FN.MIDB = FN.MID; meta.MIDB = [TEXT, 'text, start_num, num_bytes', 'Lấy byte ở giữa chuỗi.'];
def('FIND', TEXT, 'find_text, within_text, [start_num]', 'Tìm vị trí chuỗi (phân biệt hoa thường).', a => { const st = a[2] == null ? 1 : toInt(a[2]); const s = T1(a[1]); if (st < 1 || st > s.length + 1) throw E.VALUE; const i = s.indexOf(T1(a[0]), st - 1); if (i < 0) throw E.VALUE; return i + 1; }, SC); FN.FINDB = FN.FIND; meta.FINDB = meta.FIND;
def('SEARCH', TEXT, 'find_text, within_text, [start_num]', 'Tìm vị trí chuỗi (không phân biệt hoa thường, hỗ trợ * ?).', a => { const st = a[2] == null ? 1 : toInt(a[2]); const s = T1(a[1]); const re = new RegExp(wild(T1(a[0])).source.slice(1, -1), 'i'); const m = re.exec(s.slice(st - 1)); if (!m || st < 1) throw E.VALUE; return m.index + st; }, SC); FN.SEARCHB = FN.SEARCH; meta.SEARCHB = meta.SEARCH;
def('REPLACE', TEXT, 'old_text, start_num, num_chars, new_text', 'Thay thế một phần chuỗi theo vị trí.', a => { const s = T1(a[0]); const st = toInt(a[1]), n = toInt(a[2]); if (st < 1 || n < 0) throw E.VALUE; return s.slice(0, st - 1) + T1(a[3]) + s.slice(st - 1 + n); }, SC); FN.REPLACEB = FN.REPLACE; meta.REPLACEB = meta.REPLACE;
def('SUBSTITUTE', TEXT, 'text, old_text, new_text, [instance_num]', 'Thay thế chuỗi con.', a => { const s = T1(a[0]), o = T1(a[1]), n = T1(a[2]); if (!o) return s; if (a[3] == null) return s.split(o).join(n); const k = toInt(a[3]); if (k < 1) throw E.VALUE; let idx = -1; for (let i = 0; i < k; i++) { idx = s.indexOf(o, idx + 1); if (idx < 0) return s; } return s.slice(0, idx) + n + s.slice(idx + o.length); }, SC);
def('UPPER', TEXT, 'text', 'Chuyển thành chữ hoa.', a => T1(a[0]).toUpperCase(), SC);
def('LOWER', TEXT, 'text', 'Chuyển thành chữ thường.', a => T1(a[0]).toLowerCase(), SC);
def('PROPER', TEXT, 'text', 'Viết hoa chữ cái đầu mỗi từ.', a => T1(a[0]).toLowerCase().replace(/(^|[^a-zA-Z\u00C0-\u1EF9'])([a-z\u00C0-\u1EF9])/g, (m, p, c) => p + c.toUpperCase()), SC);
def('TRIM', TEXT, 'text', 'Xóa khoảng trắng thừa.', a => T1(a[0]).replace(/ +/g, ' ').replace(/^ | $/g, ''), SC);
def('CLEAN', TEXT, 'text', 'Xóa ký tự không in được.', a => T1(a[0]).replace(/[\x00-\x1f]/g, ''), SC);
def('CONCAT', TEXT, 'text1, [text2], ...', 'Nối chuỗi (cả phạm vi).', (a, x, c) => { let s = ''; for (const v of a) flatVals(c, v).forEach(z => { if (z instanceof Err) throw z; s += str(z); }); return s; });
def('CONCATENATE', TEXT, 'text1, [text2], ...', 'Nối các chuỗi.', a => a.map(v => S(v)).join(''), { scalar: true });
def('TEXTJOIN', TEXT, 'delimiter, ignore_empty, text1, ...', 'Nối chuỗi với dấu phân cách.', (a, x, c) => { const d = S(scalar(c, a[0])), ig = B(scalar(c, a[1])); const parts = []; for (const v of a.slice(2)) flatVals(c, v).forEach(z => { if (z instanceof Err) throw z; const s = str(z); if (!ig || s !== '') parts.push(s); }); return parts.join(d); });
def('TEXT', TEXT, 'value, format_text', 'Định dạng số thành văn bản.', a => { let v = a[0]; if (typeof v === 'string') { const p = parseInput(v); if (typeof p.v === 'number') v = p.v; } return format(v, S(a[1])).t; }, SC);
def('VALUE', TEXT, 'text', 'Chuyển văn bản thành số.', a => { if (typeof a[0] === 'number') return a[0]; const p = parseInput(S(a[0])); if (typeof p.v !== 'number') throw E.VALUE; return p.v; }, SC);
def('NUMBERVALUE', TEXT, 'text, [decimal_separator], [group_separator]', 'Chuyển văn bản thành số theo dấu phân cách.', a => { let s = S(a[0]); const d = a[1] == null ? '.' : S(a[1]), g = a[2] == null ? ',' : S(a[2]); s = s.split(g).join('').replace(d, '.').replace(/\s/g, ''); let pct = 0; while (/%$/.test(s)) { s = s.slice(0, -1); pct++; } const n = +s; if (s === '' || isNaN(n)) throw E.VALUE; return n / Math.pow(100, pct); }, SC);
def('FIXED', TEXT, 'number, [decimals], [no_commas]', 'Làm tròn và định dạng số thành văn bản.', a => { const d = a[1] == null ? 2 : toInt(a[1]); const nc = a[2] != null && B(a[2]); const n = roundX(N(a[0]), d, 0); return format(n, (nc ? '0' : '#,##0') + (d > 0 ? '.' + '0'.repeat(d) : '')).t; }, SC);
def('DOLLAR', TEXT, 'number, [decimals]', 'Định dạng tiền tệ thành văn bản.', a => { const d = a[1] == null ? 2 : toInt(a[1]); return format(roundX(N(a[0]), d, 0), '"$"#,##0' + (d > 0 ? '.' + '0'.repeat(d) : '') + ';("$"#,##0' + (d > 0 ? '.' + '0'.repeat(d) : '') + ')').t; }, SC);
def('REPT', TEXT, 'text, number_times', 'Lặp lại chuỗi.', a => { const n = toInt(a[1]); if (n < 0) throw E.VALUE; return T1(a[0]).repeat(n); }, SC);
def('EXACT', TEXT, 'text1, text2', 'So sánh chính xác hai chuỗi.', a => T1(a[0]) === T1(a[1]), SC);
def('CHAR', TEXT, 'number', 'Ký tự theo mã.', a => { const n = toInt(a[0]); if (n < 1 || n > 255) throw E.VALUE; return String.fromCharCode(n); }, SC);
def('CODE', TEXT, 'text', 'Mã của ký tự đầu tiên.', a => { const s = T1(a[0]); if (!s) throw E.VALUE; return s.charCodeAt(0); }, SC);
def('UNICHAR', TEXT, 'number', 'Ký tự Unicode theo mã.', a => { const n = toInt(a[0]); if (n < 1) throw E.VALUE; return String.fromCodePoint ? String.fromCodePoint(n) : String.fromCharCode(n); }, SC);
def('UNICODE', TEXT, 'text', 'Mã Unicode của ký tự đầu.', a => { const s = T1(a[0]); if (!s) throw E.VALUE; return s.codePointAt ? s.codePointAt(0) : s.charCodeAt(0); }, SC);
def('T', TEXT, 'value', 'Trả về văn bản nếu là văn bản.', a => typeof a[0] === 'string' ? a[0] : '', SC);
def('N', INFO, 'value', 'Chuyển giá trị thành số.', a => typeof a[0] === 'number' ? a[0] : a[0] === true ? 1 : a[0] instanceof Err ? (() => { throw a[0]; })() : 0, SC);
def('TEXTBEFORE', TEXT, 'text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]', 'Văn bản trước dấu phân cách.', a => textAround(a, true), SC);
def('TEXTAFTER', TEXT, 'text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]', 'Văn bản sau dấu phân cách.', a => textAround(a, false), SC);
function textAround(a, before) {
  let s = T1(a[0]), d = T1(a[1]); const k = a[2] == null ? 1 : toInt(a[2]); const ci = a[3] != null && N(a[3]) === 1; const nf = a[5];
  if (k === 0 || Math.abs(k) > s.length + 1) throw E.VALUE;
  const S2 = ci ? s.toLowerCase() : s, D2 = ci ? d.toLowerCase() : d;
  let idx = -1;
  if (k > 0) { let p = -1; for (let i = 0; i < k; i++) { p = S2.indexOf(D2, p + (i ? D2.length || 1 : 0)); if (p < 0) break; } idx = p; if (idx < 0 && a[4] != null && N(a[4]) === 1 && k === 1) idx = s.length; }
  else { let p = S2.length + 1; for (let i = 0; i < -k; i++) { p = S2.lastIndexOf(D2, p - 1 - (i ? 0 : 0)); if (p < 0) break; if (i < -k - 1) p = p; } idx = p; }
  if (idx < 0) { if (nf != null) return nf; throw E.NA; }
  return before ? s.slice(0, idx) : s.slice(idx + d.length);
}
def('TEXTSPLIT', TEXT, 'text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]', 'Tách văn bản thành mảng.', (a, x, c) => {
  const s = S(scalar(c, a[0])); const cd = flatVals(c, a[1]).map(str); const rd = a[2] == null ? [] : flatVals(c, a[2]).map(str); const ig = a[3] != null && B(scalar(c, a[3])); const pad = a[5] === undefined ? E.NA : a[5];
  const esc2 = z => z.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const rows = (rd.length ? s.split(new RegExp(rd.map(esc2).join('|'))) : [s]).map(r => { let p = cd.filter(Boolean).length ? r.split(new RegExp(cd.filter(Boolean).map(esc2).join('|'))) : [r]; if (ig) p = p.filter(z => z !== ''); return p; });
  const w = Math.max.apply(null, rows.map(r => r.length));
  return rows.map(r => { while (r.length < w) r.push(pad); return r; });
});
def('VALUETOTEXT', TEXT, 'value, [format]', 'Chuyển giá trị thành văn bản.', a => { const strict = a[1] != null && N(a[1]) === 1; const v = a[0]; if (typeof v === 'string') return strict ? '"' + v + '"' : v; return str(v instanceof Err ? v.code : v); }, SC);
def('ARRAYTOTEXT', TEXT, 'array, [format]', 'Chuyển mảng thành văn bản.', (a, x, c) => { const strict = a[1] != null && N(scalar(c, a[1])) === 1; const m = arr2(c, a[0]); const f = v => v instanceof Err ? v.code : typeof v === 'string' && strict ? '"' + v + '"' : str(v); return strict ? '{' + m.map(r => r.map(f).join(',')).join(';') + '}' : m.map(r => r.map(f).join(', ')).join(', '); });
def('REGEXTEST', TEXT, 'text, pattern, [case_sensitivity]', 'Kiểm tra văn bản khớp biểu thức chính quy.', a => { try { return new RegExp(T1(a[1]), a[2] != null && N(a[2]) === 1 ? 'i' : '').test(T1(a[0])); } catch (e) { throw E.VALUE; } }, SC);
def('REGEXEXTRACT', TEXT, 'text, pattern, [return_mode], [case_sensitivity]', 'Trích xuất văn bản theo biểu thức chính quy.', (a, x, c) => { const s = S(scalar(c, a[0])), p = S(scalar(c, a[1])); const mode = a[2] == null ? 0 : N(scalar(c, a[2])); const fl = a[3] != null && N(scalar(c, a[3])) === 1 ? 'i' : ''; let re; try { re = new RegExp(p, fl + (mode === 1 ? 'g' : '')); } catch (e) { throw E.VALUE; } if (mode === 1) { const m = s.match(re); if (!m) throw E.NA; return m.map(z => [z]); } const m = re.exec(s); if (!m) throw E.NA; if (mode === 2) return [m.slice(1).map(z => z == null ? '' : z)]; return m[0]; });
def('REGEXREPLACE', TEXT, 'text, pattern, replacement, [occurrence], [case_sensitivity]', 'Thay thế văn bản theo biểu thức chính quy.', a => { const s = T1(a[0]); const occ = a[3] == null ? 0 : toInt(a[3]); const fl = a[4] != null && N(a[4]) === 1 ? 'i' : ''; let re; try { re = new RegExp(T1(a[1]), fl + 'g'); } catch (e) { throw E.VALUE; } if (!occ) return s.replace(re, T1(a[2])); let i = 0; return s.replace(re, m => (++i === occ ? T1(a[2]) : m)); }, SC);
def('ASC', TEXT, 'text', 'Chuyển ký tự toàn độ sang bán độ.', a => T1(a[0]).replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)).replace(/\u3000/g, ' '), SC);
def('DBCS', TEXT, 'text', 'Chuyển ký tự bán độ sang toàn độ.', a => T1(a[0]).replace(/[!-~]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0xFEE0)), SC); FN.JIS = FN.DBCS; meta.JIS = meta.DBCS;
def('PHONETIC', TEXT, 'reference', 'Trích ký tự phiên âm (trả về chính văn bản).', (a, x, c) => str(scalar(c, a[0])));
def('ENCODEURL', WEB, 'text', 'Mã hóa chuỗi theo URL.', a => encodeURIComponent(T1(a[0])), SC);

// ================================================================ DATE & TIME
const P = n => serialToParts(Math.floor(N(n)));
def('DATE', DATE, 'year, month, day', 'Tạo giá trị ngày.', a => { let y = toInt(a[0]), m = toInt(a[1]), d = toInt(a[2]); if (y < 1900) y += 1900; if (y > 9999) throw E.NUM; const v = (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5; if (v < 0) throw E.NUM; return v; }, SC);
def('TIME', DATE, 'hour, minute, second', 'Tạo giá trị thời gian.', a => { const v = (toInt(a[0]) * 3600 + toInt(a[1]) * 60 + toInt(a[2])) / 86400; if (v < 0) throw E.NUM; return v - Math.floor(v); }, SC);
def('YEAR', DATE, 'serial_number', 'Năm của ngày.', a => P(a[0]).y, SC);
def('MONTH', DATE, 'serial_number', 'Tháng của ngày.', a => P(a[0]).m, SC);
def('DAY', DATE, 'serial_number', 'Ngày trong tháng.', a => P(a[0]).d, SC);
def('HOUR', DATE, 'serial_number', 'Giờ.', a => serialToParts(N(a[0])).H, SC);
def('MINUTE', DATE, 'serial_number', 'Phút.', a => serialToParts(N(a[0])).M, SC);
def('SECOND', DATE, 'serial_number', 'Giây.', a => serialToParts(N(a[0])).S, SC);
def('TODAY', DATE, '', 'Ngày hôm nay.', () => { const d = new Date(); return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); }, { volatile: true });
def('NOW', DATE, '', 'Ngày giờ hiện tại.', () => jsDateToSerial(new Date()), { volatile: true });
def('WEEKDAY', DATE, 'serial_number, [return_type]', 'Thứ trong tuần.', a => { const dow = P(a[0]).dow; const t = a[1] == null ? 1 : toInt(a[1]); switch (t) { case 1: case 17: return dow + 1; case 2: case 11: return dow === 0 ? 7 : dow; case 3: return dow === 0 ? 6 : dow - 1; case 12: return (dow + 5) % 7 + 1; case 13: return (dow + 4) % 7 + 1; case 14: return (dow + 3) % 7 + 1; case 15: return (dow + 2) % 7 + 1; case 16: return (dow + 1) % 7 + 1; } throw E.NUM; }, SC);
def('WEEKNUM', DATE, 'serial_number, [return_type]', 'Số tuần trong năm.', a => { const n = Math.floor(N(a[0])); const t = a[1] == null ? 1 : toInt(a[1]); if (t === 21) return isoWeek(n); const p = serialToParts(n); const jan1 = dateToSerial(p.y, 1, 1); const startDow = { 1: 0, 17: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6 }[t]; if (startDow === undefined) throw E.NUM; const d1 = serialToParts(jan1).dow; const off = (d1 - startDow + 7) % 7; return Math.floor((n - jan1 + off) / 7) + 1; }, SC);
function isoWeek(n) { const p = serialToParts(n); const dow = (p.dow + 6) % 7; const th = n - dow + 3; const ty = serialToParts(th).y; const j = dateToSerial(ty, 1, 1); return Math.floor((th - j) / 7) + 1; }
def('ISOWEEKNUM', DATE, 'date', 'Số tuần theo ISO.', a => isoWeek(Math.floor(N(a[0]))), SC);
const addMonths = (n, k) => { const p = serialToParts(n); const t = new Date(Date.UTC(p.y, p.m - 1 + k, 1)); const last = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0)).getUTCDate(); return dateToSerial(t.getUTCFullYear(), t.getUTCMonth() + 1, Math.min(p.d, last)); };
def('EDATE', DATE, 'start_date, months', 'Ngày cách n tháng.', a => addMonths(Math.floor(N(a[0])), toInt(a[1])), SC);
def('EOMONTH', DATE, 'start_date, months', 'Ngày cuối tháng sau n tháng.', a => { const p = serialToParts(Math.floor(N(a[0]))); const t = new Date(Date.UTC(p.y, p.m + toInt(a[1]), 0)); return dateToSerial(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); }, SC);
def('DATEDIF', DATE, 'start_date, end_date, unit', 'Khoảng cách giữa hai ngày (Y, M, D, MD, YM, YD).', a => {
  const s = Math.floor(N(a[0])), e = Math.floor(N(a[1])); if (s > e) throw E.NUM; const u = T1(a[2]).toUpperCase(); const p = serialToParts(s), q = serialToParts(e);
  let months = (q.y - p.y) * 12 + q.m - p.m; if (q.d < p.d) months--;
  switch (u) { case 'Y': return Math.floor(months / 12); case 'M': return months; case 'D': return e - s; case 'MD': { let d = q.d - p.d; if (d < 0) { const lm = new Date(Date.UTC(q.y, q.m - 1, 0)).getUTCDate(); d += lm; } return d; } case 'YM': return months % 12; case 'YD': { let ys = dateToSerial(q.y, p.m, p.d); if (ys > e) ys = dateToSerial(q.y - 1, p.m, p.d); return e - ys; } }
  throw E.NUM;
}, SC);
def('DAYS', DATE, 'end_date, start_date', 'Số ngày giữa hai ngày.', a => Math.floor(N(a[0])) - Math.floor(N(a[1])), SC);
def('DATEVALUE', DATE, 'date_text', 'Chuyển chuỗi ngày thành số seri.', a => { const p = parseInput(T1(a[0])); if (typeof p.v !== 'number') throw E.VALUE; return Math.floor(p.v); }, SC);
def('TIMEVALUE', DATE, 'time_text', 'Chuyển chuỗi giờ thành số seri.', a => { const p = parseInput(T1(a[0])); if (typeof p.v !== 'number') throw E.VALUE; return p.v - Math.floor(p.v); }, SC);
function holidaysSet(c, h) { const s = new Set(); if (h != null) flatVals(c, h).forEach(v => { if (typeof v === 'number') s.add(Math.floor(v)); }); return s; }
function weekendMask(w) { if (w == null || w === 1) return [1, 0, 0, 0, 0, 0, 1]; if (typeof w === 'string') { if (!/^[01]{7}$/.test(w)) throw E.VALUE; const m = w.split('').map(Number); return [m[6], m[0], m[1], m[2], m[3], m[4], m[5]]; } const map = { 2: [1, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6], 11: [0], 12: [1], 13: [2], 14: [3], 15: [4], 16: [5], 17: [6] }; const pairs = { 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6] }; const m = [0, 0, 0, 0, 0, 0, 0]; if (pairs[w]) pairs[w].forEach(d => { m[d] = 1; }); else if (map[w]) m[(map[w][0] + 0) % 7] = 1; else throw E.NUM; return m; }
function networkdays(s, e, mask, hol) { let n = 0; const st = Math.min(s, e), en = Math.max(s, e); for (let d = st; d <= en; d++) { const dow = serialToParts(d).dow; if (!mask[dow] && !hol.has(d)) n++; } return s > e ? -n : n; }
def('NETWORKDAYS', DATE, 'start_date, end_date, [holidays]', 'Số ngày làm việc giữa hai ngày.', (a, x, c) => networkdays(Math.floor(N(scalar(c, a[0]))), Math.floor(N(scalar(c, a[1]))), weekendMask(1), holidaysSet(c, a[2])));
def('NETWORKDAYS.INTL', DATE, 'start_date, end_date, [weekend], [holidays]', 'Số ngày làm việc với cuối tuần tùy chỉnh.', (a, x, c) => networkdays(Math.floor(N(scalar(c, a[0]))), Math.floor(N(scalar(c, a[1]))), weekendMask(a[2] == null ? 1 : scalar(c, a[2])), holidaysSet(c, a[3])));
function workday(s, k, mask, hol) { let d = s; const step = k < 0 ? -1 : 1; let left = Math.abs(k); while (left > 0) { d += step; const dow = serialToParts(d).dow; if (!mask[dow] && !hol.has(d)) left--; } return d; }
def('WORKDAY', DATE, 'start_date, days, [holidays]', 'Ngày làm việc sau n ngày.', (a, x, c) => workday(Math.floor(N(scalar(c, a[0]))), toInt(scalar(c, a[1])), weekendMask(1), holidaysSet(c, a[2])));
def('WORKDAY.INTL', DATE, 'start_date, days, [weekend], [holidays]', 'Ngày làm việc với cuối tuần tùy chỉnh.', (a, x, c) => workday(Math.floor(N(scalar(c, a[0]))), toInt(scalar(c, a[1])), weekendMask(a[2] == null ? 1 : scalar(c, a[2])), holidaysSet(c, a[3])));
def('DAYS360', DATE, 'start_date, end_date, [method]', 'Số ngày theo năm 360 ngày.', a => { const p = P(a[0]), q = P(a[1]); const eu = a[2] != null && B(a[2]); let d1 = p.d, d2 = q.d; if (eu) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; } else { const lastFeb = z => z.m === 2 && z.d === new Date(Date.UTC(z.y, 2, 0)).getUTCDate(); if (lastFeb(p)) d1 = 30; if (d1 === 31) d1 = 30; if (d2 === 31 && d1 >= 30) d2 = 30; } return (q.y - p.y) * 360 + (q.m - p.m) * 30 + d2 - d1; }, SC);

// ================================================================ LOOKUP & REFERENCE
function lookupApprox(list, v) { // list ascending; last <= v
  let lo = 0, hi = list.length - 1, ans = -1;
  while (lo <= hi) { const mid = (lo + hi) >> 1; const x = list[mid]; if (x == null || (typeof x !== typeof v && !(typeof v === 'number' && typeof x === 'number'))) { // skip mismatched types linearly
      let j = mid - 1; while (j >= lo && (list[j] == null || typeof list[j] !== typeof v)) j--; if (j < lo) { lo = mid + 1; continue; } const c2 = cmp(list[j], v); if (c2 <= 0) { ans = j; lo = mid + 1; } else hi = j - 1; continue; }
    const c2 = cmp(x, v); if (c2 <= 0) { ans = mid; lo = mid + 1; } else hi = mid - 1; }
  return ans;
}
function exactIndex(list, v) {
  if (typeof v === 'string' && /[*?~]/.test(v)) { const re = wild(v); for (let i = 0; i < list.length; i++) if (typeof list[i] === 'string' && re.test(list[i])) return i; return -1; }
  for (let i = 0; i < list.length; i++) { const x = list[i]; if (x == null) continue; if (typeof x === typeof v && cmp(x, v) === 0) return i; }
  return -1;
}
def('VLOOKUP', LOOK, 'lookup_value, table_array, col_index_num, [range_lookup]', 'Tra cứu theo cột đầu tiên và trả về giá trị cùng hàng.', (a, x, c) => {
  const v = scalar(c, a[0]); if (v instanceof Err) throw v;
  const tb = a[1]; const ci = toInt(scalar(c, a[2])); const approx = a[3] == null || a[3] === null ? true : B(scalar(c, a[3]));
  const [rows, cols] = dims(c, tb); if (ci < 1) throw E.VALUE; if (ci > cols) throw E.REF;
  let col; if (tb instanceof Ref) { const u = tb.sheet.used(); const r2 = Math.min(tb.r2, Math.max(u.r2, tb.r1)); col = []; for (let r = tb.r1; r <= r2; r++) col.push(c.cellValue(tb.sheet, r, tb.c1)); }
  else col = tb.map(r => r[0]);
  const i = approx ? lookupApprox(col, v) : exactIndex(col, v);
  if (i < 0) throw E.NA;
  return tb instanceof Ref ? c.cellValue(tb.sheet, tb.r1 + i, tb.c1 + ci - 1) : tb[i][ci - 1];
});
def('HLOOKUP', LOOK, 'lookup_value, table_array, row_index_num, [range_lookup]', 'Tra cứu theo hàng đầu tiên.', (a, x, c) => {
  const v = scalar(c, a[0]); const m = arr2(c, a[1]); const ri = toInt(scalar(c, a[2])); const approx = a[3] == null ? true : B(scalar(c, a[3]));
  if (ri < 1) throw E.VALUE; if (ri > m.length) throw E.REF;
  const i = approx ? lookupApprox(m[0], v) : exactIndex(m[0], v); if (i < 0) throw E.NA; return m[ri - 1][i];
});
def('LOOKUP', LOOK, 'lookup_value, lookup_vector, [result_vector]', 'Tra cứu trong vectơ hoặc mảng.', (a, x, c) => {
  const v = scalar(c, a[0]); const m = arr2(c, a[1]);
  let lv, rv;
  if (a[2] != null) { lv = m.length === 1 ? m[0] : m.map(r => r[0]); const R = arr2(c, a[2]); rv = R.length === 1 ? R[0] : R.map(r => r[0]); }
  else if (m[0].length > m.length) { lv = m[0]; rv = m[m.length - 1]; } else { lv = m.map(r => r[0]); rv = m.map(r => r[r.length - 1]); }
  const i = lookupApprox(lv, v); if (i < 0) throw E.NA; return rv[i];
});
def('MATCH', LOOK, 'lookup_value, lookup_array, [match_type]', 'Vị trí tương đối của giá trị trong mảng.', (a, x, c) => {
  const v = scalar(c, a[0]); const m = arr2(c, a[1]); const list = m.length === 1 ? m[0] : m.map(r => r[0]); const t = a[2] == null ? 1 : N(scalar(c, a[2]));
  let i;
  if (t === 0) i = exactIndex(list, v);
  else if (t > 0) i = lookupApprox(list, v);
  else { i = -1; for (let k = 0; k < list.length; k++) { if (list[k] == null) continue; if (cmp(list[k], v) >= 0) i = k; else break; } }
  if (i < 0) throw E.NA; return i + 1;
});
function xmatchCore(list, v, mode, search) {
  const idxs = list.map((_, i) => i); if (search === -1 || search === -2) idxs.reverse();
  if (mode === 2 && typeof v === 'string') { const re = wild(v); for (const i of idxs) if (typeof list[i] === 'string' && re.test(list[i])) return i; return -1; }
  if (mode === 3 && typeof v === 'string') { let re; try { re = new RegExp(v); } catch (e) { throw E.VALUE; } for (const i of idxs) if (typeof list[i] === 'string' && re.test(list[i])) return i; return -1; }
  for (const i of idxs) { const z = list[i]; if (z != null && typeof z === typeof v && cmp(z, v) === 0) return i; }
  if (mode === -1 || mode === 1) {
    let best = -1;
    for (const i of idxs) { const z = list[i]; if (z == null || typeof z !== typeof v) continue; const c2 = cmp(z, v); if (mode === -1 && c2 < 0 && (best < 0 || cmp(z, list[best]) > 0)) best = i; if (mode === 1 && c2 > 0 && (best < 0 || cmp(z, list[best]) < 0)) best = i; }
    return best;
  }
  return -1;
}
def('XMATCH', LOOK, 'lookup_value, lookup_array, [match_mode], [search_mode]', 'Vị trí của giá trị (hiện đại).', (a, x, c) => { const m = arr2(c, a[1]); const list = m.length === 1 ? m[0] : m.map(r => r[0]); const vv = a[0]; const one = v => { const i = xmatchCore(list, v, a[2] == null ? 0 : N(scalar(c, a[2])), a[3] == null ? 1 : N(scalar(c, a[3]))); if (i < 0) return E.NA; return i + 1; }; if (vv instanceof Ref || isArr(vv)) return arr2(c, vv).map(r => r.map(one)); const r = one(vv); if (r instanceof Err) throw r; return r; });
def('XLOOKUP', LOOK, 'lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]', 'Tra cứu linh hoạt theo hàng hoặc cột.', (a, x, c) => {
  const v = scalar(c, a[0]); const L = arr2(c, a[1]); const R = arr2(c, a[2]);
  const vertical = L[0].length === 1; const list = vertical ? L.map(r => r[0]) : L[0];
  const i = xmatchCore(list, v, a[4] == null ? 0 : N(scalar(c, a[4])), a[5] == null ? 1 : N(scalar(c, a[5])));
  if (i < 0) { if (a.length > 3 && a[3] !== null) return a[3]; throw E.NA; }
  if (vertical) { const row = R[i]; if (!row) throw E.VALUE; return row.length === 1 ? row[0] : [row]; }
  const col = R.map(r => r[i]); return col.length === 1 ? col[0] : col.map(z => [z]);
});
def('INDEX', LOOK, 'array, row_num, [column_num], [area_num]', 'Giá trị tại vị trí hàng/cột trong mảng.', (a, x, c) => {
  const src = a[0]; let r = a[1] == null ? 0 : toInt(scalar(c, a[1])), col = a[2] == null ? 0 : toInt(scalar(c, a[2]));
  const [rows, cols] = dims(c, src);
  if (rows === 1 && a[2] == null && cols > 1) { col = r; r = 1; }
  if (r < 0 || col < 0 || r > rows || col > cols) throw E.REF;
  if (src instanceof Ref) {
    if (r === 0 && col === 0) return src;
    if (r === 0) return new Ref(src.sheet, src.r1, src.c1 + col - 1, src.r2, src.c1 + col - 1);
    if (col === 0) return cols === 1 ? new Ref(src.sheet, src.r1 + r - 1, src.c1, src.r1 + r - 1, src.c1) : new Ref(src.sheet, src.r1 + r - 1, src.c1, src.r1 + r - 1, src.c2);
    return new Ref(src.sheet, src.r1 + r - 1, src.c1 + col - 1, src.r1 + r - 1, src.c1 + col - 1);
  }
  const m = toArr(src);
  if (r === 0) return m.map(row => [row[col - 1]]);
  if (col === 0) return cols === 1 ? m[r - 1][0] : [m[r - 1]];
  return m[r - 1][col - 1];
});
def('OFFSET', LOOK, 'reference, rows, cols, [height], [width]', 'Tham chiếu dịch chuyển từ một ô.', (a, x, c) => { const ref = a[0]; if (!(ref instanceof Ref)) throw E.VALUE; const dr = toInt(scalar(c, a[1])), dc = toInt(scalar(c, a[2])); const hgt = a[3] == null ? ref.rows : toInt(scalar(c, a[3])), wid = a[4] == null ? ref.cols : toInt(scalar(c, a[4])); const r1 = ref.r1 + dr, c1 = ref.c1 + dc; if (r1 < 0 || c1 < 0 || hgt < 1 || wid < 1) throw E.REF; return new Ref(ref.sheet, r1, c1, r1 + hgt - 1, c1 + wid - 1); }, { volatile: true });
def('INDIRECT', LOOK, 'ref_text, [a1]', 'Tham chiếu từ chuỗi văn bản.', (a, ctx, c) => {
  const t = S(scalar(c, a[0])); const a1 = a[1] == null || B(scalar(c, a[1]));
  let sheet = ctx.sheet, refTxt = t, sheets = [sheet]; const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(t);
  if (m) {
    const rawSheet = (m[1] || m[2]).replace(/''/g, "'"); refTxt = m[3];
    if (rawSheet.indexOf(':') >= 0) {
      const parts = rawSheet.split(':');
      const s1 = c.wb.sheet(parts[0].replace(/^'|'$/g, '').trim()), s2 = c.wb.sheet(parts[1].replace(/^'|'$/g, '').trim());
      if (!s1 || !s2) throw E.REF;
      const idx1 = c.wb.sheets.indexOf(s1), idx2 = c.wb.sheets.indexOf(s2);
      if (idx1 < 0 || idx2 < 0) throw E.REF;
      sheets = c.wb.sheets.slice(Math.min(idx1, idx2), Math.max(idx1, idx2) + 1); sheet = s1;
    } else {
      sheet = c.wb.sheet(rawSheet); if (!sheet) throw E.REF; sheets = [sheet];
    }
  }
  if (!a1) { const mm = /^R(\[?-?\d*\]?)C(\[?-?\d*\]?)$/i.exec(refTxt); if (!mm) throw E.REF; const rr = rc(mm[1], ctx.r), cc = rc(mm[2], ctx.c); return new Ref(sheet, rr, cc, rr, cc, sheets); }
  const def2 = c.wb.names[refTxt.toUpperCase()]; if (def2) { const v = c.eval(parse(String(def2.f || def2).replace(/^=/, '')), ctx); if (v instanceof Ref) return v; }
  const rg = parseRangeAddr(refTxt); if (!rg) throw E.REF; return new Ref(sheet, rg.r1, rg.c1, rg.r2, rg.c2, sheets);
}, { volatile: true });
function rc(s, base) { if (!s) return base; const m = /^\[(-?\d+)\]$/.exec(s); if (m) return base + +m[1]; return +s - 1; }
def('ADDRESS', LOOK, 'row_num, column_num, [abs_num], [a1], [sheet_text]', 'Địa chỉ ô dạng văn bản.', a => { const r = toInt(a[0]), cc = toInt(a[1]); const ab = a[2] == null ? 1 : toInt(a[2]); const a1 = a[3] == null || B(a[3]); if (r < 1 || cc < 1) throw E.VALUE; let s; if (a1) s = (ab === 1 || ab === 3 ? '$' : '') + colName(cc - 1) + (ab === 1 || ab === 2 ? '$' : '') + r; else s = 'R' + (ab === 1 || ab === 2 ? r : '[' + r + ']') + 'C' + (ab === 1 || ab === 3 ? cc : '[' + cc + ']'); if (a[4] != null) { const sh = S(a[4]); s = (/^[\w.]+$/.test(sh) ? sh : "'" + sh + "'") + '!' + s; } return s; }, SC);
def('ROW', LOOK, '[reference]', 'Số thứ tự hàng của tham chiếu.', (a, ctx) => { const r = a[0]; if (r == null) return ctx.r + 1; if (!(r instanceof Ref)) throw E.VALUE; if (r.rows === 1) return r.r1 + 1; const out = []; for (let i = r.r1; i <= r.r2; i++) out.push([i + 1]); return out; });
def('COLUMN', LOOK, '[reference]', 'Số thứ tự cột của tham chiếu.', (a, ctx) => { const r = a[0]; if (r == null) return ctx.c + 1; if (!(r instanceof Ref)) throw E.VALUE; if (r.cols === 1) return r.c1 + 1; const out = []; for (let i = r.c1; i <= r.c2; i++) out.push(i + 1); return [out]; });
def('ROWS', LOOK, 'array', 'Số hàng của mảng/tham chiếu.', (a, x, c) => dims(c, a[0])[0]);
def('COLUMNS', LOOK, 'array', 'Số cột của mảng/tham chiếu.', (a, x, c) => dims(c, a[0])[1]);
def('AREAS', LOOK, 'reference', 'Số vùng trong tham chiếu.', a => { if (!(a[0] instanceof Ref)) throw E.VALUE; return 1; });
def('TRANSPOSE', LOOK, 'array', 'Chuyển vị hàng thành cột.', (a, x, c) => { const m = arr2(c, a[0]); return m[0].map((_, j) => m.map(r => r[j] == null ? 0 : r[j])); });
def('CHOOSECOLS', LOOK, 'array, col_num1, [col_num2], ...', 'Chọn các cột của mảng.', (a, x, c) => { const m = arr2(c, a[0]); const idx = a.slice(1).map(v => { let k = toInt(scalar(c, v)); if (k < 0) k = m[0].length + k + 1; if (k < 1 || k > m[0].length) throw E.VALUE; return k - 1; }); return m.map(r => idx.map(k => r[k])); });
def('CHOOSEROWS', LOOK, 'array, row_num1, [row_num2], ...', 'Chọn các hàng của mảng.', (a, x, c) => { const m = arr2(c, a[0]); return a.slice(1).map(v => { let k = toInt(scalar(c, v)); if (k < 0) k = m.length + k + 1; if (k < 1 || k > m.length) throw E.VALUE; return m[k - 1]; }); });
def('TAKE', LOOK, 'array, rows, [columns]', 'Lấy các hàng/cột đầu hoặc cuối.', (a, x, c) => { let m = arr2(c, a[0]); if (a[1] != null) { const r = toInt(scalar(c, a[1])); m = r >= 0 ? m.slice(0, r) : m.slice(r); } if (a[2] != null) { const k = toInt(scalar(c, a[2])); m = m.map(row => k >= 0 ? row.slice(0, k) : row.slice(k)); } if (!m.length || !m[0].length) throw E.CALC; return m; });
def('DROP', LOOK, 'array, rows, [columns]', 'Bỏ các hàng/cột đầu hoặc cuối.', (a, x, c) => { let m = arr2(c, a[0]); if (a[1] != null) { const r = toInt(scalar(c, a[1])); m = r >= 0 ? m.slice(r) : m.slice(0, r); } if (a[2] != null) { const k = toInt(scalar(c, a[2])); m = m.map(row => k >= 0 ? row.slice(k) : row.slice(0, k)); } if (!m.length || !m[0].length) throw E.CALC; return m; });
def('HSTACK', LOOK, 'array1, [array2], ...', 'Ghép mảng theo chiều ngang.', (a, x, c) => { const ms = a.map(v => arr2(c, v)); const R = Math.max.apply(null, ms.map(m => m.length)); const out = []; for (let i = 0; i < R; i++) { let row = []; ms.forEach(m => { const src = m[i]; row = row.concat(src ? src : m[0].map(() => E.NA)); }); out.push(row); } return out; });
def('VSTACK', LOOK, 'array1, [array2], ...', 'Ghép mảng theo chiều dọc.', (a, x, c) => { const ms = a.map(v => arr2(c, v)); const C = Math.max.apply(null, ms.map(m => m[0].length)); const out = []; ms.forEach(m => m.forEach(r => { const row = r.slice(); while (row.length < C) row.push(E.NA); out.push(row); })); return out; });
def('TOCOL', LOOK, 'array, [ignore], [scan_by_column]', 'Chuyển mảng thành một cột.', (a, x, c) => { const m = arr2(c, a[0]); const ig = a[1] == null ? 0 : N(scalar(c, a[1])); const bycol = a[2] != null && B(scalar(c, a[2])); const out = []; const push = v => { if ((ig === 1 || ig === 3) && (v == null)) return; if ((ig === 2 || ig === 3) && v instanceof Err) return; out.push([v == null ? 0 : v]); }; if (bycol) { for (let j = 0; j < m[0].length; j++) for (let i = 0; i < m.length; i++) push(m[i][j]); } else m.forEach(r => r.forEach(push)); return out.length ? out : (() => { throw E.CALC; })(); });
def('TOROW', LOOK, 'array, [ignore], [scan_by_column]', 'Chuyển mảng thành một hàng.', (a, x, c) => { const col = FN.TOCOL.f(a, x, c); return [col.map(r => r[0])]; });
def('WRAPROWS', LOOK, 'vector, wrap_count, [pad_with]', 'Gói vectơ thành các hàng.', (a, x, c) => { const v = flatVals(c, a[0]); const n = toInt(scalar(c, a[1])); if (n < 1) throw E.NUM; const pad = a[2] === undefined ? E.NA : a[2]; const out = []; for (let i = 0; i < v.length; i += n) { const r = v.slice(i, i + n); while (r.length < n) r.push(pad); out.push(r); } return out; });
def('WRAPCOLS', LOOK, 'vector, wrap_count, [pad_with]', 'Gói vectơ thành các cột.', (a, x, c) => { const rows = FN.WRAPROWS.f(a, x, c); return rows[0].map((_, j) => rows.map(r => r[j])); });
def('EXPAND', LOOK, 'array, rows, [columns], [pad_with]', 'Mở rộng mảng tới kích thước chỉ định.', (a, x, c) => { const m = arr2(c, a[0]); const R = a[1] == null ? m.length : toInt(scalar(c, a[1])), C = a[2] == null ? m[0].length : toInt(scalar(c, a[2])); if (R < m.length || C < m[0].length) throw E.VALUE; const pad = a[3] === undefined ? E.NA : a[3]; const out = []; for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) row.push(m[i] && j < m[i].length ? m[i][j] : pad); out.push(row); } return out; });
def('FILTER', LOOK, 'array, include, [if_empty]', 'Lọc mảng theo điều kiện.', (a, x, c) => {
  const m = arr2(c, a[0]); const inc = arr2(c, a[1]); const out = [];
  if (inc.length === m.length && inc[0].length === 1) { m.forEach((r, i) => { const b = bool(inc[i][0]); if (b instanceof Err) throw b; if (b) out.push(r); }); }
  else if (inc.length === 1 && inc[0].length === m[0].length) { const keep = inc[0].map(v => bool(v) === true); m.forEach(r => out.push(r.filter((_, j) => keep[j]))); if (!out[0].length) out.length = 0; }
  else throw E.VALUE;
  if (!out.length) { if (a.length > 2) return a[2]; throw E.CALC; }
  return out;
});
function sortRows(m, keys) { const idx = m.map((_, i) => i); idx.sort((p, q) => { for (const k of keys) { const r = cmp(k.col(p), k.col(q)); if (r) return r * k.dir; } return p - q; }); return idx; }
def('SORT', LOOK, 'array, [sort_index], [sort_order], [by_col]', 'Sắp xếp mảng.', (a, x, c) => {
  let m = arr2(c, a[0]); const bycol = a[3] != null && B(scalar(c, a[3])); if (bycol) m = m[0].map((_, j) => m.map(r => r[j]));
  const si = a[1] == null ? [1] : flatVals(c, a[1]).map(v => toInt(v)); const so = a[2] == null ? [1] : flatVals(c, a[2]).map(v => toInt(v));
  const keys = si.map((k, i) => ({ col: r => m[r][k - 1], dir: (so[i] != null ? so[i] : so[0]) === -1 ? -1 : 1 }));
  let out = sortRows(m, keys).map(i => m[i]); if (bycol) out = out[0].map((_, j) => out.map(r => r[j])); return out;
});
def('SORTBY', LOOK, 'array, by_array1, [sort_order1], ...', 'Sắp xếp mảng theo mảng khác.', (a, x, c) => { const m = arr2(c, a[0]); const keys = []; for (let i = 1; i < a.length; i += 2) { const b = arr2(c, a[i]); const col = b.length === m.length ? b.map(r => r[0]) : b[0]; keys.push({ col: r => col[r], dir: a[i + 1] != null && toInt(scalar(c, a[i + 1])) === -1 ? -1 : 1 }); } return sortRows(m, keys).map(i => m[i]); });
def('UNIQUE', LOOK, 'array, [by_col], [exactly_once]', 'Trả về các giá trị duy nhất.', (a, x, c) => {
  let m = arr2(c, a[0]); const bycol = a[1] != null && B(scalar(c, a[1])); const once = a[2] != null && B(scalar(c, a[2]));
  if (bycol) m = m[0].map((_, j) => m.map(r => r[j]));
  const keyOf = r => r.map(v => typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + String(v)).join('\u0001');
  const cnt = new Map(); m.forEach(r => { const k = keyOf(r); cnt.set(k, (cnt.get(k) || 0) + 1); });
  const seen = new Set(); let out = [];
  m.forEach(r => { const k = keyOf(r); if (seen.has(k)) return; seen.add(k); if (once && cnt.get(k) > 1) return; out.push(r); });
  if (!out.length) throw E.CALC; if (bycol) out = out[0].map((_, j) => out.map(r => r[j])); return out;
});
def('FORMULATEXT', LOOK, 'reference', 'Trả về công thức dạng văn bản.', a => { const r = a[0]; if (!(r instanceof Ref)) throw E.VALUE; const cell = r.sheet.get(r.r1, r.c1); if (!cell || !cell.f) throw E.NA; return '=' + cell.f; });
def('HYPERLINK', LOOK, 'link_location, [friendly_name]', 'Tạo liên kết.', (a, x, c) => { const l = scalar(c, a[0]); return a[1] != null ? scalar(c, a[1]) : l; });
def('GETPIVOTDATA', LOOK, 'data_field, pivot_table, [field1, item1], ...', 'Lấy dữ liệu từ PivotTable.', () => { throw E.REF; });
def('TRIMRANGE', LOOK, 'range, [trim_rows], [trim_cols]', 'Loại bỏ hàng/cột trống ở biên.', (a, x, c) => { let m = arr2(c, a[0]); const empty = v => v == null || v === ''; while (m.length > 1 && m[m.length - 1].every(empty)) m.pop(); while (m.length > 1 && m[0].every(empty)) m.shift(); while (m[0].length > 1 && m.every(r => empty(r[r.length - 1]))) m = m.map(r => r.slice(0, -1)); return m.map(r => r.map(v => v == null ? 0 : v)); });
def('GROUPBY', LOOK, 'row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array]', 'Nhóm dữ liệu và tổng hợp theo hàng.', (a, ctx, c) => {
  const keys = arr2(c, a[0]); const vals = arr2(c, a[1]); const fn = a[2];
  const groups = new Map(); const order = [];
  keys.forEach((r, i) => { const k = r.map(v => str(v)).join('\u0001'); if (!groups.has(k)) { groups.set(k, { key: r, vals: [] }); order.push(k); } groups.get(k).vals.push(vals[i][0]); });
  const agg = list => { if (fn instanceof Lambda) return scalar(c, c.applyLambda(fn, [list.map(v => [v])], ctx)); const n = list.filter(v => typeof v === 'number'); return sum(n); };
  order.sort((p, q) => cmp(groups.get(p).key[0], groups.get(q).key[0]));
  const out = order.map(k => groups.get(k).key.concat([agg(groups.get(k).vals)]));
  out.push(['Total'].concat(keys[0].slice(1).map(() => '')).concat([agg(vals.map(r => r[0]))]));
  return out;
});
def('PERCENTOF', MATH, 'data_subset, data_all', 'Tỷ lệ phần trăm của tập con trên tổng.', (a, x, c) => { const s = sum(nums(c, [a[0]])), t = sum(nums(c, [a[1]])); if (!t) throw E.DIV0; return s / t; });

// ================================================================ INFORMATION
const isF = f => (a, ctx, c) => { const v = a[0]; if (v instanceof Ref && v.rows * v.cols > 1) return c.deref(v).map(r => r.map(f)); return f(v instanceof Ref ? c.cellValue(v.sheet, v.r1, v.c1) : v); };
def('ISBLANK', INFO, 'value', 'TRUE nếu ô trống.', isF(v => v == null));
def('ISNUMBER', INFO, 'value', 'TRUE nếu là số.', isF(v => typeof v === 'number'));
def('ISTEXT', INFO, 'value', 'TRUE nếu là văn bản.', isF(v => typeof v === 'string'));
def('ISNONTEXT', INFO, 'value', 'TRUE nếu không phải văn bản.', isF(v => typeof v !== 'string'));
def('ISLOGICAL', INFO, 'value', 'TRUE nếu là giá trị logic.', isF(v => typeof v === 'boolean'));
def('ISERROR', INFO, 'value', 'TRUE nếu là bất kỳ lỗi nào.', (args, ctx, c) => { try { let v = c.eval(args[0], ctx); if (v instanceof Ref) v = c.deref(v, true); if (isArr(v)) return v.map(r => r.map(z => z instanceof Err)); return v instanceof Err; } catch (e) { if (e instanceof Err) return true; throw e; } }, { special: true });
def('ISERR', INFO, 'value', 'TRUE nếu là lỗi khác #N/A.', (args, ctx, c) => { try { let v = c.eval(args[0], ctx); if (v instanceof Ref) v = c.deref(v, true); return v instanceof Err && v.code !== '#N/A'; } catch (e) { if (e instanceof Err) return e.code !== '#N/A'; throw e; } }, { special: true });
def('ISNA', INFO, 'value', 'TRUE nếu là lỗi #N/A.', (args, ctx, c) => { try { let v = c.eval(args[0], ctx); if (v instanceof Ref) v = c.deref(v, true); return v instanceof Err && v.code === '#N/A'; } catch (e) { if (e instanceof Err) return e.code === '#N/A'; throw e; } }, { special: true });
def('ISREF', INFO, 'value', 'TRUE nếu là tham chiếu.', (args, ctx, c) => { try { return c.eval(args[0], ctx) instanceof Ref; } catch (e) { return false; } }, { special: true });
def('ISFORMULA', INFO, 'reference', 'TRUE nếu ô chứa công thức.', a => { const r = a[0]; if (!(r instanceof Ref)) throw E.VALUE; const cell = r.sheet.get(r.r1, r.c1); return !!(cell && cell.f); });
def('ISEVEN', INFO, 'number', 'TRUE nếu là số chẵn.', a => Math.floor(Math.abs(N(a[0]))) % 2 === 0, SC);
def('ISODD', INFO, 'number', 'TRUE nếu là số lẻ.', a => Math.floor(Math.abs(N(a[0]))) % 2 === 1, SC);
def('TYPE', INFO, 'value', 'Kiểu dữ liệu (1 số, 2 văn bản, 4 logic, 16 lỗi, 64 mảng).', (args, ctx, c) => { let v; try { v = c.eval(args[0], ctx); } catch (e) { return 16; } if (v instanceof Ref) { if (v.rows * v.cols > 1) return 64; v = c.deref(v, true); } if (isArr(v)) return 64; return typeof v === 'number' || v == null ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : v instanceof Err ? 16 : 1; }, { special: true });
def('ERROR.TYPE', INFO, 'error_val', 'Số tương ứng với loại lỗi.', (args, ctx, c) => { let v; try { v = scalar(c, c.eval(args[0], ctx)); } catch (e) { v = e; } if (!(v instanceof Err)) throw E.NA; return { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#GETTING_DATA': 8, '#SPILL!': 9, '#CALC!': 14 }[v.code] || 8; }, { special: true });
def('NA', INFO, '', 'Trả về lỗi #N/A.', () => { throw E.NA; });
def('SHEET', INFO, '[value]', 'Số thứ tự của trang tính.', (a, ctx, c) => {
  const v = a[0];
  if (v == null) return c.wb.sheets.indexOf(ctx.sheet) + 1;
  if (v instanceof Ref) { const s = (v.sheets && v.sheets[0]) || v.sheet; return c.wb.sheets.indexOf(s) + 1; }
  const s = c.wb.sheet(S(v)); if (!s) throw E.NA; return c.wb.sheets.indexOf(s) + 1;
});
def('SHEETS', INFO, '[reference]', 'Số trang tính trong tham chiếu hoặc trong toàn bộ sổ làm việc.', (a, ctx, c) => {
  if (a[0] == null) return c.wb.sheets.length;
  if (a[0] instanceof Ref) return (a[0].sheets && a[0].sheets.length) ? a[0].sheets.length : 1;
  return 1;
});
def('CELL', INFO, 'info_type, [reference]', 'Thông tin về ô.', (a, ctx, c) => { const t = S(scalar(c, a[0])).toLowerCase(); const r = a[1] instanceof Ref ? a[1] : new Ref(ctx.sheet, ctx.r, ctx.c, ctx.r, ctx.c); const v = c.cellValue(r.sheet, r.r1, r.c1); switch (t) { case 'address': return addr(r.r1, r.c1, true, true); case 'row': return r.r1 + 1; case 'col': return r.c1 + 1; case 'contents': return v; case 'type': return v == null ? 'b' : typeof v === 'string' ? 'l' : 'v'; case 'filename': return '[' + (c.wb.props.name || 'Book1') + ']' + r.sheet.name; case 'format': return 'G'; case 'width': return Math.round(r.sheet.colWidth(r.c1) / 7); case 'protect': return 1; case 'prefix': return typeof v === 'string' ? "'" : ''; case 'color': return 0; case 'parentheses': return 0; } throw E.VALUE; });
def('INFO', INFO, 'type_text', 'Thông tin về môi trường.', a => { const t = T1(a[0]).toLowerCase(); switch (t) { case 'numfile': return 1; case 'osversion': return navigator.platform; case 'recalc': return 'Automatic'; case 'release': return '16.0'; case 'system': return /Win/.test(navigator.platform) ? 'pcdos' : 'mac'; case 'directory': return '/'; case 'origin': return '$A:$A$1'; } throw E.VALUE; }, SC);

// ================================================================ DATABASE (native: header row + criteria range, rows OR / columns AND)
function dbSelect(c, a) {
  const db = arr2(c, a[0]); const hdr = db[0].map(v => str(v).toLowerCase());
  let f = scalar(c, a[1]); let fi = -1;
  if (f != null) { if (typeof f === 'number') fi = Math.floor(f) - 1; else fi = hdr.indexOf(str(f).toLowerCase()); if (fi < 0 || fi >= hdr.length) throw E.VALUE; }
  const cr = arr2(c, a[2]); const ch = cr[0].map(v => str(v).toLowerCase());
  const conds = cr.slice(1).map(row => row.map((v, j) => ({ col: hdr.indexOf(ch[j]), pred: v == null || v === '' ? null : criteria(v) })).filter(x => x.pred));
  const rows = db.slice(1).filter(row => !conds.length || conds.some(cs => cs.every(x => x.col >= 0 && x.pred(row[x.col]))));
  return { vals: rows.map(r => fi >= 0 ? r[fi] : null), rows };
}
const dnum = v => v.filter(x => typeof x === 'number');
def('DSUM', DB, 'database, field, criteria', 'Tổng giá trị trường của các bản ghi thỏa điều kiện.', (a, x, c) => fixf(sum(dnum(dbSelect(c, a).vals))));
def('DAVERAGE', DB, 'database, field, criteria', 'Trung bình các bản ghi thỏa điều kiện.', (a, x, c) => { const v = dnum(dbSelect(c, a).vals); if (!v.length) throw E.DIV0; return sum(v) / v.length; });
def('DCOUNT', DB, 'database, [field], criteria', 'Đếm ô chứa số trong các bản ghi thỏa điều kiện.', (a, x, c) => { const r = dbSelect(c, a); return a[1] == null ? r.rows.length : dnum(r.vals).length; });
def('DCOUNTA', DB, 'database, [field], criteria', 'Đếm ô không trống trong các bản ghi thỏa điều kiện.', (a, x, c) => { const r = dbSelect(c, a); return a[1] == null ? r.rows.length : r.vals.filter(v => v != null && v !== '').length; });
def('DGET', DB, 'database, field, criteria', 'Lấy một giá trị duy nhất thỏa điều kiện.', (a, x, c) => { const v = dbSelect(c, a).vals; if (!v.length) throw E.VALUE; if (v.length > 1) throw E.NUM; return v[0]; });
def('DMAX', DB, 'database, field, criteria', 'Giá trị lớn nhất thỏa điều kiện.', (a, x, c) => { const v = dnum(dbSelect(c, a).vals); return v.length ? Math.max.apply(null, v) : 0; });
def('DMIN', DB, 'database, field, criteria', 'Giá trị nhỏ nhất thỏa điều kiện.', (a, x, c) => { const v = dnum(dbSelect(c, a).vals); return v.length ? Math.min.apply(null, v) : 0; });
def('DPRODUCT', DB, 'database, field, criteria', 'Tích các giá trị thỏa điều kiện.', (a, x, c) => dnum(dbSelect(c, a).vals).reduce((p, v) => p * v, 1));
def('DSTDEV', DB, 'database, field, criteria', 'Độ lệch chuẩn mẫu của bản ghi thỏa điều kiện.', (a, x, c) => Math.sqrt(variance(dnum(dbSelect(c, a).vals), false)));
def('DSTDEVP', DB, 'database, field, criteria', 'Độ lệch chuẩn tổng thể của bản ghi thỏa điều kiện.', (a, x, c) => Math.sqrt(variance(dnum(dbSelect(c, a).vals), true)));
def('DVAR', DB, 'database, field, criteria', 'Phương sai mẫu của bản ghi thỏa điều kiện.', (a, x, c) => variance(dnum(dbSelect(c, a).vals), false));
def('DVARP', DB, 'database, field, criteria', 'Phương sai tổng thể của bản ghi thỏa điều kiện.', (a, x, c) => variance(dnum(dbSelect(c, a).vals), true));

// ================================================================ WEB / CUBE / misc
def('FILTERXML', WEB, 'xml, xpath', 'Trích dữ liệu từ XML bằng XPath.', a => { try { const doc = new DOMParser().parseFromString(T1(a[0]), 'text/xml'); const r = doc.evaluate(T1(a[1]), doc, null, 7, null); const out = []; for (let i = 0; i < r.snapshotLength; i++) { const t = r.snapshotItem(i).textContent; const p = parseInput(t); out.push([typeof p.v === 'number' ? p.v : t]); } if (!out.length) throw E.VALUE; return out.length === 1 ? out[0][0] : out; } catch (e) { throw e instanceof Err ? e : E.VALUE; } });
def('WEBSERVICE', WEB, 'url', 'Dữ liệu từ dịch vụ web (không khả dụng khi ngoại tuyến).', () => { throw E.VALUE; });
['CUBEKPIMEMBER', 'CUBEMEMBER', 'CUBEMEMBERPROPERTY', 'CUBERANKEDMEMBER', 'CUBESET', 'CUBESETCOUNT', 'CUBEVALUE'].forEach(n => def(n, CUBE, '...', 'Hàm khối OLAP (cần kết nối máy chủ dữ liệu — không khả dụng ngoại tuyến).', () => { throw E.NA; }));
def('EUROCONVERT', ENG, 'number, source, target, [full_precision], [triangulation_precision]', 'Chuyển đổi tiền tệ châu Âu cũ sang Euro.', a => { const R = { EUR: 1, BEF: 40.3399, DEM: 1.95583, ESP: 166.386, FRF: 6.55957, IEP: 0.787564, ITL: 1936.27, LUF: 40.3399, NLG: 2.20371, ATS: 13.7603, PTE: 200.482, FIM: 5.94573, GRD: 340.75, SIT: 239.64, CYP: 0.585274, MTL: 0.4293, SKK: 30.126, EEK: 15.6466, LVL: 0.702804, LTL: 3.4528 }; const s = R[T1(a[1]).toUpperCase()], t = R[T1(a[2]).toUpperCase()]; if (!s || !t) throw E.VALUE; return N(a[0]) / s * t; }, SC);
def('CALL', 'Bổ trợ', '...', 'Gọi thủ tục DLL (không hỗ trợ vì lý do bảo mật).', () => { throw E.VALUE; });
def('REGISTER.ID', 'Bổ trợ', '...', 'Mã đăng ký DLL (không hỗ trợ).', () => { throw E.VALUE; });
def('RTD', LOOK, 'progID, server, topic1, ...', 'Dữ liệu thời gian thực (cần máy chủ COM — không hỗ trợ).', () => { throw E.NA; });
def('IMAGE', LOOK, 'source, [alt_text], [sizing], [height], [width]', 'Chèn ảnh từ URL vào ô (hiển thị văn bản thay thế khi ngoại tuyến).', (a, x, c) => a[1] != null ? S(scalar(c, a[1])) : S(scalar(c, a[0])));
def('TRANSLATE', TEXT, 'text, [source_language], [target_language]', 'Dịch văn bản (cần A.I trực tuyến; ngoại tuyến trả nguyên văn).', (a, x, c) => S(scalar(c, a[0])));
def('DETECTLANGUAGE', TEXT, 'text', 'Nhận diện ngôn ngữ (đơn giản, ngoại tuyến).', (a, x, c) => { const s = S(scalar(c, a[0])); return /[ăâđêôơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹ]/i.test(s) ? 'vi' : /[\u3040-\u30ff]/.test(s) ? 'ja' : /[\u4e00-\u9fff]/.test(s) ? 'zh' : /[\uac00-\ud7af]/.test(s) ? 'ko' : 'en'; });

// ================================================================ Formula.js adapter for the long tail (statistics, finance, engineering, database...)
function toFjs(calc, v) {
  if (v instanceof Ref) { const m = calc.deref(v); return m.length === 1 && m[0].length === 1 ? m[0][0] : m; }
  return v;
}
function fromFjs(r) {
  if (r instanceof Error) return new Err(/^#/.test(r.message) ? r.message : '#VALUE!');
  if (r instanceof Date) return jsDateToSerial(r);
  if (r === undefined) return E.VALUE;
  if (typeof r === 'number' && !isFinite(r)) return E.NUM;
  if (Array.isArray(r)) { if (!r.length) return E.CALC; if (!Array.isArray(r[0])) return r.length === 1 ? fromFjs(r[0]) : r.map(v => [fromFjs(v)]); return r.map(row => row.map(fromFjs)); }
  return r;
}
const NUMS_ONLY = new Set(['AVEDEV', 'DEVSQ', 'GEOMEAN', 'HARMEAN', 'KURT', 'SKEW', 'SKEW.P', 'TRIMMEAN', 'NPV', 'IRR', 'MIRR', 'PERCENTRANK', 'PERCENTRANK.INC', 'PERCENTRANK.EXC', 'FREQUENCY', 'STANDARDIZE', 'MULTINOMIAL', 'SERIESSUM']);
const FJS_CATS = {
  [STAT]: 'AVEDEV BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CONFIDENCE.NORM CONFIDENCE.T CORREL COVARIANCE.P COVARIANCE.S DEVSQ EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FISHER FISHERINV FORECAST FORECAST.LINEAR FREQUENCY GAMMA GAMMA.DIST GAMMA.INV GAMMALN GAMMALN.PRECISE GAUSS GEOMEAN GROWTH HARMEAN HYPGEOM.DIST INTERCEPT KURT LINEST LOGEST LOGNORM.DIST LOGNORM.INV NEGBINOM.DIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV PEARSON PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST PROB RSQ SKEW SKEW.P SLOPE STANDARDIZE STEYX T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TREND TRIMMEAN WEIBULL.DIST Z.TEST COMBINA CHISQ.DIST CHISQ.INV CONFIDENCE CONFIDENCE.NORM CONFIDENCE.T BINOM.DIST.RANGE',
  [COMPAT]: 'BETADIST BETAINV BINOMDIST CHIDIST CHIINV CHITEST CONFIDENCE COVAR CRITBINOM EXPONDIST FDIST FINV FTEST GAMMADIST GAMMAINV HYPGEOMDIST LOGINV LOGNORMDIST NEGBINOMDIST NORMDIST NORMINV NORMSDIST NORMSINV PERCENTRANK POISSON TDIST TINV TTEST WEIBULL ZTEST',
  [FIN]: 'ACCRINT ACCRINTM AMORDEGRC AMORLINC COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD CUMIPMT CUMPRINC DB DDB DISC DOLLARDE DOLLARFR DURATION EFFECT FV FVSCHEDULE INTRATE IPMT IRR ISPMT MDURATION MIRR NOMINAL NPER NPV ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD PDURATION PMT PPMT PRICE PRICEDISC PRICEMAT PV RATE RECEIVED RRI SLN SYD TBILLEQ TBILLPRICE TBILLYIELD VDB XIRR XNPV YIELD YIELDDISC YIELDMAT STOCKHISTORY',
  [ENG]: 'BESSELI BESSELJ BESSELK BESSELY BIN2DEC BIN2HEX BIN2OCT BITAND BITLSHIFT BITOR BITRSHIFT BITXOR COMPLEX CONVERT DEC2BIN DEC2HEX DEC2OCT DELTA ERF ERF.PRECISE ERFC ERFC.PRECISE GESTEP HEX2BIN HEX2DEC HEX2OCT IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMCOSH IMCOT IMCSC IMCSCH IMDIV IMEXP IMLN IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL IMSEC IMSECH IMSIN IMSINH IMSQRT IMSUB IMSUM IMTAN OCT2BIN OCT2DEC OCT2HEX',
  [DB]: 'DAVERAGE DCOUNT DCOUNTA DGET DMAX DMIN DPRODUCT DSTDEV DSTDEVP DSUM DVAR DVARP',
  [MATH]: 'ACOSH ACOT ACOTH ARABIC ASINH ATANH BASE COMBINA COS COSH COT COTH CSC CSCH DECIMAL FACTDOUBLE MULTINOMIAL ROMAN SEC SECH SERIESSUM SINH SUMX2MY2 SUMX2PY2 SUMXMY2 TANH',
  [DATE]: 'YEARFRAC',
  [TEXT]: 'BAHTTEXT'
};
const FJS_ALIAS = { 'NORM.S.DIST': 'NORMSDIST', 'NORM.S.INV': 'NORMSINV', 'NORM.DIST': 'NORMDIST', 'NORM.INV': 'NORMINV', 'T.DIST.2T': 'TDIST', 'POISSON': 'POISSONDIST', 'WEIBULL': 'WEIBULLDIST', 'CONFIDENCE': 'CONFIDENCENORM', 'CONFIDENCE.NORM': 'CONFIDENCENORM', 'COVAR': 'COVAR', 'FORECAST.LINEAR': 'FORECAST', 'ERF.PRECISE': 'ERFPRECISE', 'ERFC.PRECISE': 'ERFCPRECISE', 'GAMMALN.PRECISE': 'GAMMALNPRECISE', 'Z.TEST': 'ZTEST', 'T.TEST': 'TTEST', 'F.TEST': 'FTEST', 'CHISQ.TEST': 'CHITEST', 'PERCENTRANK': 'PERCENTRANKINC', 'CHISQ.DIST.RT': 'CHIDISTRT', 'CHISQ.INV.RT': 'CHIINVRT', 'F.DIST.RT': 'FDISTRT', 'F.INV.RT': 'FINVRT', 'T.DIST.RT': 'TDISTRT', 'BETADIST': 'BETADIST', 'CRITBINOM': 'CRITBINOM', 'BINOM.INV': 'CRITBINOM' };
const VI_DESC = {
  STAT: 'Hàm thống kê', FIN: 'Hàm tài chính', ENG: 'Hàm kỹ thuật', DB: 'Hàm cơ sở dữ liệu (bảng có tiêu đề + vùng điều kiện).'
};
const fjsDescVi = { PMT: 'Khoản thanh toán định kỳ cho khoản vay.', PV: 'Giá trị hiện tại của khoản đầu tư.', FV: 'Giá trị tương lai của khoản đầu tư.', NPV: 'Giá trị hiện tại ròng.', IRR: 'Tỷ suất hoàn vốn nội bộ.', NPER: 'Số kỳ thanh toán.', RATE: 'Lãi suất mỗi kỳ.', IPMT: 'Tiền lãi trong một kỳ.', PPMT: 'Tiền gốc trong một kỳ.', SLN: 'Khấu hao đường thẳng.', CORREL: 'Hệ số tương quan.', FORECAST: 'Dự báo theo xu hướng tuyến tính.', SLOPE: 'Độ dốc đường hồi quy.', INTERCEPT: 'Giao điểm đường hồi quy.', CONVERT: 'Chuyển đổi đơn vị đo.', ROMAN: 'Đổi số Ả Rập sang số La Mã.', ARABIC: 'Đổi số La Mã sang số Ả Rập.', DSUM: 'Tổng các bản ghi thỏa điều kiện.', DCOUNT: 'Đếm ô số trong bản ghi thỏa điều kiện.', XIRR: 'IRR cho dòng tiền không định kỳ.', XNPV: 'NPV cho dòng tiền không định kỳ.', YEARFRAC: 'Phần năm giữa hai ngày.', BAHTTEXT: 'Đổi số thành chữ tiếng Thái.' };
for (const cat in FJS_CATS) {
  for (const name of FJS_CATS[cat].split(' ')) {
    if (FN[name]) continue;
    const fname = FJS_ALIAS[name] || name.replace(/\./g, '');
    let impl = FJS[fname];
    if (typeof impl !== 'function') { impl = name.split('.').reduce((o, k) => (o && o[k] != null ? o[k] : null), FJS); }   // v0.0.6: dotted names (CHISQ.DIST, CONFIDENCE.T, BINOM.DIST.RANGE…)
    if (typeof impl !== 'function') continue;
    const numsOnly = NUMS_ONLY.has(name);
    defFn(name, (a, ctx, calc) => {
      let args = a.map(v => toFjs(calc, v));
      if (numsOnly) args = args.map(v => isArr(v) ? [].concat.apply([], v).filter(z => typeof z === 'number') : v);
      for (const v of args) if (v instanceof Err) throw v;
      args = args.map(v => (v instanceof Err ? new Error(v.code) : v));
      let r; try { r = impl.apply(null, args); } catch (e) { throw E.VALUE; }
      return fromFjs(r);
    });
    meta[name] = [cat, '...', fjsDescVi[name] || (cat + ': ' + name)];
  }
}
// SUMX2MY2 etc and stragglers present in Formula.js under exact names
['FISHER', 'FISHERINV', 'GAMMA', 'GAUSS', 'PHI', 'PROB', 'STANDARDIZE', 'KURT', 'SKEW'].forEach(n => { if (!FN[n] && FJS[n]) defFn(n, (a, x, c) => fromFjs(FJS[n].apply(null, a.map(v => toFjs(c, v))))); });

function fixf(v) { return Math.abs(v) < 1e-15 ? 0 : +v.toPrecision(15); }

export function functionCatalog(extraDesc) {
  const out = [];
  Object.keys(FN).sort().forEach(n => { const m = meta[n] || ['Khác', '...', '']; let d = m[2]; if ((!d || /: /.test(d)) && extraDesc && extraDesc[n]) d = extraDesc[n]; out.push({ name: n, cat: m[0], sig: m[1], desc: d }); });
  return out;
}
export const FUNCTION_COUNT = () => Object.keys(FN).length;
export { meta as FUNCTION_META };
