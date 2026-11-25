// Yomi-Office Open Source v0.0.6 — the last Excel 365 functions that were missing in v0.0.5 (native, offline):
// CHISQ.DIST, CHISQ.INV, CONFIDENCE, CONFIDENCE.NORM, CONFIDENCE.T, BINOM.DIST.RANGE,
// FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY, FORECAST.ETS.STAT (additive Holt-Winters / AAA),
// PIVOTBY, plus STOCKHISTORY / FIELDVALUE / COPILOT / PY which need online services (return the same errors Excel shows offline).
import { defFn, FN, Ref, Lambda, num, str, cmp, isArr } from './formula.js';
import { Err, E, errOf } from './model.js';
import { registerMeta } from './functions.js';

const chk = v => { if (v instanceof Err) throw v; return v; };
const N = v => chk(num(v));
function flat(c, v) { const out = []; if (v instanceof Ref) c.eachIn(v, x => out.push(x)); else if (isArr(v)) v.forEach(r => r.forEach(x => out.push(x))); else out.push(v); return out; }
function arr2(c, v) { if (v instanceof Ref) return c.deref(v); if (isArr(v)) return v; return [[v]]; }

// ---- special functions
function gammaln(x) { const g = [76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5]; let y = x, t = x + 5.5; t -= (x + 0.5) * Math.log(t); let s = 1.000000000190015; for (const c of g) s += c / ++y; return -t + Math.log(2.5066282746310005 * s / x); }
function lowRegGamma(a, x) { // P(a, x)
  if (x <= 0) return 0;
  if (x < a + 1) { let sum = 1 / a, del = sum, ap = a; for (let n = 1; n < 500; n++) { ap++; del *= x / ap; sum += del; if (Math.abs(del) < Math.abs(sum) * 1e-15) break; } return sum * Math.exp(-x + a * Math.log(x) - gammaln(a)); }
  let b = x + 1 - a, c = 1 / 1e-300, d = 1 / b, hh = d; for (let i = 1; i < 500; i++) { const an = -i * (i - a); b += 2; d = an * d + b; if (Math.abs(d) < 1e-300) d = 1e-300; c = b + an / c; if (Math.abs(c) < 1e-300) c = 1e-300; d = 1 / d; const del = d * c; hh *= del; if (Math.abs(del - 1) < 1e-15) break; }
  return 1 - Math.exp(-x + a * Math.log(x) - gammaln(a)) * hh;
}
function chisqCdf(x, k) { return lowRegGamma(k / 2, x / 2); }
function chisqPdf(x, k) { if (x < 0) return 0; if (x === 0) return k === 2 ? 0.5 : 0; return Math.exp((k / 2 - 1) * Math.log(x) - x / 2 - (k / 2) * Math.log(2) - gammaln(k / 2)); }
function invert(cdf, p, lo, hi) { for (let i = 0; i < 200; i++) { const m = (lo + hi) / 2; if (cdf(m) < p) lo = m; else hi = m; if (hi - lo < 1e-12 * Math.max(1, m)) break; } return (lo + hi) / 2; }
function normsinv(p) { return FN['NORM.S.INV'] ? N(FN['NORM.S.INV'].f([p])) : invert(z => 0.5 * (1 + erf(z / Math.SQRT2)), p, -40, 40); }
function erf(x) { const t = 1 / (1 + 0.5 * Math.abs(x)); const y = 1 - t * Math.exp(-x * x - 1.26551223 + t * (1.00002368 + t * (0.37409196 + t * (0.09678418 + t * (-0.18628806 + t * (0.27886807 + t * (-1.13520398 + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277))))))))); return x >= 0 ? y : -y; }
function tinv2(p, df) { return Math.abs(N(FN['T.INV'].f([p / 2, df]))); }
function lnComb(n, k) { return gammaln(n + 1) - gammaln(k + 1) - gammaln(n - k + 1); }

const STAT = 'Thống kê', LOOK = 'Tra cứu & Tham chiếu', WEB = 'Web', TEXT = 'Văn bản';
function def(name, cat, sig, desc, f) { defFn(name, f); if (registerMeta) registerMeta(name, cat, sig, desc); }

def('CHISQ.DIST', STAT, 'x, deg_freedom, cumulative', 'Phân phối chi bình phương (đuôi trái).', a => { const x = N(a[0]), k = Math.trunc(N(a[1])); if (x < 0 || k < 1 || k > 1e10) throw E.NUM; return a[2] && N(a[2]) ? chisqCdf(x, k) : chisqPdf(x, k); });
def('CHISQ.INV', STAT, 'probability, deg_freedom', 'Nghịch đảo phân phối chi bình phương (đuôi trái).', a => { const p = N(a[0]), k = Math.trunc(N(a[1])); if (p < 0 || p > 1 || k < 1) throw E.NUM; if (p === 0) return 0; return invert(x => chisqCdf(x, k), p, 0, Math.max(1000, k * 20)); });
const confNorm = a => { const al = N(a[0]), sd = N(a[1]), n = Math.trunc(N(a[2])); if (al <= 0 || al >= 1 || sd <= 0 || n < 1) throw E.NUM; return -normsinv(al / 2) * sd / Math.sqrt(n); };
def('CONFIDENCE.NORM', STAT, 'alpha, standard_dev, size', 'Khoảng tin cậy cho trung bình (phân phối chuẩn).', confNorm);
def('CONFIDENCE', 'Tương thích', 'alpha, standard_dev, size', 'Khoảng tin cậy (tương thích Excel 2007).', confNorm);
def('CONFIDENCE.T', STAT, 'alpha, standard_dev, size', 'Khoảng tin cậy cho trung bình (phân phối Student t).', a => { const al = N(a[0]), sd = N(a[1]), n = Math.trunc(N(a[2])); if (al <= 0 || al >= 1 || sd <= 0 || n < 1) throw E.NUM; if (n === 1) throw E.DIV0; return tinv2(al, n - 1) * sd / Math.sqrt(n); });
def('BINOM.DIST.RANGE', STAT, 'trials, probability_s, number_s, [number_s2]', 'Xác suất số lần thành công nằm trong khoảng.', a => { const n = Math.trunc(N(a[0])), p = N(a[1]), s1 = Math.trunc(N(a[2])), s2 = a[3] == null ? s1 : Math.trunc(N(a[3])); if (n < 0 || p < 0 || p > 1 || s1 < 0 || s1 > n || s2 < s1 || s2 > n) throw E.NUM; let t = 0; for (let k = s1; k <= s2; k++) t += Math.exp(lnComb(n, k) + (k ? k * Math.log(p) : 0) + (n - k ? (n - k) * Math.log(1 - p) : 0)); return t; });

// ---- FORECAST.ETS family: additive error / additive trend / additive seasonality (AAA) like Excel
function etsPrep(c, values, timeline, completion, aggregation) {
  const v = flat(c, values).map(x => (typeof x === 'number' ? x : null)); const t = flat(c, timeline).map(x => (typeof x === 'number' ? x : null));
  if (v.length !== t.length || v.length < 3) throw E.NA;
  const pairs = t.map((x, i) => [x, v[i]]).filter(p => p[0] != null).sort((p, q) => p[0] - q[0]);
  const steps = []; for (let i = 1; i < pairs.length; i++) steps.push(pairs[i][0] - pairs[i - 1][0]);
  const step = steps.filter(s => s > 0).reduce((m, s) => Math.min(m, s), Infinity); if (!isFinite(step) || step <= 0) throw E.NUM;
  // aggregate duplicates (AVERAGE default) and fill gaps (linear interpolation, or zeros when completion = 0)
  const map = new Map(); pairs.forEach(([x, y]) => { const k = Math.round((x - pairs[0][0]) / step); if (!map.has(k)) map.set(k, []); if (y != null) map.get(k).push(y); });
  const last = Math.max.apply(null, Array.from(map.keys())); const ys = [];
  const agg = list => { if (!list.length) return null; switch (aggregation || 1) { case 2: return list.length; case 3: return list.length; case 4: return Math.max.apply(null, list); case 5: return list.slice().sort((a, b) => a - b)[list.length >> 1]; case 6: return Math.min.apply(null, list); case 7: return list.reduce((s, x) => s + x, 0); default: return list.reduce((s, x) => s + x, 0) / list.length; } };
  for (let k = 0; k <= last; k++) ys.push(map.has(k) ? agg(map.get(k)) : null);
  for (let i = 0; i < ys.length; i++) if (ys[i] == null) { if (completion === 0) { ys[i] = 0; continue; } let j = i; while (j < ys.length && ys[j] == null) j++; const a = ys[i - 1] != null ? ys[i - 1] : ys[j], b = ys[j] != null ? ys[j] : a; for (let q = i; q < j; q++) ys[q] = a + (b - a) * (q - i + 1) / (j - i + 1); i = j; }
  return { ys, t0: pairs[0][0], step };
}
function detectSeason(ys0) { // autocorrelation peak on the de-trended series (Excel caps at 8760)
  const n = ys0.length; if (n < 6) return 1;
  const xm = (n - 1) / 2, ym = ys0.reduce((s, x) => s + x, 0) / n; let sxy = 0, sxx = 0; ys0.forEach((y, i) => { sxy += (i - xm) * (y - ym); sxx += (i - xm) * (i - xm); }); const sl = sxx ? sxy / sxx : 0;
  const ys = ys0.map((y, i) => y - (ym + sl * (i - xm))); const mean = ys.reduce((s, x) => s + x, 0) / n; const den = ys.reduce((s, x) => s + (x - mean) * (x - mean), 0) || 1;
  const maxLag = Math.min(Math.floor(n / 2), 8760); const R = [1];
  for (let lag = 1; lag <= maxLag; lag++) { let r = 0; for (let i = lag; i < n; i++) r += ys[i] * ys[i - lag]; R.push(r / den); }
  let k = 1; while (k <= maxLag && R[k] > 0) k++;                // skip the short-lag correlation of a smooth series
  let best = 1, bestR = 0.15; for (let lag = Math.max(2, k); lag <= maxLag; lag++) { if (R[lag] > bestR && R[lag] >= (R[lag - 1] || 0) && R[lag] >= (R[lag + 1] == null ? -1 : R[lag + 1])) { bestR = R[lag]; best = lag; } }
  return best;
}
function holtWinters(ys, m, al, be, ga) {
  const n = ys.length; let L, T; const S = new Array(n + m).fill(0);
  if (m > 1 && n >= 2 * m) { const avg1 = ys.slice(0, m).reduce((s, x) => s + x, 0) / m, avg2 = ys.slice(m, 2 * m).reduce((s, x) => s + x, 0) / m; T = (avg2 - avg1) / m; L = avg1 + T * (m - 1) / 2; for (let i = 0; i < m; i++) S[i] = ys[i] - (avg1 + T * (i - (m - 1) / 2)); }
  else { m = 1; L = ys[0]; T = n > 1 ? ys[1] - ys[0] : 0; }
  let sse = 0, cnt = 0; const fitted = [];
  for (let t = (m > 1 ? m : 1); t < n; t++) { const sIdx = m > 1 ? t - m : 0; const f = L + T + (m > 1 ? S[sIdx] : 0); fitted.push(f); const e = ys[t] - f; sse += e * e; cnt++; const Lp = L; L = al * (ys[t] - (m > 1 ? S[sIdx] : 0)) + (1 - al) * (L + T); T = be * (L - Lp) + (1 - be) * T; if (m > 1) S[t] = ga * (ys[t] - L) + (1 - ga) * S[sIdx]; }
  return { L, T, S, m, n, sse, cnt, rmse: Math.sqrt(sse / Math.max(1, cnt)), mae: 0 };
}
function fitEts(ys, m) {
  let best = null; const grid = [0.05, 0.15, 0.3, 0.5, 0.7, 0.9];
  for (const a of grid) for (const b of [0.01, 0.05, 0.15, 0.3]) for (const g of (m > 1 ? [0.05, 0.2, 0.4, 0.7] : [0])) { const r = holtWinters(ys, m, a, b, g); if (!best || r.sse < best.sse) best = Object.assign(r, { al: a, be: b, ga: g }); }
  // local refinement
  for (let it = 0; it < 3; it++) { const st = 0.04 / (it + 1); for (const [da, db, dg] of [[st, 0, 0], [-st, 0, 0], [0, st, 0], [0, -st, 0], [0, 0, st], [0, 0, -st]]) { const a = best.al + da, b = best.be + db, g = best.ga + dg; if (a <= 0 || a >= 1 || b < 0 || b >= 1 || g < 0 || g >= 1) continue; const r = holtWinters(ys, m, a, b, g); if (r.sse < best.sse) best = Object.assign(r, { al: a, be: b, ga: g }); } }
  return best;
}
function etsModel(a, c, sIdx) {
  const seasonArg = a[sIdx] == null ? 1 : Math.trunc(N(a[sIdx])); const completion = a[sIdx + 1] == null ? 1 : Math.trunc(N(a[sIdx + 1])); const aggregation = a[sIdx + 2] == null ? 1 : Math.trunc(N(a[sIdx + 2]));
  const { ys, t0, step } = etsPrep(c, a[1], a[2], completion, aggregation);
  let m = seasonArg === 1 ? detectSeason(ys) : seasonArg === 0 ? 1 : seasonArg; if (m < 1 || m > 8760) throw E.NUM;
  const fit = fitEts(ys, m);
  return { fit, ys, t0, step, m: fit.m };
}
function etsForecast(model, target) { const { fit, ys, t0, step } = model; const h = (target - (t0 + (ys.length - 1) * step)) / step; if (h <= 0) throw E.NUM; const hi = Math.max(1, Math.round(h)); const s = fit.m > 1 ? fit.S[ys.length - fit.m + ((hi - 1) % fit.m)] : 0; return fit.L + h * fit.T + s; }
def('FORECAST.ETS', STAT, 'target_date, values, timeline, [seasonality], [data_completion], [aggregation]', 'Dự báo theo làm trơn hàm mũ (ETS AAA).', (a, x, c) => { const md = etsModel(a, c, 3); const tg = a[0]; if (isArr(tg) || tg instanceof Ref) return arr2(c, tg).map(r => r.map(v => etsForecast(md, N(v)))); return etsForecast(md, N(tg)); });
def('FORECAST.ETS.CONFINT', STAT, 'target_date, values, timeline, [confidence_level], [seasonality], [data_completion], [aggregation]', 'Khoảng tin cậy của giá trị dự báo ETS.', (a, x, c) => { const cl = a[3] == null ? 0.95 : N(a[3]); if (cl <= 0 || cl >= 1) throw E.NUM; const md = etsModel(a, c, 4); const h = Math.max(1, (N(a[0]) - (md.t0 + (md.ys.length - 1) * md.step)) / md.step); const f = md.fit; const v = 1 + (h - 1) * (f.al * f.al * (1 + h * f.be + h * (2 * h - 1) * f.be * f.be / 6)); return normsinv(0.5 + cl / 2) * f.rmse * Math.sqrt(Math.max(1, v)); });
def('FORECAST.ETS.SEASONALITY', STAT, 'values, timeline, [data_completion], [aggregation]', 'Độ dài chu kỳ mùa vụ phát hiện được.', (a, x, c) => { const { ys } = etsPrep(c, a[0], a[1], a[2] == null ? 1 : N(a[2]), a[3] == null ? 1 : N(a[3])); const m = detectSeason(ys); return m === 1 ? 0 : m; });
def('FORECAST.ETS.STAT', STAT, 'values, timeline, statistic_type, [seasonality], [data_completion], [aggregation]', 'Thống kê của mô hình ETS (alpha, beta, gamma, MASE, SMAPE, MAE, RMSE, step).', (a, x, c) => { const md = etsModel([null, a[0], a[1], a[3], a[4], a[5]], c, 3); const k = Math.trunc(N(a[2])); const f = md.fit; const ys = md.ys; let mae = 0, smape = 0, naive = 0; const n = ys.length; for (let i = 1; i < n; i++) naive += Math.abs(ys[i] - ys[i - 1]); naive /= Math.max(1, n - 1); for (let i = 0; i < n; i++) { const fc = i ? ys[i - 1] : ys[0]; mae += Math.abs(ys[i] - fc); smape += Math.abs(ys[i] - fc) / ((Math.abs(ys[i]) + Math.abs(fc)) / 2 || 1); } mae /= n; smape /= n; switch (k) { case 1: return f.al; case 2: return f.be; case 3: return f.ga; case 4: return naive ? f.rmse / naive : 0; case 5: return smape; case 6: return mae; case 7: return f.rmse; case 8: return md.step; } throw E.NUM; });

// ---- PIVOTBY(row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order], [col_total_depth], [col_sort_order], [filter_array], [relative_to])
def('PIVOTBY', LOOK, 'row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order], [col_total_depth], [col_sort_order], [filter_array]', 'Tạo bảng tổng hợp (pivot) theo hàng và cột.', (a, ctx, c) => {
  let rows = arr2(c, a[0]), cols = arr2(c, a[1]), vals = arr2(c, a[2]); const fn = a[3];
  const hdr = a[4] == null ? 0 : Math.trunc(N(a[4])); const rowTot = a[5] == null ? 1 : Math.trunc(N(a[5])), colTot = a[7] == null ? 1 : Math.trunc(N(a[7]));
  const rs = a[6] == null ? 1 : Math.trunc(N(a[6])), cs = a[8] == null ? 1 : Math.trunc(N(a[8]));
  if (hdr === 1 || hdr === 3) { rows = rows.slice(1); cols = cols.slice(1); vals = vals.slice(1); }
  const filt = a[9] != null ? arr2(c, a[9]).map(r => !!r[0]) : null;
  const key = r => r.map(v => str(v)).join(' / ');
  const rk = [], ck = [], cell = new Map(), byR = new Map(), byC = new Map(), all = [];
  rows.forEach((r, i) => { if (filt && !filt[i]) return; const a1 = key(r), b1 = key(cols[i] || ['']); const v = vals[i] ? vals[i][0] : null; if (rk.indexOf(a1) < 0) rk.push(a1); if (ck.indexOf(b1) < 0) ck.push(b1); const k = a1 + '\u0001' + b1; if (!cell.has(k)) cell.set(k, []); cell.get(k).push(v); if (!byR.has(a1)) byR.set(a1, []); byR.get(a1).push(v); if (!byC.has(b1)) byC.set(b1, []); byC.get(b1).push(v); all.push(v); });
  const sorter = dir => (p, q) => (dir < 0 ? -1 : 1) * cmp(isNaN(+p) ? p : +p, isNaN(+q) ? q : +q);
  rk.sort(sorter(rs)); ck.sort(sorter(cs));
  const agg = list => { if (!list || !list.length) return ''; if (fn instanceof Lambda) { const v = c.applyLambda(fn, [list.map(x => [x])], ctx); return v instanceof Ref ? c.deref(v, true) : isArr(v) ? v[0][0] : v; } const n = list.filter(x => typeof x === 'number'); return n.reduce((s, x) => s + x, 0); };
  const out = [[''].concat(ck).concat(colTot ? ['Total'] : [])];
  rk.forEach(r => out.push([r].concat(ck.map(cc => agg(cell.get(r + '\u0001' + cc)))).concat(colTot ? [agg(byR.get(r))] : [])));
  if (rowTot) out.push(['Total'].concat(ck.map(cc => agg(byC.get(cc)))).concat(colTot ? [agg(all)] : []));
  return out;
});
// ---- functions that need Microsoft online services: same behaviour as Excel when offline
const BUSY = errOf('#CONNECT!');
def('STOCKHISTORY', WEB, 'stock, start_date, [end_date], [interval], [headers], [property0]…', 'Lịch sử giá cổ phiếu (cần dịch vụ trực tuyến của Microsoft — ngoại tuyến trả về #CONNECT!).', () => { throw BUSY; });
def('FIELDVALUE', LOOK, 'value, field_name', 'Lấy trường từ kiểu dữ liệu liên kết (Cổ phiếu/Địa lý) — ngoại tuyến trả về #FIELD!.', () => { throw errOf('#FIELD!'); });
def('PY', 'Python', 'python_code, return_type', 'Python trong Excel (chạy trên đám mây Microsoft) — ngoại tuyến trả về #CONNECT!.', () => { throw BUSY; });
def('COPILOT', TEXT, 'prompt_part1, [context1], …', 'Hỏi A.I ngay trong ô. Ngoại tuyến: trả lời bằng bộ A.I tích hợp (tóm tắt / từ khóa / đếm).', (a, x, c) => {
  const parts = a.map(v => (v instanceof Ref || isArr(v)) ? flat(c, v).filter(z => z != null).map(z => str(z)).join(', ') : str(v));
  const prompt = String(parts[0] || '').toLowerCase(); const ctxText = parts.slice(1).join('\n');
  if (/tóm tắt|summar/.test(prompt)) { const sents = (ctxText.match(/[^.!?]+[.!?]*/g) || []).map(x => x.trim()).filter(Boolean); return sents.slice(0, 2).join(' ') || ctxText.slice(0, 200); }
  if (/đếm|count/.test(prompt)) return ctxText ? ctxText.split(/,\s*/).length : 0;
  if (/viết hoa|upper/.test(prompt)) return ctxText.toUpperCase();
  if (/từ khóa|keyword/.test(prompt)) { const f = {}; ctxText.toLowerCase().split(/[^a-zà-ỹđ0-9]+/i).filter(w => w.length > 3).forEach(w => { f[w] = (f[w] || 0) + 1; }); return Object.keys(f).sort((p, q) => f[q] - f[p]).slice(0, 5).join(', '); }
  return (parts[0] || '') + (ctxText ? ': ' + ctxText.slice(0, 200) : '');
});
// v0.0.6 fix: T.INV.2T / TINV are two-tailed (v0.0.5 returned the one-tailed lower quantile, e.g. -1.68 instead of 2.01)
const tinv2t = a => { const p = N(a[0]), df = Math.trunc(N(a[1])); if (p <= 0 || p > 1 || df < 1) throw E.NUM; return tinv2(p, df); };
def('T.INV.2T', STAT, 'probability, deg_freedom', 'Nghịch đảo hai phía của phân phối Student t.', tinv2t);
def('TINV', 'Tương thích', 'probability, deg_freedom', 'Nghịch đảo hai phía của phân phối Student t (tương thích).', tinv2t);

// ---- array criteria lift-up (Excel 365): =COUNTIF(A1:A6, A1:A6) spills one count per criteria cell — used by AutoXL A.COUNTEQ / A.DUPLICATED
(function liftCriteria() {
  const odd = n => i => i % 2 === 1;            // COUNTIFS: criteria at 1,3,5...
  const evenFrom2 = i => i >= 2 && i % 2 === 0; // SUMIFS / AVERAGEIFS / MAXIFS / MINIFS: 2,4,6...
  const spec = { COUNTIF: i => i === 1, SUMIF: i => i === 1, AVERAGEIF: i => i === 1, COUNTIFS: odd(), SUMIFS: evenFrom2, AVERAGEIFS: evenFrom2, MAXIFS: evenFrom2, MINIFS: evenFrom2 };
  Object.keys(spec).forEach(name => {
    const def = FN[name]; if (!def || def._lifted) return;
    const orig = def.f, isCrit = spec[name];
    def.f = function (a, x, c) {
      for (let i = 0; i < a.length; i++) {
        if (!isCrit(i)) continue;
        const v = a[i]; const multi = v instanceof Ref ? v.rows * v.cols > 1 : isArr(v) && (v.length > 1 || v[0].length > 1);
        if (!multi) continue;
        return arr2(c, v).map(row => row.map(cv => { const b = a.slice(); b[i] = cv; try { return def.f(b, x, c); } catch (e) { if (e instanceof Err) return e; throw e; } }));
      }
      return orig(a, x, c);
    };
    def._lifted = true;
  });
})();
