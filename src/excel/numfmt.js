// Excel number format engine (format codes) + input parsing
import { Err, serialToParts, dateToSerial } from './model.js';

export const SEP = { dec: '.', grp: ',' }; // display like Excel en-US (matches the user's VISICONS workbooks)
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const COLORS = { black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00B050', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00' };

const cache = {};
function splitSections(code) {
  const out = []; let cur = '', q = false, br = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"' && !br) q = !q;
    else if (ch === '\\' && !q) { cur += ch + (code[i + 1] || ''); i++; continue; }
    else if (ch === '[' && !q) br = true; else if (ch === ']' && !q) br = false;
    if (ch === ';' && !q && !br) { out.push(cur); cur = ''; } else cur += ch;
  }
  out.push(cur); return out;
}
function compile(code) {
  if (cache[code]) return cache[code];
  const secs = splitSections(code).map(s => {
    const sec = { raw: s, color: null, cond: null, isDate: false, isText: false, pct: 0, sci: false, toks: [] };
    let body = s.replace(/\[([^\]]*)\]/g, (m, inner) => {
      const l = inner.toLowerCase();
      if (COLORS[l]) { sec.color = COLORS[l]; return ''; }
      if (/^color\d+$/.test(l)) return '';
      const cm = /^(<=|>=|<>|<|>|=)(-?[\d.]+)$/.exec(inner); if (cm) { sec.cond = [cm[1], +cm[2]]; return ''; }
      if (/^\$/.test(inner)) { const sym = inner.slice(1).split('-')[0]; return sym ? '"' + sym + '"' : ''; }
      if (/^(h+|m+|s+)$/i.test(inner)) { sec.isDate = true; return '\u0001' + inner.toLowerCase() + '\u0001'; }
      return '';
    });
    sec.body = body;
    // tokenize
    const toks = [];
    for (let i = 0; i < body.length; i++) {
      const ch = body[i];
      if (ch === '"') { let j = body.indexOf('"', i + 1); if (j < 0) j = body.length; toks.push({ t: 'lit', v: body.slice(i + 1, j) }); i = j; continue; }
      if (ch === '\\') { toks.push({ t: 'lit', v: body[i + 1] || '' }); i++; continue; }
      if (ch === '_') { toks.push({ t: 'lit', v: ' ' }); i++; continue; }
      if (ch === '*') { i++; continue; }
      if (ch === '\u0001') { const j = body.indexOf('\u0001', i + 1); toks.push({ t: 'elapsed', v: body.slice(i + 1, j) }); i = j; continue; }
      if (ch === '@') { toks.push({ t: 'text' }); sec.isText = true; continue; }
      const low = body.slice(i).toLowerCase();
      let m = /^(am\/pm|a\/p)/.exec(low); if (m) { toks.push({ t: 'ampm', v: m[1] }); sec.ampm = true; sec.isDate = true; i += m[1].length - 1; continue; }
      m = /^(yyyy|yy|e|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s)/.exec(low);
      if (m && !/[0#?]/.test(ch)) { toks.push({ t: 'dt', v: m[1] }); sec.isDate = true; i += m[1].length - 1; continue; }
      if (/[0#?.,%Ee+\-]/.test(ch)) {
        if ((ch === 'E' || ch === 'e') && /^[eE][+-]/.test(body.slice(i))) { toks.push({ t: 'num', v: body.slice(i, i + 2) }); sec.sci = true; i++; continue; }
        if (ch === '%') { sec.pct++; toks.push({ t: 'lit', v: '%' }); continue; }
        if (ch === '-' || ch === '+') { toks.push({ t: 'lit', v: ch }); continue; }
        toks.push({ t: 'num', v: ch }); continue;
      }
      if (ch === '/' && toks.length && toks[toks.length - 1].t === 'num' && /[0#?]/.test(body[i + 1] || '')) { toks.push({ t: 'num', v: '/' }); sec.frac = true; continue; }
      toks.push({ t: 'lit', v: ch });
    }
    // minutes vs month: 'm' after h or before s => minutes
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i]; if (t.t !== 'dt' || (t.v !== 'm' && t.v !== 'mm')) continue;
      let prev = null, next = null;
      for (let j = i - 1; j >= 0; j--) if (toks[j].t === 'dt') { prev = toks[j]; break; }
      for (let j = i + 1; j < toks.length; j++) if (toks[j].t === 'dt') { next = toks[j]; break; }
      if ((prev && /^h/.test(prev.v)) || (next && /^s/.test(next.v))) t.v = t.v === 'm' ? 'M' : 'MM';
    }
    // numeric pattern analysis
    if (!sec.isDate && !sec.isText) {
      const pat = toks.filter(t => t.t === 'num').map(t => t.v).join('');
      let p = pat; let scale = 0;
      while (/,$/.test(p) || /,\./.test(p)) { if (/,$/.test(p)) { p = p.slice(0, -1); scale++; } else { p = p.replace(/,\./, '.'); scale++; } }
      sec.scale = scale;
      const ePos = p.search(/[eE]/);
      const mant = ePos >= 0 ? p.slice(0, ePos) : p;
      const dot = mant.indexOf('.');
      const intP = dot >= 0 ? mant.slice(0, dot) : mant, decP = dot >= 0 ? mant.slice(dot + 1) : '';
      sec.grp = intP.indexOf(',') >= 0;
      sec.intMin = (intP.replace(/,/g, '').match(/0/g) || []).length;
      sec.decMin = (decP.match(/0/g) || []).length; sec.decMax = (decP.match(/[0#?]/g) || []).length;
      sec.expDigits = ePos >= 0 ? (p.slice(ePos + 2).match(/0/g) || []).length : 0;
      sec.hasNum = /[0#?]/.test(pat);
    }
    sec.toks = toks;
    return sec;
  });
  return (cache[code] = secs);
}

function fmtNumberSec(v, sec) {
  let n = Math.abs(v) * Math.pow(100, sec.pct) / Math.pow(1000, sec.scale);
  let numStr;
  if (sec.sci) {
    let e = n === 0 ? 0 : Math.floor(Math.log(n) / Math.LN10);
    let m = n / Math.pow(10, e);
    m = +m.toFixed(sec.decMax);
    if (m >= 10) { m /= 10; e++; }
    numStr = m.toFixed(sec.decMax) + 'E' + (e < 0 ? '-' : '+') + pad(Math.abs(e), Math.max(2, sec.expDigits));
  } else if (sec.frac) {
    const w = Math.floor(n), f = n - w; let best = [0, 1], err = 1;
    for (let d = 1; d <= 99; d++) { const nn = Math.round(f * d); const e2 = Math.abs(f - nn / d); if (e2 < err - 1e-12) { err = e2; best = [nn, d]; } }
    numStr = (w ? w + ' ' : '') + (best[0] ? best[0] + '/' + best[1] : (w ? '' : '0'));
  } else {
    let s = roundTo(n, sec.decMax).toFixed(sec.decMax);
    let [ip, dp] = s.split('.');
    if (dp) { dp = dp.replace(/0+$/, ''); while (dp.length < sec.decMin) dp += '0'; }
    if (ip === '0' && sec.intMin === 0) ip = '';
    while (ip.length < sec.intMin) ip = '0' + ip;
    if (sec.grp) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, SEP.grp);
    numStr = ip + (dp ? SEP.dec + dp : (sec.decMax && sec.decMin ? SEP.dec : ''));
  }
  // rebuild with literals: place number at first num token, drop the other num tokens
  let out = '', placed = false;
  for (const t of sec.toks) {
    if (t.t === 'num') { if (!placed) { out += numStr; placed = true; } }
    else if (t.t === 'lit') out += t.v;
    else if (t.t === 'text') out += '';
  }
  if (!placed) out += sec.hasNum === false ? '' : numStr;
  return out;
}
function roundTo(n, d) { const f = Math.pow(10, d); return Math.round((n + Number.EPSILON) * f) / f; }
function pad(n, w) { n = String(n); while (n.length < w) n = '0' + n; return n; }

function fmtDateSec(v, sec) {
  if (v < 0) return '#'.repeat(10);
  const p = serialToParts(v);
  const fracDays = v;
  let out = '';
  const decSec = (() => { for (let i = 0; i < sec.toks.length; i++) if (sec.toks[i].t === 'num' && sec.toks[i].v === '.') { let n = 0; for (let j = i + 1; j < sec.toks.length && sec.toks[j].t === 'num' && sec.toks[j].v === '0'; j++) n++; return n; } return 0; })();
  let skipNum = false;
  for (const t of sec.toks) {
    if (t.t === 'lit') { out += t.v; continue; }
    if (t.t === 'num') { if (t.v === '.' && decSec) { out += SEP.dec + pad(Math.floor(p.ms / Math.pow(10, 3 - decSec)), decSec); skipNum = true; } else if (!skipNum) out += t.v; continue; }
    if (t.t === 'elapsed') { const tot = fracDays * 24; out += t.v[0] === 'h' ? pad(Math.floor(tot), t.v.length) : t.v[0] === 'm' ? pad(Math.floor(tot * 60), t.v.length) : pad(Math.floor(tot * 3600), t.v.length); continue; }
    if (t.t === 'ampm') { const pm = p.H >= 12; out += t.v === 'a/p' ? (pm ? 'p' : 'a') : (pm ? 'PM' : 'AM'); continue; }
    if (t.t !== 'dt') continue;
    let H = p.H; if (sec.ampm) { H = H % 12; if (H === 0) H = 12; }
    switch (t.v) {
      case 'yyyy': case 'e': out += p.y; break;
      case 'yy': out += pad(p.y % 100, 2); break;
      case 'mmmmm': out += MONTHS[p.m - 1][0]; break;
      case 'mmmm': out += MONTHS[p.m - 1]; break;
      case 'mmm': out += MONTHS[p.m - 1].slice(0, 3); break;
      case 'mm': out += pad(p.m, 2); break;
      case 'm': out += p.m; break;
      case 'dddd': out += DAYS[p.dow]; break;
      case 'ddd': out += DAYS[p.dow].slice(0, 3); break;
      case 'dd': out += pad(p.d, 2); break;
      case 'd': out += p.d; break;
      case 'hh': out += pad(H, 2); break;
      case 'h': out += H; break;
      case 'MM': out += pad(p.M, 2); break;
      case 'M': out += p.M; break;
      case 'ss': out += pad(p.S, 2); break;
      case 's': out += p.S; break;
    }
  }
  return out;
}

export function generalNumber(v) {
  if (!isFinite(v)) return '#NUM!';
  if (v === 0) return '0';
  const a = Math.abs(v);
  if (a >= 1e11 || a < 1e-9) { let s = v.toExponential(5).replace(/\.?0+e/, 'e').replace('e', 'E'); s = s.replace(/E([+-])(\d)$/, 'E$10$2'); return s.replace('.', SEP.dec); }
  let s = String(+v.toPrecision(11));
  if (/e/.test(s)) s = v.toFixed(10).replace(/\.?0+$/, '');
  return s.replace('.', SEP.dec);
}

// format(value, code) -> {t: text, c: color|null}
export function format(v, code) {
  if (v == null) return { t: '', c: null };
  if (v instanceof Err) return { t: v.code, c: null };
  if (typeof v === 'boolean') return { t: v ? 'TRUE' : 'FALSE', c: null };
  if (!code || code === 'General' || code === 'general') {
    if (typeof v === 'number') return { t: generalNumber(v), c: null };
    return { t: String(v), c: null };
  }
  const secs = compile(code);
  if (typeof v === 'string') {
    const ts = secs.length >= 4 ? secs[3] : secs.find(s => s.isText);
    if (!ts) return { t: v, c: null };
    let out = ''; for (const t of ts.toks) out += t.t === 'text' ? v : t.t === 'lit' ? t.v : ''; return { t: out, c: ts.color };
  }
  if (typeof v !== 'number') return { t: String(v), c: null };
  let sec, neg = v < 0;
  if (secs.some(s => s.cond)) {
    sec = secs.find(s => s.cond && test(v, s.cond)) || secs.find(s => !s.cond && !s.isText) || secs[secs.length - 1];
    neg = neg && !(sec.cond);
  } else if (secs.length === 1 || (secs.length >= 2 && secs[1].isText)) { sec = secs[0]; }
  else if (v > 0 || (v === 0 && secs.length < 3)) sec = secs[0];
  else if (v < 0) { sec = secs[1]; neg = false; }
  else sec = secs[2] || secs[0];
  if (sec.isDate) return { t: fmtDateSec(v, sec), c: sec.color };
  if (!sec.hasNum && !sec.isText) { let o = ''; for (const t of sec.toks) if (t.t === 'lit') o += t.v; return { t: o, c: sec.color }; }
  let t = fmtNumberSec(v, sec);
  if (neg && v < 0 && secs.indexOf(sec) === 0) t = '-' + t;
  return { t, c: sec.color };
}
function test(v, c) { const n = c[1]; switch (c[0]) { case '<': return v < n; case '>': return v > n; case '<=': return v <= n; case '>=': return v >= n; case '=': return v === n; case '<>': return v !== n; } return false; }
export function isDateFormat(code) { if (!code || code === 'General') return false; return compile(code).some(s => s.isDate); }

// ---------------------------------------------------------------- input parsing
// returns {v, nf?} — nf is an inferred number format for the cell
export function parseInput(s) {
  if (s == null) return { v: null };
  const t = String(s);
  const tr = t.trim();
  if (tr === '') return { v: null };
  if (/^(true|false)$/i.test(tr)) return { v: /^true$/i.test(tr) };
  if (/^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!)$/i.test(tr)) return { v: new Err(tr.toUpperCase()) };
  // percent
  let m = /^([+-]?\d+(?:[.,]\d+)?)\s*%$/.exec(tr);
  if (m) { const n = parseFloat(m[1].replace(',', '.')) / 100; return { v: n, nf: /[.,]/.test(m[1]) ? '0.00%' : '0%' }; }
  // currency
  m = /^([$€£¥₫])\s*([+-]?[\d,]+(?:\.\d+)?)$/.exec(tr) || /^([+-]?[\d.,]+)\s*(₫|đ|VND|vnđ)$/i.exec(tr);
  if (m) {
    if (/^[$€£¥₫]$/.test(m[1])) { const n = +m[2].replace(/,/g, ''); if (!isNaN(n)) return { v: n, nf: (m[1] === '$' ? '"$"' : '"' + m[1] + '"') + '#,##0' + (/\./.test(m[2]) ? '.00' : '') }; }
    else { const n = +m[1].replace(/[.,]/g, ''); if (!isNaN(n)) return { v: n, nf: '#,##0 "₫"' }; }
  }
  // thousands separators 1,234,567.89
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(tr)) return { v: +tr.replace(/,/g, ''), nf: /\./.test(tr) ? '#,##0.00' : '#,##0' };
  // plain number / scientific
  if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(tr)) { if (/^0\d/.test(tr) && tr.length > 1 && !/\./.test(tr)) return { v: tr }; return { v: +tr, nf: /e/i.test(tr) ? '0.00E+00' : undefined }; }
  // negative in parentheses
  if (/^\(\d+(\.\d+)?\)$/.test(tr)) return { v: -tr.slice(1, -1) };
  // dates: d/m/yyyy (Vietnamese order), yyyy-mm-dd, d-mmm-yyyy
  m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(tr);
  if (m) {
    let d = +m[1], mo = +m[2], y = +m[3]; if (y < 100) y += y < 30 ? 2000 : 1900;
    if (mo > 12 && d <= 12) { const x = d; d = mo; mo = x; }
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
      let v = dateToSerial(y, mo, d); let nf = 'dd/mm/yyyy';
      if (m[4]) { v += (+m[4] * 3600 + +m[5] * 60 + (+m[6] || 0)) / 86400; nf = 'dd/mm/yyyy hh:mm'; }
      return { v, nf };
    }
  }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(tr);
  if (m) { let v = dateToSerial(+m[1], +m[2], +m[3]); let nf = 'yyyy-mm-dd'; if (m[4]) { v += (+m[4] * 3600 + +m[5] * 60 + (+m[6] || 0)) / 86400; nf = 'yyyy-mm-dd hh:mm'; } return { v, nf }; }
  m = /^(\d{1,2})[- ]([A-Za-z]{3})[a-z]*[- ](\d{2,4})$/.exec(tr);
  if (m) { const mi = MONTHS.findIndex(x => x.slice(0, 3).toLowerCase() === m[2].toLowerCase()); if (mi >= 0) { let y = +m[3]; if (y < 100) y += 2000; return { v: dateToSerial(y, mi + 1, +m[1]), nf: 'd-mmm-yy' }; } }
  // time
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm|sa|ch)?$/i.exec(tr);
  if (m) { let H = +m[1]; const ap = (m[4] || '').toLowerCase(); if ((ap === 'pm' || ap === 'ch') && H < 12) H += 12; if ((ap === 'am' || ap === 'sa') && H === 12) H = 0; if (H < 24 && +m[2] < 60) return { v: (H * 3600 + +m[2] * 60 + (+m[3] || 0)) / 86400, nf: m[4] ? 'h:mm AM/PM' : (m[3] ? 'hh:mm:ss' : 'hh:mm') }; }
  return { v: t };
}

// Built-in format catalogue for Format Cells dialog / ribbon
export const NUMBER_FORMATS = [
  ['General', 'Chung (General)'], ['0', 'Số nguyên'], ['0.00', 'Số (2 chữ số thập phân)'], ['#,##0', 'Số có phân cách hàng nghìn'], ['#,##0.00', 'Số có phân cách, 2 thập phân'],
  ['#,##0 "₫"', 'Tiền tệ (₫)'], ['"$"#,##0.00', 'Tiền tệ ($)'], ['_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)', 'Kế toán (Accounting)'],
  ['dd/mm/yyyy', 'Ngày ngắn (dd/mm/yyyy)'], ['dddd, d mmmm yyyy', 'Ngày dài'], ['hh:mm:ss', 'Thời gian'], ['0%', 'Phần trăm'], ['0.00%', 'Phần trăm (2 số lẻ)'],
  ['# ?/?', 'Phân số'], ['0.00E+00', 'Khoa học (Scientific)'], ['@', 'Văn bản (Text)']
];
