// Formula engine: tokenizer, Pratt parser, evaluator, reference shifting, calculation (with dynamic-array spill)
import { Err, E, errOf, isErr, colIndex, colName, key, MAX_ROWS, MAX_COLS, parseAddr } from './model.js';
import { generalNumber, parseInput } from './numfmt.js';

// ------------------------------------------------------------------ tokenizer
const RE_ERR = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)/i;
export function tokenize(src) {
  const toks = []; let i = 0; const n = src.length;
  const push = (t, v, extra) => toks.push(Object.assign({ t, v, p: i }, extra || {}));
  while (i < n) {
    const ch = src[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') { i++; continue; }
    if (ch === '"') { let j = i + 1, s = ''; while (j < n) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; } push('str', s); i = j + 1; continue; }
    if (ch === '#') { const m = RE_ERR.exec(src.slice(i)); if (m) { push('err', m[0].toUpperCase()); i += m[0].length; continue; } }
    // sheet-qualified reference: 'My Sheet'!A1 or Sheet1!A1:B2 or [Book.xlsx]Sheet1!A1
    let sheet = null, j = i;
    if (ch === "'") { let k = i + 1, s = ''; while (k < n) { if (src[k] === "'") { if (src[k + 1] === "'") { s += "'"; k += 2; continue; } break; } s += src[k++]; } if (src[k + 1] === '!') { sheet = s; j = k + 2; } }
    else { const m = /^((?:\[[^\]\r\n]+\])?[A-Za-z_À-ỹ][\w.À-ỹ]*(?::(?:\[[^\]\r\n]+\])?[A-Za-z_À-ỹ][\w.À-ỹ]*)?)!/.exec(src.slice(i)); if (m) { sheet = m[1]; j = i + m[0].length; } }
    const rest = src.slice(j);
    let m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?(?![\w(])/.exec(rest);
    if (m && colIndex(m[2]) < MAX_COLS && +m[4] >= 1) {
      const r1 = +m[4] - 1, c1 = colIndex(m[2]);
      const ref = { sheet, r1, c1, r2: r1, c2: c1, ar1: !!m[3], ac1: !!m[1], ar2: !!m[3], ac2: !!m[1], kind: 'cell' };
      if (m[6]) { ref.r2 = +m[8] - 1; ref.c2 = colIndex(m[6]); ref.ar2 = !!m[7]; ref.ac2 = !!m[5]; ref.kind = 'range'; }
      i = j + m[0].length; push('ref', ref); continue;
    }
    m = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/.exec(rest);
    if (m) {
      const c1 = colIndex(m[2]), c2 = colIndex(m[4]);
      if (c1 < MAX_COLS && c2 < MAX_COLS) { i = j + m[0].length; push('ref', { sheet, r1: 0, c1, r2: MAX_ROWS - 1, c2, ar1: true, ar2: true, ac1: !!m[1], ac2: !!m[3], kind: 'col' }); continue; }
    }
    m = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\w(])/.exec(rest);
    if (m && (sheet || toks.length === 0 || /^(op|lp|comma|semi)$/.test(toks[toks.length - 1].t))) { i = j + m[0].length; push('ref', { sheet, r1: +m[2] - 1, c1: 0, r2: +m[4] - 1, c2: MAX_COLS - 1, ar1: !!m[1], ar2: !!m[3], ac1: true, ac2: true, kind: 'row' }); continue; }
    if (sheet) { // Sheet!Name
      const mm = /^[A-Za-z_\\][\w.]*/.exec(rest); if (mm) { push('name', mm[0], { sheet }); i = j + mm[0].length; continue; }
    }
    m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i));
    if (m) { push('num', +m[0]); i += m[0].length; continue; }
    m = /^([A-Za-z_\\À-ỹ][\w.À-ỹ]*)(\s*\()?/.exec(src.slice(i));
    if (m) {
      const word = m[1];
      // structured reference Table1[Col]
      if (src[i + word.length] === '[') {
        let depth = 0, k = i + word.length; for (; k < n; k++) { if (src[k] === '[') depth++; else if (src[k] === ']') { depth--; if (!depth) break; } }
        push('sref', word, { spec: src.slice(i + word.length + 1, k) }); i = k + 1; continue;
      }
      if (m[2]) { push('fn', word.toUpperCase()); i += m[0].length; continue; }
      if (/^(TRUE|FALSE)$/i.test(word)) { push('bool', /^TRUE$/i.test(word)); i += word.length; continue; }
      push('name', word); i += word.length; continue;
    }
    const two = src.substr(i, 2);
    if (two === '<=' || two === '>=' || two === '<>') { push('op', two); i += 2; continue; }
    if ('+-*/^&=<>%:'.indexOf(ch) >= 0) { push('op', ch); i++; continue; }
    if (ch === '(') { push('lp'); i++; continue; }
    if (ch === ')') { push('rp'); i++; continue; }
    if (ch === ',') { push('comma'); i++; continue; }
    if (ch === ';') { push('semi'); i++; continue; }
    if (ch === '{') { push('lb'); i++; continue; }
    if (ch === '}') { push('rb'); i++; continue; }
    if (ch === '@') { i++; continue; } // implicit intersection operator: ignore
    throw new Err('#NAME?');
  }
  return toks;
}

// ------------------------------------------------------------------ parser
const BIN = { ':': 9, '^': 6, '*': 5, '/': 5, '+': 4, '-': 4, '&': 3, '=': 2, '<>': 2, '<': 2, '>': 2, '<=': 2, '>=': 2 };
export function parse(src) {
  const toks = tokenize(src); let p = 0;
  const peek = () => toks[p]; const next = () => toks[p++];
  function expr(minPrec) {
    let left = unary();
    for (;;) {
      const t = peek();
      if (!t) break;
      if (t.t === 'op' && t.v === '%') { next(); left = { t: 'pct', a: left }; continue; }
      if (t.t !== 'op' || BIN[t.v] === undefined || BIN[t.v] < minPrec) break;
      next();
      const prec = BIN[t.v];
      const right = expr(t.v === '^' ? prec : prec + 1);
      left = t.v === ':' ? { t: 'range', a: left, b: right } : { t: 'bin', op: t.v, a: left, b: right };
    }
    return left;
  }
  function unary() {
    const t = peek();
    if (t && t.t === 'op' && (t.v === '-' || t.v === '+')) { next(); const a = expr(7); return t.v === '-' ? { t: 'neg', a } : a; }
    return primary();
  }
  function primary() {
    let node = primary0();
    while (peek() && peek().t === 'lp' && (node.t === 'fn' || node.t === 'paren' || node.t === 'name')) { next(); node = { t: 'invoke', f: node, args: callArgs() }; }
    return node;
  }
  function primary0() {
    const t = next();
    if (!t) throw E.VALUE;
    switch (t.t) {
      case 'num': return { t: 'num', v: t.v };
      case 'str': return { t: 'str', v: t.v };
      case 'bool': return { t: 'bool', v: t.v };
      case 'err': return { t: 'err', v: errOf(t.v) };
      case 'ref': return { t: 'ref', ref: t.v };
      case 'sref': return { t: 'sref', table: t.v, spec: t.spec };
      case 'name': return { t: 'name', n: t.v, sheet: t.sheet || null };
      case 'lp': { const e = expr(0); if (!peek() || peek().t !== 'rp') throw E.VALUE; next(); return { t: 'paren', a: e }; }
      case 'lb': {
        const rows = [[]];
        for (;;) {
          const x = peek(); if (!x) throw E.VALUE;
          if (x.t === 'rb') { next(); break; }
          if (x.t === 'comma') { next(); continue; }
          if (x.t === 'semi') { next(); rows.push([]); continue; }
          const neg = x.t === 'op' && x.v === '-'; if (neg) next();
          const v = next(); let val = v.t === 'num' ? v.v : v.t === 'str' ? v.v : v.t === 'bool' ? v.v : v.t === 'err' ? errOf(v.v) : null;
          if (neg && typeof val === 'number') val = -val;
          rows[rows.length - 1].push(val);
        }
        return { t: 'arr', v: rows };
      }
      case 'fn': {
        const args = [];
        if (peek() && peek().t === 'rp') { next(); return { t: 'fn', n: t.v, args }; }
        for (;;) {
          const x = peek();
          if (x && (x.t === 'comma' || x.t === 'rp')) args.push({ t: 'missing' });
          else args.push(expr(0));
          const sep = next();
          if (!sep) throw E.VALUE;
          if (sep.t === 'rp') break;
          if (sep.t !== 'comma') throw E.VALUE;
          if (peek() && peek().t === 'rp') { next(); args.push({ t: 'missing' }); break; }
        }
        return { t: 'fn', n: t.v, args };
      }
    }
    throw E.VALUE;
  }
  const ast = expr(0);
  if (p < toks.length) throw E.VALUE;
  return ast;
  function callArgs() {
    const args = [];
    if (peek() && peek().t === 'rp') { next(); return args; }
    for (;;) { const x = peek(); if (x && (x.t === 'comma' || x.t === 'rp')) args.push({ t: 'missing' }); else args.push(expr(0)); const sep = next(); if (!sep) throw E.VALUE; if (sep.t === 'rp') break; if (sep.t !== 'comma') throw E.VALUE; }
    return args;
  }
  return ast;
}

// ------------------------------------------------------------------ values
export class Ref { constructor(sheet, r1, c1, r2, c2, sheets) { this.sheet = sheet; this.r1 = r1; this.c1 = c1; this.r2 = r2; this.c2 = c2; this.sheets = sheets || (sheet ? [sheet] : []); } get rows() { return this.r2 - this.r1 + 1; } get cols() { return this.c2 - this.c1 + 1; } }
export class Lambda { constructor(params, body, env) { this.params = params; this.body = body; this.env = env; } }
export function isArr(v) { return Array.isArray(v); }

// ------------------------------------------------------------------ coercion helpers (exported for functions)
export function num(v) {
  if (typeof v === 'number') return v;
  if (v == null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof Err) return v;
  if (typeof v === 'string') { if (v.trim() === '') return E.VALUE; const p = parseInput(v); return typeof p.v === 'number' ? p.v : E.VALUE; }
  return E.VALUE;
}
export function str(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return generalNumber(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof Err) return v;
  return String(v);
}
export function bool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v == null) return false;
  if (v instanceof Err) return v;
  if (typeof v === 'string') { const u = v.toUpperCase(); if (u === 'TRUE') return true; if (u === 'FALSE') return false; return E.VALUE; }
  return E.VALUE;
}
const typeRank = v => typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 0;
export function cmp(a, b) {
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ta = typeRank(a), tb = typeRank(b);
  if (ta !== tb) return ta - tb;
  if (ta === 2) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  if (ta === 3) return (a ? 1 : 0) - (b ? 1 : 0);
  return a < b ? -1 : a > b ? 1 : 0;
}

// ------------------------------------------------------------------ function registry
export const FN = {};
export function defFn(name, f, opts) { FN[name] = Object.assign({ f }, opts || {}); }

// ------------------------------------------------------------------ evaluation
export class Calc {
  constructor(wb) { this.wb = wb; this.stack = []; this.state = new Map(); this.circular = null; this.pass = 0; this.volatileHit = false; }
  sheetOf(name, ctx) {
    if (!name) return ctx.sheet;
    let targetWb = this.wb;
    let cleanName = name.replace(/^'|'$/g, '').trim();
    // Support cross-workbook references [Book.xlsx]Sheet or 'C:\path\[Book.xlsx]Sheet'
    const wbMatch = /^(\[[^\]]+\]|.+?\[[^\]]+\])(.*)$/.exec(cleanName);
    if (wbMatch) {
      const wbSpec = wbMatch[1];
      cleanName = wbMatch[2].trim().replace(/^'|'$/g, '');
      const bMatch = /\[([^\]]+)\]/.exec(wbSpec);
      const rawWbName = bMatch ? bMatch[1] : wbSpec;
      const baseWb = rawWbName.replace(/\.xlsx?$|\.xlsm$/i, '').toLowerCase();
      if (typeof window !== 'undefined' && window.OOS && Array.isArray(window.OOS.docs)) {
        const foundDoc = window.OOS.docs.find(d => {
          if (d.kind !== 'excel' || !d.wb) return false;
          const dn = (d.name || '').toLowerCase();
          const fn = (d.path ? d.path.split(/[\\/]/).pop() : '').toLowerCase();
          return dn === baseWb || dn === rawWbName.toLowerCase() || fn === rawWbName.toLowerCase();
        });
        if (foundDoc && foundDoc.wb) targetWb = foundDoc.wb;
      }
      if (targetWb === this.wb && typeof window !== 'undefined' && window.OOS_EXTERNAL_WORKBOOKS && window.OOS_EXTERNAL_WORKBOOKS[rawWbName]) {
        targetWb = window.OOS_EXTERNAL_WORKBOOKS[rawWbName];
      }
    }
    if (!cleanName && targetWb !== this.wb && targetWb.sheets && targetWb.sheets.length) {
      return targetWb.sheets[0];
    }
    if (cleanName.indexOf(':') >= 0) {
      const parts = cleanName.split(':');
      const s1 = targetWb.sheet(parts[0].replace(/^'|'$/g, '').trim());
      const s2 = targetWb.sheet(parts[1].replace(/^'|'$/g, '').trim());
      if (!s1 || !s2) throw E.REF;
      const idx1 = targetWb.sheets.indexOf(s1), idx2 = targetWb.sheets.indexOf(s2);
      if (idx1 < 0 || idx2 < 0) throw E.REF;
      const sheets = targetWb.sheets.slice(Math.min(idx1, idx2), Math.max(idx1, idx2) + 1);
      s1._3dSheets = sheets;
      return s1;
    }
    const s = targetWb.sheet(cleanName);
    if (!s) throw E.REF;
    return s;
  }
  // value of a cell, computing formulas on demand
  cellValue(sheet, r, c) {
    const x = sheet.cells.get(key(r, c));
    if (!x) { const sp = sheet.spill.get(key(r, c)); return sp ? sp.v : null; }
    if (x.f) {
      const sk = sheet.name + '!' + key(r, c);
      const st = this.state.get(sk);
      if (st === 2) return x.v === undefined ? null : x.v;
      if (st === 1) { this.circular = this.circular || { sheet: sheet.name, r, c }; return 0; }
      this.computeCell(sheet, r, c, x);
    }
    return x.v === undefined ? null : x.v;
  }
  computeCell(sheet, r, c, x) {
    const sk = sheet.name + '!' + key(r, c);
    this.state.set(sk, 1);
    let v;
    try {
      if (!x._ast || x._src !== x.f) { x._src = x.f; try { x._ast = parse(x.f); } catch (e) { x._ast = { t: 'err', v: e instanceof Err ? e : E.NAME }; } }
      v = this.eval(x._ast, { sheet, r, c, env: null });
      if (v instanceof Ref) v = this.deref(v, true);
      if (v instanceof Lambda) v = E.CALC;
      if (isArr(v)) v = this.spill(sheet, r, c, v);
      if (v === undefined) v = null;
    } catch (e) { v = e instanceof Err ? e : (console.warn(e), E.VALUE); }
    x.v = v;
    this.state.set(sk, 2);
    return v;
  }
  spill(sheet, r, c, arr) {
    const rows = arr.length, cols = arr[0] ? arr[0].length : 0;
    if (rows === 1 && cols === 1) return arr[0][0];
    if (!rows || !cols) return E.CALC;
    // blocked?
    for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
      if (!i && !j) continue;
      const k = key(r + i, c + j); const o = sheet.cells.get(k);
      if (o && (o.v != null || o.f)) return E.SPILL;
      const sp = sheet.spill.get(k); if (sp && sp.src !== key(r, c)) return E.SPILL;
    }
    const src = key(r, c);
    for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) { if (!i && !j) continue; let v = arr[i][j]; if (v instanceof Ref) v = this.deref(v, true); sheet.spill.set(key(r + i, c + j), { src, v: v == null ? null : v }); }
    const c0 = sheet.cells.get(src); if (c0) c0.spillTo = { r2: r + rows - 1, c2: c + cols - 1 };
    let v0 = arr[0][0]; if (v0 instanceof Ref) v0 = this.deref(v0, true);
    return v0 == null ? 0 : v0;
  }
  recalc() {
    this.state = new Map(); this.circular = null;
    const prevSpill = this.wb.sheets.map(s => s.spill.size);
    for (const s of this.wb.sheets) { s.spill = new Map(); s.cells.forEach(x => { if (x.spillTo) delete x.spillTo; }); }
    for (let pass = 0; pass < 6; pass++) {
      if (pass) this.state = new Map();
      let computedAny = false;
      for (const s of this.wb.sheets) s.cells.forEach((x, k) => { if (x.f && this.state.get(s.name + '!' + k) !== 2) { this.computeCell(s, Math.floor(k / 20000), k % 20000, x); computedAny = true; } });
      const any = this.wb.sheets.some(s => s.spill.size);
      if (!computedAny) break;
      if (pass >= 1 && !any && !prevSpill.some(n => n)) break;
    }
    return this.circular;
  }
  // evaluate a formula string in a context (used by data validation / conditional formatting / name box)
  evalFormula(src, sheet, r, c) {
    try { const v = this.eval(parse(src), { sheet, r: r || 0, c: c || 0, env: null }); return v instanceof Ref ? this.deref(v, true) : v; } catch (e) { return e instanceof Err ? e : E.VALUE; }
  }
  // materialize a Ref as 2D array (clipped to used area unless keepShape)
  deref(ref, toScalarIfSingle) {
    if (!(ref instanceof Ref)) return ref;
    if (ref.r1 === ref.r2 && ref.c1 === ref.c2) { const v = this.cellValue(ref.sheet, ref.r1, ref.c1); return toScalarIfSingle ? v : [[v]]; }
    let { r1, c1, r2, c2 } = ref;
    if (r2 - r1 > 5000 || c2 - c1 > 500) { const u = ref.sheet.used(); r2 = Math.min(r2, Math.max(u.r2, r1)); c2 = Math.min(c2, Math.max(u.c2, c1)); }
    const out = [];
    for (let r = r1; r <= r2; r++) { const row = []; for (let c = c1; c <= c2; c++) row.push(this.cellValue(ref.sheet, r, c)); out.push(row); }
    return out;
  }
  // iterate values in a Ref without allocating (skips empty cells when sparse)
  eachIn(ref, fn) {
    const sheets = (ref.sheets && ref.sheets.length) ? ref.sheets : [ref.sheet];
    for (let si = 0; si < sheets.length; si++) {
      const sheet = sheets[si];
      const area = (ref.r2 - ref.r1 + 1) * (ref.c2 - ref.c1 + 1);
      if (area > sheet.cells.size * 4 + 64) {
        const ks = [];
        sheet.cells.forEach((x, k) => { const r = Math.floor(k / 20000), c = k % 20000; if (r >= ref.r1 && r <= ref.r2 && c >= ref.c1 && c <= ref.c2) ks.push(k); });
        sheet.spill.forEach((x, k) => { const r = Math.floor(k / 20000), c = k % 20000; if (r >= ref.r1 && r <= ref.r2 && c >= ref.c1 && c <= ref.c2) ks.push(k); });
        ks.sort((a, b) => a - b);
        for (const k of ks) { const r = Math.floor(k / 20000), c = k % 20000; fn(this.cellValue(sheet, r, c), r, c, sheet); }
      } else {
        for (let r = ref.r1; r <= ref.r2; r++) for (let c = ref.c1; c <= ref.c2; c++) fn(this.cellValue(sheet, r, c), r, c, sheet);
      }
    }
  }
  toRef(node, ctx) {
    const R = node.ref; const sheet = this.sheetOf(R.sheet, ctx);
    const sheets = (sheet && sheet._3dSheets) ? sheet._3dSheets : (sheet ? [sheet] : []);
    if (sheet) delete sheet._3dSheets;
    return new Ref(sheet, Math.min(R.r1, R.r2), Math.min(R.c1, R.c2), Math.max(R.r1, R.r2), Math.max(R.c1, R.c2), sheets);
  }
  eval(node, ctx) {
    switch (node.t) {
      case 'num': case 'str': case 'bool': return node.v;
      case 'err': throw node.v;
      case 'missing': return null;
      case 'paren': return this.eval(node.a, ctx);
      case 'arr': return node.v;
      case 'ref': return this.toRef(node, ctx);
      case 'sref': return this.structRef(node, ctx);
      case 'range': {
        const a = this.eval(node.a, ctx), b = this.eval(node.b, ctx);
        if (!(a instanceof Ref) || !(b instanceof Ref)) throw E.VALUE;
        return new Ref(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
      }
      case 'name': {
        if (ctx.env) { let e = ctx.env; while (e) { if (Object.prototype.hasOwnProperty.call(e.vars, node.n.toUpperCase())) return e.vars[node.n.toUpperCase()]; e = e.parent; } }
        const def = this.wb.names[node.n.toUpperCase()] || this.wb.names[node.n];
        if (def == null) {
          // v0.0.6: eta-reduced lambdas — a bare function name used as a value (GROUPBY(..., SUM), MAP(x, ABS), PIVOTBY(..., AVERAGE))
          const fnName = node.n.toUpperCase();
          if (FN[fnName]) return new Lambda(['_ETA1', '_ETA2', '_ETA3'], { t: 'eta', n: fnName }, null);
          throw E.NAME;
        }
        if (typeof def === 'object' && def.f) return this.eval(def._ast || (def._ast = parse(def.f)), ctx);
        return this.eval(parse(String(def).replace(/^=/, '')), ctx);
      }
      case 'neg': return this.unop(this.eval(node.a, ctx), v => { const n = num(v); return n instanceof Err ? n : -n; });
      case 'pct': return this.unop(this.eval(node.a, ctx), v => { const n = num(v); return n instanceof Err ? n : n / 100; });
      case 'bin': return this.binop(node.op, this.eval(node.a, ctx), this.eval(node.b, ctx));
      case 'fn': return this.call(node, ctx);
      case 'eta': { const args = []; const env = ctx.env; ['_ETA1', '_ETA2', '_ETA3'].forEach(k => { if (env && env.vars[k] !== null && env.vars[k] !== undefined) args.push({ t: 'val', v: env.vars[k] }); }); return this.call({ t: 'fn', n: node.n, args }, ctx); }
      case 'val': return node.v;
      case 'invoke': { const f = this.eval(node.f, ctx); if (!(f instanceof Lambda)) throw E.VALUE; return this.applyLambda(f, node.args.map(a => this.eval(a, ctx)), ctx); }
    }
    throw E.VALUE;
  }
  structRef(node, ctx) {
    const t = (ctx.sheet.wb.sheets.reduce((acc, s) => acc || s.tables.find(x => x.name.toLowerCase() === node.table.toLowerCase()) && { s, x: s.tables.find(x => x.name.toLowerCase() === node.table.toLowerCase()) }, null));
    if (!t) throw E.REF;
    const { s, x } = t; const spec = node.spec.replace(/[\[\]]/g, '').split(',').map(z => z.trim());
    let r1 = x.r1 + 1, r2 = x.r2, c1 = x.c1, c2 = x.c2;
    for (const sp of spec) {
      if (/^#all$/i.test(sp)) { r1 = x.r1; } else if (/^#headers$/i.test(sp)) { r1 = r2 = x.r1; } else if (/^#data$/i.test(sp)) { /* default */ }
      else if (/^#this row$/i.test(sp) || /^@/.test(sp)) { r1 = r2 = ctx.r; const col = sp.replace(/^@/, ''); if (col) { const ci = this.tableCol(s, x, col); c1 = c2 = ci; } }
      else if (sp) { const ci = this.tableCol(s, x, sp); c1 = c2 = ci; }
    }
    return new Ref(s, r1, c1, r2, c2);
  }
  tableCol(s, x, name) { for (let c = x.c1; c <= x.c2; c++) if (String(s.value(x.r1, c)).toLowerCase() === name.toLowerCase()) return c; throw E.REF; }
  unop(v, f) {
    if (v instanceof Ref) v = this.deref(v, true);
    if (isArr(v)) return v.map(row => row.map(x => x instanceof Err ? x : f(x)));
    if (v instanceof Err) return v;
    return f(v);
  }
  binop(op, a, b) {
    if (a instanceof Ref) a = this.deref(a, true);
    if (b instanceof Ref) b = this.deref(b, true);
    if (isArr(a) || isArr(b)) {
      const A = isArr(a) ? a : [[a]], B = isArr(b) ? b : [[b]];
      const rows = Math.max(A.length, B.length), cols = Math.max(A[0].length, B[0].length);
      const out = [];
      for (let i = 0; i < rows; i++) {
        const row = [];
        for (let j = 0; j < cols; j++) {
          const x = A.length === 1 ? A[0] : A[i], y = B.length === 1 ? B[0] : B[i];
          const xv = x ? (x.length === 1 ? x[0] : x[j]) : undefined, yv = y ? (y.length === 1 ? y[0] : y[j]) : undefined;
          row.push(xv === undefined || yv === undefined ? E.NA : this.scalarOp(op, xv, yv));
        }
        out.push(row);
      }
      return out;
    }
    return this.scalarOp(op, a, b);
  }
  scalarOp(op, a, b) {
    if (a instanceof Err) return a; if (b instanceof Err) return b;
    switch (op) {
      case '&': { const x = str(a), y = str(b); return x + y; }
      case '=': return cmp(a, b) === 0;
      case '<>': return cmp(a, b) !== 0;
      case '<': return cmp(a, b) < 0;
      case '>': return cmp(a, b) > 0;
      case '<=': return cmp(a, b) <= 0;
      case '>=': return cmp(a, b) >= 0;
    }
    const x = num(a); if (x instanceof Err) return x; const y = num(b); if (y instanceof Err) return y;
    switch (op) {
      case '+': return fix(x + y);
      case '-': return fix(x - y);
      case '*': return x * y;
      case '/': return y === 0 ? E.DIV0 : x / y;
      case '^': { if (x === 0 && y === 0) return E.NUM; const r = Math.pow(x, y); return isFinite(r) && !isNaN(r) ? r : E.NUM; }
    }
    return E.VALUE;
  }
  call(node, ctx) {
    const name = node.n.replace(/^_XLFN\./, '').replace(/^_XLWS\./, '');
    // LAMBDA invocation of a name or LET variable
    let def = FN[name];
    if (!def) {
      let lam = null;
      if (ctx.env) { let e = ctx.env; while (e && !lam) { if (e.vars[name] instanceof Lambda) lam = e.vars[name]; e = e.parent; } }
      if (!lam && this.wb.names[name]) { const v = this.eval({ t: 'name', n: name }, ctx); if (v instanceof Lambda) lam = v; }
      if (lam) return this.applyLambda(lam, node.args.map(a => this.eval(a, ctx)), ctx);
      throw E.NAME;
    }
    if (def.special) return def.f(node.args, ctx, this);
    const args = node.args.map(a => this.eval(a, ctx));
    if (def.min != null && args.length < def.min) throw E.VALUE;
    // lift scalar functions over array/range args (dynamic arrays)
    if (def.scalar) {
      const shapeIdx = []; let rows = 1, cols = 1;
      for (let i = 0; i < args.length; i++) {
        if (def.scalar === 'first' && i > 0) break;
        let v = args[i];
        if (v instanceof Ref && !(v.r1 === v.r2 && v.c1 === v.c2)) { v = args[i] = this.deref(v); }
        else if (v instanceof Ref) { args[i] = this.cellValue(v.sheet, v.r1, v.c1); continue; }
        if (isArr(v)) { shapeIdx.push(i); rows = Math.max(rows, v.length); cols = Math.max(cols, v[0].length); }
      }
      for (let i = 0; i < args.length; i++) if (args[i] instanceof Ref) args[i] = this.deref(args[i], true);
      if (shapeIdx.length) {
        const out = [];
        for (let r = 0; r < rows; r++) {
          const row = [];
          for (let c = 0; c < cols; c++) {
            const a2 = args.map((v, i) => shapeIdx.indexOf(i) >= 0 ? ((v[v.length === 1 ? 0 : r] || [])[v[0].length === 1 ? 0 : c]) : v);
            try { row.push(def.f(a2, ctx, this)); } catch (e) { row.push(e instanceof Err ? e : E.VALUE); }
          }
          out.push(row);
        }
        return out;
      }
    }
    return def.f(args, ctx, this);
  }
  applyLambda(lam, vals, ctx) {
    if (vals.length > lam.params.length) throw E.VALUE;
    const vars = {}; lam.params.forEach((p, i) => { vars[p] = vals[i] === undefined ? null : vals[i]; });
    return this.eval(lam.body, Object.assign({}, ctx, { env: { vars, parent: lam.env } }));
  }
}
function fix(v) { return Math.abs(v) < 1e-15 ? 0 : +v.toPrecision(15); }

// ------------------------------------------------------------------ formula text utilities
// shift relative references (copy / fill / move). dr, dc = offset
export function shiftFormula(f, dr, dc) {
  if (!dr && !dc) return f;
  try {
    const toks = tokenize(f);
    let out = '', last = 0;
    for (const t of toks) {
      if (t.t !== 'ref') continue;
      const R = t.v; const start = srcPosOf(f, t, last);
      if (start < 0) continue;
      const len = refLen(f, start);
      out += f.slice(last, start) + refText(R, dr, dc);
      last = start + len;
    }
    return out + f.slice(last);
  } catch (e) { return f; }
}
function srcPosOf(f, t, from) { return t.p >= from ? t.p : -1; }
function refLen(f, start) {
  const m = /^(?:'(?:[^']|'')*'!|[A-Za-z_À-ỹ][\w.À-ỹ]*!)?(?:\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)/.exec(f.slice(start));
  return m ? m[0].length : 0;
}
export function refText(R, dr, dc) {
  const sh = R.sheet ? (/^[A-Za-z_][\w.]*$/.test(R.sheet) ? R.sheet : "'" + R.sheet.replace(/'/g, "''") + "'") + '!' : '';
  const sr = (r, a) => a ? r : r + (dr || 0), sc = (c, a) => a ? c : c + (dc || 0);
  const bad = v => v < 0;
  if (R.kind === 'col') { const c1 = sc(R.c1, R.ac1), c2 = sc(R.c2, R.ac2); if (bad(c1) || bad(c2)) return '#REF!'; return sh + (R.ac1 ? '$' : '') + colName(c1) + ':' + (R.ac2 ? '$' : '') + colName(c2); }
  if (R.kind === 'row') { const r1 = sr(R.r1, R.ar1), r2 = sr(R.r2, R.ar2); if (bad(r1) || bad(r2)) return '#REF!'; return sh + (R.ar1 ? '$' : '') + (r1 + 1) + ':' + (R.ar2 ? '$' : '') + (r2 + 1); }
  const r1 = sr(R.r1, R.ar1), c1 = sc(R.c1, R.ac1);
  if (bad(r1) || bad(c1)) return '#REF!';
  let s = sh + (R.ac1 ? '$' : '') + colName(c1) + (R.ar1 ? '$' : '') + (r1 + 1);
  if (R.kind === 'range') { const r2 = sr(R.r2, R.ar2), c2 = sc(R.c2, R.ac2); if (bad(r2) || bad(c2)) return '#REF!'; s += ':' + (R.ac2 ? '$' : '') + colName(c2) + (R.ar2 ? '$' : '') + (r2 + 1); }
  return s;
}
// adjust references for inserted/deleted rows or columns on `sheetName`.
// axis 'r' | 'c', at index, n (positive insert, negative delete)
export function adjustFormulaForInsert(f, formulaSheet, sheetName, axis, at, n) {
  try {
    const toks = tokenize(f);
    let out = '', last = 0;
    for (const t of toks) {
      if (t.t !== 'ref') continue;
      const R = Object.assign({}, t.v);
      const target = R.sheet || formulaSheet;
      if (target.toLowerCase() !== sheetName.toLowerCase()) continue;
      const start = t.p, len = refLen(f, start);
      let dead = false;
      const mv = (v, isEnd) => {
        if (n > 0) return v >= at ? v + n : v;
        const delEnd = at - n - 1; // deleting [at, delEnd]
        if (v < at) return v;
        if (v > delEnd) return v + n;
        return isEnd ? at - 1 : at; // inside deleted block
      };
      if (axis === 'r' && R.kind !== 'col') {
        if (n < 0 && R.r1 >= at && R.r2 <= at - n - 1) dead = true;
        R.r1 = mv(R.r1, false); R.r2 = mv(R.r2, true);
        if (R.r2 < R.r1) dead = true;
      }
      if (axis === 'c' && R.kind !== 'row') {
        if (n < 0 && R.c1 >= at && R.c2 <= at - n - 1) dead = true;
        R.c1 = mv(R.c1, false); R.c2 = mv(R.c2, true);
        if (R.c2 < R.c1) dead = true;
      }
      out += f.slice(last, start) + (dead ? '#REF!' : refText(R, 0, 0));
      last = start + len;
    }
    return out + f.slice(last);
  } catch (e) { return f; }
}
// references in a formula (for highlighting while editing & tracing precedents)
export function formulaRefs(f) {
  try { return tokenize(f).filter(t => t.t === 'ref').map(t => ({ ref: t.v, p: t.p, len: refLen(f, t.p) })); } catch (e) { return []; }
}
// toggle $ on the reference at caret (F4)
export function cycleAbs(text, caret) {
  const refs = formulaRefs(text.replace(/^=/, ''));
  const off = text[0] === '=' ? 1 : 0;
  for (const r of refs) {
    const s = r.p + off, e = s + r.len;
    if (caret >= s && caret <= e) {
      const seg = text.slice(s, e);
      const m = /^(.*!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(.*)$/i.exec(seg); if (!m) return null;
      const state = (m[2] ? 2 : 0) + (m[4] ? 1 : 0); // 0 rel, 1 row abs, 2 col abs, 3 both
      const nextState = { 0: 3, 3: 1, 1: 2, 2: 0 }[state];
      const rep = (m[1] || '') + (nextState & 2 ? '$' : '') + m[3] + (nextState & 1 ? '$' : '') + m[5] + (m[6] || '');
      return { text: text.slice(0, s) + rep + text.slice(e), caret: s + rep.length };
    }
  }
  return null;
}
