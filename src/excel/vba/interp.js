// Yomi-Office Open Source v0.0.6 — VBA (Visual Basic for Applications) engine for Excel, runs fully offline.
// Lexer + recursive-descent parser + async tree-walking interpreter (async so MsgBox/InputBox can wait for the user).
// Supports the VBA language core used by everyday Excel macros: Sub/Function (ByRef/ByVal/Optional/ParamArray), Dim/Const/Static/ReDim [Preserve],
// If/ElseIf/Else (block + single line), For/Next (Step), For Each, Do/Loop While/Until, While/Wend, Select Case (Is, To, lists), With,
// Exit, GoTo/labels, On Error Resume Next/GoTo, Set, Call, arrays (multi-dim, LBound/UBound), Collection, Dictionary (Scripting), ~120 built-ins.
// The Excel object model (Application, Workbook, Worksheets, Range, Cells, Font, Interior…) lives in ./objects.js.

export class VbaError extends Error { constructor(msg, line, num) { super(msg); this.vbaLine = line; this.number = num || 5; } }
class ExitSignal { constructor(kind) { this.kind = kind; } }
class GotoSignal { constructor(label) { this.label = label; } }
class EndSignal {}

// ---------------------------------------------------------------- lexer
const KW = new Set(('and as byref byval call case const dim do each else elseif end eqv erase exit false for function get global goto if imp in is let like loop mod new next not nothing null empty on optional or paramarray preserve private property public redim rem resume select set static step stop sub then to true until wend while with xor error option explicit compare base text binary attribute typeof').split(' '));
export function lex(src) {
  const T = []; let i = 0, line = 1; const n = src.length;
  const push = (t, v, extra) => T.push(Object.assign({ t, v, line }, extra || {}));
  while (i < n) {
    const ch = src[i];
    if (ch === ' ' || ch === '\t' || ch === '\r') { i++; continue; }
    if (ch === '_' && /[ \t]/.test(src[i - 1] || ' ')) { let j = i + 1; while (j < n && /[ \t\r]/.test(src[j])) j++; if (src[j] === '\n') { i = j + 1; line++; continue; } }
    if (ch === '\n') { push('nl'); i++; line++; continue; }
    if (ch === ':' && src[i + 1] !== '=') { push('nl', ':'); i++; continue; }
    if (ch === "'") { while (i < n && src[i] !== '\n') i++; continue; }
    if (ch === '"') { let s = ''; i++; while (i < n) { if (src[i] === '"') { if (src[i + 1] === '"') { s += '"'; i += 2; continue; } i++; break; } if (src[i] === '\n') break; s += src[i++]; } push('str', s); continue; }
    if (ch === '#' ) { const m = /^#([^#\n]+)#/.exec(src.slice(i)); if (m && /\d/.test(m[1])) { push('date', m[1]); i += m[0].length; continue; } }
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] || ''))) { const m = /^(\d*\.?\d+(?:[eE][+-]?\d+)?)[!#@%&^]?/.exec(src.slice(i)); push('num', parseFloat(m[1])); i += m[0].length; continue; }
    if (ch === '&' && /[hHoO]/.test(src[i + 1] || '')) { const m = /^&([hH][0-9a-fA-F]+|[oO][0-7]+)&?/.exec(src.slice(i)); if (m) { push('num', parseInt(m[1].slice(1), /[hH]/.test(m[1][0]) ? 16 : 8)); i += m[0].length; continue; } }
    if (/[A-Za-z_À-ỹ\[]/.test(ch)) {
      let m;
      if (ch === '[') { m = /^\[([^\]\n]*)\]/.exec(src.slice(i)); if (m) { push('bracket', m[1]); i += m[0].length; continue; } }
      m = /^[A-Za-z_À-ỹ][A-Za-z0-9_À-ỹ]*[$%&!#@]?/.exec(src.slice(i)); const w = m[0].replace(/[$%&!#@]$/, ''); const lw = w.toLowerCase();
      i += m[0].length;
      if (lw === 'rem') { while (i < n && src[i] !== '\n') i++; continue; }
      if (KW.has(lw)) push('kw', lw, { raw: w }); else push('id', w, { lc: lw });
      continue;
    }
    const two = src.substr(i, 2);
    if (two === '<>' || two === '<=' || two === '>=' || two === ':=') { push('op', two); i += 2; continue; }
    if ('+-*/\\^&=<>(),.!'.indexOf(ch) >= 0) { push('op', ch); i++; continue; }
    if (ch === ';') { push('op', ','); i++; continue; }
    throw new VbaError('Ký tự không hợp lệ: "' + ch + '"', line, 1);
  }
  push('nl'); push('eof');
  return T;
}

// ---------------------------------------------------------------- parser
export function parseModule(src, modName) {
  const T = lex(src); let p = 0;
  const peek = (k) => T[p + (k || 0)];
  const isKw = (w, k) => { const t = peek(k); return t.t === 'kw' && t.v === w; };
  const isOp = (o, k) => { const t = peek(k); return t.t === 'op' && t.v === o; };
  const isNl = () => peek().t === 'nl' || peek().t === 'eof';
  const err = (m) => { throw new VbaError('Lỗi cú pháp' + (modName ? ' (' + modName + ')' : '') + ', dòng ' + peek().line + ': ' + m, peek().line, 0); };
  const eat = (t, v) => { const x = peek(); if (x.t !== t || (v !== undefined && x.v !== v)) err('cần "' + (v || t) + '" nhưng gặp "' + (x.raw || x.v || x.t) + '"'); p++; return x; };
  const eatKw = w => eat('kw', w);
  const eatOp = o => eat('op', o);
  const skipNl = () => { while (peek().t === 'nl') p++; };
  const endStmt = () => { if (peek().t === 'eof') return; if (peek().t !== 'nl') err('thừa "' + (peek().raw || peek().v) + '"'); p++; };
  const ident = () => { const x = peek(); if (x.t === 'id') { p++; return x.v; } if (x.t === 'kw' && /^(text|binary|compare|base|explicit|error|property|get|attribute|step|string)$/.test(x.v)) { p++; return x.raw; } if (x.t === 'bracket') { p++; return x.v; } err('cần tên biến/thủ tục'); };

  const mod = { name: modName || 'Module1', procs: {}, vars: [], consts: [], order: [] };
  // ----- expressions (VBA precedence)
  function expr() { return imp(); }
  function bin(next, ops) { return function f() { let a = next(); for (;;) { const t = peek(); const o = t.t === 'op' && ops.indexOf(t.v) >= 0 ? t.v : (t.t === 'kw' && ops.indexOf(t.v) >= 0 ? t.v : null); if (!o) return a; p++; a = { k: 'bin', o, a, b: next(), line: t.line }; } }; }
  function notE() { if (isKw('not')) { const l = peek().line; p++; return { k: 'un', o: 'not', a: notE(), line: l }; } return cmpE(); }
  function unary() { if (isOp('-')) { p++; return { k: 'un', o: '-', a: unary() }; } if (isOp('+')) { p++; return unary(); } return pow(); }
  function pow() { let a = postfix(); while (isOp('^')) { p++; let b; if (isOp('-')) { p++; b = { k: 'un', o: '-', a: postfix() }; } else b = postfix(); a = { k: 'bin', o: '^', a, b }; } return a; }
  const mul = bin(unary, ['*', '/']); const idiv = bin(mul, ['\\']); const modE = bin(idiv, ['mod']); const add = bin(modE, ['+', '-']); const cat = bin(add, ['&']);
  const cmpE = bin(cat, ['=', '<>', '<', '>', '<=', '>=', 'like', 'is']);
  const andE = bin(notE, ['and']); const orE = bin(andE, ['or']); const xorE = bin(orE, ['xor']); const eqv = bin(xorE, ['eqv']); const imp = bin(eqv, ['imp']);
  function args() { // after '('
    const out = [];
    if (isOp(')')) { p++; return out; }
    for (;;) {
      if (isOp(',')) { out.push({ k: 'missing' }); p++; continue; }
      if (peek().t === 'id' && isOp(':=', 1)) { const nm = ident(); p++; out.push({ k: 'named', name: nm.toLowerCase(), e: expr() }); }
      else if (isKw('byval')) { p++; out.push(expr()); }
      else out.push(expr());
      if (isOp(',')) { p++; if (isOp(')')) out.push({ k: 'missing' }); continue; }
      eatOp(')'); return out;
    }
  }
  function primary() {
    const t = peek();
    if (t.t === 'num') { p++; return { k: 'lit', v: t.v }; }
    if (t.t === 'str') { p++; return { k: 'lit', v: t.v }; }
    if (t.t === 'date') { p++; const d = new Date(t.v); return { k: 'lit', v: isNaN(d) ? t.v : d }; }
    if (t.t === 'kw') {
      if (t.v === 'true') { p++; return { k: 'lit', v: true }; }
      if (t.v === 'false') { p++; return { k: 'lit', v: false }; }
      if (t.v === 'nothing') { p++; return { k: 'lit', v: null, nothing: true }; }
      if (t.v === 'null') { p++; return { k: 'lit', v: null }; }
      if (t.v === 'empty') { p++; return { k: 'lit', v: undefined }; }
      if (t.v === 'new') { p++; const cls = ident(); while (isOp('.')) { p++; ident(); } return { k: 'new', cls: cls.toLowerCase() }; }
      if (t.v === 'typeof') { p++; const a = postfix(); eatKw('is'); const cls = ident(); return { k: 'typeof', a, cls: cls.toLowerCase() }; }
      if (/^(text|binary|compare|base|error|string|step)$/.test(t.v)) { p++; return { k: 'id', name: t.raw, lc: t.v, line: t.line }; }
    }
    if (t.t === 'id') { p++; return { k: 'id', name: t.v, lc: t.lc, line: t.line }; }
    if (t.t === 'bracket') { p++; return { k: 'bracket', v: t.v }; }
    if (isOp('(')) { p++; const e = expr(); eatOp(')'); return { k: 'paren', e }; }
    if (isOp('.') || isOp('!')) { return { k: 'with' }; }
    err('biểu thức không hợp lệ gần "' + (t.raw || t.v || t.t) + '"');
  }
  function postfix() {
    let e = primary();
    for (;;) {
      if (isOp('.') || isOp('!')) { const bang = isOp('!'); p++; const tk = peek(); let nm; if (tk.t === 'kw') { p++; nm = tk.raw || tk.v; } else nm = ident(); e = { k: 'mem', o: e, name: nm, lc: nm.toLowerCase(), bang, line: peek().line }; continue; }
      if (isOp('(') && e.k !== 'lit') { p++; e = { k: 'call', f: e, args: args(), line: peek().line }; continue; }
      return e;
    }
  }
  // ----- statements
  function block(enders) { // returns list of statements until one of the enders (not consumed)
    const out = [];
    for (;;) {
      skipNl();
      const t = peek(); if (t.t === 'eof') err('thiếu "' + enders[0] + '"');
      if (enders.some(e => matchEnder(e))) return out;
      out.push(stmt());
    }
  }
  function matchEnder(e) { const parts = e.split(' '); for (let k = 0; k < parts.length; k++) { const t = peek(k); if (!(t.t === 'kw' && t.v === parts[k]) && !(t.t === 'id' && t.lc === parts[k])) return false; } return true; }
  function dimList(kind) {
    const out = [];
    do {
      if (isKw('withevents')) p++;
      const name = ident(); let dims = null, isNew = false, type = null;
      if (isOp('(')) { p++; dims = []; if (!isOp(')')) { for (;;) { let lo = expr(), hi = null; if (isKw('to')) { p++; hi = expr(); } else { hi = lo; lo = null; } dims.push({ lo, hi }); if (isOp(',')) { p++; continue; } break; } } eatOp(')'); }
      if (isKw('as')) { p++; if (isKw('new')) { p++; isNew = true; } type = ident(); while (isOp('.')) { p++; type = ident(); } if (isOp('*')) { p++; expr(); } }
      out.push({ name, lc: name.toLowerCase(), dims, isNew, type: type && type.toLowerCase() });
      if (isOp(',')) { p++; continue; } break;
    } while (true);
    return { k: kind || 'dim', vars: out };
  }
  function stmt() {
    const t = peek(); const line = t.line;
    // label
    if ((t.t === 'id' || t.t === 'num') && peek(1).t === 'nl' && peek(1).v === ':' ) { p += 2; return { k: 'label', name: String(t.v).toLowerCase(), line }; }
    if (t.t === 'kw') {
      switch (t.v) {
        case 'dim': case 'static': case 'private': case 'public': case 'global': {
          p++; if (isKw('const')) { p++; return constStmt(line); } if (isKw('sub') || isKw('function')) err('thủ tục lồng nhau'); const d = dimList(); d.line = line; d.isStatic = t.v === 'static'; endStmt(); return d; }
        case 'const': p++; return constStmt(line);
        case 'redim': { p++; const pres = isKw('preserve') ? (p++, true) : false; const vars = []; do { const name = ident(); eatOp('('); const dims = []; for (;;) { let lo = expr(), hi; if (isKw('to')) { p++; hi = expr(); } else { hi = lo; lo = null; } dims.push({ lo, hi }); if (isOp(',')) { p++; continue; } break; } eatOp(')'); if (isKw('as')) { p++; ident(); } vars.push({ lc: name.toLowerCase(), dims }); if (isOp(',')) { p++; continue; } break; } while (true); endStmt(); return { k: 'redim', pres, vars, line }; }
        case 'erase': { p++; const names = [ident().toLowerCase()]; while (isOp(',')) { p++; names.push(ident().toLowerCase()); } endStmt(); return { k: 'erase', names, line }; }
        case 'set': { p++; const lhs = postfix(); eatOp('='); const rhs = expr(); endStmt(); return { k: 'assign', lhs, rhs, set: true, line }; }
        case 'let': { p++; const lhs = postfix(); eatOp('='); const rhs = expr(); endStmt(); return { k: 'assign', lhs, rhs, line }; }
        case 'call': { p++; const e = postfix(); endStmt(); return { k: 'callstmt', e, args: null, line }; }
        case 'if': return ifStmt();
        case 'for': return forStmt();
        case 'do': return doStmt();
        case 'while': { p++; const c = expr(); endStmt(); const body = block(['wend']); eatKw('wend'); endStmt(); return { k: 'while', c, body, line }; }
        case 'select': { p++; eatKw('case'); const e = expr(); endStmt(); const cases = []; let els = null;
          for (;;) { skipNl(); if (matchEnder('end select')) { p += 2; endStmt(); break; } eatKw('case');
            if (isKw('else')) { p++; if (peek().t === 'nl') p++; els = block(['case', 'end select']); continue; }
            const tests = []; for (;;) { if (isKw('is')) { p++; const o = eat('op').v; tests.push({ is: o, e: expr() }); } else { const a = expr(); if (isKw('to')) { p++; tests.push({ lo: a, hi: expr() }); } else tests.push({ eq: a }); } if (isOp(',')) { p++; continue; } break; }
            if (peek().t === 'nl') p++; cases.push({ tests, body: block(['case', 'end select']) }); }
          return { k: 'select', e, cases, els, line }; }
        case 'with': { p++; const e = expr(); endStmt(); const body = block(['end with']); p += 2; endStmt(); return { k: 'with', e, body, line }; }
        case 'exit': { p++; const w = peek(); p++; endStmt(); return { k: 'exit', what: w.v, line }; }
        case 'goto': { p++; const l = peek(); p++; endStmt(); return { k: 'goto', label: String(l.v).toLowerCase(), line }; }
        case 'on': { p++; if (isKw('error')) { p++; if (isKw('resume')) { p++; eatKw('next'); endStmt(); return { k: 'onerror', mode: 'next', line }; } eatKw('goto'); const l = peek(); p++; endStmt(); return { k: 'onerror', mode: String(l.v) === '0' ? 'off' : 'goto', label: String(l.v).toLowerCase(), line }; } err('chỉ hỗ trợ On Error'); }
        case 'resume': { p++; let mode = 'same'; if (isKw('next')) { p++; mode = 'next'; } else if (!isNl()) { mode = 'label'; const l = peek(); p++; endStmt(); return { k: 'resume', mode, label: String(l.v).toLowerCase(), line }; } endStmt(); return { k: 'resume', mode, line }; }
        case 'end': { p++; endStmt(); return { k: 'end', line }; }
        case 'stop': { p++; endStmt(); return { k: 'end', line }; }
        case 'option': case 'attribute': { while (!isNl()) p++; endStmt(); return { k: 'nop' }; }
        case 'error': { p++; const e = expr(); endStmt(); return { k: 'raise', e, line }; }
      }
    }
    // assignment or call
    const lhs = postfix();
    if (isOp('=')) { p++; const rhs = expr(); endStmt(); return { k: 'assign', lhs, rhs, line }; }
    if (isNl()) { endStmt(); return { k: 'callstmt', e: lhs, args: null, line }; }
    // call without parentheses: Foo a, b   (also MsgBox ("x"), vbOK)
    const a = []; let callee = lhs;
    if (lhs.k === 'call' && isOp(',')) { callee = lhs.f; a.push.apply(a, lhs.args.length === 1 ? lhs.args : [{ k: 'paren', e: lhs.args[0] }]); p++; }
    for (;;) {
      if (isNl()) break;
      if (isOp(',')) { a.push({ k: 'missing' }); p++; continue; }
      if (peek().t === 'id' && isOp(':=', 1)) { const nm = ident(); p++; a.push({ k: 'named', name: nm.toLowerCase(), e: expr() }); }
      else { if (isKw('byval')) p++; a.push(expr()); }
      if (isOp(',')) { p++; continue; }
      break;
    }
    endStmt();
    return { k: 'callstmt', e: callee, args: a, line };
  }
  function constStmt(line) { const out = []; do { const name = ident(); if (isKw('as')) { p++; ident(); } eatOp('='); out.push({ lc: name.toLowerCase(), e: expr() }); if (isOp(',')) { p++; continue; } break; } while (true); endStmt(); return { k: 'const', items: out, line }; }
  function ifStmt() {
    const line = peek().line; eatKw('if'); const c = expr(); eatKw('then');
    if (!isNl()) { // single-line If
      const thenS = []; let elseS = null;
      const one = () => { const save = T.slice(); return stmtInline(); };
      thenS.push(stmtInline()); while (isOp(':') || (peek().t === 'nl' && peek().v === ':')) { p++; if (isKw('else')) break; thenS.push(stmtInline()); }
      if (isKw('else')) { p++; elseS = [stmtInline()]; while (peek().t === 'nl' && peek().v === ':') { p++; elseS.push(stmtInline()); } }
      if (isKw('end') && isKw('if', 1)) { p += 2; }
      endStmt();
      return { k: 'if', arms: [{ c, body: thenS }], els: elseS, line };
    }
    endStmt();
    const arms = [{ c, body: block(['elseif', 'else', 'end if']) }]; let els = null;
    for (;;) {
      if (isKw('elseif')) { p++; const c2 = expr(); eatKw('then'); endStmt(); arms.push({ c: c2, body: block(['elseif', 'else', 'end if']) }); continue; }
      if (isKw('else')) { p++; if (isKw('if')) { p++; const c2 = expr(); eatKw('then'); endStmt(); arms.push({ c: c2, body: block(['elseif', 'else', 'end if']) }); continue; } endStmt(); els = block(['end if']); }
      eatKw('end'); eatKw('if'); endStmt(); break;
    }
    return { k: 'if', arms, els, line };
  }
  // a statement inside single-line If (terminated by ':' / Else / newline)
  function stmtInline() {
    const t = peek(); const line = t.line;
    // reuse stmt() by temporarily treating Else as end: build a tiny token window
    let depth = 0, q = p; while (true) { const x = T[q]; if (x.t === 'eof' || (x.t === 'nl' && x.v !== ':')) break; if (x.t === 'nl' && x.v === ':' && depth === 0) break; if (x.t === 'kw' && x.v === 'else' && depth === 0) break; if (x.t === 'op' && x.v === '(') depth++; if (x.t === 'op' && x.v === ')') depth--; q++; }
    const sub = T.slice(p, q).concat([{ t: 'nl', line }, { t: 'eof', line }]);
    const saveT = T.slice(); const saveP = p;
    T.length = 0; sub.forEach(x => T.push(x)); p = 0;
    let s; try { s = stmt(); } finally { const consumed = q; T.length = 0; saveT.forEach(x => T.push(x)); p = consumed; }
    return s;
  }
  function forStmt() {
    const line = peek().line; eatKw('for');
    if (isKw('each')) { p++; const v = ident(); eatKw('in'); const coll = expr(); endStmt(); const body = block(['next']); eatKw('next'); if (!isNl()) ident(); endStmt(); return { k: 'foreach', v: v.toLowerCase(), coll, body, line }; }
    const v = ident(); eatOp('='); const a = expr(); eatKw('to'); const b = expr(); let st = null; if (isKw('step')) { p++; st = expr(); } endStmt();
    const body = block(['next']); eatKw('next');
    // "Next i, j" closes two loops
    if (!isNl()) { ident(); if (isOp(',')) { T[p] = { t: 'kw', v: 'next', line: T[p].line }; return { k: 'for', v: v.toLowerCase(), a, b, st, body, line }; } }
    endStmt();
    return { k: 'for', v: v.toLowerCase(), a, b, st, body, line };
  }
  function doStmt() {
    const line = peek().line; eatKw('do'); let pre = null;
    if (isKw('while') || isKw('until')) { const w = peek().v; p++; pre = { until: w === 'until', c: expr() }; }
    endStmt(); const body = block(['loop']); eatKw('loop'); let post = null;
    if (isKw('while') || isKw('until')) { const w = peek().v; p++; post = { until: w === 'until', c: expr() }; }
    endStmt(); return { k: 'do', pre, post, body, line };
  }
  function procedure(kind, isPrivate) {
    const name = ident(); const params = [];
    if (kind === 'property') { /* name already = Get/Let/Set handled by caller */ }
    if (isOp('(')) { p++; if (!isOp(')')) { for (;;) { let byval = false, optional = false, paramArray = false; for (;;) { if (isKw('byval')) { byval = true; p++; continue; } if (isKw('byref')) { p++; continue; } if (isKw('optional')) { optional = true; p++; continue; } if (isKw('paramarray')) { paramArray = true; p++; continue; } break; } const pn = ident(); let isArr = false; if (isOp('(')) { p++; eatOp(')'); isArr = true; } if (isKw('as')) { p++; ident(); while (isOp('.')) { p++; ident(); } } let def = null; if (isOp('=')) { p++; def = expr(); } params.push({ name: pn, lc: pn.toLowerCase(), byval, optional, paramArray, def, isArr }); if (isOp(',')) { p++; continue; } break; } } eatOp(')'); }
    if (isKw('as')) { p++; ident(); if (isOp('(')) { p++; eatOp(')'); } }
    endStmt();
    const endW = kind === 'sub' ? 'end sub' : kind === 'function' ? 'end function' : 'end property';
    const body = block([endW]); p += 2; endStmt();
    return { name, lc: name.toLowerCase(), kind, params, body, isPrivate, module: mod.name };
  }
  // ----- module level
  for (;;) {
    skipNl(); const t = peek(); if (t.t === 'eof') break;
    let isPrivate = false;
    if (t.t === 'kw' && (t.v === 'public' || t.v === 'private' || t.v === 'global' || t.v === 'friend' || (t.t === 'id' && t.lc === 'friend'))) {
      const nx = peek(1); if (nx.t === 'kw' && (nx.v === 'sub' || nx.v === 'function' || nx.v === 'property' || nx.v === 'static') || (nx.t === 'id' && nx.lc === 'declare')) { isPrivate = t.v === 'private'; p++; }
    }
    if (isKw('static') && (isKw('sub', 1) || isKw('function', 1))) p++;
    if (isKw('sub') || isKw('function')) { const kind = peek().v; p++; const pr = procedure(kind, isPrivate); mod.procs[pr.lc] = pr; mod.order.push(pr.lc); continue; }
    if (isKw('property')) { p++; const acc = peek(); p++; const pr = procedure('property', isPrivate); pr.acc = String(acc.v).toLowerCase(); mod.procs[(pr.acc === 'get' ? '' : pr.acc + ':') + pr.lc] = pr; continue; }
    if (peek().t === 'id' && peek().lc === 'declare') { while (!isNl()) p++; endStmt(); continue; }
    if (peek().t === 'id' && (peek().lc === 'type' || peek().lc === 'enum')) { // Enum support; Type skipped
      const isEnum = peek().lc === 'enum'; p++; ident(); endStmt(); let val = 0;
      for (;;) { skipNl(); if (isKw('end')) { p++; p++; endStmt(); break; } const nm = ident(); if (isOp('=')) { p++; const e = expr(); if (isEnum) mod.consts.push({ lc: nm.toLowerCase(), e }); val = null; } else { if (isEnum) mod.consts.push({ lc: nm.toLowerCase(), e: { k: 'enumnext' } }); else { if (isKw('as')) { p++; ident(); } } } endStmt(); }
      continue;
    }
    const s = stmt();
    if (s.k === 'dim') mod.vars.push(s); else if (s.k === 'const') mod.consts.push.apply(mod.consts, s.items); else if (s.k !== 'nop') err('câu lệnh phải nằm trong Sub/Function');
  }
  return mod;
}

// ---------------------------------------------------------------- runtime values
export class VbArray {
  constructor(bounds) { this.bounds = bounds; this.data = []; const n = bounds.reduce((a, b) => a * Math.max(0, b[1] - b[0] + 1), 1); for (let i = 0; i < n; i++) this.data.push(undefined); }
  index(ix) { if (ix.length !== this.bounds.length) throw new VbaError('Sai số chiều mảng', 0, 9); let k = 0; for (let d = 0; d < ix.length; d++) { const [lo, hi] = this.bounds[d]; const i = Math.floor(ix[d]); if (i < lo || i > hi) throw new VbaError('Chỉ số nằm ngoài phạm vi (Subscript out of range)', 0, 9); k = k * (hi - lo + 1) + (i - lo); } return k; }
  get(ix) { return this.data[this.index(ix)]; }
  set(ix, v) { this.data[this.index(ix)] = v; }
  static from1D(list, base) { const a = new VbArray([[base || 0, (base || 0) + list.length - 1]]); a.data = list.slice(); return a; }
  static from2D(rows) { const R = rows.length, C = R ? rows[0].length : 0; const a = new VbArray([[1, R], [1, C]]); for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) a.data[r * C + c] = rows[r][c]; return a; }
  to2D() { if (this.bounds.length === 1) return [this.data.slice()]; const C = this.bounds[1][1] - this.bounds[1][0] + 1; const out = []; for (let i = 0; i < this.data.length; i += C) out.push(this.data.slice(i, i + C)); return out; }
}
// instance of a class module (or the singleton of a document module: ThisWorkbook / Sheet1 / ThisDocument)
export class UserObj { constructor(mod, hostObj) { this.mod = mod; this.vars = {}; this.hostObj = hostObj; this.__user = true; } }
export class Collection { constructor() { this.items = []; this.keys = []; }
  vcall(m, a) { switch (m) {
    case 'add': { const [item, key, before, after] = a; if (key != null && this.find(key) >= 0) throw new VbaError('Khóa đã tồn tại trong Collection', 0, 457); let at = this.items.length; if (before != null) at = typeof before === 'number' ? before - 1 : this.find(before); else if (after != null) at = (typeof after === 'number' ? after - 1 : this.find(after)) + 1; this.items.splice(at, 0, item); this.keys.splice(at, 0, key == null ? null : String(key).toLowerCase()); return; }
    case 'count': return this.items.length;
    case 'remove': { const i = this.idx(a[0]); this.items.splice(i, 1); this.keys.splice(i, 1); return; }
    case 'item': case '': return this.items[this.idx(a[0])];
  } throw new VbaError('Collection không có thành viên "' + m + '"', 0, 438); }
  find(k) { return this.keys.indexOf(String(k).toLowerCase()); }
  idx(k) { const i = typeof k === 'number' ? k - 1 : this.find(k); if (i < 0 || i >= this.items.length) throw new VbaError('Chỉ số nằm ngoài phạm vi', 0, 9); return i; }
  each() { return this.items.slice(); }
}
export class Dictionary { constructor() { this.m = new Map(); this.compareText = false; }
  k(x) { return this.compareText && typeof x === 'string' ? x.toLowerCase() : (x instanceof Date ? +x : x); }
  vcall(m, a, setVal, isSet) { switch (m) {
    case 'add': if (this.m.has(this.k(a[0]))) throw new VbaError('Khóa đã tồn tại', 0, 457); this.m.set(this.k(a[0]), [a[0], a[1]]); return;
    case 'exists': return this.m.has(this.k(a[0]));
    case 'count': return this.m.size;
    case 'item': case '': if (isSet) { const e = this.m.get(this.k(a[0])); this.m.set(this.k(a[0]), [e ? e[0] : a[0], setVal]); return; } { const e = this.m.get(this.k(a[0])); if (!e) { this.m.set(this.k(a[0]), [a[0], undefined]); return undefined; } return e[1]; }
    case 'keys': return VbArray.from1D(Array.from(this.m.values()).map(e => e[0]));
    case 'items': return VbArray.from1D(Array.from(this.m.values()).map(e => e[1]));
    case 'remove': this.m.delete(this.k(a[0])); return;
    case 'removeall': this.m.clear(); return;
    case 'comparemode': if (isSet) { this.compareText = setVal === 1; return; } return this.compareText ? 1 : 0;
  } throw new VbaError('Dictionary không có thành viên "' + m + '"', 0, 438); }
  each() { return Array.from(this.m.values()).map(e => e[0]); }
}

// ---------------------------------------------------------------- interpreter
export class Interp {
  constructor(host) {
    this.host = host;               // { globals(lc) -> value|undefined, call(lc,args)->{found,value}, print(text), msgbox(), inputbox(), yieldEvery }
    this.modules = []; this.procs = {}; this.modVars = {}; this.statics = {};
    this.steps = 0; this.cancelled = false; this.deadline = 0;
  }
  load(modules) {
    this.modules = modules; this.procs = {}; this.modVars = {}; this.classes = {}; this.docMods = {}; this.docInsts = {};
    for (const m of modules) {
      const t = m.type || 'module';
      if (t === 'class') { this.classes[m.name.toLowerCase()] = m; continue; } // class procs live on instances only
      if (t !== 'module') this.docMods[m.name.toLowerCase()] = m;
      for (const k in m.procs) { const pr = m.procs[k]; const key = k; if (!this.procs[key] || !pr.isPrivate) this.procs[key] = pr; this.procs[m.name.toLowerCase() + '.' + key] = pr; }
    }
  }
  async init() {
    const scope = { vars: this.modVars, parent: null, mod: true };
    for (const m of this.modules) {
      if (m.type === 'class') continue;
      let enumVal = -1;
      for (const c of m.consts) { if (c.e && c.e.k === 'enumnext') { enumVal++; this.modVars[c.lc] = { v: enumVal, c: true }; continue; } const v = await this.ev(c.e, scope); if (typeof v === 'number') enumVal = v; this.modVars[c.lc] = { v, c: true }; }
      for (const d of m.vars) await this.exec([d], scope);
    }
  }
  // ----- class modules / document modules
  async newObj(cls) {
    const mod = this.classes[cls]; if (!mod) return this.host.create(cls);
    const o = new UserObj(mod, null); const sc = { vars: o.vars, parent: null, inst: o };
    for (const c of mod.consts) o.vars[c.lc] = { v: await this.ev(c.e, sc), c: true };
    for (const d of mod.vars) await this.exec([d], sc);
    const init = mod.procs['class_initialize']; if (init) await this.callProc(init, null, sc, [], o);
    return o;
  }
  async docInst(lc) {
    if (this.docInsts[lc]) return this.docInsts[lc];
    const mod = this.docMods[lc]; let hostObj = null;
    try { const h = await this.host.ident(lc, [], this, { vars: {}, parent: null }, false); if (h !== NOT_FOUND) hostObj = h; } catch (e) { /* no host object */ }
    if (hostObj == null && mod.type === 'sheet') { try { hostObj = await this.host.ident('__codename', [mod.name], this, { vars: {}, parent: null }, false); if (hostObj === NOT_FOUND) hostObj = null; } catch (e) { hostObj = null; } }
    const o = new UserObj(mod, hostObj); o.vars = this.modVars; this.docInsts[lc] = o; return o;
  }
  async objMember(o, lc, args, wantValue, isStmt) {
    const pr = o.mod.procs[lc];
    if (pr && (!pr.isPrivate || false)) return this.callProc(pr, null, { vars: {}, parent: null }, Array.from(args || []), o);
    const cell = o.mod.type === 'class' ? o.vars[lc] : null;
    if (cell) { if (args && args.length) { if (cell.v instanceof VbArray) return cell.v.get(Array.from(args).map(toNum)); return this.host.member(cell.v, '', args, this, wantValue); } return cell.v; }
    if (o.hostObj != null) return this.host.member(o.hostObj, lc, args, this, wantValue, isStmt);
    throw new VbaError('Đối tượng ' + o.mod.name + ' không có thành viên "' + lc + '"', 0, 438);
  }
  async objSet(o, lc, args, val, isSet) {
    const pr = (isSet && o.mod.procs['set:' + lc]) || o.mod.procs['let:' + lc] || o.mod.procs['set:' + lc];
    if (pr) return this.callProc(pr, null, { vars: {}, parent: null }, Array.from(args || []).concat([val]), o);
    const cell = o.mod.type === 'class' ? o.vars[lc] : null;
    if (cell) { cell.v = isSet ? val : coerce(this.host.value(val), cell.type); return; }
    if (o.hostObj != null) return this.host.setMember(o.hostObj, lc, args, val, this);
    throw new VbaError('Không thể gán ' + o.mod.name + '.' + lc, 0, 438);
  }
  instOf(sc) { for (let s = sc; s; s = s.parent) if (s.inst) return s.inst; return null; }
  tick(line) {
    if (this.cancelled) throw new VbaError('Đã dừng macro theo yêu cầu', line, 18);
    if (++this.steps % 20000 === 0) { if (this.deadline && Date.now() > this.deadline) throw new VbaError('Macro chạy quá lâu (có thể vòng lặp vô hạn) — đã dừng', line, 18); return new Promise(r => setTimeout(r, 0)); }
    return null;
  }
  // ----- scopes
  lookup(lc, scope) { for (let s = scope; s; s = s.parent) { if (Object.prototype.hasOwnProperty.call(s.vars, lc)) return s.vars[lc]; if (s.inst && s.inst.vars !== this.modVars && Object.prototype.hasOwnProperty.call(s.inst.vars, lc)) return s.inst.vars[lc]; } if (Object.prototype.hasOwnProperty.call(this.modVars, lc)) return this.modVars[lc]; return null; }
  // ----- run a procedure
  async callProc(pr, argNodes, scope, argVals, inst) {
    if (!inst) { const dm = this.docMods[String(pr.module || '').toLowerCase()]; if (dm) inst = await this.docInst(dm.name.toLowerCase()); }
    const vars = {}; const sc = { vars, parent: null, proc: pr, withStack: [], inst: inst || null };
    let ai = 0; const named = {};
    const list = argNodes || [];
    list.forEach(a => { if (a.k === 'named') named[a.name] = a; });
    const positional = list.filter(a => a.k !== 'named');
    for (let i = 0; i < pr.params.length; i++) {
      const prm = pr.params[i];
      if (prm.paramArray) { const rest = []; for (let j = i; j < positional.length; j++) rest.push(await this.ev(positional[j], scope)); if (argVals) rest.push.apply(rest, argVals.slice(i)); vars[prm.lc] = { v: VbArray.from1D(rest) }; break; }
      let a = named[prm.lc] || positional[i];
      if (argVals && i < argVals.length) { vars[prm.lc] = { v: argVals[i] }; continue; }
      if (!a || a.k === 'missing') { if (!prm.optional) throw new VbaError('Thiếu đối số "' + prm.name + '" khi gọi ' + pr.name, 0, 449); vars[prm.lc] = { v: prm.def ? await this.ev(prm.def, scope) : undefined, missing: !prm.def }; continue; }
      if (a.k === 'named') a = a.e;
      if (!prm.byval && a.k === 'id') { const cell = this.lookup(a.lc, scope); if (cell && !cell.c) { vars[prm.lc] = cell; continue; } }
      let v = await this.ev(a, scope); if (prm.byval && v instanceof VbArray) { const c = new VbArray(v.bounds.map(b => b.slice())); c.data = v.data.slice(); v = c; }
      vars[prm.lc] = { v };
    }
    if (pr.kind === 'function' || pr.kind === 'property') vars[pr.lc] = { v: undefined, ret: true };
    const sk = pr.module + '.' + pr.lc; if (!this.statics[sk]) this.statics[sk] = {};
    sc.statics = this.statics[sk];
    try { await this.runBody(pr.body, sc); }
    catch (e) { if (!(e instanceof ExitSignal && (e.kind === 'sub' || e.kind === 'function' || e.kind === 'property'))) throw e; }
    return (pr.kind === 'function' || pr.kind === 'property') ? vars[pr.lc].v : undefined;
  }
  // body with labels / On Error support
  async runBody(body, sc) {
    let pc = 0; sc.onError = null; sc.errLabel = null;
    const labels = {}; body.forEach((s, i) => { if (s.k === 'label') labels[s.name] = i; });
    while (pc < body.length) {
      const s = body[pc];
      try { await this.exec([s], sc); pc++; }
      catch (e) {
        if (e instanceof GotoSignal) { if (labels[e.label] == null) throw new VbaError('Không tìm thấy nhãn "' + e.label + '"', s.line, 0); pc = labels[e.label] + 1; continue; }
        if (e instanceof ExitSignal || e instanceof EndSignal) throw e;
        if (sc.onError === 'next') { this.setErr(e, s); pc++; continue; }
        if (sc.onError === 'goto' && labels[sc.errLabel] != null && !sc.inHandler) { this.setErr(e, s); sc.inHandler = true; sc.resumeAt = pc; pc = labels[sc.errLabel] + 1; continue; }
        if (e instanceof VbaError && !e.vbaLine) e.vbaLine = s.line;
        throw e;
      }
    }
  }
  setErr(e, s) { this.err = { number: (e && e.number) || 5, description: (e && e.message) || String(e), line: s && s.line }; }
  async exec(list, sc) {
    for (const s of list) {
      const w = this.tick(s.line); if (w) await w;
      switch (s.k) {
        case 'nop': case 'label': break;
        case 'dim': for (const v of s.vars) {
          const target = s.isStatic ? sc.statics : sc.vars; if (s.isStatic && sc.statics[v.lc]) { sc.vars[v.lc] = sc.statics[v.lc]; continue; }
          let val; if (v.dims) { val = v.dims.length ? new VbArray(await this.bounds(v.dims, sc)) : new VbArray([[0, -1]]); val.dynamic = !v.dims.length; }
          else if (v.isNew) val = await this.newObj(v.type);
          else val = defaultOf(v.type);
          const cell = { v: val, type: v.type }; target[v.lc] = cell; if (s.isStatic) sc.vars[v.lc] = cell; }
          break;
        case 'const': for (const c of s.items) sc.vars[c.lc] = { v: await this.ev(c.e, sc), c: true }; break;
        case 'redim': for (const v of s.vars) {
          const b = await this.bounds(v.dims, sc); let cell = this.lookup(v.lc, sc); if (!cell) { cell = { v: null }; sc.vars[v.lc] = cell; }
          const na = new VbArray(b);
          if (s.pres && cell.v instanceof VbArray) { const old = cell.v; if (old.bounds.length === 1) { for (let i = b[0][0]; i <= Math.min(b[0][1], old.bounds[0][1]); i++) if (i >= old.bounds[0][0]) na.set([i], old.get([i])); } else { const ob = old.bounds; for (let i = ob[0][0]; i <= Math.min(ob[0][1], b[0][1]); i++) for (let j = ob[1][0]; j <= Math.min(ob[1][1], b[1][1]); j++) na.set([i, j], old.get([i, j])); } }
          cell.v = na; }
          break;
        case 'erase': for (const n of s.names) { const cell = this.lookup(n, sc); if (cell && cell.v instanceof VbArray) { if (cell.v.dynamic) cell.v = new VbArray([[0, -1]]); else cell.v.data = cell.v.data.map(() => undefined); } } break;
        case 'assign': await this.assign(s.lhs, await this.ev(s.rhs, sc, !s.set), sc, s.set); break;
        case 'callstmt': await this.callStmt(s, sc); break;
        case 'if': { let done = false; for (const arm of s.arms) { if (truthy(await this.ev(arm.c, sc))) { await this.exec(arm.body, sc); done = true; break; } } if (!done && s.els) await this.exec(s.els, sc); break; }
        case 'for': {
          const a = toNum(await this.ev(s.a, sc)), b = toNum(await this.ev(s.b, sc)), st = s.st ? toNum(await this.ev(s.st, sc)) : 1;
          let cell = this.lookup(s.v, sc); if (!cell) { cell = { v: 0 }; sc.vars[s.v] = cell; }
          cell.v = a;
          try { while (st >= 0 ? cell.v <= b : cell.v >= b) { await this.exec(s.body, sc); cell.v = toNum(cell.v) + st; } }
          catch (e) { if (!(e instanceof ExitSignal && e.kind === 'for')) throw e; }
          break;
        }
        case 'foreach': {
          const coll = await this.ev(s.coll, sc); const items = this.host.each(coll);
          let cell = this.lookup(s.v, sc); if (!cell) { cell = { v: null }; sc.vars[s.v] = cell; }
          try { for (const it of items) { cell.v = it; await this.exec(s.body, sc); } }
          catch (e) { if (!(e instanceof ExitSignal && e.kind === 'for')) throw e; }
          break;
        }
        case 'do': try {
          for (;;) {
            if (s.pre) { const c = truthy(await this.ev(s.pre.c, sc)); if (s.pre.until ? c : !c) break; }
            await this.exec(s.body, sc);
            if (s.post) { const c = truthy(await this.ev(s.post.c, sc)); if (s.post.until ? c : !c) break; }
            const w2 = this.tick(s.line); if (w2) await w2;
          } } catch (e) { if (!(e instanceof ExitSignal && e.kind === 'do')) throw e; }
          break;
        case 'while': while (truthy(await this.ev(s.c, sc))) { await this.exec(s.body, sc); const w2 = this.tick(s.line); if (w2) await w2; } break;
        case 'select': {
          const v = await this.ev(s.e, sc); let hit = false;
          for (const c of s.cases) {
            for (const t of c.tests) {
              if (t.eq) { if (cmp(v, await this.ev(t.eq, sc)) === 0) { hit = true; break; } }
              else if (t.lo) { if (cmp(v, await this.ev(t.lo, sc)) >= 0 && cmp(v, await this.ev(t.hi, sc)) <= 0) { hit = true; break; } }
              else if (t.is) { const r = cmp(v, await this.ev(t.e, sc)); if ({ '=': r === 0, '<>': r !== 0, '<': r < 0, '>': r > 0, '<=': r <= 0, '>=': r >= 0 }[t.is]) { hit = true; break; } }
            }
            if (hit) { await this.exec(c.body, sc); break; }
          }
          if (!hit && s.els) await this.exec(s.els, sc);
          break;
        }
        case 'with': { const o = await this.ev(s.e, sc); sc.withStack = sc.withStack || []; sc.withStack.push(o); try { await this.exec(s.body, sc); } finally { sc.withStack.pop(); } break; }
        case 'exit': throw new ExitSignal(s.what);
        case 'goto': throw new GotoSignal(s.label);
        case 'onerror': sc.onError = s.mode === 'off' ? null : s.mode; sc.errLabel = s.label; if (s.mode === 'off') this.err = null; sc.inHandler = false; break;
        case 'resume': sc.inHandler = false; if (s.mode === 'label') throw new GotoSignal(s.label); this.err = null; break;
        case 'end': throw new EndSignal();
        case 'raise': throw new VbaError('Lỗi ' + toNum(await this.ev(s.e, sc)), s.line, toNum(await this.ev(s.e, sc)));
        default: throw new VbaError('Câu lệnh chưa hỗ trợ: ' + s.k, s.line, 0);
      }
    }
  }
  async bounds(dims, sc) { const out = []; for (const d of dims) { const lo = d.lo ? toNum(await this.ev(d.lo, sc)) : (this.optionBase || 0); out.push([lo, toNum(await this.ev(d.hi, sc))]); } return out; }
  async evArgs(list, sc) { const out = []; const named = {}; for (const a of list || []) { if (a.k === 'missing') out.push(undefined); else if (a.k === 'named') named[a.name] = await this.ev(a.e, sc); else out.push(await this.ev(a, sc)); } out.named = named; return out; }
  // statement-level call: procedure, builtin, or method on object
  async callStmt(s, sc) {
    const e = s.e;
    if (s.args) { // Foo a, b  → treat as call node
      return this.ev({ k: 'call', f: e, args: s.args, line: s.line, stmt: true }, sc);
    }
    return this.ev(e, sc, false, true);
  }
  // ----- expressions
  async ev(n, sc, wantValue, isStmt) {
    switch (n.k) {
      case 'lit': return n.v;
      case 'paren': return this.host.value(await this.ev(n.e, sc));
      case 'bin': return this.binop(n, sc);
      case 'un': { const v = this.host.value(await this.ev(n.a, sc)); if (n.o === '-') return -toNum(v); return typeof v === 'boolean' ? !v : (v == null ? true : ~toNum(v)); }
      case 'new': return this.newObj(n.cls);
      case 'typeof': { const v = await this.ev(n.a, sc); return this.host.typeName(v).toLowerCase() === n.cls; }
      case 'with': { const ws = sc.withStack; if (!ws || !ws.length) throw new VbaError('Tham chiếu "." không nằm trong khối With', n.line, 91); return ws[ws.length - 1]; }
      case 'bracket': return this.host.bracket(n.v);
      case 'id': {
        const cell = this.lookup(n.lc, sc);
        if (cell) return wantValue ? this.host.value(cell.v) : cell.v;
        const me = this.instOf(sc);
        if (n.lc === 'me' && me) return me.mod.type === 'class' ? me : (wantValue ? this.host.value(me.hostObj) : me);
        if (me && me.mod.procs[n.lc] && me.mod.type === 'class') return this.callProc(me.mod.procs[n.lc], [], sc, null, me);
        const pr = this.procs[n.lc]; if (pr) { const r = await this.callProc(pr, [], sc); return r; }
        if (this.docMods[n.lc]) return this.docInst(n.lc);
        if (n.lc === 'err') return this.errObj();
        const g = await this.host.ident(n.lc, [], this, sc, isStmt);
        if (g !== NOT_FOUND) return wantValue ? this.host.value(g) : g;
        if (sc.proc && !this.optionExplicit) { const c = { v: undefined }; sc.vars[n.lc] = c; return undefined; }
        throw new VbaError('Biến chưa khai báo: ' + n.name, n.line, 0);
      }
      case 'mem': {
        const o = await this.ev(n.o, sc);
        if (o && o.__err) return o.__err(n.lc);
        if (o instanceof UserObj) return this.objMember(o, n.lc, null, wantValue, isStmt);
        return this.host.member(o, n.lc, null, this, wantValue, isStmt);
      }
      case 'call': {
        const f = n.f;
        if (f.k === 'id') {
          const cell = this.lookup(f.lc, sc);
          if (cell && !(cell.ret && !(cell.v instanceof VbArray))) { // array index / default member (a function's own name = recursive call)
            const args = await this.evArgs(n.args, sc);
            if (cell.v instanceof VbArray) return cell.v.get(args.map(toNum));
            if (cell.v && typeof cell.v === 'object') return this.host.member(cell.v, '', args, this, wantValue);
            if (typeof cell.v === 'string' && args.length === 1) return this.host.member(cell.v, '', args, this);
          }
          const me = this.instOf(sc);
          if (me && me.mod.type === 'class' && me.mod.procs[f.lc]) return this.callProc(me.mod.procs[f.lc], n.args, sc, null, me);
          const pr = this.procs[f.lc]; if (pr) return this.callProc(pr, n.args, sc);
          if (f.lc === 'array') { const a = await this.evArgs(n.args, sc); return VbArray.from1D(Array.from(a)); }
          if (f.lc === 'isMissing' || f.lc === 'ismissing') { const a0 = n.args[0]; if (a0 && a0.k === 'id') { const c = this.lookup(a0.lc, sc); return !!(c && c.missing); } return false; }
          const args = await this.evArgs(n.args, sc);
          const g = await this.host.ident(f.lc, args, this, sc, n.stmt);
          if (g !== NOT_FOUND) return g;
          throw new VbaError('Không tìm thấy Sub/Function: ' + f.name, n.line, 35);
        }
        if (f.k === 'mem') {
          const o = await this.ev(f.o, sc); const args = await this.evArgs(n.args, sc);
          if (o === this.macroHost) { const pr = this.procs[f.lc]; if (pr) return this.callProc(pr, null, sc, Array.from(args)); }
          if (o instanceof UserObj) return this.objMember(o, f.lc, args, wantValue, n.stmt);
          return this.host.member(o, f.lc, args, this, wantValue, n.stmt);
        }
        const o = await this.ev(f, sc); const args = await this.evArgs(n.args, sc);
        if (o instanceof VbArray) return o.get(args.map(toNum));
        if (o instanceof UserObj) return this.objMember(o, '', args, wantValue);
        return this.host.member(o, '', args, this, wantValue);
      }
      case 'missing': return undefined;
    }
    throw new VbaError('Biểu thức chưa hỗ trợ: ' + n.k, n.line, 0);
  }
  errObj() { const self = this; return { __vbaErr: true, __err(m) { const e = self.err || { number: 0, description: '' }; if (m === 'number' || m === '') return e.number; if (m === 'description') return e.description; if (m === 'clear') { self.err = null; return; } if (m === 'raise') return; if (m === 'source') return 'VBAProject'; if (m === 'line' || m === 'erl') return e.line || 0; return undefined; } }; }
  async assign(lhs, val, sc, isSet) {
    if (lhs.k === 'id') {
      let cell = this.lookup(lhs.lc, sc);
      if (!cell) { const me = this.instOf(sc); if (me && me.mod.type === 'class' && (me.mod.procs['let:' + lhs.lc] || me.mod.procs['set:' + lhs.lc])) { await this.objSet(me, lhs.lc, null, val, isSet); return; }
        const pr = this.procs['let:' + lhs.lc] || this.procs['set:' + lhs.lc]; if (pr) { await this.callProc(pr, null, sc, [val]); return; }
        if (await this.host.setIdent(lhs.lc, val, this)) return; cell = { v: undefined }; sc.vars[lhs.lc] = cell; }
      if (cell.c) throw new VbaError('Không thể gán cho hằng số ' + lhs.name, lhs.line, 0);
      if (!isSet && cell.v && typeof cell.v === 'object' && !(cell.v instanceof VbArray) && !(cell.v instanceof Date) && this.host.hasDefault(cell.v)) { await this.host.setMember(cell.v, '', null, val, this); return; }
      cell.v = isSet ? val : (val instanceof VbArray ? val : coerce(this.host.value(val), cell.type));
      return;
    }
    if (lhs.k === 'call') {
      const args = await this.evArgs(lhs.args, sc);
      if (lhs.f.k === 'id') { const cell = this.lookup(lhs.f.lc, sc); if (cell && cell.v instanceof VbArray) { cell.v.set(args.map(toNum), isSet ? val : this.host.value(val)); return; } if (cell && cell.v && typeof cell.v === 'object') { await this.host.setMember(cell.v, '', args, val, this); return; } const g = await this.host.ident(lhs.f.lc, args, this, sc); if (g !== NOT_FOUND && g && typeof g === 'object') { await this.host.setMember(g, '', null, val, this); return; } throw new VbaError('Không thể gán cho ' + lhs.f.name, lhs.line, 0); }
      if (lhs.f.k === 'mem') { const o = await this.ev(lhs.f.o, sc); if (o instanceof UserObj) { await this.objSet(o, lhs.f.lc, args, val, isSet); return; } await this.host.setMember(o, lhs.f.lc, args, val, this); return; }
      const o = await this.ev(lhs.f, sc); if (o instanceof VbArray) { o.set(args.map(toNum), val); return; } await this.host.setMember(o, '', args, val, this); return;
    }
    if (lhs.k === 'mem') { const o = await this.ev(lhs.o, sc); if (o instanceof UserObj) { await this.objSet(o, lhs.lc, null, val, isSet); return; } await this.host.setMember(o, lhs.lc, null, val, this); return; }
    throw new VbaError('Vế trái của phép gán không hợp lệ', lhs.line, 0);
  }
  async binop(n, sc) {
    const o = n.o;
    if (o === 'and' || o === 'or') { const a = this.host.value(await this.ev(n.a, sc)); const b = this.host.value(await this.ev(n.b, sc)); if (typeof a === 'boolean' || typeof b === 'boolean' || a == null || b == null) { return o === 'and' ? truthy(a) && truthy(b) : truthy(a) || truthy(b); } return o === 'and' ? (toNum(a) & toNum(b)) : (toNum(a) | toNum(b)); }
    if (o === 'is') { const a = await this.ev(n.a, sc), b = await this.ev(n.b, sc); return (a == null && b == null) || this.host.same(a, b); }
    const a = this.host.value(await this.ev(n.a, sc)), b = this.host.value(await this.ev(n.b, sc));
    switch (o) {
      case '+': if (typeof a === 'string' && typeof b === 'string') return a + b; if (a instanceof Date) return new Date(+a + toNum(b) * 864e5); return toNum(a) + toNum(b);
      case '-': if (a instanceof Date && b instanceof Date) return (a - b) / 864e5; if (a instanceof Date) return new Date(+a - toNum(b) * 864e5); return toNum(a) - toNum(b);
      case '*': return toNum(a) * toNum(b);
      case '/': { const d = toNum(b); if (d === 0) throw new VbaError('Chia cho 0 (Division by zero)', n.line, 11); return toNum(a) / d; }
      case '\\': { const d = Math.round(toNum(b)); if (d === 0) throw new VbaError('Chia cho 0 (Division by zero)', n.line, 11); return Math.trunc(Math.round(toNum(a)) / d); }
      case 'mod': { const d = Math.round(toNum(b)); if (d === 0) throw new VbaError('Chia cho 0 (Division by zero)', n.line, 11); return Math.round(toNum(a)) % d; }
      case '^': return Math.pow(toNum(a), toNum(b));
      case '&': return toStr(a) + toStr(b);
      case '=': return cmp(a, b) === 0; case '<>': return cmp(a, b) !== 0;
      case '<': return cmp(a, b) < 0; case '>': return cmp(a, b) > 0; case '<=': return cmp(a, b) <= 0; case '>=': return cmp(a, b) >= 0;
      case 'like': return likeMatch(toStr(a), toStr(b));
      case 'xor': return truthy(a) !== truthy(b);
      case 'eqv': return truthy(a) === truthy(b);
      case 'imp': return !truthy(a) || truthy(b);
    }
    throw new VbaError('Toán tử chưa hỗ trợ: ' + o, n.line, 0);
  }
  // run a Sub by name
  async run(name, args) {
    const pr = this.procs[String(name).toLowerCase()] || this.procs[String(name).toLowerCase().replace(/^.*!/, '')];
    if (!pr) throw new VbaError('Không tìm thấy macro "' + name + '"', 0, 0);
    try { return await this.callProc(pr, null, { vars: {}, parent: null }, args || []); }
    catch (e) { if (e instanceof EndSignal) return; if (e instanceof ExitSignal) return; throw e; }
  }
}
export const NOT_FOUND = { notFound: true };

// ---------------------------------------------------------------- helpers
export function defaultOf(t) { if (!t) return undefined; if (/^(integer|long|longlong|single|double|currency|byte|decimal)$/.test(t)) return 0; if (t === 'string') return ''; if (t === 'boolean') return false; if (t === 'date') return new Date(1899, 11, 30); if (t === 'variant') return undefined; return null; }
export function coerce(v, t) { if (!t || t === 'variant' || t === 'object') return v; if (/^(integer|long|byte|longlong)$/.test(t)) return v == null ? 0 : bankRound(toNum(v)); if (/^(single|double|currency|decimal)$/.test(t)) return toNum(v); if (t === 'string') return toStr(v); if (t === 'boolean') return truthy(v); if (t === 'date') return toDate(v); return v; }
export function bankRound(x) { const r = Math.round(x); return (Math.abs(x % 1) === 0.5 && r % 2 !== 0) ? r - Math.sign(x) : r; }
export function toNum(v) {
  if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? -1 : 0; if (v == null) return 0;
  if (v instanceof Date) return (v - new Date(1899, 11, 30)) / 864e5;
  if (typeof v === 'string') { const s = v.trim(); if (s === '') return 0; const x = Number(s.replace(/,/g, '')); if (!isNaN(x)) return x; throw new VbaError('Kiểu dữ liệu không khớp (Type mismatch): "' + v + '"', 0, 13); }
  if (v && v.__value) return toNum(v.__value());
  throw new VbaError('Kiểu dữ liệu không khớp (Type mismatch)', 0, 13);
}
export function toStr(v) {
  if (v == null) return ''; if (typeof v === 'boolean') return v ? 'True' : 'False';
  if (typeof v === 'number') { if (Number.isInteger(v)) return String(v); const s = String(+v.toPrecision(15)); return s; }
  if (v instanceof Date) { const p = n => (n < 10 ? '0' : '') + n; const hasT = v.getHours() || v.getMinutes() || v.getSeconds(); return p(v.getDate()) + '/' + p(v.getMonth() + 1) + '/' + v.getFullYear() + (hasT ? ' ' + p(v.getHours()) + ':' + p(v.getMinutes()) + ':' + p(v.getSeconds()) : ''); }
  if (v && v.__value) return toStr(v.__value());
  return String(v);
}
export function toDate(v) { if (v instanceof Date) return v; if (typeof v === 'number') return new Date(new Date(1899, 11, 30).getTime() + Math.round(v * 864e5)); const s = toStr(v); const m = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s.trim()); if (m) { let y = +m[3]; if (y < 100) y += 2000; return new Date(y, +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)); } const d = new Date(s); if (isNaN(d)) throw new VbaError('Kiểu dữ liệu không khớp (không phải ngày): "' + s + '"', 0, 13); return d; }
export function truthy(v) { if (typeof v === 'boolean') return v; if (v == null) return false; if (typeof v === 'string') { const l = v.toLowerCase(); if (l === 'true') return true; if (l === 'false') return false; } return toNum(v) !== 0; }
export function cmp(a, b) {
  if (a == null && b == null) return 0;
  if (typeof a === 'number' && typeof b === 'string' && b !== '' && !isNaN(+b)) b = +b;
  if (typeof b === 'number' && typeof a === 'string' && a !== '' && !isNaN(+a)) a = +a;
  if (a instanceof Date || b instanceof Date) { const x = toNum(a), y = toNum(b); return x < y ? -1 : x > y ? 1 : 0; }
  if (typeof a === 'string' || typeof b === 'string') { const x = toStr(a), y = toStr(b); return x < y ? -1 : x > y ? 1 : 0; }
  const x = toNum(a), y = toNum(b); return x < y ? -1 : x > y ? 1 : 0;
}
export function likeMatch(s, pat) {
  let re = '^';
  for (let i = 0; i < pat.length; i++) { const c = pat[i];
    if (c === '*') re += '[\\s\\S]*'; else if (c === '?') re += '[\\s\\S]'; else if (c === '#') re += '\\d';
    else if (c === '[') { const j = pat.indexOf(']', i); if (j < 0) { re += '\\['; continue; } let cls = pat.slice(i + 1, j); if (cls[0] === '!') cls = '^' + cls.slice(1); re += '[' + cls.replace(/\\/g, '\\\\') + ']'; i = j; }
    else re += c.replace(/[.+^${}()|\\\]]/g, '\\$&'); }
  return new RegExp(re + '$').test(s);
}
