// Yomi-Office Open Source v0.0.6 — WPS JSA (JavaScript for Applications) macros for Excel.
// Integrated from the "WPS Office" projects of the Project GitHub sheet (okazaki112/Jsa-vibcoding-skill, Bwkyd/wps-skills — MIT):
// JSA is WPS's JavaScript macro language with the same object model as VBA (Application.ActiveSheet.Range("A1").Value2 …).
// Here JSA code runs against the very same Excel object model as our VBA engine (./objects.js) through synchronous proxies.
import { VbArray, Collection, Dictionary, VbaError, toNum, toStr } from './interp.js';
import { ExcelModel, installDispatch, member, setMember, CONSTS, VRange, VObj, rgbLong, builtins } from './objects.js';

const ACTIONS = new Set(['activate', 'select', 'clearcontents', 'clear', 'clearformats', 'clearcomments', 'clearnotes', 'delete', 'insert', 'copy', 'cut', 'pastespecial', 'autofit', 'merge', 'unmerge', 'find', 'replace', 'sort', 'removeduplicates', 'offset', 'resize', 'end', 'specialcells', 'calculate', 'protect', 'unprotect', 'add', 'item', 'move', 'save', 'saveas', 'close', 'printout', 'evaluate', 'run', 'inputbox', 'intersect', 'union', 'goto', 'wait', 'exists', 'remove', 'removeall', 'keys', 'items', 'address', 'autofilter']);
const IS_OBJ = v => v instanceof VObj || v instanceof Collection || v instanceof Dictionary || (v && (v.__coll || v.__tab || v.__vbaErr));

export function toJs(v) {
  if (v instanceof VbArray) { if (v.bounds.length === 2) return v.to2D(); return v.data.slice(); }
  if (v && v.__cverr) return v.__cverr;
  return v;
}
export function fromJs(v) {
  if (v && v.__oosTarget) return v.__oosTarget;
  if (Array.isArray(v)) { if (v.length && Array.isArray(v[0])) return VbArray.from2D(v.map(r => r.map(fromJs))); return VbArray.from2D([v.map(fromJs)]); }
  return v;
}
export function makeJsaRuntime(model) {
  const wrapCache = new WeakMap();
  const call = (o, name, args) => { const a = Array.from(args).map(fromJs); a.named = {}; const r = member(model, o, name, a); return r && r.then ? r.then(x => wrap(x)) : wrap(r); };
  function wrap(v) {
    if (!IS_OBJ(v)) return toJs(v);
    if (wrapCache.has(v)) return wrapCache.get(v);
    // a callable proxy: obj(args) = default member / item, obj.Prop = property, obj.Method(args) = method
    const target = function () {}; target.__oosTarget = v;
    const p = new Proxy(target, {
      get(t, prop) {
        if (prop === '__oosTarget') return v;
        if (typeof prop === 'symbol') { if (prop === Symbol.iterator) return function* () { const items = v.each ? v.each() : (v.__coll || []); for (const it of items) yield wrap(it); }; return undefined; }
        if (prop === 'then' || prop === 'toJSON') return undefined;
        if (prop === 'valueOf' || prop === 'toString') return () => { const x = v instanceof VRange ? model.getValue(v) : v; return toJs(x); };
        const name = String(prop).toLowerCase();
        if (name === 'count') { try { return toJs(member(model, v, 'count', [])); } catch (e) { if (v.each) return v.each().length; throw e; } }
        if (ACTIONS.has(name)) return function () { return call(v, name, arguments); };
        let val; try { val = member(model, v, name, []); } catch (e) { return function () { return call(v, name, arguments); }; }
        if (IS_OBJ(val)) { const inner = wrap(val); return new Proxy(function () {}, { apply(t2, th, args) { return call(v, name, args); }, get(t2, pr) { return inner[pr]; }, set(t2, pr, nv) { inner[pr] = nv; return true; } }); }
        return toJs(val);
      },
      set(t, prop, val) { setMember(model, v, String(prop).toLowerCase(), [], fromJs(val)); return true; },
      apply(t, th, args) { return call(v, '', args); }
    });
    wrapCache.set(v, p); return p;
  }
  const B = builtins(model, null);
  const ui = model.ui;
  const scope = {
    get Application() { return wrap(model.app); }, get ActiveSheet() { return wrap(model.sheetObj(model.activeSheet())); }, get ActiveWorkbook() { return wrap(model.book); }, get ThisWorkbook() { return wrap(model.book); },
    get ActiveCell() { return wrap(model.activeCell()); }, get Selection() { return wrap(model.selRange()); }, get WorksheetFunction() { return wrap(model.wsf); },
    Worksheets(i) { return arguments.length ? call(B.worksheets([]), '', arguments) : wrap(B.worksheets([])); }, Sheets(i) { return arguments.length ? call(B.sheets([]), '', arguments) : wrap(B.sheets([])); },
    Range(a, b) { return wrap(model.range(model.activeSheet(), fromJs(a), fromJs(b))); }, Cells(r, c) { return call(model.sheetObj(model.activeSheet()), 'cells', arguments); },
    Rows(i) { return call(model.sheetObj(model.activeSheet()), 'rows', arguments); }, Columns(i) { return call(model.sheetObj(model.activeSheet()), 'columns', arguments); },
    MsgBox(t, flags, title) { const s = toStr(toJs(t)); const f = flags | 0; if ((f & 7) === 4 || (f & 7) === 3) return window.confirm(s) ? 6 : 7; if ((f & 7) === 1) return window.confirm(s) ? 1 : 2; window.alert(s); return 1; },
    InputBox(prompt, title, def) { const r = window.prompt(toStr(prompt), def == null ? '' : toStr(def)); return r == null ? '' : r; },
    alert(t) { window.alert(toStr(toJs(t))); }, RGB: rgbLong,
    Debug: { Print() { ui.print(Array.from(arguments).map(x => toStr(toJs(x && x.__oosTarget ? model.getValue(x.__oosTarget) : x))).join(' ')); } },
    console: { log() { ui.print(Array.from(arguments).map(x => typeof x === 'object' ? JSON.stringify(x) : String(x)).join(' ')); }, error() { ui.print('Lỗi: ' + Array.from(arguments).join(' ')); }, warn() { ui.print('Cảnh báo: ' + Array.from(arguments).join(' ')); } },
    CreateObject(n) { return wrap(model.create(String(n).toLowerCase())); }, Evaluate(t) { return wrap(model.evaluate(t)); },
    Union() { return wrap(B.union(Array.from(arguments).map(fromJs))); }, Intersect() { return wrap(B.intersect(Array.from(arguments).map(fromJs))); }
  };
  // xl* / vb* constants with their usual casing (xlUp, vbYesNo …)
  const constants = {}; Object.keys(CONSTS).forEach(k => { constants[k] = CONSTS[k]; });
  return { scope, constants, wrap };
}
export function jsaFunctions(code) { const out = []; const re = /(?:^|\n)\s*(?:async\s+)?function\s+([A-Za-z_$À-ỹ][\w$À-ỹ]*)\s*\(\s*\)/g; let m; while ((m = re.exec(code))) out.push(m[1]); return out; }
export function jsaHas(code, name) { return new RegExp('function\\s+' + name.replace(/[$]/g, '\\$') + '\\s*\\(').test(code); }
// run one JSA function; all JSA modules share one global scope (like WPS)
export function runJsa(model, modules, name) {
  const { scope, constants } = makeJsaRuntime(model);
  const sandbox = new Proxy({}, {
    has(t, k) { return typeof k === 'string' && (k in scope || k.toLowerCase() in constants || /^(xl|vb|mso)[A-Z]/.test(k)); },
    get(t, k) { if (typeof k !== 'string') return undefined; if (k in scope) return scope[k]; const c = constants[k.toLowerCase()]; return c; }
  });
  const src = modules.map(m => m.code).join('\n;\n') + '\n;return (typeof ' + name + ' === "function") ? ' + name + '() : (function(){ throw new Error("Không tìm thấy hàm JSA ' + name + '"); })();';
  // eslint-disable-next-line no-new-func
  const fn = new Function('__s', 'with (__s) { return (function(){ ' + src + ' })(); }');
  return fn(sandbox);
}
// keep JSA modules inside vbaProject.bin as commented VBA modules so they survive .xlsm round trips (Excel ignores them)
export const JSA_TAG = "'@OOS-JSA";
export function jsaToVba(code) { return JSA_TAG + '\r\n' + String(code).split(/\r?\n/).map(l => "' " + l).join('\r\n') + '\r\n'; }
export function vbaToJsa(code) { if (String(code).trim().indexOf(JSA_TAG) !== 0) return null; return String(code).replace(/\r\n/g, '\n').split('\n').slice(1).map(l => l.replace(/^' ?/, '')).join('\n').replace(/\n+$/, '\n'); }
