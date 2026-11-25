// Library polyfills for Chrome 44 (Android 6.0 WebView). Syntax is handled by esbuild target chrome44.
// Injected into the bundle; executes before app code.
(function () {
  var AP = Array.prototype, SP = String.prototype, OP = Object;
  if (!AP.includes) AP.includes = function (x, from) { var n = this.length >>> 0, i = from | 0; if (i < 0) i = Math.max(0, n + i); for (; i < n; i++) { var v = this[i]; if (v === x || (v !== v && x !== x)) return true; } return false; };
  if (!AP.find) AP.find = function (f, t) { for (var i = 0; i < this.length; i++) if (f.call(t, this[i], i, this)) return this[i]; };
  if (!AP.findIndex) AP.findIndex = function (f, t) { for (var i = 0; i < this.length; i++) if (f.call(t, this[i], i, this)) return i; return -1; };
  if (!AP.fill) AP.fill = function (v, s, e) { var n = this.length >>> 0; s = s | 0; e = e === undefined ? n : e | 0; if (s < 0) s += n; if (e < 0) e += n; for (; s < e; s++) this[s] = v; return this; };
  if (!AP.flat) AP.flat = function (d) { d = d === undefined ? 1 : d | 0; var o = []; (function rec(a, dd) { for (var i = 0; i < a.length; i++) { if (dd > 0 && Array.isArray(a[i])) rec(a[i], dd - 1); else o.push(a[i]); } })(this, d); return o; };
  if (!AP.flatMap) AP.flatMap = function (f, t) { return this.map(f, t).flat(); };
  if (!Array.from) Array.from = function (it, map, t) { var o = []; if (it == null) return o; if (typeof it.length === 'number') { for (var i = 0; i < it.length; i++) o.push(map ? map.call(t, it[i], i) : it[i]); return o; } var ir = it[typeof Symbol !== 'undefined' && Symbol.iterator] && it[Symbol.iterator](); if (ir) { var r, j = 0; while (!(r = ir.next()).done) o.push(map ? map.call(t, r.value, j++) : r.value); } return o; };
  if (!Array.of) Array.of = function () { return AP.slice.call(arguments); };
  if (!SP.includes) SP.includes = function (s, i) { return this.indexOf(s, i) >= 0; };
  if (!SP.startsWith) SP.startsWith = function (s, i) { return this.substr(i | 0, s.length) === s; };
  if (!SP.endsWith) SP.endsWith = function (s, e) { e = e === undefined ? this.length : e; return this.substring(e - s.length, e) === s; };
  if (!SP.repeat) SP.repeat = function (n) { n = n | 0; var o = '', s = String(this); while (n > 0) { if (n & 1) o += s; s += s; n >>= 1; } return o; };
  if (!SP.padStart) SP.padStart = function (n, p) { p = p === undefined ? ' ' : String(p); var s = String(this); while (s.length < n && p) s = p.substr(0, n - s.length) + s; return s; };
  if (!SP.padEnd) SP.padEnd = function (n, p) { p = p === undefined ? ' ' : String(p); var s = String(this); while (s.length < n && p) s = s + p.substr(0, n - s.length); return s; };
  if (!SP.trimStart) SP.trimStart = SP.trimLeft || function () { return this.replace(/^\s+/, ''); };
  if (!SP.trimEnd) SP.trimEnd = SP.trimRight || function () { return this.replace(/\s+$/, ''); };
  if (!SP.codePointAt) SP.codePointAt = function (i) { return this.charCodeAt(i); };
  if (!String.fromCodePoint) String.fromCodePoint = function () { var o = ''; for (var i = 0; i < arguments.length; i++) { var c = arguments[i]; if (c <= 0xFFFF) o += String.fromCharCode(c); else { c -= 0x10000; o += String.fromCharCode(0xD800 + (c >> 10), 0xDC00 + (c & 0x3FF)); } } return o; };
  if (!OP.entries) OP.entries = function (o) { var r = []; for (var k in o) if (OP.prototype.hasOwnProperty.call(o, k)) r.push([k, o[k]]); return r; };
  if (!OP.values) OP.values = function (o) { var r = []; for (var k in o) if (OP.prototype.hasOwnProperty.call(o, k)) r.push(o[k]); return r; };
  if (!OP.fromEntries) OP.fromEntries = function (it) { var o = {}; Array.from(it).forEach(function (p) { o[p[0]] = p[1]; }); return o; };
  if (!OP.assign) OP.assign = function (t) { for (var i = 1; i < arguments.length; i++) { var s = arguments[i]; if (s) for (var k in s) if (OP.prototype.hasOwnProperty.call(s, k)) t[k] = s[k]; } return t; };
  if (!Number.isFinite) Number.isFinite = function (v) { return typeof v === 'number' && isFinite(v); };
  if (!Number.isNaN) Number.isNaN = function (v) { return v !== v; };
  if (!Number.isInteger) Number.isInteger = function (v) { return typeof v === 'number' && isFinite(v) && Math.floor(v) === v; };
  if (Number.EPSILON === undefined) Number.EPSILON = Math.pow(2, -52);
  if (!Math.trunc) Math.trunc = function (v) { return v < 0 ? Math.ceil(v) : Math.floor(v); };
  if (!Math.sign) Math.sign = function (v) { v = +v; return v > 0 ? 1 : v < 0 ? -1 : v; };
  if (!Math.log10) Math.log10 = function (v) { return Math.log(v) / Math.LN10; };
  if (!Math.log2) Math.log2 = function (v) { return Math.log(v) / Math.LN2; };
  if (!Math.cbrt) Math.cbrt = function (v) { var r = Math.pow(Math.abs(v), 1 / 3); return v < 0 ? -r : r; };
  if (!Math.sinh) Math.sinh = function (x) { return (Math.exp(x) - Math.exp(-x)) / 2; };
  if (!Math.cosh) Math.cosh = function (x) { return (Math.exp(x) + Math.exp(-x)) / 2; };
  if (!Math.tanh) Math.tanh = function (x) { if (x === Infinity) return 1; if (x === -Infinity) return -1; var e = Math.exp(2 * x); return (e - 1) / (e + 1); };
  if (!Math.asinh) Math.asinh = function (x) { return x === -Infinity ? x : Math.log(x + Math.sqrt(x * x + 1)); };
  if (!Math.acosh) Math.acosh = function (x) { return Math.log(x + Math.sqrt(x * x - 1)); };
  if (!Math.atanh) Math.atanh = function (x) { return Math.log((1 + x) / (1 - x)) / 2; };
  if (!Math.hypot) Math.hypot = function () { var s = 0; for (var i = 0; i < arguments.length; i++) s += arguments[i] * arguments[i]; return Math.sqrt(s); };
  if (!Math.expm1) Math.expm1 = function (x) { return Math.exp(x) - 1; };
  if (!Math.log1p) Math.log1p = function (x) { return Math.log(1 + x); };
  if (typeof Promise !== 'undefined' && !Promise.prototype.finally) Promise.prototype.finally = function (f) { return this.then(function (v) { return Promise.resolve(f()).then(function () { return v; }); }, function (e) { return Promise.resolve(f()).then(function () { throw e; }); }); };
  if (!window.requestAnimationFrame) window.requestAnimationFrame = function (cb) { return setTimeout(function () { cb(Date.now()); }, 16); };
  if (!window.cancelAnimationFrame) window.cancelAnimationFrame = function (id) { clearTimeout(id); };
  // Element.closest / matches for WebView 44
  var EP = typeof Element !== 'undefined' && Element.prototype;
  if (EP) { if (!EP.matches) EP.matches = EP.webkitMatchesSelector || EP.msMatchesSelector; if (!EP.closest) EP.closest = function (s) { var e = this; while (e && e.nodeType === 1) { if (e.matches(s)) return e; e = e.parentElement || e.parentNode; } return null; }; }
  if (typeof NodeList !== 'undefined' && !NodeList.prototype.forEach) NodeList.prototype.forEach = AP.forEach;
  // v0.0.6: TypedArray methods (Chrome 45+) used by the VBA/ODF/legacy-format readers on Android 6.0 WebView 44
  ['Int8Array', 'Uint8Array', 'Uint8ClampedArray', 'Int16Array', 'Uint16Array', 'Int32Array', 'Uint32Array', 'Float32Array', 'Float64Array'].forEach(function (n) {
    var C = window[n]; if (!C) return; var P = C.prototype;
    if (!P.fill) P.fill = AP.fill;
    if (!P.slice) P.slice = function (s, e) { var n2 = this.length; s = s === undefined ? 0 : s | 0; e = e === undefined ? n2 : e | 0; if (s < 0) s = Math.max(0, n2 + s); if (e < 0) e = Math.max(0, n2 + e); var o = new C(Math.max(0, Math.min(e, n2) - Math.min(s, n2))); for (var i = 0; i < o.length; i++) o[i] = this[s + i]; return o; };
    ['indexOf', 'lastIndexOf', 'forEach', 'join', 'some', 'every', 'reduce', 'reduceRight', 'find', 'findIndex', 'includes'].forEach(function (m) { if (!P[m] && AP[m]) P[m] = AP[m]; });
    if (!P.map) P.map = function (f, t) { var o = new C(this.length); for (var i = 0; i < this.length; i++) o[i] = f.call(t, this[i], i, this); return o; };
    if (!P.filter) P.filter = function (f, t) { var o = []; for (var i = 0; i < this.length; i++) if (f.call(t, this[i], i, this)) o.push(this[i]); return new C(o); };
    if (!C.from) C.from = function (a, f) { var l = Array.from(a, f); var o = new C(l.length); for (var i = 0; i < l.length; i++) o[i] = l[i]; return o; };
    if (!C.of) C.of = function () { return C.from(arguments); };
  });
  if (typeof IDBObjectStore !== 'undefined' && !IDBObjectStore.prototype.getAll) IDBObjectStore.prototype.getAll = function () { var req = { result: [], onsuccess: null, onerror: null }; var c = this.openCursor(); c.onsuccess = function () { var cur = c.result; if (cur) { req.result.push(cur.value); cur.continue(); } else if (req.onsuccess) req.onsuccess({ target: req }); }; c.onerror = function (e) { if (req.onerror) req.onerror(e); }; return req; };
  if (typeof Blob !== 'undefined' && !Blob.prototype.arrayBuffer) Blob.prototype.arrayBuffer = function () { var b = this; return new Promise(function (ok, no) { var r = new FileReader(); r.onload = function () { ok(r.result); }; r.onerror = no; r.readAsArrayBuffer(b); }); };
})();

// canvas.toBlob: Chrome 50+ (Android 6 stock WebView is 44)
if (typeof HTMLCanvasElement !== 'undefined' && !HTMLCanvasElement.prototype.toBlob) {
  Object.defineProperty(HTMLCanvasElement.prototype, 'toBlob', { configurable: true, writable: true, value: function (cb, type, q) {
    var self = this; setTimeout(function () { var url = self.toDataURL(type, q), b = atob(url.split(',')[1]), u = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); cb(new Blob([u], { type: type || 'image/png' })); }, 0);
  } });
}

// v0.0.6: Chrome 44 (Android 6 stock WebView) lacks RegExp.prototype.flags (49) and KeyboardEvent.key (51)
(function () {
  try {
    if (typeof RegExp !== 'undefined' && !('flags' in RegExp.prototype)) Object.defineProperty(RegExp.prototype, 'flags', { configurable: true, get: function () { return (this.global ? 'g' : '') + (this.ignoreCase ? 'i' : '') + (this.multiline ? 'm' : '') + (this.unicode ? 'u' : '') + (this.sticky ? 'y' : ''); } });
  } catch (e) { /* ignore */ }
  try {
    if (typeof KeyboardEvent !== 'undefined' && !('key' in KeyboardEvent.prototype)) {
      var NAMES = { 8: 'Backspace', 9: 'Tab', 13: 'Enter', 16: 'Shift', 17: 'Control', 18: 'Alt', 19: 'Pause', 20: 'CapsLock', 27: 'Escape', 32: ' ', 33: 'PageUp', 34: 'PageDown', 35: 'End', 36: 'Home', 37: 'ArrowLeft', 38: 'ArrowUp', 39: 'ArrowRight', 40: 'ArrowDown', 45: 'Insert', 46: 'Delete', 91: 'Meta', 93: 'ContextMenu', 229: 'Unidentified',
        186: ';', 187: '=', 188: ',', 189: '-', 190: '.', 191: '/', 192: '`', 219: '[', 220: '\\', 221: ']', 222: "'" };
      var SHIFTED = { 186: ':', 187: '+', 188: '<', 189: '_', 190: '>', 191: '?', 192: '~', 219: '{', 220: '|', 221: '}', 222: '"', 48: ')', 49: '!', 50: '@', 51: '#', 52: '$', 53: '%', 54: '^', 55: '&', 56: '*', 57: '(' };
      Object.defineProperty(KeyboardEvent.prototype, 'key', { configurable: true, get: function () {
        var c = this.keyCode || this.which || 0;
        if (this.type === 'keypress' && this.charCode) return String.fromCharCode(this.charCode);
        if (this.shiftKey && SHIFTED[c]) return SHIFTED[c];
        if (NAMES[c]) return NAMES[c];
        if (c >= 112 && c <= 123) return 'F' + (c - 111);
        if (c >= 96 && c <= 105) return String(c - 96);
        if (c >= 65 && c <= 90) { var ch = String.fromCharCode(c); return this.shiftKey ? ch : ch.toLowerCase(); }
        if (c >= 48 && c <= 57) return String.fromCharCode(c);
        var id = this.keyIdentifier; if (id && /^U\+[0-9A-F]{4}$/i.test(id)) return String.fromCharCode(parseInt(id.slice(2), 16));
        return id || 'Unidentified';
      } });
    }
  } catch (e) { /* ignore */ }
})();
