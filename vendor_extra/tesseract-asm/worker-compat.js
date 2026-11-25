/* Office Open Source: ES2015 library shims for old worker engines (Android 6 stock WebView = Chrome 44), then the Tesseract 2 worker */
(function (G) {
  function rel(i, n) { i = i >> 0; return i < 0 ? Math.max(n + i, 0) : Math.min(i, n); }
  function def(o, k, v) { if (o && !o[k]) try { Object.defineProperty(o, k, { value: v, writable: true, configurable: true }); } catch (e) { o[k] = v; } }
  def(Array, 'from', function (src, fn, self) {
    if (src == null) throw new TypeError('Array.from requires an array-like'); var out = [], i;
    if (typeof Symbol !== 'undefined' && Symbol.iterator && src[Symbol.iterator] && typeof src !== 'string') { var it = src[Symbol.iterator](), s; i = 0; while (!(s = it.next()).done) { out.push(fn ? fn.call(self, s.value, i) : s.value); i++; } return out; }
    var n = src.length >>> 0; for (i = 0; i < n; i++) out.push(fn ? fn.call(self, src[i], i) : src[i]); return out;
  });
  def(Array.prototype, 'copyWithin', function (t, s, e) { var n = this.length >>> 0; t = rel(t, n); s = rel(s, n); e = e === undefined ? n : rel(e, n); var c = Math.min(e - s, n - t); var tmp = Array.prototype.slice.call(this, s, s + c); for (var i = 0; i < tmp.length; i++) this[t + i] = tmp[i]; return this; });
  def(String.prototype, 'padStart', function (n, p) { var s = String(this); p = p === undefined ? ' ' : String(p); while (s.length < n) s = p + s; return s.slice(s.length - Math.max(n, String(this).length)); });
  def(String.prototype, 'padEnd', function (n, p) { var s = String(this); p = p === undefined ? ' ' : String(p); while (s.length < n) s += p; return s.slice(0, Math.max(n, String(this).length)); });
  def(String.prototype, 'trimStart', String.prototype.trimLeft || function () { return this.replace(/^\s+/, ''); });
  def(String.prototype, 'trimEnd', String.prototype.trimRight || function () { return this.replace(/\s+$/, ''); });
  def(Object, 'getOwnPropertyDescriptors', function (o) { var r = {}; Object.getOwnPropertyNames(o).forEach(function (k) { r[k] = Object.getOwnPropertyDescriptor(o, k); }); return r; });
  def(Math, 'cbrt', function (v) { return v < 0 ? -Math.pow(-v, 1 / 3) : Math.pow(v, 1 / 3); });
  def(Math, 'log10', function (v) { return Math.log(v) / Math.LN10; });
  def(Math, 'hypot', function () { var s = 0; for (var i = 0; i < arguments.length; i++) s += arguments[i] * arguments[i]; return Math.sqrt(s); });
  def(Number, 'isFinite', function (v) { return typeof v === 'number' && isFinite(v); });
  def(Number, 'parseFloat', parseFloat); def(Number, 'parseInt', parseInt);
  if (!Number.EPSILON) Number.EPSILON = Math.pow(2, -52);
  if (!Number.MAX_SAFE_INTEGER) Number.MAX_SAFE_INTEGER = 9007199254740991;
  def(Array, 'of', function () { return Array.prototype.slice.call(arguments); });
  def(Array.prototype, 'fill', function (v, s, e) { var n = this.length >>> 0; s = s >> 0; e = e === undefined ? n : e >> 0; for (var i = s < 0 ? Math.max(n + s, 0) : Math.min(s, n), j = e < 0 ? Math.max(n + e, 0) : Math.min(e, n); i < j; i++) this[i] = v; return this; });
  def(Array.prototype, 'find', function (f, t) { for (var i = 0; i < this.length; i++) if (f.call(t, this[i], i, this)) return this[i]; });
  def(Array.prototype, 'findIndex', function (f, t) { for (var i = 0; i < this.length; i++) if (f.call(t, this[i], i, this)) return i; return -1; });
  def(Array.prototype, 'includes', function (v) { for (var i = 0; i < this.length; i++) if (this[i] === v || (v !== v && this[i] !== this[i])) return true; return false; });
  def(Object, 'assign', function (t) { for (var i = 1; i < arguments.length; i++) { var s = arguments[i]; if (s) for (var k in s) if (Object.prototype.hasOwnProperty.call(s, k)) t[k] = s[k]; } return t; });
  def(Object, 'values', function (o) { return Object.keys(o).map(function (k) { return o[k]; }); });
  def(Object, 'entries', function (o) { return Object.keys(o).map(function (k) { return [k, o[k]]; }); });
  def(Number, 'isInteger', function (v) { return typeof v === 'number' && isFinite(v) && Math.floor(v) === v; });
  def(Number, 'isNaN', function (v) { return v !== v; });
  def(String.prototype, 'includes', function (s, p) { return this.indexOf(s, p) !== -1; });
  def(String.prototype, 'startsWith', function (s, p) { p = p || 0; return this.substr(p, s.length) === s; });
  def(String.prototype, 'endsWith', function (s) { return this.slice(-s.length) === s; });
  def(String.prototype, 'repeat', function (n) { var r = ''; for (var i = 0; i < n; i++) r += this; return r; });
  def(Math, 'trunc', function (v) { return v < 0 ? Math.ceil(v) : Math.floor(v); });
  def(Math, 'sign', function (v) { return v > 0 ? 1 : v < 0 ? -1 : +v; });
  def(Math, 'log2', function (v) { return Math.log(v) / Math.LN2; });
  def(Math, 'imul', function (a, b) { var ah = (a >>> 16) & 0xffff, al = a & 0xffff, bh = (b >>> 16) & 0xffff, bl = b & 0xffff; return ((al * bl) + (((ah * bl + al * bh) << 16) >>> 0) | 0); });
  def(Math, 'fround', function (v) { return new Float32Array([v])[0]; });
  def(Math, 'clz32', function (v) { v = v >>> 0; return v ? 31 - Math.floor(Math.log(v + 0.5) / Math.LN2) : 32; });
  ['Int8Array', 'Uint8Array', 'Uint8ClampedArray', 'Int16Array', 'Uint16Array', 'Int32Array', 'Uint32Array', 'Float32Array', 'Float64Array'].forEach(function (n) {
    var T = G[n]; if (!T) return; var P = T.prototype;
    def(P, 'slice', function (s, e) { return new T(Array.prototype.slice.call(this, s, e)); });
    // Chrome 45: copyWithin; TypedArray#set handles overlapping ranges of the same buffer correctly (memmove semantics)
    def(P, 'copyWithin', function (t, s, e) { var n = this.length; t = rel(t, n); s = rel(s, n); e = e === undefined ? n : rel(e, n); var c = Math.min(e - s, n - t); if (c > 0) this.set(this.subarray(s, s + c), t); return this; });
    def(P, 'includes', function (v, f) { return Array.prototype.indexOf.call(this, v, f) !== -1 || (v !== v && Array.prototype.some.call(this, function (x) { return x !== x; })); });
    def(P, 'find', function (f, t) { return Array.prototype.find.call(this, f, t); });
    def(P, 'findIndex', function (f, t) { return Array.prototype.findIndex.call(this, f, t); });
    def(P, 'lastIndexOf', function () { return Array.prototype.lastIndexOf.apply(this, arguments); });
    def(P, 'filter', function (f, t) { return new T(Array.prototype.filter.call(this, f, t)); });
    def(P, 'reduceRight', function () { return Array.prototype.reduceRight.apply(this, arguments); });
    def(P, 'reverse', function () { for (var i = 0, j = this.length - 1; i < j; i++, j--) { var x = this[i]; this[i] = this[j]; this[j] = x; } return this; });
    def(P, 'sort', function (cmp) { var a = Array.prototype.slice.call(this).sort(cmp || function (x, y) { return x - y; }); this.set(a); return this; });
    def(P, 'fill', function (v, s, e) { return Array.prototype.fill.call(this, v, s, e); });
    def(P, 'indexOf', function (v, f) { return Array.prototype.indexOf.call(this, v, f); });
    def(P, 'forEach', function (f, t) { return Array.prototype.forEach.call(this, f, t); });
    def(P, 'map', function (f, t) { return new T(Array.prototype.map.call(this, f, t)); });
    def(P, 'reduce', function () { return Array.prototype.reduce.apply(this, arguments); });
    def(P, 'join', function (s) { return Array.prototype.join.call(this, s); });
    def(P, 'every', function (f, t) { return Array.prototype.every.call(this, f, t); });
    def(P, 'some', function (f, t) { return Array.prototype.some.call(this, f, t); });
    def(T, 'from', function (a, f, t) { return new T(Array.from(a, f, t)); });
    def(T, 'of', function () { return new T(Array.prototype.slice.call(arguments)); });
  });
  if (typeof G.Promise === 'function' && !G.Promise.prototype['finally']) G.Promise.prototype['finally'] = function (f) { return this.then(function (v) { return Promise.resolve(f()).then(function () { return v; }); }, function (e) { return Promise.resolve(f()).then(function () { throw e; }); }); };
})(self);
importScripts('worker.min.js');
