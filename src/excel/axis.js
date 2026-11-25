// Fast row/column position math with sparse size overrides (hidden = 0)
export class Axis {
  constructor(sheet, isRow) { this.sheet = sheet; this.isRow = isRow; this.dirty = true; }
  get def() { return this.isRow ? 20 : 64; }
  size(i) { return this.isRow ? this.sheet.rowHeight(i) : this.sheet.colWidth(i); }
  build() {
    const s = this.sheet; const map = this.isRow ? s.rowH : s.colW; const hid = this.isRow ? s.hiddenR : s.hiddenC;
    const idx = {};
    for (const k in map) idx[k] = 1; for (const k in hid) if (hid[k]) idx[k] = 1;
    this.keys = Object.keys(idx).map(Number).sort((a, b) => a - b);
    this.prefix = []; let acc = 0; const d = this.def;
    for (const k of this.keys) { this.prefix.push(acc); acc += this.size(k) - d; }
    this.total = acc; this.dirty = false;
  }
  // sum of deltas for indices < i
  delta(i) {
    if (this.dirty) this.build();
    const ks = this.keys; let lo = 0, hi = ks.length - 1, ans = -1;
    while (lo <= hi) { const m = (lo + hi) >> 1; if (ks[m] < i) { ans = m; lo = m + 1; } else hi = m - 1; }
    if (ans < 0) return 0;
    return this.prefix[ans] + this.size(ks[ans]) - this.def;
  }
  pos(i) { return i * this.def + this.delta(i); }
  // index containing position x (x in sheet units)
  at(x, max) {
    if (x <= 0) return 0;
    let lo = 0, hi = max - 1;
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (this.pos(m) <= x) lo = m; else hi = m - 1; }
    // skip hidden (size 0) items forward
    while (lo < max - 1 && this.size(lo) === 0) lo++;
    return lo;
  }
}
