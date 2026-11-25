// v0.1.7: the native Windows ribbon icons (Icons.spec in OfficeUi.cs) drawn as SVG, so the Web UI shows exactly the same pictures.
// ops: r x y w h fill stroke [w] | q x y w h rad fill stroke [w] | o x y w h fill stroke [w] | l x1 y1 x2 y2 color w
//      p fill stroke w x,y x,y ... | t cx cy size color style text | a x y w h start sweep color w
import L from './ui_layout.json';
const PAL = L.palette, SPEC = L.icons;
const cache = {};
const col = t => t === '-' ? 'none' : t[0] === '#' ? t : (PAL[t[0]] || '#808080');
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
export function hasNative(name) { return !!SPEC[name]; }
export function nativeSvg(name, size) {
  const key = name + '|' + size; if (cache[key]) return cache[key];
  const spec = SPEC[name]; if (!spec) return '';
  let out = '';
  spec.split(';').forEach(raw => {
    const op = raw.trim(); if (!op) return;
    const a = op.split(' '), n = i => parseFloat(a[i]);
    try {
      switch (a[0]) {
        case 'r': case 'o': case 'q': {
          const o = a[0] === 'q' ? 1 : 0, x = n(1), y = n(2), w = n(3), h = n(4);
          const fill = col(a[5 + o]), stroke = col(a[6 + o]), sw = a.length > 7 + o ? n(7 + o) : 1.5;
          const st = ' fill="' + fill + '" stroke="' + stroke + '" stroke-width="' + sw + '"';
          if (a[0] === 'o') out += '<ellipse cx="' + (x + w / 2) + '" cy="' + (y + h / 2) + '" rx="' + (w / 2) + '" ry="' + (h / 2) + '"' + st + '/>';
          else out += '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '"' + (o ? ' rx="' + n(5) + '"' : '') + st + '/>';
          break;
        }
        case 'l': out += '<line x1="' + a[1] + '" y1="' + a[2] + '" x2="' + a[3] + '" y2="' + a[4] + '" stroke="' + col(a[5]) + '" stroke-width="' + a[6] + '" stroke-linecap="round"/>'; break;
        case 'a': {   // GDI+ DrawArc: ellipse box, start + sweep in degrees, clockwise
          const x = n(1), y = n(2), w = n(3), h = n(4), s0 = n(5) * Math.PI / 180, sw = n(6) * Math.PI / 180, cx = x + w / 2, cy = y + h / 2, rx = w / 2, ry = h / 2;
          const p = t => (cx + rx * Math.cos(t)).toFixed(2) + ' ' + (cy + ry * Math.sin(t)).toFixed(2);
          out += '<path d="M' + p(s0) + ' A' + rx + ' ' + ry + ' 0 ' + (Math.abs(sw) > Math.PI ? 1 : 0) + ' ' + (sw > 0 ? 1 : 0) + ' ' + p(s0 + sw) + '" fill="none" stroke="' + col(a[7]) + '" stroke-width="' + a[8] + '" stroke-linecap="round"/>';
          break;
        }
        case 'p': {
          const fill = col(a[1]), stroke = col(a[2]), sw = n(3), pts = a.slice(4).filter(v => v.indexOf(',') > 0).join(' ');
          out += (fill !== 'none' ? '<polygon' : '<polyline') + ' points="' + pts + '" fill="' + fill + '" stroke="' + (sw > 0 ? stroke : 'none') + '" stroke-width="' + sw + '" stroke-linejoin="round" stroke-linecap="round"/>';
          break;
        }
        case 't': {
          const txt = a.slice(6).join(' '), st = a[5];
          const fam = (txt === 'Ω' || txt === 'π' || txt === 'Σ' || txt === '¶') ? 'Cambria,Georgia,serif' : 'Segoe UI,Roboto,Arial,sans-serif';
          out += '<text x="' + a[1] + '" y="' + a[2] + '" font-size="' + a[3] + '" fill="' + col(a[4]) + '" font-family="' + fam + '"' + (st.indexOf('b') >= 0 ? ' font-weight="700"' : '') + (st.indexOf('i') >= 0 ? ' font-style="italic"' : '') + ' text-anchor="middle" dominant-baseline="central">' + esc(txt) + '</text>';
          break;
        }
      }
    } catch (e) {}
  });
  return (cache[key] = '<svg xmlns="http://www.w3.org/2000/svg" width="' + size + '" height="' + size + '" viewBox="0 0 32 32">' + out + '</svg>');
}
