// Yomi-Office Open Source v0.0.7 — PowerPoint VBA object model for the shared VBA engine (../excel/vba/interp.js).
// Application, ActivePresentation (Slides, PageSetup, Name, Save…), Slides (Count, Item, Add, AddSlide, Range), Slide (Shapes,
// SlideIndex, Name, Delete, Duplicate, MoveTo, Background, NotesPage), Shapes (AddTextbox, AddShape, AddTitle, Title, Placeholders,
// Count, Item, Range), Shape (Name, Left/Top/Width/Height, Rotation, TextFrame/TextFrame2.TextRange.Text, Font.Size/Bold/Italic/
// Name/Color.RGB, ParagraphFormat.Alignment, Fill.ForeColor.RGB, Line.Visible, Delete, Duplicate, HasTextFrame), ActiveWindow.View.Slide,
// SlideShowSettings.Run. Units are points (960×540 for 16:9) like PowerPoint.
import { VbaError, VbArray, Collection, Dictionary, NOT_FOUND, toNum, toStr, truthy } from '../excel/vba/interp.js';
import { CONSTS, builtins } from '../excel/vba/objects.js';
import { uid, esc } from '../core/util.js';

const PP = {
  pplayouttitle: 1, pplayouttext: 2, pplayouttwocolumntext: 3, pplayouttable: 4, pplayouttextandchart: 5, pplayoutblank: 12, pplayouttitleonly: 11, pplayoutsectionheader: 33, pplayoutobject: 16, pplayoutcustom: 32,
  msotextorientationhorizontal: 1, msotextorientationupward: 2, msotextorientationdownward: 3, msotextorientationvertical: 5,
  msoshaperectangle: 1, msoshaperoundedrectangle: 5, msoshapeoval: 9, msoshapeisoscelestriangle: 7, msoshaperighttriangle: 8, msoshapediamond: 4, msoshapepentagon: 51, msoshapehexagon: 10, msoshaperightarrow: 33, msoshapeleftarrow: 34, msoshapeuparrow: 35, msoshapedownarrow: 36, msoshape5pointstar: 92, msoshapeheart: 21, msoshapecloud: 179,
  ppalignleft: 1, ppaligncenter: 2, ppalignright: 3, ppalignjustify: 4, msotrue: -1, msofalse: 0, msoctrue: 1,
  ppsaveasdefault: 11, ppsaveasopenxmlpresentation: 24, ppsaveasopenxmlpresentationmacroenabled: 25, ppsaveaspdf: 32,
  ppplaceholdertitle: 1, ppplaceholderbody: 2, ppplaceholdercentertitle: 3, ppplaceholdersubtitle: 4, ppviewnormal: 9, ppviewslidesorter: 7
};
const LAYOUT_OF = { 1: 'title', 2: 'content', 3: 'two', 11: 'titleonly', 12: 'blank', 33: 'section', 16: 'content', 32: 'content', 4: 'content', 5: 'content' };
const SHAPE_OF = { 1: 'rect', 5: 'roundRect', 9: 'ellipse', 7: 'triangle', 8: 'rtTriangle', 4: 'diamond', 51: 'pentagon', 10: 'hexagon', 33: 'rightArrow', 34: 'leftArrow', 35: 'upArrow', 36: 'downArrow', 92: 'star5', 21: 'heart', 179: 'cloud' };
const ALIGN_OF = { 1: 'left', 2: 'center', 3: 'right', 4: 'justify' }; const ALIGN_NUM = { left: 1, center: 2, right: 3, justify: 4 };
const hexOf = n => { n = toNum(n); const r = n & 255, g = (n >> 8) & 255, b = (n >> 16) & 255; return '#' + [r, g, b].map(x => (x < 16 ? '0' : '') + x.toString(16)).join(''); };
const rgbNum = c => { const m = /^#?([0-9a-f]{6})$/i.exec(String(c || '')); if (!m) return 0; const v = parseInt(m[1], 16); return ((v & 255) << 16) | (v & 0xff00) | (v >> 16); };
const plain = html => { const d = document.createElement('div'); d.innerHTML = String(html || '').replace(/<br\s*\/?>/gi, '\n').replace(/<\/(div|p|li)>/gi, '\n'); return d.textContent.replace(/\n$/, ''); };
const toHtml = t => String(t).split(/\r\n|\r|\n|\v/).map(esc).join('<br>');

class PObj { }
class PIgnore extends PObj { }
class PApp extends PObj { }
class PPres extends PObj { }
class PSlides extends PObj { }
class PSlide extends PObj { constructor(s) { super(); this.s = s; } }
class PShapes extends PObj { constructor(s, list) { super(); this.s = s; this.list = list; } }
class PShape extends PObj { constructor(s, o) { super(); this.s = s; this.o = o; } }
class PText extends PObj { constructor(o, part) { super(); this.o = o; this.part = part; } } // part: frame | range | font | para | color | fill | forecolor | line | bg | bgfill | bgcolor
class PWin extends PObj { }

export class PptModel {
  constructor(doc, ui) { this.doc = doc; this.ui = ui; this.app = new PApp(); this.pres = new PPres(); this.slides = new PSlides(); this.win = new PWin(); this.ign = new PIgnore(); this.changed = false; }
  get scale() { return 1280 / 960; } // px per point
}

export function installPpt(m) {
  const d = m.doc; const S = m.scale;
  const B = builtins(Object.assign(m, { v: v => v, typeName: v => typeName(v), create, evaluate: () => { throw new VbaError('Evaluate chỉ có trong Excel', 0, 438); } }), null);
  function create(cls) { cls = String(cls).toLowerCase(); if (cls === 'collection') return new Collection(); if (/dictionary$/.test(cls)) return new Dictionary(); throw new VbaError('Không tạo được đối tượng ' + cls + ' (chế độ an toàn)', 0, 429); }
  function typeName(v) { if (v && v.__user) return v.mod.type === 'class' ? v.mod.name : 'Object'; if (v === undefined) return 'Empty'; if (v === null) return 'Nothing'; if (typeof v === 'number') return Number.isInteger(v) ? 'Long' : 'Double'; if (typeof v === 'string') return 'String'; if (typeof v === 'boolean') return 'Boolean'; if (v instanceof PSlide) return 'Slide'; if (v instanceof PShape) return 'Shape'; if (v instanceof PPres) return 'Presentation'; if (v instanceof PApp) return 'Application'; return 'Object'; }
  const touch = () => { m.changed = true; };
  const slideAt = i => { const s = d.slides[toNum(i) - 1]; if (!s) throw new VbaError('Chỉ số slide ngoài phạm vi', 0, 9); return new PSlide(s); };
  const shapesOf = s => new PShapes(s, s.objects);
  const addSlide = (idx, layout) => { const T = m.templates; const s = T.blankSlide(LAYOUT_OF[toNum(layout)] || 'content', d.theme); idx = Math.max(1, Math.min(d.slides.length + 1, toNum(idx) || d.slides.length + 1)); d.slides.splice(idx - 1, 0, s); touch(); return new PSlide(s); };
  const findShape = (s, key) => { if (typeof key === 'number' || /^\d+$/.test(String(key))) { const o = s.objects[toNum(key) - 1]; if (!o) throw new VbaError('Chỉ số hình ngoài phạm vi', 0, 9); return o; } const o = s.objects.find(x => (x.name || '').toLowerCase() === String(key).toLowerCase()); if (!o) throw new VbaError('Không tìm thấy hình "' + key + '"', 0, 9); return o; };
  const nameOf = (s, o) => o.name || ((o.type === 'text' ? 'TextBox ' : o.type === 'shape' ? 'Shape ' : o.type === 'image' ? 'Picture ' : o.type === 'table' ? 'Table ' : 'Object ') + (s.objects.indexOf(o) + 1));
  function member(o, name, a) {
    if (o == null) throw new VbaError('Biến đối tượng chưa được gán (Object variable not set) khi truy cập ".' + name + '"', 0, 91);
    if (o instanceof Collection || o instanceof Dictionary) return o.vcall(name, Array.from(a));
    if (o instanceof PIgnore) return name === 'count' ? 0 : m.ign;
    if (o instanceof PApp) { switch (name) { case 'activepresentation': return m.pres; case 'presentations': return { __count: 1, __coll: [m.pres] }; case 'activewindow': return m.win; case 'name': return 'Microsoft PowerPoint'; case 'version': return '16.0'; case 'caption': return 'Yomi-Office Open Source'; } return m.ign; }
    if (o instanceof PWin) { switch (name) { case 'view': return m.win; case 'slide': return slideAt(d.cur + 1); case 'gotoslide': d.cur = Math.max(0, Math.min(d.slides.length - 1, toNum(a[0]) - 1)); touch(); return; case 'viewtype': return 9; case 'selection': return m.ign; } return m.ign; }
    if (o instanceof PPres) {
      switch (name) {
        case 'slides': return a.length ? slideAt(a[0]) : m.slides;
        case 'name': return (d.name || 'Presentation1') + '.pptx'; case 'fullname': return d.path || d.name; case 'path': return d.path ? String(d.path).replace(/[\\/][^\\/]*$/, '') : '';
        case 'pagesetup': return new PText(null, 'pagesetup'); case 'slidemaster': case 'designs': case 'tags': case 'builtindocumentproperties': return m.ign;
        case 'slideshowsettings': return new PText(null, 'show');
        case 'saved': return !d.dirty; case 'save': m.saveRequested = 'save'; return; case 'saveas': case 'savecopyas': m.saveRequested = 'saveas'; return; case 'close': case 'printout': return;
      }
      return m.ign;
    }
    if (o instanceof PSlides) {
      switch (name) {
        case 'count': return d.slides.length; case '': case 'item': return slideAt(a[0]);
        case 'add': return addSlide(a[0], a[1]); case 'addslide': return addSlide(a[0], 2);
        case 'range': return a.length ? slideAt(a[0]) : m.slides;
        case 'findbyslideid': return slideAt(1);
      }
      return m.ign;
    }
    if (o instanceof PSlide) {
      const s = o.s;
      switch (name) {
        case 'shapes': return a.length ? new PShape(s, findShape(s, a[0])) : shapesOf(s);
        case 'slideindex': case 'slidenumber': return d.slides.indexOf(s) + 1; case 'slideid': return d.slides.indexOf(s) + 256;
        case 'name': return s.name || ('Slide' + (d.slides.indexOf(s) + 1)); case 'layout': return { title: 1, content: 2, two: 3, titleonly: 11, blank: 12, section: 33 }[s.layout] || 2;
        case 'delete': d.slides.splice(d.slides.indexOf(s), 1); if (!d.slides.length) d.slides.push(m.templates.blankSlide('blank', d.theme)); d.cur = Math.min(d.cur, d.slides.length - 1); touch(); return;
        case 'duplicate': { const n = JSON.parse(JSON.stringify(s)); n.id = uid('s'); n.objects.forEach(x => { x.id = uid('o'); }); d.slides.splice(d.slides.indexOf(s) + 1, 0, n); touch(); return new PSlide(n); }
        case 'moveto': { const i = d.slides.indexOf(s); d.slides.splice(i, 1); d.slides.splice(Math.max(0, toNum(a[0]) - 1), 0, s); touch(); return; }
        case 'select': d.cur = d.slides.indexOf(s); touch(); return;
        case 'background': return new PText({ __slide: s }, 'bg'); case 'followmasterbackground': return s.bg ? 0 : -1;
        case 'notespage': return new PText({ __slide: s }, 'notes');
        case 'slideshowtransition': case 'timeline': case 'tags': case 'customlayout': case 'design': return m.ign;
      }
      return m.ign;
    }
    if (o instanceof PShapes) {
      const s = o.s;
      switch (name) {
        case 'count': return s.objects.length; case '': case 'item': return new PShape(s, findShape(s, a[0]));
        case 'addtextbox': { const ob = { id: uid('o'), type: 'text', x: toNum(a[1]) * S, y: toNum(a[2]) * S, w: toNum(a[3]) * S, h: toNum(a[4]) * S, html: '', fontSize: 18 }; s.objects.push(ob); touch(); return new PShape(s, ob); }
        case 'addlabel': { const ob = { id: uid('o'), type: 'text', x: toNum(a[1]) * S, y: toNum(a[2]) * S, w: toNum(a[3]) * S, h: toNum(a[4]) * S, html: '', fontSize: 18 }; s.objects.push(ob); touch(); return new PShape(s, ob); }
        case 'addshape': { const ob = { id: uid('o'), type: 'shape', shape: SHAPE_OF[toNum(a[0])] || 'rect', x: toNum(a[1]) * S, y: toNum(a[2]) * S, w: toNum(a[3]) * S, h: toNum(a[4]) * S, fill: d.theme.colors[0], stroke: 'none', html: '' }; s.objects.push(ob); touch(); return new PShape(s, ob); }
        case 'addtitle': { const ob = { id: uid('o'), type: 'text', ph: 'title', x: 80, y: 40, w: 1120, h: 120, html: '', fontSize: 40, b: true }; s.objects.unshift(ob); touch(); return new PShape(s, ob); }
        case 'title': { const t = s.objects.find(x => x.ph === 'title') || s.objects.find(x => x.type === 'text'); if (!t) throw new VbaError('Slide không có tiêu đề', 0, 91); return new PShape(s, t); }
        case 'hastitle': return s.objects.some(x => x.ph === 'title') ? -1 : 0;
        case 'placeholders': { const ph = s.objects.filter(x => x.type === 'text' && (x.ph || s.objects.indexOf(x) < 2)); return a.length ? new PShape(s, ph[toNum(a[0]) - 1] || s.objects[toNum(a[0]) - 1]) : { __count: ph.length, __coll: ph.map(x => new PShape(s, x)) }; }
        case 'range': return a.length ? new PShape(s, findShape(s, a[0])) : { __count: s.objects.length, __coll: s.objects.map(x => new PShape(s, x)) };
        case 'selectall': return;
      }
      return m.ign;
    }
    if (o instanceof PShape) {
      const ob = o.o, s = o.s;
      switch (name) {
        case 'name': return nameOf(s, ob); case 'left': return ob.x / S; case 'top': return ob.y / S; case 'width': return ob.w / S; case 'height': return ob.h / S; case 'rotation': return ob.rot || 0;
        case 'hastextframe': return ob.type === 'text' || ob.type === 'shape' ? -1 : 0; case 'type': return ob.type === 'text' ? 17 : ob.type === 'shape' ? 1 : ob.type === 'image' ? 13 : ob.type === 'table' ? 19 : 7;
        case 'textframe': case 'textframe2': return new PText(ob, 'frame'); case 'fill': return new PText(ob, 'fill'); case 'line': return new PText(ob, 'line');
        case 'delete': s.objects.splice(s.objects.indexOf(ob), 1); touch(); return;
        case 'duplicate': { const n = JSON.parse(JSON.stringify(ob)); n.id = uid('o'); n.x += 14; n.y += 14; s.objects.push(n); touch(); return new PShape(s, n); }
        case 'select': d.selectedObj = ob; return; case 'zorder': { const i = s.objects.indexOf(ob); s.objects.splice(i, 1); if (toNum(a[0]) === 1 || toNum(a[0]) === 2) s.objects.push(ob); else s.objects.unshift(ob); touch(); return; }
        case 'incrementleft': ob.x += toNum(a[0]) * S; touch(); return; case 'incrementtop': ob.y += toNum(a[0]) * S; touch(); return; case 'incrementrotation': ob.rot = ((ob.rot || 0) + toNum(a[0])) % 360; touch(); return;
        case 'id': return s.objects.indexOf(ob) + 1; case 'visible': return ob.hidden ? 0 : -1; case 'alternativetext': return ob.alt || '';
        case 'table': case 'animationsettings': case 'actionsettings': case 'tags': case 'shadow': case 'threed': return m.ign;
      }
      return m.ign;
    }
    if (o instanceof PText) {
      const ob = o.o;
      switch (o.part) {
        case 'pagesetup': return name === 'slidewidth' ? 960 : name === 'slideheight' ? (d.ratio === '4:3' ? 720 : 540) : m.ign;
        case 'show': if (name === 'run') { m.runShow = true; return; } return m.ign;
        case 'frame': if (name === 'textrange' || name === 'textrange2') return new PText(ob, 'range'); if (name === 'hastext') return plain(ob.html) ? -1 : 0; if (name === 'wordwrap' || name === 'autosize') return -1; return m.ign;
        case 'range':
          switch (name) {
            case 'text': return plain(ob.html); case 'font': return new PText(ob, 'font'); case 'paragraphformat': return new PText(ob, 'para'); case 'length': return plain(ob.html).length;
            case 'insertafter': ob.html = (ob.html || '') + toHtml(a[0]); touch(); return new PText(ob, 'range'); case 'insertbefore': ob.html = toHtml(a[0]) + (ob.html || ''); touch(); return new PText(ob, 'range');
            case 'replace': { const t = plain(ob.html), f = toStr(a[0]), r = toStr(a[1]); const k = t.indexOf(f); if (k < 0) return null; ob.html = toHtml(t.slice(0, k) + r + t.slice(k + f.length)); touch(); return new PText(ob, 'range'); }
            case 'paragraphs': case 'lines': case 'words': case 'characters': return new PText(ob, 'range');
            case 'count': return plain(ob.html).split('\n').length; case 'delete': ob.html = ''; touch(); return;
          }
          return m.ign;
        case 'font': if (name === 'size') return ob.fontSize || 18; if (name === 'bold') return ob.b ? -1 : 0; if (name === 'italic') return ob.i ? -1 : 0; if (name === 'underline') return ob.u ? -1 : 0; if (name === 'name') return ob.font || d.theme.font; if (name === 'color' || name === 'fill') return new PText(ob, 'color'); return m.ign;
        case 'color': case 'forecolor': case 'bgcolor': if (name === 'rgb') return rgbNum(o.part === 'color' ? ob.color : o.part === 'bgcolor' ? (typeof ob.__slide.bg === 'string' ? ob.__slide.bg : '#ffffff') : ob.fill); if (name === 'forecolor') return new PText(ob, 'color'); return m.ign;
        case 'para': if (name === 'alignment') return ALIGN_NUM[ob.align] || 1; return m.ign;
        case 'fill': if (name === 'forecolor') return new PText(ob, 'forecolor'); if (name === 'visible') return ob.fill && ob.fill !== 'none' ? -1 : 0; if (name === 'solid') return; if (name === 'transparency') return 0; return m.ign;
        case 'line': if (name === 'visible') return ob.stroke && ob.stroke !== 'none' ? -1 : 0; if (name === 'forecolor') return new PText(ob, 'lineColor'); return m.ign;
        case 'bg': if (name === 'fill') return new PText(ob, 'bgfill'); return m.ign;
        case 'bgfill': if (name === 'forecolor') return new PText(ob, 'bgcolor'); if (name === 'solid') return; return m.ign;
        case 'notes': if (name === 'shapes' || name === 'placeholders') return new PText(ob, 'notesShapes'); return m.ign;
        case 'notesShapes': return new PText(ob, 'notesShape');
        case 'notesShape': if (name === 'textframe') return new PText(ob, 'notesFrame'); return m.ign;
        case 'notesFrame': if (name === 'textrange') return new PText(ob, 'notesRange'); return m.ign;
        case 'notesRange': if (name === 'text') return ob.__slide.notes || ''; return m.ign;
      }
      return m.ign;
    }
    throw new VbaError('Đối tượng không hỗ trợ thuộc tính "' + name + '"', 0, 438);
  }
  function setMember(o, name, a, val) {
    if (o instanceof PIgnore || o instanceof PApp || o instanceof PPres || o instanceof PWin) return;
    if (o instanceof PSlide) { if (name === 'name') { o.s.name = toStr(val); touch(); } else if (name === 'followmasterbackground') { if (truthy(val)) { o.s.bg = null; touch(); } } else if (name === 'layout') { o.s.layout = LAYOUT_OF[toNum(val)] || o.s.layout; touch(); } return; }
    if (o instanceof PShape) {
      const ob = o.o;
      switch (name) { case 'name': ob.name = toStr(val); break; case 'left': ob.x = toNum(val) * S; break; case 'top': ob.y = toNum(val) * S; break; case 'width': ob.w = toNum(val) * S; break; case 'height': ob.h = toNum(val) * S; break; case 'rotation': ob.rot = toNum(val); break; case 'visible': ob.hidden = !truthy(val); break; case 'alternativetext': ob.alt = toStr(val); break; default: return; }
      touch(); return;
    }
    if (o instanceof PText) {
      const ob = o.o;
      if (o.part === 'range' && name === 'text') { ob.html = toHtml(toStr(val)); touch(); return; }
      if (o.part === 'font') { if (name === 'size') ob.fontSize = toNum(val); else if (name === 'bold') ob.b = truthy(val); else if (name === 'italic') ob.i = truthy(val); else if (name === 'underline') ob.u = truthy(val); else if (name === 'name') ob.font = toStr(val); else return; touch(); return; }
      if ((o.part === 'color' || o.part === 'forecolor' || o.part === 'bgcolor' || o.part === 'lineColor') && name === 'rgb') { const c = hexOf(val); if (o.part === 'color') ob.color = c; else if (o.part === 'forecolor') ob.fill = c; else if (o.part === 'lineColor') ob.stroke = c; else ob.__slide.bg = c; touch(); return; }
      if (o.part === 'para' && name === 'alignment') { ob.align = ALIGN_OF[toNum(val)] || 'left'; touch(); return; }
      if (o.part === 'fill' && name === 'visible') { if (!truthy(val)) ob.fill = 'none'; touch(); return; }
      if (o.part === 'line' && name === 'visible') { ob.stroke = truthy(val) ? (ob.stroke && ob.stroke !== 'none' ? ob.stroke : '#333333') : 'none'; touch(); return; }
      if (o.part === 'notesRange' && name === 'text') { ob.__slide.notes = toStr(val); touch(); return; }
      if (o.part === 'frame' || o.part === 'pagesetup') return;
      return;
    }
    throw new VbaError('Không thể gán "' + name + '" (' + typeName(o) + ')', 0, 438);
  }
  return {
    value: v => v, create, typeName, bracket: () => { throw new VbaError('[...] chỉ dùng trong Excel', 0, 438); },
    same: (x, y) => x === y || (x instanceof PShape && y instanceof PShape && x.o === y.o) || (x instanceof PSlide && y instanceof PSlide && x.s === y.s), hasDefault: () => false,
    each: c => { if (c instanceof VbArray) return c.data.slice(); if (c instanceof PSlides) return d.slides.map(s => new PSlide(s)); if (c instanceof PShapes) return c.s.objects.map(x => new PShape(c.s, x)); if (c instanceof Collection || c instanceof Dictionary) return c.each(); if (c && c.__coll) return c.__coll; throw new VbaError('Không thể lặp For Each', 0, 451); },
    async ident(lc, args, ip) {
      if (Object.prototype.hasOwnProperty.call(CONSTS, lc)) return CONSTS[lc];
      if (Object.prototype.hasOwnProperty.call(PP, lc)) return PP[lc];
      if (/^(pp|mso)[a-z0-9]+$/.test(lc)) return 0;
      switch (lc) { case 'activepresentation': return m.pres; case 'application': return m.app; case 'activewindow': return m.win; case 'presentations': return { __count: 1, __coll: [m.pres] }; case 'debug': return { __dbg: true }; case 'slideshowwindows': case 'commandbars': case 'addins': return m.ign; }
      const f = B[lc]; if (f && ['range', 'cells', 'rows', 'columns', 'worksheets', 'sheets', 'activesheet', 'activecell', 'activeworkbook', 'thisworkbook', 'workbooks', 'worksheetfunction'].indexOf(lc) < 0) return await f(args || [], ip);
      return NOT_FOUND;
    },
    async setIdent() { return false; },
    async member(o, name, args) { if (o && o.__dbg) { if (name === 'print') m.ui.print(Array.from(args).map(x => toStr(x)).join(' ')); return; } return member(o, name, args || []); },
    async setMember(o, name, args, val) { return setMember(o, name, args || [], val); }
  };
}

export const PPT_SAMPLE = `Attribute VB_Name = "Module1"
' Macro PowerPoint mẫu — Yomi-Office Open Source (chạy ngoại tuyến). Alt+F8 để chạy.
Sub TaoSlideTongKet()
    Dim sld As Object, shp As Object, i As Integer
    Set sld = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutTitleOnly)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Tổng kết"
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 140, 840, 300)
    For i = 1 To ActivePresentation.Slides.Count - 1
        shp.TextFrame.TextRange.InsertAfter "• Slide " & i & vbCrLf
    Next i
    shp.TextFrame.TextRange.Font.Size = 24
End Sub
`;
