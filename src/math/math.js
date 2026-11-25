// MathType-style equation editor & standalone MathType application
import { h, ic, esc, $, $$, KaTeX, utf8Encode, utf8Decode, VERSION } from '../core/util.js';
import { Doc, registerApp } from '../core/shell.js';
import { dialog, toast, showMenu } from '../core/ui.js';
import Host from '../core/host.js';

export const SYMBOLS = {
  'Toán tử': ['+', '-', '\\times', '\\div', '\\pm', '\\mp', '\\cdot', '\\ast', '\\star', '\\circ', '\\bullet', '\\oplus', '\\ominus', '\\otimes', '\\oslash', '\\odot'],
  'Quan hệ': ['=', '\\neq', '\\approx', '\\equiv', '\\cong', '\\sim', '\\simeq', '\\propto', '<', '>', '\\leq', '\\geq', '\\ll', '\\gg', '\\prec', '\\succ', '\\subset', '\\supset', '\\subseteq', '\\supseteq', '\\in', '\\notin', '\\ni'],
  'Mũi tên': ['\\rightarrow', '\\leftarrow', '\\leftrightarrow', '\\Rightarrow', '\\Leftarrow', '\\Leftrightarrow', '\\uparrow', '\\downarrow', '\\mapsto', '\\to', '\\longrightarrow', '\\hookrightarrow', '\\nearrow', '\\searrow'],
  'Hy Lạp thường': ['\\alpha', '\\beta', '\\gamma', '\\delta', '\\epsilon', '\\varepsilon', '\\zeta', '\\eta', '\\theta', '\\vartheta', '\\iota', '\\kappa', '\\lambda', '\\mu', '\\nu', '\\xi', '\\pi', '\\rho', '\\sigma', '\\tau', '\\upsilon', '\\phi', '\\varphi', '\\chi', '\\psi', '\\omega'],
  'Hy Lạp hoa': ['\\Gamma', '\\Delta', '\\Theta', '\\Lambda', '\\Xi', '\\Pi', '\\Sigma', '\\Upsilon', '\\Phi', '\\Psi', '\\Omega'],
  'Lớn': ['\\sum', '\\prod', '\\coprod', '\\int', '\\iint', '\\iiint', '\\oint', '\\bigcup', '\\bigcap', '\\bigvee', '\\bigwedge', '\\bigoplus', '\\bigotimes'],
  'Ký hiệu': ['\\infty', '\\partial', '\\nabla', '\\forall', '\\exists', '\\nexists', '\\emptyset', '\\varnothing', '\\neg', '\\wedge', '\\vee', '\\angle', '\\triangle', '\\square', '\\Re', '\\Im', '\\aleph', '\\hbar', '\\ell', '\\wp', '\\degree'],
  'Dấu': ['\\hat{a}', '\\bar{a}', '\\vec{a}', '\\dot{a}', '\\ddot{a}', '\\tilde{a}', '\\acute{a}', '\\grave{a}', '\\check{a}', '\\breve{a}', "a'", 'a^{\\circ}'],
  'Chữ đặc biệt': ['\\mathbb{R}', '\\mathbb{N}', '\\mathbb{Z}', '\\mathbb{Q}', '\\mathbb{C}', '\\mathbb{P}', '\\mathcal{L}', '\\mathcal{F}', '\\mathfrak{g}']
};

export const TEMPLATES = [
  ['Phân số', '\\frac{a}{b}', 'divide'], ['Căn bậc hai', '\\sqrt{x}', 'radical'], ['Căn bậc n', '\\sqrt[n]{x}', 'radical'],
  ['Lũy thừa', 'x^{n}', 'superscript'], ['Chỉ số dưới', 'x_{n}', 'subscript'], ['Trên và dưới', 'x_{a}^{b}', 'variable'],
  ['Tổng', '\\sum_{i=1}^{n} a_i', 'sigma'], ['Tích', '\\prod_{i=1}^{n} a_i', 'x'], ['Tích phân', '\\int_{a}^{b} f(x)\\,dx', 'integral'],
  ['Giới hạn', '\\lim_{x \\to \\infty} f(x)', 'move-right'], ['Đạo hàm', '\\frac{d}{dx} f(x)', 'divide'], ['Đạo hàm riêng', '\\frac{\\partial f}{\\partial x}', 'divide'],
  ['Ma trận 2×2', '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}', 'grid-2x2'], ['Định thức', '\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}', 'grid-2x2'],
  ['Hệ phương trình', '\\begin{cases} x + y = 1 \\\\ x - y = 0 \\end{cases}', 'braces'], ['Nghiệm bậc 2', 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}', 'function-square'],
  ['Tổ hợp', '\\binom{n}{k}', 'parentheses'], ['Vectơ', '\\vec{v} = (x, y, z)', 'move-up-right']
];

export const EXAMPLES = [
  ['Phương trình bậc hai', 'ax^2 + bx + c = 0'], ['Nghiệm PT bậc 2', 'x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}'],
  ['Định lý Pythagore', 'a^2 + b^2 = c^2'], ['Khai triển nhị thức', '(a+b)^n = \\sum_{k=0}^{n} \\binom{n}{k} a^{n-k} b^k'],
  ['Hằng đẳng thức Euler', 'e^{i\\pi} + 1 = 0'], ['Đạo hàm', "f'(x) = \\lim_{h \\to 0} \\frac{f(x+h)-f(x)}{h}"],
  ['Tích phân cơ bản', '\\int_0^1 x^2\\,dx = \\frac{1}{3}'], ['Chuỗi hình học', '\\sum_{n=0}^{\\infty} ar^n = \\frac{a}{1-r}'],
  ['Tích phân suy rộng Gauss', '\\int_{-\\infty}^{\\infty} e^{-x^2}\\,dx = \\sqrt{\\pi}'],
  ['Định luật Ohm & Công suất', 'P = U \\cdot I = I^2 \\cdot R = \\frac{U^2}{R}'],
  ['Phương trình Schrödinger', 'i\\hbar\\frac{\\partial}{\\partial t}\\Psi = \\hat{H}\\Psi']
];

// Modal equation editor for inserting equations into Word, Excel, PowerPoint
export function equationEditor(initial) {
  return new Promise(resolve => {
    const input = h('textarea.eq-input', { spellcheck: false, rows: 3 }); input.value = initial || '';
    const preview = h('div.eq-preview');
    const tabsBar = h('div.eq-pal-tabs'); const palette = h('div.eq-palette'); const errLine = h('div.eq-err');
    const render = () => {
      KaTeX().then(k => {
        try {
          k.render(input.value || '\\;', preview, { throwOnError: true, displayMode: true });
          errLine.textContent = '';
        } catch (e) {
          try { k.render(input.value, preview, { throwOnError: false, displayMode: true }); } catch (e2) {}
          errLine.textContent = 'Cú pháp: ' + String(e.message || e).slice(0, 80);
        }
      });
    };
    const insert = (tex) => {
      const s = input.selectionStart, e = input.selectionEnd;
      input.value = input.value.slice(0, s) + tex + input.value.slice(e);
      const caret = tex.indexOf('{}') >= 0 ? s + tex.indexOf('{}') + 1 : s + tex.length;
      input.setSelectionRange(caret, caret);
      input.focus();
      render();
    };
    input.addEventListener('input', render);

    const tmplRow = h('div.eq-templates');
    TEMPLATES.forEach(t => {
      const b = h('button.eq-tmpl', { title: t[0] });
      KaTeX().then(k => { try { k.render(t[1].replace(/a_i|f\(x\)|x|a|b|c|d|n/g, m => m), b, { throwOnError: false }); } catch (e) { b.textContent = t[0]; } });
      b.addEventListener('click', () => insert(t[1]));
      tmplRow.appendChild(b);
    });

    let curCat = Object.keys(SYMBOLS)[0];
    const drawPal = () => {
      palette.innerHTML = '';
      SYMBOLS[curCat].forEach(sym => {
        const b = h('button.eq-sym');
        KaTeX().then(k => { try { k.render(sym, b, { throwOnError: false }); } catch (e) { b.textContent = sym; } });
        b.addEventListener('click', () => insert(sym.indexOf('a') >= 0 && /\\(hat|bar|vec|dot|ddot|tilde|acute|grave|check|breve)/.test(sym) ? sym.replace('{a}', '{}') : sym + ' '));
        palette.appendChild(b);
      });
    };
    Object.keys(SYMBOLS).forEach(cat => {
      const b = h('button.eq-pal-tab' + (cat === curCat ? '.on' : ''), cat);
      b.addEventListener('click', () => { curCat = cat; $$('.eq-pal-tab', tabsBar).forEach(x => x.classList.toggle('on', x === b)); drawPal(); });
      tabsBar.appendChild(b);
    });
    drawPal(); render();

    const examples = h('select.inp', [h('option', { value: '' }, 'Ví dụ công thức…')].concat(EXAMPLES.map(e => h('option', { value: e[1] }, e[0]))));
    examples.addEventListener('change', () => { if (examples.value) { input.value = examples.value; render(); } });

    const body = h('div.eq-editor',
      h('div.eq-preview-wrap', preview, errLine),
      h('div.eq-toolbar', examples),
      tmplRow,
      h('div.eq-pal', tabsBar, palette),
      h('label.lbl', 'Mã LaTeX (có thể gõ trực tiếp):'), input,
      h('div.hint', 'Mẹo: ^ cho lũy thừa, _ cho chỉ số dưới, \\frac{tử}{mẫu} cho phân số. MathType dùng KaTeX để hiển thị.'));

    dialog({
      title: 'MathType — Trình soạn phương trình',
      width: 640,
      body,
      cls: 'eq-dlg',
      buttons: [
        { label: 'Chèn', primary: true, run: () => { resolve(input.value.trim() || null); return true; } },
        { label: 'Hủy', value: null, run: () => { resolve(null); return true; } }
      ],
      onOpen: () => setTimeout(() => input.focus(), 50)
    });
  });
}

export function renderInline(latex) {
  return KaTeX().then(k => {
    const span = document.createElement('span');
    try { k.render(latex, span, { throwOnError: false, displayMode: false }); } catch (e) { span.textContent = latex; }
    return span.innerHTML;
  });
}

export function toImage(latex) { return renderInline(latex); }

// ============================================================================
// STANDALONE MATHTYPE APPLICATION (MathDoc)
// Full-fledged Equation Editor workspace with dedicated Ribbon, Palettes,
// Live KaTeX Canvas, MathML/LaTeX conversion, and multi-format Export.
// ============================================================================
export class MathDoc extends Doc {
  static get kind() { return 'mathtype'; }
  static templates() {
    return [
      { id: 'blank', name: 'Công thức trống', latex: '' },
      { id: 'quadratic', name: 'Phương trình bậc 2', latex: 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}' },
      { id: 'calculus', name: 'Tích phân & Đạo hàm', latex: '\\int_{a}^{b} f(x)\\,dx = F(b) - F(a)' },
      { id: 'matrix', name: 'Ma trận & Hệ phương trình', latex: '\\begin{pmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{pmatrix} \\cdot \\begin{pmatrix} x_1 \\\\ x_2 \\end{pmatrix} = \\begin{pmatrix} b_1 \\\\ b_2 \\end{pmatrix}' },
      { id: 'trig', name: 'Lượng giác & Euler', latex: 'e^{i\\theta} = \\cos\\theta + i\\sin\\theta' },
      { id: 'physics', name: 'Vật lý hiện đại & Sóng', latex: 'E = mc^2 = \\frac{m_0 c^2}{\\sqrt{1 - \\frac{v^2}{c^2}}}' }
    ];
  }

  constructor(shell, opts) {
    super(shell, opts);
    this.latex = (opts && opts.latex) || (opts && opts.template && opts.template.latex) || 'f(x) = \\int_{-\\infty}^{x} e^{-t^2}\\,dt';
    this.undoStack = [];
    this.redoStack = [];
    this.curCat = 'Toán tử';
    this.zoom = 120;
    this.buildUi();
  }

  async init() {
    this.pushHistory(true);
    this.render();
  }

  async load(bytes, ext, name) {
    const txt = utf8Decode(bytes).trim();
    if (ext === 'mml' || ext === 'xml') {
      // Basic MathML to LaTeX extraction or fallback
      const m = /<annotation encoding="application\/x-tex">([\s\S]*?)<\/annotation>/i.exec(txt);
      this.latex = m ? m[1].trim() : txt;
    } else {
      this.latex = txt.replace(/^\\\[|\\\]$/g, '').trim();
    }
    this.pushHistory(true);
    this.render();
  }

  async restore(native) {
    this.latex = native.latex || '';
    this.pushHistory(true);
    this.render();
  }

  snapshot() {
    return { latex: this.latex };
  }

  getText() {
    return this.latex;
  }

  stats() {
    const chars = this.latex.length;
    const tokens = (this.latex.match(/\\[A-Za-z]+|[-+*=^_{}]/g) || []).length;
    return {
      'Độ dài mã': chars + ' ký tự',
      'Số ký hiệu toán': tokens,
      'Định dạng': 'LaTeX / MathML KaTeX',
      'Trạng thái': 'Sẵn sàng chèn hoặc xuất'
    };
  }

  canSaveExt(ext) {
    return ['tex', 'latex', 'mml', 'svg', 'png', 'html', 'txt'].indexOf(ext) >= 0;
  }

  saveFormats() {
    return [
      ['tex', 'Mã nguồn LaTeX (*.tex)', 'file-code'],
      ['mml', 'Công thức MathML (*.mml)', 'code-2'],
      ['svg', 'Ảnh vector SVG (*.svg)', 'image'],
      ['png', 'Ảnh đồ họa PNG (*.png)', 'image'],
      ['html', 'Trang web công thức (*.html)', 'globe'],
      ['txt', 'Văn bản thuần (*.txt)', 'file']
    ];
  }

  async save(ext) {
    if (ext === 'tex' || ext === 'latex' || ext === 'txt') {
      return utf8Encode(this.latex);
    }
    if (ext === 'html') {
      const page = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + esc(this.name) + '</title>' +
        '<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex/dist/katex.min.css">' +
        '</head><body style="display:flex;align-items:center;justify-content:center;height:100vh;margin:0;font-family:sans-serif">' +
        '<div style="font-size:2em">' + (await renderInline(this.latex)) + '</div>' +
        '</body></html>';
      return utf8Encode(page);
    }
    if (ext === 'mml') {
      const k = await KaTeX();
      const mml = k.renderToString(this.latex, { output: 'mathml', throwOnError: false });
      return utf8Encode('<?xml version="1.0" encoding="utf-8"?>\n' + mml);
    }
    if (ext === 'svg' || ext === 'png') {
      const span = document.createElement('div');
      span.style.padding = '20px';
      span.style.background = '#ffffff';
      span.style.display = 'inline-block';
      const k = await KaTeX();
      k.render(this.latex, span, { displayMode: true, throwOnError: false });
      document.body.appendChild(span);
      // Construct clean SVG or HTML data URL
      const svgEl = span.querySelector('svg');
      let svgData = '';
      if (svgEl) {
        svgData = new XMLSerializer().serializeToString(svgEl);
      } else {
        svgData = '<svg xmlns="http://www.w3.org/2000/svg" width="600" height="200"><foreignObject width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml" style="font-size:24px;padding:20px">' + span.innerHTML + '</div></foreignObject></svg>';
      }
      document.body.removeChild(span);
      if (ext === 'svg') return utf8Encode(svgData);
      // For PNG, draw to canvas
      return new Promise(res => {
        const img = new Image();
        img.onload = () => {
          const c = document.createElement('canvas');
          c.width = Math.max(200, img.width || 600);
          c.height = Math.max(100, img.height || 200);
          const ctx = c.getContext('2d');
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, c.width, c.height);
          ctx.drawImage(img, 0, 0);
          c.toBlob(blob => {
            const fr = new FileReader();
            fr.onload = () => res(new Uint8Array(fr.result));
            fr.readAsArrayBuffer(blob);
          }, 'image/png');
        };
        img.src = 'data:image/svg+xml;utf8,' + encodeURIComponent(svgData);
      });
    }
    return utf8Encode(this.latex);
  }

  buildUi() {
    this.el.classList.add('mt-app-view');
    this.el.innerHTML = '';

    const wrap = h('div.mt-wrap');

    // Quick symbol & template palette header
    const topBar = h('div.mt-pal-top');
    const palTabs = h('div.mt-tabs');
    const palItems = h('div.mt-items');

    const drawPalette = () => {
      palItems.innerHTML = '';
      (SYMBOLS[this.curCat] || []).forEach(sym => {
        const b = h('button.mt-sym-b', { title: sym });
        KaTeX().then(k => { try { k.render(sym, b, { throwOnError: false }); } catch (e) { b.textContent = sym; } });
        b.addEventListener('click', () => this.insertLatex(sym.indexOf('a') >= 0 && /\\(hat|bar|vec|dot|ddot|tilde|acute|grave|check|breve)/.test(sym) ? sym.replace('{a}', '{}') : sym + ' '));
        palItems.appendChild(b);
      });
    };

    Object.keys(SYMBOLS).forEach(cat => {
      const t = h('button.mt-tab' + (cat === this.curCat ? '.on' : ''), cat);
      t.addEventListener('click', () => {
        this.curCat = cat;
        $$('.mt-tab', palTabs).forEach(x => x.classList.toggle('on', x === t));
        drawPalette();
      });
      palTabs.appendChild(t);
    });

    topBar.appendChild(palTabs);
    topBar.appendChild(palItems);
    wrap.appendChild(topBar);

    // Main workspace: Preview Canvas + Code Editor
    const work = h('div.mt-work');

    // Left/Top: Equation Canvas Preview
    const canvasWrap = h('div.mt-canvas-wrap');
    this.previewEl = h('div.mt-preview');
    this.errEl = h('div.mt-err');
    canvasWrap.appendChild(this.previewEl);
    canvasWrap.appendChild(this.errEl);

    // Right/Bottom: Code Editor & Quick Tool Strip
    const editWrap = h('div.mt-edit-wrap');
    const editHead = h('div.mt-edit-head',
      h('span.lbl', 'Mã LaTeX & MathML (Soạn trực tiếp hoặc dùng thanh công cụ):'),
      h('button.btn.small', { onclick: () => this.copyLatex() }, ic('copy', 14), ' Sao chép LaTeX'),
      h('button.btn.small', { onclick: () => this.copyMathML() }, ic('code', 14), ' Sao chép MathML')
    );

    this.input = h('textarea.mt-input', { spellcheck: false, rows: 4 });
    this.input.value = this.latex;
    this.input.addEventListener('input', () => {
      this.latex = this.input.value;
      this.markDirty();
      this.render();
    });

    editWrap.appendChild(editHead);
    editWrap.appendChild(this.input);

    work.appendChild(canvasWrap);
    work.appendChild(editWrap);
    wrap.appendChild(work);

    this.el.appendChild(wrap);
    drawPalette();

    // Status bar indicators
    this.sbStat = h('span.sb-item', 'MathType v' + VERSION);
    this.sbCount = h('span.sb-item', '0 ký hiệu');
    this.statusLeft.appendChild(this.sbStat);
    this.statusLeft.appendChild(this.sbCount);
  }

  ribbonDef() {
    return {
      tabs: [
        {
          id: 'home', label: 'Trang chủ', groups: [
            {
              label: 'Bảng tạm', items: [
                { label: 'Dán', icon: 'clipboard', run: () => this.pasteFromClip() },
                {
                  type: 'col', items: [
                    { label: 'Cắt', icon: 'scissors', run: () => this.cutText() },
                    { label: 'Sao chép LaTeX', icon: 'copy', run: () => this.copyLatex() }
                  ]
                }
              ]
            },
            {
              label: 'Lịch sử', items: [
                { label: 'Hoàn tác', icon: 'undo-2', run: () => this.undo() },
                { label: 'Làm lại', icon: 'redo-2', run: () => this.redo() },
                { label: 'Xóa trắng', icon: 'trash-2', run: () => { this.latex = ''; this.input.value = ''; this.render(); this.markDirty(); } }
              ]
            },
            {
              label: 'Cấu trúc nhanh', items: [
                { label: 'Phân số', icon: 'divide', run: () => this.insertLatex('\\frac{a}{b}') },
                { label: 'Căn thức', icon: 'radical', run: () => this.insertLatex('\\sqrt{x}') },
                { label: 'Mũ/Chỉ số', icon: 'superscript', run: () => this.insertLatex('x^{n}') },
                { label: 'Tích phân', icon: 'integral', run: () => this.insertLatex('\\int_{a}^{b} f(x)\\,dx') },
                { label: 'Tổng', icon: 'sigma', run: () => this.insertLatex('\\sum_{i=1}^{n} a_i') }
              ]
            },
            {
              label: 'Chèn vào tài liệu', items: [
                { label: 'Chèn vào Word', icon: 'file-text', run: () => this.insertToOtherDoc('word') },
                { label: 'Chèn vào Excel', icon: 'sheet', run: () => this.insertToOtherDoc('excel') },
                { label: 'Chèn vào PPT', icon: 'presentation', run: () => this.insertToOtherDoc('ppt') }
              ]
            }
          ]
        },
        {
          id: 'templates', label: 'Mẫu công thức', groups: [
            {
              label: 'Đại số & Giải tích', items: [
                { label: 'Nghiệm PT bậc 2', icon: 'function-square', run: () => this.insertLatex('x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}') },
                { label: 'Giới hạn lim', icon: 'move-right', run: () => this.insertLatex('\\lim_{x \\to x_0} f(x) = L') },
                { label: 'Đạo hàm d/dx', icon: 'divide', run: () => this.insertLatex('\\frac{df}{dx} = \\lim_{\\Delta x \\to 0} \\frac{f(x+\\Delta x)-f(x)}{\\Delta x}') }
              ]
            },
            {
              label: 'Ma trận & Hệ phương trình', items: [
                { label: 'Ma trận 2x2', icon: 'grid-2x2', run: () => this.insertLatex('\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}') },
                { label: 'Ma trận 3x3', icon: 'grid-3x3', run: () => this.insertLatex('\\begin{pmatrix} a_{11} & a_{12} & a_{13} \\\\ a_{21} & a_{22} & a_{23} \\\\ a_{31} & a_{32} & a_{33} \\end{pmatrix}') },
                { label: 'Hệ phương trình', icon: 'braces', run: () => this.insertLatex('\\begin{cases} a_1 x + b_1 y = c_1 \\\\ a_2 x + b_2 y = c_2 \\end{cases}') }
              ]
            },
            {
              label: 'Hình học & Vector', items: [
                { label: 'Vector', icon: 'move-up-right', run: () => this.insertLatex('\\vec{AB} = (x_B - x_A, y_B - y_A)') },
                { label: 'Tích có hướng', icon: 'crosshair', run: () => this.insertLatex('\\vec{u} \\times \\vec{v} = |\\vec{u}| |\\vec{v}| \\sin\\theta \\,\\vec{n}') }
              ]
            }
          ]
        },
        {
          id: 'library', label: 'Thư viện phương trình', groups: [
            {
              label: 'Phương trình mẫu', items: [
                {
                  label: 'Chọn mẫu phương trình…', icon: 'book-open', run: () => {
                    showMenu(EXAMPLES.map(e => ({
                      label: e[0],
                      run: () => { this.latex = e[1]; this.input.value = e[1]; this.render(); this.markDirty(); }
                    })), { x: 260, y: 120 });
                  }
                }
              ]
            }
          ]
        },
        {
          id: 'export', label: 'Xuất tệp', groups: [
            {
              label: 'Định dạng xuất', items: [
                { label: 'Lưu tệp LaTeX (*.tex)', icon: 'file-code', run: () => this.shell.save(this, true, 'tex') },
                { label: 'Lưu MathML (*.mml)', icon: 'code-2', run: () => this.shell.save(this, true, 'mml') },
                { label: 'Xuất ảnh SVG (*.svg)', icon: 'image', run: () => this.shell.save(this, true, 'svg') },
                { label: 'Xuất ảnh PNG (*.png)', icon: 'image', run: () => this.shell.save(this, true, 'png') }
              ]
            }
          ]
        }
      ]
    };
  }

  insertLatex(tex) {
    this.pushHistory();
    const inp = this.input;
    const s = inp.selectionStart, e = inp.selectionEnd;
    this.latex = inp.value.slice(0, s) + tex + inp.value.slice(e);
    inp.value = this.latex;
    const caret = tex.indexOf('{}') >= 0 ? s + tex.indexOf('{}') + 1 : s + tex.length;
    inp.setSelectionRange(caret, caret);
    inp.focus();
    this.markDirty();
    this.render();
  }

  render() {
    KaTeX().then(k => {
      try {
        k.render(this.latex || '\\;', this.previewEl, { throwOnError: true, displayMode: true });
        this.errEl.textContent = '';
      } catch (e) {
        try { k.render(this.latex, this.previewEl, { throwOnError: false, displayMode: true }); } catch (e2) {}
        this.errEl.textContent = 'Cú pháp: ' + String(e.message || e);
      }
    });
    this.updateStatus();
  }

  setZoom(z) {
    this.zoom = z;
    if (this.previewEl) this.previewEl.style.fontSize = (z / 100 * 2) + 'em';
  }

  updateStatus() {
    if (this.sbCount) {
      const tokens = (this.latex.match(/\\[A-Za-z]+|[-+*=^_{}]/g) || []).length;
      this.sbCount.textContent = tokens + ' ký hiệu toán';
    }
  }

  pushHistory(initial) {
    if (initial) {
      this.undoStack = [this.latex];
      this.redoStack = [];
      return;
    }
    this.undoStack.push(this.latex);
    if (this.undoStack.length > 50) this.undoStack.shift();
    this.redoStack = [];
  }

  undo() {
    if (this.undoStack.length > 1) {
      this.redoStack.push(this.undoStack.pop());
      this.latex = this.undoStack[this.undoStack.length - 1];
      this.input.value = this.latex;
      this.render();
      this.markDirty();
    }
  }

  redo() {
    if (this.redoStack.length) {
      const v = this.redoStack.pop();
      this.undoStack.push(v);
      this.latex = v;
      this.input.value = this.latex;
      this.render();
      this.markDirty();
    }
  }

  canUndo() { return this.undoStack.length > 1; }
  canRedo() { return this.redoStack.length > 0; }

  async copyLatex() {
    await Host.setClipboard(this.latex);
    toast('Đã sao chép mã LaTeX vào bộ nhớ tạm', 'ok');
  }

  async copyMathML() {
    const k = await KaTeX();
    const mml = k.renderToString(this.latex, { output: 'mathml', throwOnError: false });
    await Host.setClipboard(mml);
    toast('Đã sao chép công thức MathML', 'ok');
  }

  async pasteFromClip() {
    const text = await Host.getClipboard();
    if (text) this.insertLatex(text);
  }

  cutText() {
    const s = this.input.selectionStart, e = this.input.selectionEnd;
    if (s !== e) {
      const cut = this.input.value.slice(s, e);
      Host.setClipboard(cut);
      this.insertLatex('');
    }
  }

  async insertToOtherDoc(kind) {
    const docs = this.shell.docs.filter(d => d.kind === kind);
    if (!docs.length) {
      toast('Chưa có tài liệu ' + kind.toUpperCase() + ' nào đang mở. Hãy tạo tài liệu trước.', 'warn');
      return;
    }
    const target = docs[0];
    const imgHtml = await toImage(this.latex);
    if (kind === 'word') {
      target.insertHtml && target.insertHtml(imgHtml);
    } else if (kind === 'excel') {
      const cell = target.activeCell ? target.activeCell() : { r: 0, c: 0 };
      target.sheet && target.sheet.setCell && target.sheet.setCell(cell.r, cell.c, this.latex);
    } else if (kind === 'ppt') {
      target.addText && target.addText(this.latex);
    }
    this.shell.showDoc(target);
    toast('Đã chèn công thức vào ' + target.name, 'ok');
  }
}

// Auto-register MathDoc in the application registry
registerApp('mathtype', MathDoc);
