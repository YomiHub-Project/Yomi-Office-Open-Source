// Yomi-Office Open Source v0.1.3 — esbuild TypeScript & TSX Integration Engine
// Grounded on https://github.com/evanw/esbuild and https://esbuild.github.io/content-types/#ts-vs-tsx
import Host from './host.js';
import { dialog, toast, alertDlg } from './ui.js';
import { h, ic } from './util.js';

/**
 * esbuild Content Types: 'ts' vs 'tsx' Spec:
 * - 'ts': TypeScript without JSX. Type assertion syntax `<type>expr` is allowed.
 * - 'tsx': TypeScript with JSX. `<tag>` is always parsed as JSX element.
 *          `<type>expr` is FORBIDDEN; use `expr as type` instead.
 */

// Fallback in-browser pure-JS transformer when native server esbuild binary is unavailable (e.g. mobile/offline)
export function clientTransform(code, loader = 'ts', options = {}) {
  const isTsx = loader === 'tsx';
  let transformed = code;

  // Check for invalid <type>expr in tsx mode as specified by esbuild docs
  if (isTsx) {
    const invalidAssertion = /<([A-Z][a-zA-Z0-9_]*|string|number|boolean|any|unknown)>\s*([a-zA-Z0-9_$\(])/g;
    const match = invalidAssertion.exec(code);
    if (match && !/^(div|span|h[1-6]|p|table|tr|td|th|button|input|select|SmartCard|Widget|Card)$/.test(match[1])) {
      console.warn(`[esbuild tsx warning]: Type assertion <${match[1]}> detected in TSX mode. According to esbuild spec (https://esbuild.github.io/content-types/#ts-vs-tsx), you should use 'value as ${match[1]}' instead of '<${match[1]}>value'.`);
    }
  }

  // 1. Remove interfaces and type aliases
  transformed = transformed.replace(/export\s+interface\s+[\w\d_]+(?:\s*<[^>]+>)?(?:\s+extends\s+[^{]+)?\s*\{[\s\S]*?\}/g, '');
  transformed = transformed.replace(/interface\s+[\w\d_]+(?:\s*<[^>]+>)?(?:\s+extends\s+[^{]+)?\s*\{[\s\S]*?\}/g, '');
  transformed = transformed.replace(/export\s+type\s+[\w\d_]+(?:\s*<[^>]+>)?\s*=\s*[^;]+;/g, '');
  transformed = transformed.replace(/type\s+[\w\d_]+(?:\s*<[^>]+>)?\s*=\s*[^;]+;/g, '');

  // 2. Remove enums and convert to JS object
  transformed = transformed.replace(/(?:export\s+)?enum\s+([\w\d_]+)\s*\{([^}]+)\}/g, (m, name, body) => {
    const items = body.split(',').map(x => x.trim()).filter(Boolean);
    const obj = {};
    let val = 0;
    items.forEach(it => {
      const parts = it.split('=').map(x => x.trim());
      if (parts.length === 2) {
        val = isNaN(Number(parts[1])) ? parts[1].replace(/['"]/g, '') : Number(parts[1]);
        obj[parts[0]] = val;
      } else {
        obj[parts[0]] = val++;
      }
    });
    return `const ${name} = ${JSON.stringify(obj)};`;
  });

  // 3. Remove type assertions ('expr as Type' or '<Type>expr')
  transformed = transformed.replace(/\s+as\s+[A-Za-z0-9_<>\[\]|&\s]+/g, '');
  if (!isTsx) {
    // In 'ts' mode, strip <Type>expr
    transformed = transformed.replace(/<[A-Za-z0-9_<>\[\]|&\s]+>\s*([a-zA-Z0-9_$\(])/g, '$1');
  }

  // 4. Remove variable type annotations (const x: string = 'a')
  transformed = transformed.replace(/((?:const|let|var)\s+[\w\d_]+)\s*:\s*[A-Za-z0-9_<>\[\]|&]+(?:\s*=)/g, '$1 =');

  // 5. Remove function parameter and return types: (a: number, b: string): number =>
  transformed = transformed.replace(/\(([^)]*)\)\s*:\s*[A-Za-z0-9_<>\[\]|&]+\s*=>/g, '($1) =>');
  transformed = transformed.replace(/function(?:\s+[\w\d_]+)?\s*\(([^)]*)\)\s*:\s*[A-Za-z0-9_<>\[\]|&]+\s*\{/g, 'function $1 {');

  // Strip param types inside parentheses: (x: number, y: string)
  transformed = transformed.replace(/\(([^)]+)\)/g, (full, inner) => {
    const cleaned = inner.split(',').map(p => {
      return p.replace(/([\w\d_]+)\s*:\s*[^=,]+/g, '$1');
    }).join(',');
    return `(${cleaned})`;
  });

  // 6. If loader === 'tsx', transform JSX tags into h() calls
  if (isTsx) {
    const jsxFactory = options.jsxFactory || 'h';
    // Transform self-closing tags: <Tag prop="val" /> -> h('Tag', { prop: 'val' })
    transformed = transformed.replace(/<([A-Za-z0-9_\-]+)([^>]*?)\/>/g, (m, tag, attrs) => {
      const parsedAttrs = parseJsxAttrs(attrs);
      return `${jsxFactory}('${tag}', ${parsedAttrs})`;
    });

    // Transform open-close tags: <Tag prop="val">text</Tag> -> h('Tag', { prop: 'val' }, text)
    transformed = transformed.replace(/<([A-Za-z0-9_\-]+)([^>]*)>([\s\S]*?)<\/\1>/g, (m, tag, attrs, children) => {
      const parsedAttrs = parseJsxAttrs(attrs);
      const childExpr = parseJsxChildren(children, jsxFactory);
      return `${jsxFactory}('${tag}', ${parsedAttrs}${childExpr ? ', ' + childExpr : ''})`;
    });
  }

  return {
    code: transformed,
    loader,
    engine: 'esbuild-client-fallback'
  };
}

function parseJsxAttrs(str) {
  if (!str || !str.trim()) return 'null';
  const pairs = [];
  const re = /([A-Za-z0-9_\-]+)(?:=(?:"([^"]*)"|'([^']*)'|\{([^}]+)\}))?/g;
  let m;
  while ((m = re.exec(str))) {
    const key = m[1];
    let val = 'true';
    if (m[2] != null) val = JSON.stringify(m[2]);
    else if (m[3] != null) val = JSON.stringify(m[3]);
    else if (m[4] != null) val = m[4].trim();
    pairs.push(`"${key}": ${val}`);
  }
  return pairs.length ? `{ ${pairs.join(', ')} }` : 'null';
}

function parseJsxChildren(str, jsxFactory) {
  if (!str) return '';
  const trimmed = str.trim();
  if (!trimmed) return '';
  if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
    return trimmed.slice(1, -1);
  }
  if (trimmed.startsWith('<') && trimmed.endsWith('>')) {
    return trimmed;
  }
  return JSON.stringify(trimmed);
}

/**
 * Main Transform function with Server esbuild fallback to client transform
 */
export async function transformEsbuild(code, loader = 'ts', options = {}) {
  const start = performance.now();
  // Try server API first if running in desktop hybrid mode
  if (Host.kind === 'win' || Host.kind === 'node') {
    try {
      const res = await fetch('/api/esbuild_transform', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          code,
          loader,
          jsxFactory: options.jsxFactory || 'h',
          jsxFragment: options.jsxFragment || 'Fragment'
        })
      });
      if (res.ok) {
        const data = await res.json();
        if (data && data.ok) {
          data.timeMs = Math.round((performance.now() - start) * 100) / 100;
          data.engine = 'esbuild-native-go';
          return data;
        }
      }
    } catch (e) {
      // Fallback to client-side transformer
    }
  }

  const res = clientTransform(code, loader, options);
  res.timeMs = Math.round((performance.now() - start) * 100) / 100;
  return res;
}

/**
 * Execute transformed code inside Office Environment
 */
export async function executeEsbuildMacro(doc, code, loader = 'ts') {
  try {
    const transformed = await transformEsbuild(code, loader, { jsxFactory: 'h', jsxFragment: 'Fragment' });
    
    // Create execution context
    const appCtx = {
      Application: window.OOS,
      ActiveDocument: doc,
      ActiveSheet: (doc && doc.wb && doc.wb.cur) ? doc.wb.cur : null,
      Selection: (doc && doc.wb && doc.wb.cur) ? doc.wb.cur.sel : null,
      Range: (addr) => {
        if (!doc || !doc.wb || !doc.wb.cur) throw new Error('Không có bảng tính đang hoạt động');
        return doc.wb.cur.range ? doc.wb.cur.range(addr) : { address: addr };
      },
      h: h,
      toast: toast,
      alert: alertDlg
    };

    const func = new Function(
      'Application', 'ActiveDocument', 'ActiveSheet', 'Selection', 'Range', 'h', 'toast', 'alert',
      `"use strict";\n${transformed.code}`
    );

    const result = func(
      appCtx.Application, appCtx.ActiveDocument, appCtx.ActiveSheet, appCtx.Selection, appCtx.Range, appCtx.h, appCtx.toast, appCtx.alert
    );

    if (doc && doc.grid && doc.grid.redraw) doc.grid.redraw();
    if (doc && doc.render) doc.render();

    return { ok: true, result, transformed };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

/**
 * Open esbuild TypeScript / TSX Macro Editor Dialog
 */
export function openEsbuildEditor(doc, initialLoader = 'ts') {
  let curLoader = initialLoader;

  const samples = {
    ts: `// TypeScript (.ts) Macro for Yomi-Office Open Source
// Rule (esbuild): Type assertions <type>val, interfaces, types allowed!

interface SalesItem {
  name: string;
  quantity: number;
  unitPrice: number;
}

const items: SalesItem[] = [
  { name: "Sản phẩm A", quantity: 15, unitPrice: 250000 },
  { name: "Sản phẩm B", quantity: 40, unitPrice: 120000 },
  { name: "Dịch vụ C", quantity: 5, unitPrice: 850000 }
];

let totalRevenue: number = 0;
for (const item of items) {
  totalRevenue += item.quantity * item.unitPrice;
}

toast(\`Đã tính toán xong với esbuild: Tổng doanh thu = \${totalRevenue.toLocaleString('vi-VN')} VNĐ\`);
console.log("Total:", totalRevenue);
`,
    tsx: `// TypeScript + JSX (.tsx) Macro for Yomi-Office Open Source
// Rule (esbuild): <tag> is parsed as JSX! For type assertion use 'val as type'.

const title: string = "Thẻ Thống Kê KPI Doanh Nghiệp";
const targetValue: number = 185000000;

// JSX Element parsed and built by esbuild
const KpiCard = (
  <div style="background:#f8f9fc; border:1px solid #4e73df; border-radius:8px; padding:16px; width:300px; font-family:sans-serif;">
    <h3 style="margin:0 0 8px 0; color:#4e73df;">{title}</h3>
    <div style="font-size:24px; font-weight:bold; color:#2e59d9;">
      {targetValue.toLocaleString('vi-VN')} VNĐ
    </div>
    <p style="margin:8px 0 0 0; color:#1cc88a; font-weight:bold;">▲ Đạt 112% mục tiêu tháng</p>
  </div>
);

// If in Word, append smart card
if (ActiveDocument && ActiveDocument.editor) {
  const el = document.createElement('div');
  el.innerHTML = KpiCard;
  ActiveDocument.editor.appendChild(el);
  toast("Đã chèn JSX Component vào tài liệu Word!");
} else {
  toast("Đã biên dịch thành công JSX Widget với esbuild!");
}
`
  };

  const codeArea = h('textarea.esbuild-code', {
    style: {
      width: '100%',
      height: '240px',
      fontFamily: 'Consolas, Monaco, "Courier New", monospace',
      fontSize: '13px',
      lineHeight: '1.4',
      padding: '10px',
      background: '#1e1e1e',
      color: '#d4d4d4',
      border: '1px solid #333',
      borderRadius: '4px',
      resize: 'vertical',
      boxSizing: 'border-box'
    }
  });
  codeArea.value = samples[curLoader];

  const outArea = h('pre.esbuild-out', {
    style: {
      width: '100%',
      height: '140px',
      fontFamily: 'Consolas, Monaco, "Courier New", monospace',
      fontSize: '12px',
      lineHeight: '1.4',
      padding: '8px',
      background: '#0e0e0e',
      color: '#4ec9b0',
      border: '1px solid #222',
      borderRadius: '4px',
      overflow: 'auto',
      boxSizing: 'border-box'
    }
  }, '// Kết quả JavaScript được biên dịch bởi esbuild sẽ hiển thị tại đây...');

  const infoLabel = h('div.small.muted', { style: { marginBottom: '8px', color: '#666' } });
  function updateInfo() {
    if (curLoader === 'ts') {
      infoLabel.innerHTML = '<b>esbuild [loader: ts]</b>: TypeScript thuần. Hỗ trợ cú pháp ép kiểu <code>&lt;Type&gt;value</code>, interface, enum, types.';
    } else {
      infoLabel.innerHTML = '<b>esbuild [loader: tsx]</b>: TypeScript + JSX. Thẻ <code>&lt;Tag&gt;</code> được hiểu là JSX. Cấm <code>&lt;Type&gt;value</code> (dùng <code>value as Type</code> theo chuẩn <a href="https://esbuild.github.io/content-types/#ts-vs-tsx" target="_blank" style="color:#0f6cbd">esbuild ts-vs-tsx</a>).';
    }
  }
  updateInfo();

  const selLoader = h('select', { style: { padding: '4px 8px', borderRadius: '4px', border: '1px solid #ccc' } },
    h('option', { value: 'ts' }, 'TypeScript (.ts) — Macro logic thuần'),
    h('option', { value: 'tsx' }, 'TypeScript JSX (.tsx) — Component UI & Widget')
  );
  selLoader.value = curLoader;
  selLoader.addEventListener('change', () => {
    curLoader = selLoader.value;
    updateInfo();
    if (confirm('Bạn có muốn nạp mã mẫu cho định dạng ' + curLoader.toUpperCase() + ' không?')) {
      codeArea.value = samples[curLoader];
    }
  });

  const btnTransform = h('button.btn', { style: { marginRight: '6px' } }, ic('cpu', 14), ' Biên dịch esbuild');
  const btnRun = h('button.btn.btn-primary', { style: { marginRight: '6px' } }, ic('play', 14), ' Chạy Macro');
  const btnCopy = h('button.btn', { style: { marginRight: '6px' } }, ic('copy', 14), ' Sao chép JS');

  btnTransform.addEventListener('click', async () => {
    btnTransform.disabled = true;
    btnTransform.textContent = 'Đang biên dịch...';
    try {
      const res = await transformEsbuild(codeArea.value, curLoader);
      outArea.textContent = `// [esbuild ${res.engine}] Thời gian: ${res.timeMs}ms | Loader: ${res.loader}\n\n` + res.code;
      toast(`Biên dịch thành công với esbuild (${res.timeMs}ms)!`, 'ok');
    } catch (e) {
      outArea.textContent = '// Lỗi biên dịch esbuild:\n' + (e.message || String(e));
      toast('Lỗi biên dịch esbuild', 'err');
    } finally {
      btnTransform.disabled = false;
      btnTransform.innerHTML = '';
      btnTransform.appendChild(ic('cpu', 14));
      btnTransform.appendChild(document.createTextNode(' Biên dịch esbuild'));
    }
  });

  btnRun.addEventListener('click', async () => {
    btnRun.disabled = true;
    try {
      const res = await executeEsbuildMacro(doc, codeArea.value, curLoader);
      if (res.ok) {
        outArea.textContent = `// [esbuild] Thực thi thành công!\n` + (res.transformed ? res.transformed.code : '');
        toast('Đã thực thi Macro esbuild thành công!', 'ok');
      } else {
        outArea.textContent = '// Lỗi thực thi Macro:\n' + res.error;
        toast('Lỗi thực thi: ' + res.error, 'err');
      }
    } finally {
      btnRun.disabled = false;
    }
  });

  btnCopy.addEventListener('click', () => {
    navigator.clipboard.writeText(outArea.textContent).then(() => {
      toast('Đã sao chép mã JavaScript đầu ra');
    });
  });

  const body = h('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
    h('div', { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
      h('div', { style: { display: 'flex', alignItems: 'center', gap: '8px' } },
        h('span', { style: { fontWeight: 'bold' } }, 'Định dạng mã:'),
        selLoader
      ),
      h('span.badge', { style: { background: '#ffcf00', color: '#000', padding: '2px 8px', borderRadius: '4px', fontWeight: 'bold', fontSize: '11px' } }, 'esbuild v0.24')
    ),
    infoLabel,
    codeArea,
    h('div', { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '4px' } },
      h('span', { style: { fontWeight: 'bold', fontSize: '12px' } }, 'Mã JavaScript đầu ra (Output):'),
      h('div', null, btnTransform, btnRun, btnCopy)
    ),
    outArea
  );

  dialog({
    title: 'Trình biên dịch & Thực thi TypeScript / TSX (esbuild Engine)',
    body,
    width: 720,
    actions: [{ label: 'Đóng', default: true }]
  });
}
