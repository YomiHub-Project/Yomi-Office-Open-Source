// node test/selftest.js <distDir> [shotPrefix]
const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const dist = path.resolve(process.argv[2] || 'dist'), shot = process.argv[3];
const MIME = { '.js': 'application/javascript', '.css': 'text/css', '.html': 'text/html', '.wasm': 'application/wasm', '.json': 'application/json', '.woff2': 'font/woff2', '.png': 'image/png', '.svg': 'image/svg+xml' };
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split('?')[0].split('#')[0]); if (p === '/') p = '/index.html'; const f = path.join(dist, p); if (!fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' }); fs.createReadStream(f).pipe(r); }).listen(0, async () => {
  const port = srv.address().port; const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' }).catch(() => chromium.launch());
  const pg = await b.newPage({ viewport: { width: 1366, height: 800 } }); let done; const fin = new Promise(r => done = r); const lines = [];
  pg.on('console', m => { const t = m.text(); if (t.startsWith('OOSTEST')) { lines.push(t); console.log(t); if (t.startsWith('OOSTEST DONE')) done(); } else if (m.type() === 'error') console.log('[console.error]', t.slice(0, 300)); });
  pg.on('pageerror', e => console.log('[pageerror]', e.message));
  await pg.goto(`http://127.0.0.1:${port}/index.html${process.env.Q || ''}#selftest`);
  await Promise.race([fin, new Promise(r => setTimeout(r, +(process.env.T || 240000)))]);
  if (shot) await pg.screenshot({ path: shot + '.png' });
  await b.close(); srv.close();
  const f = lines.filter(l => l.includes('FAIL')).length; console.log('SUMMARY fail=' + f + ' total=' + lines.length); process.exit(f ? 1 : 0);
});
