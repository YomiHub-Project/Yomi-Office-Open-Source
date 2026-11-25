const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const dist = path.resolve(process.argv[2]), out = process.argv[3];
const MIME = { '.js': 'application/javascript', '.css': 'text/css', '.html': 'text/html', '.wasm': 'application/wasm', '.png': 'image/png', '.woff2': 'font/woff2' };
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split('?')[0]); if (p === '/') p = '/index.html'; const f = path.join(dist, p); if (!fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' }); fs.createReadStream(f).pipe(r); }).listen(0, async () => {
  const b = await chromium.launch(); const pg = await b.newPage({ viewport: { width: 1366, height: 800 } });
  pg.on('pageerror', e => console.log('[pageerror]', e.message)); pg.on('console', m => { if (m.type() === 'error' || m.text().startsWith('T8')) console.log(m.text().slice(0, 600)); });
  await pg.goto('http://127.0.0.1:' + srv.address().port + '/index.html#new=pdf'); await pg.waitForTimeout(1800);
  const r = await pg.evaluate(async () => {
    const d = OOS.cur; const PL = window.PDFLib || (await new Promise(res => { const s = document.createElement('script'); s.src = 'vendor/pdf-lib.min.js'; s.onload = () => res(window.PDFLib); document.head.appendChild(s); }));
    const p = await PL.PDFDocument.create(); const f = await p.embedFont(PL.StandardFonts.Helvetica);
    for (let i = 1; i <= 6; i++) { const pp = p.addPage([595, 842]); pp.drawText('Chapter ' + i, { x: 60, y: 760, size: 28, font: f }); pp.drawText('Item     Qty     Price', { x: 60, y: 700, size: 12, font: f }); for (let k = 0; k < 5; k++) pp.drawText('Row' + k + '     ' + (k + 1) * i + '     ' + (k * 10 + i), { x: 60, y: 680 - k * 18, size: 12, font: f }); pp.drawText('Body text paragraph for page ' + i + '. Lorem ipsum dolor sit amet.', { x: 60, y: 560, size: 11, font: f }); }
    p.setTitle('Tai lieu thu'); p.setAuthor('OOS test');
    await d.load(await p.save(), 'pdf', 't.pdf'); await new Promise(r => setTimeout(r, 1200));
    const res = { pages: d.pdf.numPages };
    return res;
  });
  console.log('T8 load', JSON.stringify(r));
  await pg.screenshot({ path: out + '_1default.png' });
  // tools panel + info panel
  await pg.evaluate(() => { document.querySelector('.pdf-rail-b[data-k=info]').click(); document.querySelector('.pdf-rail-tools').click(); }); await pg.waitForTimeout(800);
  await pg.screenshot({ path: out + '_2info_tools.png' });
  // two-page layout via status bar
  await pg.evaluate(() => { const b = document.querySelector('.sb-view[title="Hai trang có trang bìa"]'); b && b.click(); }); await pg.waitForTimeout(1000);
  await pg.screenshot({ path: out + '_3twocover.png' });
  const nav = await pg.evaluate(() => { const d = OOS.cur; d.goto(3); return d.cur + ' inp=' + document.querySelector('.pdf-pg-inp').value + document.querySelector('.pdf-pg-tot').textContent; }); console.log('T8 nav', nav);
  // search panel
  await pg.evaluate(() => { document.querySelector('.pdf-rail-b[data-k=search]').click(); }); await pg.waitForTimeout(300);
  await pg.evaluate(() => { const i = document.querySelector('.pdf-nav-list input[type=search]'); i.value = 'Chapter 4'; i.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' })); }); await pg.waitForTimeout(1200);
  console.log('T8 search', await pg.evaluate(() => document.querySelector('.pdf-nav-list').textContent.slice(0, 80)));
  // ribbon tabs
  const tabs = await pg.evaluate(() => Array.from(document.querySelectorAll('.rb-tab')).map(t => t.textContent.trim()).join('|')); console.log('T8 tabs', tabs);
  await pg.evaluate(() => { Array.from(document.querySelectorAll('.rb-tab')).find(t => /^Xem$/.test(t.textContent.trim())).click(); }); await pg.waitForTimeout(400);
  await pg.screenshot({ path: out + '_4viewtab.png' });
  // conversions (no save)
  const conv = await pg.evaluate(async () => {
    const d = OOS.cur; const U = d._u8 || null; const m = window.__u8; return 'skip';
  });
  // print preview via backstage
  await pg.evaluate(() => { OOS.cur.printCfg = null; OOS.backstage(OOS.cur, 'print'); }); await pg.waitForTimeout(2500);
  await pg.screenshot({ path: out + '_5print.png' });
  await pg.evaluate(() => { const s = Array.from(document.querySelectorAll('.pdf-print-opts select')).find(x => Array.from(x.options).some(o => o.value === '4')); s.value = '4'; s.dispatchEvent(new Event('change')); }); await pg.waitForTimeout(2500);
  await pg.screenshot({ path: out + '_6print4up.png' });
  await b.close(); srv.close();
});
