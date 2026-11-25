const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const dist = path.resolve('dist-modern'); const file = fs.readFileSync(process.argv[2]).toString('base64');
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split('?')[0]); if (p === '/') p = '/index.html'; const f = path.join(dist, p); if (!fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': p.endsWith('.js') ? 'application/javascript' : p.endsWith('.html') ? 'text/html' : 'application/octet-stream' }); fs.createReadStream(f).pipe(r); }).listen(0, async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' }); const pg = await b.newPage();
  pg.on('pageerror', e => console.log('[pageerror]', e.message));
  await pg.goto('http://127.0.0.1:' + srv.address().port + '/index.html'); await pg.waitForTimeout(1200);
  const r = await pg.evaluate(async (b64) => { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    const d = await OOS.openFile({ name: 'saved_by_powerpoint.pptm', bytes: u, here: true }); await new Promise(r => setTimeout(r, 800));
    const V = await window.OOS_loadVBA(); const n0 = d.slides.length; await V.runMacro(d, 'TaoSlideTongKet', { quiet: true, onError: e => { throw e; } });
    return { slides: n0, after: d.slides.length, mods: (d.vba && d.vba.modules || []).map(m => m.name), texts: d.slides.map(s => s.objects.map(o => (o.html || '').replace(/<[^>]+>/g, ' ').slice(0, 40)).join('|')), x: d.slides[1].objects.map(o => [o.ph, Math.round(o.x), Math.round(o.w)]) };
  }, file);
  console.log(JSON.stringify(r)); await b.close(); srv.close();
});
