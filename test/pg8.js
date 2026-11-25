const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const dist = path.resolve(process.argv[2]), out = process.argv[3];
const MIME = { '.js': 'application/javascript', '.css': 'text/css', '.html': 'text/html', '.wasm': 'application/wasm', '.png': 'image/png', '.woff2': 'font/woff2' };
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split('?')[0]); if (p === '/') p = '/index.html'; const f = path.join(dist, p); if (!fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' }); fs.createReadStream(f).pipe(r); }).listen(0, async () => {
  const b = await chromium.launch(); const pg = await b.newPage({ viewport: { width: 1366, height: 800 } });
  pg.on('pageerror', e => console.log('[pageerror]', e.message)); pg.on('console', m => { if (m.type() === 'error' || m.text().startsWith('T8')) console.log(m.text().slice(0, 400)); });
  await pg.goto('http://127.0.0.1:' + srv.address().port + '/index.html#new=word'); await pg.waitForTimeout(1500);
  const r = await pg.evaluate(async () => {
    const d = OOS.cur; let html = '<h1>Báo cáo phân trang</h1>';
    for (let i = 1; i <= 60; i++) html += '<p>Đoạn ' + i + '. Văn bản mẫu tiếng Việt để kiểm tra việc tách trang, chia trang và hiển thị trang của Word trong Office Open Source phiên bản 0.0.8.</p>' + (i === 20 ? '<div class="wd-pagebreak" contenteditable="false"></div>' : '');
    html += '<table>' + Array(30).fill('<tr><td>Ô A</td><td>Ô B</td></tr>').join('') + '</table>';
    d.body.innerHTML = html; d.updatePages();
    // check: no block crosses a page content boundary
    const m = d.marg, P = d.pageH, G = 24, cH = P - m[0] - m[2]; const body = d.body; const bad = [];
    const top = el => { let y = 0, e = el; while (e && e !== body) { y += e.offsetTop; e = e.offsetParent; } return y; };
    const els = Array.from(body.querySelectorAll(':scope > p, :scope > h1, tr'));
    els.forEach(el => { const t = top(el), b2 = t + el.offsetHeight; const n = Math.floor(t / (P + G)); const ct = n * (P + G) + m[0], cb = ct + cH; if (t < ct - 1 || b2 > cb + 1) bad.push(el.textContent.slice(0, 20) + ' t=' + t.toFixed(0) + ' b=' + b2.toFixed(0) + ' page ' + (n + 1)); });
    // the paragraph after the break must start at a page top
    const brk = body.querySelector('.wd-pagebreak'); const nx = brk.nextElementSibling; const nt = top(nx); const onTop = Math.abs(((nt - m[0]) % (P + G))) < 2;
    return { pages: d.pageCount, bad: bad.slice(0, 5), nbad: bad.length, afterBreakAtTop: onTop, nextText: nx.textContent.slice(0, 10) };
  });
  console.log('T8 word', JSON.stringify(r));
  await pg.evaluate(() => { OOS.cur.scroller.scrollTop = 900; }); await pg.waitForTimeout(300); await pg.screenshot({ path: out + '_word_pages.png' });
  await pg.evaluate(() => OOS.cur.gotoPage(4)); await pg.waitForTimeout(300);
  console.log('T8 goto', await pg.evaluate(() => OOS.cur.curPage + '/' + OOS.cur.pageCount));
  // Excel
  await pg.evaluate(() => OOS.newDoc('excel')); await pg.waitForTimeout(2500); console.log('T8 kind', await pg.evaluate(() => OOS.cur && OOS.cur.kind));
  const x = await pg.evaluate(async () => {
    const d = OOS.cur, s = d.wb.cur; for (let r = 0; r < 120; r++) for (let c = 0; c < 14; c++) s.setCell(r, c, { v: r === 0 ? 'Cột ' + (c + 1) : r * 10 + c });
    s.pageSetup.printTitles = '1:1'; s.showPageBreaks = true; s.sel.r = 60; s.sel.c = 0;
    d.grid.redraw(); return 'ok';
  });
  await pg.waitForTimeout(300);
  const x2 = await pg.evaluate(() => { const d = OOS.cur, s = d.wb.cur; const pb = d.pageBreaks(s); return { pages: pb.pages.length, rows: pb.rowStarts.map(b => b.at), cols: pb.colStarts.map(b => b.at), scale: pb.scale }; });
  console.log('T8 excel auto', JSON.stringify(x2));
  await pg.evaluate(() => { const d = OOS.cur; d.wb.cur.sel.r = 30; d.wb.cur.sel.c = 5; }); 
  await pg.evaluate(() => { const b = Array.from(document.querySelectorAll('.rb-tab')).find(t => /Bố trí/.test(t.textContent)); b && b.click(); });
  await pg.waitForTimeout(300);
  await pg.evaluate(() => { const b = Array.from(document.querySelectorAll('.rb-btn, button')).find(t => /Ngắt trang/.test(t.textContent)); b && b.click(); });
  await pg.waitForTimeout(300); await pg.screenshot({ path: out + '_excel_menu.png' });
  await pg.evaluate(() => { const b = Array.from(document.querySelectorAll('.menu .mi, .menu button, .menu div')).find(t => /^Chèn ngắt trang/.test(t.textContent.trim())); b && b.click(); });
  await pg.waitForTimeout(400);
  const x3 = await pg.evaluate(() => { const d = OOS.cur, s = d.wb.cur; const pb = d.pageBreaks(s); return { manualR: s.pageSetup.rowBreaks, manualC: s.pageSetup.colBreaks, pages: pb.pages.length, rows: pb.rowStarts.map(b => b.at + (b.manual ? 'M' : '')), cols: pb.colStarts.map(b => b.at + (b.manual ? 'M' : '')) }; });
  console.log('T8 excel manual', JSON.stringify(x3));
  await pg.evaluate(() => { const g = OOS.cur.grid; g.scroller.scrollTop = 500; }); await pg.waitForTimeout(300);
  await pg.screenshot({ path: out + '_excel_pbp.png' });
  // xlsx round trip
  const x4 = await pg.evaluate(async () => { const d = OOS.cur; const bytes = await d.save('xlsx'); const JSZ = window.JSZip; const z = await JSZ.loadAsync(bytes); const sx = await z.file('xl/worksheets/sheet1.xml').async('string'); const wb = await z.file('xl/workbook.xml').async('string'); return { rb: /<rowBreaks[^]*?<\/rowBreaks>/.exec(sx) && /<rowBreaks[^]*?<\/rowBreaks>/.exec(sx)[0], titles: /Print_Titles[^<]*<\/definedName>|Print_Titles[^]*?<\/definedName>/.exec(wb) && /Print_Titles[^]*?<\/definedName>/.exec(wb)[0] }; });
  console.log('T8 xlsx', JSON.stringify(x4));
  // print preview
  await pg.evaluate(() => OOS.backstage(OOS.cur, 'print')); await pg.waitForTimeout(800); await pg.screenshot({ path: out + '_excel_print.png' });
  await b.close(); srv.close();
});
