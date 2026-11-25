const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const dist = path.resolve('dist-modern');
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split('?')[0]); if (p === '/') p = '/index.html'; const f = path.join(dist, p); if (!fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': p.endsWith('.js') ? 'application/javascript' : p.endsWith('.html') ? 'text/html' : p.endsWith('.css') ? 'text/css' : 'application/octet-stream' }); fs.createReadStream(f).pipe(r); }).listen(0, async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' }); const pg = await b.newPage();
  await pg.goto('http://127.0.0.1:' + srv.address().port + '/index.html'); await pg.waitForTimeout(1200);
  const out = await pg.evaluate(async () => {
    const wait = ms => new Promise(r => setTimeout(r, ms)); const b64 = u => { let s = ''; for (let i = 0; i < u.length; i++) s += String.fromCharCode(u[i]); return btoa(s); };
    OOS.newDoc('excel'); await wait(1000); const x = OOS.cur;
    x.wb.vba = { modules: [{ name: 'Module1', code: 'Attribute VB_Name = "Module1"\nSub TestAll()\n    Dim i As Long, t As Double\n    For i = 1 To 10\n        Cells(i, 1).Value = i * i\n        t = t + i * i\n    Next i\n    Range("B1").Value = t\n    Range("B2").Value = WorksheetFunction.Sum(Range("A1:A10"))\n    Range("B3").Value = "Xin ch" & ChrW(&HE0) & "o"\n    Range("B4").Value = msoTrue\nEnd Sub\n' }], shortcuts: {} };
    const xb = await x.save('xlsm');
    OOS.newDoc('word'); await wait(1000); const w = OOS.cur; w.body.innerHTML = '<p>xin chao</p>';
    w.vba = { modules: [{ name: 'NewMacros', code: 'Attribute VB_Name = "NewMacros"\nSub WTest()\n    ActiveDocument.Content.InsertAfter "Office Open Source " & msoTrue\nEnd Sub\n' }], shortcuts: {} };
    const wb = await w.save('docm');
    return { x: b64(xb), w: b64(wb) };
  });
  fs.writeFileSync('/mnt/user-data/outputs/oos_test_macro.xlsm', Buffer.from(out.x, 'base64')); fs.writeFileSync('/mnt/user-data/outputs/oos_test_macro.docm', Buffer.from(out.w, 'base64')); console.log('ok');
  await b.close(); srv.close();
});
