const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const dist = path.resolve('dist-modern');
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split('?')[0]); if (p === '/') p = '/index.html'; const f = path.join(dist, p); if (!fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': p.endsWith('.js') ? 'application/javascript' : p.endsWith('.html') ? 'text/html' : p.endsWith('.css') ? 'text/css' : 'application/octet-stream' }); fs.createReadStream(f).pipe(r); }).listen(0, async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' }); const pg = await b.newPage();
  await pg.goto('http://127.0.0.1:' + srv.address().port + '/index.html'); await pg.waitForTimeout(1200);
  const b64 = await pg.evaluate(async () => {
    OOS.newDoc('ppt'); await new Promise(r => setTimeout(r, 1200)); const p = OOS.cur;
    p.vba = { modules: [{ name: 'Module1', code: 'Attribute VB_Name = "Module1"\nSub TaoSlideTongKet()\n    Dim sld As Object, shp As Object, i As Integer\n    Set sld = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutTitleOnly)\n    sld.Shapes.Title.TextFrame.TextRange.Text = "Tong ket"\n    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 140, 840, 300)\n    For i = 1 To ActivePresentation.Slides.Count - 1\n        shp.TextFrame.TextRange.InsertAfter "Slide " & i & vbCrLf\n    Next i\n    shp.TextFrame.TextRange.Font.Size = 24\nEnd Sub\nFunction SoSlide()\n    SoSlide = ActivePresentation.Slides.Count\nEnd Function\n' }], shortcuts: {} };
    const bytes = await p.save('pptx'); let s = ''; for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]); return btoa(s);
  });
  fs.writeFileSync('/mnt/user-data/outputs/oos_test_plain.pptx', Buffer.from(b64, 'base64')); console.log('ok', Buffer.from(b64, 'base64').length);
  await b.close(); srv.close();
});
