const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const dist = path.resolve(process.argv[2]), out = process.argv[3];
const MIME = { '.js': 'application/javascript', '.css': 'text/css', '.html': 'text/html', '.wasm': 'application/wasm', '.png': 'image/png', '.woff2': 'font/woff2' };
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split('?')[0]); if (p === '/') p = '/index.html'; const f = path.join(dist, p); if (!fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' }); fs.createReadStream(f).pipe(r); }).listen(0, async () => {
  const b = await chromium.launch(); const pg = await b.newPage({ viewport: { width: 1366, height: 800 } });
  pg.on('pageerror', e => console.log('[pageerror]', e.message)); pg.on('console', m => { if (m.type() === 'error' || m.text().startsWith('T8')) console.log(m.text().slice(0, 600)); });
  await pg.goto('http://127.0.0.1:' + srv.address().port + '/index.html#new=pdf'); await pg.waitForTimeout(1800);
  const r = await pg.evaluate(async () => {
    const d = OOS.cur; await new Promise(res => { const s = document.createElement('script'); s.src = 'vendor/pdf-lib.min.js'; s.onload = res; document.head.appendChild(s); }); const PL = window.PDFLib;
    const p = await PL.PDFDocument.create(); const f = await p.embedFont(PL.StandardFonts.Helvetica);
    for (let i = 1; i <= 3; i++) { const pp = p.addPage([595, 842]); pp.drawText('Chapter ' + i, { x: 60, y: 760, size: 28, font: f }); pp.drawText('Item', { x: 60, y: 700, size: 12, font: f }); pp.drawText('Qty', { x: 200, y: 700, size: 12, font: f }); pp.drawText('Price', { x: 300, y: 700, size: 12, font: f }); for (let k = 0; k < 5; k++) { pp.drawText('Row ' + k, { x: 60, y: 680 - k * 18, size: 12, font: f }); pp.drawText(String((k + 1) * i), { x: 200, y: 680 - k * 18, size: 12, font: f }); pp.drawText(String(k * 10 + i), { x: 300, y: 680 - k * 18, size: 12, font: f }); } }
    const form = p.getForm(); const tf = form.createTextField('ho_ten'); tf.addToPage(p.getPage(0), { x: 60, y: 400, width: 200, height: 20 }); tf.setText('Nguyen Van A'); const cb = form.createCheckBox('dong_y'); cb.addToPage(p.getPage(0), { x: 60, y: 360, width: 14, height: 14 }); cb.check();
    await d.load(await p.save(), 'pdf', 't.pdf'); await new Promise(r => setTimeout(r, 1000));
    const U = await OOS_v8.pdf(); const res = {}; const JSZ = window.JSZip || (await new Promise(res2 => { const s = document.createElement('script'); s.src = 'vendor/jszip.min.js'; s.onload = () => res2(window.JSZip); document.head.appendChild(s); }));
    const ppt = await U.toPptx(d, true); const zp = await JSZ.loadAsync(ppt); res.pptx = ppt.length + 'B slides=' + Object.keys(zp.files).filter(k => /slides\/slide\d+\.xml$/.test(k)).length;
    const xl = await U.toXlsx(d, true); const zx = await JSZ.loadAsync(xl.bytes); const s1 = await zx.file('xl/worksheets/sheet1.xml').async('string'); res.xlsx = xl.bytes.length + 'B sheets=' + xl.pages.length + ' row2=' + JSON.stringify(xl.pages[0][1]) + ' hasNum=' + /<v>11<\/v>/.test(s1);
    const pa = await U.makePdfA(d.bytes); const txt = new TextDecoder('latin1').decode(pa); res.pdfa = pa.length + 'B pdfaid=' + /pdfaid:part>2</.test(txt) + ' OutputIntent=' + /GTS_PDFA1/.test(txt) + ' ID=' + /\/ID\s*\[/.test(txt);
    // reopen PDF/A and check info panel detection
    const fonts = await U.listFonts(d); res.fonts = fonts.map(f => f.name + (f.embedded ? '+' : '-')).join(',');
    const info = await U.docInfo(d); res.info = info.version + ' ' + info.pages + 'p pdfa=' + (info.pdfa || 'no');
    // forms export/reset
    await U.resetForm(d); const PL2 = window.PDFLib; const pd = await PL2.PDFDocument.load(d.bytes); res.formReset = JSON.stringify(pd.getForm().getTextField('ho_ten').getText() || '') + ' cb=' + pd.getForm().getCheckBox('dong_y').isChecked();
    // crop / blank / reverse / repair
    await U.insertBlank(d); res.blank = d.pdf.numPages; await U.reverseOrder(d); res.rev = d.pdf.numPages; await U.repairPdf(d); res.repair = d.pdf.numPages;
    // sha256 fallback == subtle
    const C = await OOS_v8.cloud(); const enc = new TextEncoder().encode('Office Open Source 0.0.8 – PKCE'); const a = Array.from(C._sha256js(enc)).map(x => ('0' + x.toString(16)).slice(-2)).join(''); const bb = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', enc))).map(x => ('0' + x.toString(16)).slice(-2)).join(''); res.sha = a === bb;
    return res;
  });
  console.log('T8 conv', JSON.stringify(r, null, 1));
  await b.close(); srv.close();
});
