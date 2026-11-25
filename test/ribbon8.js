// Clicks every ribbon button of Word / Excel / PowerPoint / PDF and reports JS errors (item 15: toolbar items must work)
const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const dist = path.resolve(process.argv[2]);
const MIME = { '.js': 'application/javascript', '.css': 'text/css', '.html': 'text/html', '.wasm': 'application/wasm', '.png': 'image/png', '.woff2': 'font/woff2' };
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split('?')[0]); if (p === '/') p = '/index.html'; const f = path.join(dist, p); if (!fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' }); fs.createReadStream(f).pipe(r); }).listen(0, async () => {
  const b = await chromium.launch();
  const total = { clicks: 0, errors: [] };
  for (const app of (process.argv[3] || 'word,excel,ppt,pdf').split(',')) {
    const pg = await b.newPage({ viewport: { width: 1500, height: 860 } }); let cur = '';
    pg.on('pageerror', e => { if (!/Transport destroyed|cancelled/.test(e.message)) total.errors.push(app + ' › ' + cur + ': ' + e.message.slice(0, 160)); });
    pg.on('dialog', d => d.dismiss().catch(() => {}));
    pg.on('popup', p => p.close().catch(() => {}));
    await pg.goto('http://127.0.0.1:' + srv.address().port + '/index.html#new=' + app); await pg.waitForTimeout(1800);
    if (app === 'pdf') await pg.evaluate(async () => { const d = OOS.cur; const { PDFLib } = await import('data:text/javascript,export const PDFLib=1').catch(() => ({})); });
    if (app === 'pdf') await pg.evaluate(async () => { await new Promise(res => { const s = document.createElement('script'); s.src = 'vendor/pdf-lib.min.js'; s.onload = res; document.head.appendChild(s); }); const PL = window.PDFLib; const p = await PL.PDFDocument.create(); const f = await p.embedFont(PL.StandardFonts.Helvetica); for (let i = 0; i < 3; i++) p.addPage([595, 842]).drawText('Page ' + (i + 1), { x: 60, y: 760, size: 20, font: f }); await OOS.cur.load(await p.save(), 'pdf', 'x.pdf'); });
    if (app === 'word') await pg.evaluate(() => { OOS.cur.body.innerHTML = '<p>Xin chào Office Open Source</p><p>Dòng hai</p>'; const r = document.createRange(); r.selectNodeContents(OOS.cur.body.firstChild); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
    if (app === 'excel') await pg.evaluate(() => { const s = OOS.cur.wb.cur; for (let r = 0; r < 6; r++) for (let c = 0; c < 4; c++) s.setCell(r, c, { v: r === 0 ? 'H' + c : r * c }); OOS.cur.grid.redraw(); });
    const tabs = await pg.evaluate(() => Array.from(document.querySelectorAll('.rb-tab')).map(t => t.textContent.trim()).filter(t => t !== 'Tệp'));
    for (const tab of tabs) {
      await pg.evaluate(t => { const x = Array.from(document.querySelectorAll('.rb-tab')).find(z => z.textContent.trim() === t); x && x.click(); }, tab); await pg.waitForTimeout(150);
      const n = await pg.evaluate(() => document.querySelectorAll('.ribbon .rb-btn').length);
      for (let i = 0; i < n; i++) {
        const lab = await pg.evaluate(i => { const bt = document.querySelectorAll('.ribbon .rb-btn')[i]; if (!bt) return null; const l = (bt.textContent || bt.title || '').trim(); if (/^(In|Thoát|Đóng|Xuất|Lưu|Mở|Chia sẻ|Email|Đám mây|Tải|Quét|CamScanner|Ghi|Ghi âm|Ghi màn hình|Đọc chính tả|Camera|Máy ảnh|Trình chiếu|Từ đầu|Từ slide|Phản hồi|Đặt làm trình đọc PDF mặc định|Trợ lý A.I|Copilot.*|Soát lỗi A.I|Tóm tắt.*|Dịch trang)$/i.test(l)) return null; bt.click(); return l; }, i);
        if (lab == null) continue; cur = tab + ' › ' + lab; total.clicks++;
        await pg.waitForTimeout(120);
        await pg.evaluate(() => { document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); Array.from(document.querySelectorAll('.dlg-x')).forEach(x => x.click()); Array.from(document.querySelectorAll('.menu')).forEach(m => m.remove()); Array.from(document.querySelectorAll('.busy')).forEach(m => m.remove()); const bs = document.querySelector('.backstage .bs-back, .bs-close'); bs && bs.click(); });
        await pg.waitForTimeout(40);
        // tab may have changed (contextual tabs) — re-select
        await pg.evaluate(t => { const x = Array.from(document.querySelectorAll('.rb-tab')).find(z => z.textContent.trim() === t); x && !x.classList.contains('on') && x.click(); }, tab);
      }
    }
    await pg.close();
  }
  console.log('RIBBON clicks=' + total.clicks + ' errors=' + total.errors.length); total.errors.slice(0, 40).forEach(e => console.log('  ERR ' + e));
  await b.close(); srv.close();
});
