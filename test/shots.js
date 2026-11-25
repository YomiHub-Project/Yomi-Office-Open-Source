const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const dist = path.resolve(process.argv[2]), out = process.argv[3];
const MIME = { '.js': 'application/javascript', '.css': 'text/css', '.html': 'text/html', '.wasm': 'application/wasm', '.png': 'image/png', '.woff2': 'font/woff2' };
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split('?')[0]); if (p === '/') p = '/index.html'; const f = path.join(dist, p); if (!fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' }); fs.createReadStream(f).pipe(r); }).listen(0, async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' }); const pg = await b.newPage({ viewport: { width: 1366, height: 800 } });
  pg.on('pageerror', e => console.log('[pageerror]', e.message));
  await pg.goto('http://127.0.0.1:' + srv.address().port + '/index.html'); await pg.waitForTimeout(1500);
  await pg.screenshot({ path: out + '_1start.png' });
  await pg.evaluate(async () => { const S = await OOS_v7.scan(); const api = await S.openScanner(OOS); const c = document.createElement('canvas'); c.width = 1200; c.height = 900; const x = c.getContext('2d'); x.fillStyle = '#4a4036'; x.fillRect(0, 0, 1200, 900); x.fillStyle = '#efe9dc'; x.beginPath(); x.moveTo(250, 120); x.lineTo(960, 190); x.lineTo(920, 820); x.lineTo(190, 760); x.closePath(); x.fill(); x.fillStyle = '#222'; x.font = 'bold 40px Arial'; x.fillText('HỢP ĐỒNG KINH TẾ', 380, 300); x.font = '24px Arial'; for (let i = 0; i < 9; i++) x.fillText('Điều ' + (i + 1) + '. Nội dung điều khoản số ' + (i + 1) + ' .........', 300, 380 + i * 40); api.addCanvas(c); window.__scan = api; });
  await pg.waitForTimeout(600); await pg.screenshot({ path: out + '_2scan_crop.png' });
  await pg.evaluate(() => { [...document.querySelectorAll('.scan-tools .btn')].find(b => /Cắt/.test(b.textContent)).click(); });
  await pg.waitForTimeout(800); await pg.screenshot({ path: out + '_3scan_magic.png' });
  await pg.evaluate(() => { [...document.querySelectorAll('.scan-tools .btn')].find(b => /Đen trắng/.test(b.textContent)).click(); });
  await pg.waitForTimeout(800); await pg.screenshot({ path: out + '_4scan_bw.png' });
  await pg.evaluate(() => window.__scan.close());
  // reader
  await pg.evaluate(async () => { const R = await OOS_v7.reader(); const ch = []; for (let i = 1; i <= 3; i++) ch.push({ href: 'c' + i, title: 'Chương ' + i, html: '<h2>Chương ' + i + '</h2>' + Array(12).fill('<p>Trăm năm trong cõi người ta, chữ tài chữ mệnh khéo là ghét nhau. Trải qua một cuộc bể dâu, những điều trông thấy mà đau đớn lòng. Lạ gì bỉ sắc tư phong, trời xanh quen thói má hồng đánh ghen.</p>').join('') }); R.showReader(OOS, { meta: { title: 'Truyện Kiều', author: 'Nguyễn Du' }, chapters: ch, toc: ch.map(c => ({ title: c.title, href: c.href, level: 0 })) }, 'kieu.epub'); });
  await pg.waitForTimeout(600); await pg.screenshot({ path: out + '_5reader.png' });
  await pg.evaluate(() => { document.querySelector('.rd-b[title="Mục lục"]').click(); });
  await pg.waitForTimeout(300); await pg.screenshot({ path: out + '_6reader_toc.png' });
  await pg.evaluate(() => { document.querySelector('.rd-b[title^="Đóng"]').click(); });
  // AI pane in web mode
  await pg.evaluate(async () => { const A = await OOS_v7.ai(); const s = JSON.parse(localStorage.getItem('oos_settings') || '{}'); A.openAi(OOS, null); const sel = document.querySelector('.ai-prov'); sel.value = 'web:chatgpt'; sel.dispatchEvent(new Event('change')); });
  await pg.waitForTimeout(500); await pg.screenshot({ path: out + '_7ai_web.png' });
  // options page with update section
  await pg.evaluate(() => { document.querySelector('.ai-pane .pane-x[title="Đóng"]').click(); [...document.querySelectorAll('.rail-b')].find(b => /Tùy chọn/.test(b.textContent)).click(); });
  await pg.waitForTimeout(500); await pg.evaluate(() => { const h = [...document.querySelectorAll('.opt-sec h3')].find(x => /Cập nhật/.test(x.textContent)); h && h.scrollIntoView(); });
  await pg.waitForTimeout(300); await pg.screenshot({ path: out + '_8options_update.png' });
  await b.close(); srv.close();
});
