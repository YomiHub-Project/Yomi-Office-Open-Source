// v0.0.8 demo documents for screenshots / manual testing: #demo=pdf | word | excel | cloud
import { PDFLib } from './util.js';
const wait = ms => new Promise(r => setTimeout(r, ms));
async function docOf(shell, kind) { const prev = shell.cur; shell.newDoc(kind); for (let i = 0; i < 60 && !(shell.cur && shell.cur !== prev && shell.cur.kind === kind); i++) await wait(100); await wait(300); return shell.cur; }
export async function run(shell, what) {
  if (what === 'pdf') {
    const d = await docOf(shell, 'pdf'); const PL = await PDFLib(); const p = await PL.PDFDocument.create(); const f = await p.embedFont(PL.StandardFonts.Helvetica); const fb = await p.embedFont(PL.StandardFonts.HelveticaBold);
    for (let i = 1; i <= 8; i++) {
      const pg = p.addPage([595, 842]); pg.drawRectangle({ x: 0, y: 792, width: 595, height: 50, color: PL.rgb(0.85, 0.23, 0) });
      pg.drawText('Yomi-Office Open Source 0.0.8 - PDF Reader', { x: 40, y: 810, size: 16, font: fb, color: PL.rgb(1, 1, 1) });
      pg.drawText('Chapter ' + i + ': Sample document', { x: 40, y: 740, size: 24, font: fb });
      ['Item', 'Qty', 'Price', 'Total'].forEach((t, k) => pg.drawText(t, { x: 40 + k * 120, y: 690, size: 12, font: fb }));
      for (let r = 0; r < 8; r++) [String.fromCharCode(65 + r) + '-' + i, String(r + i), String((r + 1) * 1000), String((r + i) * (r + 1) * 1000)].forEach((t, k) => pg.drawText(t, { x: 40 + k * 120, y: 668 - r * 20, size: 11, font: f }));
      for (let l = 0; l < 14; l++) pg.drawText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.', { x: 40, y: 470 - l * 18, size: 11, font: f });
      pg.drawText('Page ' + i + ' / 8', { x: 270, y: 30, size: 10, font: f });
    }
    p.setTitle('Yomi-Office Open Source PDF demo'); p.setAuthor('Yomi-Office Open Source');
    await d.load(await p.save(), 'pdf', 'Demo PDF.pdf'); d.name = 'Demo PDF';
    const U = await import('../pdf/ui8.js'); await wait(500); U.togglePanel(d, 'info'); U.toggleTools(d, true); shell.updateTitle && shell.updateTitle();
  } else if (what === 'word') {
    const w = await docOf(shell, 'word'); let html = '<h1>Báo cáo phân trang Word</h1>';
    for (let i = 1; i <= 40; i++) html += '<p>Đoạn ' + i + '. Văn bản mẫu tiếng Việt để kiểm tra tách trang, hiển thị tách trang, đổi trang và chia trang trong Yomi-Office Open Source 0.0.8.</p>' + (i === 14 ? '<div class="wd-pagebreak" contenteditable="false"></div><h2>Phần 2 — bắt đầu ở trang mới</h2>' : '');
    w.body.innerHTML = html; w.updatePages(); w.markDirty(); await wait(300); w.gotoPage(1); w.scroller.scrollTop = Math.round((w.pageH - 260) * w.zoom / 100);
  } else if (what === 'excel') {
    const x = await docOf(shell, 'excel'); const s = x.wb.cur;
    for (let r = 0; r < 60; r++) for (let c = 0; c < 12; c++) s.setCell(r, c, { v: r === 0 ? 'Dữ liệu ' + (c + 1) : (r * 15 + c * 3) });
    s.pageSetup.printTitles = '1:1';
    s.pageSetup.printArea = 'A1:J40';
    s.pageSetup.rowBreaks = [20];
    s.pageSetup.colBreaks = [5];
    s.showPageBreaks = true;
    x.grid.redraw(); x.refreshUi();
    await wait(200);
    const tab = Array.prototype.slice.call(document.querySelectorAll('.rb-tab')).find(t => /Bố trí trang|Xem/.test(t.textContent));
    if (tab) tab.click();
  } else if (what === 'cloud') {
    const C = await import('./cloud.js'); C.accountsDialog();
  }
}
