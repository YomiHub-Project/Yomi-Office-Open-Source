// v0.1.7: Windows build - the Web editor window is borderless like the native editor: the title bar draws the
// minimize / maximize / close buttons (46 x 40, same glyphs and hover colours as TitleBar.cs) and dragging the empty
// part of the title bar moves the window (WebView2 postMessage -> OfficeMainWindow).
import { h } from './util.js';
import Host from './host.js';

const SVG = {
  min: '<svg width="10" height="10" viewBox="0 0 10 10"><path d="M0 5.5h10" stroke="currentColor" stroke-width="1"/></svg>',
  max: '<svg width="10" height="10" viewBox="0 0 10 10"><rect x="0.5" y="0.5" width="9" height="9" fill="none" stroke="currentColor" stroke-width="1"/></svg>',
  restore: '<svg width="10" height="10" viewBox="0 0 10 10"><rect x="0.5" y="2.5" width="7" height="7" fill="none" stroke="currentColor"/><path d="M2.5 2.5V0.5h7v7h-2" fill="none" stroke="currentColor"/></svg>',
  close: '<svg width="10" height="10" viewBox="0 0 10 10"><path d="M0 0l10 10M10 0L0 10" stroke="currentColor" stroke-width="1.1"/></svg>'
};
export function initWinCaps() {
  const wv = window.chrome && window.chrome.webview;
  if (!wv || Host.kind !== 'windows') return;
  document.body.classList.add('win-frameless');
  const send = m => { try { wv.postMessage(m); } catch (e) {} };
  const b = (id, title) => { const x = h('button.wc-b.wc-' + id, { title, html: SVG[id] }); x.addEventListener('click', () => send(id)); return x; };
  const maxB = b('max', 'Phóng to');
  const caps = h('div.win-caps', b('min', 'Thu nhỏ'), maxB, b('close', 'Đóng'));
  document.body.appendChild(caps);
  wv.addEventListener('message', ev => {
    const max = ev.data === 'state:max'; if (ev.data !== 'state:max' && ev.data !== 'state:normal') return;
    document.body.classList.toggle('win-max', max); maxB.innerHTML = max ? SVG.restore : SVG.max; maxB.title = max ? 'Khôi phục' : 'Phóng to';
  });
  send('state');
  document.addEventListener('mousedown', e => {
    if (e.button !== 0 || e.clientY > 40) return;
    if (e.target.closest('button, input, select, textarea, a, label, [contenteditable], .tb-title, .tb-search, .win-caps, .menu, .dlg-back, .bs-main, .bs-nav')) return;
    if (!e.target.closest('.titlebar, .start, .win-drag')) return;
    e.preventDefault(); send(e.detail === 2 ? 'max' : 'drag');
  }, true);
}
