// Yomi-Office Open Source v0.1.4 — application entry
import { Shell, applyTheme, applyMouseSettings, applySafeInsetsSettings } from './core/shell.js';
import Host from './core/host.js';
import './excel/excel.js';
import './word/word.js';
import './ppt/ppt.js';
import './pdf/pdf.js';
import './math/math.js';
import { runSelfTest } from './core/selftest.js';
import { initFonts } from './core/fonts.js';
import { syncNativeSettings } from './core/store.js';
import { initDexEngine } from './core/dex.js';

async function boot() {
  try { await syncNativeSettings(); } catch (e) {}
  applyTheme(); applyMouseSettings(); applySafeInsetsSettings(); initFonts(); initDexEngine();
  document.body.classList.add('host-' + Host.kind);
  const root = document.getElementById('app');
  const shell = new Shell(root);
  window.OOS = shell;
  const hash = location.hash.replace('#', '');
  const m = /^new=(word|excel|ppt|pdf|mathtype)/.exec(hash);
  const o = /^open=([\w-]+)/.exec(hash);
  const tl = /^tool=(scan|ai|pdfedit|reader|esbuild|devscan)/.exec(hash);
  const dx = /^dex=(guide|trackpad|presenter|tiled|desktop|tvhome|tvapps|multiwin|tvsettings)/.exec(hash);
  const demo = /^demo=(pdf|word|excel|cloud)/.exec(hash);
  if (dx) {
    if (dx[1] === 'guide') {
      shell.showStart();
      setTimeout(() => import('./core/dex.js').then(D => D.openDexGuideDialog()), 400);
    } else if (dx[1] === 'tvsettings') {
      shell.showStart();
      setTimeout(() => import('./core/dex.js').then(D => D.openDexControlCenterDialog('tvdesktop')), 400);
    } else if (dx[1] === 'trackpad') {
      shell.newDoc('word');
      setTimeout(() => import('./core/dex.js').then(D => { D.applyDexMode(true, true); D.openVirtualTrackpad(); }), 500);
    } else if (dx[1] === 'presenter') {
      shell.newDoc('ppt');
      setTimeout(() => import('./core/dex.js').then(D => { D.applyDexMode(true, true); D.startDualScreenPresenter(shell.cur); }), 600);
    } else if (dx[1] === 'tiled') {
      shell.newDoc('word');
      setTimeout(() => import('./core/dex.js').then(D => { D.applyDexMode(true, true); D.tileWindow('left'); }), 500);
    } else if (dx[1] === 'desktop') {
      shell.newDoc('word');
      setTimeout(() => import('./core/dex.js').then(D => D.applyDexMode(true, true)), 500);
    } else if (dx[1] === 'tvhome' || dx[1] === 'tvapps' || dx[1] === 'multiwin') {
      shell.showStart();
      setTimeout(() => import('./core/dex.js').then(D => {
        D.applyDexMode(true, true);
        D.openDexDesktopEnvironment({ mode: dx[1] });
      }), 400);
    }
  }
  else if (tl) {
    if (tl[1] === 'esbuild') {
      shell.newDoc('excel');
      setTimeout(() => import('./core/esbuild_runner.js').then(E => E.openEsbuildEditor(shell.activeDoc, 'tsx')), 1000);
    } else {
      shell.showStart();
      setTimeout(() => {
        if (tl[1] === 'scan') import('./scan/scan.js').then(S => S.openScanner(shell));
        else if (tl[1] === 'ai') import('./ai/ai.js').then(A => A.openAi(shell, null));
        else if (tl[1] === 'devscan') { const b = Array.prototype.slice.call(document.querySelectorAll('.rt')).find(x => /Trên thiết bị/.test(x.textContent)); if (b) b.click(); } // v0.1.8
        else if (tl[1] === 'pdfedit') { shell.newDoc('pdf'); setTimeout(() => { const t = Array.prototype.slice.call(document.querySelectorAll(".rb-tab")).find(x => /Chú thích/.test(x.textContent)); t && t.click(); }, 800); }
      }, 400);
    }
  }
  else if (demo) { shell.showStart(); setTimeout(() => import('./core/demo8.js').then(D => D.run(shell, demo[1])), 400); }
  else if (m) shell.newDoc(m[1]);
  else if (o) { shell.showStart(); Host.takeStash(o[1]).then(files => files.forEach(f => { f.here = true; shell.openFile(f); })); }
  else shell.showStart();
  if (hash === 'selftest') setTimeout(() => runSelfTest(shell), 600);
  else { import('./core/update.js').then(U => U.startupCheck()); import('./core/devscan.js').then(D => D.startupScan()); }
  Host.pending().then(files => files.forEach((f, i) => { if (i === 0 && !shell.docs.length) f.here = true; shell.openFile(f); }));
  const sp = document.getElementById('splash'); if (sp) setTimeout(() => sp.parentNode && sp.parentNode.removeChild(sp), 50);
}
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
// test/automation hook: load the VBA engine on demand (used by the self test and the Windows/Android test scripts)
window.OOS_loadVBA = () => import('./excel/vba/index.js');
import { FN as __FN } from './excel/formula.js';
window.OOS_functions = () => Object.keys(__FN).sort();
// v0.0.7 automation hooks (tests / screenshots)
window.OOS_v7 = { scan: () => import('./scan/scan.js'), reader: () => import('./reader/reader.js'), pdfpro: () => import('./pdf/pro.js'), ai: () => import('./ai/ai.js'), update: () => import('./core/update.js') };
// v0.0.8 automation hooks (tests / screenshots)
window.OOS_v8 = { pdf: () => import('./pdf/ui8.js'), cloud: () => import('./core/cloud.js'), pages: () => import('./excel/pages.js') };
// v0.1.3 automation hooks
window.OOS_v13 = { dex: () => import('./core/dex.js') };
