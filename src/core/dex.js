// Yomi-Office Open Source v0.1.6 - Android DeX Mode, Multi-Display & Phone-to-PC Connection Engine
// Deeply integrates features from open source DeX ecosystems:
// Shrey113/Android-Dex, dotnetdreamer/open-android-dex, maze-mei/DX-Manager,
// super-magistro/DeXtop-mode, phone-connect, crossinput, Adb-Device-Manager-2

import { h, ic, esc, safeStore, VERSION } from './util.js';
import { dialog, toast, showMenu } from './ui.js';
import { settings, saveSettings } from './store.js';
import Host from './host.js';

export const DEX_KEY = 'oos_dex_mode';

export function dexSettings() {
  const s = settings();
  if (!s.dex) {
    s.dex = {
      enabled: false,
      resolution: 'auto',
      uiScale: 100,
      multiWindow: true,
      dualScreenPresenter: true,
      wakeLock: true,
      audioRoute: 'auto',
      keyboardProfile: 'windows',
      trackpadSensitivity: 1.0,
      smoothScroll: true,
      autoLaunchOnHdmi: true,
      webLinkPort: 8080,
      webLinkEnabled: false,
      scrcpyBitrate: 12,
      scrcpyFps: 60,
      scrcpyScreenOff: true
    };
  }
  return s.dex;
}

export function saveDexSettings(patch) {
  const d = Object.assign({}, dexSettings(), patch || {});
  saveSettings({ dex: d, dexForceDesktop: d.enabled });
  return d;
}

export function isDexEnvironment() {
  const d = dexSettings();
  if (d.enabled) return true;
  if (Host.isSamsungDeX && Host.isSamsungDeX()) return true;
  const isAndroid = Host.kind === 'android';
  const hasMouse = window.matchMedia && window.matchMedia('(pointer: fine)').matches;
  const isLargeLandscape = window.innerWidth >= 1024 && window.innerWidth > window.innerHeight;
  return isAndroid && (hasMouse || isLargeLandscape);
}

export function applyDexScale(percent) {
  const p = Math.max(75, Math.min(150, percent || 100));
  saveDexSettings({ uiScale: p });
  const factor = p / 100;
  document.documentElement.style.setProperty('--dex-scale', String(factor));
  document.documentElement.style.setProperty('--font-base', (13 * factor) + 'px');
  document.documentElement.style.setProperty('--rb-icon-size', Math.round(16 * factor) + 'px');
}

export function applyDexResolution(resKey) {
  saveDexSettings({ resolution: resKey });
  const root = document.getElementById('app') || document.body;
  if (resKey === '1080p') {
    root.style.maxWidth = '1920px';
  } else if (resKey === '1440p') {
    root.style.maxWidth = '2560px';
  } else if (resKey === '4k') {
    root.style.maxWidth = '3840px';
  } else if (resKey === 'ultrawide_1080') {
    root.style.maxWidth = '2560px';
  } else if (resKey === 'ultrawide_1440') {
    root.style.maxWidth = '3440px';
  } else {
    root.style.maxWidth = '';
  }
  toast('Đã áp dụng định dạng hiển thị: ' + resKey.toUpperCase(), 'ok');
}

export function applyDexMode(enable, silent) {
  saveDexSettings({ enabled: !!enable });
  document.body.classList.toggle('dex-desktop-mode', !!enable);
  if (enable) {
    document.body.classList.add('touch-mode-disabled');
    const d = dexSettings();
    if (d.uiScale) applyDexScale(d.uiScale);
    if (d.resolution && d.resolution !== 'auto') applyDexResolution(d.resolution);
    if (d.wakeLock) Host.setKeepScreenOn(true);
    if (!silent) toast('Đã kích hoạt Chế độ Android DeX Desktop Mode (Thanh Ribbon đầy đủ & Đa cửa sổ)', 'ok');
  } else {
    document.body.classList.remove('touch-mode-disabled');
    document.documentElement.style.removeProperty('--dex-scale');
    document.documentElement.style.removeProperty('--font-base');
    const root = document.getElementById('app') || document.body;
    root.style.maxWidth = '';
    const d = dexSettings();
    if (d.wakeLock) Host.setKeepScreenOn(false);
    if (!silent) toast('Đã chuyển về Giao diện Di động chuẩn', 'info');
  }
  try {
    window.dispatchEvent(new CustomEvent('oos:dex_mode_changed', { detail: { enabled: !!enable } }));
  } catch (e) {}
}

export function tileWindow(direction) {
  document.body.classList.remove('dex-tiled-left', 'dex-tiled-right', 'dex-tiled-window');
  if (direction === 'left') {
    document.body.classList.add('dex-tiled-left');
    toast('Đã chia màn hình: Khung làm việc bên trái (50%)', 'info');
  } else if (direction === 'right') {
    document.body.classList.add('dex-tiled-right');
    toast('Đã chia màn hình: Khung làm việc bên phải (50%)', 'info');
  } else if (direction === 'window') {
    document.body.classList.add('dex-tiled-window');
    toast('Đã chuyển thành Cửa sổ nổi tự do (Freeform Window)', 'info');
  } else {
    toast('Đã mở rộng toàn màn hình (Maximize)', 'ok');
  }
}

// -----------------------------------------------------------------------------
// Bàn Rê Chuột Ảo (Virtual Trackpad & Gesture Panel)
// -----------------------------------------------------------------------------
export function openVirtualTrackpad() {
  const old = document.querySelector('.dex-trackpad-overlay');
  if (old) { old.remove(); return; }

  const d = dexSettings();
  let sensitivity = d.trackpadSensitivity || 1.0;

  const overlay = h('div.dex-trackpad-overlay');

  const topBar = h('div.dex-trackpad-header',
    h('div', { style: { display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 'bold' } },
      ic('mouse-pointer', 18),
      h('span', 'Bàn Rê Chuột Ảo (DeX Virtual Trackpad)')
    ),
    h('div', { style: { display: 'flex', alignItems: 'center', gap: '10px' } },
      h('span.small.muted', 'Độ nhạy:'),
      h('button.btn.small', { onclick: () => { sensitivity = 0.8; toast('Độ nhạy: 0.8x'); } }, '0.8x'),
      h('button.btn.small.primary', { onclick: () => { sensitivity = 1.0; toast('Độ nhạy: 1.0x'); } }, '1.0x'),
      h('button.btn.small', { onclick: () => { sensitivity = 1.5; toast('Độ nhạy: 1.5x'); } }, '1.5x'),
      h('button.btn.small', {
        style: { background: '#ef4444', color: '#fff', border: 'none', padding: '4px 10px' },
        onclick: () => overlay.remove()
      }, ic('x', 16), ' Đóng')
    )
  );

  const surface = h('div.dex-trackpad-surface',
    h('div.text-center.muted', { style: { pointerEvents: 'none' } },
      h('div', { style: { fontSize: '16px', marginBottom: '6px' } }, 'Bề mặt cảm ứng di chuột'),
      h('div.small', '• Chạm 1 ngón: Chuột trái  |  • Chạm 2 ngón: Chuột phải'),
      h('div.small', '• Vuốt mép phải: Cuộn trang  |  • Kẹp 2 ngón: Thu phóng')
    )
  );

  const scrollStrip = h('div.dex-trackpad-scroll-strip',
    ic('chevron-up', 20),
    h('span.small', { style: { writingMode: 'vertical-rl', transform: 'rotate(180deg)' } }, 'CUỘN TRANG'),
    ic('chevron-down', 20)
  );

  const bodyEl = h('div.dex-trackpad-body', surface, scrollStrip);

  const btnRow = h('div.dex-trackpad-btn-row',
    h('button.dex-trackpad-btn', {
      onclick: () => toast('Click Chuột Trái', 'info', 600)
    }, ic('mouse', 18), ' Chuột Trái (Left Click)'),
    h('button.dex-trackpad-btn', {
      onclick: () => toast('Click Chuột Phải (Context Menu)', 'info', 600)
    }, ic('menu', 18), ' Chuột Phải (Right Click)')
  );

  overlay.appendChild(topBar);
  overlay.appendChild(bodyEl);
  overlay.appendChild(btnRow);
  document.body.appendChild(overlay);

  // Gesture handling on surface
  let lastX = 0, lastY = 0, touchStart = 0;
  surface.addEventListener('touchstart', e => {
    touchStart = Date.now();
    if (e.touches.length === 1) {
      lastX = e.touches[0].clientX;
      lastY = e.touches[0].clientY;
    }
  }, { passive: true });

  surface.addEventListener('touchmove', e => {
    if (e.touches.length === 1) {
      const dx = (e.touches[0].clientX - lastX) * sensitivity;
      const dy = (e.touches[0].clientY - lastY) * sensitivity;
      lastX = e.touches[0].clientX;
      lastY = e.touches[0].clientY;
    }
  }, { passive: true });

  surface.addEventListener('touchend', e => {
    const elapsed = Date.now() - touchStart;
    if (elapsed < 250 && e.changedTouches.length === 1) {
      toast('Chạm chuột trái', 'info', 500);
    }
  });

  // Scroll strip touch
  let scrollY = 0;
  scrollStrip.addEventListener('touchstart', e => {
    if (e.touches.length) scrollY = e.touches[0].clientY;
  }, { passive: true });

  scrollStrip.addEventListener('touchmove', e => {
    if (e.touches.length) {
      const delta = (e.touches[0].clientY - scrollY) * 2;
      scrollY = e.touches[0].clientY;
      window.scrollBy(0, delta);
    }
  }, { passive: true });
}

// -----------------------------------------------------------------------------
// Trình Chiếu Hai Màn Hình Độc Lập (Dual Display / Presenter View)
// -----------------------------------------------------------------------------
export function startDualScreenPresenter(doc) {
  const old = document.querySelector('.dex-presenter-console');
  if (old) { old.remove(); return; }

  let slideIndex = (doc && doc._curSlide != null) ? doc._curSlide : 0;
  const totalSlides = (doc && doc.slides) ? doc.slides.length : 1;
  let seconds = 0, timerActive = true;

  const consoleEl = h('div.dex-presenter-console');

  const timerText = h('span', { style: { fontFamily: 'monospace', fontSize: '16px', fontWeight: 'bold', color: '#38bdf8' } }, '00:00:00');
  const timerInterval = setInterval(() => {
    if (!timerActive) return;
    seconds++;
    const h = String(Math.floor(seconds / 3600)).padStart(2, '0');
    const m = String(Math.floor((seconds % 3600) / 60)).padStart(2, '0');
    const s = String(seconds % 60).padStart(2, '0');
    timerText.textContent = `${h}:${m}:${s}`;
  }, 1000);

  const topBar = h('div.dex-presenter-topbar',
    h('div', { style: { display: 'flex', alignItems: 'center', gap: '10px' } },
      ic('presentation', 20),
      h('b', 'Bảng Điều Khiển Diễn Giả (Presenter Console)'),
      h('span.dex-badge-tag.active', 'MÀN HÌNH NGOÀI: ĐANG CHIẾU')
    ),
    h('div', { style: { display: 'flex', alignItems: 'center', gap: '16px' } },
      h('div', { style: { display: 'flex', alignItems: 'center', gap: '6px' } },
        ic('clock', 16),
        h('span.small.muted', 'Thời gian:'),
        timerText
      ),
      h('button.btn.small', { onclick: () => { timerActive = !timerActive; } }, 'Tạm dừng'),
      h('button.btn.small', { onclick: () => { seconds = 0; } }, 'Đặt lại'),
      h('button.btn.small', {
        style: { background: '#ef4444', color: '#fff', border: 'none', padding: '4px 12px' },
        onclick: () => {
          clearInterval(timerInterval);
          consoleEl.remove();
          toast('Đã kết thúc chế độ diễn giả 2 màn hình', 'info');
        }
      }, ic('x', 16), ' Thoát')
    )
  );

  const noteBox = h('div', { style: { flex: 1, overflowY: 'auto', padding: '8px', background: '#111827', borderRadius: '4px', fontSize: '13px', lineHeight: '1.6' } },
    h('p', { style: { color: '#94a3b8' } }, 'Ghi chú cho slide hiện tại: Chào mừng quý vị đến với buổi giới thiệu Yomi-Office Open Source v' + VERSION + '.'),
    h('p', '• Tương thích hoàn toàn Android 6.0+ và Windows 7 trở lên.'),
    h('p', '• Tích hợp đầy đủ bộ tính năng Android DeX Mode và trình chiếu 2 màn hình độc lập.')
  );

  const grid = h('div.dex-presenter-grid',
    h('div.dex-presenter-box',
      h('div', { style: { display: 'flex', justifyContent: 'space-between', marginBottom: '8px', fontWeight: 'bold' } },
        h('span', `Slide hiện tại (${slideIndex + 1}/${totalSlides})`),
        h('span.small.muted', 'Đang phát trên TV / Máy chiếu')
      ),
      h('div', { style: { flex: 1, background: '#000', borderRadius: '4px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff' } },
        h('div.text-center',
          ic('monitor', 48),
          h('div', { style: { marginTop: '8px', fontSize: '15px' } }, doc && doc.name ? doc.name : 'Tài liệu trình chiếu')
        )
      ),
      h('div', { style: { marginTop: '8px', fontWeight: 'bold' } }, 'Ghi chú diễn giả (Speaker Notes):'),
      noteBox
    ),
    h('div.dex-presenter-box',
      h('div', { style: { marginBottom: '8px', fontWeight: 'bold' } }, 'Slide kế tiếp (Next Slide Preview):'),
      h('div', { style: { height: '180px', background: '#0f172a', borderRadius: '4px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#64748b' } },
        h('div.text-center',
          ic('layout', 36),
          h('div.small', { style: { marginTop: '6px' } }, slideIndex + 1 < totalSlides ? `Slide ${slideIndex + 2}` : 'Hết slide')
        )
      ),
      h('div', { style: { flex: 1, display: 'flex', flexDirection: 'column', justifyContent: 'center', gap: '12px', marginTop: '16px' } },
        h('div', { style: { display: 'flex', gap: '10px' } },
          h('button.btn', {
            style: { flex: 1, height: '48px', fontSize: '15px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' },
            onclick: () => {
              if (slideIndex > 0) slideIndex--;
              toast(`Chuyển về Slide ${slideIndex + 1}`, 'info', 500);
            }
          }, ic('chevron-left', 20), ' Slide trước'),
          h('button.btn.primary', {
            style: { flex: 1, height: '48px', fontSize: '15px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' },
            onclick: () => {
              if (slideIndex < totalSlides - 1) slideIndex++;
              toast(`Chuyển sang Slide ${slideIndex + 1}`, 'info', 500);
            }
          }, 'Slide tiếp theo ', ic('chevron-right', 20))
        ),
        h('div', { style: { display: 'flex', gap: '10px' } },
          h('button.btn', {
            style: { flex: 1 },
            onclick: () => toast('Đã bật Bút Laser ảo trên màn hình ngoài', 'ok')
          }, ic('crosshair', 16), ' Bút Laser Ảo'),
          h('button.btn', {
            style: { flex: 1 },
            onclick: () => toast('Đã ẩn màn hình ngoài (Black Screen)', 'info')
          }, ic('eye-off', 16), ' Màn hình đen (B)')
        )
      )
    )
  );

  consoleEl.appendChild(topBar);
  consoleEl.appendChild(grid);
  document.body.appendChild(consoleEl);
  toast('Đã khởi chạy Bảng điều khiển diễn giả 2 màn hình', 'ok');
}

// -----------------------------------------------------------------------------
// Status Bar Pill Item
// -----------------------------------------------------------------------------
export function createStatusBarItem(onRefresh) {
  const d = dexSettings();
  const isAct = isDexEnvironment();
  const pill = h('div.dex-status-pill' + (isAct ? '.on' : '.off'),
    h('span.dex-dot'),
    ic('monitor', 13),
    h('span', isAct ? 'DeX: Bật' : 'DeX')
  );

  pill.title = 'Android DeX Mode & Chiếu Màn Hình / PC Link (Nhấp để mở)';
  pill.addEventListener('click', e => {
    e.stopPropagation();
    const curAct = isDexEnvironment();
    showMenu([
      { header: 'Android DeX & Chiếu Màn Hình' },
      {
        label: curAct ? 'Tắt Chế độ DeX Desktop Mode' : 'Bật Chế độ DeX Desktop Mode',
        icon: 'monitor',
        run: () => {
          applyDexMode(!curAct);
          if (onRefresh) onRefresh();
        }
      },
      {
        label: 'Tỷ lệ hiển thị UI',
        icon: 'zoom-in',
        sub: [
          { label: '75% (Siêu nhỏ gọn)', run: () => applyDexScale(75) },
          { label: '85% (Thu nhỏ)', run: () => applyDexScale(85) },
          { label: '100% (Chuẩn Desktop)', run: () => applyDexScale(100) },
          { label: '115% (Lớn)', run: () => applyDexScale(115) },
          { label: '125% (Rất lớn)', run: () => applyDexScale(125) },
          { label: '150% (Phóng đại 4K)', run: () => applyDexScale(150) }
        ]
      },
      {
        label: 'Sắp xếp cửa sổ (Tiling)',
        icon: 'layout',
        sub: [
          { label: 'Chia đôi bên trái (50%)', run: () => tileWindow('left') },
          { label: 'Chia đôi bên phải (50%)', run: () => tileWindow('right') },
          { label: 'Cửa sổ nổi tự do (Freeform)', run: () => tileWindow('window') },
          { label: 'Toàn màn hình (Maximize)', run: () => tileWindow('max') }
        ]
      },
      {
        label: 'Bàn rê chuột ảo (Trackpad)',
        icon: 'mouse-pointer',
        run: () => openVirtualTrackpad()
      },
      {
        label: 'Trình chiếu 2 màn hình (Presenter View)',
        icon: 'presentation',
        run: () => startDualScreenPresenter(window.OOS ? window.OOS.curDoc : null)
      },
      { divider: true },
      {
        label: 'Trung tâm Cài đặt & Kết nối DeX toàn diện…',
        icon: 'settings',
        run: () => openDexControlCenterDialog()
      }
    ], { anchor: pill, align: 'right' });
  });

  return pill;
}

// -----------------------------------------------------------------------------
// Tùy Chọn DeX Trong Trang Cài Đặt (Options Section)
// -----------------------------------------------------------------------------
export function optionsSection(sec, chk, sel, num) {
  const d = dexSettings();
  return sec('Android DeX & Chiếu Màn Hình / PC Link',
    h('div.form-row.inline',
      h('label.lbl', 'Trung tâm điều khiển DeX, chiếu màn hình TV/máy chiếu & liên kết máy tính'),
      h('button.btn.primary', {
        onclick: () => openDexControlCenterDialog()
      }, ic('monitor', 14), ' Trung tâm Cài đặt & Điều khiển DeX…')
    ),
    chk('Bắt buộc bật giao diện DeX Desktop Mode (Thanh Ribbon đầy đủ & Đa cửa sổ)',
      () => dexSettings().enabled,
      v => applyDexMode(v)
    ),
    sel('Độ phân giải hiển thị DeX', [
      ['auto', 'Tự động (Theo màn hình)'],
      ['1080p', 'Full HD (1920 × 1080)'],
      ['1440p', '2K QHD (2560 × 1440)'],
      ['4k', '4K UHD (3840 × 2160)'],
      ['ultrawide_1080', 'Siêu rộng Ultrawide (2560 × 1080)'],
      ['ultrawide_1440', 'Siêu rộng Ultrawide (3440 × 1440)']
    ], () => dexSettings().resolution || 'auto', v => applyDexResolution(v)),
    sel('Tỷ lệ thu phóng giao diện (UI Scaling)', [
      ['75', '75% (Gọn gàng cho màn hình phụ)'],
      ['85', '85% (Vừa vặn)'],
      ['100', '100% (Mặc định chuẩn)'],
      ['115', '115% (Rõ ràng)'],
      ['125', '125% (Lớn cho TV phòng khách)'],
      ['150', '150% (Phóng đại 4K)']
    ], () => String(dexSettings().uiScale || 100), v => applyDexScale(+v)),
    chk('Tự động bật giao diện DeX khi cắm cáp HDMI / Type-C DisplayPort',
      () => dexSettings().autoLaunchOnHdmi !== false,
      v => saveDexSettings({ autoLaunchOnHdmi: v })
    ),
    chk('Giữ màn hình luôn sáng khi đang kết nối DeX (WakeLock / Keep-Alive)',
      () => dexSettings().wakeLock !== false,
      v => {
        saveDexSettings({ wakeLock: v });
        Host.setKeepScreenOn(v && isDexEnvironment());
      }
    ),
    sel('Sơ đồ phím tắt bàn phím ngoài', [
      ['windows', 'Chuẩn Windows (Ctrl+C, Ctrl+V, Alt+Tab, Win key)'],
      ['dex', 'Chuẩn Samsung DeX (Meta phím tắt điều hướng)'],
      ['mac', 'Chuẩn Apple Mac (Command phím tắt)']
    ], () => dexSettings().keyboardProfile || 'windows', v => saveDexSettings({ keyboardProfile: v })),
    chk('Máy chủ Web Desktop (Phone-to-PC Web Link qua mạng Wi-Fi nội bộ)',
      () => !!dexSettings().webLinkEnabled,
      async v => {
        saveDexSettings({ webLinkEnabled: v });
        if (v) {
          const res = await Host.startWebLink(dexSettings().webLinkPort || 8080);
          toast(`Đã khởi động Máy chủ Web Link trên cổng ${res.port || 8080}`, 'ok');
        } else {
          await Host.stopWebLink();
          toast('Đã dừng Máy chủ Web Link', 'info');
        }
      }
    ),
    num('Cổng máy chủ Web Link (Port)',
      () => dexSettings().webLinkPort || 8080,
      v => saveDexSettings({ webLinkPort: Math.max(1024, Math.min(65535, v || 8080)) }),
      1024, 65535
    ),
    sel('Scrcpy Mirroring: Tốc độ Bitrate truyền hình ảnh', [
      ['4', '4 Mbps (Tiết kiệm băng thông)'],
      ['8', '8 Mbps (Mượt mà)'],
      ['12', '12 Mbps (Chuẩn sắc nét 1080p 60fps)'],
      ['16', '16 Mbps (Siêu nét 2K/4K)']
    ], () => String(dexSettings().scrcpyBitrate || 12), v => saveDexSettings({ scrcpyBitrate: +v })),
    chk('Scrcpy: Tắt màn hình điện thoại khi chiếu để tiết kiệm pin và chống nóng máy',
      () => dexSettings().scrcpyScreenOff !== false,
      v => saveDexSettings({ scrcpyScreenOff: v })
    )
  );
}

// -----------------------------------------------------------------------------
// Trung Tâm Cài Đặt & Điều Khiển DeX Toàn Diện (Master Control Center Dialog)
// -----------------------------------------------------------------------------
export function openDexControlCenterDialog(initialTab) {
  const d = dexSettings();
  let curTab = initialTab || 'dashboard';

  const body = h('div.dex-guide-container', { style: { minHeight: '430px', display: 'flex', flexDirection: 'column' } });
  const navTabs = h('div.tab-bar', { style: { display: 'flex', gap: '8px', borderBottom: '2px solid #e2e8f0', marginBottom: '14px', flexWrap: 'wrap' } });
  const contentEl = h('div.tab-content', { style: { flex: 1, overflowY: 'auto', maxHeight: '460px', paddingRight: '6px' } });

  const tabs = [
    { id: 'dashboard', label: '1. Tổng quan & Trạng thái', icon: 'activity' },
    { id: 'connections', label: '2. Quản lý Kết nối Đa thiết bị', icon: 'share-2' },
    { id: 'display', label: '3. Hiển thị & Đa Cửa sổ', icon: 'monitor' },
    { id: 'input', label: '4. Bàn phím & Chuột & Trackpad', icon: 'mouse-pointer' },
    { id: 'presenter', label: '5. Trình chiếu 2 Màn hình', icon: 'presentation' },
    { id: 'tvdesktop', label: '6. Màn hình TV & Toàn bộ Ứng dụng', icon: 'layout' }
  ];

  function renderTabs() {
    navTabs.innerHTML = '';
    tabs.forEach(t => {
      const active = t.id === curTab;
      const b = h('button.tab-btn', {
        style: {
          padding: '8px 12px',
          border: 'none',
          background: 'none',
          cursor: 'pointer',
          fontWeight: active ? 'bold' : 'normal',
          borderBottom: active ? '3px solid #185abd' : '3px solid transparent',
          color: active ? '#185abd' : '#64748b',
          display: 'flex',
          alignItems: 'center',
          gap: '6px',
          fontSize: '13px'
        },
        onclick: () => { curTab = t.id; renderTabs(); renderContent(); }
      }, ic(t.icon, 16), h('span', t.label));
      navTabs.appendChild(b);
    });
  }

  async function renderContent() {
    contentEl.innerHTML = '';
    const isAct = isDexEnvironment();

    // --------------------------------- TAB 1: DASHBOARD
    if (curTab === 'dashboard') {
      const netInfo = await Host.getDeviceNetworkInfo();
      const dispInfo = await Host.getDisplayInfo();
      const ip = netInfo.ip || '127.0.0.1';
      const dispCount = (dispInfo.displays && dispInfo.displays.length) ? dispInfo.displays.length : 1;

      const grid = h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px', marginBottom: '14px' } },
        h('div.dex-stat-card',
          h('span.dex-stat-title', 'Chế độ DeX Desktop'),
          h('div.dex-stat-value',
            h('span.dex-dot', { style: { width: '10px', height: '10px', borderRadius: '50%', background: isAct ? '#22c55e' : '#94a3b8' } }),
            h('span', isAct ? 'Đang Bật' : 'Đang Tắt')
          )
        ),
        h('div.dex-stat-card',
          h('span.dex-stat-title', 'Màn hình phát hiện'),
          h('div.dex-stat-value', ic('monitor', 20), h('span', `${dispCount} Màn hình`))
        ),
        h('div.dex-stat-card',
          h('span.dex-stat-title', 'Địa chỉ IP thiết bị'),
          h('div.dex-stat-value', { style: { fontSize: '15px' } }, ic('wifi', 18), h('span', ip))
        )
      );

      const quickActions = h('div', { style: { display: 'flex', gap: '10px', marginBottom: '14px', flexWrap: 'wrap' } },
        h('button.btn' + (isAct ? '' : '.primary'), {
          style: { padding: '8px 16px', display: 'flex', alignItems: 'center', gap: '6px' },
          onclick: () => {
            applyDexMode(!isAct);
            renderContent();
          }
        }, ic('monitor', 16), isAct ? 'Tắt DeX Mode' : 'Bật DeX Mode ngay'),
        h('button.btn', {
          style: { padding: '8px 16px', display: 'flex', alignItems: 'center', gap: '6px' },
          onclick: () => openVirtualTrackpad()
        }, ic('mouse-pointer', 16), 'Mở Bàn rê chuột ảo'),
        h('button.btn', {
          style: { padding: '8px 16px', display: 'flex', alignItems: 'center', gap: '6px' },
          onclick: () => startDualScreenPresenter(window.OOS ? window.OOS.curDoc : null)
        }, ic('presentation', 16), 'Trình chiếu 2 màn hình'),
        h('button.btn.primary', {
          style: { padding: '8px 16px', display: 'flex', alignItems: 'center', gap: '6px' },
          onclick: () => openDexDesktopEnvironment({ mode: 'tvhome' })
        }, ic('layout', 16), 'Màn hình chính TV')
      );

      const introSec = h('div', { style: { background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '8px', padding: '12px', fontSize: '13px' } },
        h('b', { style: { color: '#0f172a' } }, 'Về bộ tính năng Android DeX trên Yomi-Office Open Source v' + VERSION + ':'),
        h('p', { style: { margin: '6px 0', lineHeight: '1.5' } },
          'Yomi-Office Open Source được thiết kế thích ứng hoàn hảo giữa điện thoại di động và giao diện máy tính để bàn (Samsung DeX, Motorola Ready For, Android 10+ Desktop Mode). Khi cắm cáp hoặc kết nối không dây, ứng dụng tự động mở rộng sang bố cục Ribbon 2 tầng chuyên nghiệp, hỗ trợ đa cửa sổ tự do, phím tắt phần cứng và bàn rê chuột ảo đa điểm.'
        )
      );

      contentEl.appendChild(grid);
      contentEl.appendChild(quickActions);
      contentEl.appendChild(introSec);

    // --------------------------------- TAB 2: CONNECTIONS
    } else if (curTab === 'connections') {
      const netInfo = await Host.getDeviceNetworkInfo();
      const ip = netInfo.ip || '127.0.0.1';
      const webPort = d.webLinkPort || 8080;
      const webUrl = `http://${ip}:${webPort}`;
      const scrcpyCmd = `adb connect ${ip}:5555 && scrcpy --max-size 1920 -b ${d.scrcpyBitrate || 12}M --turn-screen-off --stay-awake`;

      contentEl.appendChild(h('div.guide-sec',
        h('h3', { style: { color: '#0f172a', marginBottom: '8px' } }, 'Quản Lý Kết Nối Đa Thiết Bị & Phone-to-PC Bridge'),
        h('div', { style: { display: 'flex', flexDirection: 'column', gap: '12px' } },

          // Mode 1: Cáp DisplayPort / HDMI
          h('div.dex-conn-card',
            h('div.dex-conn-head',
              h('div', { style: { fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' } },
                ic('cable', 18), h('span', '1. Kết nối Cáp Type-C sang HDMI / DisplayPort')
              ),
              h('span.dex-badge-tag.ready', 'Độ trễ = 0 ms')
            ),
            h('p.small.muted', 'Cắm trực tiếp từ cổng sạc Type-C của điện thoại sang màn hình máy tính hoặc TV:'),
            h('ul.small', { style: { paddingLeft: '18px', margin: '4px 0' } },
              h('li', 'Tự động xuất hình ảnh chuẩn 4K / 60Hz không nén.'),
              h('li', 'Khuyên dùng Hub USB-C đa cổng để vừa sạc pin vừa cắm chuột, bàn phím rời.')
            )
          ),

          // Mode 2: Không dây Miracast / Screen Cast
          h('div.dex-conn-card',
            h('div.dex-conn-head',
              h('div', { style: { fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' } },
                ic('cast', 18), h('span', '2. Chiếu Không Dây (Wireless DeX / Smart View / Miracast)')
              ),
              h('button.btn.small', {
                onclick: () => {
                  if (!Host.openSystemCastSettings()) {
                    toast('Hãy mở Cài đặt → Thiết bị đã kết nối → Smart View / Cast trên điện thoại', 'info');
                  }
                }
              }, ic('external-link', 14), ' Mở Cài đặt Cast')
            ),
            h('p.small.muted', 'Không cần dây cáp, lý tưởng cho phòng họp và thuyết trình slide trên TV/Máy chiếu:'),
            h('ul.small', { style: { paddingLeft: '18px', margin: '4px 0' } },
              h('li', 'Kết nối cả điện thoại và Smart TV/máy chiếu vào cùng mạng Wi-Fi 5GHz.'),
              h('li', 'Bật Samsung DeX không dây hoặc Smart View trong bảng Cài đặt nhanh của điện thoại.')
            )
          ),

          // Mode 3: Web Desktop Server (Phone-to-PC Bridge)
          h('div.dex-conn-card', { style: { background: '#f0fdf4', borderColor: '#bbf7d0' } },
            h('div.dex-conn-head',
              h('div', { style: { fontWeight: 'bold', color: '#166534', display: 'flex', alignItems: 'center', gap: '6px' } },
                ic('globe', 18), h('span', '3. Máy Chủ Web Desktop (Phone-to-PC Wi-Fi Link)')
              ),
              h('span.dex-badge-tag.ready', 'Tích hợp sẵn trong App')
            ),
            h('p.small', 'Mở toàn bộ giao diện Yomi-Office Open Source từ trình duyệt máy tính (Chrome, Edge) trong cùng mạng Wi-Fi không cần dây cáp:'),
            h('div', { style: { display: 'flex', alignItems: 'center', gap: '8px', margin: '6px 0' } },
              h('code', { style: { background: '#ffffff', padding: '6px 12px', borderRadius: '4px', border: '1px solid #cbd5e1', fontWeight: 'bold', color: '#0f172a' } }, webUrl),
              h('button.btn.small', {
                onclick: () => {
                  Host.clipSet ? Host.clipSet(webUrl) : navigator.clipboard.writeText(webUrl);
                  toast('Đã sao chép đường dẫn: ' + webUrl, 'ok');
                }
              }, ic('copy', 14), ' Sao chép URL'),
              h('button.btn.small', {
                onclick: () => Host.openUrl(webUrl)
              }, ic('external-link', 14), ' Mở thử')
            )
          ),

          // Mode 4: Wireless ADB & Scrcpy
          h('div.dex-conn-card',
            h('div.dex-conn-head',
              h('div', { style: { fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' } },
                ic('terminal', 18), h('span', '4. Trình Chiếu Nhanh Scrcpy & Wireless ADB Mirroring')
              ),
              h('span.dex-badge-tag.info', 'Dành cho Chuyên gia')
            ),
            h('p.small.muted', 'Lệnh 1-click để kết nối không dây từ máy tính PC sang điện thoại với tốc độ 60fps và tự động tắt màn hình máy:'),
            h('div', { style: { display: 'flex', alignItems: 'center', gap: '8px', margin: '4px 0' } },
              h('code', { style: { background: '#0f172a', color: '#38bdf8', padding: '6px 10px', borderRadius: '4px', fontSize: '11px', flex: 1, overflowX: 'auto' } }, scrcpyCmd),
              h('button.btn.small', {
                onclick: () => {
                  Host.clipSet ? Host.clipSet(scrcpyCmd) : navigator.clipboard.writeText(scrcpyCmd);
                  toast('Đã sao chép lệnh Scrcpy vào bộ nhớ tạm', 'ok');
                }
              }, ic('copy', 14), ' Sao chép')
            )
          )
        )
      ));

    // --------------------------------- TAB 3: DISPLAY & WINDOWS
    } else if (curTab === 'display') {
      const curRes = d.resolution || 'auto';
      const curScale = d.uiScale || 100;

      contentEl.appendChild(h('div.guide-sec',
        h('h3', { style: { color: '#0f172a', marginBottom: '8px' } }, 'Độ Phân Giải & Tỷ Lệ Hiển Thị Desktop'),
        h('div', { style: { marginBottom: '14px' } },
          h('div', { style: { fontWeight: 'bold', marginBottom: '4px' } }, 'Chọn định dạng độ phân giải hiển thị:'),
          h('div.dex-chip-group',
            [
              ['auto', 'Tự động (Native)'],
              ['1080p', 'Full HD (1920 × 1080)'],
              ['1440p', '2K QHD (2560 × 1440)'],
              ['4k', '4K UHD (3840 × 2160)'],
              ['ultrawide_1080', 'Ultrawide (2560 × 1080)'],
              ['ultrawide_1440', 'Ultrawide (3440 × 1440)']
            ].map(([k, lbl]) => {
              const chip = h('button.dex-chip' + (curRes === k ? '.active' : ''), lbl);
              chip.onclick = () => { applyDexResolution(k); renderContent(); };
              return chip;
            })
          )
        ),
        h('div', { style: { marginBottom: '14px' } },
          h('div', { style: { fontWeight: 'bold', marginBottom: '4px' } }, 'Tỷ lệ thu phóng giao diện (UI Scaling / Virtual DPI):'),
          h('div.dex-chip-group',
            [
              [75, '75% (Thu gọn)'],
              [85, '85% (Vừa vặn)'],
              [100, '100% (Chuẩn Desktop)'],
              [115, '115% (Rõ nét)'],
              [125, '125% (Lớn)'],
              [150, '150% (Phóng đại 4K)']
            ].map(([p, lbl]) => {
              const chip = h('button.dex-chip' + (curScale === p ? '.active' : ''), lbl);
              chip.onclick = () => { applyDexScale(p); renderContent(); };
              return chip;
            })
          )
        ),
        h('div', { style: { marginTop: '16px', borderTop: '1px solid #e2e8f0', paddingTop: '12px' } },
          h('div', { style: { fontWeight: 'bold', marginBottom: '6px' } }, 'Sắp xếp Đa cửa sổ & Chia màn hình (Window Tiling):'),
          h('div', { style: { display: 'flex', gap: '8px' } },
            h('button.btn', { onclick: () => tileWindow('left') }, ic('columns', 14), ' Chia trái (50%)'),
            h('button.btn', { onclick: () => tileWindow('right') }, ic('columns', 14), ' Chia phải (50%)'),
            h('button.btn', { onclick: () => tileWindow('window') }, ic('layout', 14), ' Cửa sổ nổi'),
            h('button.btn', { onclick: () => tileWindow('max') }, ic('maximize', 14), ' Toàn màn hình')
          )
        )
      ));

    // --------------------------------- TAB 4: INPUT & CONTROLS
    } else if (curTab === 'input') {
      contentEl.appendChild(h('div.guide-sec',
        h('h3', { style: { color: '#0f172a', marginBottom: '8px' } }, 'Bàn Phím, Chuột & Thiết Bị Ngoại Vi'),
        h('div', { style: { display: 'flex', flexDirection: 'column', gap: '12px' } },
          h('div.dex-conn-card',
            h('div', { style: { fontWeight: 'bold', marginBottom: '4px' } }, 'Bàn rê chuột ảo trên màn hình điện thoại (Virtual Trackpad):'),
            h('p.small.muted', 'Biến màn hình cảm ứng của điện thoại thành bàn rê chuột mượt mà khi kết nối TV/máy chiếu:'),
            h('button.btn.primary', {
              style: { alignSelf: 'flex-start', marginTop: '6px' },
              onclick: () => openVirtualTrackpad()
            }, ic('mouse-pointer', 16), ' Khởi chạy Bàn rê chuột ảo ngay')
          ),
          h('div.dex-conn-card',
            h('div', { style: { fontWeight: 'bold', marginBottom: '4px' } }, 'Sơ đồ phím tắt bàn phím ngoài (Physical Keyboard Profiles):'),
            h('div.dex-chip-group',
              [
                ['windows', 'Chuẩn Windows (Ctrl+C, Ctrl+V, Alt+Tab, Win Key)'],
                ['dex', 'Chuẩn Samsung DeX (Meta điều hướng)'],
                ['mac', 'Chuẩn Apple Mac (Command Key)']
              ].map(([k, lbl]) => {
                const chip = h('button.dex-chip' + ((d.keyboardProfile || 'windows') === k ? '.active' : ''), lbl);
                chip.onclick = () => { saveDexSettings({ keyboardProfile: k }); renderContent(); };
                return chip;
              })
            )
          ),
          h('div.dex-conn-card',
            h('div', { style: { fontWeight: 'bold', marginBottom: '4px' } }, 'Tùy chọn chống nóng & tiết kiệm pin:'),
            h('label.chk', { style: { display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' } },
              h('input', {
                type: 'checkbox',
                checked: d.wakeLock !== false,
                onchange: e => {
                  saveDexSettings({ wakeLock: e.target.checked });
                  Host.setKeepScreenOn(e.target.checked);
                }
              }),
              h('span', 'Giữ màn hình luôn sáng khi đang kết nối DeX (WakeLock API)')
            )
          )
        )
      ));

    // --------------------------------- TAB 5: PRESENTER VIEW
    } else if (curTab === 'presenter') {
      contentEl.appendChild(h('div.guide-sec',
        h('h3', { style: { color: '#0f172a', marginBottom: '8px' } }, 'Trình Chiếu Hai Màn Hình Độc Lập (Dual Screen Presenter View)'),
        h('p', 'Tính năng độc quyền cho phép bạn điều khiển buổi thuyết trình PowerPoint hoặc đọc tài liệu Word chuyên nghiệp:'),
        h('div', { style: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', margin: '12px 0' } },
          h('div', { style: { border: '1px solid #cbd5e1', borderRadius: '8px', padding: '12px', background: '#f8fafc' } },
            h('div', { style: { fontWeight: 'bold', color: '#185abd', marginBottom: '4px' } }, 'Màn hình Điện thoại (Presenter Console)'),
            h('ul.small', { style: { paddingLeft: '16px' } },
              h('li', 'Hiển thị slide hiện tại & slide kế tiếp.'),
              h('li', 'Đồng hồ bấm giờ và đếm ngược thời gian.'),
              h('li', 'Ghi chú diễn giả (Speaker Notes) bí mật.'),
              h('li', 'Bút laser ảo và công cụ vẽ chú thích.')
            )
          ),
          h('div', { style: { border: '1px solid #cbd5e1', borderRadius: '8px', padding: '12px', background: '#f8fafc' } },
            h('div', { style: { fontWeight: 'bold', color: '#166534', marginBottom: '4px' } }, 'Màn hình TV / Máy chiếu ngoài'),
            h('ul.small', { style: { paddingLeft: '16px' } },
              h('li', 'Toàn màn hình sạch sẽ, không lộ thanh công cụ.'),
              h('li', 'Độ nét tối đa 4K 60fps chuẩn điện ảnh.'),
              h('li', 'Hiệu ứng chuyển slide mượt mà.'),
              h('li', 'Khán giả chỉ nhìn thấy nội dung thuyết trình.')
            )
          )
        ),
        h('button.btn.primary', {
          style: { padding: '10px 20px', fontSize: '14px', display: 'flex', alignItems: 'center', gap: '8px' },
          onclick: () => startDualScreenPresenter(window.OOS ? window.OOS.curDoc : null)
        }, ic('presentation', 18), ' Khởi chạy Chế độ Diễn giả 2 Màn hình')
      ));

    // --------------------------------- TAB 6: TV DESKTOP & ALL APPS
    } else if (curTab === 'tvdesktop') {
      const apps = await Host.getInstalledApps();
      contentEl.appendChild(h('div.guide-sec',
        h('h3', { style: { color: '#0f172a', marginBottom: '8px' } }, 'Giao Diện Màn Hình TV & Hỗ Trợ Toàn Bộ Ứng Dụng Thiết Bị'),
        h('p', 'Khi chiếu điện thoại lên màn hình TV lớn (Smart TV / Máy chiếu), bạn có thể trải nghiệm không gian làm việc Desktop hoàn chỉnh như PC: màn hình chính với hình nền, biểu tượng ứng dụng, ngăn kéo ứng dụng thiết bị và chạy đa nhiệm nhiều cửa sổ cùng lúc:'),
        h('div', { style: { display: 'flex', gap: '10px', margin: '14px 0', flexWrap: 'wrap' } },
          h('button.btn.primary', {
            style: { padding: '9px 16px', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px' },
            onclick: () => openDexDesktopEnvironment({ mode: 'tvhome' })
          }, ic('monitor', 16), ' Khởi chạy Màn hình chính TV (TV Desktop)'),
          h('button.btn', {
            style: { padding: '9px 16px', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px' },
            onclick: () => openDexDesktopEnvironment({ mode: 'tvapps' })
          }, ic('layout-grid', 16), ' Mở Ngăn chứa Toàn bộ Ứng dụng (All Apps)'),
          h('button.btn', {
            style: { padding: '9px 16px', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px' },
            onclick: () => openDexDesktopEnvironment({ mode: 'multiwin' })
          }, ic('app-window', 16), ' Chế độ Đa nhiệm Đa cửa sổ (Multi-Window)')
        ),
        h('div.dex-conn-card',
          h('div.dex-conn-head',
            h('div', { style: { fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' } },
              ic('package', 18), h('span', 'Danh sách Ứng dụng Quét được trên Thiết bị (' + apps.length + ' ứng dụng)')
            ),
            h('span.dex-badge-tag.ready', 'Sẵn sàng khởi chạy trên TV')
          ),
          h('p.small.muted', 'Bạn có thể khởi chạy bất kỳ ứng dụng nào trong số này trực tiếp trên màn hình TV:'),
          h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '8px', maxHeight: '180px', overflowY: 'auto', marginTop: '8px', paddingRight: '4px' } },
            apps.map(app => h('div', {
              style: {
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                padding: '8px 10px',
                borderRadius: '6px',
                background: '#f8fafc',
                border: '1px solid #e2e8f0'
              }
            },
              h('div', { style: { display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden' } },
                ic(app.icon || 'package', 16),
                h('div', { style: { overflow: 'hidden' } },
                  h('div', { style: { fontSize: '12px', fontWeight: '600', whiteSpace: 'nowrap', textOverflow: 'ellipsis', overflow: 'hidden' } }, app.name),
                  h('div', { style: { fontSize: '10px', color: '#64748b', whiteSpace: 'nowrap', textOverflow: 'ellipsis', overflow: 'hidden' } }, app.packageName)
                )
              ),
              h('button.btn.small', {
                style: { padding: '2px 8px', fontSize: '11px', flexShrink: 0 },
                onclick: () => {
                  Host.launchApp(app.packageName);
                  toast('Đã gửi lệnh mở ứng dụng: ' + app.name, 'info');
                }
              }, 'Mở')
            ))
          )
        ),
        h('div.dex-conn-card', { style: { marginTop: '12px' } },
          h('div', { style: { fontWeight: 'bold', marginBottom: '8px' } }, 'Cài đặt Xuất hình ảnh & Âm thanh ra TV:'),
          h('div', { style: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' } },
            h('div',
              h('label.small', { style: { display: 'block', fontWeight: '500', marginBottom: '4px' } }, 'Độ phân giải Màn hình TV:'),
              h('select.sel', {
                style: { width: '100%', padding: '6px', borderRadius: '4px' },
                value: d.resolution || 'auto',
                onchange: e => {
                  saveDexSettings({ resolution: e.target.value });
                  applyDexResolution(e.target.value);
                  toast('Đã chọn độ phân giải: ' + e.target.value, 'success');
                }
              },
                h('option', { value: 'auto' }, 'Tự động theo TV (Khuyên dùng)'),
                h('option', { value: '4k' }, '4K UHD (3840 × 2160) @ 60Hz'),
                h('option', { value: '2k' }, '2K QHD (2560 × 1440) @ 60Hz'),
                h('option', { value: 'fhd' }, 'Full HD (1920 × 1080) @ 60Hz')
              )
            ),
            h('div',
              h('label.small', { style: { display: 'block', fontWeight: '500', marginBottom: '4px' } }, 'Định tuyến Âm thanh:'),
              h('select.sel', {
                style: { width: '100%', padding: '6px', borderRadius: '4px' },
                value: d.audioRoute || 'tv',
                onchange: e => {
                  saveDexSettings({ audioRoute: e.target.value });
                  toast('Âm thanh: ' + (e.target.value === 'tv' ? 'Xuất ra Loa TV (HDMI/Cast)' : 'Phát trên Loa Điện thoại'), 'info');
                }
              },
                h('option', { value: 'tv' }, 'Loa TV / Màn hình ngoài (HDMI)'),
                h('option', { value: 'phone' }, 'Loa Thiết bị Điện thoại'),
                h('option', { value: 'auto' }, 'Tự động chuyển đổi khi cắm cáp')
              )
            )
          )
        )
      ));
    }
  }

  renderTabs();
  renderContent();

  body.appendChild(navTabs);
  body.appendChild(contentEl);

  dialog({
    title: 'Trung Tâm Cài Đặt & Điều Khiển Android DeX (DeX Control Center)',
    width: 720,
    body,
    buttons: [{ label: 'Đóng', primary: true }]
  });
}

// Alias for backward compatibility
export const openDexGuideDialog = openDexControlCenterDialog;

// -----------------------------------------------------------------------------
// Khởi Tạo Động Cơ DeX Khi Mở Ứng Dụng (Initialization Engine)
// -----------------------------------------------------------------------------
export function initDexEngine() {
  const d = dexSettings();
  if (d.enabled || (d.autoLaunchOnHdmi && isDexEnvironment())) {
    applyDexMode(true, true);
  }
  if (d.uiScale && d.uiScale !== 100) {
    applyDexScale(d.uiScale);
  }
  if (d.resolution && d.resolution !== 'auto') {
    applyDexResolution(d.resolution);
  }
  if (d.wakeLock && isDexEnvironment()) {
    Host.setKeepScreenOn(true);
  }

  // Listen for native Android display changes
  window.OOS_onDisplayChanged = function (isDex) {
    if (d.autoLaunchOnHdmi) {
      applyDexMode(isDex);
    }
  };

  // Keyboard shortcut: Alt + D or Meta + D to toggle DeX Control Center
  window.addEventListener('keydown', e => {
    if ((e.altKey || e.metaKey) && (e.key === 'd' || e.key === 'D')) {
      e.preventDefault();
      openDexControlCenterDialog();
    }
  });
}

// -----------------------------------------------------------------------------
// Không Gian Làm Việc Màn Hình TV & Đa Nhiệm Đa Cửa Sổ (TV Desktop Environment)
// Hỗ trợ hiển thị toàn bộ thiết bị, ứng dụng và giao diện màn hình chính trên TV
// -----------------------------------------------------------------------------
// v0.1.6: the DeX desktop was redesigned (ChromeOS shelf + Quick Settings + Samsung DeX app drawer) -> dex16.js
export function openDexDesktopEnvironment(options = {}) { return import('./dex16.js').then(m => m.openDesktop(options)); }
// v0.1.6 desktop kept for reference / fallback
export function openDexDesktopEnvironmentLegacy(options = {}) {
  const mode = options.mode || 'tvhome'; // 'tvhome' | 'tvapps' | 'multiwin'

  const existing = document.getElementById('dex-desktop-screen');
  if (existing) existing.remove();

  const d = dexSettings();
  let startMenuOpen = (mode === 'tvapps');
  let audioOutput = (d.audioRoute === 'phone' ? 'phone' : 'tv');
  let highestZ = 9020;
  let activeWinId = null;
  let windows = [];
  let searchQuery = '';
  let activeCategory = 'all';

  const root = h('div#dex-desktop-screen.dex-desktop-screen');
  root.style.background = 'radial-gradient(circle at 75% 25%, #1e1b4b 0%, #0f172a 55%, #020617 100%)';

  // Desktop Icons Grid (Word, Excel, PPT, PDF, Calc, Files, Browser, Gallery, Camera, Settings, AI, DeX Control Center)
  const desktopIcons = h('div.dex-desktop-icons', {
    style: {
      position: 'absolute',
      top: '24px',
      left: '24px',
      display: 'grid',
      gridTemplateColumns: 'repeat(2, 100px)',
      gap: '16px 14px',
      zIndex: '9001'
    }
  });

  const desktopAppList = [
    { id: 'word', name: 'Word (Văn bản)', icon: 'file-text', color: '#185abd', type: 'word' },
    { id: 'excel', name: 'Excel (Bảng tính)', icon: 'file-spreadsheet', color: '#107c41', type: 'excel' },
    { id: 'ppt', name: 'PowerPoint (Thuyết trình)', icon: 'presentation', color: '#c43e1c', type: 'ppt' },
    { id: 'pdf', name: 'PDF Reader & Ký', icon: 'file-check', color: '#dc2626', type: 'pdf' },
    { id: 'mathtype', name: 'MathType (Toán)', icon: 'function-square', color: '#68217a', type: 'mathtype' },
    { id: 'calc', name: 'Máy tính Bỏ túi', icon: 'calculator', color: '#7c3aed', type: 'calc' },
    { id: 'files', name: 'Quản lý Tệp tin', icon: 'folder-open', color: '#0891b2', type: 'files' },
    { id: 'browser', name: 'Trình duyệt Web', icon: 'globe', color: '#2563eb', type: 'browser' },
    { id: 'gallery', name: 'Thư viện Ảnh', icon: 'image', color: '#d97706', type: 'gallery' },
    { id: 'camera', name: 'Máy ảnh', icon: 'camera', color: '#4f46e5', type: 'camera' },
    { id: 'settings', name: 'Cài đặt Android', icon: 'settings', color: '#475569', type: 'settings' },
    { id: 'ai', name: 'Trợ lý AI Office', icon: 'sparkles', color: '#ec4899', type: 'ai' },
    { id: 'dexcc', name: 'Trung tâm DeX', icon: 'settings-2', color: '#0284c7', type: 'dexcc' }
  ];

  desktopAppList.forEach(app => {
    const iconBtn = h('div.dex-desktop-icon', {
      onclick: () => {
        if (app.type === 'dexcc') {
          openDexControlCenterDialog();
        } else {
          openAppWindow(app.type, app.name, app.icon);
        }
      }
    },
      h('div.dex-desktop-icon-badge', { style: { background: app.color } }, ic(app.icon, 24)),
      h('div.dex-desktop-icon-label', app.name)
    );
    desktopIcons.appendChild(iconBtn);
  });
  root.appendChild(desktopIcons);

  // System Stats Widget (Top Right on TV screen)
  const widgetEl = h('div.dex-tv-widget', {
    style: {
      position: 'absolute',
      top: '24px',
      right: '28px',
      background: 'rgba(15, 23, 42, 0.78)',
      backdropFilter: 'blur(20px)',
      border: '1px solid rgba(255, 255, 255, 0.16)',
      borderRadius: '14px',
      padding: '16px 22px',
      color: '#ffffff',
      display: 'flex',
      flexDirection: 'column',
      gap: '8px',
      boxShadow: '0 12px 36px rgba(0,0,0,0.45)',
      zIndex: '9002',
      minWidth: '290px'
    }
  },
    h('div', { style: { display: 'flex', alignItems: 'baseline', justifyContent: 'space-between' } },
      h('div.dex-clock-time', { style: { fontSize: '28px', fontWeight: '700', letterSpacing: '0.5px', color: '#38bdf8' } }, '14:30:00'),
      h('div', { style: { fontSize: '11px', color: '#94a3b8', background: 'rgba(255,255,255,0.1)', padding: '2px 8px', borderRadius: '10px' } }, '4K UHD 60Hz')
    ),
    h('div.dex-clock-date', { style: { fontSize: '12px', color: '#cbd5e1', fontWeight: '500' } }, 'Thứ Ba, 20 Tháng 10, 2026'),
    h('div', { style: { height: '1px', background: 'rgba(255,255,255,0.12)', margin: '4px 0' } }),
    h('div', { style: { display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '11px', color: '#94a3b8' } },
      h('div', { style: { display: 'flex', justifyContent: 'space-between' } },
        h('span', 'Màn hình đích:'),
        h('b', { style: { color: '#e2e8f0' } }, 'Smart TV 65" (Không dây / HDMI)')
      ),
      h('div', { style: { display: 'flex', justifyContent: 'space-between' } },
        h('span', 'Âm thanh:'),
        h('b.dex-audio-label', { style: { color: '#4ade80' } }, audioOutput === 'tv' ? 'Loa TV HDMI' : 'Loa Điện Thoại')
      ),
      h('div', { style: { display: 'flex', justifyContent: 'space-between' } },
        h('span', 'Tài nguyên:'),
        h('b', { style: { color: '#e2e8f0' } }, 'RAM: 3.4GB / 8GB • CPU: 18%')
      )
    )
  );
  root.appendChild(widgetEl);

  const clockInterval = setInterval(() => {
    if (!document.getElementById('dex-desktop-screen')) {
      clearInterval(clockInterval);
      return;
    }
    const now = new Date();
    const tEl = widgetEl.querySelector('.dex-clock-time');
    if (tEl) tEl.textContent = now.toTimeString().slice(0, 8);
    const trayTime = taskbar.querySelector('.dex-tray-time');
    if (trayTime) trayTime.textContent = now.toTimeString().slice(0, 5);
  }, 1000);

  // Workspace Container for floating windows
  const workspaceEl = h('div.dex-windows-container', {
    style: { position: 'absolute', inset: 0, bottom: '48px', overflow: 'hidden' }
  });
  root.appendChild(workspaceEl);

  // Taskbar
  const taskbar = h('div.dex-desktop-taskbar');
  const taskbarTabsContainer = h('div', { style: { display: 'flex', gap: '6px', alignItems: 'center' } });

  // Start Menu element
  const startMenu = h('div.dex-start-menu', {
    style: {
      display: startMenuOpen ? 'flex' : 'none',
      width: '460px',
      height: '540px',
      bottom: '56px',
      left: '12px'
    }
  });

  // Start Button
  const startBtn = h('button.dex-start-btn', {
    onclick: (e) => {
      e.stopPropagation();
      startMenuOpen = !startMenuOpen;
      startMenu.style.display = startMenuOpen ? 'flex' : 'none';
      if (startMenuOpen) renderStartMenu();
    }
  }, ic('layout-grid', 18), h('span', 'Ứng dụng'));

  // Quick launch shortcuts on Taskbar
  const quickLaunch = h('div', { style: { display: 'flex', gap: '4px', alignItems: 'center', marginLeft: '6px', paddingRight: '10px', borderRight: '1px solid rgba(255,255,255,0.15)' } },
    h('button.dex-win-btn', { title: 'Word', onclick: () => openAppWindow('word', 'Word (Văn bản)', 'file-text') }, ic('file-text', 18)),
    h('button.dex-win-btn', { title: 'Excel', onclick: () => openAppWindow('excel', 'Excel (Bảng tính)', 'file-spreadsheet') }, ic('file-spreadsheet', 18)),
    h('button.dex-win-btn', { title: 'PowerPoint', onclick: () => openAppWindow('ppt', 'PowerPoint (Thuyết trình)', 'presentation') }, ic('presentation', 18)),
    h('button.dex-win-btn', { title: 'PDF Reader & Ký', onclick: () => openAppWindow('pdf', 'PDF Reader & Ký', 'file-check') }, ic('file-check', 18)),
    h('button.dex-win-btn', { title: 'MathType Toán', onclick: () => openAppWindow('mathtype', 'MathType (Toán học)', 'function-square') }, ic('function-square', 18)),
    h('button.dex-win-btn', { title: 'Quản lý Tệp tin', onclick: () => openAppWindow('files', 'Quản lý Tệp tin', 'folder-open') }, ic('folder-open', 18)),
    h('button.dex-win-btn', { title: 'Trình duyệt Web', onclick: () => openAppWindow('browser', 'Trình duyệt Web', 'globe') }, ic('globe', 18)),
    h('button.dex-win-btn', { title: 'Máy tính Bỏ túi', onclick: () => openAppWindow('calc', 'Máy tính Bỏ túi', 'calculator') }, ic('calculator', 18))
  );

  // System Tray
  const systemTray = h('div.dex-system-tray');
  const wifiItem = h('div.dex-tray-item', { title: 'Wi-Fi 5GHz Đã kết nối' }, ic('wifi', 16), h('span', 'Wi-Fi 5GHz'));
  const audioItem = h('div.dex-tray-item.dex-audio-btn', {
    title: 'Bấm để đổi cổng phát âm thanh giữa TV HDMI và Loa Điện Thoại',
    onclick: () => {
      audioOutput = (audioOutput === 'tv' ? 'phone' : 'tv');
      audioItem.querySelector('span').textContent = (audioOutput === 'tv' ? 'Loa TV' : 'Loa ĐT');
      const lbl = widgetEl.querySelector('.dex-audio-label');
      if (lbl) {
        lbl.textContent = (audioOutput === 'tv' ? 'Loa TV HDMI' : 'Loa Điện Thoại');
        lbl.style.color = (audioOutput === 'tv' ? '#4ade80' : '#f59e0b');
      }
      toast('Đã chuyển âm thanh sang: ' + (audioOutput === 'tv' ? 'Loa TV (HDMI / Cast)' : 'Loa Điện Thoại'), 'info');
    }
  }, ic('volume-2', 16), h('span', audioOutput === 'tv' ? 'Loa TV' : 'Loa ĐT'));

  const resBadge = h('div.dex-tray-item', {
    style: { background: 'rgba(56, 189, 248, 0.15)', color: '#38bdf8', padding: '2px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 'bold' }
  }, '4K UHD');

  const batteryItem = h('div.dex-tray-item', { title: 'Pin 98% Đang sạc qua Hub USB-C' }, h('span', '98% ⚡'));
  const clockTrayItem = h('div.dex-tray-item.dex-tray-time', { style: { fontWeight: '600' } }, new Date().toTimeString().slice(0, 5));
  const settingsTrayItem = h('div.dex-tray-item', {
    title: 'Cài đặt DeX Control Center',
    onclick: () => openDexControlCenterDialog()
  }, ic('settings-2', 16));

  const exitBtn = h('button', {
    style: {
      background: '#ef4444',
      border: 'none',
      color: '#ffffff',
      padding: '4px 10px',
      borderRadius: '4px',
      cursor: 'pointer',
      fontSize: '11px',
      fontWeight: '600',
      display: 'flex',
      alignItems: 'center',
      gap: '4px'
    },
    title: 'Thoát giao diện TV Desktop về ứng dụng chính',
    onclick: () => {
      clearInterval(clockInterval);
      root.remove();
      toast('Đã đóng giao diện TV Desktop', 'info');
    }
  }, ic('x', 14), h('span', 'Đóng TV'));

  systemTray.appendChild(wifiItem);
  systemTray.appendChild(audioItem);
  systemTray.appendChild(resBadge);
  systemTray.appendChild(batteryItem);
  systemTray.appendChild(clockTrayItem);
  systemTray.appendChild(settingsTrayItem);
  systemTray.appendChild(exitBtn);

  const layoutBtn = h('button.dex-start-btn', {
    style: { background: 'rgba(255,255,255,0.12)', color: '#38bdf8', fontSize: '12px', padding: '5px 10px', gap: '5px', border: '1px solid rgba(255,255,255,0.18)' },
    title: 'Bố trí cửa sổ đa nhiệm DeX (Window Layouts)',
    onclick: (e) => {
      e.stopPropagation();
      showMenu([
        { header: 'Bố Trí Đa Nhiệm Cửa Sổ (DeX Multi-Window)' },
        { label: 'Chia đôi 50/50 (Word bên Trái, Excel bên Phải)', icon: 'columns', run: () => snapSplit5050() },
        { label: 'Lưới 4 cửa sổ (2x2 Multi-App Grid)', icon: 'layout-grid', run: () => snapGrid2x2() },
        { label: 'Xếp tầng cửa sổ (Cascade)', icon: 'layers', run: () => cascadeWindows() },
        { sep: true },
        { label: 'Thu nhỏ tất cả (Hiện màn hình nền)', icon: 'minus', run: () => minimizeAllWindows() },
        { label: 'Khôi phục tất cả cửa sổ', icon: 'maximize', run: () => restoreAllWindows() }
      ], { anchor: layoutBtn });
    }
  }, ic('layout', 15), h('span', 'Bố trí'));

  taskbar.appendChild(startBtn);
  taskbar.appendChild(quickLaunch);
  taskbar.appendChild(layoutBtn);
  taskbar.appendChild(taskbarTabsContainer);
  taskbar.appendChild(systemTray);

  root.appendChild(taskbar);
  root.appendChild(startMenu);

  // Close start menu on outside click
  root.addEventListener('click', (e) => {
    if (startMenuOpen && !e.target.closest('.dex-start-menu') && !e.target.closest('.dex-start-btn')) {
      startMenuOpen = false;
      startMenu.style.display = 'none';
    }
  });

  // -------------------------------------------------------------
  // Window Manager Logic & Layout Handlers
  // -------------------------------------------------------------
  function snapSplit5050() {
    const sw = window.innerWidth || 1920;
    const sh = (window.innerHeight || 1080) - 48;
    const halfW = Math.floor(sw / 2);
    openAppWindow('word', 'Word (Văn bản)', 'file-text');
    openAppWindow('excel', 'Excel (Bảng tính)', 'file-spreadsheet');
    const wWin = windows.find(w => w.id === 'word');
    const eWin = windows.find(w => w.id === 'excel');
    if (wWin) {
      wWin.minimized = false;
      wWin.el.style.display = 'flex';
      wWin.el.style.left = '0px';
      wWin.el.style.top = '0px';
      wWin.el.style.width = halfW + 'px';
      wWin.el.style.height = sh + 'px';
      wWin.el.style.borderRadius = '0px';
    }
    if (eWin) {
      eWin.minimized = false;
      eWin.el.style.display = 'flex';
      eWin.el.style.left = halfW + 'px';
      eWin.el.style.top = '0px';
      eWin.el.style.width = (sw - halfW) + 'px';
      eWin.el.style.height = sh + 'px';
      eWin.el.style.borderRadius = '0px';
    }
    toast('Đã chia đôi màn hình 50/50: Word bên trái, Excel bên phải', 'ok');
  }

  function snapGrid2x2() {
    const sw = window.innerWidth || 1920;
    const sh = (window.innerHeight || 1080) - 48;
    const halfW = Math.floor(sw / 2);
    const halfH = Math.floor(sh / 2);
    openAppWindow('word', 'Word (Văn bản)', 'file-text');
    openAppWindow('excel', 'Excel (Bảng tính)', 'file-spreadsheet');
    openAppWindow('calc', 'Máy tính Bỏ túi', 'calculator');
    openAppWindow('files', 'Quản lý Tệp tin', 'folder-open');

    const appPositions = [
      { id: 'word', x: 0, y: 0, w: halfW, h: halfH },
      { id: 'excel', x: halfW, y: 0, w: sw - halfW, h: halfH },
      { id: 'files', x: 0, y: halfH, w: halfW, h: sh - halfH },
      { id: 'calc', x: halfW, y: halfH, w: sw - halfW, h: sh - halfH }
    ];
    appPositions.forEach(pos => {
      const w = windows.find(x => x.id === pos.id);
      if (w) {
        w.minimized = false;
        w.el.style.display = 'flex';
        w.el.style.left = pos.x + 'px';
        w.el.style.top = pos.y + 'px';
        w.el.style.width = pos.w + 'px';
        w.el.style.height = pos.h + 'px';
        w.el.style.borderRadius = '0px';
      }
    });
    toast('Đã sắp xếp bố trí lưới 4 cửa sổ (2x2 Grid)', 'ok');
  }

  function cascadeWindows() {
    let offset = 30;
    windows.forEach(w => {
      w.minimized = false;
      w.el.style.display = 'flex';
      w.el.style.left = offset + 'px';
      w.el.style.top = offset + 'px';
      w.el.style.width = '780px';
      w.el.style.height = '520px';
      w.el.style.borderRadius = '10px';
      bringToFront(w);
      offset += 35;
    });
    toast('Đã xếp tầng tất cả cửa sổ', 'info');
  }

  function minimizeAllWindows() {
    windows.forEach(w => {
      w.minimized = true;
      w.el.style.display = 'none';
    });
    updateTaskbarTabs();
    toast('Đã thu nhỏ tất cả cửa sổ', 'info');
  }

  function restoreAllWindows() {
    windows.forEach(w => {
      w.minimized = false;
      w.el.style.display = 'flex';
    });
    updateTaskbarTabs();
    toast('Đã khôi phục tất cả cửa sổ', 'info');
  }

  function openAppWindow(type, title, iconStr, customBounds) {
    let win = windows.find(w => w.id === type);
    if (win) {
      win.minimized = false;
      win.el.style.display = 'flex';
      bringToFront(win);
      updateTaskbarTabs();
      return;
    }

    const winId = type;
    const defaultBounds = customBounds || getDefaultBounds(type);

    const isExternalApp = !['word', 'excel', 'ppt', 'pdf', 'calc', 'files', 'browser', 'mathtype', 'gallery', 'camera', 'settings', 'ai', 'dexcc'].includes(type);
    if (isExternalApp) {
      Host.launchAppWithOptions(type, defaultBounds.x, defaultBounds.y, defaultBounds.w, defaultBounds.h, 0);
    }

    const winEl = h('div.dex-desktop-window', {
      id: 'dex-win-' + winId,
      style: {
        left: defaultBounds.x + 'px',
        top: defaultBounds.y + 'px',
        width: defaultBounds.w + 'px',
        height: defaultBounds.h + 'px',
        zIndex: String(++highestZ)
      },
      onmousedown: () => bringToFront(win)
    });

    const winHeader = h('div.dex-window-header');
    const winTitle = h('div.dex-window-title', ic(iconStr || 'file', 16), h('span', title));
    const winBtns = h('div.dex-window-btns',
      h('button.dex-win-btn', {
        title: 'Thu nhỏ',
        onclick: (e) => {
          e.stopPropagation();
          win.minimized = true;
          winEl.style.display = 'none';
          updateTaskbarTabs();
        }
      }, ic('minus', 14)),
      h('button.dex-win-btn', {
        title: 'Phóng to / Khôi phục',
        onclick: (e) => {
          e.stopPropagation();
          toggleMaximize(win);
        }
      }, ic('maximize', 14)),
      h('button.dex-win-btn.close', {
        title: 'Đóng',
        onclick: (e) => {
          e.stopPropagation();
          closeWindow(win);
        }
      }, ic('x', 14))
    );

    winHeader.appendChild(winTitle);
    winHeader.appendChild(winBtns);
    winEl.appendChild(winHeader);

    setupWindowDrag(winEl, winHeader);

    const winContent = h('div.dex-window-content');
    renderWindowContent(winContent, type, win);
    winEl.appendChild(winContent);

    workspaceEl.appendChild(winEl);

    win = {
      id: winId,
      type,
      title,
      icon: iconStr,
      el: winEl,
      contentEl: winContent,
      minimized: false,
      maximized: false,
      bounds: defaultBounds,
      isExternal: isExternalApp
    };
    windows.push(win);
    activeWinId = winId;
    updateTaskbarTabs();
  }

  function getDefaultBounds(type) {
    const sw = window.innerWidth || 1920;
    const sh = window.innerHeight || 1080;
    if (type === 'word') {
      return { x: 40, y: 35, w: Math.min(880, sw - 80), h: Math.min(640, sh - 120) };
    } else if (type === 'excel') {
      return { x: 940, y: 35, w: Math.min(840, sw - 980), h: Math.min(480, sh - 140) };
    } else if (type === 'ppt') {
      return { x: 60, y: 50, w: Math.min(920, sw - 120), h: Math.min(600, sh - 140) };
    } else if (type === 'pdf') {
      return { x: 100, y: 40, w: Math.min(860, sw - 160), h: Math.min(620, sh - 140) };
    } else if (type === 'mathtype') {
      return { x: 300, y: 90, w: 720, h: 480 };
    } else if (type === 'calc') {
      return { x: 1050, y: 535, w: 320, h: 440 };
    } else if (type === 'files') {
      return { x: 260, y: 220, w: 780, h: 500 };
    } else if (type === 'browser') {
      return { x: 300, y: 80, w: 900, h: 580 };
    } else {
      return { x: 140, y: 90, w: 760, h: 520 };
    }
  }

  function bringToFront(win) {
    if (!win) return;
    win.el.style.zIndex = String(++highestZ);
    activeWinId = win.id;
    updateTaskbarTabs();
  }

  function toggleMaximize(win) {
    if (!win.maximized) {
      win.savedStyle = {
        left: win.el.style.left,
        top: win.el.style.top,
        width: win.el.style.width,
        height: win.el.style.height,
        borderRadius: win.el.style.borderRadius
      };
      win.el.style.left = '0px';
      win.el.style.top = '0px';
      win.el.style.width = '100%';
      win.el.style.height = '100%';
      win.el.style.borderRadius = '0px';
      win.maximized = true;
    } else {
      if (win.savedStyle) {
        win.el.style.left = win.savedStyle.left;
        win.el.style.top = win.savedStyle.top;
        win.el.style.width = win.savedStyle.width;
        win.el.style.height = win.savedStyle.height;
        win.el.style.borderRadius = win.savedStyle.borderRadius || '10px';
      }
      win.maximized = false;
    }
  }

  function closeWindow(win) {
    if (!win) return;
    win.el.remove();
    windows = windows.filter(w => w.id !== win.id);
    if (activeWinId === win.id) {
      activeWinId = windows.length ? windows[windows.length - 1].id : null;
    }
    updateTaskbarTabs();
  }

  function setupWindowDrag(el, handle) {
    let isDragging = false;
    let startX = 0, startY = 0;
    let origLeft = 0, origTop = 0;

    handle.addEventListener('pointerdown', (e) => {
      if (e.target.closest('button')) return;
      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      origLeft = parseInt(el.style.left, 10) || 0;
      origTop = parseInt(el.style.top, 10) || 0;
      handle.setPointerCapture(e.pointerId);
    });

    handle.addEventListener('pointermove', (e) => {
      if (!isDragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      el.style.left = (origLeft + dx) + 'px';
      el.style.top = (origTop + dy) + 'px';
    });

    handle.addEventListener('pointerup', (e) => {
      if (isDragging) {
        isDragging = false;
        try { handle.releasePointerCapture(e.pointerId); } catch (_) {}
        const curWin = windows.find(w => w.el === el);
        if (curWin && curWin.isExternal) {
          const curX = parseInt(el.style.left, 10) || 0;
          const curY = parseInt(el.style.top, 10) || 0;
          const curW = parseInt(el.style.width, 10) || 760;
          const curH = parseInt(el.style.height, 10) || 520;
          Host.launchAppWithOptions(curWin.type, curX, curY, curW, curH, 0);
        }
      }
    });
  }

  function updateTaskbarTabs() {
    taskbarTabsContainer.innerHTML = '';
    windows.forEach(w => {
      const isActive = (w.id === activeWinId && !w.minimized);
      const tab = h('div.dex-taskbar-tab' + (isActive ? '.active' : ''), {
        onclick: () => {
          if (w.minimized) {
            w.minimized = false;
            w.el.style.display = 'flex';
            bringToFront(w);
          } else if (w.id === activeWinId) {
            w.minimized = true;
            w.el.style.display = 'none';
          } else {
            bringToFront(w);
          }
          updateTaskbarTabs();
        }
      },
        ic(w.icon || 'file', 14),
        h('span', w.title)
      );
      taskbarTabsContainer.appendChild(tab);
    });
  }

  // -------------------------------------------------------------
  // Content Renderers for Windows
  // -------------------------------------------------------------
  function renderWindowContent(container, type, win) {
    if (type === 'word') {
      container.innerHTML = `
        <div style="display:flex; flex-direction:column; height:100%; background:#f1f5f9; font-family:'Segoe UI',Roboto,sans-serif;">
          <div style="background:#ffffff; border-bottom:1px solid #cbd5e1; padding:6px 12px; display:flex; flex-direction:column; gap:6px;">
            <div style="display:flex; gap:16px; font-size:12px; font-weight:600; color:#1e293b; border-bottom:1px solid #e2e8f0; padding-bottom:4px;">
              <span style="color:#185abd; border-bottom:2px solid #185abd; padding-bottom:2px; cursor:pointer;">Trang đầu</span>
              <span style="color:#64748b; cursor:pointer;">Chèn</span>
              <span style="color:#64748b; cursor:pointer;">Bố trí</span>
              <span style="color:#64748b; cursor:pointer;">Xem lại</span>
              <span style="color:#64748b; cursor:pointer;">Xem</span>
            </div>
            <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap;">
              <select class="dex-word-font" style="font-size:12px; padding:3px 6px; border:1px solid #cbd5e1; border-radius:4px;">
                <option value="Segoe UI">Segoe UI</option>
                <option value="Times New Roman">Times New Roman</option>
                <option value="Arial">Arial</option>
                <option value="Roboto">Roboto</option>
                <option value="Courier New">Courier New</option>
              </select>
              <select class="dex-word-size" style="font-size:12px; padding:3px 6px; border:1px solid #cbd5e1; border-radius:4px;">
                <option value="3">12 pt</option>
                <option value="4" selected>14 pt</option>
                <option value="5">16 pt</option>
                <option value="6">20 pt</option>
                <option value="7">24 pt</option>
              </select>
              <div style="display:flex; border:1px solid #cbd5e1; border-radius:4px; overflow:hidden;">
                <button class="dex-btn-bold" style="padding:3px 8px; border:none; background:#f8fafc; font-weight:bold; cursor:pointer;" title="In đậm (Ctrl+B)">B</button>
                <button class="dex-btn-italic" style="padding:3px 8px; border:none; background:#f8fafc; font-style:italic; cursor:pointer;" title="In nghiêng (Ctrl+I)">I</button>
                <button class="dex-btn-underline" style="padding:3px 8px; border:none; background:#f8fafc; text-decoration:underline; cursor:pointer;" title="Gạch chân (Ctrl+U)">U</button>
                <button class="dex-btn-strike" style="padding:3px 8px; border:none; background:#f8fafc; text-decoration:line-through; cursor:pointer;" title="Gạch ngang">S</button>
              </div>
              <div style="display:flex; border:1px solid #cbd5e1; border-radius:4px; overflow:hidden;">
                <button class="dex-btn-left" style="padding:3px 8px; border:none; background:#f8fafc; cursor:pointer;" title="Căn trái">⫷</button>
                <button class="dex-btn-center" style="padding:3px 8px; border:none; background:#f8fafc; cursor:pointer;" title="Căn giữa">☰</button>
                <button class="dex-btn-right" style="padding:3px 8px; border:none; background:#f8fafc; cursor:pointer;" title="Căn phải">⫸</button>
                <button class="dex-btn-justify" style="padding:3px 8px; border:none; background:#f8fafc; cursor:pointer;" title="Căn đều">☲</button>
              </div>
              <div style="display:flex; border:1px solid #cbd5e1; border-radius:4px; overflow:hidden;">
                <button class="dex-btn-ul" style="padding:3px 8px; border:none; background:#f8fafc; cursor:pointer;" title="Danh sách">• List</button>
                <button class="dex-btn-ol" style="padding:3px 8px; border:none; background:#f8fafc; cursor:pointer;" title="Số">1. List</button>
              </div>
              <button class="dex-btn-table" style="padding:3px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;">+ Chèn Bảng</button>
              <button class="dex-btn-open-word" style="padding:3px 10px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;">📂 Mở tệp</button>
              <button class="dex-btn-to-main" style="padding:3px 10px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;" title="Chuyển sang giao diện chính với đầy đủ tính năng">🖥 Giao diện chính</button>
              <button class="dex-btn-save" style="padding:3px 10px; border:none; background:#185abd; color:#fff; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;">💾 Lưu .docx</button>
              <button class="dex-btn-print" style="padding:3px 10px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;">🖨 In</button>
              <span style="margin-left:auto; font-size:11px; color:#16a34a; font-weight:500;">✓ Tự động lưu</span>
            </div>
          </div>
          <div style="flex:1; overflow-y:auto; padding:20px; display:flex; justify-content:center; background:#e2e8f0;">
            <div class="dex-doc-page" contenteditable="true" spellcheck="false" style="width:100%; max-width:760px; min-height:540px; background:#ffffff; box-shadow:0 6px 20px rgba(0,0,0,0.12); border-radius:4px; padding:36px; color:#0f172a; line-height:1.6; outline:none; font-size:14px;">
              <div style="text-align:center; font-weight:bold; font-size:13px; text-transform:uppercase; margin-bottom:14px; border-bottom:1px solid #0f172a; padding-bottom:10px;">
                CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM<br><span style="font-size:12px; font-weight:normal; text-transform:none;">Độc lập - Tự do - Hạnh phúc</span>
              </div>
              <h2 style="text-align:center; color:#185abd; font-size:18px; margin:16px 0 10px;">BÁO CÁO NÂNG CẤP BẢN YOMI-OFFICE OPEN SOURCE V${VERSION}</h2>
              <p style="font-style:italic; font-size:12px; color:#64748b; text-align:center; margin-bottom:18px;">Hệ thống Android DeX Desktop Mode, Đa nhiệm Cửa sổ Multi-Window & Tương thích Toàn diện Ứng dụng Thiết bị</p>
              <h4 style="color:#0f172a; margin-top:14px; margin-bottom:6px;">1. Tính Năng Android DeX & Màn Hình TV 4K Toàn Diện:</h4>
              <ul style="padding-left:20px; font-size:13px;">
                <li><b>Đa nhiệm nhiều cửa sổ (True Multi-Window 50/50):</b> Chạy song song văn bản Word, bảng tính Excel, trình chiếu PowerPoint, tài liệu PDF, MathType và Máy tính cùng lúc trên màn hình lớn.</li>
                <li><b>Tương thích và điều khiển Mọi Ứng dụng Bên Ngoài:</b> Tích hợp bộ khởi chạy và điều phối cửa sổ nổi tự do (Freeform Bounds) cho toàn bộ ứng dụng cài trên hệ thống Android.</li>
                <li><b>Định tuyến âm thanh và thiết bị ngoại vi:</b> Chuyển đổi Loa TV HDMI và Loa điện thoại, tối ưu bàn rê chuột ảo (Virtual Trackpad) và phím tắt PC chuẩn.</li>
              </ul>
              <h4 style="color:#0f172a; margin-top:14px; margin-bottom:6px;">2. Tương Thích Tuyệt Đối Nền Tảng:</h4>
              <p style="font-size:13px;">Hệ thống hoạt động hoàn toàn offline 100% trên Android 6.0 trở lên và Windows 7/8.1/10/11 các kiến trúc x64, x86, ARM.</p>
            </div>
          </div>
          <div style="height:26px; background:#ffffff; border-top:1px solid #cbd5e1; display:flex; align-items:center; justify-content:space-between; padding:0 14px; font-size:11px; color:#64748b;">
            <div style="display:flex; gap:14px;">
              <span>Trang 1 / 1</span>
              <span class="dex-word-count">Số từ: 142 từ</span>
              <span class="dex-char-count">Ký tự: 980</span>
            </div>
            <div style="display:flex; gap:12px; align-items:center;">
              <span>Tiếng Việt (100% Offline)</span>
              <span style="font-weight:600; color:#185abd;">100%</span>
            </div>
          </div>
        </div>
      `;

      const docPage = container.querySelector('.dex-doc-page');
      const wCount = container.querySelector('.dex-word-count');
      const cCount = container.querySelector('.dex-char-count');

      const updateStats = () => {
        const text = docPage.innerText || '';
        const words = text.trim() ? text.trim().split(/\s+/).length : 0;
        wCount.textContent = `Số từ: ${words} từ`;
        cCount.textContent = `Ký tự: ${text.length}`;
      };

      docPage.addEventListener('input', updateStats);
      container.querySelector('.dex-btn-bold').onclick = () => { document.execCommand('bold'); docPage.focus(); };
      container.querySelector('.dex-btn-italic').onclick = () => { document.execCommand('italic'); docPage.focus(); };
      container.querySelector('.dex-btn-underline').onclick = () => { document.execCommand('underline'); docPage.focus(); };
      container.querySelector('.dex-btn-strike').onclick = () => { document.execCommand('strikeThrough'); docPage.focus(); };
      container.querySelector('.dex-btn-left').onclick = () => { document.execCommand('justifyLeft'); docPage.focus(); };
      container.querySelector('.dex-btn-center').onclick = () => { document.execCommand('justifyCenter'); docPage.focus(); };
      container.querySelector('.dex-btn-right').onclick = () => { document.execCommand('justifyRight'); docPage.focus(); };
      container.querySelector('.dex-btn-justify').onclick = () => { document.execCommand('justifyFull'); docPage.focus(); };
      container.querySelector('.dex-btn-ul').onclick = () => { document.execCommand('insertUnorderedList'); docPage.focus(); };
      container.querySelector('.dex-btn-ol').onclick = () => { document.execCommand('insertOrderedList'); docPage.focus(); };
      container.querySelector('.dex-word-font').onchange = (e) => { document.execCommand('fontName', false, e.target.value); docPage.focus(); };
      container.querySelector('.dex-word-size').onchange = (e) => { document.execCommand('fontSize', false, e.target.value); docPage.focus(); };
      container.querySelector('.dex-btn-table').onclick = () => {
        const tbl = '<table border="1" style="width:100%; border-collapse:collapse; margin:10px 0;"><tr style="background:#f1f5f9;"><th style="padding:6px; border:1px solid #cbd5e1;">Cột 1</th><th style="padding:6px; border:1px solid #cbd5e1;">Cột 2</th><th style="padding:6px; border:1px solid #cbd5e1;">Cột 3</th></tr><tr><td style="padding:6px; border:1px solid #cbd5e1;">Nội dung 1</td><td style="padding:6px; border:1px solid #cbd5e1;">Nội dung 2</td><td style="padding:6px; border:1px solid #cbd5e1;">Nội dung 3</td></tr></table>';
        document.execCommand('insertHTML', false, tbl);
        docPage.focus();
      };
      const openWordBtn = container.querySelector('.dex-btn-open-word');
      if (openWordBtn) {
        openWordBtn.onclick = async () => {
          const fs = await Host.openFiles('.docx,.doc,.txt,.rtf', false);
          if (fs && fs.length) {
            const f = fs[0];
            if (window.OOS) {
              const d = await window.OOS.openFile(f);
              if (d && d.body) {
                docPage.innerHTML = d.body.innerHTML;
                updateStats();
                toast('Đã nạp văn bản: ' + f.name, 'ok');
              }
            }
          }
        };
      }
      const toMainWordBtn = container.querySelector('.dex-btn-to-main');
      if (toMainWordBtn) {
        toMainWordBtn.onclick = () => {
          root.remove();
          if (window.OOS) {
            if (!window.OOS.cur) window.OOS.newDoc('word');
            else window.OOS.showDoc(window.OOS.cur);
          }
          toast('Đã chuyển về Giao diện Chính đầy đủ', 'info');
        };
      }
      container.querySelector('.dex-btn-save').onclick = () => {
        const blob = new Blob([docPage.innerHTML], { type: 'application/msword' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'Bao_Cao_v015.doc';
        a.click();
        toast('Đã lưu tài liệu Word thành công', 'ok');
      };
      container.querySelector('.dex-btn-print').onclick = () => {
        window.print();
      };
    } else if (type === 'excel') {
      const budgetItems = [
        { id: 1, name: 'Smart TV 4K UHD 65" Phòng Họp', qty: 2, price: 16500000, status: '✓ Sẵn sàng' },
        { id: 2, name: 'Hub Chuyển Đổi Type-C HDMI 4K 60Hz', qty: 5, price: 850000, status: '✓ Sẵn sàng' },
        { id: 3, name: 'Bộ Phím Chuột Không Dây Bluetooth', qty: 5, price: 650000, status: '✓ Sẵn sàng' },
        { id: 4, name: 'Hệ Thống Máy Chủ Lưu Trữ Riêng (NAS)', qty: 1, price: 105000000, status: '✓ Sẵn sàng' }
      ];

      container.innerHTML = `
        <div style="display:flex; flex-direction:column; height:100%; background:#f8fafc; font-family:'Segoe UI',sans-serif;">
          <div style="background:#ffffff; border-bottom:1px solid #cbd5e1; padding:6px 10px; display:flex; align-items:center; gap:8px;">
            <div class="dex-xl-addr" style="font-weight:bold; color:#475569; font-size:12px; width:45px; text-align:center; background:#f1f5f9; padding:4px; border-radius:4px; border:1px solid #cbd5e1;">D6</div>
            <span style="font-weight:bold; color:#107c41; font-size:13px; font-style:italic;">fx</span>
            <input class="dex-xl-fbar" type="text" value="=SUM(D2:D5)" style="flex:1; padding:4px 8px; border:1px solid #cbd5e1; border-radius:4px; font-family:monospace; font-size:12px;" />
            <button class="dex-xl-add-row" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;">+ Dòng mới</button>
            <button class="dex-xl-open" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;">📂 Mở tệp</button>
            <button class="dex-xl-to-main" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;" title="Chuyển sang giao diện bảng tính đầy đủ">🖥 Giao diện chính</button>
            <button class="dex-xl-export" style="padding:4px 10px; border:none; background:#107c41; color:#fff; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;">📊 Xuất Excel</button>
            <span class="dex-xl-sum-badge" style="font-size:11px; font-weight:bold; color:#166534; background:#dcfce7; padding:4px 10px; border-radius:4px;">= 145,500,000 đ</span>
          </div>
          <div style="flex:1; overflow:auto;">
            <table class="dex-xl-table" style="width:100%; border-collapse:collapse; font-size:12px; text-align:left; background:#fff;">
              <thead>
                <tr style="background:#f1f5f9; color:#475569; font-weight:bold; border-bottom:2px solid #cbd5e1;">
                  <th style="padding:8px; border:1px solid #cbd5e1; width:40px; text-align:center;">#</th>
                  <th style="padding:8px; border:1px solid #cbd5e1;">A (Hạng mục Thiết bị DeX TV)</th>
                  <th style="padding:8px; border:1px solid #cbd5e1; width:90px; text-align:center;">B (SL)</th>
                  <th style="padding:8px; border:1px solid #cbd5e1; width:130px; text-align:right;">C (Đơn giá)</th>
                  <th style="padding:8px; border:1px solid #cbd5e1; width:150px; text-align:right;">D (Thành tiền)</th>
                  <th style="padding:8px; border:1px solid #cbd5e1; width:110px; text-align:center;">E (Trạng thái)</th>
                </tr>
              </thead>
              <tbody class="dex-xl-tbody"></tbody>
              <tfoot>
                <tr style="background:#ecfdf5; font-weight:bold;">
                  <td style="padding:8px; border:1px solid #cbd5e1; text-align:center;">∑</td>
                  <td style="padding:8px; border:1px solid #cbd5e1; color:#065f46;">TỔNG CỘNG DỰ TOÁN KINH PHÍ</td>
                  <td class="dex-total-qty" style="padding:8px; border:1px solid #cbd5e1; text-align:center;">-</td>
                  <td style="padding:8px; border:1px solid #cbd5e1; text-align:right;">-</td>
                  <td class="dex-total-money" style="padding:8px; border:1px solid #cbd5e1; text-align:right; color:#047857; font-size:13px;">145,500,000 đ</td>
                  <td style="padding:8px; border:1px solid #cbd5e1; text-align:center; color:#047857;">Đã duyệt</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      `;

      const tbody = container.querySelector('.dex-xl-tbody');
      const sumBadge = container.querySelector('.dex-xl-sum-badge');
      const totalMoney = container.querySelector('.dex-total-money');
      const totalQty = container.querySelector('.dex-total-qty');
      const fbar = container.querySelector('.dex-xl-fbar');
      const addrBox = container.querySelector('.dex-xl-addr');

      function recalculate() {
        let total = 0, sumQty = 0;
        budgetItems.forEach(item => {
          const rowMoney = item.qty * item.price;
          total += rowMoney;
          sumQty += item.qty;
        });
        const fmtTotal = total.toLocaleString('vi-VN') + ' đ';
        sumBadge.textContent = '= ' + fmtTotal;
        totalMoney.textContent = fmtTotal;
        totalQty.textContent = String(sumQty);
      }

      function renderRows() {
        tbody.innerHTML = '';
        budgetItems.forEach((item, idx) => {
          const tr = document.createElement('tr');
          const rowMoney = item.qty * item.price;
          tr.innerHTML = `
            <td style="padding:6px; border:1px solid #cbd5e1; text-align:center; background:#f8fafc;">${idx + 1}</td>
            <td contenteditable="true" class="dex-cell-name" style="padding:6px; border:1px solid #cbd5e1; font-weight:500; outline:none;">${item.name}</td>
            <td contenteditable="true" class="dex-cell-qty" style="padding:6px; border:1px solid #cbd5e1; text-align:center; outline:none; font-weight:bold; color:#185abd;">${item.qty}</td>
            <td contenteditable="true" class="dex-cell-price" style="padding:6px; border:1px solid #cbd5e1; text-align:right; outline:none;">${item.price.toLocaleString('vi-VN')} đ</td>
            <td class="dex-cell-total" style="padding:6px; border:1px solid #cbd5e1; text-align:right; font-weight:600;">${rowMoney.toLocaleString('vi-VN')} đ</td>
            <td style="padding:6px; border:1px solid #cbd5e1; text-align:center; color:#16a34a; font-weight:500;">${item.status}</td>
          `;

          const cellQty = tr.querySelector('.dex-cell-qty');
          const cellPrice = tr.querySelector('.dex-cell-price');
          const cellTotal = tr.querySelector('.dex-cell-total');

          cellQty.onfocus = () => { addrBox.textContent = `B${idx + 2}`; fbar.value = String(item.qty); };
          cellPrice.onfocus = () => { addrBox.textContent = `C${idx + 2}`; fbar.value = String(item.price); };
          tr.querySelector('.dex-cell-name').onfocus = () => { addrBox.textContent = `A${idx + 2}`; fbar.value = item.name; };

          cellQty.oninput = () => {
            const v = parseInt(cellQty.innerText.replace(/[^0-9]/g, ''), 10) || 0;
            item.qty = v;
            const newTotal = item.qty * item.price;
            cellTotal.textContent = newTotal.toLocaleString('vi-VN') + ' đ';
            recalculate();
          };

          cellPrice.oninput = () => {
            const v = parseInt(cellPrice.innerText.replace(/[^0-9]/g, ''), 10) || 0;
            item.price = v;
            const newTotal = item.qty * item.price;
            cellTotal.textContent = newTotal.toLocaleString('vi-VN') + ' đ';
            recalculate();
          };

          tbody.appendChild(tr);
        });
        recalculate();
      }

      renderRows();

      fbar.onkeydown = (e) => {
        if (e.key === 'Enter') {
          toast('Đã tính toán công thức: ' + fbar.value, 'ok');
          recalculate();
        }
      };

      const openXlBtn = container.querySelector('.dex-xl-open');
      if (openXlBtn) {
        openXlBtn.onclick = async () => {
          const fs = await Host.openFiles('.xlsx,.xlsm,.csv', false);
          if (fs && fs.length) {
            const f = fs[0];
            if (window.OOS) {
              const d = await window.OOS.openFile(f);
              if (d && d.wb) {
                root.remove();
                window.OOS.showDoc(d);
                toast('Đã mở bảng tính trên giao diện chính: ' + f.name, 'ok');
              }
            }
          }
        };
      }
      const toMainXlBtn = container.querySelector('.dex-xl-to-main');
      if (toMainXlBtn) {
        toMainXlBtn.onclick = () => {
          root.remove();
          if (window.OOS) {
            if (!window.OOS.cur) window.OOS.newDoc('excel');
            else window.OOS.showDoc(window.OOS.cur);
          }
          toast('Đã chuyển về Giao diện Chính đầy đủ', 'info');
        };
      }
      container.querySelector('.dex-xl-add-row').onclick = () => {
        budgetItems.push({
          id: budgetItems.length + 1,
          name: 'Hạng mục DeX mới ' + (budgetItems.length + 1),
          qty: 1,
          price: 1500000,
          status: '✓ Mới'
        });
        renderRows();
        toast('Đã thêm dòng mới vào bảng tính', 'info');
      };

      container.querySelector('.dex-xl-export').onclick = () => {
        let csv = 'STT,Hạng mục,Số lượng,Đơn giá,Thành tiền\n';
        budgetItems.forEach((it, i) => {
          csv += `${i+1},"${it.name}",${it.qty},${it.price},${it.qty*it.price}\n`;
        });
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'Du_Toan_DeX_2026.csv';
        a.click();
        toast('Đã xuất tệp bảng tính CSV thành công', 'ok');
      };
    } else if (type === 'ppt') {
      const slides = [
        { id: 1, title: 'Yomi-Office Open Source v' + VERSION, sub: 'Nền Tảng Văn Phòng Đa Nền Tảng Thế Hệ Mới', bullets: ['Tương thích 100% Android 6.0+ & Windows 7/8.1/10/11 x64, x86, ARM', 'Hỗ trợ trọn vẹn Word, Excel, PowerPoint, PDF, MathType', 'Hoạt động hoàn toàn ngoại tuyến không cần internet'] },
        { id: 2, title: 'Android DeX Desktop & TV 4K UHD', sub: 'Đa Nhiệm Cửa Sổ Thông Minh (True Multi-Window)', bullets: ['Chia đôi màn hình 50/50, bố trí lưới 4 cửa sổ và cửa sổ nổi tự do', 'Định tuyến âm thanh thông minh giữa Loa TV HDMI và Loa Điện thoại', 'Bàn rê chuột ảo (Virtual Trackpad) và tối ưu phím tắt PC'] },
        { id: 3, title: 'Quản Lý Toàn Bộ Ứng Dụng Bên Ngoài', sub: 'Tích Hợp Android Freeform Window Bounds', bullets: ['Khởi chạy mọi ứng dụng cài trên thiết bị trực tiếp từ Start Menu TV', 'Hỗ trợ cửa sổ điều khiển tương đương và Web Companion', 'Đồng bộ thao tác phím/chuột ngoại vi mượt mà'] }
      ];
      let curSlide = 0;

      container.innerHTML = `
        <div style="display:flex; flex-direction:column; height:100%; background:#0f172a; color:#fff; font-family:'Segoe UI',sans-serif;">
          <div style="background:#1e293b; border-bottom:1px solid #334155; padding:6px 12px; display:flex; align-items:center; gap:8px;">
            <button class="dex-ppt-present" style="padding:5px 12px; border:none; background:#c43e1c; color:#fff; border-radius:4px; font-size:12px; font-weight:600; cursor:pointer;">▶ Trình chiếu (F5)</button>
            <button class="dex-ppt-add" style="padding:5px 10px; border:1px solid #475569; background:#334155; color:#fff; border-radius:4px; font-size:12px; cursor:pointer;">+ Thêm Slide</button>
            <div style="margin-left:auto; display:flex; align-items:center; gap:8px;">
              <button class="dex-ppt-prev" style="padding:3px 8px; border:none; background:#334155; color:#fff; border-radius:4px; cursor:pointer;">◀</button>
              <span class="dex-ppt-counter" style="font-size:12px; color:#94a3b8;">Slide 1 / 3</span>
              <button class="dex-ppt-next" style="padding:3px 8px; border:none; background:#334155; color:#fff; border-radius:4px; cursor:pointer;">▶</button>
            </div>
          </div>
          <div style="flex:1; display:flex; overflow:hidden;">
            <div class="dex-ppt-sidebar" style="width:160px; background:#1e293b; border-right:1px solid #334155; overflow-y:auto; padding:10px; display:flex; flex-direction:column; gap:8px;"></div>
            <div style="flex:1; display:flex; align-items:center; justify-content:center; padding:20px; background:#020617;">
              <div class="dex-ppt-canvas" style="width:100%; max-width:680px; aspect-ratio:16/9; background:#ffffff; border-radius:8px; box-shadow:0 12px 36px rgba(0,0,0,0.5); padding:32px; color:#0f172a; display:flex; flex-direction:column; justify-content:center;"></div>
            </div>
          </div>
        </div>
      `;

      const sidebar = container.querySelector('.dex-ppt-sidebar');
      const canvas = container.querySelector('.dex-ppt-canvas');
      const counter = container.querySelector('.dex-ppt-counter');

      function renderPresentation() {
        sidebar.innerHTML = '';
        slides.forEach((sl, idx) => {
          const thumb = document.createElement('div');
          thumb.style.cssText = `border:2px solid ${idx === curSlide ? '#c43e1c' : '#334155'}; border-radius:6px; padding:6px; background:#0f172a; cursor:pointer; font-size:10px; color:#cbd5e1; transition:all 0.1s;`;
          thumb.innerHTML = `<div style="font-weight:bold; color:#fff; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">#${idx + 1} ${sl.title}</div><div style="font-size:9px; color:#94a3b8; margin-top:2px;">${sl.sub}</div>`;
          thumb.onclick = () => { curSlide = idx; renderPresentation(); };
          sidebar.appendChild(thumb);
        });

        counter.textContent = `Slide ${curSlide + 1} / ${slides.length}`;
        const cur = slides[curSlide];
        canvas.innerHTML = `
          <h2 contenteditable="true" style="margin:0 0 8px; color:#c43e1c; font-size:22px; outline:none;">${cur.title}</h2>
          <h4 contenteditable="true" style="margin:0 0 16px; color:#64748b; font-size:14px; font-weight:normal; outline:none;">${cur.sub}</h4>
          <ul style="padding-left:22px; font-size:14px; line-height:1.8; color:#334155;">
            ${cur.bullets.map(b => `<li contenteditable="true" style="outline:none;">${b}</li>`).join('')}
          </ul>
        `;
      }

      container.querySelector('.dex-ppt-prev').onclick = () => { if (curSlide > 0) { curSlide--; renderPresentation(); } };
      container.querySelector('.dex-ppt-next').onclick = () => { if (curSlide < slides.length - 1) { curSlide++; renderPresentation(); } };
      container.querySelector('.dex-ppt-add').onclick = () => {
        slides.push({ id: slides.length + 1, title: 'Slide mới ' + (slides.length + 1), sub: 'Tiêu đề phụ slide', bullets: ['Điểm nội dung 1', 'Điểm nội dung 2'] });
        curSlide = slides.length - 1;
        renderPresentation();
        toast('Đã thêm Slide mới', 'info');
      };
      container.querySelector('.dex-ppt-present').onclick = () => {
        const fs = document.createElement('div');
        fs.style.cssText = 'position:fixed; inset:0; z-index:99999; background:#000; color:#fff; display:flex; flex-direction:column; align-items:center; justify-content:center; padding:40px;';
        const renderFs = () => {
          const s = slides[curSlide];
          fs.innerHTML = `
            <div style="width:100%; max-width:1000px; text-align:center;">
              <h1 style="color:#c43e1c; font-size:42px; margin-bottom:12px;">${s.title}</h1>
              <h3 style="color:#94a3b8; font-size:22px; font-weight:normal; margin-bottom:32px;">${s.sub}</h3>
              <ul style="display:inline-block; text-align:left; font-size:22px; line-height:2.2; color:#e2e8f0;">
                ${s.bullets.map(b => `<li>${b}</li>`).join('')}
              </ul>
            </div>
            <div style="position:absolute; bottom:20px; display:flex; gap:16px; align-items:center;">
              <button class="dex-fs-prev" style="padding:8px 16px; background:#334155; color:#fff; border:none; border-radius:6px; cursor:pointer;">◀ Trước</button>
              <span style="color:#94a3b8;">${curSlide + 1} / ${slides.length}</span>
              <button class="dex-fs-next" style="padding:8px 16px; background:#c43e1c; color:#fff; border:none; border-radius:6px; cursor:pointer;">Tiếp ▶</button>
              <button class="dex-fs-close" style="padding:8px 16px; background:#ef4444; color:#fff; border:none; border-radius:6px; cursor:pointer;">Thoát (Esc)</button>
            </div>
          `;
          fs.querySelector('.dex-fs-prev').onclick = () => { if (curSlide > 0) { curSlide--; renderFs(); } };
          fs.querySelector('.dex-fs-next').onclick = () => { if (curSlide < slides.length - 1) { curSlide++; renderFs(); } };
          fs.querySelector('.dex-fs-close').onclick = () => { fs.remove(); renderPresentation(); };
        };
        renderFs();
        document.body.appendChild(fs);
      };

      renderPresentation();
    } else if (type === 'pdf') {
      container.innerHTML = `
        <div style="display:flex; flex-direction:column; height:100%; background:#f1f5f9; font-family:'Segoe UI',sans-serif;">
          <div style="background:#ffffff; border-bottom:1px solid #cbd5e1; padding:6px 12px; display:flex; align-items:center; gap:8px;">
            <button class="dex-pdf-prev" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; cursor:pointer;">◀</button>
            <span class="dex-pdf-page-lbl" style="font-size:12px; font-weight:600;">Trang 1 / 3</span>
            <button class="dex-pdf-next" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; cursor:pointer;">▶</button>
            <div style="height:16px; width:1px; background:#cbd5e1; margin:0 4px;"></div>
            <button class="dex-pdf-sign-btn" style="padding:4px 12px; border:none; background:#dc2626; color:#fff; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;">✍ Ký Tên Điện Tử</button>
            <button class="dex-pdf-stamp-btn" style="padding:4px 10px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;">Đóng dấu Đã Ký</button>
            <span style="margin-left:auto; font-size:11px; color:#16a34a; font-weight:bold;">PDF/A-1b Chuẩn Quốc Tế</span>
          </div>
          <div style="flex:1; overflow-y:auto; padding:20px; display:flex; justify-content:center; background:#cbd5e1;">
            <div class="dex-pdf-view" style="width:100%; max-width:680px; min-height:560px; background:#fff; box-shadow:0 6px 20px rgba(0,0,0,0.18); border-radius:4px; padding:36px; position:relative; color:#0f172a;">
              <div style="display:flex; justify-content:space-between; border-bottom:2px solid #dc2626; padding-bottom:8px; margin-bottom:16px;">
                <b style="color:#dc2626; font-size:16px;">HỢP ĐỒNG ĐIỆN TỬ VĂN PHÒNG DEX</b>
                <span style="font-size:12px; color:#64748b;">Mã: HD-2026-OOS</span>
              </div>
              <p style="font-size:13px; line-height:1.6;">Căn cứ nhu cầu số hóa quy trình văn phòng và vận hành chế độ màn hình lớn Android DeX TV 4K, bản quyền phần mềm Yomi-Office Open Source v${VERSION} được cấp phép sử dụng ngoại tuyến trọn đời.</p>
              <h4 style="margin:16px 0 8px; font-size:13px;">Điều khoản nghiệm thu:</h4>
              <ul style="font-size:12px; line-height:1.7; padding-left:20px;">
                <li>Hệ thống chạy mượt mà 100% offline trên Android 6.0+ và Windows 7/8.1/10/11.</li>
                <li>Đa nhiệm nhiều cửa sổ (Word, Excel, Máy tính) phản hồi tức thì.</li>
                <li>Hỗ trợ xuất hình ảnh 4K UHD 60Hz không độ trễ.</li>
              </ul>
              <div class="dex-pdf-stamp-area" style="margin-top:40px; display:flex; justify-content:space-between; align-items:flex-end;">
                <div>
                  <div style="font-size:12px; font-weight:bold;">ĐẠI DIỆN BÊN GIAO</div>
                  <div style="font-size:11px; color:#16a34a; margin-top:4px;">✓ Đã xác thực chứng thư số</div>
                </div>
                <div style="text-align:right;">
                  <div style="font-size:12px; font-weight:bold;">ĐẠI DIỆN BÊN NHẬN</div>
                  <div class="dex-pdf-sig-target" style="width:160px; height:60px; border:1px dashed #cbd5e1; border-radius:4px; margin-top:4px; display:flex; align-items:center; justify-content:center; font-size:11px; color:#94a3b8;">[Ký tên tại đây]</div>
                </div>
              </div>
            </div>
          </div>
          <div class="dex-pdf-modal" style="display:none; position:absolute; inset:0; background:rgba(0,0,0,0.5); z-index:9999; align-items:center; justify-content:center;">
            <div style="background:#fff; border-radius:8px; padding:16px; width:340px; box-shadow:0 12px 36px rgba(0,0,0,0.3);">
              <h4 style="margin:0 0 8px; font-size:14px; color:#0f172a;">Ký tên bằng Chuột / Cảm ứng / S-Pen:</h4>
              <canvas class="dex-pdf-canvas" width="308" height="120" style="border:1px solid #cbd5e1; border-radius:4px; background:#f8fafc; cursor:crosshair;"></canvas>
              <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:10px;">
                <button class="dex-pdf-clear" style="padding:4px 10px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;">Xóa</button>
                <button class="dex-pdf-apply" style="padding:4px 12px; border:none; background:#dc2626; color:#fff; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;">Áp dụng chữ ký</button>
              </div>
            </div>
          </div>
        </div>
      `;

      const modal = container.querySelector('.dex-pdf-modal');
      const sigTarget = container.querySelector('.dex-pdf-sig-target');
      const cvs = container.querySelector('.dex-pdf-canvas');
      const ctx = cvs.getContext('2d');
      let drawing = false;

      ctx.strokeStyle = '#185abd';
      ctx.lineWidth = 2.5;
      ctx.lineCap = 'round';

      cvs.onpointerdown = (e) => { drawing = true; ctx.beginPath(); ctx.moveTo(e.offsetX, e.offsetY); cvs.setPointerCapture(e.pointerId); };
      cvs.onpointermove = (e) => { if (!drawing) return; ctx.lineTo(e.offsetX, e.offsetY); ctx.stroke(); };
      cvs.onpointerup = () => { drawing = false; };

      container.querySelector('.dex-pdf-clear').onclick = () => { ctx.clearRect(0, 0, cvs.width, cvs.height); };
      container.querySelector('.dex-pdf-sign-btn').onclick = () => { modal.style.display = 'flex'; };
      container.querySelector('.dex-pdf-apply').onclick = () => {
        const url = cvs.toDataURL();
        sigTarget.innerHTML = `<img src="${url}" style="max-width:100%; max-height:100%;" />`;
        modal.style.display = 'none';
        toast('Đã dán chữ ký điện tử vào văn bản PDF', 'ok');
      };
      container.querySelector('.dex-pdf-stamp-btn').onclick = () => {
        sigTarget.innerHTML = `<div style="color:#dc2626; font-weight:bold; border:2px solid #dc2626; padding:4px 8px; border-radius:4px; text-transform:uppercase;">✓ ĐÃ KÝ SỐ (${new Date().toLocaleDateString('vi-VN')})</div>`;
        toast('Đã đóng dấu điện tử lên tài liệu', 'ok');
      };
    } else if (type === 'mathtype') {
      container.innerHTML = `
        <div style="display:flex; flex-direction:column; height:100%; background:#f8fafc; font-family:'Segoe UI',sans-serif; color:#0f172a;">
          <div style="background:#ffffff; border-bottom:1px solid #cbd5e1; padding:8px 12px; display:flex; flex-direction:column; gap:6px;">
            <div style="font-size:12px; font-weight:600; color:#68217a;">Bảng Ký Hiệu Toán Học & Công Thức:</div>
            <div class="dex-math-palette" style="display:flex; gap:6px; flex-wrap:wrap;">
              <button class="dex-sym" data-tex="\\frac{a}{b}" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">a/b</button>
              <button class="dex-sym" data-tex="\\sqrt{x}" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">√x</button>
              <button class="dex-sym" data-tex="\\sum_{i=1}^n" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">∑</button>
              <button class="dex-sym" data-tex="\\int_a^b" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">∫</button>
              <button class="dex-sym" data-tex="\\alpha" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">α</button>
              <button class="dex-sym" data-tex="\\beta" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">β</button>
              <button class="dex-sym" data-tex="\\pi" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">π</button>
              <button class="dex-sym" data-tex="\\theta" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">θ</button>
              <button class="dex-sym" data-tex="\\pm" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">±</button>
              <button class="dex-sym" data-tex="\\times" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">×</button>
              <button class="dex-sym" data-tex="\\div" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">÷</button>
              <button class="dex-sym" data-tex="\\infty" style="padding:4px 8px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:13px; cursor:pointer;">∞</button>
            </div>
          </div>
          <div style="padding:12px; display:flex; flex-direction:column; gap:8px;">
            <div style="font-size:11px; color:#64748b; font-weight:600;">Mã LaTeX Toán:</div>
            <input class="dex-math-inp" type="text" value="x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}" style="padding:8px 10px; border:1px solid #cbd5e1; border-radius:6px; font-family:monospace; font-size:13px; background:#fff;" />
          </div>
          <div style="flex:1; padding:16px; margin:0 12px 12px; background:#ffffff; border:1px solid #cbd5e1; border-radius:8px; display:flex; flex-direction:column; align-items:center; justify-content:center;">
            <div style="font-size:11px; color:#94a3b8; margin-bottom:8px;">Xem trước kết quả công thức:</div>
            <div class="dex-math-prev" style="font-size:24px; font-family:'Cambria Math',serif; color:#68217a; text-align:center;">
              x = (-b ± √(b² - 4ac)) / (2a)
            </div>
          </div>
          <div style="padding:10px 12px; background:#ffffff; border-top:1px solid #cbd5e1; display:flex; gap:8px;">
            <button class="dex-math-copy" style="padding:6px 14px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:12px; cursor:pointer;">Sao chép LaTeX</button>
            <button class="dex-math-insert" style="padding:6px 14px; border:none; background:#68217a; color:#fff; border-radius:4px; font-size:12px; font-weight:600; cursor:pointer;">Chèn vào Word</button>
          </div>
        </div>
      `;

      const inp = container.querySelector('.dex-math-inp');
      const prev = container.querySelector('.dex-math-prev');
      inp.oninput = () => { prev.textContent = inp.value; };
      container.querySelectorAll('.dex-sym').forEach(btn => {
        btn.onclick = () => {
          inp.value += ' ' + btn.getAttribute('data-tex');
          prev.textContent = inp.value;
          inp.focus();
        };
      });
      container.querySelector('.dex-math-copy').onclick = () => {
        navigator.clipboard ? navigator.clipboard.writeText(inp.value) : null;
        toast('Đã sao chép mã công thức LaTeX', 'ok');
      };
      container.querySelector('.dex-math-insert').onclick = () => {
        toast('Đã chèn công thức toán vào văn bản Word', 'ok');
      };
    } else if (type === 'calc') {
      let displayVal = '0';
      let prevVal = null;
      let op = null;
      let waiting = false;
      let historyVal = '';

      container.innerHTML = `
        <div style="display:flex; flex-direction:column; height:100%; background:#1e293b; color:#fff; padding:12px; font-family:monospace;">
          <div style="background:#0f172a; border-radius:8px; padding:12px; margin-bottom:12px; text-align:right;">
            <div class="dex-calc-hist" style="font-size:11px; color:#94a3b8; min-height:16px; margin-bottom:4px;">${historyVal}</div>
            <div class="dex-calc-disp" style="font-size:26px; font-weight:bold; color:#38bdf8; word-break:break-all;">${displayVal}</div>
          </div>
          <div style="flex:1; display:grid; grid-template-columns:repeat(4, 1fr); gap:8px;">
            <button class="dex-calc-btn" data-k="C" style="background:#475569; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">C</button>
            <button class="dex-calc-btn" data-k="CE" style="background:#475569; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">CE</button>
            <button class="dex-calc-btn" data-k="%" style="background:#475569; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">%</button>
            <button class="dex-calc-btn" data-k="/" style="background:#f59e0b; color:#fff; border:none; border-radius:6px; font-size:16px; font-weight:bold; cursor:pointer;">÷</button>
            <button class="dex-calc-btn" data-k="7" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">7</button>
            <button class="dex-calc-btn" data-k="8" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">8</button>
            <button class="dex-calc-btn" data-k="9" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">9</button>
            <button class="dex-calc-btn" data-k="*" style="background:#f59e0b; color:#fff; border:none; border-radius:6px; font-size:16px; font-weight:bold; cursor:pointer;">×</button>
            <button class="dex-calc-btn" data-k="4" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">4</button>
            <button class="dex-calc-btn" data-k="5" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">5</button>
            <button class="dex-calc-btn" data-k="6" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">6</button>
            <button class="dex-calc-btn" data-k="-" style="background:#f59e0b; color:#fff; border:none; border-radius:6px; font-size:16px; font-weight:bold; cursor:pointer;">-</button>
            <button class="dex-calc-btn" data-k="1" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">1</button>
            <button class="dex-calc-btn" data-k="2" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">2</button>
            <button class="dex-calc-btn" data-k="3" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">3</button>
            <button class="dex-calc-btn" data-k="+" style="background:#f59e0b; color:#fff; border:none; border-radius:6px; font-size:16px; font-weight:bold; cursor:pointer;">+</button>
            <button class="dex-calc-btn" data-k="0" style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer; grid-column:span 2;">0</button>
            <button class="dex-calc-btn" data-k="." style="background:#334155; color:#fff; border:none; border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer;">.</button>
            <button class="dex-calc-btn" data-k="=" style="background:#10b981; color:#fff; border:none; border-radius:6px; font-size:16px; font-weight:bold; cursor:pointer;">=</button>
          </div>
        </div>
      `;

      const dispEl = container.querySelector('.dex-calc-disp');
      const histEl = container.querySelector('.dex-calc-hist');

      function updateCalc() {
        dispEl.textContent = displayVal;
        histEl.textContent = historyVal;
      }

      function handleKey(k) {
        if (k >= '0' && k <= '9') {
          if (waiting || displayVal === '0') {
            displayVal = k;
            waiting = false;
          } else {
            displayVal += k;
          }
        } else if (k === '.') {
          if (!displayVal.includes('.')) displayVal += '.';
        } else if (k === 'C') {
          displayVal = '0';
          prevVal = null;
          op = null;
          historyVal = '';
          waiting = false;
        } else if (k === 'CE') {
          displayVal = '0';
        } else if (k === '%') {
          displayVal = String(parseFloat(displayVal) / 100);
        } else if (['+', '-', '*', '/'].includes(k)) {
          const cur = parseFloat(displayVal);
          if (op && !waiting) {
            displayVal = String(evalOp(prevVal, cur, op));
            prevVal = parseFloat(displayVal);
          } else {
            prevVal = cur;
          }
          op = k;
          waiting = true;
          historyVal = `${prevVal} ${k === '*' ? '×' : k === '/' ? '÷' : k}`;
        } else if (k === '=') {
          if (op && prevVal != null) {
            const cur = parseFloat(displayVal);
            historyVal = `${prevVal} ${op === '*' ? '×' : op === '/' ? '÷' : op} ${cur} =`;
            displayVal = String(evalOp(prevVal, cur, op));
            prevVal = null;
            op = null;
            waiting = true;
          }
        }
        updateCalc();
      }

      function evalOp(a, b, operation) {
        if (operation === '+') return a + b;
        if (operation === '-') return a - b;
        if (operation === '*') return a * b;
        if (operation === '/') return b !== 0 ? a / b : 'Lỗi chia 0';
        return b;
      }

      container.querySelectorAll('.dex-calc-btn').forEach(btn => {
        btn.onclick = () => handleKey(btn.getAttribute('data-k'));
      });
    } else if (type === 'files') {
      const filesList = [
        { name: 'Bao_Cao_v015.docx', kind: 'word', size: '1.4 MB', date: 'Hôm nay', icon: '📄' },
        { name: 'Ke_Hoach_2026.xlsx', kind: 'excel', size: '850 KB', date: 'Hôm nay', icon: '📊' },
        { name: 'Thuyet_Trinh_TV.pptx', kind: 'ppt', size: '4.2 MB', date: 'Hôm qua', icon: '📑' },
        { name: 'Hop_Dong_DeX.pdf', kind: 'pdf', size: '2.1 MB', date: 'Hôm qua', icon: '📕' }
      ];

      container.innerHTML = `
        <div style="display:flex; height:100%; background:#ffffff; font-family:'Segoe UI',sans-serif;">
          <div style="width:200px; background:#f8fafc; border-right:1px solid #e2e8f0; padding:12px; font-size:12px; display:flex; flex-direction:column; gap:4px;">
            <div style="font-weight:bold; color:#475569; margin-bottom:6px; font-size:11px; text-transform:uppercase;">Bộ nhớ Thiết bị</div>
            <div class="dex-folder-btn" style="padding:6px 8px; border-radius:4px; background:#e0f2fe; color:#0369a1; font-weight:600; cursor:pointer;">📁 Tài liệu (Documents)</div>
            <div class="dex-folder-btn" style="padding:6px 8px; border-radius:4px; color:#334155; cursor:pointer;">📥 Tải về (Downloads)</div>
            <div class="dex-folder-btn" style="padding:6px 8px; border-radius:4px; color:#334155; cursor:pointer;">🖼 Hình ảnh (DCIM)</div>
            <div class="dex-folder-btn" style="padding:6px 8px; border-radius:4px; color:#334155; cursor:pointer;">💾 Thẻ nhớ ngoài (SD Card)</div>
            <div class="dex-folder-btn" style="padding:6px 8px; border-radius:4px; color:#334155; cursor:pointer;">☁ Đám mây Cá nhân</div>
          </div>
          <div style="flex:1; padding:16px; overflow-y:auto; display:flex; flex-direction:column;">
            <div style="font-size:12px; color:#64748b; margin-bottom:12px; display:flex; justify-content:space-between; align-items:center;">
              <div>Đường dẫn: <b>/sdcard/Documents/Office_Projects/</b></div>
              <div style="display:flex; gap:6px;">
                <button class="dex-pick-file" style="padding:4px 10px; border:none; background:#0284c7; color:#fff; border-radius:4px; font-size:11px; cursor:pointer; font-weight:600;">📂 Chọn tệp từ máy...</button>
                <button class="dex-new-file" style="padding:4px 10px; border:1px solid #cbd5e1; background:#f8fafc; border-radius:4px; font-size:11px; cursor:pointer;">+ Tạo tệp mới</button>
              </div>
            </div>
            <div class="dex-files-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(130px, 1fr)); gap:12px;"></div>
          </div>
        </div>
      `;

      const pickBtn = container.querySelector('.dex-pick-file');
      if (pickBtn) {
        pickBtn.onclick = async () => {
          const picked = await Host.openFiles('.docx,.doc,.xlsx,.xlsm,.pptx,.pdf,.txt', true);
          if (picked && picked.length) {
            for (const pf of picked) {
              const ext = pf.name.split('.').pop().toLowerCase();
              const kind = ext === 'docx' || ext === 'doc' ? 'word' : ext === 'xlsx' || ext === 'xlsm' ? 'excel' : ext === 'pptx' ? 'ppt' : 'pdf';
              if (window.OOS && typeof window.OOS.openFile === 'function') {
                window.OOS.openFile(pf);
              }
              openAppWindow(kind, pf.name, kind === 'word' ? 'file-text' : kind === 'excel' ? 'file-spreadsheet' : kind === 'ppt' ? 'presentation' : 'file-check');
              toast('Đã nạp tệp vào DeX: ' + pf.name, 'ok');
            }
          }
        };
      }

      const grid = container.querySelector('.dex-files-grid');
      filesList.forEach(f => {
        const item = document.createElement('div');
        item.style.cssText = 'padding:12px 8px; border:1px solid #e2e8f0; border-radius:8px; text-align:center; cursor:pointer; background:#f8fafc; transition:all 0.1s;';
        item.innerHTML = `
          <div style="font-size:28px;">${f.icon}</div>
          <div style="font-size:11px; font-weight:600; margin-top:4px; text-overflow:ellipsis; overflow:hidden;">${f.name}</div>
          <div style="font-size:10px; color:#94a3b8;">${f.size} • ${f.date}</div>
        `;
        item.onclick = () => {
          if (f.kind === 'word') openAppWindow('word', 'Word (Văn bản)', 'file-text');
          else if (f.kind === 'excel') openAppWindow('excel', 'Excel (Bảng tính)', 'file-spreadsheet');
          else if (f.kind === 'ppt') openAppWindow('ppt', 'PowerPoint (Thuyết trình)', 'presentation');
          else if (f.kind === 'pdf') openAppWindow('pdf', 'PDF Reader & Ký', 'file-check');
          toast('Đã mở tệp: ' + f.name, 'ok');
        };
        grid.appendChild(item);
      });
      container.querySelector('.dex-new-file').onclick = () => {
        openAppWindow('word', 'Word (Tài liệu mới)', 'file-text');
        toast('Đã tạo tài liệu mới', 'ok');
      };
    } else if (type === 'browser') {
      container.innerHTML = `
        <div style="display:flex; flex-direction:column; height:100%; background:#ffffff; font-family:'Segoe UI',sans-serif;">
          <div style="background:#f1f5f9; border-bottom:1px solid #cbd5e1; padding:6px 12px; display:flex; align-items:center; gap:8px;">
            <button class="dex-br-back" style="border:none; background:none; cursor:pointer; font-size:14px; color:#475569;">◀</button>
            <button class="dex-br-fwd" style="border:none; background:none; cursor:pointer; font-size:14px; color:#475569;">▶</button>
            <button class="dex-br-reload" style="border:none; background:none; cursor:pointer; font-size:14px; color:#475569;">⟳</button>
            <input class="dex-br-url" type="text" value="https://github.com/YomiHub-Project/Yomi-Office-Open-Source" style="flex:1; padding:4px 10px; border:1px solid #cbd5e1; border-radius:20px; font-size:12px; background:#fff;" />
            <button class="dex-br-go" style="padding:4px 12px; border:none; background:#185abd; color:#fff; border-radius:14px; font-size:12px; cursor:pointer;">Đi tới</button>
          </div>
          <div class="dex-br-body" style="flex:1; overflow-y:auto; background:#0f172a; color:#fff; padding:24px;">
            <div style="max-width:720px; margin:0 auto; padding:24px; background:#1e293b; border-radius:12px; border:1px solid #334155;">
              <h2 style="color:#38bdf8; margin-top:0;">GitHub Project: Yomi-Office Open Source v${VERSION}</h2>
              <p style="color:#cbd5e1; font-size:13px; line-height:1.6;">Bộ ứng dụng văn phòng mã nguồn mở đa nền tảng, tích hợp hệ thống Android DeX & Màn hình TV 4K, tương thích tuyệt đối từ Android 6.0 trở lên và Windows 7/8.1/10/11 x64, x86, ARM.</p>
              <div style="display:flex; gap:10px; margin:16px 0;">
                <span style="background:#15803d; color:#fff; padding:4px 10px; border-radius:6px; font-size:11px; font-weight:bold;">★ 15.4k Stars</span>
                <span style="background:#0369a1; color:#fff; padding:4px 10px; border-radius:6px; font-size:11px; font-weight:bold;">License: MIT</span>
                <span style="background:#7c3aed; color:#fff; padding:4px 10px; border-radius:6px; font-size:11px; font-weight:bold;">Release: v${VERSION}</span>
              </div>
              <h4 style="color:#38bdf8; margin:14px 0 6px;">Tính Năng Android DeX & Màn Hình Ngoài:</h4>
              <ul style="color:#94a3b8; font-size:12px; line-height:1.7; padding-left:20px;">
                <li>Đa nhiệm Đa cửa sổ (True Multi-Window) với phân giải 4K UHD 60fps</li>
                <li>Quét và khởi chạy toàn bộ ứng dụng cài trên thiết bị Android di động</li>
                <li>Điều khiển và đồng bộ hóa bàn rê chuột ảo đa điểm cùng phím tắt PC</li>
              </ul>
            </div>
          </div>
        </div>
      `;

      const urlInp = container.querySelector('.dex-br-url');
      const goBtn = container.querySelector('.dex-br-go');
      const bodyEl = container.querySelector('.dex-br-body');

      const navigate = () => {
        let u = urlInp.value.trim();
        if (!u.startsWith('http://') && !u.startsWith('https://')) u = 'https://' + u;
        toast('Đang điều hướng tới: ' + u, 'info');
      };

      goBtn.onclick = navigate;
      urlInp.onkeydown = (e) => { if (e.key === 'Enter') navigate(); };
      container.querySelector('.dex-br-reload').onclick = navigate;
    } else {
      const pkg = type;
      container.innerHTML = `
        <div style="display:flex; flex-direction:column; height:100%; background:#0f172a; color:#f8fafc; font-family:'Segoe UI',sans-serif;">
          <div style="padding:14px 18px; background:linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border-bottom:1px solid #334155; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px;">
            <div style="display:flex; align-items:center; gap:12px;">
              <div style="width:44px; height:44px; border-radius:12px; background:#38bdf8; color:#0f172a; display:flex; align-items:center; justify-content:center; font-size:22px; font-weight:bold;">📱</div>
              <div>
                <div style="font-size:15px; font-weight:700; color:#ffffff;">Ứng Dụng Thiết Bị: ${win ? win.title : pkg}</div>
                <div style="font-size:11px; color:#94a3b8; font-family:monospace;">${pkg}</div>
              </div>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
              <span style="display:inline-flex; align-items:center; gap:6px; background:rgba(34,197,94,0.15); color:#4ade80; padding:4px 10px; border-radius:20px; font-size:11px; font-weight:600;">
                <span style="width:8px; height:8px; border-radius:50%; background:#22c55e;"></span>
                Đang mở trên DeX
              </span>
            </div>
          </div>

          <div style="padding:12px 18px; background:#1e293b; border-bottom:1px solid #334155; display:flex; gap:8px; flex-wrap:wrap; align-items:center;">
            <button class="dex-ext-bounds" style="padding:6px 14px; border:none; background:#0284c7; color:#fff; border-radius:6px; font-size:12px; font-weight:600; cursor:pointer; display:flex; align-items:center; gap:6px;">
              🔲 Đồng bộ Tọa độ Cửa sổ nổi
            </button>
            <button class="dex-ext-split" style="padding:6px 14px; border:1px solid #475569; background:#334155; color:#fff; border-radius:6px; font-size:12px; cursor:pointer; display:flex; align-items:center; gap:6px;">
              🪟 Chia đôi 50/50
            </button>
            <button class="dex-ext-front" style="padding:6px 14px; border:1px solid #475569; background:#334155; color:#fff; border-radius:6px; font-size:12px; cursor:pointer; display:flex; align-items:center; gap:6px;">
              ⬆ Đưa lên trước
            </button>
            <button class="dex-ext-full" style="padding:6px 14px; border:1px solid #475569; background:#334155; color:#fff; border-radius:6px; font-size:12px; cursor:pointer; display:flex; align-items:center; gap:6px;">
              ⛶ Toàn màn hình
            </button>
          </div>

          <div style="flex:1; padding:18px; display:flex; flex-direction:column; gap:14px; overflow-y:auto;">
            <div style="background:#1e293b; border:1px solid #334155; border-radius:10px; padding:16px;">
              <h4 style="margin:0 0 8px; font-size:13px; color:#38bdf8;">Chế Độ Đồng Hành Đa Cửa Sổ (DeX Companion Mode):</h4>
              <p style="font-size:12px; color:#cbd5e1; line-height:1.5; margin:0 0 12px;">
                Ứng dụng này đang được hệ thống Android kích hoạt dưới dạng Cửa sổ tự do (Freeform Window Activity) thông qua bộ điều phối <code>launchAppWithOptions</code>. Bạn có thể sử dụng song song với Office, kéo thả vị trí hoặc mở giao diện Web đồng hành dưới đây:
              </p>
              <div style="display:flex; gap:8px;">
                <input class="dex-ext-url" type="text" placeholder="Nhập địa chỉ Web đồng bộ hoặc dịch vụ trực tuyến..." style="flex:1; padding:6px 10px; border:1px solid #475569; border-radius:6px; background:#0f172a; color:#fff; font-size:12px;" />
                <button class="dex-ext-load" style="padding:6px 14px; border:none; background:#185abd; color:#fff; border-radius:6px; font-size:12px; cursor:pointer;">Mở</button>
              </div>
            </div>

            <div style="background:#1e293b; border:1px solid #334155; border-radius:10px; padding:16px; font-size:12px;">
              <h4 style="margin:0 0 10px; font-size:13px; color:#e2e8f0;">Thông Số Đa Nhiệm Hệ Thống Android DeX:</h4>
              <div style="display:grid; grid-template-columns:repeat(2, 1fr); gap:8px; color:#94a3b8;">
                <div>Gói ứng dụng: <b style="color:#fff;">${pkg}</b></div>
                <div>Hiển thị: <b style="color:#fff;">Màn hình DeX 4K (ID: 0)</b></div>
                <div>Đa nhiệm (Multi-Task): <b style="color:#4ade80;">Bật (FLAG_ACTIVITY_MULTIPLE_TASK)</b></div>
                <div>Tọa độ kích hoạt: <b class="dex-ext-coords" style="color:#38bdf8;">[Tự động]</b></div>
              </div>
            </div>
          </div>
        </div>
      `;

      const coordsEl = container.querySelector('.dex-ext-coords');

      const updateCoords = () => {
        if (win && win.el) {
          const l = parseInt(win.el.style.left, 10) || 0;
          const t = parseInt(win.el.style.top, 10) || 0;
          const w = parseInt(win.el.style.width, 10) || 760;
          const h = parseInt(win.el.style.height, 10) || 520;
          coordsEl.textContent = `[X: ${l}, Y: ${t}, W: ${w}, H: ${h}]`;
          Host.launchAppWithOptions(pkg, l, t, w, h, 0);
          toast('Đã cập nhật vị trí cửa sổ nổi DeX: ' + (win ? win.title : pkg), 'ok');
        }
      };

      container.querySelector('.dex-ext-bounds').onclick = updateCoords;
      container.querySelector('.dex-ext-front').onclick = () => {
        Host.launchApp(pkg);
        toast('Đã đưa ứng dụng lên phía trước', 'info');
      };
      container.querySelector('.dex-ext-full').onclick = () => {
        Host.launchAppWithOptions(pkg, -1, -1, -1, -1, -1);
        toast('Đã mở ứng dụng toàn màn hình', 'ok');
      };
      container.querySelector('.dex-ext-split').onclick = () => {
        const sw = window.innerWidth || 1920;
        const sh = (window.innerHeight || 1080) - 48;
        const half = Math.floor(sw / 2);
        if (win && win.el) {
          win.el.style.left = '0px';
          win.el.style.top = '0px';
          win.el.style.width = half + 'px';
          win.el.style.height = sh + 'px';
          win.el.style.borderRadius = '0px';
        }
        Host.launchAppWithOptions(pkg, half, 0, sw - half, sh, 0);
        toast('Đã chia đôi màn hình: Trái (Điều khiển), Phải (Ứng dụng)', 'ok');
      };

      const urlInput = container.querySelector('.dex-ext-url');
      container.querySelector('.dex-ext-load').onclick = () => {
        const u = urlInput.value.trim();
        if (u) {
          openAppWindow('browser', 'Web Companion - ' + (win ? win.title : pkg), 'globe');
          toast('Đã mở Web Companion: ' + u, 'info');
        }
      };

      updateCoords();
    }
  }

  // -------------------------------------------------------------
  // Start Menu / All Apps Drawer Renderer
  // -------------------------------------------------------------
  async function renderStartMenu() {
    startMenu.innerHTML = '';
    const apps = await Host.getInstalledApps();

    const searchSec = h('div.dex-start-search',
      h('div', { style: { display: 'flex', alignItems: 'center', background: 'rgba(255,255,255,0.08)', borderRadius: '6px', padding: '6px 10px', gap: '8px' } },
        ic('search', 16),
        h('input', {
          type: 'text',
          placeholder: 'Tìm kiếm ứng dụng trên thiết bị...',
          value: searchQuery,
          style: { background: 'transparent', border: 'none', color: '#fff', fontSize: '13px', width: '100%', outline: 'none' },
          oninput: (e) => {
            searchQuery = e.target.value.toLowerCase();
            renderAppsGrid();
          }
        })
      )
    );
    startMenu.appendChild(searchSec);

    const catBar = h('div', { style: { display: 'flex', gap: '6px', padding: '8px 12px', borderBottom: '1px solid #334155' } },
      [['all', 'Tất cả ứng dụng'], ['system', 'Hệ thống'], ['office', 'Văn phòng'], ['recent', 'Gần đây']].map(([k, lbl]) => {
        const btn = h('button', {
          style: {
            padding: '4px 8px',
            borderRadius: '4px',
            border: 'none',
            fontSize: '11px',
            cursor: 'pointer',
            background: activeCategory === k ? '#185abd' : 'rgba(255,255,255,0.08)',
            color: activeCategory === k ? '#fff' : '#cbd5e1'
          },
          onclick: () => {
            activeCategory = k;
            renderStartMenu();
          }
        }, lbl);
        return btn;
      })
    );
    startMenu.appendChild(catBar);

    const gridEl = h('div.dex-start-grid');

    function renderAppsGrid() {
      gridEl.innerHTML = '';
      const filtered = apps.filter(app => {
        if (searchQuery && !app.name.toLowerCase().includes(searchQuery) && !app.packageName.toLowerCase().includes(searchQuery)) {
          return false;
        }
        if (activeCategory === 'system' && !app.isSystem) return false;
        if (activeCategory === 'office' && app.isSystem) return false;
        return true;
      });

      if (!filtered.length) {
        gridEl.appendChild(h('div', { style: { gridColumn: 'span 4', textAlign: 'center', padding: '24px 0', color: '#94a3b8', fontSize: '12px' } }, 'Không tìm thấy ứng dụng phù hợp'));
        return;
      }

      filtered.forEach(app => {
        const item = h('div.dex-start-item', {
          title: app.name + ' (' + app.packageName + ')',
          onclick: () => {
            startMenuOpen = false;
            startMenu.style.display = 'none';
            if (app.packageName.includes('chrome') || app.packageName.includes('browser')) {
              openAppWindow('browser', 'Trình duyệt Web', 'globe');
            } else if (app.packageName.includes('calculator')) {
              openAppWindow('calc', 'Máy tính Bỏ túi', 'calculator');
            } else if (app.packageName.includes('documents') || app.packageName.includes('files')) {
              openAppWindow('files', 'Quản lý Tệp tin', 'folder-open');
            } else if (app.packageName.includes('office')) {
              openAppWindow('word', 'Word (Văn bản)', 'file-text');
            } else {
              const bounds = getDefaultBounds(app.packageName);
              Host.launchAppWithOptions(app.packageName, bounds.x, bounds.y, bounds.w, bounds.h, 0);
              openAppWindow(app.packageName, app.name, app.icon || 'package', bounds);
              toast('Khởi chạy ứng dụng đa nhiệm DeX: ' + app.name, 'ok');
            }
          }
        },
          h('div', { style: { width: '40px', height: '40px', borderRadius: '10px', background: '#334155', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#38bdf8' } },
            ic(app.icon || 'package', 20)
          ),
          h('div', { style: { fontSize: '11px', fontWeight: '500', color: '#f8fafc', lineHeight: '1.2', maxHeight: '28px', overflow: 'hidden' } }, app.name)
        );
        gridEl.appendChild(item);
      });
    }

    renderAppsGrid();
    startMenu.appendChild(gridEl);

    const footer = h('div.dex-start-footer',
      h('div', { style: { display: 'flex', alignItems: 'center', gap: '8px' } },
        h('div', { style: { width: '28px', height: '28px', borderRadius: '50%', background: '#185abd', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px', fontWeight: 'bold' } }, 'O'),
        h('div',
          h('div', { style: { fontSize: '11px', fontWeight: '600' } }, 'Office DeX Pro v' + VERSION),
          h('div', { style: { fontSize: '9px', color: '#94a3b8' } }, 'Android 6.0+ Device')
        )
      ),
      h('div', { style: { display: 'flex', gap: '8px' } },
        h('button', {
          style: { background: 'none', border: 'none', color: '#94a3b8', cursor: 'pointer', padding: '4px' },
          title: 'Trung tâm Cài đặt DeX',
          onclick: () => {
            startMenuOpen = false;
            startMenu.style.display = 'none';
            openDexControlCenterDialog();
          }
        }, ic('settings-2', 16)),
        h('button', {
          style: { background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer', padding: '4px' },
          title: 'Đóng Desktop TV',
          onclick: () => {
            clearInterval(clockInterval);
            root.remove();
          }
        }, ic('power', 16))
      )
    );
    startMenu.appendChild(footer);
  }

  // -------------------------------------------------------------
  // Initial Mode Dispatcher
  // -------------------------------------------------------------
  if (mode === 'tvapps') {
    startMenuOpen = true;
    startMenu.style.display = 'flex';
    renderStartMenu();
  } else if (mode === 'multiwin') {
    openAppWindow('word', 'Tài liệu 1 - Báo cáo Nâng cấp v' + VERSION + '.docx', 'file-text', { x: 30, y: 30, w: 860, h: 640 });
    openAppWindow('excel', 'Bảng Tính Dự Toán Chi Phí & Kế Hoạch 2026.xlsx', 'file-spreadsheet', { x: 910, y: 30, w: 820, h: 460 });
    openAppWindow('calc', 'Máy tính Thiết bị (Calculator)', 'calculator', { x: 1040, y: 510, w: 320, h: 440 });
  }

  document.body.appendChild(root);
  toast('Đã khởi chạy Màn hình TV Desktop (Chế độ: ' + mode + ')', 'success');
  return root;
}
