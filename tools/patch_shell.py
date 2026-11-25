import sys
p=sys.argv[1]; s=open(p,encoding='utf-8').read()
def rep(a,b,cnt=1):
    global s
    assert a in s, a[:80]
    s=s.replace(a,b,cnt)
rep("import Host from './host.js';","import Host from './host.js';\nimport { bsPage as ubPage, bsAction as ubAction, optionsDialog, SPEC as USPEC, friendlyDate } from './bstage.js';")
rep("    const nav = h('div.bs-nav', { style: { background: doc.meta.color } });\n    const pages = [\n      ['home', 'Trang chủ', 'house'], ['new', 'Mới', 'file-plus'], ['open', 'Mở', 'folder-open'], null,\n      ['info', 'Thông tin', 'info'], ['save', 'Lưu', 'save'], ['saveas', 'Lưu Như', 'save-all'], ['print', 'In', 'printer'], ['share', 'Chia sẻ', 'share-2'], ['export', 'Xuất', 'file-output'], ['close', 'Đóng', 'x'], null,\n      ['account', 'Tài khoản', 'user'], ['options', 'Tùy chọn', 'settings']\n    ];\n    nav.appendChild(h('button.bs-back', { title: 'Quay lại (Esc)', onclick: close }, ic('circle-arrow-left', 30)));",
    "    const nav = h('div.bs-nav.ub-nav');   // v0.1.7: same navigation as the native editor (ui_settings.json)\n    const pages = USPEC.nav;\n    nav.appendChild(h('button.bs-back', { title: 'Quay lại (Esc)', onclick: close }, ic('n:back', 32)));")
rep("    pages.forEach(p => {\n      if (!p) { nav.appendChild(h('div.bs-sep')); return; }\n      const b = h('button.bs-item', ic(p[2], 18), h('span', p[1]));",
    "    let bottom = false;\n    pages.forEach((p, i) => {\n      if (!p) { if (i > 4) { bottom = true; nav.appendChild(h('div.bs-fill')); } else nav.appendChild(h('div.bs-sep')); return; }\n      const b = h('button.bs-item' + (bottom ? '.bottom' : ''), ic('n:' + p[2], 16), h('span', p[1]));")
rep("      if (id === 'close') { close(); return this.closeDoc(doc); }","      if (id === 'close') { close(); return this.closeDoc(doc); }\n      if (id === 'options') { Object.keys(btns).forEach(k => btns[k].classList.toggle('on', k === cur)); go(cur); return optionsDialog(doc, this); }\n      cur = id;")
rep("    const go = async (id) => {","    let cur = 'info';\n    const go = async (id) => {")
rep("      main.appendChild(await this.bsPage(doc, id, close));","      main.appendChild(await ubPage(this, doc, id, close));")
rep("  async bsPage(doc, id, close) {","  bsAction(doc, id) { return ubAction(this, doc, id); }\n  async bsPageLegacy(doc, id, close) {")
rep("    go(page || 'info');","    go(page === 'options' ? 'info' : (page || 'home'));\n    if (page === 'options') optionsDialog(doc, this);")
rep("  async save(doc, saveAs, ext) {","  async save(doc, saveAs, ext, toPath) {")
rep("      const r = await Host.saveFile(name, bytes, { path: !saveAs && doc.fileExt === ext ? doc.path : null, saveAs: !!saveAs || !doc.path || doc.fileExt !== ext });",
    "      const r = toPath ? await Host.saveFile(name, bytes, { path: toPath, saveAs: false }) : await Host.saveFile(name, bytes, { path: !saveAs && doc.fileExt === ext ? doc.path : null, saveAs: !!saveAs || !doc.path || doc.fileExt !== ext });")
rep("onclick: () => dialog({ title: 'Tùy chọn', width: 640, body: optionsPage(null, this), buttons: [{ label: 'Đóng', primary: true }] }) }","onclick: () => optionsDialog(null, this) }")
rep("h('span.c-date', relTime(r.ts))","h('span.c-date', friendlyDate(r.ts))")
open(p,'w',encoding='utf-8').write(s)
print('patched')
