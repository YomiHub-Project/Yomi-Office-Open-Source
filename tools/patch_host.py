import sys
p=sys.argv[1]; s=open(p,encoding='utf-8').read()
a="  ping() {"
assert a in s
s=s.replace(a,"""  // v0.1.7: Backstage "Lưu thành" folder browser + "Mở vị trí tệp" (same as the native editor; Windows only)
  async listDir(dir) { if (this.kind !== 'windows') return null; try { return await api('/api/list_dir', { dir: dir || '' }); } catch (e) { return null; } },
  async pickFolder() { if (this.kind !== 'windows') return null; try { const r = await api('/api/pick_folder', {}); return r && r.ok ? r.path : null; } catch (e) { return null; } },
  reveal(path) { if (this.kind === 'windows' && path) api('/api/reveal', { path }).catch(() => {}); },
"""+a,1)
open(p,'w',encoding='utf-8').write(s); print('host ok')
