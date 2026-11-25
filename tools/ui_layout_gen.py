import json,re,sys
def norm(s):
    s=(s or '').lower(); s=re.sub(r'\(.*?\)','',s); s=re.sub(r'[…\.:]+','',s); s=s.replace('&','và'); s=re.sub(r'\s+',' ',s).strip(); return s
SYN={ # native normalized -> web normalized
 'chổi định dạng':'sao định dạng','cỡ phông':'cỡ chữ','tăng cỡ phông':'tăng cỡ chữ','giảm cỡ phông':'giảm cỡ chữ','đổi chữ hoa/thường':'đổi kiểu chữ hoa/thường',
 'gạch dưới':'gạch chân','gạch ngang chữ':'gạch ngang','màu tô sáng văn bản':'màu tô sáng','giảm mức thụt lề':'giảm thụt lề','tăng mức thụt lề':'tăng thụt lề',
 'hiện/ẩn dấu đoạn văn và ký hiệu định dạng ẩn':'hiện/ẩn ký hiệu định dạng','giãn cách dòng và đoạn':'giãn dòng và khoảng cách đoạn','màu tô ô':'màu tô',
 'căn giữa theo chiều dọc':'căn giữa','ngắt dòng văn bản':'ngắt dòng','gộp và căn giữa':'gộp và căn giữa','định dạng tiền tệ':'định dạng số kế toán',
 'dấu phân cách hàng nghìn':'kiểu dấu phẩy','trợ lý ai':'copilot','bình luận':'chú thích','trường thời gian':'ngày và giờ','khoảng cách đoạn':'giãn dòng đoạn',
 'tăng khoảng cách':'cách sau +6pt','giảm khoảng cách':'cách trước +6pt','cập nhật bảng':'cập nhật','chèn chú thích cuối trang':'chèn chú thích chân','chèn chú thích cuối tài liệu':'chú thích cuối tài liệu',
 'chèn trường phối':'chèn trường trộn','hoàn tất và phối':'hoàn tất và trộn','chính tả':'chính tả và ngữ pháp','thống kê từ':'đếm số từ và thống kê','bình luận mới':'chú thích mới',
 'xóa tất cả':'xóa','chấp nhận':'chấp nhận tất cả','từ chối':'từ chối tất cả','bố cục in':'bố trí in','bố cục web':'bố trí web','ký tự định dạng':'ký hiệu định dạng',
 'thường':'bình thường','dạng xem đọc':'đọc','bố trí':'bố cục','màu phông chữ':'màu chữ','đường viền hình':'viền hình','kích thước trang chiếu':'cỡ trang chiếu',
 'ngăn hoạt hình':'ngăn hoạt ảnh','từ trang chiếu hiện tại':'từ trang hiện tại','biểu đồ':'biểu đồ đề xuất','tính toán tự động':'tùy chọn tính toán','tính lại toàn bộ':'tính ngay',
 'bộ lọc':'lọc','ghi chú mới':'ghi chú','đậm':'in đậm','nghiêng':'in nghiêng','dấu đầu dòng':'danh sách dấu đầu dòng','đánh số':'danh sách đánh số','căn đều':'căn đều hai bên',
 'ngắt dòng văn bản@bố cục':'ngắt chữ quanh ảnh','từ đầu':'từ đầu','sắp xếp trang chiếu':'sắp xếp trang chiếu','trang ghi chú':'trang ghi chú','tách':'cố định ngăn@x',
}
GSYN={'căn chỉnh':'căn lề','hình minh họa':'hình ảnh','bình luận':'chú thích','bắt đầu phối thư':'bắt đầu trộn thư','ghi và chèn trường':'viết và chèn trường','hoạt hình':'hoạt ảnh','hoạt hình nâng cao':'hoạt ảnh nâng cao','thư':'thư tín','định dạng tài liệu':'định dạng tài liệu','chuyển tiếp sang trang chiếu này':'chuyển tiếp sang trang chiếu này','tùy chỉnh':'tùy chỉnh','dạng xem':'dạng xem sổ làm việc','hiển thị':'hiện','bắt đầu trình chiếu':'bắt đầu trình chiếu'}
TSYN={'thư':'thư tín','hoạt hình':'hoạt ảnh'}
NICON_FOR_WEB={}  # lucide -> native spec name when a native icon of the same meaning exists
def leaves(items):
    for it in items or []:
        if not it: continue
        if it.get('type') in ('col','row') and isinstance(it.get('items'),list): yield from leaves(it['items'])
        else: yield it
def wkey(it): return (it.get('label') or '')+'\u0001'+(it.get('title') or '')
def nkey(it): return (it.get('t') or '')+'\u0001'+(it.get('tip') or '')
W=json.load(open('web_ui.json')); out={}
for kind in ['word','excel','ppt']:
    N=json.load(open(f'native_{kind}.json',encoding='utf-8'))
    wt=[t for t in W[kind]['tabs'] if not t.get('ctx')]
    wtab={norm(t['label']):t for t in wt}
    used=set(); tabs=[]; stat={'both':0,'nat':0,'web':0}
    def find_web(tabw, gname, txt, tip, kindn, nat_items_in_group):
        cands=[]
        if tabw:
            for g in tabw['groups']:
                for it in leaves(g['items']): cands.append((g,it))
        allw=[(g,it) for t in wt for g in t['groups'] for it in leaves(g['items'])]
        keys=[norm(txt),norm(tip)]
        keys2=[]
        for k in keys:
            if not k: continue
            keys2.append(k)
            if k+'@'+norm(gname) in SYN: keys2.append(SYN[k+'@'+norm(gname)])
            if k in SYN: keys2.append(SYN[k])
        for pool in ((cands,) if tabw else (allw,)):
            # same group first
            for g,it in pool:
                if id(it) in used: continue
                wl=[norm(it.get('label')),norm(it.get('title'))]
                if any(k and k in wl for k in keys2): return g,it
        return None,None
    for t in N['tabs']:
        if t['name']=='Tệp': continue
        tw=wtab.get(norm(t['name'])) or wtab.get(TSYN.get(norm(t['name']),'#'))
        groups=[]
        for g in t['groups']:
            cols=[]; wgroups=set()
            for c in g['cols']:
                def conv(it,kind_):
                    if it['k']=='Gallery' and tw:
                        # pair with the web gallery/custom of the matching group
                        for wg in tw['groups']:
                            if norm(wg['label']) in (norm(g['name']),GSYN.get(norm(g['name']),'#')) or not wg['label']:
                                for wi in leaves(wg['items']):
                                    if id(wi) not in used and (wi['type'] in ('gallery','custom')):
                                        used.add(id(wi)); wgroups.add(id(wg)); stat['both']+=1
                                        return dict(t=it['t'] or wi.get('label') or '', tip=it['tip'] or wi.get('title') or '', i='n:'+(g.get('icon') or 'styles'), nat=nkey(it), web=wkey(wi), k='gallery')
                    wg,wi=find_web(tw,g['name'],it['t'],it['tip'],it['k'],None)
                    if wi: used.add(id(wi)); wgroups.add(id(wg)); stat['both']+=1
                    else: stat['nat']+=1
                    icon=('n:'+it['i']) if it.get('i') else (('l:'+wi['icon']) if wi and wi.get('icon') else '')
                    return dict(t=it['t'] or '', tip=it['tip'] or (wi.get('title') if wi else '') or '', i=icon, nat=nkey(it), web=wkey(wi) if wi else None,
                                k={'Large':'large','Small':'small','Mini':'mini','Combo':'combo','Gallery':'gallery'}[it['k']], menu=bool(it.get('menu')), split=it.get('split'), color=it.get('color'))
                if c['type']=='rows': cols.append(dict(type='rows',rows=[[conv(x,'') for x in r] for r in c['rows']]))
                else: cols.append(dict(type=c['type'],items=[conv(x,'') for x in c['items']]))
            # web group with the same name -> its unmatched items join this group
            if tw:
                for wg in tw['groups']:
                    if norm(wg['label']) in (norm(g['name']),GSYN.get(norm(g['name']),'#')): wgroups.add(id(wg))
            groups.append(dict(name=g['name'],icon=g.get('icon'),launcher=g.get('launcher'),nlauncher=g['name'] if g.get('launcher') is not None else None,cols=cols,_w=wgroups))
        tabs.append(dict(name=t['name'],groups=groups,_tw=tw))
    # web-only items / groups / tabs
    def webcols(items):
        cols=[]; stack=[]; icons=[]
        def flush_icons():
            nonlocal icons
            if icons:
                rows=[icons[i:i+4] for i in range(0,len(icons),4)]
                for i in range(0,len(rows),3): cols.append(dict(type='rows',rows=rows[i:i+3]))
                icons=[]
        def flush_stack():
            nonlocal stack
            if stack: cols.append(dict(type='stack',items=stack)); stack=[]
        for it in items:
            if id(it) in used: continue
            used.add(id(it)); stat['web']+=1
            e=dict(t=it.get('label') or '',tip=it.get('title') or '',i=('l:'+it['icon']) if it.get('icon') else '',nat=None,web=wkey(it),
                   k='large' if it.get('big') else ('combo' if it['type']=='combo' else 'gallery' if it['type'] in ('gallery','custom') else 'small'),menu=bool(it.get('menu')) or it['type'] in ('drop','split','color'),split=it['type'] in ('split','color'),color=it['type']=='color')
            if e['k'] in ('large','gallery'):
                flush_icons(); flush_stack()
                cols.append(dict(type='large' if e['k']=='large' else 'gallery',items=[e]))
            elif not e['t'] and e['k']!='combo':
                flush_stack(); e['k']='mini'; icons.append(e)
            else:
                flush_icons(); stack.append(e)
                if len(stack)==3: flush_stack()
        flush_icons(); flush_stack()
        return cols
    for wi,t in enumerate(wt):
        ut=next((u for u in tabs if u['_tw'] is t),None)
        if ut is None:
            ut=dict(name=t['label'],groups=[],_tw=t)
            prev=wt[wi-1] if wi>0 else None
            pos=next((i+1 for i,u in enumerate(tabs) if prev is not None and u['_tw'] is prev),len(tabs))
            tabs.insert(pos,ut)
        last=None
        for wg in t['groups']:
            ug=next((g for g in ut['groups'] if id(wg) in g['_w']),None)
            cols=webcols(list(leaves(wg['items'])))
            if ug is not None:
                ug['cols']+=cols
                if wg.get('launcher') and not ug.get('launcher') is not None: ug['wlauncher']=wg['label']
                if wg.get('launcher'): ug['wlauncher']=wg['label']
                last=ug
            elif cols:
                ng=dict(name=wg['label'],icon=None,launcher=None,wlauncher=wg['label'] if wg.get('launcher') else None,cols=cols,_w={id(wg)})
                pos=ut['groups'].index(last)+1 if last in ut['groups'] else len(ut['groups'])
                ut['groups'].insert(pos,ng); last=ng
    # help tab last
    for u in tabs:
        u.pop('_tw',None)
        for g in u['groups']: g.pop('_w',None)
    h=[u for u in tabs if u['name']=='Trợ giúp']
    tabs=[u for u in tabs if u['name']!='Trợ giúp']+h
    # ids
    n=0
    for u in tabs:
        for g in u['groups']:
            for c in g['cols']:
                for it in (c.get('items') or [x for r in c.get('rows',[]) for x in r]):
                    n+=1; it['id']=f'{kind}.{n}'
    out[kind]=dict(tabs=tabs,status=N['status'],views=N['views'])
    print(kind,stat,[u['name'] for u in tabs])
out['icons']=N['icons']; out['palette']=N['palette']
import json as _j
src=open(sys.argv[1] if len(sys.argv)>1 else 'icons.gen.js',encoding='utf-8').read()
LUC=_j.loads(src[src.index('{'):src.rindex('}')+1])
used=set()
def walk(o):
    if isinstance(o,dict):
        i=o.get('i')
        if isinstance(i,str) and i.startswith('l:'): used.add(i[2:])
        for v in o.values(): walk(v)
    elif isinstance(o,list):
        for v in o: walk(v)
walk({k:out[k] for k in ['word','excel','ppt']})
out['lucide']={k:LUC[k] for k in sorted(used) if k in LUC}
print('lucide used',len(used),'missing',[k for k in used if k not in LUC])
json.dump(out,open('ui_layout.json','w',encoding='utf-8'),ensure_ascii=False,indent=0)
