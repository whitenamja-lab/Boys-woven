"""Read-only catalog adapters; publish immutable dated exports and a download index."""
import argparse, datetime, hashlib, html, io, json, re, shutil, zipfile
from pathlib import Path
from urllib.request import urlopen
from PIL import Image as PIL
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
ROOT=Path(__file__).resolve().parents[1]
LEGACY={'76':'bigboys','32-205':'biggirls','33':'toddlerboys','31':'toddlergirls'}
NEW=['18','46','38','328']
UUID=re.compile(r'^[a-fA-F0-9-]{36}$')
CACHE={}
OFFLINE=False

def photo(uid):
    if not uid or not UUID.fullmatch(str(uid)) or OFFLINE: return None
    if uid not in CACHE:
        try:
            with urlopen(f'https://target.scene7.com/is/image/Target/GUEST_{uid}?wid=180&hei=180&fmt=jpeg',timeout=8) as r: data=r.read(2000000)
            CACHE[uid]=thumb(data)
        except Exception: CACHE[uid]=None
    return CACHE[uid]

def thumb(data):
    im=PIL.open(io.BytesIO(data)).convert('RGB'); im.thumbnail((120,120)); out=io.BytesIO(); im.save(out,'JPEG',quality=80); return out.getvalue()

def legacy(dept,folder):
    files=sorted((ROOT/'reports'/folder).glob('*.xlsx'),key=lambda p:p.name)
    if not files: raise ValueError(f'Missing legacy Excel {dept}')
    p=files[-1]; date=re.search(r'\d{4}-\d{2}-\d{2}',p.name).group(); rows=[]
    w=load_workbook(p)
    for s in w:
        pics={}
        for im in s._images:
            if im.anchor._from.row==3: pics[im.anchor._from.col+1]=thumb(im._data())
        for c in range(2,s.max_column+1):
            v=lambda r:s.cell(r,c).value
            if not v(3): continue
            aid=re.search(r'A-(\d+)',str(v(8)))
            colors=str(v(7) or '').split('\n',1)
            names=colors[1].split(' / ') if len(colors)>1 else []
            rows.append(dict(aid=aid.group(1) if aid else '',dpci=v(10),name=v(3),price=v(5),material=v(6),type=s.title,colors=[{'n':n} for n in names if n and '단일' not in n],photo=pics.get(c),status='ACTIVE',date=date))
    # D76 has verified color identifiers, but only merge matching active AIDs.
    cat=ROOT/'reports/tcom/catalog.json'
    if dept=='76' and cat.exists():
        lookup={str(p['aid']):p for p in json.loads(cat.read_text())['items'] if p.get('last')==date}
        for r in rows:
            match=lookup.get(r['aid'])
            if match: r.update({k:match[k] for k in ['colors','uuid','fit','fabric','nsize','sku'] if k in match})
    return date,rows,[p]

def newer(dept):
    folder=ROOT/'reports'/f'dept-{dept}'; cat=folder/'catalog.json'
    if cat.exists():
        d=json.loads(cat.read_text()); date=d['generated']; rows=d['items']; sources=[cat]
    else:
        p=ROOT/'new4_baseline_raw_2026-11-20_23.zip'; sources=[p]
        with zipfile.ZipFile(p) as z:
            date=json.loads(z.read('BASELINE_MANIFEST.json'))['departments'][dept]
            prefix=f'dept-{dept}/'; data=json.loads(z.read(prefix+f'pdp_baseline_{date}.json')); prices=json.loads(z.read(prefix+f'grid_prices_{date}.json'))
            rows=[dict(r,price=prices.get(a,r.get('price')),status='ACTIVE') for a,r in data.items() if str(r.get('dpci','')).startswith(dept.zfill(3)+'-') and not r.get('bot')]
    for r in rows:
        r['date']=date; r['uuid']=r.get('uuid') or r.get('img'); r['type']=r.get('type') or r.get('item_type'); r['nsize']=r.get('nsize',r.get('sizes')); r['sku']=r.get('sku',r.get('skus'))
    if not rows: raise ValueError(f'Empty D{dept}')
    return date,rows,sources

def colors(r):
    cs=r.get('colors') or []
    for c in cs:
        if isinstance(c,dict): yield c.get('n',''),c.get('u',''),c.get('rgb','')
        elif isinstance(c,list): yield (c+['','','',''])[0],(c+['','','',''])[1],(c+['','','',''])[3]
        else: yield str(c),'',''
    if not cs: yield '(single or unspecified)','',''

def sheet(w,title,headers,rows):
    s=w.create_sheet(title); s.append(headers); s.freeze_panes='C2'; s.sheet_view.showGridLines=False
    for c in s[1]: c.fill=PatternFill('solid',fgColor='174E45'); c.font=Font(color='FFFFFF',bold=True)
    for i,h in enumerate(headers,1): s.column_dimensions[get_column_letter(i)].width= {'Image':19,'Product':55,'Material':40,'URL':48,'Image status':25}.get(h,18)
    for vals,pic,rgb in rows:
        s.append(vals); n=s.max_row; s.row_dimensions[n].height=94 if pic else 32
        for c in s[n]:
            if c.data_type=='f': c.data_type='s'
            c.alignment=Alignment(vertical='center',wrap_text=True)
        if pic:
            im=Image(io.BytesIO(pic)); s.add_image(im,f'A{n}')
        elif rgb and re.fullmatch(r'\d{1,3},\d{1,3},\d{1,3}',str(rgb)):
            nums=list(map(int,rgb.split(',')))
            if max(nums)<=255: s.cell(n,1).fill=PatternFill('solid',fgColor=''.join(f'{x:02X}' for x in nums))
        for c in s[n]:
            if isinstance(c.value,str) and c.value.startswith('https://www.target.com/'): c.hyperlink=c.value; c.style='Hyperlink'
    s.auto_filter.ref=s.dimensions
    return s

HEAD=['Image','Department','AID','DPCI','Product','Price','Material','Fit','Fabric','Type','Sizes count','SKU count','Status','Source date','URL','Image status']
def base(d,r): return [d,str(r.get('aid','')),r.get('dpci'),html.unescape(str(r.get('name',''))),r.get('price'),r.get('material'),r.get('fit'),r.get('fabric'),r.get('type'),r.get('nsize'),r.get('sku'),r.get('status'),r['date'],'https://www.target.com/p/-/A-'+str(r.get('aid',''))]
def product_rows(d,rs):
    for r in rs:
        pic=r.get('photo') or photo(r.get('uuid')); yield [None]+base(d,r)+['Embedded' if pic else 'Unavailable'],pic,None

def color_rows(d,rs):
    for r in rs:
        for name,uid,rgb in colors(r):
            pic=photo(uid); yield [None]+base(d,r)+['Embedded' if pic else 'Swatch only' if rgb else 'Unavailable',name],pic,rgb

def workbook(path,sets):
    w=Workbook(); s=w.active; s.title='Summary'; s.append(['Department','Source date','Active','Other','Color rows']);
    for d,(date,rs,_) in sets.items(): s.append([d,date,sum(r.get('status')=='ACTIVE' for r in rs),sum(r.get('status')!='ACTIVE' for r in rs),sum(len(list(colors(r))) for r in rs)])
    s.append(['Note','Color photos only when a matching image identifier exists; no substitute photo.']); s.append(['Note','Dates are source dates, not a synchronized collection.'])
    for c in s[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='174E45')
    for col in 'ABCDE': s.column_dimensions[col].width=24
    s.column_dimensions['B'].width=85
    for d,(_,rs,_) in sets.items(): sheet(w,'D'+d,HEAD,product_rows(d,rs))
    sheet(w,'All_Colorways',HEAD+['Color'],(row for d,(_,rs,_) in sets.items() for row in color_rows(d,rs)))
    w.save(path)

def buyer_workbook(path, dept, date, rows):
    from openpyxl.styles import Border, Side
    from collections import defaultdict
    w=Workbook(); w.remove(w.active)
    groups=defaultdict(list)
    for r in rows:
        groups[str(r.get('type') or 'Other')].append(r)
    labels=['Department '+dept, 'Source date', 'Description', 'Photo', 'Retail Price', 'F. Contents', 'Colorways', 'Link', 'Model Shot', 'DPCI', 'Dept', 'Swatch', 'Fit', 'Fabric', 'Sizes count', 'SKU count', 'Status']
    heights=[18,18,45,149.25,22,42,60,32,40,18,18,48,32,24,18,18,20]
    edge=Side(style='thin',color='B7B7B7')
    for typ, items in sorted(groups.items()):
        title=re.sub(r'[\\/*?:\[\]]','-',typ)[:31] or 'Other'
        sh=w.create_sheet(title); sh.freeze_panes='B4'; sh.sheet_view.showGridLines=False
        sh.column_dimensions['A'].width=18
        for i,(label,height) in enumerate(zip(labels,heights),1):
            sh.cell(i,1,label); sh.row_dimensions[i].height=height
        for c,r in enumerate(sorted(items,key=lambda x:str(x.get('name','')).lower()),2):
            letter=get_column_letter(c); sh.column_dimensions[letter].width=27
            cs=list(colors(r)); names=[str(n) for n,u,rgb in cs]
            vals={1:typ,2:date,3:html.unescape(str(r.get('name',''))),5:r.get('price'),6:r.get('material'),7:str(len(cs))+' colors\n'+' / '.join(names),8:'https://www.target.com/p/-/A-'+str(r.get('aid','')),9:'Not separately collected',10:r.get('dpci'),11:dept.zfill(3),13:r.get('fit'),14:r.get('fabric'),15:r.get('nsize'),16:r.get('sku'),17:r.get('status')}
            for row,v in vals.items():
                cell=sh.cell(row,c,v)
                if cell.data_type=='f': cell.data_type='s'
            sh.cell(8,c).hyperlink=vals[8]
            pic=r.get('photo') or photo(r.get('uuid'))
            if pic:
                im=Image(io.BytesIO(pic)); im.width=im.height=180; sh.add_image(im,letter+'4')
            else: sh.cell(4,c,'Image unavailable')
            # A compact strip uses only actual color-specific images or known RGB swatches.
            from PIL import ImageDraw
            strip=PIL.new('RGB',(192,60),'white'); draw=ImageDraw.Draw(strip); shown=0
            for j,(n,uid,rgb) in enumerate(cs[:6]):
                data=photo(uid); x=j*32
                if data:
                    im=PIL.open(io.BytesIO(data)).convert('RGB'); im.thumbnail((30,56)); strip.paste(im,(x,0)); shown+=1
                elif re.fullmatch(r'\d{1,3},\d{1,3},\d{1,3}',str(rgb)) and max(map(int,rgb.split(',')))<=255:
                    draw.rectangle((x,8,x+29,40),fill=tuple(map(int,rgb.split(',')))); shown+=1
            if shown:
                buf=io.BytesIO(); strip.save(buf,'PNG'); buf.seek(0); sh.add_image(Image(buf),letter+'12')
            else: sh.cell(12,c,'See All_Colorways')
        for row in sh.iter_rows(min_row=1,max_row=len(labels),max_col=len(items)+1):
            for cell in row:
                cell.border=Border(left=edge,right=edge,top=edge,bottom=edge)
                cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
                cell.font=Font(name='Arial',size=10,bold=cell.column==1 or cell.row==1,color='0563C1' if cell.row==8 and cell.column>1 else '000000')
                if cell.column==1 or cell.row==1: cell.fill=PatternFill('solid',fgColor='E7E6E6')
    sheet(w,'All_Colorways',HEAD+['Color'],color_rows(dept,rows))
    w.save(path)

def main():
    global OFFLINE
    a=argparse.ArgumentParser(); a.add_argument('--offline',action='store_true'); args=a.parse_args(); OFFLINE=args.offline
    sets={d:legacy(d,f) for d,f in LEGACY.items()}; sets.update({d:newer(d) for d in NEW})
    digest=hashlib.sha256(b"buyer-layout-v1")
    for d,(date,rs,src) in sets.items():
        digest.update(d.encode()); digest.update(date.encode())
        for p in src: digest.update(p.read_bytes())
        print(d,date,len(rs))
    out=ROOT/'exports'; key=max(v[0] for v in sets.values())+'-'+digest.hexdigest()[:10]; dest=out/key
    if not dest.exists():
        temp=out/('.tmp-'+key); temp.mkdir(parents=True,exist_ok=True)
        workbook(temp/'All_Departments.xlsx',sets)
        for d in sets:
            if d in NEW: buyer_workbook(temp/f'D{d}.xlsx',d,sets[d][0],sets[d][1])
            else: workbook(temp/f'D{d}.xlsx',{d:sets[d]})
        (temp/'manifest.json').write_text(json.dumps({'source_dates':{d:v[0] for d,v in sets.items()},'offline':OFFLINE},indent=2))
        temp.rename(dest)
    rounds=sorted((p for p in out.iterdir() if p.is_dir() and (p/'manifest.json').exists()),reverse=True)
    links=lambda p:' · '.join(f'<a href="{p.name}/{f.name}">{f.stem}</a>' for f in sorted(p.glob('*.xlsx')))
    (out/'index.html').write_text('<!doctype html><meta charset="utf-8"><title>Catalog Excel</title><style>body{font:16px system-ui;max-width:1000px;margin:40px auto;padding:20px}section{padding:18px;border-bottom:1px solid #ddd}a{color:#174e45}</style><a href="../">Dashboard</a><h1>Catalog Excel downloads</h1><h2>Latest export</h2>'+links(dest)+'<h2>Archive</h2>'+''.join(f'<section><b>{p.name}</b><p>{links(p)}</p></section>' for p in rounds),encoding='utf-8')
    (out/'latest.json').write_text(json.dumps({'path':key+'/All_Departments.xlsx'}))
    p=ROOT/'index.html'; text=p.read_text(); marker='<!-- catalog-excel-export -->'
    if marker not in text:
        block=marker+'<div style="padding:12px;margin:8px 0;background:#e8f4ed;border-radius:8px"><a id="catalog-excel-latest" href="exports/">📊 전체 8개 Department Excel 다운로드</a> · <a href="exports/">부서별 / 과거 회차</a></div><script>fetch("exports/latest.json").then(r=>r.json()).then(d=>document.getElementById("catalog-excel-latest").href="exports/"+d.path).catch(()=>{});</script>'
        anchor='<div class="sec s-track">'; assert anchor in text; p.write_text(text.replace(anchor,anchor+block,1),encoding='utf-8')
if __name__=='__main__': main()
