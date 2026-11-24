#!/usr/bin/env python3
"""Build verified Target department reports from browser-exported JSON, no network calls.

Input: incoming/YYYY-MM-DD/dept-N.json with complete:true, expected_results and
products list. Each product MUST have a PDP-verified dpci. Never infer department
from category URLs. A complete run is required to mark products DROPPED.
"""
import argparse
import datetime as dt
import html
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CFG = json.loads((ROOT / 'config/departments.json').read_text(encoding='utf-8'))
AID = re.compile(r'^\d{6,15}$')
PRICE = re.compile(r'^\$?\s*(\d+(?:\.\d{1,2})?)')

def load(path):
    return json.loads(path.read_text(encoding='utf-8')) if path.exists() else None

def save(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')

def price_number(value):
    m = PRICE.match(str(value))
    if not m: raise ValueError(f'Invalid price: {value!r}')
    return float(m.group(1))

def normalize(p, dept):
    aid = str(p.get('aid', '')).strip()
    dpci = str(p.get('dpci', '')).strip()
    if not AID.fullmatch(aid): raise ValueError(f'Invalid AID: {aid!r}')
    if not re.fullmatch(r'\d{3}-\d{2}-\d{4}', dpci):
        raise ValueError(f'PDP-verified DPCI required for {aid}: {dpci!r}')
    if dpci[:3] != dept: return None
    if not str(p.get('name', '')).strip(): raise ValueError(f'Name missing: {aid}')
    price = str(p.get('price', '')).strip()
    price_number(price)
    return {"aid":aid,"dpci":dpci,"name":str(p['name']).strip(),"price":price,
            "material":str(p.get('material','')),"fit":str(p.get('fit','')),
            "fabric":str(p.get('fabric','')),"type":str(p.get('type','')),
            "uuid":str(p.get('uuid','')),"colors":p.get('colors',[]),
            "ncolor":int(p.get('ncolor',len(p.get('colors',[])))),
            "nsize":int(p.get('nsize',0)),"sku":int(p.get('sku',0))}

def excel_download(dept):
    if str(dept) not in {'18', '46', '38', '328'}:
        raise ValueError(f'Unsupported department: {dept}')
    return """<div style="display:flex;align-items:center;justify-content:space-between;background:#fff;border:1px solid #ddd;border-radius:10px;padding:12px;margin:12px 0;gap:12px;flex-wrap:wrap">
<div><b>📊 카탈로그 엑셀</b><div style="font-size:12px;color:#666">상품 사진 · 상세 정보 · Colorways 포함</div></div>
<a id="dept-excel" href="../../exports/" style="background:#1565c0;color:white;padding:8px 14px;border-radius:8px;text-decoration:none;font-weight:bold">⬇ 내려받기</a></div>
<script>(()=>{const link=document.getElementById('dept-excel');
fetch('../../exports/latest.json',{cache:'no-store'}).then(r=>{if(!r.ok)throw Error('Export unavailable');return r.json();}).then(d=>{
if(typeof d.path!=='string'||!/^\\d{4}-\\d{2}-\\d{2}-[a-f0-9]+\\/All_Departments\\.xlsx$/.test(d.path))throw Error('Invalid export path');
link.href='../../exports/'+d.path.replace('All_Departments.xlsx','D__DEPT__.xlsx');
}).catch(()=>{link.textContent='엑셀 목록 열기';});})();</script>""".replace('__DEPT__', str(dept))

def report_html(name, date, current, events, dept):
    def esc(x): return html.escape(str(x), quote=True)
    rows = []
    for p in sorted(current.values(), key=lambda x:(x['name'].lower(),x['aid'])):
        url = 'https://www.target.com/p/-/A-' + p['aid']
        rows.append('<tr>' + ''.join('<td>'+c+'</td>' for c in [
            esc(p['dpci']),f'<a href="{esc(url)}">{esc(p["name"])}</a>',esc(p['price']),
            esc(p['material']),esc(p['fit']),esc(p['fabric']),esc(p['ncolor'])]) + '</tr>')
    summary = ' · '.join(f'{key}: {len(events[key])}' for key in ['new','dropped','returned','changes'])
    return ('<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
            '<title>'+esc(name)+' '+esc(date)+'</title><style>body{font:16px system-ui;margin:2rem auto;max-width:1200px;padding:0 1rem;background:#f8fafc;color:#172033}table{border-collapse:collapse;width:100%;background:white}td,th{border:1px solid #ddd;padding:9px;text-align:left}th{background:#e8eef8}tr:nth-child(even){background:#f7f9fc}a{color:#1458a7}</style>'
            '<h1>'+esc(name)+'</h1>'+excel_download(dept)+'<p>Verified PDP DPCI · '+esc(date)+'</p><p>Current: '+str(len(current))+' · '+esc(summary)+'</p>'
            '<table><thead><tr><th>DPCI</th><th>Product</th><th>Retail</th><th>Material</th><th>Fit</th><th>Fabric</th><th>Colorways</th></tr></thead><tbody>'
            +''.join(rows)+'</tbody></table></html>')

def process(dept, date, root):
    spec = CFG[dept]
    src = root / 'incoming' / date / f'dept-{dept}.json'
    raw = load(src)
    if raw is None:
        print(f'SKIP D{dept}: missing {src}')
        return False
    if raw.get('complete') is not True: raise ValueError(f'D{dept}: complete:true required')
    if raw.get('date') != date: raise ValueError(f'D{dept}: date mismatch')
    products = raw.get('products')
    if not isinstance(products,list) or not products: raise ValueError(f'D{dept}: empty products')
    expected = raw.get('expected_results')
    if type(expected) is not int or expected != len(products):
        raise ValueError(f'D{dept}: expected_results must match total unique category AIDs (pre DPCI filtering)')
    if len({str(p.get('aid','')) for p in products}) != len(products):
        raise ValueError(f'D{dept}: duplicate AID in input')
    current = {}
    for p in products:
        item = normalize(p,spec['dpci'])
        if item: current[item['aid']] = item
    if not current: raise ValueError(f'D{dept}: no matching PDP DPCI; refuse to publish')
    dest = root / 'reports' / f'dept-{dept}'
    history_file = dest / 'history.json'
    history = load(history_file) or {"items":{},"rounds":[]}
    if date in history['rounds']: raise ValueError(f'D{dept}: date already processed')
    known = history['items']
    prev_active = {aid for aid,p in known.items() if p['status']=='ACTIVE'}
    curr_ids = set(current)
    new = sorted(curr_ids-set(known))
    dropped = sorted(prev_active-curr_ids)
    returned = sorted(aid for aid in curr_ids & set(known) if known[aid]['status']=='DROPPED')
    changes = []
    for aid,p in current.items():
        if aid in known and abs(price_number(known[aid]['price'])-price_number(p['price']))>0.004:
            changes.append({'aid':aid,'old':known[aid]['price'],'new':p['price']})
    for aid in dropped: known[aid]['status']='DROPPED'
    for aid,p in current.items():
        first = known[aid]['first'] if aid in known else date
        known[aid] = {**known.get(aid,{}),**p,'first':first,'last':date,'status':'ACTIVE'}
    history['rounds'].append(date)
    events = {'new':new,'dropped':dropped,'returned':returned,'changes':changes}
    save(history_file,history)
    save(dest/'catalog.json',{'generated':date,'count':len(known),'items':list(known.values())})
    save(dest/f'{date}-diff.json',events)
    page = report_html(spec['name'],date,current,events,dept)
    (dest/f'{date}.html').write_text(page,encoding='utf-8')
    (dest/'index.html').write_text(page,encoding='utf-8')
    print(f'D{dept}: {len(current)} active, {len(new)} new, {len(dropped)} dropped, {len(returned)} returned, {len(changes)} price changes')
    return True

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--date',default=dt.date.today().isoformat())
    parser.add_argument('--root',type=Path,default=ROOT)
    args=parser.parse_args()
    dt.date.fromisoformat(args.date)
    errors=[]
    for dept in CFG:
        try: process(dept,args.date,args.root)
        except (ValueError,TypeError,KeyError) as exc: errors.append(f'D{dept}: {exc}')
    if errors: raise SystemExit('\n'.join(errors))

if __name__=='__main__': main()
