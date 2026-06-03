# -*- coding: utf-8 -*-
from __future__ import annotations
import datetime, io, re, zipfile, csv
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable, List

NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
METRICS=['商品访客数','商品浏览量','商品加购人数','商品加购件数','支付买家数','支付件数','支付金额','成功退款金额']


def _load_shared(z: zipfile.ZipFile):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    root=ET.fromstring(z.read('xl/sharedStrings.xml'))
    return [''.join((t.text or '') for t in si.iter(NS+'t')) for si in root.findall(NS+'si')]


def _sheet_paths(z: zipfile.ZipFile):
    names=z.namelist(); out={}
    try:
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rid_to_target={r.attrib['Id']:r.attrib['Target'] for r in rels}
        for s in wb.find(NS+'sheets'):
            name=s.attrib.get('name','')
            rid=s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            target=rid_to_target.get(rid,'')
            path='xl/'+target.lstrip('/'); path=path.replace('xl/xl/','xl/')
            if name in ('天猫数据源','京东抖音数据源') and path in names:
                out[name]=path
    except Exception:
        pass
    if not out:
        if 'xl/worksheets/sheet1.xml' in names: out['天猫数据源']='xl/worksheets/sheet1.xml'
        if 'xl/worksheets/sheet2.xml' in names: out['京东抖音数据源']='xl/worksheets/sheet2.xml'
    return out


def _col_idx(ref: str) -> int:
    m=re.match(r'([A-Z]+)', ref or '')
    if not m: return 0
    n=0
    for ch in m.group(1): n=n*26+ord(ch)-64
    return n-1


def _cell_value(c, ss):
    v=c.find(NS+'v')
    if v is None: return None
    txt=v.text
    if c.attrib.get('t')=='s':
        try: return ss[int(txt)]
        except Exception: return txt
    return txt


def _iter_rows(z: zipfile.ZipFile, sheet_path: str, ss):
    with z.open(sheet_path) as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if elem.tag==NS+'row':
                row={}
                for c in elem.findall(NS+'c'):
                    row[_col_idx(c.attrib.get('r',''))]=_cell_value(c, ss)
                yield row
                elem.clear()


def _num(v) -> float:
    if v is None or v=='' or v=='-': return 0.0
    s=str(v).replace(',','').replace('￥','').replace('¥','').strip()
    if s.endswith('%'):
        try: return float(s[:-1])/100
        except Exception: return 0.0
    try: return float(s)
    except Exception: return 0.0


def _norm(v) -> str:
    s='' if v is None else str(v).strip()
    return s if s and s!='-' else '未标注'


def _date_str(v):
    if v is None or v=='': return None
    s=str(v).strip().replace('/','-')
    try:
        x=float(s)
        if 30000<x<65000:
            return (datetime.datetime(1899,12,30)+datetime.timedelta(days=x)).date().isoformat()
    except Exception:
        pass
    m=re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        y,mo,d=map(int,m.groups())
        try: return datetime.date(y,mo,d).isoformat()
        except Exception: return None
    return None


def _empty(): return defaultdict(float)

def _add(agg, key, vals):
    d=agg[key]
    for m in METRICS: d[m]+=vals.get(m,0.0)


def enrich(r: Dict[str,Any]):
    r['客单价']=round(r.get('支付金额',0)/r.get('支付买家数',0),2) if r.get('支付买家数',0) else 0
    r['支付转化率']=round(r.get('支付买家数',0)/r.get('商品访客数',0),4) if r.get('商品访客数',0) else 0
    r['加购率']=round(r.get('商品加购人数',0)/r.get('商品访客数',0),4) if r.get('商品访客数',0) else 0
    r['浏览深度']=round(r.get('商品浏览量',0)/r.get('商品访客数',0),2) if r.get('商品访客数',0) else 0
    r['退款率']=round(r.get('成功退款金额',0)/r.get('支付金额',0),4) if r.get('支付金额',0) else 0
    return r


def _final(agg, keys, sort_key=None, limit=None, desc=False):
    out=[]
    for key, vals in agg.items():
        if not isinstance(key, tuple): key=(key,)
        r={k:v for k,v in zip(keys,key)}
        for m in METRICS: r[m]=round(vals[m],2)
        enrich(r); out.append(r)
    if sort_key: out.sort(key=lambda x:x.get(sort_key,0), reverse=desc)
    if limit: out=out[:limit]
    return out


def _open_zip(src):
    if isinstance(src, (str, Path)):
        return zipfile.ZipFile(src)
    if isinstance(src, bytes):
        return zipfile.ZipFile(io.BytesIO(src))
    try:
        return zipfile.ZipFile(io.BytesIO(src.getvalue()))
    except Exception:
        return zipfile.ZipFile(io.BytesIO(src.read()))


def month_shift(month: str, delta: int) -> str:
    y,m=map(int, month.split('-'))
    m+=delta
    while m<1: y-=1; m+=12
    while m>12: y+=1; m-=12
    return f'{y:04d}-{m:02d}'


def parse_sales_workbook(src) -> Dict[str,Any]:
    by_day=defaultdict(_empty); by_month=defaultdict(_empty); by_channel=defaultdict(_empty); by_store=defaultdict(_empty); by_cat=defaultdict(_empty); by_model=defaultdict(_empty); by_style=defaultdict(_empty); by_product=defaultdict(_empty)
    total=defaultdict(float); sets={k:set() for k in ['channels','stores','categories','models']}
    min_date=max_date=None; rows_count=0; used_sheets=[]
    with _open_zip(src) as z:
        ss=_load_shared(z); paths=_sheet_paths(z)
        if not paths: raise ValueError('未找到数据源工作表。请确认 Excel 中包含「天猫数据源」或「京东抖音数据源」。')
        for sheet_name, sp in paths.items():
            rows=_iter_rows(z, sp, ss)
            try: header=next(rows)
            except StopIteration: continue
            headers={str(v).strip():k for k,v in header.items() if v is not None and str(v).strip()}
            if '统计日期' not in headers: continue
            used_sheets.append(sheet_name)
            def get(row,h):
                i=headers.get(h); return row.get(i) if i is not None else None
            for row in rows:
                ds=_date_str(get(row,'统计日期'))
                if not ds: continue
                rows_count+=1
                cat=_norm(get(row,'品类')); model=_norm(get(row,'型号')); channel=_norm(get(row,'渠道')); store=_norm(get(row,'店铺')); style=_norm(get(row,'款式'))
                product=_norm(get(row,'商品名称')); pid=_norm(get(row,'商品ID'))
                vals={m:_num(get(row,m)) for m in METRICS}
                if '成功退款金额' not in headers: vals['成功退款金额']=0.0
                for m,v in vals.items(): total[m]+=v
                sets['channels'].add(channel); sets['stores'].add(store); sets['categories'].add(cat); sets['models'].add(model)
                min_date=ds if min_date is None or ds<min_date else min_date
                max_date=ds if max_date is None or ds>max_date else max_date
                _add(by_day,(ds,channel,store,cat,model),vals); _add(by_month,(ds[:7],channel,store,cat,model),vals)
                _add(by_channel,channel,vals); _add(by_store,store,vals); _add(by_cat,cat,vals); _add(by_model,(model,channel,cat,store),vals); _add(by_style,(style,channel,cat,model),vals); _add(by_product,(product[:80],pid,channel,cat,model),vals)
    if rows_count == 0:
        raise ValueError('没有解析到有效数据行。请检查表头是否包含「统计日期」以及销售指标字段。\n如果您上传的是目标拆解 Excel（含"X年X月目标拆解及登记"工作表），请切换到左侧【销售目标】数据类型，不要使用【销售数据】上传。')
    tot={m:round(total[m],2) for m in METRICS}; enrich(tot)
    monthly=_final(by_month,['月份','渠道','店铺','品类','型号'],'月份')
    by_month_all=defaultdict(_empty)
    for r in monthly: _add(by_month_all,r['月份'],{m:r[m] for m in METRICS})
    all_months=_final(by_month_all,['月份'],'月份'); all_months.sort(key=lambda r:r['月份'])
    return {
        'meta':{'rows':rows_count,'dateRange':[min_date,max_date],'generatedAt':datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'usedSheets':used_sheets},
        'filters':{k:sorted(v) for k,v in sets.items()}, 'totals':tot, 'daily':_final(by_day,['日期','渠道','店铺','品类','型号'],'日期'), 'monthly':monthly, 'all_months':all_months,
        'channels':_final(by_channel,['渠道'],'支付金额',None,True), 'stores':_final(by_store,['店铺'],'支付金额',None,True), 'categories':_final(by_cat,['品类'],'支付金额',None,True),
        'models':_final(by_model,['型号','渠道','品类','店铺'],'支付金额',600,True), 'styles':_final(by_style,['款式','渠道','品类','型号'],'支付金额',1000,True), 'products':_final(by_product,['商品名称','商品ID','渠道','品类','型号'],'支付金额',1200,True)
    }


def rows_to_csv(rows: List[Dict[str,Any]], cols: List[str]) -> bytes:
    out=io.StringIO(); w=csv.DictWriter(out, fieldnames=cols, extrasaction='ignore'); w.writeheader(); w.writerows(rows)
    return ('\ufeff'+out.getvalue()).encode('utf-8-sig')
