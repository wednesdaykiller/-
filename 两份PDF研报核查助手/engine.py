"""Evidence-first two-PDF pipeline. No external standards or human decisions required."""
import copy,hashlib,json,math,re,sys,time,unicodedata
from pathlib import Path
from decimal import Decimal,ROUND_HALF_UP
ROOT=Path(__file__).parent
sys.path.insert(0,str(ROOT/'vendor'))
import pymupdf

ALIASES={
'营业收入':['营业收入','营业总收入'], '营业成本':['营业成本'],
'归母净利':['归属于上市公司股东的净利润','归属于母公司股东的净利润','归属母公司净利润','归母净利润'],
'扣非归母':['归属于上市公司股东的扣除非经常性损益的净利润','扣非归母净利润'],
'净利润':['净利润'], '总资产':['总资产','资产总计'],
'归母净资产':['归属于上市公司股东的净资产','归属于母公司股东权益合计','归属母公司股东权益'],
'应收账款':['应收账款'],'存货':['存货'],'经营性现金流':['经营活动产生的现金流量净额','经营活动现金流量净额','经营活动现金流'],
'财务费用':['财务费用'],'销售费用':['销售费用'],'管理费用':['管理费用'],'研发费用':['研发费用'],
'基本每股收益':['基本每股收益','EPS','每股收益'],'稀释每股收益':['稀释每股收益'],
'市盈率':['市盈率P/E','市盈率','P/E','PE'],'市净率':['市净率P/B','市净率','P/B','PB'],
'毛利率':['毛利率'],'净利率':['净利率'],'资产负债率':['资产负债率'],
'营业收入同比':['营业收入同比','营业收入同比增速'],'归母净利同比':['归母净利润同比','归母净利同比']}
SCALE={'元':1,'千元':1000,'万元':10000,'百万元':1000000,'亿元':100000000}

def clean(s):return re.sub(r'\s+','',unicodedata.normalize('NFKC',str(s or ''))).replace('−','-')
def num(s):
 s=clean(s).replace('，',',')
 if not re.fullmatch(r'[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?',s):return None
 v=float(s.replace(',',''));return v if math.isfinite(v) else None
def canonical(label,previous=None):
 s=clean(label);s=re.sub(r'^[一二三四五六七八九十0-9]+[、.．]','',s);s=re.sub(r'^其中[:：]','',s)
 base=re.sub(r'\([^)]*\)','',s)
 if base in ['同比','同比%','增长率YoY%','增长率YOY%']:return {'营业收入':'营业收入同比','归母净利':'归母净利同比'}.get(previous)
 base=base.removesuffix('%')
 for field,aliases in ALIASES.items():
  if base in aliases:return field
 return None

def period(s):
 s=clean(s);m=re.search(r'(20\d{2})(?:A|E|年|FY|H[12]|Q[1-4])',s,re.I)
 if not m:return None
 year=m.group(1)
 if '半年' in s or 'H1' in s.upper():return year+'H1'
 if 'H2' in s.upper():return year+'H2'
 q=re.search(r'Q([1-4])',s,re.I)
 if q:return year+'Q'+q.group(1)
 if any(x in s for x in ['季度','月']) and not ('12月31' in s):return None
 return year+'FY'
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def statement(text):
 patterns=[('parent_balance','母公司资产负债表'),('parent_income','母公司利润表'),('parent_cash','母公司现金流量表'),('cash_reconcile','现金流量表补充资料'),('balance','合并资产负债表'),('income','合并利润表'),('cash','合并现金流量表'),('metrics','主要会计数据'),('metrics','主要财务指标')]
 found=[(text.rfind(word),kind) for kind,word in patterns if word in text]
 return max(found)[1] if found else None

def identity(text):
 s=clean(text)
 m=re.search(r'([\u4e00-\u9fffA-Za-z]{2,30})股份有限公司20\d{2}年',s)
 name=m.group(1) if m else None
 m=re.search(r'(?:股票|证券)代码[:：]?(\d{6})',s)
 return {'name':name,'code':m.group(1) if m else None}

def extract_annual(path,progress):
 facts=[];pages=[];warnings=[];last=None;context=None;lastpage=-2
 with pymupdf.open(path) as doc:
  if doc.is_encrypted:raise ValueError('财报PDF已加密，请上传可读取的版本')
  if len(doc)>600:raise ValueError('本版支持600页以内财报')
  text0=''.join(p.get_text() for p in list(doc)[:5]);issuer=identity(text0)
  ym=re.search(r'(20\d{2})\s*年\s*(年度|半年度)',text0)
  year=int(ym.group(1)) if ym else None
  for n,page in enumerate(doc):
   text=page.get_text();pages.append({'page':n+1,'chars':len(text.strip())})
   if len(text.strip())<20:
    warnings.append({'page':n+1,'reason':'财报页面缺少文本层，本版不自动把OCR财报当作标准依据'});last=None;continue
   if n%20==0:progress('财报依据提取',f'正在检查财报第 {n+1}/{len(doc)} 页')
   # Financial tables only; no model invented standard values.
   if not any(a in clean(text) for aliases in ALIASES.values() for a in aliases):last=None;continue
   try:tables=page.find_tables().tables
   except Exception: warnings.append({'page':n+1,'reason':'表格解析失败'});last=None;continue
   for ti,t in enumerate(tables):
    matrix=t.extract();prefix=page.get_text(clip=pymupdf.Rect(0,0,page.rect.width,t.bbox[1]))
    ctx=statement(prefix) or (context if n-lastpage<=1 else None);context=ctx
    headers=None;unit=None
    units=re.findall(r'单位\s*[:：]\s*(亿元|百万元|万元|千元|元)',prefix)
    if units:unit=units[-1]
    if ctx in ('income','balance','cash','cash_reconcile','parent_income','parent_balance','parent_cash') and n-lastpage<=1 and last and len(matrix[0])==last['cols'] and ctx==last['context']:
     headers=last['headers'];unit=unit or last['unit']
    for ri,row in enumerate(matrix):
     hp={j:period(v) for j,v in enumerate(row) if j>0 and period(v)}
     # Only a header row, never an arbitrary row containing a year in its label.
     if len(hp)>=2 and len(set(hp.values()))>=2:headers=hp;continue
     if not headers and year and ym and ym.group(2)=='年度':
      current=[j for j,v in enumerate(row) if clean(v) in ('本期金额','本期发生额')]
      previous=[j for j,v in enumerate(row) if clean(v) in ('上期金额','上期发生额')]
      if current and previous:headers={current[0]:f'{year}FY',previous[0]:f'{year-1}FY'};continue
     if not headers or not row or not row[0]:continue
     label=clean(row[0]);field=canonical(label)
     if not field or not ctx or ctx.startswith('parent'):continue
     u='元/股' if '每股' in field else '%' if field.endswith('率') else unit
     if not u:continue
     # Facts from annual summary and consolidated statements only; notes use explicit matching context.
     for j,per in headers.items():
      if j>=len(row):continue
      value=num(row[j])
      if value is None:continue
      facts.append({'field':field,'period':per,'number':value,'unit':u,'page':n+1,'row':ri+1,'label':label,'raw':str(row[j]),'context':ctx,'bbox':list(t.bbox),'method':'native_table'})
    if headers:last={'headers':headers,'unit':unit,'cols':len(matrix[0]),'context':ctx};lastpage=n
   if not tables:last=None
 # Derive only from exact aligned period/source/context, retaining input evidence.
 base=list(facts)
 for per in sorted({f['period'] for f in base}):
  def unique(field,p=per,contexts=('metrics','income')):
   fs=[f for f in base if f['field']==field and f['period']==p and f['context'] in contexts and f['unit'] in SCALE]
   vs={round(f['number']*SCALE[f['unit']],4) for f in fs}
   return fs[0] if len(vs)==1 else None
  rev=unique('营业收入');cost=unique('营业成本',contexts=('income',))
  if rev and cost and rev['number']>0:
   value=(1-cost['number']*SCALE[cost['unit']]/(rev['number']*SCALE[rev['unit']]))*100
   facts.append({**rev,'field':'毛利率','number':value,'unit':'%','label':'合并毛利率','method':'derived','formula':'(营业收入-营业成本)/营业收入×100%','inputs':[rev,cost]})
  if per.endswith('FY'):
   for field in ('营业收入','归母净利'):
    now=unique(field);prev=unique(field,str(int(per[:4])-1)+'FY')
    if now and prev and prev['number']>0:
     facts.append({**now,'field':field+'同比','number':(now['number']*SCALE[now['unit']]/(prev['number']*SCALE[prev['unit']])-1)*100,'unit':'%','method':'derived','formula':'(本期/上期-1)×100%','inputs':[now,prev]})
 return {'issuer':issuer,'facts':facts,'pages':pages,'warnings':warnings,'sha256':sha(path)}

def research_data(path,progress,ocr_path=None,use_cache=True):
 from pdf_extractor import extract_pdf,build_quality_report
 cache=ROOT/'cache';manifest=cache/'manifest.json'
 if use_cache and manifest.exists() and json.loads(manifest.read_text())['sha256']==sha(path):
  progress('研报解析','命中同一PDF的提取缓存，保留原始OCR审计；未载入任何人工决定')
  data=json.loads((cache/'extraction.json').read_text());quality=json.loads((cache/'quality.json').read_text());data['file_path']=str(path)
  return data,quality,True
 if ocr_path and Path(ocr_path).is_dir():sys.path.insert(0,str(ocr_path))
 audit=[];data=extract_pdf(str(path),ocr_audit=audit);return data,build_quality_report(data,audit),False

def research_facts(data,quality):
 scanned={p['page_num'] for p in quality.get('ocr',[])};items=[]
 for t in data['tables']:
  previous=None
  for ri,row in enumerate(t['rows']):
   label=row[t['headers'][0]]['raw'];field=canonical(label,previous)
   if field and not field.endswith('同比'):previous=field
   elif not field:previous=None
   if not field:continue
   for h in t['headers'][1:]:
    per=period(h)
    if not per:continue
    c=row[h];value=c.get('number');value=value if type(value) in (int,float) and math.isfinite(value) else num(c['raw'])
    unit=c.get('unit')
    if not unit:
     m=re.search(r'(亿元|百万元|万元|千元|元/股|倍|%)',clean(label));unit=m.group(1) if m else None
    key=hashlib.sha256(f"{t['table_id']}:{ri}:{h}".encode()).hexdigest()[:20]
    name=clean(t.get('table_name',''))
    ctx='cash_reconcile' if '现金流' in name else 'balance' if '资产负债' in name else 'income' if '利润表' in name else 'metrics'
    items.append({'id':key,'field':field,'period':per,'period_raw':h,'label':label,'number':value,'unit':unit,'raw':c['raw'],'page':t['page_num'],'table_id':t['table_id'],'row':ri+1,'bbox':t['bbox'],'context':ctx,'estimate':bool(re.search(r'E\b',h,re.I)),'trusted':t['page_num'] not in scanned,'recognition':'原生文本' if t['page_num'] not in scanned else 'OCR待复核'})
    # Explicit table headings identify statement scope; field names do not.
    if any(title in name for title in ('现金流','资产负债','利润表')):
     items[-1].update(source_context=ctx,source_table_name=t['table_name'])
 return items

def compare(item,annual):
 r=copy.deepcopy(item);r.update(status='待确认',reason='',evidence=[])
 if item['estimate']:r.update(status='预测值',reason='不与历史实绩直接比较');return r
 if item['field'] in ('基本每股收益','稀释每股收益','市盈率','市净率'):
  r['reason']='需要股本基准、每股口径或估值时点，暂不自动判错';return r
 if item['number'] is None or not item['unit']:r['reason']='缺少可靠数值或单位';return r
 candidates=[f for f in annual['facts'] if f['field']==item['field'] and f['period']==item['period']]
 if item['field']=='财务费用':candidates=[f for f in candidates if f['context']==item['context']]
 else:candidates=[f for f in candidates if f['context'] in ('metrics','income','balance','cash','cash_reconcile')]
 comparable=[]
 for f in candidates:
  if f['unit'] in SCALE and item['unit'] in SCALE:v=f['number']*SCALE[f['unit']]/SCALE[item['unit']]
  elif f['unit']==item['unit']:v=f['number']
  else:continue
  comparable.append((v,f))
 if not comparable:r.update(status='缺少依据',reason='上传财报未提取到同字段、期间和口径的可比依据');return r
 if len({round(v,6) for v,f in comparable})!=1:r['reason']='财报同字段同期间存在不同数值，需核实表格口径';r['evidence']=[f for v,f in comparable];return r
 value,f=comparable[0];r['evidence']=[f];r['standard_in_report_unit']=value
 if not item['trusted']:r['reason']=item.get('verification_reason','已找到财报依据，但扫描识别结果尚未经独立核验');return r
 raw=clean(item['raw']).replace(',','').removesuffix('%');places=len(raw.split('.')[-1]) if '.' in raw and num(raw) is not None else 0
 rounded=float(Decimal(str(value)).quantize(Decimal(1).scaleb(-places),rounding=ROUND_HALF_UP))
 if abs(item['number']-rounded)<1e-8:r.update(status='相符',reason='按研报显示精度换算和舍入后相符')
 else:r.update(status='差异待核实',reason='与财报依据存在数值差异，请核实引用口径；尚未直接认定研报错误',difference=item['number']-value)
 return r

def pipeline(annual_path,research_path,out,progress=lambda stage,detail:None,use_model=False,ocr_path=None,use_cache=True):
 start=time.time();out=Path(out);out.mkdir(exist_ok=True,parents=True)
 progress('检查输入','检查两份PDF及公司信息')
 for p in (annual_path,research_path):
  with pymupdf.open(p) as doc:
   if not doc.is_pdf or not len(doc):raise ValueError('请输入有效PDF')
   if len(doc)>600:raise ValueError('本版每份PDF最多600页')
 progress('财报依据提取','从财报原生表格构建依据，保留页码、单位和统计口径')
 annual=extract_annual(annual_path,progress)
 progress('研报解析','提取文字和表格；扫描页面调用本地OCR')
 data,quality,cached=research_data(research_path,progress,ocr_path,use_cache)
 text=''.join(p['text'] for p in data['pages'][:2]);name=annual['issuer']['name'];code=annual['issuer']['code']
 matched=bool((name and name in clean(text)) or (code and re.search(r'(?<!\d)'+re.escape(code)+r'(?!\d)',text)))
 if not matched:raise ValueError('未能可靠确认两份PDF属于同一公司，请检查文件；本版不会强行继续比对')
 progress('字段归一化','匹配字段、期间、单位，并保留预测与报表口径')
 items=research_facts(data,quality);model_audit=[]
 if use_model:
  progress('模型辅助','独立读取扫描表格；不同读数保留待确认')
  from visual_batches import run
  model_audit=run(items,data,research_path,out,progress,canonical,num,clean)
 progress('证据比对','按同公司、字段、期间及口径核查')
 result={'company':name or code,'annual_sha256':annual['sha256'],'research_sha256':sha(research_path),'annual':annual,'research_facts':items,'results':[compare(i,annual) for i in items],'model_audit':model_audit,'model_requested':use_model,'model_summary':{'completed_batches':sum(a.get('status')=='completed' for a in model_audit),'failed_batches':sum(a.get('status')=='failed' for a in model_audit),'verified_cells':sum(a.get('verified_cells',0) for a in model_audit)},'research_cache_hit':cached,'elapsed_seconds':round(time.time()-start,2),'coverage':{'research_pages':data['total_pages'],'tables':data['total_tables'],'recognized_fields':len(items),'body_automatic_check':False,'note':'本版优先核查表格财务指标，正文陈述未全部覆盖；待确认、预测和缺少依据不计为通过。'},'human_reviews':[]}
 progress('生成报告','汇总结果及证据，生成可下载报告')
 save_result(result,out);return result

def save_result(result,out):
 from collections import Counter
 import html
 out=Path(out);result['summary']=dict(Counter(r['status'] for r in result['results']))
 (out/'result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2))
 e=lambda x:html.escape(str(x))
 rows=''.join('<tr>'+''.join('<td>'+e(x)+'</td>' for x in (r['field'],r['period_raw'],r['number'],r['unit'],r['status'],r['reason'],r['page'],','.join(str(f['page']) for f in r['evidence'])))+'</tr>' for r in result['results'])
 (out/'report.html').write_text('<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>研报核查报告</title><style>body{font:15px/1.6 system-ui;margin:32px;color:#24364a}td,th{border:1px solid #ccd5df;padding:8px}table{border-collapse:collapse;width:100%}</style><h1>'+e(result['company'])+'研报核查报告</h1><p>'+e(result['summary'])+'</p><p>'+e(result['coverage']['note']+(' 模型核验存在失败，请查看model_audit.json；待确认不代表识别错误。' if result.get('model_summary',{}).get('failed_batches') else ''))+'</p><table><tr>'+''.join('<th>'+h+'</th>' for h in ['字段','期间','数值','单位','状态','依据/待确认原因','研报页','财报页'])+'</tr>'+rows+'</table></html>')

def review(result,record_id,value,unit,reviewer,note):
 if not reviewer.strip() or not note.strip():raise ValueError('请填写复核人和核对依据')
 if type(value) not in (int,float) or not math.isfinite(value):raise ValueError('请输入有限数值')
 if unit not in {*SCALE,'%','倍','元/股'}:raise ValueError('单位暂不支持')
 item=next((i for i in result['research_facts'] if i['id']==record_id),None)
 if item is None:raise ValueError('记录不存在')
 old=copy.deepcopy(item);item.update(number=value,unit=unit,trusted=True,recognition='人工确认')
 result['human_reviews'].append({'id':record_id,'reviewer':reviewer,'note':note,'before':old,'after':copy.deepcopy(item),'time':time.time()})
 result['results']=[compare(i,result['annual']) for i in result['research_facts']]
