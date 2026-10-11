"""Local-only upload UI and asynchronous orchestration."""
import argparse,json,os,re,threading,uuid,webbrowser
from concurrent.futures import ThreadPoolExecutor
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from engine import ROOT,pipeline,review,save_result
JOBS={};LOCK=threading.RLock();POOL=ThreadPoolExecutor(max_workers=1);RUNS=ROOT/'runs'
MAX_BYTES=128*1024*1024

def model_ready():return all(os.getenv(k,'').strip() for k in ('LLM_API_KEY','LLM_BASE_URL','LLM_MODEL'))
def upload_parts(content_type,body):
 if not content_type.startswith('multipart/form-data;'):raise ValueError('请上传两份PDF文件')
 msg=BytesParser(policy=default).parsebytes(('Content-Type: '+content_type+'\r\nMIME-Version: 1.0\r\n\r\n').encode()+body)
 values={}
 for part in msg.iter_parts():
  name=part.get_param('name',header='content-disposition')
  if name not in ('annual','research','use_model') or name in values:raise ValueError('上传字段无效或重复')
  values[name]=part.get_payload(decode=True)
 for name in ('annual','research'):
  b=values.get(name,b'')
  if not b or b'%PDF-' not in b[:1024]:raise ValueError('财报和研报都必须是PDF文件')
 return values

def start_job(parts):
 use_model=parts.get('use_model')==b'true'
 if use_model and not model_ready():raise ValueError('尚未配置模型接口，请先关闭模型辅助或在启动环境配置接口')
 with LOCK:
  if any(v['state'] in ('queued','running') for v in JOBS.values()):raise ValueError('已有任务正在运行，请完成后再提交')
  jid=uuid.uuid4().hex;folder=RUNS/jid;folder.mkdir(parents=True)
  for name in ('annual','research'):(folder/(name+'.pdf')).write_bytes(parts[name])
  job={'id':jid,'state':'queued','stage':'排队','detail':'等待开始','history':[]};JOBS[jid]=job
 def progress(stage,detail):
  with LOCK:job.update(stage=stage,detail=detail);job['history'].append({'stage':stage,'detail':detail})
 def run():
  with LOCK:job['state']='running'
  try:
   result=pipeline(folder/'annual.pdf',folder/'research.pdf',folder,progress,use_model,os.getenv('OCR_DEPS_PATH'))
   with LOCK:job.update(state='done',stage='模型核验未完成' if result.get('model_summary',{}).get('failed_batches') else '核查完成',detail='模型调用失败，已停止后续请求；不需要把接口故障逐项人工复核。' if result.get('model_summary',{}).get('failed_batches') else '请查看覆盖范围和待确认项',result=result)
  except Exception as exc:
   msg=str(exc) if isinstance(exc,ValueError) else '处理失败：'+type(exc).__name__+'。请检查PDF和OCR依赖。'
   with LOCK:job.update(state='failed',stage='未完成',detail=msg)
  (folder/'job.json').write_text(json.dumps({k:v for k,v in job.items() if k!='result'},ensure_ascii=False,indent=2))
 POOL.submit(run);return jid

class Handler(BaseHTTPRequestHandler):
 def log_message(self,*args):pass
 def respond(self,status,payload,ctype='application/json; charset=utf-8',download=None):
  body=json.dumps(payload,ensure_ascii=False).encode() if isinstance(payload,(dict,list)) else payload
  self.send_response(status);self.send_header('Content-Type',ctype);self.send_header('Content-Length',str(len(body)));self.send_header('X-Content-Type-Options','nosniff');self.send_header('Cache-Control','no-store')
  if download:self.send_header('Content-Disposition','attachment; filename="'+download+'"')
  self.end_headers();self.wfile.write(body)
 def do_GET(self):
  path=urlparse(self.path).path
  if path=='/':return self.respond(200,(ROOT/'static/index.html').read_bytes(),'text/html; charset=utf-8')
  if path=='/api/config':return self.respond(200,{'model_ready':model_ready(),'model':os.getenv('LLM_MODEL','') if model_ready() else None})
  if path.startswith('/static/'):
   name=path.removeprefix('/static/')
   if name not in ('app.js','style.css'):return self.respond(404,{'error':'不存在'})
   return self.respond(200,(ROOT/'static'/name).read_bytes(),'application/javascript; charset=utf-8' if name.endswith('.js') else 'text/css; charset=utf-8')
  m=re.fullmatch(r'/api/jobs/([a-f0-9]{32})(?:/(result|report|annual|research))?',path)
  if not m:return self.respond(404,{'error':'不存在'})
  jid,kind=m.groups()
  with LOCK:job=JOBS.get(jid)
  if not job:return self.respond(404,{'error':'任务不存在；重启后可在runs目录查看已保存报告'})
  if not kind:
   with LOCK:return self.respond(200,{k:v for k,v in job.items() if k!='result'})
  files={'annual':('annual.pdf','application/pdf'),'research':('research.pdf','application/pdf'),'result':('result.json','application/json; charset=utf-8'),'report':('report.html','text/html; charset=utf-8')}
  name,ctype=files[kind];f=RUNS/jid/name
  if not f.exists():return self.respond(409,{'error':'结果尚未生成'})
  return self.respond(200,f.read_bytes(),ctype,name if kind in ('report',) else None)
 def do_POST(self):
  try:
   origin=self.headers.get('Origin');host=self.headers.get('Host')
   if self.headers.get('X-Requested-With')!='ReportAgent' or (origin and origin!='http://'+str(host)):return self.respond(403,{'error':'请从本地应用页面提交'})
   size=int(self.headers.get('Content-Length','0'))
   if size<=0 or size>MAX_BYTES:return self.respond(413,{'error':'两份文件合计需小于128MB'})
   body=self.rfile.read(size);path=urlparse(self.path).path
   if path=='/api/jobs':return self.respond(202,{'id':start_job(upload_parts(self.headers.get('Content-Type',''),body))})
   m=re.fullmatch(r'/api/jobs/([a-f0-9]{32})/review',path)
   if m:
    with LOCK:
     job=JOBS.get(m.group(1))
     if not job or job['state']!='done':raise ValueError('任务尚未完成')
     req=json.loads(body);review(job['result'],req['id'],req['number'],req['unit'],req['reviewer'],req['note']);save_result(job['result'],RUNS/job['id'])
     return self.respond(200,{'ok':True})
   return self.respond(404,{'error':'不存在'})
  except (ValueError,KeyError,TypeError) as exc:return self.respond(400,{'error':str(exc) if isinstance(exc,ValueError) else '请求格式错误'})
  except Exception:return self.respond(500,{'error':'处理失败，请保留文件并重试'})

def main():
 p=argparse.ArgumentParser();p.add_argument('--port',type=int,default=8765);p.add_argument('--no-browser',action='store_true');args=p.parse_args()
 server=ThreadingHTTPServer(('127.0.0.1',args.port),Handler)
 print(f'打开 http://127.0.0.1:{args.port}；按 Ctrl+C 停止。')
 if not args.no_browser:threading.Timer(.5,lambda:webbrowser.open(f'http://127.0.0.1:{args.port}')).start()
 try:server.serve_forever()
 except KeyboardInterrupt:server.server_close()
if __name__=='__main__':main()
