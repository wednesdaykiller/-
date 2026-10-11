"""Same pipeline as the UI, for reproducibility and automated evaluation."""
import argparse,os,json
from pathlib import Path
from engine import pipeline
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('annual');p.add_argument('research');p.add_argument('--out',required=True);p.add_argument('--model',action='store_true');p.add_argument('--no-cache',action='store_true');args=p.parse_args()
 out=Path(args.out)
 if out.exists():p.error('请指定新输出目录，避免覆盖已有核查记录')
 r=pipeline(args.annual,args.research,out,lambda a,b:print(a+'：'+b,flush=True),args.model,os.getenv('OCR_DEPS_PATH'),not args.no_cache)
 print(json.dumps(r['summary'],ensure_ascii=False))
