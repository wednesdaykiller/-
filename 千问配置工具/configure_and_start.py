"""Interactive local setup; keys stay in memory and inherited process environment only."""
import base64,getpass,json,os,sys
from pathlib import Path

def main():
 app=Path(sys.argv[1]);sys.path.insert(0,str(app/'vendor'))
 print('先在原来的应用终端按 Ctrl+C 停止旧服务。')
 print('本次会发送一张程序生成的数字图片，测试视觉识别及 JSON 接口；不发送你的研报。调用可能产生少量服务费用。')
 print('请从服务商控制台复制对应地区/业务空间的兼容接口地址和视觉模型名称。')
 base='https://ws-g6d9ih787hhw1uz8.cn-beijing.maas.aliyuncs.com/compatible-mode/v1'
 print('已填入你提供的北京业务空间接口地址。')
 model=input('视觉模型名称（直接按回车使用 qwen3-vl-plus）：').strip() or 'qwen3-vl-plus'
 key=getpass.getpass('API Key（输入不显示，不保存到文件）：').strip()
 if not base or not model or not key:raise ValueError('三个配置项都不能为空')
 os.environ.update(LLM_BASE_URL=base,LLM_MODEL=model,LLM_API_KEY=key)
 from llm_assist import request_json,_client
 import pymupdf
 doc=pymupdf.open();page=doc.new_page(width=300,height=100);page.insert_text((30,65),'123.45',fontsize=38)
 data=page.get_pixmap(matrix=pymupdf.Matrix(2,2)).tobytes('png');doc.close()
 print('正在测试连接和图片识别…')
 try:
  obj,_=request_json([{'role':'system','content':'Read the number in the image. Return only JSON object {"number": numeric_value}.'},{'role':'user','content':[{'type':'image_url','image_url':{'url':'data:image/png;base64,'+base64.b64encode(data).decode()}}]}],_client())
 except Exception as exc:
  status=getattr(exc,'status_code',None)
  tips={401:'密钥无效，或密钥与接口地区不匹配。',403:'账号或业务空间没有对应模型调用权限。',404:'接口地址或模型名称不存在。',429:'额度、余额或限流问题，请查看控制台。',400:'请求参数、视觉输入或 JSON 输出不受支持，请核对模型与接口。'}
  print('连接测试未通过：'+tips.get(status,'请检查网络、模型配置和依赖。')+' 错误类型：'+type(exc).__name__)
  return 1
 if type(obj.get('number')) not in (int,float) or obj['number']!=123.45:
  print('接口已响应，但图像/JSON测试未通过，暂不启动带模型的应用。');return 1
 print('基础连接、图像输入和JSON响应测试通过。这不代表研报核查准确率已验证。')
 print('即将启动应用，请在页面勾选“启用国产视觉模型辅助”，再上传两份PDF。')
 if not os.environ.get('OCR_DEPS_PATH'):
  deps=app.parent/'.ocr_deps'
  if deps.is_dir():os.environ['OCR_DEPS_PATH']=str(deps)
 os.chdir(app)
 os.execv(sys.executable,[sys.executable,str(app/'app.py')])
if __name__=='__main__':
 try:sys.exit(main())
 except (KeyboardInterrupt,EOFError):print('\n已取消，未保存密钥。');sys.exit(1)
 except Exception as exc:print('未完成：'+(str(exc) if isinstance(exc,ValueError) else type(exc).__name__));sys.exit(1)
