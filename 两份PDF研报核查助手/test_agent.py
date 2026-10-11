import copy,io,json,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
import engine,app
class AgentTests(unittest.TestCase):
 def setUp(self):
  self.item={'id':'x','field':'营业收入','period':'2024FY','period_raw':'2024A','number':16831,'unit':'百万元','raw':'16,831','trusted':True,'estimate':False,'context':'metrics','page':2,'recognition':'原生文本'}
  self.fact={'field':'营业收入','period':'2024FY','number':16830541086.13,'unit':'元','page':9,'context':'metrics','label':'营业收入'}
  self.annual={'facts':[self.fact]}
 def test_percentage_precision(self):
  i={**self.item,'field':'毛利率','number':43.2,'unit':'%','raw':'43.2%'}
  f={**self.fact,'field':'毛利率','number':43.247165,'unit':'%'}
  self.assertEqual(engine.compare(i,{'facts':[f]})['status'],'相符')
 def test_growth_alias(self):
  self.assertEqual(engine.canonical('增长率YoY%','营业收入'),'营业收入同比')
  self.assertEqual(engine.canonical('毛利率%'),'毛利率')
 def test_unit_and_rounding(self):self.assertEqual(engine.compare(self.item,self.annual)['status'],'相符')
 def test_ocr_not_silently_passed(self):self.assertEqual(engine.compare({**self.item,'trusted':False},self.annual)['status'],'待确认')
 def test_forecast(self):self.assertEqual(engine.compare({**self.item,'estimate':True},self.annual)['status'],'预测值')
 def test_missing(self):self.assertEqual(engine.compare(self.item,{'facts':[]})['status'],'缺少依据')
 def test_difference(self):self.assertEqual(engine.compare({**self.item,'number':123},self.annual)['status'],'差异待核实')
 def test_conflicting_standard(self):self.assertEqual(engine.compare(self.item,{'facts':[self.fact,{**self.fact,'number':100}]})['status'],'待确认')
 def test_cash_flow_not_income(self):
  i={**self.item,'field':'财务费用','context':'cash_reconcile'};f={**self.fact,'field':'财务费用','context':'income'}
  self.assertEqual(engine.compare(i,{'facts':[f]})['status'],'缺少依据')
 def test_eps_not_generic(self):self.assertEqual(engine.compare({**self.item,'field':'基本每股收益'},self.annual)['status'],'待确认')
 def test_year(self):self.assertEqual(engine.compare({**self.item,'period':'2023FY'},self.annual)['status'],'缺少依据')
 def test_periods(self):
  self.assertEqual(engine.period('2024年12月31日'),'2024FY');self.assertEqual(engine.period('2025H1'),'2025H1');self.assertIsNone(engine.period('2025年3月31日'))
 def test_alias_not_pe_substring(self):self.assertIsNone(engine.canonical('EBITDA'));self.assertEqual(engine.canonical('P/B'),'市净率')
 def test_manual_review_logged(self):
  r={'research_facts':[{**self.item,'trusted':False}],'annual':self.annual,'human_reviews':[]}
  engine.review(r,'x',16831,'百万元','测试复核人','仅验证机制，不作为真实人工决定')
  self.assertEqual(len(r['human_reviews']),1);self.assertEqual(r['results'][0]['status'],'相符')
 def test_review_invalid(self):
  with self.assertRaises(ValueError):engine.review({},'x',float('nan'),'元','a','b')
 def test_upload_validation(self):
  with self.assertRaises(ValueError):app.upload_parts('application/json',b'{}')
 def test_upload_parse(self):
  body=b'--foo\r\nContent-Disposition: form-data; name="annual"; filename="../../bad.pdf"\r\n\r\n%PDF-1.7\r\n--foo\r\nContent-Disposition: form-data; name="research"; filename="r.pdf"\r\n\r\n%PDF-1.7\r\n--foo--\r\n'
  parts=app.upload_parts('multipart/form-data; boundary=foo',body);self.assertEqual(set(parts),{'annual','research'})
 def test_csrf(self):
  handler=object.__new__(app.Handler);handler.headers={'Origin':'http://evil.example','Host':'127.0.0.1:8765','X-Requested-With':'ReportAgent'};got=[];handler.respond=lambda status,*args:got.append(status);handler.do_POST();self.assertEqual(got,[403])
 def test_static_path_traversal(self):
  handler=object.__new__(app.Handler);handler.path='/static/../../engine.py';got=[];handler.respond=lambda status,*args:got.append(status);handler.do_GET();self.assertEqual(got,[404])
class RealSampleTests(unittest.TestCase):
 @classmethod
 def setUpClass(cls):cls.r=json.loads((engine.ROOT/'demo_run/result.json').read_text())
 def test_no_quarter_inherits_annual(self):self.assertFalse(any(f['page']==10 for f in self.r['annual']['facts']))
 def test_actual_annual_values(self):
  fs=self.r['annual']['facts']
  self.assertTrue(any(f['field']=='营业收入' and f['period']=='2024FY' and f['number']==16830541086.13 for f in fs))
  self.assertTrue(any(f['field']=='财务费用' and f['context']=='cash_reconcile' and f['number']==-95614916.8 for f in fs))
 def test_no_ocr_unreviewed_passes(self):self.assertNotIn('相符',self.r['summary']);self.assertEqual(self.r['human_reviews'],[])
 def test_company(self):self.assertEqual(self.r['company'],'公牛集团')
 def test_derived_evidence(self):
  fs=[f for f in self.r['annual']['facts'] if f['field']=='毛利率'];self.assertEqual(len(fs),2);self.assertTrue(all(len(f['inputs'])==2 for f in fs))
if __name__=='__main__':unittest.main()
