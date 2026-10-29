# coding: utf-8
"""无效指标编码须在正式绑定时纠正，不能作为成功身份流入后续。"""
from copy import deepcopy
import json
from pathlib import Path
import unittest
import uuid
from unittest.mock import patch

from 附注更新.A语义阅读 import validate_bindings,validate_review
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.设置 import load_settings


class A回执格式纠正测试(unittest.TestCase):
    def setUp(self):
        self.table={'id':'任意表号','sheet':'表','range':'A1:B2','data_ranges':['B2']}
        self.snapshot={'path':'原件.xlsx','sha256':'a'*64,'sheets':[{'name':'表','merges':[],
            'cells':{a:{'row':r,'column':c,'value':v} for a,r,c,v in
                     [('A1',1,1,'项目'),('B1',1,2,'期末'),('A2',2,1,'待查项目'),('B2',2,2,None)]}}]}
        self.bad={'tables':[{'id':'任意表号','metrics':[{'id':'任意表号','ranges':['B2']}],
                            'dimensions':[],'unresolved':[]}]}
        self.fixed={'tables':[{'id':'任意表号','metrics':[],'dimensions':[],
                              'unresolved':[{'range':'B2','reason':'尚未找到对应标准，保留待查'}]}]}
        self.directory=Path(__file__).resolve().parents[1]/'测试结果'/('A格式纠正_'+uuid.uuid4().hex)
        self.directory.mkdir()
        self.engine=UnifiedSemanticEngine({'base_url':'codex://local','model':'gpt-6-luna','model_reasoning_effort':'high'},
                                          load_settings()['gold_path'],self.directory,lambda _:None,None)
        self.engine.source_hash='a'*64

    def check(self,answer):
        return validate_bindings(answer,[self.table],self.snapshot,[],self.engine.standard,candidate=True,require_identity=True)

    def test_表号不能作为指标而明确未决可保留(self):
        with self.assertRaisesRegex(ValueError,'指标'):
            self.check(self.bad)
        self.assertEqual(self.check(self.fixed),self.fixed)

    def test_正式绑定仅允许一次带错误证据的格式纠正(self):
        calls=[]
        def respond(system,payload):
            calls.append(deepcopy(payload))
            return deepcopy(self.bad if len(calls)==1 else self.fixed)
        with patch.object(self.engine,'_http',side_effect=respond):
            result=self.engine._request('只按原文绑定，不存在的指标不能填写',{'task':'a_semantic_bind'},self.check)
        self.assertEqual(result,self.fixed)
        self.assertEqual(len(calls),2)
        self.assertEqual(calls[1]['correction']['previous_result'],self.bad)
        self.assertIn('指标',calls[1]['correction']['validation_error'])
        receipt=json.loads(self.engine._review_receipts[-1].read_text(encoding='utf-8-sig'))
        self.assertTrue(receipt['correction_id'])
        self.assertTrue(list(self.directory.rglob('失败_*.json')))

    def test_两次无效即停止且不产生成功回执(self):
        with patch.object(self.engine,'_http',return_value=self.bad) as request:
            with self.assertRaises(ValueError):
                self.engine._request('只按原文绑定',{'task':'a_semantic_bind'},self.check)
        self.assertEqual(request.call_count,2)
        self.assertFalse(self.engine._review_receipts)
        self.assertEqual(len(list(self.directory.rglob('失败_*.json'))),2)

    def test_核查排除声明须明确范围和类别且只纠正一次(self):
        bad={'tables':[{'id':'任意表号','corrections':{'non_business':[
            {'ranges':['A1:A2'],'reason':'行项目标签','evidence':['A1']}]},'issues':[]}]}
        fixed=deepcopy(bad)
        fixed['tables'][0]['corrections']['non_business']=[
            {'range':'A1:A2','category':'label','reason':'行项目标签','evidence':['A1']}]
        def check(answer):return validate_review(answer,[self.table],self.snapshot,[],require_identity=True)
        with self.assertRaisesRegex(ValueError,'non_business'):check(bad)
        with patch.object(self.engine,'_http',side_effect=[bad,fixed]) as request:
            result=self.engine._request('只核查原结果并返回完整协议',{'task':'a_semantic_review'},check)
        self.assertEqual(result,fixed)
        self.assertEqual(request.call_count,2)
        self.assertIn('non_business',request.call_args_list[1].args[1]['correction']['validation_error'])

    def test_无更正声明不能掩盖原身份的无效维度(self):
        from 附注更新.A语义阅读 import validate_review_result,standard_material
        table={**self.table,'meaning':'余额'}
        metric='metric.inventory.gross_balance'
        binding={'tables':[{'id':table['id'],'metrics':[{'id':metric,'ranges':['B2']}],
            'dimensions':[{'id':'currency','ranges':['B2'],'value':'人民币','evidence':['A1']}],'unresolved':[]}]}
        payload={'task':'a_semantic_review','reading':[table],'standards':standard_material([metric],self.engine.standard)}
        bad={'tables':[{'id':table['id'],'corrections':{},'issues':[]}]}
        fixed=deepcopy(bad)
        fixed['tables'][0]['corrections']={'patches':[{'id':'currency','ranges':['B2'],'value':'CNY','evidence':['A1']}]}
        def check(answer):return validate_review_result(answer,binding,[table],self.snapshot,[],self.engine.standard,payload)
        with self.assertRaisesRegex(ValueError,'B2'):check(bad)
        with patch.object(self.engine,'_http',side_effect=[bad,fixed]) as request:
            result=self.engine._request('按完整定义核查，不允许保留无效格式却声明通过',payload,check)
        self.assertEqual(result,fixed)
        self.assertEqual(request.call_count,2)
        self.assertIn('B2',request.call_args_list[1].args[1]['correction']['validation_error'])
