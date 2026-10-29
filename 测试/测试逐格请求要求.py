# coding: utf-8
"""发给模型的本次要求须与代码一致，并保持原标准及旧证据可核验。"""
from copy import deepcopy
from pathlib import Path
import unittest
from unittest.mock import patch

from 附注更新.设置 import load_settings
from 附注更新.统一语义 import read_standard_v2,required_cell_dimensions
from 附注更新.统一勾稽 import load_fact_mapping
from 附注更新.A语义阅读 import binding_payload,focused_review_payload,review_material


class 逐格请求要求测试(unittest.TestCase):
    def setUp(self):
        self.standard=read_standard_v2(load_settings()['gold_path'])
        self.table={'id':'本表','sheet':'资料','range':'A1:B3','subject':'项目','meaning':'各行已有业务含义','data_ranges':['B2:B3']}
        self.snapshot={'path':'原件.xlsx','sha256':'a'*64,'semantic_slot_bindings':True,
            'sheets':[{'name':'资料','merges':[],'cells':{a:{'row':r,'column':c,'value':v} for a,r,c,v in
                [('A1',1,1,'预付款项'),('B1',1,2,'期末'),('A2',2,1,'甲项目'),('B2',2,2,None),('A3',3,1,'乙项目'),('B3',3,2,None)]}}]}

    def material(self,ids):
        return binding_payload([self.table],{'tables':[{'id':'本表','metric_ids':ids}]},self.snapshot,[],self.standard)[0]

    def test_逐格要求不以报告口径设门槛且原标准不变(self):
        key='metric.receivable.loss_allowance.coverage_ratio';original=deepcopy(self.standard['metric'][key])
        payload=self.material([key]);rules=payload['cell_requirements']
        self.assertEqual(rules['complete_fact_dimensions'][key],sorted(required_cell_dimensions(original,self.standard)))
        self.assertNotIn('report_scope',rules['complete_fact_dimensions'][key])
        self.assertFalse(rules['report_scope']['required_for_cell'])
        self.assertFalse(rules['report_scope']['used_for_matching'])
        self.assertEqual(rules['ratio_definitions'][key],original['definition']['ratio'])
        self.assertEqual(self.standard['metric'][key],original)
        self.assertTrue(next(r for r in payload['standards']['metrics'][0]['dimension_refs'] if r['id']=='report_scope')['required'])

    def test_局部核查的要求与保留的指标同步(self):
        a='metric.prepayment.gross_balance';b='metric.prepayment.gross_balance.share_ratio'
        payload=self.material([a,b]);payload['recognized']={'tables':[{'id':'本表','metrics':[
            {'id':a,'ranges':['B2']},{'id':b,'ranges':['B3']}],'dimensions':[]}]}
        result=focused_review_payload(payload,{'本表':['B3']},self.standard)
        self.assertEqual(set(result['cell_requirements']['complete_fact_dimensions']),{b})
        self.assertEqual(set(result['cell_requirements']['ratio_definitions']),{b})

    def test_无效维度诊断指出具体字段和所给值(self):
        key='metric.prepayment.gross_balance';payload=self.material([key])
        binding={'tables':[{'id':'本表','metrics':[{'id':key,'ranges':['B2:B3']}],
            'dimensions':[{'id':'currency','value':'人民币','ranges':['*'],'evidence':['A1']}],'unresolved':[]}]}
        result=review_material(payload,binding,[self.table],self.snapshot,[],self.standard)
        issues=[r for r in result['program_issues'] if r.get('dimension')=='currency']
        self.assertTrue(issues)
        self.assertEqual(issues[0]['provided'],'人民币')
        self.assertIn('币种',issues[0]['reason'])

    def test_按标准比例关系补充候选但不能自动绑定(self):
        ids=['metric.receivable.gross_balance','metric.receivable.loss_allowance.balance']
        binding={'tables':[{'id':'本表','metrics':[{'id':key,'ranges':[cell]} for key,cell in zip(ids,['B2','B3'])],
                           'dimensions':[],'unresolved':[]}]}
        before=deepcopy(binding)
        result=review_material(self.material(ids),binding,[self.table],self.snapshot,[],self.standard)
        key='metric.receivable.loss_allowance.coverage_ratio'
        self.assertIn(key,result['review_candidates'])
        self.assertIn(key,[m['id'] for m in result['standards']['metrics']])
        self.assertEqual(binding,before)
        self.assertEqual(result['recognized'],before)
        focused=focused_review_payload(result,{'本表':['B3']},self.standard)
        self.assertIn(key,[m['id'] for m in focused['standards']['metrics']])

    def test_正式回读拒绝伪造本次逐格要求(self):
        from 测试.测试未决语义自动纠正 import 未决语义纠正测试
        from 附注更新.A语义阅读 import _stage_receipt
        engine,result,calls=未决语义纠正测试().run_case(True)
        def tampered(ref,mapping,task):
            request,answer=_stage_receipt(ref,mapping,task)
            if task=='a_semantic_bind':
                request=deepcopy(request);request['payload']['cell_requirements']['report_scope']['required_for_cell']=True
            return request,answer
        with patch('附注更新.A语义阅读._stage_receipt',side_effect=tampered),self.assertRaises(ValueError):
            load_fact_mapping(result['mapping_path'],engine.standard)
