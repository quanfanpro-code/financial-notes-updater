# coding: utf-8
"""未确定实例的槽位按目标列明，不阻断其他来源明确的目标。"""
import copy
import unittest
from 测试 import 测试统一勾稽 as 夹具
from 附注更新 import 统一勾稽 as api


class 取数部分维度测试(unittest.TestCase):
    def test_取数优先已有逐格候选但候选遗漏不遮蔽确定来源(self):
        case=夹具.MatchTests();case.setUp()
        a=夹具.mapping('甲表!D8');b=夹具.mapping('来源!J3',900.5)
        other=copy.deepcopy(b['facts'][0]);other['id']='来源!J4';other['source_reference']['location']='来源!J4';other['raw_value']=1200.5
        b['facts'].append(other)
        b['recognition_scope']={'details':[{'targets':[{'id':'甲表!D8','candidates':['来源!J4']}]}]}
        result=api.build_fact_update_plan(a,b,case.standard)
        self.assertEqual([r['value'] for r in result['updates']],[1200.5])
        b['recognition_scope']['details'][0]['targets'][0]['candidates']=[]
        result=api.build_fact_update_plan(a,b,case.standard)
        self.assertEqual(result['issues'][0]['kind'],'conflicting_sources')
        b['facts']=b['facts'][:1]
        self.assertEqual(api.build_fact_update_plan(a,b,case.standard)['updates'][0]['value'],900.5)

    def test_缺少成员的来源不阻断明确来源也不冒充具体成员(self):
        case=夹具.MatchTests();case.setUp()
        a=夹具.mapping('甲表!D8');b=夹具.mapping('来源!J3',900.5)
        known=copy.deepcopy(a['facts'][0]);known['dimensions'].pop('asset_selection');known['dimension_evidence'].pop('asset_selection')
        b.update(slot_binding_version=1,semantic_reading={'synthetic':True},slot_bindings=[{'sheet':'来源','cell':'J9','partial_semantics':{
            **{k:known[k] for k in ('metric_id','dimensions','dimension_evidence')},'missing_dimensions':['asset_selection'],'evidence_cells':['A1']}}])
        result=api.build_fact_update_plan(a,b,case.standard)
        self.assertEqual(result['updates'][0]['value'],900.5)
        self.assertEqual(result['issues'],[])
        b['facts']=[]
        result=api.build_fact_update_plan(a,b,case.standard)
        self.assertEqual(result['updates'],[])
        self.assertEqual(result['issues'][0]['kind'],'source_dimensions_unresolved')
        self.assertEqual(result['issues'][0]['source_locations'],['来源!J9'])


if __name__=='__main__':unittest.main()
