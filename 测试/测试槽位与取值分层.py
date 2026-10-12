# coding: utf-8
"""槽位确认不依赖具体成员名称；实例取值及不确定状态仍完整保留。"""
import copy
import unittest
from unittest.mock import patch
from openpyxl import load_workbook
from 测试.测试统一语义 import records, fact
from 测试 import 测试B范围定位 as 范围定位
from 附注更新 import 统一语义 as semantics
from 附注更新.统一识别 import UnifiedSemanticEngine, apply_identity_review
from 附注更新.统一勾稽 import load_fact_mapping


class 槽位与取值分层测试(unittest.TestCase):
    def test_槽位不随成员取值改变且不混同不同计量(self):
        standard = semantics.validate_records(records())
        first, second = fact(), fact()
        second['dimensions']['asset_selection'] = {
            'domain':'固定资产', 'mode':'members', 'members':['合成设备乙'], 'completeness':'complete'}
        self.assertEqual(semantics.slot_key(first['metric_id'], standard),
                         semantics.slot_key(second['metric_id'], standard))
        self.assertNotEqual(semantics.semantic_key(first, standard), semantics.semantic_key(second, standard))
        changed = copy.deepcopy(standard)
        changed['metric']['固定资产原值']['definition']['measurement_basis'] = '累计折旧'
        self.assertNotEqual(semantics.slot_key(first['metric_id'], standard),
                            semantics.slot_key(first['metric_id'], changed))

    def test_仅经原文核查通过的已知槽位才可采纳(self):
        standard = semantics.validate_records(records()); known = fact()
        dims = copy.deepcopy(known['dimensions']); dims.pop('asset_selection')
        row = {'kind':'partial_metric_fact', 'reason':'名称尚未填写', 'partial_semantics':{
            'metric_id':known['metric_id'], 'dimensions':dims,
            'dimension_evidence':{k:['A1'] for k in dims}, 'missing_dimensions':['asset_selection'],
            'evidence_cells':['A2']}}
        review = {'checked_cells':['B3'], 'issues':[]}
        approved = apply_identity_review({'B3':row}, review, standard, confirm_slots=True)
        self.assertEqual(len(approved['slot_bindings']), 1)
        self.assertEqual(approved['facts'], [])
        self.assertEqual(approved['unresolved'], [])
        self.assertEqual(approved['slot_bindings'][0]['partial_semantics'], row['partial_semantics'])
        for rejected in [{'checked_cells':[], 'issues':[]},
                         {'checked_cells':['B3'], 'issues':[{'cells':['B3'],'reason':'计量含义未确认'}]}]:
            result = apply_identity_review({'B3':row}, rejected, standard, confirm_slots=True)
            self.assertEqual(result['slot_bindings'], [])
            self.assertEqual(len(result['unresolved']), 1)

    def test_名称和金额均空白仍经普通B入口保存槽位并正式回读(self):
        case = 范围定位.B范围定位测试(); case.setUp()
        with patch.object(UnifiedSemanticEngine, '_http', lambda e,s,p:case.response(e,s,p)):
            target = case.case.target()
        source = case.case.folder/'名称及金额均空白.xlsx'
        book = load_workbook(case.case.source); sheet = book['综合']
        sheet['A2'] = '固定资产名称'; sheet['A3'] = None; sheet['B3'] = None
        book.save(source); book.close()
        def answer(engine, system, payload):
            result = case.response(engine, system, payload)
            if payload['task']=='a_semantic_read':
                result['tables'][0]['meaning']='按固定资产名称逐项列示期末原值，具体名称与金额尚未填写。'
            if payload['task']=='a_semantic_bind':
                for d in result['tables'][0]['dimensions']:
                    d['evidence']=['A1','A2','B2']
                    if d['id']=='asset_selection':
                        d['value']={'domain':'固定资产','mode':'unknown','completeness':'unknown'}
            return result
        with patch.object(UnifiedSemanticEngine, '_http', answer):
            engine=case.case.engine(); result=engine.recognize_workbook(source,target_mapping=target)
        saved=load_fact_mapping(result['mapping_path'],engine.standard)
        self.assertEqual(saved['facts'], [])
        self.assertEqual(saved['unresolved'], [])
        self.assertEqual(len(saved['slot_bindings']), 1)
        self.assertEqual(saved['slot_bindings'][0]['partial_semantics']['dimensions']['asset_selection']['mode'],'unknown')
        self.assertEqual(saved['slot_bindings'][0]['cell'],'B3')
        self.assertTrue(saved['b_identity_coverage']['complete'])
        coverage=saved['recognition_scope']['coverage']
        self.assertTrue(coverage['complete']); self.assertTrue(coverage['satisfied'])
        self.assertEqual(coverage['targets'][0]['standard_sources'], ['综合!B3'])
        self.assertEqual(coverage['targets'][0]['exact_sources'], [])
        from 附注更新.统一勾稽 import build_fact_update_plan
        with self.assertRaisesRegex(ValueError,'实际维度取值未齐'):
            build_fact_update_plan(target,saved,engine.standard)
        broken=copy.deepcopy(saved); broken['slot_bindings'][0]['partial_semantics']['metric_id']='不存在的标准'
        from 附注更新.A语义阅读 import verify_mapping
        with self.assertRaises(ValueError):verify_mapping(broken,engine.standard)


if __name__=='__main__':unittest.main()
