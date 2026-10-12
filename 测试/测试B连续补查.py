# coding: utf-8-sig
"""验证补查按未满足的身份继续，不能依赖第一次候选或固定位置。"""
import copy
import json
import unittest
from unittest.mock import patch
from openpyxl import load_workbook

from 测试.测试B范围定位 import B范围定位测试
from 测试.测试统一语义 import fact, records
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一勾稽 import load_fact_mapping


class 连续补查测试(unittest.TestCase):
    def run_case(self, initially_empty=False):
        case = B范围定位测试(); case.setUp()
        standards=records()
        for name in ('固定资产累计折旧','固定资产减值准备'):
            metric=copy.deepcopy(standards[-1]);metric['id']=name
            metric['definition']['measurement_basis']=name
            standards.append(metric)
        case.case.gold=case.case.folder/'不同计量标准.jsonl'
        case.case.gold.write_text('\n'.join(json.dumps(r,ensure_ascii=False) for r in standards),encoding='utf-8-sig')
        metrics={'B':'固定资产累计折旧','C':'固定资产减值准备','D':'固定资产原值'}
        book = load_workbook(case.case.source)
        sheet = book['综合']
        for column, metric in metrics.items():
            sheet[column+'2'] = metric
            sheet[column+'3'] = 71 if column == 'D' else 36
        source = case.case.folder/'不同编排.xlsx'
        book.save(source); book.close()
        searched = []

        def reply(engine, system, payload):
            task = payload['task']
            if task == 'b_candidate_followup':
                searched.append(bool(payload.get('identified_candidates')))
                column = ('C' if len(searched) == 1 else 'D') if not initially_empty else ('B' if len(searched) == 1 else 'D')
                candidates = [] if initially_empty and len(searched) == 1 else ['综合!'+column+'3']
                return {'read_ranges': [], 'ranges': [{'sheet':'综合','range':'A1:D3'}],
                        'structure':'检查各计量列，逐次补入新来源', 'evidence':['综合!'+column+'2'],
                        'targets':[{'id':r['id'],'candidates':candidates,'reason':'按原文计量含义查找'} for r in payload['targets']]}
            if task == 'a_semantic_read':
                return {'tables':[{'id':'原值','sheet':'综合','range':'A1:D3','subject':'固定资产',
                    'meaning':'累计折旧、减值准备、原值分列，均为相同时点',
                    'data_ranges':[p.rpartition('!')[2] for p in payload['candidate_locations']]}], 'notes':'按各列实际计量绑定'}
            if task == 'a_semantic_select':
                return {'tables':[{'id':'原值','metric_ids':list(metrics.values()),'missing':''}]}
            if task == 'a_semantic_bind':
                cells = payload['reading'][0]['data_ranges']; dimensions = []
                for cell in cells:
                    dims = copy.deepcopy(fact()['dimensions'])
                    dimensions.extend({'ranges':[cell],'id':key,'value':value,'evidence':['A1',cell[0]+'2','A3']} for key,value in dims.items())
                return {'tables':[{'id':'原值','metrics':[{'ranges':[cell],'id':metrics[cell[0]]} for cell in cells], 'dimensions':dimensions, 'unresolved':[]}]}
            result = case.response(engine, system, payload)
            if task == 'b_locate' and result['areas']:
                result['areas'][0]['ranges'] = [{'sheet':'综合','range':'A1:D3'}]
            if task == 'subtable_detail' and initially_empty:
                for row in result['targets']: row.update(candidates=[], reason='首次漏列')
            return result

        with patch.object(UnifiedSemanticEngine, '_http', reply):
            target = case.case.target(); engine = case.case.engine()
            result = engine.recognize_workbook(source, target_mapping=target)
        self.assertEqual(result['recognition_scope']['coverage']['targets'][0]['exact_sources'], ['综合!D3'])
        saved = load_fact_mapping(result['mapping_path'], engine.standard)
        self.assertEqual(saved['recognition_scope']['coverage']['targets'][0]['exact_sources'], ['综合!D3'])
        self.assertEqual(len(searched), 2)
        return saved

    def test_新补入的候选仍为其他计量则继续补查(self):
        self.run_case()

    def test_首轮及首次补查均为空仍进入身份后的补查(self):
        self.run_case(initially_empty=True)


if __name__ == '__main__':
    unittest.main()
