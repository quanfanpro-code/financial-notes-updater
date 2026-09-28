# coding: utf-8
"""验证单格回执错误不会丢弃其他格，也不会成为补充金标准的依据。"""
import copy
import json
import unittest
import uuid
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook

from 附注更新.统一识别 import UnifiedSemanticEngine,validate_response,agree_rounds
from 附注更新.统一语义 import validate_records,read_standard_v2
from 附注更新.统一勾稽 import load_fact_mapping
from 附注更新.分步流程 import run_step
from 测试.测试统一识别 import packet,answer
from 测试.测试统一语义 import records


class CellIsolationTests(unittest.TestCase):
    def test_financial_self_reference_cannot_hide_among_valid_headers(self):
        standard=validate_records(records());raw=answer()
        raw['cells']['D7']['dimension_evidence']['asset_selection']=['A1','D7']
        with self.assertRaisesRegex(ValueError,'财务值本身'):
            validate_response(raw,packet(),standard)
        checked=validate_response(raw,packet(),standard,isolate_errors=True)
        self.assertEqual(checked['D7']['kind'],'invalid_response')
        self.assertEqual(checked['F9']['kind'],'metric_fact')
        result=agree_rounds(checked,checked,standard)
        self.assertEqual(len(result['facts']),1)
        self.assertEqual(result['unresolved'][0]['cell'],'D7')

    def test_heading_can_still_quote_its_own_actual_text(self):
        source=packet();source['candidate_cells']=['A1']
        raw={'cells':{'A1':{'kind':'non_business','category':'title','reason':'标题原文',
                            'evidence_cells':['A1']}}}
        checked=validate_response(raw,source,validate_records(records()),isolate_errors=True)
        self.assertEqual(checked['A1']['kind'],'non_business')

    def test_invalid_evidence_keeps_valid_fact_and_stays_unresolved(self):
        standard=validate_records(records());raw=answer();raw['cells']['F9']['dimension_evidence']['entity']=['F9']
        with self.assertRaises(ValueError):validate_response(raw,packet(),standard)
        checked=validate_response(raw,packet(),standard,isolate_errors=True)
        result=agree_rounds(checked,checked,standard)
        self.assertEqual(len(result['facts']),1)
        self.assertEqual(result['facts'][0]['source_reference']['location'],'任意布局!D7')
        self.assertEqual(result['facts'][0]['raw_value'],125.0)
        self.assertEqual(result['excluded'],[])
        self.assertEqual(result['unresolved'][0]['cell'],'F9')
        self.assertIn(checked['F9']['reason'],result['unresolved'][0]['reason'])
        self.assertEqual(checked['F9']['kind'],'invalid_response')

    def test_wrong_numeric_exclusion_and_unknown_metric_are_never_accepted(self):
        standard=validate_records(records())
        for bad in ({'kind':'non_business','category':'empty_padding','reason':'错误回执','evidence_cells':['A1']},
                    {**answer()['cells']['D7'],'metric_id':'不存在的指标'}):
            raw=answer();raw['cells']['D7']=bad
            checked=validate_response(raw,packet(),standard,isolate_errors=True)
            self.assertEqual(checked['D7']['kind'],'invalid_response')
            self.assertEqual(checked['F9']['kind'],'metric_fact')

    def test_invalid_outer_coverage_still_rejects_whole_response(self):
        standard=validate_records(records())
        for kind in ('missing','extra'):
            raw=answer()
            if kind=='missing':raw['cells'].pop('F9')
            else:raw['cells']['Z99']=copy.deepcopy(raw['cells']['D7'])
            with self.assertRaisesRegex(ValueError,'覆盖'):
                validate_response(raw,packet(),standard,isolate_errors=True)

    def test_formal_entry_saves_valid_cells_without_proposing_gold_for_bad_receipt(self):
        root=Path(__file__).resolve().parents[1]/'测试结果'/('逐格回执隔离_'+uuid.uuid4().hex);root.mkdir()
        gold=root/'标准.jsonl';gold.write_text('\n'.join(json.dumps(row,ensure_ascii=False) for row in records()),encoding='utf-8')
        source=root/'原表.xlsx';book=Workbook();sheet=book.active;sheet.title='任意布局'
        sheet['A1']=packet()['cells']['A1']['value'];sheet['B2']=125;sheet['C2']=77;book.save(source);book.close()
        calls=[]
        class ControlledEngine(UnifiedSemanticEngine):
            def _request(self,system,payload,validator):
                calls.append(payload['task'])
                if payload['task']!='unified_classify':raise AssertionError('无效回执不能据此新增业务定义')
                response={'cells':{}}
                for address in payload['candidate_cells']:
                    if address=='B2':row=copy.deepcopy(answer()['cells']['D7'])
                    elif address=='C2':row={'kind':'non_business','category':'empty_padding','reason':'错误地排除真实数值','evidence_cells':['A1']}
                    else:row={'kind':'non_business','category':'title' if address=='A1' else 'empty_padding','reason':'受控测试','evidence_cells':['A1']}
                    response['cells'][address]=row
                return validator(response)
        with patch.object(UnifiedSemanticEngine,'_request',ControlledEngine._request):
            result=run_step('recognize_b',{'gold_path':str(gold),'b_path':str(source),'output_dir':str(root/'正式入口'),
                                         'base_url':'http://unused.example/v1','model':'受控测试','api_key':'','timeout':600})
        self.assertEqual(result['status'],'partial')
        mapping=load_fact_mapping(result['files']['b_mapping'],read_standard_v2(gold))
        self.assertEqual(len(mapping['facts']),1)
        self.assertEqual(mapping['facts'][0]['source_reference']['location'],'任意布局!B2')
        self.assertEqual([row['cell'] for row in mapping['unresolved']],['C2'])
        self.assertFalse(mapping['complete'])
        self.assertTrue(calls and set(calls)=={'unified_classify'})


if __name__=='__main__':unittest.main()
