# coding: utf-8
"""诊断用请求表示必须可完整还原，不能删业务定义或更改原格值。"""
import copy
import importlib
import json
import sys
import unittest
from pathlib import Path

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'金标准/全量重整审查_20260920_1710'))


class PayloadRepresentationTests(unittest.TestCase):
    def setUp(self):
        self.api=importlib.import_module('诊断语义请求重复内容')

    def payload(self):
        rule='所有范围必须根据实际业务含义识别，不能根据坐标或金额匹配。'*8
        return {'task':'unified_classify','candidate_cells':['B2'],
                'cells':{'A1':{'value':'合成标题','hidden':False,'formula':None},
                         'B2':{'value':-31.25,'hidden':False,'formula':None}},
                'dimensions':[{'id':name,'value_type':'selection','meaning':name+'对象范围。'+rule} for name in ['存货','固定资产','应收款']],
                'metrics':[{'id':'不参与删减的指标'}],'document_context':[{'text':'实际适用说明'}]}

    def test_repeated_content_is_shorter_and_exactly_reversible(self):
        original=self.payload();before=copy.deepcopy(original)
        compact=self.api.compact_payload(original)
        restored=self.api.expand_payload(compact)
        self.assertEqual(json.dumps(restored,sort_keys=True,ensure_ascii=False),json.dumps(original,sort_keys=True,ensure_ascii=False))
        self.assertLess(len(json.dumps(compact,ensure_ascii=False)),len(json.dumps(original,ensure_ascii=False)))
        self.assertEqual(original,before)
        self.assertEqual(compact['metrics'],original['metrics'])
        self.assertEqual(compact['candidate_cells'],original['candidate_cells'])

    def test_boolean_zero_blank_and_explicit_override_are_preserved(self):
        original=self.payload();original['cells']['B2']['hidden']=0
        original['cells']['C3']={'value':None,'hidden':False,'formula':'=1+2'}
        compact=self.api.compact_payload(original)
        self.assertNotIn('hidden',compact.get('cell_defaults',{}))
        restored=self.api.expand_payload(compact)
        self.assertIs(type(restored['cells']['B2']['hidden']),int)
        self.assertEqual(restored['cells'],original['cells'])

    def test_unique_meanings_and_missing_fields_do_not_gain_shared_properties(self):
        original=self.payload();original['dimensions']=original['dimensions'][:1]
        del original['cells']['B2']['hidden']
        compact=self.api.compact_payload(original)
        self.assertNotIn('shared_dimension_meaning_suffix',compact)
        self.assertNotIn('hidden',compact.get('cell_defaults',{}))
        self.assertEqual(self.api.expand_payload(compact),original)

    def test_reserved_representation_fields_cannot_be_overwritten(self):
        original=self.payload();original['cell_defaults']={'value':'不得覆盖'}
        with self.assertRaises(ValueError):self.api.compact_payload(original)


if __name__=='__main__':unittest.main()
