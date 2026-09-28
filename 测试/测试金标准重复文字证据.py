# coding: utf-8
"""同一表头文字出现在不同原格，不应阻断缺失语义定义的保存。"""
import copy
import unittest
from pathlib import Path
from 附注更新.统一补充 import _proposal
from 测试.测试统一语义 import records


class RepeatedEvidenceTests(unittest.TestCase):
    def test_equal_text_at_distinct_cells_preserves_decisions_and_saves_definition(self):
        original = records()
        metric = copy.deepcopy(original[-1])
        metric.pop('source_refs')
        metric['id'] = '固定资产累计折旧'
        metric['definition'].update(measure='累计折旧余额', measurement_basis='已经计提的累计折旧')
        payload = {'source_hash': 'a'*64, 'mapping_hash': 'b'*64, 'sheet': '重复表头',
            'candidate_cells': ['B2', 'B3'], 'text_evidence_cells': ['A2', 'A3'],
            'cells': {'A2': {'value': '累计折旧'}, 'A3': {'value': '累计折旧'}}, 'context_evidence': {}}
        decisions = [{'cell': cell, 'decision': 'revision', 'reason': '计量基础不同于原值',
            'record_ids': [metric['id']], 'evidence_cells': [evidence]} for cell, evidence in [('B2', 'A2'), ('B3', 'A3')]]
        result = _proposal({'records': [metric], 'decisions': decisions}, payload, original, Path('本地审查记录.json'))
        self.assertEqual(result['decisions'], decisions)
        source = next(r for r in result['rows'] if r['id'] not in {x['id'] for x in original} and r['record_type'] == 'source_example')
        self.assertEqual(source['text_evidence'], ['累计折旧'])
        self.assertEqual(source['origin']['location'], '重复表头!A2')


if __name__ == '__main__':
    unittest.main()
