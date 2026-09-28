# coding: utf-8
"""自动补充不得推翻已停用决定；有效定义和真实新定义仍可使用。"""
import copy
import unittest
from pathlib import Path
from 附注更新.统一补充 import _proposal
from 测试.测试统一语义 import records


class RetiredDefinitionTests(unittest.TestCase):
    def test_retired_cannot_reactivate_or_be_reused(self):
        rows = records()
        prior = rows[-1]
        payload = {'source_hash': 'a'*64, 'mapping_hash': 'b'*64, 'sheet': '测试',
                   'candidate_cells': ['B2'], 'text_evidence_cells': ['A2'],
                   'cells': {'A2': {'value': '固定资产原值'}}, 'context_evidence': {}}
        decision = {'cell': 'B2', 'decision': 'existing', 'reason': '使用已有计量定义',
                    'record_ids': [prior['id']], 'evidence_cells': ['A2']}
        _proposal({'records': [], 'decisions': [decision]}, payload, rows, Path('复核.json'))
        prior['status'] = 'retired'
        with self.assertRaisesRegex(ValueError, '复用决定'):
            _proposal({'records': [], 'decisions': [decision]}, payload, rows, Path('复核.json'))
        proposed = copy.deepcopy(prior)
        proposed.pop('source_refs')
        proposed['status'] = 'active'
        decision['decision'] = 'revision'
        with self.assertRaisesRegex(ValueError, '重新启用'):
            _proposal({'records': [proposed], 'decisions': [decision]}, payload, rows, Path('复核.json'))
        proposed['id'] = prior['id'] + '_换号'
        proposed['dimension_refs'].reverse()
        decision['record_ids'] = [proposed['id']]
        for status in ('retired', 'active'):
            prior['status'] = status
            with self.subTest(status=status), self.assertRaisesRegex(ValueError, '换号重建'):
                _proposal({'records': [proposed], 'decisions': [decision]}, payload, rows, Path('复核.json'))


if __name__ == '__main__':
    unittest.main()
