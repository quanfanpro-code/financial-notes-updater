"""保留人工示例库供检查；实验检索不自动注入正式识别请求。"""
import copy
import unittest
from 附注更新.统一识别 import build_payload
from 附注更新.人工映射复用 import reviewed_examples
from 附注更新.统一语义 import validate_records
from 测试.测试统一语义 import records


class 人工映射复用测试(unittest.TestCase):
    def test_人工成果保留但实验检索不进入正式请求(self):
        packet = {'sheet': '任意新样式', 'candidate_cells': ['B2'], 'cells': {
            'A1': {'value': '货币资金 库存现金 期末余额'}, 'B2': {'value': 987654321}}}
        standard = validate_records(records())
        result = build_payload(packet, standard)
        self.assertNotIn('reviewed_examples', result)
        examples = reviewed_examples(packet)
        self.assertTrue(any(x['legacy_slot_id'] == 'C-N087-T001-S001' for x in examples['examples']))
        self.assertEqual(examples['corpus_business_cells'], 27967)
        shifted = copy.deepcopy(packet)
        shifted['cells'] = {'Z90': {'value': '货币资金 库存现金 期末余额'}, 'AC95': {'value': -7}}
        shifted['candidate_cells'] = ['AC95']
        self.assertEqual(reviewed_examples(shifted), examples)
        for item in examples['examples']:
            self.assertNotIn('raw_value', item)
            self.assertNotIn('cached_value', item)
            self.assertIn(item['legacy_slot_id'], examples['legacy_definitions'])
        unknown = {'sheet':'', 'cells': {'A1':{'value':123}}, 'candidate_cells':['A1']}
        self.assertEqual(reviewed_examples(unknown)['examples'], [])


if __name__ == '__main__':
    unittest.main()
