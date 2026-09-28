"""验证模型声明的行列语义展开，防止坐标被当作隐含业务规则。"""
import copy
import unittest

from 附注更新.回执格式 import normalize_response


class 结构展开测试(unittest.TestCase):
    def setUp(self):
        self.payload = {'task': 'unified_classify', 'response_contract': 'unified_structure_v1',
                        'candidate_cells': ['B2', 'C2', 'B3', 'C3'],
                        'cells': {a: {'value': v} for a, v in
                                  [('A1', '项目'), ('B1', '期末'), ('C1', '期初'),
                                   ('A2', '现金'), ('A3', '存款'),
                                   ('B2', 8), ('C2', None), ('B3', '—'), ('C3', 9)]},
                        'dimensions': [{'id': 'period'}, {'id': 'category'}]}
        self.reply = {
            'dimension_values': [
                {'dimension_id': 'period', 'value': '2025', 'evidence': ['B1']},
                {'dimension_id': 'period', 'value': '2024', 'evidence': ['C1']},
                {'dimension_id': 'category', 'value': '现金', 'evidence': ['A2']},
                {'dimension_id': 'category', 'value': '存款', 'evidence': ['A3']}],
            'regions': [{'range': 'B2:C3',
                         'record': {'kind': 'metric_fact', 'metric_id': 'metric.balance',
                                    'dimension_refs': [], 'reason': '行项目与列表头共同说明余额'},
                         'rows': [{'indices': [2], 'record': {'dimension_refs': [2]}},
                                  {'indices': [3], 'record': {'dimension_refs': [3]}}],
                         'columns': [{'indices': ['B'], 'record': {'dimension_refs': [0]}},
                                     {'indices': ['C'], 'record': {'dimension_refs': [1]}}]}],
            'cells': {}}

    def test_行列语义展开与单格例外(self):
        self.reply['cells']['C3'] = {'kind': 'unresolved', 'reason': '原说明有歧义', 'evidence_cells': ['A3']}
        result = normalize_response(self.payload, self.reply)['cells']
        self.assertEqual(set(result), {'B2', 'C2', 'B3', 'C3'})
        self.assertEqual(result['B2']['dimensions'], {'period': '2025', 'category': '现金'})
        self.assertEqual(result['C2']['dimensions'], {'period': '2024', 'category': '现金'})
        self.assertEqual(result['B3']['dimensions'], {'period': '2025', 'category': '存款'})
        self.assertEqual(result['C3']['kind'], 'unresolved')
        self.assertNotIn('value', result['B2'])

    def test_冲突越界不能静默覆盖(self):
        for change in ('重复区域', '重复维度', '越界'):
            with self.subTest(change=change):
                reply = copy.deepcopy(self.reply)
                if change == '重复区域': reply['regions'].append(copy.deepcopy(reply['regions'][0]))
                elif change == '重复维度': reply['regions'][0]['record']['dimension_refs'] = [0]
                else: reply['regions'][0]['range'] = 'B2:D3'
                with self.assertRaises(ValueError): normalize_response(self.payload, reply)


if __name__ == '__main__': unittest.main()
