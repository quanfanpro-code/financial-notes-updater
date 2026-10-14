# coding: utf-8
"""同科目共用阅读材料，独立子表与已核查成果保持完整。"""
import unittest

from 附注更新.A语义阅读 import semantic_groups


class 科目分组测试(unittest.TestCase):
    def test_同科目不同位置共同处理且已核查组不变(self):
        tables = [
            {'id': '旧账龄', 'subject': '应收款', 'sheet': '甲', 'range': 'A1:D6'},
            {'id': '分类', 'subject': '应收款', 'sheet': '甲', 'range': 'A20:D30'},
            {'id': '资本', 'subject': '实收资本', 'sheet': '乙', 'range': 'A1:G8'},
            {'id': '准备变动', 'subject': '应收款', 'sheet': '乙', 'range': 'A40:F48'},
        ]
        fresh = semantic_groups(tables, {})
        self.assertEqual([[t['id'] for t in group] for group in fresh],
                         [['旧账龄', '分类', '准备变动'], ['资本']])
        resumed = semantic_groups(tables, {'旧账龄': [tables[0]]})
        self.assertEqual([[t['id'] for t in group] for group in resumed],
                         [['旧账龄'], ['分类', '准备变动'], ['资本']])
        self.assertEqual(sorted(t['id'] for g in resumed for t in g),
                         sorted(t['id'] for t in tables))
        self.assertEqual(next(t for g in resumed for t in g if t['id'] == '准备变动'), tables[3])
        self.assertEqual(semantic_groups([], {}), [])


if __name__ == '__main__':
    unittest.main()
