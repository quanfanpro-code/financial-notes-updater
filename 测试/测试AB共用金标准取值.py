# coding: utf-8
"""用模拟传输检查共用标准、分层保存及无效原字段保护，不调用模型。"""
from copy import deepcopy
import json
from pathlib import Path
import unittest
from unittest.mock import patch

from 测试 import 测试直接识别 as fixture
from 测试.测试统一语义 import records
from 附注更新.A合集 import _records, save_direct_answer
from 附注更新.分步流程 import run_step
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一勾稽 import load_fact_mapping


class AB共用金标准取值测试(unittest.TestCase):
    def setUp(self):
        self.case = fixture.直接识别测试()
        self.case.setUp()
        self.config = dict(self.case.config)
        self.selection = {'domain': '固定资产', 'mode': 'members', 'members': ['设备乙', '设备甲']}

    def answer(self, **changes):
        row = {'location': '合成表!B3', 'metric_id': '固定资产原值',
               'currency': 'CNY', 'period': {'kind': 'instant', 'date': '2025-12-31'},
               'asset_selection': self.selection}
        row.update(changes)
        return json.dumps(row, ensure_ascii=False)

    def recognize(self, side, reply, **extra):
        with patch.object(UnifiedSemanticEngine, 'request_text', return_value=reply) as call:
            result = run_step('recognize_' + side, {**self.config, **extra})
        self.assertEqual(result['status'], 'complete', result)
        call.assert_called_once()
        return load_fact_mapping(result['files'][side + '_mapping'], {}), call.call_args.args

    def test_AB同一提示规则且只带完整相关定义(self):
        gold = self.case.folder / '含无关维度合成金标准.jsonl'
        rows = records() + [{'record_type': 'dimension', 'id': 'unrelated', 'domain': '无关',
                             'meaning': '本次未使用的维度', 'value_type': 'text'}]
        gold.write_text('\n'.join(json.dumps(r, ensure_ascii=False) for r in rows), encoding='utf-8')
        self.config['gold_path'] = str(gold)
        a_path = None
        for side in ('a', 'b'):
            extra = {side + '_allowed_metrics': ['固定资产原值', 'asset_selection']}
            if side == 'b':
                extra['a_mapping'] = a_path
            saved, (prompt, materials) = self.recognize(side, self.answer(), **extra)
            if side == 'a':
                a_path = saved['mapping_path']
            self.assertIn('槽位编号、维度定义、维度实际值是三层', prompt)
            self.assertIn('A、B共用所给金标准原文', prompt)
            self.assertIn('标准没有规定的答复包装或额外字段不能成为填写要求', prompt)
            self.assertNotIn('domain使用本维度金标准的domain', prompt)
            supplied = {json.loads(line)['id']: json.loads(line)
                        for line in materials['金标准.jsonl'].splitlines()}
            self.assertNotIn('unrelated', supplied)
            for definition in records():
                if definition['record_type'] in ('metric', 'dimension'):
                    self.assertEqual(supplied[definition['id']], definition)

    def test_AB同格冲突不选择也不覆盖原有值(self):
        a, _ = self.recognize('a', self.answer())
        b, _ = self.recognize('b', self.answer(), a_mapping=a['mapping_path'])
        for side, old in [('A', a), ('B', b)]:
            raw = json.dumps({'location': '合成表!B3', 'currency': 'CNY',
                              'dimensions': {'currency': 'USD'}}, ensure_ascii=False)
            with patch.object(UnifiedSemanticEngine, 'request_text',
                              side_effect=AssertionError('保存不得调用模型')):
                saved = save_direct_answer(old, raw, self.case.folder / f'{side}冲突字段.json',
                                           field_scope={'合成表!B3': ['currency']})
            self.assertEqual(dict(_records(saved))['合成表!B3']['dimensions'],
                             dict(_records(old))['合成表!B3']['dimensions'])
            issues = [row for table in saved['outline']['subtables']
                      for row in table['unresolved'] if row.get('id') == 'currency']
            self.assertTrue(any(row.get('conflicting_values') == ['CNY', 'USD'] for row in issues))
            self.assertEqual(saved['本次逐格原答复'], raw)
            self.assertEqual(saved['机械记录'], old['机械记录'])

    def test_原文实际日期不受程序包装要求拦截(self):
        saved, _ = self.recognize('a', self.answer(period='2025-12-31'))
        row = dict(_records(saved))['合成表!B3']
        self.assertEqual(row['metric_id'], '固定资产原值')
        self.assertEqual(row['dimensions']['period'], '2025-12-31')
        self.assertNotIn('entity', row['dimensions'])

    def test_槽位编号和维度编号不能互相代选(self):
        for side in ('a', 'b'):
            raw = self.answer(dimension_id='asset_selection')
            a, _ = self.recognize('a', self.answer())
            extra = {'a_mapping': a['mapping_path']} if side == 'b' else {}
            saved, _ = self.recognize(side, raw, **extra)
            self.assertNotIn('合成表!B3', dict(_records(saved)))
            issues = [row for table in saved['outline']['subtables'] for row in table['unresolved']]
            self.assertTrue(any('同时给出槽位编号和维度编号' in row['reason'] for row in issues))
            self.assertEqual(saved['本次逐格原答复'], raw)

    def test_同格维度冲突只影响本维度(self):
        raw = self.answer(dimensions={'currency': 'USD'})
        saved, _ = self.recognize('a', raw)
        row = dict(_records(saved))['合成表!B3']
        self.assertEqual(row['metric_id'], '固定资产原值')
        self.assertEqual(row['dimensions']['asset_selection'], self.selection)
        self.assertNotIn('currency', row['dimensions'])
        self.assertEqual(row['dimensions']['period'], {'kind': 'instant', 'date': '2025-12-31'})

    def test_正确值原样保存不排序不填充(self):
        saved, _ = self.recognize('a', self.answer())
        row = dict(_records(saved))['合成表!B3']
        self.assertEqual(row['dimensions']['asset_selection'], self.selection)
        self.assertNotIn('completeness', row['dimensions']['asset_selection'])
        self.assertNotIn('entity', row['dimensions'])


if __name__ == '__main__':
    unittest.main()
