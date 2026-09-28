# coding: utf-8
"""只验证改义后的局部接续及依据校验，不作为模型准确率证据。"""
import copy
import json
import unittest
import uuid
from pathlib import Path
from openpyxl import Workbook
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一语义 import read_standard_v2
from 附注更新.统一勾稽 import load_fact_mapping
from 测试.测试统一识别 import answer
from 测试.测试统一语义 import records


class RevisionReuseTests(unittest.TestCase):
    def test_changed_meaning_is_reidentified_while_other_classification_is_reused(self):
        root = Path(__file__).resolve().parents[1]/'测试结果'/('金标准局部复用_'+uuid.uuid4().hex)
        root.mkdir()
        rows = records()
        added = copy.deepcopy(rows[-1])
        added['id'] = '固定资产累计折旧'
        added['definition']['measurement_basis'] = '累计折旧'
        rows.append(added)
        gold = root/'原标准.jsonl'
        gold.write_text('\n'.join(json.dumps(r, ensure_ascii=False) for r in rows), encoding='utf-8')
        source = root/'人工原表.xlsx'
        book = Workbook()
        book.active['A1'] = '合成企业甲，单户人民币元，2025年12月31日固定资产原值与累计折旧'
        book.active['B2'] = 125
        book.active['C2'] = 32
        book.save(source)
        book.close()
        calls = []

        class ControlledEngine(UnifiedSemanticEngine):
            def _request(self, system, payload, validator):
                calls.append(payload['candidate_cells'])
                response = {'cells': {}}
                for address in payload['candidate_cells']:
                    if address in ('B2', 'C2'):
                        row = copy.deepcopy(answer()['cells']['D7'])
                        row['metric_id'] = '固定资产原值' if address == 'B2' else '固定资产累计折旧'
                    else:
                        row = {'kind': 'non_business', 'category': 'title' if address == 'A1' else 'empty_padding',
                               'reason': '受控本地资料，不代表模型准确率', 'evidence_cells': ['A1']}
                    response['cells'][address] = row
                return validator(response)

        first = ControlledEngine({}, gold, root/'首次识别').recognize_workbook(source)
        # 新增无关槽位后，整表可引用未改义的历史局部包，原包版本不改写。
        expanded_rows = copy.deepcopy(rows)
        extra = copy.deepcopy(rows[-1])
        extra['id'] = '未使用的新槽位'
        extra['definition']['measurement_basis'] = '账面净值'
        expanded_rows.append(extra)
        expanded_gold = root/'增加无关槽位.jsonl'
        expanded_gold.write_text('\n'.join(json.dumps(r, ensure_ascii=False) for r in expanded_rows), encoding='utf-8')
        combined = copy.deepcopy(first)
        combined.update(gold_path=str(expanded_gold), gold_hash=read_standard_v2(expanded_gold)['sha256'])
        combined_path = root/'保留历史包的整表.json'
        combined_path.write_text(json.dumps(combined, ensure_ascii=False), encoding='utf-8-sig')
        loaded = load_fact_mapping(combined_path, read_standard_v2(expanded_gold))
        self.assertEqual(loaded['facts'], first['facts'])
        self.assertEqual(loaded['packet_mappings'], first['packet_mappings'])
        for label, index, key in [('指标', -3, 'definition'), ('维度', 0, 'meaning')]:
            incompatible = copy.deepcopy(expanded_rows)
            if label == '指标':
                incompatible[index][key]['measurement_basis'] += '；已改变含义'
            else:
                incompatible[index][key] += '；已改变含义'
            changed_gold = root/(label+'改变.jsonl')
            changed_gold.write_text('\n'.join(json.dumps(r, ensure_ascii=False) for r in incompatible), encoding='utf-8')
            combined.update(gold_path=str(changed_gold), gold_hash=read_standard_v2(changed_gold)['sha256'])
            rejected_path = root/(label+'改变后不可复用.json')
            rejected_path.write_text(json.dumps(combined, ensure_ascii=False), encoding='utf-8-sig')
            with self.assertRaisesRegex(ValueError, '定义已改变'):
                load_fact_mapping(rejected_path, read_standard_v2(changed_gold))
        rows[-2]['definition']['measurement_basis'] += '；修订后明确的计量边界'
        revised = root/'修订标准.jsonl'
        revised.write_text('\n'.join(json.dumps(r, ensure_ascii=False) for r in rows), encoding='utf-8')
        calls.clear()
        second = ControlledEngine({}, revised, root/'接续识别').recognize_workbook(source, previous_result=first['mapping_path'])
        self.assertEqual(calls, [['B2'], ['B2']])
        self.assertTrue(second['complete'])
        old_stable = next(r for r in first['facts'] if r['metric_id'] == '固定资产累计折旧')
        self.assertIn(old_stable, second['facts'])
        self.assertEqual(second['excluded'], first['excluded'])
        loaded = load_fact_mapping(second['mapping_path'], read_standard_v2(revised))
        self.assertEqual(len(loaded['facts']), 2)
        changed = copy.deepcopy(second)
        changed['facts'][0]['raw_value'] = 987
        tampered = root/'原值被篡改的映射.json'
        tampered.write_text(json.dumps(changed, ensure_ascii=False), encoding='utf-8-sig')
        with self.assertRaises(ValueError):
            load_fact_mapping(tampered, read_standard_v2(revised))


if __name__ == '__main__':
    unittest.main()
