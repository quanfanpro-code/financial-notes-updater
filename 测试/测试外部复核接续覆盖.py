# coding: utf-8
"""纠正实例取值须刷新覆盖；历史定位仍由已验证的原映射约束。"""
import copy
import json
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import load_workbook
from 测试.测试B范围定位 import B范围定位测试
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一语义 import definition_hash, read_standard_v2
from 附注更新.统一外部复核 import preview_external_review, apply_external_review
from 附注更新.统一勾稽 import load_fact_mapping
from 附注更新.表格 import file_hash


class 外部复核接续覆盖测试(unittest.TestCase):
    def check_case(self, new_standard=False):
        fixture = B范围定位测试(); fixture.setUp()
        with patch.object(UnifiedSemanticEngine, '_http', lambda e, s, p: fixture.response(e, s, p)):
            target = fixture.case.target()
            book = load_workbook(fixture.case.source)
            book['综合']['A1'] = '合成企业甲，人民币元，2024年12月31日'
            source = fixture.case.folder / '期间误识别原文.xlsx'
            book.save(source); book.close()
            engine = fixture.case.engine()
            parent = engine.recognize_workbook(source, target_mapping=target)
        parent['carrier'] = 'B'
        parent['mapping_path'] = str(fixture.case.folder / '独立B映射.json')
        Path(parent['mapping_path']).write_text(json.dumps(parent, ensure_ascii=False), encoding='utf-8-sig')
        original = copy.deepcopy(parent['facts'][0])
        gold = Path(parent['gold_path'])
        if new_standard:
            records = [json.loads(line) for line in gold.read_text(encoding='utf-8-sig').splitlines() if line.strip()]
            added = copy.deepcopy(records[-1]); added['id'] = '累计折旧'
            added['definition']['measurement_basis'] = '累计折旧余额'
            records.append(added)
            gold = fixture.case.folder / '只新增定义的新标准.jsonl'
            gold.write_text('\n'.join(json.dumps(r, ensure_ascii=False) for r in records), encoding='utf-8-sig')
        standard = read_standard_v2(gold)
        dims = copy.deepcopy(original['dimensions']); dims['period']['date'] = '2024-12-31'
        item = {'sheet': '综合', 'cell': 'B3', 'metric_id': original['metric_id'], 'dimensions': dims,
                'dimension_evidence': {k: [r.removeprefix('综合!') for r in refs] for k, refs in original['dimension_evidence'].items()},
                'replace_fact_hash': definition_hash(original), 'reason': '原文A1明确2024，纠正错误实际期间'}
        opinion = {'schema': '附注外部语义复核-v2', 'carrier': 'B', 'mapping_sha256': file_hash(parent['mapping_path']),
                   'source_hash': parent['source_hash'], 'gold_hash': standard['sha256'],
                   'reviewer': {'name': '受控复核测试', 'method': 'human', 'reviewed_at': '2026-10-11T20:13:00+08:00'}, 'items': [item]}
        path = fixture.case.folder / '复核意见.json'; path.write_text(json.dumps(opinion, ensure_ascii=False), encoding='utf-8-sig')
        preview = preview_external_review(parent['mapping_path'], path, gold)
        result = apply_external_review(preview, fixture.case.folder, confirmed=True, confirm_method='测试明确采纳')
        loaded = load_fact_mapping(result['files']['b_mapping'], standard)
        need = loaded['recognition_scope']['coverage']['targets'][0]
        self.assertEqual(need['exact_sources'], [])
        self.assertEqual(need['standard_sources'], ['综合!B3'])
        self.assertEqual(loaded['facts'][0]['raw_value'], original['raw_value'])
        from 附注更新.表格编排 import source_ready
        self.assertTrue(source_ready(loaded,target,standard))
        # 更正不赋予修改历史候选或覆盖声明的权限。
        for name, alter in [('历史候选', lambda m: m['recognition_scope']['candidate_locations'].append('综合!C3')),
                            ('覆盖声明', lambda m: m['recognition_scope']['coverage']['targets'][0].update(exact_sources=['综合!B3']))]:
            bad = copy.deepcopy(loaded); alter(bad)
            bad_path = fixture.case.folder / (name + '.json'); bad_path.write_text(json.dumps(bad, ensure_ascii=False), encoding='utf-8-sig')
            with self.assertRaisesRegex(ValueError, 'recognition_scope'):
                load_fact_mapping(bad_path, standard)

    def test_纠正期间后槽位不变但实例覆盖刷新(self):
        self.check_case()

    def test_新标准接续仍核验旧定位及新覆盖(self):
        self.check_case(new_standard=True)


if __name__ == '__main__':
    unittest.main()
