# coding: utf-8
"""用受控回执验证缺项补充、拒绝重复及后续重识别；不是模型准确率测试。"""
import copy
import importlib
import json
import unittest
from unittest.mock import patch
import uuid
from pathlib import Path
from openpyxl import Workbook
from 测试.测试统一语义 import records, fact
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一语义 import read_standard_v2
from 附注更新.统一勾稽 import load_fact_mapping
from 附注更新.表格 import file_hash


class ControlledEngine(UnifiedSemanticEngine):
    deny = False
    malicious = False
    calls = None

    def _http(self, system, payload):
        self.calls.append(copy.deepcopy(payload))
        if payload['task'] == 'unified_classify':
            available = {row['id'] for row in payload['metrics']}; cells = {}
            for address in payload['candidate_cells']:
                metric = {'B2': '固定资产原值', 'B3': '固定资产累计折旧'}.get(address)
                if not metric:
                    cells[address] = {'kind': 'non_business', 'category': 'header', 'reason': '标题或项目说明', 'evidence_cells': ['A1']}
                elif metric not in available:
                    cells[address] = {'kind': 'unresolved', 'reason': '原值不能表达累计折旧', 'evidence_cells': ['A3']}
                else:
                    cells[address] = {'kind': 'metric_fact', 'metric_id': metric, 'dimensions': fact()['dimensions'],
                                      'dimension_evidence': {key: ['A1', 'A2' if address == 'B2' else 'A3'] for key in fact()['dimensions']},
                                      'reason': '标题、项目及计量基础有明确原文'}
            return {'cells': cells}
        if payload['task'] == 'unified_standard_proposal':
            metric = copy.deepcopy(records()[-1]); metric.pop('source_refs')
            metric['id'] = '固定资产累计折旧'; metric['definition']['measurement_basis'] = '累计已计提折旧'
            if self.malicious: metric['raw_value'] = 999
            return {'records': [metric], 'decisions': [{'cell': address, 'decision': 'revision', 'reason': '不同计量基础，不是新增资产类别',
                    'record_ids': [metric['id']], 'evidence_cells': ['A3']} for address in payload['candidate_cells']]}
        if payload['task'] == 'unified_standard_review':
            return {'checks': [{'record_id': row['id'], 'accepted': not self.deny,
                               'layering_correct': not self.deny, 'instance_only': self.deny,
                               'existing_equivalent_ids': ['固定资产原值'] if self.deny else [],
                               'reason': '拒绝重复或实例拆号' if self.deny else '累计折旧与原值计量基础不同，维度沿用',
                               'evidence_cells': ['A3']} for row in payload['proposed_records']],
                    'metric_scope_checks': [{'metric_id': row['id'], 'can_partition': True, 'dimension_ids': ['asset_selection'],
                                             'reason': '资产类别改变时沿用同一计量指标，通过资产集合说明范围'}
                                            for row in payload['proposed_records'] if row['record_type'] == 'metric']}
        raise AssertionError(payload['task'])


class SupplementTests(unittest.TestCase):
    def setUp(self):
        self.api = importlib.import_module('附注更新.统一补充')
        self.root = Path(__file__).resolve().parents[1]/'测试结果'/('统一语义补充_'+uuid.uuid4().hex)
        self.root.mkdir()
        self.gold = self.root/'原标准.jsonl'
        self.gold.write_text('\n'.join(json.dumps(row, ensure_ascii=False) for row in records()), encoding='utf-8')
        self.source = self.root/'原表.xlsx'
        book = Workbook(); sheet = book.active; sheet.title = '表'
        sheet['A1'] = '合成企业甲；单户人民币元；2025年12月31日；全部固定资产'
        sheet['A2'] = '固定资产原值'; sheet['B2'] = 125
        sheet['A3'] = '固定资产累计折旧'; sheet['B3'] = 25
        book.save(self.source); book.close()
        self.original = {str(path): file_hash(path) for path in (self.gold, self.source)}
        self.engine = ControlledEngine({'model': '受控模型', 'base_url': 'http://localhost'}, self.gold, self.root/'识别')
        self.engine.calls = []
        self.mapping = self.engine.recognize_workbook(self.source)

    def test_new_metric_is_reviewed_published_and_unresolved_cell_resumes(self):
        result = self.api.review_missing_standard(self.engine, self.mapping, self.root/'补充', publish_root=self.root)
        self.assertEqual(result['new_ids'], ['固定资产累计折旧'])
        self.assertTrue(Path(result['report_path']).is_file())
        current = read_standard_v2(result['path'])
        self.assertEqual(len(current['metric']), 2)
        self.assertEqual(len(current['dimension']), 6)
        second = ControlledEngine(self.engine.settings, result['path'], self.root/'接续'); second.calls = []
        updated = second.recognize_workbook(self.source, previous_result=self.mapping['mapping_path'])
        self.assertTrue(updated['complete'])
        self.assertEqual([p['candidate_cells'] for p in second.calls], [['B3'], ['B3']])
        self.assertEqual(updated['facts'][0], self.mapping['facts'][0])
        self.assertEqual(next(f['raw_value'] for f in updated['facts'] if f['metric_id'] == '固定资产累计折旧'), 25)
        load_fact_mapping(updated['mapping_path'], current)
        for path, digest in self.original.items(): self.assertEqual(file_hash(path), digest)
        reviews = [p for p in self.engine.calls if p['task'] == 'unified_standard_review']
        self.assertEqual(len(reviews), 2)
        self.assertTrue(all(any(row['id'] == '固定资产原值' for row in p['existing_metrics']) for p in reviews))

    def test_saved_packet_uses_its_own_sheet_for_missing_definition_review(self):
        path = self.mapping['packet_mappings'][0]['path']
        packet = load_fact_mapping(path, self.engine.standard)
        before = file_hash(path)
        self.assertTrue(packet['unresolved'])
        self.assertNotIn('sheet', packet['unresolved'][0])
        result = self.api.review_missing_standard(self.engine, packet, self.root/'原包补充', publish_root=self.root)
        self.assertEqual(result['new_ids'], ['固定资产累计折旧'])
        self.assertEqual(file_hash(path), before)
        self.assertEqual(result['reviews'][0]['sheet'], '表')

    def test_repeated_definition_text_is_preserved_in_proposal_and_both_reviews(self):
        rows = records()
        shared = '这段合成通用定义用于核实原文完整保留，各项维度的独立含义仍须分别读取。' * 8
        for row in rows:
            if row['record_type'] == 'dimension': row['meaning'] += ' ' + shared
        gold = self.root/'含重复段落的标准.jsonl'
        gold.write_text('\n'.join(json.dumps(row, ensure_ascii=False) for row in rows), encoding='utf-8-sig')
        engine = ControlledEngine(self.engine.settings, gold, self.root/'重复段落识别'); engine.calls = []
        mapping = engine.recognize_workbook(self.source)
        before = file_hash(gold)
        result = self.api.review_missing_standard(engine, mapping, self.root/'重复段落补充', publish_root=self.root)
        requests = [p for p in engine.calls if p['task'].startswith('unified_standard_')]
        self.assertEqual(len(requests), 3)
        expected = [row for row in rows if row['record_type'] == 'dimension']
        for payload in requests:
            self.assertTrue(payload.get('shared_dimension_rules'))
            restored = copy.deepcopy(payload['existing_dimensions'])
            for row in restored:
                if 'meaning_parts' in row:
                    row['meaning'] = ' '.join(payload['shared_dimension_rules'][part['rule']]
                                             if isinstance(part, dict) else part for part in row.pop('meaning_parts'))
            self.assertEqual(restored, expected)
            self.assertEqual(payload['candidate_cells'], ['B3'])
            self.assertEqual(payload['cells']['B3']['value'], None)
        self.assertEqual(result['new_ids'], ['固定资产累计折旧'])
        self.assertEqual(read_standard_v2(result['path'])['dimension'], engine.standard['dimension'])
        self.assertEqual(file_hash(gold), before)

    def test_reviewer_rejects_duplicate_or_instance_without_publishing(self):
        self.engine.deny = True
        result = self.api.review_missing_standard(self.engine, self.mapping, self.root/'补充', publish_root=self.root)
        self.assertEqual(result['path'], str(self.gold))
        self.assertEqual(result['new_ids'], [])
        self.assertFalse((self.root/'金标准/版本').exists())
        self.assertTrue(result['reviews'])
        for path, digest in self.original.items(): self.assertEqual(file_hash(path), digest)

    def test_complete_dimension_disagreement_is_reviewed_without_treating_missing_context_as_new_definition(self):
        class DisagreementEngine(ControlledEngine):
            missing_context = False

            def _http(self, system, payload):
                if payload['task'] == 'unified_standard_proposal':
                    self.calls.append(copy.deepcopy(payload))
                    return {'records': [], 'decisions': [
                        {'cell': cell, 'decision': 'context', 'reason': '不能仅凭识别分歧认定标准有错',
                         'record_ids': [], 'evidence_cells': ['A1']}
                        for cell in payload['candidate_cells']]}
                answer = super()._http(system, payload)
                row = answer['cells'].get('B2')
                if row and row['kind'] == 'metric_fact':
                    if self.missing_context:
                        row.update(kind='partial_metric_fact', missing_dimensions=['period'], evidence_cells=['A1'])
                        row['dimensions'].pop('period')
                        row['dimension_evidence'].pop('period')
                    else:
                        row['dimensions']['period']['date'] = '2024-12-31' if payload['round'] == 1 else '2025-01-01'
                return answer

        for missing in (False, True):
            with self.subTest(missing_context=missing):
                source = self.root/('缺少日期.xlsx' if missing else '日期解释分歧.xlsx')
                book = Workbook(); sheet = book.active; sheet.title = '表'
                sheet['A1'] = '合成企业甲；单户人民币元；全部固定资产；' + (
                    '日期未说明' if missing else '上年末2024年12月31日、本年初2025年1月1日')
                sheet['A2'] = '固定资产原值'; sheet['B2'] = 125
                book.save(source); book.close()
                before = file_hash(source)
                engine = DisagreementEngine(self.engine.settings, self.gold, self.root/('识别缺失' if missing else '识别分歧'))
                engine.calls = []; engine.missing_context = missing
                mapping = engine.recognize_workbook(source)
                self.assertEqual(len(mapping['unresolved']), 1)
                self.assertIn('partial_semantics', mapping['unresolved'][0])
                result = self.api.review_missing_standard(engine, mapping, self.root/('补充缺失' if missing else '补充分歧'), publish_root=self.root)
                requests = [p for p in engine.calls if p['task'] == 'unified_standard_proposal']
                self.assertEqual(len(requests), 0 if missing else 1)
                if not missing:
                    self.assertEqual(requests[0]['semantic_review_issues'], [
                        {'cell': 'B2', 'metric_id': '固定资产原值', 'disputed_dimensions': ['period']}])
                    self.assertIsNone(requests[0]['cells']['B2']['value'])
                self.assertEqual(result['new_ids'], [])
                self.assertEqual(result['changed_ids'], [])
                self.assertEqual(result['path'], str(self.gold))
                self.assertEqual(file_hash(source), before)
                self.assertEqual(load_fact_mapping(mapping['mapping_path'], engine.standard)['facts'], [])

    def test_model_cannot_insert_a_financial_value_into_definition(self):
        self.engine.malicious = True
        result = self.api.review_missing_standard(self.engine, self.mapping, self.root/'补充', publish_root=self.root)
        self.assertEqual(result['new_ids'], [])
        self.assertFalse((self.root/'金标准/版本').exists())
        self.assertTrue(any(row.get('error') for row in result['reviews']))
        for path, digest in self.original.items(): self.assertEqual(file_hash(path), digest)

    def test_formal_recognition_step_returns_new_standard_and_completed_mapping(self):
        from 附注更新.分步流程 import run_step
        from 附注更新.金标准 import publish_revision
        ControlledEngine.calls = []
        def isolated_publish(current, candidate, root=None, expected_hashes=None):
            return publish_revision(current, candidate, root=self.root, expected_hashes=expected_hashes)
        config = {'mode': 'online', 'model': '受控模型', 'base_url': 'http://localhost', 'api_key': '受控测试',
                  'gold_path': str(self.gold), 'b_path': str(self.source), 'output_dir': str(self.root/'正式入口')}
        with patch('附注更新.统一识别.UnifiedSemanticEngine', ControlledEngine), \
             patch('附注更新.统一补充.publish_revision', isolated_publish):
            result = run_step('recognize_b', config)
        self.assertEqual(result['status'], 'complete', result)
        self.assertIn('gold_path', result['files'])
        self.assertIn('b_gold_changes', result['files'])
        mapped = load_fact_mapping(result['files']['b_mapping'], read_standard_v2(result['files']['gold_path']))
        self.assertTrue(mapped['complete'])
        self.assertEqual(len(mapped['facts']), 2)
        self.assertTrue(Path(result['files']['b_review']).is_file())

    def test_changed_dimension_cannot_reuse_legacy_conversion_approval(self):
        rows = records()
        legacy = {'record_type': 'legacy_conversion', 'id': '转换:旧固定资产原值', 'source_gold_hash': 'a'*64,
                  'legacy_id': '旧固定资产原值', 'legacy_definition_hash': 'b'*64, 'status': 'approved',
                  'target': {'record_type': 'metric_fact', 'metric_id': '固定资产原值',
                             'dimension_bindings': {ref['id']: {'op': 'context', 'key': ref['id']} for ref in rows[-1]['dimension_refs']}},
                  'evidence_refs': ['来源一'], 'review': {'reviewer': '原复核者', 'reason': '基于旧集合定义的批准'}}
        rows.append(legacy); original = copy.deepcopy(rows)
        dimension = copy.deepcopy(next(row for row in rows if row['id'] == 'asset_selection'))
        dimension['meaning'] = '修订后的固定资产选取范围，原迁移必须按新含义重新核实'
        raw = {'records': [dimension], 'decisions': [{'cell': 'B3', 'decision': 'revision', 'reason': '原集合定义不充分',
               'record_ids': ['asset_selection'], 'evidence_cells': ['A3']}]}
        payload = {'source_hash': file_hash(self.source), 'mapping_hash': file_hash(self.mapping['mapping_path']), 'sheet': '表',
                   'candidate_cells': ['B3'], 'cells': {'A3': {'value': '固定资产类别及纳入范围'}}, 'text_evidence_cells': ['A3']}
        proposed = self.api._proposal(raw, payload, rows, self.root/'语义复核记录.json')
        revised = next(row for row in proposed['rows'] if row['id'] == legacy['id'])
        self.assertEqual(revised['status'], 'unresolved')
        self.assertIsNone(revised['target'])
        self.assertEqual(revised['legacy_definition_hash'], legacy['legacy_definition_hash'])
        self.assertEqual(rows, original)

    def test_partitionable_metric_without_scope_dimension_cannot_be_approved(self):
        metric = copy.deepcopy(records()[-1])
        metric['dimension_refs'] = [ref for ref in metric['dimension_refs'] if ref['id'] != 'asset_selection']
        payload = {'proposed_records': [metric], 'existing_metrics': [records()[-1]],
                   'existing_dimensions': [row for row in records() if row['record_type'] == 'dimension'],
                   'text_evidence_cells': ['A3']}
        raw = {'checks': [{'record_id': metric['id'], 'accepted': True, 'layering_correct': True, 'instance_only': False,
                           'existing_equivalent_ids': [], 'reason': '当前全体数据可识别，但遗漏可分组范围', 'evidence_cells': ['A3']}],
               'metric_scope_checks': [{'metric_id': metric['id'], 'can_partition': True, 'dimension_ids': [],
                                        'reason': '资产可按类别分组，不能固定全体'}]}
        with self.assertRaisesRegex(ValueError, '范围维度'):
            self.api._review(raw, payload)


if __name__ == '__main__':
    unittest.main()
