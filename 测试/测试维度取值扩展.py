# coding: utf-8
"""新增可选值保留旧语义；改义、删值及指标限制变化仍须重新识别。"""
import copy
import json
import unittest
from pathlib import Path
from uuid import uuid4

from 附注更新 import 统一语义 as api
from 测试.测试统一语义 import records, fact


class DimensionExpansionTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(__file__).resolve().parents[1] / '测试结果' / ('维度取值扩展_' + uuid4().hex)
        self.root.mkdir()
        self.rows = records()
        metric = self.rows[-1]
        metric['definition'].update(measure='变动金额', time_kind='duration')
        metric['dimension_refs'].append({'id': 'change_direction', 'required': True})
        self.rows.append({'record_type': 'dimension', 'id': 'change_direction',
                          'meaning': '原文明确的变动方向，不从金额正负推断',
                          'value_type': 'text', 'domain': '变动方向', 'allowed_values': ['增加', '减少']})
        path, old = self.save('原标准', self.rows)
        amount = fact()
        amount.update(id='金额格', owner_reference='对象一', binding_refs={'change_direction': '方向格'})
        amount['dimensions'].update(period={'kind': 'duration', 'start': '2025-01-01', 'end': '2025-12-31'}, change_direction='增加')
        amount['dimension_evidence']['change_direction'] = ['原表方向表头']
        binding = {'id': '方向格', 'record_type': 'dimension_binding', 'dimension_id': 'change_direction',
                   'value': '增加', 'owner_reference': '对象一', 'source_reference': {'file_hash': 'b' * 64, 'location': '当前表!C8'},
                   'recognition_provenance': {'method': 'synthetic', 'record_id': '合成方向'}}
        self.mapping = {'schema_version': 2, 'gold_path': str(path), 'gold_hash': old['sha256'],
                        'facts': [binding, amount], 'owners': {'对象一': copy.deepcopy(amount['dimensions'])}}
        api.validate_facts(self.mapping['facts'], old, self.mapping['owners'])

    def save(self, label, rows):
        path = self.root / (label + '.jsonl')
        with path.open('x', encoding='utf-8-sig') as stream:
            stream.write('\n'.join(json.dumps(row, ensure_ascii=False) for row in rows))
        return path, api.read_standard_v2(path)

    def test_added_choice_preserves_amount_binding_owner_and_original_value(self):
        rows = copy.deepcopy(self.rows)
        rows[-1]['allowed_values'].append('净变动')
        _, expanded = self.save('仅扩展取值', rows)
        self.assertTrue(api.ensure_standard_compatible(self.mapping, expanded))
        selected = api.compatible_classification(self.mapping, expanded)
        self.assertEqual(selected['facts'], self.mapping['facts'])
        self.assertEqual(selected['owners'], self.mapping['owners'])
        # 转置及更改金额只改变载体和原值，同一标准下业务身份仍一致。
        changed = copy.deepcopy(self.mapping['facts'][1])
        changed['raw_value'] = -321
        changed['source_reference']['location'] = '新布局!AB90'
        self.assertEqual(api.semantic_key(changed, expanded), api.semantic_key(self.mapping['facts'][1], expanded))

    def test_partial_meaning_remains_pending_after_choice_expansion(self):
        rows = copy.deepcopy(self.rows)
        rows[-1]['allowed_values'].append('净变动')
        _, expanded = self.save('扩展部分语义', rows)
        known = {key: copy.deepcopy(self.mapping['facts'][1][key]) for key in ('metric_id', 'dimensions', 'dimension_evidence')}
        del known['dimensions']['asset_selection']
        del known['dimension_evidence']['asset_selection']
        known.update(missing_dimensions=['asset_selection'], evidence_cells=['当前表!A1'])
        mapping = copy.deepcopy(self.mapping)
        mapping['unresolved'] = [{'partial_semantics': known}]
        self.assertTrue(api.ensure_standard_compatible(mapping, expanded))
        self.assertEqual(mapping['unresolved'][0]['partial_semantics']['missing_dimensions'], ['asset_selection'])

    def test_meaning_removal_and_metric_constraint_changes_are_not_expansions(self):
        for label in ('改义', '删除未用取值', '取消取值限制', '改变指标限制'):
            rows = copy.deepcopy(self.rows)
            if label == '改义':
                rows[-1]['meaning'] = '现金收付方向'
            elif label == '删除未用取值':
                rows[-1]['allowed_values'] = ['增加']
            elif label == '取消取值限制':
                del rows[-1]['allowed_values']
            else:
                rows[-2]['definition']['constraints'] = [{'dimension_id': 'change_direction', 'allowed_values': ['增加']}]
            _, changed = self.save(label, rows)
            with self.subTest(label=label):
                with self.assertRaisesRegex(ValueError, '定义已改变'):
                    api.ensure_standard_compatible(self.mapping, changed)
                selected = api.compatible_classification(self.mapping, changed)
                self.assertFalse(any(row['record_type'] == 'metric_fact' for row in selected['facts']))

    def test_ratio_cannot_reuse_changed_underlying_measurement(self):
        rows = copy.deepcopy(self.rows)
        ratio = copy.deepcopy(rows[-2])
        ratio['id'] = '变动占比'
        ratio['definition'].update(value_type='percentage', measure='比例', ratio={
            side: {'metric_id': rows[-2]['id'], 'dimension_bindings': {
                ref['id']: {'op': 'dimension', 'key': ref['id']} for ref in ratio['dimension_refs']}}
            for side in ('numerator', 'denominator')})
        rows.append(ratio)
        path, old = self.save('比例原标准', rows)
        mapping = copy.deepcopy(self.mapping)
        mapping.update(gold_path=str(path), gold_hash=old['sha256'])
        mapping['facts'][1]['metric_id'] = ratio['id']
        rows[-3]['definition']['measurement_basis'] = '账面净额'
        _, changed = self.save('比例底层改义', rows)
        with self.assertRaisesRegex(ValueError, '定义已改变'):
            api.ensure_standard_compatible(mapping, changed)
        self.assertFalse(any(row['record_type'] == 'metric_fact'
                             for row in api.compatible_classification(mapping, changed)['facts']))

    def test_formal_resume_reuses_expanded_dimension_without_model_requests(self):
        from openpyxl import Workbook
        from 附注更新.统一识别 import UnifiedSemanticEngine
        from 附注更新.统一勾稽 import load_fact_mapping

        source = self.root / '人工合成数据.xlsx'
        book = Workbook()
        book.active['A1'] = '合成企业甲，单户人民币元，2025年度固定资产原值增加金额'
        book.active['B2'] = -321
        book.save(source)
        book.close()
        calls = []
        amount = self.mapping['facts'][1]

        class ControlledEngine(UnifiedSemanticEngine):
            def _request(self, system, payload, validator):
                calls.append(payload['candidate_cells'])
                response = {}
                for address in payload['candidate_cells']:
                    if address == 'B2':
                        response[address] = {'kind': 'metric_fact', 'metric_id': amount['metric_id'],
                            'dimensions': copy.deepcopy(amount['dimensions']),
                            'dimension_evidence': {key: ['A1'] for key in amount['dimensions']},
                            'reason': '合成数据只验证正式接续，不代表真实模型识别准确率'}
                    else:
                        response[address] = {'kind': 'non_business', 'category': 'title' if address == 'A1' else 'empty_padding',
                                             'reason': '合成资料的表头或填充', 'evidence_cells': ['A1']}
                return validator({'cells': response})

        first = ControlledEngine({}, self.mapping['gold_path'], self.root / '首次识别').recognize_workbook(source)
        rows = copy.deepcopy(self.rows)
        rows[-1]['allowed_values'].append('净变动')
        path, expanded = self.save('正式接续扩展标准', rows)
        calls.clear()
        second = ControlledEngine({}, path, self.root / '接续识别').recognize_workbook(source, previous_result=first['mapping_path'])
        self.assertEqual(calls, [])
        self.assertTrue(second['complete'])
        self.assertEqual(second['facts'], first['facts'])
        loaded = load_fact_mapping(second['mapping_path'], expanded)
        self.assertEqual(loaded['facts'][0]['raw_value'], -321)


if __name__ == '__main__':
    unittest.main()
