# coding: utf-8
"""初选之外的实际绑定指标须进入核查材料，并能正式保存回读。"""
import copy
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from openpyxl import Workbook
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一勾稽 import load_fact_mapping
from 测试.测试统一语义 import records, fact


class 核查补选标准检查(unittest.TestCase):
    def test_补选标准及维度传给核查且保存回读(self):
        folder = Path(tempfile.mkdtemp(prefix='核查补选标准_'))
        rows = records()
        extra = copy.deepcopy(rows[-1])
        extra['id'] = '固定资产账面价值'
        extra['definition'].update(measure='账面价值', measurement_basis='原值扣除累计折旧及减值准备')
        rows.append(extra)
        gold = folder / '合成标准.jsonl'
        gold.write_text('\n'.join(json.dumps(row, ensure_ascii=False) for row in rows), encoding='utf-8')
        source = folder / '合成原表.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = '原表'
        sheet.append(['合成企业甲，单户人民币元，2025年12月31日固定资产账面价值合计', '期末'])
        sheet.append(['固定资产账面价值合计', 50])
        book.save(source)
        book.close()
        table = {'id': 't', 'sheet': '原表', 'range': 'A1:B2', 'data_ranges': ['B2'],
                 'subject': '固定资产', 'meaning': '实际列示扣除折旧及减值后的固定资产账面价值'}
        reviews = []

        def reply(engine, system, payload):
            if payload['task'] == 'a_semantic_read':
                return {'tables': [table], 'notes': '公司、单位与时点见A1'}
            if payload['task'] == 'a_semantic_select':
                return {'tables': [{'id': 't', 'metric_ids': ['固定资产原值'], 'missing': ''}]}
            if payload['task'] == 'a_semantic_bind':
                return {'tables': [{'id': 't', 'metrics': [{'ranges': ['B2'], 'id': extra['id']}],
                    'dimensions': [{'ranges': ['B2'], 'id': key, 'value': value, 'evidence': ['A1']}
                                   for key, value in fact()['dimensions'].items()], 'unresolved': []}]}
            self.assertEqual(payload['task'], 'a_semantic_review')
            reviews.append(payload)
            return {'tables': [{'id': 't', 'corrections': {}, 'issues': []}]}

        engine = UnifiedSemanticEngine({'model': '合成回执', 'base_url': 'http://unused.local'}, gold, folder / '回执')
        with patch.object(UnifiedSemanticEngine, '_http', reply):
            result = engine.recognize_workbook(source, artifact_metadata={'carrier': 'A'})
        self.assertEqual(len(reviews), 1)
        supplied = {item['id']: item for item in reviews[0]['standards']['metrics']}
        self.assertIn(extra['id'], supplied, '核查材料漏掉识别时补选的标准')
        self.assertEqual(supplied[extra['id']]['definition'], extra['definition'])
        self.assertEqual({item['id'] for item in reviews[0]['standards']['dimensions']},
                         {item['id'] for item in extra['dimension_refs']})
        loaded = load_fact_mapping(result['mapping_path'], engine.standard)
        self.assertTrue(loaded['complete'])
        self.assertEqual(loaded['facts'][0]['metric_id'], extra['id'])


if __name__ == '__main__':
    unittest.main()
