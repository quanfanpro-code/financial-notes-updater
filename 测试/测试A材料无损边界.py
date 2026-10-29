# coding: utf-8
"""第7节第二步的机械反例；不调用模型，不构成真实识别验收。"""
import copy
import collections
import json
import unittest
from unittest.mock import patch

from 附注更新.A语义阅读 import source_material, standard_material
from 附注更新.统一语义 import validate_records
from 测试.测试统一语义 import records, fact
from 测试.测试引用来源语义 import 原格
from 测试.测试Codex语义调用 import test_directory


class A材料无损边界测试(unittest.TestCase):
    def test_重复长定义逐字保留且不共享可变对象(self):
        standard = validate_records(records())
        paragraph = '本段定义必须保留其完整条件及明确例外。' * 10
        for i in range(10):
            standard['dimension']['参考' + str(i)] = {'meaning': paragraph}
        standard['dimension']['asset_selection']['meaning'] = '原职责  ' + paragraph + ' \n例外保持。'
        before = copy.deepcopy(standard)
        material = standard_material(['固定资产原值'], standard)
        dimension = next(d for d in material['dimensions'] if d['id'] == 'asset_selection')
        self.assertEqual(dimension['meaning'], standard['dimension']['asset_selection']['meaning'])
        material['metrics'][0]['dimension_refs'][0]['required'] = False
        self.assertEqual(standard, before)

    def test_比例定义及其依赖闭合(self):
        standard = validate_records(records())
        ratio = copy.deepcopy(standard['metric']['固定资产原值'])
        ratio.update(id='比例')
        ratio['definition']['ratio'] = {
            'numerator': {'metric_id': '固定资产原值', 'dimension_bindings': {}},
            'denominator': {'metric_id': '固定资产原值', 'dimension_bindings': {}}}
        standard['metric']['比例'] = ratio
        material = standard_material(['比例'], standard)
        self.assertEqual({m['id'] for m in material['metrics']}, {'比例', '固定资产原值'})
        self.assertEqual(material['metrics'][0]['definition'], ratio['definition'])

    def test_数值缓存不截断下层公式与原文(self):
        snapshot = {'path': '机械小例.xlsx', 'semantic_formula_sources': 5, 'sheets': [
            {'name': '展示', 'merges': [], 'cells': {
                'A2': 原格(2, 1, '项目'), 'B2': 原格(2, 2, formula="='中间'!B2")}},
            {'name': '中间', 'merges': [], 'cells': {
                'A2': 原格(2, 1, '范围说明'), 'B2': 原格(2, 2, formula="='明细'!B2")}},
            {'name': '明细', 'merges': [], 'cells': {
                'A2': 原格(2, 1, '下层真实限定'), 'B2': 原格(2, 2, 17)}}]}
        tables = [{'sheet': '展示', 'range': 'A2:B2', 'data_ranges': ['B2']}]
        material, _ = source_material(snapshot, [], tables, semantic_only=True)
        proof = material['公式引用依据']
        self.assertIn('明细!B2', proof['原文'])
        self.assertIn('下层真实限定', json.dumps(proof, ensure_ascii=False))
        self.assertNotIn('未递归数值计算', proof)
        snapshot['sheets'][1]['cells']['B2']['cached_value'] = '缓存文字'
        changed, _ = source_material(snapshot, [], tables, semantic_only=True)
        self.assertEqual(proof, changed['公式引用依据'])

    def test_真实引用及标题公式保留而不展开范围外数值公式(self):
        snapshot={'path':'机械小例.xlsx','semantic_formula_sources':5,'sheets':[
            {'name':'展示','merges':[],'cells':{'B4':原格(4,2,formula="='中间'!B4")}},
            {'name':'中间','merges':['B1:C1'],'cells':{
                'B1':原格(1,2,formula="='标题'!A1"),
                'B2':原格(2,2,formula="='无关'!B2"),
                'B3':{**原格(3,2,formula="='标题'!A3"),'cached_value':'缓存标题'},
                'A4':原格(4,1,formula="='标题'!A4"),
                'B4':原格(4,2,formula="='明细'!B4"),
                'C4':原格(4,3,formula="='无关'!C4")}},
            {'name':'明细','merges':[],'cells':{'A4':原格(4,1,'真实限定'),'B4':原格(4,2,17)}},
            {'name':'标题','merges':[],'cells':{f'A{r}':原格(r,1,f'标题{r}') for r in (1,3,4)}},
            {'name':'无关','merges':[],'cells':{'B2':原格(2,2,19),'C4':原格(4,3,23)}}]}
        material,_=source_material(snapshot,[],[{'sheet':'展示','range':'B4','data_ranges':['B4']}])
        proof=material['公式引用依据']
        self.assertTrue({'展示!B4','中间!B4','中间!B1','中间!B3','中间!A4'}<=set(proof['引用']))
        self.assertTrue({'明细!B4','标题!A1','标题!A3','标题!A4'}<=set(proof['原文']))
        self.assertFalse({'中间!B2','中间!C4'}&set(proof['引用']))
        self.assertNotIn('无关!',json.dumps(proof,ensure_ascii=False))

    def test_缺少公式引用时普通请求入口不提交(self):
        from 附注更新.语义 import SemanticEngine
        engine = object.__new__(SemanticEngine)
        engine.cancel = None
        engine.settings = {'base_url': 'codex://local', 'model': 'gpt-6-luna'}
        engine.cache_source_hash = ''
        engine.source_hash = 'a' * 64
        engine.gold_hash = 'b' * 64
        engine.usage = collections.Counter()
        engine._review_receipts = []
        engine.log = lambda text: None
        payload = {'task': 'a_semantic_bind', 'source': {'公式引用依据': {
            '引用': {}, '原文': {}, '未展开': [{'位置': '表!B2', '引用': '未定义名称'}]}}}
        with test_directory() as directory, patch.object(engine, '_http', return_value={}):
            engine.work_dir = directory
            with self.assertRaisesRegex(ValueError, '公式.*未展开'):
                engine._request('未改动的提示词', payload, lambda x: x)

    def test_超限在启动Codex之前停止(self):
        from 附注更新.Codex语义调用 import call_codex
        with test_directory() as directory, \
             patch('附注更新.Codex语义调用.shutil.which', return_value=r'C:\本机\codex.exe'), \
             patch('附注更新.Codex语义调用.subprocess.Popen', side_effect=AssertionError('超限仍尝试启动模型')):
            with self.assertRaisesRegex(ValueError, '超过.*1048576'):
                call_codex({'model': 'gpt-6-luna'}, directory,
                           {'messages': [{'content': '字' * 1048577}]}, lambda: None)

    def test_完整性报告保留缺项未知和原分类(self):
        import 附注更新.统一语义 as semantics
        self.assertTrue(hasattr(semantics, 'identity_completeness'))
        row = fact()
        row['dimensions'].pop('report_scope')
        row['dimensions'].pop('entity')
        row['dimensions']['asset_selection'] = {
            'domain': '固定资产', 'mode': 'unknown', 'completeness': 'unknown'}
        before = copy.deepcopy(row)
        result = semantics.identity_completeness(row, validate_records(records()))
        self.assertEqual(result['missing_dimensions'], ['entity', 'report_scope'])
        self.assertEqual(result['unknown_dimensions'], ['asset_selection'])
        self.assertFalse(result['complete'])
        self.assertEqual(row, before)

    def test_动态地址公式保留为未展开(self):
        snapshot = {'path': '机械小例.xlsx', 'semantic_formula_sources': 5, 'sheets': [
            {'name': '展示', 'merges': [], 'cells': {
                'A2': 原格(2, 1, '项目'), 'B2': 原格(2, 2, formula='=INDIRECT("C2")'),
                'C2': 原格(2, 3, '不能猜该公式的实际引用')}}]}
        material, _ = source_material(snapshot, [], [{'sheet': '展示', 'range': 'A2:B2', 'data_ranges': ['B2']}])
        self.assertEqual(material['公式引用依据']['未展开'], [{'位置': '展示!B2', '引用': '=INDIRECT("C2")'}])

    def test_只有引用名而未给内容也停止(self):
        from 附注更新.语义 import SemanticEngine
        engine = object.__new__(SemanticEngine)
        engine.cancel = None
        engine.settings = {'base_url': 'codex://local', 'model': 'gpt-6-luna'}
        engine.cache_source_hash = ''
        engine.source_hash = 'a' * 64
        engine.gold_hash = 'b' * 64
        engine.usage = collections.Counter()
        engine._review_receipts = []
        engine.log = lambda text: None
        source = {'公式原文': {'表!B2': '=C2'}, '公式引用依据': {
            '引用': {'表!B2': ['表!C2']}, '原文': {}, '未展开': []}}
        with test_directory() as directory, patch.object(engine, '_http', return_value={}):
            engine.work_dir = directory
            with self.assertRaisesRegex(ValueError, '公式.*缺少'):
                engine._request('未改动的提示词', {'task': 'a_semantic_read', 'source': source}, lambda x: x)


if __name__ == '__main__':
    unittest.main()
