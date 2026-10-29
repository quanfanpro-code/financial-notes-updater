# coding: utf-8
"""公式引用同一批文字时只发送一次，保留来源、层级和适用范围。"""
import json
import unittest
from 附注更新.A语义阅读 import source_material
from 测试.测试引用来源语义 import 原格


class 公式共用原文测试(unittest.TestCase):
    def test_语义阅读保留直接引用及文字标题链但不递归数字计算(self):
        from copy import deepcopy
        snapshot={'path':'小例.xlsx','semantic_formula_sources':4,'sheets':[
            {'name':'展示','merges':[],'cells':{'A2':原格(2,1,'业务项目'),
                'B2':原格(2,2,formula="='汇总'!B2")}},
            {'name':'汇总','merges':['B1:C1'],'cells':{
                'A2':原格(2,1,'实际对象'),
                'B1':{**原格(1,2,formula="='信息'!A1"),'cached_value':'本期标题'},
                'B2':原格(2,2,formula="=SUM('底稿'!B2:B301)")}},
            {'name':'信息','merges':[],'cells':{'A1':原格(1,1,'本期标题')}},
            {'name':'底稿','merges':[],'cells':{f'B{r}':原格(r,2,r) for r in range(2,302)}}]}
        tables=[{'sheet':'展示','range':'A2:B2','data_ranges':['B2']}]
        material,_=source_material(snapshot,[],tables,semantic_only=True)
        proof=material['公式引用依据'];text=json.dumps(proof,ensure_ascii=False)
        self.assertIn('实际对象',text);self.assertIn('本期标题',text)
        self.assertIn('信息!A1',proof['原文']);self.assertIn('汇总!B2',proof['原文'])
        self.assertEqual(proof['原文']['汇总!B2']['公式'],"=SUM('底稿'!B2:B301)")
        self.assertFalse('底稿!B301' in proof['原文'])
        self.assertIn('汇总!B2',proof['未递归数值计算'])
        self.assertIn('行列文字规则',proof)
        legacy=deepcopy(snapshot);legacy['semantic_formula_sources']=3
        original,_=source_material(legacy,[],tables,semantic_only=True)
        self.assertIn('底稿!B301',original['公式引用依据']['原文'])
        self.assertEqual(material['公式原文'],original['公式原文'])

    def test_整段引用不重复携带共同文字(self):
        cells = {'A1': 原格(1, 1, '父项共同说明')}
        cells.update({f'B{row}': 原格(row, 2, f'分类项目{row:03}的原始说明') for row in range(2, 202)})
        snapshot = {'path': '小例.xlsx', 'semantic_formula_sources': 3, 'sheets': [
            {'name': '展示', 'merges': [], 'cells': {'A2': 原格(2, 1, '合计'),
                'B2': 原格(2, 2, formula="=SUM('分类'!B2:B201)")}},
            {'name': '分类', 'merges': ['A1:A201'], 'cells': cells}]}
        material, _ = source_material(snapshot, [], [{'sheet': '展示', 'range': 'A2:B2',
            'data_ranges': ['B2']}], semantic_only=True)
        proof = material['公式引用依据']
        text = json.dumps(proof, ensure_ascii=False)
        self.assertEqual(proof['引用']['展示!B2'], [{'范围引用': '分类!B2:B201'}])
        self.assertEqual(len(proof['引用范围']['分类!B2:B201']), 200)
        self.assertEqual(text.count('父项共同说明'), 1)
        for row in range(2, 202):
            self.assertEqual(text.count(f'分类项目{row:03}的原始说明'), 1)
        self.assertIn('分类!B201', proof['原文'])
        self.assertLess(len(text), 65000)
        snapshot['semantic_formula_sources'] = 2
        legacy, _ = source_material(snapshot, [], [{'sheet': '展示', 'range': 'A2:B2',
            'data_ranges': ['B2']}], semantic_only=True)
        self.assertEqual(len(legacy['公式引用依据']['引用']['展示!B2']), 200)
        self.assertEqual(legacy['公式引用依据']['原文']['分类!B201']['文字'], '分类项目201的原始说明')
        self.assertNotIn('共同行列文字', legacy['公式引用依据'])


if __name__ == '__main__':
    unittest.main()
