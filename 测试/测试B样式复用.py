# coding: utf-8
"""用实际五样式和实际B核对语义复用，不把结构相似当业务相同。"""
import copy
import json
import unittest
from pathlib import Path


class B样式复用测试(unittest.TestCase):
    def test_旧语义转为当前指标且只补当前上下文(self):
        from 附注更新.B样式复用 import reuse_templates, apply_reuse
        from 附注更新.统一语义 import read_standard_v2
        root=Path(__file__).resolve().parents[1]
        standard=read_standard_v2(root/'金标准/版本/财务报表附注语义金标准_20261009_100733_874487_5a72e839.jsonl')
        comparison=json.loads(Path(r'C:\Users\27651\Documents\Codex\2026-10-08\du\outputs\B实际金标准绑定_20261008_204954\实际请求与回执\B机械比较_c41222115aca4321b5561747f68b927c.json').read_text(encoding='utf-8-sig'))
        templates=reuse_templates(comparison,standard)
        template=next(t for t in templates if t['location']=='word附注!B4' and t['row_semantics']=='库存现金')
        self.assertEqual(template['metric_id'],'metric.monetary_funds.balance')
        self.assertEqual(template['constant_dimensions']['cash_selection']['members'],['库存现金'])
        self.assertNotIn('period',template['constant_dimensions'])
        self.assertIn('period',template['context_bindings'])
        table={'id':'货币资金','sheet':'word附注','range':'A2:C8','data_ranges':['B4','B5']}
        row={'id':'货币资金','metrics':[{'ranges':['B5'],'id':'metric.monetary_funds.balance'}],
             'dimensions':[{'ranges':['*'],'id':'period','value':{'kind':'instant','date':'2026-12-31'},'evidence':['B3']}],
             'style_reuse':[{'template_id':template['id'],'evidence':['A2','A4','B3'],
                             'reason':'本表货币资金中库存现金含义与旧表一致；使用本次实际期间。'}]}
        result=apply_reuse(row,table,templates)
        self.assertEqual(len(result['metrics']),2)
        self.assertEqual(result['dimensions'][0]['value']['date'],'2026-12-31')
        self.assertEqual(result['dimensions'][-1]['value']['members'],['库存现金'])
        bad=copy.deepcopy(row);bad['style_reuse'][0]['template_id']='不存在'
        with self.assertRaises(ValueError):apply_reuse(bad,table,templates)
        bad=copy.deepcopy(row);bad['style_reuse'][0]['evidence']=['B4']
        with self.assertRaises(ValueError):apply_reuse(bad,table,templates)
        duplicate=copy.deepcopy(row);duplicate['metrics'].append({'ranges':['B4'],'id':'metric.monetary_funds.balance'})
        with self.assertRaises(ValueError):apply_reuse(duplicate,table,templates)


if __name__=='__main__':unittest.main()
