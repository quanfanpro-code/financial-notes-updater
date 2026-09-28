"""正式识别使用已批准的语义表达；未决转换不能混作金标准知识。"""
import copy
from pathlib import Path
import unittest
from 附注更新.设置 import load_settings
from 附注更新.统一语义 import read_standard_v2
from 附注更新.统一识别 import build_payload


class 已审语义入口测试(unittest.TestCase):
    def test_已审对应进入实际请求而未决对应被排除(self):
        standard=read_standard_v2(load_settings()['gold_path'])
        packet={'sheet':'任意新表式','candidate_cells':['D8'],'cells':{
            'A1':{'value':'银行存款 期末余额'},'D8':{'value':123}}}
        request=build_payload(packet,standard)
        self.assertIn('standard_expressions',request)
        knowledge=request['standard_expressions']
        row=next(x for x in knowledge['expressions'] if x['legacy_id']=='C-N087-T001-S003')
        rule=knowledge['rules'][row['rule']]
        self.assertEqual(rule['metric_id'],'metric.monetary_funds.balance')
        self.assertEqual(rule['dimension_bindings']['cash_selection']['value']['members'],['银行存款'])
        self.assertTrue(row['manual_forms'])
        self.assertNotIn('raw_value',row)
        changed=copy.deepcopy(standard)
        conversion=changed['legacy_conversion'][row['conversion_id']]
        conversion.update(status='unresolved',target=None)
        updated=build_payload(packet,changed)['standard_expressions']
        self.assertNotIn('C-N087-T001-S003',{x['legacy_id'] for x in updated['expressions']})
        shifted={'sheet':'另一种布局','candidate_cells':['Z100'],'cells':{
            'G50':{'value':'银行存款 期末余额'},'Z100':{'value':-5}}}
        self.assertEqual(build_payload(shifted,standard)['standard_expressions'],knowledge)


if __name__=='__main__':unittest.main()
