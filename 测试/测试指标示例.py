"""已有指标的模板表现须保留完整类别及实际缺项。"""
import unittest
from copy import deepcopy
from 测试统一语义 import records
from 附注更新.统一语义 import validate_records
import 附注更新.人工映射复用 as 复用


class 指标示例(unittest.TestCase):
    def test_模板缺日期不丢类别且非法目标维度拒绝(self):
        self.assertTrue(hasattr(复用,'metric_template'), '尚未支持已有指标模板')
        standard=validate_records(records())
        row={'target_metric_id':'固定资产原值','target_standard_sha256':'a'*64,
             'dimensions':{'asset_selection':{'domain':'固定资产','mode':'members',
                 'members':['甲设备','乙设备'],'completeness':'complete'}},
             'dimension_evidence':{'asset_selection':['原表!A3']},'worksheet':'原表','cell':'C3',
             'template_period_role':'closing','confirmation':{'说明':'原表完整证据'}}
        out=复用.metric_template(row,standard,'a'*64)
        self.assertEqual(out['dimensions'],row['dimensions'])
        self.assertIn('period',out['missing_dimensions'])
        for kind in ('双目标','错版本','无依据','非法维度','停用'):
            r=deepcopy(row);s=deepcopy(standard)
            if kind=='双目标':r['target_slot_id']='old'
            elif kind=='错版本':r['target_standard_sha256']='b'*64
            elif kind=='无依据':r['confirmation']={}
            elif kind=='非法维度':r['dimensions']['invented']='x'
            else:s['metric']['固定资产原值']['status']='retired'
            with self.subTest(kind=kind),self.assertRaises(ValueError):复用.metric_template(r,s,'a'*64)

if __name__=='__main__':unittest.main(verbosity=2)
