"""局部验证保留整份未决，并按实际范围及本次用途建立对应。"""
import copy
import unittest
from 测试.测试统一勾稽 import mapping
from 测试.测试统一语义 import records
from 附注更新.统一语义 import validate_records
from 附注更新 import 统一勾稽 as api


class 局部范围测试(unittest.TestCase):
    def test_局部选择不伪造整份完成也不丢范围内未决(self):
        source = mapping('来源!B2', 10)
        source.update(complete=False, coverage={'complete': False, 'expected_cells': ['来源!B2','来源!B3','来源!D9']},
                      unresolved=[{'sheet':'来源','cell':'B3','reason':'类别未明'}, {'sheet':'来源','cell':'D9','reason':'尚未处理'}])
        original = copy.deepcopy(source)
        view = api.select_mapping_scope(source, [{'sheet':'来源','range':'B2:B3'}])
        self.assertEqual(source, original)
        self.assertFalse(view['complete'])
        self.assertFalse(view['scope_coverage']['complete'])
        self.assertEqual(view['scope_coverage']['expected_count'], 2)
        self.assertEqual(view['unresolved'], [source['unresolved'][0]])
        self.assertEqual(view['scope_coverage']['outside_unresolved_count'], 1)

    def test_局部范围遗漏分类不能声明通过(self):
        source = mapping('来源!B2', 10)
        source['coverage'] = {'complete':False,'expected_cells':['来源!B2','来源!B3']}
        with self.assertRaisesRegex(ValueError, '遗漏'):
            api.select_mapping_scope(source, [{'sheet':'来源','range':'B2:B3'}])

    def test_模板用途单独保存原主体且不改变来源事实(self):
        standard=validate_records(records());standard['sha256']='a'*64
        target=mapping('附注!D8', 7);source=mapping('来源!J3', -99)
        source['facts'][0]['dimensions']['entity']='合成企业乙';source['owners']['对象']['entity']='合成企业乙'
        original=copy.deepcopy(target)
        purpose={'entity':{'from':'合成企业甲','to':'合成企业乙'},'reason':'用户确认借用甲模板填乙单户数据'}
        plan=api.build_fact_update_plan(target,source,standard,update_purpose=purpose)
        self.assertEqual(target,original)
        self.assertEqual(plan['updates'][0]['value'],-99)
        self.assertEqual(plan['updates'][0]['original_dimensions']['entity'],'合成企业甲')
        self.assertEqual(plan['updates'][0]['dimensions']['entity'],'合成企业乙')
        source['facts'][0]['dimensions']['period']['date']='2023-12-31'
        source['owners']['对象']=copy.deepcopy(source['facts'][0]['dimensions'])
        self.assertEqual(api.build_fact_update_plan(target,source,standard,update_purpose=purpose)['updates'],[])


if __name__=='__main__':
    unittest.main()
