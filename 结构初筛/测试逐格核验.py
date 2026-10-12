"""固定逐格业务预期的拒错能力，不调用比较器生成答案。"""
import unittest
from copy import deepcopy
import 核验真实样本 as 核验


class 逐格核验(unittest.TestCase):
    def test_正常及各类破坏(self):
        expected = {'旧表编号':'t','样式':1,'旧工作表':'旧','旧坐标':'C3','新工作表':'新','新坐标':'D4',
                    '语义引用':{'mapping_id':'m','representation_id':'r','semantic_id':'s',
                               '定义旧ID':'s','维度':{'period_role':'closing','category':['甲','乙']},
                               '确认依据':{'纠错记录':'有原件依据'}}}
        d = {k:expected[k] for k in ('旧工作表','旧坐标','新工作表','新坐标','语义引用')}
        d.update(角色='业务数值',待通读确认=True)
        good = {'候选':[{'旧表编号':'t','样式来源':1,'新工作表':'新','新表范围':'A1:D5',
                        '结构对应':[d],'歧义':[]}]}
        self.assertTrue(hasattr(核验,'核验逐格'), '尚未实现严格逐格核验')
        self.assertEqual(核验.核验逐格(good,[expected])[0]['结论'],'通过')
        for change in ('空','漏表名','期间','范围','目标','依据','待确认','第二落点'):
            r = deepcopy(good); c = r['候选'][0]; item = c['结构对应'][0]
            if change == '空': c['结构对应']=[]
            elif change == '漏表名': c['新工作表']=item['新工作表']='错表'
            elif change == '期间': item['语义引用']['维度']['period_role']='opening'
            elif change == '范围': item['语义引用']['维度']['category']=['甲']
            elif change == '目标': item['语义引用']['semantic_id']='other'
            elif change == '依据': item['语义引用']['确认依据']={}
            elif change == '待确认': item['待通读确认']=False
            else: c['结构对应'].insert(0,{**deepcopy(item),'新坐标':'Z99'})
            with self.subTest(change=change):
                self.assertEqual(核验.核验逐格(r,[expected])[0]['结论'],'失败')
        e = {**expected,'允许其他落点':[['新','Z99']]}
        r=deepcopy(good); other=deepcopy(r['候选'][0]);other['结构对应'][0]['新坐标']='Z99'
        r['候选'].append(other)
        self.assertEqual(核验.核验逐格(r,[e])[0]['结论'],'失败')
        other['歧义']=[{'不同对应':[{'旧坐标':'C3','本候选新坐标':'Z99','另一候选新坐标':'D4'}]}]
        checked=核验.核验逐格(r,[e])[0]
        self.assertEqual(checked['结论'],'通过')
        self.assertFalse(checked['唯一落点'])


if __name__ == '__main__': unittest.main(verbosity=2)
