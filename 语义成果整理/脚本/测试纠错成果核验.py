"""固定纠错凭证必须落实到表现及完整沿革。"""
import unittest
from copy import deepcopy
import 核验整理 as 核验
import 统一整理 as 整理

class 纠错成果核验(unittest.TestCase):
    def test_退出业务后逐格验证上下文和历史去向(self):
        ex=整理.结构排除证据({'mapping_file':'源','mapping_id':'m','worksheet':'表',
            'cell':'I781','status':'结构格排除','raw_value':'=C781-G764','confirmation':{'纠错ID':'f'}})
        proof={'纠错ID':'f','纠错类型':'业务格改为结构格排除',
               '来源位置':ex['来源位置'],'旧表现':{'representation_id':'old'},
               '新排除证据摘要':整理.稳定摘要(ex),'新映射':{'confirmation':ex['confirmation']}}
        chain=[{'旧表现ID':'old','新表现ID':None,'新排除证据ID':ex['排除证据ID'],'纠错引用':'f'}]
        self.assertEqual(核验.核验纠错结果([proof],[],chain,[ex]),[])
        for contexts,history,active in (([],chain,[]),([ex],[],[]),([ex,ex],chain,[]),
               ([{**ex,'raw_value':'=C781'}],chain,[]),
               ([ex],chain,[{'source_locator':{'mapping_file':'源','mapping_id':'m'}}])):
            self.assertTrue(核验.核验纠错结果([proof],active,history,contexts))

    def test_目标期间与沿革不能漏检(self):
        self.assertTrue(hasattr(核验,'核验纠错结果'),'缺少纠错成果逐条核验')
        location={'mapping_file':'源','mapping_id':'m','sheet':'表','cell':'C3'}
        rep={'representation_id':'new','source_locator':location,'semantic_id':'正确目标',
             'dimension_values':{'period_role':'closing'},'规范维度值':{'period_role':'closing'},
             '确认依据':{'纠错ID':'f'}}
        proof={'纠错ID':'f','来源位置':{'来源':'源','mapping_id':'m','sheet':'表','cell':'C3'},
               '旧表现':{'representation_id':'old'},'新semantic_id':'正确目标',
               '新业务维度':rep['dimension_values'],'新规范维度':rep['规范维度值'],
               '新映射':{'confirmation':rep['确认依据']}}
        chain=[{'旧表现ID':'old','新表现ID':'new','纠错引用':'f'}]
        self.assertEqual(核验.核验纠错结果([proof],[rep],chain),[])
        wrong=deepcopy(rep);wrong['规范维度值']['period_role']='opening'
        self.assertTrue(核验.核验纠错结果([proof],[wrong],chain))
        self.assertTrue(核验.核验纠错结果([proof],[rep],[]))

if __name__=='__main__':unittest.main(verbosity=2)
