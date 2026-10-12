"""含义变化必须有逐条证据；历史链不能因纠错断开。"""
import unittest
from copy import deepcopy
import 统一整理 as 整理


class 纠错沿革(unittest.TestCase):
    def test_纠错有据承接并保留早期编号(self):
        self.assertTrue(hasattr(整理,'承接表现沿革'),'尚未实现有据纠错沿革')
        old={'representation_id':'old','source_kind':'five_style','source_locator':{'mapping_file':'a',
             'mapping_id':'m','sheet':'表','cell':'C3'},'semantic_id':'wrong',
             'dimension_values':{'period_role':'closing'},'规范维度值':{'period_role':'closing'}}
        new={**deepcopy(old),'representation_id':'new','semantic_id':'right'}
        proof={'纠错ID':'fix','结论':'已确认','旧表现摘要':整理.稳定摘要(old),
               '来源位置':整理.来源位置(old),'新semantic_id':'right','新业务维度':new['dimension_values'],
               '新规范维度':new['规范维度值'],'依据':'原行列与完整标准逐条核对'}
        chain=[{'旧表现ID':'early','新表现ID':'old','旧记录':{'representation_id':'early'}}]
        result=整理.承接表现沿革([old],[new],chain,[proof])
        self.assertEqual({r['旧表现ID']:r['新表现ID'] for r in result},{'old':'new','early':'new'})
        for kind in ('无依据','摘要错误','目标错误','多目标','同ID期间变更'):
            p=deepcopy(proof); n=deepcopy(new); proofs=[p];newrows=[n]
            if kind=='无依据':proofs=[]
            elif kind=='摘要错误':p['旧表现摘要']='bad'
            elif kind=='目标错误':p['新semantic_id']='other'
            elif kind=='多目标':newrows.append({**n,'representation_id':'another'})
            else:n['semantic_id']='wrong';n['dimension_values']['period_role']='opening';proofs=[]
            with self.subTest(kind=kind),self.assertRaises(ValueError):
                整理.承接表现沿革([old],newrows,chain,proofs)
        self.assertEqual(整理.承接表现沿革([new],[new],result,[proof]),result)

    def test_差值校验格退出业务仍保留全部编号(self):
        old={'representation_id':'old','source_kind':'five_style','source_locator':{
            'mapping_file':'a','mapping_id':'m','sheet':'表','cell':'I781'},
            'semantic_id':'错误余额','dimension_values':{},'规范维度值':{}}
        raw={'mapping_file':'a','mapping_id':'m','worksheet':'表','cell':'I781',
             'status':'结构格排除','raw_value':'=C781-G764','confirmation':{'纠错ID':'f'}}
        excluded=整理.结构排除证据(raw)
        proof={'纠错ID':'f','结论':'已确认','纠错类型':'业务格改为结构格排除',
               '旧表现摘要':整理.稳定摘要(old),'来源位置':整理.来源位置(old),
               '新排除证据摘要':整理.稳定摘要(excluded),'依据':'两表准备余额相减'}
        early=[{'旧表现ID':'early','新表现ID':'old','旧记录':{'representation_id':'early'}}]
        result=整理.承接表现沿革([old],[],early,[proof],[excluded])
        self.assertEqual({r['旧表现ID'] for r in result},{'old','early'})
        self.assertTrue(all(r['新表现ID'] is None and
                           r['新排除证据ID']==excluded['排除证据ID'] for r in result))
        self.assertEqual(整理.承接表现沿革([],[],result,[proof],[excluded]),result)
        for kind in ('无凭证','旧记录篡改','排除公式篡改','位置篡改','双重去向','排除证据丢失'):
            p=deepcopy(proof);ex=deepcopy(excluded);proofs=[p];active=[];contexts=[ex]
            if kind=='无凭证':proofs=[]
            elif kind=='旧记录篡改':p['旧表现摘要']='错误'
            elif kind=='排除公式篡改':ex['raw_value']='=C781'
            elif kind=='位置篡改':ex['cell']='J781'
            elif kind=='双重去向':active=[old]
            else:contexts=[]
            with self.subTest(kind=kind),self.assertRaises(ValueError):
                整理.承接表现沿革([old],active,early,proofs,contexts)
        with self.assertRaises(ValueError):
            整理.承接表现沿革([],[],result,[proof],[])

    def test_缺失维度改变不能无依据迁移(self):
        old={'representation_id':'old','source_kind':'five_style','source_locator':{'mapping_file':'a',
             'mapping_id':'m','sheet':'表','cell':'C3'},'semantic_id':'metric.a',
             'dimension_values':{},'规范维度值':{},'占位与缺失':[{'dimension':'period'}]}
        new={**deepcopy(old),'representation_id':'new','占位与缺失':[]}
        with self.assertRaisesRegex(ValueError,'纠错依据'):
            整理.承接表现沿革([old],[new],[],[])

if __name__=='__main__':unittest.main(verbosity=2)
