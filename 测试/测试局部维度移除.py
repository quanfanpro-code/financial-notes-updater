# coding: utf-8
"""模型只移除所声明范围的维度，不能抹掉其他行的已确认内容。"""
from copy import deepcopy
import unittest
from 附注更新.A语义阅读 import validate_review,reviewed_binding


class 局部维度移除测试(unittest.TestCase):
    def setUp(self):
        self.table={'id':'表甲','sheet':'资料','range':'A1:B3','data_ranges':['B2:B3']}
        self.snapshot={'path':'原件.xlsx','sha256':'a'*64,'sheets':[{'name':'资料','merges':[],
            'cells':{a:{'row':r,'column':c,'value':v} for a,r,c,v in
                     [('A1',1,1,'指标'),('B1',1,2,'比例'),('A2',2,1,'类别一'),('B2',2,2,None),('A3',3,1,'类别二'),('B3',3,2,None)]}}]}
        self.binding={'tables':[{'id':'表甲','metrics':[],'unresolved':[],
            'dimensions':[{'id':key,'ranges':['B2:B3'],'value':value,'evidence':['B1']}
                          for key,value in [('currency','CNY'),('unit_scale',1)]]}]}
        self.answer={'tables':[{'id':'表甲','corrections':{'remove_dimensions':[{'id':'unit_scale','ranges':['B2']}]},'issues':[]}]}

    def test_只移除声明格并保留其余维度(self):
        before=deepcopy(self.binding)
        checked=validate_review(self.answer,[self.table],self.snapshot,[],{'表甲':['B2']})
        result=reviewed_binding(self.binding,checked,[self.table],self.snapshot)
        dims=result['tables'][0]['dimensions']
        self.assertEqual(next(x for x in dims if x['id']=='unit_scale')['ranges'],['B3'])
        self.assertEqual(next(x for x in dims if x['id']=='currency'),before['tables'][0]['dimensions'][0])
        self.assertEqual(self.binding,before)

    def test_局部核查不能接受整表或越界移除(self):
        for removal in ('unit_scale',{'id':'unit_scale','ranges':['B3']},{'id':'unit_scale'}):
            answer=deepcopy(self.answer);answer['tables'][0]['corrections']['remove_dimensions']=[removal]
            with self.subTest(removal=removal),self.assertRaises(ValueError):
                validate_review(answer,[self.table],self.snapshot,[],{'表甲':['B2']})

    def test_原整表移除仍能正常回读(self):
        answer=deepcopy(self.answer);answer['tables'][0]['corrections']['remove_dimensions']=['unit_scale']
        checked=validate_review(answer,[self.table],self.snapshot,[])
        result=reviewed_binding(self.binding,checked,[self.table],self.snapshot)
        self.assertEqual([x['id'] for x in result['tables'][0]['dimensions']],['currency'])
