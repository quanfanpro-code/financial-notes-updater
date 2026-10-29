# coding: utf-8
"""共用核查的格式由实际标准约束，不代替模型判断业务或取值。"""
from copy import deepcopy
import unittest
from jsonschema import Draft202012Validator

from 附注更新.A语义阅读 import standard_material
from 附注更新.回执格式 import response_format
from 附注更新.设置 import load_settings
from 附注更新.统一语义 import read_standard_v2


class A核查回执格式测试(unittest.TestCase):
    def setUp(self):
        standard=read_standard_v2(load_settings()['gold_path'])
        self.payload={'task':'a_semantic_review','reading':[{'id':'任意表号'}],
                      'standards':standard_material(['metric.inventory.gross_balance'],standard)}
        self.reply={'tables':[{'id':'任意表号','corrections':{
            'metrics':[],'patches':[{'id':'period','ranges':['C8'],
                'value':{'kind':'instant','date':'2037-09-30'},'evidence':['B7']}],
            'resolved_ranges':[],'non_business':[],'remove_dimensions':[]},'issues':[]}]}

    def test_日期字符串被结构约束拒绝但模型给出的实际日期结构可用(self):
        form=response_format(self.payload)
        self.assertEqual(form['type'],'json_schema')
        schema=form['json_schema']['schema'];Draft202012Validator.check_schema(schema)
        validator=Draft202012Validator(schema)
        self.assertFalse(list(validator.iter_errors(self.reply)))
        wrong=deepcopy(self.reply);wrong['tables'][0]['corrections']['patches'][0]['value']='2037-09-30'
        self.assertTrue(list(validator.iter_errors(wrong)))

    def test_只约束本次给出的编码和维度类型不填写业务值(self):
        schema=response_format(self.payload)['json_schema']['schema']
        self.assertNotIn('2037',str(schema))
        wrong=deepcopy(self.reply);wrong['tables'][0]['corrections']['patches'][0]['id']='不存在的维度'
        self.assertTrue(list(Draft202012Validator(schema).iter_errors(wrong)))

    def test_其他阅读步骤保持原有协议(self):
        self.assertEqual(response_format({'task':'a_semantic_read'}),{'type':'json_object'})

    def test_局部回执不能在更正或问题中加入范围外位置(self):
        self.payload['review_focus']={'任意表号':['C8:C9']}
        schema=response_format(self.payload)['json_schema']['schema'];validator=Draft202012Validator(schema)
        self.assertFalse(list(validator.iter_errors(self.reply)))
        wrong=deepcopy(self.reply)
        wrong['tables'][0]['issues']=[{'range':'D8','reason':'范围外疑问','evidence':['B7']}]
        self.assertTrue(list(validator.iter_errors(wrong)))
        wrong=deepcopy(self.reply);wrong['tables'][0]['corrections']['patches'][0]['ranges']=['D8']
        self.assertTrue(list(validator.iter_errors(wrong)))
