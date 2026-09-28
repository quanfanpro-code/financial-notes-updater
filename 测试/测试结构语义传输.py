"""结构回执须能通过实际HTTP格式约束，并还原为逐格结果。"""
import unittest
from jsonschema import Draft202012Validator
from 附注更新.回执格式 import response_format, normalize_response, transport_instruction
from 测试.测试结构语义展开 import 结构展开测试


class 结构传输测试(unittest.TestCase):
    def test_共同维度与行列维度可通过传输协议(self):
        sample=结构展开测试();sample.setUp();payload=sample.payload
        payload['dimensions']=[{'id':key,'value_type':'text'} for key in ('period','category')]
        payload['metrics']=[{'id':'metric.balance','dimension_refs':[{'id':key,'required':True} for key in ('period','category')]}]
        schema=response_format(payload)['json_schema']['schema']
        Draft202012Validator.check_schema(schema)
        errors=list(Draft202012Validator(schema).iter_errors(sample.reply))
        self.assertEqual([error.message for error in errors],[])
        self.assertIn('regions',transport_instruction(payload))
        result=normalize_response(payload,sample.reply)
        self.assertEqual(result['cells']['C3']['dimensions'],{'period':'2024','category':'存款'})


if __name__=='__main__':unittest.main()
