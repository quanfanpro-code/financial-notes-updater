# coding: utf-8
"""空白、零值和金额不同均不得改变已确认的格子语义对应。"""
import copy
import unittest
from unittest.mock import patch
from openpyxl import load_workbook
from 测试.测试B范围定位 import B范围定位测试
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一勾稽 import load_fact_mapping
from 附注更新.统一语义 import semantic_key


class 空白格语义对应测试(unittest.TestCase):
    def test_双方空白和不同数值均完整保留身份及对应(self):
        identities = []
        for a_value, b_value in [(None,None), (None,0), (300,19)]:
            with self.subTest(A=a_value, B=b_value):
                case = B范围定位测试(); case.setUp()
                source = case.case.folder/'空白财务格.xlsx'
                book = load_workbook(case.case.source); book['综合']['B3'] = b_value
                book.save(source); book.close()
                requests = []
                def reply(engine, system, payload):
                    requests.append(copy.deepcopy(payload))
                    return case.response(engine,system,payload)
                with patch.object(UnifiedSemanticEngine, '_http', reply):
                    target = case.case.target(a_value); engine = case.case.engine()
                    result = engine.recognize_workbook(source, target_mapping=target)
                detail = next(p for p in requests if p['task']=='subtable_detail')
                self.assertIn('B3',detail['positions']['综合'])
                saved = load_fact_mapping(result['mapping_path'],engine.standard)
                self.assertEqual(saved['facts'][0]['raw_value'], b_value)
                self.assertEqual(saved['recognition_scope']['coverage']['targets'][0]['exact_sources'], ['综合!B3'])
                self.assertTrue(saved['recognition_scope']['coverage']['satisfied'])
                a_key = semantic_key(target['facts'][0],engine.standard)
                b_key = semantic_key(saved['facts'][0],engine.standard)
                self.assertEqual(a_key,b_key); identities.append(a_key)
        self.assertEqual(identities[0],identities[1])
        self.assertEqual(identities[1],identities[2])


if __name__ == '__main__':
    unittest.main()
