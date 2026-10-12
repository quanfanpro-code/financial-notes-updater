# coding: utf-8
"""B重新定位后，回执路径变化不应迫使已经核实的同表重新绑定。"""
import unittest
from unittest.mock import patch
from 测试 import 测试B范围定位 as 原定位测试
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一勾稽 import load_fact_mapping


class B接续识别测试(unittest.TestCase):
    def test_重新定位沿用同一原表并可正式回读(self):
        case=原定位测试.B范围定位测试();case.setUp()
        def reply(engine,system,payload):return case.response(engine,system,payload)
        with patch.object(UnifiedSemanticEngine,'_http',reply):
            target=case.case.target()
            first_engine=case.case.engine()
            first=first_engine.recognize_workbook(case.case.source,target_mapping=target)
            second_engine=case.case.engine()
            second=second_engine.recognize_workbook(case.case.source,target_mapping=target,previous_result=first['mapping_path'])
        self.assertEqual(second['usage'].get('reused_confirmed_tables'),1)
        self.assertEqual(len(load_fact_mapping(second['mapping_path'],second_engine.standard)['facts']),1)


if __name__=='__main__':unittest.main()
