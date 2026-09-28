"""仅补正格式错误格；有效格不重识别，失败时保留原未决。"""
import copy
from pathlib import Path
import unittest
from 附注更新.统一识别 import UnifiedSemanticEngine,build_payload
from 附注更新.统一语义 import validate_records
from 测试.测试统一识别 import packet,answer
from 测试.测试统一语义 import records


class 格式补正测试(unittest.TestCase):
    def test_只重试坏格且不修改已经通过的结果(self):
        self.assertTrue(hasattr(UnifiedSemanticEngine,'_classification_round'))
        standard=validate_records(records());source=packet()
        payload=build_payload(source,standard);payload['round']=1
        calls=[];valid=answer();broken=copy.deepcopy(valid)
        broken['cells']['F9']['dimensions']['asset_selection'].pop('completeness')
        class Controlled(UnifiedSemanticEngine):
            def _request(self,system,payload,validator):
                calls.append(payload['candidate_cells'])
                return validator(broken if len(calls)==1 else {'cells':{'F9':valid['cells']['F9']}})
        engine=object.__new__(Controlled);engine.standard=standard;engine.log=lambda message:None;engine.settings={}
        result,repair=engine._classification_round(payload,source)
        self.assertEqual(calls,[['D7','F9'],['F9']])
        self.assertEqual(result['D7']['kind'],'metric_fact')
        self.assertEqual(result['F9']['kind'],'metric_fact')
        self.assertEqual(repair['cells'],['F9'])
        self.assertEqual(repair['remaining_invalid'],[])
        calls.clear()
        class Failed(Controlled):
            def _request(self,system,payload,validator):
                if payload.get('format_correction'):raise ValueError('合成补正请求失败')
                return super()._request(system,payload,validator)
        failed=object.__new__(Failed);failed.standard=standard;failed.log=lambda message:None;failed.settings={}
        preserved,report=failed._classification_round(payload,source)
        self.assertEqual(preserved['D7'],result['D7'])
        self.assertEqual(preserved['F9']['kind'],'invalid_response')
        self.assertEqual(report['remaining_invalid'],['F9'])


if __name__=='__main__':unittest.main()
