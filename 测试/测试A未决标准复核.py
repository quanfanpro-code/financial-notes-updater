# coding: utf-8
"""正常入口仅在整份核查及定向纠正结束后复核现有标准。"""
from copy import deepcopy
from pathlib import Path
import uuid
import unittest
from unittest.mock import patch

from 附注更新.分步流程 import _recognize_v2


class A未决标准复核测试(unittest.TestCase):
    def run_case(self,incomplete,freeze=False):
        from 测试.测试未决语义自动纠正 import 未决语义纠正测试
        engine,original,_=未决语义纠正测试().run_case(False)
        result=deepcopy(original)
        if incomplete:result['a_semantic_reading']['batches']=[]
        folder=Path(__file__).resolve().parents[1]/'测试结果'/('A标准复核入口_'+uuid.uuid4().hex)
        folder.mkdir()
        learned={'path':engine.gold_path,'new_ids':[],'changed_ids':[],'reviews':[]}
        with patch('附注更新.统一识别.UnifiedSemanticEngine',return_value=engine), \
             patch.object(engine,'recognize_workbook',return_value=result), \
             patch('附注更新.统一补充.review_missing_standard',return_value=learned) as review:
            report=_recognize_v2({'a_path':result['source_path'],'gold_path':engine.gold_path,
                'output_dir':str(folder),'base_url':'codex://local','model':'gpt-6-luna','freeze_standard':freeze},'A',folder,lambda _:None,None)
        self.assertEqual(report['status'],'partial')
        self.assertEqual(review.call_count,0 if incomplete or freeze else 1)
        self.assertNotIn('gold_path',report['files'])
        return report

    def test_整份未决仅复核一次无新定义不循环(self):
        self.run_case(False)

    def test_核查中断不能转去生成标准(self):
        self.run_case(True)

    def test_冻结标准保留未决且不生成标准(self):
        self.run_case(False,freeze=True)
