"""验证完整回执格式恢复不改变业务内容，也不接受冲突字段。"""
import io
import json
import unittest
from unittest.mock import Mock, patch
from 附注更新.语义 import SemanticError, _repair_complete_response


class 完整回执恢复测试(unittest.TestCase):
    def test_只恢复相同范围清单和结束围栏(self):
        item = '"metric_id":"m","can_partition":true,"dimension_ids":["d"],"reason":"范围","dimension_ids":["d"]'
        text = '{"metric_scope_checks":[{' + item + '}]}'
        for tail in ('', '\n``', '\n```'):
            decoded, repairs = _repair_complete_response(text + tail)
            self.assertEqual(decoded['metric_scope_checks'][0]['dimension_ids'], ['d'])
            self.assertEqual(len(repairs), 1 + bool(tail))
        for invalid in (text.replace('"dimension_ids":["d"]}', '"dimension_ids":["x"]}'),
                        '{"dimension_ids":[],"dimension_ids":[]}',
                        '{"cells":{},"cells":{}}', text + '{}', text + '\n其他答案', text[:-2]):
            with self.subTest(invalid=invalid), self.assertRaises((ValueError, SemanticError)):
                _repair_complete_response(invalid)


if __name__ == '__main__':
    unittest.main()
