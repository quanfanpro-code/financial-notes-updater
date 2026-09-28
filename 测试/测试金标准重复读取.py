# coding: utf-8
"""重复读取可复用校验，但文件变化及调用者修改不能污染语义定义。"""
import copy
import hashlib
import json
import unittest
from pathlib import Path
from unittest.mock import patch
from uuid import uuid4

from 附注更新 import 统一语义 as api
from 测试.测试统一语义 import records


def content():
    rows = records()
    rows[-2]['review_refs'] = ['独立测试-' + uuid4().hex]
    return ('\n'.join(json.dumps(row, ensure_ascii=False) for row in rows)).encode('utf-8')


class StandardReadTests(unittest.TestCase):
    def test_same_content_reuses_validation_and_returns_independent_definitions(self):
        raw = content()
        with patch.object(Path, 'read_bytes', return_value=raw) as read, \
             patch.object(api, 'validate_records', wraps=api.validate_records) as validate:
            first = api.read_standard_v2('同一金标准.jsonl')
            expected = copy.deepcopy(first)
            first['dimension']['entity']['meaning'] = '调用者的临时修改'
            first['metric']['固定资产原值']['definition']['constraints'].append('临时限制')
            second = api.read_standard_v2('同一金标准.jsonl')
        self.assertEqual(second, expected)
        self.assertIsNot(first, second)
        self.assertEqual(read.call_count, 2, '每次须读取实际内容以发现文件变化')
        self.assertEqual(validate.call_count, 1, '完全相同内容不应重复执行全库校验')

    def test_same_path_changed_contents_never_reuses_old_meaning(self):
        first, second = content(), content()
        with patch.object(Path, 'read_bytes', side_effect=[first, second, first]), \
             patch.object(api, 'validate_records', wraps=api.validate_records) as validate:
            results = [api.read_standard_v2('同一路径.jsonl') for _ in range(3)]
        self.assertEqual([item['sha256'] for item in results],
                         [hashlib.sha256(raw).hexdigest() for raw in (first, second, first)])
        self.assertEqual(results[0], results[2])
        self.assertNotEqual(results[0]['source_example'], results[1]['source_example'])
        self.assertEqual(validate.call_count, 2)

    def test_invalid_replacement_and_missing_file_still_fail(self):
        raw = content()
        invalid_rows = records()
        invalid_rows[-1]['dimension_refs'].append({'id': '不存在的维度', 'required': True})
        invalid = '\n'.join(json.dumps(row, ensure_ascii=False) for row in invalid_rows).encode('utf-8')
        with patch.object(Path, 'read_bytes', side_effect=[raw, invalid, FileNotFoundError('文件已移走')]):
            api.read_standard_v2('同一路径.jsonl')
            with self.assertRaises(ValueError):
                api.read_standard_v2('同一路径.jsonl')
            with self.assertRaises(FileNotFoundError):
                api.read_standard_v2('同一路径.jsonl')


if __name__ == '__main__':
    unittest.main()
