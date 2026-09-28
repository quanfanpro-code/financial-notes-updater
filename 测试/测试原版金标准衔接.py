"""完整定义相同才可沿用已批准转换，不能按编号或名称放行。"""
import copy
import importlib
import importlib.util
import unittest
from 测试.测试统一迁移 import fixture
from 附注更新.统一迁移 import migrate_fact
from 附注更新.统一语义 import validate_records


class 原版衔接测试(unittest.TestCase):
    def test_原定义相同复用转换并保留原批准依据(self):
        self.assertIsNotNone(importlib.util.find_spec('附注更新.原版金标准衔接'))
        api = importlib.import_module('附注更新.原版金标准衔接')
        original, definition, standard = fixture()
        rows = [r for kind in ('metric','dimension','source_example','legacy_conversion') for r in standard[kind].values()]
        before = copy.deepcopy(rows)
        result = api.bridge_definitions({definition['id']: definition}, 'a'*64, rows)
        self.assertEqual(rows, before)
        self.assertEqual(len(result['linked']), 1)
        new = validate_records(result['records']); new['sha256'] = 'f'*64
        migrated = migrate_fact(original, definition, 'a'*64, new, {})
        self.assertEqual(migrated['fact']['raw_value'], '125.00')
        self.assertEqual(migrated['fact']['dimensions']['asset_selection']['members'], ['办公设备'])
        self.assertEqual(new['legacy_conversion']['旧办公设备转换'], standard['legacy_conversion']['旧办公设备转换'])
        self.assertEqual(api.bridge_definitions({definition['id']:definition}, 'a'*64, result['records'])['records'], result['records'])
        changed = copy.deepcopy(definition); changed['slot']['name'] = '净额'
        self.assertEqual(api.bridge_definitions({definition['id']:changed}, 'a'*64, rows)['linked'], [])
        pending = copy.deepcopy(rows)
        cv = next(r for r in pending if r['record_type']=='legacy_conversion')
        cv.update(status='unresolved', target=None)
        self.assertEqual(api.bridge_definitions({definition['id']:definition}, 'a'*64, pending)['linked'], [])


if __name__ == '__main__': unittest.main()
