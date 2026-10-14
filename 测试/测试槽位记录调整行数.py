# coding: utf-8
"""槽位记录须参与名单行数调整，新增格保留已知身份及真实空值。"""
import copy
import unittest

from 测试 import 测试统一结构 as 夹具
from 附注更新.模板布局 import project_layout
from 附注更新.统一勾稽 import _template_cell_records
from 附注更新.表格 import file_hash


def 槽位记录(mapping):
    result = copy.deepcopy(mapping)
    result['slot_bindings'] = []
    for fact in result['facts']:
        known = {key: copy.deepcopy(fact[key]) for key in ('metric_id', 'dimensions', 'dimension_evidence')}
        known['dimensions'].pop('entity')
        known['dimension_evidence'].pop('entity')
        known.update(missing_dimensions=['entity'], evidence_cells=['明细!A1'])
        sheet, cell = fact['id'].rsplit('!', 1)
        result['slot_bindings'].append({'sheet': sheet, 'cell': cell, 'partial_semantics': known,
                                       'recognition_provenance': fact['recognition_provenance']})
    result['facts'] = []
    result['owners'] = {}
    return result


class 槽位记录行数检查(unittest.TestCase):
    def test_已识别槽位参与删除新增且无同槽位来源不能删除(self):
        case = 夹具.UnifiedStructureTests(); case.setUp()
        target = 槽位记录(case.target)
        source = 槽位记录(case.make('新完整名单', [['完整名单', '年末余额'], ['甲类设备', 77], ['丙类设备', 88], ['合计', 165]],
                                    [('B2', '甲类设备'), ('B3', '丙类设备'), ('B4', None)]))
        delete = {'range_name': '附注表', 'axis': 'row', 'action': 'delete', 'index': 2, 'count': 1,
                  'template_index': 1, 'source_evidence': ['明细!A1'], 'coverage_reason': '完整新名单替换旧名单'}
        insert = {'range_name': '附注表', 'axis': 'row', 'action': 'insert', 'index': 2, 'count': 1, 'template_index': 1}
        view = project_layout(target, source, case.standard, [delete, insert], source_locations={'明细!B2', '明细!B3', '明细!B4'})
        self.assertEqual(view['_template_layout']['removed'], ['明细!B3'])
        self.assertEqual({row['cell'] for row in view['slot_bindings']}, {'B2', 'B3', 'B4'})
        rows = _template_cell_records(view)
        self.assertEqual(rows['明细!B3']['raw_value_state'], 'blank')
        self.assertIsNone(rows['明细!B3']['raw_value'])
        self.assertEqual(rows['明细!B3']['dimension_evidence']['asset_selection'], ['明细!A3'])
        self.assertEqual(rows['明细!B4']['raw_value'], 33)
        empty = copy.deepcopy(source); empty['slot_bindings'] = []
        with self.assertRaisesRegex(ValueError, '来源未提供该业务'):
            project_layout(target, empty, case.standard, [delete])
        self.assertEqual(file_hash(target['source_path']), target['source_hash'])
        self.assertEqual(file_hash(source['source_path']), source['source_hash'])


if __name__ == '__main__':
    unittest.main()
