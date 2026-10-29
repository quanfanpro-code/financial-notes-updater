# coding: utf-8
"""只检查文字、合并和位置传递，不调用模型。"""
import json
from pathlib import Path
from uuid import uuid4
from zipfile import ZipFile
import unittest
from openpyxl import Workbook

from 附注更新.Word原文材料 import word_source_material, word_request_material
from 附注更新.表格 import file_hash


class Word原文材料测试(unittest.TestCase):
    def setUp(self):
        self.folder = Path(__file__).resolve().parents[1] / '输出' / '主线接续_20261027_213431' / '合成离线验证' / ('Word材料_' + uuid4().hex)
        self.folder.mkdir(parents=True)
        self.word = self.folder / '合成原文.docx'
        xml = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>表前说明</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>项目</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>余额</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>金额</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>比例</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc><w:p><w:r><w:t>原文 | 不拆开</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>
</w:tbl><w:p><w:r><w:t>表后说明</w:t></w:r></w:p></w:body></w:document>'''
        with ZipFile(self.word, 'w') as z:
            z.writestr('word/document.xml', xml.encode('utf-8'))
        self.excel = self.folder / '合成原表.xlsx'
        book = Workbook()
        book.save(self.excel)
        book.close()
        self.link = self.folder / '已有位置关联.json'
        self.data = {'schema': '附注Word关联-v1', 'word_path': str(self.word),
                     'word_hash': file_hash(self.word), 'linked_word': str(self.word),
                     'linked_word_hash': file_hash(self.word), 'a_path': str(self.excel),
                     'a_hash': file_hash(self.excel), 'context': {'tables': [{
                         'table_id': '/word/document.xml#body/1', 'range_name': '原表一',
                         'sheet': '合成表', 'first_row': 10, 'first_column': 2,
                         'row_count': 3, 'column_count': 3}]}}
        self.link.write_text(json.dumps(self.data, ensure_ascii=False), encoding='utf-8')

    def test_横向纵向合并表号及Excel位置保留(self):
        result = word_source_material(self.link)
        self.assertEqual([b['kind'] for b in result['blocks']], ['paragraph', 'table', 'paragraph'])
        table = result['blocks'][1]
        self.assertEqual((table['table_number'], table['range_name'], table['grid_columns']), (1, '原表一', 3))
        self.assertEqual(table['rows'][0]['cells'][0]['row_span'], 2)
        self.assertEqual(table['rows'][0]['cells'][1]['column_span'], 2)
        self.assertEqual(table['rows'][0]['cells'][1]['source_excel'], '合成表!C10')
        self.assertEqual(table['rows'][1]['cells'][0]['merge_anchor'], [1, 1])
        self.assertEqual(table['rows'][1]['cells'][1]['source_excel'], '合成表!C11')

    def test_省略位置和空白不伪造成其他格(self):
        table = word_source_material(self.link)['blocks'][1]
        row = table['rows'][2]
        self.assertEqual(row['grid_before'], 1)
        self.assertEqual(row['cells'][0]['source_excel'], '合成表!C12')
        self.assertEqual(row['cells'][0]['content'][0]['text'], '原文 | 不拆开')
        self.assertEqual(row['cells'][1]['content'][0]['text'], '')
        self.assertEqual(len(row['cells']), 2)

    def test_当前坐标只用已有映射不猜未登记位置(self):
        positions = {'合成表!C12': '合成表!C20'}
        table = word_source_material(self.link, positions)['blocks'][1]
        self.assertEqual(table['rows'][2]['cells'][0]['current_excel'], '合成表!C20')
        self.assertIsNone(table['rows'][0]['cells'][0]['current_excel'])
        self.assertEqual(table['rows'][0]['cells'][0]['source_excel'], '合成表!B10')
        self.assertEqual(file_hash(self.word), self.data['word_hash'])
        self.assertEqual(file_hash(self.excel), self.data['a_hash'])

    def test_共享重复字段后全部内容可原样还原(self):
        from copy import deepcopy
        original = word_source_material(self.link, {'合成表!C12': '合成表!C20'})
        packed = word_request_material(original)
        restored = deepcopy(packed)
        fields = restored.pop('cell_fields')[:-1]
        restored.pop('content_fields')
        restored.pop('paragraph_fields')
        body = restored.pop('paragraph_position_base')
        restored.pop('excel_coordinates_relative_to_table_sheet')
        for index, table in enumerate(restored['blocks']):
            if isinstance(table, list):
                restored['blocks'][index] = {'kind': 'paragraph', 'word_position': body + table[0], 'text': table[1]}
                continue
            if table['kind'] != 'table':
                continue
            for row in table['rows']:
                cells = []
                for values in row['cells']:
                    cell = dict(zip(fields, values[:-1]))
                    cell.update(values[-1])
                    cell['word_row'] = row['word_row']
                    for key in ('source_excel', 'current_excel'):
                        if cell.get(key) and '!' not in cell[key]:
                            cell[key] = table['sheet'] + '!' + cell[key]
                    cell['content'] = [dict(zip(['kind', 'word_position', 'text'],
                                               [block[0], table['table_id'] + block[1], block[2]]))
                                       if isinstance(block, list) else block for block in cell['content']]
                    cells.append(cell)
                row['cells'] = cells
        self.assertEqual(restored, original)
        self.assertLess(len(json.dumps(packed)), len(json.dumps(original)))


if __name__ == '__main__':
    unittest.main()
