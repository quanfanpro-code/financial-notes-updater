# coding: utf-8
"""检查原Word章节层级及已存身份的原样传递，不调用业务模型。"""
import json
from pathlib import Path
from uuid import uuid4
from zipfile import ZipFile
import unittest
from unittest.mock import patch

from 测试 import 测试直接识别 as fixture
from 附注更新.A合集 import _records
from 附注更新.分步流程 import run_step
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.表格 import file_hash
from 附注更新.Word原文材料 import word_source_material, word_text_material, word_request_material


class 章节与既有身份参照测试(unittest.TestCase):
    def setUp(self):
        self.folder = Path(__file__).resolve().parents[1] / '输出' / '主线接续_20261027_213431' / '合成离线验证' / ('章节与参照_' + uuid4().hex)
        with patch.object(fixture, 'WORK', self.folder):
            self.case = fixture.直接识别测试()
            self.case.setUp()
        self.selection = {'mode': 'predicate', 'predicate': '设备甲与设备乙中用于生产的设备'}

    def recognize(self, side, year, currency='CNY', **extra):
        raw = json.dumps({'location': '合成表!B3', 'metric_id': '固定资产原值',
                          'dimensions': {'asset_selection': self.selection, 'currency': currency,
                                         'period': str(year) + '-12-31'}}, ensure_ascii=False)
        with patch.object(UnifiedSemanticEngine, 'request_text', return_value=raw) as call:
            result = run_step('recognize_' + side, {**self.case.config,
                side+'_target_table':{'id':'本表','sheet':'合成表','range':'A2:C4','data_ranges':['B3']},**extra})
        self.assertEqual(result['status'], 'complete', result)
        call.assert_called_once()
        saved = json.loads(Path(result['files'][side + '_mapping']).read_text('utf-8-sig'))
        return saved, call.call_args.args

    def test_A有既有B也只读本方原文及金标准(self):
        a, _ = self.recognize('a', 2025)
        b, _ = self.recognize('b', 2024, a_mapping=a['mapping_path'])
        result, (prompt, materials) = self.recognize(
            'a', 2025, 'USD', b_mapping=b['mapping_path'], a_allowed_metrics=['固定资产原值'])
        self.assertNotIn('既有另一方身份证表达.json',materials)
        self.assertIn('依据本方完整原文及相关金标准', prompt)
        own = dict(_records(result))['合成表!B3']['dimensions']
        self.assertEqual((own['period'], own['currency']), ('2025-12-31', 'USD'))
        self.assertEqual(own['asset_selection'], self.selection)

    def test_不可用B参照不成为A识别门槛(self):
        result, (_, materials) = self.recognize('a', 2025, b_mapping=str(self.folder / '不存在的B.json'))
        self.assertNotIn('既有另一方身份证表达.json', materials)
        self.assertEqual(dict(_records(result))['合成表!B3']['dimensions']['asset_selection'], self.selection)

    def test_Word直接及样式继承提纲原样保留且不产生业务身份(self):
        word = self.folder / '分级标题原文.docx'
        xml = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>母级标题</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="子级"/></w:pPr><w:r><w:t>子级标题</w:t></w:r></w:p>
<w:p><w:pPr><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>普通正文</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>原格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
</w:body></w:document>'''
        styles = '''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:styleId="父级"><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>
<w:style w:styleId="子级"><w:basedOn w:val="父级"/></w:style></w:styles>'''
        with ZipFile(word, 'w') as package:
            package.writestr('word/document.xml', xml.encode('utf-8'))
            package.writestr('word/styles.xml', styles.encode('utf-8'))
        link = self.folder / '分级标题位置关联.json'
        data = {'schema': '附注Word关联-v1', 'word_path': str(word), 'word_hash': file_hash(word),
                'linked_word': str(word), 'linked_word_hash': file_hash(word),
                'a_path': str(self.case.source), 'a_hash': file_hash(self.case.source),
                'context': {'tables': [{'table_id': '/word/document.xml#body/3', 'range_name': '原表',
                                       'sheet': '合成表', 'first_row': 3, 'first_column': 1,
                                       'row_count': 1, 'column_count': 1}]}}
        link.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
        material = word_source_material(link)
        self.assertEqual([b['word_outline_level'] for b in material['blocks'][:3]], [0, 2, 9])
        text = word_text_material(material)
        self.assertIn('# 母级标题', text)
        self.assertIn('### 子级标题', text)
        self.assertIn('提纲级别=9', text)
        self.assertNotIn('# 普通正文', text)
        self.assertNotIn('report_scope', text)
        self.assertEqual(word_request_material(material)['blocks'][:3], material['blocks'][:3])


if __name__ == '__main__':
    unittest.main()
