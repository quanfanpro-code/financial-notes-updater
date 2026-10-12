"""产物须绑定生成时版本；同目录换索引不能沿用旧验收。"""
import unittest
from pathlib import Path
import tempfile
import 比较新表 as 比较


class 运行版本(unittest.TestCase):
    def test_内容变化即使同目录也拒绝(self):
        self.assertTrue(hasattr(比较, '生成版本'), '尚未登记生成版本')
        root=Path(tempfile.mkdtemp(prefix='初筛版本证据_'))
        book=root/'输入.xlsx'; book.write_bytes(b'fixed workbook')
        index=root/'索引';index.mkdir()
        for name in ('索引清单.json','小表记录.jsonl','文字查找.json','去向明细.jsonl'):
            (index/name).write_text('{}',encoding='utf-8')
        version=比较.生成版本(book,index)
        self.assertEqual(version,比较.生成版本(book,index))
        with (index/'小表记录.jsonl').open('a',encoding='utf-8') as f:f.write('\n{}')
        with self.assertRaisesRegex(ValueError,'版本'):
            比较.核对生成版本({'生成版本':version},book,index)

if __name__=='__main__':unittest.main(verbosity=2)
