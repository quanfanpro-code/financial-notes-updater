"""同一小表不再每24格重复发送整份上下文；完整覆盖仍须保存。"""
import json
from pathlib import Path
import unittest
import uuid
from openpyxl import Workbook
from 附注更新.统一识别 import UnifiedSemanticEngine
from 测试.测试统一语义 import records


class 整表合批测试(unittest.TestCase):
    def test_一百二十格只需一个任务包的两轮识别(self):
        root=Path(__file__).resolve().parents[1]/'测试结果'/('合批验证_'+uuid.uuid4().hex)
        root.mkdir();book=Workbook();sheet=book.active
        for row in range(1,121):sheet.cell(row,1,'合成说明文本')
        source=root/'合成表.xlsx';book.save(source);book.close()
        gold=root/'合成金标准.jsonl';gold.write_text('\n'.join(json.dumps(x,ensure_ascii=False) for x in records()),encoding='utf-8')
        calls=[]
        class Controlled(UnifiedSemanticEngine):
            def _request(self,system,payload,validator):
                calls.append(list(payload['candidate_cells']))
                return validator({'cells':{address:{'kind':'non_business','category':'annotation',
                    'reason':'合成说明','evidence_cells':['A1']} for address in payload['candidate_cells']}})
        result=Controlled({},gold,root/'结果').recognize_workbook(source)
        self.assertTrue(result['complete'])
        self.assertEqual(len(result['excluded']),120)
        self.assertEqual(len(calls),2)


if __name__=='__main__':unittest.main()
