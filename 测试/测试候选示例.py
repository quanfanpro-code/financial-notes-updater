"""候选语义示例生成隔离验证；临时资料保留。"""
from pathlib import Path
import tempfile
import unittest
import json
from 附注更新.人工映射复用 import install, CORPUS


class 候选示例(unittest.TestCase):
    def test_候选输出不修改正式且重复位置拒绝(self):
        root = Path(tempfile.mkdtemp(prefix='候选示例验证_'))
        (root/'语义金标准').mkdir()
        (root/'附注样式示例').mkdir()
        (root/'语义金标准/财务报表附注语义金标准-v1.jsonl').write_text(
            json.dumps({'id':'s','status':'active'})+'\n',encoding='utf-8-sig')
        for i in range(1,6):
            (root/f'附注样式示例/附注样式{i}.xlsx').write_bytes(b'fixed-input')
            p = root/f'work/excel-mapping-v1/style{i}-manual-review-v2'
            p.mkdir(parents=True)
            name = '一项人工映射.jsonl' if i == 2 else f'样式{i}-合并映射.jsonl'
            row = {'mapping_id':str(i),'worksheet':'表','cell':'C3','target_slot_id':'s',
                   'dimensions':{'period_role':'closing'}}
            (p/name).write_text(json.dumps(row)+'\n',encoding='utf-8-sig')
        before = CORPUS.read_bytes() if CORPUS.exists() else None
        out = root/'候选'
        try:
            result = install(root, output_dir=out)
        except TypeError as e:
            self.fail('候选输出接口尚不支持：'+str(e))
        self.assertEqual(result['business_cells'],5)
        self.assertEqual(Path(result['path']).parent,out)
        self.assertEqual(before,CORPUS.read_bytes() if CORPUS.exists() else None)
        with self.assertRaises(ValueError): install(root, output_dir=out)
        source = root/'work/excel-mapping-v1/style1-manual-review-v2/样式1-合并映射.jsonl'
        with source.open('a',encoding='utf-8') as f:
            f.write(json.dumps({'mapping_id':'duplicate','worksheet':'表','cell':'C3','target_slot_id':'s'})+'\n')
        with self.assertRaisesRegex(ValueError,'重复来源'): install(root, output_dir=root/'重复')

if __name__ == '__main__': unittest.main(verbosity=2)
