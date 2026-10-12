# coding: utf-8
"""兼容只允许已核实的具体读取器版本，未知改动仍须拒绝。"""
import unittest
from pathlib import Path
from unittest.mock import patch
from 结构初筛 import 建立结构索引 as index


class 索引读取兼容测试(unittest.TestCase):
    def test_允许兼容不等于任意读取器版本放行(self):
        folder=Path(__file__).resolve().parents[1]/'结构初筛/索引/工作表语义版_正式'
        result=index.校验索引目录(folder,允许读取器版本变化=True)
        self.assertEqual(len(result['记录']),714)
        original=index.文件指纹
        def unknown(path):
            return 'f'*64 if Path(path).resolve()==index.机械读取模块路径.resolve() else original(path)
        with patch.object(index,'文件指纹',unknown):
            with self.assertRaisesRegex(ValueError,'读取器.*兼容|来源指纹不符'):
                index.校验索引目录(folder,允许读取器版本变化=True)

    def test_索引不兼容时保留定位范围继续正常识别(self):
        from 附注更新.B机械比较 import compare_located, comparison_material
        snapshot={'sha256':'a'*64,'sheets':[{'name':'表','max_row':2,'max_column':2,
                   'cells':{'A1':{'value':'项目'},'B2':{'value':1}},'merges':[]}]}
        table={'id':'本表','ranges':[{'sheet':'表','range':'A1:B2'}]}
        outline={'subtables':[table],'context_ranges':[]}
        with patch('附注更新.B机械比较.载入索引',side_effect=ValueError('读取器未经兼容核验')):
            result=compare_located(snapshot,outline)
        self.assertEqual(result['本次定位范围'],table['ranges'])
        self.assertEqual(comparison_material(result,table)['候选'],[])
        self.assertIn('正常识别',result['说明'])


if __name__=='__main__':unittest.main()
