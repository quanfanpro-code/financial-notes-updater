"""独立小样本回归：全部材料虚构，不依赖开发者电脑或收费模型。"""
import copy
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from 建立结构索引 import 读取工作簿, 规范文字
from 比较新表 import 比较工作簿, _候选问题
import 运行初筛


def 虚构参照():
    表 = {'小表编号': '演示资金表', '小表名称': '演示货币资金', '样式': 1,
         '来源': {'工作表': '演示表'}, '节点': [], '合并': [], '项目顺序': [], '上下文': [], '表头': []}
    数据 = [('演示货币资金', None, None), ('项目', '期末余额', '期初余额'),
            ('现金', 100, 90), ('银行存款', 200, 180), ('合计', 300, 270)]
    for r, 行 in enumerate(数据, 1):
        for c, 值 in enumerate(行, 1):
            if 值 is None:
                continue
            文 = isinstance(值, str)
            n = {'坐标': f'{get_column_letter(c)}{r}', '行': r, '列': c, '原文': 值,
                 '角色': ('标题' if r == 1 else '表头' if r == 2 else '项目') if 文 else '业务数值'}
            if 文:
                n['查找文字'] = 规范文字(值)
            else:
                n['语义引用'] = {'说明': '完全虚构的演示位置，不是已批准的业务语义'}
            表['节点'].append(n)
    return 表


def 内存索引(表):
    正文 = {}
    for n in 表['节点']:
        if n.get('查找文字'):
            正文.setdefault(n['查找文字'], []).append({'小表编号': 表['小表编号']})
    return {'目录': '虚构演示', '规则版本': '演示', '正文': 正文, '别名': {}, '记录': [表],
            '小表': {表['小表编号']: 表}, '索引清单': {}}


class 收尾回归(unittest.TestCase):
    def setUp(self):
        self.目录 = Path(tempfile.mkdtemp(prefix='附注收尾验证_'))
        self.表 = 虚构参照()
        self.索引 = 内存索引(self.表)

    def 新表(self, offsets=((0, 0),), merge=None, sheets=('演示表',)):
        w = Workbook()
        for i, 名 in enumerate(sheets):
            s = w.active if i == 0 else w.create_sheet()
            s.title = 名
            for dr, dc in offsets:
                for n in self.表['节点']:
                    s.cell(n['行'] + dr, n['列'] + dc, n['原文'])
            if merge:
                s.merge_cells(merge)
        p = self.目录 / '虚构.xlsx'
        w.save(p)
        w.close()
        return p

    def test_左右并排两个位置都保留(self):
        r = 比较工作簿(读取工作簿(self.新表(((0, 0), (0, 5)))), self.索引)
        self.assertEqual({x['新表范围'] for x in r['候选']}, {'A1:C5', 'F1:H5'})

    def test_重复十二处全部保留(self):
        r = 比较工作簿(读取工作簿(self.新表(tuple((10*k, 0) for k in range(12)))), self.索引)
        self.assertEqual({x['新表范围'] for x in r['候选']}, {f'A{1+10*k}:C{5+10*k}' for k in range(12)})

    def test_跨出候选边界的合并仍需细读(self):
        r = 比较工作簿(读取工作簿(self.新表(merge='A1:E1')), self.索引)
        c = next(x for x in r['候选'] if x['种子偏移'] == [0, 0])
        self.assertTrue(any(x.get('新合并') == 'A1:E1' for x in c['关系核对']['合并']))
        self.assertTrue(_候选问题(c))

    def test_单词重复两行不能冒充两个识别词(self):
        for n in self.表['节点']:
            if n.get('查找文字'):
                n['原文'] = n['查找文字'] = '现金' if n['行'] in (3, 4) else '不存在' + n['坐标']
        self.索引 = 内存索引(self.表)
        w = Workbook(); s = w.active; s['A3'] = '现金'; s['A4'] = '现金'
        p = self.目录 / '单词.xlsx'; w.save(p); w.close()
        self.assertEqual(比较工作簿(读取工作簿(p), self.索引)['候选'], [])

    def test_基准跨工作表遗漏不能返回成功(self):
        p = self.新表(sheets=('甲表', '乙表'))
        真比较 = 比较工作簿
        def 故障注入(记录, 索引, 预筛选=True):
            r = 真比较(记录, 索引, 预筛选)
            if 预筛选:
                r['候选'] = [c for c in r['候选'] if c['新工作表'] == '甲表']
                cells = list(记录['sheets'][1]['cells'])
                r['未匹配范围'].append({'工作表': '乙表', '隐藏': False, '坐标数': len(cells),
                    '分段': [{'范围': 'A1:C5', '坐标数': len(cells), '坐标清单': cells, '原文摘录': []}]})
            return r
        输出 = self.目录 / '交接材料'
        with patch('比较新表.载入索引', return_value=self.索引), patch('比较新表.比较工作簿', side_effect=故障注入):
            码 = 运行初筛.执行比较(p, '虚构', 输出, True)
        r = json.loads((输出 / '初筛结果.json').read_text(encoding='utf-8-sig'))
        self.assertNotEqual(码, 0)
        self.assertEqual(r['基准对照']['基准候选数'], 2)
        self.assertEqual(len(r['基准对照']['遗漏']), 1)

    def test_一个词重复四次仍不构成部分候选(self):
        self.表['节点'] = [{'坐标': f'A{r}', '行': r, '列': 1, '角色': '项目',
            '原文': t, '查找文字': t} for r, t in enumerate(['现金']*4 + ['存款', '票据', '合计'], 1)]
        self.索引 = 内存索引(self.表)
        w = Workbook(); s = w.active
        for r in range(1, 5): s.cell(r, 1, '现金')
        p = self.目录 / '重复同词.xlsx'; w.save(p); w.close()
        self.assertEqual(比较工作簿(读取工作簿(p), self.索引)['候选'], [])

    def test_输入清单过期必须在整簿比较之前失败(self):
        import hashlib
        import 核验真实样本
        原件 = self.目录 / '虚构输入.txt'
        原件.write_text('只核对指纹，不读取为工作簿', encoding='utf-8')
        样本 = {'文件': str(原件), 'sha256': hashlib.sha256(原件.read_bytes()).hexdigest()}
        清单 = self.目录 / '固定样本.json'
        清单.write_text(json.dumps({'五样式原样查回': [样本], '真实A端': 样本}), encoding='utf-8')
        输入 = self.目录 / '过期输入.json'
        输入.write_text(json.dumps({'条目': [{'用途': '复用代码:机械读取模块', '路径': str(原件), 'sha256': '0'*64}]}), encoding='utf-8')
        with patch.object(核验真实样本, '读取工作簿') as 读取:
            with self.assertRaisesRegex(ValueError, '指纹不符'):
                核验真实样本.main(['--索引目录', str(self.目录/'未用索引'), '--样本清单', str(清单), '--输入清单', str(输入), '--输出', str(self.目录/'核验输出')])
            读取.assert_not_called()

    def test_同一通用词两行全对不得被预筛选挡住(self):
        self.表['节点'] = [{'坐标': f'A{r}', '行': r, '列': 1, '角色': '项目',
            '原文': '合计', '查找文字': '合计'} for r in (1, 2)]
        索引 = 内存索引(self.表)
        索引['正文']['合计'] = [{'小表编号': self.表['小表编号']}] * 11
        记录 = 读取工作簿(self.新表())
        self.assertEqual(len(比较工作簿(记录, 索引, False)['候选']), 1)
        self.assertEqual(len(比较工作簿(记录, 索引, True)['候选']), 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
