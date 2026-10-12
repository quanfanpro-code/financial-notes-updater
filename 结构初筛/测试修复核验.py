"""修复批次的行为反例；预期手工固定，纯内存，不清理文件。"""
import unittest
from copy import deepcopy
from datetime import datetime
from 比较新表 import 核验预期对应, 完整性检查, 比较工作簿
from 测试独立复核比较 import 参照, 工作簿, 索引


class 修复核验(unittest.TestCase):
    def setUp(self):
        self.预期 = {('旧表', 'C3'): ('新表', 'D8'), ('旧表', 'C4'): ('新表', 'D9')}
        self.正常 = {'候选': [{'小表名称': '测试小表', '结构对应': [
            {'旧工作表': a, '旧坐标': b, '新工作表': c, '新坐标': d}
            for (a, b), (c, d) in self.预期.items()]}]}

    def test_非空预期不能空过或漏格(self):
        self.assertTrue(核验预期对应(self.正常, self.预期, '测试小表'))
        for count in (0, 1):
            r = deepcopy(self.正常)
            r['候选'][0]['结构对应'] = r['候选'][0]['结构对应'][:count]
            with self.subTest(count=count), self.assertRaisesRegex(AssertionError, '遗漏'):
                核验预期对应(r, self.预期, '测试小表')

    def test_先错后对不能覆盖冲突(self):
        r = deepcopy(self.正常)
        bad = {**r['候选'][0]['结构对应'][0], '新坐标': 'Z99'}
        r['候选'][0]['结构对应'].insert(0, bad)
        with self.assertRaisesRegex(AssertionError, '冲突'):
            核验预期对应(r, self.预期, '测试小表')

    def test_允许其他正确结构对应(self):
        required = {('旧表', 'C3'): ('新表', 'D8')}
        self.assertTrue(核验预期对应(self.正常, required, '测试小表', 允许关系=self.预期))

    def test_候选内部内容不能只有矩形没有去向(self):
        old, data = 参照('cash', '货币资金', ['库存现金', '银行存款'])
        for v in ('新增文字', 0, False, datetime(2026, 1, 1), '=1+1'):
            book = 工作簿('货币资金', {**data, 'B3': v})
            if v == '=1+1':
                book['sheets'][0]['cells']['B3'].update(formula=v, cached_value=None)
            r = 比较工作簿(book, 索引([old]))
            self.assertTrue(完整性检查(book, r))
            for candidate in r['候选']:
                candidate['未对应内容'] = []
            with self.subTest(value=v), self.assertRaisesRegex(AssertionError, 'B3'):
                完整性检查(book, r)

    def test_未对应原文公式缓存不能被改写(self):
        old, data = 参照('cash', '货币资金', ['库存现金', '银行存款'])
        book = 工作簿('货币资金', {**data, 'B3': '新增说明'})
        r = 比较工作簿(book, 索引([old]))
        r['候选'][0]['未对应内容'][0]['原文'] = '不同说明'
        with self.assertRaisesRegex(AssertionError, '原文'):
            完整性检查(book, r)

    def test_已知业务空格也必须有去向(self):
        old, data = 参照('cash', '货币资金', ['库存现金', '银行存款'])
        book = 工作簿('货币资金', {**data, 'C3': None})
        r = 比较工作簿(book, 索引([old]))
        for c in r['候选']:
            c['结构对应'] = [d for d in c['结构对应'] if d['新坐标'] != 'C3']
        with self.assertRaisesRegex(AssertionError, 'C3'):
            完整性检查(book, r, 必需业务位置={('货币资金', 'C3')})

    def test_基准对照必须比较具体对应而非只有矩形(self):
        import 比较新表 as 模块
        old, data = 参照('cash', '货币资金', ['库存现金', '银行存款'])
        r = 比较工作簿(工作簿('货币资金', data), 索引([old]))
        other = deepcopy(r)
        other['候选'][0]['结构对应'][0]['新坐标'] = 'Z99'
        self.assertTrue(hasattr(模块, '比较完整事实'), '缺少完整事实对照')
        self.assertNotEqual(模块.比较完整事实(r), 模块.比较完整事实(other))
        other = deepcopy(r)
        other['运行统计']['耗时秒'] = 999
        self.assertEqual(模块.比较完整事实(r), 模块.比较完整事实(other))


if __name__ == '__main__':
    unittest.main(verbosity=2)
