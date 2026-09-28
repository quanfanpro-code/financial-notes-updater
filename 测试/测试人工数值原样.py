"""只按语义对应原样搬运，数值及显示格式均由用户决定。"""
import unittest
from 附注更新.表格 import _convert, WorkbookError


class 原值测试(unittest.TestCase):
    def test_旧入口不同倍率仍原样搬运(self):
        target = {'value_type': 'monetary', 'dimensions': {'unit': '元', 'scale': 1, 'currency': 'CNY'}}
        source = {'value_type': 'monetary', 'dimensions': {'unit': '万元', 'scale': 10000, 'currency': 'CNY'}}
        self.assertEqual(_convert(-1.25, source, target, {}, {}), -1.25)

    def test_旧入口保留人工金额文字及符号(self):
        mapping = {'value_type': 'monetary', 'dimensions': {'unit': '元', 'scale': 1, 'currency': 'CNY'}}
        for value in ('-1.2500', -1.25, 0, '—'):
            with self.subTest(value=value):
                self.assertEqual(_convert(value, mapping, mapping, {}, {}), value)

    def test_百分比不能因显示格式自行乘除一百(self):
        mapping = {'value_type': 'percentage', 'dimensions': {'scale': 1}}
        self.assertEqual(_convert(0.12, mapping, mapping, {'number_format': '0%'}, {'number_format': '0.00'}), 0.12)

    def test_更新不能根据金额推断并调整目标显示格式(self):
        from 测试.测试表格 import WorkbookTests
        from 附注更新.表格 import build_update_plan
        _,_,a,b,am,bm=WorkbookTests().cash_format_fixture()
        plan=build_update_plan(a,b,am,bm)
        self.assertFalse(plan['issues'])
        self.assertTrue(plan['updates'])
        self.assertTrue(all('number_format_update' not in row for row in plan['updates']))


if __name__ == '__main__':
    unittest.main()
