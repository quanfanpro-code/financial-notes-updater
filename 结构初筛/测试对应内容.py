"""结构对应必须保存并核对当前新表内容，不能只保留旧模板原值。"""
import unittest
from 比较新表 import 比较工作簿, 完整性检查
from 测试独立复核比较 import 参照, 工作簿, 索引

class 对应内容(unittest.TestCase):
    def test_业务对应的当前公式和缓存要完整保留(self):
        old,data=参照('cash','货币资金',['库存现金','银行存款'])
        book=工作簿('货币资金',{**data,'C3':'=2+3'})
        book['sheets'][0]['cells']['C3'].update(formula='=2+3',cached_value=None)
        result=比较工作簿(book,索引([old]))
        point=next(d for d in result['候选'][0]['结构对应'] if d['新坐标']=='C3')
        self.assertEqual(point.get('新表内容'),{'原文':'=2+3','公式':'=2+3','缓存':None})
        self.assertTrue(完整性检查(book,result))
        point['新表内容']['公式']='=2+4'
        with self.assertRaisesRegex(AssertionError,'对应内容'):
            完整性检查(book,result)

if __name__=='__main__':unittest.main(verbosity=2)
