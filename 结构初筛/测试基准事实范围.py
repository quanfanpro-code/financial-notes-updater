"""搜索过程清单可不同，候选对应及业务范围必须一致。"""
from copy import deepcopy
import unittest
from 比较新表 import 比较完整事实

class 基准事实范围(unittest.TestCase):
    def test_搜索排除清单不是业务结果(self):
        a={'工作表':[{'名称':'表','上层语义':{'业务路径':[{'附注':'现金'}],
            '未进入比较':[{'旧表编号':'原表甲'}]}}],
           '候选':[{'结构对应':[{'旧坐标':'B3','新坐标':'C3'}]}]}
        b=deepcopy(a);b['工作表'][0]['上层语义']['未进入比较'].append({'旧表编号':'原表乙'})
        self.assertEqual(比较完整事实(a),比较完整事实(b))
        b['候选'][0]['结构对应'][0]['新坐标']='D3'
        self.assertNotEqual(比较完整事实(a),比较完整事实(b))
        b=deepcopy(a);b['工作表'][0]['上层语义']['业务路径']=[{'附注':'存货'}]
        self.assertNotEqual(比较完整事实(a),比较完整事实(b))

if __name__=='__main__':unittest.main(verbosity=2)
