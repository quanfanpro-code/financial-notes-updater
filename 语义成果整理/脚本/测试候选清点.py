"""固定候选参数不得悄悄写回正式路径。"""
import unittest
import 清点资产

class 候选清点(unittest.TestCase):
    def test_候选清点接口必须存在(self):
        self.assertTrue(hasattr(清点资产,'清点候选'),'尚无固定候选输入接口')

if __name__=='__main__':unittest.main(verbosity=2)
