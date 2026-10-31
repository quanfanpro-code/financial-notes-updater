"""旧双轮和格式补问入口拒绝执行，不生成第二次请求。"""
import unittest
from unittest.mock import patch
from 附注更新.语义 import SemanticEngine, SemanticError
from 附注更新.统一识别 import UnifiedSemanticEngine


class 格式补正测试(unittest.TestCase):
    def test_旧入口不执行双轮或格式补问(self):
        for cls, method in [(SemanticEngine, 'recognize'), (UnifiedSemanticEngine, 'classify_packet')]:
            engine=object.__new__(cls)
            with patch.object(engine,'_http') as request:
                with self.assertRaisesRegex(SemanticError,'旧.*识别入口已停用'):
                    getattr(engine,method)({})
                request.assert_not_called()
        self.assertFalse(hasattr(UnifiedSemanticEngine,'_classification_round'))
        self.assertFalse(hasattr(UnifiedSemanticEngine,'_classify_reviewed_packet'))

    def test_旧位置取数和名单安排不能保存为对应(self):
        from 附注更新.A合集 import save_source_relations, save_list_relations
        from 附注更新.正常入口识别 import answer_relations
        for call in [lambda:save_source_relations({}, {}, '旧安排', [], '不应写入'),
                     lambda:save_list_relations({}, {}, '旧安排', '范围', '不应写入')]:
            with self.assertRaisesRegex(ValueError,'旧.*安排入口已停用'):call()
        self.assertEqual(answer_relations({}, {'识别答复':'A 原表!B3采用 B 来源!C3'}, [], [], {}, []),[])



if __name__=='__main__':unittest.main()
