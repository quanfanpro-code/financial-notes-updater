# coding: utf-8
"""复核请求去重不丢原文，不串用依据，旧回执继续可查。"""
import copy
import json
import unittest
from pathlib import Path
from 测试 import 测试同义复核 as fixtures


class SharedReviewTests(unittest.TestCase):
    def test_request_restores_original_definitions_and_evidence(self):
        case = fixtures.EquivalenceTests(); case.setUp()
        folder, pair = case.fixture(); engine = case.engine(folder)
        reverse = {'left': copy.deepcopy(pair['right']), 'right': copy.deepcopy(pair['left']),
                   'left_path': pair['right_path'], 'right_path': pair['left_path']}
        pairs = [pair, reverse]
        original = case.api.build_payload(pairs, engine.standard, same_cell=False)
        before = copy.deepcopy(original)
        calls = []
        def respond(system, payload):
            calls.append(copy.deepcopy(payload))
            result = case.response(payload)
            result['pairs'][case.api.pair_id(reverse)]['decision'] = 'different'
            return result
        engine._http = respond
        proof = case.api.review_pairs(engine, pairs, same_cell=False)
        self.assertIn('shared_records', calls[0])
        restored = copy.deepcopy(calls[0]); restored.pop('round')
        pool = restored.pop('shared_records')
        for item in restored['pairs'].values():
            for key in ('metric', 'dimension_definitions'):
                value = item[key]
                if isinstance(value, dict) and set(value) == {'shared_record'}:
                    item[key] = pool[value['shared_record']]
            for side in ('left', 'right'):
                for ref, value in list(item[side]['evidence'].items()):
                    if isinstance(value, dict) and set(value) == {'shared_record'}:
                        item[side]['evidence'][ref] = pool[value['shared_record']]
        self.assertEqual(restored, original)
        self.assertEqual(original, before)
        self.assertEqual(case.api.verify_proof(proof, pairs, engine.standard, same_cell=False), {case.api.pair_id(pair)})
        self.assertLess(len(json.dumps(calls[0], ensure_ascii=False)), len(json.dumps(original, ensure_ascii=False)))
        # 原共享定义被篡改，即使更新文件哈希，也不能与原表及标准重建的请求一致。
        from 附注更新.表格 import file_hash
        from 附注更新.语义 import _hash
        tampered = copy.deepcopy(proof)
        task = json.loads(Path(proof['rounds'][0]['task']['path']).read_text(encoding='utf-8-sig'))
        task['payload']['shared_records'][next(iter(pool))] = '不是原文'
        task['task_id'] = _hash({k:v for k,v in task.items() if k != 'task_id'})
        path = folder/'篡改共享原文任务.json'; path.write_text(json.dumps(task,ensure_ascii=False),encoding='utf-8-sig')
        tampered['rounds'][0]['task'] = {'path':str(path),'sha256':file_hash(path)}
        with self.assertRaises(ValueError):
            case.api.verify_proof(tampered, pairs, engine.standard, same_cell=False)


if __name__ == '__main__':
    unittest.main()
