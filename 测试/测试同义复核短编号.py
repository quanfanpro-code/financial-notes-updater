# coding: utf-8
"""模型回执使用短编号，仍须绑定原事实身份、两轮意见和各自证据。"""
import copy
import json
import unittest
from pathlib import Path
from openpyxl import Workbook
from 测试 import 测试同义复核 as fixtures


class ShortReviewTests(unittest.TestCase):
    def fixture(self):
        case = fixtures.EquivalenceTests()
        case.setUp()
        folder, pair = case.fixture()
        engine = case.engine(folder)
        return case, engine, pair

    def test_short_request_keeps_two_original_identities_and_report_decisions(self):
        case, engine, first = self.fixture()
        second = {'left': copy.deepcopy(first['right']), 'right': copy.deepcopy(first['left']),
            'left_path': first['right_path'], 'right_path': first['left_path']}
        captured = []
        def respond(system, payload):
            captured.append(copy.deepcopy(payload))
            result = case.response(payload)
            result['pairs']['P02']['decision'] = 'different'
            return result
        engine._http = respond
        proof = case.api.review_pairs(engine, [first, second], same_cell=False, numbered=True)
        self.assertEqual(len(captured), 2)
        self.assertEqual(set(captured[0]['pairs']), {'P01', 'P02'})
        self.assertEqual(set(captured[0]['pair_ids'].values()), {case.api.pair_id(first), case.api.pair_id(second)})
        self.assertEqual(case.api.verify_proof(proof, [first, second], engine.standard, same_cell=False),
                         {captured[0]['pair_ids']['P01']})
        self.assertEqual(proof['pairs'], [first, second])
        book = Workbook()
        case.api.append_review_sheet(book, [proof])
        adopted = [row[9].value for row in list(book['维度同义复核'].rows)[1:]]
        self.assertCountEqual(adopted, ['两轮同意', '未确认同义'])
        book.close()
        changed = copy.deepcopy(proof); changed['schema'] = case.api.PROOF_SCHEMA
        with self.assertRaises(ValueError):
            case.api.verify_proof(changed, [first, second], engine.standard, same_cell=False)

    def test_short_number_does_not_allow_other_facts_evidence(self):
        case, engine, pair = self.fixture()
        def respond(system, payload):
            answer = case.response(payload)
            answer['pairs']['P01']['left_evidence'] = ['当前表!A999']
            return answer
        engine._http = respond
        with self.assertRaisesRegex(ValueError, '其他事实'):
            case.api.review_pairs(engine, [pair], same_cell=False, numbered=True)

    def test_partial_short_review_never_completes_missing_identity(self):
        case, engine, pair = self.fixture()
        for side in ('left', 'right'):
            fact = pair[side]
            known = {'metric_id': fact['metric_id'], 'dimensions': {'asset_selection': fact['dimensions']['asset_selection']},
                'dimension_evidence': {'asset_selection': ['当前表!A1']}, 'evidence_cells': ['当前表!A1'],
                'missing_dimensions': ['entity', 'period', 'report_scope', 'currency', 'unit_scale']}
            pair[side] = case.api.partial_review_fact(known, pair['left']['source_reference'])
        proof = case.api.review_pairs(engine, [pair], same_cell=True, numbered=True)
        self.assertEqual(case.api.verify_proof(proof, [pair], engine.standard, same_cell=True), {case.api.pair_id(pair)})
        self.assertFalse(case.api.reviewable(pair['left'], pair['right'], engine.standard, same_cell=False))
        self.assertEqual(len(proof['pairs'][0]['left']['missing_dimensions']), 5)


if __name__ == '__main__':
    unittest.main()
