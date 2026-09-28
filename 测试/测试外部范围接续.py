"""核验真实来源范围确认进入正式识别与接续；全部使用本地虚构资料。"""
import copy
import json
import unittest
from pathlib import Path
from unittest.mock import patch

from 测试 import 测试外部范围 as fixtures
from 附注更新 import 分步流程 as flow
from 附注更新.语义 import SemanticEngine
from 附注更新.表格 import read_workbook, file_hash


class ExternalScopeResumeTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.ExternalScopeTests()
        self.fixture.setUp()
        self.work = self.fixture.work

    def selection(self):
        opinion = self.fixture.opinion()
        opinion["areas"].pop(1)
        opinion["areas"][1]["range"] = "A3:B5"
        opinion["areas"][1]["reason"] = "本地虚构人员统计区域及其前置留白，不是货币资金业务"
        preview = self.fixture.preview(opinion)
        result = self.fixture.api().apply_external_scope_review(preview, self.work, confirmed=True,
                    confirm_method="本地集成验收的模拟人工确认")
        return json.loads(Path(result["files"]["b_scope"]).read_text(encoding="utf-8-sig"))

    def engine(self):
        return SemanticEngine({"base_url":"http://127.0.0.1:1/v1", "model":"本地测试", "review":True,
                               "layout_mode":"cells_v1"}, str(self.fixture.gold), str(self.work))

    def recognize(self, selection):
        engine = self.engine()
        snapshot = read_workbook(self.fixture.source)
        owned = []
        def local_layout(payload, sheet, candidates, batch):
            owned.extend(candidates)
            return {"tables":[], "non_business":[], "unresolved":[]}
        with patch("附注更新.识别范围.select_source_scope", side_effect=AssertionError("外部确认不可冒充模型双轮")), \
             patch.object(engine, "_request_layout_cells", side_effect=local_layout), \
             patch.object(engine, "_request", side_effect=AssertionError("本地验收禁止网络")):
            result = engine.recognize(snapshot, target_scope=selection["target_meanings"], scope_selection=selection)
        return engine, snapshot, result, owned

    def test_confirmed_blank_exclusions_are_not_added_back(self):
        selection = self.selection()
        engine, snapshot, result, owned = self.recognize(selection)
        self.assertEqual(set(owned), {"A1", "B1", "A2", "B2"})
        self.assertEqual({x["cell"] for x in result["out_of_scope"]}, {"A3", "B3", "A4", "B4", "A5", "B5"})
        self.assertEqual(len(result["unresolved"]), 4)
        self.assertEqual(len(snapshot["sheets"][0]["cells"]), 10)
        self.assertEqual(result["recognition_basis"]["candidate_blanks"], "source_grid_v1")

    def test_formal_load_and_nonformal_resume_recheck_external_proof(self):
        selection = self.selection()
        engine, snapshot, result, _ = self.recognize(selection)
        record = flow._record(snapshot, result, "B", self.fixture.gold, self.fixture.source, [], None)
        path = self.fixture.save("正式映射", record)
        loaded, actual = flow._load_mapping(path, "B", complete=False)
        self.assertEqual(len(loaded["out_of_scope"]), 6)
        missing=copy.deepcopy(record)
        missing.pop("scope_selection")
        with self.assertRaises(ValueError):
            flow._load_mapping(self.fixture.save("缺失范围来源",missing),"B",complete=False)
        kept, proof = engine._reuse_previous({"mapping_path":str(path), "mapping_sha256":file_hash(path)},
                                             copy.deepcopy(actual), result["recognition_basis"], selection)
        self.assertFalse(proof["recheck"])
        bad = copy.deepcopy(record)
        bad["scope_selection"]["external_scope_review"]["confirm_method"] = "已改变的确认来源"
        with self.assertRaises(ValueError):
            flow._load_mapping(self.fixture.save("篡改正式映射", bad), "B", complete=False)
        broken = copy.deepcopy(result)
        broken["scope_selection"]["external_scope_review"]["confirm_method"] = "已改变的确认来源"
        bare = {k:v for k,v in broken.items() if k not in {"evidence_path", "evidence_hash"}}
        evidence = self.fixture.save("篡改中途成果", bare)
        broken.update(evidence_path=str(evidence), evidence_hash=file_hash(evidence))
        with self.assertRaises(ValueError):
            engine._reuse_previous(broken, copy.deepcopy(actual), result["recognition_basis"], selection)

        for change in (lambda r:r.pop("scope_selection"),
                       lambda r:r["out_of_scope"][0].update(reason="改写了原范围意见")):
            broken=copy.deepcopy(result)
            change(broken)
            bare={k:v for k,v in broken.items() if k not in {"evidence_path","evidence_hash"}}
            evidence=self.fixture.save("来源缺口回归",bare)
            broken.update(evidence_path=str(evidence),evidence_hash=file_hash(evidence))
            with self.subTest(change=change),self.assertRaises(ValueError):
                engine._reuse_previous(broken,copy.deepcopy(actual),result["recognition_basis"],selection)


if __name__ == "__main__":
    unittest.main()
