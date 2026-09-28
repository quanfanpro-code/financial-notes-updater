"""金标准本地版本维护测试，结构检查不代表业务语义审核。"""
import copy
import hashlib
import importlib
import json
import unittest
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REAL = ROOT / "金标准" / "财务报表附注语义金标准-v1.jsonl"


def slot(number=1):
    return {"id": f"C-N001-T001-S{number:03}", "status": "active", "scope": "consolidated",
        "note": {"id": "C-N001", "name": "货币资金", "meaning": "资金披露"},
        "table": {"id": "C-N001-T001", "name": "余额", "meaning": "余额构成"},
        "slot": {"name": f"项目{number}期末余额", "role": "measure"}, "value_type": "monetary",
        "dimensions": [{"name": "currency", "required": True}], "superseded_by": None,
        "aggregation": {"parent_slot_id": None, "component_slot_ids": []},
        "row_path": [f"项目{number}"], "column_path": ["期末余额"],
        "applicability": {"industries": ["general"], "optional": False},
        "扩展资料": {"人工审核意见": "测试内容原样保留"}}


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


class GoldTests(unittest.TestCase):
    def setUp(self):
        self.work = ROOT / "测试结果" / ("金标准_" + uuid.uuid4().hex[:8])
        self.work.mkdir(parents=True)

    def api(self):
        try:
            return importlib.import_module("附注更新.金标准")
        except ImportError:
            self.fail("金标准版本维护尚未实现")

    def file(self, rows):
        path = self.work / (uuid.uuid4().hex[:8] + ".jsonl")
        path.write_text("\n".join(json.dumps(row, ensure_ascii=False) for row in rows) + "\n", encoding="utf-8-sig")
        return path

    def record(self, path):
        return {"gold_path": str(path), "gold_hash": sha(path), "mappings": [{"slot_id": slot()["id"], "scope": "consolidated"}]}

    def unified_candidate(self):
        from 测试.测试统一语义 import records
        original=records();candidate=copy.deepcopy(original)
        metric=copy.deepcopy(candidate[-1]);metric['id']='固定资产账面净值'
        metric['definition']['measurement_basis']='账面净值';candidate.append(metric)
        candidate.append({'record_type':'dimension','id':'asset_location','meaning':'实际资产所在地',
                          'value_type':'text','domain':'资产所在地'})
        return original,candidate

    def test_unified_revision_preserves_layers_and_publishes_from_existing_entry(self):
        original,candidate=self.unified_candidate();current=self.file(original);new=self.file(candidate)
        preview=self.api().preview_revision(current,new)
        self.assertEqual(preview['added'],2)
        self.assertEqual({row['after']['record_type'] for row in preview['changes']},{'metric','dimension'})
        published=self.api().publish_revision(current,new,root=self.work,expected_hashes=preview['source_hashes'])
        self.assertEqual(Path(published['path']).read_bytes(),new.read_bytes())
        self.assertEqual(sha(current),preview['source_hashes']['current'])
        from 附注更新.界面 import format_revision_preview
        shown=format_revision_preview(preview)
        self.assertIn('维度定义',shown);self.assertIn('账面净值',shown)
        self.assertIn('资产所在地',shown)

    def test_unified_revision_rejects_duplicate_dimensions_and_overwritten_sources(self):
        original,candidate=self.unified_candidate();current=self.file(original)
        duplicate=copy.deepcopy(original[0]);duplicate['id']='同义期间维度';candidate.append(duplicate)
        with self.assertRaisesRegex(ValueError,'重复'):
            self.api().preview_revision(current,self.file(candidate))
        changed=copy.deepcopy(original);example=next(row for row in changed if row['record_type']=='source_example')
        example['text_evidence']=['改写了历史原文']
        with self.assertRaisesRegex(ValueError,'来源|证据'):
            self.api().preview_revision(current,self.file(changed))

    def test_redundant_metric_normalizes_only_against_its_gold_definition(self):
        api=self.api();row=slot()
        row["aliases"]={"column_labels":["年末数"]}
        expected={"slot_id":row["id"],"scope":"consolidated","semantic_field":"保留展示文字",
                  "dimensions":{"currency":"CNY","period":"2025-12-31","counterparty":"客户甲"}}
        for alias in (None,row["slot"]["name"],"期末余额","年末数"):
            source=copy.deepcopy(expected)
            if alias is not None:source["dimensions"]["metric"]=alias
            original=copy.deepcopy(source)
            self.assertEqual(api.normalize_mapping_metric(row,source),expected)
            self.assertEqual(source,original)
        for metric in ("坏账准备",None,123,""):
            with self.subTest(metric=metric),self.assertRaisesRegex(ValueError,"metric|指标"):
                api.normalize_mapping_metric(row,{**expected,"dimensions":{"metric":metric}})

    def test_declared_metric_and_distinct_slots_remain_distinct(self):
        api=self.api();row=slot()
        from 附注更新.表格 import semantic_key
        plain={"slot_id":row["id"],"scope":"consolidated","value_type":"monetary","dimensions":{"currency":"CNY"}}
        legacy=copy.deepcopy(plain);legacy["dimensions"]["metric"]=row["slot"]["name"]
        self.assertEqual(semantic_key(api.normalize_mapping_metric(row,plain)),
                         semantic_key(api.normalize_mapping_metric(row,legacy)))
        other=copy.deepcopy(plain);other["slot_id"]=slot(2)["id"]
        self.assertNotEqual(semantic_key(plain),semantic_key(other))
        row["dimensions"].append({"name":"metric","required":True,"open":True})
        one=copy.deepcopy(plain);one["dimensions"]["metric"]="指标一"
        two=copy.deepcopy(plain);two["dimensions"]["metric"]="指标二"
        self.assertEqual(api.normalize_mapping_metric(row,one),one)
        self.assertNotEqual(semantic_key(api.normalize_mapping_metric(row,one)),
                            semantic_key(api.normalize_mapping_metric(row,two)))

    def test_standalone_requires_explicit_reviewed_applicability(self):
        api=self.api();row=slot()
        self.assertEqual(api.allowed_scopes(row),{"consolidated"})
        self.assertIs(api.validate_mapping_scope(row,{"scope":"consolidated"}),True)
        self.assertIs(api.validate_mapping_scope(row,{"scope":"consolidated","definition_scope":"consolidated"}),True)
        mapping={"scope":"standalone","definition_scope":"consolidated"}
        with self.assertRaisesRegex(ValueError,"适用|口径"):api.validate_mapping_scope(row,mapping)
        row["applicability"]["report_scopes"]=[{"scope":"standalone","reason":"库存现金余额指标不依赖合并抵销或少数股东身份",
            "reviewed_by":"测试审核人","reviewed_at":"2026-09-19T21:00:00+08:00"}]
        self.assertEqual(api.allowed_scopes(row),{"consolidated","standalone"})
        self.assertIs(api.validate_mapping_scope(row,mapping),True)
        for invalid in [{"scope":"standalone"},{"scope":"standalone","definition_scope":"parent"},
                        {"scope":"parent","definition_scope":"consolidated"},
                        {"scope":"consolidated","definition_scope":"parent"},{}]:
            with self.subTest(invalid=invalid):
                with self.assertRaisesRegex(ValueError,"口径"):api.validate_mapping_scope(row,invalid)
        saved=self.file([row]);loaded=api.read_standard(saved)[row["id"]]
        self.assertEqual(loaded,row)

    def test_applicability_metadata_is_strict_and_not_a_scope_wildcard(self):
        api=self.api()
        valid={"scope":"standalone","reason":"具体适用性依据","reviewed_by":"测试审核人","reviewed_at":"2026-09-19"}
        invalids=[None,"standalone",[None],[dict(valid,scope="parent")],[dict(valid,scope="consolidated")],
                  [dict(valid,scope="unknown")],[valid,valid],[dict(valid,reason=" ")],
                  [dict(valid,reviewed_by="")],[dict(valid,reviewed_at="2026-02-30")],
                  [{k:v for k,v in valid.items() if k!="reason"}],[dict(valid,approved=True)]]
        for extension in invalids:
            with self.subTest(extension=extension):
                row=slot();row["applicability"]["report_scopes"]=extension
                with self.assertRaisesRegex(ValueError,"适用|applicability"):
                    api.read_standard(self.file([row]))
        for container in (None, [], "standalone"):
            row=slot();row["applicability"]=container
            with self.assertRaisesRegex(ValueError,"applicability"):
                api.read_standard(self.file([row]))

    def test_compatibility_checks_actual_scope_and_reviewed_extension(self):
        api=self.api();row=slot();extension={"scope":"standalone","reason":"具体适用性依据",
            "reviewed_by":"测试审核人","reviewed_at":"2026-09-19"}
        row["applicability"]["report_scopes"]=[extension];old=self.file([row]);record=self.record(old)
        record["mappings"][0].update(scope="standalone",definition_scope="consolidated")
        self.assertIs(api.ensure_compatible(record,self.file([row,slot(2)])),True)
        forged=self.record(self.file([slot()]))
        forged["mappings"][0].update(scope="standalone",definition_scope="consolidated")
        with self.assertRaisesRegex(ValueError,"适用|口径"):
            api.ensure_compatible(forged,forged["gold_path"])
        changed=copy.deepcopy(row);changed["applicability"]["report_scopes"][0]["reason"]="审核依据已改变"
        with self.assertRaisesRegex(ValueError,"定义已改变"):
            api.ensure_compatible(record,self.file([changed]))
        missing=copy.deepcopy(record);missing["mappings"][0].pop("definition_scope")
        with self.assertRaisesRegex(ValueError,"口径"):api.ensure_compatible(missing,old)

    def test_existing_full_standard_is_accepted_without_modification(self):
        api = self.api()
        path = REAL if REAL.is_file() else Path(r"D:\BaiduSyncdisk\数字资产\审计数据自动搬运系统\语义金标准\财务报表附注语义金标准-v1.jsonl")
        before = sha(path)
        items = api.read_standard(path)
        self.assertEqual(len(items), 12794)
        self.assertEqual(sum(item["status"] == "active" for item in items.values()), 9726)
        self.assertEqual(sha(path), before)

    def test_duplicate_id_and_invalid_active_structures_are_rejected(self):
        api = self.api()
        with self.assertRaisesRegex(ValueError, "重复"):
            api.read_standard(self.file([slot(), slot()]))
        for field, value in [("scope", "unknown"), ("note", {}), ("table", {}), ("slot", {}),
                ("value_type", "unknown"), ("dimensions", [{"name": "currency", "required": "yes"}])]:
            with self.subTest(field=field):
                item = slot();item[field] = value
                with self.assertRaises(ValueError):
                    api.read_standard(self.file([item]))

    def test_dangling_references_are_rejected_including_nested_variants(self):
        api = self.api()
        for field in ("aggregation", "calculation"):
            with self.subTest(field=field):
                item = slot()
                item[field] = {"component_variants": [{"component_slot_ids": [slot(99)["id"]]}]}
                with self.assertRaisesRegex(ValueError, "引用"):
                    api.read_standard(self.file([item]))
        item = slot();item.update(status="superseded", superseded_by=[slot(99)["id"]])
        with self.assertRaisesRegex(ValueError, "引用"):
            api.read_standard(self.file([item, slot(2)]))

    def test_preview_reports_addition_change_retirement_and_keeps_originals(self):
        api = self.api()
        current = self.file([slot(1), slot(2), slot(3)])
        changed = slot(1);changed["slot"]["name"] = "修订指标定义"
        retired = slot(2);retired["status"] = "deprecated"
        candidate = self.file([changed, retired, slot(3), slot(4)])
        before = (sha(current), sha(candidate))
        preview = api.preview_revision(current, candidate)
        self.assertEqual((preview["added"], preview["changed"], preview["retired"]), (1, 1, 1))
        self.assertEqual(preview["added_ids"], [slot(4)["id"]])
        self.assertEqual(preview["changed_ids"], [slot(1)["id"]])
        self.assertEqual(preview["retired_ids"], [slot(2)["id"]])
        self.assertEqual(preview["source_hashes"], {"current": before[0], "candidate": before[1]})
        self.assertEqual(len(preview["changes"]), 3)
        revised = next(change for change in preview["changes"] if change["kind"] == "changed")
        self.assertEqual(revised["before"]["note_definition"]["meaning"], "资金披露")
        self.assertEqual(revised["after"]["slot_definition"]["name"], "修订指标定义")
        self.assertIn("aggregation", revised["after"])
        self.assertIn("扩展资料", revised["after"])
        self.assertEqual((sha(current), sha(candidate)), before)

    def test_complete_revision_cannot_silently_omit_existing_id(self):
        api = self.api()
        with self.assertRaisesRegex(ValueError, "遗漏|缺少"):
            api.preview_revision(self.file([slot(1), slot(2)]), self.file([slot(1)]))

    def test_revision_cannot_silently_drop_existing_extension_fields(self):
        api = self.api()
        old = self.file([slot(1)])
        modified = slot(1)
        modified.pop("扩展资料")
        with self.assertRaisesRegex(ValueError, "遗漏已有字段"):
            api.preview_revision(old, self.file([modified]))

    def test_publish_creates_unique_version_and_report_preserving_unknown_fields(self):
        api = self.api()
        current = self.file([slot(1)])
        candidate = self.file([slot(1), slot(2)])
        original = (current.read_bytes(), candidate.read_bytes())
        preview = api.preview_revision(current, candidate)
        first = api.publish_revision(current, candidate, root=self.work, expected_hashes=preview["source_hashes"])
        second = api.publish_revision(current, candidate, root=self.work, expected_hashes=preview["source_hashes"])
        self.assertNotEqual(first["path"], second["path"])
        self.assertEqual(Path(first["path"]).parent, self.work / "金标准" / "版本")
        self.assertEqual(Path(first["path"]).read_bytes(), original[1])
        self.assertEqual(first["sha256"], sha(candidate))
        self.assertTrue(Path(first["report_path"]).is_file())
        self.assertEqual(first["report"]["added"], 1)
        self.assertIn("语义", json.dumps(first["report"], ensure_ascii=False))
        self.assertEqual((current.read_bytes(), candidate.read_bytes()), original)

    def test_publish_rechecks_preview_hash_before_writing(self):
        api = self.api()
        current = self.file([slot(1)])
        candidate = self.file([slot(1), slot(2)])
        preview = api.preview_revision(current, candidate)
        changed_candidate = self.file([slot(1), slot(2), slot(3)])
        with self.assertRaisesRegex(ValueError, "变化|哈希|预览"):
            api.publish_revision(current, changed_candidate, root=self.work, expected_hashes=preview["source_hashes"])
        self.assertFalse((self.work / "金标准" / "版本").exists())

    def test_compatible_append_reuses_but_changed_retired_or_missing_definition_does_not(self):
        api = self.api()
        old = self.file([slot(1)])
        self.assertIs(api.ensure_compatible(self.record(old), self.file([slot(1), slot(2)])), True)
        variants = []
        changed = slot(1);changed["slot"]["name"] = "同ID不同定义";variants.append([changed, slot(2)])
        changed = slot(1);changed["扩展资料"]["人工审核意见"] = "变更内容";variants.append([changed, slot(2)])
        retired = slot(1);retired["status"] = "deprecated";variants.append([retired, slot(2)])
        variants.append([slot(2)])
        for rows in variants:
            with self.subTest(rows=rows):
                with self.assertRaisesRegex(ValueError, "定义|停用|不存在|缺少"):
                    api.ensure_compatible(self.record(old), self.file(rows))

    def test_old_mapping_gold_hash_is_verified_even_for_same_version(self):
        api = self.api()
        old = self.file([slot(1)]);record = self.record(old);record["gold_hash"] = "0" * 64
        with self.assertRaisesRegex(ValueError, "变化|哈希|不一致"):
            api.ensure_compatible(record, old)


if __name__ == "__main__":
    unittest.main()
