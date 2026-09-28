# coding: utf-8
"""旧成果显式转换的行为测试：不改旧值，不借新定义丢掉旧维度。"""
import copy
import importlib
import importlib.util
import unittest

from 附注更新.统一语义 import definition_hash, validate_records
from 测试.测试统一语义 import records, fact


def fixture():
    old_definition = {"id": "旧办公设备余额", "status": "active", "value_type": "monetary", "slot": {"name": "办公设备期末原值"}}
    rows = records()
    conversion = {"record_type": "legacy_conversion", "id": "旧办公设备转换", "legacy_id": old_definition["id"],
                  "source_gold_hash": "c" * 64, "legacy_definition_hash": definition_hash(old_definition), "status": "approved",
                  "target": {"record_type": "metric_fact", "metric_id": "固定资产原值", "dimension_bindings": {
                      key: {"op": "dimension", "key": key} for key in fact()["dimensions"]}},
                  "evidence_refs": ["来源一"], "review": {"reviewer": "合成复核", "reason": "测试显式迁移"}}
    conversion["target"]["dimension_bindings"]["asset_selection"] = {"op": "constant", "value": {
        "domain": "固定资产", "mode": "members", "members": ["办公设备"], "completeness": "complete"}}
    rows.append(conversion)
    standard = validate_records(rows); standard["sha256"] = "d" * 64
    original = fact(); original.pop("record_type"); original.pop("metric_id")
    original["legacy_id"] = old_definition["id"]
    original["legacy_mapping_hash"] = "e" * 64
    del original["dimensions"]["asset_selection"]; del original["dimension_evidence"]["asset_selection"]
    return original, old_definition, standard


class MigrationTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(importlib.util.find_spec("附注更新.统一迁移"), "缺少显式迁移执行入口")
        self.api = importlib.import_module("附注更新.统一迁移")

    def test_preserves_source_value_original_review_and_three_bindings(self):
        original, definition, standard = fixture(); before = copy.deepcopy(original)
        result = self.api.migrate_fact(original, definition, "c" * 64, standard, {})
        self.assertEqual(original, before)
        self.assertEqual(result["fact"]["raw_value"], "125.00")
        self.assertEqual(result["fact"]["dimensions"]["asset_selection"]["members"], ["办公设备"])
        self.assertEqual(result["receipt"]["source_mapping_hash"], "e" * 64)
        self.assertEqual(result["receipt"]["source_gold_hash"], "c" * 64)
        self.assertEqual(result["receipt"]["source_provenance"], original["recognition_provenance"])
        self.assertEqual(result["receipt"]["target_gold_hash"], "d" * 64)

    def test_refuses_changed_definition_version_and_unreviewed_conversion(self):
        original, definition, standard = fixture()
        changed = copy.deepcopy(definition); changed["slot"]["name"] = "净额"
        for source_definition, source_hash in ((changed, "c" * 64), (definition, "b" * 64)):
            with self.subTest(source_hash=source_hash), self.assertRaises(ValueError):
                self.api.migrate_fact(original, source_definition, source_hash, standard, {})
        standard["legacy_conversion"]["旧办公设备转换"]["status"] = "unresolved"
        standard["legacy_conversion"]["旧办公设备转换"]["target"] = None
        with self.assertRaises(ValueError):
            self.api.migrate_fact(original, definition, "c" * 64, standard, {})

    def test_does_not_overwrite_conflicting_real_dimension_or_discard_extra(self):
        for key, value in (("asset_selection", {"domain": "固定资产", "mode": "all", "completeness": "complete"}), ("project", "项目甲")):
            original, definition, standard = fixture(); original["dimensions"][key] = value
            original["dimension_evidence"][key] = ["原格已确认维度"]
            with self.subTest(key=key), self.assertRaises(ValueError):
                self.api.migrate_fact(original, definition, "c" * 64, standard, {})

    def test_context_must_have_evidence_and_cannot_supply_example_year(self):
        original, definition, standard = fixture()
        rule = standard["legacy_conversion"]["旧办公设备转换"]["target"]["dimension_bindings"]
        rule["period"] = {"op": "context", "key": "实际期末日期"}
        del original["dimensions"]["period"]; del original["dimension_evidence"]["period"]
        for context in ({}, {"实际期末日期": {"value": {"kind": "instant", "date": "2025-12-31"}, "evidence_refs": []}}):
            with self.subTest(context=context), self.assertRaises(ValueError):
                self.api.migrate_fact(original, definition, "c" * 64, standard, context)
        context = {"实际期末日期": {"value": {"kind": "instant", "date": "2025-12-31"}, "evidence_refs": ["原报告日期说明"]}}
        result = self.api.migrate_fact(original, definition, "c" * 64, standard, context)
        self.assertEqual(result["fact"]["dimensions"]["period"]["date"], "2025-12-31")

    def test_dimension_alias_copy_is_checked_without_losing_original_dimension(self):
        original, definition, standard = fixture()
        original["dimensions"]["currency_original"] = original["dimensions"].pop("currency")
        original["dimension_evidence"]["currency_original"] = original["dimension_evidence"].pop("currency")
        standard["legacy_conversion"]["旧办公设备转换"]["target"]["dimension_bindings"]["currency"] = {"op": "dimension", "key": "currency_original"}
        result = self.api.migrate_fact(original, definition, "c" * 64, standard, {})
        self.assertEqual(result["fact"]["dimensions"]["currency"], "CNY")
        self.assertIn("currency_original", result["receipt"]["consumed_dimensions"])


if __name__ == "__main__":
    unittest.main()
