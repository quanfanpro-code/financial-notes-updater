# coding: utf-8
"""验证槽位、维度定义和实例分层；例值为合成数据，不是模型验收答案。"""
import copy
import importlib
import importlib.util
import json
import unittest
from pathlib import Path


NAME = "附注更新.统一语义"


def records():
    definitions = [
        ("period", "实际报告期间", "period", "报告期间"),
        ("entity", "实际报告主体", "text", "报告主体"),
        ("report_scope", "实际报表口径", "text", "报告口径"),
        ("currency", "币种", "text", "币种"),
        ("unit_scale", "金额倍率", "number", "金额倍率"),
        ("asset_selection", "所计量的固定资产集合", "selection", "固定资产"),
    ]
    rows = [{"record_type": "dimension", "id": key, "meaning": meaning,
             "value_type": kind, "domain": domain} for key, meaning, kind, domain in definitions]
    rows.append({"record_type": "source_example", "id": "来源一",
                 "origin": {"file_hash": "a" * 64, "location": "原表!C7", "scope": "consolidated"},
                 "text_evidence": ["固定资产原值", "期末"], "review_refs": ["合成测试来源"]})
    rows.append({"record_type": "metric", "id": "固定资产原值", "status": "active",
                 "definition": {"economic_object": "固定资产", "measure": "余额",
                                "measurement_basis": "账面原值", "value_type": "monetary",
                                "time_kind": "instant", "constraints": []},
                 "dimension_refs": [{"id": key, "required": True} for key, *_ in definitions],
                 "source_refs": ["来源一"]})
    return rows


def fact():
    dimensions = {"period": {"kind": "instant", "date": "2025-12-31"},
                  "entity": "合成企业甲", "report_scope": "standalone", "currency": "CNY",
                  "unit_scale": "1", "asset_selection": {"domain": "固定资产", "mode": "all", "completeness": "complete"}}
    return {"record_type": "metric_fact", "metric_id": "固定资产原值", "dimensions": dimensions,
            "dimension_evidence": {key: ["原资料对应表头或报告信息"] for key in dimensions},
            "source_reference": {"file_hash": "b" * 64, "location": "当前表!D8"},
            "raw_value_state": "value", "raw_value": "125.00",
            "recognition_provenance": {"method": "synthetic", "record_id": "规则测试"}}


class UnifiedSemanticsTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(importlib.util.find_spec(NAME), "缺少统一语义模块，无法保证槽位与维度分层")
        self.api = importlib.import_module(NAME)
        self.standard = self.api.validate_records(records())

    def test_layout_and_value_change_do_not_change_business_identity(self):
        left, right = fact(), fact()
        right["source_reference"]["location"] = "转置后的新表!AJ91"
        right["raw_value"] = "256.00"
        right["dimension_evidence"]["period"] = ["另一个实际表头"]
        self.assertEqual(self.api.semantic_key(left, self.standard), self.api.semantic_key(right, self.standard))

    def test_instance_changes_do_not_require_new_metric_but_do_change_fact(self):
        for key, replacement in [("report_scope", "consolidated"),
                                 ("period", {"kind": "instant", "date": "2024-12-31"}),
                                 ("asset_selection", {"domain": "固定资产", "mode": "members", "members": ["办公设备"], "completeness": "complete"})]:
            with self.subTest(key=key):
                other = fact()
                other["dimensions"][key] = replacement
                self.api.validate_fact(other, self.standard)
                self.assertNotEqual(self.api.semantic_key(fact(), self.standard), self.api.semantic_key(other, self.standard))

    def test_主体原文保留但不作为额外对应门槛(self):
        other = fact()
        other['dimensions']['entity'] = '合成企业乙'
        self.api.validate_fact(other, self.standard)
        self.assertEqual(self.api.semantic_key(fact(), self.standard), self.api.semantic_key(other, self.standard))
        self.assertEqual(other['dimensions']['entity'], '合成企业乙')

    def test_source_scope_is_not_business_exclusivity(self):
        self.api.validate_fact(fact(), self.standard)

    def test_unknown_or_incomplete_population_never_matches_itself(self):
        for selection in [{"domain": "固定资产", "mode": "unknown", "completeness": "unknown"},
                          {"domain": "固定资产", "mode": "members", "members": ["设备"], "completeness": "partial"}]:
            other = fact()
            other["dimensions"]["asset_selection"] = selection
            self.api.validate_fact(other, self.standard)
            with self.assertRaises(ValueError):
                self.api.semantic_key(other, self.standard)

    def test_complement_retains_base_and_exclusions_and_set_order_is_irrelevant(self):
        left, right = fact(), fact()
        left["dimensions"]["asset_selection"] = {"domain": "固定资产", "mode": "complement", "base": "全部固定资产",
                                                  "excluded": ["房屋", "运输工具"], "completeness": "complete"}
        right["dimensions"]["asset_selection"] = copy.deepcopy(left["dimensions"]["asset_selection"])
        right["dimensions"]["asset_selection"]["excluded"].reverse()
        self.assertEqual(self.api.semantic_key(left, self.standard), self.api.semantic_key(right, self.standard))
        right["dimensions"]["asset_selection"]["base"] = "融资租赁固定资产"
        self.assertNotEqual(self.api.semantic_key(left, self.standard), self.api.semantic_key(right, self.standard))

    def test_missing_extra_and_unproved_dimensions_rejected(self):
        variants = []
        one = fact(); del one["dimensions"]["period"]; variants.append(one)
        one = fact(); one["dimensions"]["任意指标"] = "净额"; variants.append(one)
        one = fact(); one["dimension_evidence"]["entity"] = []; variants.append(one)
        one = fact(); one["dimensions"]["period"]["date"] = "2025-02-30"; variants.append(one)
        one = fact(); one["dimensions"]["unit_scale"] = True; variants.append(one)
        for other in variants:
            with self.subTest(other=other), self.assertRaises(ValueError):
                self.api.validate_fact(other, self.standard)

    def test_measurement_basis_distinguishes_metrics_and_identical_new_id_rejected(self):
        rows = records()
        duplicate = copy.deepcopy(rows[-1]); duplicate["id"] = "另一个期初编号"
        with self.assertRaises(ValueError):
            self.api.validate_records(rows + [duplicate])
        duplicate["definition"]["measurement_basis"] = "减值准备余额"
        standard = self.api.validate_records(rows + [duplicate])
        other = fact(); other["metric_id"] = duplicate["id"]
        self.assertNotEqual(self.api.semantic_key(fact(), standard), self.api.semantic_key(other, standard))

    def test_definition_cannot_embed_layout_or_instance_fields(self):
        for key, value in [("row_path", ["期末"]), ("actual_values", {"客户": "合成甲"}), ("scope", "consolidated")]:
            rows = records(); rows[-1][key] = value
            with self.subTest(key=key), self.assertRaises(ValueError):
                self.api.validate_records(rows)

    def test_named_dimension_cell_keeps_owner_without_becoming_metric_fact(self):
        binding = {"record_type": "dimension_binding", "dimension_id": "entity", "value": "合成企业甲",
                   "owner_reference": "该载体内的业务对象一", "source_reference": fact()["source_reference"],
                   "recognition_provenance": fact()["recognition_provenance"]}
        self.api.validate_fact(binding, self.standard)
        with self.assertRaises(ValueError):
            self.api.semantic_key(binding, self.standard)

    def test_actual_text_or_date_disclosure_can_remain_a_metric(self):
        rows = records()
        rows[-1]["definition"] = {"economic_object": "借款合同", "measure": "约定还款日", "measurement_basis": "合同约定",
                                    "value_type": "date", "time_kind": "none", "constraints": []}
        rows[-1]["dimension_refs"] = [{"id": "entity", "required": True}, {"id": "report_scope", "required": True}]
        standard = self.api.validate_records(rows)
        other = fact(); other["dimensions"] = {"entity": "合成企业甲", "report_scope": "standalone"}
        other["dimension_evidence"] = {"entity": ["主体信息"], "report_scope": ["单户说明"]}; other["raw_value"] = "2027-03-01"
        self.api.validate_fact(other, standard)

    def test_legacy_conversion_needs_review_and_all_three_source_bindings(self):
        old = {"id": "C-N001-T001-S001", "status": "active", "slot": {"name": "期末原值"}}
        conversion = {"record_type": "legacy_conversion", "id": "转换一", "source_gold_hash": "c" * 64,
                      "legacy_id": old["id"], "legacy_definition_hash": self.api.definition_hash(old),
                      "status": "approved", "target": {"record_type": "dimension_binding", "dimension_id": "entity",
                          "value_rule": {"op": "context", "key": "主体原文"}, "owner_rule": {"op": "context", "key": "所属对象"}},
                      "evidence_refs": ["来源一"], "review": {"reviewer": "独立合成复核", "reason": "这里只验证协议"}}
        standard = self.api.validate_records(records() + [conversion])
        self.api.validate_legacy_bindings(standard, {old["id"]: old}, "c" * 64)
        changed = copy.deepcopy(old); changed["slot"]["name"] = "本期折旧"
        with self.assertRaises(ValueError):
            self.api.validate_legacy_bindings(standard, {old["id"]: changed}, "c" * 64)
        with self.assertRaises(ValueError):
            self.api.validate_legacy_bindings(standard, {old["id"]: old}, "d" * 64)
        conversion["review"]["reason"] = ""
        with self.assertRaises(ValueError):
            self.api.validate_records(records() + [conversion])

    def test_ratio_requires_business_numerator_and_denominator(self):
        rows = records(); ratio = copy.deepcopy(rows[-1]); ratio["id"] = "固定资产占比"
        ratio["definition"]["value_type"] = "percentage"
        ratio["definition"]["measure"] = "余额占比"
        with self.assertRaises(ValueError):
            self.api.validate_records(rows + [ratio])

    def test_period_kind_and_raw_blank_are_not_silently_coerced(self):
        other = fact(); other["dimensions"]["period"] = {"kind": "duration", "start": "2025-01-01", "end": "2025-12-31"}
        with self.assertRaises(ValueError):
            self.api.validate_fact(other, self.standard)
        other = fact(); other["raw_value_state"] = "blank"; other["raw_value"] = 0
        with self.assertRaises(ValueError):
            self.api.validate_fact(other, self.standard)

    def test_same_id_with_changed_definition_cannot_match_across_versions(self):
        rows = records(); rows[-1]["definition"]["measurement_basis"] = "扣除累计折旧后"
        changed = self.api.validate_records(rows)
        self.assertNotEqual(self.api.semantic_key(fact(), self.standard), self.api.semantic_key(fact(), changed))
        rows = records(); rows[1]["meaning"] = "担保人主体"
        changed = self.api.validate_records(rows)
        self.assertNotEqual(self.api.semantic_key(fact(), self.standard), self.api.semantic_key(fact(), changed))

    def test_known_money_units_match_and_raw_value_is_preserved(self):
        source, target = fact(), fact()
        source["dimensions"]["unit_scale"] = "10000"; source["raw_value"] = "1.25"
        self.assertEqual(self.api.semantic_key(source, self.standard), self.api.semantic_key(target, self.standard))
        self.assertEqual(self.api.value_for_target(source, target, self.standard), "1.25")
        source["raw_value_state"] = "unknown"; source["raw_value"] = None
        with self.assertRaises(ValueError):
            self.api.value_for_target(source, target, self.standard)

    def test_full_review_register_does_not_approve_by_role_or_lose_history(self):
        self.assertTrue(hasattr(self.api, "prepare_legacy_review"), "缺少全量转换审查建账")
        old = {"旧一": {"id": "旧一", "status": "superseded", "superseded_by": ["旧二"], "slot": {"name": "旧期末", "role": "dimension"}},
               "旧二": {"id": "旧二", "status": "active", "slot": {"name": "新名称", "role": "descriptive"}, "value_type": "monetary"}}
        before = copy.deepcopy(old)
        rows = self.api.prepare_legacy_review(old, "f" * 64)
        standard = self.api.validate_records(rows)
        self.api.validate_legacy_bindings(standard, old, "f" * 64)
        self.assertEqual(len(standard["legacy_conversion"]), 2)
        self.assertEqual(standard["metric"], {})
        self.assertEqual({row["legacy_id"] for row in standard["legacy_conversion"].values()}, {"旧一", "旧二"})
        self.assertTrue(all(row["target"] is None and row["status"] == "unresolved" for row in standard["legacy_conversion"].values()))
        self.assertEqual(old, before)

    def test_financial_metrics_require_actual_entity_and_scope(self):
        for omitted in ("entity", "report_scope"):
            rows = records(); rows[-1]["dimension_refs"] = [ref for ref in rows[-1]["dimension_refs"] if ref["id"] != omitted]
            with self.subTest(omitted=omitted), self.assertRaises(ValueError):
                self.api.validate_records(rows)

    def test_true_business_constraint_is_enforced_instead_of_inferred_from_source(self):
        rows = records()
        rows[-1]["definition"]["constraints"] = [{"dimension_id": "report_scope", "allowed_values": ["consolidated"]}]
        standard = self.api.validate_records(rows)
        with self.assertRaises(ValueError):
            self.api.semantic_key(fact(), standard)
        other = fact(); other["dimensions"]["report_scope"] = "consolidated"
        self.api.semantic_key(other, standard)

    def test_reserved_dimension_types_cannot_turn_currency_into_boolean(self):
        for identifier, bad_type in (("currency", "boolean"), ("unit_scale", "text"), ("entity", "number")):
            rows = records()
            next(row for row in rows if row["id"] == identifier)["value_type"] = bad_type
            with self.subTest(identifier=identifier), self.assertRaises(ValueError):
                self.api.validate_records(rows)

    def test_ratio_checks_binding_domain_numeric_meaning_and_cycles(self):
        rows = records()
        ratio = copy.deepcopy(rows[-1]); ratio["id"] = "占比"
        ratio["definition"].update(value_type="percentage", measure="比例", ratio={
            side: {"metric_id": "固定资产原值", "dimension_bindings": {ref["id"]: {"op": "dimension", "key": ref["id"]} for ref in ratio["dimension_refs"]}}
            for side in ("numerator", "denominator")})
        self.api.validate_records(rows + [ratio])
        invalid = copy.deepcopy(ratio)
        invalid["definition"]["ratio"]["numerator"]["dimension_bindings"]["period"]["key"] = "entity"
        with self.assertRaises(ValueError):
            self.api.validate_records(rows + [invalid])
        optional = copy.deepcopy(ratio)
        next(ref for ref in optional["dimension_refs"] if ref["id"] == "asset_selection")["required"] = False
        with self.assertRaises(ValueError):
            self.api.validate_records(rows + [optional])
        invalid_rows = copy.deepcopy(rows); invalid_rows[-1]["definition"]["value_type"] = "text"
        with self.assertRaises(ValueError):
            self.api.validate_records(invalid_rows + [ratio])
        second = copy.deepcopy(ratio); second["id"] = "另一比例"; second["definition"]["measure"] = "第二比例"
        ratio["definition"]["ratio"]["numerator"]["metric_id"] = second["id"]
        second["definition"]["ratio"]["numerator"]["metric_id"] = ratio["id"]
        with self.assertRaises(ValueError):
            self.api.validate_records(rows + [ratio, second])

    def test_ratio_identity_and_actual_constraints_include_referenced_definitions(self):
        rows = records()
        ratio = copy.deepcopy(rows[-1]); ratio["id"] = "占比"
        ratio["definition"].update(value_type="percentage", measure="比例", ratio={
            side: {"metric_id": "固定资产原值", "dimension_bindings": {ref["id"]: {"op": "dimension", "key": ref["id"]} for ref in ratio["dimension_refs"]}}
            for side in ("numerator", "denominator")})
        first = self.api.validate_records(rows + [ratio])
        amount = fact(); amount["metric_id"] = "占比"
        rows[-1]["definition"]["measurement_basis"] = "账面净额"
        second = self.api.validate_records(rows + [ratio])
        self.assertNotEqual(self.api.semantic_key(amount, first), self.api.semantic_key(amount, second))
        rows[-1]["definition"]["constraints"] = [{"dimension_id": "report_scope", "allowed_values": ["consolidated"]}]
        third = self.api.validate_records(rows + [ratio])
        with self.assertRaises(ValueError):
            self.api.validate_fact(amount, third)

    def test_binding_cannot_reference_missing_or_conflicting_owner(self):
        self.assertTrue(hasattr(self.api, "validate_facts"), "缺少载体内维度绑定和业务对象的一致性核验")
        binding = {"id": "名称格", "record_type": "dimension_binding", "dimension_id": "entity", "value": "合成企业甲",
                   "owner_reference": "对象一", "source_reference": fact()["source_reference"],
                   "recognition_provenance": fact()["recognition_provenance"]}
        amount = fact(); amount.update(id="金额格", owner_reference="对象一", binding_refs={"entity": "名称格"})
        owner = {"对象一": {"entity": "合成企业甲", "report_scope": "standalone"}}
        self.api.validate_facts([binding, amount], self.standard, owner)
        with self.assertRaises(ValueError):
            self.api.validate_facts([binding, amount], self.standard, {})
        owner["对象一"]["entity"] = "合成企业乙"
        with self.assertRaises(ValueError):
            self.api.validate_facts([binding, amount], self.standard, owner)


if __name__ == "__main__":
    unittest.main()
