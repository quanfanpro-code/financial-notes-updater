# coding: utf-8
"""执行已复核的单值旧定义转换；原成果和复核记录保持不变。"""
from __future__ import annotations

import copy
from .统一语义 import (_dimension_value, _fields, _hash, _require, _texts,
                     definition_hash, validate_fact)


def migrate_fact(original, legacy_definition, source_gold_hash, standard, context):
    """输入由文件读取入口验证的旧事实及上下文，输出新事实和独立迁移凭据。

    文件级调用者负责核验原映射哈希、来源坐标和上下文原文；此处核验转换版本、
    定义、维度冲突及值类型。不得以此函数替代原复核链或原件回读。
    """
    _hash(source_gold_hash); _hash(standard.get("sha256"))
    _fields(original, {"legacy_id", "legacy_mapping_hash", "dimensions", "dimension_evidence",
                       "source_reference", "raw_value_state", "raw_value", "recognition_provenance"})
    _hash(original["legacy_mapping_hash"])
    _require(legacy_definition.get("id") == original["legacy_id"], "原映射与旧定义编号不一致")
    candidates = [row for row in standard["legacy_conversion"].values()
                  if row["source_gold_hash"] == source_gold_hash and row["legacy_id"] == original["legacy_id"]]
    _require(len(candidates) == 1, "缺少唯一的已绑定版本转换")
    conversion = candidates[0]
    _require(conversion["status"] == "approved" and conversion["target"] is not None, "该旧定义仍待复核，不能迁移")
    _require(conversion["legacy_definition_hash"] == definition_hash(legacy_definition), "旧完整定义已变化")
    dimensions, proofs = original["dimensions"], original["dimension_evidence"]
    _require(isinstance(dimensions, dict) and isinstance(proofs, dict) and set(dimensions) == set(proofs), "原实际维度必须有逐项证据")
    _require(isinstance(context, dict), "转换上下文须为对象")
    consumed, used_context = set(), {}

    def resolve(rule):
        op = rule["op"]
        if op == "constant":
            return copy.deepcopy(rule["value"]), list(conversion["evidence_refs"])
        key = rule["key"]
        if op == "dimension":
            _require(key in dimensions, "原映射缺少实际维度：" + key)
            _texts(proofs[key]); consumed.add(key)
            return copy.deepcopy(dimensions[key]), list(proofs[key])
        _require(op == "context" and key in context, "缺少有依据的实际上下文：" + key)
        supplied = context[key]
        _fields(supplied, {"value", "evidence_refs"}); _texts(supplied["evidence_refs"])
        used_context[key] = copy.deepcopy(supplied)
        return copy.deepcopy(supplied["value"]), list(supplied["evidence_refs"])

    target = conversion["target"]
    _require(target["record_type"] == "metric_fact", "维度单元格应连同业务对象迁移，不通过金额事实入口强转")
    result_dimensions, evidence = {}, {}
    for key, rule in target["dimension_bindings"].items():
        value, references = resolve(rule)
        canonical = _dimension_value(value, standard["dimension"][key])
        if key in dimensions:
            _require(_dimension_value(dimensions[key], standard["dimension"][key]) == canonical,
                     "转换不能覆盖已经确认的实际维度：" + key)
            _texts(proofs[key]); consumed.add(key)
        result_dimensions[key], evidence[key] = canonical, references
    _require(consumed == set(dimensions), "转换不得静默丢弃原业务维度：" + "、".join(sorted(set(dimensions) - consumed)))
    receipt = {"source_mapping_hash": original["legacy_mapping_hash"], "source_gold_hash": source_gold_hash,
               "legacy_id": original["legacy_id"], "legacy_definition_hash": definition_hash(legacy_definition),
               "source_fact_hash": definition_hash(original), "source_provenance": copy.deepcopy(original["recognition_provenance"]),
               "target_gold_hash": standard["sha256"], "conversion_hash": definition_hash(conversion),
               "conversion_id": conversion["id"], "consumed_dimensions": sorted(consumed),
               "context_dependencies": used_context}
    fact = {"record_type": "metric_fact", "metric_id": target["metric_id"], "dimensions": result_dimensions,
            "dimension_evidence": evidence, "source_reference": copy.deepcopy(original["source_reference"]),
            "raw_value_state": original["raw_value_state"], "raw_value": copy.deepcopy(original["raw_value"]),
            "recognition_provenance": {"method": "legacy_conversion", "record_id": definition_hash(receipt)}}
    validate_fact(fact, standard)
    receipt["target_fact_hash"] = definition_hash(fact)
    return {"fact": fact, "receipt": receipt}
