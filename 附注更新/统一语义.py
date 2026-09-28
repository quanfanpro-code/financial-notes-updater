# coding: utf-8
"""统一标准的定义与实例校验；不推断业务等价，不执行未经复核的旧标准迁移。"""
from __future__ import annotations

import copy
import hashlib
import json
import re
from datetime import date
from decimal import Decimal, InvalidOperation
from functools import lru_cache
from pathlib import Path

from .金标准 import _invalid_number, _pairs

RESERVED_DIMENSIONS = {"entity": "text", "report_scope": "text", "currency": "text", "unit_scale": "number", "period": "period"}


def _json(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False)


def definition_hash(value):
    return hashlib.sha256(_json(value).encode("utf-8")).hexdigest()


def _require(condition, reason):
    if not condition:
        raise ValueError(reason)


def _text(value):
    return isinstance(value, str) and bool(value.strip())


def _fields(value, required, optional=()):
    _require(isinstance(value, dict), "定义或实例须为对象")
    missing = set(required) - set(value)
    extra = set(value) - set(required) - set(optional)
    _require(not missing and not extra, "缺少字段：" + str(sorted(missing)) + "；未定义字段：" + str(sorted(extra)))


def _texts(value, *, empty=False):
    _require(isinstance(value, list) and (empty or bool(value)) and all(_text(x) for x in value), "须为有效文字列表")
    _require(len(value) == len(set(value)), "列表不能含重复项")


def _hash(value):
    _require(isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value), "版本或来源须绑定完整SHA256")


def _source(value):
    _fields(value, {"file_hash", "location"}, {"scope", "legacy_id", "legacy_definition_hash", "source_gold_hash"})
    _hash(value["file_hash"])
    _require(_text(value["location"]), "来源必须有精确定位")
    for key in ("legacy_definition_hash", "source_gold_hash"):
        if key in value:
            _hash(value[key])


def _date(value):
    _require(isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value), "日期须为实际ISO日期")
    try:
        date.fromisoformat(value)
    except ValueError:
        raise ValueError("日期不存在") from None
    return value


def _number(value):
    _require(type(value) in {int, float, str}, "数字不能为布尔值或对象")
    try:
        number = Decimal(str(value))
    except InvalidOperation:
        raise ValueError("数值格式无效") from None
    _require(number.is_finite(), "数字须有限")
    return format(number.normalize(), "f")


def _dimension_value(value, definition, *, matching=False):
    kind = definition["value_type"]
    if kind == "text":
        _require(_text(value), "维度取值不能为空")
        result = value.strip()
        if definition["id"] == "report_scope":
            _require(result in {"standalone", "consolidated", "parent"}, "实际报告口径无效")
        if definition["id"] == "currency":
            _require(re.fullmatch(r"[A-Z]{3}", result), "币种须先按原资料解析为统一币种代码")
    elif kind == "number":
        result = _number(value)
        if definition["id"] == "unit_scale":
            _require(Decimal(result) > 0, "金额倍率须大于零")
    elif kind == "boolean":
        _require(type(value) is bool, "布尔维度须为布尔值")
        result = value
    elif kind == "date":
        result = _date(value)
    elif kind == "period":
        _require(isinstance(value, dict), "期间须为实际日期结构，不能只写期初/期末")
        if value.get("kind") == "instant":
            _fields(value, {"kind", "date"}); _date(value["date"])
        else:
            _fields(value, {"kind", "start", "end"})
            _require(value["kind"] == "duration", "期间类型无效")
            _require(_date(value["start"]) <= _date(value["end"]), "期间起止顺序错误")
        result = copy.deepcopy(value)
    else:
        _require(isinstance(value, dict), "集合维度须明确业务总体和选取方式")
        mode = value.get("mode")
        extras = {"all": set(), "members": {"members"}, "predicate": {"predicate"},
                  "complement": {"base", "excluded"}, "unknown": set()}
        _require(mode in extras, "集合选取方式无效")
        _fields(value, {"domain", "mode", "completeness"} | extras[mode])
        _require(value["domain"] == definition["domain"], "集合超出该维度声明的业务总体")
        _require(value["completeness"] in {"complete", "partial", "unknown"}, "集合完整性无效")
        if mode == "unknown":
            _require(value["completeness"] == "unknown", "未知集合不能声称完整")
        if matching:
            _require(mode != "unknown" and value["completeness"] == "complete", "集合范围未明，不能建立确定匹配")
        result = copy.deepcopy(value)
        for key in ("members", "excluded"):
            if key in value:
                _texts(value[key]); result[key] = sorted(value[key])
        for key in ("base", "predicate"):
            if key in value:
                _require(_text(value[key]), "集合总体及条件须明确")
    if "allowed_values" in definition:
        _require(result in definition["allowed_values"], "维度取值不在声明范围内")
    return result


def _rule(rule, definition=None):
    _require(isinstance(rule, dict), "转换规则须为对象")
    op = rule.get("op")
    _require(op in {"constant", "dimension", "context"}, "转换只支持审定常量、原维度或有依据上下文，不执行代码")
    _fields(rule, {"op", "value"} if op == "constant" else {"op", "key"})
    if op == "constant" and definition:
        _require(definition["value_type"] != "period", "实际期间须从原资料解析，不能以样例年份作转换常量")
        _dimension_value(rule["value"], definition)
    elif op != "constant":
        _require(_text(rule["key"]), "转换来源字段不能为空")


def validate_records(records):
    """验证完整v2记录集，返回按类型索引的独立副本；空指标目录可用于未发布审查稿。"""
    _require(isinstance(records, list) and bool(records), "标准不能为空")
    result = {kind: {} for kind in ("metric", "dimension", "source_example", "legacy_conversion")}
    identifiers = set()
    for original in records:
        row = copy.deepcopy(original)
        _require(isinstance(row, dict), "标准记录须为对象")
        kind, identifier = row.get("record_type"), row.get("id")
        _require(kind in result and _text(identifier), "标准记录类型或ID无效")
        _require(identifier not in identifiers, "标准存在重复ID：" + identifier)
        identifiers.add(identifier)
        result[kind][identifier] = row
    seen_dimensions=set()
    for row in result["dimension"].values():
        _fields(row, {"record_type", "id", "meaning", "value_type", "domain"}, {"allowed_values"})
        _require(_text(row["meaning"]) and _text(row["domain"]), "维度须声明含义和业务域")
        _require(row["value_type"] in {"text", "number", "boolean", "date", "period", "selection"}, "维度类型无效")
        if row["id"] in RESERVED_DIMENSIONS:
            _require(row["value_type"] == RESERVED_DIMENSIONS[row["id"]], "报告主体、口径、币种、倍率及期间不能改变基本类型")
        if "allowed_values" in row:
            _require(isinstance(row["allowed_values"], list) and bool(row["allowed_values"]), "限定值清单不能为空")
            unrestricted = {key: value for key, value in row.items() if key != "allowed_values"}
            normalized = [_dimension_value(value, unrestricted) for value in row["allowed_values"]]
            _require(len(set(map(_json, normalized))) == len(normalized), "维度限定值重复")
            row["allowed_values"] = sorted(normalized,key=_json)
        signature=_json({key:value for key,value in row.items() if key!='id'})
        _require(signature not in seen_dimensions,'相同维度定义不能换编号重复新增')
        seen_dimensions.add(signature)
    for row in result["source_example"].values():
        _fields(row, {"record_type", "id", "origin", "text_evidence", "review_refs"})
        _source(row["origin"]); _texts(row["text_evidence"]); _texts(row["review_refs"])
    seen_definitions = {}
    for row in result["metric"].values():
        _fields(row, {"record_type", "id", "status", "definition", "dimension_refs", "source_refs"})
        _require(row["status"] in {"active", "retired"}, "统一指标状态无效")
        definition = row["definition"]
        _fields(definition, {"economic_object", "measure", "measurement_basis", "value_type", "time_kind", "constraints"}, {"ratio"})
        for key in ("economic_object", "measure", "measurement_basis"):
            _require(_text(definition[key]), "指标须明确经济对象、计量内容及基础")
        _require(definition["value_type"] in {"monetary", "number", "percentage", "text", "date", "boolean", "enum"}, "指标值类型无效")
        _require(definition["time_kind"] in {"instant", "duration", "none"}, "指标期间性质无效")
        _require(isinstance(definition["constraints"], list), "业务限制须为可验证的维度条件清单")
        refs = row["dimension_refs"]
        _require(isinstance(refs, list), "指标须明确维度引用清单")
        ref_ids = set()
        for ref in refs:
            _fields(ref, {"id", "required"})
            _require(ref["id"] in result["dimension"] and ref["id"] not in ref_ids, "指标维度引用不存在或重复")
            _require(type(ref["required"]) is bool, "required须为布尔值")
            ref_ids.add(ref["id"])
        required = {ref["id"] for ref in refs if ref["required"]}
        _require({"entity", "report_scope"} <= required, "业务事实须明确实际报告主体及口径，来源口径不能替代")
        constraint_ids = set()
        for constraint in definition["constraints"]:
            _fields(constraint, {"dimension_id", "allowed_values"})
            key = constraint["dimension_id"]
            _require(key in required and key not in constraint_ids, "业务限制须引用不重复的必填维度")
            constraint_ids.add(key)
            _require(isinstance(constraint["allowed_values"], list) and bool(constraint["allowed_values"]), "业务允许范围不能为空")
            values = [_dimension_value(value, result["dimension"][key]) for value in constraint["allowed_values"]]
            _require(len(set(map(_json, values))) == len(values), "业务限制值重复")
            constraint["allowed_values"] = sorted(values, key=_json)
        definition["constraints"].sort(key=lambda item: item["dimension_id"])
        if definition["time_kind"] != "none":
            _require("period" in required and result["dimension"]["period"]["value_type"] == "period", "时点或期间指标须声明实际期间维度")
        if definition["value_type"] == "monetary":
            _require({"currency", "unit_scale"} <= required, "金额指标须声明币种及倍率")
        _texts(row["source_refs"])
        _require(all(ref in result["source_example"] for ref in row["source_refs"]), "指标来源引用不存在")
        if definition["value_type"] == "percentage":
            _require("ratio" in definition, "比例必须明确业务分子和分母")
        if "ratio" in definition:
            _fields(definition["ratio"], {"numerator", "denominator"})
            for side in definition["ratio"].values():
                _fields(side, {"metric_id", "dimension_bindings"})
                _require(side["metric_id"] in result["metric"] and side["metric_id"] != row["id"], "比例引用指标不存在或自引用")
                target = result["metric"][side["metric_id"]]
                _require(target["status"] == "active" and target["definition"]["value_type"] in {"monetary", "number", "percentage"}, "比例分子分母须为有效数值指标")
                declared = {ref["id"] for ref in target["dimension_refs"]}
                mandatory = {ref["id"] for ref in target["dimension_refs"] if ref["required"]}
                _require(isinstance(side["dimension_bindings"], dict) and mandatory <= set(side["dimension_bindings"]) <= declared, "比例两侧须完整绑定业务维度")
                for key, rule in side["dimension_bindings"].items():
                    _rule(rule, result["dimension"][key])
                    _require(rule["op"] != "context", "比例定义不得依赖未说明的实例上下文")
                    if rule["op"] == "dimension":
                        _require(rule["key"] in required, "比例两侧使用的实例维度须在比例指标中声明为必需")
                        origin_dimension, target_dimension = result["dimension"][rule["key"]], result["dimension"][key]
                        _require(origin_dimension["value_type"] == target_dimension["value_type"] and origin_dimension["domain"] == target_dimension["domain"], "比例维度绑定的类型或业务域不一致")
        if row["status"] == "active":
            key = _json(definition)
            _require(key not in seen_definitions, "同一指标定义不能因为编号、样式或维度引用不同重复建号")
            seen_definitions[key] = row["id"]
    # 比例含义必须最终落到可计量指标，循环引用不能构成有效定义。
    completed = set()
    def visit(identifier, visiting):
        _require(identifier not in visiting, "比例指标存在循环引用")
        if identifier in completed:
            return
        for side in result["metric"][identifier]["definition"].get("ratio", {}).values():
            visit(side["metric_id"], visiting | {identifier})
        completed.add(identifier)
    for identifier in result["metric"]:
        visit(identifier, set())
    legacy_keys = set()
    for row in result["legacy_conversion"].values():
        _fields(row, {"record_type", "id", "source_gold_hash", "legacy_id", "legacy_definition_hash", "status", "target", "evidence_refs", "review"})
        _hash(row["source_gold_hash"]); _hash(row["legacy_definition_hash"])
        _require(_text(row["legacy_id"]), "转换缺少旧编号")
        key = row["source_gold_hash"], row["legacy_id"]
        _require(key not in legacy_keys, "一个旧版本旧编号不能有多个未消歧转换")
        legacy_keys.add(key)
        _require(row["status"] in {"approved", "unresolved"}, "转换状态无效")
        _fields(row["review"], {"reviewer", "reason"})
        _require(_text(row["review"]["reason"]), "转换须说明依据或待审原因")
        _texts(row["evidence_refs"])
        _require(all(ref in result["source_example"] for ref in row["evidence_refs"]), "转换依据引用不存在")
        if row["status"] == "unresolved":
            _require(row["target"] is None, "未决转换不能提供可执行目标")
            continue
        _require(_text(row["review"]["reviewer"]), "已批准转换须保留复核身份")
        target = row["target"]
        _require(isinstance(target, dict), "转换缺少确定目标")
        if target.get("record_type") == "metric_fact":
            _fields(target, {"record_type", "metric_id", "dimension_bindings"})
            _require(target["metric_id"] in result["metric"], "转换指标不存在")
            metric = result["metric"][target["metric_id"]]
            _require(metric["status"] == "active", "转换不能指向停用指标")
            declared = {ref["id"] for ref in metric["dimension_refs"]}
            required = {ref["id"] for ref in metric["dimension_refs"] if ref["required"]}
            bindings = target["dimension_bindings"]
            _require(isinstance(bindings, dict) and required <= set(bindings) <= declared, "转换须逐维度声明取值依据，不能丢弃或臆补")
            for key, rule in bindings.items():
                _rule(rule, result["dimension"][key])
        else:
            _fields(target, {"record_type", "dimension_id", "value_rule", "owner_rule"})
            _require(target["record_type"] == "dimension_binding" and target["dimension_id"] in result["dimension"], "维度绑定转换目标无效")
            _rule(target["value_rule"], result["dimension"][target["dimension_id"]])
            _rule(target["owner_rule"])
            _require(target["owner_rule"]["op"] == "context", "维度归属须从当前载体上下文确定")
    return result


def read_standard_v2(path):
    # 每次读取实际内容；返回独立副本，避免调用者修改已校验的定义。
    return copy.deepcopy(_validated_standard_content(Path(path).read_bytes()))


@lru_cache(maxsize=2)
def _validated_standard_content(raw):
    rows = [json.loads(line, object_pairs_hook=_pairs, parse_constant=_invalid_number)
            for line in raw.decode("utf-8-sig").splitlines() if line.strip()]
    result = validate_records(rows)
    result["sha256"] = hashlib.sha256(raw).hexdigest()
    return result


def validate_legacy_bindings(standard, legacy_rows, legacy_gold_hash):
    """校验被迁移版本及完整定义；不改写旧意见，也不替用户作语义等价判断。"""
    _hash(legacy_gold_hash)
    for row in standard["legacy_conversion"].values():
        _require(row["source_gold_hash"] == legacy_gold_hash, "旧标准版本不匹配")
        original = legacy_rows.get(row["legacy_id"])
        _require(original is not None and definition_hash(original) == row["legacy_definition_hash"], "旧完整定义已变化，不能沿用转换")
    return True


def partial_pending_dimensions(known, standard):
    return sorted(set(known['missing_dimensions'])|{key for key,value in known['dimensions'].items()
        if standard['dimension'][key]['value_type']=='selection' and (value['mode']=='unknown' or value['completeness']!='complete')})


def validate_partial_semantics(known, standard):
    """已知含义独立保存；缺少实际维度时不能冒充可匹配的完整事实。"""
    _fields(known, {'metric_id','dimensions','dimension_evidence','missing_dimensions','evidence_cells'})
    _require(known['metric_id'] in standard['metric'],'部分语义须引用现有指标')
    metric=standard['metric'][known['metric_id']]
    _require(metric['status']=='active','不能使用停用指标')
    declared={ref['id'] for ref in metric['dimension_refs']}
    required={ref['id'] for ref in metric['dimension_refs'] if ref['required']}
    dims=known['dimensions'];evidence=known['dimension_evidence']
    _require(isinstance(dims,dict) and set(dims)<=declared,'部分语义包含未声明维度')
    _require(isinstance(evidence,dict) and set(evidence)==set(dims),'每项已知维度须有证据')
    _texts(known['missing_dimensions'],empty=True);_texts(known['evidence_cells'])
    _require(set(known['missing_dimensions'])==required-set(dims),'缺少维度须精确列出尚未确定的必需项')
    for key,value in dims.items():
        _dimension_value(value,standard['dimension'][key]);_texts(evidence[key])
    _require(partial_pending_dimensions(known,standard),'全部维度和范围已明确时应保存完整事实')
    definition=metric['definition']
    if 'period' in dims and definition['time_kind']!='none':
        _require(dims['period']['kind']==definition['time_kind'],'已知期间与指标时点/期间性质不符')
    for constraint in definition['constraints']:
        key=constraint['dimension_id']
        if key in dims:_require(_dimension_value(dims[key],standard['dimension'][key]) in constraint['allowed_values'],'已知维度不满足指标适用限制')
    return True


def validate_fact(fact, standard):
    """允许保存未知集合，但不允许将其作为确定匹配；实例证据不进入业务身份。"""
    _require(isinstance(fact, dict), "事实须为对象")
    common = {"record_type", "source_reference", "recognition_provenance"}
    _source(fact.get("source_reference"))
    provenance = fact.get("recognition_provenance")
    _require(isinstance(provenance, dict) and _text(provenance.get("method")) and _text(provenance.get("record_id")), "事实须保留识别或复核来源")
    if fact.get("record_type") == "dimension_binding":
        _fields(fact, common | {"dimension_id", "value", "owner_reference"}, {"id"})
        _require(fact["dimension_id"] in standard["dimension"] and _text(fact["owner_reference"]), "维度绑定须有已声明维度和所属对象")
        _dimension_value(fact["value"], standard["dimension"][fact["dimension_id"]])
        return True
    _fields(fact, common | {"metric_id", "dimensions", "dimension_evidence", "raw_value_state", "raw_value"}, {"id", "owner_reference", "binding_refs"})
    _require(fact["record_type"] == "metric_fact" and fact["metric_id"] in standard["metric"], "事实须引用统一指标")
    metric = standard["metric"][fact["metric_id"]]
    _require(metric["status"] == "active", "不能使用停用指标")
    dims, evidence = fact["dimensions"], fact["dimension_evidence"]
    declared = {ref["id"] for ref in metric["dimension_refs"]}
    required = {ref["id"] for ref in metric["dimension_refs"] if ref["required"]}
    _require(isinstance(dims, dict) and required <= set(dims) <= declared, "缺少必需维度或使用了未声明维度")
    _require(isinstance(evidence, dict) and set(evidence) == set(dims), "每项实际维度必须有证据")
    for key, value in dims.items():
        _dimension_value(value, standard["dimension"][key]); _texts(evidence[key])
    definition = metric["definition"]
    for constraint in definition["constraints"]:
        key = constraint["dimension_id"]
        _require(_dimension_value(dims[key], standard["dimension"][key]) in constraint["allowed_values"], "实际维度不满足指标的业务适用限制")
    if definition["time_kind"] != "none":
        _require(dims["period"]["kind"] == definition["time_kind"], "指标时点/期间性质与实际期间不一致")
    for side in definition.get("ratio", {}).values():
        projected = {key: copy.deepcopy(rule["value"] if rule["op"] == "constant" else dims[rule["key"]])
                     for key, rule in side["dimension_bindings"].items()}
        # 只检查分子分母的实际维度适用性，不伪造它们的金额。
        validate_fact({"record_type": "metric_fact", "metric_id": side["metric_id"], "dimensions": projected,
                       "dimension_evidence": {key: ["比例定义的已验证维度绑定"] for key in projected},
                       "source_reference": fact["source_reference"], "recognition_provenance": provenance,
                       "raw_value_state": "unknown", "raw_value": None}, standard)
    state, value = fact["raw_value_state"], fact["raw_value"]
    _require(state in {"value", "blank", "unknown", "formula"}, "原值状态无效")
    if state in {"blank", "unknown"}:
        _require(value is None or isinstance(value, str) and not value.strip(), "原空白或未知不能擅自填零")
    elif state == "formula":
        _require(isinstance(value, str) and value.startswith("="), "公式状态须保留原公式")
    else:
        # 金标准的值类型描述计量含义，不据此否定已识别的语义或修正人工原值。
        # 这里只校验原值容器及状态；源文件一致性仍由映射加载器逐格核对。
        _require(type(value) in {str, int, float, bool}, "原值须为单个单元格的原始值")
        if isinstance(value, str):
            _require(_text(value), "空白原值须使用空白状态")
            _require(not value.startswith("="), "公式原值须使用公式状态")
        elif type(value) in {int, float}:
            _number(value)
    return True


def validate_facts(facts, standard, owners):
    """整份载体核验：绑定必须回指存在的业务对象，引用的维度值和归属必须一致。"""
    _require(isinstance(facts, list) and isinstance(owners, dict), "载体事实和业务对象清单无效")
    normalized_owners = {}
    for identifier, dimensions in owners.items():
        _require(_text(identifier) and isinstance(dimensions, dict) and {"entity", "report_scope"} <= set(dimensions), "业务对象须有主体及口径")
        _require(set(dimensions) <= set(standard["dimension"]), "业务对象含未声明维度")
        normalized_owners[identifier] = {key: _dimension_value(value, standard["dimension"][key]) for key, value in dimensions.items()}
    indexed = {}
    for fact in facts:
        validate_fact(fact, standard)
        identifier, owner = fact.get("id"), fact.get("owner_reference")
        _require(_text(identifier) and identifier not in indexed, "载体事实须有不重复的本地ID")
        _require(_text(owner) and owner in normalized_owners, "事实或维度绑定引用了不存在的所属对象")
        indexed[identifier] = fact
        dimensions = fact["dimensions"] if fact["record_type"] == "metric_fact" else {fact["dimension_id"]: fact["value"]}
        if fact["record_type"] == "metric_fact":
            _require(set(normalized_owners[owner]) <= set(dimensions), "事实缺少所属业务对象的身份维度")
        for key, expected in normalized_owners[owner].items():
            if key in dimensions:
                _require(_dimension_value(dimensions[key], standard["dimension"][key]) == expected, "单元格维度与所属业务对象冲突")
    for fact in facts:
        if fact["record_type"] != "metric_fact":
            continue
        bindings = fact.get("binding_refs", {})
        _require(isinstance(bindings, dict) and set(bindings) <= set(fact["dimensions"]), "事实绑定引用了未使用的维度")
        for key, identifier in bindings.items():
            binding = indexed.get(identifier)
            _require(binding is not None and binding["record_type"] == "dimension_binding" and binding["dimension_id"] == key, "维度证据须引用存在的相应维度单元格")
            _require(binding["owner_reference"] == fact["owner_reference"], "维度单元格不属于该业务对象")
            definition = standard["dimension"][key]
            _require(_dimension_value(binding["value"], definition) == _dimension_value(fact["dimensions"][key], definition), "维度绑定值与事实取值不一致")
    return True


def _metric_meaning(identifier, standard):
    row=standard['metric'][identifier]
    return {'definition':row['definition'],
            'dimensions':[{'required':ref['required'],'definition':standard['dimension'][ref['id']]}
                          for ref in sorted(row['dimension_refs'],key=lambda ref:ref['id'])],
            'ratio_meanings':{side:_metric_meaning(target['metric_id'],standard)
                              for side,target in row['definition'].get('ratio',{}).items()}}


def _dimension_compatible(old, new):
    """仅扩充允许取值不改变旧取值含义；定义其余部分仍须完全相同。"""
    if old == new:
        return True
    if not new or 'allowed_values' not in old or 'allowed_values' not in new:
        return False
    return ({key: value for key, value in old.items() if key != 'allowed_values'} ==
            {key: value for key, value in new.items() if key != 'allowed_values'} and
            {_json(value) for value in old['allowed_values']} <= {_json(value) for value in new['allowed_values']})


def _metric_compatible(identifier, old, new):
    previous = old['metric'][identifier]
    current = new['metric'].get(identifier)
    if not current or current['status'] != 'active' or previous['definition'] != current['definition']:
        return False
    refs = sorted(previous['dimension_refs'], key=lambda ref: ref['id'])
    return (refs == sorted(current['dimension_refs'], key=lambda ref: ref['id']) and
            all(_dimension_compatible(old['dimension'][ref['id']], new['dimension'].get(ref['id'])) for ref in refs) and
            all(_metric_compatible(target['metric_id'], old, new)
                for target in previous['definition'].get('ratio', {}).values()))


def compatible_classification(mapping, standard):
    """从已按旧版完整回读的成果选取未改义的分类；不迁移已改变的语义。"""
    old = read_standard_v2(mapping['gold_path'])
    _require(old['sha256'] == mapping['gold_hash'], '复用依据的旧金标准已变化')
    validate_facts(mapping['facts'], old, mapping['owners'])
    unchanged_metrics = set()
    for identifier in old['metric'].keys() & standard['metric'].keys():
        if _metric_compatible(identifier, old, standard):
            unchanged_metrics.add(identifier)
    unchanged_dimensions = {k for k, v in old['dimension'].items() if _dimension_compatible(v, standard['dimension'].get(k))}
    unchanged_owners = {k for k, v in mapping['owners'].items() if set(v) <= unchanged_dimensions}
    conversions_unchanged = all(standard['legacy_conversion'].get(k) == v for k, v in old['legacy_conversion'].items())
    facts = []
    for fact in mapping['facts']:
        if fact['owner_reference'] not in unchanged_owners:
            continue
        if fact['recognition_provenance']['method'] == 'legacy_conversion' and not conversions_unchanged:
            continue
        if ((fact['record_type'] == 'metric_fact' and fact['metric_id'] in unchanged_metrics)
                or (fact['record_type'] == 'dimension_binding' and fact['dimension_id'] in unchanged_dimensions)):
            facts.append(copy.deepcopy(fact))
    kept_ids = {fact['id'] for fact in facts}
    facts = [fact for fact in facts if set(fact.get('binding_refs', {}).values()) <= kept_ids]
    owners = {fact['owner_reference']: copy.deepcopy(mapping['owners'][fact['owner_reference']]) for fact in facts}
    validate_facts(facts, standard, owners)
    return {'facts': facts, 'owners': owners, 'excluded': copy.deepcopy(mapping.get('excluded', []))}


def ensure_standard_compatible(mapping, standard):
    """新增定义或扩充维度可选值可复用；已有含义、必需维度及指标限制必须保持。"""
    _require(mapping.get('schema_version')==2,'映射不是统一标准格式')
    validate_facts(mapping['facts'],standard,mapping['owners'])
    partials=[row['partial_semantics'] for row in mapping.get('unresolved',[]) if row.get('partial_semantics')]
    for known in partials:validate_partial_semantics(known,standard)
    if mapping.get('gold_hash')==standard['sha256']:return True
    _require(mapping.get('gold_path'),'映射缺少原金标准版本，不能核实兼容性')
    old=read_standard_v2(mapping['gold_path'])
    _require(old['sha256']==mapping['gold_hash'],'映射原金标准文件已变化')
    validate_facts(mapping['facts'],old,mapping['owners'])
    for known in partials:
        validate_partial_semantics(known,old)
        _require(_metric_compatible(known['metric_id'],old,standard),
                 '部分语义所用定义已改变，须重新识别：'+known['metric_id'])
    for fact in mapping['facts']:
        if fact['record_type']=='metric_fact':
            identifier=fact['metric_id']
            _require(_metric_compatible(identifier,old,standard),
                     '所用指标或维度定义已改变，须重新识别：'+identifier)
        else:
            identifier=fact['dimension_id']
            _require(_dimension_compatible(old['dimension'][identifier],standard['dimension'][identifier]),
                     '所用维度定义已改变，须重新识别：'+identifier)
        if fact['recognition_provenance']['method']=='legacy_conversion':
            _require(all(standard['legacy_conversion'].get(key)==value for key,value in old['legacy_conversion'].items()),
                     '旧标准转换意见已变化，须先复核已迁移事实')
    return True


def semantic_key(fact, standard):
    """指标相同不等于事实相同；完整实例维度一致才得到相同键。"""
    validate_fact(fact, standard)
    _require(fact["record_type"] == "metric_fact", "维度单元格只建立归属绑定，不能作为金额事实匹配")
    dimensions = {key: _dimension_value(value, standard["dimension"][key], matching=True)
                  for key, value in fact["dimensions"].items() if key != 'entity'}
    metric = standard["metric"][fact["metric_id"]]
    if metric["definition"]["value_type"] == "monetary":
        # 主体名称和金额倍率不决定财务单元格的对应；值始终由用户决定。
        del dimensions["unit_scale"]
    signature = _metric_meaning(fact["metric_id"],standard)
    return fact["metric_id"], definition_hash(signature), _json(dimensions)


def value_for_target(source, target, standard):
    """只搬运已确定同一业务事实的明确值；不将空白或未知变成零。"""
    _require(semantic_key(source, standard) == semantic_key(target, standard), "源与目标不是同一业务事实")
    _require(source["raw_value_state"] == "value", "源值未明确，不能写入目标")
    return copy.deepcopy(source["raw_value"])


def prepare_legacy_review(legacy_rows, source_gold_hash):
    """全库建立待审转换，包含历史定义；建账本身不批准任何业务等价或迁移。"""
    _hash(source_gold_hash)
    _require(isinstance(legacy_rows, dict) and bool(legacy_rows), "缺少完整旧标准")
    records = []
    for identifier, original in legacy_rows.items():
        _require(original.get("id") == identifier, "旧标准索引与定义编号不一致")
        definition_sha = definition_hash(original)
        source_id = "source:" + identifier
        origin = {"file_hash": source_gold_hash, "location": "旧定义ID:" + identifier,
                  "legacy_id": identifier, "source_gold_hash": source_gold_hash,
                  "legacy_definition_hash": definition_sha}
        if original.get("scope"):
            origin["scope"] = original["scope"]
        description = {"原角色": original.get("slot", {}).get("role"), "原值类型": original.get("value_type"),
                       "原状态": original.get("status"), "历史后继": original.get("superseded_by", []),
                       "原定义": original}
        records.append({"record_type": "source_example", "id": source_id, "origin": origin,
                        "text_evidence": [_json(description)], "review_refs": ["全量只读建账，业务等价尚未复核"]})
        records.append({"record_type": "legacy_conversion", "id": "conversion:" + identifier,
                        "source_gold_hash": source_gold_hash, "legacy_id": identifier,
                        "legacy_definition_hash": definition_sha, "status": "unresolved", "target": None,
                        "evidence_refs": [source_id], "review": {"reviewer": "", "reason":
                            "需保留历史关系并复核去向，不作为新识别候选" if original.get("status") != "active" else
                            "需复核指标、维度及实际业务限制；旧角色或相似名称不能直接决定转换"}})
    standard = validate_records(records)
    validate_legacy_bindings(standard, legacy_rows, source_gold_hash)
    return records
