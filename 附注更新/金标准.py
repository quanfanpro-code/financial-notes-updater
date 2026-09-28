"""金标准完整修订的本地版本维护；结构校验不代替业务语义审核。"""
from __future__ import annotations

import copy
import hashlib
import json
import re
import uuid
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VALUE_TYPES = {"monetary", "number", "percentage", "text", "date", "boolean", "enum"}
STATUSES = {"active", "superseded", "deprecated"}


def _pairs(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("JSON对象含重复字段：" + key)
        result[key] = value
    return result


def _invalid_number(value):
    raise ValueError("JSON不能包含非有限数字：" + value)


def _text(value):
    return isinstance(value, str) and bool(value.strip())


def _references(value, prefix=""):
    if isinstance(value, dict):
        for key, item in value.items():
            location = prefix + "." + key
            if key.endswith("_slot_id"):
                if item is not None and not _text(item):
                    raise ValueError("单个槽位引用必须是ID字符串：" + location)
                if item:
                    yield item, location
            elif key.endswith("_slot_ids") or key == "superseded_by":
                if item is not None:
                    if not isinstance(item, list) or any(not _text(entry) for entry in item):
                        raise ValueError("槽位引用清单必须是ID列表：" + location)
                    for entry in item:
                        yield entry, location
            yield from _references(item, location)
    elif isinstance(value, list):
        for item in value:
            yield from _references(item, prefix + "[]")


def allowed_scopes(row):
    """返回该原定义及已逐项审核扩展允许的实际报表口径，不自动推定可复用。"""
    original = row.get("scope")
    if original not in {"consolidated", "parent"}:
        raise ValueError("金标准原定义口径无效")
    scopes = {original}
    applicability = row.get("applicability", {})
    if not isinstance(applicability, dict):
        raise ValueError("applicability须保留原有适用行业对象")
    entries = applicability.get("report_scopes", [])
    if not isinstance(entries, list):
        raise ValueError("applicability.report_scopes适用口径扩展必须为列表")
    for entry in entries:
        fields = {"scope", "reason", "reviewed_by", "reviewed_at"}
        if not isinstance(entry, dict) or set(entry) != fields:
            raise ValueError("applicability适用性扩展须完整包含scope、reason、reviewed_by、reviewed_at")
        scope = entry["scope"]
        if scope != "standalone":
            raise ValueError("applicability当前仅支持审核后的standalone适用性，不能混入其他口径")
        if scope in scopes:
            raise ValueError("applicability适用性口径不能重复")
        if any(not _text(entry[field]) for field in ("reason", "reviewed_by", "reviewed_at")):
            raise ValueError("applicability适用性依据、审核人和审核日期不能为空")
        stamp = entry["reviewed_at"]
        try:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:T.+)?", stamp):
                raise ValueError()
            datetime.fromisoformat(stamp)
        except ValueError:
            raise ValueError("applicability审核日期须为有效ISO日期或日期时间") from None
        scopes.add(scope)
    return scopes


def validate_mapping_scope(row, mapping):
    """实际口径须经原定义允许；跨口径映射必须明确保留原定义口径。"""
    if not isinstance(mapping, dict):
        raise ValueError("映射缺少报表口径")
    actual = mapping.get("scope")
    if not isinstance(actual, str) or actual not in allowed_scopes(row):
        raise ValueError("映射实际报表口径不在该金标准定义的适用范围")
    original = row["scope"]
    definition = mapping.get("definition_scope", original if actual == original else None)
    if definition != original:
        raise ValueError("映射definition_scope须保留金标准原定义口径")
    return True


def normalize_mapping_metric(row, mapping):
    """依据绑定的金标准移除旧记录中重复的指标名称，保留真正声明的维度。"""
    result = copy.deepcopy(mapping)
    dimensions = result.get("dimensions")
    if dimensions is None and "dimensions" not in result:
        return result
    if not isinstance(dimensions, dict):
        raise ValueError("映射维度必须为对象")
    if any(item.get("name") == "metric" for item in row.get("dimensions") or []):
        return result
    if "metric" not in dimensions:
        return result
    default = str((row.get("slot") or {}).get("name") or "").strip()
    aliases = [default, " / ".join(row.get("column_path") or [])]
    aliases.extend((row.get("aliases") or {}).get("column_labels") or [])
    supplied = dimensions["metric"]
    if not isinstance(supplied, str) or not supplied.strip() or supplied not in aliases:
        raise ValueError("该金标准未声明metric维度，不能用其他指标替代原定义；请核实或补充金标准")
    del dimensions["metric"]
    return result


def _validate_active(row):
    identifier = row["id"]
    if row.get("scope") not in {"consolidated", "parent"}:
        raise ValueError(identifier + " 的scope须为consolidated或parent")
    allowed_scopes(row)
    for field in ("note", "table", "slot"):
        obj = row.get(field)
        if not isinstance(obj, dict) or not _text(obj.get("name")):
            raise ValueError(identifier + " 缺少有效的" + field + "定义名称")
        if field != "slot" and not _text(obj.get("id")):
            raise ValueError(identifier + " 缺少" + field + "定义ID")
    prefix = "C" if row["scope"] == "consolidated" else "P"
    note, table = row["note"]["id"], row["table"]["id"]
    if not re.fullmatch(prefix + r"-N\d+", note) or not re.fullmatch(re.escape(note) + r"-T\d+", table) or not re.fullmatch(re.escape(table) + r"-S\d+", identifier):
        raise ValueError(identifier + " 的报表口径、note/table/slot层级ID不一致")
    if row.get("value_type") not in VALUE_TYPES:
        raise ValueError(identifier + " 的value_type不受当前识别程序支持")
    dimensions = row.get("dimensions")
    if not isinstance(dimensions, list):
        raise ValueError(identifier + " 的dimensions必须为维度定义列表")
    names = set()
    for item in dimensions:
        if not isinstance(item, dict) or not _text(item.get("name")) or type(item.get("required")) is not bool:
            raise ValueError(identifier + " 的维度须包含name和布尔required")
        if item["name"] in names:
            raise ValueError(identifier + " 存在重复维度：" + item["name"])
        names.add(item["name"])
    for field in ("row_path", "column_path"):
        if field in row and (not isinstance(row[field], list) or any(not isinstance(item, str) for item in row[field])):
            raise ValueError(identifier + " 的" + field + "必须是文字列表")
    if row.get("superseded_by"):
        raise ValueError(identifier + " 为active，不能同时声明已被替代")


def _load(path):
    path = Path(path).resolve()
    if not path.is_file():
        raise ValueError("金标准文件不存在：" + str(path))
    raw = path.read_bytes()
    try:
        content = raw.decode("utf-8-sig")
    except UnicodeError as error:
        raise ValueError("金标准须为UTF-8编码") from error
    rows = {}
    for number, line in enumerate(content.splitlines(), 1):
        if not line.strip():
            continue
        try:
            row = json.loads(line, object_pairs_hook=_pairs, parse_constant=_invalid_number)
        except ValueError as error:
            raise ValueError(f"金标准第{number}行不是有效JSON：{error}") from error
        if not isinstance(row, dict) or not _text(row.get("id")):
            raise ValueError(f"金标准第{number}行缺少ID")
        identifier = row["id"]
        if identifier in rows:
            raise ValueError("金标准存在重复ID：" + identifier)
        if row.get("status") not in STATUSES:
            raise ValueError(identifier + " 的status须为active/superseded/deprecated")
        if row["status"] == "active":
            _validate_active(row)
        elif row["status"] == "superseded" and not row.get("superseded_by"):
            raise ValueError(identifier + " 标记替代后须列明superseded_by引用")
        rows[identifier] = row
    if not rows or not any(row["status"] == "active" for row in rows.values()):
        raise ValueError("金标准没有可用的active定义")
    for identifier, row in rows.items():
        for target, location in _references(row):
            if target not in rows:
                raise ValueError(identifier + " 的引用不存在：" + target + "（" + location + "）")
            if location == ".superseded_by" and target == identifier:
                raise ValueError(identifier + " 不能引用自身作为替代定义")
    return rows, hashlib.sha256(raw).hexdigest(), raw


def read_standard(path):
    """返回全部ID到原始完整定义的字典，保留扩展字段与历史停用记录。"""
    return _load(path)[0]


def _canonical(row):
    return json.dumps(row, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _missing_fields(before, after, prefix=""):
    if isinstance(before, dict) and isinstance(after, dict):
        for key, value in before.items():
            name = prefix + key
            if key not in after:
                yield name
            else:
                yield from _missing_fields(value, after[key], name + ".")


def _summary(row):
    if row is None:
        return None
    if row.get('record_type'):return copy.deepcopy(row)
    # 名称供界面简明展示，完整定义及扩展字段同时保留，避免维度等修订不可见。
    return {**row, "note": (row.get("note") or {}).get("name", ""),
        "table": (row.get("table") or {}).get("name", ""),
        "slot": (row.get("slot") or {}).get("name", ""),
        "note_definition": row.get("note"), "table_definition": row.get("table"),
        "slot_definition": row.get("slot")}


def _preview(current, candidate, old, new, old_hash, new_hash):
    unified=bool(next(iter(old.values())).get('record_type'))
    if unified!=bool(next(iter(new.values())).get('record_type')):
        raise ValueError('新旧标准层次不同，须通过完整迁移及复核，不能直接当作同格式修订')
    missing = sorted(set(old) - set(new))
    if missing:
        raise ValueError("完整修订不能遗漏旧ID，请保留并用status声明停用或替代：" + "、".join(missing[:20]))
    groups = {"added": [], "changed": [], "retired": []}
    changes = []
    for identifier, row in new.items():
        previous = old.get(identifier)
        if previous is None:
            kind = "added"
        else:
            if unified and previous['record_type']!=row['record_type']:
                raise ValueError(identifier+' 不能把原指标、维度或来源编号改成其他层次')
            if unified and previous['record_type']=='source_example' and _canonical(previous)!=_canonical(row):
                raise ValueError(identifier+' 的历史来源证据须保留；更正依据请另增来源记录')
            if unified and previous['record_type']=='legacy_conversion' and any(previous[key]!=row[key] for key in ('source_gold_hash','legacy_id','legacy_definition_hash')):
                raise ValueError(identifier+' 的旧定义身份不得改写')
            missing_fields = [] if unified else list(_missing_fields(previous, row))
            if missing_fields:
                raise ValueError(identifier + " 的修订遗漏已有字段，请保留并明确修订内容：" + "、".join(missing_fields[:20]))
            if _canonical(previous) == _canonical(row):
                continue
            kind = "retired" if previous.get("status") == "active" and row.get("status") != "active" else "changed"
        groups[kind].append(identifier)
        changes.append({"id": identifier, "kind": kind, "before": _summary(previous), "after": _summary(row),
            "fields": sorted(key for key in set(previous or {}) | set(row) if _canonical((previous or {}).get(key)) != _canonical(row.get(key)))})
    return {**{key: len(value) for key, value in groups.items()},
        **{key + "_ids": sorted(value) for key, value in groups.items()},
        "changes": changes, "source_hashes": {"current": old_hash, "candidate": new_hash},
        "current_path": str(Path(current).resolve()), "candidate_path": str(Path(candidate).resolve()),
        "validation_note": "仅完成结构与引用检查，不代表语义内容已经审核；修订含义应由人工或外部高级AI核实。"}


def _load_revision(path):
    raw=Path(path).read_bytes()
    rows=[json.loads(line,object_pairs_hook=_pairs,parse_constant=_invalid_number)
          for line in raw.decode('utf-8-sig').splitlines() if line.strip()]
    if any(isinstance(row,dict) and row.get('record_type') for row in rows):
        from .统一语义 import validate_records
        validate_records(rows)
        return {row['id']:row for row in rows},hashlib.sha256(raw).hexdigest(),raw
    return _load(path)


def preview_revision(current, candidate):
    """比较人工或外部高级AI提供的完整修订文件，不改写任何文件。"""
    old, old_hash, _ = _load_revision(current)
    new, new_hash, _ = _load_revision(candidate)
    return _preview(current, candidate, old, new, old_hash, new_hash)


def publish_revision(current, candidate, root=None, expected_hashes=None):
    """保存唯一新版本与变更记录，旧版不覆盖；可绑定用户刚检查的预览哈希。"""
    old, old_hash, _ = _load_revision(current)
    new, new_hash, candidate_bytes = _load_revision(candidate)
    report = _preview(current, candidate, old, new, old_hash, new_hash)
    if expected_hashes is not None and expected_hashes != report["source_hashes"]:
        raise ValueError("金标准文件自预览后发生变化或哈希不一致，请重新预览")
    folder = Path(root or ROOT).resolve() / "金标准" / "版本"
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f") + "_" + uuid.uuid4().hex[:8]
    output = folder / ("财务报表附注语义金标准_" + stamp + ".jsonl")
    report_path = folder / ("变更记录_" + stamp + ".json")
    with output.open("xb") as handle:
        handle.write(candidate_bytes)
    if hashlib.sha256(output.read_bytes()).hexdigest() != new_hash:
        raise ValueError("新版本保存后哈希核验失败，未完成版本发布")
    report.update({"published_path": str(output), "published_hash": new_hash,
        "published_at": datetime.now().isoformat(timespec="seconds")})
    with report_path.open("x", encoding="utf-8-sig") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2)
    return {"path": str(output), "sha256": new_hash, "report": report, "report_path": str(report_path)}


def ensure_compatible(record, new_gold_path):
    """旧映射仅在所用ID的新旧完整定义完全相同且仍active时可复用。"""
    if isinstance(record,dict) and record.get('schema_version')==2:
        from .统一语义 import read_standard_v2,ensure_standard_compatible
        return ensure_standard_compatible(record,read_standard_v2(new_gold_path))
    if not isinstance(record, dict) or not record.get("gold_path") or not record.get("gold_hash"):
        raise ValueError("映射缺少原金标准文件与哈希，无法核验定义")
    old, old_hash, _ = _load(record["gold_path"])
    if old_hash != record["gold_hash"]:
        raise ValueError("映射原金标准文件已变化，哈希与记录不一致")
    new, _, _ = _load(new_gold_path)
    mappings = record.get("mappings")
    if not isinstance(mappings, list):
        raise ValueError("映射缺少业务单元格清单")
    for mapping in mappings:
        identifier = mapping.get("slot_id") if isinstance(mapping, dict) else None
        if not _text(identifier) or identifier not in old or identifier not in new:
            raise ValueError("映射所用定义在新旧金标准中不存在：" + str(identifier))
        if old[identifier]["status"] != "active" or new[identifier]["status"] != "active":
            raise ValueError("映射所用定义已停用或替代，须重新识别：" + identifier)
        if _canonical(old[identifier]) != _canonical(new[identifier]):
            raise ValueError("映射所用定义已改变，不能仅按同ID复用，须重新识别：" + identifier)
        validate_mapping_scope(new[identifier], mapping)
    return True
