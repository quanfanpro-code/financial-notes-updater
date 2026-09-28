# -*- coding: utf-8 -*-
"""建立查找索引（按2026-09-24修订版设计第五节）。用法：python 建立索引.py [输出根]

索引方向：
- by_semantic：语义 → 维度定义、表现列表；
- by_legacy_versioned：（源标准哈希，旧ID）→ 本次语义与维度对应（带版本）；
- by_locator：源文件版本+定位 → 表现列表（允许多个历史修订）；
- by_dimension：维度 → 适用语义、实际取值及次数（不含模板占位）；
- by_text：科目、业务表、行列文字、已确认别名 → 相关语义列表。
"""
import json
import os
import sys
from collections import Counter, defaultdict

根 = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def 读jsonl(路径):
    with open(路径, encoding="utf-8-sig") as f:
        return [json.loads(行) for 行 in f if 行.strip()]


def main():
    语义 = 读jsonl(os.path.join(根, "数据", "统一语义记录.jsonl"))
    表现 = 读jsonl(os.path.join(根, "数据", "表现记录.jsonl"))
    旧新 = 读jsonl(os.path.join(根, "对应", "旧新对应.jsonl"))
    规则路径 = os.path.join(根, "对应", "维度规范对应.json")
    规则 = json.load(open(规则路径, encoding="utf-8-sig"))
    示例 = 读jsonl(os.path.join(根, "对应", "来源示例.jsonl"))
    按id = {r["semantic_id"]: r for r in 语义}

    表现按语义 = defaultdict(list)
    for p in 表现:
        表现按语义[p["semantic_id"]].append(p["representation_id"])
    旧到目标 = {}
    for r in 旧新:
        if r.get("版本") == "当前v1":
            t = (r.get("本次规范对应") or {}).get("目标semantic_id")
            if t:
                旧到目标[r["legacy_id"]] = t
    by_semantic = {}
    for r in 语义:
        表现列表 = list(表现按语义[r["semantic_id"]])
        汇集旧 = []
        if r["semantic_kind"] == "metric":
            汇集旧 = sorted(旧 for 旧, 目标 in 旧到目标.items() if 目标 == r["semantic_id"])
            for 旧 in 汇集旧:
                表现列表 += 表现按语义[旧]
        by_semantic[r["semantic_id"]] = {
            "semantic_kind": r["semantic_kind"], "lifecycle": r["lifecycle"],
            "dimension_definition_ids": r["dimension_definition_ids"],
            "表现": 表现列表,
        }
        if 汇集旧:
            by_semantic[r["semantic_id"]]["汇集对应旧ID"] = 汇集旧

    by_legacy_versioned = {}
    by_definition = {}
    for r in 旧新:
        三元键 = f"{r.get('源标准哈希')}|{r['legacy_id']}|{r.get('旧定义哈希')}"
        if 三元键 in by_definition:
            raise ValueError("版本定义重复：" + 三元键)
        by_definition[三元键] = {k: r.get(k) for k in ("对应ID", "版本", "旧定义引用", "v1状态")}
        by_definition[三元键]["目标semantic_id"] = (r.get("本次规范对应") or {}).get("目标semantic_id")
        版本键 = f"{r.get('源标准哈希')}|{r['legacy_id']}"
        if 版本键 in by_legacy_versioned:
            raise ValueError("同版本旧ID有多个定义，请用完整定义指纹消歧：" + 版本键)
        by_legacy_versioned[版本键] = {
            "版本": r.get("版本"), "v1状态": r["v1状态"], "superseded_by": r.get("superseded_by"),
            "旧定义哈希": r.get("旧定义哈希"),
            "本次核验结论": r.get("本次核验结论"), "本次规范对应": r.get("本次规范对应"),
            "转换ID列表": [c.get("id") for c in r.get("历史转换", [])],
            "五样式表现数": r.get("五样式表现数", 0),
        }

    by_locator = defaultdict(list)
    for p in 表现:
        loc = p["source_locator"]
        if p["source_kind"] == "five_style":
            by_locator[f"{loc['mapping_sha256']}|{loc.get('sheet')}|{loc.get('cell')}"].append(p["representation_id"])
            by_locator[f"{loc['mapping_file']}|{loc.get('sheet')}|{loc.get('cell')}"].append(p["representation_id"])
            by_locator[f"mapping_id|{loc['mapping_id']}"].append(p["representation_id"])
        elif p["source_kind"] == "a_fact":
            by_locator[f"{loc['sha256']}|{loc['sheet_cell']}"].append(p["representation_id"])
        elif p["source_kind"] == "v1_source":
            by_locator[f"v1|{loc.get('source_table_id')}"].append(p["representation_id"])
            by_locator[f"v1|{loc.get('company')}|{loc.get('report_file')}"].append(p["representation_id"])

    by_dimension = defaultdict(lambda: defaultdict(Counter))
    for p in 表现:
        for k, v in p.get("规范维度值", {}).items():
            键 = json.dumps(v, ensure_ascii=False, sort_keys=True) if not isinstance(v, str) else v
            by_dimension[k][p["semantic_id"]][键] += 1

    by_text = defaultdict(set)
    for r in 语义:
        m = r["meaning"]
        for 字段 in ("业务对象", "业务表", "指标含义", "指标"):
            if m.get(字段):
                by_text[str(m[字段])].add(r["semantic_id"])
        for t in (m.get("原表达") or {}).get("row_path", []) + (m.get("原表达") or {}).get("column_path", []):
            by_text[str(t)].add(r["semantic_id"])
        for t in sum((v for v in (m.get("原表达") or {}).get("aliases", {}).values()), []):
            by_text[str(t)].add(r["semantic_id"])
    for p in 表现:
        for t in (p.get("table_form") or {}).values():
            if isinstance(t, str) and t:
                by_text[t].add(p["semantic_id"])

    语义证据 = defaultdict(list)
    for e in 示例:
        for sid in e["去向"].get("关联规范语义", []):
            语义证据[sid].append(e["示例ID"])
    索引 = {
        "说明": "查询不要求先选择v1/v2或A/B；by_locator键含源文件版本哈希；by_dimension与表现维度值均按规范维度，"
              "原名经维度别名归一，取值经取值别名归一；metric语义的表现已沿本次规范对应汇集旧定义表现。",
        "维度别名": {原名: 对["规范维度"] for 原名, 对 in 规则["原名对应"].items()},
        "取值别名": 规则["取值别名"],
        "表现维度值": {p["representation_id"]: p["规范维度值"] for p in 表现 if p.get("规范维度值")},
        "by_semantic": by_semantic,
        "by_legacy_versioned": by_legacy_versioned,
        "by_definition": by_definition,
        "by_source_example": {e["示例ID"]: e["去向"] for e in 示例},
        "by_semantic_evidence": {s: sorted(es) for s, es in 语义证据.items()},
        "by_locator": {k: v for k, v in sorted(by_locator.items())},
        "by_dimension": {d: {s: dict(c.most_common()) for s, c in 语义们.items()}
                        for d, 语义们 in sorted(by_dimension.items())},
        "by_text": {t: sorted(s) for t, s in sorted(by_text.items())},
    }
    os.makedirs(os.path.join(根, "索引"), exist_ok=True)
    with open(os.path.join(根, "索引", "查找索引.json"), "w", encoding="utf-8-sig", newline="\n") as f:
        json.dump(索引, f, ensure_ascii=False, indent=1, sort_keys=True)
    print(f"索引已写出：by_semantic {len(by_semantic)}，by_legacy_versioned {len(by_legacy_versioned)}，"
          f"by_locator {len(by_locator)}，by_dimension {len(by_dimension)}，by_text {len(by_text)}")


def 规范取值(索引, 维度名, 值):
    return 索引.get("取值别名", {}).get(维度名, {}).get(值, 值) if isinstance(值, str) else 值


def 维度符合(实际, 条件):
    """结构条件须完整相同；文字简称只允许完整、单成员集合，不能把合计或补集当成明细。"""
    if 实际 == 条件:
        return True
    return (isinstance(条件, str) and isinstance(实际, dict)
            and 实际.get("mode") == "members" and 实际.get("completeness") == "complete"
            and 实际.get("members") == [条件])


def 查询语义(索引, 语义ID, 维度筛选=None):
    """按规范语义查询表现ID集合。维度筛选的键与取值先经索引中的别名表解析；
    别名或规范维度值缺失时按无匹配处理（查询能力由索引内容决定）。"""
    条目 = 索引.get("by_semantic", {}).get(语义ID)
    if not 条目:
        return set()
    命中 = set(条目.get("表现", []))
    if 维度筛选:
        规范筛选 = {}
        for k, v in 维度筛选.items():
            规范键 = 索引.get("维度别名", {}).get(k, k)
            if isinstance(规范键, list):
                raise ValueError(f"筛选维度名有歧义，候选：{规范键}")
            值 = 规范取值(索引, 规范键, v)
            if 规范键 in 规范筛选 and 规范筛选[规范键] != 值:
                return set()
            规范筛选[规范键] = 值
        维度值表 = 索引.get("表现维度值", {})
        命中 = {i for i in 命中
                if all(k in 维度值表.get(i, {}) and 维度符合(维度值表[i][k], v)
                       for k, v in 规范筛选.items())}
    return 命中


def 查询维度(索引, 维度名, 取值, 适用语义=None):
    """按维度名与取值查询表现ID集合。名称对应多个规范维度时返回候选及条件，不静默选择。"""
    目标 = 索引.get("维度别名", {}).get(维度名, 维度名)
    if isinstance(目标, list):
        return {"候选": 目标, "说明": "维度名对应多个规范维度，返回候选及条件，未静默选择"}
    值 = 规范取值(索引, 目标, 取值)
    维度值表 = 索引.get("表现维度值", {})
    命中 = {pid for pid, 值们 in 维度值表.items() if 目标 in 值们 and 维度符合(值们[目标], 值)}
    if 适用语义 is not None:
        命中 &= set(索引.get("by_semantic", {}).get(适用语义, {}).get("表现", []))
    return 命中


if __name__ == "__main__":
    main()
