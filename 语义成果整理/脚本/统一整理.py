# -*- coding: utf-8 -*-
"""统一整理主脚本（按2026-09-24修订版需求/设计）：

- 全部已确认成果按同一套规范字段整理；v2 approved转换只作带版本核验的对应，不是准入门槛。
- 输入固定清单并校验哈希，运行中输入变化则不发布。
- 输出：数据/统一语义记录、维度字典、表现记录、上下文证据；对应/旧新对应；问题/问题清单。
- 幂等：数据文件不含时间等元信息，规范序列化，重复运行字节一致。
"""
import hashlib
import copy
import json
import os
import re
import sys
from collections import Counter, defaultdict

项目 = r"C:\Users\27651\Desktop\附注自动更新系统"
数字资产 = r"D:\BaiduSyncdisk\数字资产\审计数据自动搬运系统"


def 解析参数(argv):
    位置, 输入清单, 已有成果 = [], None, None
    i = 0
    while i < len(argv):
        if argv[i] == "--输入清单" and i + 1 < len(argv):
            输入清单 = argv[i + 1]; i += 2
        elif argv[i] == "--已有成果" and i + 1 < len(argv):
            已有成果 = argv[i + 1]; i += 2
        else:
            位置.append(argv[i]); i += 1
    return (位置[0] if 位置 else None), 输入清单, 已有成果


输出根, 输入清单路径, 已有成果路径 = 解析参数(sys.argv[1:])
输出根 = 输出根 or os.path.join(项目, "语义成果整理")
V1路径 = os.path.join(项目, "金标准", "财务报表附注语义金标准-v1.jsonl")
V2路径 = os.path.join(项目, "金标准", "版本", "财务报表附注语义金标准_20260922_132731_236607_5547978f.jsonl")
索引路径 = os.path.join(项目, "金标准", "五种样式人工映射", "语义示例索引.json")
映射根 = os.path.join(数字资产, "work", "excel-mapping-v1")
A路径 = os.path.join(项目, "测试结果", "真实语义识别续跑_A_20260922_050354_458720",
                 "导入外部复核_20260922_051437_942503_b323c5", "附注表格语义映射.json")
A复核路径 = os.path.join(项目, "测试结果", "真实语义识别续跑_A_20260922_050354_458720",
                   "导入外部复核_20260922_051437_942503_b323c5", "外部复核确认记录.json")

追踪字段 = {"source_row_path", "source_column_path", "source_value_type", "field_path"}
占位 = re.compile(r"^<.+>$")
A确认方法 = {"model_v2_two_rounds": "模型双轮识别",
             "model_v2_two_rounds_equivalence": "模型双轮识别＋同义复核",
             "external_semantic_review_v2": "外部复核"}


def sha256(路径):
    h = hashlib.sha256()
    with open(路径, "rb") as f:
        for 块 in iter(lambda: f.read(1 << 20), b""):
            h.update(块)
    return h.hexdigest()


def 读jsonl(路径):
    with open(路径, encoding="utf-8-sig") as f:
        return [json.loads(行) for 行 in f if 行.strip()]


def 写jsonl(路径, 记录们):
    os.makedirs(os.path.dirname(路径), exist_ok=True)
    with open(路径, "w", encoding="utf-8-sig", newline="\n") as f:
        for r in 记录们:
            f.write(json.dumps(r, ensure_ascii=False, sort_keys=True) + "\n")


def 稳定摘要(内容):
    return hashlib.sha256(json.dumps(内容, ensure_ascii=False, sort_keys=True,
                                    separators=(",", ":")).encode("utf-8")).hexdigest()


def 来源位置(表现):
    """逻辑来源与格子/表中行列；文件内容哈希只作版本证据，不作表现身份。"""
    loc, 表式 = 表现.get("source_locator", {}), 表现.get("table_form", {})
    if 表现["source_kind"] == "v1_source":
        return {"来源": "附注证据", **{k: loc.get(k) for k in
                ("company", "report_file", "source_table_id", "section")},
                "行": 表式.get("row_path"), "列": 表式.get("column_path"),
                "表名": 表式.get("table_name"), "附注": 表式.get("note_name")}
    if 表现["source_kind"] == "five_style":
        return {"来源": loc.get("mapping_file"), "sheet": loc.get("sheet"),
                "cell": loc.get("cell"), "mapping_id": loc.get("mapping_id")}
    return {"来源": loc.get("file"), "位置": loc.get("location") or loc.get("sheet_cell")}


def 表现身份(表现):
    return "rep:" + 稳定摘要({"来源位置": 来源位置(表现),
                "语义": 表现["semantic_id"], "业务维度": 表现.get("dimension_values", {}),
                "规范维度": 表现.get("规范维度值", {}), "占位与缺失": 表现.get("占位与缺失", [])})


def main():
    # ---------- 固定输入清单并校验哈希 ----------
    if os.path.exists(os.path.join(输出根, "数据", "表现记录.jsonl")):
        raise SystemExit("输出目录已有成果，请使用新的候选目录；核验后再备份并交付")
    global V1路径, V2路径, 索引路径, A路径, A复核路径
    映射文件表 = None  # [(相对路径, 绝对路径)]
    if 输入清单路径:
        单 = json.load(open(输入清单路径, encoding="utf-8-sig"))
        固 = 单["固定输入"]
        V1路径, V2路径, 索引路径 = 固["v1"]["路径"], 固["v2当前版本"]["路径"], 固["五样式索引"]["路径"]
        A路径, A复核路径 = 固["A端映射"]["路径"], 固["A端外部复核记录"]["路径"]
        映射文件表 = [(e["相对路径"], e["路径"]) for e in 固["五样式原映射"]]
        期望清单 = {固[k]["路径"]: 固[k]["sha256"] for k in ("v1", "v2当前版本", "五样式索引", "A端映射", "A端外部复核记录")}
        期望清单.update({e["路径"]: e["sha256"] for e in 固["五样式原映射"]})
        for p, 期望 in 期望清单.items():
            if not os.path.exists(p):
                raise SystemExit(f"输入缺失：{p}")
            if sha256(p) != 期望:
                raise SystemExit(f"输入哈希与固定清单不一致，本批不发布：{p}")
        清单 = 期望清单
    else:
        清单 = {p: sha256(p) for p in (V1路径, V2路径, 索引路径, A路径, A复核路径)}
    索引 = json.load(open(索引路径, encoding="utf-8-sig"))
    if 映射文件表 is None:
        映射文件表 = []
        for 相对路径, 期望 in 索引["files"].items():
            if not 相对路径.startswith("work/"):
                continue
            全 = os.path.join(数字资产, 相对路径)
            if not os.path.exists(全):
                raise SystemExit(f"输入缺失：{全}")
            实际 = sha256(全)
            if 实际 != 期望:
                raise SystemExit(f"输入哈希变化，本批不发布：{全}")
            清单[全] = 实际
            映射文件表.append((相对路径, 全))
    规则路径 = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "对应", "维度规范对应.json")
    if 输入清单路径 and "维度规范规则" in 固:
        规则路径 = 固["维度规范规则"]["路径"]
        if sha256(规则路径) != 固["维度规范规则"]["sha256"]:
            raise SystemExit("维度规则与固定输入清单不一致，本批不发布")
    清单[规则路径] = sha256(规则路径)
    print(f"输入清单固定：{len(清单)} 个文件哈希一致")

    # ---------- 读取来源 ----------
    v1记录 = 读jsonl(V1路径)
    v1 = {r["id"]: r for r in v1记录}
    if len(v1) != len(v1记录):
        raise SystemExit("同一输入版本出现重复旧ID，本批不发布")
    v1哈希 = 清单[V1路径]
    v2指标, v2维度, 转换记录们, v2示例们 = {}, {}, [], []
    for r in 读jsonl(V2路径):
        if r["record_type"] == "metric":
            v2指标[r["id"]] = r
        elif r["record_type"] == "dimension":
            v2维度[r["id"]] = r
        elif r["record_type"] == "legacy_conversion":
            转换记录们.append(r)
        elif r["record_type"] == "source_example":
            v2示例们.append(r)
    A映射 = json.load(open(A路径, encoding="utf-8-sig"))
    A源哈希 = A映射.get("source_hash") or ""

    # 五样式原映射按 (文件, mapping_id) 索引，用于补充原值状态
    原映射 = {}
    排除证据 = []
    for 相对路径, 全路径 in 映射文件表:
        for r in 读jsonl(全路径):
            if r.get("status") == "结构格排除" or "target_slot_id" not in r:
                排除证据.append({"mapping_file": 相对路径, **r})
            else:
                原映射[(相对路径, r["mapping_id"])] = r

    # ---------- 来源示例解析与版本定义证据（任务1） ----------
    def 定义哈希(r):
        return hashlib.sha256(json.dumps(r, sort_keys=True, ensure_ascii=False,
                                         separators=(",", ":")).encode()).hexdigest()

    批准定义原文 = {}  # 供按内容承接：只来自固定输入中可重算哈希的原定义
    内嵌定义 = {}   # (源标准哈希, 旧ID, 定义哈希) -> {定义, 示例ID, 原状态}
    示例解析 = {}   # 示例ID -> [解析条目]
    for e in v2示例们:
        o = e.get("origin", {})
        条目们 = []
        for t in e.get("text_evidence", []):
            try:
                内 = json.loads(t)
            except Exception:
                条目们.append({"类型": "非结构化原文"})
                continue
            if not isinstance(内, dict):
                条目们.append({"类型": "结构化非对象"})
                continue
            条 = {"类型": "结构化", "键": sorted(内.keys())}
            if "原定义" in 内 and isinstance(内["原定义"], dict):
                条["原定义哈希"] = 定义哈希(内["原定义"])
                for k in ("原状态", "原角色", "原值类型", "历史后继"):
                    条[k] = 内.get(k)
                源哈希 = o.get("source_gold_hash") or o.get("file_hash")
                if o.get("legacy_id") and 源哈希:
                    内嵌定义[(源哈希, o["legacy_id"], 条["原定义哈希"])] = {
                        "定义": 内["原定义"], "示例ID": e["id"], "原状态": 内.get("原状态")}
            if "原定义" in 内 and isinstance(内["原定义"], dict):
                d = 内["原定义"]
                批准定义原文[(d.get("id") or o.get("legacy_id"), 定义哈希(d))] = d
            if isinstance(内.get("original_definitions"), list):
                定义列表 = []
                for d in 内["original_definitions"]:
                    if not isinstance(d, dict) or not d.get("id"):
                        continue
                    h = 定义哈希(d)
                    批准定义原文[(d["id"], h)] = d
                    定义列表.append({"旧ID": d["id"], "定义哈希": h,
                                     "源标准哈希": o.get("source_gold_hash") or o.get("file_hash")})
                条["原定义列表"] = 定义列表
            条目们.append(条)
        示例解析[e["id"]] = 条目们

    转换按旧 = defaultdict(list)
    for r in 转换记录们:
        转换按旧[r["legacy_id"]].append(r)

    def 差异字段(a, b):
        return sorted(k for k in set(a) | set(b) if a.get(k) != b.get(k))

    def 适用条件(定义):
        适用 = 定义.get("applicability") or {}
        条件 = {}
        if 适用.get("industries"):
            条件["行业"] = 适用["industries"]
        if 适用.get("optional") is not None:
            条件["可选披露"] = 适用["optional"]
        范围 = 适用.get("report_scopes") or []
        if 范围:
            条件["报表范围"] = [{"scope": x.get("scope"), "依据": x.get("reason"),
                                "审查": x.get("reviewed_by"), "审查时间": x.get("reviewed_at")}
                               for x in 范围]
        return 条件 or None

    def 建对应(旧ID, 定义, 定义哈希值, 适用转换, 版本说明):
        """从适用转换中核验并建立本次规范对应。目标、绑定和条件逐条来自本条适用转换。"""
        approved适用 = [c for c in 适用转换 if c.get("status") == "approved"]
        if not approved适用:
            return None, None
        目标组 = defaultdict(list)
        维度绑定 = []
        for c in approved适用:
            t = c.get("target") or {}
            if t.get("record_type") == "metric_fact" and t.get("metric_id"):
                目标组[t["metric_id"]].append(c)
            elif t.get("record_type") == "dimension_binding":
                维度绑定.append({"转换ID": c.get("id"), "target": t})
        有效 = {m for m in 目标组 if m in v2指标 and v2指标[m].get("status") == "active"}
        停用 = {m for m in 目标组 if m in v2指标 and v2指标[m].get("status") != "active"}
        缺目标 = {m for m in 目标组 if m not in v2指标}
        结论差异 = []
        对应 = None
        if len(目标组) > 1:
            结论差异.append(f"多个approved目标：{sorted(目标组)}，保留各条依据不硬选")
        if 停用:
            结论差异.append(f"目标已停用：{sorted(停用)}，不作当前有效对应")
        if 缺目标:
            结论差异.append(f"目标指标不存在：{sorted(缺目标)}")
        if len(有效) == 1 and len(目标组) == 1 and not 维度绑定:
            目标 = next(iter(有效))
            绑定集 = {稳定摘要(c["target"].get("dimension_bindings")) for c in 目标组[目标]}
            if len(绑定集) != 1:
                return None, "同一目标的approved维度绑定相互冲突，未静默选第一条"
            记录 = 目标组[目标][0]
            绑定 = 记录["target"].get("dimension_bindings") or {}
            声明 = {d["id"] for d in v2指标[目标].get("dimension_refs", [])}
            必需 = {d["id"] for d in v2指标[目标].get("dimension_refs", []) if d.get("required")}
            if not 必需 <= set(绑定) or not set(绑定) <= 声明 or any(
                    not isinstance(b, dict) or b.get("op") not in ("constant", "context")
                    or (b["op"] == "constant" and "value" not in b)
                    or (b["op"] == "context" and not b.get("key")) for b in 绑定.values()):
                return None, "批准记录的维度绑定与目标声明不符，保留原记录及冲突"
            对应 = {"目标semantic_id": 目标,
                    "固定维度限制": 记录["target"].get("dimension_bindings"),
                    "适用条件": 适用条件(定义),
                    "证据": [c.get("evidence_refs") for c in 目标组[目标]],
                    "处理说明": f"已建立对应（{版本说明}；目标状态与维度绑定已逐条核对）"}
        elif not 目标组 and 维度绑定:
            对应 = {"目标semantic_id": None, "维度绑定对应": 维度绑定,
                    "适用条件": 适用条件(定义),
                    "证据": [c.get("evidence_refs") for c in approved适用],
                    "处理说明": f"已建立维度绑定（{版本说明}），无指标目标"}
        return 对应, "；".join(结论差异) or None

    def 核验版本定义(旧ID, 定义, 源哈希, 定义哈希值, 版本名, 定义引用):
        """一条（源标准哈希，旧ID，定义哈希）版本定义的完整核验。"""
        适用, 不适用 = [], []
        for c in 转换按旧.get(旧ID, []):
            if c.get("source_gold_hash") != 源哈希:
                不适用.append({"转换ID": c.get("id"), "原因": "转换绑定的源标准版本不同"})
                continue
            if c.get("legacy_definition_hash") != 定义哈希值:
                不适用.append({"转换ID": c.get("id"), "原因": "同版本但定义内容哈希不一致"})
                continue
            适用.append(c)
        对应, 差异 = 建对应(旧ID, 定义, 定义哈希值, 适用, f"转换绑定{版本名}定义")
        return 适用, 不适用, 对应, 差异

    历史定义按旧 = defaultdict(list)  # 旧ID -> [(源哈希, 定义哈希, 定义, 示例ID)]
    for (源哈希, 旧ID, 定义哈希值), 证 in 内嵌定义.items():
        历史定义按旧[旧ID].append((源哈希, 定义哈希值, 证["定义"], 证["示例ID"]))

    旧新 = []
    对应表 = {}  # 对应ID -> 本次规范对应（当前版本条目的，供表现引用）
    五样式引用 = Counter(e["legacy_slot_id"] for e in 索引["examples"])

    # 当前版本：v1全部记录逐条（含非active的历史保留）
    for 旧ID, r in v1.items():
        当前哈希 = 定义哈希(r)
        条目 = {"record_type": "legacy_link", "legacy_id": 旧ID, "版本": "当前v1",
                "源标准哈希": v1哈希, "旧定义哈希": 当前哈希,
                "旧定义引用": {"文件": V1路径, "sha256": v1哈希, "记录id": 旧ID},
                "v1状态": r["status"], "superseded_by": r.get("superseded_by"),
                "五样式表现数": 五样式引用.get(旧ID, 0)}
        对应ID = "link:" + 稳定摘要([v1哈希, 旧ID, 当前哈希])
        条目["对应ID"] = 对应ID
        if r["status"] != "active":
            条目["历史转换"] = 转换按旧.get(旧ID, [])
            条目["本次核验结论"] = "非active定义，仅作历史追溯，不作当前有效对应"
            条目["本次规范对应"] = {"处理": "历史保留（被替代）" if r["status"] == "superseded" else "历史保留（停用）"}
            旧新.append(条目)
            continue
        适用, 不适用, 对应, 差异 = 核验版本定义(旧ID, r, v1哈希, 当前哈希, "当前版本", 条目["旧定义引用"])
        结论们 = []
        if 差异:
            结论们.append(差异)
        if 对应 is None:
            # 文件增删无关记录或重排时，整文件哈希会变，但单条完整定义仍可核对。
            # 若仅来源证据增加，必须有原定义全文证明业务字段未变，不能只凭旧ID承接。
            承接 = []
            for c in 转换按旧.get(旧ID, []):
                if c.get("status") != "approved":
                    continue
                原指纹 = c.get("legacy_definition_hash")
                基线 = 批准定义原文.get((旧ID, 原指纹))
                if 原指纹 == 当前哈希 or (基线 is not None and set(差异字段(r, 基线)) <= {"sources"}):
                    承接.append(c)
            if 承接:
                对应, 承接差异 = 建对应(旧ID, r, 当前哈希, 承接, "定义内容相同或仅来源证据变化")
                if 对应:
                    适用 = 承接
                    对应["版本适用依据"] = [{"转换ID": c["id"],
                        "转换原标准哈希": c["source_gold_hash"], "转换原定义哈希": c.get("legacy_definition_hash"),
                        "当前定义哈希": 当前哈希,
                        "差异字段": 差异字段(r, 批准定义原文[(旧ID, c["legacy_definition_hash"])])
                            if (旧ID, c["legacy_definition_hash"]) in 批准定义原文 else []} for c in 承接]
                elif 承接差异:
                    结论们.append(承接差异)
        if 对应 is None:
            # 版本兼容路径：历史转换绑定的定义与当前定义内容一致，或仅差applicability
            for 源哈希, 历史哈希, 历史定义, 示例ID in 历史定义按旧.get(旧ID, []):
                差异集 = 差异字段(r, 历史定义)
                if not 差异集:
                    h适用, _, h对应, h差异 = 核验版本定义(旧ID, 历史定义, 源哈希, 历史哈希, "历史版本", None)
                    if h对应:
                        对应 = {**h对应,
                                "处理说明": "已建立对应（版本兼容：历史定义内容与当前一致；目标状态与维度绑定已逐条核对）"}
                        适用 = h适用
                        break
                elif 差异集 == ["applicability"]:
                    h适用, _, h对应, h差异 = 核验版本定义(旧ID, 历史定义, 源哈希, 历史哈希, "历史版本", None)
                    if h对应:
                        对应 = {**h对应,
                                "处理说明": "已建立对应（历史定义与当前仅差applicability单体适用审查记录，"
                                        "业务含义一致；适用范围按审查记录注明，未扩展）"}
                        适用 = h适用
                        break
                    else:
                        结论们.append("与历史定义仅applicability不同，但历史转换未批准")
                else:
                    结论们.append("当前定义相对历史定义已演变（差异字段：" + "、".join(差异集)
                                + "），历史转换不延伸适用当前定义")
        适用IDs = {c["id"] for c in 适用}
        不适用 = [c for c in 不适用 if c["转换ID"] not in 适用IDs]
        条目["历史转换"] = 适用
        条目["不适用转换"] = 不适用
        条目["本次核验结论"] = "；".join(结论们) if 结论们 else (
            "版本核验通过" if 对应 else "无approved转换适用当前定义")
        条目["本次规范对应"] = 对应 or {"目标semantic_id": None,
                                    "处理": "未建立指标对应，已完成规范拆分（旧定义自身作为规范语义整理）"}
        条目["对应ID"] = 对应ID
        对应表[对应ID] = 条目["本次规范对应"]
        旧新.append(条目)

    # 历史版本：内嵌定义逐条成条（完整覆盖，不仅抽样）
    for 旧ID, 版本们 in sorted(历史定义按旧.items()):
        for 源哈希, 历史哈希, 历史定义, 示例ID in 版本们:
            当前 = v1.get(旧ID)
            与当前一致 = 当前 is not None and 定义哈希(当前) == 历史哈希
            适用, 不适用, 对应, 差异 = 核验版本定义(旧ID, 历史定义, 源哈希, 历史哈希, "历史版本", None)
            结论们 = []
            if 差异:
                结论们.append(差异)
            if 当前 is not None and not 与当前一致:
                差异集 = 差异字段(当前, 历史定义)
                if 差异集 == ["applicability"]:
                    结论们.append("历史定义与当前定义仅差applicability单体适用审查记录，业务含义一致")
                else:
                    结论们.append("历史定义与当前定义内容不同（差异字段：" + "、".join(差异集) + "）")
            elif 与当前一致:
                结论们.append("历史定义内容与当前版本一致（版本兼容证据）")
            else:
                结论们.append("旧ID不在当前v1，仅历史追溯")
            旧新.append({
                "record_type": "legacy_link", "legacy_id": 旧ID, "版本": "历史v1",
                "源标准哈希": 源哈希, "旧定义哈希": 历史哈希,
                "旧定义引用": {"来源示例": 示例ID, "说明": "历史定义取自v2 source_example内嵌定义，哈希已重算核对"},
                "对应ID": "link:" + 稳定摘要([源哈希, 旧ID, 历史哈希]),
                "v1状态": 内嵌定义[(源哈希, 旧ID, 历史哈希)]["原状态"],
                "superseded_by": (当前 or {}).get("superseded_by"),
                "历史转换": 适用, "不适用转换": 不适用,
                "本次核验结论": "；".join(结论们),
                "本次规范对应": 对应 or {"目标semantic_id": None,
                                    "处理": "历史转换未决或无历史转换，历史定义仅作追溯"},
                "与当前定义一致": 与当前一致,
                "五样式表现数": 0,
            })

    # ---------- 统一语义记录（legacy与metric同一套meaning字段） ----------
    规则 = json.load(open(规则路径, encoding="utf-8-sig"))
    规范维度 = 规则["规范维度"]
    原名对应 = 规则["原名对应"]
    取值别名表 = 规则["取值别名"]
    v1v2关系 = 规则["v1与v2维度关系"]

    值类型说明映射 = {"monetary": "货币金额", "percentage": "比率/百分比", "text": "文本",
                 "number": "数值", "boolean": "是/否", "date": "日期", "enum": "枚举"}

    def 期间类型(声明们):
        for d in 声明们:
            if d["name"] == "period":
                return d.get("type") or "原资料未给出"
        return "不适用（无期间维度声明）"

    def 业务限制(r):
        限制 = {}
        bp = r.get("blank_policy") or {}
        if bp.get("default_interpretation"):
            限制["空白处理"] = bp["default_interpretation"]
        if (r.get("calculation") or {}).get("expression"):
            限制["计算式"] = r["calculation"]["expression"]
        if r.get("scope"):
            限制["编制范围"] = r["scope"]
        return 限制 or "原资料未给出特别限制"

    定义计量对应 = {x["legacy_id"]:x for x in 旧新 if x["版本"]=="当前v1"}
    语义们 = {}
    for 旧ID, r in v1.items():
        if r["status"] != "active":
            continue
        已对应 = 定义计量对应[旧ID]
        目标指标 = (已对应.get("本次规范对应") or {}).get("目标semantic_id")
        已知基础 = (v2指标.get(目标指标, {}).get("definition") or {}).get("measurement_basis")
        基础依据 = ({"来源":"适用的已确认指标对应","metric_id":目标指标,"对应引用":已对应["对应ID"]}
                    if 已知基础 is not None else
                    {"来源":"原v1未单列计量基础属性","旧ID":旧ID,
                     "说明":"保留原指标含义和原定义；此属性未单列，不表示该语义尚未确认，不用值类型冒充"})
        语义们[旧ID] = {
            "record_type": "unified_semantic", "semantic_id": 旧ID, "semantic_kind": "legacy_slot",
            "meaning": {
                "业务对象": r["note"]["name"], "业务表": r["table"]["name"],
                "指标含义": r["slot"]["name"], "槽位角色": r["slot"]["role"],
                "值类型": r["value_type"],
                "值类型说明": 值类型说明映射.get(r["value_type"], r["value_type"]),
                "计量基础": 已知基础,
                "计量基础依据": 基础依据,
                "期间类型": 期间类型(r.get("dimensions", [])),
                "期间角色": next((d.get("role") for d in r.get("dimensions", [])
                                  if d["name"] == "period"), None),
                "适用范围": r.get("applicability") or "原资料未给出",
                "业务限制": 业务限制(r),
                "汇总角色": r.get("aggregation", {}).get("role"),
                "汇总关系": r.get("aggregation"),
                "原表达": {"row_path": r["row_path"], "column_path": r["column_path"],
                           "aliases": r.get("aliases")},
            },
            "dimension_definition_ids": [d["name"] for d in r.get("dimensions", [])],
            "definition_evidence_refs": [{"类型": "v1记录", "旧ID": 旧ID, "文件": V1路径, "sha256": v1哈希}],
            "lifecycle": {"状态": "active", "来源标准": "v1", "来源标准哈希": v1哈希},
        }
    for 指标ID, m in v2指标.items():
        d = m.get("definition", {})
        语义们[指标ID] = {
            "record_type": "unified_semantic", "semantic_id": 指标ID, "semantic_kind": "metric",
            "meaning": {
                "业务对象": d.get("economic_object") or "原资料未给出",
                "指标含义": d.get("measure") or "原资料未给出",
                "槽位角色": "measure",
                "值类型": d.get("value_type") or "原资料未给出",
                "计量基础": d.get("measurement_basis") or "原资料未给出",
                "期间类型": d.get("time_kind") or "原资料未给出",
                "适用范围": m.get("applicability") or "原资料未给出",
                "业务限制": d.get("constraints") or "原资料未给出特别限制",
            },
            "dimension_definition_ids": [x["id"] for x in m.get("dimension_refs", [])],
            "definition_evidence_refs": [{"类型": "v2指标", "文件": V2路径, "sha256": 清单[V2路径], "id": 指标ID}],
            "lifecycle": {"状态": m.get("status"), "来源标准": "v2", "来源标准哈希": 清单[V2路径]},
        }

    # ---------- 表现记录 ----------
    表现们 = []
    当前对应按旧ID = {条目["legacy_id"]: 条目["本次规范对应"]
                   for 条目 in 旧新 if 条目.get("版本") == "当前v1"}

    def 声明期间角色(r):
        for d in r.get("dimensions", []):
            if d["name"] == "period" and d.get("role"):
                return d["role"]
        return None

    def 定义固定维度值(旧ID, r):
        """只复制适用转换明示的常量，不从末级行名猜类别或合计范围。"""
        对应 = 当前对应按旧ID.get(旧ID) or {}
        return {k: copy.deepcopy(b["value"]) for k, b in (对应.get("固定维度限制") or {}).items()
                if b.get("op") == "constant" and "value" in b}

    def 规范原名值(原名值):
        值 = {}
        for k, v in 原名值.items():
            对 = 原名对应.get(k)
            if 对 is None:
                if k not in 规范维度:
                    continue  # 未归纳名不静默选择，原值保留在dimension_values
                目标 = k
            elif isinstance(对["规范维度"], list):
                continue  # 歧义名不静默选择，原值保留在dimension_values
            else:
                目标 = 对["规范维度"]
            值[目标] = 取值别名表.get(目标, {}).get(str(v), v)
        return 值

    表现按身份 = {}
    重复证据 = []

    def 挂表现(语义ID, 表现):
        assert 语义ID in 语义们, f"表现指向不存在的语义: {语义ID}"
        表现["semantic_id"] = 语义ID
        表现["representation_id"] = 表现身份(表现)
        pid = 表现["representation_id"]
        if pid in 表现按身份:
            既有 = 表现按身份[pid]
            for k in ("证据引用", "来源原文证据", "问题引用"):
                for e in 表现.get(k, []):
                    if e not in 既有.setdefault(k, []):
                        既有[k].append(e)
            重复证据.append({"表现ID": pid, "处理": "同一表现合并全部不同证据，身份不变"})
            return
        表现按身份[pid] = 表现
        表现们.append(表现)

    def 内容哈希(*部分):
        return 稳定摘要(部分)

    当前引用 = {x["legacy_id"]: x["对应ID"] for x in 旧新 if x["版本"] == "当前v1"}

    for 旧ID, r in v1.items():
        if r["status"] != "active":
            continue
        for s in r.get("sources", []):
            身份内容 = {"旧ID": 旧ID, "company": s.get("company"), "report_file": s.get("report_file"),
                      "source_table_id": s.get("source_table_id"), "section": s.get("section"),
                      "row": s.get("original_row_path"), "col": s.get("original_column_path"),
                      "excerpt": s.get("searchable_excerpt")}
            表现ID = "v1src:" + 内容哈希(
                "v1_source", 旧ID, s.get("company"), s.get("report_file"), s.get("source_table_id"),
                s.get("section"), s.get("original_row_path"), s.get("original_column_path"),
                s.get("searchable_excerpt"))
            规范值 = {}
            角色 = 声明期间角色(r)
            if 角色:
                规范值["period_role"] = 角色
            规范值.update(定义固定维度值(旧ID, r))
            表现 = {
                "record_type": "representation", "representation_id": 表现ID,
                "source_kind": "v1_source", "定义旧ID": 旧ID,
                "对应引用": 当前引用[旧ID],
                "table_form": {"note_name": s.get("note_title"), "table_name": s.get("table_context"),
                               "row_path": s.get("original_row_path"),
                               "column_path": s.get("original_column_path"),
                               "searchable_excerpt": s.get("searchable_excerpt")},
                "dimension_values": {}, "规范维度值": 规范值, "占位与缺失": [],
                "source_locator": {"company": s.get("company"), "report_file": s.get("report_file"),
                                   "source_table_id": s.get("source_table_id"),
                                   "section": s.get("section"), "raw_unit_text": s.get("raw_unit_text")},
                "raw_state": "evidence",
                "证据引用": [f"v1:{旧ID}"],
                "确认依据": {"方法": "原金标准v1来源证据"},
                "问题引用": [],
            }
            表现["来源原文证据"] = [s]
            挂表现(旧ID, 表现)

    未匹配原映射 = []
    for e in 索引["examples"]:
        旧ID = e["legacy_slot_id"]
        o = e["origin"]
        键 = (o["mapping_file"], o["mapping_id"])
        原 = 原映射.get(键)
        if 原 is None:
            未匹配原映射.append(str(键))
        业务维度, 追踪, 占位们 = {}, {}, []
        for k, v in e.get("legacy_dimensions", {}).items():
            if k in 追踪字段:
                追踪[k] = v
            elif isinstance(v, str) and 占位.match(v):
                占位们.append({"dimension": k, "原值": v, "处理": "模板占位，未登记为真实取值"})
            else:
                业务维度[k] = v
        规范值 = 规范原名值(业务维度)
        角色 = 声明期间角色(v1[旧ID]) if 旧ID in v1 else None
        if 角色 and "period_role" not in 规范值:
            规范值["period_role"] = 角色
        if 旧ID in v1:
            规范值.update(定义固定维度值(旧ID, v1[旧ID]))
        挂表现(旧ID, {
            "record_type": "representation",
            "representation_id": "fs:" + 内容哈希("five_style", o["mapping_id"], 旧ID),
            "source_kind": "five_style", "定义旧ID": 旧ID,
            "对应引用": 当前引用[旧ID],
            "table_form": {"row_semantics": e.get("row_semantics"),
                           "column_semantics": e.get("column_semantics"),
                           "table_semantics": e.get("table_semantics")},
            "dimension_values": 业务维度, "规范维度值": 规范值, "占位与缺失": 占位们,
            "derivation": 追踪,
            "source_locator": {"mapping_file": o["mapping_file"], "mapping_sha256": o["mapping_sha256"],
                               "mapping_id": o["mapping_id"], "sheet": o.get("sheet"), "cell": o.get("cell")},
            "raw_state": ({"is_formula": 原.get("is_formula"), "raw_blank": 原.get("raw_blank"),
                           "raw_value": 原.get("raw_value"), "cached_value": 原.get("cached_value")}
                          if 原 else "未回读到原映射"),
            "证据引用": [f"五样式索引:{o['mapping_id']}", f"原映射:{o['mapping_file']}#{o['mapping_id']}"],
            "确认依据": {"方法": "五样式人工逐格映射"},
            "问题引用": (["样式2应收票据raw_blank历史标志误为false，原表实际为空"]
                       if o["mapping_file"].endswith("004-应收票据-样式2-人工映射.jsonl")
                       and o.get("cell") in ("C47", "D47", "C48", "D48") else []),
        })

    for f in A映射["facts"]:
        指标ID = f["metric_id"]
        方法 = A确认方法.get(f.get("recognition_provenance", {}).get("method"), "未知方法")
        规范值 = copy.deepcopy(f.get("dimensions", {}))
        # 只应用本来源已核实的明文期间说明；不猜日期，不改变原period结构。
        期间说明 = 规则.get("来源期间对应", {})
        if 期间说明.get("source_path") == A映射.get("source_path"):
            位置 = f.get("source_reference", {}).get("location", f["id"])
            匹配 = re.fullmatch(r"(.+)!([A-Z]+)([0-9]+)", 位置)
            上下文 = [c for c in A映射.get("source_context", []) if 匹配 and isinstance(c, dict)
                      and c.get("sheet") == 匹配[1]
                      and c.get("first_row", 0) <= int(匹配[3]) < c.get("first_row", 0) + c.get("row_count", 0)
                      and 期间说明.get("依据原文") in c.get("chapter_context", [])]
            if len(上下文) == 1:
                for 对 in 期间说明.get("对应", []):
                    if f.get("dimensions", {}).get("period") == 对["period"]:
                        规范值["period_role"] = 对["period_role"]
        挂表现(指标ID, {
            "record_type": "representation",
            "representation_id": "afact:" + 内容哈希(
                "a_fact", f.get("source_reference", {}).get("file_hash"),
                f.get("source_reference", {}).get("location"), 指标ID),
            "source_kind": "a_fact",
            "table_form": {"metric_definition": 语义们[指标ID]["meaning"]},
            "dimension_values": f.get("dimensions", {}), "规范维度值": 规范值, "占位与缺失": [],
            "dimension_evidence": f.get("dimension_evidence", {}),
            "source_locator": {"file": A映射.get("source_path"), "sha256": A源哈希,
                               "sheet_cell": f["id"],
                               "file_hash": f.get("source_reference", {}).get("file_hash"),
                               "location": f.get("source_reference", {}).get("location")},
            "raw_state": f.get("raw_value_state"),
            "证据引用": [f"A端映射:{f['id']}", f"确认方法:{f.get('recognition_provenance', {}).get('method')}"],
            "确认依据": {"方法": 方法, "原确认记录": f.get("recognition_provenance"),
                       "owner_reference": f.get("owner_reference")},
            "问题引用": [],
        })

    # ---------- 维度字典（含义、值类型、名称与取值对应按规则文件逐名归一） ----------
    维度按id = {}
    维度缺口 = []

    def 加维度(维度ID, 含义, 值类型, 来源, 原名条目, 对应v2=None, 对应依据=None,
              规范值=None, 取值表达要求=None, 说明=None):
        if 维度ID in 维度按id:
            维度按id[维度ID]["原字段名称及对应依据"].extend(原名条目)
            return
        条目 = {"dimension_id": 维度ID, "含义": 含义, "值类型": 值类型,
                "definition_source": 来源, "原字段名称及对应依据": 原名条目,
                "对应v2维度": 对应v2, "对应依据": 对应依据}
        if 规范值:
            条目["规范值"] = 规范值
        if 取值表达要求:
            条目["取值表达要求"] = 取值表达要求
        取别 = 取值别名表.get(维度ID)
        if 取别:
            条目["取值对应"] = 取别
        if 说明:
            条目["说明"] = 说明
        维度按id[维度ID] = 条目

    for 维度ID, d in sorted(v2维度.items()):
        加维度(维度ID, d.get("meaning"), d.get("value_type"), "v2",
              [{"来源": "v2", "名称": 维度ID, "类型": "同名"}], 对应v2=维度ID,
              规范值=(规范维度.get(维度ID) or {}).get("规范值"),
              取值表达要求=d.get("domain"))
    v1维度观察 = defaultdict(list)
    for 旧ID, r in v1.items():
        if r["status"] != "active":
            continue
        for d in r.get("dimensions", []):
            v1维度观察[d["name"]].append(d)
    for 名, 声明们 in sorted(v1维度观察.items()):
        规 = 规范维度.get(名)
        if 规 is None:
            加维度(名, None, None, "v1",
                  [{"来源": "v1", "名称": 名, "类型": "同名",
                    "依据": f"active槽位声明{len(声明们)}次", "声明样例": 声明们[:3]}])
            维度缺口.append({"名称": 名, "来源": "v1声明", "原因": "规则文件未收录，含义未归纳",
                          "依据": f"active槽位声明{len(声明们)}次"})
            continue
        关系 = v1v2关系.get(名, {})
        对应v2 = 关系.get("对应v2维度") or (名 if 名 in v2维度 else None)
        加维度(名, 规.get("含义"), 规.get("值类型"), "v1",
              [{"来源": "v1", "名称": 名, "类型": "同名",
                "依据": f"active槽位声明{len(声明们)}次", "声明样例": 声明们[:3]}],
              对应v2=对应v2, 对应依据=关系.get("依据") or ("同名同含义" if 对应v2 else None),
              规范值=规.get("规范值"), 取值表达要求=规.get("取值表达要求"))
    五样式维度观察 = defaultdict(Counter)
    for e in 索引["examples"]:
        for k, v in e.get("legacy_dimensions", {}).items():
            if k not in 追踪字段 and not (isinstance(v, str) and 占位.match(v)):
                五样式维度观察[k][str(v)[:40]] += 1
    for 名, 值计数 in sorted(五样式维度观察.items()):
        对 = 原名对应.get(名)
        原名条目 = {"来源": "五样式", "名称": 名,
                 "依据": f"观察{sum(值计数.values())}格",
                 "取值样例": dict(值计数.most_common(10))}
        if 对 is None:
            if 名 in 规范维度:
                原名条目["类型"] = "同名"
                原名条目["对应依据"] = "与规范维度同名，含义按v1归纳"
                if 名 in 维度按id:
                    加维度(名, None, None, None, [原名条目])
                else:
                    规 = 规范维度[名]
                    加维度(名, 规.get("含义"), 规.get("值类型"), "五样式观察", [原名条目],
                          对应v2=名 if 名 in v2维度 else None,
                          规范值=规.get("规范值"), 取值表达要求=规.get("取值表达要求"))
                continue
            加维度(名, None, None, "五样式观察", [原名条目])
            维度缺口.append({"名称": 名, "来源": "五样式观察", "原因": "规则文件未收录，含义未归纳",
                          "取值样例": dict(值计数.most_common(5))})
            continue
        规范名 = 对["规范维度"]
        if isinstance(规范名, list):
            维度缺口.append({"名称": 名, "来源": "五样式观察", "原因": "名称有歧义，候选：" + "、".join(规范名)})
            continue
        原名条目["类型"] = 对["类型"]
        原名条目["对应依据"] = 对["依据"]
        if 规范名 in 维度按id:
            加维度(规范名, None, None, None, [原名条目])
        else:
            规 = 规范维度.get(规范名) or {}
            加维度(规范名, 规.get("含义"), 规.get("值类型"), "五样式观察", [原名条目],
                  对应v2=规范名 if 规范名 in v2维度 else None,
                  规范值=规.get("规范值"), 取值表达要求=规.get("取值表达要求"))
    维度字典 = [维度按id[k] for k in sorted(维度按id)]

    # ---------- 问题清单 ----------
    差异调查 = []
    for 条目 in 旧新:
        if 条目.get("版本") == "历史v1" and not 条目.get("与当前定义一致") \
                and (条目["本次规范对应"] or {}).get("目标semantic_id"):
            差异调查.append({"旧ID": 条目["legacy_id"], "差异": 条目["本次核验结论"],
                            "适用条件": 条目["本次规范对应"].get("适用条件"),
                            "处理": 条目["本次规范对应"].get("处理说明"),
                            "出处": 条目["旧定义引用"]})
    B检查点 = os.path.join(项目, "测试结果", "真实语义识别续跑_B_20260922_071726_598567")
    问题单 = {
        "原本尚未完成的识别": {
            "说明": "这些是原本未完成的识别工作或验证材料，本次不补识别，也不作为五样式成果的缺口",
            "B端未决": {"数量": 25425, "位置": B检查点, "依据": "最新检查点00094；B端来自五样式之一，不重复收录"},
            "五样式新系统核验": ["五样式固定资产首批语义：未通过（维度未确认）",
                              "五样式收入成本语义：指标通过、维度不全"],
        },
        "已确认的上下文资料": {"A端排除": {"数量": len(A映射.get("excluded", [])),
                        "位置": A路径, "去向": "数据/上下文证据.json：A端排除项"}},
        "本次整理未解决": {
            "哈希差异调查结论": 差异调查,
            "五样式原映射未匹配": 未匹配原映射,
            "维度缺口": 维度缺口,
        },
        "证据合并": {
            "说明": "完全相同的表现合并引用，不重复建条；证据增加不改变表现身份",
            "重复证据合并": 重复证据,
        },
        "无法读取或核验的来源": {
            "识别快照损坏": {"路径": os.path.join(项目, "金标准", "识别快照", "6bbf75f7b8d8904712caf7d68505fe7efee321492faa74cfd3d1af49927feab0.jsonl"),
                          "问题": "首行无法解析，属中间产物，不影响正式来源"},
            "样式2应收票据四格": {"格子": ["C47", "D47", "C48", "D48"],
                             "问题": "raw_blank历史标志误为false，原表实际为空；语义和目标ID不受影响",
                             "出处": "审计数据自动搬运系统-语义金标准与附注样式1至5映射完整说明.md第4.11节"},
        },
        "历史转换状态": {
            "说明": "仅作追溯，不作为旧成果降级理由",
            "approved记录": sum(1 for r in 转换记录们 if r.get("status") == "approved"),
            "unresolved记录": sum(1 for r in 转换记录们 if r.get("status") == "unresolved"),
            "approved绑定旧原库哈希e79babc0": sum(1 for r in 转换记录们
                                          if r.get("status") == "approved" and r["source_gold_hash"].startswith("e79babc0")),
            "approved绑定当前原库哈希5edec086": sum(1 for r in 转换记录们
                                           if r.get("status") == "approved" and r["source_gold_hash"].startswith("5edec086")),
        },
    }

    # ---------- 来源示例整理去向（任务1：逐条登记，不只另存一份） ----------
    表现按旧ID = defaultdict(list)
    for p in 表现们:
        if p.get("定义旧ID"):
            表现按旧ID[p["定义旧ID"]].append(p["representation_id"])
    引用反查 = defaultdict(list)
    for c in 转换记录们:
        for ref in c.get("evidence_refs") or []:
            引用反查[str(ref)].append({"类型": "转换", "id": c.get("id")})
    当前对应按旧 = {条目["legacy_id"]: 条目["本次规范对应"]
                 for 条目 in 旧新 if 条目.get("版本") == "当前v1"}
    来源示例去向 = []
    按版本定义 = {(x["源标准哈希"], x["legacy_id"], x["旧定义哈希"]): x for x in 旧新}
    for e in v2示例们:
        o = e.get("origin", {})
        旧ID = o.get("legacy_id")
        解析条目 = 示例解析[e["id"]]
        原定义哈希值 = next((x["原定义哈希"] for x in 解析条目 if x.get("原定义哈希")), None)
        if 旧ID:
            源版本 = o.get("source_gold_hash") or o.get("file_hash")
            定义指纹 = o.get("legacy_definition_hash") or 原定义哈希值
            条 = 按版本定义.get((源版本, 旧ID, 定义指纹))
            assert 条 is not None, f"来源示例无法关联版本定义: {e['id']}"
            对应 = 条.get("本次规范对应") or {}
            语义集 = set()
            当前可用 = 旧ID in 语义们 and (条["版本"] == "当前v1" or 条.get("与当前定义一致"))
            if 当前可用:
                语义集.add(旧ID)
            if 对应.get("目标semantic_id") in 语义们:
                语义集.add(对应["目标semantic_id"])
            去向 = {"类型": "旧定义证据",
                    "带版本旧定义": {"源标准哈希": 源版本, "旧ID": 旧ID, "定义哈希": 定义指纹},
                    "对应引用": 条["对应ID"], "关联规范语义": sorted(语义集),
                    "表现": 表现按旧ID.get(旧ID, []) if 当前可用 else [],
                    "被引用": 引用反查.get(e["id"], []),
                    "说明": "历史定义证据按原版本保存；只关联有版本依据的现行语义及表现"}
        else:
            去向 = {"类型": "审查说明证据",
                    "说明": "不含业务表现的审查说明，按定义证据保留，不虚构单元格",
                    "被引用": 引用反查.get(e["id"], [])}
        来源示例去向.append({
            "record_type": "来源示例去向", "示例ID": e["id"],
            "来源标准哈希": o.get("source_gold_hash") or o.get("file_hash"),
            "origin": o, "review_refs": e.get("review_refs"),
            "text_evidence原文": e.get("text_evidence"),
            "解析": {"条目": 解析条目, "原定义哈希": 原定义哈希值},
            "去向": 去向})

    # ---------- 上下文证据（已确认排除项逐条保留，不笼统丢掉） ----------
    上下文证据 = {
        "说明": "已确认的排除/非业务格作为上下文证据逐条保留；其他样式的排除数量取自完整说明，未逐格保存",
        "A端排除项": A映射.get("excluded", []),
        "A端来源上下文": A映射.get("source_context"),
        "A端补充上下文": A映射.get("additional_context"),
        "A端归属记录": A映射.get("owners"),
        "样式2结构格排除记录": 排除证据,
        "其他样式排除计数": {"样式1非业务格": 3119, "样式3排除": 258,
                          "样式4可见非空结构或检查格": 5997, "样式5可见非空结构或检查格": 8539},
    }

    # ---------- 旧表现ID一次性迁移（--已有成果；此后重跑不得再换ID） ----------
    迁移 = []
    if 已有成果路径:
        旧表现路径 = os.path.join(已有成果路径, "数据", "表现记录.jsonl")
        if os.path.exists(旧表现路径):
            新按位置 = defaultdict(list)
            for p in 表现们:
                新按位置[(稳定摘要(来源位置(p)), p["semantic_id"])].append(p["representation_id"])
            沿革路径 = os.path.join(已有成果路径, "对应", "表现身份迁移.jsonl")
            旧资料 = 读jsonl(旧表现路径)
            if os.path.exists(沿革路径):
                for 旧迁移 in 读jsonl(沿革路径):
                    if 旧迁移.get("旧记录"):
                        旧资料.append(旧迁移["旧记录"])
            已迁 = set()
            for op in 旧资料:
                if op["representation_id"] in 已迁:
                    continue
                已迁.add(op["representation_id"])
                候选 = 新按位置.get((稳定摘要(来源位置(op)), op["semantic_id"]), [])
                if len(候选) != 1:
                    raise SystemExit(f"旧表现迁移无法唯一定位：{op['representation_id']}，候选{len(候选)}")
                if op["representation_id"] == 候选[0]:
                    continue
                迁移.append({"旧表现ID": op["representation_id"], "新表现ID": 候选[0],
                            "核对": "逻辑来源、位置及原语义一致；完整结构维度见新记录",
                            "旧记录": op})

            # 旧成果本身还有更早的迁移记录时，沿链更新到本批ID，不遗失早期别名。
            到本批 = {x["旧表现ID"]: x["新表现ID"] for x in 迁移}
            到本批.update({p["representation_id"]: p["representation_id"] for p in 表现们})
            if os.path.exists(沿革路径):
                for old in 读jsonl(沿革路径):
                    if old["旧表现ID"] in 到本批:
                        continue
                    new = 到本批.get(old["新表现ID"])
                    if new is None:
                        raise SystemExit("更早的表现迁移链无法承接：" + old["旧表现ID"])
                    迁移.append({"旧表现ID": old["旧表现ID"], "新表现ID": new,
                                 "沿革来源记录": old, "核对":"从既有迁移记录沿链承接，未重造语义"})
                    到本批[old["旧表现ID"]] = new

    # ---------- 写出前重核输入哈希（读取后输入变化则本批不发布） ----------
    for p, 期望 in 清单.items():
        if sha256(p) != 期望:
            raise SystemExit(f"读取后输入发生变化，本批不发布，保留诊断：{p}")

    # ---------- 写出 ----------
    规则输出 = os.path.join(输出根, "对应", "维度规范对应.json")
    os.makedirs(os.path.dirname(规则输出), exist_ok=True)
    if os.path.abspath(规则输出) != os.path.abspath(规则路径):
        with open(规则输出, "wb") as f:
            f.write(open(规则路径, "rb").read())
    写jsonl(os.path.join(输出根, "数据", "统一语义记录.jsonl"), 语义们.values())
    写jsonl(os.path.join(输出根, "数据", "表现记录.jsonl"), 表现们)
    写jsonl(os.path.join(输出根, "数据", "维度字典.jsonl"), 维度字典)
    写jsonl(os.path.join(输出根, "对应", "旧新对应.jsonl"), 旧新)
    写jsonl(os.path.join(输出根, "对应", "来源示例.jsonl"), 来源示例去向)
    if 迁移:
        写jsonl(os.path.join(输出根, "对应", "表现身份迁移.jsonl"), 迁移)
    os.makedirs(os.path.join(输出根, "问题"), exist_ok=True)
    with open(os.path.join(输出根, "问题", "问题清单.json"), "w", encoding="utf-8-sig", newline="\n") as f:
        json.dump(问题单, f, ensure_ascii=False, indent=1, sort_keys=True)
    with open(os.path.join(输出根, "数据", "上下文证据.json"), "w", encoding="utf-8-sig", newline="\n") as f:
        json.dump(上下文证据, f, ensure_ascii=False, indent=1, sort_keys=True)

    print(f"语义记录 {len(语义们)}（active旧定义 {sum(1 for r in 语义们.values() if r['semantic_kind']=='legacy_slot')}，指标 {len(v2指标)}）")
    print(f"表现记录 {len(表现们)}（v1来源 {sum(1 for r in 表现们 if r['source_kind']=='v1_source')}，"
          f"五样式 {sum(1 for r in 表现们 if r['source_kind']=='five_style')}，"
          f"A {sum(1 for r in 表现们 if r['source_kind']=='a_fact')}）")
    已建立 = sum(1 for r in 旧新 if (r.get("本次规范对应") or {}).get("目标semantic_id"))
    print(f"维度字典 {len(维度字典)}；旧新对应 {len(旧新)}（已建立对应 {已建立}）；"
          f"来源示例去向 {len(来源示例去向)}；排除证据 {len(排除证据)}；未匹配原映射 {len(未匹配原映射)}")


if __name__ == "__main__":
    main()
