# -*- coding: utf-8 -*-
"""统一整理的行为检查（按2026-09-25修补计划第二节）。

用法：python 测试整理.py [输出根] [--输入清单 资产清单.json]

期望值手工取自原记录或按修补计划的业务预期写明，不用被测代码计算。
检查累积执行，末尾统一报告并以退出码区分通过/失败。
"""
import hashlib
import importlib.util
import json
import os
import re
import sys
from collections import Counter

项目 = r"C:\Users\27651\Desktop\附注自动更新系统"


def 解析参数(argv):
    位置, 输入清单 = [], None
    i = 0
    while i < len(argv):
        if argv[i] == "--输入清单" and i + 1 < len(argv):
            输入清单 = argv[i + 1]; i += 2
        else:
            位置.append(argv[i]); i += 1
    return (位置[0] if 位置 else os.path.join(项目, "语义成果整理")), 输入清单


根, 输入清单路径 = 解析参数(sys.argv[1:])


def 读jsonl(路径):
    with open(路径, encoding="utf-8-sig") as f:
        return [json.loads(行) for 行 in f if 行.strip()]


def 定义哈希(r):
    return hashlib.sha256(json.dumps(r, sort_keys=True, ensure_ascii=False,
                                     separators=(",", ":")).encode()).hexdigest()


def 载入查询模块():
    路径 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "建立索引.py")
    spec = importlib.util.spec_from_file_location("建立索引", 路径)
    模块 = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(模块)
    return 模块


结果们 = []


def 检查(名, 条件, 详情=""):
    结果们.append({"名称": 名, "通过": bool(条件), "详情": str(详情)[:600]})
    print(("PASS " if 条件 else "FAIL ") + 名 + (" | " + str(详情)[:250] if 详情 else ""))


def main():
    global 结果们
    # ---------- 输入（按固定清单，否则默认路径） ----------
    if 输入清单路径:
        固 = json.load(open(输入清单路径, encoding="utf-8-sig"))["固定输入"]
        V1路径 = 固["v1"]["路径"]; V2路径 = 固["v2当前版本"]["路径"]
        索引路径 = 固["五样式索引"]["路径"]; A路径 = 固["A端映射"]["路径"]
    else:
        V1路径 = os.path.join(项目, "金标准", "财务报表附注语义金标准-v1.jsonl")
        V2路径 = os.path.join(项目, "金标准", "版本",
                          "财务报表附注语义金标准_20260922_132731_236607_5547978f.jsonl")
        索引路径 = os.path.join(项目, "金标准", "五种样式人工映射", "语义示例索引.json")
        A路径 = os.path.join(项目, "测试结果", "真实语义识别续跑_A_20260922_050354_458720",
                         "导入外部复核_20260922_051437_942503_b323c5", "附注表格语义映射.json")
    v1 = {r["id"]: r for r in 读jsonl(V1路径)}
    v2转换, v2指标, v2维度, v2示例 = [], {}, {}, {}
    for r in 读jsonl(V2路径):
        if r["record_type"] == "legacy_conversion":
            v2转换.append(r)
        elif r["record_type"] == "metric":
            v2指标[r["id"]] = r
        elif r["record_type"] == "dimension":
            v2维度[r["id"]] = r
        elif r["record_type"] == "source_example":
            v2示例[r["id"]] = r
    五样式索引 = json.load(open(索引路径, encoding="utf-8-sig"))
    A映射 = json.load(open(A路径, encoding="utf-8-sig"))

    # ---------- 整理输出 ----------
    语义 = 读jsonl(os.path.join(根, "数据", "统一语义记录.jsonl"))
    表现 = 读jsonl(os.path.join(根, "数据", "表现记录.jsonl"))
    维字典 = 读jsonl(os.path.join(根, "数据", "维度字典.jsonl"))
    旧新 = 读jsonl(os.path.join(根, "对应", "旧新对应.jsonl"))
    问题单 = json.load(open(os.path.join(根, "问题", "问题清单.json"), encoding="utf-8-sig"))
    索引文件 = json.load(open(os.path.join(根, "索引", "查找索引.json"), encoding="utf-8-sig"))
    按id = {r["semantic_id"]: r for r in 语义}
    旧新按id = {}
    for r in 旧新:
        旧新按id.setdefault(r["legacy_id"], []).append(r)
    查询 = 载入查询模块()

    # ========== 一、保留的有效检查（输入覆盖与身份区分） ==========
    active = {i for i, r in v1.items() if r["status"] == "active"}
    检查("v1全部active旧定义有语义记录", active <= set(按id), f"active {len(active)}")
    原来源 = {(r["id"], 定义哈希(e)) for r in v1.values() if r["status"]=="active" for e in r.get("sources",[])}
    成果来源 = {(p.get("定义旧ID"), 定义哈希(e)) for p in 表现 if p["source_kind"]=="v1_source"
                for e in p.get("来源原文证据",[])}
    检查("v1全部来源对象成为表现", 原来源==成果来源)
    检查("五样式表现全数整理",
         len([p for p in 表现 if p["source_kind"] == "five_style"]) == len(五样式索引["examples"]))
    A表现 = [p for p in 表现 if p["source_kind"] == "a_fact"]
    检查("A端表现全数收录", len(A表现) == len(A映射["facts"]))
    方法分布 = Counter(p["确认依据"]["方法"] for p in A表现)
    检查("A端确认方法分布为519/282/23",
         方法分布 == {"模型双轮识别": 519, "模型双轮识别＋同义复核": 282, "外部复核": 23}, dict(方法分布))
    检查("115个指标全部成记录", set(v2指标) <= set(按id), f"{len(v2指标)}个")
    检查("无b_fact表现（B端不重复收录）", not any(p["source_kind"] == "b_fact" for p in 表现))
    检查("每条表现都有有效语义", all(p["semantic_id"] in 按id for p in 表现))
    停用 = [r for r in 语义 if r["semantic_id"] in (
        "metric.net_profit_reconciliation_adjustment.amount",
        "metric.deferred_tax_liability.net_increase", "metric.operating_payable.net_increase")]
    检查("3个停用指标保留且不复活", len(停用) == 3
         and all(r["lifecycle"]["状态"] == "retired" for r in 停用))
    占位 = re.compile(r"^<.+>$")
    检查("模板占位未登记为真实取值",
         all(not (isinstance(v, str) and 占位.match(v))
             for p in 表现 for v in p.get("dimension_values", {}).values()))
    检查("语义记录ID无重复", len(按id) == len(语义))
    检查("表现记录ID无重复", len({p["representation_id"] for p in 表现}) == len(表现))

    # ========== 二、真实例子的业务含义（修正张冠李戴后的类别×期间） ==========
    # 样式1货币资金：S001=库存现金/期末，S002=库存现金/期初，S003=银行存款/期末（手工核对v1原记录）
    预期三槽 = {"C-N087-T001-S001": ("库存现金", "closing"),
              "C-N087-T001-S002": ("库存现金", "opening"),
              "C-N087-T001-S003": ("银行存款", "closing")}
    槽问题 = []
    for 旧ID, (类别, 期间) in 预期三槽.items():
        r = 按id.get(旧ID)
        if r is None:
            槽问题.append(f"{旧ID}缺记录"); continue
        m = r["meaning"]
        if 类别 not in json.dumps(m, ensure_ascii=False):
            槽问题.append(f"{旧ID}类别不是{类别}")
        if m.get("期间角色") != 期间:
            槽问题.append(f"{旧ID}期间类型={m.get('期间角色')}≠{期间}")
    检查("样式1现金S001库存现金期末/S002库存现金期初/S003银行存款期末", not 槽问题, "；".join(槽问题))

    # 受限资产：同一语义下按受限类型区分，维度含义完整
    受限表现 = [r for r in 表现 if r["source_kind"] == "five_style"
               and r["source_locator"].get("mapping_id") in ("S1-REVIEW-0011", "S1-REVIEW-0013", "S1-REVIEW-0015")]
    取值们 = {r["dimension_values"].get("restriction_type") for r in 受限表现}
    检查("受限资产保证金实例按受限类型区分",
         {"银行承兑汇票保证金", "履约保证金", "保函保证金"} <= 取值们
         and all(r["semantic_id"] == "C-N057-T003-S019" for r in 受限表现), str(取值们))
    规范取值们 = {r.get("规范维度值", {}).get("restriction_type") for r in 受限表现}
    检查("受限类型进入规范维度值可统一查询",
         {"银行承兑汇票保证金", "履约保证金", "保函保证金"} <= 规范取值们, str(规范取值们))
    检查("维度取值未误建为指标", not any("银行承兑汇票保证金" in sid for sid in 按id))
    维按id = {d["dimension_id"]: d for d in 维字典}
    rt = 维按id.get("restriction_type", {})
    检查("restriction_type维度含义与类型完整", bool(rt.get("含义")) and bool(rt.get("值类型")),
         f"含义={rt.get('含义')}")
    # 合计保留汇总范围；开放项目保留dimension角色
    合计007 = 按id.get("C-N087-T001-S007", {})
    检查("C-N087-T001-S007保留汇总角色与来源特定组成",
         合计007.get("meaning", {}).get("汇总角色") == "total"
         and 合计007.get("meaning", {}).get("汇总关系", {}).get("component_variants") is not None)
    检查("C-N009-T002-S001保留dimension角色",
         按id.get("C-N009-T002-S001", {}).get("meaning", {}).get("槽位角色") == "dimension")

    # ========== 三、跨来源查询接通（复核问题1） ==========
    mf = 查询.查询语义(索引文件, "metric.monetary_funds.balance")
    表现按pid = {p["representation_id"]: p for p in 表现}
    mf来源 = Counter(表现按pid[i]["source_kind"] for i in mf if i in 表现按pid)
    S001来源数 = len(v1["C-N087-T001-S001"].get("sources", []))
    S001五样式数 = sum(1 for e in 五样式索引["examples"] if e["legacy_slot_id"] == "C-N087-T001-S001")
    检查("查询货币资金余额指标汇集v1来源、五样式与A端三来源",
         mf来源.get("v1_source", 0) >= S001来源数
         and mf来源.get("five_style", 0) >= S001五样式数
         and mf来源.get("a_fact", 0) >= 26,
         f"实得 {dict(mf来源)}；应至少 v1来源{S001来源数}(S001)+五样式{S001五样式数}(S001)+A端26")
    S001表现数 = sum(1 for p in 表现 if p.get("定义旧ID") == "C-N087-T001-S001")
    检查("S001的9条v1来源与5条五样式全部在指标查询结果中",
         S001表现数 >= S001来源数 + S001五样式数
         and all(pid in mf for pid, p in 表现按pid.items()
                 if p.get("定义旧ID") == "C-N087-T001-S001"),
         f"S001表现登记{S001表现数}条（应≥{S001来源数 + S001五样式数}）")
    # 具体类别+期间筛选：库存现金/期末不混入银行存款或期初
    # （T002-S015经核实为"类别明细"表的库存现金/期末余额，与S001同含义不同表结构，属于合法命中）
    筛 = 查询.查询语义(索引文件, "metric.monetary_funds.balance",
                     {"cash_selection": "库存现金", "period_role": "closing"})
    筛旧id = {表现按pid[i].get("定义旧ID") for i in 筛 if i in 表现按pid}
    检查("指标下筛选库存现金与期末：不混入银行存款或库存现金期初",
         len(筛) >= S001来源数 + S001五样式数
         and 筛旧id <= {"C-N087-T001-S001", "C-N087-T002-S015", None}
         and not ({"C-N087-T001-S002", "C-N087-T001-S003"} & 筛旧id),
         f"命中{len(筛)}条，涉及旧ID {sorted(x for x in 筛旧id if x)}")
    # 旧ID反查不扩大范围
    s001表现 = set(索引文件.get("by_semantic", {}).get("C-N087-T001-S001", {}).get("表现", []))
    s001旧集 = {表现按pid[i].get("定义旧ID") for i in s001表现 if i in 表现按pid}
    检查("旧ID查询不扩大为全部货币资金",
         len(s001表现) > 0 and s001旧集 <= {"C-N087-T001-S001", None},
         f"{len(s001表现)}条")

    # ========== 四、期间异写归一（复核问题2） ==========
    q1 = 查询.查询维度(索引文件, "期间角色", "期末")
    q2 = 查询.查询维度(索引文件, "period_role", "closing")
    q3 = 查询.查询维度(索引文件, "period_role", "opening")
    检查("期间角色/期末 与 period_role/closing 查询同集合",
         isinstance(q1, set) and isinstance(q2, set) and q1 == q2 and len(q2) > 0,
         f"{len(q1) if isinstance(q1, set) else q1} vs {len(q2) if isinstance(q2, set) else q2}")
    检查("期初不混入期末查询", isinstance(q3, set) and not (q2 & q3) and len(q3) > 0)
    检查("期间角色不再作为独立维度ID（归一为period_role别名）",
         "期间角色" not in 维按id and "period_role" in 维按id)
    pr = 维按id.get("period_role", {})
    检查("period_role规范值closing/opening/current/prior齐备",
         {"closing", "opening", "current", "prior"} <= set(pr.get("规范值", [])),
         str(pr.get("规范值")))

    # ========== 五、来源示例整理去向（复核问题3） ==========
    示例去向路径 = os.path.join(根, "对应", "来源示例.jsonl")
    去向 = 读jsonl(示例去向路径) if os.path.exists(示例去向路径) else []
    去向按id = {r.get("示例ID"): r for r in 去向}
    检查("全部来源示例逐条登记整理去向",
         len(去向) == len(v2示例) and all(eid in 去向按id for eid in v2示例),
         f"已登记 {len(去向)} / 输入 {len(v2示例)}")
    引用示例 = {ref for r in v2转换 for ref in (r.get("evidence_refs") or [])
              if str(ref).startswith("source:")}
    检查("转换的来源示例引用都能解析到去向记录", 引用示例 <= set(去向按id),
         f"引用{len(引用示例)}，未解析{len(引用示例 - set(去向按id))}")
    # 内嵌定义逐条解析（不假定一示例一旧ID）
    内嵌 = [e for e in v2示例.values() if e["origin"].get("legacy_id")]
    坏解析 = [r["示例ID"] for r in 去向
             if r["示例ID"] in {e["id"] for e in 内嵌} and not r.get("解析", {}).get("原定义哈希")]
    检查("内嵌定义的示例都有解析结果与原定义哈希",
         去向 and len([r for r in 去向 if r["示例ID"] in {e["id"] for e in 内嵌}]) == len(内嵌)
         and not 坏解析, f"缺解析{len(坏解析)}条")

    # ========== 六、版本对应（复核问题4） ==========
    # 双版本旧ID按版本分别成条
    按旧版本 = {}
    for r in v2转换:
        if r["status"] == "approved":
            按旧版本.setdefault(r["legacy_id"], set()).add(r["source_gold_hash"])
    双版本 = {k for k, v in 按旧版本.items() if len(v) > 1}
    错分条 = [k for k in 双版本
             if len({(r.get("源标准哈希"), r.get("旧定义哈希")) for r in 旧新按id.get(k, [])}) < 2]
    检查("双版本旧ID按版本分别成条", not 错分条, f"双版本{len(双版本)}个，未分条{len(错分条)}个")
    # 当前版本条目定义哈希=当前重算哈希（全量）
    混配 = [r["legacy_id"] for r in 旧新
           if r.get("源标准哈希", "").startswith("5edec086") and r["legacy_id"] in v1
           and r.get("旧定义哈希") and r["旧定义哈希"] != 定义哈希(v1[r["legacy_id"]])]
    检查("当前版本引用与重算哈希全量一致", not 混配, f"混配{len(混配)}条，原958条口径现逐条核验")
    # 历史版本反查
    历史键 = [k for k in 索引文件.get("by_legacy_versioned", {}) if k.startswith("e79babc0")]
    检查("按历史版本哈希+旧ID能反查定义与转换", len(历史键) > 0,
         f"历史版本键{len(历史键)}个")

    # ========== 七、8条哈希差异逐项处理（不预定处理数量） ==========
    差异8 = ["C-N019-T004-S015", "C-N035-T001-S003", "C-N035-T001-S004", "C-N035-T001-S036",
            "C-N103-T001-S003", "C-N103-T001-S004", "C-N103-T001-S019", "C-N103-T001-S020"]
    未处理 = []
    for 旧ID in 差异8:
        条目 = 旧新按id.get(旧ID, [])
        结论们 = [r.get("本次核验结论") or r.get("本次整理", {}).get("处理结果", "") for r in 条目]
        if not any(c and "需复核" not in c for c in 结论们):
            未处理.append(旧ID)
    检查("8条哈希差异逐项有差异说明与处理结论", not 未处理, f"仍笼统待复核：{未处理}")
    s036 = [r for r in 旧新按id.get("C-N035-T001-S036", []) if r.get("源标准哈希", "").startswith("e79babc0")]
    检查("S036历史版本条目带条件对应（目标与适用条件齐备）",
         bool(s036) and bool((s036[0].get("本次规范对应") or {}).get("目标semantic_id"))
         and bool((s036[0].get("本次规范对应") or {}).get("适用条件")),
         "")

    # ========== 八、规范含义与维度（复核问题2） ==========
    八字段 = ["业务对象", "指标含义", "槽位角色", "值类型", "计量基础", "期间类型", "适用范围", "业务限制"]
    缺字段 = [r["semantic_id"] for r in 语义
             if sum(1 for k in 八字段 if k in r["meaning"]) < 8 and r["semantic_kind"] in ("legacy_slot", "metric")]
    检查("legacy与metric记录使用同一套meaning共同字段", not 缺字段,
         f"缺字段记录{len(缺字段)}条，示例{缺字段[:3]}")
    空含义 = [d["dimension_id"] for d in 维字典 if not d.get("含义")]
    缺口登记 = (问题单.get("本次整理未解决", {}) or {}).get("维度缺口", [])
    检查("维度空含义逐条解决或明确登记缺口",
         not 空含义 and not 缺口登记,
         f"空含义{len(空含义)}条，缺口登记{len(缺口登记)}条")
    # 表现关联实际采用的版本对应
    有对应旧 = {r["legacy_id"] for r in 旧新
              if (r.get("本次规范对应") or {}).get("目标semantic_id")
              or r.get("本次整理", {}).get("目标semantic_id")}
    应带 = [p for p in 表现 if p["source_kind"] in ("v1_source", "five_style")
            and p.get("定义旧ID") in 有对应旧]
    缺引用 = [p for p in 应带 if not p.get("对应引用")]
    检查("表现关联实际采用的版本对应", len(应带) > 0 and not 缺引用,
         f"应带{len(应带)}，缺引用{len(缺引用)}")

    # ========== 九、上下文证据（A端538条排除全分类保留） ==========
    上下文 = json.load(open(os.path.join(根, "数据", "上下文证据.json"), encoding="utf-8-sig"))
    A排除 = 上下文.get("A端排除项", [])
    A类 = Counter(e.get("category") for e in A排除)
    检查("A端538条排除逐条保留分类、位置与理由",
         len(A排除) == len(A映射.get("excluded", []))
         and all(e.get("reason") and e.get("cell") for e in A排除),
         f"保留{len(A排除)}/{len(A映射.get('excluded', []))}，分类{dict(A类)}")

    # ========== 十、表现身份不含整文件哈希（复核：增量换ID） ==========
    禁含 = ["5edec0861b757b0a", "e79babc0af823d82"]
    A源哈希 = (A映射.get("source_hash") or "")[:16]
    if A源哈希:
        禁含.append(A源哈希)
    带文件哈希 = [p["representation_id"] for p in 表现
                if any(h in p["representation_id"] for h in 禁含)]
    检查("表现ID不含整份金标准或来源文件哈希", not 带文件哈希,
         f"违规{len(带文件哈希)}条，示例{带文件哈希[:2]}")
    顺序号 = [p["representation_id"] for p in 表现 if re.search(r"#\d+$", p["representation_id"])]
    检查("表现ID不用出现顺序编号", not 顺序号, f"{len(顺序号)}条")

    from 二次复核回归 import 检查集
    for r in 检查集(根, 输入清单路径):
        检查("二次复核:" + r["名称"], r["通过"], r["详情"])

    # ========== 汇总 ==========
    失败 = [r for r in 结果们 if not r["通过"]]
    print(f"\n检查 {len(结果们)} 项，通过 {len(结果们) - len(失败)}，失败 {len(失败)}")
    for r in 失败:
        print("  未通过：" + r["名称"])
    sys.exit(1 if 失败 else 0)


if __name__ == "__main__":
    main()
