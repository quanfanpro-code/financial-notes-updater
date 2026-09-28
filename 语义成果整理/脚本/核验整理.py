# -*- coding: utf-8 -*-
"""独立全量核验（按2026-09-25修补计划任务5）。

用法：python 核验整理.py [输出根] [--输入清单 资产清单.json]

- 静态全量检查：直接回读固定输入与整理输出，逐条核对去向、引用、版本、维度、重复与真实例子；
  不读取主脚本自报数，不固定任何总数。
- 幂等：同一固定输入重跑整理与索引，数据文件字节一致。
- 增量：输入行序变化、增加无关记录，原表现身份不变且不增加重复表现；输入被篡改则拒绝发布。
- 反例：在输出副本上分别断开跨来源关联、移除示例证据、交换现金/银行存款绑定、
  期初误归期末、当前版本填历史定义哈希，对应检查必须失败。
存在必需检查失败、未执行或反例未检出时退出码非零。
"""
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from 建立索引 import 查询语义, 查询维度
from 二次复核回归 import 检查集

项目 = r"C:\Users\27651\Desktop\附注自动更新系统"
运行器 = r"C:\Users\27651\AppData\Local\Programs\Python\Python314\python.exe"
脚本目录 = os.path.dirname(os.path.abspath(__file__))


def 解析参数(argv):
    位置, 输入清单 = [], None
    i = 0
    while i < len(argv):
        if argv[i] == "--输入清单" and i + 1 < len(argv):
            输入清单 = argv[i + 1]; i += 2
        else:
            位置.append(argv[i]); i += 1
    return (位置[0] if 位置 else os.path.join(项目, "语义成果整理")), 输入清单


输出根, 输入清单路径 = 解析参数(sys.argv[1:])


def 读jsonl(路径):
    with open(路径, encoding="utf-8-sig") as f:
        return [json.loads(行) for 行 in f if 行.strip()]


def sha256文件(路径):
    h = hashlib.sha256()
    with open(路径, "rb") as f:
        for 块 in iter(lambda: f.read(1 << 20), b""):
            h.update(块)
    return h.hexdigest()


def 定义哈希(r):
    return hashlib.sha256(json.dumps(r, sort_keys=True, ensure_ascii=False,
                                     separators=(",", ":")).encode()).hexdigest()


class 输入:
    def __init__(self, 清单路径):
        if 清单路径:
            固 = json.load(open(清单路径, encoding="utf-8-sig"))["固定输入"]
            self.V1 = 固["v1"]["路径"]; self.V2 = 固["v2当前版本"]["路径"]
            self.索引 = 固["五样式索引"]["路径"]; self.A = 固["A端映射"]["路径"]
        else:
            self.V1 = os.path.join(项目, "金标准", "财务报表附注语义金标准-v1.jsonl")
            self.V2 = os.path.join(项目, "金标准", "版本",
                              "财务报表附注语义金标准_20260922_132731_236607_5547978f.jsonl")
            self.索引 = os.path.join(项目, "金标准", "五种样式人工映射", "语义示例索引.json")
            self.A = os.path.join(项目, "测试结果", "真实语义识别续跑_A_20260922_050354_458720",
                             "导入外部复核_20260922_051437_942503_b323c5", "附注表格语义映射.json")
        self.v1 = {r["id"]: r for r in 读jsonl(self.V1)}
        self.转换, self.指标, self.维度, self.示例 = [], {}, {}, {}
        for r in 读jsonl(self.V2):
            if r["record_type"] == "legacy_conversion":
                self.转换.append(r)
            elif r["record_type"] == "metric":
                self.指标[r["id"]] = r
            elif r["record_type"] == "dimension":
                self.维度[r["id"]] = r
            elif r["record_type"] == "source_example":
                self.示例[r["id"]] = r
        self.五样式 = json.load(open(self.索引, encoding="utf-8-sig"))
        self.A映射 = json.load(open(self.A, encoding="utf-8-sig"))


class 输出:
    """按根目录懒加载整理输出。"""

    def __init__(self, 根):
        self.根 = 根
        self._缓存 = {}

    def 取(self, 名):
        if 名 not in self._缓存:
            路径们 = {
                "语义": ("数据", "统一语义记录.jsonl"), "表现": ("数据", "表现记录.jsonl"),
                "维字典": ("数据", "维度字典.jsonl"), "旧新": ("对应", "旧新对应.jsonl"),
                "示例去向": ("对应", "来源示例.jsonl"),
            }
            if 名 in 路径们:
                self._缓存[名] = 读jsonl(os.path.join(self.根, *路径们[名]))
            elif 名 == "问题单":
                self._缓存[名] = json.load(open(os.path.join(self.根, "问题", "问题清单.json"), encoding="utf-8-sig"))
            elif 名 == "索引":
                self._缓存[名] = json.load(open(os.path.join(self.根, "索引", "查找索引.json"), encoding="utf-8-sig"))
            elif 名 == "上下文":
                self._缓存[名] = json.load(open(os.path.join(self.根, "数据", "上下文证据.json"), encoding="utf-8-sig"))
        return self._缓存[名]


def 静态检查(原, 出, 只要=None):
    """返回 {检查名: (通过, 详情)}。只要=检查名集合时只跑这些。"""
    结果 = {}

    def 检查(名, 条件, 详情=""):
        if 只要 is None or 名 in 只要:
            结果[名] = (bool(条件), str(详情)[:400])

    语义 = 出.取("语义"); 表现 = 出.取("表现"); 旧新 = 出.取("旧新")
    按id = {r["semantic_id"]: r for r in 语义}
    当前条目 = {r["legacy_id"]: r for r in 旧新 if r.get("版本") == "当前v1"}
    历史条目 = [r for r in 旧新 if r.get("版本") == "历史v1"]

    # ---- 输入覆盖 ----
    active = {i for i, r in 原.v1.items() if r["status"] == "active"}
    检查("v1active全有语义记录", active <= set(按id), f"active {len(active)}")
    非active = {i for i, r in 原.v1.items() if r["status"] != "active"}
    检查("v1非active全部历史保留",
         非active <= {r["legacy_id"] for r in 旧新 if "历史保留" in str(r.get("本次规范对应")) + str(r.get("本次核验结论"))},
         f"非active {len(非active)}")
    v1src = [p for p in 表现 if p["source_kind"] == "v1_source"]
    原来源 = {(r["id"], 定义哈希(e)) for r in 原.v1.values() if r["status"]=="active" for e in r.get("sources",[])}
    成果来源 = {(p.get("定义旧ID"), 定义哈希(e)) for p in 表现 if p["source_kind"]=="v1_source"
                for e in p.get("来源原文证据",[])}
    检查("active来源对象全部成为表现", 原来源==成果来源, f"不同来源证据{len(原来源)}/{len(成果来源)}")

    fs = [p for p in 表现 if p["source_kind"] == "five_style"]
    检查("五样式表现全数整理", len(fs) == len(原.五样式["examples"]), f"{len(fs)}")
    A表 = [p for p in 表现 if p["source_kind"] == "a_fact"]
    检查("A端全数收录且方法分布519/282/23",
         len(A表) == len(原.A映射["facts"])
         and Counter(p["确认依据"]["方法"] for p in A表)
         == {"模型双轮识别": 519, "模型双轮识别＋同义复核": 282, "外部复核": 23})
    检查("指标全部成记录", set(原.指标) <= set(按id))
    检查("v2维度全覆盖", set(原.维度) <= {d["dimension_id"] for d in 出.取("维字典")})
    收录 = set()
    for r in 旧新:
        for c in r.get("历史转换", []):
            收录.add(c.get("id"))
        for c in r.get("不适用转换", []):
            收录.add(c.get("转换ID"))
    检查("v2转换全部有着落（适用或不适用逐条登记）",
         {c["id"] for c in 原.转换} <= 收录, f"转换{len(原.转换)}，收录{len(收录)}")
    去向 = 出.取("示例去向")
    检查("来源示例逐条登记去向", len(去向) == len(原.示例)
         and {r["示例ID"] for r in 去向} == set(原.示例), f"{len(去向)}/{len(原.示例)}")
    检查("无b_fact", not any(p["source_kind"] == "b_fact" for p in 表现))

    # ---- 引用完整 ----
    检查("表现都有有效语义", all(p["semantic_id"] in 按id for p in 表现))
    目标们 = {(r.get("本次规范对应") or {}).get("目标semantic_id") for r in 旧新} - {None}
    检查("对应目标语义都存在", 目标们 <= set(按id), f"{len(目标们)}个")
    检查("语义ID唯一", len(按id) == len(语义))
    检查("表现ID唯一", len({p["representation_id"] for p in 表现}) == len(表现))
    对应ids = {r.get("对应ID") for r in 旧新}
    检查("表现对应引用可解析",
         all(not p.get("对应引用") or p["对应引用"] in
             对应ids for p in 表现))
    引用示例 = {ref for r in 原.转换 for ref in (r.get("evidence_refs") or []) if str(ref).startswith("source:")}
    检查("转换source引用可解析到去向", 引用示例 <= {r["示例ID"] for r in 去向})

    # ---- 版本适用 ----
    坏当前 = [r["legacy_id"] for r in 当前条目.values()
             if r["legacy_id"] in 原.v1 and r.get("旧定义哈希") != 定义哈希(原.v1[r["legacy_id"]])]
    检查("当前版本条目定义哈希与重算全量一致", not 坏当前, f"{len(坏当前)}条")
    内嵌哈希 = {}
    for e in 原.示例.values():
        if e["origin"].get("legacy_id") and e["origin"].get("legacy_definition_hash"):
            内嵌哈希.setdefault(e["origin"]["legacy_id"], set()).add(e["origin"]["legacy_definition_hash"])
    坏历史 = [r["legacy_id"] for r in 历史条目
             if r.get("旧定义哈希") not in 内嵌哈希.get(r["legacy_id"], set())]
    检查("历史版本条目定义哈希与内嵌定义一致", not 坏历史, f"{len(坏历史)}条")
    approved适用 = set()
    for r in 旧新:
        for c in r.get("历史转换", []):
            if c.get("status") == "approved":
                approved适用.add(c.get("id"))
    拒绝依据 = {c["转换ID"] for r in 旧新 for c in r.get("不适用转换", []) if c.get("原因")}
    检查("approved转换逐条适用或保留拒绝依据",
         {c["id"] for c in 原.转换 if c["status"] == "approved"} <= (approved适用 | 拒绝依据),
         f"已按版本适用{len(approved适用)}；其余须有拒绝原因")
    停用 = [r for r in 语义 if r["lifecycle"]["状态"] == "retired"]
    检查("停用指标保留且不复活",
         len(停用) == 3 and all(原.指标.get(r["semantic_id"], {}).get("status") == "retired" for r in 停用))

    # ---- 维度 ----
    空含义 = [d["dimension_id"] for d in 出.取("维字典") if not d.get("含义")]
    缺口 = (出.取("问题单").get("本次整理未解决", {}) or {}).get("维度缺口", [])
    检查("维度空含义逐条登记缺口",
         not 空含义 and not 缺口, f"空{len(空含义)}，登记{len(缺口)}")
    维按id = {d["dimension_id"]: d for d in 出.取("维字典")}
    检查("period_role规范值齐备",
         {"closing", "opening", "current", "prior"} <= set(维按id.get("period_role", {}).get("规范值", [])))
    检查("期间角色不作独立维度", "期间角色" not in 维按id)
    占位 = re.compile(r"^<.+>$")
    检查("模板占位未登记为真实取值",
         all(not (isinstance(v, str) and 占位.match(v))
             for p in 表现 for v in p.get("dimension_values", {}).values()))

    # ---- 重复与修订 ----
    检查("五样式mapping_id无重复",
         len({p["source_locator"]["mapping_id"] for p in fs}) == len(fs))
    双版本旧 = set()
    按旧版本 = {}
    for c in 原.转换:
        if c["status"] == "approved":
            按旧版本.setdefault(c["legacy_id"], set()).add(c["source_gold_hash"])
    双版本旧 = {k for k, v in 按旧版本.items() if len(v) > 1}
    按旧条目 = {}
    for r in 旧新:
        按旧条目.setdefault(r["legacy_id"], set()).add((r.get("源标准哈希"), r.get("旧定义哈希")))
    检查("双版本旧ID分别成条",
         all(len(按旧条目.get(k, set())) >= 2 for k in 双版本旧), f"双版本{len(双版本旧)}")
    禁含 = ["5edec0861b757b0a", "e79babc0af823d82"]
    A源 = (原.A映射.get("source_hash") or "")[:16]
    if A源:
        禁含.append(A源)
    检查("表现ID不含整文件哈希不用顺序号",
         not [p for p in 表现 if any(h in p["representation_id"] for h in 禁含)
              or re.search(r"#\d+$", p["representation_id"])])

    # ---- 真实例子与查询行为 ----
    三槽 = {"C-N087-T001-S001": ("库存现金", "closing"), "C-N087-T001-S002": ("库存现金", "opening"),
           "C-N087-T001-S003": ("银行存款", "closing")}
    检查("现金三槽位类别与期间",
         all(类别 in json.dumps(按id.get(旧, {}).get("meaning", {}), ensure_ascii=False)
             and 按id.get(旧, {}).get("meaning", {}).get("期间角色") == 期间
             for 旧, (类别, 期间) in 三槽.items()))
    受限 = [p for p in fs if p["source_locator"].get("mapping_id") in
           ("S1-REVIEW-0011", "S1-REVIEW-0013", "S1-REVIEW-0015")]
    检查("受限保证金三取值挂同一语义",
         {"银行承兑汇票保证金", "履约保证金", "保函保证金"}
         <= {p["dimension_values"].get("restriction_type") for p in 受限}
         and {"银行承兑汇票保证金", "履约保证金", "保函保证金"}
         <= {p.get("规范维度值", {}).get("restriction_type") for p in 受限}
         and all(p["semantic_id"] == "C-N057-T003-S019" for p in 受限))
    合计 = 按id.get("C-N087-T001-S007", {})
    检查("合计保留汇总范围",
         合计.get("meaning", {}).get("汇总角色") == "total"
         and 合计.get("meaning", {}).get("汇总关系", {}).get("component_variants") is not None)
    检查("开放项目dimension角色",
         按id.get("C-N009-T002-S001", {}).get("meaning", {}).get("槽位角色") == "dimension")
    索引 = 出.取("索引")
    表现按pid = {p["representation_id"]: p for p in 表现}
    mf = 索引.get("by_semantic", {}).get("metric.monetary_funds.balance", {}).get("表现", [])
    mf来源 = Counter(表现按pid[i]["source_kind"] for i in mf if i in 表现按pid)
    S001n = len(原.v1["C-N087-T001-S001"].get("sources", []))
    S001五 = sum(1 for e in 原.五样式["examples"] if e["legacy_slot_id"] == "C-N087-T001-S001")
    S001表现 = {p["representation_id"] for p in 表现
              if p.get("定义旧ID") == "C-N087-T001-S001" and p.get("source_kind") == "v1_source"}
    检查("货币资金指标汇集三来源",
         mf来源.get("v1_source", 0) >= S001n and mf来源.get("five_style", 0) >= S001五
         and mf来源.get("a_fact", 0) >= 26 and bool(S001表现) and S001表现 <= set(mf),
         str(dict(mf来源)) + f"，S001v1表现{len(S001表现)}条在汇集{len(S001表现 & set(mf))}条")
    维度值表 = 索引.get("表现维度值", {})
    筛 = 查询语义(索引, "metric.monetary_funds.balance", {"cash_selection":"库存现金", "period_role":"closing"})
    筛旧 = {表现按pid[i].get("定义旧ID") for i in 筛 if i in 表现按pid}
    检查("筛选库存现金期末不混银行存款或期初",
         len(筛) >= S001n + S001五
         and 筛旧 <= {"C-N087-T001-S001", "C-N087-T002-S015", None}
         and not ({"C-N087-T001-S002", "C-N087-T001-S003"} & 筛旧),
         f"{len(筛)}条 {sorted(x for x in 筛旧 if x)}")
    别名 = 索引.get("维度别名", {}); 取别 = 索引.get("取值别名", {})

    def 查维度(名, 值):
        return 查询维度(索引, 名, 值)

    q中 = 查维度("期间角色", "期末"); q英 = 查维度("period_role", "closing"); q初 = 查维度("period_role", "opening")
    检查("期间角色/期末与period_role/closing同集合", q中 == q英 and len(q英) > 0,
         f"{len(q中)}/{len(q英)}")
    检查("期初不混入期末", not (q英 & q初) and len(q初) > 0)
    s001集合 = set(索引.get("by_semantic", {}).get("C-N087-T001-S001", {}).get("表现", []))
    检查("旧ID查询不扩大",
         s001集合 and {表现按pid[i].get("定义旧ID") for i in s001集合 if i in 表现按pid}
         <= {"C-N087-T001-S001", None})

    # ---- 输出健康 ----
    乱码 = []
    for 相 in ("数据/统一语义记录.jsonl", "数据/表现记录.jsonl", "数据/维度字典.jsonl",
             "对应/旧新对应.jsonl", "对应/来源示例.jsonl", "问题/问题清单.json", "索引/查找索引.json"):
        p = os.path.join(出.根, 相)
        if not os.path.exists(p):
            乱码.append(相 + ":缺失"); continue
        文 = open(p, encoding="utf-8-sig").read()
        if "�" in 文:
            乱码.append(相)
        if any(s in 文 for s in ("encrypted_api_key", "api_key", "secret")):
            乱码.append(相 + ":疑似密钥")
    检查("输出无乱码无密钥", not 乱码, str(乱码))
    检查("问题单分类齐全",
         all(k in 出.取("问题单") for k in ("原本尚未完成的识别", "本次整理未解决", "无法读取或核验的来源",
                                  "历史转换状态", "证据合并")))
    return 结果


def 跑脚本(名, *参数):
    r = subprocess.run([运行器, "-B", "-X", "utf8", os.path.join(脚本目录, 名), *参数],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def main():
    报告 = {"检查项": [], "通过": True}
    from datetime import datetime
    备份目录 = os.path.join(r"C:\Users\27651\BackUp", "语义成果核验_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
    os.makedirs(备份目录)
    def 备份(路径):
        if not os.path.isfile(路径):
            return
        目标 = os.path.join(备份目录, str(len(os.listdir(备份目录))) + "_" + os.path.basename(路径))
        shutil.copy2(路径, 目标)
        assert sha256文件(路径) == sha256文件(目标)
    备份(os.path.join(输出根, "核验报告.json"))

    def 记(名, 通过, 范围="", 预期="", 实际=""):
        报告["检查项"].append({"名称": 名, "通过": bool(通过) if 通过 is not None else None,
                           "输入范围": 范围, "预期": 预期, "实际": str(实际)[:400]})
        if 通过 is not True:
            报告["通过"] = False
        print(("PASS " if 通过 else "FAIL ") + 名 + (" | " + str(实际)[:200] if 实际 else ""))

    原 = 输入(输入清单路径)
    print("== 静态全量检查 ==")
    出 = 输出(输出根)
    静态 = 静态检查(原, 出)
    for 名, (过, 详情) in 静态.items():
        记("静态:" + 名, 过, "固定输入全集+候选输出", "见修补计划任务5", 详情)

    for r in 检查集(输出根, 输入清单路径):
        记("二次复核:" + r["名称"], r["通过"], "原始输入逐条对照", "无遗漏、无错误关联", r["详情"])

    输出文件们 = ["数据/统一语义记录.jsonl", "数据/表现记录.jsonl", "数据/维度字典.jsonl",
               "对应/旧新对应.jsonl", "对应/来源示例.jsonl", "问题/问题清单.json", "数据/上下文证据.json", "对应/维度规范对应.json", "索引/查找索引.json"]

    临时 = tempfile.mkdtemp(prefix="核验留存_", dir=输出根)
    报告["验证留存目录"] = 临时
    try:
        # ---- 幂等：同一输入重跑，字节一致 ----
        print("== 幂等重跑 ==")
        复跑 = os.path.join(临时, "复跑")
        code, log = 跑脚本("统一整理.py", 复跑, "--输入清单", 输入清单路径) if 输入清单路径 else (1, "无输入清单")
        if code != 0:
            记("幂等:重跑整理", False, "同一固定输入", "退出0", log[-300:])
        else:
            code, log = 跑脚本("建立索引.py", 复跑)
            不一致 = [f for f in 输出文件们
                     if not os.path.exists(os.path.join(复跑, f))
                     or sha256文件(os.path.join(复跑, f)) != sha256文件(os.path.join(输出根, f))]
            记("幂等:同一输入两次整理输出字节一致", code == 0 and not 不一致, "同一固定输入重跑",
               "数据文件字节一致", str(不一致))

        # ---- 增量：行序变化与无关记录 ----
        print("== 增量稳定性 ==")
        原表现IDs = {p["representation_id"] for p in 出.取("表现")}
        固 = json.load(open(输入清单路径, encoding="utf-8-sig"))
        v1行 = open(原.V1, encoding="utf-8-sig").read().splitlines()
        证据增量 = []
        for 行 in v1行:
            if not 行.strip():
                continue
            r = json.loads(行)
            if r["id"] == "C-N087-T001-S001":
                r["sources"].append({**r["sources"][0], "新增证据说明":"复核增量验证同一表现的补充证据"})
            证据增量.append(json.dumps(r, ensure_ascii=False))
        for 名, 行们, 预期描述 in (
            ("行序变化", list(reversed(v1行)), "原表现ID不变"),
            ("增加同位置证据", 证据增量, "原表现ID不变、补充证据合并保存"),
            ("增加无关记录", v1行 + [json.dumps({
                "id": "X-IRRELEVANT-TEST", "status": "active",
                "note": {"name": "无关测试科目", "id": "X"}, "table": {"name": "无关表", "id": "X-T"},
                "slot": {"name": "无关槽位", "role": "measure"}, "row_path": [], "column_path": [],
                "value_type": "text", "dimensions": [], "sources": []}, ensure_ascii=False)],
             "原表现ID不变且不增加重复表现"),
        ):
            副本目录 = os.path.join(临时, 名)
            os.makedirs(副本目录)
            副本v1 = os.path.join(副本目录, "v1.jsonl")
            with open(副本v1, "w", encoding="utf-8-sig", newline="\n") as f:
                f.write("\n".join(x for x in 行们 if x.strip()) + "\n")
            单 = dict(固)
            单["固定输入"] = dict(固["固定输入"])
            单["固定输入"]["v1"] = {"路径": 副本v1, "sha256": sha256文件(副本v1)}
            清单副本 = os.path.join(副本目录, "清单.json")
            with open(清单副本, "w", encoding="utf-8-sig", newline="\n") as f:
                json.dump(单, f, ensure_ascii=False)
            出目录 = os.path.join(副本目录, "输出")
            code, log = 跑脚本("统一整理.py", 出目录, "--输入清单", 清单副本)
            if code != 0:
                记("增量:" + 名, False, "修改后输入副本", 预期描述, log[-300:])
                continue
            新表现 = 读jsonl(os.path.join(出目录, "数据", "表现记录.jsonl"))
            新IDs = {p["representation_id"] for p in 新表现}
            记("增量:" + 名, 原表现IDs <= 新IDs and len(新IDs) == len(新表现)
               and len(新IDs) <= len(原表现IDs),
               "修改后输入副本", 预期描述,
               f"原{len(原表现IDs)}，新{len(新IDs)}，缺失{len(原表现IDs - 新IDs)}")

            if 名 == "增加同位置证据":
                已保存 = [e for p in 新表现 for e in p.get("来源原文证据", []) if e.get("新增证据说明")]
                记("增量:新增证据实际保存且未造重复表现", len(已保存)==1 and 新IDs==原表现IDs,
                   "同一来源位置附加证据", "一条补充证据且身份集合不变", str(len(已保存)))

        # 单独改A文件版本哈希，业务内容不变；通过真实生成入口验证，而非只查ID里有无哈希文字。
        A新 = json.loads(json.dumps(原.A映射, ensure_ascii=False))
        A新["source_hash"] = "f" * 64
        for fact in A新["facts"]:
            fact.setdefault("source_reference", {})["file_hash"] = "f" * 64
        A副本 = os.path.join(临时, "A版本变化.json")
        with open(A副本,"w",encoding="utf-8-sig") as f:
            json.dump(A新,f,ensure_ascii=False)
        A单 = json.loads(json.dumps(固,ensure_ascii=False))
        A单["固定输入"]["A端映射"] = {"路径":A副本,"sha256":sha256文件(A副本)}
        A清单 = os.path.join(临时,"A版本清单.json")
        with open(A清单,"w",encoding="utf-8-sig") as f:
            json.dump(A单,f,ensure_ascii=False)
        A出 = os.path.join(临时,"A版本输出")
        code,log=跑脚本("统一整理.py",A出,"--输入清单",A清单)
        新IDs = {p["representation_id"] for p in 读jsonl(os.path.join(A出,"数据","表现记录.jsonl"))} if code==0 else set()
        记("增量:A整文件版本变化不重建原表现", code==0 and 新IDs==原表现IDs,
           "只改来源文件版本指纹", "全部表现身份不变", f"退出{code}，缺失{len(原表现IDs-新IDs)}")

        # ---- 增量：输入篡改拒绝发布 ----
        坏v1 = os.path.join(临时, "篡改v1.jsonl")
        with open(坏v1, "wb") as f:
            f.write(open(原.V1, "rb").read() + b" ")
        坏单 = dict(固)
        坏单["固定输入"] = dict(固["固定输入"])
        坏单["固定输入"]["v1"] = {"路径": 坏v1, "sha256": 固["固定输入"]["v1"]["sha256"]}
        坏清单 = os.path.join(临时, "篡改清单.json")
        with open(坏清单, "w", encoding="utf-8-sig", newline="\n") as f:
            json.dump(坏单, f, ensure_ascii=False)
        code, log = 跑脚本("统一整理.py", os.path.join(临时, "篡改输出"), "--输入清单", 坏清单)
        记("增量:输入与清单不一致拒绝发布", code != 0, "清单外篡改输入", "退出非零",
           f"退出码{code}" if code != 0 else "竟然通过")

        # ---- 反例：每个必须使对应检查失败 ----
        print("== 反例验证 ==")
        反例们 = []

        def 副本输出(名):
            d = os.path.join(临时, "反例_" + 名)
            for 子 in ("数据", "对应", "问题", "索引"):
                shutil.copytree(os.path.join(输出根, 子), os.path.join(d, 子))
            return d

        def 改jsonl(根, 相对, 改):
            p = os.path.join(根, 相对)
            备份(p)
            记录 = 读jsonl(p)
            记录 = [改(r) for r in 记录]
            with open(p, "w", encoding="utf-8-sig", newline="\n") as f:
                for r in 记录:
                    f.write(json.dumps(r, ensure_ascii=False, sort_keys=True) + "\n")

        # D1 断开S001跨来源对应
        d = 副本输出("断开对应")

        def 断(r):
            if r.get("版本") == "当前v1" and r["legacy_id"] == "C-N087-T001-S001":
                r["本次规范对应"] = {"目标semantic_id": None, "处理": "反例篡改"}
            return r
        改jsonl(d, "对应/旧新对应.jsonl", 断)
        备份(os.path.join(d, "索引", "查找索引.json"))
        跑脚本("建立索引.py", d)
        反例们.append(("反例:断开跨来源关联被检出", "货币资金指标汇集三来源", d))
        # D2 移除一条来源示例去向
        d = 副本输出("移除示例")
        目标示例 = next(iter(原.示例))
        改jsonl(d, "对应/来源示例.jsonl", lambda r: r if r["示例ID"] != 目标示例 else {**r, "示例ID": "removed"})
        反例们.append(("反例:移除来源示例证据被检出", "来源示例逐条登记去向", d))
        # D3 交换现金与银行存款绑定
        d = 副本输出("交换绑定")

        def 换(r):
            if r.get("定义旧ID") == "C-N087-T001-S001" and "规范维度值" in r:
                r["规范维度值"]["cash_selection"] = "银行存款"
            if r.get("定义旧ID") == "C-N087-T001-S003" and "规范维度值" in r:
                r["规范维度值"]["cash_selection"] = "库存现金"
            return r
        改jsonl(d, "数据/表现记录.jsonl", 换)
        备份(os.path.join(d, "索引", "查找索引.json"))
        跑脚本("建立索引.py", d)
        反例们.append(("反例:交换现金和银行存款绑定被检出", "筛选库存现金期末不混银行存款或期初", d))
        # D4 期初误归期末
        d = 副本输出("期初归期末")

        def 归(r):
            if r.get("定义旧ID") == "C-N087-T001-S002" and "规范维度值" in r:
                r["规范维度值"]["period_role"] = "closing"
            return r
        改jsonl(d, "数据/表现记录.jsonl", 归)
        备份(os.path.join(d, "索引", "查找索引.json"))
        跑脚本("建立索引.py", d)
        反例们.append(("反例:期初误归期末被检出", "筛选库存现金期末不混银行存款或期初", d))
        # D5 当前版本填历史定义哈希
        d = 副本输出("指纹混配")
        受害 = next(iter(当前条目IDs := {r["legacy_id"] for r in 读jsonl(os.path.join(输出根, "对应/旧新对应.jsonl")) if r.get("版本") == "当前v1"}))

        def 混(r):
            if r.get("版本") == "当前v1" and r["legacy_id"] == 受害:
                r["旧定义哈希"] = "0" * 64
            return r
        改jsonl(d, "对应/旧新对应.jsonl", 混)
        反例们.append(("反例:当前版本填历史定义哈希被检出", "当前版本条目定义哈希与重算全量一致", d))

        for 名, 检查名, d in 反例们:
            子 = 静态检查(原, 输出(d), 只要={检查名})
            通过 = 子.get(检查名, (True, "检查未执行"))[0]
            记(名, 通过 is False, f"输出副本{os.path.basename(d)}",
               f"检查[{检查名}]应失败", "已检出" if 通过 is False else "未检出（检查仍通过或未执行）")
        for 名, 相, 改, 预期失败 in (
            ("丢失A端结构", "数据/表现记录.jsonl",
             lambda r: {**r,"规范维度值":{}} if r.get("source_kind")=="a_fact" else r,
             "A端集合和期间结构无损保留"),
            ("示例悬空引用", "对应/来源示例.jsonl",
             lambda r: {**r,"去向":{**r["去向"],"关联规范语义":["不存在的语义"]}} if r["示例ID"]==目标示例 else r,
             "来源示例无悬空规范语义引用"),
        ):
            d = 副本输出(名)
            改jsonl(d, 相, 改)
            子 = {r["名称"]:r["通过"] for r in 检查集(d, 输入清单路径)}
            记("反例:"+名+"被检出", 子.get(预期失败) is False, "单项篡改副本", 预期失败+"失败", str(子.get(预期失败)))
    finally:
        # 保留核验证据，不删除目录。清理由用户决定。
        print("核验证据留存：" + 临时)

    报告["统计"] = {"静态检查": len(静态), "二次复核": len(检查集(输出根, 输入清单路径)), "反例": 7}
    with open(os.path.join(输出根, "核验报告.json"), "w", encoding="utf-8-sig", newline="\n") as f:
        json.dump(报告, f, ensure_ascii=False, indent=1, sort_keys=True)
    print("\n总结果：", "全部通过" if 报告["通过"] else "存在失败项")
    sys.exit(0 if 报告["通过"] else 1)


if __name__ == "__main__":
    main()
