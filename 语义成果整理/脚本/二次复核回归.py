# -*- coding: utf-8 -*-
"""针对二次复核发现的真实缺口；读取交付与原输入，独立推导预期。"""
import copy
import hashlib
import importlib.util
import json
import sys
from pathlib import Path

项目根 = Path(__file__).resolve().parents[1]

def 读行(p):
    return [json.loads(s) for s in p.read_text(encoding="utf-8-sig").splitlines() if s.strip()]

def 检查集(根, 清单路径=None):
    根 = Path(根)
    清单路径 = Path(清单路径 or 项目根 / "资产清单.json")
    固 = json.loads(清单路径.read_text(encoding="utf-8-sig"))["固定输入"]
    A = json.loads(Path(固["A端映射"]["路径"]).read_text(encoding="utf-8-sig"))
    表现 = 读行(根 / "数据/表现记录.jsonl")
    语义 = {r["semantic_id"]: r for r in 读行(根 / "数据/统一语义记录.jsonl")}
    维度 = 读行(根 / "数据/维度字典.jsonl")
    示例 = 读行(根 / "对应/来源示例.jsonl")
    对应 = 读行(根 / "对应/旧新对应.jsonl")
    索引 = json.loads((根 / "索引/查找索引.json").read_text(encoding="utf-8-sig"))
    spec = importlib.util.spec_from_file_location("复核查询", 项目根 / "脚本/建立索引.py")
    查询 = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(查询)
    结果 = []
    def 记(名, 过, 详情=""):
        结果.append({"名称": 名, "通过": bool(过), "详情": str(详情)[:300]})
    A按格 = {p["source_locator"]["sheet_cell"]: p for p in 表现 if p["source_kind"] == "a_fact"}
    缺 = [(f["id"], k) for f in A["facts"] for k,v in f["dimensions"].items()
          if isinstance(v,(dict,list)) and A按格[f["id"]].get("规范维度值",{}).get(k) != v]
    记("A端集合和期间结构无损保留", not 缺, 缺[:4])
    try:
        f = next(f for f in A["facts"] if f["id"] == "sheet1!C3")
        pid = A按格[f["id"]]["representation_id"]
        got = 查询.查询语义(索引, f["metric_id"], {"cash_selection": f["dimensions"]["cash_selection"]})
        记("结构化条件可查询且命中原事实", pid in got)
    except Exception as e:
        记("结构化条件可查询且命中原事实", False, repr(e))
    try:
        c = A按格["sheet1!C3"]["representation_id"]
        d = A按格["sheet1!D3"]["representation_id"]
        got = 查询.查询语义(索引,"metric.monetary_funds.balance",{"期间角色":"期末","cash_selection":"库存现金"})
        记("现金期末筛选包含A端且不含期初", c in got and d not in got)
    except Exception as e:
        记("现金期末筛选包含A端且不含期初",False,repr(e))
    坏 = [(e["示例ID"],s) for e in 示例 for s in e["去向"].get("关联规范语义",[]) if s not in 语义]
    记("来源示例无悬空规范语义引用", not 坏, f"{len(坏)}条")
    对应ids = {r.get("对应ID") for r in 对应}
    记("表现对应引用精确解析",all(not p.get("对应引用") or p["对应引用"] in 对应ids for p in 表现))
    记("维度没有未完成定义",all(d.get("含义") and d.get("值类型") for d in 维度))
    记("期间类型与期间角色分开",语义["C-N087-T001-S001"]["meaning"].get("期间类型") == "instant")
    问题 = json.loads((根/"问题/问题清单.json").read_text(encoding="utf-8-sig"))
    记("A排除不误称未完成识别","A端排除" not in 问题.get("原本尚未完成的识别",{}))
    记("表现身份使用完整摘要",all(len(p["representation_id"].split(":")[-1])==64 for p in 表现))
    # 从输入转换逐项核对常量，不用行名推导被测结果。
    当前 = {x["legacy_id"]: x for x in 对应 if x["版本"] == "当前v1"}
    常量错误 = []
    for p in 表现:
        link = 当前.get(p.get("定义旧ID"), {})
        for k,b in (link.get("本次规范对应", {}).get("固定维度限制") or {}).items():
            if b.get("op") == "constant" and p.get("规范维度值", {}).get(k) != b.get("value"):
                常量错误.append((p["representation_id"], k))
    记("已批准常量维度含合计范围完整落到表现", not 常量错误, 常量错误[:3])
    # 结构不相等时不能退化成单成员文字误命中。
    基础 = {"domain":"货币资金", "mode":"members", "members":["库存现金"], "completeness":"complete"}
    反例 = [{**基础,"mode":"complement"}, {**基础,"mode":"all"},
            {**基础,"members":["库存现金","银行存款"]}, {**基础,"completeness":"partial"}]
    记("集合范围不同不误命中明细", all(not 查询.维度符合(v,"库存现金") for v in 反例))
    同义冲突 = 查询.查询语义(索引,"metric.monetary_funds.balance",{"期间角色":"期末","period_role":"opening"})
    记("同一维度别名条件冲突不互相覆盖", not 同义冲突)
    spec2 = importlib.util.spec_from_file_location("复核身份", 项目根/"脚本/统一整理.py")
    身份 = importlib.util.module_from_spec(spec2); spec2.loader.exec_module(身份)
    身份错=[]
    for 类 in ("v1_source","five_style","a_fact"):
        p=copy.deepcopy(next(p for p in 表现 if p["source_kind"]==类))
        old=身份.表现身份(p)
        p["source_locator"].update({"file_hash":"changed","sha256":"changed","mapping_sha256":"changed"})
        p["证据引用"].append("新增证据")
        if 身份.表现身份(p)!=old: 身份错.append(类+"整文件或证据变化换ID")
        p["规范维度值"]["复核专用维度"]="语义内容已变"
        if 身份.表现身份(p)==old: 身份错.append(类+"语义变化仍覆盖旧ID")
    记("整文件与证据变化不换ID而语义变化换ID",not 身份错,身份错)
    ref={x["对应ID"]:x for x in 对应}
    错=[]
    for e in 示例:
        d=e["去向"]
        if d["类型"]!="旧定义证据": continue
        x=ref.get(d.get("对应引用"),{})
        o=d["带版本旧定义"]
        if (x.get("源标准哈希"),x.get("legacy_id"),x.get("旧定义哈希"))!=(o["源标准哈希"],o["旧ID"],o["定义哈希"]):
            错.append(e["示例ID"])
        for pid in d.get("表现",[]):
            if pid not in 索引["by_semantic"].get(o["旧ID"],{}).get("表现",[]): 错.append(pid)
    记("来源示例按完整版本定义关联且表现可解析",not 错,错[:3])
    原示例=[json.loads(l) for l in Path(固["v2当前版本"]["路径"]).read_text(encoding="utf-8-sig").splitlines()
            if l.strip() and json.loads(l).get("record_type")=="source_example"]
    原表={x["id"]:x for x in 原示例}
    记("来源示例原文逐条无损并进入索引",all(e["text_evidence原文"]==原表[e["示例ID"]].get("text_evidence")
        and 索引.get("by_source_example",{}).get(e["示例ID"])==e["去向"] for e in 示例))
    历史原={}
    for e in 原示例:
        for text in e.get("text_evidence",[]):
            try: obj=json.loads(text)
            except (ValueError,TypeError): continue
            if isinstance(obj,dict) and isinstance(obj.get("原定义"),dict):
                origin=e["origin"]
                digest=hashlib.sha256(json.dumps(obj["原定义"],ensure_ascii=False,sort_keys=True,separators=(",",":")).encode()).hexdigest()
                历史原[(origin.get("source_gold_hash") or origin.get("file_hash"),origin.get("legacy_id"),digest)]=obj.get("原状态")
    错历史=[x["legacy_id"] for x in 对应 if x["版本"]=="历史v1" and
         (x["源标准哈希"],x["legacy_id"],x["旧定义哈希"]) not in 历史原]
    错状态=[x["legacy_id"] for x in 对应 if x["版本"]=="历史v1" and
         x["v1状态"]!=历史原.get((x["源标准哈希"],x["legacy_id"],x["旧定义哈希"]))]
    记("历史指纹与状态均来自该版本原文",not 错历史 and not 错状态,f"指纹{len(错历史)}，状态{len(错状态)}")
    上下文=json.loads((根/"数据/上下文证据.json").read_text(encoding="utf-8-sig"))
    记("A端确认与上下文依据完整保留",上下文.get("A端来源上下文")==A.get("source_context")
        and 上下文.get("A端归属记录")==A.get("owners")
        and all(A按格[f["id"]]["确认依据"].get("原确认记录")==f.get("recognition_provenance") for f in A["facts"]))
    伪计量基础={"货币金额","比率/百分比","文本","数值","是/否","日期","枚举"}
    错基础=[r["semantic_id"] for r in 语义.values() if r["semantic_kind"]=="legacy_slot"
           and (r["meaning"].get("计量基础") in 伪计量基础 or not r["meaning"].get("计量基础依据"))]
    记("计量基础不冒充值类型且有明确来源",not 错基础,f"{len(错基础)}条")
    return 结果

if __name__ == "__main__":
    根 = sys.argv[1] if len(sys.argv)>1 else 项目根
    结果 = 检查集(根)
    for r in 结果:
        print(("PASS " if r["通过"] else "FAIL ")+r["名称"]+" "+r["详情"])
    sys.exit(0 if all(r["通过"] for r in 结果) else 1)
