# -*- coding: utf-8 -*-
"""核验真实样本：按事先固定的样本清单跑五样式原样查回、真实A端入口、性能测量。

用法：
    核验真实样本.py --索引目录 索引\首版 --样本清单 验收证据\真实样本清单.json --输出 验收证据\真实核验

样本与预期以清单为准（先于本程序运行固定）；任何核对失败退出码1并写明原因。
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

本批目录 = Path(__file__).resolve().parent
sys.path.insert(0, str(本批目录))

from openpyxl.utils import range_boundaries

from 建立结构索引 import 文件指纹, 读取工作簿, 建立索引, 核对输入清单
from 比较新表 import 载入索引, 比较工作簿, 写出交接材料


def 指纹表(路径们):
    return {str(p): 文件指纹(p) for p in 路径们}


def 范围相交(范围a, 范围b):
    a = range_boundaries(范围a)
    b = range_boundaries(范围b)
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def 核验原样查回(样本, 索引, 输出根, 结论):
    """完整文件入口：原样文件应把逐项核对小表恒等查回自身。"""
    记录 = 读取工作簿(样本["文件"])
    结果 = 比较工作簿(记录, 索引)
    目录 = 输出根 / 样本["样本名"]
    写出交接材料(记录, 结果, str(目录))
    for 核对 in 样本["逐项核对小表"]:
        候选 = [c for c in 结果["候选"]
              if c["样式来源"] == 核对["样式"] and c["小表名称"] == 核对["小表名称"]]
        条目 = {"样本": 样本["样本名"], "小表": 核对["小表名称"], "核对点": 核对["核对点"]}
        # 设计A07：合理的相似候选全部保留，不靠排名消歧；因此不要求候选唯一，
        # 要求"恒等查回候选"（旧坐标==新坐标且旧表文字覆盖满分）存在且唯一
        恒等 = []
        for c in 候选:
            非恒等 = [d for d in c["结构对应"]
                   if d["旧工作表"] != d["新工作表"] or d["旧坐标"] != d["新坐标"]]
            覆盖 = c["覆盖计算"]["旧表文字覆盖"]
            if not 非恒等 and 覆盖["分子"] == 覆盖["分母"]:
                恒等.append(c)
        条目["候选数"] = len(候选)
        条目["并存候选"] = [f"{c['新工作表']}!{c['新表范围']} 覆盖{c['覆盖计算']['旧表文字覆盖']['分子']}/{c['覆盖计算']['旧表文字覆盖']['分母']}"
                       for c in 候选]
        if len(恒等) != 1:
            条目["结论"] = "失败"
            条目["原因"] = f"恒等查回候选数={len(恒等)}，应为1（候选总数{len(候选)}）"
        else:
            c = 恒等[0]
            条目["候选范围"] = f"{c['新工作表']}!{c['新表范围']}"
            条目["结构对应数"] = len(c["结构对应"])
            条目["合并核对"] = c["关系核对"]["合并"]
            条目["结论"] = "通过"
        结论.append(条目)
    return 结果


def 核验真实A端(a端, 索引, 输出根, 结论):
    记录 = 读取工作簿(a端["文件"])
    结果 = 比较工作簿(记录, 索引)
    目录 = 输出根 / "真实A端"
    写出交接材料(记录, 结果, str(目录))
    for 区域 in a端["共同表达代表区域"]:
        范围 = 区域["范围"].split("!")[-1]
        命中 = [c for c in 结果["候选"]
              if any(c["样式来源"] == r["样式"] and c["小表名称"] == r["小表名称"] for r in 区域["应找到参照"])
              and 范围相交(c["新表范围"], 范围)]
        结论.append({"样本": "真实A端", "区域": 区域["区域名"], "范围": 区域["范围"],
                   "依据": 区域["共有文字依据"],
                   "结论": "通过" if 命中 else "失败",
                   "命中": [f"样式{c['样式来源']} {c['小表名称']} {c['新表范围']}" for c in 命中],
                   **({} if 命中 else {"原因": "代表区域未实际查回任何事先固定的参照"})})
    未匹配格数 = sum(m["坐标数"] for m in 结果["未匹配范围"])
    结论.append({"样本": "真实A端", "区域": "未知结构如实列出",
               "结论": "通过" if 结果["未匹配范围"] else "注意",
               "说明": f"未匹配范围{len(结果['未匹配范围'])}段共{未匹配格数}格，已写入初筛结果.json与LLM交接说明.md"})
    return 记录, 结果


def main(argv=None):
    解析 = argparse.ArgumentParser(description="真实样本核验")
    解析.add_argument("--索引目录", required=True)
    解析.add_argument("--样本清单", required=True)
    解析.add_argument("--输出", required=True)
    解析.add_argument("--输入清单", default=str(本批目录 / "输入清单.json"))
    参数 = 解析.parse_args(argv)

    输出 = Path(参数.输出)
    if 输出.exists():
        print(f"输出目录已存在，拒绝覆盖：{输出}", file=sys.stderr)
        return 2
    输出.mkdir(parents=True)

    清单 = json.loads(Path(参数.样本清单).read_text(encoding="utf-8-sig"))
    涉及文件 = [s["文件"] for s in 清单["五样式原样查回"]] + [清单["真实A端"]["文件"]]
    for s in 清单["五样式原样查回"]:
        实际 = 文件指纹(s["文件"])
        if 实际 != s["sha256"]:
            print(f"样本指纹不符：{s['样本名']} {s['文件']}", file=sys.stderr)
            return 2
    if 文件指纹(清单["真实A端"]["文件"]) != 清单["真实A端"]["sha256"]:
        print("真实A端文件指纹不符", file=sys.stderr)
        return 2

    输入清单 = json.loads(Path(参数.输入清单).read_text(encoding="utf-8-sig"))
    # 先核对包含复用代码在内的全部输入，避免长时间比较后才在冷建库时发现旧指纹。
    核对输入清单(输入清单)
    原件们 = [e["路径"] for e in 输入清单["条目"]] + 涉及文件
    前指纹 = 指纹表(原件们)

    索引 = 载入索引(参数.索引目录)
    结论 = []
    开始 = time.perf_counter()

    # 一、五样式原样查回（完整文件入口，不传表区）
    for 样本 in 清单["五样式原样查回"]:
        核验原样查回(样本, 索引, 输出, 结论)

    # 二、真实A端入口
    a端记录, a端结果 = 核验真实A端(清单["真实A端"], 索引, 输出, 结论)

    # 三、性能：冷建库一次 + 同一新文件重复比较3次
    t0 = time.perf_counter()
    冷建目录 = 输出 / "冷建索引"
    建库结果 = 建立索引(参数.输入清单, str(冷建目录))
    冷建秒 = round(time.perf_counter() - t0, 3)

    重复 = []
    for i in range(3):
        t读 = time.perf_counter()
        记录 = 读取工作簿(清单["真实A端"]["文件"])
        读取秒 = time.perf_counter() - t读
        t比 = time.perf_counter()
        结果 = 比较工作簿(记录, 索引)
        比较秒 = time.perf_counter() - t比
        t写 = time.perf_counter()
        写出交接材料(记录, 结果, str(输出 / "性能" / f"运行{i + 1}"))
        输出秒 = time.perf_counter() - t写
        重复.append({"次序": i + 1, "读取秒": round(读取秒, 3),
                  "预筛选秒": 结果["运行统计"]["阶段耗时"]["预筛选秒"],
                  "候选对齐秒": 结果["运行统计"]["阶段耗时"]["候选对齐秒"],
                  "比较秒": round(比较秒, 3), "输出秒": round(输出秒, 3),
                  "候选数": 结果["运行统计"]["候选数"],
                  "比较小表数": 结果["运行统计"]["比较的小表数"]})

    t基 = time.perf_counter()
    基准 = 比较工作簿(a端记录, 索引, 预筛选=False)
    基准秒 = round(time.perf_counter() - t基, 3)
    快速集合 = {(c["旧表编号"], c["新工作表"], c["新表范围"]) for c in a端结果["候选"]}
    基准集合 = {(c["旧表编号"], c["新工作表"], c["新表范围"]) for c in 基准["候选"]}
    遗漏 = 基准集合 - 快速集合

    中位 = sorted(r["比较秒"] for r in 重复)[1]
    性能 = {
        "冷建库": {"耗时秒": 冷建秒, "小表数": 建库结果["小表数"], "映射去向": 建库结果["映射去向"]},
        "重复运行": 重复,
        "比较秒中位": 中位,
        "基准对照": {"基准耗时秒": 基准秒, "基准比较小表数": 基准["运行统计"]["比较的小表数"],
                  "基准候选数": len(基准集合), "快速候选数": len(快速集合),
                  "候选遗漏": sorted(遗漏), "结论": "一致" if not 遗漏 else "快速筛选漏候选"},
        "数据规模": {"索引小表数": len(索引["记录"]),
                  "新表工作表数": len(a端记录["sheets"]),
                  "新表内容格数": sum(len(s["cells"]) for s in a端记录["sheets"])},
    }
    (输出 / "性能记录.json").write_text(
        json.dumps(性能, ensure_ascii=False, indent=1), encoding="utf-8-sig")

    后指纹 = 指纹表(原件们)
    改动 = {p: (前指纹[p], 后指纹[p]) for p in 前指纹 if 前指纹[p] != 后指纹[p]}
    (输出 / "原件核对.json").write_text(json.dumps(
        {"说明": "核验运行前后全部固定输入与样本原件的SHA-256对照",
         "前": 前指纹, "后": 后指纹, "不一致": 改动},
        ensure_ascii=False, indent=1), encoding="utf-8-sig")

    失败 = [c for c in 结论 if c["结论"] == "失败"]
    汇总 = {"核验时间秒": round(time.perf_counter() - 开始, 3),
          "核对项数": len(结论), "通过": sum(1 for c in 结论 if c["结论"] == "通过"),
          "失败": len(失败), "注意": sum(1 for c in 结论 if c["结论"] == "注意"),
          "原件不一致": 改动, "明细": 结论}
    (输出 / "核验结论.json").write_text(
        json.dumps(汇总, ensure_ascii=False, indent=1), encoding="utf-8-sig")

    print(f"核验完成：{汇总['通过']}/{汇总['核对项数']}通过，失败{汇总['失败']}，注意{汇总['注意']}")
    print(f"冷建库{冷建秒}秒；重复比较中位{中位}秒；基准对照：{性能['基准对照']['结论']}")
    print(f"原件不一致：{len(改动)}个")
    print(f"结果目录：{输出}")
    return 1 if 失败 or 改动 or 遗漏 else 0


if __name__ == "__main__":
    sys.exit(main())
