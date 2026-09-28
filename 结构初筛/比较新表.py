# -*- coding: utf-8 -*-
"""比较新表：在完整新工作簿中查找旧小表候选、比对文字与排列、输出差异。

接口约定（02_设计.md 第8节）：
    比较工作簿(工作簿记录, 索引, 预筛选=True) -> dict
    写出交接材料(工作簿记录, 初筛结果, 输出目录) -> None
本模块不调用任何模型；结构对应只说明文字与排列可对应，不构成新的语义确认。
"""
from __future__ import annotations

import json
import time
from collections import defaultdict
from pathlib import Path

from openpyxl.utils import get_column_letter, range_boundaries

from 建立结构索引 import 规范文字, 读jsonl, _json默认

通用词频上限 = 10  # 索引中超过此数小表共有的文字区分作用弱，不能单独作候选依据
比较规则版本 = "结构比较规则v4"


def 载入索引(索引目录: str) -> dict:
    目录 = Path(索引目录)
    记录 = 读jsonl(目录 / "小表记录.jsonl")
    查找 = json.loads((目录 / "文字查找.json").read_text(encoding="utf-8-sig"))
    清单 = json.loads((目录 / "索引清单.json").read_text(encoding="utf-8-sig"))
    return {"目录": str(目录), "记录": 记录, "小表": {r["小表编号"]: r for r in 记录},
            "正文": 查找["正文"], "别名": 查找.get("别名", {}),
            "规则版本": 清单.get("规则版本"), "索引清单": 清单}


def _工作表数据(工作表记录):
    """每表：文字→坐标列表、坐标→显示文字、坐标→原文（未规范化）、内容格集合、数字位置。"""
    文字位置 = defaultdict(list)
    文字列行 = defaultdict(list)
    显示 = {}
    原文 = {}
    内容格 = []
    数字位置 = []
    for 坐标, 格 in 工作表记录["cells"].items():
        值 = 格.get("cached_value") if 格.get("formula") else 格.get("value")
        if 格.get("formula") or 格.get("value") is not None:
            内容格.append(坐标)
        if isinstance(值, str):
            规范 = 规范文字(值)
            if 规范:
                文字位置[规范].append(坐标)
                显示[坐标] = 规范
                原文[坐标] = 值
                列, 行, _, _ = range_boundaries(坐标)
                文字列行[(规范, 列)].append(行)
        elif isinstance(值, (int, float)) and not isinstance(值, bool):
            数字位置.append({"坐标": 坐标, "原值": 值})
    return {"文字位置": 文字位置, "文字列行": 文字列行, "行查找缓存": {},
            "显示": 显示, "原文": 原文, "内容格": 内容格, "数字位置": 数字位置,
            "数字坐标": {d["坐标"] for d in 数字位置},
            "合并": [range_boundaries(m) for m in 工作表记录.get("merges", [])],
            "记录": 工作表记录}


def _旧行序列(小表):
    """旧小表内带查找文字的节点按行分组，保持行序。"""
    行组 = defaultdict(list)
    for n in 小表["节点"]:
        if n.get("查找文字") and n["角色"] in ("标题", "表头", "项目", "说明", "检验", "结构排除"):
            行组[n["行"]].append((n["列"], n["查找文字"], n["角色"], n["坐标"]))
    return [(r, sorted(行组[r])) for r in sorted(行组)]


def _提候选范围(小表, 表数据):
    """从文字命中提出种子偏移。返回 [(dr, dc, 支持文字数)]，不去重候选。"""
    支持 = defaultdict(set)  # (dr,dc) -> set(旧坐标)
    强线索数 = 0
    for n in 小表["节点"]:
        文字 = n.get("查找文字")
        if not 文字 or n["角色"] not in ("标题", "表头", "项目", "说明", "检验", "结构排除"):
            continue
        新位置 = 表数据["文字位置"].get(文字)
        if not 新位置:
            continue
        if len(索引正文计数.get(文字, [])) <= 通用词频上限:
            强线索数 += 1
        for 新坐标 in 新位置:
            c, r, _, _ = range_boundaries(新坐标)
            支持[(r - n["行"], c - n["列"])].add(n["坐标"])
    return [(dr, dc, len(旧集)) for (dr, dc), 旧集 in sorted(支持.items(), key=lambda x: -len(x[1]))
            if len(旧集) >= 2 or (强线索数 >= 1 and len(旧集) >= 1)]


索引正文计数 = {}


def _对齐候选(小表, 工作表名, 表数据, 种子, 弱候选收集=None):
    """以一个种子偏移为起点做允许插入/缺失的顺序对齐，返回候选dict或None。"""
    dr, dc, _ = 种子
    旧行们 = _旧行序列(小表)
    if not 旧行们:
        return None
    显示 = 表数据["显示"]

    def 新文字(行, 列):
        if 行 < 1 or 列 < 1:
            return None
        return 显示.get(f"{get_column_letter(列)}{行}")

    # 表内词频：同名行在同一张旧小表内多次出现时，不能单独作锚
    表内词频 = defaultdict(int)
    for _, 文字组 in 旧行们:
        for _, 文字, _, _ in 文字组:
            表内词频[文字] += 1

    def 锚定强度(文字组):
        return sum(1 for _, 文字, 角色, _ in 文字组 if 表内词频[文字] == 1 and 角色 in ("标题", "项目", "表头", "说明"))

    def 行匹配(新行, 期望):
        命中 = [(列, 文字, 角色, 旧坐标) for 列, 文字, 角色, 旧坐标 in 期望 if 新文字(新行, 列) == 文字]
        return 命中

    def 有文字的行(期望, 起, 止):
        # 只枚举原循环可能命中的行，窗口及判定保持不变；每种列偏移只整理一次。
        键 = tuple((列, 文字) for 列, 文字, _, _ in 期望)
        缓存 = 表数据["行查找缓存"]
        if 键 not in 缓存:
            缓存[键] = sorted({r for 列, 文字 in 键 for r in 表数据["文字列行"].get((文字, 列), ())})
        from bisect import bisect_left, bisect_right
        行们 = 缓存[键]
        return 行们[bisect_left(行们, 起):bisect_right(行们, 止)]

    行映射 = {}
    缺失行 = []
    新增行 = []
    歧义 = []
    底 = max(n["行"] for n in 小表["节点"])
    高度 = 底 - 旧行们[0][0]

    # 第一遍：只对齐有表内唯一锚文字的行
    指针 = max(1, 旧行们[0][0] + dr)
    for 旧行, 文字组 in 旧行们:
        if not 锚定强度(文字组):
            continue
        期望 = [(列 + dc, 文字, 角色, 旧坐标) for 列, 文字, 角色, 旧坐标 in 文字组]
        备选 = []
        预测 = 旧行 + dr
        上限 = min(指针 + 高度 + 20, 表数据["记录"]["max_row"])
        for q in 有文字的行(期望, 指针, 上限):
            命中 = 行匹配(q, 期望)
            if 命中 and (len(命中) == len(期望) or any(表内词频[t] == 1 and r in ("标题", "项目", "表头", "说明") for _, t, r, _ in 命中)):
                # 已过预测位且比已记录的命中更远：q递增，后面只会更远，提前停
                if 备选 and q > 预测 and q - 预测 > abs(备选[-1] - 预测):
                    break
                备选.append(q)
        if not 备选:
            缺失行.append((旧行, 期望))
            continue
        # 同名文字可能属于别的小表（如原值/累计折旧表同项目名）：
        # 取离种子预测位（旧行+dr）最近者，其余如实记入歧义，不取最先出现者替代所属关系
        备选.sort(key=lambda x: (abs(x - (旧行 + dr)), x))
        命中行 = 备选.pop(0)
        if 备选:
            歧义.append({"旧行": 旧行, "选中": 命中行, "备选": 备选, "原因": "同名文字在窗口内多次出现，按种子预测位就近选定，其余位置存疑"})
        行映射[旧行] = 命中行
        指针 = 命中行 + 1

    # 第二遍：非锚行只在相邻锚之间插值查找，首尾外探3行，防止串到相邻小表
    已对齐 = sorted(行映射.items())
    for 旧行, 文字组 in 旧行们:
        if 旧行 in 行映射 or any(r == 旧行 for r, _ in 缺失行):
            continue
        期望 = [(列 + dc, 文字, 角色, 旧坐标) for 列, 文字, 角色, 旧坐标 in 文字组]
        前 = [q for r, q in 已对齐 if r < 旧行]
        后 = [q for r, q in 已对齐 if r > 旧行]
        # 表头外探与表尾对称：取"首锚前3行"与"旧行原位-3"的较宽者
        起 = (前[-1] + 1) if 前 else max(1, min(后[0], 旧行 + dr + 3) - 3 if 后 else 旧行 + dr - 3)
        # 表尾外探：取"末锚后3行"与"旧行原位+3"的较宽者，
        # 兼容合计/检验行与末锚之间隔空行（如填充空行后再合计）的版式
        止 = (后[0] - 1) if 后 else max(前[-1] + 3 if 前 else 0, 旧行 + dr + 3)
        备选 = []
        for q in 有文字的行(期望, 起, 止):
            if 行匹配(q, 期望):
                备选.append(q)
        if not 备选:
            缺失行.append((旧行, 期望))
            continue
        # 表内重名行（如各组合都有的"1年以内"）取离种子预测位最近者，其余记入歧义
        备选.sort(key=lambda x: (abs(x - (旧行 + dr)), x))
        命中行 = 备选.pop(0)
        if 备选:
            歧义.append({"旧行": 旧行, "选中": 命中行, "备选": 备选, "原因": "非锚行在锚间有多个位置，按种子预测位就近选定，其余位置存疑"})
        行映射[旧行] = 命中行
        已对齐 = sorted(行映射.items())

    if len(行映射) < 2:
        return None

    # 列映射：由已对齐行上的文字投票
    列票 = defaultdict(lambda: defaultdict(int))
    差异 = []
    for 旧行, 文字组 in 旧行们:
        if 旧行 not in 行映射:
            continue
        新行 = 行映射[旧行]
        for 列, 文字, 角色, 旧坐标 in 文字组:
            if 新文字(新行, 列 + dc) == 文字:
                列票[列][列 + dc] += 1
            else:
                别处 = [c for c in range(1, 表数据["记录"]["max_column"] + 1) if 新文字(新行, c) == 文字]
                if 别处:
                    列票[列][别处[0]] += 1
                    差异.append({"旧坐标": 旧坐标, "查找文字": 文字, "类型": "列位置变化",
                               "说明": f"旧列{get_column_letter(列)}在新表改到{get_column_letter(别处[0])}列"})
                else:
                    原位 = 新文字(新行, 列 + dc)
                    差异.append({"旧坐标": 旧坐标, "查找文字": 文字, "类型": "文字缺失或改写",
                               "说明": f"新表对应行未找到'{文字}'" + (f"，原位置现为'{原位}'" if 原位 else "")})
    列映射 = {}
    for 列, 票 in 列票.items():
        最高 = max(票.values())
        前列 = [c for c, v in 票.items() if v == 最高]
        列映射[列] = 前列[0]
        if len(前列) > 1:
            歧义.append({"旧列": get_column_letter(列), "候选列": [get_column_letter(c) for c in 前列],
                       "原因": "列对应票数相同，未唯一确定"})

    结构对应 = []
    对上旧文字 = set()
    for n in 小表["节点"]:
        if n["行"] not in 行映射 or n["列"] not in 列映射:
            if n["角色"] == "业务数值":
                差异.append({"旧坐标": n["坐标"], "类型": "缺对应依据",
                           "说明": "业务格缺行或列对应依据，未产生结构对应"})
            continue
        新坐标 = f"{get_column_letter(列映射[n['列']])}{行映射[n['行']]}"
        条目 = {"旧工作表": 小表["来源"]["工作表"], "旧坐标": n["坐标"],
              "新工作表": 工作表名, "新坐标": 新坐标, "角色": n["角色"]}
        if n.get("查找文字"):
            if 新文字(行映射[n["行"]], 列映射[n["列"]]) == n["查找文字"]:
                条目["原文"] = n.get("原文")
                条目["查找文字"] = n["查找文字"]
                对上旧文字.add(n["坐标"])
                结构对应.append(条目)
            else:
                新原文 = 表数据["显示"].get(f"{get_column_letter(列映射[n['列']])}{行映射[n['行']]}")
                差异.append({"旧坐标": n["坐标"], "查找文字": n["查找文字"], "类型": "文字缺失或改写",
                           "说明": f"行列已对应但文字不一致：旧{n['查找文字']!r} 新{新原文!r}"})
        elif n["角色"] == "业务数值":
            条目["语义引用"] = n.get("语义引用")
            条目["待通读确认"] = True
            条目["原文"] = n.get("原文")
            结构对应.append(条目)

    # 排版差异：规范文字相等但原文不同（缩进/空白/全半角），保留给后续通读，不判含义
    for 条目 in 结构对应:
        if not 条目.get("查找文字"):
            continue
        旧原文 = 条目.get("原文")
        新原文 = 表数据["原文"].get(条目["新坐标"])
        if isinstance(旧原文, str) and isinstance(新原文, str) and 旧原文 != 新原文:
            差异.append({"旧坐标": 条目["旧坐标"], "查找文字": 条目["查找文字"], "类型": "排版差异",
                       "说明": f"规范文字一致但原文不同：旧{旧原文[:30]!r} 新{新原文[:30]!r}"})

    # 项目层级核对：索引保存的缩进层级与实际对齐位置的缩进对比
    层级冲突 = []
    for e in 小表.get("项目顺序", []):
        if e["行"] not in 行映射:
            continue
        c0, _, _, _ = range_boundaries(e["坐标"])
        if c0 not in 列映射:
            continue
        新原文 = 表数据["原文"].get(f"{get_column_letter(列映射[c0])}{行映射[e['行']]}")
        if not isinstance(新原文, str):
            continue
        新缩进 = len(新原文) - len(新原文.lstrip(" 　"))
        if 新缩进 != e.get("缩进", 0):
            层级冲突.append({"旧坐标": e["坐标"], "旧缩进": e.get("缩进", 0), "新缩进": 新缩进,
                         "说明": "项目缩进层级与旧表不一致，所属关系可能变化，待通读确认"})

    旧文字节点 = [n for n in 小表["节点"] if n.get("查找文字")]
    旧覆盖分子 = sorted(对上旧文字)
    # 成立规则：候选要声称自己是这张旧小表，必须满足以下之一——
    # 1) 登记文字全部对上（完整排列即依据，不强求稀有词）；
    # 2) 至少对上两个不同识别性文字（索引中不超过通用词频上限的小表共有），
    #    且对上数达到半数或不少于4个（部分匹配必须有区分性锚点组合，单一弱词不够）。
    # 只靠几个通用词巧合对齐的偏移不算候选，记入弱候选备查。
    识别坐标 = {n["坐标"] for n in 旧文字节点
              if len(索引正文计数.get(n["查找文字"], [])) <= 通用词频上限}
    对上数 = len(旧覆盖分子)
    锚对上 = [c for c in 旧覆盖分子 if c in 识别坐标]
    识别词对上 = {n["查找文字"] for n in 旧文字节点 if n["坐标"] in 锚对上}
    全对 = 对上数 == len(旧文字节点) and 对上数 >= 2
    达标 = 对上数 >= max(2, -(-len(旧文字节点) // 2)) or 对上数 >= 4
    if not (全对 or (len(识别词对上) >= 2 and 达标)):
        if 弱候选收集 is not None:
            弱候选收集.append({"小表编号": 小表["小表编号"], "小表名称": 小表["小表名称"],
                           "新工作表": 工作表名, "种子偏移": [dr, dc],
                           "对上文字数": 对上数, "识别性文字对上数": len(识别词对上),
                           "旧文字节点数": len(旧文字节点),
                           "原因": "无识别性文字对上" if not 锚对上 else "对上数量不足"})
        return None
    # 候选成立后才展开新增行；未达标的种子只需留下弱候选依据。
    对齐序列 = sorted(行映射.items())
    已用行 = set(行映射.values())
    for (前旧, 前新), (后旧, 后新) in zip([(None, 旧行们[0][0] + dr - 1)] + 对齐序列, 对齐序列 + [(None, 对齐序列[-1][1] + 1)]):
        if 前旧 is not None and 后旧 is not None and 后旧 - 前旧 == 1 and 后新 - 前新 == 1:
            continue
        for 跳过 in range(前新 + 1, 后新):
            if 跳过 in 已用行:
                continue
            行文字 = []
            for 列 in range(1, 表数据["记录"]["max_column"] + 1):
                t = 新文字(跳过, 列)
                if t:
                    坐标 = f"{get_column_letter(列)}{跳过}"
                    原文 = 表数据["原文"].get(坐标, t)
                    行文字.append((列, t, 原文))
            if 行文字:
                新增行.append({"行": 跳过, "文字": 行文字})
    新范围坐标 = [d["新坐标"] for d in 结构对应]
    新范围坐标 += [f"{get_column_letter(列)}{e['行']}" for e in 新增行 for 列, _, _ in e["文字"]]
    if not 新范围坐标:
        return None
    盒们 = [range_boundaries(c) for c in 新范围坐标]
    左, 顶, 右, 底2 = min(b[0] for b in 盒们), min(b[1] for b in 盒们), max(b[2] for b in 盒们), max(b[3] for b in 盒们)
    新表范围 = f"{get_column_letter(左)}{顶}:{get_column_letter(右)}{底2}"

    载荷位置 = {d["新坐标"] for d in 结构对应 if d["角色"] == "业务数值"}
    范围内文字 = []
    范围内数字 = []
    for 行 in range(顶, 底2 + 1):
        for 列 in range(左, 右 + 1):
            坐标 = f"{get_column_letter(列)}{行}"
            if 坐标 in 载荷位置:
                continue
            if 坐标 in 显示:
                范围内文字.append(坐标)
            elif 坐标 in 表数据["数字坐标"]:
                范围内数字.append(坐标)
    对上新文字 = {d["新坐标"] for d in 结构对应 if d.get("查找文字")}
    新增项目 = []
    for e in 新增行:
        标签 = next((原 for 列, _, 原 in e["文字"]), None)
        if 标签:
            新增项目.append(标签)

    新增影响 = []
    if 新增项目:
        对齐项目 = [(n, 行映射.get(n["行"])) for n in 小表["节点"] if n["角色"] == "项目" and n["行"] in 行映射]
        for e, 原文 in zip(新增行, 新增项目):
            前行 = [n for n, q in 对齐项目 if q and q < e["行"]]
            后合计 = [n for n, q in 对齐项目 if q and q > e["行"] and n.get("原文") and 规范文字(str(n["原文"])).startswith("合计")]
            影响 = {"新增": 原文, "行": e["行"]}
            if 前行:
                影响["父项候选"] = 前行[-1].get("原文")
            if 后合计:
                影响["说明"] = f"新增行位于合计行'{后合计[0].get('原文')}'之前，可能影响父项目与合计口径，待通读确认"
            else:
                影响["说明"] = "新增行影响待通读确认，不能自动认定无影响"
            新增影响.append(影响)

    冲突 = [d for d in 差异 if d["类型"] == "列位置变化"]
    顺序冲突 = []
    项目行序 = [(n["行"], 行映射.get(n["行"])) for n in sorted(小表["节点"], key=lambda n: n["行"])
              if n["角色"] == "项目" and n["行"] in 行映射]
    for (前行, 前新), (后行, 后新) in zip(项目行序, 项目行序[1:]):
        if 后新 <= 前新:
            顺序冲突.append({"旧行": [前行, 后行], "说明": "项目先后顺序与新表不一致"})
    合并核对 = []
    预期合并 = set()
    for m in 小表.get("合并", []):
        ml, mt, mr, mb = range_boundaries(m)
        if mt in 行映射 and mb in 行映射 and ml in 列映射 and mr in 列映射:
            目标 = (列映射[ml], 行映射[mt], 列映射[mr], 行映射[mb])
            预期合并.add(目标)
            一致 = 目标 in {(b[0], b[1], b[2], b[3]) for b in 表数据["合并"]}
            合并核对.append({"旧合并": m, "结果": "一致" if 一致 else "冲突"})
        else:
            合并核对.append({"旧合并": m, "结果": "缺依据"})
    # 反向核对：候选范围内新表新增的合并（旧表没有对应），不得静默放过
    for b in 表数据["合并"]:
        if b[0] <= 右 and b[2] >= 左 and b[1] <= 底2 and b[3] >= 顶:
            if (b[0], b[1], b[2], b[3]) not in 预期合并:
                合并核对.append({"新合并": f"{get_column_letter(b[0])}{b[1]}:{get_column_letter(b[2])}{b[3]}",
                              "结果": "新增合并，旧表无对应"})

    上下文核对 = {"一致": [], "冲突": [], "缺依据": [], "冲突原因": []}
    # 共用表头：外部所属依据独立核对，不计入小表正文覆盖
    共用表头核对 = {"一致": [], "冲突": [], "缺依据": []}
    for h in 小表.get("表头", []):
        if not h.get("共用"):
            continue
        c, r, _, _ = range_boundaries(h["坐标"])
        新值 = 新文字(r + dr, c + dc)
        if 新值 == h["查找文字"]:
            共用表头核对["一致"].append(h["坐标"])
        elif 新值:
            共用表头核对["冲突"].append({"旧坐标": h["坐标"], "旧": h["查找文字"], "新": 新值})
        else:
            共用表头核对["缺依据"].append(h["坐标"])
    for ctx in 小表.get("上下文", []):
        c, r, _, _ = range_boundaries(ctx["坐标"])
        期望行, 期望列 = r + dr, c + dc
        新值 = 新文字(期望行, 期望列)
        if 新值 == ctx["查找文字"]:
            上下文核对["一致"].append(ctx["坐标"])
        elif 新值:
            上下文核对["冲突"].append({"旧坐标": ctx["坐标"], "旧": ctx["查找文字"], "新": 新值})
            上下文核对["冲突原因"].append(f"上级标题变化：旧'{ctx['查找文字']}'在新表{期望行}行位置为'{新值}'")
        else:
            附近 = [q for q in range(max(1, 期望行 - 3), 期望行 + 1)
                   if any(新文字(q, cc) == ctx["查找文字"] for cc in range(1, 表数据["记录"]["max_column"] + 1))]
            if 附近:
                上下文核对["一致"].append(ctx["坐标"])
            else:
                上下文核对["缺依据"].append(ctx["坐标"])

    return {
        "小表名称": 小表["小表名称"], "旧表编号": 小表["小表编号"], "样式来源": 小表["样式"],
        "新工作表": 工作表名, "新表范围": 新表范围,
        "种子偏移": [dr, dc],
        "结构对应": 结构对应,
        "新增项目原文": 新增项目, "新增影响": 新增影响,
        "缺失项目原文": [规范文字(str(n.get("原文"))) for n in 小表["节点"]
                     if n["角色"] == "项目" and n["行"] in {r for r, _ in 缺失行}],
        "差异": 差异,
        "冲突": 冲突,
        "关系核对": {"顺序": {"已核对": len(项目行序) - len(顺序冲突), "冲突": 顺序冲突},
                   "合并": 合并核对, "变化": 冲突, "层级": {"冲突": 层级冲突},
                   "同列": {get_column_letter(k): get_column_letter(v) for k, v in 列映射.items()}},
        "上下文核对": 上下文核对,
        "共用表头核对": 共用表头核对,
        "覆盖计算": {
            "旧表文字覆盖": {"分子": len(旧覆盖分子), "分母": len(旧文字节点), "对上项目": 旧覆盖分子,
                        "未对上项目": sorted(n["坐标"] for n in 旧文字节点 if n["坐标"] not in 对上旧文字)},
            "新表文字覆盖": {"分子": len(对上新文字), "分母": len(范围内文字),
                        "未对上项目": sorted(c for c in 范围内文字 if c not in 对上新文字)},
            "未知用途数字": {"数量": len(范围内数字), "坐标": sorted(范围内数字),
                        "说明": "文字覆盖比例不衡量这些数字"},
        },
        "歧义": 歧义,
        "语义引用说明": "结构对应沿用旧语义引用，均为待通读确认，不构成新表已批准语义",
    }


def 比较工作簿(工作簿记录: dict, 索引: dict, 预筛选: bool = True) -> dict:
    起始 = time.perf_counter()
    global 索引正文计数
    索引正文计数 = {t: v for t, v in 索引["正文"].items()}
    结果 = {"输入版本": {"路径": 工作簿记录.get("path"), "sha256": 工作簿记录.get("sha256")},
          "比较规则版本": 比较规则版本,
          "索引版本": {"目录": 索引["目录"], "规则版本": 索引["规则版本"]},
          "工作表": [], "候选": [], "弱候选": [], "未匹配范围": [], "缺口": [],
          "样式命中": {}, "运行统计": {}}
    比较小表数 = 0
    预筛选秒 = 0.0
    对齐秒 = 0.0
    for 表 in 工作簿记录["sheets"]:
        表数据 = _工作表数据(表)
        结果["工作表"].append({"名称": 表["name"], "隐藏": 表.get("hidden", False),
                          "内容格数": len(表数据["内容格"]), "文字种类": len(表数据["文字位置"]),
                          "隐藏行": 表.get("hidden_rows", []), "隐藏列": 表.get("hidden_columns", [])})
        if 预筛选:
            t0 = time.perf_counter()
            命中计数 = defaultdict(int)
            强线索 = defaultdict(int)
            for 文字 in 表数据["文字位置"]:
                for 条目 in 索引["正文"].get(文字, []):
                    命中计数[条目["小表编号"]] += 1
                    if len(索引["正文"][文字]) <= 通用词频上限:
                        强线索[条目["小表编号"]] += 1
            待比 = [编号 for 编号, n in 命中计数.items() if (n >= 2 and 强线索[编号] >= 1) or 强线索[编号] >= 2]
            # 候选成立允许"全部对上"（不要求识别性文字），预筛选不得把它挡在门外：
            # 登记文字全部出现在新表的小表，即使全是通用词也纳入待比
            for 编号, n in 命中计数.items():
                if 编号 in 待比:
                    continue
                文字们 = [nd["查找文字"] for nd in 索引["小表"][编号]["节点"] if nd.get("查找文字")]
                if len(文字们) >= 2 and all(t in 表数据["文字位置"] for t in 文字们):
                    待比.append(编号)
            # 别名只作扩展入口：其文字对应的旧小表也纳入待比
            for 文字 in 表数据["文字位置"]:
                for 条目 in 索引.get("别名", {}).get(文字, []):
                    if 条目["小表编号"] not in 待比:
                        待比.append(条目["小表编号"])
            预筛选秒 += time.perf_counter() - t0
        else:
            待比 = list(索引["小表"].keys())
        比较小表数 += len(待比)
        候选们 = []
        t1 = time.perf_counter()
        for 编号 in 待比:
            小表 = 索引["小表"][编号]
            种子们 = _提候选范围(小表, 表数据)
            # 种子按 (dr,dc) 唯一、按支持度排序，逐一尝试，不按固定数量截断独立位置；
            # 只有全部文字逐格对应且偏移相同，才证明该二维位置已经比较过。
            # 仅凭候选范围包含预测行段，会漏掉左右并排及范围内的独立位置。
            已发现 = set()
            for 种子 in 种子们:
                if 种子[:2] in 已发现:
                    continue
                候选 = _对齐候选(小表, 表["name"], 表数据, 种子, 结果["弱候选"])
                if 候选:
                    候选们.append(候选)
                    覆盖 = 候选["覆盖计算"]["旧表文字覆盖"]
                    if 覆盖["分子"] == 覆盖["分母"]:
                        偏移 = set()
                        for d in 候选["结构对应"]:
                            if d.get("查找文字"):
                                旧列, 旧行, _, _ = range_boundaries(d["旧坐标"])
                                新列, 新行, _, _ = range_boundaries(d["新坐标"])
                                偏移.add((新行 - 旧行, 新列 - 旧列))
                        if len(偏移) == 1:
                            已发现.update(偏移)
        候选们 = _合并候选(候选们)
        对齐秒 += time.perf_counter() - t1
        结果["候选"].extend(候选们)
        # 未匹配：候选范围之外的内容格
        覆盖盒 = []
        for c in 候选们:
            覆盖盒.append(range_boundaries(c["新表范围"]))
        未覆盖 = []
        for 坐标 in 表数据["内容格"]:
            c, r, _, _ = range_boundaries(坐标)
            if not any(b[0] <= c <= b[2] and b[1] <= r <= b[3] for b in 覆盖盒):
                未覆盖.append(坐标)
        if 未覆盖:
            行组 = defaultdict(list)
            for 坐标 in 未覆盖:
                c, r, _, _ = range_boundaries(坐标)
                行组[r].append(坐标)
            段 = []
            for r in sorted(行组):
                if 段 and r == 段[-1]["行止"] + 1:
                    段[-1]["行止"] = r
                    段[-1]["坐标清单"].extend(sorted(行组[r]))
                else:
                    段.append({"行起": r, "行止": r, "坐标清单": sorted(行组[r])})
            for s in 段:
                s["坐标清单"].sort(key=lambda a: (range_boundaries(a)[1], range_boundaries(a)[0]))
                s["坐标数"] = len(s["坐标清单"])
                s["原文摘录"] = [表数据["显示"].get(c) or str((表["cells"][c].get("cached_value") if 表["cells"][c].get("formula") else 表["cells"][c].get("value")))[:40]
                              for c in s["坐标清单"][:8]]
                s["范围"] = f"{get_column_letter(min(range_boundaries(c)[0] for c in s['坐标清单']))}{s['行起']}:{get_column_letter(max(range_boundaries(c)[0] for c in s['坐标清单']))}{s['行止']}"
            结果["未匹配范围"].append({"工作表": 表["name"], "隐藏": 表.get("hidden", False), "分段": 段,
                                  "坐标数": len(未覆盖)})
    for i in range(1, 6):
        命中 = [c for c in 结果["候选"] if c["样式来源"] == i]
        结果["样式命中"][str(i)] = {"候选数": len(命中),
                              "命中小表": sorted({c["小表名称"] for c in 命中}),
                              "未命中说明": "该样式无候选命中" if not 命中 else ""}
    结果["运行统计"] = {"比较的小表数": 比较小表数, "候选数": len(结果["候选"]),
                   "阶段耗时": {"预筛选秒": round(预筛选秒, 3), "候选对齐秒": round(对齐秒, 3)},
                   "耗时秒": round(time.perf_counter() - 起始, 3)}
    return 结果


def _合并候选(候选们):
    """同旧小表、共享新坐标的候选是同一处发现的重复种子，合并保留最大证据；范围不交的保留为多个候选。"""
    按表 = defaultdict(list)
    for c in 候选们:
        按表[c["旧表编号"]].append(c)
    结果 = []
    for 编号, 组 in 按表.items():
        组.sort(key=lambda c: -len(c["结构对应"]))
        保留 = []
        for c in 组:
            坐标集 = {d["新坐标"] for d in c["结构对应"]}
            并入 = None
            for r in 保留:
                已有 = {d["新坐标"] for d in r["结构对应"]}
                if 坐标集 & 已有:
                    并入 = r
                    break
            if 并入 is None:
                保留.append(c)
            else:
                if len(c["结构对应"]) > len(并入["结构对应"]):
                    c.setdefault("歧义", []).append({"原因": "同一种小表的多个种子偏移收敛到同一范围，已合并", "合并前候选数": len(组)})
                    保留[保留.index(并入)] = c
                else:
                    并入.setdefault("歧义", []).append({"原因": "同一种小表的多个种子偏移收敛到同一范围，已合并", "合并前候选数": len(组)})
        结果.extend(保留)
    return 结果


def 核验预期对应(结果, 预期, 小表名称):
    """独立核对：候选的结构对应必须落在事先固定的变换预期内。"""
    候选 = [c for c in 结果["候选"] if c["小表名称"] == 小表名称]
    assert len(候选) == 1, f"候选数量不是1：{len(候选)}"
    对应 = {(d["旧工作表"], d["旧坐标"]): (d["新工作表"], d["新坐标"]) for d in 候选[0]["结构对应"]}
    for k, v in 对应.items():
        assert k in 预期, f"候选包含预期外旧坐标{k}"
        assert v == 预期[k], f"{k}对应{v}与预期{预期[k]}不符"
    return True


def 完整性检查(工作簿记录, 初筛结果):
    """以工作簿读取结果为底稿：每个实际内容格必须在候选范围或未匹配清单中可回查。"""
    问题 = []
    候选盒 = defaultdict(list)
    for c in 初筛结果["候选"]:
        候选盒[c["新工作表"]].append(range_boundaries(c["新表范围"]))
    未匹配坐标 = defaultdict(set)
    for m in 初筛结果["未匹配范围"]:
        for s in m["分段"]:
            未匹配坐标[m["工作表"]].update(s["坐标清单"])
    for 表 in 工作簿记录["sheets"]:
        名 = 表["name"]
        for 坐标, 格 in 表["cells"].items():
            if 格.get("value") is None and not 格.get("formula"):
                continue
            c, r, _, _ = range_boundaries(坐标)
            if any(b[0] <= c <= b[2] and b[1] <= r <= b[3] for b in 候选盒.get(名, [])):
                continue
            if 坐标 in 未匹配坐标.get(名, set()):
                continue
            问题.append(f"{名}!{坐标} 既不在候选范围也不在未匹配清单")
    if 问题:
        raise AssertionError("完整性检查失败：\n" + "\n".join(问题[:20]))
    return True


def 写出交接材料(工作簿记录: dict, 初筛结果: dict, 输出目录: str) -> None:
    输出目录 = Path(输出目录)
    if 输出目录.exists():
        raise ValueError(f"输出目录已存在，拒绝覆盖：{输出目录}")
    输出目录.mkdir(parents=True)

    工作簿结构 = {"说明": "完整工作簿结构材料；实际内容位置均可按坐标回查，未截断",
              "路径": 工作簿记录.get("path"), "sha256": 工作簿记录.get("sha256"),
              "外部引用": 工作簿记录.get("external_references", []),
              "工作表": [{"名称": s["name"], "隐藏": s.get("hidden", False),
                       "隐藏行": s.get("hidden_rows", []), "隐藏列": s.get("hidden_columns", []),
                       "最大行": s.get("max_row"), "最大列": s.get("max_column"),
                       "合并": s.get("merges", []),
                       "单元格": {k: {kk: vv for kk, vv in v.items() if kk in
                                   ("row", "column", "value", "formula", "cached_value", "hidden")}
                                for k, v in s["cells"].items()}}
                      for s in 工作簿记录["sheets"]]}
    (输出目录 / "工作簿结构.json").write_text(
        json.dumps(工作簿结构, ensure_ascii=False, indent=1, default=_json默认), encoding="utf-8-sig")

    with open(输出目录 / "初筛结果.json", "w", encoding="utf-8-sig", newline="\n") as f:
        json.dump(初筛结果, f, ensure_ascii=False, indent=1, default=_json默认)

    (输出目录 / "初筛报告.md").write_text(_生成报告(初筛结果), encoding="utf-8-sig")
    (输出目录 / "LLM交接说明.md").write_text(_生成交接说明(初筛结果), encoding="utf-8-sig")

    完整性 = "通过"
    try:
        完整性检查(工作簿记录, 初筛结果)
    except AssertionError as e:
        完整性 = str(e)
    记录 = {"输入": 初筛结果["输入版本"], "索引版本": 初筛结果["索引版本"],
          "规则版本": 初筛结果["索引版本"]["规则版本"], "运行统计": 初筛结果["运行统计"],
          "完整性结论": 完整性}
    (输出目录 / "运行记录.json").write_text(
        json.dumps(记录, ensure_ascii=False, indent=1, default=_json默认), encoding="utf-8-sig")
    if 完整性 != "通过":
        raise AssertionError(完整性)


def _生成报告(结果):
    行 = ["# 初筛报告", "",
        f"- 输入：{结果['输入版本']['路径']}（{结果['输入版本']['sha256'][:12]}…）",
        f"- 索引：{结果['索引版本']['目录']}（规则 {结果['索引版本']['规则版本']}）",
        f"- 候选数：{结果['运行统计']['候选数']}；比较的小表数：{结果['运行统计']['比较的小表数']}", "",
        "## 五样式命中情况", ""]
    for 样式, 情况 in sorted(结果["样式命中"].items()):
        行.append(f"- 样式{样式}：候选{情况['候选数']}个" +
                (f"，命中：{'、'.join(情况['命中小表'][:8])}" if 情况["命中小表"] else "，无命中"))
    行 += ["", "## 候选明细", ""]
    for c in 结果["候选"]:
        旧 = c["覆盖计算"]["旧表文字覆盖"]
        新 = c["覆盖计算"]["新表文字覆盖"]
        问题 = _候选问题(c)
        行 += [f"### {c['小表名称']}（样式{c['样式来源']}，{c['新工作表']}!{c['新表范围']}）",
             f"- 旧表文字覆盖 {旧['分子']}/{旧['分母']}；新表候选范围文字覆盖 {新['分子']}/{新['分母']}",
             f"- 新增：{c['新增项目原文'] or '无'}；缺失：{c['缺失项目原文'] or '无'}",
             f"- 差异：{len(c['差异'])}项；冲突：{len(c['冲突'])}项；歧义：{len(c['歧义'])}项",
             f"- 上下文：一致{len(c['上下文核对']['一致'])}，冲突{len(c['上下文核对']['冲突'])}，缺依据{len(c['上下文核对']['缺依据'])}",
             f"- 待处理问题：{'；'.join(问题) if 问题 else '无'}",
             ""]
        for d in c["差异"][:10]:
            行.append(f"  - 差异 {d['旧坐标']}：{d['说明']}")
        for x in c["新增影响"][:10]:
            行.append(f"  - 新增影响：{x['新增']}（{x['说明']}）")
        if c["上下文核对"]["冲突原因"]:
            for x in c["上下文核对"]["冲突原因"][:5]:
                行.append(f"  - {x}")
    行 += ["", "## 未匹配范围", ""]
    for m in 结果["未匹配范围"]:
        行.append(f"- 工作表 {m['工作表']}：{m['坐标数']}格未匹配" + ("（隐藏表）" if m["隐藏"] else ""))
        for s in m["分段"][:10]:
            行.append(f"  - {s['范围']}（{s['坐标数']}格）：{'；'.join(str(x) for x in s['原文摘录'][:4])}")
    行.append("")
    行.append("说明：以上比例是文字覆盖情况，不是业务含义正确率；结构对应带旧语义引用，均为待通读确认。")
    return "\n".join(行)


def _候选问题(c):
    """一个候选全部需要后续处理的问题，JSON/报告/交接说明共用同一份判定。"""
    问题 = []
    if c["缺失项目原文"]:
        问题.append(f"缺失项目{len(c['缺失项目原文'])}项：{c['缺失项目原文']}")
    if c["差异"]:
        按类 = defaultdict(int)
        for d in c["差异"]:
            按类[d["类型"]] += 1
        问题.append("差异" + "、".join(f"{t}{n}项" for t, n in sorted(按类.items())))
    if c["新增影响"]:
        问题.append(f"新增影响{len(c['新增影响'])}项")
    if c["冲突"]:
        问题.append(f"列位置冲突{len(c['冲突'])}项")
    if c["歧义"]:
        问题.append(f"歧义{len(c['歧义'])}项")
    if c["上下文核对"]["冲突"]:
        问题.append(f"上下文冲突{len(c['上下文核对']['冲突'])}项")
    if c["上下文核对"]["缺依据"]:
        问题.append(f"上下文缺依据{len(c['上下文核对']['缺依据'])}项")
    合并未一致 = [m for m in c["关系核对"]["合并"] if m["结果"] != "一致"]
    if 合并未一致:
        问题.append(f"合并核对未一致{len(合并未一致)}项")
    if c["关系核对"].get("层级", {}).get("冲突"):
        问题.append(f"层级冲突{len(c['关系核对']['层级']['冲突'])}项")
    if c["关系核对"]["顺序"]["冲突"]:
        问题.append(f"顺序冲突{len(c['关系核对']['顺序']['冲突'])}项")
    共用 = c.get("共用表头核对", {})
    if 共用.get("冲突") or 共用.get("缺依据"):
        问题.append(f"共用表头冲突{len(共用.get('冲突', []))}项、缺依据{len(共用.get('缺依据', []))}项")
    return 问题


def _生成交接说明(结果):
    未解决 = {id(c): _候选问题(c) for c in 结果["候选"]}
    未解决 = {k: v for k, v in 未解决.items() if v}
    行 = ["# 后续LLM通读交接说明", "",
        "## 使用方法", "",
        "1. 先通读 `工作簿结构.json` 的完整新表结构与 `初筛结果.json` 的全部候选、差异、未匹配范围。",
        "2. 候选中的结构对应带有旧小表的语义引用（representation_id / semantic_id），可在统一成果库回查；"
        "这些是待确认线索，不是新表已批准语义。",
        "3. 只细读下列仍未解决的部分；通读已能确认的不要重复识别。",
        "4. 对查不到的部分，查阅已有标准区分：同义的新表现、已有维度的新取值、新增维度、真正新增语义；查不到绝不自动新增槽位。",
        "",
        "## 需要进一步处理的范围", ""]
    if 未解决 or 结果["未匹配范围"]:
        行.append(f"状态：存在必须细读的范围（候选未解决问题{len(未解决)}个，未匹配范围{len(结果['未匹配范围'])}段）。")
        行.append("")
    if 未解决:
        行.append("### 候选内未解决问题")
        for c in 结果["候选"]:
            if id(c) in 未解决:
                行.append(f"- {c['小表名称']}（{c['新工作表']}!{c['新表范围']}）：")
                for p in 未解决[id(c)]:
                    行.append(f"  - {p}")
    if 结果["未匹配范围"]:
        行.append("")
        行.append("### 未匹配范围（无任何旧参照，必须细读）")
        for m in 结果["未匹配范围"]:
            行.append(f"- {m['工作表']}：{m['坐标数']}格，分段见 初筛结果.json")
    if not 未解决 and not 结果["未匹配范围"]:
        行.append("本次没有未解决问题（全部候选无差异、无冲突、无歧义、无合并/层级/上下文/共用表头问题，"
                "且无未匹配范围），不强制产生全表细读任务；通读确认候选边界后即可复用。")
    return "\n".join(行) + "\n"
