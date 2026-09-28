# -*- coding: utf-8 -*-
"""运行初筛：用户入口。只读输入，每次比较新建带时间戳的运行目录，不覆盖旧结果。

用法：
    运行初筛.py                                      # 无参数：弹出Windows文件选择窗口（双击启动初筛.cmd同此）
    运行初筛.py 清点 --成果库 目录 --五样式目录 目录 --输出 输入清单.json
    运行初筛.py 建库 --输入清单 输入清单.json --输出 索引\首版 [--旧索引目录 索引\上一版]
    运行初筛.py 比较 --新表 新表.xlsx --索引目录 索引\首版 [--输出 目录] [--基准对照]
    运行初筛.py 选择文件 [--索引目录 索引\首版]      # 弹出Windows文件选择窗口后比较
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

本批目录 = Path(__file__).resolve().parent
sys.path.insert(0, str(本批目录))

默认成果库 = 本批目录.parent / "语义成果整理"
默认五样式目录 = 本批目录.parent / "参考资料" / "附注样式示例"
默认索引 = 本批目录 / "索引" / "首版"
新版索引 = 本批目录 / "索引" / "工作表语义版_正式"
if (新版索引 / "索引清单.json").is_file():
    默认索引 = 新版索引
运行结果根 = 本批目录 / "运行结果"


def _运行目录(输出=None):
    if 输出:
        return Path(输出)
    目录 = 运行结果根 / time.strftime("运行_%Y%m%d_%H%M%S")
    n = 2
    while 目录.exists():  # 同一秒内的连续运行不互相覆盖
        目录 = 运行结果根 / f"{time.strftime('运行_%Y%m%d_%H%M%S')}_{n}"
        n += 1
    return 目录


def 命令清点(参数):
    from 建立结构索引 import 清点
    清单 = 清点(参数.成果库, 参数.五样式目录, 参数.输出)
    print(f"清点完成：{len(清单['条目'])}个输入已登记到 {参数.输出}")
    return 0


def 命令建库(参数):
    from 建立结构索引 import 建立索引
    结果 = 建立索引(参数.输入清单, 参数.输出, 旧索引目录=参数.旧索引目录)
    print(f"建库完成：{结果['小表数']}个小表记录，映射去向{结果['映射去向']}，缺口{结果['缺口数']}条")
    print(f"索引目录：{结果['输出目录']}")
    return 0


def 执行比较(输入路径, 索引目录, 输出=None, 基准对照=False):
    from 建立结构索引 import 读取工作簿
    from 比较新表 import 载入索引, 比较工作簿, 写出交接材料
    输入路径 = Path(输入路径)
    if not 输入路径.is_file():
        print(f"输入文件不存在：{输入路径}", file=sys.stderr)
        return 2
    目录 = _运行目录(输出)
    t0 = time.perf_counter()
    记录 = 读取工作簿(str(输入路径))
    读取秒 = time.perf_counter() - t0
    索引 = 载入索引(str(索引目录))
    t1 = time.perf_counter()
    结果 = 比较工作簿(记录, 索引)
    比较秒 = time.perf_counter() - t1
    if 基准对照:
        基准 = 比较工作簿(记录, 索引, 预筛选=False)
        快速集合 = {(c["旧表编号"], c["新工作表"], c["新表范围"]) for c in 结果["候选"]}
        基准集合 = {(c["旧表编号"], c["新工作表"], c["新表范围"]) for c in 基准["候选"]}
        遗漏 = 基准集合 - 快速集合
        多出 = 快速集合 - 基准集合
        结果["基准对照"] = {"基准候选数": len(基准集合), "快速候选数": len(快速集合),
                       "基准比较小表数": 基准["运行统计"]["比较的小表数"],
                       "遗漏": sorted(遗漏), "多出": sorted(多出),
                       "结论": "一致" if 快速集合 == 基准集合 else "候选不一致"}
    结果["运行统计"].update({"读取秒": round(读取秒, 3), "比较秒": round(比较秒, 3)})
    t2 = time.perf_counter()
    写出交接材料(记录, 结果, str(目录))
    输出秒 = time.perf_counter() - t2
    统计 = 结果["运行统计"]
    print(f"比较完成：候选{统计['候选数']}个，比较小表{统计['比较的小表数']}张，"
          f"读取{统计['读取秒']}秒、比较{统计['比较秒']}秒、输出{输出秒:.3f}秒")
    print(f"未匹配范围：{len(结果['未匹配范围'])}段；弱候选：{len(结果.get('弱候选', []))}个")
    print(f"结果目录：{目录}")
    if 基准对照 and 结果["基准对照"]["结论"] != "一致":
        print("基准对照不一致，详见初筛结果中的遗漏和多出清单。", file=sys.stderr)
        return 1
    return 0


def 命令比较(参数):
    return 执行比较(参数.新表, 参数.索引目录, 参数.输出, 参数.基准对照)


def 命令选择文件(参数):
    import tkinter as tk
    from tkinter import filedialog
    根 = tk.Tk()
    根.withdraw()
    路径 = filedialog.askopenfilename(title="选择要初筛的新Excel文件",
                                  filetypes=[("Excel工作簿", "*.xlsx"), ("所有文件", "*.*")])
    索引目录 = 参数.索引目录
    if 路径 and not (Path(索引目录) / "索引清单.json").is_file():
        索引目录 = filedialog.askdirectory(title="选择已建立的结构索引目录（内有索引清单.json）")
    根.destroy()
    if not 路径 or not 索引目录:
        print("未选择文件，已取消。")
        return 1
    print(f"已选择：{路径}")
    return 执行比较(路径, 索引目录)


def main(argv=None):
    解析 = argparse.ArgumentParser(description="附注结构初筛：为完整新Excel查找旧小表候选并输出比对材料")
    子 = 解析.add_subparsers(dest="命令", required=True)

    p = 子.add_parser("清点", help="登记本批固定输入并生成输入清单")
    p.add_argument("--成果库", default=str(默认成果库))
    p.add_argument("--五样式目录", default=str(默认五样式目录))
    p.add_argument("--输出", default=str(本批目录 / "输入清单.json"))
    p.set_defaults(执行=命令清点)

    p = 子.add_parser("建库", help="从固定输入建立结构索引")
    p.add_argument("--输入清单", default=str(本批目录 / "输入清单.json"))
    p.add_argument("--输出", default=str(默认索引))
    p.add_argument("--旧索引目录", default=None)
    p.set_defaults(执行=命令建库)

    p = 子.add_parser("比较", help="对完整新Excel执行初筛比较")
    p.add_argument("--新表", required=True)
    p.add_argument("--索引目录", default=str(默认索引))
    p.add_argument("--输出", default=None, help="缺省在 运行结果\\运行_年月日_时分秒 新建")
    p.add_argument("--基准对照", action="store_true", help="相同上层语义规则下，关闭文字预筛选并核对候选集合一致")
    p.set_defaults(执行=命令比较)

    p = 子.add_parser("选择文件", help="弹出Windows选择窗口挑选新Excel后比较")
    p.add_argument("--索引目录", default=str(默认索引))
    p.set_defaults(执行=命令选择文件)

    if argv is None:
        argv = sys.argv[1:]
    if not argv:
        argv = ["选择文件"]  # 无参数（含双击.cmd）默认弹出文件选择窗口
    参数 = 解析.parse_args(argv)
    try:
        return 参数.执行(参数)
    except Exception as e:
        print(f"失败：{e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
