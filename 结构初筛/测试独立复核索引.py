# -*- coding: utf-8 -*-
"""自包含的小样本回归；全部资料虚构，临时证据保留，不自动删除。"""
import copy
import json
import tempfile
import unittest
from collections import Counter, defaultdict
from pathlib import Path
from unittest.mock import patch

import 建立结构索引 as 库


def 写json(路径, 对象):
    路径.write_text(json.dumps(对象, ensure_ascii=False, indent=2), encoding="utf-8-sig")


def 小表(样式=1, 名称="虚构表"):
    return {"小表编号": f"虚构{样式}", "样式": 样式, "小表名称": 名称,
            "来源": {"工作表": "虚构工作表"}, "节点": [
                {"坐标": "A1", "行": 1, "列": 1, "角色": "标题", "原文": 名称, "查找文字": 名称},
                {"坐标": "B2", "行": 2, "列": 2, "角色": "业务数值", "原文": None,
                 "语义引用": {"mapping_id": f"m{样式}", "representation_id": f"r{样式}", "semantic_id": f"s{样式}",
                            "维度": {}, "确认依据": {"说明": "虚构"}}}]}


def 写索引(目录, 表们=None, 调整=None, 版本=None):
    目录.mkdir()
    表们 = 表们 or [小表()]
    正文 = defaultdict(list)
    去向 = []
    for t in 表们:
        for n in t["节点"]:
            if n.get("查找文字"):
                正文[n["查找文字"]].append({"小表编号": t["小表编号"], "样式": t["样式"],
                                         "工作表": t["来源"]["工作表"], "坐标": n["坐标"], "角色": n["角色"]})
            if n["角色"] == "业务数值":
                去向.append({"样式": t["样式"], "mapping_id": n["语义引用"]["mapping_id"], "去向": "业务格进入索引"})
    版本 = 版本 or 库.规则版本
    清单 = {"规则版本": 版本, "小表数": len(表们), "按样式": dict(Counter(str(t["样式"]) for t in 表们)),
          "映射去向": dict(Counter(d["去向"] for d in 去向)), "输入清单": "不存在的原作者路径.json", "缺口": []}
    查找 = {"规则版本": 版本, "正文": dict(正文), "别名": {}}
    if 调整:
        调整(清单, 查找, 去向, 表们)
    (目录 / "小表记录.jsonl").write_text("".join(json.dumps(t, ensure_ascii=False) + "\n" for t in 表们), encoding="utf-8-sig")
    (目录 / "去向明细.jsonl").write_text("".join(json.dumps(t, ensure_ascii=False) + "\n" for t in 去向), encoding="utf-8-sig")
    写json(目录 / "文字查找.json", 查找)
    清单.setdefault("生成物指纹", 库.索引文件指纹(目录))
    if 版本 == "结构初筛规则v4":
        清单.pop("生成物指纹")
    写json(目录 / "索引清单.json", 清单)
    return 目录


class 索引复核(unittest.TestCase):
    def setUp(self):
        self.根 = Path(tempfile.mkdtemp(prefix="初筛索引独立复核_"))

    def test_便携索引可核内部完整性且明示来源限度(self):
        r = 库.校验索引目录(写索引(self.根 / "正常"))
        self.assertEqual(len(r["记录"]), 1)
        self.assertTrue(any("不代表重新审定" in s for s in r["校验说明"]))

    def test_生成物指纹错误拒绝载入(self):
        p = 写索引(self.根 / "错误", 调整=lambda c, *_: c.update(生成物指纹={"小表记录.jsonl": "0" * 64}))
        with self.assertRaisesRegex(ValueError, "指纹"):
            库.校验索引目录(p)

    def test_旧版也拒绝截断小表和错误去向(self):
        for i, 改法 in enumerate((lambda c, *_: c.update(小表数=2),
                                 lambda c, d, 去向, *_: 去向[0].update(mapping_id="不是该业务格"))):
            p = 写索引(self.根 / str(i), 调整=改法, 版本="结构初筛规则v4")
            with self.assertRaisesRegex(ValueError, "数量|映射身份"):
                库.校验索引目录(p)

    def test_规则冲突和文字目录错位被拒绝(self):
        改法们 = (lambda c, d, *_: d.update(规则版本="冲突版本"),
                 lambda c, d, *_: d["正文"]["虚构表"][0].update(坐标="Z99"))
        for i, 改法 in enumerate(改法们):
            with self.assertRaisesRegex(ValueError, "版本|逐项对应"):
                库.校验索引目录(写索引(self.根 / str(i), 调整=改法))

    def test_断开的业务引用被拒绝(self):
        def 改法(c, d, g, ts):
            ts[0]["节点"][1]["语义引用"].pop("representation_id")
        with self.assertRaisesRegex(ValueError, "语义引用不完整"):
            库.校验索引目录(写索引(self.根 / "断开", 调整=改法))

    def test_同词表头与项目分开且父合并覆盖子列(self):
        t, gaps = self.提取示例()
        按格 = {n["坐标"]: n for n in t["节点"]}
        self.assertEqual(按格["C2"]["角色"], "表头")
        self.assertEqual(按格["A3"]["角色"], "项目")
        self.assertTrue(any(h["坐标"] == "B1" and h["覆盖列"] == "C" for h in t["表头"]))
        self.assertNotIn("C1", t["空白位置"])
        self.assertFalse(any(g["问题"] == "该业务列缺表头归属依据" for g in gaps))

    def test_合计后的其中不擅自认定父项目(self):
        t, gaps = self.提取示例(其中=True)
        其中 = next(n for n in t["项目顺序"] if n["坐标"] == "A5")
        self.assertIsNone(其中["上级"])
        self.assertEqual(其中["候选上级"], "A4")
        self.assertTrue(any(g.get("影响业务格") == ["B5", "C5"] for g in gaps))

    def test_公式表头缺缓存保留公式及缺口(self):
        t, gaps = self.提取示例(无缓存=True)
        n = next(n for n in t["节点"] if n["坐标"] == "C2")
        self.assertEqual(n["公式"], '="其他"')
        self.assertIsNone(n["查找文字"])
        self.assertTrue(any(g["问题"] == "公式结构格缺少缓存显示" and "C3" in g["影响业务格"] for g in gaps))

    def test_合并父表头无缓存影响全部覆盖业务列(self):
        t, gaps = self.提取示例(合并无缓存=True)
        问题 = next(g for g in gaps if g.get("坐标") == "B1" and g["问题"] == "公式结构格缺少缓存显示")
        self.assertEqual(set(问题["影响业务格"]), {"B3", "C3"})
        self.assertEqual({h["覆盖列"] for h in t["表头"] if h["坐标"] == "B1"}, {"B", "C"})

    def test_业务区未映射公式不封锁同列其他业务格(self):
        cells = {a: {'row': 库.range_boundaries(a)[1], 'column': 库.range_boundaries(a)[0],
                     'value': v, 'formula': None, 'cached_value': v}
                 for a, v in {'A1':'项目','B1':'比例','A2':'甲','A3':'补充计算','A4':'乙'}.items()}
        cells['B3'] = {'row':3, 'column':2, 'value':'=B2/100', 'formula':'=B2/100', 'cached_value':None}
        maps = [{'mapping_id':a,'cell':a,'target_slot_id':a,'row_semantics':name,
                 'column_semantics':'比例','__文件':'虚构.jsonl'} for a,name in [('B2','甲'),('B4','乙')]]
        reps = {a:{'representation_id':a,'semantic_id':a,'规范维度值':{},'确认依据':{'说明':'虚构'}} for a in ('B2','B4')}
        gaps = []
        t = 库._提取小表(1,'虚构.xlsx','虚构','比例表',maps,(1,1,2,4),'明确范围',
                       {'cells':cells,'merges':[]},reps,'A',gaps,[])
        self.assertTrue(any(n.get('公式') == '=B2/100' for n in t['节点']))
        self.assertFalse(any(set(g.get('影响业务格',[])) & {'B2','B4'} for g in gaps))

    def 提取示例(self, 其中=False, 无缓存=False, 合并无缓存=False):
        cells = {}
        for a, v in {"A1": "项目", "B1": "本期增加", "B2": "计提", "C2": "其他", "A3": "其他",
                     **({"A4": "合计", "A5": "其中：某项"} if 其中 else {})}.items():
            row, col = 库.range_boundaries(a)[1], 库.range_boundaries(a)[0]
            cells[a] = {"row": row, "column": col, "value": v, "formula": None, "cached_value": v, "hidden": False}
        if 无缓存:
            cells["C2"].update(value='="其他"', formula='="其他"', cached_value=None)
        if 合并无缓存:
            cells["B1"].update(value='="本期增加"', formula='="本期增加"', cached_value=None)
        maps, reps = [], {}
        for r in (range(3, 6) if 其中 else [3]):
            for c in ("B", "C"):
                a = f"{c}{r}"
                maps.append({"mapping_id": a, "cell": a, "target_slot_id": a, "row_semantics": cells[f"A{r}"]["value"],
                             "column_semantics": "本期增加/其他", "__文件": "虚构.jsonl"})
                reps[a] = {"representation_id": a, "semantic_id": a, "规范维度值": {}, "确认依据": {"说明": "虚构"}}
        gaps = []
        t = 库._提取小表(1, "虚构.xlsx", "虚构", "虚构小表", maps, (1, 1, 3, 5 if 其中 else 3), "虚构明确范围",
                     {"cells": cells, "merges": ["B1:C1"]}, reps, "A", gaps, [])
        return t, gaps

    def test_源表改动只重提取对应样式且读表模块变化全部失效(self):
        entries = []
        for i in range(1, 6):
            p = self.根 / f"原件{i}.txt"
            p.write_text(f"虚构样式{i}", encoding="utf-8")
            entries.append({"用途": f"五样式原Excel:附注样式{i}.xlsx", "路径": str(p), "sha256": 库.文件指纹(p)})
        for name in ("表现记录", "统一语义记录"):
            p = self.根 / f"{name}.jsonl"
            p.write_text("", encoding="utf-8")
            entries.append({"用途": f"统一成果:{name}", "路径": str(p), "sha256": 库.文件指纹(p)})
        p = self.根 / "读取模块.py"
        p.write_text("# 虚构读取规则第一版\n", encoding="utf-8")
        entries.append({"用途": "复用代码:机械读取模块", "路径": str(p), "sha256": 库.文件指纹(p)})
        清单 = {"条目": entries}
        calls = []
        def 提取(style, path, *args, **kwargs):
            calls.append((style, kwargs.get("机械读取路径")))
            t = 小表(style, Path(path).read_text(encoding="utf-8"))
            t["节点"] = t["节点"][:1]
            return [t], [], []
        with patch.object(库, "_提取样式", side_effect=提取):
            库.建立索引(清单, self.根 / "首版")
            self.assertEqual([s for s, _ in calls], [1, 2, 3, 4, 5])
            # 本测试只改自己刚生成的虚构输入；旧内容由首版结构记录保留。
            Path(entries[0]["路径"]).write_text("虚构样式1增加内容", encoding="utf-8")
            entries[0]["sha256"] = 库.文件指纹(entries[0]["路径"])
            calls.clear()
            库.建立索引(清单, self.根 / "单表变更", self.根 / "首版")
            self.assertEqual([s for s, _ in calls], [1])
            Path(entries[-1]["路径"]).write_text("# 虚构读取规则第二版\n", encoding="utf-8")
            entries[-1]["sha256"] = 库.文件指纹(entries[-1]["路径"])
            calls.clear()
            库.建立索引(清单, self.根 / "规则变更", self.根 / "单表变更")
            self.assertEqual([s for s, _ in calls], [1, 2, 3, 4, 5])
            self.assertTrue(all(p == entries[-1]["路径"] for _, p in calls))


if __name__ == "__main__":
    unittest.main(verbosity=2)
