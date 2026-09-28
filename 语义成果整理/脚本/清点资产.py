# -*- coding: utf-8 -*-
"""清点语义资产（按2026-09-24修订版需求/设计）：固定本批输入，逐来源登记完整哈希、角色、版本、
确认依据位置和纳入/未纳入原因；拆清统计口径；核对五样式用作B端的来源关系。只读，不修改任何来源。"""
import hashlib
import json
import os
from collections import Counter
from datetime import datetime

项目 = r"C:\Users\27651\Desktop\附注自动更新系统"
数字资产 = r"D:\BaiduSyncdisk\数字资产\审计数据自动搬运系统"
输出 = os.path.join(项目, "语义成果整理", "资产清单.json")


def sha256(路径):
    h = hashlib.sha256()
    with open(路径, "rb") as f:
        for 块 in iter(lambda: f.read(1 << 20), b""):
            h.update(块)
    return h.hexdigest()


def main():
    资产 = []

    def 登记(路径, 角色, 纳入, 原因, 版本="", 确认依据=""):
        if not os.path.exists(路径):
            资产.append({"路径": 路径, "角色": 角色, "纳入": 纳入, "原因": "路径不存在", "状态": "缺失"})
            return None
        if os.path.isdir(路径):
            资产.append({"路径": 路径, "角色": 角色, "纳入": 纳入, "原因": 原因, "版本": 版本,
                         "确认依据": 确认依据, "状态": "目录"})
            return None
        资产.append({"路径": 路径, "角色": 角色, "纳入": 纳入, "原因": 原因, "版本": 版本,
                     "确认依据": 确认依据, "字节": os.path.getsize(路径), "sha256": sha256(路径),
                     "状态": "文件"})
        return 资产[-1]["sha256"]

    # ---- 主要输入：原金标准v1及来源证据 ----
    登记(os.path.join(数字资产, "语义金标准", "财务报表附注语义金标准-v1.jsonl"),
         "正式原件", "纳入", "原金标准v1唯一结构化正式数据源，语义与来源证据的主输入", "v1")
    登记(os.path.join(项目, "金标准", "财务报表附注语义金标准-v1.jsonl"),
         "副本", "纳入", "项目内v1副本，与正式原件同哈希只整理一次", "v1")
    登记(os.path.join(项目, "金标准", "五种样式人工映射", "语义金标准", "财务报表附注语义金标准-v1.jsonl"),
         "副本", "不重复整理", "五样式映射所附v1副本，哈希相同", "v1")

    # ---- 五样式最终映射 ----
    索引路径 = os.path.join(项目, "金标准", "五种样式人工映射", "语义示例索引.json")
    登记(索引路径, "正式", "纳入", "五样式语义示例索引（表现主输入）")
    for 样式 in range(1, 6):
        登记(os.path.join(数字资产, "work", "excel-mapping-v1", f"style{样式}-manual-review-v2"),
             "正式", "纳入", f"样式{样式}最终人工映射原文件（核对索引、补充原值状态）")
    for 样式 in range(1, 6):
        登记(os.path.join(数字资产, "附注样式示例", f"附注样式{样式}.xlsx"),
             "原始工作簿", "必要时回读", "五样式原表，索引缺字段时回读")
    登记(os.path.join(项目, "金标准", "五种样式人工映射", "work", "excel-mapping-v1"),
         "副本", "不重复整理", "项目内五样式映射目录，与D盘原件核对后去重")

    # ---- 选定新标准及版本链 ----
    版本目录 = os.path.join(项目, "金标准", "版本")
    配置版本 = "财务报表附注语义金标准_20260922_132731_236607_5547978f.jsonl"
    for f in sorted(os.listdir(版本目录)):
        p = os.path.join(版本目录, f)
        if f.endswith(".jsonl"):
            登记(p, "当前版本" if f == 配置版本 else "历史版本",
                 "纳入" if f == 配置版本 else "追溯用",
                 "配置gold_path指向" if f == 配置版本 else "版本沿革追溯",
                 "v2", "变更记录链")
        elif f.startswith("变更记录"):
            资产.append({"路径": p, "角色": "版本链", "纳入": "追溯用", "原因": "版本继承依据",
                         "sha256": sha256(p), "状态": "文件"})
    登记(os.path.join(项目, "金标准", "财务报表附注语义金标准-v2.jsonl"),
         "历史快照", "追溯用", "v2较早快照（90指标/78维度），与当前版本哈希不同")
    登记(os.path.join(项目, "金标准", "识别快照"), "中间产物", "不纳入",
         "v1格式识别快照6个，其中一个首行解析失败，不影响正式来源")
    登记(os.path.join(项目, "金标准", "全量重整审查_20260920_1710"), "审查证据库", "纳入",
         "v2转换与定义审查的证据来源（107条source_example的出处）")
    资产.append({"路径": "（历史v1原件 e79babc0…）", "角色": "历史版本", "纳入": "追溯用",
                 "原因": "历史v1原件不在项目及D盘正式目录内；历史定义取自v2当前版本source_example内嵌定义，"
                        "其legacy_definition_hash已与内嵌原定义逐条重算核对一致",
                 "状态": "以可核实内嵌定义代替"})

    # ---- A端补充（原稿已列入，只核对有效补充） ----
    A目录 = os.path.join(项目, "测试结果", "真实语义识别续跑_A_20260922_050354_458720", "导入外部复核_20260922_051437_942503_b323c5")
    登记(os.path.join(A目录, "附注表格语义映射.json"), "正式", "纳入（补充核对）",
         "A端824条事实；逐条区分确认依据（双轮519/同义复核282/外部复核23）",
         "v2", "recognition_provenance.method逐条记录")
    登记(os.path.join(A目录, "外部复核确认记录.json"), "证据", "纳入（补充核对）",
         "A端外部复核确认记录（2格正式采纳等）")
    登记(os.path.join(A目录, "附注表格语义识别核对.xlsx"), "核对表", "追溯用", "A端逐格核对表")

    # ---- B端：按用户说明来自五样式之一，不单独收录 ----
    登记(os.path.join(项目, "测试结果", "真实语义识别续跑_B_20260922_071726_598567"),
         "B端运行产物", "不纳入", "按用户说明B端来自五样式之一，以五样式最终映射为依据，不追收各次检查点识别结果")
    登记(os.path.join(项目, "测试结果", "真实金融资产范围高级AI复核_20260921_233024_233186"),
         "B端运行产物", "不纳入", "同上；GUI历史选择452条，不另列一批")
    登记(r"C:\Users\27651\Desktop\附注数据示例\示例表格B.xlsx", "B端来源文件", "不纳入",
         "按用户说明为五样式之一的另存副本；哈希与五样式原表均不同，工作表结构未直接对应，以用户说明为来源关系依据")

    # ---- 五样式新系统核验材料 ----
    for d, 说明 in [("五样式固定资产首批语义_20260920_232856_243821", "未通过：维度未确认"),
                    ("五样式固定资产接续核验_20260920_234502_546341", "接续核验材料"),
                    ("五样式收入成本语义_20260921_024024_125189", "指标通过、维度不全"),
                    ("人工映射复用真实识别_20260922_131351", "复用验证材料"),
                    ("两表当前语义覆盖核查_20260922_050708_614215", "覆盖核查材料")]:
        登记(os.path.join(项目, "测试结果", d), "验证材料", "不纳入", 说明 + "；不否定五样式已有最终映射")

    # ---- 测试结果整体规模（按目录名模式归类，不深读） ----
    测试 = os.path.join(项目, "测试结果")
    模式 = Counter()
    for f in os.listdir(测试):
        if os.path.isdir(os.path.join(测试, f)):
            if f.startswith("金标准_"):
                模式["金标准快照"] += 1
            elif f.startswith("任务3_识别更新数据"):
                模式["任务3识别运行"] += 1
            else:
                模式["其他语义运行"] += 1
        else:
            模式["散文件"] += 1
    资产.append({"路径": 测试, "角色": "运行产物汇总", "纳入": "不纳入",
                 "原因": "历次运行中间产物，不按修改时间或目录名认定为有效成果", "状态": "目录",
                 "统计": dict(模式)})

    # ---- 配置（只记路径） ----
    cfg = json.load(open(os.path.join(项目, "用户配置.json"), encoding="utf-8-sig"))
    资产.append({"路径": os.path.join(项目, "用户配置.json"), "角色": "配置", "纳入": "参考",
                 "原因": f"gold_path={cfg.get('gold_path')}；a_mapping={cfg.get('a_mapping')}",
                 "说明": "只读取必要路径字段，未读取密钥", "状态": "文件"})

    # ---- 统计口径核对（直接读源文件重算） ----
    v1 = [json.loads(l) for l in open(os.path.join(项目, "金标准", "财务报表附注语义金标准-v1.jsonl"), encoding="utf-8-sig") if l.strip()]
    v1状态 = Counter(r["status"] for r in v1)
    v1来源 = sum(len(r.get("sources", [])) for r in v1)
    v1active来源 = sum(len(r.get("sources", [])) for r in v1 if r["status"] == "active")
    索引 = json.load(open(索引路径, encoding="utf-8-sig"))
    转换记录 = v2指标 = v2维度 = v2示例 = 0
    转换状态 = Counter()
    转换旧id = set()
    approved旧id = set()
    unresolved旧id = set()
    指标状态 = Counter()
    for l in open(os.path.join(版本目录, 配置版本), encoding="utf-8-sig"):
        r = json.loads(l)
        if r["record_type"] == "legacy_conversion":
            转换记录 += 1
            转换状态[r["status"]] += 1
            转换旧id.add(r["legacy_id"])
            (approved旧id if r["status"] == "approved" else unresolved旧id).add(r["legacy_id"])
        elif r["record_type"] == "metric":
            v2指标 += 1
            指标状态[r.get("status")] += 1
        elif r["record_type"] == "dimension":
            v2维度 += 1
        elif r["record_type"] == "source_example":
            v2示例 += 1
    引用路径 = {e["origin"]["mapping_file"] for e in 索引["examples"]}
    口径 = {
        "v1": {"总记录": len(v1), "状态分布": dict(v1状态), "来源对象": v1来源, "active来源对象": v1active来源},
        "五样式": {"表现数": len(索引["examples"]), "引用旧ID": len(索引["legacy_definitions"]),
                  "files条目": len(索引["files"]), "examples引用映射路径": len(引用路径),
                  "零表现映射文件": sorted(k for k in 索引["files"] if k.startswith("work/") and k not in 引用路径),
                  "各样式表现数": dict(Counter(str(e["style"]) for e in 索引["examples"]))},
        "v2当前版本": {"metric": v2指标, "指标状态": dict(指标状态), "dimension": v2维度,
                     "转换记录": 转换记录, "转换状态": dict(转换状态),
                     "转换涉及旧ID": len(转换旧id), "approved旧ID": len(approved旧id),
                     "unresolved旧ID": len(unresolved旧id),
                     "approved且unresolved交叉": len(approved旧id & unresolved旧id),
                     "仅unresolved旧ID": len(unresolved旧id - approved旧id),
                     "source_example": v2示例},
    }

    # ---- 固定输入清单（统一整理/测试/核验只按此清单读取并校验输入哈希） ----
    def 固定(路径):
        return {"路径": 路径, "sha256": sha256(路径)}
    五样式映射文件 = []
    for 相对路径, 期望 in 索引["files"].items():
        if 相对路径.startswith("work/"):
            全 = os.path.join(数字资产, 相对路径)
            五样式映射文件.append({"相对路径": 相对路径, "路径": 全, "sha256": sha256(全)})
    固定输入 = {
        "维度规范规则": 固定(os.path.join(项目, "语义成果整理", "对应", "维度规范对应.json")),
        "v1": 固定(os.path.join(项目, "金标准", "财务报表附注语义金标准-v1.jsonl")),
        "v2当前版本": 固定(os.path.join(版本目录, 配置版本)),
        "五样式索引": 固定(索引路径),
        "五样式原映射": 五样式映射文件,
        "A端映射": 固定(os.path.join(A目录, "附注表格语义映射.json")),
        "A端外部复核记录": 固定(os.path.join(A目录, "外部复核确认记录.json")),
        "历史v1定义证据": {"说明": "历史v1原件(e79babc0…)不在项目及D盘正式目录；历史定义取自v2当前版本"
                             "source_example内嵌定义，legacy_definition_hash已与内嵌原定义逐条重算核对一致",
                       "来源": "v2当前版本"},
    }

    结果 = {"生成时间": datetime.now().isoformat(timespec="seconds"),
            "说明": "只读清点，未修改任何来源；配置仅记录路径，未读取密钥字段。",
            "资产": 资产, "统计口径": 口径, "固定输入": 固定输入}
    os.makedirs(os.path.dirname(输出), exist_ok=True)
    with open(输出, "w", encoding="utf-8-sig", newline="\n") as f:
        json.dump(结果, f, ensure_ascii=False, indent=1, sort_keys=True)
    print(f"资产清单：{len(资产)} 项")
    print(json.dumps(口径, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
