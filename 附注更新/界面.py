"""以五个可独立执行、可复用文件的步骤组织附注更新。"""
from __future__ import annotations

import copy
import json
import os
import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .设置 import DEFAULT_GOLD_PATH, DEFAULT_TIMEOUT, PROJECT_DIR, load_settings, save_settings, validate_settings

BG = "#EEF2F6"
TEXT = "#203047"
MUTED = "#65758A"
BLUE = "#245FC2"
GREEN = "#246B54"
BORDER = "#DDE5ED"

FILE_FIELDS = {
    "word_path": ("原始附注 Word", "*.docx"),
    "a_path": ("附注表格", "*.xlsx *.xlsm"),
    "b_path": ("更新数据表", "*.xlsx *.xlsm"),
    "a_link": ("附注表格与Word位置关联", "*.json"),
    "a_mapping": ("附注表格语义映射", "*.json"),
    "a_review": ("附注表格逐格语义核对表", "*.xlsx"),
    "a_standard_review": ("待处理语义与金标准补充材料", "*.json"),
    "b_mapping": ("更新数据语义映射", "*.json"),
    "b_review": ("更新数据逐格语义核对表", "*.xlsx"),
    "b_scope": ("更新数据识别范围与章外清单", "*.json"),
    "a_gold_changes": ("语义复核与金标准补充记录", "*.json"),
    "b_gold_changes": ("语义复核与金标准补充记录", "*.json"),
    "b_standard_review": ("待处理语义与金标准补充材料", "*.json"),
    "update_plan": ("更新清单", "*.json"),
    "structure_review": ("行列调整的语义依据与复核", "*.json"),
    "unused_source_review": ("本次未采用来源的语义依据", "*.json"),
    "adjusted_a": ("调整行列后的附注表格（待填数）", "*.xlsx *.xlsm"),
    "adjusted_a_mapping": ("调整行列后的独立语义映射", "*.json"),
    "a_prime": ("更新后附注表格", "*.xlsx *.xlsm"),
    "a_prime_mapping": ("更新后附注表格语义映射", "*.json"),
    "word": ("更新后附注 Word", "*.docx"),
    "word_link": ("更新后附注表格与Word位置关联", "*.json"),
    "trace": ("逐格对应与更新核对", "*.xlsx"),
    "gold_path": ("语义金标准", "*.jsonl"),
}
STEPS = {
    "extract_a": {
        "number": "01", "title": "提取附注表格", "short": "提取附注表格",
        "description": "只提取“财务报表重要项目的说明”章节中的表格，并保存与 Word 的位置关联。",
        "required": ("word_path",), "inputs": ("word_path",), "optional": (),
        "outputs": ("a_path", "a_link"), "button": "提取附注表格与位置关联",
        "hint": "已有附注表格时，可直接进入第 2 步；已有位置关联也可在下方选择复用。",
    },
    "recognize_a": {
        "number": "02", "title": "识别附注表格语义", "short": "识别附注表格语义",
        "description": "逐格识别指标槽位与实际维度，独立保存映射。新客户、新日期是维度取值，不因此新增槽位；真正缺少的指标经核实后补入金标准。",
        "required": ("a_path",), "inputs": ("a_path",), "optional": ("a_link",),
        "outputs": ("a_mapping", "a_review", "a_gold_changes", "a_standard_review"), "button": "识别附注表格并保存映射",
        "hint": "已有完整语义映射可直接用于建立对应；部分成果可选择后继续识别，已确认格会先核验再复用。",
    },
    "recognize_b": {
        "number": "03", "title": "识别更新数据语义", "short": "识别更新数据语义",
        "description": "按金标准独立识别完整更新数据，保存每格的指标槽位与实际维度。新客户、新日期是维度取值，不因此新增槽位；真正缺少的指标经核实后补入金标准。",
        "required": ("b_path",), "inputs": ("b_path",), "optional": (),
        "outputs": ("b_mapping", "b_review", "b_scope", "b_gold_changes", "b_standard_review"), "button": "识别更新数据并保存映射",
        "hint": "每份更新数据表各自保存语义映射。已有这份数据的映射时，可直接选择复用。",
    },
    "match": {
        "number": "04", "title": "建立单元格对应", "short": "建立单元格对应",
        "description": "读取两个已保存的语义映射，勾稽相同含义的单元格，保存更新清单。",
        "required": ("a_mapping", "b_mapping"), "inputs": ("a_mapping", "b_mapping"),
        "optional": ("a_path", "b_path", "a_link"),
        "outputs": ("update_plan", "trace", "structure_review", "unused_source_review", "adjusted_a", "adjusted_a_mapping"), "button": "建立对应并保存清单",
        "hint": "需要增减行列时，先保存调整后的表格并重新识别语义，核对原格和新增格的含义后再形成清单。此步不填写新增金额。",
    },
    "write": {
        "number": "05", "title": "生成更新后的附注", "short": "生成更新后的附注",
        "description": "按清单更新“财务报表重要项目的说明”章节；其余章节保持原样，同时保存更新后附注表格和语义映射。",
        "required": ("update_plan",), "inputs": ("update_plan",), "optional": ("a_link",),
        "outputs": ("a_prime", "a_prime_mapping", "word", "word_link", "trace"), "button": "生成更新后的附注",
        "hint": "此步不重新识别语义。未提供 Word 位置关联时，可独立生成更新后附注表格。",
    },
}
MODEL_STEPS = {"recognize_a", "recognize_b"}


FINAL_FIELDS = ("a_prime", "a_prime_mapping", "word", "word_link", "trace")
STRUCTURE_FIELDS = ("structure_review", "adjusted_a", "adjusted_a_mapping", "unused_source_review")
INVALIDATED = {
    "word_path": ("report_scope", "a_context_text", "a_path", "a_link", "a_mapping", "a_review", "a_gold_changes", "a_standard_review", "b_scope", "update_plan") + FINAL_FIELDS,
    "a_path": ("a_context_text", "a_link", "a_mapping", "a_review", "a_gold_changes", "a_standard_review", "b_scope", "update_plan") + FINAL_FIELDS,
    "b_path": ("report_scope", "b_context_text", "b_mapping", "b_review", "b_scope", "b_gold_changes", "b_standard_review", "update_plan") + FINAL_FIELDS,
    "a_mapping": ("a_review", "a_gold_changes", "a_standard_review", "update_plan") + FINAL_FIELDS,
    "b_mapping": ("b_review", "b_scope", "b_gold_changes", "b_standard_review", "update_plan") + FINAL_FIELDS,
    "b_scope": ("b_mapping", "b_review", "b_gold_changes", "b_standard_review", "update_plan") + FINAL_FIELDS,
    "gold_path": ("update_plan",) + FINAL_FIELDS,
    "a_link": ("b_scope",) + FINAL_FIELDS,
    "update_plan": FINAL_FIELDS,
}
STEP_INVALIDATED = {
    "extract_a": ("a_path", "a_link", "a_mapping", "a_review", "a_gold_changes", "a_standard_review", "b_scope", "update_plan") + FINAL_FIELDS,
    "recognize_a": ("a_mapping", "a_review", "a_gold_changes", "a_standard_review", "update_plan") + FINAL_FIELDS,
    "recognize_b": ("b_mapping", "b_review", "b_scope", "b_gold_changes", "b_standard_review", "update_plan") + FINAL_FIELDS,
    "match": ("update_plan",) + FINAL_FIELDS,
    "write": FINAL_FIELDS,
}
for _key in INVALIDATED:INVALIDATED[_key] += STRUCTURE_FIELDS
for _key in STEP_INVALIDATED:
    if _key != 'write':STEP_INVALIDATED[_key] += STRUCTURE_FIELDS


def _read_record(path, schema, carrier=None):
    try:
        record = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        raise ValueError("无法读取所选成果文件：" + str(exc)) from exc
    accepted=(schema,) if isinstance(schema,str) else schema
    unified=isinstance(record,dict) and schema=='附注语义映射-v1' and record.get('schema_version')==2 and all(
        key in record for key in ('facts','owners','source_path','source_hash','gold_path','gold_hash'))
    if not isinstance(record, dict) or not (record.get("schema") in accepted or unified):
        raise ValueError("请选择本程序保存的相应成果文件")
    if carrier and record.get("carrier") != carrier:
        raise ValueError("选择的文件不是" + {"A": "附注表格", "B": "更新数据"}.get(carrier, "所选资料") + "的语义映射")
    return record


def _mapping_sources(path, carrier):
    record = _read_record(path, "附注语义映射-v1", carrier)
    updates = {"a_path" if carrier == "A" else "b_path": record.get("original_path") or record.get("source_path") or ""}
    updates["a_review" if carrier == "A" else "b_review"] = record.get("review_path") or ""
    updates["a_standard_review" if carrier == "A" else "b_standard_review"] = record.get("standard_review_path") or ""
    updates["a_context_text" if carrier == "A" else "b_context_text"] = record.get("additional_context") or ""
    updates["a_gold_changes" if carrier == "A" else "b_gold_changes"] = record.get("gold_learning_path") or ""
    updates["b_scope"] = (record.get("scope_selection_path") or "") if carrier == "B" else ""
    if carrier == "A":
        updates["a_link"] = record.get("a_link") or ""
    return updates


def file_selection_updates(key, path, current):
    """更换材料时同时更新其绑定来源，清掉不再适用的旧成果。"""
    path = str(path or "")
    if path == current.get(key, ""):
        return {}
    updates = {field: "" for field in INVALIDATED.get(key, ())}
    if path and key in {"a_mapping", "b_mapping"}:
        updates.update(_mapping_sources(path, "A" if key == "a_mapping" else "B"))
    if path and key == "update_plan":
        record = _read_record(path, ("附注更新清单-v1","附注更新清单-v2"))
        for mapping, carrier in (("a_mapping", "A"), ("b_mapping", "B")):
            bound = record.get(mapping)
            if not bound:
                raise ValueError("更新清单缺少绑定的" + FILE_FIELDS[mapping][0])
            updates[mapping] = bound
            updates.update(_mapping_sources(bound, carrier))
        updates["a_link"] = record.get("a_link") or ""
        structure=record.get('structure',{})
        updates['unused_source_review']=record.get('unused_source_review_path','')
        updates.update(structure_review=structure.get('report_path',''),adjusted_a_mapping=structure.get('mapping_path',''),
                       adjusted_a=structure.get('layout',{}).get('source_path',''))
    updates[key] = path
    return updates


def format_revision_preview(preview):
    if not preview:
        return "选择修订后的完整标准，先查看与当前标准的差异。尚未保存或切换版本。"
    lines = [f"新增 {preview.get('added', 0)} 项；修订 {preview.get('changed', 0)} 项；停用 {preview.get('retired', 0)} 项。",
             "结构检查不代表语义正确。以下含义需要由处理人员确认。",
             "指标定义与维度要求分开核实；新客户、新日期等维度取值不因此新增槽位。", ""]
    def meaning(item):
        if not item:
            return "无"
        if item.get('record_type'):
            kind=item['record_type']
            if kind=='metric':
                return ('指标定义：'+json.dumps(item['definition'],ensure_ascii=False,indent=2)+
                        '\n维度要求：'+json.dumps(item['dimension_refs'],ensure_ascii=False,indent=2)+
                        '\n状态：'+item['status']+'；来源：'+json.dumps(item['source_refs'],ensure_ascii=False))
            if kind=='dimension':return '维度定义：'+json.dumps(item,ensure_ascii=False,indent=2)
            return {'source_example':'来源示例：','legacy_conversion':'旧定义转换：'}.get(kind,'定义：')+json.dumps(item,ensure_ascii=False,indent=2)
        parts = []
        for key in ("note", "table"):
            value = item.get(key)
            if isinstance(value, dict):
                parts.append(" / ".join(str(value.get(k)) for k in ("name", "meaning") if value.get(k)))
            elif value:
                parts.append(str(value))
        text = "所属附注 / 标准表：" + " → ".join(part for part in parts if part) + "；状态：" + str(item.get("status") or "未注明")
        text += "\n指标定义：" + json.dumps(item.get("slot") or {}, ensure_ascii=False, indent=2)
        text += "\n维度要求：" + json.dumps(item.get("dimensions") or [], ensure_ascii=False, indent=2)
        details = {key: value for key, value in item.items() if key not in {"note", "table", "slot", "dimensions", "status", "id"}}
        if details:
            text += "\n其他定义与来源：" + json.dumps(details, ensure_ascii=False, indent=2)
        return text
    for change in preview.get("changes", []):
        lines.extend([{"added":"新增", "changed":"修订", "retired":"停用"}.get(change.get("kind"), "变更") + "：" + str(change.get("id")),
                      "原定义：" + meaning(change.get("before")), "新定义：" + meaning(change.get("after")), ""])
    return "\n".join(lines)


def format_external_review_preview(preview):
    reviewer = preview["reviewer"]
    method = {"human": "人工", "advanced_ai": "高级 AI"}.get(reviewer["method"], reviewer["method"])
    lines = ["审核者：" + reviewer["name"], "审核方式：" + method, "审核时间：" + reviewer["reviewed_at"],
        f"本次拟采纳 {preview['count']} 格；采纳后仍待核实 {preview['remaining_pending']} 格。",
        "确认后另存新映射和核对材料，不修改原件或金标准。取消不会采纳意见。",
        "原语义映射：" + preview["mapping_path"], "本次核验使用的金标准：" + preview["gold_path"],
        "复核意见：" + preview["opinion_path"], ""]
    for change in preview["changes"]:
        mapping = change["mapping"]
        lines.append(change["sheet"] + "!" + change["cell"])
        if preview.get('schema')=='附注外部复核预览-v2':
            previous=change.get('previous_mapping')
            if previous:
                lines.extend(['更正已确认语义','原指标槽位：'+previous['metric_id'],
                              '原实际维度：'+json.dumps(previous['dimensions'],ensure_ascii=False)])
            lines.extend(['指标槽位：'+mapping['metric_id'],'实际维度：'+json.dumps(mapping['dimensions'],ensure_ascii=False,indent=2),
                          '原值（不会修改）：'+str(mapping['raw_value']),
                          '逐项语义依据：'+json.dumps(mapping['dimension_evidence'],ensure_ascii=False),
                          '复核理由：'+change['reason'],''])
            continue
        if change.get("business_context"):
            lines.extend([method + "补充原格业务上下文", "原格内容：" + str(change.get("source_value")),
                "上下文依据：" + change["business_context"]["reason"], "回读原文："])
            lines.extend(item["cell"] + "：" + str(item["text"]) for item in change["context_evidence"])
        if change.get("previous_mapping"):
            previous = change["previous_mapping"]
            lines.extend(["更正已确认语义", "原指标槽位：" + previous["slot_id"] + "；" + str(previous.get("semantic_field") or ""),
                "原实际维度：" + json.dumps(previous["dimensions"], ensure_ascii=False, indent=2)])
        lines.extend(["原识别依据：" + change["previous_reason"],
            "采用的原表范围：" + change["sheet"] + "!" + str(mapping.get("table_range") or ""),
            "原表业务说明：" + str(mapping.get("table_semantic") or ""),
            "指标槽位：" + mapping["slot_id"] + "；" + str(mapping.get("semantic_field") or ""),
            "实际口径：" + {"standalone": "单户", "consolidated": "合并", "parent": "母公司"}.get(mapping["scope"], mapping["scope"]),
            "实际维度：" + json.dumps(mapping["dimensions"], ensure_ascii=False, indent=2),
            "原格证据：" + "、".join(mapping.get("evidence_cells") or []), "复核依据：" + mapping["reason"], ""])
    return "\n".join(lines)


def format_scope_review_preview(preview):
    reviewer = preview["reviewer"]
    method = {"human": "人工", "advanced_ai": "高级 AI"}.get(reviewer["method"], reviewer["method"])
    coverage = preview["coverage"]
    lines = ["审核者：" + reviewer["name"], "审核方式：" + method, "审核时间：" + reviewer["reviewed_at"],
        f"全部候选 {coverage['candidate_count']} 格；保留识别 {coverage['selected_count']} 格；明确章外 {coverage['out_of_scope_count']} 格。",
        f"保留识别中仍有 {preview['uncertain_count']} 格范围不确定，不按章外排除。",
        "本次只确认来源范围，未完成业务语义识别。",
        "确认后另存范围并开启来源筛选。旧更新数据语义映射、核对材料及后续更新成果的选择会清空，原文件保留；请按新范围重新识别更新数据。",
        "不修改原件或金标准。取消不会采纳意见。",
        "更新数据表：" + preview["source_path"], "附注表格与 Word 位置关联：" + preview["a_link_path"],
        "本次核验使用的金标准：" + preview["gold_path"], "范围复核意见：" + preview["opinion_path"], ""]
    for change in preview["changes"]:
        decision = {"include": "保留识别", "out_of_scope": "明确章外", "uncertain": "不确定，保留识别"}[change["decision"]]
        lines.extend([change["sheet"] + "!" + change["range"] + "：" + decision,
            "判断依据：" + change["reason"], "来源原格证据：" + "、".join(change["evidence_cells"])])
        comparison = change.get("semantic_comparison")
        if comparison:
            lines.extend(["来源业务：" + comparison["source_business"], "附注业务：" + comparison["target_business"],
                "业务差异：" + comparison["reason"], "附注原文证据：" + "、".join(comparison["target_evidence"])])
        lines.append("")
    return "\n".join(lines)


def missing_inputs(step, config):
    """只检查当前操作的必要材料，不让模型设置阻挡提取和回写。"""
    if step not in STEPS:
        return ["请选择要执行的步骤"]
    missing = []
    for key in STEPS[step]["required"]:
        value = str(config.get(key) or "")
        path = Path(value)
        if not value or not path.is_file():
            missing.append("请选择有效的" + FILE_FIELDS[key][0] + "文件")
    if not config.get("output_dir"):
        missing.append("请选择成果保存文件夹")
    elif Path(config["output_dir"]).exists() and not Path(config["output_dir"]).is_dir():
        missing.append("成果保存位置应为文件夹")
    if step in MODEL_STEPS:
        if not Path(config.get("gold_path") or "").is_file():
            missing.append("请在设置中选择语义金标准")
        if not config.get("base_url") or not config.get("model"):
            missing.append("请在设置中填写模型地址和模型名称")
    return missing


class Application:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("附注自动更新系统")
        root.configure(bg=BG)
        root.geometry(f"{min(1320, root.winfo_screenwidth()-70)}x{min(960, root.winfo_screenheight()-100)}")
        root.minsize(1000, 680)
        root.option_add("*Font", ("Microsoft YaHei UI", 10))
        root.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("TEntry", padding=4, fieldbackground="white")
        style.configure("TProgressbar", troughcolor="#E7EDF5", background=BLUE, borderwidth=0, thickness=5)
        self.messages = queue.Queue()
        self.cancel = threading.Event()
        self.worker = None
        self.running = False
        self.running_step = None
        self.closing = False
        self.secret = ""
        self.last_output = None
        self.global_editable = []
        self.page_editable = []
        self.path_open_buttons = {}
        self.path_entries = {}
        self.mapping_summary_labels = {}
        self.context_panels = {}
        self.context_toggles = {}
        self.context_text_widgets = {}
        self.outputs_card = None
        self.step_states = {}
        self.step_messages = {}
        error = None
        try:
            saved = load_settings()
        except ValueError as exc:
            error = str(exc)
            saved = {}
        defaults = {"gold_path": DEFAULT_GOLD_PATH, "output_dir": str(PROJECT_DIR / "输出"), "timeout": DEFAULT_TIMEOUT, "thinking_mode": "default"}
        fields = tuple(FILE_FIELDS) + ("output_dir", "base_url", "model", "api_key", "timeout", "thinking_mode", "a_context_text", "b_context_text", "report_scope")
        self.values = {key: tk.StringVar(value=saved.get(key, defaults.get(key, ""))) for key in fields}
        self.scope_prefilter = tk.BooleanVar(value=saved.get("scope_prefilter", False))
        self.current_step = saved.get("current_step", "extract_a")
        if self.current_step not in STEPS:
            self.current_step = "extract_a"
        self.revision_preview = None
        self.revision_current = ""
        self.revision_candidate = tk.StringVar(value="")
        self.show_key = tk.BooleanVar(value=False)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)
        self.build_header()
        self.build_output_location()
        self.build_workspace()
        self.build_log()
        self.show_step(self.current_step)
        for key, value in self.values.items():
            value.trace_add("write", lambda *_, field=key: self.on_value_change(field))
        self.root.after(100, self.poll)
        self.append_log(error or "按需要选择一个步骤。已有 Excel、语义映射、位置关联或更新清单均可直接复用。")
        if error:
            self.append_log("旧配置未改动，请到设置页重新填写。")

    def label(self, parent, text, **options):
        return tk.Label(parent, **{"text": text, "bg": "white", "fg": TEXT, "anchor": "w", **options})

    def button(self, parent, text, command, primary=False, choose=False):
        color = BLUE if primary else GREEN if choose else "#EAF0F6"
        return tk.Button(parent, text=text, command=command, bg=color,
                         fg="white" if primary or choose else TEXT, relief="flat", borderwidth=0,
                         activebackground="#DCE8F4", activeforeground=TEXT, padx=10, pady=3,
                         cursor="hand2", disabledforeground="#929FAC")

    def card(self, parent, **options):
        return tk.Frame(parent, **{"bg": "white", "highlightthickness": 1,
                        "highlightbackground": BORDER, "padx": 12, "pady": 7, **options})

    def build_header(self):
        card = self.card(self.root)
        card.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 5))
        card.columnconfigure(0, weight=1)
        self.label(card, "附注自动更新", font=("Microsoft YaHei UI", 15, "bold")).grid(row=0, column=0, sticky="w")
        settings = self.button(card, "模型与项目设置", lambda: self.show_step("settings"))
        settings.grid(row=0, column=1, padx=(15, 0))
        self.global_editable.append((settings, "normal"))

    def build_output_location(self):
        card = self.card(self.root, pady=4)
        card.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 5))
        card.columnconfigure(1, weight=1)
        self.label(card, "成果保存位置").grid(row=0, column=0, padx=(0, 12))
        entry = ttk.Entry(card, textvariable=self.values["output_dir"], state="readonly")
        entry.grid(row=0, column=1, sticky="ew")
        choose = self.button(card, "选择文件夹", self.choose_output, choose=True)
        choose.grid(row=0, column=2, padx=(8, 0))
        self.button(card, "打开", lambda: self.open_path("output_dir")).grid(row=0, column=3, padx=(6, 0))
        self.global_editable.extend(((entry, "readonly"), (choose, "normal")))

    def build_workspace(self):
        workspace = tk.Frame(self.root, bg=BG)
        workspace.grid(row=2, column=0, sticky="nsew", padx=12)
        workspace.rowconfigure(0, weight=1)
        workspace.columnconfigure(1, weight=1)
        sidebar = self.card(workspace)
        sidebar.grid(row=0, column=0, sticky="ns", padx=(0, 8))
        self.label(sidebar, "工作流程", font=("Microsoft YaHei UI", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))
        self.nav_buttons = {}
        for row, (step, definition) in enumerate(STEPS.items(), 1):
            button = tk.Button(sidebar, text=definition["number"] + "  " + definition["short"],
                               command=lambda key=step: self.show_step(key), anchor="w",
                               bg="white", fg=TEXT, relief="flat", bd=0, padx=8, pady=9,
                               width=18, cursor="hand2", font=("Microsoft YaHei UI", 10))
            button.grid(row=row, column=0, sticky="ew", pady=2)
            self.nav_buttons[step] = button
        self.label(sidebar, "已有成果可直接选择复用，\n不必从第一步重来。",
                   justify="left", fg=MUTED, font=("Microsoft YaHei UI", 9)).grid(row=6, column=0, sticky="w", pady=(12, 0))

        pane = tk.Frame(workspace, bg=BG)
        pane.grid(row=0, column=1, sticky="nsew")
        pane.columnconfigure(0, weight=1)
        pane.rowconfigure(1, weight=1)
        heading = self.card(pane)
        heading.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        heading.columnconfigure(0, weight=1)
        self.page_title = self.label(heading, "", font=("Microsoft YaHei UI", 12, "bold"))
        self.page_title.grid(row=0, column=0, sticky="w")
        self.page_description = self.label(heading, "", fg=MUTED, justify="left", wraplength=790, font=("Microsoft YaHei UI", 9))
        self.page_description.grid(row=1, column=0, sticky="ew", pady=(3, 0))
        self.page_description.bind("<Configure>", lambda event: self.page_description.configure(wraplength=max(250,event.width)))

        content = tk.Frame(pane, bg=BG)
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(content, bg=BG, highlightthickness=0, bd=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(content, command=self.canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.page = tk.Frame(self.canvas, bg=BG)
        self.page.columnconfigure(0, weight=1)
        self.page_window = self.canvas.create_window((0, 0), window=self.page, anchor="nw")
        self.page.bind("<Configure>", lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda event: self.canvas.itemconfigure(self.page_window, width=event.width))
        self.root.bind("<MouseWheel>", self.on_content_wheel, add="+")

        execution = self.card(pane, pady=5)
        execution.grid(row=2, column=0, sticky="ew", pady=(5, 0))
        execution.columnconfigure(2, weight=1)
        self.start_button = self.button(execution, "", self.start, primary=True)
        self.start_button.grid(row=0, column=0, rowspan=2, padx=(0, 8), sticky="ns")
        self.cancel_button = self.button(execution, "取消本步", self.request_cancel)
        self.cancel_button.grid(row=0, column=1, rowspan=2, padx=(0, 12), sticky="ns")
        self.cancel_button.configure(state="disabled")
        self.status = self.label(execution, "", fg=MUTED, wraplength=500, justify="left", font=("Microsoft YaHei UI", 9))
        self.status.grid(row=0, column=2, sticky="ew")
        self.status.bind("<Configure>", lambda event: self.status.configure(wraplength=max(100,event.width)))
        self.progress = ttk.Progressbar(execution, mode="indeterminate")
        self.progress.grid(row=1, column=2, sticky="ew", pady=(6, 0))

    def build_log(self):
        card = self.card(self.root, pady=5)
        card.grid(row=3, column=0, sticky="ew", padx=12, pady=(6, 8))
        card.columnconfigure(0, weight=1)
        heading = tk.Frame(card, bg="white")
        heading.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        self.label(heading, "本次执行记录", font=("Microsoft YaHei UI", 10, "bold")).pack(side="left")
        self.button(heading, "打开最近成果文件夹", self.open_last_output).pack(side="right")
        self.log = tk.Text(card, height=2, wrap="word", state="disabled", bg="#F8FAFD", fg=TEXT,
                           relief="flat", padx=8, pady=3, font=("Microsoft YaHei UI", 9))
        self.log.grid(row=1, column=0, sticky="ew")
        scroll = ttk.Scrollbar(card, command=self.log.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)

    def file_row(self, parent, row, key, optional=False, reusable=False, output=False):
        name, pattern = FILE_FIELDS[key]
        field = tk.Frame(parent, bg="white")
        field.grid(row=row, column=0, columnspan=5, sticky="ew", pady=3)
        field.columnconfigure(0 if output else 1, weight=1)
        title = self.label(field, name + ("（可选）" if optional else ""), font=("Microsoft YaHei UI", 9))
        title.grid(row=0, column=0, columnspan=4 if output else 1, sticky="w", padx=(0, 8), pady=(0, 2) if output else 0)
        if not output:
            title.configure(wraplength=150)
        line, column = (1, 0) if output else (0, 1)
        entry = ttk.Entry(field, textvariable=self.values[key], state="readonly", width=20)
        entry.grid(row=line, column=column, sticky="ew")
        choose = self.button(field, "选择已有" if reusable else "选择文件",
                             lambda: self.choose_file(key), choose=True)
        choose.grid(row=line, column=column+1, padx=(6, 0))
        opening = self.button(field, "打开", lambda: self.open_path(key))
        opening.grid(row=line, column=column+2, padx=(5, 0))
        self.path_open_buttons[key] = opening
        self.path_entries[key] = entry
        if key in {"a_mapping", "b_mapping", "a_prime_mapping"}:
            summary = self.label(field, "", fg=MUTED, wraplength=760, justify="left", font=("Microsoft YaHei UI", 9))
            summary.grid(row=line+1, column=0, columnspan=4, sticky="w", pady=(3, 0))
            self.mapping_summary_labels[key] = summary
            self.refresh_mapping_summary(key)
        self.page_editable.extend(((entry, "readonly"), (choose, "normal")))
        if optional or key == "b_mapping":
            clear = self.button(field, "清空", lambda: self.select_file(key, ""))
            clear.grid(row=line, column=column+3, padx=(5, 0))
            self.page_editable.append((clear, "normal"))

    def on_value_change(self, key):
        self.refresh_mapping_summary(key)
        self.refresh_ready()

    def refresh_mapping_summary(self, key):
        label = self.mapping_summary_labels.get(key)
        if label is None or not label.winfo_exists():
            return
        path = self.values[key].get()
        if not path:
            label.configure(text="尚未选择语义映射", fg=MUTED)
            return
        try:
            record = _read_record(path, "附注语义映射-v1", "B" if key == "b_mapping" else "A")
            business_key = 'facts' if record.get('schema_version') == 2 else 'mappings'
            if any(not isinstance(record.get(category), list) for category in (business_key, "excluded", "unresolved")):
                raise ValueError("映射缺少有效的业务、非业务或待核实清单")
            if type(record.get("complete", False)) is not bool or record.get("complete") and record["unresolved"]:
                raise ValueError("映射的完整标记与待核实清单不一致")
            state = "完整成果" if record.get("complete") is True else "部分成果"
            business_count = sum(row.get('record_type') == 'metric_fact' for row in record['facts']) if business_key == 'facts' else len(record['mappings'])
            text = state + "：" + "；".join(f"{name} {count} 格" for name, count in
                (("已确认业务", business_count), ("非业务", len(record['excluded'])), ("待核实", len(record['unresolved']))))
            label.configure(text=text + "。执行时仍会核验来源与金标准。", fg=MUTED)
        except (OSError, ValueError, TypeError) as error:
            label.configure(text="不能读取映射成果：" + str(error), fg="#A65B05")

    def open_pending_instructions(self, carrier):
        key = "a_standard_review" if carrier == "A" else "b_standard_review"
        try:
            package = _read_record(self.values[key].get(), ("附注待处理语义-v1", "统一金标准补充材料-v2"))
            if package.get('schema') == '统一金标准补充材料-v2':
                instructions = package.get('review_instructions')
                if not isinstance(instructions, list) or not instructions or any(not isinstance(item, str) or not item.strip() for item in instructions):
                    raise ValueError('待处理材料缺少处理说明')
                messagebox.showinfo('待处理语义与金标准补充说明', '\n\n'.join(instructions), parent=self.root)
                return
            if package.get('carrier') != carrier:
                raise ValueError('待处理材料与本表不一致')
            path = package.get("instructions_path")
            if not isinstance(path, str) or not path.strip() or not Path(path).is_file():
                raise ValueError("待处理材料未提供可打开的说明文件")
            os.startfile(path)
        except (OSError, ValueError, TypeError) as error:
            messagebox.showerror("待处理说明未打开", "不能读取待处理说明：" + str(error), parent=self.root)

    def export_external_review(self, carrier):
        if self.running:
            messagebox.showinfo("当前步骤正在执行", "请等待本步结束后导出复核意见格式。", parent=self.root)
            return
        key = "a_mapping" if carrier == "A" else "b_mapping"
        mapping = self.values[key].get()
        if not mapping:
            messagebox.showinfo("尚未选择映射", "请先选择本表的语义映射。", parent=self.root)
            return
        path = filedialog.asksaveasfilename(parent=self.root, title="导出复核意见格式", defaultextension=".json",
            initialfile=("附注表格" if carrier == "A" else "更新数据") + "复核意见.json", filetypes=(("复核意见", "*.json"),))
        if not path:
            return
        try:
            from .外部复核 import export_review_template
            saved = export_review_template(mapping, path, gold_path=self.values["gold_path"].get(), include_confirmed=True)
        except (ValueError, OSError) as error:
            messagebox.showerror("复核意见格式未导出", str(error), parent=self.root)
            return
        self.append_log("已导出复核意见格式：" + str(saved))

    def import_external_review(self, carrier):
        if self.running:
            messagebox.showinfo("当前步骤正在执行", "请等待本步结束后导入复核意见，运行中不能替换映射。", parent=self.root)
            return
        key = "a_mapping" if carrier == "A" else "b_mapping"
        mapping = self.values[key].get()
        if not mapping:
            messagebox.showinfo("尚未选择映射", "请先选择本表的语义映射。", parent=self.root)
            return
        path = filedialog.askopenfilename(parent=self.root, title="导入人工或高级 AI 复核意见", filetypes=(("复核意见", "*.json"),))
        if not path:
            return
        try:
            from .外部复核 import preview_external_review
            preview = preview_external_review(mapping, path, gold_path=self.values["gold_path"].get())
            text = format_external_review_preview(preview)
        except (ValueError, OSError, KeyError, TypeError) as error:
            messagebox.showerror("复核意见不能预览", str(error), parent=self.root)
            return
        return self.review_preview_window("复核意见预览 · 确认前不会采纳", text, "确认另存并选用",
            lambda window: self.confirm_external_review(preview, window))

    def review_preview_window(self, title, text, confirm_text, confirm):
        window = tk.Toplevel(self.root)
        window.title(title)
        window.geometry("880x620");window.transient(self.root)
        window.columnconfigure(0, weight=1);window.rowconfigure(0, weight=1)
        body = tk.Text(window, wrap="word", padx=12, pady=10, font=("Microsoft YaHei UI", 10))
        body.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(window, command=body.yview);scroll.grid(row=0, column=1, sticky="ns")
        body.configure(yscrollcommand=scroll.set);body.insert("1.0", text);body.configure(state="disabled")
        actions = tk.Frame(window);actions.grid(row=1, column=0, columnspan=2, sticky="e", padx=12, pady=10)
        self.button(actions, "取消", window.destroy).pack(side="right", padx=(8, 0))
        self.button(actions, confirm_text, lambda: confirm(window), primary=True).pack(side="right")
        window.protocol("WM_DELETE_WINDOW", window.destroy);window.grab_set()
        return window

    def confirm_external_review(self, preview, window):
        if self.running:
            messagebox.showinfo("当前步骤正在执行", "运行中不能采纳复核意见或替换映射。", parent=window)
            return False
        key = "a_mapping" if preview["carrier"] == "A" else "b_mapping"
        step = "recognize_a" if preview["carrier"] == "A" else "recognize_b"
        try:
            if Path(self.values[key].get()).resolve() != Path(preview["mapping_path"]).resolve():
                raise ValueError("当前选择的语义映射已变化，请重新导入复核意见并预览。")
            if Path(self.values["gold_path"].get()).resolve() != Path(preview["gold_path"]).resolve():
                raise ValueError("当前选择的金标准已变化，请重新导入复核意见并预览。")
            output = self.values["output_dir"].get()
            if not output or not Path(output).is_dir():
                raise ValueError("请先选择有效的成果保存文件夹。")
            from .外部复核 import apply_external_review
            result = apply_external_review(preview, output, confirmed=True, confirm_method="界面确认导入外部复核意见")
            files = result["files"]
            updates = file_selection_updates(key, files[key], self.config())
            updates.update({field: path for field, path in files.items() if field in STEPS[step]["outputs"]})
        except (ValueError, OSError, KeyError, TypeError) as error:
            messagebox.showerror("复核意见未采纳", str(error), parent=window)
            return False
        self.apply_values(updates)
        self.last_output = str(Path(files[key]).parent)
        self.step_states[step] = result["status"]
        self.step_messages[step] = result["message"]
        try:
            save_settings(self.config())
        except (ValueError, OSError) as error:
            messagebox.showerror("复核结果已保存，设置尚未保存", str(error) + "\n新映射：" + files[key], parent=window)
        self.append_log(result["message"] + "\n新映射：" + files[key])
        window.destroy();self.show_step(self.current_step)
        return True

    def scope_review_inputs(self):
        inputs = {key: self.values[key].get() for key in ("b_path", "a_link", "gold_path")}
        for key, path in inputs.items():
            if not path or not Path(path).is_file():
                raise ValueError("请先选择有效的" + FILE_FIELDS[key][0] + "。")
        return inputs

    def export_scope_review(self):
        if self.running:
            messagebox.showinfo("当前步骤正在执行", "请等待本步结束后导出范围复核格式。", parent=self.root)
            return
        try:
            inputs = self.scope_review_inputs()
            path = filedialog.asksaveasfilename(parent=self.root, title="导出范围复核格式", defaultextension=".json",
                initialfile="更新数据范围复核意见.json", filetypes=(("范围复核意见", "*.json"),))
            if not path:
                return
            from .外部范围 import export_scope_review_template
            saved = export_scope_review_template(inputs["b_path"], inputs["a_link"], path, gold_path=inputs["gold_path"])
        except (ValueError, OSError) as error:
            messagebox.showerror("范围复核格式未导出", str(error), parent=self.root)
            return
        self.append_log("已导出范围复核格式：" + str(saved))

    def import_scope_review(self):
        if self.running:
            messagebox.showinfo("当前步骤正在执行", "运行中不能导入范围意见或替换来源范围。", parent=self.root)
            return
        try:
            inputs = self.scope_review_inputs()
            path = filedialog.askopenfilename(parent=self.root, title="导入人工或高级 AI 范围复核意见", filetypes=(("范围复核意见", "*.json"),))
            if not path:
                return
            from .外部范围 import preview_external_scope_review
            preview = preview_external_scope_review(inputs["b_path"], inputs["a_link"], path, gold_path=inputs["gold_path"])
            text = format_scope_review_preview(preview)
        except (ValueError, OSError, KeyError, TypeError) as error:
            messagebox.showerror("范围意见不能预览", str(error), parent=self.root)
            return
        return self.review_preview_window("范围复核预览 · 确认前不会采纳", text, "确认范围并选用",
            lambda window: self.confirm_scope_review(preview, window))

    def confirm_scope_review(self, preview, window):
        if self.running:
            messagebox.showinfo("当前步骤正在执行", "运行中不能采纳范围意见或替换来源范围。", parent=window)
            return False
        try:
            for key, bound in (("b_path", "source_path"), ("a_link", "a_link_path"), ("gold_path", "gold_path")):
                if not self.values[key].get() or Path(self.values[key].get()).resolve() != Path(preview[bound]).resolve():
                    raise ValueError("当前选择的" + FILE_FIELDS[key][0] + "已变化，请重新导入范围意见并预览。")
            output = self.values["output_dir"].get()
            if not output or not Path(output).is_dir():
                raise ValueError("请先选择有效的成果保存文件夹。")
            from .外部范围 import apply_external_scope_review
            result = apply_external_scope_review(preview, output, confirmed=True, confirm_method="界面确认导入范围复核意见")
            path = result["files"]["b_scope"]
            updates = file_selection_updates("b_scope", path, self.config())
        except (ValueError, OSError, KeyError, TypeError) as error:
            messagebox.showerror("范围意见未采纳", str(error), parent=window)
            return False
        self.apply_values(updates)
        self.scope_prefilter.set(True)
        self.last_output = str(Path(path).parent)
        self.step_states.pop("recognize_b", None)
        self.step_messages["recognize_b"] = "范围已保存并选用，尚未完成业务语义识别；请按新范围重新识别更新数据。"
        try:
            save_settings(self.config())
        except (ValueError, OSError) as error:
            messagebox.showerror("范围已保存，设置尚未保存", str(error) + "\n新范围：" + path, parent=window)
        self.append_log(self.step_messages["recognize_b"] + "\n新范围：" + path)
        window.destroy();self.show_step(self.current_step)
        return True

    def on_content_wheel(self, event):
        widget = event.widget
        if isinstance(widget, (tk.Text, tk.Listbox, ttk.Treeview, ttk.Scrollbar)):
            return None
        current = widget
        while current is not None:
            if current == self.page or current == self.canvas:
                delta = int(getattr(event, "delta", 0))
                if delta:
                    self.canvas.yview_scroll(-1 if delta > 0 else 1, "units")
                    return "break"
                return None
            current = getattr(current, "master", None)
        return None

    def scroll_to_outputs(self):
        if self.outputs_card is None or not self.outputs_card.winfo_exists():
            return
        self.root.update_idletasks()
        height = max(1, self.page.winfo_height())
        self.canvas.yview_moveto(max(0, self.outputs_card.winfo_y() - 3) / height)
        for key in STEPS.get(self.current_step, {}).get("outputs", ()):
            entry = self.path_entries.get(key)
            if entry and entry.winfo_exists():
                entry.xview_moveto(1)

    def show_step(self, step):
        if self.running or (step not in STEPS and step != "settings"):
            return
        self.current_step = step
        for widget in self.page.winfo_children():
            widget.destroy()
        self.page_editable = []
        self.path_open_buttons = {}
        self.path_entries = {}
        self.mapping_summary_labels = {}
        self.context_panels = {}
        self.context_toggles = {}
        self.context_text_widgets = {}
        self.outputs_card = None
        self.canvas.yview_moveto(0)
        for key, button in self.nav_buttons.items():
            definition = STEPS[key]
            state = "  已完成" if self.step_states.get(key) == "complete" else ""
            button.configure(bg="#E5EEFF" if key == step else "white",
                             fg=BLUE if key == step else TEXT,
                             text=definition["number"] + "  " + definition["short"] + state)
        if step == "settings":
            self.build_settings_page()
            self.start_button.configure(text="保存设置")
        else:
            definition = STEPS[step]
            self.page_title.configure(text=definition["number"] + "  " + definition["title"])
            self.page_description.configure(text=definition["description"])
            inputs = self.card(self.page)
            inputs.grid(row=0, column=0, sticky="ew", pady=(0, 8))
            inputs.columnconfigure(1, weight=1)
            self.label(inputs, "本步使用的文件", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 4))
            row = 1
            for key in definition["inputs"]:
                self.file_row(inputs, row, key, reusable=key in {"a_mapping", "b_mapping", "update_plan"})
                row += 1
            if definition["optional"]:
                optional = ttk.LabelFrame(inputs, text="可选的来源与关联", padding=8)
                optional.grid(row=row, column=0, columnspan=5, sticky="ew", pady=(7, 0))
                optional.columnconfigure(1, weight=1)
                for index, key in enumerate(definition["optional"]):
                    self.file_row(optional, index, key, optional=True, reusable=True)
            unified_gold = False
            if step == "recognize_b":
                try:
                    with Path(self.values['gold_path'].get()).open(encoding='utf-8-sig') as handle:
                        first = next((json.loads(line) for line in handle if line.strip()), {})
                    unified_gold = isinstance(first, dict) and bool(first.get('record_type'))
                except (OSError, ValueError):
                    pass
            if step == "recognize_b" and not unified_gold:
                prefilter = tk.Checkbutton(inputs, text="先按附注内容筛选来源（可选）", bg="white",
                    variable=self.scope_prefilter, command=self.change_scope_prefilter)
                prefilter.grid(row=row, column=0, columnspan=5, sticky="w", pady=(7, 0))
                self.page_editable.append((prefilter, "normal"))
                self.label(inputs, "未勾选时，按金标准独立识别完整更新数据；附注中已有的客户、项目和年份不是排除其他数据的白名单。勾选且有 Word 位置关联时才先筛选，不能确定的内容仍保留识别。",
                    fg=MUTED, wraplength=760, justify="left", font=("Microsoft YaHei UI", 9)).grid(
                    row=row+1, column=0, columnspan=5, sticky="w", pady=(2, 3))
                scope_actions = tk.Frame(inputs, bg="white")
                scope_actions.grid(row=row+2, column=0, columnspan=5, sticky="w", pady=(3, 3))
                for text, action in (("导出范围复核格式", self.export_scope_review), ("导入范围复核意见", self.import_scope_review)):
                    button = self.button(scope_actions, text, action)
                    button.pack(side="left", padx=(0, 8));self.page_editable.append((button, "normal"))
                row += 3
            elif step == "recognize_b":
                self.scope_prefilter.set(False)
                definition = {**definition, 'outputs': tuple(key for key in definition['outputs'] if key != 'b_scope')}
                self.label(inputs, "按新版金标准独立识别完整更新数据，保存每格的槽位、实际维度和依据；与附注表格建立对应在下一步进行。",
                    fg=MUTED, wraplength=760, justify="left", font=("Microsoft YaHei UI", 9)).grid(
                    row=row, column=0, columnspan=5, sticky="w", pady=(7, 3))
                row += 1
            if step in MODEL_STEPS:
                self.context_field(inputs, row + 1, "a_context_text" if step == "recognize_a" else "b_context_text")
            outputs = self.card(self.page)
            self.outputs_card = outputs
            outputs.grid(row=1, column=0, sticky="ew")
            outputs.columnconfigure(1, weight=1)
            self.label(outputs, "本步成果 · 可选择已有文件复用", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 4))
            for row, key in enumerate(definition["outputs"], 1):
                self.file_row(outputs, row, key, reusable=True, output=True)
            self.label(outputs, definition["hint"], fg=MUTED, wraplength=760, justify="left",
                       font=("Microsoft YaHei UI", 9)).grid(row=len(definition["outputs"])+1, column=0, columnspan=5, sticky="ew", pady=(8, 0))
            if step in MODEL_STEPS:
                self.label(outputs, "指标槽位说明“是什么指标”，实际维度说明“哪个客户、哪个日期等”。新维度取值沿用适用的原槽位；真正缺少的指标才补入金标准，仍有分歧的内容保留待核实。",
                           fg=MUTED, wraplength=700, justify="left", font=("Microsoft YaHei UI", 9)).grid(
                           row=len(definition["outputs"])+2, column=0, columnspan=5, sticky="ew", pady=(8, 0))
                actions = tk.Frame(outputs, bg="white")
                actions.grid(row=len(definition["outputs"])+3, column=0, columnspan=5, sticky="w", pady=(8, 0))
                manage = self.button(actions, "管理金标准", lambda: self.show_step("settings"))
                manage.pack(side="left")
                self.button(actions, "打开待处理说明", lambda: self.open_pending_instructions("A" if step == "recognize_a" else "B")).pack(side="left", padx=(8, 0))
                self.page_editable.append((manage, "normal"))
                review_actions = tk.Frame(outputs, bg="white")
                review_actions.grid(row=len(definition["outputs"])+4, column=0, columnspan=5, sticky="w", pady=(8, 0))
                carrier = "A" if step == "recognize_a" else "B"
                for text, action in (("导出复核意见格式", self.export_external_review), ("导入复核意见", self.import_external_review)):
                    button = self.button(review_actions, text, lambda action=action, carrier=carrier: action(carrier))
                    button.pack(side="left", padx=(0, 8));self.page_editable.append((button, "normal"))
            if step == "write":
                sources = self.card(self.page)
                sources.grid(row=2, column=0, sticky="ew", pady=(8, 0))
                sources.columnconfigure(1, weight=1)
                self.label(sources, "当前清单绑定的来源", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
                for index, key in enumerate(("a_mapping", "b_mapping", "a_path", "b_path"), 1):
                    self.label(sources, FILE_FIELDS[key][0], font=("Microsoft YaHei UI", 9)).grid(row=index, column=0, padx=(0, 10), pady=3, sticky="w")
                    ttk.Entry(sources, textvariable=self.values[key], state="readonly").grid(row=index, column=1, sticky="ew", pady=3)
                reuse = self.button(outputs, "将更新后附注表格用于下一轮", self.use_updated_a)
                reuse.grid(row=len(definition["outputs"])+2, column=0, columnspan=5, sticky="w", pady=(10, 0))
                self.page_editable.append((reuse, "normal"))
            self.start_button.configure(text=definition["button"])
        self.refresh_ready()

    def build_settings_page(self):
        self.page_title.configure(text="模型与项目设置")
        self.page_description.configure(text="语义识别在第 2、3 步调用模型。提取和最终写入可独立执行。")
        model = self.card(self.page)
        model.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        model.columnconfigure(1, weight=1)
        for row, (title, key) in enumerate((("服务地址", "base_url"), ("模型名称", "model"), ("API Key", "api_key"), ("请求超时（秒）", "timeout"))):
            self.label(model, title).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=5)
            entry = ttk.Entry(model, textvariable=self.values[key], show="●" if key == "api_key" and not self.show_key.get() else "")
            entry.grid(row=row, column=1, sticky="ew", pady=5)
            self.page_editable.append((entry, "normal"))
            if key == "api_key":
                self.key_entry = entry
        self.label(model, "模型思考模式").grid(row=4, column=0, sticky="w", padx=(0, 12), pady=5)
        thinking_options = {"服务默认":"default", "开启":"enabled", "关闭":"disabled"}
        self.thinking_choice = tk.StringVar(value=next(name for name,value in thinking_options.items() if value == self.values["thinking_mode"].get()))
        thinking = ttk.Combobox(model, textvariable=self.thinking_choice, values=tuple(thinking_options), state="readonly")
        thinking.grid(row=4, column=1, sticky="ew", pady=5)
        thinking.bind("<<ComboboxSelected>>", lambda _: self.values["thinking_mode"].set(thinking_options[self.thinking_choice.get()]))
        self.page_editable.append((thinking, "readonly"))
        show = tk.Checkbutton(model, text="显示密钥", bg="white", variable=self.show_key,
                              command=lambda: self.key_entry.configure(show="" if self.show_key.get() else "●"))
        show.grid(row=5, column=1, sticky="w")
        self.page_editable.append((show, "normal"))
        self.label(model, "密钥只由当前 Windows 用户加密保存。", fg=MUTED, font=("Microsoft YaHei UI", 9)).grid(row=6, column=0, columnspan=2, sticky="w", pady=(5, 0))
        self.label(model, "API Key 按服务要求填写；无需认证的服务可以留空。", fg=MUTED,
                   font=("Microsoft YaHei UI", 9)).grid(row=7, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self.label(model, "报告口径（可指定）").grid(row=8, column=0, sticky="w", padx=(0, 12), pady=5)
        options={"按文件识别":"","单户":"standalone","合并":"consolidated","母公司":"parent"}
        self.scope_choice=tk.StringVar(value=next((name for name,value in options.items() if value==self.values["report_scope"].get()),"按文件识别"))
        scope=ttk.Combobox(model,textvariable=self.scope_choice,values=tuple(options),state="readonly")
        scope.grid(row=8,column=1,sticky="ew",pady=5)
        def change_scope(_):
            value=options[self.scope_choice.get()]
            if value!=self.values["report_scope"].get():
                fields=set(STEP_INVALIDATED["recognize_a"]+STEP_INVALIDATED["recognize_b"])
                self.apply_values({**{key:"" for key in fields},"report_scope":value})
        scope.bind("<<ComboboxSelected>>",change_scope)
        self.page_editable.append((scope,"readonly"))
        self.label(model,"指定口径会同时约束两份表的识别；更换原始资料后重新确认。",fg=MUTED,font=("Microsoft YaHei UI",9)).grid(row=9,column=0,columnspan=2,sticky="w",pady=(4,0))
        self.build_standard_settings()

    def context_field(self, parent, row, key):
        box = tk.Frame(parent, bg="white")
        box.grid(row=row, column=0, columnspan=5, sticky="ew", pady=(6, 0))
        box.columnconfigure(0, weight=1)
        expanded = bool(self.values[key].get().strip())
        toggle = self.button(box, ("收起" if expanded else "展开") + "：文件未注明的信息（可选）", lambda: self.toggle_context(key))
        toggle.grid(row=0, column=0, sticky="w")
        details = tk.Frame(box, bg="white")
        details.grid(row=1, column=0, sticky="ew", pady=(5, 0));details.columnconfigure(0, weight=1)
        self.label(details, "只补充主体、实际日期、币种、单位及单户/合并/母公司口径；文件已写明时留空。",
                   fg=MUTED, wraplength=700, justify="left", font=("Microsoft YaHei UI", 9)).grid(row=0, column=0, sticky="w", pady=(0, 4))
        context = tk.Text(details, height=3, wrap="word", relief="solid", borderwidth=1, bg="white", fg=TEXT)
        context.grid(row=1, column=0, sticky="ew")
        context.insert("1.0", self.values[key].get())
        context.edit_modified(False)
        def change(_):
            if context.edit_modified():
                self.set_context(key, context.get("1.0", "end-1c"))
                context.edit_modified(False)
        context.bind("<<Modified>>", change)
        self.context_panels[key] = details
        self.context_toggles[key] = toggle
        self.context_text_widgets[key] = context
        self.page_editable.extend(((toggle, "normal"), (context, "normal")))
        if not expanded:
            details.grid_remove()

    def toggle_context(self, key):
        details = self.context_panels[key]
        if details.winfo_manager():
            details.grid_remove()
            self.context_toggles[key].configure(text="展开：文件未注明的信息（可选）")
        else:
            details.grid()
            self.context_toggles[key].configure(text="收起：文件未注明的信息（可选）")

    def set_context(self, key, value):
        if key not in {"a_context_text", "b_context_text"} or value == self.values[key].get():
            return
        step = "recognize_a" if key == "a_context_text" else "recognize_b"
        self.apply_values({**{field: "" for field in STEP_INVALIDATED[step]}, key: value})

    def build_standard_settings(self):
        project = self.card(self.page)
        project.grid(row=1, column=0, sticky="ew")
        project.columnconfigure(0, weight=1)
        self.label(project, "语义金标准", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, sticky="w")
        current = tk.Frame(project, bg="white")
        current.grid(row=1, column=0, sticky="ew", pady=(7, 0));current.columnconfigure(0, weight=1)
        ttk.Entry(current, textvariable=self.values["gold_path"], state="readonly").grid(row=0, column=0, sticky="ew")
        self.button(current, "打开标准", lambda: self.open_path("gold_path")).grid(row=0, column=1, padx=(8, 0))
        self.button(current, "打开所在文件夹", self.open_standard_folder).grid(row=0, column=2, padx=(8, 0))
        self.label(project, "金标准分别定义指标槽位和维度要求；新客户、新日期等实际取值保存在映射中，不因此新增槽位。确需修订标准时，可导入人工或高级 AI 修订的完整金标准，预览后保存新版本。",
                   fg=MUTED, wraplength=720, justify="left", font=("Microsoft YaHei UI", 9)).grid(row=2, column=0, sticky="ew", pady=(10, 0))
        candidate = tk.Frame(project, bg="white")
        candidate.grid(row=3, column=0, sticky="ew", pady=(9, 0));candidate.columnconfigure(0, weight=1)
        ttk.Entry(candidate, textvariable=self.revision_candidate, state="readonly").grid(row=0, column=0, sticky="ew")
        choose = self.button(candidate, "选择修订文件并预览", self.choose_revision, choose=True)
        choose.grid(row=0, column=1, padx=(8, 0));self.page_editable.append((choose, "normal"))
        preview = tk.Frame(project, bg="white")
        preview.grid(row=4, column=0, sticky="ew", pady=(9, 0));preview.columnconfigure(0, weight=1)
        self.revision_text = tk.Text(preview, height=10, wrap="word", state="disabled", bg="#F8FAFD", relief="flat", padx=8, pady=7, font=("Microsoft YaHei UI", 9))
        self.revision_text.grid(row=0, column=0, sticky="ew")
        scroll = ttk.Scrollbar(preview, command=self.revision_text.yview);scroll.grid(row=0, column=1, sticky="ns")
        self.revision_text.configure(yscrollcommand=scroll.set)
        self.publish_button = self.button(project, "保存并使用新版本", self.publish_gold_revision, primary=True)
        self.publish_button.grid(row=5, column=0, sticky="w", pady=(9, 0))
        self.update_revision_display()
        self.label(project, "结构检查不代表语义正确。请核实新增、修订和停用的含义后再保存。新版本不会覆盖原标准；已引用变更定义的映射需要重新识别。",
                   fg=MUTED, wraplength=720, justify="left", font=("Microsoft YaHei UI", 9)).grid(row=6, column=0, sticky="ew", pady=(9, 0))

    def open_standard_folder(self):
        parent = Path(self.values["gold_path"].get()).parent
        if parent.is_dir():
            try:
                os.startfile(str(parent))
            except OSError as exc:
                messagebox.showerror("无法打开文件夹", str(exc), parent=self.root)

    def choose_revision(self):
        candidate = filedialog.askopenfilename(parent=self.root, title="选择人工或高级 AI 修订后的完整金标准",
                                               filetypes=(("完整金标准", "*.jsonl"), ("所有文件", "*.*")))
        if candidate:
            self.preview_gold_revision(candidate)

    def preview_gold_revision(self, candidate):
        from .金标准 import preview_revision
        self.revision_preview = None
        self.revision_candidate.set(str(candidate))
        self.revision_current = self.values["gold_path"].get()
        try:
            self.revision_preview = preview_revision(self.revision_current, candidate)
        except (ValueError, OSError) as exc:
            messagebox.showerror("金标准修订文件需检查", str(exc), parent=self.root)
        self.update_revision_display()
        return self.revision_preview is not None

    def update_revision_display(self):
        if not hasattr(self, "revision_text") or not self.revision_text.winfo_exists():
            return
        self.revision_text.configure(state="normal")
        self.revision_text.delete("1.0", "end")
        self.revision_text.insert("1.0", format_revision_preview(self.revision_preview))
        self.revision_text.configure(state="disabled")
        self.publish_button.configure(state="normal" if self.revision_preview and not self.running else "disabled")

    def publish_gold_revision(self):
        from .金标准 import publish_revision
        if not self.revision_preview or self.running:
            return
        if self.revision_current != self.values["gold_path"].get():
            self.revision_preview = None;self.update_revision_display()
            messagebox.showinfo("请重新预览", "当前使用的标准已变化，请重新预览修订文件。", parent=self.root)
            return
        try:
            published = publish_revision(self.revision_current, self.revision_candidate.get(), root=PROJECT_DIR,
                                         expected_hashes=self.revision_preview["source_hashes"])
        except (ValueError, OSError) as exc:
            self.revision_preview = None
            self.update_revision_display()
            messagebox.showerror("新金标准未保存", str(exc) + "\n请重新选择修订文件并预览。", parent=self.root)
            return
        self.apply_values(file_selection_updates("gold_path", str(published["path"]), self.config()))
        self.revision_preview = None;self.revision_candidate.set("");self.update_revision_display()
        self.append_log("已保存并使用新金标准：" + str(published["path"]))
        self.append_log("附注表格和更新数据的语义映射已保留；后续执行会逐项检查引用定义是否变化，变化项需重新识别。旧更新清单及最终成果已从当前选择中清空。")
        try:
            save_settings(self.config())
        except (ValueError, OSError) as exc:
            messagebox.showerror("标准已保存，设置尚未保存", "新标准已在本窗口使用，请修正设置后点击保存设置。\n" + str(exc), parent=self.root)
        self.status.configure(text="已使用新版本；引用定义变化的映射需重新识别。", fg=GREEN)

    def change_scope_prefilter(self):
        self.apply_values({key: "" for key in STEP_INVALIDATED["recognize_b"]})
        self.append_log("来源筛选方式已改变，已清空更新数据及后续成果的选择；原成果文件仍保留，请重新识别更新数据。")

    def config(self):
        return {**{key: value.get().strip() for key, value in self.values.items()},
                "scope_prefilter": self.scope_prefilter.get(), "review": True, "current_step": self.current_step if self.current_step in STEPS else "extract_a"}

    def refresh_ready(self):
        if not hasattr(self, "status"):
            return
        for key, button in self.path_open_buttons.items():
            button.configure(state="normal" if self.values[key].get() else "disabled")
        if self.running:
            return
        if self.current_step == "settings":
            self.status.configure(text="设置单独保存，不会启动任何处理步骤。", fg=MUTED)
            return
        messages = missing_inputs(self.current_step, self.config())
        previous = self.step_messages.get(self.current_step)
        if messages:
            self.status.configure(text="尚缺：" + "；".join(messages[:2]), fg="#A65B05")
        elif previous:
            self.status.configure(text=previous, fg=GREEN if self.step_states.get(self.current_step) == "complete" else MUTED)
        else:
            self.status.configure(text="材料已选，可单独执行本步；已有成果也可直接复用。", fg=MUTED)

    def choose_file(self, key):
        title, pattern = FILE_FIELDS[key]
        path = filedialog.askopenfilename(parent=self.root, title="选择" + title,
                                           filetypes=((title, pattern), ("所有文件", "*.*")))
        if path:
            self.select_file(key, path)

    def select_file(self, key, path):
        try:
            updates = file_selection_updates(key, path, self.config())
        except (ValueError, OSError) as exc:
            messagebox.showerror("成果文件需检查", str(exc), parent=self.root)
            return False
        self.apply_values(updates)
        self.refresh_mapping_summary(key)
        if updates:
            self.append_log(("已选择" if path else "已清空") + FILE_FIELDS[key][0] + ("：" + path if path else ""))
            if key in {"a_mapping", "b_mapping", "update_plan"}:
                for source in ("a_path", "b_path", "a_mapping", "b_mapping", "a_link"):
                    if source in updates and updates[source]:
                        self.append_log("关联来源 · " + FILE_FIELDS[source][0] + "：" + updates[source])
            self.show_step(self.current_step)
        return True

    def apply_values(self, updates):
        for key, value in updates.items():
            if key in self.values:
                self.values[key].set(str(value or ""))
        for step, definition in STEPS.items():
            if any(key in updates for key in definition["outputs"] if key not in {"trace", "a_review", "b_review", "a_standard_review", "b_standard_review"}):
                self.step_states.pop(step, None)
                self.step_messages.pop(step, None)
        self.refresh_ready()

    def use_updated_a(self):
        if self.running:
            return
        prime = self.values["a_prime"].get()
        mapping = self.values["a_prime_mapping"].get()
        try:
            if not prime or not Path(prime).is_file() or not mapping or not Path(mapping).is_file():
                raise ValueError("请先生成或选齐更新后附注表格及其语义映射")
            record = _read_record(mapping, "附注语义映射-v1", "A")
            bound = record.get("original_path") or record.get("source_path") or ""
            if Path(bound).resolve() != Path(prime).resolve():
                raise ValueError("当前更新后附注表格与所选语义映射的来源不一致")
            link = record.get("a_link") or ""
            if link and not Path(link).is_file():
                raise ValueError("更新后附注表格语义映射中的 Word 位置关联不存在")
            linked_word = ""
            if link:
                association = _read_record(link, "附注Word关联-v1")
                if Path(association.get("a_path") or "").resolve() != Path(prime).resolve():
                    raise ValueError("下一轮 Word 位置关联与当前更新后附注表格不一致")
                linked_word = association.get("word_path") or ""
        except (ValueError, OSError) as exc:
            messagebox.showinfo("尚不能开始下一轮", str(exc), parent=self.root)
            return
        self.apply_values({"a_path": prime, "a_mapping": mapping, "a_link": link,
                           "word_path": linked_word,
                           "a_review": "", "a_gold_changes": "", "a_standard_review": "", "b_path": "", "b_mapping": "", "b_review": "",
                           "b_scope": "", "b_gold_changes": "", "b_standard_review": "",
                           "a_context_text": record.get("additional_context") or "", "b_context_text": "", "update_plan": "", "trace": "",
                           **{key:'' for key in STRUCTURE_FIELDS}})
        self.append_log("已将更新后附注表格及其语义映射用于下一轮；请选择新的更新数据表。没有自动运行任何步骤。")
        self.show_step("recognize_b")

    def choose_output(self):
        path = filedialog.askdirectory(parent=self.root, title="选择成果保存文件夹", mustexist=True)
        if path:
            self.values["output_dir"].set(path)

    def save(self):
        try:
            config = validate_settings(self.config(), require_model=False)
            save_settings(config)
            self.values["base_url"].set(config["base_url"])
            self.values["model"].set(config["model"])
            self.values["timeout"].set(f"{config['timeout']:g}")
            self.secret = config["api_key"]
            self.append_log("设置和材料位置已保存；没有启动处理步骤。")
            self.status.configure(text="设置已保存", fg=GREEN)
        except (ValueError, OSError) as exc:
            messagebox.showerror("设置未保存", str(exc), parent=self.root)

    def start(self):
        if self.current_step == "settings":
            self.save()
            return
        if self.running or (self.worker and self.worker.is_alive()):
            return
        config = self.config()
        missing = missing_inputs(self.current_step, config)
        if missing:
            messagebox.showinfo("本步缺少材料", "\n".join(missing), parent=self.root)
            return
        try:
            if self.current_step in MODEL_STEPS:
                config = validate_settings(config)
        except ValueError as exc:
            messagebox.showerror("模型设置需检查", str(exc), parent=self.root)
            return
        self.update_step_files(self.current_step, {}, preserve=True)
        self.step_states.pop(self.current_step, None)
        self.step_messages.pop(self.current_step, None)
        self.running_step = self.current_step
        self.nav_buttons[self.running_step].configure(text=STEPS[self.running_step]["number"] + "  " + STEPS[self.running_step]["short"])
        self.secret = config.get("api_key", "")
        self.cancel.clear()
        self.last_output = config["output_dir"]
        self.set_running(True)
        self.status.configure(text="正在执行：" + STEPS[self.running_step]["title"], fg=BLUE)
        self.append_log("开始第 " + STEPS[self.running_step]["number"] + " 步：" + STEPS[self.running_step]["title"])
        self.worker = threading.Thread(target=self.run, args=(self.running_step, config), name="附注分步处理", daemon=False)
        self.worker.start()

    def run(self, step, config):
        try:
            from .分步流程 import run_step
            result = run_step(step, config, lambda message: self.messages.put(("log", str(message))), self.cancel,
                              progress=lambda checkpoint: self.messages.put(("progress", copy.deepcopy(checkpoint))))
            if not isinstance(result, dict):
                raise RuntimeError("当前步骤没有返回结果记录")
            self.messages.put(("done", (step, result)))
        except Exception as exc:
            self.messages.put(("done", (step, {"status": "failed", "message": f"{type(exc).__name__}：{exc}", "files": {}})))

    def set_running(self, running):
        self.running = running
        for widget, state in self.global_editable + self.page_editable:
            widget.configure(state="disabled" if running else state)
        for widget in self.nav_buttons.values():
            widget.configure(state="disabled" if running else "normal")
        self.start_button.configure(state="disabled" if running else "normal")
        self.cancel_button.configure(state="normal" if running else "disabled")
        if running:
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.configure(value=0)
            self.refresh_ready()

    def update_step_files(self, step, files, *, preserve=False):
        clear = STEP_INVALIDATED[step]
        if preserve and step in MODEL_STEPS:
            mapping = "a_mapping" if step == "recognize_a" else "b_mapping"
            if not files.get(mapping) or str(files[mapping]) == self.values[mapping].get():
                # 同一次材料继续识别时，未有新映射就保留此前可接续的本步成果。
                clear = tuple(key for key in clear if key not in STEPS[step]["outputs"])
        self.apply_values({key: "" for key in clear})
        for key, path in files.items():
            if key in self.values and isinstance(path, (str, Path)) and str(path):
                self.values[key].set(str(path))

    def save_reusable_mapping(self, step):
        if step not in MODEL_STEPS:
            return
        key = "a_mapping" if step == "recognize_a" else "b_mapping"
        if not Path(self.values[key].get()).is_file():
            return
        try:
            save_settings(self.config())
        except (ValueError, OSError) as error:
            self.append_log("映射文件已保存，但接续路径设置保存失败：" + str(error))

    def apply_progress(self, result):
        step = result.get("step")
        if (not self.running or step != self.running_step or step not in MODEL_STEPS
                or result.get("status") != "running" or result.get("checkpoint") is not True):
            return
        files = result.get("files") or {}
        mapping = "a_mapping" if step == "recognize_a" else "b_mapping"
        path = files.get(mapping)
        if not isinstance(path, (str, Path)) or not str(path) or not Path(path).is_file():
            return
        # 仅接收本步文件及后端已核实发布的新标准；不从映射或候选提案猜测标准版本。
        updates = {key: value for key, value in files.items() if key in STEPS[step]["outputs"]}
        published_gold = files.get("gold_path")
        if isinstance(published_gold, (str, Path)) and str(published_gold) and Path(published_gold).is_file():
            updates["gold_path"] = str(published_gold)
        self.update_step_files(step, updates, preserve=True)
        self.last_output = result.get("output_dir") or self.last_output
        counts = result.get("counts") or {}
        details = "；".join(label + " " + str(counts[key]) + " 格" for key, label in
                           (("mappings", "业务"), ("excluded", "非业务"), ("unresolved", "待核实"))
                           if type(counts.get(key)) is int and counts[key] >= 0)
        message = str(result.get("message") or "已保存本步部分映射，可继续识别")
        if details:
            message += "；" + details
        if self.cancel.is_set():
            message = "正在取消；" + message
        self.step_states[step] = "running"
        self.step_messages[step] = message
        self.status.configure(text=message, fg=BLUE)
        self.save_reusable_mapping(step)

    def apply_result(self, step, result):
        files = result.get("files") or {}
        self.update_step_files(step, files, preserve=True)
        if files.get("gold_path"):
            try:
                save_settings(self.config())
                self.append_log("已保存并启用本次扩充的金标准，下次使用时继续复用。")
            except (ValueError,OSError) as error:
                self.append_log("新金标准文件已保存，但默认设置保存失败："+str(error))
        if not files.get("gold_path"):
            self.save_reusable_mapping(step)
        self.last_output = result.get("output_dir") or self.last_output
        status = str(result.get("status", "partial")).lower()
        self.step_states[step] = status
        self.step_messages[step] = str(result.get("message") or ("本步已完成" if status == "complete" else "请查看执行记录"))
        self.set_running(False)
        self.show_step(self.current_step)
        if files and self.current_step == step:
            self.root.after_idle(self.scroll_to_outputs)
        self.append_log(self.step_messages[step])
        for key, path in files.items():
            if key in FILE_FIELDS and path:
                self.append_log(FILE_FIELDS[key][0] + "：" + str(path))
        if self.last_output:
            self.append_log("本步成果文件夹：" + str(self.last_output))

    def append_log(self, text):
        text = str(text)
        secret = self.secret or self.values["api_key"].get()
        if secret:
            text = text.replace(secret, "[密钥已隐藏]")
        self.log.configure(state="normal")
        self.log.insert("end", datetime.now().strftime("%H:%M:%S") + "  " + text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def poll(self):
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                if kind == "log":
                    self.append_log(payload)
                elif kind == "progress":
                    self.apply_progress(payload)
                elif kind == "done":
                    self.apply_result(*payload)
        except queue.Empty:
            pass
        if self.closing and not (self.worker and self.worker.is_alive()):
            self.root.destroy()
            return
        self.root.after(100, self.poll)

    def request_cancel(self):
        self.cancel.set()
        self.cancel_button.configure(state="disabled")
        self.status.configure(text="正在取消当前步骤，已有成果会保留…", fg=MUTED)
        self.append_log("已请求取消本步。")

    def open_path(self, key):
        value = self.values[key].get()
        if value and Path(value).exists():
            try:
                os.startfile(value)
            except OSError as exc:
                messagebox.showerror("无法打开", str(exc), parent=self.root)
        else:
            messagebox.showinfo("文件尚未就绪", "请先生成或选择已有文件；当前路径不存在。", parent=self.root)

    def open_last_output(self):
        path = self.last_output or self.values["output_dir"].get()
        if path and Path(path).is_dir():
            try:
                os.startfile(path)
            except OSError as exc:
                messagebox.showerror("无法打开", str(exc), parent=self.root)
        else:
            messagebox.showinfo("尚无成果文件夹", "完成一个步骤后，可在这里打开它的成果文件夹。", parent=self.root)

    def close(self):
        if self.worker and self.worker.is_alive():
            self.closing = True
            self.request_cancel()
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    Application(root)
    root.mainloop()
