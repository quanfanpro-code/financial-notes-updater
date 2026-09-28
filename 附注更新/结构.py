"""根据已确认语义规划小范围行列增减，在新Excel副本上验证，不修改原A。"""
from __future__ import annotations

import copy
import hashlib
import json
import pathlib
import re
import uuid

from openpyxl.utils.cell import get_column_letter, range_boundaries

from .语义 import SemanticCancelled
from .表格 import file_hash, read_workbook, semantic_key, write_updated_workbook

STRUCTURE_PROMPT = """你负责根据共同金标准语义为附注A规划必要的行列增减。A是目标样式，B是更新数据。
只有给定A/B已确认语义可作为依据。不能根据数值、行数或名字相近猜测对应。不能把B其他业务子表全部加入A。
保持A已有披露；B仍有同语义时不得删除。合计、小计、仍需的表头不得删除。插入继承A真实明细行/指标列的样式。
以下ranges是A可调整的命名区域，index、template_index均是每步操作之前该区域内的0基位置。每个操作count=1。
操作按列表顺序执行；bindings和labels的目标地址是全部操作完成后的最终Excel地址。
只输出JSON：
{"operations":[{"range_name":"名称","axis":"row或column","action":"insert或delete","index":0,"count":1,"template_index":0,"reason":"业务语义依据"}],
"bindings":[{"source_sheet":"B表名","source_cell":"B地址","target_sheet":"A表名","target_cell":"最终目标地址","template_sheet":"原A表名","template_cell":"原A已识别业务格","reason":"新增语义依据"}],
"labels":[{"sheet":"A表名","cell":"最终地址","value":"新增标签文字","source_sheet":"B原始表名","source_cell":"B原始标签地址","dimension_key":"如取自B已识别维度则填键名，否则省略","reason":"标签依据"}],
"unresolved":[]}。
原A映射由程序机械移位，不要在bindings中重复；bindings仅增加新插入行列中的业务格。
bindings只能引用B已识别格和原A样式模板，不要返回slot_id、dimensions或任何财务数值。程序从B重新取得业务身份。
labels只能写新插入格，内容必须逐字来自B真实文本或B已确认维度；需要哪些标签就逐项列出，不能依赖复制旧客户名称。
不得在插入后的空格中遗留原模板客户、日期、旧金额。新表头与新对象的必要文字均要有来源证据。
无法证明安全时返回unresolved，不要用手工限制替代可实现的正确方案。"""


def _cell(snapshot, sheet, address):
    for item in snapshot.get("sheets", []):
        if item["name"] == sheet:
            return item.get("cells", {}).get(address)
    return None


def _point(address):
    match = re.fullmatch(r"([A-Z]+)([1-9]\d*)", str(address))
    if not match:
        raise ValueError("无效单元格坐标")
    col = 0
    for char in match[1]:
        col = 26 * col + ord(char) - 64
    return int(match[2]) - 1, col - 1


def _address(point):
    return get_column_letter(point[1] + 1) + str(point[0] + 1)


def _bounds(text):
    c1, r1, c2, r2 = range_boundaries(text.replace("$", ""))
    return r1 - 1, c1 - 1, r2 - 1, c2 - 1


def _in(point, bounds):
    return bounds[0] <= point[0] <= bounds[2] and bounds[1] <= point[1] <= bounds[3]


def _strings(value):
    if isinstance(value, str):
        return {value}
    if isinstance(value, dict):
        return set().union(*(_strings(v) for v in value.values())) if value else set()
    if isinstance(value, list):
        return set().union(*(_strings(v) for v in value)) if value else set()
    return set()


def _key(mapping):
    return json.dumps(semantic_key(mapping), ensure_ascii=False, sort_keys=True, default=str)


def _summary(mapping):
    return {k: mapping[k] for k in ("sheet", "cell", "slot_id", "scope", "dimensions", "semantic_field", "value_type", "row_label", "column_label", "table_range", "table_semantic", "aggregation") if k in mapping}


def _normal_proposal(raw):
    if not isinstance(raw, dict) or any(not isinstance(raw.get(key), list) for key in ("operations", "bindings", "labels", "unresolved")):
        raise ValueError("结构回执必须包含操作、绑定、标签和待核实清单")
    result = copy.deepcopy(raw)
    for operation in result["operations"]:
        if operation.get("axis") not in {"row", "column"} or operation.get("action") not in {"insert", "delete"}:
            raise ValueError("无效的行列增减操作")
        if type(operation.get("index")) is not int or operation.get("count") != 1:
            raise ValueError("结构操作须按一行或一列逐项指定0基位置")
        if not operation.get("reason"):
            raise ValueError("结构调整缺少业务依据")
    return result


def _decision(proposal):
    clean = {}
    for group in ("operations", "bindings", "labels"):
        clean[group] = [{k: v for k, v in item.items() if k != "reason"} for item in proposal[group]]
        if group != "operations":
            clean[group].sort(key=lambda x: json.dumps(x, ensure_ascii=False, sort_keys=True))
    clean["unresolved"] = proposal["unresolved"]
    return clean


def _ranges(snapshot):
    result = {}
    for name, destinations in snapshot.get("names", {}).items():
        if name.startswith("_xlnm") or not isinstance(destinations, list) or len(destinations) != 1:
            continue
        target = destinations[0]
        if not isinstance(target, dict) or not target.get("sheet") or not target.get("range"):
            continue
        bounds = _bounds(target["range"])
        result[name] = {"sheet": target["sheet"], "range": target["range"], "bounds": bounds, "height": bounds[2] - bounds[0] + 1, "width": bounds[3] - bounds[1] + 1}
    return result


def _mapping_range(mapping, ranges):
    candidates = [(name, area) for name, area in ranges.items() if mapping["sheet"] == area["sheet"] and _in(_point(mapping["cell"]), area["bounds"])]
    exact = [(name, area) for name, area in candidates if mapping.get("table_range", "").replace("$", "") == area["range"].replace("$", "")]
    choices = exact or candidates
    if len(choices) != 1:
        raise ValueError(f"{mapping['sheet']}!{mapping['cell']}不能唯一确定Word命名表范围")
    return choices[0][0]


def _simulate(a, amaps, bmaps, ranges, operations):
    states = copy.deepcopy(ranges)
    a_by_position = {(m["sheet"], m["cell"]): m for m in amaps}
    b_keys = {_key(m) for m in bmaps}
    for name, state in states.items():
        state["originals"] = {}
        state["added"] = {}
        for sheet in a["sheets"]:
            if sheet["name"] != state["sheet"]:
                continue
            for address in sheet["cells"]:
                point = _point(address)
                if _in(point, state["bounds"]):
                    state["originals"][address] = (point[0] - state["bounds"][0], point[1] - state["bounds"][1])
    for mapping in amaps:
        _mapping_range(mapping, ranges)
    for operation in operations:
        name = operation.get("range_name")
        if name not in states:
            raise ValueError("结构操作引用不存在或不唯一的命名区域")
        state = states[name]
        axis = 0 if operation["axis"] == "row" else 1
        size = state["height"] if axis == 0 else state["width"]
        index = operation["index"]
        if not 0 <= index <= size or (operation["action"] == "delete" and index >= size):
            raise ValueError("结构操作位置超出命名表")
        on_slice = [(address, a_by_position.get((state["sheet"], address))) for address, point in state["originals"].items() if point[axis] == index]
        if operation["action"] == "delete":
            removed = [m for _, m in on_slice if m is not None]
            if index == 0 or not removed:
                raise ValueError("不能删除表首标签或没有已确认业务含义的行列")
            if any(not any(other["slot_id"].rsplit("-S", 1)[0] == m["slot_id"].rsplit("-S", 1)[0] and other.get("scope") == m.get("scope") for other in bmaps) for m in removed):
                raise ValueError("更新数据表缺少整个对应业务表，不能把未提供解释为删除")
            if any(_key(m) in b_keys for m in removed):
                raise ValueError("拟删除行列中仍有更新数据表明确披露的同语义数据")
            if any("total" in str(m.get("aggregation", "")) or any(word in m.get("row_label", "") for word in ("合计", "小计")) for m in removed):
                raise ValueError("不能自动删除仍需保留的合计或小计")
            allowed = set().union(*(_strings(m.get("dimensions", {})) | {m.get("row_label", ""), m.get("column_label", "")} for m in removed))
            for address, mapping in on_slice:
                value = _cell(a, state["sheet"], address).get("value")
                if mapping is None and value not in (None, "") and (not isinstance(value, str) or value not in allowed):
                    raise ValueError("拟删除范围包含不能由被移除语义解释的原文或数据")
            if any(point[axis] == index for point in state["added"]):
                raise ValueError("同一方案不能先插入后再删除该新行列")
            delta = -1
        else:
            template = operation.get("template_index")
            if type(template) is not int or not 0 <= template < size:
                raise ValueError("插入行列缺少有效的原表样式模板")
            template_maps = [a_by_position[(state["sheet"], address)] for address, point in state["originals"].items() if point[axis] == template and (state["sheet"], address) in a_by_position]
            component = [m for m in template_maps if "total" not in str(m.get("aggregation", "")) and not any(word in m.get("row_label", "") for word in ("合计", "小计"))]
            if not component or (axis == 0 and len(component) != len(template_maps)):
                raise ValueError("新增明细必须继承原表已确认的明细行或指标列样式")
            inverse = {point: address for address, point in state["originals"].items()}
            inverse.update(state["added"])
            copied = {}
            for other in range(state["width"] if axis == 0 else state["height"]):
                source = (template, other) if axis == 0 else (other, template)
                target = (index, other) if axis == 0 else (other, index)
                copied[target] = inverse.get(source)
            delta = 1
        def shift(point):
            mutable = list(point)
            if mutable[axis] >= index:
                mutable[axis] += delta
            return tuple(mutable)
        state["originals"] = {address: shift(point) for address, point in state["originals"].items() if not (delta < 0 and point[axis] == index)}
        state["added"] = {shift(point): source for point, source in state["added"].items()}
        if delta > 0:
            state["added"].update(copied)
        state["height" if axis == 0 else "width"] += delta
    return states


def _check_bindings(proposal, a, b, amaps, bmaps, ranges, states):
    a_positions = {(m["sheet"], m["cell"]): m for m in amaps}
    b_positions = {(m["sheet"], m["cell"]): m for m in bmaps}
    targets = set()
    checked = []
    for binding in proposal["bindings"]:
        source = (binding.get("source_sheet"), binding.get("source_cell"))
        template = (binding.get("template_sheet"), binding.get("template_cell"))
        target = (binding.get("target_sheet"), binding.get("target_cell"))
        if source not in b_positions or template not in a_positions:
            raise ValueError("新增绑定必须引用更新数据表的已识别单元格和附注表格的已识别样式模板")
        if target in targets:
            raise ValueError("多个新增语义不能写入同一目标格")
        targets.add(target)
        src, dst = b_positions[source], a_positions[template]
        if src.get("scope") != dst.get("scope") or src.get("value_type") != dst.get("value_type"):
            raise ValueError("新增语义与目标模板的口径或实际值类型不一致")
        if src["slot_id"].rsplit("-S", 1)[0] != dst["slot_id"].rsplit("-S", 1)[0]:
            raise ValueError("新增语义不属于原模板的同一金标准业务表")
        name = _mapping_range(dst, ranges)
        if target[0] != states[name]["sheet"]:
            raise ValueError("新增语义不能写到目标模板之外的工作表")
        _point(target[1])
        checked.append({"binding": binding, "source": src, "template": dst, "range_name": name})
    labels = []
    for label in proposal["labels"]:
        if not isinstance(label.get("value"), str) or not label.get("value"):
            raise ValueError("新增标签必须是有依据的文本")
        source = (label.get("source_sheet"), label.get("source_cell"))
        source_cell = _cell(b, *source)
        if source_cell is None:
            raise ValueError("新增标签没有真实的更新数据来源单元格")
        if label.get("dimension_key"):
            mapping = b_positions.get(source)
            value = (mapping or {}).get("dimensions", {}).get(label["dimension_key"])
        else:
            value = source_cell.get("value")
        if value != label["value"]:
            raise ValueError("新增标签不等于更新数据表原文或已确认维度，拒绝捏造")
        linked = set().union(*(_strings(m["source"].get("dimensions", {})) | {m["source"].get("row_label", ""), m["source"].get("column_label", "")} for m in checked)) if checked else set()
        if label["value"] not in linked:
            raise ValueError("新增标签与新增绑定的业务对象或指标无关")
        target = (label.get("sheet"), label.get("cell"))
        _point(target[1])
        if target in targets:
            raise ValueError("新增标签与其他绑定或标签位置冲突")
        targets.add(target)
        labels.append({key: label[key] for key in ("sheet", "cell", "value", "reason") if key in label})
    return checked, labels


def resolve_structure(engine, a, b, a_result, b_result, initial_plan, work_dir):
    """返回经真实布局副本验证的新A；任何不确定项都留issues，由主流程阻止发布。"""
    fallback = {"source_a": a["path"], "a_snapshot": a, "a_mappings": copy.deepcopy(a_result.get("mappings", [])), "operations": [], "issues": []}
    try:
        engine._check_cancel()
        if not a_result.get("complete") or not b_result.get("complete"):
            raise ValueError("附注表格或更新数据表还有未完成的语义识别，不能据此调整结构")
        amaps, bmaps = a_result["mappings"], b_result["mappings"]
        ranges = _ranges(a)
        if not ranges:
            raise ValueError("附注表格缺少可追溯到Word的唯一命名表范围，暂不能自动增减结构")
        a_keys, b_keys = {_key(m) for m in amaps}, {_key(m) for m in bmaps}
        extra_tables = {m["slot_id"].rsplit("-S", 1)[0] for m in bmaps if _key(m) not in a_keys}
        relevant_names = {_mapping_range(m, ranges) for m in amaps if _key(m) not in b_keys or m["slot_id"].rsplit("-S", 1)[0] in extra_tables}
        if not relevant_names:
            return fallback
        related_a = [m for m in amaps if _mapping_range(m, ranges) in relevant_names]
        related_tables = {m["slot_id"].rsplit("-S", 1)[0] for m in related_a}
        related_b = [m for m in bmaps if m["slot_id"].rsplit("-S", 1)[0] in related_tables]
        visible_ranges = {name: {k: v for k, v in value.items() if k != "bounds"} for name, value in ranges.items() if name in relevant_names}
        labels = {"A": [], "B": []}
        for label, snapshot in (("A", a), ("B", b)):
            for sheet in snapshot["sheets"]:
                relevant_mappings = related_a if label == "A" else related_b
                areas = [_bounds(m["table_range"]) for m in relevant_mappings if m["sheet"] == sheet["name"] and m.get("table_range")]
                labels[label].extend({"sheet": sheet["name"], "cell": address, "text": c["value"]} for address, c in sheet["cells"].items() if any(_in(_point(address), area) for area in areas) and isinstance(c.get("value"), str) and not c.get("formula") and not re.fullmatch(r"[\d,.%％()（）\s+-]+", c["value"]))
        proposals = []
        for round_number in (1, 2):
            engine._check_cancel()
            payload = {"task": "structure", "round": round_number, "ranges": visible_ranges, "a_mappings": [_summary(m) for m in related_a], "b_mappings": [_summary(m) for m in related_b], "labels": labels, "instruction": "独立从原始语义判断，不提供另一轮方案"}
            proposals.append(engine._request(STRUCTURE_PROMPT, payload, _normal_proposal))
        if any(p["unresolved"] for p in proposals):
            raise ValueError("结构方案仍有待核实事项：" + json.dumps([p["unresolved"] for p in proposals], ensure_ascii=False))
        if _decision(proposals[0]) != _decision(proposals[1]):
            raise ValueError("两轮独立结构方案不一致，未执行行列调整")
        proposal = proposals[0]
        if not proposal["operations"]:
            if proposal["bindings"] or proposal["labels"]:
                raise ValueError("没有插入行列时不能新增绑定或改写原标签")
            return fallback
        if any(op["range_name"] not in relevant_names for op in proposal["operations"]):
            raise ValueError("结构方案修改了没有语义差异的业务表")
        context = copy.deepcopy(a.get("context", []))
        context_tables = context.get("tables", []) if isinstance(context, dict) else context
        context_tables = [table for table in context_tables if isinstance(table, dict)] if isinstance(context_tables, list) else []
        changed_names = {op["range_name"] for op in proposal["operations"]}
        for table in context_tables:
            if table.get("structural_omissions"):
                if table.get("range_name") not in ranges:
                    raise ValueError("Word 结构省略位置缺少可追溯的命名表范围")
                if table["range_name"] in changed_names:
                    raise ValueError("含 Word 结构省略位置的表格暂不能自动增减行列")
        states = _simulate(a, amaps, bmaps, ranges, proposal["operations"])
        bindings, new_labels = _check_bindings(proposal, a, b, amaps, bmaps, ranges, states)
        for snapshot in (a, b):
            if read_workbook(snapshot["path"])["sha256"] != snapshot["sha256"]:
                raise ValueError("结构规划期间输入文件发生变化，请重新读取")
        engine._check_cancel()
        work = pathlib.Path(work_dir)
        work.mkdir(parents=True, exist_ok=True)
        output = work / ("附注表格_结构调整_" + uuid.uuid4().hex + pathlib.Path(a["path"]).suffix)
        write_updated_workbook(a["path"], str(output), {"operations": proposal["operations"], "labels": new_labels, "updates": []})
        updated = read_workbook(str(output))
        final_ranges = _ranges(updated)
        for table in context_tables:
            name = table.get("range_name")
            if name not in ranges:
                continue
            if name not in final_ranges:
                raise ValueError("结构调整后 Word 表格关联范围丢失")
            old, final = ranges[name], final_ranges[name]
            row_shift = final["bounds"][0] - old["bounds"][0]
            column_shift = final["bounds"][1] - old["bounds"][1]
            table.update(sheet=final["sheet"], first_row=final["bounds"][0] + 1,
                         first_column=final["bounds"][1] + 1, row_count=final["height"],
                         column_count=final["width"])
            for omission in table.get("structural_omissions", []):
                top, left, bottom, right = _bounds(omission["range"])
                start = _address((top + row_shift, left + column_shift))
                end = _address((bottom + row_shift, right + column_shift))
                omission.update(sheet=final["sheet"], range=start if start == end else start + ":" + end)
        if context:
            updated = read_workbook(str(output), context=context)
        for name, state in states.items():
            final = final_ranges.get(name)
            if final is None or final["height"] != state["height"] or final["width"] != state["width"]:
                raise ValueError("实际命名表范围与结构方案尺寸不一致")
        def absolute(name, relative):
            area = final_ranges[name]
            return area["sheet"], _address((area["bounds"][0] + relative[0], area["bounds"][1] + relative[1]))
        added_positions = {}
        for name, state in states.items():
            for relative, original in state["added"].items():
                added_positions[absolute(name, relative)] = (name, original)
        outmaps = []
        for mapping in amaps:
            name = _mapping_range(mapping, ranges)
            if mapping["cell"] not in states[name]["originals"]:
                continue
            target = absolute(name, states[name]["originals"][mapping["cell"]])
            old_cell, new_cell = _cell(a, mapping["sheet"], mapping["cell"]), _cell(updated, *target)
            if new_cell is None or (not old_cell.get("formula") and old_cell.get("value") != new_cell.get("value")):
                raise ValueError("结构调整改变了既有业务格内容或使其丢失")
            item = copy.deepcopy(mapping)
            item.update({"sheet": target[0], "cell": target[1], "table_range": final_ranges[name]["range"]})
            outmaps.append(item)
        mapped_added = set()
        for detail in bindings:
            binding, source, template, name = detail["binding"], detail["source"], detail["template"], detail["range_name"]
            target = (binding["target_sheet"], binding["target_cell"])
            if target not in added_positions or added_positions[target] != (name, template["cell"]):
                raise ValueError("新增绑定未落在对应模板新插入的业务格内")
            if _cell(updated, *target) is None:
                raise ValueError("新增业务格在实际布局副本中不存在")
            item = copy.deepcopy(source)
            item["dimensions"] = copy.deepcopy(source.get("dimensions", {}))
            for dimension in ("unit", "scale"):
                if dimension in template.get("dimensions", {}):
                    item["dimensions"][dimension] = template["dimensions"][dimension]
            item.update({"sheet": target[0], "cell": target[1], "table_range": final_ranges[name]["range"], "table_semantic": template.get("table_semantic", ""), "reason": binding["reason"], "reviewed": True})
            if any(_key(old) == _key(item) and _mapping_range(old, final_ranges) == name for old in outmaps):
                raise ValueError("同一目标表新增了已经存在的同语义业务格")
            outmaps.append(item)
            mapped_added.add(target)
        label_positions = {(item["sheet"], item["cell"]) for item in new_labels}
        for label in new_labels:
            target = (label["sheet"], label["cell"])
            if target not in added_positions or _cell(updated, *target).get("value") != label["value"]:
                raise ValueError("新增标签未准确写入新插入的格子")
        original_mapping_positions = {(m["sheet"], m["cell"]) for m in amaps}
        for target, (name, original) in added_positions.items():
            if (states[name]["sheet"], original) in original_mapping_positions and target not in mapped_added:
                raise ValueError("新插入行列尚有原模板业务格没有对应的更新数据语义")
            if original is not None:
                old_cell = _cell(a, states[name]["sheet"], original)
                if isinstance(old_cell.get("value"), str) and not old_cell.get("formula") and old_cell["value"].strip() and target not in label_positions and target not in mapped_added:
                    raise ValueError("新插入位置的对象或表头标签未提供更新数据依据，不能沿用旧文字")
        engine._check_cancel()
        return {"source_a": str(output), "a_snapshot": updated, "a_mappings": outmaps, "operations": proposal["operations"], "issues": []}
    except SemanticCancelled:
        raise
    except Exception as error:
        fallback["issues"] = [str(error)]
        return fallback

UNUSED_SOURCE_PROMPT = """复核更新来源中尚无接收位置的业务数据，判断是否属于本次附注明确不列示的额外指标。
这是财务披露范围判断，不能把真实业务数据称为非业务格。不能仅因槽位、表ID不同或附注未出现该行就判为不采用。
只有原附注真实业务边界与来源原文能支持时才列入 unused_sources，必须引用两侧真实文字坐标及原文。
不得根据金额为0、为空、金额大小、来源缺失或旧主体名称决定不采用。不能忽略附注已具有完全相同业务身份的数据。
同一槽位的新客户、新项目、新账龄或新期间通常需要保留并增补结构；不能因为原附注未列出而忽略。
若同槽位来源确实在明确限定的披露范围外，必须给 scope_limit：限定维度、允许值及附注完整限制说明的真实坐标原文。
原表只列“甲客户”或“1年以内”、已有行列标签、维度值本身均不是禁止披露其他客户或账龄的证据。
限制说明须明确“仅/只/限于”的允许范围，allowed_values逐项来自已确认目标维度，且覆盖原表同槽位全部允许值。“不包括甲”是排除条件，不能解读成“仅允许甲”。
不能给出完整范围依据、需要增补行列或有争议的来源放 unresolved；不猜测不存在于输入的披露政策。
独立复核，不参考另一轮结论。只输出JSON：
{"unused_sources":[{"sheet":"更新来源工作表","cell":"格址","reason":"原附注与更新来源业务范围依据",
"evidence":[{"carrier":"A或B","sheet":"工作表","cell":"文字格址","text":"真实完整文字"}],
"scope_limit":{"dimension":"受明确限制的维度名","allowed_values":["原附注允许的维度值"],"evidence":{"carrier":"A","sheet":"工作表","cell":"限制说明格址","text":"真实完整限制说明"}}}],
"unresolved":[]}。不同指标没有scope_limit时可省略该字段，不得返回或推断金额。"""


def _unused_json(value):
    return json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(",",":"),default=str)


def _unused_text(cell):
    value=(cell or {}).get("value")
    return (isinstance(value,str) and bool(value.strip()) and not cell.get("formula")
            and not re.fullmatch(r"[\d\s,.，%％()（）+\-—－–]+(?:元|万元|亿元|户|个)?",value.strip()))


def _unused_fingerprint(mappings):
    rows=[{**_summary(m),"semantic_key":semantic_key(m)} for m in mappings]
    return hashlib.sha256(_unused_json(sorted(rows,key=_unused_json)).encode("utf-8")).hexdigest()


def _unused_inputs(a,b,amaps,bmaps,plan):
    for snapshot,key in ((a,"source_a_hash"),(b,"source_b_hash")):
        if not snapshot.get("sha256") or plan.get(key)!=snapshot["sha256"]:
            raise ValueError("未采用来源清单与当前工作簿身份不一致")
        if file_hash(snapshot["path"])!=snapshot["sha256"]:
            raise ValueError("复核未采用来源时原工作簿已变化")
    positions={}
    for mapping in bmaps:
        point=(mapping["sheet"],mapping["cell"])
        if point in positions:raise ValueError("更新来源同一格存在重复映射，不能判为不采用")
        positions[point]=mapping
    extras={}
    for mapping in plan.get("unplaced_sources",[]):
        point=(mapping["sheet"],mapping["cell"])
        if point in extras or point not in positions or semantic_key(mapping)!=semantic_key(positions[point]):
            raise ValueError("未接收来源的坐标或业务身份与已保存映射不一致")
        extras[point]=positions[point]
    used={(source.get("sheet"),source.get("cell")) for update in plan.get("updates",[]) for source in update.get("sources",[])}
    used.update((binding.get("source_sheet"),binding.get("source_cell")) for binding in plan.get("bindings",[]))
    metadata={"source_a_hash":a["sha256"],"source_b_hash":b["sha256"],
              "a_mappings_hash":_unused_fingerprint(amaps),"b_mappings_hash":_unused_fingerprint(bmaps),"review_count":2}
    return extras,used,metadata


def _unused_evidence(proof,a,b,amaps,source):
    if not isinstance(proof,dict) or proof.get("carrier") not in {"A","B"}:
        raise ValueError("未采用来源缺少附注或更新来源的文字依据")
    carrier=proof["carrier"];snapshot=a if carrier=="A" else b
    sheet,address=proof.get("sheet"),proof.get("cell")
    cell=_cell(snapshot,sheet,address)
    if not _unused_text(cell) or proof.get("text")!=cell["value"]:
        raise ValueError("未采用来源依据不是可回读的原始文字，不能引用数值或伪造文字")
    related=amaps if carrier=="A" else [source]
    related=[m for m in related if m["sheet"]==sheet]
    if not related or not any(not m.get("table_range") or _in(_point(address),_bounds(m["table_range"])) for m in related):
        raise ValueError("未采用来源依据不属于被复核的附注或来源业务范围")
    return {"carrier":carrier,"sheet":sheet,"cell":address,"text":cell["value"]}


def _unused_limit(limit,a,b,amaps,source):
    same=[m for m in amaps if m["slot_id"]==source["slot_id"]]
    if not same and limit is None:return None
    if not isinstance(limit,dict):
        raise ValueError("同槽位新增客户、账龄或期间没有明确披露限制，不能判为未采用")
    dimension=limit.get("dimension");allowed=limit.get("allowed_values")
    if not isinstance(dimension,str) or not isinstance(allowed,list) or not allowed or dimension in {"unit","scale","currency","metric"}:
        raise ValueError("披露限制缺少可核验的业务维度与允许值")
    targets=same or amaps
    values=[m.get("dimensions",{}).get(dimension) for m in targets]
    keys={_unused_json(value) for value in allowed}
    if len(keys)!=len(allowed) or any(value is None for value in values) or keys!={_unused_json(value) for value in values}:
        raise ValueError("披露限制允许值必须与该槽位原附注已确认维度一致")
    source_value=source.get("dimensions",{}).get(dimension)
    if source_value is None or _unused_json(source_value) in keys:
        raise ValueError("拟不采用来源并未处于已明确限定的维度范围之外")
    proof=_unused_evidence(limit.get("evidence"),a,b,targets,source)
    text=proof["text"]
    plain=set().union(*(_strings(m.get("dimensions",{}))|{m.get("row_label",""),m.get("column_label","")} for m in targets))
    if proof["carrier"]!="A" or text.strip() in plain or not re.search(r"仅|限于|只(?:列|披露|包含|包括|涉及|统计)",text) or re.search(r"(?:不|非)(?:仅|只|限于)",text):
        raise ValueError("现有行列标签或维度值不是完整披露限制说明")
    for value in allowed:
        if isinstance(value,str):
            normal=lambda value:re.sub(r"[\s年月日/.-]","",value)
            if not normal(value) or normal(value) not in normal(text):
                raise ValueError("完整限制说明并未列明所声称的允许维度值")
    return {"dimension":dimension,"allowed_values":sorted(copy.deepcopy(allowed),key=_unused_json),"evidence":proof}


def _checked_unused(a,b,amaps,bmaps,plan,saved,verify_saved):
    if not isinstance(saved,list):raise ValueError("未采用来源必须为逐格清单")
    extras,used,metadata=_unused_inputs(a,b,amaps,bmaps,plan)
    akeys={semantic_key(m) for m in amaps};seen=set();records=[]
    for item in saved:
        if not isinstance(item,dict):raise ValueError("未采用来源记录格式错误")
        point=(item.get("sheet"),item.get("cell"))
        if point in seen or point not in extras or point in used:
            raise ValueError("未采用来源重复、不属于剩余来源或已经用于更新绑定")
        source=extras[point];key=semantic_key(source)
        if key in akeys:raise ValueError("原附注已有完全相同业务身份，不能将该来源列为不采用")
        if _cell(b,*point) is None:raise ValueError("未采用来源实际单元格不存在")
        reason=item.get("reason")
        if not isinstance(reason,str) or not reason.strip():raise ValueError("未采用来源缺少业务范围理由")
        evidence=item.get("evidence")
        if not isinstance(evidence,list):raise ValueError("未采用来源缺少两侧文字依据")
        evidence=[_unused_evidence(proof,a,b,amaps,source) for proof in evidence]
        if {proof["carrier"] for proof in evidence}!={"A","B"}:
            raise ValueError("未采用来源必须同时引用附注与更新来源的真实文字")
        if len({_unused_json(proof) for proof in evidence})!=len(evidence):raise ValueError("文字证据重复")
        scope_limit=_unused_limit(item.get("scope_limit"),a,b,amaps,source)
        record={**metadata,"sheet":point[0],"cell":point[1],"semantic_key":key,"reason":reason.strip(),
                "evidence":sorted(evidence,key=_unused_json)}
        if scope_limit is not None:record["scope_limit"]=scope_limit
        if verify_saved and (any(item.get(field)!=value for field,value in metadata.items()) or item.get("semantic_key")!=key):
            raise ValueError("已保存未采用来源的文件、映射身份或双轮核验记录已变化")
        records.append(record);seen.add(point)
    return sorted(records,key=lambda item:(item["sheet"],item["cell"]))


def validate_unused_sources(a,b,a_mappings,b_mappings,plan,saved):
    """对已密封清单重新核对真实文件、语义身份、文字证据与来源是否已经使用，不调用模型。"""
    return _checked_unused(a,b,a_mappings,b_mappings,plan,saved,True)


def _unused_context(amaps,bmaps,a,b,batch):
    # 目录保留全部实际披露，详细上下文按已有共同业务域取回；目录不是ID白名单。
    catalogue={};groups={}
    notes={m["slot_id"].split("-T",1)[0] for m in batch}
    related=[m for m in amaps if m["slot_id"].split("-T",1)[0] in notes]
    for mapping in amaps:
        table=(mapping["sheet"],mapping.get("table_range",""),mapping.get("table_semantic",""))
        catalogue.setdefault(table,{"sheet":table[0],"range":table[1],"meaning":table[2]})
    for mapping in related:
        identity=(mapping["sheet"],mapping.get("table_range",""),mapping["slot_id"],mapping.get("scope"),mapping.get("value_type"))
        group=groups.setdefault(identity,{"sheet":identity[0],"table_range":identity[1],"slot_id":identity[2],"scope":identity[3],
            "value_type":identity[4],"meaning":mapping.get("dimensions",{}).get("metric") or mapping.get("semantic_field"),"dimensions":{}})
        for key,value in mapping.get("dimensions",{}).items():
            group["dimensions"].setdefault(key,{})[_unused_json(value)]=value
    for group in groups.values():
        group["dimensions"]={key:{"values":list(values.values())[:12],"value_count":len(values)} for key,values in group["dimensions"].items()}
    labels=[]
    for carrier,snapshot,mappings in (("A",a,amaps),("B",b,batch)):
        for sheet in snapshot["sheets"]:
            items=[m for m in mappings if m["sheet"]==sheet["name"]]
            full=[m for m in (amaps if carrier=="A" else bmaps) if m["sheet"]==sheet["name"]]
            business_rows={_point(m["cell"])[0] for m in full}
            selected_rows={_point(m["cell"])[0] for m in items}
            if carrier=="A":
                selected_rows={_point(m["cell"])[0] for m in related if m["sheet"]==sheet["name"]}
                selected_rows=set(sorted(selected_rows)[:12])
            for address,cell in sheet["cells"].items():
                point=_point(address)
                if (_unused_text(cell) and (point[0] not in business_rows or point[0] in selected_rows)
                        and any(not m.get("table_range") or _in(point,_bounds(m["table_range"])) for m in items)):
                    labels.append({"carrier":carrier,"sheet":sheet["name"],"cell":address,"text":cell["value"]})
    return {"a_disclosures":list(catalogue.values()),"a_mappings":list(groups.values()),
            "unplaced_sources":[_summary(m) for m in batch],"labels":labels,
            "context_note":"实际披露目录完整。详细指标按共同定义业务域提供，重复对象仅给部分文字和值并标value_count；未展示不代表原附注未披露。若不足以证明不采用，请逐格标为unresolved。"}


def review_unused_sources(engine,a,b,a_mappings,b_mappings,plan):
    """逐包独立双审剩余来源；不修改原表，未覆盖、上下文不足或分歧均留待核实。"""
    result={"unused_sources":[],"issues":[]}
    if not plan.get("unplaced_sources"):return result
    try:
        engine._check_cancel()
        extras,_,_=_unused_inputs(a,b,a_mappings,b_mappings,plan)
    except SemanticCancelled:raise
    except Exception as error:
        result["issues"]=[{"kind":"unused_source_review","reason":str(error)}];return result
    grouped={}
    for mapping in extras.values():grouped.setdefault((mapping["sheet"],mapping.get("table_range","")),[]).append(mapping)
    pending=[items[index:index+24] for items in grouped.values() for index in range(0,len(items),24)]
    while pending:
        batch=pending.pop(0);positions={(m["sheet"],m["cell"]) for m in batch}
        receipt_start=len(getattr(engine,"_review_receipts",[]));issue_start=len(result["issues"])
        try:
            engine._check_cancel()
            context=_unused_context(a_mappings,b_mappings,a,b,batch)
            payload={"task":"review_unused_sources","round":1,**context,
                     "instruction":"本包每个来源必须逐格进入unused_sources或unresolved，不能遗漏；独立依据原始业务含义与文字判断，不提供另一轮结论或金额"}
            if len(json.dumps(payload,ensure_ascii=False))+len(UNUSED_SOURCE_PROMPT)>24000:
                if len(batch)>1:
                    middle=len(batch)//2;pending[0:0]=[batch[:middle],batch[middle:]];continue
                raise ValueError("该来源的披露上下文仍过长，尚不能完整复核其范围，未判为不采用")
            proofs={_unused_json(proof) for proof in context["labels"]}
            def validate(raw):
                if not isinstance(raw,dict) or not isinstance(raw.get("unused_sources"),list) or not isinstance(raw.get("unresolved"),list):
                    raise ValueError("未采用来源复核须返回已确认清单与待核实清单")
                covered=[]
                for item in raw["unused_sources"]+raw["unresolved"]:
                    if not isinstance(item,dict):raise ValueError("剩余来源须按真实坐标逐格复核")
                    covered.append((item.get("sheet"),item.get("cell")))
                if set(covered)!=positions or len(covered)!=len(positions):
                    raise ValueError("该包剩余来源未逐格覆盖、重复或引用了其他包的来源")
                for item in raw["unused_sources"]:
                    evidence=list(item.get("evidence") or [])
                    if item.get("scope_limit"):evidence.append(item["scope_limit"].get("evidence"))
                    if any(_unused_json(proof) not in proofs for proof in evidence):
                        raise ValueError("未采用来源引用了本包未提供的文字，不能跨包猜测证据")
                checked=_checked_unused(a,b,a_mappings,b_mappings,plan,raw["unused_sources"],False)
                unresolved=[]
                for item in raw["unresolved"]:
                    if not isinstance(item.get("reason"),str) or not item["reason"].strip():raise ValueError("待核实来源缺少原因")
                    unresolved.append({"sheet":item["sheet"],"cell":item["cell"],"reason":item["reason"]})
                return {"unused_sources":checked,"unresolved":unresolved}
            decisions=[]
            for round_number in (1,2):
                engine._check_cancel();payload={**payload,"round":round_number}
                decisions.append(engine._request(UNUSED_SOURCE_PROMPT,payload,validate))
            rounds=[{(item["sheet"],item["cell"]):item for item in decision["unused_sources"]} for decision in decisions]
            for point in sorted(positions):
                left,right=rounds[0].get(point),rounds[1].get(point)
                if left and right and {k:v for k,v in left.items() if k!="reason"}=={k:v for k,v in right.items() if k!="reason"}:
                    result["unused_sources"].append(left)
                else:
                    reasons=[item["reason"] for decision in decisions for item in decision["unresolved"] if (item["sheet"],item["cell"])==point]
                    result["issues"].append({"kind":"unused_source_review","sheet":point[0],"cell":point[1],
                        "reason":"两轮未共同确认该来源可不采用"+("："+"；".join(dict.fromkeys(reasons)) if reasons else "")})
        except SemanticCancelled:raise
        except Exception as error:
            result["issues"].extend({"kind":"unused_source_review","sheet":point[0],"cell":point[1],"reason":str(error)} for point in sorted(positions))
        if len(result["issues"])>issue_start and callable(getattr(engine,"_retry_receipts",None)):
            engine._retry_receipts(receipt_start)
    return result
