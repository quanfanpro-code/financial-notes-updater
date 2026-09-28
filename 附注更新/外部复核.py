"""采纳真实人工或高级AI对待决格的既有槽位意见；无模型调用，不修改原件和标准。"""
from __future__ import annotations
import copy
import json
from pathlib import Path
from datetime import datetime
import uuid
from . import 分步流程 as flow
from .语义 import SemanticEngine, _inside, _range, _hash, _unique_json_object
from .金标准 import read_standard, ensure_compatible
from .表格 import file_hash
from .识别范围 import _entry

SCHEMA="附注外部语义复核-v1"
PREVIEW_SCHEMA="附注外部复核预览-v1"
METHOD="external_semantic_review"


def _read(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"),object_pairs_hook=_unique_json_object)


def _text(value):return isinstance(value,str) and bool(value.strip())


def _coverage(record,snapshot):
    expected={(s["name"],a) for s in snapshot["sheets"] for a in s["cells"]}
    positions=[(m.get("sheet"),m.get("cell")) for key in ("mappings","excluded","unresolved","out_of_scope") for m in record.get(key,[])]
    if set(positions)!=expected or len(positions)!=len(set(positions)):
        raise ValueError("原映射存在遗漏、重复或包外格，不能导入外部意见")


def _engine(record,directory):
    flow._check(record["gold_path"],record["gold_hash"],"外部复核依据金标准")
    # 复用原校验器；目录已存在，构造不创建模型任务，也不读取服务配置。
    return SemanticEngine({},record["gold_path"],str(directory))


def _selected_gold(record,gold_path=None):
    selected=copy.deepcopy(record);path=Path(gold_path or record["gold_path"]).resolve()
    ensure_compatible(record,path)
    selected.update(gold_path=str(path),gold_hash=file_hash(path))
    return selected


def _confirmed_scope(record):
    values=[record.get("recognition_basis",{}).get("report_scope")]
    values.extend(item.get("report_scope") for item in record.get("context",[]) if isinstance(item,dict) and "report_scope" in item)
    scopes={value for value in values if isinstance(value,str) and value}
    if any(value is not None and value!="" and not isinstance(value,str) for value in values) or len(scopes)!=1 or not scopes<={"standalone","consolidated","parent"}:
        raise ValueError("补充业务上下文须有一致的原结构化已确认报告口径")
    return next(iter(scopes))


def _new_business_context(record,snapshot,sheet,raw,engine):
    context=raw["business_context"];address=raw["cell"]
    fields={"range","note_ids","scope","table_semantic","header_cells","reason"}
    if not isinstance(context,dict) or set(context)!=fields or not _text(context.get("reason")):
        raise ValueError("补充业务上下文须完整提供单格范围、科目、实际口径、业务含义、原表头及依据")
    if any(key in raw for key in ("replace_mapping_sha256","table_basis_hash")):
        raise ValueError("补充上下文不能更正已确认格或冒用原业务依据标识")
    if raw.get("table_range",context["range"])!=context["range"]:
        raise ValueError("补充上下文与意见单格范围不一致")
    # 必须检查采纳前原对象；不能用采纳后新登记的范围反过来证明自己。
    for area in record.get("confirmed_business_ranges",{}).get(sheet["name"],[]):
        if _inside(address,_range(area["table_range"])):raise ValueError("原格已有业务范围，不能用补充上下文接管")
    for area in sheet.get("confirmed_business_ranges",[]):
        if _inside(address,_range(area["table_range"])):raise ValueError("原格已有业务范围，不能用补充上下文接管")
    if address not in sheet["cells"]:raise ValueError("补充上下文只能引用实际已有原格")
    if record["carrier"]=="A" and not sheet.get("word_table_ranges"):
        raise ValueError("原附注表格缺少可回读的Word物理表关联")
    for merged in sheet.get("merges",[]):
        if _inside(address,_range(merged)) and address!=merged.split(":")[0]:raise ValueError("合并从格不能补充业务上下文")
    if any(_inside(address,_range(item["range"])) for item in sheet.get("structural_omissions",[])):
        raise ValueError("Word结构省略位置不能补充业务上下文")
    selection=record.get("scope_selection")
    if selection and address not in selection.get("selected_cells",{}).get(sheet["name"],[]):
        raise ValueError("补充上下文的原格不在已确认识别范围内")
    scope=_confirmed_scope(record)
    if context["scope"]!=scope:raise ValueError("补充上下文与原已确认报告口径不一致")
    headers=context["header_cells"]
    if not isinstance(headers,list) or not headers or any(not isinstance(a,str) or a not in sheet["cells"] or _entry(sheet["cells"][a]).get("kind")!="text" for a in headers) or len(headers)!=len(set(headers)):
        raise ValueError("补充上下文表头须为不重复的真实原文文字格")
    previous_scope=engine.report_scope
    try:
        engine.report_scope=scope
        engine._validate_cell_table(context,sheet,address)
    finally:engine.report_scope=previous_scope
    return {"source_hash":snapshot["sha256"],"table_range":context["range"],
        **{key:copy.deepcopy(context[key]) for key in ("table_semantic","note_ids","scope","header_cells")}}


def _checked_items(record,snapshot,opinion,mapping_sha256,engine):
    if not isinstance(opinion,dict) or opinion.get("schema")!=SCHEMA:raise ValueError("请选择外部语义复核意见文件")
    for key,expected in (("carrier",record["carrier"]),("mapping_sha256",mapping_sha256),("source_hash",snapshot["sha256"]),("gold_hash",record["gold_hash"])):
        if opinion.get(key)!=expected:raise ValueError("外部意见绑定不一致："+key)
    reviewer=opinion.get("reviewer")
    if not isinstance(reviewer,dict) or not _text(reviewer.get("name")) or reviewer.get("method") not in {"human","advanced_ai"}:
        raise ValueError("请填写真实审核者和人工或高级AI审核方式")
    try:
        if not _text(reviewer.get("reviewed_at")):raise ValueError()
        datetime.fromisoformat(reviewer["reviewed_at"])
    except (ValueError,TypeError):raise ValueError("请填写真实审核时间") from None
    items=opinion.get("items")
    if not isinstance(items,list) or not items:raise ValueError("外部意见没有待采纳的逐格结论")
    _coverage(record,snapshot)
    pending={(m["sheet"],m["cell"]):m for m in record.get("unresolved",[])}
    confirmed={(m["sheet"],m["cell"]):m for m in record.get("mappings",[])}
    sheets={s["name"]:s for s in snapshot["sheets"]};seen=set();changes=[]
    permitted={"sheet","cell","decision","slot_id","scope","dimensions","value_type","semantic_field","row_label","column_label","evidence_cells","metric_evidence","reason","previous_reason","table_range","table_basis_hash","replace_mapping_sha256","business_context"}
    for raw in items:
        if not isinstance(raw,dict) or set(raw)-permitted or raw.get("decision")!="existing":raise ValueError("仅支持既有槽位意见，不能导入确认标志、新标准或排除结论")
        pos=(raw.get("sheet"),raw.get("cell"))
        previous=confirmed.get(pos)
        if pos in seen or (pos not in pending and previous is None):raise ValueError("外部意见引用非业务待决格、非已确认业务格或重复格")
        if previous is not None:
            if raw.get("replace_mapping_sha256")!=_hash(previous):raise ValueError("更正已确认格须明确绑定原映射，不能隐式覆盖")
        elif "replace_mapping_sha256" in raw:raise ValueError("待决格没有可更正的原业务映射")
        original=previous if previous is not None else pending[pos]
        seen.add(pos);sheet=sheets[pos[0]]
        new_context="business_context" in raw
        if new_context and previous is not None:raise ValueError("只允许对原未决格补充业务上下文")
        areas={flow._json(a):a for a in sheet.get("confirmed_business_ranges",[]) if a.get("source_hash")==snapshot["sha256"] and a.get("table_range") and _inside(pos[1],_range(a["table_range"]))}
        if new_context:
            area=_new_business_context(record,snapshot,sheet,raw,engine)
            areas={flow._json(area):area}
        selected=raw.get("table_range")
        if "table_range" in raw:
            if not _text(selected):raise ValueError("审核者选择的原业务范围须为非空地址")
            # 审核者只能选择已有范围；原重叠证据和待决原因仍保存在原映射及意见中。
            areas={key:area for key,area in areas.items() if area["table_range"]==selected}
        if "table_basis_hash" in raw:
            if not _text(selected) or not _text(raw["table_basis_hash"]):raise ValueError("原业务依据标识必须与明确的原范围一起提供")
            areas={key:area for key,area in areas.items() if _hash(area)==raw["table_basis_hash"]}
        if len(areas)!=1:raise ValueError("待决格没有唯一且可回读的原业务范围")
        area=next(iter(areas.values()));table={**copy.deepcopy(area),"range":area["table_range"]}
        if not new_context and selected is None and original.get("table_range") and original["table_range"]!=table["range"]:raise ValueError("复核格与原业务范围不一致")
        engine._validate_layout({"tables":[table]},sheet,[])
        evidence=raw.get("evidence_cells")
        metric=raw.get("metric_evidence",[])
        if not isinstance(evidence,list) or not evidence or any(not isinstance(a,str) for a in evidence) or len(evidence)!=len(set(evidence)):raise ValueError("请提供不重复的原始文字证据格")
        if not isinstance(metric,list) or any(not isinstance(a,str) for a in metric):raise ValueError("指标依据须为原始文字格地址清单")
        for address in evidence+metric:
            if not isinstance(address,str) or address not in sheet["cells"] or _entry(sheet["cells"][address]).get("kind")!="text":raise ValueError("复核证据必须是可回读的原始文字，不能用数字或空白证明")
            if address not in table.get("header_cells",[]) and not _inside(address,_range(table["range"])):raise ValueError("复核文字证据不属于原业务范围或原表头")
            if any(_inside(address,_range(p["range"])) and not _inside(pos[1],_range(p["range"])) for p in sheet.get("word_table_ranges",[])):raise ValueError("复核证据不得借用另一张Word物理表")
        allowed={slot["id"] for slot in engine._candidate_slots(table)}
        suggestion=copy.deepcopy(raw)
        if not suggestion.get("semantic_field") and suggestion.get("slot_id") in engine.slots:suggestion["semantic_field"]=engine.slots[suggestion["slot_id"]]["slot"]["name"]
        checked=engine._validate_classification({"mappings":[suggestion],"excluded":[],"unresolved":[]},[pos[1]],allowed,sheet,table)[pos[1]][1]
        if "table_basis_hash" in raw or new_context:checked["table_basis_hash"]=_hash(area)
        change={"sheet":pos[0],"cell":pos[1],"previous_reason":original.get("reason",""),"mapping":checked}
        if previous is not None:change["previous_mapping"]=copy.deepcopy(previous)
        if new_context:change.update(business_context=copy.deepcopy(raw["business_context"]),business_range=area,
            source_value=copy.deepcopy(sheet["cells"][pos[1]]["value"]),
            context_evidence=[{"cell":a,"text":sheet["cells"][a]["value"]} for a in area["header_cells"]])
        changes.append(change)
    return changes


def export_review_template(mapping_path,output_path,gold_path=None,*,include_confirmed=False):
    if _read(mapping_path).get('schema_version')==2:
        from .统一外部复核 import export_review_template as unified_export
        return unified_export(mapping_path,output_path,gold_path,include_confirmed=include_confirmed)
    record,snapshot=flow._load_mapping(mapping_path,complete=False);_coverage(record,snapshot)
    # 更正绑定磁盘原对象，不能用加载器规范后的等价表示替代原记录。
    record["mappings"]=_read(mapping_path)["mappings"]
    record=_selected_gold(record,gold_path)
    value={"schema":SCHEMA,"carrier":record["carrier"],"mapping_sha256":file_hash(mapping_path),"source_hash":record["source_hash"],"gold_hash":record["gold_hash"],
           "reviewer":{"name":"","method":"","reviewed_at":""},"items":[{"sheet":item["sheet"],"cell":item["cell"],"previous_reason":item.get("reason",""),"decision":"existing","slot_id":"","scope":"","dimensions":{},"value_type":"","evidence_cells":[],"reason":""} for item in record.get("unresolved",[])]}
    if include_confirmed:
        value["confirmed_mapping_options"]=copy.deepcopy(record.get("mappings",[]))
        value["items"].extend({"sheet":item["sheet"],"cell":item["cell"],"previous_reason":item.get("reason",""),"replace_mapping_sha256":_hash(item),"decision":"existing","slot_id":"","scope":"","dimensions":{},"value_type":"","evidence_cells":[],"reason":""} for item in record.get("mappings",[]))
    # 业务说明和原表头供审核者选择；标识仅绑定原对象，不生成新的布局证据。
    value["business_range_options"]=[{"sheet":sheet["name"],"range":area["table_range"],
        "table_semantic":area["table_semantic"],"header_cells":copy.deepcopy(area.get("header_cells",[])),
        "note_ids":copy.deepcopy(area["note_ids"]),"scope":area["scope"],"basis_hash":_hash(area)}
        for sheet in snapshot["sheets"] for area in sheet.get("confirmed_business_ranges",[])]
    value["business_context_instructions"]={"applies_to":"仅限原未决、没有任何原业务范围覆盖且能匹配现有有效槽位的真实单格；有原范围时仍从business_range_options选择。",
        "rules":"在该格items中增加business_context。实际口径须与原已确认结构化事实一致；header_cells只填同表原文字地址。不得更改已排除、章外或合并从格，不得同时填写table_basis_hash或replace_mapping_sha256。原文字由程序回读，采纳前显示预览。",
        "example":{"range":"本格:本格","note_ids":["既有科目ID"],"scope":"原已确认实际口径",
            "table_semantic":"该原格的真实业务总体及层级","header_cells":["原行项目地址","原期间表头地址"],"reason":"原行、列及总额或分项关系的具体依据"}}
    output=Path(output_path).resolve()
    with output.open("x",encoding="utf-8-sig") as stream:json.dump(value,stream,ensure_ascii=False,indent=2)
    return str(output)


def preview_external_review(mapping_path,opinion_path,gold_path=None):
    if _read(mapping_path).get('schema_version')==2:
        from .统一外部复核 import preview_external_review as unified_preview
        return unified_preview(mapping_path,opinion_path,gold_path)
    mapping=Path(mapping_path).resolve();opinion_file=Path(opinion_path).resolve()
    digest=file_hash(mapping);opinion_hash=file_hash(opinion_file)
    record,snapshot=flow._load_mapping(mapping,complete=False);opinion=_read(opinion_file)
    record["mappings"]=_read(mapping)["mappings"]
    record=_selected_gold(record,gold_path)
    changes=_checked_items(record,snapshot,opinion,digest,_engine(record,mapping.parent))
    if digest!=file_hash(mapping) or opinion_hash!=file_hash(opinion_file):raise ValueError("预检期间输入文件发生变化")
    return {"schema":PREVIEW_SCHEMA,"mapping_path":str(mapping),"mapping_sha256":digest,"opinion_path":str(opinion_file),"opinion_sha256":opinion_hash,
            "carrier":record["carrier"],"source_hash":record["source_hash"],"gold_path":record["gold_path"],"gold_hash":record["gold_hash"],
            "reviewer":copy.deepcopy(opinion["reviewer"]),"count":len(changes),"remaining_pending":len(record["unresolved"])-sum("previous_mapping" not in x for x in changes),"changes":changes}


def validate_external_reviews(record,snapshot,*,_state=None):
    """只读重验外部来源，不回调映射加载器，避免接续校验递归。"""
    external=[m for m in record.get("mappings",[]) if m.get("review_method")==METHOD or "external_review" in m]
    if not external:return False
    if record.get("source_hash")!=snapshot["sha256"]:raise ValueError("外部复核来源与本轮快照不一致")
    state=_state if _state is not None else {"adoptions":{},"checking":set()}
    current=read_standard(flow._check(record["gold_path"],record["gold_hash"],"当前金标准"));validated=set();requires_context=False
    for mapping in external:
        proof=mapping.get("external_review")
        if mapping.get("review_method")!=METHOD or mapping.get("reviewed") is not True or not isinstance(proof,dict):raise ValueError("外部语义缺少实际复核来源")
        receipt=_read(flow._check(proof.get("adoption_path",""),proof.get("adoption_sha256"),"外部复核确认记录"))
        if receipt!={k:v for k,v in proof.items() if k not in {"adoption_path","adoption_sha256","item_index"}}:raise ValueError("外部复核确认方式或来源已变化")
        if receipt.get("schema")!="附注外部复核采纳-v1" or not _text(receipt.get("confirm_method")) or not _text(receipt.get("confirmed_at")):raise ValueError("外部意见未经明确确认")
        cache_key=proof["adoption_sha256"]
        if cache_key in state["checking"]:raise ValueError("外部复核历史来源出现循环引用")
        if cache_key not in state["adoptions"]:
            state["checking"].add(cache_key)
            opinion_path=flow._check(receipt["opinion_path"],receipt["opinion_sha256"],"外部意见副本")
            original_path=flow._check(receipt["mapping_path"],receipt["mapping_sha256"],"复核所依据的原映射副本")
            original=_read(original_path);opinion=_read(opinion_path)
            if original.get("schema")!=flow.MAP_SCHEMA or original.get("carrier") not in {"A","B"} or original.get("source_hash")!=snapshot["sha256"] or (record.get("carrier") is not None and original["carrier"]!=record["carrier"]):raise ValueError("外部复核原映射不是当前来源")
            if receipt.get("source_hash")!=original.get("source_hash") or receipt.get("reviewer")!=opinion.get("reviewer"):raise ValueError("外部意见的审核者或来源不一致")
            selected=_selected_gold(original,flow._check(receipt.get("gold_path",""),receipt.get("gold_hash"),"实际复核金标准"))
            engine=_engine(selected,original_path.parent)
            adds_context=any("business_context" in item for item in opinion.get("items",[]))
            source_snapshot=flow._snapshot(original)
            # 仅沿本次被更正格的旧来源追验；共享采纳哈希缓存，不逐格遍历全历史。
            replaced={(item.get("sheet"),item.get("cell")) for item in opinion.get("items",[]) if "replace_mapping_sha256" in item}
            predecessors=[item for item in original.get("mappings",[]) if (item.get("sheet"),item.get("cell")) in replaced and (item.get("review_method")==METHOD or "external_review" in item)]
            inherited=validate_external_reviews({**original,"mappings":predecessors},source_snapshot,_state=state) if predecessors else False
            changes=_checked_items(selected,source_snapshot,opinion,receipt["mapping_sha256"],engine)
            if receipt.get("accepted_cells")!=[[x["sheet"],x["cell"]] for x in changes]:raise ValueError("确认清单与实际复核意见不一致")
            state["adoptions"][cache_key]=(changes,read_standard(selected["gold_path"]),source_snapshot,original,adds_context or inherited)
            state["checking"].remove(cache_key)
        changes,definitions,source_snapshot,original,bound_context=state["adoptions"][cache_key]
        requires_context=requires_context or bound_context
        if cache_key not in validated:
            if bound_context and (_hash(source_snapshot.get("context",[]))!=_hash(snapshot.get("context",[])) or _confirmed_scope(original)!=_confirmed_scope(record)):
                raise ValueError("补充业务上下文的原结构化事实已变化")
            for change in changes:
                name=change["sheet"];checked=change["mapping"]
                current_sheet=next((s for s in snapshot["sheets"] if s["name"]==name),None)
                source_sheet=next(s for s in source_snapshot["sheets"] if s["name"]==name)
                areas=[area for area in source_sheet.get("confirmed_business_ranges",[]) if area["table_range"]==checked["table_range"] and area["table_semantic"]==checked["table_semantic"] and ("table_basis_hash" not in checked or _hash(area)==checked["table_basis_hash"])]
                if "business_range" in change:areas=[change["business_range"]]
                if len(areas)!=1 or current_sheet is None or areas[0] not in current_sheet.get("confirmed_business_ranges",[]):
                    raise ValueError("外部复核业务范围缺失或与采纳前原意见不一致")
                area=areas[0]
                if "business_range" in change and record.get("confirmed_business_ranges",{}).get(name,[]).count(area)!=1:
                    raise ValueError("外部补充业务范围缺失或与采纳前原意见不一致")
                if bound_context and source_sheet.get("word_table_ranges")!=current_sheet.get("word_table_ranges"):
                    raise ValueError("补充上下文的原Word物理表关联已变化")
                if any(current_sheet["cells"].get(a)!=source_sheet["cells"][a] for a in [change["cell"]]+area["header_cells"]):
                    raise ValueError("外部补充上下文的原格或原文证据已变化")
            validated.add(cache_key)
        index=proof.get("item_index")
        if type(index) is not int or not 0<=index<len(changes):raise ValueError("外部复核项索引无效")
        checked=changes[index]["mapping"]
        if {k:v for k,v in mapping.items() if k not in {"reviewed","review_method","external_review"}}!=checked:raise ValueError("已采纳映射与外部原意见、维度或证据不一致")
        identifier=mapping["slot_id"]
        if identifier not in current or current[identifier].get("status")!="active" or _hash(current[identifier])!=_hash(definitions[identifier]):raise ValueError("外部复核所依据的完整槽位定义已变化，须重新审核")
    return requires_context


def apply_external_review(preview,output_dir,*,confirmed=False,confirm_method=""):
    if isinstance(preview,dict) and preview.get('schema')=='附注外部复核预览-v2':
        from .统一外部复核 import apply_external_review as unified_apply
        return unified_apply(preview,output_dir,confirmed=confirmed,confirm_method=confirm_method)
    if confirmed is not True or not _text(confirm_method):raise ValueError("请先预览并明确确认采纳外部意见")
    if not isinstance(preview,dict) or preview.get("schema")!=PREVIEW_SCHEMA:raise ValueError("缺少外部复核预览")
    fresh=preview_external_review(preview["mapping_path"],preview["opinion_path"],gold_path=preview["gold_path"])
    if fresh!=preview:raise ValueError("意见、原映射或预览已变化，请重新预览")
    record,snapshot=flow._load_mapping(fresh["mapping_path"],complete=False);original=_read(fresh["mapping_path"])
    record=_selected_gold(record,fresh["gold_path"])
    run=Path(output_dir).resolve()/("导入外部复核_"+datetime.now().strftime("%Y%m%d_%H%M%S_%f")+"_"+uuid.uuid4().hex[:6]);run.mkdir(parents=True)
    paths={}
    for key,name in (("mapping","复核所依据的原映射.json"),("opinion","外部复核原始意见.json")):
        raw=flow._check(fresh[key+"_path"],fresh[key+"_sha256"],"复核输入").read_bytes();target=run/name
        with target.open("xb") as stream:stream.write(raw)
        if file_hash(target)!=fresh[key+"_sha256"]:raise ValueError("意见或原映射副本校验失败")
        paths[key+"_path"]=str(target)
    receipt={"schema":"附注外部复核采纳-v1",**paths,"mapping_sha256":fresh["mapping_sha256"],"opinion_sha256":fresh["opinion_sha256"],
             "source_hash":fresh["source_hash"],"gold_path":fresh["gold_path"],"gold_hash":fresh["gold_hash"],"reviewer":fresh["reviewer"],
             "confirm_method":confirm_method,"confirmed_at":datetime.now().astimezone().isoformat(),"accepted_cells":[[x["sheet"],x["cell"]] for x in fresh["changes"]]}
    adoption=flow.save_json(run/"外部复核确认记录.json",receipt)
    accepted=[]
    for index,change in enumerate(fresh["changes"]):
        mapping=copy.deepcopy(change["mapping"]);mapping.update(reviewed=True,review_method=METHOD,
            external_review={**copy.deepcopy(receipt),"adoption_path":adoption,"adoption_sha256":file_hash(adoption),"item_index":index})
        accepted.append(mapping)
    result=copy.deepcopy(record);positions={(m["sheet"],m["cell"]) for m in accepted}
    for change in fresh["changes"]:
        if "business_context" not in change:continue
        area=change["business_range"];name=change["sheet"]
        next(s for s in snapshot["sheets"] if s["name"]==name).setdefault("confirmed_business_ranges",[]).append(copy.deepcopy(area))
        result.setdefault("confirmed_business_ranges",{}).setdefault(name,[]).append(copy.deepcopy(area))
    unchanged=[m for m in original["mappings"] if (m["sheet"],m["cell"]) not in positions]
    result["mappings"]=copy.deepcopy(unchanged)+accepted
    result["unresolved"]=[m for m in result["unresolved"] if (m["sheet"],m["cell"]) not in positions];result["complete"]=not result["unresolved"]
    result.setdefault("external_semantic_reviews",[]).append({"path":adoption,"sha256":file_hash(adoption),"accepted_count":len(accepted)})
    result.pop("evidence_path",None);result.pop("evidence_hash",None)
    _coverage(result,snapshot);validate_external_reviews(result,snapshot)
    item=flow._record(snapshot,result,record["carrier"],record["gold_path"],record["original_path"],record.get("context",[]),record.get("a_link"))
    item["mappings"]=copy.deepcopy(unchanged)+accepted
    prefix=flow.FILE_PREFIXES[record["carrier"]];key=record["carrier"].lower()
    item["review_path"]=flow._mapping_review(item,snapshot,run/(prefix+"语义核对表.xlsx"))
    structure=flow.save_json(run/(prefix+"实际结构.json"),snapshot)
    item["structure_path"]=structure;item.pop("standard_review_path",None)
    files={key+"_review":item["review_path"]}
    if item["unresolved"]:
        item["standard_review_path"]=flow._standard_review(item,snapshot,run,structure);files[key+"_standard_review"]=item["standard_review_path"]
    files[key+"_mapping"]=flow.save_json(run/(prefix+"语义映射.json"),item)
    loaded,saved_snapshot=flow._load_mapping(files[key+"_mapping"],complete=False);validate_external_reviews(loaded,saved_snapshot)
    if _read(files[key+"_mapping"])["mappings"][:len(unchanged)]!=unchanged:raise ValueError("未列入本次更正的原已确认格发生变化，未接受本次成果")
    return {"status":"complete" if item["complete"] else "partial","message":f"已采纳{len(accepted)}格外部复核意见，仍有{len(item['unresolved'])}格待核实；未调用模型或修改金标准。","files":files,"accepted_count":len(accepted)}
