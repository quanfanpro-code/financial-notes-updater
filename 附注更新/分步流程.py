"""以独立成果文件连接提取、A/B识别、勾稽和回写，支持复用与下一轮更新。"""
from __future__ import annotations
import copy
from collections import Counter
import hashlib
import json
import threading
import uuid
from datetime import datetime
from pathlib import Path
from .流程 import ROOT,bridge,save_json,trace_workbook,check_cancel
from .设置 import validate_settings
from .语义 import SemanticEngine,SemanticCancelled,SemanticProgressError,_normal_dimensions,recognition_progress
from .表格 import read_workbook,file_hash,build_update_plan,write_updated_workbook,recalculate_copy,materialize_business_blanks,cell_at
from .结构 import resolve_structure
from .金标准 import ensure_compatible,validate_mapping_scope,normalize_mapping_metric,read_standard

MAP_SCHEMA="附注语义映射-v1"
LINK_SCHEMA="附注Word关联-v1"
PLAN_SCHEMA="附注更新清单-v1"
CHAPTER_TITLE="财务报表重要项目的说明"
STEPS={"extract_a":"提取附注表格","recognize_a":"识别附注表格","recognize_b":"识别更新数据","match":"生成更新清单","write":"生成更新文件"}
CARRIER_NAMES={"A":"附注表格","B":"更新数据表"}
FILE_PREFIXES={"A":"附注表格","B":"更新数据"}

def _read(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))

def _json(value):
    return json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(",",":"),default=str)

def _path(config,key,label):
    value=str(config.get(key) or "").strip()
    if not value or not Path(value).is_file():raise ValueError("请选择存在的"+label)
    return Path(value).resolve()

def _check(path,digest,label):
    path=Path(path)
    if not path.is_file() or file_hash(path)!=digest:
        raise ValueError(label+"已变化或不存在，请重新生成相应成果")
    return path

def _gold(config,record=None):
    path=Path(config.get("gold_path") or (record or {}).get("gold_path") or "")
    if not path.is_file():raise ValueError("请选择正式语义金标准")
    digest=file_hash(path)
    if record:ensure_compatible(record,path)
    return path.resolve(),digest

def _load_link(path,a_path=None):
    item=_read(path)
    if item.get("schema")!=LINK_SCHEMA:raise ValueError("这不是本程序保存的附注表格与Word位置关联文件")
    scope=item.get("scope") or item.get("context",{}).get("scope",{})
    if scope.get("chapter_title")!=CHAPTER_TITLE:
        raise ValueError("位置关联未限定财务报表重要项目的说明章节，请重新提取本章节")
    _check(item["word_path"],item["word_hash"],"关联的Word原件")
    _check(item["linked_word"],item["linked_word_hash"],"带连接的Word")
    _check(item["a_path"],item["a_hash"],"关联的附注表格")
    if a_path and file_hash(a_path)!=item["a_hash"]:raise ValueError("附注表格与Word位置关联文件不属于当前附注表格")
    return item

def _engine(config,gold,identity,log,cancel):
    settings=validate_settings(config)
    identity=dict(identity)
    if identity.get("carrier")=="B":
        settings["layout_mode"]="cells_v1"
        identity["layout_mode"]="cells_v1"
    key=hashlib.sha256(_json(identity).encode("utf-8")).hexdigest()[:24]
    cache=Path(config.get("output_dir") or ROOT/"输出")/"语义缓存"/key
    return SemanticEngine(settings,str(gold),str(cache),log,cancel)

def _gold_snapshot(source,digest):
    """按内容保存一份识别依据，避免外部改写选择文件破坏既有映射。"""
    import shutil
    folder=ROOT/"金标准"/"识别快照";folder.mkdir(parents=True,exist_ok=True)
    target=folder/(digest+".jsonl")
    if target.exists():return _check(target,digest,"识别用金标准快照")
    temporary=folder/(digest+"_"+uuid.uuid4().hex+".tmp")
    shutil.copyfile(source,temporary)
    _check(temporary,digest,"新金标准快照")
    try:temporary.rename(target)
    except FileExistsError:
        # 并发识别相同版本时保留临时副本，核验已经建立的正式快照。
        _check(target,digest,"并发建立的金标准快照")
    return target

def _normalize_mapping(mapping,slot):
    """成果读写共用兼容规则；只在内存中规范已证明等价的表示。"""
    normalized=normalize_mapping_metric(slot,mapping)
    if "dimensions" in normalized:
        normalized["dimensions"]=_normal_dimensions(normalized["dimensions"],slot)
    return normalized


def _record(snapshot,result,carrier,gold,original_path,context,a_link=None):
    item=copy.deepcopy(result)
    definitions=read_standard(gold)
    normalized=[]
    for mapping in item.get("mappings",[]):
        slot=definitions.get(mapping.get("slot_id"))
        if not slot or slot.get("status")!="active":
            raise ValueError("语义映射引用了不在当前active金标准中的ID："+str(mapping.get("slot_id")))
        normalized.append(_normalize_mapping(mapping,slot))
    item["mappings"]=normalized
    item.update(schema=MAP_SCHEMA,carrier=carrier,source_path=snapshot["path"],
        source_hash=snapshot["sha256"],original_path=str(original_path),
        original_hash=file_hash(original_path),gold_path=str(gold),gold_hash=file_hash(gold),
        context=context,recalculated=bool(snapshot.get("recalculated")),
        recalculation=copy.deepcopy(snapshot.get("recalculation",{})))
    if a_link:item["a_link"]=str(a_link)
    item["confirmed_business_ranges"]={s["name"]:copy.deepcopy(s.get("confirmed_business_ranges",[])) for s in snapshot["sheets"]}
    return item


def _mapping_review(record,snapshot,path):
    """给使用者查看的逐格清单；正式后续输入仍为语义映射文件。"""
    from openpyxl import Workbook
    from .表格 import actual_value
    gold={r.get("id"):r for line in Path(record["gold_path"]).read_text(encoding="utf-8-sig").splitlines() if line.strip() for r in [json.loads(line)]}
    book=Workbook();sheet=book.active;sheet.title="业务格语义"
    sheet.append(["工作表","单元格","源内容或公式","可核实的值","金标准ID","指标槽位名称","标准附注","标准表","期间","币种","单位","倍率","实际维度（完整记录）","判断依据"])
    for m in record["mappings"]:
        c=cell_at(snapshot,m["sheet"],m["cell"]);slot=gold.get(m["slot_id"],{});d=m.get("dimensions",{})
        try:value=actual_value(snapshot,m["sheet"],m["cell"])
        except ValueError:value="公式结果尚未核实"
        raw=c.get("value")
        sheet.append([m["sheet"],m["cell"],"空白" if raw is None else raw,"空白" if value is None else value,
            m["slot_id"],m.get("semantic_field"),(slot.get("note") or {}).get("name"),(slot.get("table") or {}).get("name"),
            _json(d.get("period")) if isinstance(d.get("period"),dict) else d.get("period"),
            d.get("currency"),d.get("unit"),d.get("scale"),_json(d),m.get("reason")])
    for title,key in (("待核实","unresolved"),("非业务格","excluded"),("本次范围外","out_of_scope")):
        ws=book.create_sheet(title);ws.append(["工作表","单元格","分类","依据或待核实原因"])
        for item in record.get(key,[]):ws.append([item.get("sheet"),item.get("cell"),item.get("category"),item.get("reason")])
    for ws in book:
        ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
        for row in ws:
            for c in row:
                if isinstance(c.value,str):c.data_type="s"
        for col,width in {"A":23,"B":13,"C":30,"D":24,"E":26,"F":42,"G":24,"H":30,"I":24,"J":14,"K":14,"L":14,"M":70,"N":65}.items():
            ws.column_dimensions[col].width=width
    with Path(path).open("xb") as stream:book.save(stream)
    book.close()
    return str(path)



def _standard_review(record,snapshot,run,structure_path):
    """保留未确认项的来源证据，交给人工或高级AI处理；不自动认定缺标准。"""
    import shutil
    from .表格 import actual_value
    carrier=record["carrier"]
    gold_copy=run/"待处理时使用的金标准.jsonl"
    shutil.copyfile(record["gold_path"],gold_copy)
    _check(gold_copy,record["gold_hash"],"待处理材料中的金标准副本")
    items=[]
    for unresolved in record.get("unresolved",[]):
        item=copy.deepcopy(unresolved)
        try:
            source=cell_at(snapshot,item.get("sheet"),item.get("cell"))
            item["source_cell"]=copy.deepcopy(source)
            try:
                item["verified_value"]=actual_value(snapshot,item["sheet"],item["cell"])
                item["value_verified"]=True
            except ValueError as error:
                item.update(value_verified=False,value_note=str(error))
        except (ValueError,KeyError,TypeError):
            item["source_note"]="涉及范围或结构判断，请按实际结构文件查看，未凭空补造单元格"
        items.append(item)
    instructions=run/"待处理语义说明.md"
    instructions.write_text(
        "# 处理未确认语义\n\n"
        "本目录保存本次识别的语义映射、实际表格结构、核对表和当时使用的金标准副本。"
        "JSON文件中的路径用于本机追溯；将材料交给外部高级AI时，应同时提供这些文件。\n\n"
        "先根据原表的表头、所在附注、口径和实际维度解释每个未确认项。"
        "依次判断：现有active标准是否已覆盖；是否只是开放维度的新取值；是否缺少文件未注明的信息；"
        "是否仍有判断分歧；最后才判断需要新增或修订金标准。不可按新客户、新期间、新排版重复建标准。\n\n"
        "如果现有标准已经覆盖，请给出具体ID、维度和单元格依据，补充必要事实后重新识别。"
        "若确需改标准，请用本目录金标准副本作为起点，保留全部旧ID和字段，"
        "将完整修订结果另存为新的JSONL文件。新增定义按原格式提供语义、口径、维度、来源依据和引用关系；"
        "修改或停用旧定义须说明原因。不得用金额、坐标或某个样式替代语义定义。\n\n"
        "人工或高级AI完成语义判断后，仅在确需修订标准时，才在程序设置的金标准管理中选择修订文件，"
        "检查新增、修订、停用清单，再保存并使用新版本。程序检查结构和引用，不能代替语义判断。"
        "新版本保存后重新识别未完成的附注表格或更新数据表；已有映射所用定义未变时可以继续复用。\n",
        encoding="utf-8-sig")
    package={
        "schema":"附注待处理语义-v1","carrier":carrier,"source_path":record["source_path"],
        "source_hash":record["source_hash"],"original_path":record["original_path"],"original_hash":record["original_hash"],
        "gold_path":str(gold_copy),"gold_hash":record["gold_hash"],"original_gold_path":record["gold_path"],
        "mapping_path":str(run/(FILE_PREFIXES[carrier]+"语义映射.json")),"structure_path":str(structure_path),
        "review_path":record["review_path"],"instructions_path":str(instructions),
        "additional_context":record.get("additional_context",""),"context":record["context"],
        "items":items,"proposed_changes":[],"notice":"未确认不等于金标准缺项，须先检查现有定义和开放维度"}
    return save_json(run/(FILE_PREFIXES[carrier]+"待处理语义与金标准补充材料.json"),package)

def _snapshot(record):
    source=_check(record["source_path"],record["source_hash"],"语义映射对应的工作簿")
    if record.get("original_path"):
        _check(record["original_path"],record["original_hash"],"识别时的原始资料")
    snapshot=read_workbook(source,context=record.get("context",[]))
    basis=record.get("recognition_basis",{})
    if basis.get("candidate_blanks"):
        if basis.get("layout_mode")!="cells_v1" or basis["candidate_blanks"]!="source_grid_v1":
            raise ValueError("不支持这份映射的空白候选识别方式")
        from .表格 import materialize_candidate_blanks
        materialize_candidate_blanks(snapshot)
    for sheet,areas in record.get("confirmed_business_ranges",{}).items():
        for area in areas:
            if area.get("source_hash")!=record["source_hash"]:
                raise ValueError("业务空白的来源证据与工作簿不一致")
            table={**copy.deepcopy(area),"range":area["table_range"]}
            if "layout_batch_id" in area or "layout_review" in area:
                actual_sheet=next((s for s in snapshot["sheets"] if s["name"]==sheet),None)
                if actual_sheet is None:raise ValueError("逐格布局引用未知工作表")
                SemanticEngine._restore_cell_table(snapshot,actual_sheet,table)
            else:
                materialize_business_blanks(snapshot,sheet,table)
    if record.get("recalculated"):
        proof=record.get("recalculation",{})
        if proof.get("method")!="Excel.CalculateFullRebuild" or proof.get("source_sha256")!=record.get("original_hash"):
            raise ValueError("公式重算缺少可核验的原文件身份")
        snapshot.update(recalculated=True,recalculation=proof)
    return snapshot

def _load_mapping(path,carrier=None,complete=True):
    record=_read(path)
    if record.get("schema")!=MAP_SCHEMA:raise ValueError("请选择本程序保存的语义映射文件")
    if carrier and record.get("carrier")!=carrier:raise ValueError("选择的语义映射不是"+CARRIER_NAMES[carrier]+"的映射")
    if complete and (not record.get("complete") or record.get("unresolved")):
        raise ValueError("语义映射还有未确认项，请先完成该表识别")
    if not isinstance(record.get("mappings"),list) or (complete and not record["mappings"]):
        raise ValueError("语义映射没有已确认的业务格")
    gold=Path(record.get("gold_path") or "")
    _check(gold,record.get("gold_hash"),"映射所用金标准")
    active={row.get("id"):row for line in gold.read_text(encoding="utf-8-sig").splitlines() if line.strip()
        for row in [json.loads(line)] if row.get("status")=="active" and row.get("id")}
    snapshot=_snapshot(record)
    selection=record.get("scope_selection")
    if not selection and (record.get("out_of_scope") or record.get("recognition_basis",{}).get("selection_hash") not in (None,SemanticEngine._selection_identity(None))):
        raise ValueError("语义映射缺少实际识别范围来源，不能只保留章外结论")
    if selection:
        if selection.get("source_hash")!=record["source_hash"] or selection.get("target_hash")!=hashlib.sha256(_json(selection.get("target_meanings")).encode("utf-8")).hexdigest():
            raise ValueError("语义映射的识别范围来源或目标含义校验失败")
        from .识别范围 import read_verified_scope
        verified_scope=read_verified_scope(selection,snapshot,selection.get("target_meanings"))
        if verified_scope is not None:
            basis=record.get("recognition_basis",{})
            if basis.get("layout_mode")!="cells_v1" or basis.get("candidate_blanks")!="source_grid_v1":
                raise ValueError("外部范围确认与映射的完整候选格依据不一致")
        expected={(s["name"],a) for s in snapshot["sheets"] for a,c in s["cells"].items() if verified_scope is not None or not c.get("derived_blank")}
        selected=[(sheet,a) for sheet,addresses in selection.get("selected_cells",{}).items() for a in addresses]
        omitted=[(area["sheet"],a) for area in selection.get("out_of_scope",[]) for a in area.get("cells",[])]
        if set(selected+omitted)!=expected or len(set(selected+omitted))!=len(selected+omitted):
            raise ValueError("语义映射的识别范围存在遗漏或重复")
        if any((m["sheet"],m["cell"]) in set(omitted) for m in record["mappings"]):
            raise ValueError("范围外数据被错误地纳入业务语义映射")
        if verified_scope is not None:
            confirmed_omitted=[dict(sheet=area["sheet"],cell=a,category="out_of_scope",reason=area["reason"],evidence_cells=area["evidence_cells"]) for area in selection["out_of_scope"] for a in area["cells"]]
            if record.get("out_of_scope",[])!=confirmed_omitted:
                raise ValueError("映射的章外清单与实际外部确认不一致")
            positions=[(item.get("sheet"),item.get("cell")) for category in ("mappings","excluded","unresolved","out_of_scope") for item in record.get(category,[])]
            if set(positions)!=expected or len(positions)!=len(set(positions)):
                raise ValueError("外部范围映射存在遗漏、重复或包外格")
    seen=set()
    for item in record["mappings"]:
        pos=(item.get("sheet"),item.get("cell"))
        if pos in seen:raise ValueError("语义映射包含重复业务格")
        seen.add(pos);cell_at(snapshot,*pos)
        slot=item.get("slot_id")
        if slot not in active:raise ValueError("语义映射引用了不在当前active金标准中的ID："+str(slot))
        validate_mapping_scope(active[slot],item)
    if any(m.get("review_method")=="external_semantic_review" or "external_review" in m for m in record["mappings"]):
        from .外部复核 import validate_external_reviews
        validate_external_reviews(record,snapshot)
    # 先完成原件、金标准与原记录校验，再规范内存副本，不改写旧成果文件。
    record["mappings"]=[_normalize_mapping(item,active[item["slot_id"]]) for item in record["mappings"]]
    return record,snapshot


def _selected_sources(config,ar,br):
    """界面当前选择与落盘成果必须指向相同资料，空路径允许纯成果文件复用。"""
    for key,record,label in (("a_path",ar,"A"),("b_path",br,"B")):
        if str(config.get(key) or "").strip():
            source=_path(config,key,CARRIER_NAMES[label])
            if file_hash(source) not in {record.get("source_hash"),record.get("original_hash")}:
                raise ValueError("当前选择的"+CARRIER_NAMES[label]+"与已保存语义映射不是同一份资料，请重新识别或选择对应映射")


def _input_records(paths):
    return {str(Path(path).resolve()):file_hash(path) for path in paths if path}

def _verify_inputs(inputs):
    for path,digest in inputs.items():_check(path,digest,"该步骤依赖的文件")

def _extract(config,run,log,cancel):
    word=_path(config,"word_path","原始附注 Word")
    if word.suffix.lower()!=".docx":raise ValueError("原始附注 Word需要为.docx")
    before=file_hash(word);a=run/"附注表格.xlsx";linked=run/"原始附注_带连接.docx";context=run/"Word表格上下文.json"
    log("只提取财务报表重要项目的说明章节，并单独保存附注表格与Word的位置关联")
    result=bridge({"action":"extract","word":str(word),"excel":str(a),"linked_word":str(linked),"context":str(context),"chapter_title":CHAPTER_TITLE},run,cancel)
    _check(word,before,"原始附注 Word")
    link={"schema":LINK_SCHEMA,"word_path":str(word),"word_hash":before,"linked_word":str(linked),
        "linked_word_hash":file_hash(linked),"a_path":str(a),"a_hash":file_hash(a),"context":_read(context)}
    link["scope"]=link["context"].get("scope",{})
    if link["scope"].get("chapter_title")!=CHAPTER_TITLE:raise ValueError("Word提取未返回已确认的章节范围")
    link_path=save_json(run/"附注表格与Word位置关联.json",link)
    return {"status":"complete","message":f"已提取指定章节的{result['table_count']}张表。下一步识别附注表格语义。",
        "files":{"a_path":str(a),"a_link":link_path},"table_count":result["table_count"]}


def _apply_reviewed_existing(result,learned,engine,snapshot):
    """独立复核已确认现有槽位时接入原成果；不改旧结果、原数据或金标准。"""
    proposed=learned.get("reviewed_existing_mappings") or []
    if not proposed:return result
    report=_check(learned.get("report_path",""),learned.get("report_sha256"),"已有语义复核报告")
    saved=_read(report)
    if _json(saved)!=_json({k:v for k,v in learned.items() if k not in {"report_path","report_sha256"}}):
        raise ValueError("已有语义复核意见与已保存报告不一致")
    if learned.get("source_hash")!=snapshot["sha256"] or learned.get("gold_hash")!=engine.gold_hash:
        raise ValueError("已有语义复核的原表或金标准版本不一致")
    pending={(m.get("sheet"),m.get("cell")) for m in result.get("unresolved",[])}
    accepted=[];seen=set();sheets={s["name"]:s for s in snapshot["sheets"]}
    for raw in proposed:
        pos=(raw.get("sheet"),raw.get("cell"))
        if pos not in pending or pos in seen:raise ValueError("已有语义复核引用非待确认格或出现重复格")
        seen.add(pos)
        review=raw.get("semantic_review") or {}
        second=review.get("independent_review") or {}
        if (raw.get("reviewed") is not True or raw.get("review_method")!="existing_slot_proposal_and_independent_review"
                or not review.get("proposal") or second.get("accepted") is not True or not str(second.get("reason") or "").strip()):
            raise ValueError("已有语义缺少独立复核通过的真实依据")
        sheet=sheets.get(pos[0])
        if sheet is None:raise ValueError("已有语义复核引用未知工作表")
        areas={_json(a):a for a in sheet.get("confirmed_business_ranges",[])+result.get("confirmed_business_ranges",{}).get(pos[0],[])
               if a.get("source_hash")==snapshot["sha256"] and a.get("table_range")==raw.get("table_range")}
        if len(areas)!=1:raise ValueError("已有语义复核缺少唯一原业务范围")
        area=next(iter(areas.values()));table={**area,"range":area["table_range"]}
        engine._validate_layout({"tables":[table]},sheet,[])
        from .语义 import _inside,_range
        if not _inside(pos[1],_range(table["range"])):raise ValueError("已有语义复核引用原业务范围外单元格")
        for opinion in (review["proposal"],second):
            evidence=opinion.get("evidence_cells")
            if not isinstance(evidence,list) or not evidence or any(a not in sheet["cells"] for a in evidence):
                raise ValueError("已有语义复核缺少可回读的原格证据")
        allowed={s["id"] for s in engine._candidate_slots(table)}
        checked=engine._validate_classification({"mappings":[raw],"excluded":[],"unresolved":[]},[pos[1]],allowed,sheet,table)[pos[1]][1]
        checked.update(reviewed=True,review_method=raw["review_method"],semantic_review=copy.deepcopy(review))
        accepted.append(checked)
    updated=copy.deepcopy(result)
    updated["mappings"].extend(accepted)
    updated["unresolved"]=[m for m in updated["unresolved"] if (m["sheet"],m["cell"]) not in seen]
    expected={(s["name"],a) for s in snapshot["sheets"] for a in s["cells"]}
    positions=[(i["sheet"],i["cell"]) for k in ("mappings","excluded","unresolved","out_of_scope") for i in updated.get(k,[])]
    if set(positions)!=expected or len(positions)!=len(set(positions)):
        raise ValueError("已有语义复核后存在遗漏或重复，未采用")
    updated["complete"]=not updated["unresolved"]
    updated.setdefault("existing_semantic_reviews",[]).append({"path":str(report),"sha256":learned["report_sha256"],"accepted_cells":len(accepted)})
    # 新证据另存；后续新增金标准接续时仍可核实完整原结果，不改写旧模型回执。
    updated.pop("evidence_path",None);updated.pop("evidence_hash",None)
    proof=engine.work_dir/("识别结果_已有语义复核_"+uuid.uuid4().hex+".json")
    engine._save(proof,updated)
    updated.update(evidence_path=str(proof.resolve()),evidence_hash=file_hash(proof))
    return updated


def _check_b_mapping_mode(config, path):
    """范围映射不能冒充完整来源识别成果；原文件保持不变。"""
    if not config.get("scope_prefilter", False) and path:
        record = _read(path)
        if record.get("scope_selection") or record.get("scope_selection_path"):
            raise ValueError("所选更新数据映射曾按附注筛选，不能作为完整更新数据映射。请在第 3 步清空旧映射选择后重新识别；需要沿用原筛选方式时，先勾选“先按附注内容筛选来源（可选）”。")


def _recognize_v2(config,carrier,run,log,cancel,progress=None):
    from .统一识别 import UnifiedSemanticEngine
    from openpyxl import Workbook
    key='a_path' if carrier=='A' else 'b_path'
    original=_path(config,key,CARRIER_NAMES[carrier]);gold,digest=_gold(config)
    settings=validate_settings(config)
    context=[];link_path=None
    if carrier=='A' and config.get('a_link'):
        link_path=_path(config,'a_link','附注表格与Word位置关联文件')
        link=_load_link(link_path,original);context.extend(link['context'].get('tables',[]))
    additional=str(config.get('a_context_text' if carrier=='A' else 'b_context_text','') or '').strip()
    if additional:context.append(additional)
    identity=hashlib.sha256(_json({'source':file_hash(original),'gold':digest,'context':context,'carrier':carrier}).encode('utf-8')).hexdigest()
    cache=Path(config.get('output_dir') or ROOT/'输出')/'统一语义缓存'/identity
    engine=UnifiedSemanticEngine(settings,gold,cache,log,cancel)
    mapping_key='a_mapping' if carrier=='A' else 'b_mapping'
    latest_gold=None;learning=[];usage=Counter()
    metadata={'carrier':carrier,'original_path':str(original),'original_hash':file_hash(original),'additional_context':additional}
    if link_path:metadata['a_link']=str(link_path)
    def publish(record):
        counts={'mappings':len(record['facts']),'excluded':len(record['excluded']),'unresolved':len(record['unresolved'])}
        if progress:
            progress({'step':'recognize_a' if carrier=='A' else 'recognize_b','status':'running','checkpoint':True,
                      'files':{mapping_key:record['mapping_path'],**({'gold_path':latest_gold} if latest_gold else {})},'counts':counts,'output_dir':str(run),
                      'message':'已保存整份表格识别进度，未决格和未处理格均在映射中保留'})
    log('按统一指标及实际维度识别整份'+CARRIER_NAMES[carrier]+'，独立保存每格分类及证据')
    previous=config.get(mapping_key)
    result=engine.recognize_workbook(original,context=context,progress=publish,artifact_metadata=metadata,
                                     **({'previous_result':previous} if previous else {}))
    while result['unresolved']:
        from .统一补充 import review_missing_standard
        check_cancel(cancel)
        pending_count=len(result['unresolved'])
        log('复核未决项是已有定义的新取值，还是金标准确实缺项；不生成或修改金额')
        try:
            learned=review_missing_standard(engine,result,run)
        except SemanticCancelled:
            raise
        except Exception as error:
            message=engine._redact(str(error));learning.append({'error':message,'new_ids':[],'changed_ids':[]})
            log('金标准补充尚未完成，保留已识别映射和待审原文：'+message)
            break
        learning.append(learned)
        if not learned['new_ids'] and not learned['changed_ids']:break
        # 先登记已发布版本，接续失败或取消时界面仍能保留该版本及旧映射。
        latest_gold=learned['path'];publish(result)
        usage.update(dict(engine.usage));gold=Path(latest_gold);digest=file_hash(gold)
        next_identity=hashlib.sha256(_json({'source':file_hash(original),'gold':digest,'context':context,'carrier':carrier}).encode('utf-8')).hexdigest()
        cache=Path(config.get('output_dir') or ROOT/'输出')/'统一语义缓存'/next_identity
        engine=UnifiedSemanticEngine(settings,gold,cache,log,cancel)
        log('新金标准已另存，核实已有映射兼容性后继续识别待处理单元格')
        result=engine.recognize_workbook(original,context=context,progress=publish,previous_result=result['mapping_path'],artifact_metadata=metadata)
        if len(result['unresolved'])>=pending_count:
            log('新版本未减少待审格，保存本次结果，避免重复生成定义或循环调用模型')
            break
    usage.update(dict(engine.usage));result['usage']=dict(usage)
    learning_path=None
    if learning:
        learning_path=save_json(run/'本次统一金标准补充记录.json',{'reviews':learning,'latest_published_gold_path':latest_gold})
        result['gold_learning_path']=str(learning_path)
    from .统一勾稽 import calculate_mapping_formulas
    formula_facts=[fact for fact in result['facts'] if fact['record_type']=='metric_fact' and fact['raw_value_state']=='formula']
    if formula_facts:
        check_cancel(cancel);log('在独立副本中重算已识别的业务公式，原公式与计算结果分别保存')
        try:calculate_mapping_formulas(result,run/'公式计算依据')
        except Exception as error:
            result['calculation_issues']=[{'location':fact['source_reference']['location'],'reason':str(error)} for fact in formula_facts]
        else:
            result['calculation_issues']=[{'location':location,'reason':value['reason']}
                                          for location,value in result['formula_values'].items() if value['state']=='error']
        check_cancel(cancel)
    result.update(carrier=carrier,original_path=str(original),original_hash=file_hash(original))
    if link_path:result['a_link']=str(link_path)
    mapping_path=run/(FILE_PREFIXES[carrier]+'语义映射.json')
    review=run/(FILE_PREFIXES[carrier]+'语义识别核对.xlsx')
    pending=run/(FILE_PREFIXES[carrier]+'待处理语义与金标准补充材料.json') if result['unresolved'] else None
    result.update(review_path=str(review),additional_context=additional)
    if pending:result['standard_review_path']=str(pending)
    result['mapping_path']=str(mapping_path);save_json(mapping_path,result)
    if pending:
        from .统一识别 import standard_review_material
        save_json(pending,standard_review_material(result,engine.standard))
    book=Workbook();sheet=book.active;sheet.title='逐格语义识别'
    sheet.append(['来源位置','识别结果','指标槽位或类别','实际维度','原值或原公式','识别依据','核实计算结果','计算情况'])
    for fact in result['facts']:
        calculated=result.get('formula_values',{}).get(fact['source_reference']['location'],{})
        sheet.append([fact['source_reference']['location'],'已识别',fact.get('metric_id') or fact.get('dimension_id'),
                      _json(fact.get('dimensions',{})),fact.get('raw_value'),_json(fact.get('dimension_evidence',{})),
                      calculated.get('value'),calculated.get('reason')])
    original_cells={}
    if result['unresolved']:
        original_snapshot=read_workbook(original)
        if original_snapshot['sha256']!=result['source_hash']:raise ValueError('导出核对表时原件已变化')
        original_cells={part['name']:part['cells'] for part in original_snapshot['sheets']}
    for kind,label in (('excluded','标题或说明等'),('unresolved','待核实')):
        for row in result[kind]:
            known=row.get('partial_semantics',{})
            sheet.append([row['sheet']+'!'+row['cell'],'已识别部分语义，待核实' if known else label,
                          known.get('metric_id',row.get('category')),_json(known['dimensions']) if known else None,
                          original_cells.get(row['sheet'],{}).get(row['cell'],{}).get('value'),row['reason']])
    for row in sheet:
        for cell in row:
            if isinstance(cell.value,str):cell.data_type='s'
    sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
    for column,width in {'A':30,'B':18,'C':40,'D':70,'E':30,'F':80,'G':20,'H':45}.items():sheet.column_dimensions[column].width=width
    from .同义复核 import append_review_sheet,mapping_review_proofs
    append_review_sheet(book,mapping_review_proofs(result))
    from .统一证据 import append_context_sheet
    append_context_sheet(book,result.get('source_context',[]),original)
    with review.open('xb') as handle:book.save(handle)
    book.close()
    return {'status':'complete' if result['complete'] and not result.get('calculation_issues') else 'partial',
            'message':f"{CARRIER_NAMES[carrier]}独立映射已保存：业务{len(result['facts'])}格，标题说明等{len(result['excluded'])}格，语义待核实{len(result['unresolved'])}格，公式计算待核实{len(result.get('calculation_issues',[]))}格。尚未更新金额或Word。",
            'files':{mapping_key:str(mapping_path),'a_review' if carrier=='A' else 'b_review':str(review),
                     **({'gold_path':latest_gold} if latest_gold else {}),
                     **({'a_gold_changes' if carrier=='A' else 'b_gold_changes':str(learning_path)} if learning_path else {}),
                     **({'a_standard_review' if carrier=='A' else 'b_standard_review':str(pending)} if pending else {})},'usage':result['usage']}


def _recognize(config,carrier,run,log,cancel,progress=None):
    selected_gold,_=_gold(config)
    with selected_gold.open(encoding='utf-8-sig') as handle:
        first=next((json.loads(line) for line in handle if line.strip()),{})
    if first.get('record_type'):
        return _recognize_v2(config,carrier,run,log,cancel,progress)
    if carrier=="B":_check_b_mapping_mode(config,config.get("b_mapping"))
    key="a_path" if carrier=="A" else "b_path"
    original=_path(config,key,CARRIER_NAMES[carrier])
    if original.suffix.lower() not in {".xlsx",".xlsm"}:raise ValueError("识别对象需要为Excel工作簿")
    gold,digest=_gold(config);before=file_hash(original)
    context=[];link_path=None
    if carrier=="A" and config.get("a_link"):
        link_path=_path(config,"a_link","附注表格与Word位置关联文件")
        link=_load_link(link_path,original);context.extend(link["context"].get("tables",[]))
    context_key="a_context_text" if carrier=="A" else "b_context_text"
    additional_context=str(config.get(context_key,config.get("context_text","")) or "").strip()
    if additional_context:context.append(additional_context)
    if config.get("report_scope"):
        context.append({"report_scope":config["report_scope"],"basis":additional_context,"source":"本次已确认的报告口径"})
    snapshot=read_workbook(original,context=context)
    if carrier=="B" and any(c.get("formula") for s in snapshot["sheets"] for c in s["cells"].values()):
        log("对更新数据表的独立副本重算公式，保留原文件")
        recalculated=run/("更新数据_公式重算"+original.suffix)
        recalculate_copy(original,recalculated);snapshot=read_workbook(recalculated,context=context)
    snapshot["original_hash"]=before;snapshot["original_path"]=str(original)
    reference=None
    if carrier=="B" and config.get("a_mapping"):
        try:
            a_record,_=_load_mapping(config["a_mapping"],"A")
            ensure_compatible(a_record,gold)
            reference=a_record["mappings"]
        except (ValueError,OSError,KeyError,TypeError) as error:
            log("现有附注表格映射暂不能作参考，更新数据表仍按当前金标准独立识别："+str(error))
    engine=_engine(config,gold,{"original":before,"gold":digest,"carrier":carrier,
        "context":context,"report_scope":config.get("report_scope"),"reference":reference},log,cancel)
    log("逐格识别"+CARRIER_NAMES[carrier]+"并保存独立语义映射；本步骤不更新任何金额")
    target_scope=None
    if carrier=="B" and config.get("scope_prefilter",False) and config.get("a_link"):
        from .识别范围 import build_target_meanings
        target_link=_load_link(config["a_link"],config.get("a_path") or None)
        if config.get("a_mapping"):
            target_record,_=_load_mapping(config["a_mapping"],"A",complete=False)
            if target_link["a_hash"] not in {target_record.get("source_hash"),target_record.get("original_hash")}:
                raise ValueError("识别范围引用的位置关联与当前附注表格映射不一致")
        target_snapshot=read_workbook(target_link["a_path"],context=target_link["context"].get("tables",[]))
        target_scope=build_target_meanings(target_link,target_snapshot)
    scope_arguments={"target_scope":target_scope} if target_scope else {}
    if target_scope and config.get("b_scope"):
        # 范围筛选仅使用来源文字和目标披露说明；后续增加金标准槽位不改变这两项证据。
        scope_arguments["scope_selection"]=_read(config["b_scope"])
    previous_path=config.get("a_mapping" if carrier=="A" else "b_mapping")
    if previous_path:
        log("核验已选择的本表语义映射，保留仍然适用的已确认成果")
        scope_arguments["previous_result"]={"mapping_path":str(Path(previous_path).resolve()),"mapping_sha256":file_hash(previous_path)}
    last_progress_identity=None
    last_progress_event={}
    latest_published_gold=None;latest_published_digest=None
    evidence_dir=run/"识别进度"/"识别依据"
    def save_progress(current_snapshot,current_result,current_gold):
        nonlocal last_progress_identity,last_progress_event
        partial=recognition_progress(current_snapshot,current_result)
        current_digest=file_hash(current_gold)
        if partial.get("gold_hash")!=current_digest:
            raise ValueError("识别进度与本轮金标准版本不一致")
        _check(original,before,"识别原件")
        identity=hashlib.sha256(_json({"result":partial,"gold":current_digest,"published_gold":latest_published_gold,"published_gold_hash":latest_published_digest}).encode("utf-8")).hexdigest()
        if identity==last_progress_identity:return
        evidence_dir.mkdir(parents=True,exist_ok=True)
        frozen_gold=evidence_dir/("金标准_"+current_digest+".jsonl")
        if not frozen_gold.exists():
            with Path(current_gold).open("rb") as source,frozen_gold.open("xb") as target:
                import shutil
                shutil.copyfileobj(source,target)
        _check(frozen_gold,current_digest,"进度所用金标准副本")
        structure_path=evidence_dir/"原始实际结构.json"
        if not structure_path.exists():save_json(structure_path,current_snapshot)
        item=_record(current_snapshot,partial,carrier,frozen_gold,original,context,link_path)
        item.update(additional_context=additional_context,selected_gold_path=str(current_gold),
                    structure_path=str(structure_path),checkpoint=True)
        files={}
        if latest_published_gold:
            _check(latest_published_gold,latest_published_digest,"本次已发布的新金标准")
            item.update(latest_published_gold_path=latest_published_gold,latest_published_gold_hash=latest_published_digest)
            files["gold_path"]=latest_published_gold
        if item.get("scope_selection"):
            scope_digest=hashlib.sha256(_json(item["scope_selection"]).encode("utf-8")).hexdigest()
            scope_path=evidence_dir/("更新数据识别范围_"+scope_digest+".json")
            if not scope_path.exists():save_json(scope_path,item["scope_selection"])
            if _json(_read(scope_path))!=_json(item["scope_selection"]):raise ValueError("进度识别范围回读不一致")
            item["scope_selection_path"]=str(scope_path);files["b_scope"]=str(scope_path)
        folder=run/"识别进度"/(datetime.now().strftime("%Y%m%d_%H%M%S_%f")+"_"+uuid.uuid4().hex[:6])
        path=folder/(FILE_PREFIXES[carrier]+"语义映射.json")
        temporary=path.with_suffix(".tmp")
        save_json(temporary,item)
        if _json(_read(temporary))!=_json(item):raise ValueError("识别进度回读校验失败")
        expected_hash=file_hash(temporary)
        temporary.rename(path)
        _check(path,expected_hash,"进度映射")
        _check(original,before,"识别原件")
        _check(current_snapshot["path"],current_snapshot["sha256"],"识别工作簿")
        files["a_mapping" if carrier=="A" else "b_mapping"]=str(path)
        counts={key:len(item.get(key,[])) for key in ("mappings","excluded","unresolved")}
        last_progress_identity=identity
        event={"status":"running","step":"recognize_a" if carrier=="A" else "recognize_b", "output_dir":str(run),
               "message":f"已保存识别进度：业务格{counts['mappings']}格，非业务格{counts['excluded']}格，待处理{counts['unresolved']}格。",
               "files":files,"file_hashes":{str(path):expected_hash},"counts":counts,"checkpoint":True}
        last_progress_event=copy.deepcopy(event)
        log(event["message"])
        if progress is not None:progress(event)
    def publish_progress(current_snapshot,current_result,current_gold):
        try:save_progress(current_snapshot,current_result,current_gold)
        except SemanticProgressError:raise
        except Exception as error:raise SemanticProgressError("保存识别进度失败："+str(error)) from error
    result=engine.recognize(snapshot,reference=reference,progress=lambda snap,result:publish_progress(snap,result,gold),**scope_arguments)
    publish_progress(snapshot,result,gold)
    learning_reports=[]; learned_ids=[]; seen_learning=set(); usage=Counter()
    while result.get("unresolved"):
        signature=(digest,tuple(sorted((str(i.get("sheet")),str(i.get("cell"))) for i in result["unresolved"])))
        if signature in seen_learning:break
        seen_learning.add(signature)
        from .语义补充 import review_missing_semantics
        log("核查未确认项是否为真正缺少的新语义；与已有定义及开放维度区别后保存新版本")
        published_this_review=False
        try:
            learned=review_missing_semantics(engine,snapshot,result,str(gold),str(run))
            if learned.get("new_ids"):
                # 发布是已经发生的状态变化，必须先登记；后续应用或保存失败也要保留接续版本。
                next_gold=Path(learned["path"]);next_digest=file_hash(next_gold)
                definitions=read_standard(next_gold)
                if any(definitions.get(identifier,{}).get("status")!="active" for identifier in learned["new_ids"]):
                    raise ValueError("补充结果的新ID不在已发布版本的有效定义中")
                _check(next_gold,next_digest,"刚发布的新金标准")
                latest_published_gold=str(next_gold);latest_published_digest=next_digest
                published_this_review=True
                event=copy.deepcopy(last_progress_event)
                for path,expected_hash in event.get("file_hashes",{}).items():_check(path,expected_hash,"已保存进度")
                event["files"]["gold_path"]=latest_published_gold
                event["file_hashes"][latest_published_gold]=next_digest
                event["message"]="已登记本次发布的新金标准；已有映射保留实际识别版本，接续时核验兼容性。"
                last_progress_event=copy.deepcopy(event)
                try:
                    if progress is not None:progress(event)
                except Exception as error:raise SemanticProgressError("通知已发布金标准失败："+str(error)) from error
                log(event["message"])
            previous_pending=len(result["unresolved"])
            result=_apply_reviewed_existing(result,learned,engine,snapshot)
            publish_progress(snapshot,result,gold)
            if len(result["unresolved"])<previous_pending:
                log(f"独立复核确认已有槽位适用，解决{previous_pending-len(result['unresolved'])}个待核实格；未新增金标准")
        except (SemanticCancelled,SemanticProgressError):
            raise
        except Exception as error:
            if published_this_review:raise
            message=f"{type(error).__name__}：{error}"
            secret=str(config.get("api_key") or "")
            if secret:message=message.replace(secret,"[密钥已隐藏]")
            learning_reports.append({"status":"failed","reason":message,"new_ids":[],"gold_path":str(gold)})
            log("金标准补充未完成，仍保存本次已识别语义和待核实事项："+message)
            break
        learning_reports.append(learned)
        if not learned.get("new_ids"):break
        _check(original,before,"识别原件");_check(gold,digest,"补充前金标准")
        learned_ids.extend(learned["new_ids"]);usage.update(dict(engine.usage))
        log(f"本地金标准新增{len(learned['new_ids'])}个经核实的语义槽位，继续识别本次资料")
        previous_selection=result.get("scope_selection")
        engine=_engine(config,next_gold,{"original":before,"gold":next_digest,"carrier":carrier,"context":context,"report_scope":config.get("report_scope"),"reference":reference},log,cancel)
        arguments={"target_scope":target_scope,"scope_selection":previous_selection} if target_scope else {}
        next_snapshot=copy.deepcopy(snapshot)
        # 后续引擎失败时直接返回最近已验进度，不能再用旧result生成较早版本的成果。
        next_result=engine.recognize(next_snapshot,reference=reference,previous_result=result,progress=lambda snap,result:publish_progress(snap,result,next_gold),**arguments)
        snapshot=next_snapshot;result=next_result;gold=next_gold;digest=next_digest
        publish_progress(snapshot,result,gold)
    usage.update(dict(engine.usage));result["usage"]=dict(usage)
    _check(original,before,"识别原件");_check(gold,digest,"金标准")
    frozen_gold=_gold_snapshot(gold,digest)
    item=_record(snapshot,result,carrier,frozen_gold,original,context,link_path)
    item["selected_gold_path"]=str(gold)
    learning_path=None
    if learning_reports:
        learning_path=save_json(run/"本次金标准补充记录.json",{"new_ids":learned_ids,"gold_path":str(gold),"gold_hash":digest,"latest_published_gold_path":latest_published_gold,"reviews":learning_reports})
        item["gold_learning_path"]=learning_path
        item["new_gold_ids"]=learned_ids
    item["additional_context"]=additional_context
    review=_mapping_review(item,snapshot,run/(FILE_PREFIXES[carrier]+"语义核对表.xlsx"))
    item["review_path"]=review
    structure_path=save_json(run/(FILE_PREFIXES[carrier]+"实际结构.json"),snapshot)
    pending=None
    if item.get("unresolved") or not item.get("complete"):
        pending=_standard_review(item,snapshot,run,structure_path)
        item["standard_review_path"]=pending
    scope_path=None
    if item.get("scope_selection"):
        scope_path=save_json(run/"更新数据识别范围.json",item["scope_selection"])
        item["scope_selection_path"]=scope_path
    mapping_path=save_json(run/(FILE_PREFIXES[carrier]+"语义映射.json"),item)
    count=len(item["mappings"]);unresolved=len(item.get("unresolved",[]))
    return {"status":"complete" if item.get("complete") else "partial",
        "message":f"{FILE_PREFIXES[carrier]}语义映射已独立保存：确认{count}格，待核实{unresolved}项。"+(f"本次金标准新增{len(learned_ids)}个语义槽位。" if learned_ids else ""),
        "files":{"a_mapping" if carrier=="A" else "b_mapping":mapping_path,
                 "a_review" if carrier=="A" else "b_review":review,
                 **({"a_standard_review" if carrier=="A" else "b_standard_review":pending} if pending else {}),
                 **({"b_scope":scope_path} if scope_path else {}),
                 **({"a_gold_changes" if carrier=="A" else "b_gold_changes":learning_path} if learning_path else {}),
                 **({"gold_path":latest_published_gold} if latest_published_gold else {})},
        "usage":dict(item.get("usage",{}))}

def _match_v2(config,run,log,cancel,ap,bp):
    """新版独立映射先接入同一个勾稽入口；范围及结构未验收时明确保留待处理。"""
    from .统一语义 import read_standard_v2
    from .统一勾稽 import load_fact_mapping,build_fact_update_plan,calculation_files,select_mapping_scope,target_for_purpose
    gold,digest=_gold(config)
    standard=read_standard_v2(gold)
    ar=load_fact_mapping(ap,standard); br=load_fact_mapping(bp,standard)
    for key,record in (('a_path',ar),('b_path',br)):
        if config.get(key):_check(config[key],record['source_hash'],CARRIER_NAMES['A' if key=='a_path' else 'B'])
    paths=[ap,bp,gold,ar['source_path'],br['source_path']]
    for record in (ar,br):
        for field in ('request_receipts','packet_mappings'):
            paths.extend(receipt['path'] for receipt in record.get(field,[]))
        paths.extend(calculation_files(record))
    selected_link=config.get('a_link') if 'a_link' in config else ar.get('a_link')
    if selected_link:
        link=_load_link(selected_link,ar['source_path'])
        paths.extend([selected_link,link['word_path'],link['linked_word'],link['a_path']])
    inputs=_input_records(paths)
    if digest!=standard['sha256'] or inputs[str(gold)]!=digest:
        raise ValueError('勾稽期间金标准已变化')
    for record in (ar,br):
        if inputs[record['mapping_path']]!=record['mapping_hash'] or inputs[str(Path(record['source_path']).resolve())]!=record['source_hash']:
            raise ValueError('勾稽期间原件或语义映射已变化')
    log('读取两份统一语义映射，核对原件后按指标及实际维度生成对应清单')
    scope=config.get('validation_scope');purpose=config.get('update_purpose')
    if scope:
        if not isinstance(scope.get('label'),str) or not scope['label'].strip():raise ValueError('局部验证须明确名称')
        ar=select_mapping_scope(ar,scope['target']);br=select_mapping_scope(br,scope['source'])
        log('本次仅作局部验证：'+scope['label']+'；原整份语义及范围外未决保留')
    from .同义复核 import mapping_pairs,review_pairs
    from .统一识别 import UnifiedSemanticEngine
    pairs=mapping_pairs(target_for_purpose(ar,purpose),br,standard,routing=config.get("match_routing"));proofs=[];usage={};engine=None
    if pairs:
        engine=UnifiedSemanticEngine(validate_settings(config),gold,Path(config.get('review_cache_dir') or (run/'开放维度复核')),log,cancel)
        engine.source_hash=br['source_hash']
        equivalence_failures=[]
        # DeepSeek 官方接口经用户批准允许 4 路并发（实测无 429 限流）；火山接口仍须串行。
        from concurrent.futures import ThreadPoolExecutor
        def _review_batch(batch):
            check_cancel(cancel)
            try:
                return review_pairs(engine,batch,same_cell=False),None
            except SemanticCancelled:
                raise
            except Exception as error:
                return None,{'kind':'equivalence_review_failed','reason':str(error)[:300],
                    'pairs':[{'left':row['left']['id'],'right':row['right']['id']} for row in batch]}
        # 复核并发数按接口实测配置：火山/mimo 套餐并发即限流须串行；DeepSeek 官方实测可 4 路。
        _review_workers=max(1,int(config.get('review_concurrency') or 1))
        batches=[pairs[start:start+12] for start in range(0,len(pairs),12)]
        if _review_workers==1:
            for batch in batches:
                proof,failure=_review_batch(batch)
                if proof is not None:proofs.append(proof)
                if failure is not None:equivalence_failures.append(failure)
        else:
            with ThreadPoolExecutor(max_workers=_review_workers) as _executor:
                for proof,failure in _executor.map(_review_batch,batches):
                    if proof is not None:proofs.append(proof)
                    if failure is not None:equivalence_failures.append(failure)
        usage=dict(engine.usage)
        review_paths=[]
        for proof in proofs:
            for review in proof['rounds']:
                review_paths.extend(review[key]['path'] for key in ('task','receipt'))
        inputs.update(_input_records(review_paths))
    plan=build_fact_update_plan(ar,br,standard,equivalence_proofs=proofs,update_purpose=purpose)
    plan['issues'].extend(equivalence_failures)
    structure=None;structure_files={}
    needs=bool(plan['unplaced_sources'] or any(row.get('kind')=='missing_source' for row in plan['issues']))
    if not scope and needs and all(record.get('coverage',{}).get('complete') for record in (ar,br)) and read_workbook(ar['source_path']).get('names'):
        from .统一结构 import plan_structural_layout,recognize_adjusted_layout,load_structural_target
        if engine is None:engine=UnifiedSemanticEngine(validate_settings(config),gold,run/'结构语义识别',log,cancel)
        log('按已识别含义复核必要行列变化；调整后的表格将独立重新识别，新增财务格暂不填数')
        report=plan_structural_layout(engine,ar,br,run/'结构调整')
        structure_files['structure_review']=report['report_path']
        if report['accepted']:
            structure_files['adjusted_a']=report['layout']['source_path']
            try:
                def keep_adjusted_mapping(record):structure_files['adjusted_a_mapping']=record['mapping_path']
                structure,working=recognize_adjusted_layout(engine,ar,br,report,run/'结构调整',progress=keep_adjusted_mapping)
                load_structural_target(structure,ar,br,standard)
            except ValueError as error:
                structure=None
                plan['issues'].append({'kind':'structure_semantics_unverified','reason':str(error)})
            else:
                structure_files['adjusted_a_mapping']=working['mapping_path']
                inputs.update(_input_records([report['report_path'],working['source_path'],working['mapping_path']]))
                proofs=[];pairs=mapping_pairs(working,br,standard,routing=config.get("match_routing"))
                engine.source_hash=br['source_hash']
                for start in range(0,len(pairs),12):
                    check_cancel(cancel);proofs.append(review_pairs(engine,pairs[start:start+12],same_cell=False))
                plan=build_fact_update_plan(working,br,standard,equivalence_proofs=proofs)
        elif report['proposal']['operations'] or report['proposal']['unresolved']:
            plan['issues'].append({'kind':'structure_not_approved','reason':'行列变化的语义依据尚未通过独立复核'})
        usage=dict(engine.usage)
    if proofs:plan['equivalence_proofs']=proofs
    if structure:
        plan['structure']=structure;plan['word_operations']=structure['layout']['operations']
    plan['unused_sources']=[]
    if not scope and plan['unplaced_sources'] and all(record.get('coverage',{}).get('complete') for record in (ar,br)):
        from .统一未采用来源 import review_unused_sources
        if engine is None:engine=UnifiedSemanticEngine(validate_settings(config),gold,run/'来源采用复核',log,cancel)
        log('按双方语义和原文复核额外来源；同业务的新成员仍保留给行列调整')
        reviewed=review_unused_sources(engine,working if structure else ar,br,plan,run/'未采用来源依据')
        plan['unused_source_reviews']=reviewed['proofs'];plan['unused_sources']=reviewed['unused_sources']
        unused_ids={row['source_id'] for row in plan['unused_sources']}
        plan['unplaced_sources']=[row for row in plan['unplaced_sources'] if row['id'] not in unused_ids]
        plan['issues'].extend(reviewed['issues'])
        review_path=save_json(run/'未采用来源语义复核.json',reviewed)
        plan['unused_source_review_path']=review_path
        structure_files['unused_source_review']=review_path
        review_paths=[review_path]
        for proof in reviewed['proofs']:
            for record in proof['rounds']:review_paths.extend(record[key]['path'] for key in ('task','receipt'))
        inputs.update(_input_records(review_paths));usage=dict(engine.usage)
        if plan['unused_sources'] and not plan['unplaced_sources'] and not any(row.get('kind')=='missing_source' for row in plan['issues']):
            # 额外来源已有不采用依据、目标也不缺来源时，不再需要此前被否决的增行提案。
            plan['resolved_structure_issues']=[row for row in plan['issues'] if row.get('kind')=='structure_not_approved']
            plan['issues']=[row for row in plan['issues'] if row.get('kind')!='structure_not_approved']
    if not all(record.get('scope_coverage' if scope else 'coverage',{}).get('complete') for record in (ar,br)):
        plan['issues'].append({'kind':'coverage_not_verified','reason':'当前映射尚未证明整份载体的语义覆盖完整，须处理未决格或继续识别'})
    for row in plan['unplaced_sources']:
        plan['issues'].append({'kind':'unplaced_source','location':row['source_reference']['location'],
                               'reason':'此业务事实未找到目标位置，须判断新增披露、行列变化或本次不采用'})
    for row in plan['updates']:
        definition=standard['metric'][row['metric_id']]['definition']
        row['semantic_field']=' / '.join(definition[key] for key in ('economic_object','measure','measurement_basis'))
    ready=not plan['issues'] and bool(plan['updates'] or plan['retained_blanks'])
    plan.update(schema='附注更新清单-v2',complete=ready,inputs=inputs,a_mapping=str(ap),b_mapping=str(bp),
                gold_path=str(gold),gold_hash=digest,a_link=str(Path(selected_link).resolve()) if selected_link else None)
    if scope:plan.update(validation_scope=copy.deepcopy(scope),scope_coverage={'target':ar['scope_coverage'],'source':br['scope_coverage']})
    if purpose:plan['update_purpose']=copy.deepcopy(purpose)
    _verify_inputs(inputs);check_cancel(cancel)
    plan_path=save_json(run/'更新清单.json',plan)
    save_json(run/'更新清单校验.json',{'path':plan_path,'sha256':file_hash(plan_path)})
    save_json(run/'待核实事项.json',plan['issues'])
    trace=run/'逐格对应与更新核对.xlsx';trace_workbook(trace,plan)
    return {'status':'complete' if ready else 'partial','message':('局部验证：' if scope else '')+f"新版语义更新清单已保存：拟更新{len(plan['updates'])}格，本次不采用{len(plan['unused_sources'])}个来源，{len(plan['issues'])}项待核实。尚未生成更新文件。",
            'files':{'update_plan':plan_path,'trace':str(trace),**structure_files},'usage':usage}


def _match(config,run,log,cancel):
    ap=_path(config,"a_mapping","附注表格语义映射文件");bp=_path(config,"b_mapping","更新数据语义映射文件")
    versions=[_read(path).get('schema_version') for path in (ap,bp)]
    if 2 in versions:
        if versions!=[2,2]:raise ValueError('两份映射使用不同标准层次，须完成旧定义迁移后再勾稽')
        return _match_v2(config,run,log,cancel,ap,bp)
    _check_b_mapping_mode(config,bp)
    ar,a=_load_mapping(ap,"A");br,b=_load_mapping(bp,"B")
    _selected_sources(config,ar,br)
    gold,gold_hash=_gold(config,ar)
    ensure_compatible(br,gold)
    selected_link=config.get("a_link") if "a_link" in config else ar.get("a_link")
    link=None
    if selected_link:link=_load_link(selected_link,a["path"])
    paths=[ap,bp,gold,ar["gold_path"],br["gold_path"],ar["source_path"],br["source_path"],ar["original_path"],br["original_path"]]
    if br.get("scope_selection"):
        from .识别范围 import build_target_meanings
        current_target=build_target_meanings(link or {"scope":{"chapter_title":CHAPTER_TITLE},"tables":ar.get("context",[])},a)
        target_hash=hashlib.sha256(_json(current_target).encode("utf-8")).hexdigest()
        if not current_target.get("disclosures") or target_hash!=br["scope_selection"].get("target_hash"):
            raise ValueError("更新数据映射的识别范围不适用于当前附注披露，请按当前附注重新识别更新数据")
    if link:paths.extend([selected_link,link["word_path"],link["linked_word"],link["a_path"]])
    inputs=_input_records(paths)
    log("读取两份已保存映射，以相同金标准语义建立对应；从更新数据表的实际单元格取数")
    plan=build_update_plan(a,b,ar["mappings"],br["mappings"])
    operations=[];working=copy.deepcopy(ar);working_snapshot=a;usage={};engine=None
    working_gold=_gold_snapshot(gold,gold_hash)
    working.update(gold_path=str(working_gold),gold_hash=gold_hash,selected_gold_path=str(gold))
    inputs.update(_input_records([working_gold]))
    needs=bool(plan.get("unplaced_sources") or any(i.get("kind")=="missing_source" for i in plan.get("issues",[]) if isinstance(i,dict)))
    if needs and a.get("names"):
        engine=_engine(config,gold,{"a_map":file_hash(ap),"b_map":file_hash(bp),"action":"structure"},log,cancel)
        engine.source_hash=b["sha256"];engine.cache_source_hash=hashlib.sha256(_json(inputs).encode()).hexdigest()
        log("核验语义差异是否需要增减行列；已保存的附注表格和更新数据语义映射直接复用")
        structure=resolve_structure(engine,a,b,ar,br,plan,str(run))
        usage=dict(getattr(engine,"usage",{}))
        if structure.get("issues"):plan["issues"].extend(structure["issues"])
        else:
            operations=structure["operations"];working_snapshot=structure["a_snapshot"]
            working.update(source_path=working_snapshot["path"],source_hash=working_snapshot["sha256"],
                original_path=working_snapshot["path"],original_hash=working_snapshot["sha256"],
                mappings=structure["a_mappings"],context=working_snapshot.get("context",[]),
                confirmed_business_ranges={},recalculated=False,recalculation={})
            plan=build_update_plan(working_snapshot,b,working["mappings"],br["mappings"])
            inputs.update(_input_records([working_snapshot["path"]]))
    plan["unused_sources"]=[]
    if plan.get("unplaced_sources"):
        from .结构 import review_unused_sources,validate_unused_sources
        if engine is None:
            engine=_engine(config,gold,{"a_map":file_hash(ap),"b_map":file_hash(bp),"action":"source_selection"},log,cancel)
            engine.source_hash=b["sha256"];engine.cache_source_hash=hashlib.sha256(_json(inputs).encode()).hexdigest()
        log("逐项核实更新数据中尚无接收位置的业务，区分必要新增行列和本次附注不采用的来源")
        decisions=review_unused_sources(engine,working_snapshot,b,working["mappings"],br["mappings"],plan)
        approved=validate_unused_sources(working_snapshot,b,working["mappings"],br["mappings"],plan,decisions.get("unused_sources",[]))
        plan["unused_sources"]=approved
        positions={(m["sheet"],m["cell"]) for m in approved}
        plan["unplaced_sources"]=[m for m in plan["unplaced_sources"] if (m["sheet"],m["cell"]) not in positions]
        plan["issues"].extend(decisions.get("issues",[]));usage=dict(getattr(engine,"usage",{}))
        plan["issues"].extend({"kind":"unplaced_source","sheet":m["sheet"],"cell":m["cell"],
            "reason":"更新数据表已确认的业务格尚未确定接收位置或本次不采用的明确依据"} for m in plan["unplaced_sources"])
    ready=not plan["issues"] and bool(plan["updates"] or plan.get("retained_blanks"))
    plan.update(schema=PLAN_SCHEMA,complete=ready,inputs=inputs,a_mapping=str(ap),b_mapping=str(bp),
        working_a=working,gold_path=str(gold),gold_hash=gold_hash,word_operations=operations,a_link=str(selected_link) if selected_link else None,usage=usage)
    _verify_inputs(inputs);check_cancel(cancel)
    plan_path=save_json(run/"更新清单.json",plan)
    save_json(run/"更新清单校验.json",{"path":plan_path,"sha256":file_hash(plan_path)})
    trace=run/"逐格对应与更新核对.xlsx";trace_workbook(trace,plan)
    if plan["issues"]:save_json(run/"待核实事项.json",plan["issues"])
    return {"status":"complete" if ready else "partial",
        "message":f"更新清单已保存：拟更新{len(plan['updates'])}格，原空白保留{len(plan.get('retained_blanks',[]))}格，本次不采用{len(plan.get('unused_sources',[]))}个来源，{len(plan['issues'])}项待核实。尚未生成更新后附注表格或更新Word。",
        "files":{"update_plan":plan_path,"trace":str(trace)},"usage":usage}

def _write_v2(config,run,log,cancel,path,plan):
    from .统一语义 import read_standard_v2
    from .统一勾稽 import load_fact_mapping,verify_fact_update_plan,derive_fact_mapping
    plan_hash=file_hash(path)
    _verify_inputs(plan['inputs'])
    _check(plan['gold_path'],plan['gold_hash'],'更新清单绑定的金标准')
    if config.get('gold_path'):_check(config['gold_path'],plan['gold_hash'],'当前金标准')
    standard=read_standard_v2(plan['gold_path'])
    ar=load_fact_mapping(plan['a_mapping'],standard);br=load_fact_mapping(plan['b_mapping'],standard)
    for key,record in (('a_path',ar),('b_path',br),('a_mapping',ar),('b_mapping',br)):
        if config.get(key):
            _check(config[key],record['mapping_hash'] if key.endswith('mapping') else record['source_hash'],'当前选择的表格或语义映射')
    selected_link=config.get('a_link') if 'a_link' in config else plan.get('a_link')
    old_link=None
    if selected_link:
        if not plan.get('a_link') or Path(selected_link).resolve()!=Path(plan['a_link']).resolve():
            raise ValueError('当前Word位置关联与更新清单不同，请重新生成更新清单')
        if str(Path(selected_link).resolve()) not in plan['inputs']:
            raise ValueError('更新清单未绑定Word位置关联')
        old_link=_load_link(selected_link,ar['source_path'])
    fresh=verify_fact_update_plan(plan,ar,br,standard)
    working=fresh.pop('_working_target')
    check_cancel(cancel);log('重新核对语义与取数后生成更新文件，并保存可复用的逐格语义映射')
    prime=run/(('局部验证_' if plan.get('validation_scope') else '')+'更新后附注表格'+Path(working['source_path']).suffix)
    write_updated_workbook(working['source_path'],prime,fresh)
    _verify_inputs(plan['inputs']);_check(path,plan_hash,'本次更新清单');check_cancel(cancel)
    files={'a_prime':str(prime)}
    if old_link:
        word=run/'更新后附注.docx'
        result=bridge({'action':'update','word':old_link['linked_word'],'excel':str(prime),'output':str(word),
                       'operations':plan.get('word_operations',[]),'chapter_title':CHAPTER_TITLE},run,cancel)
        if not result.get('output_verified'):raise ValueError('Word尚未通过回读核验')
        context=bridge({'action':'read','word':str(word),'chapter_title':CHAPTER_TITLE},run,cancel)
        new_link={'schema':LINK_SCHEMA,'word_path':str(word),'word_hash':file_hash(word),
                  'linked_word':str(word),'linked_word_hash':file_hash(word),'a_path':str(prime),
                  'a_hash':file_hash(prime),'context':context,'scope':context.get('scope',{})}
        files.update(word=str(word),word_link=save_json(run/'更新后附注表格与Word位置关联.json',new_link))
    derived=derive_fact_mapping(working,prime,fresh,path,files.get('word_link'))
    mapping_path=save_json(run/'更新后附注表格语义映射.json',derived)
    load_fact_mapping(mapping_path,standard)
    files['a_prime_mapping']=mapping_path
    trace=run/'本次更新核对.xlsx';trace_workbook(trace,fresh);files['trace']=str(trace)
    _verify_inputs(plan['inputs']);_check(path,plan_hash,'本次更新清单');check_cancel(cancel)
    return {'status':'complete','message':f"已按金标准语义更新{len(fresh['updates'])}个业务格，保存更新后表格及其独立语义映射。",
            'files':files,'updated_cells':len(fresh['updates']),'retained_blank_cells':len(fresh['retained_blanks'])}


def _write(config,run,log,cancel):
    path=_path(config,"update_plan","更新清单")
    seal=path.with_name("更新清单校验.json")
    if not seal.is_file() or _read(seal).get("sha256")!=file_hash(path):
        raise ValueError("更新清单缺少校验记录或已被改动，请重新生成清单")
    plan=_read(path)
    if plan.get('schema')=='附注更新清单-v2':return _write_v2(config,run,log,cancel,path,plan)
    if plan.get("schema")!=PLAN_SCHEMA or not plan.get("complete") or plan.get("issues"):
        raise ValueError("更新清单尚未确认完整，不能生成最终文件")
    _verify_inputs(plan["inputs"])
    locked_gold=plan.get("gold_path") or plan["working_a"]["gold_path"]
    locked_hash=plan.get("gold_hash") or plan["working_a"]["gold_hash"]
    _check(locked_gold,locked_hash,"更新清单绑定的金标准")
    if config.get("gold_path") and file_hash(config["gold_path"])!=locked_hash:
        raise ValueError("当前金标准版本与更新清单不同，请重新建立对应后生成文件")
    ensure_compatible(_read(plan["a_mapping"]),locked_gold)
    ensure_compatible(_read(plan["b_mapping"]),locked_gold)
    selected_link=config.get("a_link") if "a_link" in config else plan.get("a_link")
    if selected_link and (not plan.get("a_link") or Path(selected_link).resolve()!=Path(plan["a_link"]).resolve()):
        raise ValueError("当前Word位置关联与更新清单不同，请重新生成更新清单")
    br,b=_load_mapping(plan["b_mapping"],"B")
    working=plan["working_a"];a=_snapshot(working)
    _selected_sources(config,_read(plan["a_mapping"]),br)
    fresh=build_update_plan(a,b,working["mappings"],br["mappings"])
    from .结构 import validate_unused_sources
    approved=validate_unused_sources(a,b,working["mappings"],br["mappings"],fresh,plan.get("unused_sources",[]))
    unused_positions={(m["sheet"],m["cell"]) for m in approved}
    remaining=[m for m in fresh["unplaced_sources"] if (m["sheet"],m["cell"]) not in unused_positions]
    if fresh["issues"] or remaining or _json(fresh["updates"])!=_json(plan["updates"]) or _json(fresh.get("retained_blanks",[]))!=_json(plan.get("retained_blanks",[])):
        raise ValueError("清单与当前来源重新取数或空白保留依据不一致，请重新建立对应")
    fresh["unused_sources"]=approved;fresh["unplaced_sources"]=remaining
    check_cancel(cancel);log("按已保存更新清单生成更新后附注表格；不再调用模型识别")
    prime=run/("更新后附注表格"+Path(a["path"]).suffix)
    write_updated_workbook(a["path"],prime,fresh)
    _verify_inputs(plan["inputs"]);check_cancel(cancel)
    files={"a_prime":str(prime)}
    new_link=None
    selected_link=config.get("a_link") if "a_link" in config else plan.get("a_link")
    if selected_link:
        if not plan.get("a_link") or Path(selected_link).resolve()!=Path(plan["a_link"]).resolve():
            raise ValueError("当前Word位置关联与更新清单不同，请重新生成更新清单")
        old_link=_load_link(plan["a_link"])
        word=run/"更新后附注.docx"
        result=bridge({"action":"update","word":old_link["linked_word"],"excel":str(prime),"output":str(word),
            "operations":plan.get("word_operations",[]),"chapter_title":CHAPTER_TITLE},run,cancel)
        if not result.get("output_verified"):raise ValueError("Word尚未通过回读核验")
        context=bridge({"action":"read","word":str(word),"chapter_title":CHAPTER_TITLE},run,cancel)
        new_link={"schema":LINK_SCHEMA,"word_path":str(word),"word_hash":file_hash(word),
            "linked_word":str(word),"linked_word_hash":file_hash(word),"a_path":str(prime),
            "a_hash":file_hash(prime),"context":context,"scope":context.get("scope",{})}
        link_path=save_json(run/"更新后附注表格与Word位置关联.json",new_link)
        files.update(word=str(word),word_link=link_path)
    _verify_inputs(plan["inputs"])
    updated_context=new_link["context"].get("tables",[]) if new_link else working.get("context",[])
    snapshot=read_workbook(prime,context=updated_context)
    new_mapping=copy.deepcopy(working)
    # 新布局已由前一步确认；只换来源身份和位置，不重新猜测业务语义。
    new_mapping.update(source_path=str(prime),source_hash=snapshot["sha256"],original_path=str(prime),
        original_hash=snapshot["sha256"],carrier="A",recalculated=False,recalculation={},confirmed_business_ranges={},
        complete=True,unresolved=[],excluded=[],usage={},derived_from_mapping=plan["a_mapping"])
    # 结构变更可能使旧表范围证据失效；用新映射的范围恢复真实业务空白候选。
    for m in new_mapping["mappings"]:
        try:cell_at(snapshot,m["sheet"],m["cell"])
        except ValueError:
            area=m.get("table_range")
            if not area:raise ValueError("更新后附注表格映射包含没有范围依据的空白格")
            table={"range":area,"table_semantic":m.get("table_semantic") or m.get("semantic_field"),
                "note_ids":[m["slot_id"].split("-T",1)[0]],"scope":m["scope"],"header_cells":[]}
            materialize_business_blanks(snapshot,m["sheet"],table)
            cell_at(snapshot,m["sheet"],m["cell"])
    new_mapping["confirmed_business_ranges"]={s["name"]:s.get("confirmed_business_ranges",[]) for s in snapshot["sheets"]}
    if new_link:
        new_mapping["a_link"]=files["word_link"];new_mapping["context"]=new_link["context"].get("tables",[])
    else:new_mapping.pop("a_link",None)
    new_mapping.pop("review_path",None)
    new_mapping.pop("standard_review_path",None)
    map_path=save_json(run/"更新后附注表格语义映射.json",new_mapping)
    files["a_prime_mapping"]=map_path
    trace=run/"本次更新核对.xlsx";trace_workbook(trace,fresh);files["trace"]=str(trace)
    return {"status":"complete","message":f"已按清单更新{len(fresh['updates'])}个业务格；{len(fresh.get('retained_blanks',[]))}格因来源未提供而保持原空白，未计为已更新；{len(approved)}个来源经核实本次不采用。已保存更新后附注表格语义映射供下一轮复用。",
        "files":files,"updated_cells":len(fresh["updates"]),"retained_blank_cells":len(fresh.get("retained_blanks",[])),"unused_source_cells":len(approved)}

def run_step(step,config,log=None,cancel=None,progress=None):
    log=log or (lambda message:None);cancel=cancel or threading.Event();run=None
    latest_progress={}
    def publish(event):
        latest_progress.clear();latest_progress.update(copy.deepcopy(event))
        if progress is not None:progress(copy.deepcopy(event))
    try:
        check_cancel(cancel)
        if step not in STEPS:raise ValueError("未知流程步骤")
        output=Path(config.get("output_dir") or ROOT/"输出").resolve()
        output.mkdir(parents=True,exist_ok=True)
        run=output/(STEPS[step]+"_"+datetime.now().strftime("%Y%m%d_%H%M%S")+"_"+uuid.uuid4().hex[:6])
        run.mkdir()
        if step=="extract_a":result=_extract(config,run,log,cancel)
        elif step=="recognize_a":result=_recognize(config,"A",run,log,cancel,progress=publish)
        elif step=="recognize_b":result=_recognize(config,"B",run,log,cancel,progress=publish)
        elif step=="match":result=_match(config,run,log,cancel)
        else:result=_write(config,run,log,cancel)
        result.update(step=step,output_dir=str(run))
    except SemanticCancelled as error:
        result={"status":"cancelled","step":step,"message":str(error),"files":copy.deepcopy(latest_progress.get("files",{})),"output_dir":str(run or config.get("output_dir") or ROOT/"输出")}
    except Exception as error:
        message=f"{type(error).__name__}：{error}";secret=str(config.get("api_key") or "")
        if secret:message=message.replace(secret,"[密钥已隐藏]")
        result={"status":"failed","step":step,"message":message,"files":copy.deepcopy(latest_progress.get("files",{})),"output_dir":str(run or config.get("output_dir") or ROOT/"输出")}
    if run:save_json(run/"步骤结果.json",result)
    log(result["message"])
    return result
