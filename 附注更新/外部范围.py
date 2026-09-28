"""采纳真实人工或高级AI范围意见；只确认识别范围，不生成财务语义或模型轮次。"""
from __future__ import annotations
import copy
import json
from datetime import datetime
from pathlib import Path
import uuid
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple,range_boundaries
from . import 分步流程 as flow
from .识别范围 import build_target_meanings,_entry,_contains,_grid_rows,_hash
from .表格 import read_workbook,file_hash,materialize_candidate_blanks
from .金标准 import read_standard
from .语义 import _unique_json_object

SCHEMA="附注外部范围复核-v1"
PREVIEW_SCHEMA="附注外部范围预览-v1"
METHOD="external_scope_review"
BASIS="source_grid_v1"
INSTRUCTIONS="逐区核对完整第六章的指标、业务总体和金额/比例/数量类型；指标与维度分层。相同指标的新增客户、项目、币种、期间、排名数量不能因目标未出现而排除；主表或附注等载体、空白或零值也不能作为排除理由。不确定必须保留。只有业务总体、指标或值类型确实不同，且完整目标没有需要的对应业务时，才能排除；同时引用来源与目标原文说明差异。整表混有相关与无关业务时须拆区。"


def _read(path):return json.loads(Path(path).read_text(encoding="utf-8-sig"),object_pairs_hook=_unique_json_object)
def _text(v):return isinstance(v,str) and bool(v.strip())


def _save(path,value):
    path=Path(path).resolve()
    with path.open("x",encoding="utf-8-sig") as f:json.dump(value,f,ensure_ascii=False,indent=2)
    if _read(path)!=value:raise ValueError("外部范围文件保存回读不一致")
    return str(path)


def _candidate_snapshot(snapshot):
    candidate=copy.deepcopy(snapshot)
    for sheet in candidate["sheets"]:
        sheet.pop("confirmed_business_ranges",None)
        sheet["cells"]={a:c for a,c in sheet["cells"].items() if not c.get("derived_blank")}
    materialize_candidate_blanks(candidate)
    rebuilt={s["name"]:s for s in candidate["sheets"]}
    for sheet in snapshot["sheets"]:
        for address,cell in sheet["cells"].items():
            if coordinate_to_tuple(address)!=(cell.get("row"),cell.get("column")):raise ValueError("范围来源格地址与真实行列不一致")
            if cell.get("derived_blank") and cell!=rebuilt[sheet["name"]]["cells"].get(address):raise ValueError("候选空白无法从当前真实来源重建，不能采用范围")
    return candidate


def _catalogue(snapshot):
    return [{"sheet":s["name"],"max_row":s["max_row"],"max_column":s["max_column"],
             "merges":copy.deepcopy(s.get("merges",[])),"grid_rows":_grid_rows({a:_entry(c) for a,c in s["cells"].items()})}
            for s in snapshot["sheets"]]


def _prepare(source_path,a_link_path,gold_path):
    source=Path(source_path).resolve();link_path=Path(a_link_path).resolve();gold=Path(gold_path).resolve()
    link=flow._load_link(link_path);read_standard(gold)
    paths=[source,link_path,gold,Path(link["a_path"]),Path(link["word_path"]),Path(link["linked_word"])]
    inputs={str(p.resolve()):file_hash(p) for p in paths}
    snapshot=_candidate_snapshot(read_workbook(source))
    target=build_target_meanings(link,read_workbook(link["a_path"],context=link["context"]["tables"]))
    if not target.get("word_tables"):raise ValueError("完整目标缺少可回读的Word实际表格")
    catalogue=_catalogue(snapshot)
    meta={"source_path":str(source),"a_link_path":str(link_path),"gold_path":str(gold),"inputs":inputs,
          "source_hash":snapshot["sha256"],"target_hash":_hash(target),"gold_hash":inputs[str(gold)],
          "candidate_basis":BASIS,"candidate_hash":_hash(catalogue),"candidate_count":sum(len(s["cells"]) for s in snapshot["sheets"])}
    for path,digest in inputs.items():flow._check(path,digest,"范围依据")
    return meta,snapshot,target,catalogue


def _checked(opinion,meta,snapshot,target):
    if not isinstance(opinion,dict) or opinion.get("schema")!=SCHEMA:raise ValueError("请选择外部范围复核意见")
    for key,value in {"carrier":"B",**{k:meta[k] for k in ("source_hash","target_hash","gold_hash","candidate_basis","candidate_hash","candidate_count")}}.items():
        if opinion.get(key)!=value:raise ValueError("范围意见绑定不一致："+key)
    if opinion.get("target_meanings")!=target or opinion.get("source_catalogue")!=_catalogue(snapshot):raise ValueError("意见所附原始文字或完整目标已变化")
    reviewer=opinion.get("reviewer")
    if not isinstance(reviewer,dict) or not _text(reviewer.get("name")) or reviewer.get("method") not in {"human","advanced_ai"}:raise ValueError("请填写真实审核者和审核方式")
    try:
        if not _text(reviewer.get("reviewed_at")):raise ValueError()
        datetime.fromisoformat(reviewer["reviewed_at"])
    except (ValueError,TypeError):raise ValueError("请填写真实审核时间") from None
    raw=opinion.get("areas")
    if not isinstance(raw,list):raise ValueError("范围意见须为区域清单")
    sheets={s["name"]:s for s in snapshot["sheets"]};seen=set();areas=[];out=[];uncertain=[];selected={s:[] for s in sheets}
    target_text={t for k in ("disclosures","chapter_context","table_business") for t in target.get(k,[]) if isinstance(t,str)}
    target_text.update(t for tab in target.get("word_tables",[]) for row in tab.get("text_rows",[]) for _,t in row["cells"] if isinstance(t,str))
    for item in raw:
        if not isinstance(item,dict) or set(item)-{"sheet","range","decision","reason","evidence_cells","semantic_comparison"}:raise ValueError("外部范围不得声明模型轮次或额外结论")
        name=item.get("sheet");sheet=sheets.get(name)
        if sheet is None:raise ValueError("范围引用未知工作表")
        region=item.get("range")
        if not _text(region):raise ValueError("范围缺少真实坐标")
        region=region.replace("$","").upper()
        try:
            bounds=range_boundaries(region)
            if not all(type(v) is int for v in bounds) or not(1<=bounds[0]<=bounds[2]<=sheet["max_column"] and 1<=bounds[1]<=bounds[3]<=sheet["max_row"]):raise ValueError()
        except (ValueError,TypeError):raise ValueError("范围超出真实工作表边界") from None
        cells=sorted((a for a in sheet["cells"] if _contains(a,bounds)),key=coordinate_to_tuple)
        if not cells or any((name,a) in seen for a in cells):raise ValueError("范围没有真实候选或重复覆盖原格")
        seen.update((name,a) for a in cells)
        decision=item.get("decision");reason=item.get("reason");evidence=item.get("evidence_cells")
        if decision not in {"include","out_of_scope","uncertain"} or not _text(reason):raise ValueError("范围判断缺少状态或业务理由")
        if not isinstance(evidence,list) or any(not isinstance(a,str) or a not in sheet["cells"] or _entry(sheet["cells"][a]).get("kind")!="text" for a in evidence) or len(evidence)!=len(set(evidence)):raise ValueError("范围证据须为同表不重复的真实文字格")
        if decision!="uncertain" and not evidence:raise ValueError("确定范围必须有原始文字依据")
        area={"sheet":name,"range":region,"decision":decision,"reason":reason.strip(),"evidence_cells":list(evidence)}
        comparison=item.get("semantic_comparison")
        if decision=="out_of_scope":
            fields={"source_business","target_business","difference_kind","reason","target_evidence"}
            if not isinstance(comparison,dict) or set(comparison)!=fields:raise ValueError("排除须解释来源和完整目标的业务差异")
            if comparison["difference_kind"] not in {"business_population","indicator","value_type"}:raise ValueError("开放维度取值、载体、空白或金额差异不能用于排除")
            if any(not _text(comparison[k]) for k in ("source_business","target_business","reason")) or comparison["source_business"].strip()==comparison["target_business"].strip():raise ValueError("不能排除相同业务，须说明真实指标差异")
            cited=comparison["target_evidence"]
            if not isinstance(cited,list) or not cited or any(not isinstance(t,str) or t not in target_text for t in cited):raise ValueError("目标依据必须逐字引用完整附注原文")
            area["semantic_comparison"]=copy.deepcopy(comparison)
            out.append({"sheet":name,"range":region,"cells":cells,"reason":reason.strip(),"evidence_cells":list(evidence),"area_indexes":[len(areas)]})
        elif comparison is not None:raise ValueError("业务差异说明仅用于排除意见")
        else:
            selected[name].extend(cells)
            if decision=="uncertain":uncertain.append({"sheet":name,"cells":cells,"reason":reason.strip()})
        areas.append(area)
    expected={(s["name"],a) for s in snapshot["sheets"] for a in s["cells"]}
    if seen!=expected:raise ValueError("范围意见没有完整覆盖所有原格和候选空白；遗漏格必须列为不确定并保留")
    selected={name:sorted(cells,key=coordinate_to_tuple) for name,cells in selected.items()}
    return {"source_hash":meta["source_hash"],"target_meanings":copy.deepcopy(target),"target_hash":meta["target_hash"],"gold_hash":meta["gold_hash"],
            "candidate_basis":BASIS,"candidate_hash":meta["candidate_hash"],"areas":areas,"out_of_scope":out,"unresolved":uncertain,"selected_cells":selected,
            "coverage":{"candidate_count":len(expected),"selected_count":sum(map(len,selected.values())),"out_of_scope_count":sum(len(a["cells"]) for a in out)}}


def export_scope_review_template(source_path,a_link_path,output_path,*,gold_path):
    meta,snapshot,target,catalogue=_prepare(source_path,a_link_path,gold_path)
    value={"schema":SCHEMA,"carrier":"B",**{k:meta[k] for k in ("source_hash","target_hash","gold_hash","candidate_basis","candidate_hash","candidate_count")},
           "instructions":INSTRUCTIONS,"reviewer":{"name":"","method":"","reviewed_at":""},"target_meanings":target,"source_catalogue":catalogue,
           "areas":[{"sheet":s["name"],"range":f"A1:{get_column_letter(s['max_column'])}{s['max_row']}","decision":"uncertain","reason":"","evidence_cells":[]} for s in snapshot["sheets"] if s["cells"]]}
    return _save(output_path,value)


def preview_external_scope_review(source_path,a_link_path,opinion_path,*,gold_path):
    opinion_path=Path(opinion_path).resolve();digest=file_hash(opinion_path)
    meta,snapshot,target,_=_prepare(source_path,a_link_path,gold_path);opinion=_read(opinion_path)
    selection=_checked(opinion,meta,snapshot,target)
    flow._check(opinion_path,digest,"外部范围意见")
    return {"schema":PREVIEW_SCHEMA,**meta,"opinion_path":str(opinion_path),"opinion_sha256":digest,"reviewer":copy.deepcopy(opinion["reviewer"]),
            "coverage":selection["coverage"],"uncertain_count":sum(len(i["cells"]) for i in selection["unresolved"]),"changes":selection["areas"],"selection":selection}


def apply_external_scope_review(preview,output_dir,*,confirmed=False,confirm_method=""):
    if confirmed is not True or not _text(confirm_method):raise ValueError("请预览并明确确认外部范围意见")
    if not isinstance(preview,dict) or preview.get("schema")!=PREVIEW_SCHEMA:raise ValueError("缺少外部范围预览")
    fresh=preview_external_scope_review(preview["source_path"],preview["a_link_path"],preview["opinion_path"],gold_path=preview["gold_path"])
    if fresh!=preview:raise ValueError("范围意见、来源或预览已变化，请重新预览")
    run=Path(output_dir).resolve()/("导入外部范围_"+datetime.now().strftime("%Y%m%d_%H%M%S_%f")+"_"+uuid.uuid4().hex[:6]);run.mkdir(parents=True)
    opinion=run/"外部范围原始意见.json"
    with opinion.open("xb") as f:f.write(flow._check(fresh["opinion_path"],fresh["opinion_sha256"],"外部意见").read_bytes())
    flow._check(opinion,fresh["opinion_sha256"],"意见副本")
    receipt={"schema":"附注外部范围采纳-v1",**{k:fresh[k] for k in ("source_path","a_link_path","gold_path","inputs","source_hash","target_hash","gold_hash","candidate_basis","candidate_hash","reviewer")},
             "opinion_path":str(opinion),"opinion_sha256":fresh["opinion_sha256"],"confirmed_at":datetime.now().astimezone().isoformat(),"confirm_method":confirm_method}
    adoption=_save(run/"外部范围确认记录.json",receipt)
    result=copy.deepcopy(fresh["selection"]);result.update(review_method=METHOD,external_scope_review={**receipt,"adoption_path":adoption,"adoption_sha256":file_hash(adoption)})
    read_verified_scope(result,read_workbook(fresh["source_path"]),result["target_meanings"])
    path=_save(run/"更新数据识别范围.json",result)
    return {"status":"partial" if result["unresolved"] else "complete","message":f"范围意见已保存：保留{result['coverage']['selected_count']}格，范围外{result['coverage']['out_of_scope_count']}格；不确定格全部保留，尚未识别财务语义。","files":{"b_scope":path},"coverage":result["coverage"]}


def read_verified_scope(selection,snapshot,target_meanings):
    """外部来源严格重验；旧模型范围返回None，由调用方保留原双轮路径。"""
    if not isinstance(selection,dict) or ("external_scope_review" not in selection and selection.get("review_method")!=METHOD):return None
    proof=selection.get("external_scope_review")
    if selection.get("review_method")!=METHOD or selection.get("candidate_basis")!=BASIS or not isinstance(proof,dict):raise ValueError("外部范围缺少实际审核来源或候选依据")
    receipt=_read(flow._check(proof.get("adoption_path",""),proof.get("adoption_sha256"),"外部范围确认记录"))
    if receipt!={k:v for k,v in proof.items() if k not in {"adoption_path","adoption_sha256"}}:raise ValueError("范围审核来源或确认方式已变化")
    if receipt.get("schema")!="附注外部范围采纳-v1" or not _text(receipt.get("confirm_method")) or not _text(receipt.get("confirmed_at")):raise ValueError("外部范围未经明确确认")
    for path,digest in receipt["inputs"].items():flow._check(path,digest,"原范围依据")
    meta,actual,target,_=_prepare(receipt["source_path"],receipt["a_link_path"],receipt["gold_path"])
    if any(receipt.get(k)!=meta[k] for k in ("inputs","source_hash","target_hash","gold_hash","candidate_basis","candidate_hash")):raise ValueError("原范围来源或目标身份不一致")
    if snapshot.get("sha256")!=meta["source_hash"] or _hash(_catalogue(_candidate_snapshot(snapshot)))!=meta["candidate_hash"]:raise ValueError("范围不适用于当前实际来源或候选空白")
    if target_meanings!=target:raise ValueError("范围不适用于当前完整目标披露")
    opinion=_read(flow._check(receipt["opinion_path"],receipt["opinion_sha256"],"外部范围意见副本"))
    if opinion.get("reviewer")!=receipt.get("reviewer"):raise ValueError("范围审核者与原意见不一致")
    rebuilt=_checked(opinion,meta,actual,target)
    if {k:v for k,v in selection.items() if k not in {"review_method","external_scope_review"}}!=rebuilt:raise ValueError("范围清单与真实外部意见或原格完整分区不一致")
    return copy.deepcopy(selection)
