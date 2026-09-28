"""将语义识别、确定性勾稽、更新后附注表格及Word回写组成一次可恢复作业。"""
from __future__ import annotations
import hashlib
from collections import Counter
import json
import shutil
import subprocess
import threading
import uuid
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from .语义 import SemanticEngine, SemanticCancelled
from .表格 import read_workbook,file_hash,build_update_plan,write_updated_workbook,recalculate_copy
from .设置 import validate_settings

ROOT=Path(__file__).resolve().parents[1]
BRIDGE=ROOT/"Word桥接"/"发布"/"WordBridge.exe"

def save_json(path,value):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    with path.open("x",encoding="utf-8-sig") as f:
        json.dump(value,f,ensure_ascii=False,indent=2,default=str)
    return str(path)

def check_cancel(cancel):
    if cancel is not None and cancel.is_set(): raise SemanticCancelled("用户取消；已完成的识别包已保留")

def bridge(request,work_dir,cancel=None):
    if not BRIDGE.is_file(): raise FileNotFoundError("Word桥接未构建，请运行项目中的构建与验证说明")
    request_file=Path(work_dir)/("Word请求_"+uuid.uuid4().hex[:8]+".json")
    save_json(request_file,request)
    check_cancel(cancel)
    completed=subprocess.run([str(BRIDGE),str(request_file)],capture_output=True,encoding="utf-8",
                             errors="strict",creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0))
    check_cancel(cancel)
    try: result=json.loads(completed.stdout.lstrip("\ufeff").strip())
    except json.JSONDecodeError as e: raise RuntimeError("Word桥接没有返回有效结果："+completed.stderr[:1000]) from e
    if completed.returncode or not result.get("ok"):
        raise ValueError(result.get("error") or completed.stderr or "Word操作失败")
    return result

def trace_workbook(path,plan):
    book=Workbook();sheet=book.active;sheet.title="更新记录"
    sheet.append(["目标工作表","目标单元格","金标准ID","业务含义","原数据","新数据","更新数据来源","维度"])
    for item in plan.get("updates",[]):
        row=[item["sheet"],item["cell"],item.get("metric_id") or item.get("slot_id"),item.get("semantic_field"),
             item.get("old_value"),item.get("value"),
             json.dumps({"实际取数":item.get("sources",[]),"同义来源未填":item.get("blank_sources",[])},ensure_ascii=False,default=str),
             json.dumps(item.get("dimensions",{}),ensure_ascii=False,default=str)]
        sheet.append(row)
    for title, key in (("原空白保留","retained_blanks"),("本次未采用来源","unused_sources")):
        extra=book.create_sheet(title);extra.append(["工作表","单元格","金标准ID或完整语义身份","业务含义","处理依据","维度或文字证据"])
        for item in plan.get(key,[]):
            extra.append([item.get("sheet"),item.get("cell"),item.get("slot_id") or item.get("semantic_key"),item.get("semantic_field"),item.get("reason"),
                json.dumps(item.get("dimensions") or item.get("evidence") or {},ensure_ascii=False,default=str)])
    from .同义复核 import append_review_sheet
    append_review_sheet(book,plan.get('equivalence_proofs',[]))
    for ws in book:
        for row in ws:
            for cell in row:
                if isinstance(cell.value,str):cell.data_type="s"
        ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
        for col,width in {"A":24,"B":15,"C":28,"D":40,"E":20,"F":20,"G":65,"H":65}.items():ws.column_dimensions[col].width=width
    with Path(path).open("xb") as handle:book.save(handle)
    book.close()

def run_job(config,log=None,cancel=None):
    log=log or (lambda message:None);cancel=cancel or threading.Event()
    run_dir=None
    try:
        check_cancel(cancel)
        settings=validate_settings(config)
        mode=config.get("mode","update")
        if mode not in {"update","recognize_b"}:raise ValueError("未知运行方式")
        source_b=Path(config.get("source_b","")).resolve()
        if not source_b.is_file():raise ValueError("请选择存在的更新数据表")
        gold=Path(config.get("gold_path","")).resolve()
        if not gold.is_file():raise ValueError("请选择存在的正式金标准JSONL")
        source_a=Path(config.get("source_a","")).resolve() if mode=="update" else None
        if source_a and not source_a.is_file():raise ValueError("请选择存在的原始附注 Word或附注表格")
        if source_a and source_a==source_b:raise ValueError("附注表格和更新数据表须为不同文件")
        output=Path(config.get("output_dir") or ROOT/"输出").resolve()
        output.mkdir(parents=True,exist_ok=True)
        original_inputs={str(p):file_hash(p) for p in (source_a,source_b,gold) if p}
        identity={"inputs":original_inputs,"model":settings.get("model"),"base_url":settings.get("base_url"),
                  "context":config.get("context_text",""),"review":config.get("review",True),"mode":mode}
        cache_key=hashlib.sha256(json.dumps(identity,sort_keys=True,ensure_ascii=False).encode()).hexdigest()[:24]
        cache_dir=output/"识别缓存"/cache_key
        run_dir=output/("更新_"+datetime.now().strftime("%Y%m%d_%H%M%S")+"_"+uuid.uuid4().hex[:6])
        run_dir.mkdir()
        save_json(run_dir/"输入记录.json",{"files":original_inputs,"model":settings.get("model"),
                  "mode":mode,"context":config.get("context_text",""),"cache_key":cache_key})
        context=[config["context_text"]] if config.get("context_text","").strip() else []
        log("读取更新数据表的实际单元格、合并表头和可见范围")
        b=read_workbook(source_b,context=context)
        if any(c.get("formula") for s in b["sheets"] for c in s["cells"].values()):
            check_cancel(cancel);log("使用本机Excel重新计算更新数据表的独立副本，保留原文件")
            calc_path=run_dir/("更新数据表_重算副本"+source_b.suffix)
            recalculate_copy(source_b,calc_path)
            b=read_workbook(calc_path,context=context);b["recalculated"]=True
            b["original_path"]=str(source_b);b["original_hash"]=original_inputs[str(source_b)]
        save_json(run_dir/"更新数据实际结构.json",b)
        a=None;linked_word=None;word_context=None
        engine=SemanticEngine(settings,str(gold),str(cache_dir),log,cancel)
        if mode=="update":
            check_cancel(cancel)
            if source_a.suffix.lower()==".docx":
                log("提取原始附注 Word表格，保留Word连接和章节上下文")
                a_path=run_dir/"附注表格.xlsx";linked_word=run_dir/"原始附注_带连接.docx"
                context_path=run_dir/"Word表格上下文.json"
                extracted=bridge({"action":"extract","word":str(source_a),"excel":str(a_path),
                    "linked_word":str(linked_word),"context":str(context_path)},run_dir,cancel)
                word_context=json.loads(context_path.read_text(encoding="utf-8-sig"))
                for warning in extracted.get("warnings",[]):log("提取提示："+str(warning))
                a=read_workbook(a_path,context=word_context.get("tables",[])+context)
            elif source_a.suffix.lower() in {".xlsx",".xlsm"}:
                a_path=run_dir/("附注表格"+source_a.suffix)
                with source_a.open("rb") as src,a_path.open("xb") as dest:shutil.copyfileobj(src,dest)
                a=read_workbook(a_path,context=context)
            else:raise ValueError("原始附注 Word或附注表格仅支持.docx、.xlsx、.xlsm")
            a["original_hash"]=original_inputs[str(source_a)]
            a["original_path"]=str(source_a)
            save_json(run_dir/"附注表格实际结构.json",a)
            log("识别附注表格中每个业务单元格的金标准语义")
            ar=engine.recognize(a)
            save_json(run_dir/"附注表格语义识别.json",ar)
        else:ar=None
        check_cancel(cancel)
        log("识别新更新数据表的金标准语义，并独立复核")
        br=engine.recognize(b,reference=ar["mappings"] if ar else None)
        br["source_path"]=b["path"]
        if b.get("original_hash"):br["original_hash"]=b["original_hash"]
        save_json(run_dir/"更新数据语义识别.json",br)
        def finish(status,message,**kwargs):
            usage=Counter(ar.get("usage",{}) if ar else {})
            usage.update(getattr(engine,"usage",{}))
            result={"status":status,"message":message,"output_dir":str(run_dir),"usage":dict(usage),**kwargs}
            save_json(run_dir/"运行结果.json",result);log(message);return result
        if mode=="recognize_b":
            return finish("complete" if br.get("complete") else "partial",
                f"更新数据表已确认{len(br['mappings'])}个业务格；待核实{len(br.get('unresolved',[]))}项",
                prediction=str(run_dir/"更新数据语义识别.json"))
        # 完整性必须在最终写入前成立，识别疑点不会被旧数或空值掩盖。
        if not ar.get("complete") or not br.get("complete"):
            save_json(run_dir/"待核实事项.json",{"A":ar.get("unresolved",[]),"B":br.get("unresolved",[])})
            return finish("partial","识别结果已保存；存在待核实语义，尚未发布更新后的附注")
        for path,digest in original_inputs.items():
            if file_hash(path)!=digest:raise ValueError("识别期间原始资料发生变化，本次写入停止")
        check_cancel(cancel);log("按共同金标准语义勾稽附注表格与更新数据表")
        plan=build_update_plan(a,b,ar["mappings"],br["mappings"])
        operations=[];working_a=a
        structure_needed=bool(plan.get("unplaced_sources") or any(i.get("kind")=="missing_source" for i in plan.get("issues",[])))
        if structure_needed and a.get("names"):
            from .结构 import resolve_structure
            log("根据语义差异核验是否需要增减目标表行列")
            structure=resolve_structure(engine,a,b,ar,br,plan,str(run_dir))
            operations=structure.get("operations",[])
            if structure.get("issues"):
                plan["issues"].extend(structure["issues"])
            else:
                working_a=structure["a_snapshot"]
                plan=build_update_plan(working_a,b,structure["a_mappings"],br["mappings"])
        if plan.get("unplaced_sources"):
            plan["issues"].extend({"kind":"unplaced_source","sheet":m["sheet"],"cell":m["cell"],
                "reason":"更新数据表存在已识别业务数据，但附注表格中没有已确认接收位置","slot_id":m["slot_id"]} for m in plan["unplaced_sources"])
        plan["word_operations"]=operations
        save_json(run_dir/"更新清单.json",plan)
        trace_workbook(run_dir/"更新核对表.xlsx",plan)
        if plan.get("issues"):
            save_json(run_dir/"待核实事项.json",plan["issues"])
            return finish("partial",f"已生成核对表；{len(plan['issues'])}项对应关系待核实，尚未发布最终文件")
        if not plan["updates"]:
            return finish("partial","未识别出可更新的财务数据单元格，尚未发布最终文件")
        check_cancel(cancel);log("将更新数据表原始单元格数据写入附注表格的对应语义位置")
        a_prime=run_dir/("更新后附注表格"+Path(working_a["path"]).suffix)
        write_updated_workbook(working_a["path"],a_prime,plan)
        for path,digest in original_inputs.items():
            if file_hash(path)!=digest:raise ValueError("写入前后原始资料发生变化，未发布Word")
        result_files={"excel":str(a_prime),"trace":str(run_dir/"更新核对表.xlsx"),"updated_cells":len(plan["updates"])}
        if linked_word:
            check_cancel(cancel);log("将更新后附注表格回写Word原表并重新读取验证")
            final_word=run_dir/"更新后附注.docx"
            written=bridge({"action":"update","word":str(linked_word),"excel":str(a_prime),
                "output":str(final_word),"operations":operations},run_dir,cancel)
            if not written.get("output_verified"):raise ValueError("Word回写尚未通过重新读取验证")
            result_files["word"]=str(final_word)
        return finish("complete",f"完成{len(plan['updates'])}个业务数据格更新；原始文件保持不变",**result_files)
    except SemanticCancelled as e:
        result={"status":"cancelled","message":str(e),"output_dir":str(run_dir or config.get("output_dir") or ROOT/"输出")}
        if run_dir:save_json(run_dir/("取消记录_"+uuid.uuid4().hex[:6]+".json"),result)
        log(result["message"]);return result
    except Exception as e:
        message=f"{type(e).__name__}：{e}"
        secret=str(config.get("api_key") or "")
        if secret:message=message.replace(secret,"[密钥已隐藏]")
        result={"status":"failed","message":message,"output_dir":str(run_dir or config.get("output_dir") or ROOT/"输出")}
        if run_dir:save_json(run_dir/"运行结果.json",result)
        log("本次未完成："+message);return result
