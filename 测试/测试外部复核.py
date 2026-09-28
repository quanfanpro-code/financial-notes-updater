"""外部复核导入本地测试；意见均为明确的模拟审核，不访问模型。"""
import copy
import importlib
import json
from pathlib import Path
import unittest
import uuid
from unittest.mock import patch
from openpyxl import Workbook
from 附注更新 import 分步流程 as flow
from 附注更新.表格 import read_workbook, file_hash, materialize_business_blanks

ROOT=Path(__file__).resolve().parents[1]

class ExternalReviewTests(unittest.TestCase):
    def setUp(self):
        self.work=ROOT/"测试结果"/("外部复核_"+uuid.uuid4().hex);self.work.mkdir(parents=True)
        self.source=self.work/"本地虚构来源.xlsx";book=Workbook();sheet=book.active
        for row in [["项目","期末余额"],["库存现金",1],["银行存款",2],["其他货币资金",None]]:sheet.append(row)
        sheet["B4"].number_format="#,##0.00";book.save(self.source);book.close()
        self.gold=self.work/"本地测试标准.jsonl"
        slot={"id":"C-N087-T001-S001","status":"active","scope":"consolidated","note":{"id":"C-N087","name":"货币资金"},"table":{"id":"C-N087-T001","name":"余额构成"},"slot":{"name":"金额余额"},"value_type":"monetary","dimensions":[{"name":"currency","required":True},{"name":"period","required":True,"type":"instant","role":"closing"},{"name":"unit","required":True},{"name":"scale","required":True}],"aggregation":{"role":"component"}}
        self.gold.write_text(json.dumps(slot,ensure_ascii=False)+"\n",encoding="utf-8-sig")
        snapshot=read_workbook(self.source);self.sheet=snapshot["sheets"][0]["name"]
        table={"range":"A1:B4","note_ids":["C-N087"],"scope":"consolidated","table_semantic":"货币资金余额","header_cells":["A1","B1","A2","A3","A4"]}
        materialize_business_blanks(snapshot,self.sheet,table)
        common={"sheet":self.sheet,"slot_id":slot["id"],"scope":"consolidated","dimensions":{"currency":"CNY","period":"2025-12-31","unit":"元","scale":1},"value_type":"monetary","semantic_field":"金额余额","reason":"本地模拟已确认依据","table_range":"A1:B4","table_semantic":"货币资金余额","reviewed":True}
        self.original={**common,"cell":"B2"}
        result={"mappings":[self.original],"excluded":[{"sheet":self.sheet,"cell":a,"category":"label","reason":"本地模拟文字标签","reviewed":True} for a in snapshot["sheets"][0]["cells"] if a not in ["B2","B3","B4"]],"unresolved":[{"sheet":self.sheet,"cell":a,"reason":"本地模拟待核实","table_range":"A1:B4"} for a in ["B3","B4"]],"out_of_scope":[],"complete":False,"gold_path":str(self.gold),"gold_hash":file_hash(self.gold),"source_hash":snapshot["sha256"]}
        self.record=flow._record(snapshot,result,"B",self.gold,self.source,[])
        self.mapping=self.write_json("原映射",self.record)
        item={k:v for k,v in common.items() if k not in ["reviewed","table_range","table_semantic"]}
        item.update(cell="B3",decision="existing",evidence_cells=["A3","B1"],reason="本地模拟外部审核：银行存款余额与期间文字一致")
        self.opinion={"schema":"附注外部语义复核-v1","carrier":"B","mapping_sha256":file_hash(self.mapping),"source_hash":file_hash(self.source),"gold_hash":file_hash(self.gold),"reviewer":{"name":"本地模拟审核者","method":"advanced_ai","reviewed_at":"2026-09-20T10:00:00+08:00"},"items":[item]}
        self.opinion_path=self.write_json("外部意见",self.opinion)
        self.before={str(p):file_hash(p) for p in [self.mapping,self.source,self.gold,self.opinion_path]}
    def write_json(self,name,value):
        p=self.work/(name+uuid.uuid4().hex+".json")
        with p.open("x",encoding="utf-8-sig") as f:json.dump(value,f,ensure_ascii=False,indent=2)
        return p
    def api(self):return importlib.import_module("附注更新.外部复核")

    def missing_context_fixture(self):
        record=copy.deepcopy(self.record)
        record["confirmed_business_ranges"][self.sheet][0]["table_range"]="B2:B2"
        record["mappings"][0]["table_range"]="B2:B2"
        for item in record["unresolved"]:item.pop("table_range",None)
        record["context"]=[{"report_scope":"consolidated","basis":"本地已确认合并口径","source":"本次已确认的报告口径"}]
        record["recognition_basis"]={"report_scope":"consolidated"}
        mapping=self.write_json("缺业务范围的原映射",record)
        op=copy.deepcopy(self.opinion);op["mapping_sha256"]=file_hash(mapping)
        op["items"][0]["business_context"]={"range":"B3:B3","note_ids":["C-N087"],"scope":"consolidated",
            "table_semantic":"银行存款期末余额","header_cells":["A3","B1"],"reason":"原A3是银行存款，B1说明期末余额，B3为其原金额格"}
        return record,mapping,op

    def test_missing_business_context_can_preview_save_reload_and_resume_without_model(self):
        api=self.api();original,mapping,op=self.missing_context_fixture()
        opinion=self.write_json("缺范围真实原文模拟意见",op)
        with patch("urllib.request.build_opener",side_effect=AssertionError("禁止网络")):
            preview=api.preview_external_review(mapping,opinion)
            change=preview["changes"][0]
            self.assertEqual(change["business_context"],op["items"][0]["business_context"])
            self.assertEqual(change["source_value"],2)
            self.assertEqual(change["context_evidence"],[{"cell":"A3","text":"银行存款"},{"cell":"B1","text":"期末余额"}])
            self.assertEqual(change["mapping"]["table_range"],"B3:B3")
            self.assertEqual(preview["remaining_pending"],1)
            result=api.apply_external_review(preview,self.work/"补上下文成果",confirmed=True,confirm_method="本地模拟确认补充原业务上下文")
            loaded,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
            self.assertEqual(loaded["mappings"][0],original["mappings"][0])
            self.assertEqual(loaded["confirmed_business_ranges"][self.sheet][0],original["confirmed_business_ranges"][self.sheet][0])
            area=loaded["confirmed_business_ranges"][self.sheet][-1]
            self.assertEqual(area,{"source_hash":file_hash(self.source),"table_range":"B3:B3","table_semantic":"银行存款期末余额",
                "note_ids":["C-N087"],"scope":"consolidated","header_cells":["A3","B1"]})
            self.assertEqual(len(snapshot["sheets"][0]["cells"]),8)
            accepted=loaded["mappings"][-1]
            self.assertEqual(accepted["table_basis_hash"],api._hash(area))
            for key in ("layout_review","layout_batch_id","independent_layout_cells","model_rounds"):self.assertNotIn(key,accepted)
            before=json.loads(Path(accepted["external_review"]["mapping_path"]).read_text(encoding="utf-8-sig"))
            self.assertEqual(before,original)
            api.validate_external_reviews(loaded,snapshot)
            engine=api._engine(loaded,self.work);engine.source_hash=snapshot["sha256"]
            snapshot["original_hash"]=loaded["original_hash"]
            kept,_=engine._reuse_previous({"mapping_path":result["files"]["b_mapping"],"mapping_sha256":file_hash(result["files"]["b_mapping"])},
                copy.deepcopy(snapshot),{"content_hash":engine._source_identity(snapshot)},None)
            self.assertIn("B3",[m["cell"] for m in kept["mappings"]])
        self.assertEqual(original,json.loads(mapping.read_text(encoding="utf-8-sig")))
        self.assertEqual(self.before,{p:file_hash(p) for p in self.before})

    def test_missing_context_rejects_conflicts_false_headers_and_invalid_slots(self):
        api=self.api();record,mapping,op=self.missing_context_fixture();snapshot=flow._snapshot(record)
        engine=api._engine(record,self.work)
        # 先证明相同夹具确实可采纳，避免所有反例只因入口不通而虚假通过。
        self.assertEqual(len(api._checked_items(record,snapshot,op,file_hash(mapping),engine)),1)
        for mutate in [lambda x:x["business_context"].update(range="B3:B4"),
                lambda x:x["business_context"].update(header_cells=[]),lambda x:x["business_context"].update(header_cells=["B3"]),
                lambda x:x["business_context"].update(header_cells=["A3","A3"]),lambda x:x["business_context"].update(header_cells=["A99"]),
                lambda x:x["business_context"].update(note_ids=["C-N087","C-N087"]),lambda x:x["business_context"].update(reason=""),
                lambda x:x["business_context"].update(layout_review={"rounds":2}),lambda x:x["business_context"].update(scope="standalone"),
                lambda x:x.update(table_basis_hash="0"*64),lambda x:x.update(replace_mapping_sha256="0"*64),
                lambda x:x.update(table_range="A1:B3"),lambda x:x.update(scope="standalone"),
                lambda x:x.update(slot_id="非现有槽位"),lambda x:x["dimensions"].pop("period"),
                lambda x:x.update(value_type="percentage"),lambda x:x.update(evidence_cells=["B3"]),
                lambda x:x.update(metric_evidence=["A2"]),lambda x:x.update(cell="A3")]:
            bad=copy.deepcopy(op);mutate(bad["items"][0])
            with self.subTest(item=bad["items"][0]),self.assertRaises(ValueError):api._checked_items(record,snapshot,bad,file_hash(mapping),engine)
        for label,mutate in [
                ("原范围已覆盖",lambda r,s:r["confirmed_business_ranges"][self.sheet][0].update(table_range="A1:B4")),
                ("合并从格",lambda r,s:s["sheets"][0].update(merges=["A3:B3"])),
                ("结构省略",lambda r,s:s["sheets"][0].update(structural_omissions=[{"range":"B3:B3"}])),
                ("章外未选",lambda r,s:r.update(scope_selection={"selected_cells":{self.sheet:["B2"]}})),
                ("报告口径冲突",lambda r,s:r["context"].append({"report_scope":"standalone"})),
                ("口径缺失",lambda r,s:(r.update(context=[]),r.pop("recognition_basis"))),
                ("口径非标",lambda r,s:r["recognition_basis"].update(report_scope=[]))]:
            changed=copy.deepcopy(record);shown=copy.deepcopy(snapshot);mutate(changed,shown)
            with self.subTest(label=label),self.assertRaises(ValueError):api._checked_items(changed,shown,op,file_hash(mapping),engine)
        for category in ("excluded","out_of_scope"):
            changed=copy.deepcopy(record);item=changed["unresolved"].pop(0);changed[category].append(item)
            with self.subTest(category=category),self.assertRaises(ValueError):api._checked_items(changed,snapshot,op,file_hash(mapping),engine)
        for value,formula in [(None,None),("",None),("=\"银行存款\"","=\"银行存款\"")]:
            shown=copy.deepcopy(snapshot);shown["sheets"][0]["cells"]["A3"].update(value=value,formula=formula)
            with self.subTest(value=value),self.assertRaises(ValueError):api._checked_items(record,shown,op,file_hash(mapping),engine)

    def test_new_a_context_requires_and_keeps_real_word_physical_boundaries(self):
        api=self.api();record,mapping,op=self.missing_context_fixture()
        record["carrier"]="A";op["carrier"]="A";snapshot=flow._snapshot(record);engine=api._engine(record,self.work)
        with self.assertRaisesRegex(ValueError,"Word.*物理"):
            api._checked_items(record,snapshot,op,file_hash(mapping),engine)
        sheet=snapshot["sheets"][0];sheet["word_table_ranges"]=[{"range":"A1:B4"}]
        self.assertEqual(len(api._checked_items(record,snapshot,op,file_hash(mapping),engine)),1)
        for physical in [[],[{"range":"A1:B2"},{"range":"A3:B4"}],[{"range":"A1:B2"}]]:
            sheet["word_table_ranges"]=physical
            with self.subTest(physical=physical),self.assertRaises(ValueError):api._checked_items(record,snapshot,op,file_hash(mapping),engine)

    def test_context_revalidation_uses_pre_adoption_source_and_rejects_removed_or_changed_area(self):
        api=self.api();record,mapping,op=self.missing_context_fixture()
        result=api.apply_external_review(api.preview_external_review(mapping,self.write_json("上下文复验意见",op)),
            self.work/"上下文复验成果",confirmed=True,confirm_method="本地模拟确认")
        loaded,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
        api.validate_external_reviews(loaded,snapshot)
        for label,mutate in [
                ("移除范围",lambda r:r["confirmed_business_ranges"][self.sheet].pop()),
                ("改表头",lambda r:r["confirmed_business_ranges"][self.sheet][-1].update(header_cells=["A4","B1"])),
                ("扩大范围",lambda r:r["confirmed_business_ranges"][self.sheet][-1].update(table_range="B3:B4")),
                ("原结构化口径被改",lambda r:r["recognition_basis"].update(report_scope="standalone")),
                ("原上下文被移除",lambda r:r.update(context=[]))]:
            bad=copy.deepcopy(loaded);mutate(bad)
            with self.subTest(label=label),self.assertRaises(ValueError):flow._load_mapping(self.write_json("篡改上下文成果",bad),complete=False)
        changed=copy.deepcopy(snapshot);changed["sheets"][0]["cells"]["A3"]["value"]="原文被改"
        with self.assertRaises(ValueError):api.validate_external_reviews(loaded,changed)
        self.assertEqual(json.loads(mapping.read_text(encoding="utf-8-sig")),record)

    def test_export_explains_optional_context_without_fabricating_reviewed_ranges(self):
        api=self.api();record,mapping,op=self.missing_context_fixture();path=self.work/"缺范围格式.json"
        api.export_review_template(mapping,path)
        doc=json.loads(path.read_text(encoding="utf-8-sig"))
        self.assertIn("business_context_instructions",doc)
        self.assertEqual(doc["business_context_instructions"]["example"]["range"],"本格:本格")
        self.assertTrue(all("business_context" not in item for item in doc["items"]))
        self.assertEqual([a["range"] for a in doc["business_range_options"]],["B2:B2"])

    def test_two_context_adoptions_preserve_earlier_proof_and_resume_all_cells(self):
        api=self.api();record,mapping,op=self.missing_context_fixture()
        with patch("urllib.request.build_opener",side_effect=AssertionError("禁止网络")):
            first=api.apply_external_review(api.preview_external_review(mapping,self.write_json("第一次补上下文",op)),
                self.work/"第一次上下文成果",confirmed=True,confirm_method="本地模拟第一次确认")
            first_path=first["files"]["b_mapping"];first_record=json.loads(Path(first_path).read_text(encoding="utf-8-sig"))
            second=copy.deepcopy(op);second["mapping_sha256"]=file_hash(first_path)
            item=second["items"][0];item.update(cell="B4",evidence_cells=["A4","B1"],reason="原其他货币资金期末余额为空，保留原空白")
            item["business_context"].update(range="B4:B4",header_cells=["A4","B1"],table_semantic="其他货币资金期末余额",reason="A4为其他货币资金，B1为期末余额")
            final=api.apply_external_review(api.preview_external_review(first_path,self.write_json("第二次补上下文",second)),
                self.work/"第二次上下文成果",confirmed=True,confirm_method="本地模拟第二次确认")
            loaded,snapshot=flow._load_mapping(final["files"]["b_mapping"])
            self.assertEqual(loaded["mappings"][:2],first_record["mappings"])
            self.assertEqual([m["cell"] for m in loaded["mappings"]],["B2","B3","B4"])
            self.assertEqual(len(loaded["confirmed_business_ranges"][self.sheet]),3)
            self.assertIsNone(snapshot["sheets"][0]["cells"]["B4"]["value"])
            api.validate_external_reviews(loaded,snapshot)
            engine=api._engine(loaded,self.work);engine.source_hash=snapshot["sha256"]
            snapshot["original_hash"]=loaded["original_hash"]
            kept,proof=engine._reuse_previous({"mapping_path":final["files"]["b_mapping"],"mapping_sha256":file_hash(final["files"]["b_mapping"])},
                copy.deepcopy(snapshot),{"content_hash":engine._source_identity(snapshot)},None)
            self.assertEqual([m["cell"] for m in kept["mappings"]],["B2","B3","B4"])
            self.assertFalse([m for m in proof["recheck"] if m["cell"] in {"B2","B3","B4"}])

    def test_saved_a_context_cannot_lose_word_physical_origin(self):
        api=self.api();record,mapping,op=self.missing_context_fixture();record["carrier"]="A"
        record["context"].append({"table_id":"真实物理表1","range_name":"本地测试Word表","sheet":self.sheet,
            "first_row":1,"first_column":1,"row_count":4,"column_count":2})
        mapping=self.write_json("带物理表的A原映射",record);op.update(carrier="A",mapping_sha256=file_hash(mapping))
        result=api.apply_external_review(api.preview_external_review(mapping,self.write_json("A原物理表上下文意见",op)),
            self.work/"A上下文成果",confirmed=True,confirm_method="本地模拟A确认")
        loaded,snapshot=flow._load_mapping(result["files"]["a_mapping"],complete=False)
        self.assertEqual(snapshot["sheets"][0]["word_table_ranges"][0]["range"],"A1:B4")
        for mutate in [lambda r:r["context"].pop(),lambda r:r["context"][-1].update(row_count=3)]:
            bad=copy.deepcopy(loaded);mutate(bad)
            with self.assertRaises(ValueError):flow._load_mapping(self.write_json("丢失原Word物理事实",bad),complete=False)
        shown=copy.deepcopy(snapshot);shown["sheets"][0].pop("word_table_ranges")
        with self.assertRaisesRegex(ValueError,"Word物理"):api.validate_external_reviews(loaded,shown)

    def test_correction_keeps_first_context_proof_required_after_old_mapping_is_replaced(self):
        api=self.api();record,mapping,op=self.missing_context_fixture()
        first=api.apply_external_review(api.preview_external_review(mapping,self.write_json("上下文首层意见",op)),
            self.work/"上下文首层",confirmed=True,confirm_method="本地模拟首次确认")
        first_path=first["files"]["b_mapping"];first_record=json.loads(Path(first_path).read_text(encoding="utf-8-sig"))
        old=first_record["mappings"][-1];old_opinion=old["external_review"]["opinion_path"]
        correction=copy.deepcopy(op);correction["mapping_sha256"]=file_hash(first_path)
        item=correction["items"][0];item.pop("business_context")
        item.update(replace_mapping_sha256=api._hash(old),reason="本地模拟更正期间但保留原业务总体")
        item["dimensions"]["period"]="2024-12-31"
        final=api.apply_external_review(api.preview_external_review(first_path,self.write_json("覆盖原业务格的更正",correction)),
            self.work/"上下文二层",confirmed=True,confirm_method="本地模拟第二次更正")
        path=final["files"]["b_mapping"];loaded,_=flow._load_mapping(path,complete=False)
        self.assertEqual(loaded["mappings"][-1]["dimensions"]["period"],"2024-12-31")
        changed=copy.deepcopy(loaded);changed["recognition_basis"]["report_scope"]="standalone"
        changed["context"][0]["report_scope"]="standalone"
        with self.assertRaisesRegex(ValueError,"结构化事实"):
            flow._load_mapping(self.write_json("后续更正后试改实际口径",changed),complete=False)
        check=flow._check
        def missing_first_opinion(path,*args,**kwargs):
            if str(path)==old_opinion:raise ValueError("测试模拟第一层原意见丢失")
            return check(path,*args,**kwargs)
        with patch.object(flow,"_check",side_effect=missing_first_opinion),self.assertRaisesRegex(ValueError,"第一层原意见丢失"):
            flow._load_mapping(path,complete=False)

    def test_same_meaning_approved_note_survives_external_review_and_resume(self):
        api=self.api();base=json.loads(self.gold.read_text(encoding="utf-8-sig"))
        base["note"]["meaning"]="货币资金余额构成"
        base["applicability"]={"report_scopes":[{"scope":"standalone","reason":"本地测试余额适用单户","reviewed_by":"模拟审核","reviewed_at":"2026-09-20"}]}
        peer=copy.deepcopy(base);peer.update(id="P-N009-T001-S001",scope="parent")
        peer["note"]["id"]="P-N009";peer["table"]["id"]="P-N009-T001"
        gold=self.work/"同名科目测试标准.jsonl"
        gold.write_text("\n".join(json.dumps(x,ensure_ascii=False) for x in (base,peer)),encoding="utf-8-sig")
        record=copy.deepcopy(self.record);record.update(gold_path=str(gold),gold_hash=file_hash(gold))
        for item in record["mappings"]:item.update(scope="standalone",definition_scope="consolidated")
        for areas in record["confirmed_business_ranges"].values():
            for area in areas:area["scope"]="standalone"
        mapping=self.write_json("同名科目原映射",record)
        op=copy.deepcopy(self.opinion);op.update(mapping_sha256=file_hash(mapping),gold_hash=file_hash(gold))
        op["items"][0].update(slot_id=peer["id"],scope="standalone")
        with patch("urllib.request.build_opener",side_effect=AssertionError("禁止网络")):
            preview=api.preview_external_review(mapping,self.write_json("同名科目复核意见",op))
            result=api.apply_external_review(preview,self.work/"同名科目采纳",confirmed=True,confirm_method="本地模拟审核")
            loaded,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
            chosen=next(m for m in loaded["mappings"] if m["cell"]=="B3")
            self.assertEqual((chosen["scope"],chosen["definition_scope"]),("standalone","parent"))
            self.assertEqual(loaded["confirmed_business_ranges"],record["confirmed_business_ranges"])
            engine=api._engine(loaded,self.work);engine.source_hash=snapshot["sha256"]
            basis={"content_hash":engine._source_identity(snapshot)}
            pending=copy.deepcopy(loaded);pending.pop("schema");pending.pop("carrier");pending["recognition_basis"]=basis
            evidence=self.write_json("同名科目接续证据",pending)
            kept,proof=engine._reuse_previous({**pending,"evidence_path":str(evidence),"evidence_hash":file_hash(evidence)},copy.deepcopy(snapshot),basis,None)
        self.assertEqual([m["cell"] for m in kept["mappings"]],["B2","B3"])
        self.assertFalse([m for m in proof["recheck"] if m["cell"]=="B3"])
        self.assertEqual(self.before,{p:file_hash(p) for p in self.before})

    def test_export_template_binds_source_and_contains_only_pending(self):
        api=self.api();p=self.work/"待填写意见.json";api.export_review_template(self.mapping,p)
        doc=json.loads(p.read_text(encoding="utf-8-sig"));self.assertEqual(doc["mapping_sha256"],file_hash(self.mapping))
        self.assertEqual([i["cell"] for i in doc["items"]],["B3","B4"]);self.assertFalse(doc["reviewer"]["name"])
        self.assertTrue(all(not i.get("slot_id") and not i.get("reviewed") for i in doc["items"]))
        with self.assertRaises(FileExistsError):api.export_review_template(self.mapping,p)
        with self.assertRaises(ValueError):api.preview_external_review(self.mapping,p)
    def test_explicit_correction_preserves_prior_mapping_and_can_mix_pending(self):
        api=self.api();op=copy.deepcopy(self.opinion)
        corrected=copy.deepcopy(op["items"][0]);corrected.update(cell="B2",evidence_cells=["A2","B1"],replace_mapping_sha256=api._hash(self.original),reason="本地模拟高级AI更正：原期间错误，应使用原表注明的上一年日期")
        corrected["dimensions"]["period"]="2024-12-31";op["items"].append(corrected)
        opinion=self.write_json("显式更正及补齐未决",op)
        preview=api.preview_external_review(self.mapping,opinion)
        self.assertEqual(preview["remaining_pending"],1)
        self.assertEqual(preview["changes"][1]["previous_mapping"],self.original)
        with patch("urllib.request.build_opener",side_effect=AssertionError("禁止网络")):
            result=api.apply_external_review(preview,self.work/"更正成果",confirmed=True,confirm_method="本地模拟明确更正")
            record,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
            self.assertEqual(len(record["mappings"]),2)
            self.assertEqual(next(m for m in record["mappings"] if m["cell"]=="B2")["dimensions"]["period"],"2024-12-31")
            self.assertEqual([m["cell"] for m in record["unresolved"]],["B4"])
            api.validate_external_reviews(record,snapshot)
            proof=record["mappings"][0]["external_review"]
            saved=json.loads(Path(proof["mapping_path"]).read_text(encoding="utf-8-sig"))
            self.assertEqual(saved["mappings"],[self.original])
            tampered=copy.deepcopy(record);tampered["mappings"][0]["dimensions"]["period"]="1999-12-31"
            with self.assertRaises(ValueError):api.validate_external_reviews(tampered,snapshot)
        self.assertEqual(self.before,{p:file_hash(p) for p in self.before})
    def test_correction_requires_exact_old_mapping_and_cannot_reclassify_excluded(self):
        api=self.api()
        for cell,digest in [("B2",None),("B2","0"*64),("B3",api._hash(self.original)),("A2",api._hash(self.original))]:
            op=copy.deepcopy(self.opinion);op["items"][0].update(cell=cell)
            if digest is not None:op["items"][0]["replace_mapping_sha256"]=digest
            with self.subTest(cell=cell,digest=digest),self.assertRaises(ValueError):
                api.preview_external_review(self.mapping,self.write_json("错误更正意图",op))
    def test_export_can_explicitly_include_confirmed_cells_without_claiming_review(self):
        api=self.api();p=self.work/"含已确认格的复核格式.json"
        api.export_review_template(self.mapping,p,include_confirmed=True)
        doc=json.loads(p.read_text(encoding="utf-8-sig"))
        item=next(x for x in doc["items"] if x["cell"]=="B2")
        self.assertEqual(item["replace_mapping_sha256"],api._hash(self.original))
        self.assertEqual(doc["confirmed_mapping_options"],[self.original])
        self.assertFalse(item["slot_id"]);self.assertFalse(doc["reviewer"]["name"])
        self.assertEqual([x["cell"] for x in doc["items"]],["B3","B4","B2"])
    def test_correction_binds_disk_mapping_when_legacy_dimensions_normalize(self):
        api=self.api();record=copy.deepcopy(self.record)
        record["mappings"][0]["dimensions"].update(currency="人民币",scale="1")
        old_mapping=record["mappings"][0]
        legacy=self.write_json("维度尚未规范的旧映射",record);original_bytes=legacy.read_bytes()
        loaded,_=flow._load_mapping(legacy,complete=False)
        self.assertNotEqual(loaded["mappings"][0],old_mapping)
        template=self.work/"旧格式映射更正格式.json"
        api.export_review_template(legacy,template,include_confirmed=True)
        exported=json.loads(template.read_text(encoding="utf-8-sig"))
        binding=next(x for x in exported["items"] if x["cell"]=="B2")["replace_mapping_sha256"]
        self.assertEqual(binding,api._hash(old_mapping))
        self.assertEqual(exported["confirmed_mapping_options"],[old_mapping])
        opinion=copy.deepcopy(self.opinion);opinion["mapping_sha256"]=file_hash(legacy)
        corrected=copy.deepcopy(opinion["items"][0]);corrected.update(cell="B2",evidence_cells=["A2","B1"],
            replace_mapping_sha256=binding,reason="本地模拟审核更正旧格式记录的实际期间")
        corrected["dimensions"]["period"]="2024-12-31";opinion["items"]=[corrected]
        preview=api.preview_external_review(legacy,self.write_json("旧格式映射更正意见",opinion))
        self.assertEqual(preview["changes"][0]["previous_mapping"],old_mapping)
        self.assertEqual(preview["remaining_pending"],2)
        with patch("urllib.request.build_opener",side_effect=AssertionError("禁止网络")):
            result=api.apply_external_review(preview,self.work/"旧格式更正成果",confirmed=True,confirm_method="本地模拟明确更正")
            saved,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
            api.validate_external_reviews(saved,snapshot)
        mapping=saved["mappings"][0]
        self.assertEqual(mapping["dimensions"],{"currency":"CNY","period":"2024-12-31","scale":1,"unit":"元"})
        self.assertEqual(Path(mapping["external_review"]["mapping_path"]).read_bytes(),original_bytes)
        self.assertEqual(legacy.read_bytes(),original_bytes)
        self.assertEqual(self.before,{p:file_hash(p) for p in self.before})
    def test_preview_and_partial_apply_preserve_old_data_and_real_origin(self):
        api=self.api()
        with patch("urllib.request.build_opener",side_effect=AssertionError("禁止网络")):
            preview=api.preview_external_review(self.mapping,self.opinion_path)
            self.assertEqual(preview["count"],1);self.assertEqual(preview["remaining_pending"],1)
            with self.assertRaises(ValueError):api.apply_external_review(preview,self.work/"拒绝未确认")
            result=api.apply_external_review(preview,self.work/"成果",confirmed=True,confirm_method="本地模拟确认导入")
            self.assertEqual(result["status"],"partial");self.assertEqual(result["accepted_count"],1)
            output,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
            self.assertEqual(output["mappings"][0],self.original);self.assertEqual(output["unresolved"][0]["cell"],"B4")
            item=output["mappings"][1];self.assertTrue(item["reviewed"]);self.assertEqual(item["review_method"],"external_semantic_review")
            self.assertNotIn("independent_review",item);self.assertNotIn("model_rounds",item)
            api.validate_external_reviews(output,snapshot)
            self.assertTrue(Path(result["files"]["b_review"]).is_file());self.assertTrue(Path(result["files"]["b_standard_review"]).is_file())
        self.assertEqual(self.before,{p:file_hash(p) for p in self.before})
    def test_complete_apply_keeps_source_blank(self):
        api=self.api();op=copy.deepcopy(self.opinion);second=copy.deepcopy(op["items"][0]);second.update(cell="B4",evidence_cells=["A4","B1"]);op["items"].append(second)
        preview=api.preview_external_review(self.mapping,self.write_json("完整意见",op))
        result=api.apply_external_review(preview,self.work/"完整成果",confirmed=True,confirm_method="本地模拟人工确认")
        self.assertEqual(result["status"],"complete");record,snapshot=flow._load_mapping(result["files"]["b_mapping"])
        self.assertIsNone(snapshot["sheets"][0]["cells"]["B4"]["value"]);self.assertEqual(len(record["mappings"]),3)
    def test_completed_review_clears_previous_pending_material_and_updates_structure(self):
        api=self.api();old_material=self.write_json("旧待决材料",{"pending":copy.deepcopy(self.record["unresolved"])})
        record=copy.deepcopy(self.record);record.update(standard_review_path=str(old_material),structure_path="旧结构.json")
        mapping=self.write_json("含旧材料映射",record)
        op=copy.deepcopy(self.opinion);op["mapping_sha256"]=file_hash(mapping)
        second=copy.deepcopy(op["items"][0]);second.update(cell="B4",evidence_cells=["A4","B1"]);op["items"].append(second)
        opinion=self.write_json("解决全部待决意见",op)
        result=api.apply_external_review(api.preview_external_review(mapping,opinion),self.work/"完成材料成果",confirmed=True,confirm_method="本地模拟确认")
        final=json.loads(Path(result["files"]["b_mapping"]).read_text(encoding="utf-8-sig"))
        self.assertTrue(final["complete"]);self.assertEqual(final["unresolved"],[])
        self.assertNotIn("standard_review_path",final)
        from 附注更新.界面 import _mapping_sources
        self.assertEqual(_mapping_sources(result["files"]["b_mapping"],"B")["b_standard_review"],"")
        self.assertTrue(Path(final["structure_path"]).is_file())
        self.assertEqual(Path(final["structure_path"]).parent,Path(result["files"]["b_mapping"]).parent)
        self.assertEqual(len(json.loads(old_material.read_text(encoding="utf-8-sig"))["pending"]),2)

    def test_explicit_existing_range_resolves_overlap_and_preserves_review_evidence(self):
        api=self.api();record=copy.deepcopy(self.record)
        area=copy.deepcopy(record["confirmed_business_ranges"][self.sheet][0])
        area["table_range"]="A1:B3";record["confirmed_business_ranges"][self.sheet].append(area)
        record["unresolved"][0]["table_range"]="A1:B3"
        mapping=self.write_json("同一原表内两个重叠范围",record)
        op=copy.deepcopy(self.opinion);op["mapping_sha256"]=file_hash(mapping)
        with self.assertRaisesRegex(ValueError,"唯一"):
            api.preview_external_review(mapping,self.write_json("未明确选择范围",op))
        op["items"][0].update(table_range="A1:B4",reason="本地模拟审核：原始完整表含四行，明确选择既有A1:B4范围；银行存款余额和期间文字一致")
        opinion=self.write_json("明确选择完整原范围",op)
        with patch("urllib.request.build_opener",side_effect=AssertionError("禁止网络")):
            preview=api.preview_external_review(mapping,opinion)
            self.assertEqual(preview["changes"][0]["mapping"]["table_range"],"A1:B4")
            result=api.apply_external_review(preview,self.work/"选范围成果",confirmed=True,confirm_method="本地模拟确认既有范围")
            loaded,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
            self.assertEqual(loaded["confirmed_business_ranges"],record["confirmed_business_ranges"])
            self.assertEqual(loaded["mappings"][0],self.original)
            accepted=loaded["mappings"][-1]
            self.assertEqual(accepted["reason"],op["items"][0]["reason"])
            self.assertNotIn("model_rounds",accepted);self.assertNotIn("independent_review",accepted)
            saved=json.loads(Path(accepted["external_review"]["opinion_path"]).read_text(encoding="utf-8-sig"))
            self.assertEqual(saved,op);api.validate_external_reviews(loaded,snapshot)
            engine=api._engine(loaded,self.work);snapshot["original_hash"]=loaded.get("original_hash") or loaded["source_hash"]
            engine.source_hash=snapshot["sha256"]
            previous={"mapping_path":result["files"]["b_mapping"],"mapping_sha256":file_hash(result["files"]["b_mapping"])}
            kept,_=engine._reuse_previous(previous,copy.deepcopy(snapshot),{"content_hash":engine._source_identity(snapshot)},None)
            self.assertIn("B3",[m["cell"] for m in kept["mappings"]])
        self.assertEqual(self.before,{p:file_hash(p) for p in self.before})

    def test_explicit_range_cannot_invent_ownership_or_bypass_original_guards(self):
        api=self.api();record,snapshot=flow._load_mapping(self.mapping,complete=False)
        engine=api._engine(record,self.work);sheet=snapshot["sheets"][0]
        before=copy.deepcopy(snapshot)
        for selected in ["A1:B2","A2:B3","A1:B99","",None,123]:
            op=copy.deepcopy(self.opinion);op["items"][0]["table_range"]=selected
            with self.subTest(selected=selected),self.assertRaises(ValueError):
                api._checked_items(record,snapshot,op,file_hash(self.mapping),engine)
        op=copy.deepcopy(self.opinion);op["items"][0]["table_range"]="A1:B4"
        ambiguous=copy.deepcopy(snapshot);extra=copy.deepcopy(ambiguous["sheets"][0]["confirmed_business_ranges"][0]);extra["table_semantic"]="不同原业务说明"
        ambiguous["sheets"][0]["confirmed_business_ranges"].append(extra)
        with self.assertRaises(ValueError):api._checked_items(record,ambiguous,op,file_hash(self.mapping),engine)
        mismatched=copy.deepcopy(record);mismatched["unresolved"][0]["table_range"]="A1:B3"
        with self.assertRaisesRegex(ValueError,"范围不一致"):
            api._checked_items(mismatched,snapshot,self.opinion,file_hash(self.mapping),engine)
        divided=copy.deepcopy(snapshot);divided["sheets"][0]["word_table_ranges"]=[{"range":"A1:B2"},{"range":"A3:B4"}]
        with self.assertRaisesRegex(ValueError,"Word实际表格边界"):
            api._checked_items(record,divided,op,file_hash(self.mapping),engine)
        for change in [lambda x:x.update(evidence_cells=["B3"]),lambda x:x.update(scope="standalone"),lambda x:x.update(slot_id="未知槽位"),lambda x:x["dimensions"].pop("period")]:
            bad=copy.deepcopy(op);change(bad["items"][0])
            with self.subTest(item=bad["items"][0]),self.assertRaises(ValueError):
                api._checked_items(record,snapshot,bad,file_hash(self.mapping),engine)
        self.assertEqual(snapshot,before)

    def test_same_range_basis_selects_one_real_proof_and_export_lists_options(self):
        api=self.api();record=copy.deepcopy(self.record)
        first=record["confirmed_business_ranges"][self.sheet][0]
        second=copy.deepcopy(first);second.update(table_semantic="逐行复核原表后的货币资金余额表",header_cells=["A1","B1"])
        record["confirmed_business_ranges"][self.sheet].append(second)
        mapping=self.write_json("相同地址不同原业务依据",record)
        op=copy.deepcopy(self.opinion);op["mapping_sha256"]=file_hash(mapping);op["items"][0]["table_range"]="A1:B4"
        with self.assertRaisesRegex(ValueError,"唯一"):
            api.preview_external_review(mapping,self.write_json("只选择相同地址",op))
        template=self.work/"含原依据选择的模板.json";api.export_review_template(mapping,template)
        exported=json.loads(template.read_text(encoding="utf-8-sig"))
        self.assertEqual(exported["business_range_options"],[{"sheet":self.sheet,"range":a["table_range"],"table_semantic":a["table_semantic"],"header_cells":a["header_cells"],"note_ids":a["note_ids"],"scope":a["scope"],"basis_hash":api._hash(a)} for a in [first,second]])
        op["items"][0]["table_basis_hash"]=api._hash(second)
        opinion=self.write_json("明确选择第二份真实依据",op)
        with patch("urllib.request.build_opener",side_effect=AssertionError("禁止网络")):
            preview=api.preview_external_review(mapping,opinion)
            checked=preview["changes"][0]["mapping"]
            self.assertEqual(checked["table_semantic"],second["table_semantic"])
            self.assertEqual(checked["table_basis_hash"],api._hash(second))
            result=api.apply_external_review(preview,self.work/"按既有依据成果",confirmed=True,confirm_method="本地模拟明确选择第二原依据")
            loaded,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
            self.assertEqual(loaded["confirmed_business_ranges"],record["confirmed_business_ranges"])
            api.validate_external_reviews(loaded,snapshot)
            bad=copy.deepcopy(loaded);bad["mappings"][-1]["table_basis_hash"]=api._hash(first)
            with self.assertRaises(ValueError):api.validate_external_reviews(bad,snapshot)
            # 预览后选择的原依据变化，旧哈希必须失效，不能沿用原确认。
            changed=copy.deepcopy(snapshot);changed["sheets"][0]["confirmed_business_ranges"][1]["header_cells"]=["B1"]
            with self.assertRaises(ValueError):api._checked_items(record,changed,op,file_hash(mapping),api._engine(record,self.work))
        self.assertEqual(self.before,{p:file_hash(p) for p in self.before})

    def test_table_basis_hash_requires_matching_range_and_same_sheet_proof(self):
        api=self.api();record,snapshot=flow._load_mapping(self.mapping,complete=False)
        engine=api._engine(record,self.work);area=snapshot["sheets"][0]["confirmed_business_ranges"][0]
        actual=api._hash(area);other=copy.deepcopy(area);other["table_semantic"]="另一张表的原业务依据"
        snapshot["sheets"].append({"name":"另一张表","cells":{},"confirmed_business_ranges":[other]})
        for extra in [{"table_basis_hash":actual},{"table_range":"A1:B4","table_basis_hash":"0"*64},{"table_range":"A1:B4","table_basis_hash":None},{"table_range":"A1:B4","table_basis_hash":api._hash(other)},{"table_range":"A1:B3","table_basis_hash":actual}]:
            op=copy.deepcopy(self.opinion);op["items"][0].update(extra)
            with self.subTest(extra=extra),self.assertRaises(ValueError):
                api._checked_items(record,snapshot,op,file_hash(self.mapping),engine)

    def test_rejects_stale_binding_unreviewed_or_invalid_suggestions(self):
        api=self.api()
        changes=[lambda x:x.update(mapping_sha256="0"*64),lambda x:x.update(source_hash="0"*64),lambda x:x.update(gold_hash="0"*64),lambda x:x["reviewer"].update(name=""),lambda x:x["reviewer"].update(method="model_two_rounds"),lambda x:x["reviewer"].update(reviewed_at="不是时间"),lambda x:x["items"][0].update(cell="B2"),lambda x:x["items"].append(copy.deepcopy(x["items"][0])),lambda x:x["items"][0]["dimensions"].pop("period"),lambda x:x["items"][0].update(scope="standalone"),lambda x:x["items"][0].update(value_type="percentage"),lambda x:x["items"][0].update(slot_id="未知ID"),lambda x:x["items"][0].update(evidence_cells=["B3"]),lambda x:x["items"][0].update(evidence_cells=["A99"]),lambda x:x["items"][0].update(decision="excluded"),lambda x:x["items"][0].update(reviewed=True)]
        for change in changes:
            op=copy.deepcopy(self.opinion);change(op)
            with self.subTest(op=op),self.assertRaises(ValueError):api.preview_external_review(self.mapping,self.write_json("无效意见",op))
    def test_apply_rejects_modified_preview(self):
        api=self.api();preview=api.preview_external_review(self.mapping,self.opinion_path);preview["changes"][0]["mapping"]["dimensions"]["period"]="2024-12-31"
        with self.assertRaises(ValueError):api.apply_external_review(preview,self.work/"拒绝篡改",confirmed=True,confirm_method="本地模拟确认")
    def test_external_proof_rejects_missing_changed_or_forged_evidence(self):
        api=self.api();result=api.apply_external_review(api.preview_external_review(self.mapping,self.opinion_path),self.work/"证明成果",confirmed=True,confirm_method="本地模拟确认")
        record,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
        for change in [lambda m:m.pop("external_review"),lambda m:m["external_review"].update(opinion_sha256="0"*64),lambda m:m.update(reviewed=False),lambda m:m["dimensions"].update(period="2024-12-31"),lambda m:m["external_review"].update(confirm_method="伪造确认方式")]:
            bad=copy.deepcopy(record);change(bad["mappings"][-1])
            with self.subTest(mapping=bad["mappings"][-1]),self.assertRaises(ValueError):api.validate_external_reviews(bad,snapshot)
    def test_formal_and_nonformal_resume_recheck_external_proof(self):
        api=self.api();result=api.apply_external_review(api.preview_external_review(self.mapping,self.opinion_path),self.work/"接续成果",confirmed=True,confirm_method="本地模拟确认")
        record,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
        tampered=copy.deepcopy(record);tampered["mappings"][-1]["dimensions"]["period"]="2024-12-31"
        with self.assertRaises(ValueError):flow._load_mapping(self.write_json("篡改正式映射",tampered),complete=False)
        from 附注更新.语义 import SemanticEngine
        engine=SemanticEngine({},str(self.gold),str(self.work));engine.source_hash=snapshot["sha256"]
        basis={"content_hash":engine._source_identity(snapshot)}
        nonformal=copy.deepcopy(record);nonformal.pop("schema");nonformal.pop("carrier");nonformal["recognition_basis"]=basis
        evidence=self.write_json("实际非正式结果",nonformal)
        previous={**nonformal,"evidence_path":str(evidence),"evidence_hash":file_hash(evidence)}
        kept,proof=engine._reuse_previous(previous,copy.deepcopy(snapshot),basis,None)
        self.assertEqual([m["cell"] for m in kept["mappings"]],["B2","B3"])
        self.assertFalse([m for m in proof["recheck"] if m["cell"]=="B3"])
        bad=copy.deepcopy(nonformal);bad["mappings"][-1]["external_review"]["confirm_method"]="假来源"
        evidence=self.write_json("损坏非正式结果",bad)
        with self.assertRaises(ValueError):engine._reuse_previous({**bad,"evidence_path":str(evidence),"evidence_hash":file_hash(evidence)},copy.deepcopy(snapshot),basis,None)

    def test_selected_review_gold_can_add_slot_without_relabeling_old_mapping(self):
        api=self.api();original=json.loads(self.gold.read_text(encoding="utf-8-sig"));added=copy.deepcopy(original)
        added["id"]="C-N087-T001-S002";added["slot"]["name"]="银行存款余额"
        selected=self.work/"实际审核标准.jsonl";selected.write_text("\n".join(json.dumps(x,ensure_ascii=False) for x in [original,added]),encoding="utf-8-sig")
        template=self.work/"当前标准意见模板.json";api.export_review_template(self.mapping,template,gold_path=selected)
        self.assertEqual(json.loads(template.read_text(encoding="utf-8-sig"))["gold_hash"],file_hash(selected))
        op=copy.deepcopy(self.opinion);op["gold_hash"]=file_hash(selected);op["items"][0]["slot_id"]=added["id"]
        opinion=self.write_json("使用当前标准意见",op)
        preview=api.preview_external_review(self.mapping,opinion,gold_path=selected)
        self.assertEqual(preview["gold_path"],str(selected));self.assertEqual(preview["gold_hash"],file_hash(selected))
        result=api.apply_external_review(preview,self.work/"新标准成果",confirmed=True,confirm_method="本地模拟确认新标准")
        record,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
        self.assertEqual(record["gold_hash"],file_hash(selected));self.assertEqual(record["mappings"][0],self.original)
        proof=record["mappings"][-1]["external_review"]
        self.assertEqual(proof["gold_hash"],file_hash(selected));self.assertEqual(json.loads(Path(proof["mapping_path"]).read_text(encoding="utf-8-sig"))["gold_hash"],file_hash(self.gold))
        api.validate_external_reviews(record,snapshot)
        changed=copy.deepcopy(original);changed["slot"]["name"]="已确认指标被改"
        bad=self.work/"不能兼容标准.jsonl";bad.write_text(json.dumps(changed,ensure_ascii=False),encoding="utf-8-sig")
        with self.assertRaises(ValueError):api.export_review_template(self.mapping,self.work/"拒绝模板.json",gold_path=bad)

    def test_external_proof_allows_new_gold_only_when_used_definition_unchanged(self):
        api=self.api();result=api.apply_external_review(api.preview_external_review(self.mapping,self.opinion_path),self.work/"兼容成果",confirmed=True,confirm_method="本地模拟确认")
        record,snapshot=flow._load_mapping(result["files"]["b_mapping"],complete=False)
        original=json.loads(self.gold.read_text(encoding="utf-8-sig"));second=copy.deepcopy(original);second["id"]="C-N087-T001-S002"
        p=self.work/"增加其它定义.jsonl";p.write_text("\n".join(json.dumps(x,ensure_ascii=False) for x in [original,second]),encoding="utf-8-sig")
        record.update(gold_path=str(p),gold_hash=file_hash(p));api.validate_external_reviews(record,snapshot)
        changed=copy.deepcopy(original);changed["slot"]["name"]="别的指标";q=self.work/"所用定义改变.jsonl";q.write_text(json.dumps(changed,ensure_ascii=False),encoding="utf-8-sig");record.update(gold_path=str(q),gold_hash=file_hash(q))
        with self.assertRaises(ValueError):api.validate_external_reviews(record,snapshot)

if __name__=="__main__":unittest.main()
