"""真实金标准字段的小型本地回执验收，不调用在线模型。"""
import copy
import hashlib
import json
import unittest
import uuid
from pathlib import Path
from unittest.mock import patch
from 附注更新.语义补充 import review_missing_semantics, _similar, _same_indicator, _proposal
from 附注更新.金标准 import read_standard
from 附注更新.语义 import SemanticCancelled, SemanticEngine
ROOT=Path(__file__).resolve().parents[1]


def base_slot():
    return {"id":"C-N001-T001-S001","scope":"consolidated","status":"active",
        "note":{"id":"C-N001","name":"货币资金","meaning":"货币资金披露"},
        "table":{"id":"C-N001-T001","name":"资金构成","meaning":"资金构成和受限情况"},
        "slot":{"name":"银行存款 / 期末余额","role":"measure","meaning":"期末银行存款账面余额","display_order":1},
        "row_path":["银行存款"],"column_path":["期末余额"],"value_type":"monetary",
        "dimensions":[{"name":"currency","required":True},{"name":"period","required":True,"role":"closing","type":"instant"},{"name":"scale","required":True},{"name":"unit","required":True}],
        "blank_policy":{"default_interpretation":"unknown","raw_blank_preserved":True},
        "aggregation":{"role":"component","parent_slot_id":None,"parent_variants":[],"component_slot_ids":[],"component_variants":[]},
        "calculation":{"expression":None,"component_slot_ids":[]},
        "aliases":{"row_labels":["银行存款"],"column_labels":["期末余额"]},
        "applicability":{"industries":["general"],"optional":True},"sources":[],
        "introduced_version":"v1","last_modified_version":"v1","superseded_by":None,"扩展资料":{"保留":"原始完整字段"}}


def proposal():
    d=base_slot()
    definition={k:copy.deepcopy(d[k]) for k in ("scope","slot","row_path","column_path","value_type","dimensions","blank_policy","aggregation","calculation","aliases","applicability")}
    definition.update(note_id="C-N001",table_id="C-N001-T001",slot={"name":"数字人民币钱包 / 期末受限余额","meaning":"数字人民币钱包中使用受到限制的期末金额","role":"measure"},
        row_path=["数字人民币钱包"],column_path=["期末受限余额"],aliases={"row_labels":["数字人民币钱包"],"column_labels":["期末受限余额"]})
    return {"cell":"B3","decision":"new_slot","base_slot_id":"C-N001-T001-S001","definition":definition,
        "evidence_cells":["A3","B2"],"reason":"原始表头明确披露独立受限余额指标",
        "existing_not_sufficient_reason":"已有银行存款总额不是钱包受限金额，也没有对应开放指标", "standalone_applicable":True,
        "scope_reason":"该指标为货币资金自身受限余额，不依赖合并抵销和少数股东身份"}


def fixture():
    values={"A1":"单体财务报表货币资金", "A2":"项目", "B2":"期末受限余额", "A3":"数字人民币钱包", "B3":123.45}
    table={"table_range":"A1:B3","note_ids":["C-N001"],"scope":"standalone","table_semantic":"货币资金构成及受限情况","header_cells":["A1","A2","B2"],"source_hash":"source"}
    snapshot={"sha256":"source","context":[],"sheets":[{"name":"资料","max_row":3,"max_column":2,"merges":[],"confirmed_business_ranges":[table],"cells":{a:{"row":int(a[1:]),"column":ord(a[0])-64,"value":v,"formula":None} for a,v in values.items()}}]}
    result={"mappings":[],"unresolved":[{"sheet":"资料","cell":"B3","reason":"没有确定的金标准定义"}],"confirmed_business_ranges":{"资料":[table]}}
    return snapshot,result


def cell_fixture():
    """同批两个不连续原格，中间格不是它们的业务或文字证据。"""
    snapshot,result=fixture();sheet=snapshot["sheets"][0]
    sheet["max_row"]=8;sheet["max_column"]=4
    sheet["cells"].update({"A8":{"row":8,"column":1,"value":"银行存款","formula":None},
                           "D8":{"row":8,"column":4,"value":678,"formula":None},
                           "C5":{"row":5,"column":3,"value":"中间无关业务","formula":None}})
    proof=sheet["confirmed_business_ranges"][0]
    proofs=[{**copy.deepcopy(proof),"table_range":f"{address}:{address}","scope":"consolidated",
             "header_cells":["A1","B2",label],"table_semantic":f"{label}所述余额",
             "layout_batch_id":"同一原包","layout_review":{"rounds":[1,2]}}
            for address,label in (("B3","A3"),("D8","A8"))]
    sheet["confirmed_business_ranges"]=proofs
    result["confirmed_business_ranges"]={"资料":copy.deepcopy(proofs)}
    result["unresolved"]=[{"sheet":"资料","cell":a,"reason":"待复核"} for a in ("B3","D8")]
    return snapshot,result


class Stub(SemanticEngine):
    def __init__(self,gold,mode="new"):
        self.gold_hash=hashlib.sha256(gold.read_bytes()).hexdigest();self.mode=mode;self.calls=[];self.settings={"model":"local-test"}
        self.slots={key:row for key,row in read_standard(gold).items() if row.get("status")=="active"}
        self.notes={row["note"]["id"]:copy.deepcopy(row["note"]) for row in self.slots.values()}
        self.report_scope=""
    def _check_cancel(self):pass
    def _request(self,system,payload,validator):
        self.calls.append(copy.deepcopy(payload))
        if self.mode=="cancel":raise SemanticCancelled("停止")
        if payload["task"]=="tables":return validator({"table_ids":[t["id"] for t in payload["catalogue"]]})
        if payload["task"]=="missing_semantics_existing_check":
            if self.mode=="failed_check":raise ValueError("失败")
            ids=[x["id"] for x in payload["existing_slots"]]
            return validator({"checked_ids":ids,"existing_ids":ids[:1] if self.mode=="synonym_check" else [],"open_dimension_ids":[],"uncertain":False,"reason":"逐定义比较后未见同义指标"})
        if payload["task"]=="missing_semantics_proposal_review":
            decisions=[]
            for item in payload["proposals"]:
                accepted=self.mode!="disagree"
                decisions.append({"cell":item["cell"],"accepted":accepted,"reason":"复核原表和候选定义后确认含义一致" if accepted else "候选定义与原格不是同一指标","evidence_cells":["Z99"] if self.mode=="bad_review_evidence" else ["A3","B2"]})
            return validator({"decisions":decisions})
        item=proposal()
        if self.mode=="existing":item={"cell":"B3","decision":"existing","slot_id":"C-N001-T001-S001","dimensions":{"disclosure_item":"数字人民币钱包"},"evidence_cells":["A3","B2"],"reason":"现有开放指标已能表达"}
        if self.mode=="paraphrase" and payload["round"]==2:item["definition"]["slot"]["meaning"]="钱包中使用受限制的期末资金金额"
        if self.mode=="disagree" and payload["round"]==2:item["definition"]["slot"]["name"]="其他不同指标"
        if self.mode=="new_percentage":
            item["definition"].update(value_type="percentage",dimensions=[{"name":"period","required":True,"role":"closing","type":"instant"}],
                slot={"name":"数字人民币钱包 / 受限比例","meaning":"受限钱包金额占钱包余额的比例","role":"measure"},
                column_path=["受限比例"],aliases={"row_labels":["数字人民币钱包"],"column_labels":["受限比例"]})
        if self.mode=="new_duration":
            item["definition"]["slot"].update(name="数字人民币钱包 / 本期受限资金变动",meaning="本期内钱包受限资金的净增加金额")
            item["definition"]["column_path"]=["本期受限资金变动"];item["definition"]["aliases"]["column_labels"]=["本期受限资金变动"]
            for dim in item["definition"]["dimensions"]:
                if dim["name"]=="period":dim.update(role="current",type="duration")
        if self.mode=="open_period":
            item["definition"]["slot"].update(name="数字人民币钱包 / 受限余额",meaning="数字人民币钱包中在实际余额日期使用受限的金额")
            item["definition"]["column_path"]=["受限余额"];item["definition"]["aliases"]["column_labels"]=["受限余额"]
            for dim in item["definition"]["dimensions"]:
                if dim["name"]=="period":dim.pop("role");dim["open"]=True
        if self.mode=="dimension_only":
            base=base_slot()
            for key in ("slot","row_path","column_path","aliases"):item["definition"][key]=copy.deepcopy(base[key])
            item["definition"]["slot"].update(name="银行存款 / 期初余额",meaning="期初银行存款账面余额")
            item["definition"]["column_path"]=["期初余额"];item["definition"]["aliases"]["column_labels"]=["期初余额"]
            for dim in item["definition"]["dimensions"]:
                if dim["name"]=="period":dim["role"]="opening"
        if self.mode=="metric_escape":item["definition"]["dimensions"].append({"name":"metric","required":True,"open":True})
        if self.mode=="fixed_object":item["definition"]["dimensions"].append({"name":"counterparty","required":True,"open":False,"value":"客户甲"})
        if self.mode=="invalid_age":item["definition"]["dimensions"].append({"name":"age_bucket","required":True,"lower":2,"upper":1,"lower_inclusive":True,"upper_inclusive":True})
        if self.mode=="invalid_period":
            for dim in item["definition"]["dimensions"]:
                if dim["name"]=="period":dim.update(role="closing",type="duration")
        if self.mode=="bad_evidence":item["evidence_cells"]=["Z99"]
        if self.mode=="missing_dimension":item["definition"]["dimensions"]=[]
        if self.mode=="wrong_table":item["definition"]["table_id"]="C-N001-T099"
        if self.mode=="scope_unreviewed":item["standalone_applicable"]=False
        if self.mode=="bad_reference":item["definition"]["aggregation"]["parent_slot_id"]="C-N999-T999-S999"
        if self.mode=="exact_duplicate":
            d=base_slot();item["definition"].update(slot=copy.deepcopy(d["slot"]),row_path=d["row_path"],column_path=d["column_path"],aliases=d["aliases"])
        if self.mode=="open_object":
            item["definition"].update(slot={"name":"客户甲 / 期末余额","meaning":"客户甲期末余额","role":"measure"},row_path=["客户甲"],column_path=["期末余额"],aliases={"row_labels":["客户甲"],"column_labels":["期末余额"]})
        decisions=[]
        for address in payload["candidate_cells"]:
            current=copy.deepcopy(item);current["cell"]=address
            if self.mode=="conflicting_proposals" and address=="B4":current["definition"]["slot"]["meaning"]="同名指标却有不同业务解释"
            if self.mode=="same_indicator_period_pair" and address=="B4":
                current["definition"]["slot"]["name"]="数字人民币钱包 / 期初受限余额"
                current["definition"]["column_path"]=["期初受限余额"]
                current["definition"]["aliases"]["column_labels"]=["期初受限余额"]
                for dim in current["definition"]["dimensions"]:
                    if dim["name"]=="period":dim["role"]="opening"
            if self.mode=="bad_role":current["definition"]["aggregation"]["role"]="无依据关系"
            decisions.append(current)
        return validator({"decisions":decisions})


class SemanticExtensionTests(unittest.TestCase):
    def setUp(self):
        self.work=ROOT/"测试结果"/("语义补充_"+uuid.uuid4().hex[:10]);self.work.mkdir(parents=True)
        self.gold=self.work/"完整金标准.jsonl"
        retired=copy.deepcopy(base_slot());retired.update(id="C-N001-T001-S099",status="deprecated")
        self.write_gold([base_slot(),retired])
    def write_gold(self,rows):
        if self.gold.exists(): self.gold=self.work/("完整金标准_"+uuid.uuid4().hex[:8]+".jsonl")
        self.gold.write_text("\n".join(json.dumps(r,ensure_ascii=False) for r in rows)+"\n",encoding="utf-8-sig")
    def run_review(self,mode="new",snapshot=None,result=None):
        s,r=fixture();s=s if snapshot is None else snapshot;r=r if result is None else result
        engine=Stub(self.gold,mode)
        with patch("附注更新.语义补充.ROOT",self.work):
            value=review_missing_semantics(engine,s,r,self.gold,self.work/"本次审阅")
        return value,engine
    def _review_cell_proofs(self,snapshot=None,result=None,new=False,bad_evidence=None):
        snapshot,result=cell_fixture() if snapshot is None else (snapshot,result)
        class CellStub(Stub):
            def _request(engine,system,payload,validator):
                engine.calls.append(copy.deepcopy(payload))
                task=payload["task"]
                if task=="missing_semantics_existing_check":
                    return validator({"checked_ids":[s["id"] for s in payload["existing_slots"]],
                        "existing_ids":[],"open_dimension_ids":[],"uncertain":False,"reason":"完整目录核对无同义"})
                decisions=[]
                for address in payload["candidate_cells"]:
                    header="A3" if address=="B3" else "A8"
                    cited=[header,"B2"]
                    if bad_evidence and bad_evidence[0]==task and address=="B3":cited=[bad_evidence[1]]
                    if task=="missing_semantics_proposal_review":
                        item={"cell":address,"accepted":True,"reason":"原格文字和完整定义一致","evidence_cells":cited}
                    elif new:
                        item={**proposal(),"cell":address,"evidence_cells":cited}
                    else:
                        item={"cell":address,"decision":"existing","slot_id":"C-N001-T001-S001","scope":"consolidated",
                            "dimensions":{"currency":"人民币","period":{"type":"instant","date":"2025-12-31"},"unit":"元","scale":"1"},
                            "value_type":"monetary","semantic_field":"银行存款 / 期末余额",
                            "reason":"原表银行存款期末余额及币种单位明确","evidence_cells":cited}
                    decisions.append(item)
                return validator({"decisions":decisions})
        engine=CellStub(self.gold)
        with patch("附注更新.语义补充.ROOT",self.work):
            value=review_missing_semantics(engine,snapshot,result,self.gold,self.work/"逐格合批")
        return value,engine

    def test_same_batch_nonadjacent_cells_keep_original_tables_and_do_not_expand_evidence(self):
        snapshot,result=cell_fixture();before=copy.deepcopy((snapshot,result))
        value,engine=self._review_cell_proofs(snapshot,result)
        requests=[p for p in engine.calls if p["task"]=="missing_semantics_review"]
        self.assertEqual(len(requests),1)
        request=requests[0]
        self.assertEqual(request["candidate_cells"],["B3","D8"])
        self.assertEqual(request["layout_batch_id"],"同一原包")
        self.assertEqual(request["context_range"],"B3:D8")
        self.assertNotIn("C5",request["cells"])
        self.assertEqual(set(request["cells"]),{"A1","B2","A3","B3","A8","D8"})
        self.assertEqual(set(request["cell_tables"]),{"B3","D8"})
        for mapping in value["reviewed_existing_mappings"]:
            address=mapping["cell"];proof=next(p for p in snapshot["sheets"][0]["confirmed_business_ranges"] if p["table_range"]==f"{address}:{address}")
            self.assertEqual(mapping["table_range"],proof["table_range"])
            self.assertEqual(request["cell_tables"][address],{**proof,"range":proof["table_range"]})
            self.assertEqual(mapping["table_semantic"],proof["table_semantic"])
        self.assertEqual(len(value["reviewed_existing_mappings"]),2,value)
        self.assertFalse(value["unresolved"])
        self.assertEqual((snapshot,result),before)

    def test_cell_batches_keep_scope_note_original_packet_and_word_table_boundaries(self):
        second=copy.deepcopy(base_slot());second["id"]="C-N002-T001-S001"
        second["note"].update(id="C-N002",name="其他业务");second["table"]["id"]="C-N002-T001"
        self.write_gold([base_slot(),second])
        for boundary in ("scope","note","batch","word","ambiguous_word","outside_word"):
            with self.subTest(boundary=boundary):
                snapshot,result=cell_fixture();sheet=snapshot["sheets"][0]
                proof=sheet["confirmed_business_ranges"][1]
                if boundary=="scope":proof["scope"]="parent"
                elif boundary=="note":proof["note_ids"]=["C-N002"]
                elif boundary=="batch":proof["layout_batch_id"]="另一原包"
                elif boundary=="word":sheet["word_table_ranges"]=[{"range":"A1:D3"},{"range":"A8:D8"}]
                elif boundary=="ambiguous_word":sheet["word_table_ranges"]=[{"range":"A1:D8"},{"range":"A8:D8"}]
                else:sheet["word_table_ranges"]=[{"range":"A1:D3"}]
                result["confirmed_business_ranges"]={"资料":copy.deepcopy(sheet["confirmed_business_ranges"])}
                value,engine=self._review_cell_proofs(snapshot,result)
                requests=[p for p in engine.calls if p["task"]=="missing_semantics_review"]
                self.assertTrue(all(len(p["candidate_cells"])==1 for p in requests))
                self.assertEqual(len(requests),1 if boundary in {"scope","ambiguous_word","outside_word"} else 2)
                if boundary=="scope":
                    self.assertEqual(requests[0]["candidate_cells"],["B3"])
                    self.assertTrue(any("D8" in u["cells"] for u in value["unresolved"]))
                    self.assertFalse(any(m["cell"]=="D8" for m in value["reviewed_existing_mappings"]))
                if boundary in {"ambiguous_word","outside_word"}:
                    self.assertTrue(any("D8" in u["cells"] and "Word" in u["reason"] for u in value["unresolved"]))

    def test_cell_batch_cannot_borrow_another_cells_text_or_gap_text_in_either_round(self):
        for task in ("missing_semantics_review","missing_semantics_proposal_review"):
            for wrong in ("A8","C5"):
                with self.subTest(task=task,wrong=wrong):
                    value,_=self._review_cell_proofs(bad_evidence=(task,wrong))
                    self.assertFalse(value["new_ids"])
                    self.assertFalse(value["reviewed_existing_mappings"])
                    self.assertEqual({a for u in value["unresolved"] for a in u["cells"]},{"B3","D8"})

    def test_new_slot_sources_retain_each_original_cell_range_and_own_evidence(self):
        value,engine=self._review_cell_proofs(new=True)
        self.assertEqual(len(value["new_ids"]),1,value)
        row=read_standard(value["path"])[value["new_ids"][0]]
        self.assertEqual({s["cell"]:s["table_range"] for s in row["sources"]},{"B3":"B3:B3","D8":"D8:D8"})
        self.assertEqual({s["cell"]:set(s["evidence_cells"]) for s in row["sources"]},{"B3":{"A3","B2"},"D8":{"A8","B2"}})
        self.assertEqual(sum(p["task"]=="missing_semantics_review" for p in engine.calls),1)

    def test_cell_batch_stays_within_forty_cells_and_never_mixes_sheets(self):
        snapshot,result=cell_fixture();sheet=snapshot["sheets"][0]
        template=sheet["confirmed_business_ranges"][0]
        addresses=[f"B{row}" for row in range(3,85,2)]
        sheet["max_row"]=83
        sheet["confirmed_business_ranges"]=[]
        for address in addresses:
            sheet["cells"][address]={"row":int(address[1:]),"column":2,"value":1,"formula":None}
            sheet["confirmed_business_ranges"].append({**copy.deepcopy(template),"table_range":f"{address}:{address}"})
        second=copy.deepcopy(sheet);second["name"]="另一工作表"
        snapshot["sheets"].append(second)
        result["confirmed_business_ranges"]={}
        result["unresolved"]=[{"sheet":name,"cell":a} for name in ("资料","另一工作表") for a in addresses]
        class ContextStub(Stub):
            def _request(engine,system,payload,validator):
                engine.calls.append(copy.deepcopy(payload))
                self.assertEqual(set(payload["cell_tables"]),set(payload["candidate_cells"]))
                self.assertEqual(set(payload["cells"]),set(payload["candidate_cells"])|{"A1","B2","A3"})
                return validator({"decisions":[{"cell":a,"decision":"context","reason":"尚缺原表期间文字","evidence_cells":[]}
                                               for a in payload["candidate_cells"]]})
        engine=ContextStub(self.gold)
        value=review_missing_semantics(engine,snapshot,result,self.gold,self.work/"四十格分包")
        self.assertEqual([(p["sheet"],len(p["candidate_cells"])) for p in engine.calls],
                         [("资料",40),("资料",1),("另一工作表",40),("另一工作表",1)])
        self.assertEqual(len(value["unresolved"]),82)
        self.assertFalse(value["new_ids"]);self.assertFalse(value["reviewed_existing_mappings"])

    def test_old_rectangular_proofs_keep_the_existing_group_and_evidence_range(self):
        snapshot,result=fixture();sheet=snapshot["sheets"][0]
        sheet["cells"]["B4"]={"row":4,"column":2,"value":5,"formula":None};sheet["max_row"]=4
        for area in (sheet["confirmed_business_ranges"][0],result["confirmed_business_ranges"]["资料"][0]):area["table_range"]="A1:B4"
        result["unresolved"].append({"sheet":"资料","cell":"B4"})
        value,engine=self.run_review(snapshot=snapshot,result=result)
        requests=[p for p in engine.calls if p["task"]=="missing_semantics_review"]
        self.assertEqual(len(requests),1)
        self.assertEqual(requests[0]["table_range"],"A1:B4")
        self.assertNotIn("cell_tables",requests[0])
        self.assertEqual({s["table_range"] for s in read_standard(value["path"])[value["new_ids"][0]]["sources"]},{"A1:B4"})

    def test_evidence_errors_identify_candidate_and_invalid_address_in_both_rounds(self):
        from 附注更新.语义补充 import _validate, _validate_proposal_review
        labels={"A3":{"kind":"text","text":"银行存款"},"B3":{"kind":"number"},
                "C3":{"kind":"blank"},"D3":{"kind":"formula"},"E3":{"kind":"number_text"}}
        for address,kind in (("B3","number"),("C3","blank"),("D3","formula"),
                             ("E3","number_text"),("C-N001-T001-S001","不在本包"),("Z99","不在本包")):
            for stage in ("context","existing","new_slot","同意复核","拒绝复核"):
                with self.subTest(address=address,stage=stage):
                    raw={"cell":"B8","reason":"原表文字待核实","evidence_cells":["A3",address]}
                    with self.assertRaises(ValueError) as raised:
                        if stage in ("同意复核","拒绝复核"):
                            raw["accepted"]=stage=="同意复核"
                            _validate_proposal_review({"decisions":[raw]},["B8"],labels)
                        else:
                            raw["decision"]=stage
                            _validate({"decisions":[raw]},["B8"],{}, {},labels,None,{})
                    message=str(raised.exception)
                    self.assertIn("B8",message)
                    self.assertIn(address,message)
                    self.assertIn(kind,message)

    def test_empty_context_evidence_does_not_relax_positive_review_requirements(self):
        from 附注更新.语义补充 import _validate, _validate_proposal_review
        labels={"A3":{"kind":"text","text":"银行存款"}}
        raw={"cell":"B3","decision":"context","reason":"尚缺期间文字","evidence_cells":[]}
        self.assertEqual(_validate({"decisions":[raw]},["B3"],{}, {},labels,None,{})["B3"]["decision"],"context")
        for decision in ("existing","new_slot"):
            with self.subTest(decision=decision), self.assertRaises(ValueError):
                _validate({"decisions":[{**raw,"decision":decision}]},["B3"],{}, {},labels,None,{})
        review={"cell":"B3","accepted":False,"reason":"尚缺期间文字","evidence_cells":[]}
        self.assertFalse(_validate_proposal_review({"decisions":[review]},["B3"],labels)["B3"]["accepted"])
        with self.assertRaises(ValueError):
            _validate_proposal_review({"decisions":[{**review,"accepted":True}]},["B3"],labels)
        for invalid in ("A3",None,{"cell":"A3"},[{}]):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                _validate({"decisions":[{**raw,"evidence_cells":invalid}]},["B3"],{}, {},labels,None,{})

    def test_both_review_requests_supply_real_text_evidence_addresses(self):
        snapshot,result=fixture();sheet=snapshot["sheets"][0]
        sheet["max_row"]=5
        sheet["cells"].update({"A4":{"row":4,"column":1,"value":None,"formula":None},
                              "B4":{"row":4,"column":2,"value":"=1","formula":"=1"},
                              "A5":{"row":5,"column":1,"value":"123.00","formula":None}})
        for area in (sheet["confirmed_business_ranges"][0],result["confirmed_business_ranges"]["资料"][0]):
            area["table_range"]="A1:B5"
        class PromptStub(Stub):
            def _request(engine,system,payload,validator):
                if payload["task"] in {"missing_semantics_review","missing_semantics_proposal_review"}:
                    for wording in ("text_evidence_cells","kind=text","金标准ID","数字格","空白格","自身"):
                        self.assertIn(wording,system)
                    self.assertEqual(payload.get("text_evidence_cells"),["A1","A2","B2","A3"])
                    self.assertTrue(all(payload["cells"][a]["kind"]=="text" for a in payload["text_evidence_cells"]))
                return super()._request(system,payload,validator)
        before=copy.deepcopy((snapshot,result));engine=PromptStub(self.gold)
        with patch("附注更新.语义补充.ROOT",self.work):
            value=review_missing_semantics(engine,snapshot,result,self.gold,self.work/"文字证据请求核验")
        self.assertTrue(value["new_ids"])
        self.assertEqual(sum(p["task"]=="missing_semantics_review" for p in engine.calls),1)
        self.assertEqual(sum(p["task"]=="missing_semantics_proposal_review" for p in engine.calls),1)
        self.assertEqual((snapshot,result),before)

    def test_two_reviews_publish_full_new_version_without_overwriting(self):
        before=self.gold.read_bytes();source,result=fixture();unchanged=copy.deepcopy(source)
        value,engine=self.run_review(snapshot=source,result=result)
        self.assertEqual(value["new_ids"],["C-N001-T001-S100"])
        self.assertNotEqual(Path(value["path"]),self.gold)
        rows=read_standard(value["path"]);new=rows[value["new_ids"][0]]
        self.assertEqual(rows["C-N001-T001-S001"],base_slot())
        self.assertEqual(new["blank_policy"],proposal()["definition"]["blank_policy"])
        self.assertEqual(new["scope"],"consolidated")
        self.assertEqual(new["applicability"]["report_scopes"][0]["scope"],"standalone")
        self.assertTrue(new["sources"])
        self.assertEqual(self.gold.read_bytes(),before);self.assertEqual(source,unchanged)
        self.assertEqual(sum(p["task"]=="missing_semantics_review" for p in engine.calls),1)
        reviews=[p for p in engine.calls if p["task"]=="missing_semantics_proposal_review"]
        self.assertEqual(len(reviews),1)
        self.assertEqual(reviews[0]["proposals"][0]["definition"]["slot"]["meaning"],proposal()["definition"]["slot"]["meaning"])
        self.assertNotIn("123.45",json.dumps(engine.calls,ensure_ascii=False))
    def test_equal_labels_in_different_business_populations_are_not_mechanically_identical(self):
        left=base_slot();right=copy.deepcopy(left)
        right["table"]={"id":"C-N001-T002","name":"特定客户组合","meaning":"仅该组合内客户的余额"}
        self.assertFalse(_similar(left,right))
        self.assertTrue(_similar(left,copy.deepcopy(left)))

    def test_open_period_is_not_an_open_row_object(self):
        original=base_slot();candidate=copy.deepcopy(original)
        for row in (original,candidate):
            row["column_path"]=["余额"];row["aliases"]["column_labels"]=["余额"]
            for dim in row["dimensions"]:
                if dim["name"]=="period":dim.pop("role");dim["open"]=True
        candidate["slot"].update(name="数字人民币钱包余额",meaning="数字人民币钱包持有金额")
        candidate["row_path"]=["数字人民币钱包"];candidate["aliases"]["row_labels"]=["数字人民币钱包"]
        self.assertFalse(_similar(candidate,original))

    def test_proposal_does_not_treat_open_date_or_currency_as_open_business_row(self):
        for dimension in ("period","currency"):
            with self.subTest(dimension=dimension):
                original=base_slot();original["row_path"]=["<其他项目>"]
                original["slot"]["name"]="<其他项目> / 期末余额"
                original["aliases"]["row_labels"]=["<其他项目>"]
                for dim in original["dimensions"]:
                    if dim["name"]==dimension:dim.pop("role",None);dim["open"]=True
                item=proposal();item["definition"].update(
                    dimensions=copy.deepcopy(original["dimensions"]),column_path=["期末余额"],
                    slot={"name":"数字人民币钱包 / 期末余额","meaning":"数字人民币钱包实际余额","role":"measure"})
                result=_proposal(item,{original["id"]:original},{"note_names":{"货币资金"},"scope":"consolidated"},["A3","B2"])
                self.assertEqual(result["definition"]["row_path"],["数字人民币钱包"])

    def test_actual_project_fair_value_and_total_are_not_duplicate_indicators(self):
        gold=ROOT/"金标准/版本/财务报表附注语义金标准_20260920_095458_480775_cc3efdf5.jsonl"
        rows=read_standard(gold)
        part,total=(rows[key] for key in ("C-N009-T002-S002","C-N009-T002-S006"))
        before=copy.deepcopy((part,total))
        for left,right in ((part,total),(total,part)):
            for check in (_same_indicator,_similar):
                with self.subTest(check=check.__name__,candidate=left["id"]):self.assertFalse(check(left,right))
        # 复用真实总项的定义形状，候选增加必要说明；不改标准原件。
        base=copy.deepcopy(total);base["id"]="C-N009-T002-S099"
        base["slot"]["name"]="合计 / 持有成本";base["column_path"]=["持有成本"]
        base["aliases"]["column_labels"]=["持有成本"]
        item=proposal();item["base_slot_id"]=base["id"]
        item["definition"]={key:copy.deepcopy(total[key]) for key in item["definition"] if key in total}
        item["definition"].update(note_id=total["note"]["id"],table_id=total["table"]["id"])
        item["definition"]["slot"]["meaning"]="全部受限交易性金融资产公允价值合计"
        item["definition"]["blank_policy"]["default_interpretation"]="unknown"
        result=_proposal(item,{part["id"]:part,base["id"]:base},{"note_names":{"交易性金融资产"},"scope":"consolidated"},["A3","B2"])
        self.assertEqual(result["definition"]["aggregation"]["role"],"total")
        self.assertEqual((part,total),before)

    def test_open_object_cannot_erase_fixed_business_hierarchy_or_table_population(self):
        original=base_slot();original["dimensions"].append({"name":"project","required":True,"open":True})
        original["row_path"]=["受限资产","<项目>"];original["slot"]["name"]="受限资产 / <项目> / 期末余额"
        candidate=copy.deepcopy(original);candidate["row_path"][0]="未受限资产";candidate["slot"]["name"]="未受限资产 / <项目> / 期末余额"
        for check in (_same_indicator,_similar):self.assertFalse(check(candidate,original))
        candidate=copy.deepcopy(original);candidate["table"]["id"]="C-N001-T002"
        for check in (_same_indicator,_similar):self.assertFalse(check(candidate,original))

    def test_proposal_still_rejects_new_object_or_age_values_of_same_indicator(self):
        for name in ("counterparty","project","age_bucket"):
            with self.subTest(name=name):
                original=base_slot();original["row_path"]=["<披露对象>"]
                original["slot"]["name"]="<披露对象> / 期末余额"
                original["dimensions"].append({"name":name,"required":True,"open":True})
                item=proposal();item["definition"].update(
                    dimensions=copy.deepcopy(original["dimensions"]),row_path=["<新对象>"],column_path=["期末余额"],
                    slot={"name":"<新对象> / 期末余额","meaning":"同一指标按新增对象披露","role":"measure"})
                with self.assertRaisesRegex(ValueError,"已有指标|已有同义"):
                    _proposal(item,{original["id"]:original},{"note_names":{"货币资金"},"scope":"consolidated"},["A3","B2"])

    def _review_two_standard_tables(self,distinct_population=False,uncertain=False):
        first=base_slot();second=copy.deepcopy(first)
        second["id"]="C-N001-T002-S001"
        second["table"]={"id":"C-N001-T002","name":"履约保证金钱包明细" if distinct_population else "货币资金补充说明",
            "meaning":"仅用于履约保证金的钱包集合" if distinct_population else "同一货币资金总体的补充说明，没有不同对象集合"}
        self.write_gold([first,second]);snapshot,result=fixture()
        another=copy.deepcopy(snapshot["sheets"][0]);another["name"]="补充表";snapshot["sheets"].append(another)
        result["unresolved"].append({"sheet":"补充表","cell":"B3","reason":"缺少定义"})
        result["confirmed_business_ranges"]["补充表"]=copy.deepcopy(another["confirmed_business_ranges"])
        class TwoTableStub(Stub):
            def _request(engine,system,payload,validator):
                if payload["task"]=="missing_semantics_review":
                    engine.calls.append(copy.deepcopy(payload));item=proposal()
                    if payload["sheet"]=="补充表":
                        item["base_slot_id"]=second["id"];item["definition"]["table_id"]=second["table"]["id"]
                        if distinct_population:item["definition"]["slot"]["meaning"]="仅用于履约保证金的钱包中受限金额，不包含全部钱包"
                    return validator({"decisions":[item]})
                if payload["task"]=="missing_semantics_existing_check" and payload.get("checking_pending"):
                    engine.calls.append(copy.deepcopy(payload));ids=[d["id"] for d in payload["existing_slots"]]
                    return validator({"checked_ids":ids,"existing_ids":[] if distinct_population or uncertain else ids,
                        "open_dimension_ids":[],"uncertain":uncertain,"reason":"依据指标及对象集合区分，不按表名或行名判断"})
                return super()._request(system,payload,validator)
        engine=TwoTableStub(self.gold)
        with patch("附注更新.语义补充.ROOT",self.work):
            value=review_missing_semantics(engine,snapshot,result,self.gold,self.work/"跨表审阅")
        return value,engine

    def test_same_new_indicator_across_two_tables_publishes_only_first_definition(self):
        value,engine=self._review_two_standard_tables()
        self.assertEqual(len(value["new_ids"]),1,value)
        self.assertEqual(value["unresolved"][0]["sheet"],"补充表")
        checks=[p for p in engine.calls if p.get("checking_pending")]
        self.assertEqual([p["round"] for p in checks],[1,2])
        self.assertEqual(sum(p["task"]=="missing_semantics_existing_check" and not p.get("checking_pending") for p in engine.calls),4)

    def test_same_table_synonyms_are_reviewed_before_a_second_id_is_added(self):
        snapshot,result=fixture()
        another=copy.deepcopy(snapshot["sheets"][0]);another["name"]="同义列示"
        another["cells"]["A3"]["value"]="数字人民币钱包受限资金"
        another["cells"]["B2"]["value"]="不能动用的资金金额"
        snapshot["sheets"].append(another)
        result["unresolved"].append({"sheet":"同义列示","cell":"B3","reason":"同一业务的另一种列示"})
        result["confirmed_business_ranges"]["同义列示"]=copy.deepcopy(another["confirmed_business_ranges"])
        class SameTableStub(Stub):
            def _request(engine,system,payload,validator):
                if payload["task"]=="missing_semantics_review" and payload["sheet"]=="同义列示":
                    engine.calls.append(copy.deepcopy(payload));item=proposal()
                    item["definition"]["slot"]["name"]="数字人民币钱包 / 不能动用的资金金额"
                    item["definition"].update(row_path=["数字人民币钱包受限资金"],column_path=["不能动用的资金金额"],
                        aliases={"row_labels":["数字人民币钱包受限资金"],"column_labels":["不能动用的资金金额"]})
                    return validator({"decisions":[item]})
                if payload["task"]=="missing_semantics_existing_check" and payload.get("checking_pending"):
                    engine.calls.append(copy.deepcopy(payload));ids=[d["id"] for d in payload["existing_slots"]]
                    return validator({"checked_ids":ids,"existing_ids":ids,"open_dimension_ids":[],"uncertain":False,
                        "reason":"同一标准表、同一钱包总体、同一余额日期的受限金额，仅列示措辞不同"})
                return super()._request(system,payload,validator)
        engine=SameTableStub(self.gold)
        with patch("附注更新.语义补充.ROOT",self.work):
            value=review_missing_semantics(engine,snapshot,result,self.gold,self.work/"同表同义审阅")
        self.assertEqual(len(value["new_ids"]),1,value)
        self.assertEqual(value["unresolved"][0]["sheet"],"同义列示")
        checks=[p for p in engine.calls if p.get("checking_pending")]
        self.assertEqual([p["round"] for p in checks],[1,2])
        self.assertTrue(all(p["existing_slots"][0]["table"]["id"]==p["proposal"]["table"]["id"] for p in checks))
        self.assertEqual(sum(p["task"]=="missing_semantics_existing_check" and not p.get("checking_pending") for p in engine.calls),4)

    def test_equal_names_in_distinct_pending_populations_can_both_publish(self):
        value,engine=self._review_two_standard_tables(distinct_population=True)
        self.assertEqual(len(value["new_ids"]),2,value);self.assertFalse(value["unresolved"])
        self.assertEqual(len([p for p in engine.calls if p.get("checking_pending")]),2)

    def test_uncertain_pending_population_keeps_first_but_does_not_append_second(self):
        value,engine=self._review_two_standard_tables(uncertain=True)
        self.assertEqual(len(value["new_ids"]),1,value);self.assertTrue(value["unresolved"])
        self.assertEqual(len([p for p in engine.calls if p.get("checking_pending")]),2)

    def test_new_measure_uses_one_open_period_definition(self):
        value,_=self.run_review("open_period")
        self.assertEqual(len(value["new_ids"]),1,value)
        row=read_standard(value["path"])[value["new_ids"][0]]
        period=next(d for d in row["dimensions"] if d["name"]=="period")
        self.assertEqual(period,{"name":"period","required":True,"open":True,"type":"instant"})

    def test_same_batch_cannot_publish_period_variants_as_different_measures(self):
        snapshot,result=fixture()
        snapshot["sheets"][0]["max_row"]=4
        snapshot["sheets"][0]["cells"]["B4"]={"row":4,"column":2,"value":456.78,"formula":None}
        for table in snapshot["sheets"][0]["confirmed_business_ranges"]:table["table_range"]="A1:B4"
        result["unresolved"].append({"sheet":"资料","cell":"B4","reason":"尚未确定语义"})
        value,_=self.run_review("same_indicator_period_pair",snapshot=snapshot,result=result)
        self.assertFalse(value["new_ids"],value)
        self.assertEqual(len(value["unresolved"]),2)

    def test_independent_reviewer_evaluates_same_complete_definition(self):
        value,engine=self.run_review("paraphrase")
        self.assertEqual(len(value["new_ids"]),1)
        saved=read_standard(value["path"])[value["new_ids"][0]]
        self.assertEqual(saved["slot"]["meaning"],proposal()["definition"]["slot"]["meaning"])
        self.assertTrue(saved["semantic_review"]["reasons"][0]["independent_review"])

    def test_true_new_measure_can_have_its_own_value_type_and_period(self):
        for mode,label,kind,period in (("new_percentage","受限比例","percentage",{"name":"period","required":True,"role":"closing","type":"instant"}),
                                     ("new_duration","本期受限资金变动","monetary",{"name":"period","required":True,"role":"current","type":"duration"})):
            with self.subTest(mode=mode):
                snapshot,result=fixture();snapshot["sheets"][0]["cells"]["B2"]["value"]=label
                value,_=self.run_review(mode,snapshot=snapshot,result=result)
                self.assertEqual(len(value["new_ids"]),1,value)
                added=read_standard(value["path"])[value["new_ids"][0]]
                self.assertEqual(added["value_type"],kind)
                self.assertIn(period,added["dimensions"])
                if kind=="percentage":self.assertNotIn("currency",{d["name"] for d in added["dimensions"]})

    def test_invalid_dimensions_and_open_metric_never_publish(self):
        for mode in ("metric_escape","fixed_object","invalid_age","invalid_period"):
            with self.subTest(mode=mode):
                value,_=self.run_review(mode)
                self.assertFalse(value["new_ids"],value)
                self.assertTrue(value["unresolved"])

    def test_same_measure_with_only_different_period_is_not_a_new_slot(self):
        snapshot,result=fixture();snapshot["sheets"][0]["cells"]["A3"]["value"]="银行存款";snapshot["sheets"][0]["cells"]["B2"]["value"]="期初余额"
        value,_=self.run_review("dimension_only",snapshot=snapshot,result=result)
        self.assertFalse(value["new_ids"]);self.assertTrue(value["unresolved"])

    def test_templates_keep_different_period_and_object_dimensions(self):
        opening=copy.deepcopy(base_slot());opening["id"]="C-N001-T001-S002"
        for dim in opening["dimensions"]:
            if dim["name"]=="period":dim["role"]="opening"
        self.write_gold([base_slot(),opening])
        _,engine=self.run_review()
        presented=next(p["templates"] for p in engine.calls if p["task"]=="missing_semantics_review")
        self.assertEqual({x["id"] for x in presented},{base_slot()["id"],opening["id"]})

    def _review_complete_templates(self, rows, selected=None, mode="context", batch_size=2, target=None):
        self.write_gold(rows)
        snapshot,result=fixture()
        class BatchStub(Stub):
            def _slot_batches(engine, slots):
                if batch_size is None:
                    yield from super()._slot_batches(slots)
                    return
                for offset in range(0,len(slots),batch_size):yield slots[offset:offset+batch_size]
            def _request(engine,system,payload,validator):
                if payload["task"]=="tables":
                    engine.calls.append(copy.deepcopy(payload))
                    return validator({"table_ids":selected or []})
                if payload["task"]=="missing_semantics_review":
                    engine.calls.append(copy.deepcopy(payload))
                    identifiers=[t["id"] for t in payload["templates"]]
                    if mode=="failed" and identifiers[-1].endswith("S003"):raise ValueError("该目录包未完成")
                    matching=target if target in identifiers else identifiers[0] if mode=="conflict" else None
                    if matching:
                        raw={"cell":"B3","decision":"existing","slot_id":matching,"scope":"standalone",
                            "dimensions":{"currency":"CNY","period":"2025-12-31","unit":"元","scale":1},
                            "value_type":"monetary","semantic_field":"银行存款余额",
                            "reason":"完整指标和原表证据一致","evidence_cells":["A3","B2"]}
                    elif mode in {"new_then_existing","failed","new_conflict"}:
                        raw=proposal();raw["base_slot_id"]=identifiers[0]
                        base=engine.slots[identifiers[0]]
                        raw["definition"].update(note_id=base["note"]["id"],table_id=base["table"]["id"])
                        if mode=="new_conflict":raw["definition"]["slot"]["meaning"]+=identifiers[0]
                    else:raw={"cell":"B3","decision":"context","reason":"本子目录尚未确定","evidence_cells":[]}
                    return validator({"decisions":[raw]})
                return super()._request(system,payload,validator)
        engine=BatchStub(self.gold)
        with patch("附注更新.语义补充.ROOT",self.work):
            value=review_missing_semantics(engine,snapshot,result,self.gold,self.work/"完整定义分包")
        return value,engine

    def _approved_rows(self, count=3):
        rows=[]
        for i in range(1,count+1):
            row=base_slot();row["id"]="C-N001-T001-S"+str(i).zfill(3)
            row["slot"]["name"]="不同稳定指标"+str(i)
            row["applicability"]["report_scopes"]=[{"scope":"standalone","reason":"单户同义","reviewed_by":"本地测试","reviewed_at":"2026-09-20"}]
            rows.append(row)
        return rows

    def test_templates_preserve_all_stable_ids_with_identical_dimensions_and_batch(self):
        rows=self._approved_rows(5)
        _,engine=self._review_complete_templates(rows)
        packets=[p for p in engine.calls if p["task"]=="missing_semantics_review"]
        shown=[t["id"] for p in packets for t in p["templates"]]
        self.assertEqual(shown,[r["id"] for r in rows])
        self.assertEqual([len(p["templates"]) for p in packets],[2,2,1])
        self.assertTrue(all(p["existing_slot_ids"]==[t["id"] for t in p["templates"]] for p in packets))

    def test_existing_scope_candidates_and_unapproved_new_slot_templates_are_distinct(self):
        rows=self._approved_rows(2);rows[1]["applicability"].pop("report_scopes")
        other=copy.deepcopy(rows[0]);other["id"]="P-N002-T001-S001";other["scope"]="parent"
        other["note"]["id"]="P-N002";other["table"]["id"]="P-N002-T001"
        _,engine=self._review_complete_templates(rows+[other],selected=[rows[0]["table"]["id"],other["table"]["id"]])
        packets=[p for p in engine.calls if p["task"]=="missing_semantics_review"]
        self.assertEqual([t["id"] for p in packets for t in p["templates"]],[r["id"] for r in rows+[other]])
        self.assertEqual([i for p in packets for i in p["existing_slot_ids"]],[rows[0]["id"],other["id"]])
        value,_=self.run_review()
        self.assertEqual(len(value["new_ids"]),1,"未获准单户的旧定义仍可作新指标基底，须独立审查新指标适用性")

    def test_existing_same_meaning_note_is_reviewed_without_new_slot(self):
        rows=self._approved_rows(1)
        peer=copy.deepcopy(rows[0]);peer["id"]="P-N002-T001-S001";peer["scope"]="parent"
        peer["note"]["id"]="P-N002";peer["table"]["id"]="P-N002-T001"
        value,_=self._review_complete_templates(rows+[peer],target=peer["id"])
        self.assertFalse(value["new_ids"],value)
        self.assertEqual([m["slot_id"] for m in value["reviewed_existing_mappings"]],[peer["id"]])
        accepted=value["reviewed_existing_mappings"][0]
        self.assertEqual((accepted["scope"],accepted["definition_scope"]),("standalone","parent"))

    def test_selected_table_local_new_proposal_expands_to_later_existing_definition(self):
        rows=self._approved_rows(3)
        rows[2]["table"]["id"]="C-N001-T002";rows[2]["id"]="C-N001-T002-S001"
        value,engine=self._review_complete_templates(rows,selected=["C-N001-T001"],mode="new_then_existing",target=rows[2]["id"])
        self.assertFalse(value["new_ids"])
        self.assertEqual([m["slot_id"] for m in value["reviewed_existing_mappings"]],[rows[2]["id"]])
        self.assertFalse(any(p["task"]=="missing_semantics_existing_check" for p in engine.calls))
        self.assertEqual(sum(p["task"]=="missing_semantics_proposal_review" for p in engine.calls),1)

    def test_empty_selection_searches_all_legal_existing_and_selected_hit_avoids_unrelated_tables(self):
        rows=self._approved_rows(3)
        rows[2]["table"]["id"]="C-N001-T002";rows[2]["id"]="C-N001-T002-S001"
        unavailable=copy.deepcopy(rows[2]);unavailable["id"]="C-N001-T002-S002";unavailable["applicability"].pop("report_scopes")
        _,engine=self._review_complete_templates(rows+[unavailable])
        shown=[t["id"] for p in engine.calls if p["task"]=="missing_semantics_review" for t in p["templates"]]
        self.assertEqual(shown,[r["id"] for r in rows])
        value,engine=self._review_complete_templates(rows,selected=["C-N001-T001"],target=rows[0]["id"])
        self.assertEqual(len(value["reviewed_existing_mappings"]),1)
        shown=[t["id"] for p in engine.calls if p["task"]=="missing_semantics_review" for t in p["templates"]]
        self.assertEqual(shown,[r["id"] for r in rows[:2]])

    def test_real_size_split_preserves_definitions_and_second_review_counts_full_proposals(self):
        rows=self._approved_rows(3)
        for row in rows:row["slot"]["meaning"]="独立完整业务定义"*3000
        _,engine=self._review_complete_templates(rows,batch_size=None)
        packets=[p for p in engine.calls if p["task"]=="missing_semantics_review"]
        self.assertGreater(len(packets),1)
        self.assertEqual([t["id"] for p in packets for t in p["templates"]],[r["id"] for r in rows])
        self.assertTrue(all(t["slot"]["meaning"]==rows[0]["slot"]["meaning"] for p in packets for t in p["templates"]))
        from 附注更新.语义补充 import _review_template_proposals, _summary
        class ReviewStub(Stub):
            def _request(local,system,payload,validator):
                local.calls.append(copy.deepcopy(payload))
                return validator({"decisions":[{"cell":a,"accepted":True,"reason":"完整业务定义一致","evidence_cells":["A3"]} for a in payload["candidate_cells"]]})
        actual=ReviewStub(self.gold)
        proposed={"B"+str(i+3):{"decision":"existing","definition":copy.deepcopy(row),"reason":"已有指标"} for i,row in enumerate(rows)}
        checked=_review_template_proposals(actual,proposed,list(proposed),{r["id"]:r for r in rows},{},
                                          {"A3":{"kind":"text","text":"原表项目"}}, {})
        self.assertEqual(set(checked),set(proposed))
        self.assertGreater(len(actual.calls),1)
        self.assertTrue(all(p["templates"]==[] for p in actual.calls))
        self.assertEqual([p["definition"] for call in actual.calls for p in call["proposals"]],[_summary(r) for r in rows])

    def test_template_table_selection_reuses_strict_tables_protocol(self):
        from 附注更新.回执格式 import response_format, normalize_response
        rows=self._approved_rows(3)
        rows[2]["table"]["id"]="C-N001-T002";rows[2]["id"]="C-N001-T002-S001"
        _,engine=self._review_complete_templates(rows,selected=["C-N001-T001"])
        payload=next(p for p in engine.calls if p["task"]=="tables")
        self.assertEqual(payload["purpose"],"missing_semantics_review")
        protocol=response_format(payload)
        self.assertEqual(protocol["type"],"json_schema")
        self.assertTrue(protocol["json_schema"]["strict"])
        schema=protocol["json_schema"]["schema"]
        self.assertEqual(schema["required"],["table_ids"])
        self.assertFalse(schema["additionalProperties"])
        self.assertEqual(set(schema["properties"]["table_ids"]["items"]["enum"]),{"C-N001-T001","C-N001-T002"})
        raw={"table_ids":["C-N001-T001"]}
        self.assertEqual(normalize_response(payload,raw),raw)

    def test_cross_batch_existing_conflict_failure_and_new_conflict_remain_unresolved(self):
        for mode in ("conflict","failed","new_conflict"):
            with self.subTest(mode=mode):
                value,engine=self._review_complete_templates(self._approved_rows(3),mode=mode,batch_size=1)
                self.assertFalse(value["new_ids"])
                self.assertFalse(value["reviewed_existing_mappings"])
                self.assertTrue(value["unresolved"])
                self.assertFalse(any(p["task"]=="missing_semantics_existing_check" for p in engine.calls))

    def test_fixed_dimensions_are_part_of_exact_duplicate_check(self):
        original=base_slot();other=copy.deepcopy(original)
        other["dimensions"].append({"name":"age_bucket","required":True,"lower":1,"upper":2,"lower_inclusive":False,"upper_inclusive":True})
        self.assertFalse(_similar(other,original))
        other=copy.deepcopy(original);other["value_type"]="percentage"
        self.assertFalse(_similar(other,original))
        self.assertTrue(_similar(original,copy.deepcopy(original)))

    def test_invalid_or_disagreeing_proposals_never_publish(self):
        for mode in ("disagree","bad_evidence","bad_review_evidence","missing_dimension","wrong_table","scope_unreviewed","bad_reference","failed_check","synonym_check","exact_duplicate","bad_role"):
            with self.subTest(mode=mode):
                value,_=self.run_review(mode)
                self.assertFalse(value["new_ids"]);self.assertEqual(Path(value["path"]),self.gold)
                self.assertTrue(value["unresolved"])
    def _review_existing(self,change=None,accept=True,review_evidence=None,shared_header=False):
        first=base_slot();first["slot"].update(name="库存现金 / 期末余额",meaning="期末库存现金余额")
        existing=base_slot();existing["id"]="C-N001-T001-S002"
        existing["sources"]=[{"file":"原始格式1.xlsx","sheet":"货币资金","cell":"B3"}]
        self.write_gold([first,existing])
        snapshot,result=fixture();sheet=snapshot["sheets"][0]
        sheet["cells"]["A3"]["value"]="银行存款";sheet["cells"]["B2"]["value"]="期末余额"
        sheet["cells"]["B3"]["value"]=None
        for group in (sheet["confirmed_business_ranges"][0],result["confirmed_business_ranges"]["资料"][0]):
            group["scope"]="consolidated"
            if shared_header:group["table_range"]="A3:B3"
        raw={"cell":"B3","decision":"existing","slot_id":existing["id"],"scope":"consolidated",
             "dimensions":{"currency":"人民币","period":{"type":"instant","date":"2025-12-31"},"unit":"元","scale":"1"},
             "value_type":"monetary","semantic_field":"银行存款 / 期末余额",
             "evidence_cells":["A3","B2"],"reason":"银行存款期末余额及原表日期、币种单位明确"}
        if change:change(raw,snapshot)
        unchanged=copy.deepcopy((snapshot,result));gold_before=self.gold.read_bytes()
        class ExistingStub(Stub):
            def _request(engine,system,payload,validator):
                engine.calls.append(copy.deepcopy(payload))
                if payload["task"]=="missing_semantics_review":return validator({"decisions":[raw]})
                if payload["task"]=="missing_semantics_proposal_review":
                    if accept is None:return validator({"decisions":[]})
                    definition=payload["proposals"][0].get("definition",{})
                    fields={"id","status","scope","note","table","slot","row_path","column_path","value_type","dimensions","blank_policy","aggregation","calculation","aliases","applicability"}
                    complete=definition=={key:existing[key] for key in fields}
                    header=payload["cells"].get("B2",{}).get("text")=="期末余额"
                    return validator({"decisions":[{"cell":"B3","accepted":bool(accept and complete and header),
                        "reason":"完整旧定义与原表同为银行存款期末余额，维度和实际口径一致",
                        "evidence_cells":["A3","B2"] if review_evidence is None else review_evidence}]})
                raise AssertionError("复用已有指标不得启动新增查重或发布")
        engine=ExistingStub(self.gold)
        value=review_missing_semantics(engine,snapshot,result,self.gold,self.work/"旧指标复核")
        self.assertEqual((snapshot,result),unchanged)
        self.assertEqual(self.gold.read_bytes(),gold_before)
        return value,engine,existing

    def test_existing_review_returns_verified_mapping_and_preserves_blank_and_inputs(self):
        value,engine,existing=self._review_existing(shared_header=True)
        self.assertFalse(value["new_ids"]);self.assertEqual(Path(value["path"]),self.gold)
        mapped=value["reviewed_existing_mappings"]
        self.assertEqual(len(mapped),1,value)
        item=mapped[0]
        self.assertEqual((item["sheet"],item["cell"],item["slot_id"]),("资料","B3",existing["id"]))
        self.assertEqual(item["dimensions"],{"currency":"CNY","period":"2025-12-31","unit":"元","scale":1})
        self.assertEqual(item["scope"],"consolidated")
        self.assertTrue(item["reviewed"])
        self.assertEqual(item["review_method"],"existing_slot_proposal_and_independent_review")
        self.assertEqual(item["semantic_review"]["proposal"]["definition"],existing)
        self.assertTrue(item["semantic_review"]["independent_review"]["accepted"])
        self.assertEqual(item["semantic_review"]["evidence_cells"]["B2"],"期末余额")
        self.assertNotIn("value",item);self.assertFalse(value["unresolved"])
        first=next(p for p in engine.calls if p["task"]=="missing_semantics_review")
        self.assertIn(existing["id"],{t["id"] for t in first["templates"]})
        second=next(p for p in engine.calls if p["task"]=="missing_semantics_proposal_review")
        definition=second["proposals"][0]["definition"]
        self.assertNotIn("sources",definition);self.assertNotIn("introduced_version",definition)
        self.assertEqual(set(definition),{"id","status","scope","note","table","slot","row_path","column_path","value_type","dimensions","blank_policy","aggregation","calculation","aliases","applicability"})
        self.assertEqual(item["semantic_review"]["proposal"]["definition"]["sources"],existing["sources"])
        report=Path(value["report_path"])
        self.assertEqual(value["report_sha256"],hashlib.sha256(report.read_bytes()).hexdigest())
        saved=json.loads(report.read_text(encoding="utf-8-sig"))
        self.assertEqual(saved["reviewed_existing_mappings"],mapped)
        self.assertNotIn("report_path",saved);self.assertNotIn("report_sha256",saved)
        # 经正式应用助手另存，原待定结果与原格空白仍保持不变。
        from 附注更新.分步流程 import _apply_reviewed_existing
        snapshot,result=fixture();sheet=snapshot["sheets"][0]
        sheet["cells"]["A3"]["value"]="银行存款";sheet["cells"]["B2"]["value"]="期末余额"
        sheet["cells"]["B3"]["value"]=None
        for group in (sheet["confirmed_business_ranges"][0],result["confirmed_business_ranges"]["资料"][0]):
            group.update(scope="consolidated",table_range="A3:B3")
        result["excluded"]=[{"sheet":"资料","cell":a,"category":"label","reason":"原表文字"}
                            for a in sheet["cells"] if a!="B3"]
        before=copy.deepcopy((snapshot,result))
        actual=SemanticEngine({},str(self.gold),str(self.work/"正式应用"))
        applied=_apply_reviewed_existing(result,value,actual,snapshot)
        self.assertEqual((snapshot,result),before)
        self.assertTrue(applied["complete"]);self.assertFalse(applied["unresolved"])
        self.assertEqual(applied["mappings"][0]["slot_id"],existing["id"])
        self.assertEqual(applied["mappings"][0]["dimensions"]["period"],"2025-12-31")
        self.assertEqual(applied["evidence_hash"],hashlib.sha256(Path(applied["evidence_path"]).read_bytes()).hexdigest())

    def test_existing_review_rejects_invalid_mapping_and_evidence(self):
        changes={
            "缺必填日期":lambda raw,s:raw["dimensions"].pop("period"),
            "错误实际口径":lambda raw,s:raw.update(scope="parent"),
            "单户适用性未核准":lambda raw,s:s["sheets"][0]["confirmed_business_ranges"][0].update(scope="standalone"),
            "错误值类型":lambda raw,s:raw.update(value_type="percentage"),
            "金额原格是文字":lambda raw,s:s["sheets"][0]["cells"]["B3"].update(value="缺少资料"),
            "虚构文字证据":lambda raw,s:raw.update(evidence_cells=["Z99"]),
        }
        for label,change in changes.items():
            with self.subTest(label=label):
                value,_,_=self._review_existing(change=change)
                self.assertFalse(value["new_ids"]);self.assertFalse(value["reviewed_existing_mappings"])
                self.assertTrue(value["unresolved"])

    def test_existing_review_requires_independent_acceptance_and_real_text(self):
        for accept,evidence in ((False,None),(None,None),(True,[]),(True,["B3"]),(True,["Z99"])):
            with self.subTest(accept=accept,evidence=evidence):
                value,_,_=self._review_existing(accept=accept,review_evidence=evidence)
                self.assertFalse(value["new_ids"]);self.assertFalse(value["reviewed_existing_mappings"])
                self.assertTrue(value["unresolved"])
    def test_new_object_under_open_dimension_does_not_create_slot(self):
        row=base_slot();row["dimensions"].append({"name":"counterparty","required":True,"open":True})
        self.write_gold([row])
        value,_=self.run_review("open_object")
        self.assertFalse(value["new_ids"]);self.assertTrue(value["unresolved"])
    def test_missing_table_context_stays_unresolved(self):
        s,r=fixture();s["sheets"][0]["confirmed_business_ranges"]=[];r["confirmed_business_ranges"]={}
        value,engine=self.run_review(snapshot=s,result=r)
        self.assertFalse(value["new_ids"]);self.assertFalse(engine.calls)
        self.assertTrue(value["unresolved"])
    def test_identical_new_meaning_for_two_cells_uses_one_new_id(self):
        snapshot,result=fixture()
        snapshot["sheets"][0]["max_row"]=4
        snapshot["sheets"][0]["cells"]["B4"]={"row":4,"column":2,"value":99,"formula":None}
        for area in (snapshot["sheets"][0]["confirmed_business_ranges"][0],result["confirmed_business_ranges"]["资料"][0]):area["table_range"]="A1:B4"
        result["unresolved"].append({"sheet":"资料","cell":"B4","reason":"缺少定义"})
        value,_=self.run_review(snapshot=snapshot,result=result)
        self.assertEqual(len(value["new_ids"]),1)
        self.assertEqual(len(read_standard(value["path"])[value["new_ids"][0]]["sources"]),2)

    def test_conflicting_pending_definitions_publish_neither(self):
        snapshot,result=fixture()
        snapshot["sheets"][0]["max_row"]=4
        snapshot["sheets"][0]["cells"]["B4"]={"row":4,"column":2,"value":99,"formula":None}
        for area in (snapshot["sheets"][0]["confirmed_business_ranges"][0],result["confirmed_business_ranges"]["资料"][0]):area["table_range"]="A1:B4"
        result["unresolved"].append({"sheet":"资料","cell":"B4","reason":"缺少定义"})
        value,_=self.run_review("conflicting_proposals",snapshot=snapshot,result=result)
        self.assertFalse(value["new_ids"])
        self.assertEqual({a for x in value["unresolved"] for a in x.get("cells",[])},{"B3","B4"})

    def test_stale_layout_evidence_does_not_add_definitions(self):
        snapshot,result=fixture()
        for area in (snapshot["sheets"][0]["confirmed_business_ranges"][0],result["confirmed_business_ranges"]["资料"][0]):area["source_hash"]="different-source"
        value,engine=self.run_review(snapshot=snapshot,result=result)
        self.assertFalse(value["new_ids"]);self.assertFalse(engine.calls)
        self.assertTrue(value["unresolved"])

    def test_publisher_failure_does_not_adopt_candidate(self):
        with patch("附注更新.语义补充.publish_revision",side_effect=ValueError("模拟发布校验失败")):
            value,_=self.run_review()
        self.assertFalse(value["new_ids"]);self.assertEqual(Path(value["path"]),self.gold)
        self.assertFalse((self.work/"金标准"/"版本").exists())

    def test_cancel_never_publishes(self):
        with self.assertRaises(SemanticCancelled):self.run_review("cancel")
        self.assertFalse((self.work/"金标准"/"版本").exists())


if __name__=="__main__":unittest.main()
