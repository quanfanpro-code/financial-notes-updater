"""回执格式约束只管JSON形状，业务含义继续由原校验器判断。"""
import copy
import json
import unittest
from jsonschema import Draft202012Validator, ValidationError
from 附注更新.回执格式 import response_format


def scope_payload():
    return {"task":"source_scope","cells":{"A1":{"kind":"text","text":"业务说明"},"B2":{"kind":"number"}},
            "context_cells":{"C1":{"kind":"text","text":"表头"}},"candidate_cells":["A1","B2"]}


class ResponseFormatTests(unittest.TestCase):
    def schema(self,payload):
        result=response_format(payload)
        self.assertEqual(result["type"],"json_schema")
        self.assertIs(result["json_schema"]["strict"],True)
        schema=result["json_schema"]["schema"];Draft202012Validator.check_schema(schema)
        self.assertEqual(schema["type"],"object");self.assertIs(schema["additionalProperties"],False)
        return Draft202012Validator(schema)

    def test_source_scope_limits_categories_and_real_text_evidence(self):
        validator=self.schema(scope_payload())
        value={"areas":[{"range":"A1:B2","decision":"include","reason":"已见业务","evidence_cells":["A1"]}]}
        validator.validate(value)
        for change in ({"decision":"financial"},{"evidence_cells":["B2"]},{"evidence_cells":["Z99"]},{"extra":True}):
            broken=copy.deepcopy(value);broken["areas"][0].update(change)
            with self.assertRaises(ValidationError):validator.validate(broken)
        validator.validate({"areas":[]})

    def test_batch_keeps_object_wrapper_packet_ids_and_candidate_count(self):
        first=scope_payload();first.update(packet_id="1")
        second=scope_payload();second.update(packet_id="2")
        validator=self.schema({"task":"source_scope_batch","tasks":[first,second]})
        value={"results":[{"packet_id":"1","areas":[]},{"packet_id":"2","areas":[]}]};validator.validate(value)
        for broken in (value["results"],{"results":value["results"][:1]}, {"results":[{"packet_id":"3","areas":[]},value["results"][1]]}):
            with self.assertRaises(ValidationError):validator.validate(broken)

    def test_scope_refinement_comparison_is_explicit_and_does_not_change_old_protocol(self):
        payload=scope_payload();payload['refinement']={'candidate_cells':['A1','B2']}
        comparison={'source_business':'应收账款单项准备','target_business':'其他应收款单项准备',
            'difference_kind':'business_population','reason':'债权业务总体不同',
            'target_evidence':['其他应收款']}
        area={'range':'A1:B2','decision':'out_of_scope','reason':'完整业务不同',
              'evidence_cells':['A1'],'semantic_comparison':comparison}
        value={'areas':[area]};before=copy.deepcopy(payload)
        validator=self.schema(payload);validator.validate(value)
        self.assertEqual(payload,before)
        with self.assertRaises(ValidationError):self.schema(scope_payload()).validate(value)
        for change in ({'difference_kind':'counterparty'},{'target_evidence':[]},
                       {'target_evidence':[123]},{'source_business':''}):
            broken=copy.deepcopy(value);broken['areas'][0]['semantic_comparison'].update(change)
            with self.subTest(change=change):
                with self.assertRaises(ValidationError):validator.validate(broken)
        broken=copy.deepcopy(value);del broken['areas'][0]['semantic_comparison']['reason']
        with self.assertRaises(ValidationError):validator.validate(broken)
        # 未决及真正没有同词冲突的区域不用编造一份比较结论。
        without=copy.deepcopy(value);del without['areas'][0]['semantic_comparison'];validator.validate(without)
        first={**scope_payload(),'packet_id':'1'};second={**payload,'packet_id':'2'}
        batch=self.schema({'task':'source_scope_batch','tasks':[first,second]})
        batch.validate({'results':[{'packet_id':'1','areas':[]},{'packet_id':'2','areas':[area]}]})

    def test_fixed_scope_cells_cover_owned_only_and_preserve_independent_business_decisions(self):
        from 附注更新.回执格式 import normalize_response,transport_instruction
        payload={**scope_payload(),"scope_response_mode":"cells_v1","refinement":{"attempt":1}}
        value={"cells":{"A1":{"decision":"out_of_scope","reason":"人数指标另行披露","evidence_cells":["A1"]},
                        "B2":{"decision":"include","reason":"同一行右侧薪酬金额属于本章","evidence_cells":["C1"]}}}
        before=copy.deepcopy((payload,value));validator=self.schema(payload);validator.validate(value)
        result=normalize_response(payload,value)
        self.assertEqual(result,{"areas":[{"range":a,**v} for a,v in value["cells"].items()]})
        self.assertEqual((payload,value),before)
        self.assertIn("不生成range",transport_instruction(payload))
        for invalid in ({"cells":{"A1":value["cells"]["A1"]}},
                        {"cells":{**value["cells"],"Z99":value["cells"]["A1"]}},
                        {"cells":value["cells"],"areas":[]},
                        {"areas":result["areas"]}):
            with self.subTest(invalid=invalid):
                with self.assertRaises(ValidationError):validator.validate(invalid)
                with self.assertRaises(ValueError):normalize_response(payload,invalid)
        for change in ({"range":"A1:B2"},{"decision":"exclude"},{"reason":" "},{"evidence_cells":["B2"]}):
            broken=copy.deepcopy(value);broken["cells"]["A1"].update(change)
            with self.subTest(change=change):
                with self.assertRaises(ValueError):normalize_response(payload,broken)

    def test_fixed_scope_batch_binds_each_packet_and_rejects_foreign_evidence(self):
        from 附注更新.回执格式 import normalize_response
        first={**scope_payload(),"packet_id":"left","scope_response_mode":"cells_v1","refinement":{"attempt":1}}
        second={"packet_id":"right","scope_response_mode":"cells_v1","refinement":{"attempt":1},
            "candidate_cells":["Z2"],"cells":{"Z2":{"kind":"number"}},
            "context_cells":{"Z1":{"kind":"text","text":"工资金额"}}}
        payload={"task":"source_scope_batch","tasks":[first,second]}
        decision={"decision":"include","reason":"同一业务","evidence_cells":["A1"]}
        value={"packets":{"left":{"cells":{"A1":decision,"B2":decision}},
                          "right":{"cells":{"Z2":{**decision,"evidence_cells":["Z1"]}}}}}
        before=copy.deepcopy((payload,value));validator=self.schema(payload);validator.validate(value)
        normalized=normalize_response(payload,value)
        self.assertEqual([r["packet_id"] for r in normalized["results"]],["left","right"])
        self.assertEqual(normalized["results"][1]["areas"][0]["range"],"Z2")
        self.assertEqual((payload,value),before)
        missing=copy.deepcopy(value);del missing["packets"]["right"]
        extra=copy.deepcopy(value);extra["packets"]["other"]=extra["packets"]["right"]
        foreign=copy.deepcopy(value);foreign["packets"]["right"]["cells"]["Z2"]["evidence_cells"]=["A1"]
        for invalid in (missing,extra,foreign):
            with self.subTest(invalid=invalid):
                with self.assertRaises(ValueError):normalize_response(payload,invalid)
        # 旧任务未选择新模式，仍保持旧数组协议。
        legacy={"areas":[]}
        self.assertEqual(normalize_response(scope_payload(),legacy),legacy)

    def test_fixed_scope_shape_still_requires_original_business_check(self):
        from 附注更新.回执格式 import normalize_response
        from 附注更新.识别范围 import _validate
        payload={**scope_payload(),"scope_response_mode":"cells_v1","refinement":{"attempt":1},
                 "target_meanings":{"disclosures":["营业收入"]}}
        payload["cells"]["A1"]["text"]="营业收入"
        wire={"cells":{a:{"decision":"out_of_scope","reason":"只是主表不是附注","evidence_cells":["A1"]}
                       for a in payload["candidate_cells"]}}
        self.schema(payload).validate(wire)
        sheet={"name":"主表","max_row":2,"max_column":3,"cells":payload["cells"]}
        with self.assertRaises(ValueError):_validate(normalize_response(payload,wire),sheet,payload)

    def test_tables_selects_only_catalogue_ids_but_allows_empty(self):
        validator=self.schema({"task":"tables","catalogue":[{"id":"C-N001-T001"},{"id":"P-N001-T002"}]})
        validator.validate({"table_ids":[]});validator.validate({"table_ids":["P-N001-T002"]})
        for broken in ({"table_ids":["S-N001-T003"]},["C-N001-T001"],{"table_ids":[],"explanation":"附加文字"}):
            with self.assertRaises(ValidationError):validator.validate(broken)

    def test_proposal_review_uses_owned_cells_boolean_and_required_fields(self):
        payload={"task":"missing_semantics_proposal_review","candidate_cells":["B3"],"cells":{"A3":{"kind":"text","text":"指标"},"B3":{"kind":"number"}}}
        validator=self.schema(payload);valid={"decisions":[{"cell":"B3","accepted":True,"reason":"原文支持","evidence_cells":["A3"]}]}
        validator.validate(valid)
        for change in ({"cell":"B4"},{"accepted":"true"},{"evidence_cells":["B3"]}):
            invalid=copy.deepcopy(valid);invalid["decisions"][0].update(change)
            with self.assertRaises(ValidationError):validator.validate(invalid)
        invalid=copy.deepcopy(valid);del invalid["decisions"][0]["reason"]
        with self.assertRaises(ValidationError):validator.validate(invalid)

    def test_existing_check_retains_object_even_though_validator_returns_boolean(self):
        payload={"task":"missing_semantics_existing_check","existing_slots":[{"id":"C-N001-T001-S001"},{"id":"pending-example"}]}
        validator=self.schema(payload);valid={"checked_ids":["C-N001-T001-S001","pending-example"],"existing_ids":[],"open_dimension_ids":[],"uncertain":False,"reason":"逐项比较"}
        validator.validate(valid)
        for change in ({"checked_ids":["C-N001-T001-S001"]},{"existing_ids":["wrong-id"]},{"uncertain":"false"}):
            with self.assertRaises(ValidationError):validator.validate({**valid,**change})
        with self.assertRaises(ValidationError):validator.validate(True)

    def classify_fixture(self):
        payload={"task":"classify","candidate_cells":["A1","B2","C3"],"gold_slots":[{"id":"C-N001-T001-S001"}],"cells":{"A1":{"value":"表头"},"B2":{"value":1},"C3":{"value":2}}}
        valid={"cells":{"A1":{"kind":"excluded","category":"header","reason":"表头"},
            "B2":{"kind":"mapping","slot_id":"C-N001-T001-S001","scope":"standalone","dimensions":{"period":{"date":"2026-12-31"},"counterparty":"新对象","future_dimension":{"nested":[1,True]}},"semantic_field":"资金余额","value_type":"monetary","reason":"原格证据","row_label":"银行存款","column_label":"期末余额"},
            "C3":{"kind":"unresolved","reason":"证据不足"}}}
        return payload,valid

    def test_classify_restricts_addresses_ids_and_categories_without_freezing_dimensions(self):
        payload,valid=self.classify_fixture();validator=self.schema(payload);validator.validate(valid)
        for address,change in [("B2",{"slot_id":"missing"}),("B2",{"scope":"single"}),("B2",{"value_type":"amount"}),("A1",{"category":"financial"}),("A1",{"kind":"unknown"}),("B2",{"cell":"Z99"})]:
            invalid=copy.deepcopy(valid);invalid["cells"][address].update(change)
            with self.assertRaises(ValidationError):validator.validate(invalid)
        record=validator.schema["$defs"]["cell_record"]
        self.assertEqual(len(record["oneOf"]),3)
        mapping=next(branch for branch in record["oneOf"] if branch["properties"]["kind"]["enum"]==["mapping"])
        self.assertEqual(set(mapping["required"]),{"kind","reason","slot_id","scope","dimensions","semantic_field","value_type","row_label","column_label"})
        self.assertIs(mapping["properties"]["dimensions"]["additionalProperties"],True)
        self.assertEqual(set(validator.schema["properties"]["cells"]["required"]),set(payload["candidate_cells"]))
        for address in payload["candidate_cells"]:
            self.assertEqual(validator.schema["properties"]["cells"]["properties"][address],{"$ref":"#/$defs/cell_record"})

    def test_classify_wire_requires_every_candidate_once_and_forbids_extra_cells(self):
        payload,valid=self.classify_fixture();validator=self.schema(payload)
        absent=copy.deepcopy(valid);del absent["cells"]["B2"]
        outside=copy.deepcopy(valid);outside["cells"]["Z99"]={"kind":"unresolved","reason":"包外"}
        double=copy.deepcopy(valid);double["cells"]["B2"]=[valid["cells"]["B2"],{"kind":"excluded","reason":"另一个状态"}]
        for invalid in (absent,outside,double,{"cells":{}},{"cells":valid["cells"],"mappings":[]}):
            with self.assertRaises(ValidationError):validator.validate(invalid)
        self.schema({"task":"classify","candidate_cells":[],"gold_slots":[],"cells":{}}).validate({"cells":{}})

    def test_classify_normalization_preserves_business_content_without_mutation(self):
        from 附注更新.回执格式 import normalize_response
        payload,wire=self.classify_fixture();before=copy.deepcopy((payload,wire))
        result=normalize_response(payload,wire)
        self.assertEqual((payload,wire),before)
        self.assertEqual(result["mappings"],[{**{k:v for k,v in wire["cells"]["B2"].items() if k!="kind"},"cell":"B2"}])
        self.assertEqual(result["excluded"][0]["cell"],"A1")
        self.assertEqual(result["unresolved"],[{"cell":"C3","reason":"证据不足"}])
        result["mappings"][0]["dimensions"]["counterparty"]="改动结果不能影响原回执"
        self.assertEqual((payload,wire),before)

    def test_classify_normalization_rejects_missing_outside_mixed_and_malformed(self):
        from 附注更新.回执格式 import normalize_response
        payload,wire=self.classify_fixture()
        absent=copy.deepcopy(wire);del absent["cells"]["B2"]
        outside=copy.deepcopy(wire);outside["cells"]["Z99"]={"kind":"unresolved","reason":"包外"}
        invalids=[absent,outside,{"cells":[]},{"cells":wire["cells"],"mappings":[]},{"cells":wire["cells"],"extra":True},[],{}]
        for change in ({"kind":"wrong"},{"kind":["mapping","excluded"]},{"reason":" "},{"cell":"C3"},{"value":999}):
            broken=copy.deepcopy(wire);broken["cells"]["B2"].update(change);invalids.append(broken)
        broken=copy.deepcopy(wire);broken["cells"]["B2"]=[];invalids.append(broken)
        for invalid in invalids:
            with self.subTest(invalid=invalid):
                with self.assertRaises(ValueError):normalize_response(payload,invalid)

    def test_classify_old_arrays_remain_compatible_and_other_tasks_unchanged(self):
        from 附注更新.回执格式 import normalize_response,transport_instruction
        payload,wire=self.classify_fixture();legacy=normalize_response(payload,wire)
        self.assertEqual(normalize_response(payload,legacy),legacy)
        # 旧协议的跨数组重复不由转换器吞掉，继续交原校验器报错。
        repeated=copy.deepcopy(legacy);repeated["excluded"].append({"cell":"B2","category":"header","reason":"重复"})
        self.assertEqual(normalize_response(payload,repeated),repeated)
        for task in ("layout","source_scope","tables",None):
            other={"task":task};value={"tables":[]}
            self.assertEqual(normalize_response(other,value),value)
            if task!="layout":self.assertEqual(transport_instruction(other),"")
        instruction=transport_instruction(payload)
        self.assertIn('cells',instruction);self.assertIn('kind',instruction)
        self.assertIn('candidate_cells',instruction)
        self.assertEqual(instruction,transport_instruction({**payload,"correction":{"validation_error":"重复格"}}))

    def test_classify_valid_shape_still_needs_original_business_validation(self):
        from 附注更新.回执格式 import normalize_response
        from 附注更新.语义 import SemanticEngine,SemanticError
        payload={"task":"classify","candidate_cells":["B2"],"gold_slots":[{"id":"C-N001-T001-S001"}],"cells":{"B2":{"value":123}}}
        engine=SemanticEngine.__new__(SemanticEngine)
        slot={"id":"C-N001-T001-S001","scope":"consolidated","value_type":"monetary","slot":{"name":"银行存款"},"row_path":["银行存款"],"column_path":[],"dimensions":[]}
        engine.slots={slot["id"]:slot}
        sheet={"name":"附注","cells":payload["cells"]}
        table={"range":"A1:B2","scope":"consolidated","table_semantic":"货币资金"}
        wire={"cells":{"B2":{"kind":"mapping","reason":"原表证据"}}}
        with self.assertRaises(ValidationError):self.schema(payload).validate(wire)
        with self.assertRaises(ValueError):normalize_response(payload,wire)
        wire["cells"]["B2"].update(slot_id=slot["id"],scope="consolidated",dimensions={},semantic_field="银行存款",value_type="monetary",row_label="银行存款",column_label="期末余额")
        self.schema(payload).validate(wire)
        result=engine._validate_classification(normalize_response(payload,wire),["B2"],{slot["id"]},sheet,table)
        self.assertEqual(result["B2"][0],"mappings")
        wire["cells"]["B2"]={"kind":"excluded","reason":"误把金额当表头","category":"header"}
        self.schema(payload).validate(wire)
        with self.assertRaises(SemanticError):engine._validate_classification(normalize_response(payload,wire),["B2"],{slot["id"]},sheet,table)

    def test_classify_schema_size_shares_record_and_does_not_change_payload(self):
        payload,wire=self.classify_fixture();before=copy.deepcopy(payload)
        first=response_format(payload)
        self.assertEqual(first,response_format({**payload,"correction":{"validation_error":"重复格"}}))
        self.assertEqual(payload,before)
        many={**payload,"candidate_cells":[f"B{n}" for n in range(1,401)]}
        schema=response_format(many)["json_schema"]["schema"]
        self.assertEqual(len(schema["$defs"]),1)
        self.assertEqual(json.dumps(schema).count('\"oneOf\"'),1)
        self.assertEqual(len(schema['$defs']['cell_record']['oneOf']),3)
        self.assertLess(len(json.dumps(schema)),30000)

    def layout_fixture(self):
        from 附注更新.语义 import SemanticEngine
        cells={"A1":{"value":"项目","formula":None},"B1":{"value":"期末余额","formula":None},
               "A2":{"value":"库存现金","formula":None},"B2":{"value":123,"formula":None}}
        payload={"task":"layout","notes":[{"id":"C-N001","name":"货币资金"}],"cells":cells,
                 "candidate_cells":list(cells),"max_row":2,"max_column":2}
        table={"range":"A1:B2","note_ids":["C-N001"],"scope":"standalone","table_semantic":"货币资金明细","header_cells":["A1","B1"]}
        engine=SemanticEngine.__new__(SemanticEngine);engine.notes={"C-N001":payload["notes"][0]};engine.report_scope="standalone"
        sheet={"name":"附注","max_row":2,"max_column":2,"cells":cells}
        return payload,table,engine,sheet

    def test_new_layout_http_requires_three_arrays_but_old_receipt_stays_readable(self):
        from 附注更新.回执格式 import normalize_response
        from 附注更新.语义 import SemanticError
        payload,table,engine,sheet=self.layout_fixture();validator=self.schema(payload)
        complete={"tables":[table],"non_business":[],"unresolved":[]}
        validator.validate(complete)
        for field in ("tables","non_business","unresolved"):
            broken=copy.deepcopy(complete);del broken[field]
            with self.subTest(field=field):
                with self.assertRaises(ValidationError):validator.validate(broken)
        legacy={"tables":[table]};before=copy.deepcopy(legacy)
        normalized=normalize_response(payload,legacy)
        self.assertEqual(normalized,before);self.assertEqual(legacy,before)
        engine._validate_layout(normalized,sheet,payload["candidate_cells"])
        with self.assertRaises(SemanticError):
            engine._validate_layout(normalize_response(payload,{"tables":[]}),sheet,payload["candidate_cells"])

    def test_layout_requires_real_nonempty_evidence_and_valid_table_fields(self):
        payload,table,engine,sheet=self.layout_fixture();before=copy.deepcopy(payload);validator=self.schema(payload)
        valid={"tables":[table],"non_business":[],"unresolved":[],"carry_context":"人民币元"}
        validator.validate(valid);engine._validate_layout(valid,sheet,payload["candidate_cells"])
        self.assertEqual(payload,before)
        self.assertEqual(response_format(payload),response_format({**payload,"correction":{"validation_error":"缺少证据"}}))
        invalid={"tables":[table],"non_business":[{"range":"A1:B2","category":"empty_padding","reason":"候选格之外的空白区域","evidence_cells":[]}],"unresolved":[]}
        with self.assertRaises(ValidationError):validator.validate(invalid)
        for change in ({"evidence_cells":["Z99"]},{"reason":""},{"category":"not_business"}):
            altered=copy.deepcopy(invalid);altered["non_business"][0].update(evidence_cells=["A1"]);altered["non_business"][0].update(change)
            with self.assertRaises(ValidationError):validator.validate(altered)
        for change in ({"note_ids":["S-N999"]},{"note_ids":[]},{"scope":"single"},{"header_cells":["Z99"]},{"table_semantic":""}):
            altered=copy.deepcopy(valid);altered["tables"][0].update(change)
            with self.assertRaises(ValidationError):validator.validate(altered)

    def test_layout_new_schema_requires_one_note_without_rewriting_old_layout(self):
        from 附注更新.回执格式 import normalize_response
        payload,table,engine,sheet=self.layout_fixture()
        payload["notes"].append({"id":"C-N002","name":"应付账款"})
        validator=self.schema(payload);validator.validate({"tables":[table],"non_business":[],"unresolved":[]})
        for ids in ([],["C-N001","C-N002"],["C-N001","C-N001"]):
            old={"tables":[{**table,"note_ids":ids}]};before=copy.deepcopy(old)
            with self.subTest(note_ids=ids):
                with self.assertRaises(ValidationError):validator.validate({**old,"non_business":[],"unresolved":[]})
                # 旧保存证据不受新请求schema改写，继续交原业务校验器解释。
                self.assertEqual(normalize_response(payload,old),before)
                self.assertEqual(old,before)

    def test_layout_unresolved_regions_require_reason_and_real_evidence(self):
        from 附注更新.回执格式 import normalize_response
        payload,table,engine,sheet=self.layout_fixture();validator=self.schema(payload)
        valid={"tables":[],"non_business":[{"range":"A1:B1","category":"header","reason":"原表列头","evidence_cells":["A1","B1"]}],
               "unresolved":[{"range":"A2:B2","reason":"业务已见，但未确定金标准科目","evidence_cells":["A2"]}]}
        before=copy.deepcopy(valid);validator.validate(valid)
        self.assertEqual(normalize_response(payload,valid),before)
        self.assertEqual(valid,before)
        invalids=[]
        for key in ("range","reason","evidence_cells"):
            broken=copy.deepcopy(valid);del broken["unresolved"][0][key];invalids.append(broken)
        for change in ({"range":""},{"reason":""},{"evidence_cells":[]},
                       {"evidence_cells":["Z99"]},{"note_ids":["C-N001"]}):
            broken=copy.deepcopy(valid);broken["unresolved"][0].update(change);invalids.append(broken)
        for broken in invalids:
            with self.subTest(broken=broken):
                with self.assertRaises(ValidationError):validator.validate(broken)
        # 旧布局不要求补出新字段；区域覆盖仍由原业务校验器决定。
        for legacy in ({"tables":[table]},{"tables":[],"non_business":[]}):
            self.assertEqual(normalize_response(payload,legacy),legacy)
        empty=self.schema({"task":"layout","notes":[],"cells":{}})
        empty.validate({"tables":[],"non_business":[],"unresolved":[]})
        with self.assertRaises(ValidationError):
            empty.validate({"tables":[],"non_business":[],"unresolved":[{"range":"A1","reason":"无原格证据","evidence_cells":[]}]})

    def test_layout_transport_explains_complete_three_way_ownership(self):
        from 附注更新.回执格式 import transport_instruction
        payload,table,engine,sheet=self.layout_fixture()
        instruction=transport_instruction(payload)
        self.assertIn("三个数组，均可为空但不能省略",instruction)
        self.assertNotIn("可省略",instruction)
        for field in ("tables","non_business","unresolved","range","reason","evidence_cells","candidate_cells"):
            self.assertIn(field,instruction)
        self.assertEqual(instruction,transport_instruction({**payload,"correction":{"validation_error":"遗漏"}}))

    def test_layout_schema_valid_evidence_cannot_hide_numeric_cells(self):
        from 附注更新.语义 import SemanticError
        payload,table,engine,sheet=self.layout_fixture();validator=self.schema(payload)
        # 格式完整且证据真实，仍不能把含金额的原区段称为留白。
        invalid={"tables":[],"non_business":[{"range":"A1:B2","category":"empty_padding","reason":"误判为空白","evidence_cells":["A1"]}],"unresolved":[]}
        validator.validate(invalid)
        with self.assertRaises(SemanticError):engine._validate_layout(invalid,sheet,payload["candidate_cells"])

    def test_layout_with_no_displayed_cells_cannot_generate_nonbusiness_regions(self):
        validator=self.schema({"task":"layout","notes":[],"cells":{}})
        validator.validate({"tables":[],"non_business":[],"unresolved":[]})
        with self.assertRaises(ValidationError):validator.validate({"tables":[],"non_business":[{"range":"A1","category":"annotation","reason":"无真实原格","evidence_cells":[]}],"unresolved":[]})

    def fixed_layout_fixture(self):
        payload={"task":"layout_cells","candidate_cells":["B2","C2","A2","D2"],
            "cells":{"A2":{"value":"本期增加"},"B2":{"value":10},"C2":{"value":20},"D2":{"value":30}},
            "context":{"preceding_cells":{"B1":{"value":"实收资本"},"C1":{"value":"资本公积"}}},
            "notes":[{"id":"C-N105"},{"id":"C-N091"}]}
        wire={"cells":{
            "B2":{"kind":"business","reason":"原列为实收资本","note_id":"C-N105","scope":"standalone","header_cells":["B1","A2"],"table_semantic":"实收资本本期增加"},
            "C2":{"kind":"business","reason":"相邻列为资本公积","note_id":"C-N091","scope":"standalone","header_cells":["C1","A2"],"table_semantic":"资本公积本期增加"},
            "A2":{"kind":"non_business","reason":"行标签","category":"label","evidence_cells":["A2"]},
            "D2":{"kind":"unresolved","reason":"缺少本列指标表头","evidence_cells":["D2"]}}}
        return payload,wire

    def test_layout_cells_keeps_adjacent_notes_and_original_shared_evidence(self):
        from 附注更新.回执格式 import normalize_response
        payload,wire=self.fixed_layout_fixture();before=copy.deepcopy((payload,wire))
        validator=self.schema(payload);validator.validate(wire)
        result=normalize_response(payload,wire)
        self.assertEqual([t["range"] for t in result["tables"]],["B2","C2"])
        self.assertEqual([t["note_ids"] for t in result["tables"]],[["C-N105"],["C-N091"]])
        self.assertEqual(result["tables"][0]["header_cells"],["B1","A2"])
        self.assertEqual(result["non_business"],[{"range":"A2","reason":"行标签","category":"label","evidence_cells":["A2"]}])
        self.assertEqual(result["unresolved"],[{"range":"D2","reason":"缺少本列指标表头","evidence_cells":["D2"]}])
        self.assertEqual((payload,wire),before)
        result["tables"][0]["header_cells"].append("C1")
        self.assertEqual((payload,wire),before)
        schema=validator.schema
        self.assertEqual(list(schema["properties"]),["cells"])
        self.assertEqual(set(schema["properties"]["cells"]["required"]),set(payload["candidate_cells"]))
        self.assertEqual(json.dumps(schema).count('"oneOf"'),1)
        for branch in next(iter(schema["$defs"].values()))["oneOf"]:
            self.assertEqual(next(iter(branch["properties"])),"kind")

    def test_layout_cells_rejects_missing_outside_and_wrong_kind_fields(self):
        from 附注更新.回执格式 import normalize_response
        payload,wire=self.fixed_layout_fixture();validator=self.schema(payload);invalids=[]
        missing=copy.deepcopy(wire);del missing["cells"]["C2"];invalids.append(missing)
        outside=copy.deepcopy(wire);outside["cells"]["Z9"]=outside["cells"]["D2"];invalids.append(outside)
        invalids.extend([{"cells":{}},{"tables":[],"non_business":[],"unresolved":[]},{**wire,"tables":[]}])
        for address in wire["cells"]:
            for key in wire["cells"][address]:
                broken=copy.deepcopy(wire);del broken["cells"][address][key];invalids.append(broken)
        for address,change in [("B2",{"note_id":"C-N999"}),("B2",{"scope":"single"}),
            ("B2",{"header_cells":["Z99"]}),("B2",{"header_cells":[]}),("B2",{"header_cells":"B1"}),
            ("B2",{"table_semantic":7}),("B2",{"value":123}),("B2",{"range":"B2:C2"}),
            ("B2",{"kind":["business"]}),("A2",{"category":"amount"}),
            ("A2",{"note_id":"C-N105"}),("D2",{"evidence_cells":[]}),
            ("D2",{"evidence_cells":["Z99"]}),("D2",{"scope":"standalone"})]:
            broken=copy.deepcopy(wire);broken["cells"][address].update(change);invalids.append(broken)
        for broken in invalids:
            with self.subTest(broken=broken):
                with self.assertRaises(ValidationError):validator.validate(broken)
                with self.assertRaises(ValueError):normalize_response(payload,broken)
        broken=copy.deepcopy(wire);broken["cells"]["B2"]["reason"]=" "
        with self.assertRaises(ValueError):normalize_response(payload,broken)

    def test_layout_cells_rejects_repeated_candidates_and_retains_json_duplicate_guard(self):
        from 附注更新.回执格式 import normalize_response
        from 附注更新.语义 import _unique_json_object,SemanticError
        payload,wire=self.fixed_layout_fixture()
        for candidates in (["B2","B2"],["Z99"],["B2",None]):
            invalid={**payload,"candidate_cells":candidates}
            with self.assertRaises(ValueError):response_format(invalid)
            with self.assertRaises(ValueError):normalize_response(invalid,wire)
        with self.assertRaises(SemanticError):
            json.loads('{"cells":{"B2":{"kind":"unresolved"},"B2":{"kind":"business"}}}',object_pairs_hook=_unique_json_object)
        empty={"task":"layout_cells","candidate_cells":[],"cells":{},"notes":[]}
        self.schema(empty).validate({"cells":{}})
        self.assertEqual(normalize_response(empty,{"cells":{}}),{"tables":[],"non_business":[],"unresolved":[]})
        no_notes={**payload,"notes":[]}
        with self.assertRaises(ValidationError):self.schema(no_notes).validate(wire)
        with self.assertRaises(ValueError):normalize_response(no_notes,wire)

    def test_layout_cells_format_cannot_approve_amount_exclusion_or_report_scope(self):
        from 附注更新.回执格式 import normalize_response
        from 附注更新.语义 import SemanticEngine,SemanticError
        payload,wire=self.fixed_layout_fixture()
        engine=SemanticEngine.__new__(SemanticEngine);engine.report_scope="standalone"
        engine.notes={n["id"]:n for n in payload["notes"]}
        sheet={"name":"权益表","max_row":2,"max_column":4,
            "cells":{**payload["context"]["preceding_cells"],**payload["cells"]}}
        result=normalize_response(payload,wire)
        engine._validate_layout(result,sheet,payload["candidate_cells"])
        wrong=copy.deepcopy(wire);wrong["cells"]["B2"]={"kind":"non_business","reason":"误把金额当标题","category":"header","evidence_cells":["B1"]}
        self.schema(payload).validate(wrong)
        with self.assertRaises(SemanticError):engine._validate_layout(normalize_response(payload,wrong),sheet,payload["candidate_cells"])
        wrong=copy.deepcopy(wire);wrong["cells"]["B2"]["scope"]="parent"
        self.schema(payload).validate(wrong)
        with self.assertRaises(SemanticError):engine._validate_layout(normalize_response(payload,wrong),sheet,payload["candidate_cells"])

    def test_layout_cells_instruction_preserves_old_layout_and_correction_identity(self):
        from 附注更新.回执格式 import normalize_response,transport_instruction
        payload,wire=self.fixed_layout_fixture();before=copy.deepcopy(payload)
        instruction=transport_instruction(payload)
        for term in ["candidate_cells","不生成矩形","不输出金额","business","non_business","unresolved","note_id","preceding_cells"]:
            self.assertIn(term,instruction)
        corrected={**payload,"correction":{"validation_error":"科目错误","previous_result":wire}}
        self.assertEqual(response_format(payload),response_format(corrected))
        self.assertEqual(instruction,transport_instruction(corrected));self.assertEqual(payload,before)
        legacy={"tables":[]}
        self.assertEqual(normalize_response({"task":"layout"},legacy),legacy)
        self.assertEqual(response_format({"task":"layout","notes":[],"cells":{}})["json_schema"]["name"],"reply_layout")

    def test_complex_tasks_keep_json_object_mode(self):
        for task in ["missing_semantics_review","structure","review_unused_sources","future_task",None]:
            self.assertEqual(response_format({"task":task}),{"type":"json_object"})

    def test_format_does_not_change_payload_or_correction_identity(self):
        payload=scope_payload();before=copy.deepcopy(payload);first=response_format(payload)
        self.assertEqual(payload,before)
        self.assertEqual(first,response_format({**payload,"correction":{"previous_result":{},"validation_error":"分类错"}}))
        self.assertNotIn("response_format",payload)

    def test_each_call_owns_its_schema_and_does_not_embed_financial_text(self):
        payload=scope_payload();payload["cells"]["A1"]["text"]="不得复制的业务名称与金额1234567"
        result=response_format(payload);result["json_schema"]["schema"]["properties"].clear()
        self.assertIn("areas",response_format(payload)["json_schema"]["schema"]["properties"])
        self.assertNotIn("1234567",json.dumps(response_format(payload),ensure_ascii=False))

    def test_empty_catalogue_and_evidence_use_valid_empty_arrays_not_empty_enum(self):
        self.schema({"task":"tables","catalogue":[]}).validate({"table_ids":[]})
        validator=self.schema({"task":"source_scope","cells":{},"context_cells":{}})
        validator.validate({"areas":[{"range":"A1","decision":"uncertain","reason":"缺原文","evidence_cells":[]}]})
        with self.assertRaises(ValidationError):validator.validate({"areas":[{"range":"A1","decision":"uncertain","reason":"缺原文","evidence_cells":["A1"]}]})

    def test_schema_does_not_replace_cross_packet_business_validation(self):
        first=scope_payload();first["packet_id"]="1"
        second={"packet_id":"2","cells":{"Z1":{"kind":"text","text":"另一原表"}},"context_cells":{}}
        validator=self.schema({"task":"source_scope_batch","tasks":[first,second]})
        # 精简schema只限制本批文字地址的并集，原校验器仍须逐包禁止跨任务引用。
        validator.validate({"results":[{"packet_id":"1","areas":[{"range":"A1","decision":"uncertain","reason":"待核实","evidence_cells":["Z1"]}]},{"packet_id":"2","areas":[]}]})



    def test_classify_kind_branches_require_real_fields_without_filling_missing_data(self):
        from 附注更新.回执格式 import normalize_response,transport_instruction
        payload,valid=self.classify_fixture()
        valid["cells"]["B2"].update(row_label="银行存款",column_label="期末余额")
        validator=self.schema(payload)
        validator.validate(valid)
        required=("slot_id","scope","dimensions","semantic_field","value_type","row_label","column_label")
        broken_samples=[]
        for key in required:
            broken=copy.deepcopy(valid);del broken["cells"]["B2"][key];broken_samples.append(broken)
        # 真实首次回执同类错误：已给ID和类型，却未提供口径及维度。
        actual_shape=copy.deepcopy(valid)
        actual_shape["cells"]["B2"]={"kind":"mapping","reason":"原表证据","slot_id":"C-N001-T001-S001","value_type":"monetary"}
        broken_samples.append(actual_shape)
        missing_category=copy.deepcopy(valid);del missing_category["cells"]["A1"]["category"];broken_samples.append(missing_category)
        for broken in broken_samples:
            before=copy.deepcopy(broken)
            with self.subTest(broken=broken):
                with self.assertRaises(ValidationError):validator.validate(broken)
                with self.assertRaises(ValueError):normalize_response(payload,broken)
                self.assertEqual(broken,before)
        for address,field,value in (("B2","category","header"),("A1","slot_id","C-N001-T001-S001"),("C3","dimensions",{})):
            broken=copy.deepcopy(valid);broken["cells"][address][field]=value
            with self.assertRaises(ValidationError):validator.validate(broken)
            with self.assertRaises(ValueError):normalize_response(payload,broken)
        instruction=transport_instruction(payload)
        self.assertIn("必填",instruction)
        for field in required:self.assertIn(field,instruction)


class UnifiedFormatTests(unittest.TestCase):
    schema = ResponseFormatTests.schema

    def test_shared_dimension_values_restore_each_cell_without_defaults_or_overrides(self):
        from 附注更新.回执格式 import normalize_response
        from 附注更新.统一识别 import build_payload, validate_response
        from 附注更新.统一语义 import validate_records
        from 测试.测试统一识别 import packet, answer
        from 测试.测试统一语义 import records
        source=packet();standard=validate_records(records());payload=build_payload(source,standard)
        payload['response_contract']='unified_dimension_refs_v1'
        expected=answer();wire=copy.deepcopy(expected);pool=[];seen={}
        for row in wire['cells'].values():
            refs=[]
            for key,value in row.pop('dimensions').items():
                entry={'dimension_id':key,'value':value,'evidence':row['dimension_evidence'][key]}
                signature=json.dumps(entry,sort_keys=True)
                if signature not in seen:seen[signature]=len(pool);pool.append(entry)
                refs.append(seen[signature])
            row.pop('dimension_evidence');row['dimension_refs']=refs
        wire['dimension_values']=pool;before=copy.deepcopy(wire)
        self.schema(payload).validate(wire)
        self.assertEqual(normalize_response(payload,wire),expected)
        self.assertEqual(validate_response(normalize_response(payload,wire),source,standard),validate_response(expected,source,standard))
        self.assertEqual(wire,before)
        address=payload['candidate_cells'][0]
        named=copy.deepcopy(wire)
        for row in named['cells'].values():
            row['dimension_refs']={pool[index]['dimension_id']:index for index in row['dimension_refs']}
        named_before=copy.deepcopy(named)
        self.assertEqual(normalize_response(payload,named),expected)
        self.assertEqual(validate_response(normalize_response(payload,named),source,standard),validate_response(expected,source,standard))
        self.assertEqual(named,named_before)
        key=next(iter(named['cells'][address]['dimension_refs']))
        for invalid in (-1,True,0.5,len(pool),'0'):
            broken=copy.deepcopy(named);broken['cells'][address]['dimension_refs'][key]=invalid
            with self.subTest(named_invalid=invalid),self.assertRaises(ValueError):normalize_response(payload,broken)
        mismatch=copy.deepcopy(named)
        mismatch['cells'][address]['dimension_refs']['错误维度']=mismatch['cells'][address]['dimension_refs'].pop(key)
        with self.assertRaises(ValueError):normalize_response(payload,mismatch)
        for invalid in (-1,True,0.5,len(pool)):
            broken=copy.deepcopy(wire);broken['cells'][address]['dimension_refs'][0]=invalid
            with self.subTest(invalid=invalid),self.assertRaises(ValueError):normalize_response(payload,broken)
        duplicate=copy.deepcopy(wire);duplicate['cells'][address]['dimension_refs'].append(0)
        with self.assertRaises(ValueError):normalize_response(payload,duplicate)
        missing=copy.deepcopy(wire);missing['cells'][address]['dimension_refs'].pop()
        with self.assertRaises(ValueError):validate_response(normalize_response(payload,missing),source,standard)
        partial=copy.deepcopy(wire)
        entity_ref=next(i for i,entry in enumerate(pool) if entry['dimension_id']=='entity')
        for row in partial['cells'].values():
            row['dimension_refs'].remove(entity_ref)
            row.update(kind='partial_metric_fact',missing_dimensions=['entity'],evidence_cells=['A1'])
        self.schema(payload).validate(partial)
        checked=validate_response(normalize_response(payload,partial),source,standard)
        self.assertTrue(all(row['kind']=='partial_metric_fact' for row in checked.values()))
        broken_samples=[]
        foreign=copy.deepcopy(wire);foreign['cells']['Z99']=foreign['cells'][address];broken_samples.append(foreign)
        unknown=copy.deepcopy(wire);unknown['dimension_values'][0]['dimension_id']='未知维度';broken_samples.append(unknown)
        evidence=copy.deepcopy(wire);evidence['dimension_values'][0]['evidence']=['Z99'];broken_samples.append(evidence)
        mixed=copy.deepcopy(wire);mixed['cells'][address]['dimensions']={};broken_samples.append(mixed)
        root=copy.deepcopy(wire);root['extra']=True;broken_samples.append(root)
        for broken in broken_samples:
            with self.assertRaises(ValueError):normalize_response(payload,broken)
        self.assertEqual(normalize_response(payload,expected),expected)

    def test_unified_format_shares_evidence_and_dimension_definitions_without_changing_payload(self):
        from 附注更新.统一识别 import build_payload
        from 附注更新.统一语义 import validate_records
        from 测试.测试统一识别 import packet
        from 测试.测试统一语义 import records
        payload=build_payload(packet(),validate_records(records()))
        metric=payload['metrics'][0]
        payload['metrics']=[{**copy.deepcopy(metric),'id':f'metric.test_{i}'} for i in range(20)]
        payload['cells'].update({f'Z{i}':{'value':'合成表头'} for i in range(1,302)})
        before=copy.deepcopy(payload)
        result=response_format(payload)
        encoded=json.dumps(result,ensure_ascii=False)
        self.assertEqual(encoded.count('"Z301"'),1,'相同证据地址目录应只保存一次')
        self.assertEqual(payload,before)
        Draft202012Validator.check_schema(result['json_schema']['schema'])
        # 各请求须各有自己的定义，不能被先前请求的修改污染。
        result['json_schema']['schema']['$defs'].clear()
        self.assertEqual(response_format(payload)['json_schema']['schema']['properties']['cells']['required'],payload['candidate_cells'])

    def test_unified_dimensions_use_declared_shapes_without_values_or_locations_as_meaning(self):
        from 附注更新.统一识别 import build_payload
        from 附注更新.统一语义 import validate_records
        from 测试.测试统一识别 import packet, answer
        from 测试.测试统一语义 import records
        payload=build_payload(packet(),validate_records(records()))
        validator=self.schema(payload); valid=answer(); validator.validate(valid)
        for change in ('missing_domain','string_selection','wrong_metric','financial_value','foreign_dimension','foreign_evidence'):
            bad=copy.deepcopy(valid); row=bad['cells']['D7']
            if change=='missing_domain':del row['dimensions']['asset_selection']['domain']
            elif change=='string_selection':row['dimensions']['asset_selection']='全部'
            elif change=='wrong_metric':row['metric_id']='不存在的指标'
            elif change=='financial_value':row['raw_value']=123
            elif change=='foreign_dimension':row['dimensions']['unknown_dimension']='分类'
            else:row['dimension_evidence']['entity']=['Z99']
            with self.subTest(change=change), self.assertRaises(ValidationError):validator.validate(bad)
        unresolved={"cells":{a:{"kind":"unresolved","reason":"证据不足","evidence_cells":["A1"]} for a in payload['candidate_cells']}}
        validator.validate(unresolved)


if __name__=="__main__":unittest.main()
