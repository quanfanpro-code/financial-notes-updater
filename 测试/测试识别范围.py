"""更新数据范围筛选的模拟回执测试，不代表真实模型准确率。"""
import copy
import json
import unittest
from 附注更新.识别范围 import build_target_meanings, select_source_scope
from 附注更新.语义 import SemanticCancelled


def sample():
    values={"A1":"货币资金", "A2":"项目", "B2":"期末余额", "A3":"新增银行账户", "B3":123,
            "A5":"企业人力资源统计", "A6":"人员类别", "B6":9, "A7":"年末人数", "B7":8}
    return {"sha256":"source-1", "sheets":[{"name":"混合资料", "max_row":7,"max_column":2,"merges":[],"cells":{
        a:{"row":int(a[1:]),"column":ord(a[0])-64,"value":v,"formula":None} for a,v in values.items()}}]}


class Stub:
    gold_hash="gold-1"
    def __init__(self,callback):self.callback=callback;self.calls=[]
    def _check_cancel(self):pass
    def _request(self,system,payload,validator):
        self.calls.append(copy.deepcopy(payload))
        return validator(self.callback(payload))


def area(range,decision,evidence):
    return {"range":range,"decision":decision,"reason":"依据实际业务文字确定是否属于本次披露", "evidence_cells":evidence}



class CellStub(Stub):
    def _request(self,system,payload,validator):
        from 附注更新.回执格式 import normalize_response
        self.calls.append(copy.deepcopy(payload))
        return validator(normalize_response(payload,self.callback(payload)))


def cell_reply(payload):
    def single(p):
        return {"cells":{a:{"decision":"include","reason":"依据保留的完整原表文字判断为本次业务","evidence_cells":["A1"]}
                         for a in p["candidate_cells"]}}
    return {"packets":{t["packet_id"]:single(t) for t in payload["tasks"]}} if payload["task"]=="source_scope_batch" else single(payload)


class SourceScopeTests(unittest.TestCase):
    def test_mixed_sheet_preserves_new_objects_and_excludes_other_business(self):
        source=sample();before=copy.deepcopy(source)
        engine=Stub(lambda p:{"areas":[area("A1:B3","include",["A1"]),area("A5:B7","out_of_scope",["A5"])]})
        result=select_source_scope(engine,source,{"chapter_title":"财务报表重要项目的说明","disclosures":["货币资金"]})
        self.assertEqual(source,before)
        self.assertIn("B3",result["selected_cells"]["混合资料"])
        self.assertNotIn("B6",result["selected_cells"]["混合资料"])
        self.assertEqual({a for r in result["out_of_scope"] for a in r["cells"]},{"A5","A6","B6","A7","B7"})
        self.assertEqual(len(engine.calls),2)
        self.assertEqual(result["source_hash"],"source-1")
        self.assertEqual(result["gold_hash"],"gold-1")
        self.assertEqual(len(result["target_hash"]),64)
        self.assertNotIn('123',json.dumps(engine.calls,ensure_ascii=False))

    def test_matching_business_cannot_be_removed_because_it_is_a_primary_statement(self):
        source=sample()
        source["sheets"][0]["cells"]["A1"]["value"]="一、货币资金"
        engine=Stub(lambda p:{"areas":[area("A1:B3","out_of_scope",["A1"]),area("A5:B7","out_of_scope",["A5"])]})
        result=select_source_scope(engine,source,{"disclosures":["货币资金"]})
        self.assertIn("B3",result["selected_cells"]["混合资料"])
        self.assertNotIn("B7",result["selected_cells"]["混合资料"])
        self.assertTrue(result["unresolved"])
        self.assertTrue(any("同名业务" in a["reason"] for a in result["areas"]))

    def test_explicit_outer_row_or_column_evidence_protects_only_referenced_area(self):
        for evidence,label in [("B2","固定资产减值准备"),("C1","信用减值损失")]:
            with self.subTest(evidence=evidence):
                values={evidence:label,"C2":1,"D2":2,"C3":3,"D3":4,
                        "B5":"人力资源统计","C5":5,"D5":6,"C6":7,"D6":8}
                source={"sha256":"outer-header","sheets":[{"name":"混合","max_row":6,"max_column":4,"merges":[],
                        "cells":{a:{"value":v,"formula":None} for a,v in values.items()}}]}
                engine=Stub(lambda p:{"areas":[area("C2:D3","out_of_scope",[evidence]),area("C5:D6","out_of_scope",["B5"])]})
                result=select_source_scope(engine,source,{"table_business_terms":[label]})
                self.assertIn("C2",result["selected_cells"]["混合"])
                self.assertNotIn("C5",result["selected_cells"]["混合"])
                self.assertTrue(all(a["decision"]=="uncertain" for a in result["areas"] if a["range"]=="C2:D3"))
                self.assertTrue(all(a["decision"]=="out_of_scope" for a in result["areas"] if a["range"]=="C5:D6"))

    def test_many_small_sheets_share_requests_without_crossing_evidence(self):
        source=sample()
        source["sheets"]=[{**copy.deepcopy(source["sheets"][0]),"name":"资料"+str(i)} for i in range(20)]
        def reply(payload):
            return {"results":[{"packet_id":p["packet_id"],"areas":[area("A1:B3","include",["A1"]),area("A5:B7","out_of_scope",["A5"])]} for p in payload["tasks"]]}
        engine=Stub(reply)
        result=select_source_scope(engine,source,{"disclosures":["货币资金"]})
        self.assertLess(len(engine.calls),10)
        self.assertEqual(result["coverage"]["candidate_count"],200)
        self.assertEqual(result["coverage"]["selected_count"],100)
        for selected in result["selected_cells"].values():
            self.assertIn("B3",selected)
            self.assertNotIn("B7",selected)

    def test_disagreement_or_missing_coverage_stays_selected(self):
        engine=Stub(lambda p:{"areas":[area("A1:B7","out_of_scope" if p["round"]==1 else "include",["A1"])]})
        result=select_source_scope(engine,sample(),{"disclosures":["货币资金"]})
        self.assertEqual(set(result["selected_cells"]["混合资料"]),set(sample()["sheets"][0]["cells"]))
        self.assertFalse(result["out_of_scope"])
        self.assertTrue(result["unresolved"])
        blank=select_source_scope(Stub(lambda p:{"areas":[]}),sample(),{"disclosures":["货币资金"]})
        self.assertEqual(len(blank["selected_cells"]["混合资料"]),10)

    def test_exclusion_requires_existing_presented_text_evidence(self):
        for evidence in (["Z999"],["B3"],[]):
            with self.subTest(evidence=evidence):
                result=select_source_scope(Stub(lambda p:{"areas":[area("A1:B7","out_of_scope",evidence)]}),sample(),{"disclosures":["货币资金"]})
                self.assertFalse(result["out_of_scope"])
                self.assertEqual(len(result["selected_cells"]["混合资料"]),10)
                self.assertTrue(result["unresolved"])

    def test_all_cells_all_sheets_are_catalogued_with_bounded_packets(self):
        source=sample()
        cells={}
        for row in range(1,401):
            cells[f"A{row}"]={"row":row,"column":1,"value":"其他披露业务标签"*10+str(row),"formula":None}
            cells[f"B{row}"]={"row":row,"column":2,"value":row,"formula":None}
        source["sheets"].append({"name":"隐藏资料", "hidden":True,"max_row":400,"max_column":2,"merges":["C2:D2"],"cells":cells})
        engine=Stub(lambda p:{"areas":[]})
        result=select_source_scope(engine,source,{"disclosures":["货币资金"]})
        seen={(p["sheet"],a) for p in engine.calls for a in p["candidate_cells"]}
        expected={(s["name"],a) for s in source["sheets"] for a in s["cells"]}
        self.assertEqual(seen,expected)
        self.assertEqual(sum(len(v) for v in result["selected_cells"].values()),len(expected))
        from 附注更新.识别范围 import _source_chars
        self.assertTrue(all(_source_chars(p)<=12000 for p in engine.calls))
        self.assertTrue(any(p.get("hidden") for p in engine.calls))

    def test_cross_packet_include_overrides_previous_exclusion(self):
        cells={f"A{i}":{"row":i,"column":1,"value":"披露业务文字"*35+str(i),"formula":None} for i in range(1,151)}
        source={"sha256":"large", "sheets":[{"name":"混合", "max_row":150,"max_column":1,"cells":cells,"merges":[]}]}
        def reply(p):
            decision="out_of_scope" if p["packet"]==1 else "include"
            return {"areas":[area("A1:A150",decision,[p["candidate_cells"][0]])]}
        result=select_source_scope(Stub(reply),source,{"disclosures":["货币资金"]})
        self.assertFalse(result["out_of_scope"])
        self.assertEqual(len(result["selected_cells"]["混合"]),150)

    def test_target_builder_keeps_business_text_not_amounts_or_paths(self):
        context={"scope":{"chapter_title":"财务报表重要项目的说明"},"source":"SECRET_PATH",
                 "tables":[{"title":"货币资金","chapter_context":["第六章","单位：人民币元"],"source":"SECRET_PATH","cells":{"A1":987654321}},
                           {"title":"续：","chapter_context":["第六章","单位：人民币元","坏账准备"],"first_row":25}]}
        result=build_target_meanings(context)
        encoded=json.dumps(result,ensure_ascii=False)
        self.assertIn("货币资金",encoded);self.assertIn("坏账准备",encoded)
        self.assertNotIn("SECRET_PATH",encoded);self.assertNotIn("987654321",encoded)
        self.assertEqual(result["chapter_context"].count("第六章"),1)

    def test_target_builder_reads_only_confirmed_word_table_text(self):
        context={"tables":[{"table_id":"word-table-1","range_name":"TA_001","sheet":"附注","title":"固定资产",
                            "first_row":1,"first_column":1,"row_count":2,"column_count":2}]}
        values={"A1":"表外不得加入","B3":"固定资产清理","C3":"信用减值损失","B4":"存货", "C4":123456789,
                "B5":"123,456.78","C5":"人民币123456.78元","B6":"=SUM(C4)","C6":r"C:\秘密\资料.xlsx",
                "B7":"B111","C7":"固定资产清理","B8":"1年以内","C8":"——"}
        source={"names":{},"sheets":[{"name":"附注","cells":{a:{"value":v,"formula":v if str(v).startswith("=") else None} for a,v in values.items()},
                  "word_table_ranges":[{"table_id":"word-table-1","range_name":"TA_001","range":"B3:C8"}]}]}
        original=copy.deepcopy(source);result=build_target_meanings(context,source)
        self.assertEqual(source,original)
        self.assertEqual(set(result["table_business"]),{"固定资产清理","信用减值损失","存货","1年以内"})
        encoded=json.dumps({k:v for k,v in result.items() if k!="word_tables"},ensure_ascii=False)
        for private in ("表外不得加入","123456789","123,456.78","123456.78","SUM(","秘密","B111","B3:C8"):
            self.assertNotIn(private,encoded)
        self.assertEqual(result["table_business"].count("固定资产清理"),1)

    def test_target_builder_requires_word_identity_and_uses_current_named_range(self):
        context={"tables":[{"table_id":"word-table-1","range_name":"TA_001","sheet":"附注","title":"固定资产",
                            "first_row":1,"first_column":1,"row_count":1,"column_count":1}]}
        source={"names":{"TA_001":[{"sheet":"附注","range":"$B$5:$C$5"}]},"sheets":[{"name":"附注","cells":{
            "A1":{"value":"旧位置不取"},"B5":{"value":"固定资产清理"},"C5":{"value":200}}}]}
        self.assertEqual(build_target_meanings(context,source)["table_business"],["固定资产清理"])
        unknown=copy.deepcopy(source);unknown["names"]={}
        self.assertEqual(build_target_meanings(context,unknown)["table_business"],[])
        self.assertEqual(build_target_meanings({},source)["table_business"],[])
        self.assertNotIn("table_business",build_target_meanings(context))

    def test_word_table_evidence_preserves_each_table_headers_context_and_exact_range(self):
        tables=[{"sheet":"附注","table_id":"t1","range_name":"TA1","title":"货币资金",
                 "chapter_context":["第六章","单位：人民币元"]},
                {"sheet":"附注","table_id":"t2","range_name":"TA2","title":"应付职工薪酬",
                 "chapter_context":["第六章","职工薪酬说明"]}]
        values={"A1":"第七章不得加入","B2":"项目","B3":"库存现金","C3":987654321,
                "B4":"合计","C4":"人民币123456.78元","B8":"项目","B9":"新增薪酬指标","C9":None,
                "B10":"合计","C10":"=SECRET_SUM()","B20":"相邻章外表"}
        ranges=[{"table_id":"t1","range_name":"TA1","range":"B2:C4"},
                {"table_id":"t2","range_name":"TA2","range":"B8:C10"},
                {"table_id":"t3","range_name":"TA3","range":"B20:C22"}]
        snapshot={"sheets":[{"name":"附注","cells":{a:{"value":v,"formula":v if str(v).startswith("=") else None}
                    for a,v in values.items()},"merges":["B2:C2","B8:C8","B20:C20"],"word_table_ranges":ranges}]}
        before=copy.deepcopy((tables,snapshot));target=build_target_meanings({"tables":tables},snapshot)
        self.assertEqual(len(target["word_tables"]),2)
        for table,original,expected in zip(target["word_tables"],tables,({"B2":"项目","B3":"库存现金","B4":"合计"},
                                                                    {"B8":"项目","B9":"新增薪酬指标","B10":"合计"})):
            self.assertEqual(table["title"],original["title"])
            self.assertEqual([target["chapter_context"][i] for i in table["chapter_context_index"]],original["chapter_context"])
            self.assertEqual({a:text for row in table["text_rows"] for a,text in row["cells"]},expected)
            self.assertEqual(table["range"],next(r["range"] for r in ranges if r["table_id"]==original["table_id"]))
            self.assertEqual(table["merges"],["B2:C2"] if original["table_id"]=="t1" else ["B8:C8"])
        self.assertEqual(target["word_tables"][0]["table_id"],"t1")
        encoded=json.dumps(target,ensure_ascii=False)
        for value in ("987654321","123456.78","SECRET_SUM","相邻章外表","第七章不得加入"):
            self.assertNotIn(value,encoded)
        self.assertEqual((tables,snapshot),before)

    def test_word_table_evidence_requires_unique_physical_binding_and_does_not_fallback(self):
        context={"tables":[{"sheet":"附注","table_id":"t1","range_name":"TA1","title":"固定资产"}]}
        snapshot={"names":{"TA1":[{"sheet":"附注","range":"B5:C5"}]},"sheets":[{"name":"附注",
            "cells":{"B5":{"value":"原项目"},"C5":{"value":123}},"word_table_ranges":[]}]}
        self.assertEqual(build_target_meanings(context,snapshot)["word_tables"],[])
        del snapshot["sheets"][0]["word_table_ranges"]
        self.assertEqual(build_target_meanings(context,snapshot)["word_tables"][0]["range"],"B5:C5")
        snapshot["sheets"][0]["word_table_ranges"]=[{"table_id":"t1","range_name":"TA1","range":"B5:C5"},
            {"table_id":"t1","range_name":"TA1","range":"B8:C8"}]
        with self.assertRaisesRegex(ValueError,"唯一"):
            build_target_meanings(context,snapshot)

    def test_word_table_evidence_rejects_overlapping_bindings_for_both_range_sources(self):
        tables=[{"sheet":"附注","table_id":"t1","range_name":"TA1","title":"资金"},
                {"sheet":"附注","table_id":"t2","range_name":"TA2","title":"薪酬"}]
        snapshot={"names":{"TA1":[{"sheet":"附注","range":"A1:B2"}],"TA2":[{"sheet":"附注","range":"B2:C3"}]},
            "sheets":[{"name":"附注","cells":{"A1":{"value":"资金"},"B2":{"value":"共享格不得重复绑定"},"C3":{"value":"薪酬"}}}]}
        for physical in (False,True):
            current=copy.deepcopy(snapshot)
            if physical:
                current["sheets"][0]["word_table_ranges"]=[{"table_id":t["table_id"],"range_name":t["range_name"],
                    "range":current["names"][t["range_name"]][0]["range"]} for t in tables]
            with self.subTest(physical=physical),self.assertRaisesRegex(ValueError,"重叠"):
                build_target_meanings({"tables":tables},current)

    def test_word_named_ranges_allow_adjacent_tables_and_same_addresses_on_different_sheets(self):
        tables=[{"sheet":"附注","table_id":"t1","range_name":"TA1","title":"资金"},
                {"sheet":"附注","table_id":"t2","range_name":"TA2","title":"薪酬"}]
        snapshot={"names":{"TA1":[{"sheet":"附注","range":"A1:B2"}],"TA2":[{"sheet":"附注","range":"C1:D2"}]},
            "sheets":[{"name":"附注","cells":{"A1":{"value":"资金"},"C1":{"value":"薪酬"}}}]}
        self.assertEqual(len(build_target_meanings({"tables":tables},snapshot)["word_tables"]),2)
        tables[1]["sheet"]="其他附注";snapshot["names"]["TA2"]=[{"sheet":"其他附注","range":"A1:B2"}]
        snapshot["sheets"].append({"name":"其他附注","cells":{"A1":{"value":"薪酬"}}})
        self.assertEqual(len(build_target_meanings({"tables":tables},snapshot)["word_tables"]),2)

    def test_table_structure_changes_target_hash_and_rejects_old_scope_without_requests(self):
        context={"tables":[{"sheet":"附注","table_id":"t1","range_name":"TA1","title":"货币资金"}]}
        snapshot={"sheets":[{"name":"附注","cells":{"B2":{"value":"项目"},"B3":{"value":"库存现金"}},
            "word_table_ranges":[{"table_id":"t1","range_name":"TA1","range":"B2:C3"}]}]}
        target=build_target_meanings(context,snapshot)
        self.assertTrue(target["word_tables"])
        old={k:v for k,v in target.items() if k!="word_tables"}
        source=sample();previous=select_source_scope(Stub(lambda p:{"areas":[]}),source,old)
        engine=Stub(lambda p:{"areas":[]})
        with self.assertRaisesRegex(ValueError,"目标披露说明不同"):
            select_source_scope(engine,source,target,previous_selection=previous)
        self.assertFalse(engine.calls)
        current=select_source_scope(engine,source,target)
        self.assertNotEqual(current["target_hash"],previous["target_hash"])

    def test_shared_target_size_does_not_change_source_packets_or_truncate_target(self):
        from 附注更新.识别范围 import _json, MAX_PACKET_CHARS, MAX_REQUEST_CHARS, _SYSTEM
        source={"sha256":"stable-packets","sheets":[{"name":"资料","max_row":140,"max_column":1,"merges":[],
            "cells":{f"A{i}":{"value":"实际业务文字"*20+str(i)} for i in range(1,141)}}]}
        def reply(p):return {"areas":[area(a,"include",[a]) for a in p["candidate_cells"]]}
        small=Stub(reply);big=Stub(reply)
        select_source_scope(small,source,{"disclosures":["货币资金"]})
        target={"disclosures":["货币资金"],"chapter_context":["真实目标原文"*6000]}
        result=select_source_scope(big,source,target)
        self.assertEqual([p["candidate_cells"] for p in small.calls],[p["candidate_cells"] for p in big.calls])
        self.assertTrue(big.calls)
        for payload in big.calls:
            self.assertEqual(payload["target_meanings"],target)
            source_part={k:v for k,v in payload.items() if k not in {"source_hash","target_meanings","target_hash","gold_hash"}}
            self.assertLessEqual(len(_json(source_part)),MAX_PACKET_CHARS)
            self.assertLessEqual(len(_SYSTEM)+len(_json(payload)),MAX_REQUEST_CHARS)
        self.assertFalse(result["unresolved"])

    def test_complete_scope_request_budget_counts_system_and_keeps_all_when_exceeded(self):
        from 附注更新.识别范围 import _ScopeBatch, MAX_REQUEST_CHARS
        engine=Stub(lambda p:{"areas":[]})
        batch=_ScopeBatch(engine,sample(),{},batching=False)
        with self.assertRaisesRegex(ValueError,"完整范围请求"):
            batch._send("正文"*MAX_REQUEST_CHARS,{},lambda p:p)
        self.assertFalse(engine.calls)
        result=select_source_scope(engine,sample(),{"disclosures":["目标"*MAX_REQUEST_CHARS]})
        self.assertFalse(engine.calls)
        self.assertEqual(result["coverage"]["selected_count"],10)
        self.assertFalse(result["out_of_scope"])
        self.assertTrue(any("完整范围请求" in u["reason"] for u in result["unresolved"]))

    def test_structural_identifiers_without_target_business_text_do_not_trigger_requests(self):
        context={"tables":[{"sheet":"附注","table_id":"t1","range_name":"TA1","title":""}]}
        snapshot={"sheets":[{"name":"附注","cells":{"B2":{"value":456}},
            "word_table_ranges":[{"table_id":"t1","range_name":"TA1","range":"B2:C3"}]}]}
        target=build_target_meanings(context,snapshot)
        self.assertTrue(target["word_tables"])
        engine=Stub(lambda p:{"areas":[]});result=select_source_scope(engine,sample(),target)
        self.assertFalse(engine.calls)
        self.assertEqual(result["coverage"]["selected_count"],10)
        self.assertTrue(result["unresolved"])

    def test_target_text_projection_omits_amount_strings_and_formula_text_but_keeps_dimensions(self):
        from 附注更新.识别范围 import _word_text
        for text in ("金额：123,456.78","123,456.78美元","人民币123456.78元","=SUM(A1:A2)","123,456.78",r"C:\秘密\资料.xlsx"):
            self.assertIsNone(_word_text({"value":text}),text)
        for text in ("1年以内","2025年","前五名","存货","人民币元"):
            self.assertEqual(_word_text({"value":text}),text)

    def test_word_table_context_indexes_keep_exact_original_strings_and_order(self):
        paragraphs=[" 第六章 ","单位：人民币元"," 第六章 "]
        context={"tables":[{"sheet":"附注","table_id":"t1","range_name":"TA1","title":"资金","chapter_context":paragraphs}]}
        snapshot={"sheets":[{"name":"附注","cells":{"B2":{"value":"项目"}},
            "word_table_ranges":[{"table_id":"t1","range_name":"TA1","range":"B2:C3"}]}]}
        target=build_target_meanings(context,snapshot);indexes=target["word_tables"][0]["chapter_context_index"]
        self.assertEqual([target["chapter_context"][i] for i in indexes],paragraphs)
        self.assertEqual(indexes[0],indexes[2])

    def test_scope_budget_includes_transport_general_system_and_full_correction(self):
        from unittest.mock import patch
        from 附注更新 import 识别范围 as scope
        from 附注更新.语义 import SYSTEM, CORRECTION
        from 附注更新.回执格式 import transport_instruction
        payload={"task":"source_scope","scope_response_mode":"cells_v1","candidate_cells":["A1"],"cells":{"A1":{"kind":"text","text":"业务"}}}
        system=scope._SYSTEM+scope._REFINEMENT+scope._GRID_VIEW
        expected=len(SYSTEM+"\n"+system+"\n"+transport_instruction(payload))+len(scope._json(payload))
        self.assertEqual(scope.ensure_scope_request_budget(system,payload),expected)
        correction={**payload,"correction":{"previous_result":{"unresolved":["原始纠正回执"]},"validation_error":"缺少原格"}}
        with patch.object(scope,"MAX_REQUEST_CHARS",expected):
            scope.ensure_scope_request_budget(system,payload)
            with self.assertRaisesRegex(ValueError,"完整范围请求"):
                scope.ensure_scope_request_budget(system+"\n"+CORRECTION,correction)

    def test_target_row_metric_and_two_character_business_veto_only_its_own_area(self):
        for label,source_title in [("固定资产清理","固定资产清理明细表"),("信用减值损失","一、信用减值损失情况"),("存货","DB-08 存货明细表")]:
            with self.subTest(label=label):
                source=sample();sheet=source["sheets"][0]
                sheet["cells"]["A1"]["value"]=source_title
                sheet["cells"]["A6"]["value"]="项目";sheet["cells"]["A7"]["value"]="合计"
                engine=Stub(lambda p:{"areas":[area("A1:B3","out_of_scope",["A1"]),area("A5:B7","out_of_scope",["A5"])]})
                target={"disclosures":["财务数据"],"table_business":[label,"项目","合计","期末余额"]}
                result=select_source_scope(engine,source,target)
                self.assertIn("B3",result["selected_cells"][sheet["name"]])
                self.assertNotIn("B7",result["selected_cells"][sheet["name"]])
                self.assertTrue(all(a["decision"]=="uncertain" for a in result["areas"] if a["range"]=="A1:B3"))
                self.assertNotIn("mappings",result)

    def test_customer_names_and_generic_headers_are_context_not_business_protection(self):
        context={"tables":[{"table_id":"word-table-1","range_name":"TA_001","sheet":"附注","title":"财务披露"}]}
        snapshot={"sheets":[{"name":"附注","word_table_ranges":[{"table_id":"word-table-1","range_name":"TA_001","range":"B2:D4"}],"cells":{
            "B2":{"value":"项目"},"C2":{"value":"客户名称"},"D2":{"value":"期末余额"},
            "B3":{"value":"固定资产清理"},"C3":{"value":"张三"},"D3":{"value":1},
            "B4":{"value":"其他"},"C4":{"value":"甲有限公司"},"D4":{"value":2}}}]}
        target=build_target_meanings(context,snapshot)
        self.assertIn("张三",target["table_business"])
        self.assertEqual(target["table_business_terms"],["固定资产清理"])
        source=sample();source["sheets"][0]["cells"]["A1"]["value"]="固定资产清理"
        source["sheets"][0]["cells"]["A6"]["value"]="张三";source["sheets"][0]["cells"]["A7"]["value"]="其他"
        result=select_source_scope(Stub(lambda p:{"areas":[area("A1:B3","out_of_scope",["A1"]),area("A5:B7","out_of_scope",["A5"])]}),source,target)
        self.assertIn("B3",result["selected_cells"]["混合资料"])
        self.assertNotIn("B7",result["selected_cells"]["混合资料"])

    def test_ranking_count_and_source_form_do_not_remove_same_business(self):
        source=sample();source["sheets"][0]["cells"]["A1"]["value"]="按欠款方归集的期末余额前五名应收账款明细表"
        engine=Stub(lambda p:{"areas":[{**area("A1:B3","out_of_scope",["A1"]),"reason":"来源为前五名，而目标只有前一名，且来源表名不同"}]})
        result=select_source_scope(engine,source,{"disclosures":["按欠款方归集的期末余额前一名应收账款"]})
        self.assertIn("B3",result["selected_cells"]["混合资料"])
        self.assertFalse(result["out_of_scope"])
        self.assertTrue(result["unresolved"])

    def test_enriched_target_rejects_old_target_resume_and_still_fits_packets(self):
        source=sample();old={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[]}),source,old)
        target={**old,"table_business":["实际披露业务项目"+str(i) for i in range(235)]}
        engine=Stub(lambda p:{"areas":[]})
        with self.assertRaisesRegex(ValueError,"目标披露说明不同"):
            select_source_scope(engine,source,target,previous_selection=previous)
        self.assertFalse(engine.calls)
        result=select_source_scope(engine,source,target)
        self.assertEqual(len(engine.calls),2)
        self.assertEqual(result["coverage"]["candidate_count"],10)
        from 附注更新.识别范围 import _source_chars
        self.assertTrue(all(_source_chars(p)<=12000 for p in engine.calls))

    def test_failure_keeps_candidates_and_cancel_propagates(self):
        def fail(p):raise ValueError("模拟服务失败")
        result=select_source_scope(Stub(fail),sample(),{"disclosures":["货币资金"]})
        self.assertFalse(result["out_of_scope"])
        self.assertTrue(result["unresolved"])
        def cancel(p):raise SemanticCancelled("停止")
        with self.assertRaises(SemanticCancelled):select_source_scope(Stub(cancel),sample(),{"disclosures":["货币资金"]})

    def test_empty_structured_target_keeps_all_without_model(self):
        engine=Stub(lambda p:{"areas":[area("A1:B7","out_of_scope",["A1"])]})
        result=select_source_scope(engine,sample(),build_target_meanings({}))
        self.assertFalse(engine.calls)
        self.assertEqual(len(result["selected_cells"]["混合资料"]),10)

    def test_oversized_context_or_text_never_causes_silent_omission(self):
        engine=Stub(lambda p:{"areas":[]})
        result=select_source_scope(engine,sample(),{"context":"过长说明"*40000})
        self.assertFalse(engine.calls)
        self.assertEqual(len(result["selected_cells"]["混合资料"]),10)
        self.assertTrue(result["unresolved"])
        source=sample();source["sheets"][0]["cells"]["A1"]["value"]="超长业务文字"*3000
        engine=Stub(lambda p:{"areas":[]})
        result=select_source_scope(engine,source,{"disclosures":["货币资金"]})
        self.assertEqual(len(result["selected_cells"]["混合资料"]),10)
        from 附注更新.识别范围 import _source_chars
        self.assertTrue(all(_source_chars(p)<=12000 for p in engine.calls))



    def test_resume_small_sheet_only_retries_missing_round_with_full_context(self):
        target={"disclosures":["货币资金"]}
        def first(p):
            if p["round"]==2:raise ValueError("第二轮失败")
            return {"areas":[area("A1:B3","include",["A1"]),area("A5:B7","out_of_scope",["A5"])]}
        source=sample();previous=select_source_scope(Stub(first),source,target)
        engine=Stub(lambda p:{"areas":[area("A1:B3","include",["A1"]),area("A5:B7","out_of_scope",["A5"])]})
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertEqual([(p["sheet"],p["packet"],p["round"]) for p in engine.calls],[("混合资料",1,2)])
        self.assertEqual(set(engine.calls[0]["candidate_cells"]),set(source["sheets"][0]["cells"]))
        self.assertEqual(engine.calls[0]["cells"]["B2"]["text"],"期末余额")
        self.assertEqual(result["coverage"],{"candidate_count":10,"selected_count":5,"out_of_scope_count":5})
        self.assertEqual(result["resume"]["reused_round_count"],1)

    def test_resume_many_sheets_never_requests_successful_rounds_or_packets(self):
        source=sample()
        source["sheets"]=[{**copy.deepcopy(source["sheets"][0]),"name":"资料"+str(i)} for i in range(20)]
        def reply(p):
            return {"results":[{"packet_id":t["packet_id"],"areas":[area("A1:B3","include",["A1"]),area("A5:B7","out_of_scope",["A5"])]} for t in p["tasks"]]}
        target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(reply),source,target)
        previous["areas"]=[a for a in previous["areas"] if not(a["sheet"]=="资料7" and a["round"]==2)]
        engine=Stub(reply)
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertEqual(len(engine.calls),1)
        self.assertEqual(engine.calls[0]["round"],2)
        self.assertEqual([t["sheet"] for t in engine.calls[0]["tasks"]],["资料7"])
        self.assertEqual(result["coverage"],previous["coverage"])
        self.assertEqual(result["selected_cells"],previous["selected_cells"])
        self.assertEqual(result["resume"]["reused_round_count"],39)

    def test_resume_retries_uncertain_and_uncovered_cells_with_new_task_identity(self):
        source=sample();target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B3","uncertain",["A1"])]}),source,target)
        before=copy.deepcopy(previous)
        engine=Stub(lambda p:{"areas":[area("A1:B3","include",["A1"]),area("A5:B7","out_of_scope",["A5"])]})
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertEqual(len(engine.calls),2)
        self.assertTrue(all(p.get("refinement") for p in engine.calls))
        self.assertEqual(previous,before)
        self.assertEqual(result["coverage"],{"candidate_count":10,"selected_count":5,"out_of_scope_count":5})
        self.assertFalse(result["unresolved"])

    def test_resume_preserves_confirmed_cells_and_only_rechecks_pending_with_shared_text(self):
        source=sample();target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B3","include",["A1"]),area("A5:B7","uncertain",["A5"])]}),source,target)
        before=copy.deepcopy(previous)
        engine=Stub(lambda p:{"areas":[area("A5:B7","out_of_scope",["A5"])]})
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertEqual(len(engine.calls),2)
        for p in engine.calls:
            self.assertEqual(set(p["candidate_cells"]),{"A5","A6","B6","A7","B7"})
            self.assertEqual(p["target_meanings"],target)
            self.assertEqual(p["cells"]["B2"]["text"],"期末余额")
            self.assertTrue(p["refinement"]["previous_scope_hash"])
        self.assertEqual(previous,before)
        self.assertFalse(result["unresolved"])
        self.assertEqual(result["coverage"]["selected_count"],5)

    def test_resume_retries_opposing_votes_but_preserves_other_confirmed_area(self):
        source=sample();target={"disclosures":["另一业务"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B3","include" if p["round"]==1 else "out_of_scope",["A1"]),area("A5:B7","out_of_scope",["A5"])]}),source,target)
        engine=Stub(lambda p:{"areas":[area("A1:B3","include",["A1"])]})
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertEqual([p["round"] for p in engine.calls],[1,2])
        self.assertTrue(all(set(p["candidate_cells"])=={"A1","A2","B2","A3","B3"} for p in engine.calls))
        self.assertFalse(result["unresolved"])
        self.assertEqual(result["coverage"]["out_of_scope_count"],5)

    def test_failed_refinement_keeps_confirmed_regions_and_changes_next_attempt(self):
        source=sample();target={"disclosures":["货币资金"]}
        first=Stub(lambda p:{"areas":[area("A1:B3","include",["A1"]),area("A5:B7","uncertain",["A5"])]})
        previous=select_source_scope(first,source,target)
        engine=Stub(lambda p:{"areas":[]})
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertTrue(all(set(u["cells"])<={"A5","A6","B6","A7","B7"} for u in result["unresolved"]))
        again=Stub(lambda p:{"areas":[]})
        select_source_scope(again,source,target,previous_selection=result)
        self.assertGreater(again.calls[0]["refinement"]["attempt"],engine.calls[0]["refinement"]["attempt"])
        self.assertNotEqual(again.calls[0]["refinement"]["previous_scope_hash"],engine.calls[0]["refinement"]["previous_scope_hash"])

    def test_refinement_cannot_overwrite_confirmed_cells_with_a_broad_rectangle(self):
        source=sample();target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B3","include",["A1"]),area("A5:B7","uncertain",["A5"])]}),source,target)
        result=select_source_scope(Stub(lambda p:{"areas":[area("A1:B7","out_of_scope",["A5"])]}),source,target,previous_selection=previous)
        self.assertEqual(result["coverage"]["selected_count"],10)
        self.assertTrue(all(set(u["cells"])<={"A5","A6","B6","A7","B7"} for u in result["unresolved"]))

    def test_business_comparison_can_resolve_shared_ratio_label_without_changing_population(self):
        values={"A1":"期末单项计提坏账准备的应收账款","A3":"债务人名称","B3":"账面余额","C3":"坏账准备","E3":"计提比例（%）","B4":0,"C4":None,"E4":None}
        source={"sha256":"qcf25","sheets":[{"name":"QCF25","max_row":4,"max_column":5,"merges":[],"cells":{a:{"value":v,"formula":None} for a,v in values.items()}}]}
        target={"disclosures":["其他应收款"],"table_business":["计提比例（%）"],"table_business_terms":["计提比例（%）"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:E4","out_of_scope",["A1","E3"])]}),source,target)
        comparison={"source_business":"应收账款单项坏账准备","target_business":"其他应收款单项坏账准备","difference_kind":"business_population",
                    "reason":"应收销售债权与其他应收债权属于不同业务总体，并非客户名称不同","target_evidence":["其他应收款","计提比例（%）"]}
        def reply(p):return {"areas":[{**area("A1:E4","out_of_scope",["A1","E3"]),"semantic_comparison":comparison}]}
        result=select_source_scope(Stub(reply),source,target,previous_selection=previous)
        self.assertEqual(result["coverage"]["out_of_scope_count"],8)
        self.assertFalse(result["unresolved"])
        cached=Stub(lambda p:self.fail("两轮确认后的区域不得重跑"))
        self.assertEqual(select_source_scope(cached,source,target,previous_selection=result)["coverage"],result["coverage"])
        for invalid in ("period","counterparty","ranking","carrier","blank","missing_slot"):
            bad=copy.deepcopy(comparison);bad["difference_kind"]=invalid
            rejected=select_source_scope(Stub(lambda p:{"areas":[{**area("A1:E4","out_of_scope",["A1","E3"]),"semantic_comparison":bad}]}),source,target,previous_selection=previous)
            self.assertFalse(rejected["out_of_scope"])
        bad=copy.deepcopy(comparison);bad["target_evidence"]=["凭空编造的目标业务"]
        rejected=select_source_scope(Stub(lambda p:{"areas":[{**area("A1:E4","out_of_scope",["A1","E3"]),"semantic_comparison":bad}]}),source,target,previous_selection=previous)
        self.assertFalse(rejected["out_of_scope"])

    def test_refinement_retains_z02_same_indicator_despite_claimed_carrier_difference(self):
        source=sample();source["sheets"][0]["cells"]["A1"]["value"]="税金及附加"
        target={"disclosures":["税金及附加"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B7","out_of_scope",["A1"])]}),source,target)
        comparison={"source_business":"利润表税金及附加","target_business":"附注税金及附加","difference_kind":"indicator",
                    "reason":"来源是主表，目标是附注","target_evidence":["税金及附加"]}
        result=select_source_scope(Stub(lambda p:{"areas":[{**area("A1:B7","out_of_scope",["A1"]),"semantic_comparison":comparison}]}),source,target,previous_selection=previous)
        self.assertFalse(result["out_of_scope"])

    def test_refinement_splits_z09_person_counts_from_payroll_with_shared_headers(self):
        values={"A1":"人力资源情况表","A3":"项目","B3":"本年数","C3":"项目","D3":"本年金额",
                "A4":"参加基本养老保险的年末职工人数","B4":0,"C4":"本年应发职工薪酬总额","D4":None}
        source={"sha256":"z09","sheets":[{"name":"Z09","max_row":4,"max_column":4,"merges":[],"cells":{a:{"value":v,"formula":None} for a,v in values.items()}}]}
        target={"disclosures":["应付职工薪酬"],"table_business_terms":["基本养老保险","职工薪酬"],"table_business":["基本养老保险","职工薪酬"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:D4","out_of_scope",["A1","A4","C4"])]}),source,target)
        comparison={"source_business":"参加养老保险的职工人数","target_business":"基本养老保险金额","difference_kind":"value_type",
                    "reason":"人数统计的单位是人，目标为养老保险薪酬金额，不是期间或新增职工对象的不同","target_evidence":["基本养老保险","应付职工薪酬"]}
        def reply(p):return {"areas":[area("A1:D3","include",["A1"]),{**area("A4:B4","out_of_scope",["A4"]),"semantic_comparison":comparison},area("C4:D4","include",["C4"])]}
        result=select_source_scope(Stub(reply),source,target,previous_selection=previous)
        self.assertEqual({a for v in result["out_of_scope"] for a in v["cells"]},{"A4","B4"})
        self.assertIn("D4",result["selected_cells"]["Z09"])
        self.assertFalse(result["unresolved"])

    def test_resume_rejects_source_target_or_partition_tampering_before_request(self):
        source=sample();target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B7","include",["A1"])]}),source,target)
        variants=[]
        x=copy.deepcopy(previous);x["source_hash"]="别的原文件";variants.append(x)
        x=copy.deepcopy(previous);x["target_hash"]="别的目标";variants.append(x)
        x=copy.deepcopy(previous);x["target_meanings"]={"disclosures":["篡改"]};variants.append(x)
        x=copy.deepcopy(previous);x["selected_cells"]["混合资料"].pop();variants.append(x)
        x=copy.deepcopy(previous);x["selected_cells"]["混合资料"].append("B3");variants.append(x)
        x=copy.deepcopy(previous);x["coverage"]["candidate_count"]=999;variants.append(x)
        for old in variants:
            with self.subTest(old=old):
                engine=Stub(lambda p:self.fail("完整性失败必须在请求前拒绝"))
                with self.assertRaises(ValueError):select_source_scope(engine,source,target,previous_selection=old)
                self.assertFalse(engine.calls)

    def test_resume_revalidates_old_evidence_against_real_payload_and_retains_on_failure(self):
        source=sample();target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B3","include",["A1"]),area("A5:B7","out_of_scope",["A5"])]}),source,target)
        next(a for a in previous["areas"] if a["round"]==2 and a["decision"]=="out_of_scope")["evidence_cells"]=["B7"]
        def fail(p):raise ValueError("重试暂时失败")
        engine=Stub(fail)
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertEqual([p["round"] for p in engine.calls],[2])
        self.assertFalse(result["out_of_scope"])
        self.assertEqual(result["coverage"]["selected_count"],10)
        self.assertEqual(result["resume"]["reused_round_count"],1)
        self.assertTrue(result["resume"]["rejected_rounds"])

    def test_resume_preserves_cross_packet_veto_from_valid_old_answers(self):
        cells={f"A{i}":{"row":i,"column":1,"value":"披露业务文字"*35+str(i),"formula":None} for i in range(1,151)}
        source={"sha256":"large","sheets":[{"name":"混合","max_row":150,"max_column":1,"cells":cells,"merges":[]}]}
        target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:A150","out_of_scope" if p["packet"]==1 else "include",[p["candidate_cells"][0]])]}),source,target)
        engine=Stub(lambda p:{"areas":[]})
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertTrue(engine.calls)
        self.assertTrue(all(p.get("refinement") for p in engine.calls))
        self.assertFalse(result["out_of_scope"])
        self.assertEqual(result["coverage"]["selected_count"],150)


    def test_batch_resume_only_rechecks_pending_cells_in_affected_sheet(self):
        source=sample();source["sheets"]=[{**copy.deepcopy(source["sheets"][0]),"name":"资料"+str(i)} for i in range(12)]
        target={"disclosures":["货币资金"]}
        def first(p):
            return {"results":[{"packet_id":t["packet_id"],"areas":[area("A1:B3","include",["A1"]),area("A5:B7","uncertain" if t["sheet"]=="资料7" else "out_of_scope",["A5"])]} for t in p["tasks"]]}
        previous=select_source_scope(Stub(first),source,target)
        def second(p):
            self.assertEqual(len(p["tasks"]),1)
            item=p["tasks"][0];self.assertEqual(item["sheet"],"资料7")
            self.assertEqual(set(item["candidate_cells"]),{"A5","A6","B6","A7","B7"})
            return {"results":[{"packet_id":item["packet_id"],"areas":[area("A5:B7","out_of_scope",["A5"])]}]}
        engine=Stub(second);result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertEqual(len(engine.calls),2)
        self.assertEqual(result["coverage"],{"candidate_count":120,"selected_count":60,"out_of_scope_count":60})
        self.assertFalse(result["unresolved"])

    def test_batch_validation_error_identifies_original_task_and_bad_evidence(self):
        source=sample()
        source["sheets"]=[{**copy.deepcopy(source["sheets"][0]),"name":"资料"+str(i)} for i in range(20)]
        errors=[]
        def reply(payload):
            return {"results":[{"packet_id":t["packet_id"],"areas":[area("A1:B7","out_of_scope",
                ["B7"]+[f"Z{i}" for i in range(1,13)] if t["sheet"]=="资料1" else ["A5"])]}
                for t in payload["tasks"]]}
        class Inspect(Stub):
            def _request(self,system,payload,validator):
                try:
                    return super()._request(system,payload,validator)
                except ValueError as error:
                    errors.append(str(error))
                    raise
        result=select_source_scope(Inspect(reply),source,{"disclosures":["货币资金"]})
        self.assertTrue(errors)
        for marker in ('packet_id=2','sheet=资料1','packet=1','round=1','B7','Z9'):
            self.assertIn(marker,errors[0])
        self.assertNotIn('"Z10"',errors[0])
        self.assertNotIn('"Z12"',errors[0])
        self.assertEqual(result["coverage"]["selected_count"],200)
        self.assertFalse(result["out_of_scope"])


    def test_grid_rows_keeps_numeric_coordinates_and_exact_sparse_entries(self):
        from 附注更新.识别范围 import _grid_rows, _json
        cells={"AA10":{"kind":"number"},"A10":{"kind":"text","text":"养老人数"},
               "Z10":{"kind":"text","text":"工资总额"},"C3":{"kind":"blank"},"A3":{"kind":"formula"}}
        context={"A1":{"kind":"text","text":"合并标题"},"A10":copy.deepcopy(cells["A10"])}
        before=copy.deepcopy((cells,context));grid=_grid_rows(cells,context)
        decoded=json.loads(_json({"grid_rows":grid}))["grid_rows"]
        self.assertEqual([r["row"] for r in decoded],[1,3,10])
        self.assertEqual([a for r in decoded for a,_ in r["cells"]],["A1","A3","C3","A10","Z10","AA10"])
        restored={a:v for r in decoded for a,v in r["cells"]}
        self.assertEqual(restored,{**context,**cells})
        self.assertEqual(restored["C3"],{"kind":"blank"})
        self.assertNotIn("B1",restored)  # 合并从格及没有原格的位置均不补造。
        self.assertNotIn("B3",restored)
        grid[0]["cells"][0][1]["text"]="只改展示副本"
        self.assertEqual((cells,context),before)

    def test_grid_rows_rejects_conflicting_same_address_evidence(self):
        from 附注更新.识别范围 import _grid_rows
        with self.assertRaisesRegex(ValueError,"A1"):
            _grid_rows({"A1":{"kind":"blank"}},{"A1":{"kind":"text","text":"标题"}})

    def test_fixed_requests_add_reversible_grid_without_changing_initial_packets(self):
        values={"A1":"货币资金","A3":"库存现金","B3":987654321,"C3":None,
                "A10":"其他资金","Z10":"期末余额","AA10":None}
        source={"sha256":"sparse-grid","sheets":[{"name":"资料","max_row":10,"max_column":27,
            "merges":["A1:C1"],"cells":{a:{"value":v,"formula":None} for a,v in values.items()}}]}
        source["sheets"][0]["cells"]["AA10"]["formula"]="=SECRET_FORMULA()"
        before=copy.deepcopy(source);target={"disclosures":["货币资金"]}
        initial=Stub(lambda p:{"areas":[area("A1:AA10","uncertain",["A1"])]})
        previous=select_source_scope(initial,source,target)
        self.assertTrue(initial.calls)
        self.assertTrue(all("grid_rows" not in p for p in initial.calls))
        systems=[]
        class GridStub(CellStub):
            def _request(self,system,payload,validator):
                systems.append(system)
                return super()._request(system,payload,validator)
        engine=GridStub(cell_reply)
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertEqual(len(engine.calls),2)
        for request in engine.calls:
            restored={a:v for r in request["grid_rows"] for a,v in r["cells"]}
            self.assertEqual(restored,{**request["context_cells"],**request["cells"]})
            self.assertEqual(request["merges"],["A1:C1"])
            self.assertEqual(set(request["candidate_cells"]),set(values))
            self.assertEqual(restored["B3"],{"kind":"number"})
            self.assertEqual(restored["C3"],{"kind":"blank"})
            self.assertEqual(restored["AA10"],{"kind":"formula"})
            self.assertNotIn("B1",restored)
            self.assertNotIn("B10",restored)
        self.assertTrue(all("grid_rows" in system for system in systems))
        shown=json.dumps(engine.calls,ensure_ascii=False)
        self.assertNotIn("987654321",shown)
        self.assertNotIn("SECRET_FORMULA",shown)
        self.assertFalse(result["unresolved"])
        self.assertEqual(source,before)

    def test_cell_refinement_limits_each_request_to_40_and_keeps_full_original_context(self):
        source={"sha256":"wide-cell-scope","sheets":[{"name":"资料","max_row":43,"max_column":2,"merges":[],
            "cells":{f"{column}{row}":{"value":("货币资金" if row==1 else "业务原文") if column=="A" else (None if row==2 else row),"formula":None}
                     for row in range(1,44) for column in ("A","B")}}]}
        before=copy.deepcopy(source);target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B43","uncertain",["A1"])]}),source,target)
        engine=CellStub(lambda p:cell_reply(p))
        result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertTrue(engine.calls)
        self.assertTrue(all(len(p["candidate_cells"])<=40 for p in engine.calls))
        self.assertTrue(all(p.get("scope_response_mode")=="cells_v1" for p in engine.calls))
        self.assertEqual([len(p["candidate_cells"]) for p in engine.calls],[40,40,6,40,40,6])
        for p in engine.calls:
            self.assertEqual(set(p["cells"]),set(source["sheets"][0]["cells"]))
            self.assertEqual(p["cells"]["A43"]["text"],"业务原文")
            self.assertEqual(p["cells"]["B2"],{"kind":"blank"})
            self.assertEqual(p["target_meanings"],target)
        self.assertFalse(result["unresolved"])
        self.assertEqual(len(result["areas"]),172)
        self.assertTrue(all(":" not in a["range"] for a in result["areas"]))
        self.assertEqual(source,before)

    def test_cell_refinement_failure_preserves_other_chunks_and_retries_only_failed_round(self):
        source={"sha256":"chunk-failure","sheets":[{"name":"资料","max_row":43,"max_column":2,"merges":[],
            "cells":{f"{column}{row}":{"value":"货币资金" if column=="A" else row,"formula":None}
                     for row in range(1,44) for column in ("A","B")}}]}
        target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B43","uncertain",["A1"])]}),source,target)
        def reply(p):
            if p["round"]==1 and p["candidate_cells"][0]=="A21":raise ValueError("该子请求失败")
            return cell_reply(p)
        engine=CellStub(reply);result=select_source_scope(engine,source,target,previous_selection=previous)
        expected={f"{column}{row}" for row in range(21,41) for column in ("A","B")}
        self.assertEqual({a for u in result["unresolved"] for a in u["cells"]},expected)
        again=CellStub(cell_reply)
        complete=select_source_scope(again,source,target,previous_selection=result)
        self.assertEqual(len(again.calls),1)
        self.assertEqual(again.calls[0]["round"],1)
        self.assertEqual(set(again.calls[0]["candidate_cells"]),expected)
        self.assertFalse(complete["unresolved"])

    def test_cell_batch_limits_total_candidates_not_each_member_and_propagates_mode(self):
        source=sample();source["sheets"]=[{**copy.deepcopy(source["sheets"][0]),"name":"资料"+str(i)} for i in range(12)]
        target={"disclosures":["货币资金"]}
        def previous_reply(p):
            return {"results":[{"packet_id":t["packet_id"],"areas":[area("A1:B7","uncertain",["A1"])]} for t in p["tasks"]]}
        previous=select_source_scope(Stub(previous_reply),source,target)
        engine=CellStub(cell_reply);result=select_source_scope(engine,source,target,previous_selection=previous)
        self.assertTrue(engine.calls)
        self.assertTrue(all(sum(len(t["candidate_cells"]) for t in p["tasks"])<=40 for p in engine.calls))
        self.assertEqual(len(engine.calls),6)
        for p in engine.calls:
            self.assertEqual(p["scope_response_mode"],"cells_v1")
            self.assertTrue(all(t["scope_response_mode"]=="cells_v1" for t in p["tasks"]))
            for task in p["tasks"]:
                self.assertEqual({a:v for row in task["grid_rows"] for a,v in row["cells"]},
                                 {**task["context_cells"],**task["cells"]})
        self.assertFalse(result["unresolved"])
        self.assertEqual(result["coverage"]["selected_count"],120)

    def test_cell_refinement_rejects_missing_extra_and_duplicate_normalized_cells(self):
        source=sample();target={"disclosures":["货币资金"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:B7","uncertain",["A1"])]}),source,target)
        def missing(p):
            answer=cell_reply(p);answer["cells"].pop(p["candidate_cells"][-1]);return answer
        def extra(p):
            answer=cell_reply(p);answer["cells"]["Z99"]=answer["cells"][p["candidate_cells"][0]];return answer
        for reply in (missing,extra):
            result=select_source_scope(CellStub(reply),source,target,previous_selection=previous)
            self.assertFalse(result["out_of_scope"])
            self.assertEqual({a for u in result["unresolved"] for a in u["cells"]},set(source["sheets"][0]["cells"]))
        result=select_source_scope(Stub(lambda p:{"areas":[area(a,"include",["A1"]) for a in p["candidate_cells"]]+[area("B3","include",["A1"])]}),
                                   source,target,previous_selection=previous)
        self.assertTrue(result["unresolved"])


    def test_fixed_cell_protocol_separates_person_counts_and_blank_payroll_in_same_row(self):
        values={"A1":"人力资源情况表","A3":"项目","B3":"本年数","C3":"项目","D3":"本年金额",
                "A4":"参加基本养老保险的年末职工人数","B4":0,"C4":"本年应发职工薪酬总额","D4":None}
        source={"sha256":"z09-fixed","sheets":[{"name":"Z09","max_row":4,"max_column":4,"merges":[],
                "cells":{a:{"value":v,"formula":None} for a,v in values.items()}}]}
        before=copy.deepcopy(source)
        target={"disclosures":["应付职工薪酬"],"table_business_terms":["基本养老保险","职工薪酬"],
                "table_business":["基本养老保险","职工薪酬"]}
        previous=select_source_scope(Stub(lambda p:{"areas":[area("A1:D4","uncertain",["A1"])]}),source,target)
        def reply(p):
            answer=cell_reply(p)
            comparison={"source_business":"参加养老保险的职工人数","target_business":"基本养老保险金额",
                "difference_kind":"value_type","reason":"人数与金额不同，不以零值或空白判断业务",
                "target_evidence":["基本养老保险","应付职工薪酬"]}
            for address in ("A4","B4"):
                answer["cells"][address].update(decision="out_of_scope",evidence_cells=["A4"],
                                                semantic_comparison=comparison)
            answer["cells"]["C4"]["evidence_cells"]=["C4"]
            answer["cells"]["D4"]["evidence_cells"]=["C4","D3"]
            return answer
        result=select_source_scope(CellStub(reply),source,target,previous_selection=previous)
        self.assertEqual({a for item in result["out_of_scope"] for a in item["cells"]},{"A4","B4"})
        self.assertIn("D4",result["selected_cells"]["Z09"])
        self.assertFalse(result["unresolved"])
        self.assertTrue(all(":" not in a["range"] for a in result["areas"]))
        self.assertEqual(source,before)

if __name__=="__main__":unittest.main()
