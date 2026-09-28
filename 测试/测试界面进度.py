"""实际构建 Tk，以受控后台回调检查中途保存、取消及重启；不请求模型。"""
import copy
import hashlib
import json
import sys
import tempfile
import threading
import tkinter as tk
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from 附注更新 import 界面
from 附注更新.设置 import load_settings, save_settings


class 界面进度测试(unittest.TestCase):
    def setUp(self):
        self.folder=Path(tempfile.mkdtemp(prefix="附注界面进度_"))
        self.config_path=self.folder/"独立测试配置.json"
        self.roots=[];self.apps=[];self.patches=[]
        self.user_config=界面.PROJECT_DIR/"用户配置.json"
        self.user_hash=hashlib.sha256(self.user_config.read_bytes()).hexdigest() if self.user_config.exists() else None
        for name in ("原附注.xlsx","更新数据.xlsx","测试金标准.jsonl"):(self.folder/name).write_text("独立测试材料",encoding="utf-8-sig")
        self.base={"a_path":str(self.folder/"原附注.xlsx"),"b_path":str(self.folder/"更新数据.xlsx"),
            "gold_path":str(self.folder/"测试金标准.jsonl"),"output_dir":str(self.folder),
            "base_url":"http://127.0.0.1:1/v1","model":"受控测试","api_key":"",
            "a_context_text":"已确认的附注事实","b_context_text":"已确认的更新数据事实"}
        for name in ("showerror","showinfo"):
            p=patch.object(界面.messagebox,name);p.start();self.patches.append(p)
        p=patch.object(界面,"save_settings",side_effect=lambda config:save_settings(config,self.config_path))
        p.start();self.patches.append(p)

    def tearDown(self):
        for app in self.apps:
            app.cancel.set()
            if app.worker and app.worker.is_alive():app.worker.join(3)
        for root in self.roots:
            for pending in root.tk.call("after","info"):root.tk.call("after","cancel",pending)
            root.destroy()
        for p in reversed(self.patches):p.stop()
        actual=hashlib.sha256(self.user_config.read_bytes()).hexdigest() if self.user_config.exists() else None
        self.assertEqual(actual,self.user_hash,"不得修改真实用户配置")

    def make_app(self,settings):
        root=tk.Tk();root.withdraw();self.roots.append(root)
        with patch.object(界面,"load_settings",return_value=settings):app=界面.Application(root)
        self.apps.append(app);root.update_idletasks();return app

    def mapping(self,carrier,name):
        path=self.folder/name
        record={"schema":"附注语义映射-v1","carrier":carrier,
            "source_path":self.base["a_path" if carrier=="A" else "b_path"],
            "additional_context":"本轮已确认事实","mappings":[{"cell":"B3"}],"excluded":[],"unresolved":[{"cell":"B4"}],"complete":False}
        path.write_text(json.dumps(record,ensure_ascii=False),encoding="utf-8-sig");return str(path)

    def test_启动取消没有新成果仍保留旧映射并能重启加载(self):
        for carrier,step,key in (("A","recognize_a","a_mapping"),("B","recognize_b","b_mapping")):
            with self.subTest(carrier=carrier):
                old=self.mapping(carrier,carrier+"旧映射.json");settings={**self.base,key:old,"current_step":step,"update_plan":"旧清单.json"}
                app=self.make_app(settings);started=threading.Event()
                app.step_states[step]="complete";app.show_step(step)
                def run_step(step,config,log=None,cancel=None,progress=None):
                    self.assertEqual(config[key],old);started.set();cancel.wait(3)
                    return {"status":"cancelled","step":step,"message":"已取消，先前成果保留","files":{}}
                with patch.dict(sys.modules,{"附注更新.分步流程":types.SimpleNamespace(run_step=run_step)}):
                    app.start();self.assertTrue(started.wait(2));self.assertTrue(app.running)
                    self.assertEqual(app.values[key].get(),old);self.assertEqual(app.values["update_plan"].get(),"")
                    self.assertNotIn("已完成",app.nav_buttons[step].cget("text"))
                    app.request_cancel();app.worker.join(3);app.poll()
                self.assertFalse(app.running);self.assertEqual(app.values[key].get(),old)
                self.assertEqual(app.step_states[step],"cancelled")
                reopened=self.make_app(load_settings(self.config_path));self.assertEqual(reopened.values[key].get(),old)

    def test_进度保存映射和数量不结束任务且取消后继续(self):
        for carrier,step,key in (("A","recognize_a","a_mapping"),("B","recognize_b","b_mapping")):
            with self.subTest(carrier=carrier):
                old=self.mapping(carrier,carrier+"进度旧映射.json");new=self.mapping(carrier,carrier+"进度新映射.json")
                review="a_review" if carrier=="A" else "b_review"
                app=self.make_app({**self.base,key:old,review:"旧核对.xlsx","current_step":step})
                queued=threading.Event();received=[]
                def run_step(step,config,log=None,cancel=None,progress=None):
                    received.append(config);progress({"status":"running","step":step,"output_dir":str(self.folder),
                        "message":"本包已保存，可接续","files":{key:new},"counts":{"mappings":1,"excluded":0,"unresolved":1},"checkpoint":True})
                    queued.set();cancel.wait(3)
                    return {"status":"cancelled","step":step,"message":"已取消","files":{}}
                with patch.dict(sys.modules,{"附注更新.分步流程":types.SimpleNamespace(run_step=run_step)}):
                    app.start();self.assertTrue(queued.wait(2))
                    with patch.object(app,"apply_result",wraps=app.apply_result) as done:
                        app.poll();done.assert_not_called()
                    self.assertTrue(app.running);self.assertEqual(app.values[key].get(),new)
                    self.assertEqual(app.values[review].get(),"")
                    self.assertIn("业务 1",app.status.cget("text"));self.assertIn("待核实 1",app.status.cget("text"))
                    self.assertEqual(app.values["a_context_text"].get(),self.base["a_context_text"])
                    self.assertEqual(app.values["b_context_text"].get(),self.base["b_context_text"])
                    self.assertEqual(load_settings(self.config_path)[key],new)
                    app.request_cancel();app.worker.join(3);app.poll()
                self.assertEqual(app.values[key].get(),new);self.assertEqual(app.step_states[step],"cancelled")
                restarted=self.make_app(load_settings(self.config_path));self.assertEqual(restarted.values[key].get(),new)
                passed=[]
                def resumed(step,config,log=None,cancel=None,progress=None):
                    passed.append(config[key]);return {"status":"cancelled","step":step,"message":"受控接续验证结束","files":{}}
                with patch.dict(sys.modules,{"附注更新.分步流程":types.SimpleNamespace(run_step=resumed)}):
                    restarted.start();restarted.worker.join(3);restarted.poll()
                self.assertEqual(passed,[new])

    def test_失败结果携带最新映射应保留而不声称完成(self):
        path=self.mapping("B","失败时已保存.json");app=self.make_app({**self.base,"current_step":"recognize_b"})
        app.running_step="recognize_b";app.set_running(True)
        app.messages.put(("done",("recognize_b",{"status":"failed","step":"recognize_b","message":"后续包失败",
            "files":{"b_mapping":path},"output_dir":str(self.folder)})))
        app.poll();self.assertEqual(app.values["b_mapping"].get(),path)
        self.assertEqual(app.step_states["recognize_b"],"failed");self.assertFalse(app.running)
        self.assertNotIn("已完成",app.nav_buttons["recognize_b"].cget("text"))
        self.assertEqual(load_settings(self.config_path)["b_mapping"],path)

    def test_进度只更新本步成果且采用本次绑定的范围文件(self):
        old=self.mapping("B","带旧范围的映射.json");new=self.mapping("B","带新范围的映射.json")
        scope=self.folder/"本次范围.json";scope.write_text("{}",encoding="utf-8-sig")
        app=self.make_app({**self.base,"b_mapping":old,"b_scope":"旧范围.json","b_review":"旧核对.xlsx","current_step":"recognize_b"})
        app.running_step="recognize_b";app.set_running(True)
        app.apply_progress({"status":"running","step":"recognize_b","checkpoint":True,
            "files":{"b_mapping":new,"b_scope":str(scope),"gold_path":"不能由进度改变.jsonl","b_path":"不能由进度改变.xlsx"},
            "counts":{"mappings":1,"excluded":0,"unresolved":1}})
        self.assertEqual(app.values["b_mapping"].get(),new);self.assertEqual(app.values["b_scope"].get(),str(scope))
        self.assertEqual(app.values["b_review"].get(),"")
        self.assertEqual(app.values["gold_path"].get(),self.base["gold_path"])
        self.assertEqual(app.values["b_path"].get(),self.base["b_path"])
        app.apply_progress({"status":"running","step":"recognize_b","checkpoint":True,"files":{"b_mapping":str(self.folder/"未保存文件.json")}})
        self.assertEqual(app.values["b_mapping"].get(),new)

    def test_已发布金标准随进度保存且未收终态重启也能接续(self):
        for carrier,step,key in (("A","recognize_a","a_mapping"),("B","recognize_b","b_mapping")):
            with self.subTest(carrier=carrier):
                mapping=self.mapping(carrier,carrier+"发布前已有映射.json")
                published=self.folder/(carrier+"已发布的新金标准.jsonl")
                published.write_text("已发布的独立测试标准",encoding="utf-8-sig")
                record=json.loads(Path(mapping).read_text(encoding="utf-8-sig"))
                record["gold_path"]=self.base["gold_path"]
                record["latest_published_gold_path"]=str(published)
                # 初始材料一次写出；映射忠实绑定旧标准，不能仅凭其候选字段切换设置。
                initial=self.folder/(carrier+"映射仍引用原定义.json")
                initial.write_text(json.dumps(record,ensure_ascii=False),encoding="utf-8-sig")
                app=self.make_app({**self.base,"current_step":step})
                app.running_step=step;app.set_running(True)
                event={"status":"running","step":step,"checkpoint":True,"files":{key:str(initial)},
                       "counts":{"mappings":1,"excluded":0,"unresolved":1}}
                app.messages.put(("progress",event));app.poll()
                self.assertEqual(app.values["gold_path"].get(),self.base["gold_path"])
                published_event=copy.deepcopy(event);published_event["files"]["gold_path"]=str(published)
                app.messages.put(("progress",published_event))
                with patch.object(app,"apply_result",wraps=app.apply_result) as done:
                    app.poll();done.assert_not_called()
                self.assertTrue(app.running)
                self.assertEqual(app.values["gold_path"].get(),str(published))
                saved=load_settings(self.config_path)
                self.assertEqual(saved[key],str(initial));self.assertEqual(saved["gold_path"],str(published))
                # 不发送 done，直接由独立配置重建窗口，模拟收到中途成果后异常退出。
                reopened=self.make_app(saved)
                self.assertEqual(reopened.values[key].get(),str(initial))
                self.assertEqual(reopened.values["gold_path"].get(),str(published))
                self.assertEqual(reopened.values["a_context_text"].get(),self.base["a_context_text"])
                self.assertEqual(reopened.values["b_context_text"].get(),self.base["b_context_text"])
                for status in ("cancelled","failed"):
                    app.apply_result(step,{"status":status,"step":step,"message":"受控终态",
                        "files":{key:str(initial),"gold_path":str(published)}})
                    self.assertEqual(app.step_states[step],status)
                    saved=load_settings(self.config_path)
                    self.assertEqual(saved[key],str(initial));self.assertEqual(saved["gold_path"],str(published))

    def test_真正更换材料仍清除旧映射且旧步骤进度不能串入(self):
        old=self.mapping("B","输入改变前.json");app=self.make_app({**self.base,"b_mapping":old,"current_step":"recognize_b"})
        other=self.folder/"另一份更新数据.xlsx";other.write_text("独立测试",encoding="utf-8-sig")
        self.assertTrue(app.select_file("b_path",str(other)));self.assertEqual(app.values["b_mapping"].get(),"")
        app.values["b_mapping"].set(old);app.set_context("b_context_text","业务事实改变")
        self.assertEqual(app.values["b_mapping"].get(),"")
        app.running_step="recognize_a";app.set_running(True)
        app.messages.put(("progress",{"status":"running","step":"recognize_b","checkpoint":True,"files":{"b_mapping":old},"counts":{}}))
        app.poll();self.assertEqual(app.values["b_mapping"].get(),"")


    def test_两份所选映射分别展示数量且重启仍注明部分成果(self):
        a=self.mapping("A","附注部分映射.json")
        b=self.folder/"更新部分映射.json"
        b.write_text(json.dumps({"schema":"附注语义映射-v1","carrier":"B","mappings":[{},{}],
            "excluded":[{},{},{}],"unresolved":[],"complete":False}),encoding="utf-8-sig")
        app=self.make_app({**self.base,"a_mapping":a,"b_mapping":str(b),"current_step":"match"})
        self.assertEqual(set(app.mapping_summary_labels),{"a_mapping","b_mapping"})
        for key,business,excluded,pending in (("a_mapping",1,0,1),("b_mapping",2,3,0)):
            text=app.mapping_summary_labels[key].cget("text")
            for expected in ("部分成果",f"业务 {business} 格",f"非业务 {excluded} 格",f"待核实 {pending} 格"):
                self.assertIn(expected,text)
        reopened=self.make_app(app.config())
        self.assertIn("部分成果",reopened.mapping_summary_labels["b_mapping"].cget("text"))
        app.show_step("recognize_a")
        replacement=self.mapping("A","替换附注映射.json")
        self.assertTrue(app.select_file("a_mapping",replacement))
        self.assertIn("待核实 1 格",app.mapping_summary_labels["a_mapping"].cget("text"))
        self.assertNotIn("已完成",app.nav_buttons["recognize_a"].cget("text"))

    def test_映射读取失败或分类清单矛盾不能显示完成(self):
        app=self.make_app({**self.base,"current_step":"recognize_a"})
        cases=[None,"损坏JSON",{"schema":"附注语义映射-v1","carrier":"B"},
            {"schema":"附注语义映射-v1","carrier":"A","mappings":{},"excluded":[],"unresolved":[]},
            {"schema":"附注语义映射-v1","carrier":"A","mappings":[],"excluded":[],"unresolved":[{}],"complete":True}]
        for index,record in enumerate(cases):
            path=self.folder/f"无法读取映射{index}.json"
            if record is not None:path.write_text(record if isinstance(record,str) else json.dumps(record),encoding="utf-8-sig")
            app.values["a_mapping"].set(str(path))
            text=app.mapping_summary_labels["a_mapping"].cget("text")
            self.assertIn("不能读取",text);self.assertNotIn("完整成果",text)

    def test_管理标准与打开待处理说明分别执行且不保存配置(self):
        for carrier,step,key in (("A","recognize_a","a_standard_review"),("B","recognize_b","b_standard_review")):
            instructions=self.folder/(carrier+"待处理说明.md");instructions.write_text("只复用现有指标或修订真正缺失定义",encoding="utf-8-sig")
            package=self.folder/(carrier+"待处理材料.json")
            package.write_text(json.dumps({"schema":"附注待处理语义-v1","carrier":carrier,"instructions_path":str(instructions)},ensure_ascii=False),encoding="utf-8-sig")
            app=self.make_app({**self.base,key:str(package),"current_step":step})
            def descendants(parent):
                for child in parent.winfo_children():
                    yield child;yield from descendants(child)
            buttons={w.cget("text"):w for w in descendants(app.page) if isinstance(w,tk.Button)}
            self.assertIn("管理金标准",buttons);self.assertIn("打开待处理说明",buttons)
            self.assertNotIn("处理未确认项 / 管理金标准",buttons)
            with patch.object(界面.os,"startfile") as opening,patch.object(界面,"save_settings") as saving:
                buttons["打开待处理说明"].invoke();opening.assert_called_once_with(str(instructions));saving.assert_not_called()
                app.values[key].set(str(self.folder/"说明材料不存在.json"));opening.reset_mock()
                buttons["打开待处理说明"].invoke();opening.assert_not_called()
                self.assertIn("不能读取",界面.messagebox.showerror.call_args.args[1])

    def test_指标与实际维度说明和标准预览分层(self):
        for step in ("recognize_a","recognize_b"):
            text=界面.STEPS[step]["description"]
            self.assertIn("指标槽位",text);self.assertIn("实际维度",text)
            self.assertIn("新客户、新日期",text)
        before={"id":"S1","status":"active","note":{"name":"应收账款"},"table":{"name":"账龄分析"},
            "slot":{"name":"账面余额","meaning":"应收账款余额"},"dimensions":[{"name":"counterparty","open":True,"required":True}]}
        after=copy.deepcopy(before);after["dimensions"].append({"name":"period","type":"instant","required":True})
        text=界面.format_revision_preview({"changed":1,"changes":[{"kind":"changed","id":"S1","before":before,"after":after}]})
        self.assertEqual(text.count("指标定义："),2);self.assertEqual(text.count("维度要求："),2)
        self.assertIn("账面余额",text);self.assertIn("counterparty",text);self.assertIn('"open": true',text)
        self.assertIn("维度取值",text)


    def external_fixture(self,carrier):
        old=self.mapping(carrier,carrier+"复核前映射.json")
        opinion=self.folder/(carrier+"外部复核意见.json");opinion.write_text("独立测试意见",encoding="utf-8-sig")
        new=self.folder/(carrier+"复核后映射.json")
        review=self.folder/(carrier+"复核后核对表.xlsx");review.write_text("独立测试核对文件",encoding="utf-8-sig")
        package=self.folder/(carrier+"复核后待处理.json");package.write_text("{}",encoding="utf-8-sig")
        record={"schema":"附注语义映射-v1","carrier":carrier,"original_path":self.base["a_path" if carrier=="A" else "b_path"],
            "additional_context":"本轮已确认事实","mappings":[{"cell":"B3"},{"cell":"B4"}],"excluded":[],"unresolved":[],"complete":True,
            "review_path":str(review),"standard_review_path":str(package)}
        new.write_text(json.dumps(record,ensure_ascii=False),encoding="utf-8-sig")
        preview={"schema":"附注外部复核预览-v1","mapping_path":old,"mapping_sha256":"原映射哈希","opinion_path":str(opinion),
            "opinion_sha256":"意见哈希","gold_path":self.base["gold_path"],"carrier":carrier,"reviewer":{"name":"独立复核者","method":"advanced_ai","reviewed_at":"2026-09-20"},
            "count":1,"remaining_pending":0,"changes":[{"sheet":"表一","cell":"B4","previous_reason":"原两轮未确认",
                "mapping":{"slot_id":"C-N001-T001-S001","semantic_field":"账面余额","scope":"standalone",
                    "dimensions":{"period":"2025-12-31","counterparty":"示例客户"},"value_type":"monetary","evidence_cells":["A4"],"reason":"已核原始行列证据"}}]}
        prefix="a" if carrier=="A" else "b"
        result={"status":"complete","message":"已按复核意见另存","accepted_count":1,
            "files":{prefix+"_mapping":str(new),prefix+"_review":str(review),prefix+"_standard_review":str(package)}}
        return old,str(opinion),preview,result

    def descendants(self,parent):
        for child in parent.winfo_children():
            yield child;yield from self.descendants(child)

    def test_外部复核导出使用当前映射和保存窗口且取消不导出(self):
        from unittest.mock import Mock
        for carrier,step,key in (("A","recognize_a","a_mapping"),("B","recognize_b","b_mapping")):
            old,opinion,preview,result=self.external_fixture(carrier)
            app=self.make_app({**self.base,key:old,"current_step":step})
            buttons={w.cget("text"):w for w in self.descendants(app.page) if isinstance(w,tk.Button)}
            self.assertIn("导出复核意见格式",buttons);self.assertIn("导入复核意见",buttons)
            export=Mock(return_value=str(self.folder/(carrier+"导出格式.json")))
            with patch.dict(sys.modules,{"附注更新.外部复核":types.SimpleNamespace(export_review_template=export)}), \
                 patch.object(界面.filedialog,"asksaveasfilename",return_value=export.return_value) as dialog,patch.object(界面,"save_settings") as saving:
                buttons["导出复核意见格式"].invoke();export.assert_called_once_with(old,export.return_value,gold_path=self.base["gold_path"],include_confirmed=True)
                self.assertEqual(dialog.call_args.kwargs["defaultextension"],".json")
                self.assertEqual(app.values[key].get(),old);saving.assert_not_called()
                export.reset_mock();dialog.return_value="";buttons["导出复核意见格式"].invoke();export.assert_not_called()

    def test_外部复核仅预览可看到指标和维度取消不采纳(self):
        from unittest.mock import Mock
        old,opinion,preview,result=self.external_fixture("A")
        preview["changes"][0]["mapping"].update(table_range="A1:B4",table_semantic="货币资金期末余额明细",table_basis_hash="a"*64)
        app=self.make_app({**self.base,"a_mapping":old,"current_step":"recognize_a"})
        apply=Mock(return_value=result);backend=types.SimpleNamespace(preview_external_review=Mock(return_value=preview),apply_external_review=apply)
        with patch.dict(sys.modules,{"附注更新.外部复核":backend}),patch.object(界面.filedialog,"askopenfilename",return_value=opinion),patch.object(界面,"save_settings") as saving:
            window=app.import_external_review("A");self.assertIsInstance(window,tk.Toplevel)
            backend.preview_external_review.assert_called_once_with(old,opinion,gold_path=self.base["gold_path"])
            text="\n".join(w.get("1.0","end") for w in self.descendants(window) if isinstance(w,tk.Text))
            for expected in ("独立复核者","高级 AI","2026-09-20","1","0","指标槽位","C-N001-T001-S001","实际维度","2025-12-31","示例客户","不修改原件或金标准"):
                self.assertIn(expected,text)
            self.assertIn("采用的原表范围：表一!A1:B4",text)
            self.assertIn("原表业务说明：货币资金期末余额明细",text)
            self.assertNotIn("a"*64,text);self.assertNotIn("table_basis_hash",text)
            apply.assert_not_called();saving.assert_not_called();self.assertEqual(app.values["a_mapping"].get(),old)
            next(w for w in self.descendants(window) if isinstance(w,tk.Button) and w.cget("text")=="取消").invoke()
            self.assertFalse(window.winfo_exists());apply.assert_not_called();saving.assert_not_called()

    def test_更正预览明确展示原新指标和原维度(self):
        old,opinion,preview,result=self.external_fixture("B")
        preview["changes"][0]["previous_mapping"]={"slot_id":"C-N019-T005-S001","semantic_field":"狭义其他应付款","dimensions":{"period":"2024-12-31"}}
        text=界面.format_external_review_preview(preview)
        self.assertIn("更正已确认语义",text)
        self.assertIn("原指标槽位：C-N019-T005-S001；狭义其他应付款",text)
        self.assertIn("原实际维度",text);self.assertIn("2024-12-31",text)
        self.assertIn("指标槽位：C-N001-T001-S001",text)

    def test_补充业务上下文预览区分真实审核来源并展示回读原文(self):
        old,opinion,preview,result=self.external_fixture("B")
        preview["changes"][0].update(source_value=0,
            business_context={"range":"B4:B4","note_ids":["C-N001"],"scope":"standalone",
                "table_semantic":"应付职工薪酬总额","header_cells":["A4","B1"],"reason":"工资与福利费是下级分项"},
            context_evidence=[{"cell":"A4","text":"应付职工薪酬"},{"cell":"B1","text":"期末余额"}])
        for method,label in [("advanced_ai","高级 AI"),("human","人工")]:
            preview["reviewer"]["method"]=method
            text=界面.format_external_review_preview(preview)
            for expected in [label+"补充原格业务上下文","原格内容：0","A4：应付职工薪酬","B1：期末余额","工资与福利费是下级分项"]:
                self.assertIn(expected,text)
            self.assertNotIn("模型双轮",text)

    def test_确认外部复核才另存并选用新映射及配套文件(self):
        from unittest.mock import Mock
        for carrier,step,key in (("A","recognize_a","a_mapping"),("B","recognize_b","b_mapping")):
            old,opinion,preview,result=self.external_fixture(carrier)
            app=self.make_app({**self.base,key:old,"current_step":step})
            apply=Mock(return_value=result);backend=types.SimpleNamespace(preview_external_review=Mock(return_value=preview),apply_external_review=apply)
            with patch.dict(sys.modules,{"附注更新.外部复核":backend}),patch.object(界面.filedialog,"askopenfilename",return_value=opinion):
                window=app.import_external_review(carrier)
                next(w for w in self.descendants(window) if isinstance(w,tk.Button) and w.cget("text")=="确认另存并选用").invoke()
            apply.assert_called_once_with(preview,str(self.folder),confirmed=True,confirm_method="界面确认导入外部复核意见")
            for field,path in result["files"].items():self.assertEqual(app.values[field].get(),path)
            self.assertEqual(load_settings(self.config_path)[key],result["files"][key])
            self.assertIn("业务 2 格",app.mapping_summary_labels[key].cget("text"))
            self.assertFalse(window.winfo_exists());self.assertFalse(app.running)

    def test_外部复核运行中不可导入且过期意见明确拒绝(self):
        from unittest.mock import Mock
        old,opinion,preview,result=self.external_fixture("B")
        app=self.make_app({**self.base,"b_mapping":old,"current_step":"recognize_b"})
        apply=Mock(side_effect=ValueError("原映射内容已变化，预览过期"))
        backend=types.SimpleNamespace(preview_external_review=Mock(return_value=preview),apply_external_review=apply)
        with patch.dict(sys.modules,{"附注更新.外部复核":backend}),patch.object(界面.filedialog,"askopenfilename",return_value=opinion) as dialog,patch.object(界面,"save_settings") as saving:
            app.running_step="recognize_b";app.set_running(True)
            self.assertIsNone(app.import_external_review("B"));dialog.assert_not_called();apply.assert_not_called()
            app.set_running(False);window=app.import_external_review("B")
            app.values["gold_path"].set(str(self.folder/"另选标准.jsonl"))
            next(w for w in self.descendants(window) if isinstance(w,tk.Button) and w.cget("text")=="确认另存并选用").invoke()
            apply.assert_not_called();self.assertIn("金标准已变化",界面.messagebox.showerror.call_args.args[1])
            app.values["gold_path"].set(self.base["gold_path"])
            next(w for w in self.descendants(window) if isinstance(w,tk.Button) and w.cget("text")=="确认另存并选用").invoke()
            self.assertIn("预览过期",界面.messagebox.showerror.call_args.args[1])
            self.assertEqual(app.values["b_mapping"].get(),old);saving.assert_not_called();window.destroy()
            backend.preview_external_review.side_effect=ValueError("复核意见缺少实际期间")
            self.assertIsNone(app.import_external_review("B"));self.assertIn("实际期间",界面.messagebox.showerror.call_args.args[1])



    def scope_fixture(self):
        old=self.mapping("B","范围复核前映射.json")
        link=self.folder/"虚构位置关联.json";link.write_text("{}",encoding="utf-8-sig")
        opinion=self.folder/"范围意见.json";opinion.write_text("{}",encoding="utf-8-sig")
        scope=self.folder/"范围确认成果.json";scope.write_text("{}",encoding="utf-8-sig")
        preview={"source_path":self.base["b_path"],"a_link_path":str(link),"gold_path":self.base["gold_path"],
            "opinion_path":str(opinion),"reviewer":{"name":"范围审核者","method":"advanced_ai","reviewed_at":"2026-09-20"},
            "coverage":{"candidate_count":10,"selected_count":6,"out_of_scope_count":4},"uncertain_count":2,
            "changes":[{"sheet":"来源","range":"A4:B5","decision":"out_of_scope","reason":"人数与货币金额含义不同",
                "evidence_cells":["A4","B4"],"semantic_comparison":{"source_business":"职工人数","target_business":"货币资金",
                "difference_kind":"value_type","reason":"计量性质不同","target_evidence":["货币资金"]}}]}
        settings={**self.base,"a_link":str(link),"b_mapping":old,"b_review":"旧核对表.xlsx","b_scope":"旧范围.json",
            "b_gold_changes":"旧补充.json","b_standard_review":"旧待处理.json","update_plan":"旧清单.json",
            "a_mapping":"原附注映射.json","a_prime":"旧更新表.xlsx","word":"旧输出.docx","current_step":"recognize_b"}
        result={"status":"complete","message":"范围已保存","files":{"b_scope":str(scope)},"coverage":preview["coverage"]}
        return settings,preview,result

    def test_范围复核导出使用原更新数据位置关联与当前标准(self):
        from unittest.mock import Mock
        settings,preview,result=self.scope_fixture();app=self.make_app(settings)
        buttons={w.cget("text"):w for w in self.descendants(app.page) if isinstance(w,tk.Button)}
        self.assertIn("导出范围复核格式",buttons);self.assertIn("导入范围复核意见",buttons)
        exported=str(self.folder/"导出范围意见.json");export=Mock(return_value=exported)
        with patch.dict(sys.modules,{"附注更新.外部范围":types.SimpleNamespace(export_scope_review_template=export)}), \
             patch.object(界面.filedialog,"asksaveasfilename",return_value=exported) as dialog,patch.object(界面,"save_settings") as saving:
            buttons["导出范围复核格式"].invoke()
            export.assert_called_once_with(settings["b_path"],settings["a_link"],exported,gold_path=settings["gold_path"])
            self.assertEqual(dialog.call_args.kwargs["defaultextension"],".json")
            saving.assert_not_called();self.assertEqual(app.values["b_mapping"].get(),settings["b_mapping"])
            export.reset_mock();dialog.return_value="";buttons["导出范围复核格式"].invoke();export.assert_not_called()

    def test_范围复核只预览可取消并明确旧映射需重新识别(self):
        from unittest.mock import Mock
        settings,preview,result=self.scope_fixture();app=self.make_app(settings)
        apply=Mock(return_value=result);backend=types.SimpleNamespace(preview_external_scope_review=Mock(return_value=preview),apply_external_scope_review=apply)
        with patch.dict(sys.modules,{"附注更新.外部范围":backend}),patch.object(界面.filedialog,"askopenfilename",return_value=preview["opinion_path"]),patch.object(界面,"save_settings") as saving:
            window=app.import_scope_review();self.assertIsInstance(window,tk.Toplevel)
            backend.preview_external_scope_review.assert_called_once_with(settings["b_path"],settings["a_link"],preview["opinion_path"],gold_path=settings["gold_path"])
            text="\n".join(w.get("1.0","end") for w in self.descendants(window) if isinstance(w,tk.Text))
            for expected in ("范围审核者","高级 AI","10","6","4","2","原文件保留","重新识别","未完成业务语义识别","职工人数","货币资金","A4"):
                self.assertIn(expected,text)
            next(w for w in self.descendants(window) if isinstance(w,tk.Button) and w.cget("text")=="取消").invoke()
            apply.assert_not_called();saving.assert_not_called();self.assertEqual(app.values["b_mapping"].get(),settings["b_mapping"])

    def test_确认范围清旧更新映射与下游并保存重启但不标识别完成(self):
        from unittest.mock import Mock
        settings,preview,result=self.scope_fixture();app=self.make_app(settings)
        apply=Mock(return_value=result);backend=types.SimpleNamespace(preview_external_scope_review=Mock(return_value=preview),apply_external_scope_review=apply)
        app.step_states["recognize_b"]="complete"
        with patch.dict(sys.modules,{"附注更新.外部范围":backend}),patch.object(界面.filedialog,"askopenfilename",return_value=preview["opinion_path"]):
            window=app.import_scope_review()
            next(w for w in self.descendants(window) if isinstance(w,tk.Button) and w.cget("text")=="确认范围并选用").invoke()
        apply.assert_called_once_with(preview,str(self.folder),confirmed=True,confirm_method="界面确认导入范围复核意见")
        for key in ("b_mapping","b_review","b_gold_changes","b_standard_review","update_plan")+界面.FINAL_FIELDS:self.assertEqual(app.values[key].get(),"")
        for key in ("a_mapping","a_link","b_path","b_context_text","gold_path"):self.assertEqual(app.values[key].get(),settings[key])
        self.assertEqual(app.values["b_scope"].get(),result["files"]["b_scope"]);self.assertTrue(app.scope_prefilter.get())
        self.assertNotEqual(app.step_states.get("recognize_b"),"complete")
        self.assertNotIn("已完成",app.nav_buttons["recognize_b"].cget("text"));self.assertTrue(Path(settings["b_mapping"]).is_file())
        saved=load_settings(self.config_path);reopened=self.make_app(saved)
        self.assertEqual(reopened.values["b_mapping"].get(),"");self.assertEqual(reopened.values["b_scope"].get(),result["files"]["b_scope"])
        self.assertTrue(reopened.scope_prefilter.get())

    def test_范围复核忙时和预览后材料变化及过期意见不能采纳(self):
        from unittest.mock import Mock
        settings,preview,result=self.scope_fixture();app=self.make_app(settings)
        apply=Mock(side_effect=ValueError("意见或原件已经变化，请重新预览"))
        backend=types.SimpleNamespace(preview_external_scope_review=Mock(return_value=preview),apply_external_scope_review=apply)
        with patch.dict(sys.modules,{"附注更新.外部范围":backend}),patch.object(界面.filedialog,"askopenfilename",return_value=preview["opinion_path"]) as dialog,patch.object(界面,"save_settings") as saving:
            app.running_step="recognize_b";app.set_running(True)
            self.assertIsNone(app.import_scope_review());dialog.assert_not_called();apply.assert_not_called()
            app.set_running(False);window=app.import_scope_review()
            for key in ("b_path","a_link","gold_path"):
                app.values[key].set(str(self.folder/(key+"变更")))
                self.assertFalse(app.confirm_scope_review(preview,window));apply.assert_not_called()
                app.values[key].set(settings[key])
            app.set_running(True);self.assertFalse(app.confirm_scope_review(preview,window));apply.assert_not_called();app.set_running(False)
            self.assertFalse(app.confirm_scope_review(preview,window));self.assertIn("已经变化",界面.messagebox.showerror.call_args.args[1])
            self.assertEqual(app.values["b_mapping"].get(),settings["b_mapping"]);saving.assert_not_called();window.destroy()
            backend.preview_external_scope_review.side_effect=ValueError("范围意见遗漏原格")
            self.assertIsNone(app.import_scope_review());self.assertIn("遗漏原格",界面.messagebox.showerror.call_args.args[1])

    def test_真实范围后端与Tk确认保存可重载且不请求网络(self):
        from 测试.测试外部范围 import ExternalScopeTests
        from 附注更新.外部范围 import read_verified_scope
        from 附注更新.表格 import read_workbook,file_hash
        case=ExternalScopeTests();case.setUp()
        with patch("urllib.request.build_opener",side_effect=AssertionError("禁止外部模型请求")):
            opinion=case.save("界面模拟范围意见",case.opinion())
            old=self.mapping("B","真实入口前旧映射.json")
            settings={**self.base,"b_path":str(case.source),"a_link":str(case.link),"gold_path":str(case.gold),
                "b_mapping":old,"current_step":"recognize_b"}
            app=self.make_app(settings)
            with patch.object(界面.filedialog,"askopenfilename",return_value=str(opinion)):
                window=app.import_scope_review()
                next(w for w in self.descendants(window) if isinstance(w,tk.Button) and w.cget("text")=="确认范围并选用").invoke()
            selected=json.loads(Path(app.values["b_scope"].get()).read_text(encoding="utf-8-sig"))
            checked=read_verified_scope(selected,read_workbook(case.source),selected["target_meanings"])
            self.assertEqual(checked["coverage"],{"candidate_count":10,"selected_count":6,"out_of_scope_count":4})
            self.assertIn("B3",checked["selected_cells"]["混合资料"])
            self.assertEqual(app.values["b_mapping"].get(),"");self.assertTrue(app.scope_prefilter.get())
            self.assertNotEqual(app.step_states.get("recognize_b"),"complete")
            saved=load_settings(self.config_path);self.assertEqual(saved["b_scope"],app.values["b_scope"].get())
            self.assertTrue(saved["scope_prefilter"])
            self.assertEqual(case.original,{p:file_hash(p) for p in case.original})

if __name__=="__main__":unittest.main()
