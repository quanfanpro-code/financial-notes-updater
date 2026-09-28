"""模型设置的输入边界、密钥保护和持久化验证。"""
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from 附注更新.设置 import load_settings, protect_secret, save_settings, unprotect_secret, validate_settings


def _destroy_test_window(root):
    """测试实例退出前取消全部定时和空闲回调，避免串入下一个 Tk 窗口。"""
    for pending in root.tk.call("after", "info"):
        root.tk.call("after", "cancel", pending)
    root.destroy()


class 设置测试(unittest.TestCase):
    def test_结构调整成果在勾稽步骤独立显示保存并随来源变更清空(self):
        from 附注更新.界面 import STEPS,FILE_FIELDS,file_selection_updates
        fields=('structure_review','adjusted_a','adjusted_a_mapping')
        for field in fields:
            self.assertIn(field,STEPS['match']['outputs']);self.assertIn(field,FILE_FIELDS)
        selected={field:'此前成果文件' for field in fields}
        for changed in ('a_path','b_path','gold_path','a_link','update_plan'):
            changed_values=file_selection_updates(changed,'', {**selected,changed:'旧材料'})
            self.assertTrue(all(changed_values.get(field)=='' for field in fields))
        folder=Path(tempfile.mkdtemp(prefix='结构成果设置_'));path=folder/'设置.json'
        save_settings({**selected,'base_url':'http://localhost','model':'受控模型','api_key':''},path)
        loaded=load_settings(path)
        self.assertEqual({field:loaded[field] for field in fields},selected)

    def test_界面选择新版独立映射和清单会恢复各自来源(self):
        from 附注更新.界面 import file_selection_updates
        folder=Path(tempfile.mkdtemp(prefix='新版界面来源_'))
        paths=[]
        for carrier in ('A','B'):
            path=folder/(carrier+'映射.json')
            record={'schema_version':2,'carrier':carrier,'source_path':carrier+'.xlsx',
                    'source_hash':'a'*64,'gold_path':'统一标准.jsonl','gold_hash':'b'*64,'facts':[],'owners':{},
                    'review_path':carrier+'核对.xlsx','standard_review_path':carrier+'待审.json'}
            if carrier=='A':record['a_link']='关联.json'
            path.write_text(json.dumps(record,ensure_ascii=False),encoding='utf-8-sig');paths.append(path)
        result=file_selection_updates('b_mapping',str(paths[1]),{})
        self.assertEqual(result['b_path'],'B.xlsx');self.assertEqual(result['b_review'],'B核对.xlsx')
        self.assertEqual(result['b_standard_review'],'B待审.json')
        plan=folder/'更新清单.json';plan.write_text(json.dumps({'schema':'附注更新清单-v2','a_mapping':str(paths[0]),
            'b_mapping':str(paths[1]),'a_link':'关联.json','structure':{'report_path':'结构复核.json','mapping_path':'结构后映射.json',
            'layout':{'source_path':'结构后.xlsx'}}},ensure_ascii=False),encoding='utf-8-sig')
        result=file_selection_updates('update_plan',str(plan),{})
        self.assertEqual(result['a_path'],'A.xlsx');self.assertEqual(result['b_path'],'B.xlsx')
        self.assertEqual(result['a_link'],'关联.json')
        self.assertEqual(result['adjusted_a_mapping'],'结构后映射.json');self.assertEqual(result['adjusted_a'],'结构后.xlsx')
        self.assertEqual(result['structure_review'],'结构复核.json')
        with self.assertRaisesRegex(ValueError,'更新数据'):file_selection_updates('b_mapping',str(paths[0]),{})

    def test_拒绝非法模型地址(self):
        for url in ("", "file:///secret", "https://", "https://user:secret@example.com/v1", "https://api.example.com/v1?key=secret"):
            with self.subTest(url=url), self.assertRaises(ValueError):
                validate_settings({"base_url": url, "model": "模型"})

    def test_拒绝缺少模型(self):
        with self.assertRaisesRegex(ValueError, "模型名称"):
            validate_settings({"base_url": "https://api.example.com/v1", "model": " "})

    def test_允许本机接口且规范化地址(self):
        result = validate_settings({"base_url": " http://127.0.0.1:8000/v1/ ", "model": " test-model ", "api_key": ""})
        self.assertEqual(result["base_url"], "http://127.0.0.1:8000/v1")
        self.assertEqual(result["model"], "test-model")

    def test_请求超时默认600秒且允许正数(self):
        self.assertEqual(validate_settings({}, require_model=False)["timeout"], 600)
        for value, expected in ((600, 600), (" 900.5 ", 900.5), (0.25, 0.25)):
            with self.subTest(value=value):
                self.assertEqual(validate_settings({"timeout": value}, require_model=False)["timeout"], expected)
        for value in (0, -1, True, False, None, "", "abc", "nan", "inf", float("-inf"), {}, 10 ** 1000):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "请求超时"):
                validate_settings({"timeout": value}, require_model=False)

    def test_请求超时持久化并兼容旧配置(self):
        folder = Path(tempfile.mkdtemp(prefix="请求超时设置_"))
        old = folder / "旧配置.json"
        self.assertEqual(load_settings(old)["timeout"], 600)
        old.write_text("{}", encoding="utf-8-sig")
        self.assertEqual(load_settings(old)["timeout"], 600)
        target = folder / "新配置.json"
        save_settings({"timeout": "900.5"}, target)
        self.assertEqual(json.loads(target.read_text(encoding="utf-8-sig"))["timeout"], 900.5)
        self.assertEqual(load_settings(target)["timeout"], 900.5)
        for index, value in enumerate((0, -1, True, None, "nan", "abc")):
            invalid = folder / f"无效配置{index}.json"
            invalid.write_text(json.dumps({"timeout": value}), encoding="utf-8-sig")
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "请求超时"):
                load_settings(invalid)

    def test_界面请求超时可编辑并校验保存(self):
        import tkinter as tk
        from unittest.mock import patch
        from 附注更新 import 界面
        root = tk.Tk(); root.withdraw()
        try:
            with patch.object(界面, "load_settings", return_value={}):
                app = 界面.Application(root)
            self.assertEqual(app.values["timeout"].get(), "600")
            app.show_step("settings")
            entries = [widget for widget, _ in app.page_editable
                       if isinstance(widget, 界面.ttk.Entry) and str(widget.cget("textvariable")) == str(app.values["timeout"])]
            self.assertEqual(len(entries), 1)
            labels = [child.cget("text") for child in entries[0].master.winfo_children() if isinstance(child, tk.Label)]
            self.assertIn("请求超时（秒）", labels)
            app.values["timeout"].set("900.5")
            with patch.object(界面, "save_settings") as save:
                app.save()
                self.assertEqual(save.call_args.args[0]["timeout"], 900.5)
            app.values["timeout"].set("0")
            with patch.object(界面, "save_settings") as save, patch.object(界面.messagebox, "showerror") as error:
                app.save()
                save.assert_not_called()
                self.assertIn("请求超时", error.call_args.args[1])
        finally:
            _destroy_test_window(root)

    def test_思考模式默认与持久化及非法值(self):
        self.assertEqual(validate_settings({},require_model=False).get("thinking_mode"),"default")
        folder=Path(tempfile.mkdtemp(prefix="思考模式测试_"))
        self.assertEqual(load_settings(folder/"尚无配置.json").get("thinking_mode"),"default")
        for mode in ("default","enabled","disabled"):
            path=folder/(mode+".json")
            save_settings({"thinking_mode":mode},path)
            self.assertEqual(load_settings(path)["thinking_mode"],mode)
        for index,value in enumerate(("",None,True,False,"on","关闭",{})):
            with self.subTest(value=value):
                with self.assertRaises(ValueError):validate_settings({"thinking_mode":value},require_model=False)
                path=folder/("无效模式"+str(index)+".json")
                path.write_text(json.dumps({"thinking_mode":value}),encoding="utf-8-sig")
                with self.assertRaises(ValueError):load_settings(path)

    def test_界面思考模式中文选择保存恢复(self):
        import tkinter as tk
        from unittest.mock import patch
        from 附注更新 import 界面
        root=tk.Tk();root.withdraw()
        try:
            with patch.object(界面,"load_settings",return_value={"thinking_mode":"disabled"}):app=界面.Application(root)
            self.assertIn("thinking_mode",app.values)
            app.show_step("settings")
            combos=[w for w,_ in app.page_editable if isinstance(w,界面.ttk.Combobox) and tuple(w.cget("values"))==("服务默认","开启","关闭")]
            self.assertEqual(len(combos),1)
            combo=combos[0];self.assertEqual(combo.get(),"关闭")
            self.assertEqual(str(combo.cget("state")),"readonly")
            combo.set("开启");combo.event_generate("<<ComboboxSelected>>")
            with patch.object(界面,"save_settings") as save:
                app.save();self.assertEqual(save.call_args.args[0]["thinking_mode"],"enabled")
            app.show_step("recognize_a");app.show_step("settings")
            combos=[w for w,_ in app.page_editable if isinstance(w,界面.ttk.Combobox) and tuple(w.cget("values"))==("服务默认","开启","关闭")]
            self.assertEqual(combos[0].get(),"开启")
        finally:_destroy_test_window(root)

    def test_DPAPI往返(self):
        secret = "测试密钥-仅供测试-123456"
        encrypted = protect_secret(secret)
        self.assertNotIn(secret, encrypted)
        self.assertEqual(unprotect_secret(encrypted), secret)
        self.assertEqual(unprotect_secret(protect_secret("")), "")

    def test_配置及备份不存明文密钥(self):
        # 保留测试产物以便核验，不删除临时目录。
        folder = Path(tempfile.mkdtemp(prefix="附注更新设置测试_"))
        path = folder / "用户配置.json"
        secret = "test-secret-only-20260919"
        config = {"base_url": "https://api.example.com/v1", "model": "demo", "api_key": secret, "gold_path": "示例.jsonl"}
        save_settings(config, path)
        saved = path.read_text(encoding="utf-8-sig")
        self.assertNotIn(secret, saved)
        self.assertNotIn("api_key", json.loads(saved))
        self.assertEqual(load_settings(path)["api_key"], secret)
        save_settings(dict(config, model="demo2"), path)
        self.assertEqual(load_settings(path)["model"], "demo2")

    def test_没有模型也可保存五步材料路径(self):
        folder = Path(tempfile.mkdtemp(prefix="分步材料配置_"))
        path = folder / "用户配置.json"
        fields = {"word_path":"文档一.docx","a_path":"表格A.xlsx","b_path":"表格B.xlsx",
                  "a_link":"位置关联.json","a_mapping":"A映射.json","b_mapping":"B映射.json",
                  "update_plan":"更新清单.json","a_prime":"表格A更新.xlsx",
                  "a_prime_mapping":"A更新映射.json","word":"更新Word.docx","word_link":"更新位置关联.json","trace":"逐格核对.xlsx","a_review":"A核对.xlsx","b_review":"B核对.xlsx"}
        try:
            save_settings({**fields,"base_url":"","model":"","api_key":""},path)
        except ValueError as error:
            self.fail("非模型步骤的材料配置不应要求模型：" + str(error))
        saved = load_settings(path)
        for key,value in fields.items():self.assertEqual(saved[key],value)

    def test_旧配置迁移到相应步骤但不替换已有新字段(self):
        folder = Path(tempfile.mkdtemp(prefix="分步迁移配置_"))
        path = folder / "旧Word配置.json"
        path.write_text(json.dumps({"source_a":"原报告.docx","source_b":"新数据.xlsx"},ensure_ascii=False),encoding="utf-8-sig")
        value = load_settings(path)
        self.assertEqual(value.get("word_path"),"原报告.docx")
        self.assertEqual(value.get("b_path"),"新数据.xlsx")
        second = folder / "旧Excel配置.json"
        second.write_text(json.dumps({"source_a":"旧A.xlsx","a_path":"已选A.xlsx"},ensure_ascii=False),encoding="utf-8-sig")
        self.assertEqual(load_settings(second).get("a_path"),"已选A.xlsx")

    def test_五步界面按本步骤检查前置文件(self):
        from 附注更新 import 界面
        self.assertTrue(hasattr(界面,"missing_inputs"),"界面尚未具备分步前置文件检查")
        folder = Path(tempfile.mkdtemp(prefix="步骤前置文件_"))
        word = folder / "文档一.docx";word.write_bytes(b"test")
        plan = folder / "更新清单.json";plan.write_text("{}",encoding="utf-8-sig")
        config = {"word_path":str(word),"update_plan":str(plan),"output_dir":str(folder)}
        self.assertEqual(界面.missing_inputs("extract_a",config),[])
        self.assertEqual(界面.missing_inputs("write",config),[])
        self.assertIn("附注表格语义映射", "；".join(界面.missing_inputs("match",config)))
        self.assertIn("模型", "；".join(界面.missing_inputs("recognize_a",config)))

    def test_隐藏窗口可切换步骤并显示独立成果(self):
        import tkinter as tk
        from unittest.mock import patch
        from 附注更新 import 界面
        root = tk.Tk();root.withdraw()
        try:
            with patch.object(界面,"load_settings",return_value={}):
                app = 界面.Application(root)
            self.assertTrue(hasattr(app,"show_step"),"界面尚未支持五步导航")
            self.assertTrue(hasattr(app,"apply_result"),"界面尚未支持独立成果路径回填")
            app.show_step("recognize_b")
            self.assertEqual(app.current_step,"recognize_b")
            self.assertIn("更新数据",app.start_button.cget("text"))
            app.apply_result("extract_a",{"status":"complete","message":"已提取",
                "files":{"a_path":r"C:\\材料\\A.xlsx","a_link":r"C:\\材料\\位置.json"}})
            self.assertEqual(app.values["a_path"].get(),r"C:\\材料\\A.xlsx")
            app.show_step("write")
            app.apply_result("write",{"status":"complete","message":"已生成",
                "files":{"a_prime":r"C:\\结果\\A更新.xlsx","a_prime_mapping":r"C:\\结果\\A更新映射.json",
                         "word":r"C:\\结果\\报告.docx","word_link":r"C:\\结果\\位置.json"}})
            self.assertEqual(app.values["word_link"].get(),r"C:\\结果\\位置.json")
            self.assertEqual(app.step_states["write"],"complete")
        finally:_destroy_test_window(root)

    def test_更换上游材料会清除下游成果(self):
        from 附注更新.界面 import file_selection_updates
        before={"word_path":"旧.docx","a_path":"旧A.xlsx","b_path":"旧B.xlsx","a_context_text":"A旧信息","b_context_text":"B旧信息",
                "a_link":"旧关联.json","a_mapping":"A映射.json","b_mapping":"B映射.json","a_review":"A核对.xlsx","b_review":"B核对.xlsx",
                "update_plan":"旧清单.json","a_prime":"旧成果.xlsx","trace":"旧核对.xlsx"}
        for key,value,empty,kept in (
            ("b_path","新B.xlsx",("b_mapping","b_review","b_context_text","update_plan","a_prime","trace"),("a_mapping","a_link")),
            ("a_path","新A.xlsx",("a_mapping","a_review","a_context_text","a_link","update_plan","a_prime"),("b_mapping",)),
            ("word_path","新.docx",("a_path","a_mapping","a_review","a_context_text","a_link","update_plan","a_prime"),("b_mapping",))):
            result={**before,**file_selection_updates(key,value,before)}
            for field in empty:self.assertEqual(result[field],"",(key,field))
            for field in kept:self.assertEqual(result[field],before[field],(key,field))
        self.assertEqual(file_selection_updates("a_path",before["a_path"],before),{})

    def test_复用映射或清单展示其绑定来源(self):
        from 附注更新.界面 import file_selection_updates
        folder=Path(tempfile.mkdtemp(prefix="界面来源复用_"))
        a=folder/"A映射.json";b=folder/"B映射.json";plan=folder/"清单.json"
        a.write_text(json.dumps({"schema":"附注语义映射-v1","carrier":"A","source_path":"A.xlsx","a_link":"关联.json"}),encoding="utf-8-sig")
        b.write_text(json.dumps({"schema":"附注语义映射-v1","carrier":"B","source_path":"重算B.xlsx","original_path":"原B.xlsx","review_path":"B核对.xlsx"}),encoding="utf-8-sig")
        plan.write_text(json.dumps({"schema":"附注更新清单-v1","a_mapping":str(a),"b_mapping":str(b),"a_link":"关联.json"}),encoding="utf-8-sig")
        result=file_selection_updates("b_mapping",str(b),{"b_path":"无关B.xlsx"})
        self.assertEqual(result["b_path"],"原B.xlsx");self.assertEqual(result["b_review"],"B核对.xlsx")
        result=file_selection_updates("update_plan",str(plan),{"a_path":"无关A.xlsx","b_path":"无关B.xlsx"})
        self.assertEqual(result["a_mapping"],str(a));self.assertEqual(result["b_mapping"],str(b))
        self.assertEqual(result["a_path"],"A.xlsx");self.assertEqual(result["b_path"],"原B.xlsx")
        self.assertEqual(result["a_link"],"关联.json")
        with self.assertRaisesRegex(ValueError,"更新数据"):
            file_selection_updates("b_mapping",str(a),{})

    def test_换目标资料清除旧范围但保留独立更新数据映射(self):
        from 附注更新.界面 import file_selection_updates
        before={"word_path":"原.docx","a_path":"原附注.xlsx","a_link":"原关联.json",
                "b_scope":"旧范围.json","b_mapping":"有效更新映射.json","b_path":"更新数据.xlsx",
                "a_gold_changes":"原附注新增.json","b_gold_changes":"更新新增.json"}
        for key,value in (("word_path","新.docx"),("a_path","新附注.xlsx"),("a_link","新关联.json")):
            with self.subTest(key=key):
                result={**before,**file_selection_updates(key,value,before)}
                self.assertEqual(result["b_scope"],"")
                for kept in ("b_path","b_mapping","b_gold_changes"):
                    self.assertEqual(result[kept],before[kept])
                if key!="a_link":self.assertEqual(result["a_gold_changes"],"")
        self.assertEqual(file_selection_updates("a_link",before["a_link"],before),{})

    def test_加载映射恢复自己的范围及标准记录且缺字段清空旧路径(self):
        from unittest.mock import patch
        from 附注更新.界面 import file_selection_updates
        before={"b_scope":"无关范围.json","a_gold_changes":"无关附注新增.json",
                "b_gold_changes":"无关更新新增.json","a_path":"旧附注.xlsx","b_path":"旧更新.xlsx"}
        for carrier,key in (("A","a_mapping"),("B","b_mapping")):
            for has_paths in (True,False):
                record={"schema":"附注语义映射-v1","carrier":carrier,"source_path":"本份.xlsx"}
                if has_paths:record.update(scope_selection_path="本份范围.json",gold_learning_path="本份新增.json")
                with self.subTest(carrier=carrier,has_paths=has_paths), patch("附注更新.界面._read_record",return_value=record):
                    result={**before,**file_selection_updates(key,"所选映射.json",before)}
                own="a_gold_changes" if carrier=="A" else "b_gold_changes"
                other="b_gold_changes" if carrier=="A" else "a_gold_changes"
                self.assertEqual(result[own],"本份新增.json" if has_paths else "")
                self.assertEqual(result[other],before[other])
                self.assertEqual(result["b_scope"],"本份范围.json" if carrier=="B" and has_paths else "")

    def test_进入下一轮清除旧范围和上轮新增记录而不启动窗口(self):
        from types import SimpleNamespace
        from unittest.mock import patch
        from 附注更新.界面 import Application
        prime=str(Path(r"C:\材料\更新后附注.xlsx"))
        values={key:SimpleNamespace(get=lambda value=value:value) for key,value in {
            "a_prime":prime,"a_prime_mapping":"新映射.json"}.items()}
        changes={}
        app=SimpleNamespace(running=False,values=values,root=None,
            apply_values=changes.update,append_log=lambda _:None,show_step=lambda _:None)
        record={"schema":"附注语义映射-v1","carrier":"A","source_path":prime}
        with patch("附注更新.界面._read_record",return_value=record),patch.object(Path,"is_file",return_value=True):
            Application.use_updated_a(app)
        for key in ("b_scope","a_gold_changes","b_gold_changes"):
            self.assertEqual(changes.get(key),"")
        self.assertEqual(changes["a_path"],prime)

    def test_只执行选中步骤且不混入上轮Word成果(self):
        import tkinter as tk
        from unittest.mock import patch
        from 附注更新 import 界面
        root=tk.Tk();root.withdraw()
        try:
            with patch.object(界面,"load_settings",return_value={}):app=界面.Application(root)
            with patch("附注更新.分步流程.run_step",return_value={"status":"complete","files":{}}) as call:
                app.run("extract_a",{"word_path":"原.docx"})
            self.assertEqual(call.call_count,1);self.assertEqual(call.call_args.args[0],"extract_a")
            app.apply_result("match",{"status":"complete","files":{"update_plan":"清单.json","trace":"核对.xlsx"}})
            app.values["word"].set("旧报告.docx");app.values["word_link"].set("旧关联.json")
            app.apply_result("write",{"status":"complete","files":{"a_prime":"新A.xlsx","a_prime_mapping":"新映射.json"}})
            self.assertEqual(app.values["word"].get(),"");self.assertEqual(app.values["word_link"].get(),"")
            self.assertEqual(app.step_states.get("match"),"complete")
        finally:_destroy_test_window(root)

    def test_明确选择A更新后才开启下一轮(self):
        import tkinter as tk
        from unittest.mock import patch
        from 附注更新 import 界面
        folder=Path(tempfile.mkdtemp(prefix="界面下一轮_"))
        prime=folder/"新A.xlsx";prime.write_bytes(b"test")
        mapping=folder/"新A映射.json"
        mapping.write_text(json.dumps({"schema":"附注语义映射-v1","carrier":"A","source_path":str(prime)}),encoding="utf-8-sig")
        root=tk.Tk();root.withdraw()
        try:
            with patch.object(界面,"load_settings",return_value={}):app=界面.Application(root)
            app.values["a_path"].set("原A.xlsx")
            app.values["b_path"].set("旧B.xlsx");app.values["b_mapping"].set("旧B映射.json")
            app.values["update_plan"].set("旧清单.json")
            app.apply_result("write",{"status":"complete","files":{"a_prime":str(prime),"a_prime_mapping":str(mapping)}})
            self.assertEqual(app.values["a_path"].get(),"原A.xlsx")
            app.use_updated_a()
            self.assertEqual(app.values["a_path"].get(),str(prime));self.assertEqual(app.values["a_mapping"].get(),str(mapping))
            for key in ("b_path","b_mapping","b_review","update_plan","a_link","a_review"):self.assertEqual(app.values[key].get(),"")
            self.assertEqual(app.current_step,"recognize_b");self.assertIsNone(app.worker)
            link=folder/"新位置关联.json"
            link.write_text(json.dumps({"schema":"附注Word关联-v1","word_path":"真正新Word.docx","a_path":str(prime)}),encoding="utf-8-sig")
            linked_map=folder/"含位置新A映射.json"
            linked_map.write_text(json.dumps({"schema":"附注语义映射-v1","carrier":"A","source_path":str(prime),"a_link":str(link)}),encoding="utf-8-sig")
            app.values["a_prime_mapping"].set(str(linked_map));app.values["word"].set("无关Word.docx")
            app.use_updated_a()
            self.assertEqual(app.values["word_path"].get(),"真正新Word.docx")
        finally:_destroy_test_window(root)

    def test_程序内金标准与分别补充信息迁移(self):
        from 附注更新.设置 import DEFAULT_GOLD_PATH,PROJECT_DIR
        folder=Path(tempfile.mkdtemp(prefix="标准及说明迁移_"));old=folder/"旧配置.json"
        previous=r"D:\BaiduSyncdisk\数字资产\审计数据自动搬运系统\语义金标准\财务报表附注语义金标准-v1.jsonl"
        old.write_text(json.dumps({"gold_path":previous,"context_text":"原共同说明","a_context_text":"A独有信息"},ensure_ascii=False),encoding="utf-8-sig")
        values=load_settings(old)
        完整标准 = PROJECT_DIR/"金标准"/"财务报表附注语义金标准-v2.jsonl"
        预期标准 = 完整标准 if 完整标准.is_file() else PROJECT_DIR/"金标准"/"公开语义金标准-v2.jsonl"
        self.assertEqual(Path(DEFAULT_GOLD_PATH), 预期标准)
        self.assertTrue(预期标准.is_file())
        self.assertEqual(values["gold_path"],str(DEFAULT_GOLD_PATH))
        self.assertEqual(values["a_context_text"],"A独有信息");self.assertEqual(values["b_context_text"],"原共同说明")
        target=folder/"新配置.json"
        save_settings({**values,"b_context_text":"B自己的单位","a_standard_review":"A待处理.json","b_standard_review":"B待处理.json"},target)
        payload=json.loads(target.read_text(encoding="utf-8-sig"))
        self.assertNotIn("context_text",payload)
        self.assertEqual(payload["a_context_text"],"A独有信息");self.assertEqual(payload["b_context_text"],"B自己的单位")
        self.assertEqual(payload["b_standard_review"],"B待处理.json")
        custom=folder/"自选配置.json";custom.write_text(json.dumps({"gold_path":"用户自选标准.jsonl"}),encoding="utf-8-sig")
        self.assertEqual(load_settings(custom)["gold_path"],"用户自选标准.jsonl")

    def test_无需认证的模型不被界面阻拦(self):
        from 附注更新.界面 import missing_inputs
        folder=Path(tempfile.mkdtemp(prefix="免认证识别材料_"));a=folder/"A.xlsx";gold=folder/"标准.jsonl"
        a.write_bytes(b"test");gold.write_text("{}",encoding="utf-8-sig")
        self.assertEqual(missing_inputs("recognize_a",{"a_path":str(a),"gold_path":str(gold),"output_dir":str(folder),
                          "base_url":"http://127.0.0.1:8000/v1","model":"local-model","api_key":""}),[])

    def test_切换金标准保留两份映射但清掉旧更新计划(self):
        from 附注更新.界面 import file_selection_updates
        old={"gold_path":"旧标准.jsonl","a_mapping":"A映射.json","b_mapping":"B映射.json","a_review":"A核对.xlsx",
             "update_plan":"清单.json","a_prime":"A更新.xlsx","word":"更新.docx"}
        new={**old,**file_selection_updates("gold_path","新标准.jsonl",old)}
        for key in ("a_mapping","b_mapping","a_review"):self.assertEqual(new[key],old[key])
        for key in ("update_plan","a_prime","word"):self.assertEqual(new[key],"")

    def test_两份文件补充信息独立编辑并清理各自旧成果(self):
        import tkinter as tk
        from unittest.mock import patch
        from 附注更新 import 界面
        root=tk.Tk();root.withdraw()
        try:
            with patch.object(界面,"load_settings",return_value={}):app=界面.Application(root)
            app.values["a_mapping"].set("A映射.json");app.values["b_mapping"].set("B映射.json")
            app.values["a_standard_review"].set("A待处理.json")
            app.set_context("a_context_text","A为母公司，人民币元")
            app.show_step("recognize_b");app.set_context("b_context_text","B以万元列示")
            values=app.config()
            self.assertNotIn("context_text",values)
            self.assertEqual(values["a_context_text"],"A为母公司，人民币元");self.assertEqual(values["b_context_text"],"B以万元列示")
            self.assertEqual(values["a_mapping"],"");self.assertEqual(values["a_standard_review"],"")
            self.assertEqual(values["b_mapping"],"")
        finally:_destroy_test_window(root)

    def test_标准预览展示影响语义的维度差异(self):
        from 附注更新.界面 import format_revision_preview
        report={"added":0,"changed":1,"retired":0,"changes":[{"id":"测试槽位","kind":"changed",
                "before":{"slot":{"name":"应收款"},"dimensions":{"required":["period"]}},
                "after":{"slot":{"name":"应收款"},"dimensions":{"required":["period","counterparty"]}}}]}
        text=format_revision_preview(report)
        self.assertIn("结构检查不代表语义正确",text);self.assertIn("counterparty",text)
        self.assertIn("修订 1 项",text)

    def test_标准只在确认发布后切换且复验预览身份(self):
        import tkinter as tk
        import types
        from unittest.mock import patch,Mock
        from 附注更新 import 界面
        report={"added":1,"changed":0,"retired":0,"changes":[],"source_hashes":{"current":"old-hash","candidate":"new-hash"}}
        module=types.ModuleType("附注更新.金标准")
        module.preview_revision=Mock(return_value=report)
        module.publish_revision=Mock(return_value={"path":"新正式版本.jsonl","sha256":"new-hash"})
        root=tk.Tk();root.withdraw()
        try:
            with patch.object(界面,"load_settings",return_value={}):app=界面.Application(root)
            app.values["gold_path"].set("原标准.jsonl");app.values["a_mapping"].set("A映射.json")
            app.values["update_plan"].set("原清单.json");app.show_step("settings")
            with patch.dict(sys.modules,{"附注更新.金标准":module}),patch.object(界面,"save_settings") as save:
                self.assertTrue(app.preview_gold_revision("外部修订.jsonl"))
                self.assertEqual(app.values["gold_path"].get(),"原标准.jsonl");module.publish_revision.assert_not_called();save.assert_not_called()
                app.publish_gold_revision()
                self.assertEqual(module.publish_revision.call_args.kwargs["expected_hashes"],report["source_hashes"])
                self.assertEqual(app.values["gold_path"].get(),"新正式版本.jsonl")
                self.assertEqual(app.values["a_mapping"].get(),"A映射.json");self.assertEqual(app.values["update_plan"].get(),"")
                save.assert_called_once();self.assertIsNone(app.revision_preview)
                app.publish_gold_revision();self.assertEqual(module.publish_revision.call_count,1)
        finally:_destroy_test_window(root)

    def test_修订文件变动导致发布拒绝后要求重新预览(self):
        import tkinter as tk
        import types
        from unittest.mock import patch,Mock
        from 附注更新 import 界面
        report={"added":0,"changed":1,"retired":0,"changes":[],"source_hashes":{"current":"old","candidate":"candidate"}}
        module=types.ModuleType("附注更新.金标准");module.publish_revision=Mock(side_effect=ValueError("文件在预览后变化"))
        root=tk.Tk();root.withdraw()
        try:
            with patch.object(界面,"load_settings",return_value={}):app=界面.Application(root)
            app.values["gold_path"].set("旧标准.jsonl");app.revision_current="旧标准.jsonl";app.revision_candidate.set("修订.jsonl");app.revision_preview=report
            app.show_step("settings")
            with patch.dict(sys.modules,{"附注更新.金标准":module}),patch.object(界面.messagebox,"showerror") as error:
                app.publish_gold_revision();error.assert_called_once()
            self.assertEqual(app.values["gold_path"].get(),"旧标准.jsonl");self.assertIsNone(app.revision_preview)
            self.assertEqual(str(app.publish_button.cget("state")),"disabled")
        finally:_destroy_test_window(root)

    def test_补充信息默认折叠已有文字自动展开且日志紧凑(self):
        import tkinter as tk
        from unittest.mock import patch
        from 附注更新 import 界面
        root=tk.Tk();root.withdraw()
        try:
            with patch.object(界面,"load_settings",return_value={}):app=界面.Application(root)
            self.assertLessEqual(int(app.log.cget("height")),3)
            app.show_step("recognize_b")
            self.assertEqual(app.context_panels["b_context_text"].winfo_manager(),"")
            app.toggle_context("b_context_text")
            self.assertEqual(app.context_panels["b_context_text"].winfo_manager(),"grid")
            app.values["a_context_text"].set("人民币元")
            app.show_step("recognize_a")
            self.assertEqual(app.context_panels["a_context_text"].winfo_manager(),"grid")
        finally:_destroy_test_window(root)

    def test_成果生成后定位到成果且子控件可滚动(self):
        import tkinter as tk
        from types import SimpleNamespace
        from unittest.mock import patch
        from 附注更新 import 界面
        root=tk.Tk();root.withdraw()
        try:
            with patch.object(界面,"load_settings",return_value={}):app=界面.Application(root)
            app.show_step("recognize_b")
            self.assertTrue(hasattr(app,"scroll_to_outputs"))
            with patch.object(app,"scroll_to_outputs") as focus:
                app.apply_result("recognize_b",{"status":"complete","files":{"b_mapping":"B映射.json","b_review":"B核对.xlsx"}})
                root.update_idletasks();focus.assert_called_once()
            with patch.object(app.canvas,"yview_scroll") as scroll:
                self.assertEqual(app.on_content_wheel(SimpleNamespace(widget=app.path_entries["b_mapping"],delta=-120)),"break")
                scroll.assert_called_once_with(1,"units")
                scroll.reset_mock()
                self.assertIsNone(app.on_content_wheel(SimpleNamespace(widget=app.context_text_widgets["b_context_text"],delta=-120)))
                scroll.assert_not_called()
        finally:_destroy_test_window(root)


    def test_来源筛选默认关闭且只接受布尔值并持久化(self):
        self.assertIs(validate_settings({}, require_model=False).get("scope_prefilter"), False)
        folder = Path(tempfile.mkdtemp(prefix="来源筛选设置_"))
        self.assertIs(load_settings(folder/"不存在.json").get("scope_prefilter"), False)
        old = folder/"旧配置.json";old.write_text("{}", encoding="utf-8-sig")
        self.assertIs(load_settings(old).get("scope_prefilter"), False)
        for value in (False, True):
            path = folder/(str(value)+".json")
            save_settings({"scope_prefilter":value}, path)
            self.assertIs(load_settings(path)["scope_prefilter"], value)
        for index, value in enumerate(("false", "true", 0, 1, None, {}, [])):
            with self.subTest(value=value):
                with self.assertRaisesRegex(ValueError, "来源筛选"):
                    validate_settings({"scope_prefilter":value}, require_model=False)
                path = folder/(str(index)+".json")
                path.write_text(json.dumps({"scope_prefilter":value}), encoding="utf-8-sig")
                with self.assertRaisesRegex(ValueError, "来源筛选"):load_settings(path)

    def test_B页筛选默认关闭切换保留A并清理B及下游选择(self):
        import tkinter as tk
        from unittest.mock import patch
        from 附注更新 import 界面
        legacy=Path(tempfile.mkdtemp(prefix="旧版范围界面_"))/"旧版标准.jsonl"
        legacy.write_text(json.dumps({'id':'旧版定义'}),encoding='utf-8-sig')
        root = tk.Tk();root.withdraw()
        try:
            with patch.object(界面, "load_settings", return_value={'gold_path':str(legacy)}):app = 界面.Application(root)
            app.show_step("recognize_b")
            self.assertIs(app.config().get("scope_prefilter"), False)
            checks = [w for w, _ in app.page_editable if isinstance(w, tk.Checkbutton)
                      and w.cget("text") == "先按附注内容筛选来源（可选）"]
            self.assertEqual(len(checks), 1)
            for key in ("a_link", "a_mapping", "b_path", "b_mapping", "b_scope", "b_review", "update_plan", "word"):
                app.values[key].set(key+".测试")
            checks[0].invoke()
            self.assertIs(app.config()["scope_prefilter"], True)
            for key in ("a_link", "a_mapping", "b_path"):self.assertEqual(app.values[key].get(), key+".测试")
            for key in ("b_mapping", "b_scope", "b_review", "update_plan", "word"):self.assertEqual(app.values[key].get(), "")
            app.show_step("recognize_a");app.show_step("recognize_b")
            self.assertIs(app.config()["scope_prefilter"], True)
            with patch.object(界面, "save_settings") as save:
                app.save();self.assertIs(save.call_args.args[0]["scope_prefilter"], True)
            checks = [w for w, _ in app.page_editable if isinstance(w, tk.Checkbutton)
                      and w.cget("text") == "先按附注内容筛选来源（可选）"]
            app.values["b_mapping"].set("新范围映射.json");checks[0].invoke()
            self.assertIs(app.config()["scope_prefilter"], False)
            self.assertEqual(app.values["b_mapping"].get(), "")
            self.assertEqual(app.values["a_link"].get(), "a_link.测试")
            path = Path(tempfile.mkdtemp(prefix="界面来源筛选保存_"))/"独立测试配置.json"
            with patch.object(界面, "save_settings", side_effect=lambda config:save_settings(config, path)):
                app.save()
            saved = load_settings(path)
            self.assertIs(saved["scope_prefilter"], False)
        finally:_destroy_test_window(root)
        root = tk.Tk();root.withdraw()
        try:
            with patch.object(界面, "load_settings", return_value=saved):app = 界面.Application(root)
            app.show_step("recognize_b")
            self.assertIs(app.config()["scope_prefilter"], False)
            self.assertEqual(app.values["a_link"].get(), "a_link.测试")
            checks = [w for w, _ in app.page_editable if isinstance(w, tk.Checkbutton)
                      and w.cget("text") == "先按附注内容筛选来源（可选）"]
            self.assertEqual(len(checks), 1)
            self.assertEqual(root.getvar(checks[0].cget("variable")), 0)
            clear = [w for w, _ in app.page_editable if isinstance(w, tk.Button)
                     and w.cget("text") == "清空" and w.master == app.path_entries["b_mapping"].master]
            self.assertEqual(len(clear), 1)
            app.values["b_mapping"].set("旧范围映射.json");clear[0].invoke()
            self.assertEqual(app.values["b_mapping"].get(), "")
            self.assertEqual(app.values["a_link"].get(), "a_link.测试")
        finally:_destroy_test_window(root)

    def test_新版更新数据页只呈现已接通的独立识别流程(self):
        import tkinter as tk
        from unittest.mock import patch
        from 附注更新 import 界面
        from 附注更新.设置 import DEFAULT_GOLD_PATH
        root=tk.Tk();root.withdraw()
        try:
            with patch.object(界面,'load_settings',return_value={'gold_path':DEFAULT_GOLD_PATH,
                    'scope_prefilter':True,'b_mapping':'已保存映射.json'}):
                app=界面.Application(root)
            app.show_step('recognize_b')
            texts=[w.cget('text') for w,_ in app.page_editable if 'text' in w.keys()]
            for unsupported in ('先按附注内容筛选来源（可选）','导出范围复核格式','导入范围复核意见'):
                self.assertNotIn(unsupported,texts)
            self.assertFalse(app.config()['scope_prefilter'])
            self.assertNotIn('b_scope',app.path_entries)
            self.assertEqual(app.values['b_mapping'].get(),'已保存映射.json')
            self.assertIsNone(app.worker)
        finally:_destroy_test_window(root)

    def test_损坏配置明确报错(self):
        folder = Path(tempfile.mkdtemp(prefix="附注更新损坏设置测试_"))
        path = folder / "用户配置.json"
        path.write_text("不是 JSON", encoding="utf-8-sig")
        with self.assertRaisesRegex(ValueError, "配置"):
            load_settings(path)


if __name__ == "__main__":
    unittest.main()
