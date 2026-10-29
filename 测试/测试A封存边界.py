# coding: utf-8
"""第二步的离线反例；响应均为测试替身，不构成财务识别结论。"""
import copy
import unittest
import json
from pathlib import Path
from unittest.mock import patch
from 测试.测试统一语义 import records
from 附注更新.统一语义 import validate_records, required_cell_dimensions
from 附注更新.A语义阅读 import expand_scopes, semantic_questions, source_material, normalize_evidence


class A封存边界测试(unittest.TestCase):
    def test_新A完整性不扣报表口径(self):
        standard=validate_records(records());standard['_complete_identity']=True
        self.assertIn('report_scope',required_cell_dimensions(standard['metric']['固定资产原值'],standard))

    def test_程序不补集合domain和完整性(self):
        standard=validate_records(records());standard['_complete_identity']=True
        row={'metrics':[{'id':'固定资产原值','ranges':['B2']}],
             'dimensions':[{'id':'asset_selection','ranges':['B2'],'value':{'mode':'all'},'evidence':['A2']}]}
        result=expand_scopes(row,{'range':'A1:B2','data_ranges':['B2']},{'B2':{'value':1}},standard)
        self.assertEqual(result['B2']['dimensions']['asset_selection']['value'],{'mode':'all'})

    def test_已传跨表和同文引用不按行距裁判(self):
        cells={a:{'row':r,'column':c,'value':v} for a,r,c,v in [('A1',1,1,'同文'),('A2',2,1,'同文'),('B2',2,2,1)]}
        snapshot={'path':'仅内存.xlsx','semantic_identity_version':2,'sheets':[
            {'name':'甲','cells':cells,'merges':[]},{'name':'乙','cells':{'A1':{'row':1,'column':1,'value':'跨表原文'}},'merges':[]}]}
        _,aliases=source_material(snapshot,[],[{'sheet':'甲','range':'A1:B2'},{'sheet':'乙','range':'A1'}])
        self.assertEqual(normalize_evidence(['乙!A1'],'甲',aliases,'B2'),['乙!A1'])
        self.assertEqual(set(normalize_evidence(['同文'],'甲',aliases,'B2')),{'甲!A1','甲!A2'})
        with self.assertRaises(ValueError):normalize_evidence(['乙!B99'],'甲',aliases,'B2')

    def test_新A不按财务词和标题生成纠正(self):
        # 即使旧路径会索取表数据，新路径也不以这些字段作业务判断。
        self.assertEqual(semantic_questions({'tables':[{'id':'t'}]},[],{}, {'_complete_identity':True}),[])

    def test_新A共同说明不按物理范围裁掉(self):
        snapshot={'path':'仅内存.xlsx','semantic_identity_version':2,'sheets':[
            {'name':'甲','cells':{'B2':{'row':2,'column':2,'value':1}},'merges':[]}]}
        context=[{'sheet':'乙','first_row':1,'first_column':1,'row_count':1,'column_count':1,'text':'另一表的完整共同说明'}]
        material,_=source_material(snapshot,context,[{'sheet':'甲','range':'B2'}])
        self.assertIn('另一表的完整共同说明',json.dumps(material,ensure_ascii=False))
        text='这是两处共用的完整原文，不按重复次数删掉，也不因物理位置不同替模型决定含义。'*2
        context=[{'title':text,'text':text}, {'sheet':'乙','first_row':1,'first_column':1,'row_count':1,'column_count':1,'text':text}]
        material,aliases=source_material(snapshot,context,[{'sheet':'甲','range':'B2'}])
        self.assertEqual(material['共同原文']['T0'],text)
        self.assertEqual(len(normalize_evidence(['@T0'],'甲',aliases,'B2')),2)

    def test_跨表共同范围能通过新阅读答复约束(self):
        from 附注更新.回执格式 import response_format
        import jsonschema
        payload={'task':'a_semantic_read','identity_collection_version':1,
                 'reading_ranges':[{'sheet':'甲','range':'A1:B2'}], 'candidate_locations':['甲!B2'],
                 'source':{'sheets':[{'sheet':'甲'},{'sheet':'乙'}]}}
        schema=response_format(payload)['json_schema']['schema']
        answer={'tables':[{'id':'t','sheet':'甲','range':'A1:B2','subject':'仅测试','meaning':'仅测试',
                          'data_ranges':['B2'],'context_ranges':[],
                          'context_references':[{'sheet':'乙','range':'A1'}]}],'notes':''}
        jsonschema.validate(answer,schema)

    def test_关系证据例外和未决范围各自检查(self):
        from 附注更新.A合集 import validate_relations,_relation_blocked
        table={'id':'t','sheet':'表','data_ranges':['B2:B3']}
        relation={'content':'原有关系','record_ranges':[{'table_id':'t','ranges':['B2']}],'exceptions':[],'evidence':['表!A1']}
        binding={'tables':[{'id':'t','relations':[relation]}]}
        review={'tables':[{'id':'t','checked_relations':[0],'corrections':{},'issues':[]}]}
        aliases={'_all_evidence':{'表!A1':{'value':'实际原文'}}}
        validate_relations(binding,review,[table],aliases)
        relation['evidence']=[]
        with self.assertRaisesRegex(ValueError,'证据'):validate_relations(binding,review,[table],aliases)
        relation['evidence']=['表!A1'];relation['exceptions']=[{'table_id':'t','ranges':['B3'],'content':''}]
        with self.assertRaisesRegex(ValueError,'例外'):validate_relations(binding,review,[table],aliases)
        relation['exceptions']=[]
        review['tables'][0]['issues']=[{'range':'B3'}]
        self.assertFalse(_relation_blocked(relation,review,[table]))
        review['tables'][0]['issues']=[{'range':'B2'}]
        self.assertTrue(_relation_blocked(relation,review,[table]))

    def test_同一失败输入不按时间取起点也不重答(self):
        from collections import Counter
        from 附注更新.语义 import SemanticEngine
        from 测试.测试Codex语义调用 import test_directory
        engine=object.__new__(SemanticEngine)
        engine.cancel=None;engine.settings={'base_url':'codex://local','model':'test'}
        engine.source_hash='a'*64;engine.cache_source_hash='';engine.gold_hash='b'*64
        engine.usage=Counter();engine._review_receipts=[];engine.log=lambda _:None
        def reject(result):raise ValueError('原答复缺少字段')
        with test_directory() as directory,patch.object(engine,'_http',return_value={'原答复':'保留'}) as http:
            engine.work_dir=directory
            payload={'task':'a_semantic_bind','identity_collection_version':1}
            with self.assertRaisesRegex(ValueError,'原答复缺少'):engine._request('测试替身',payload,reject)
            original={p:p.read_bytes() for p in directory.rglob('*.json')}
            with self.assertRaisesRegex(ValueError,'不自动重答'):engine._request('测试替身',payload,reject)
            self.assertEqual(http.call_count,1)
            self.assertEqual(original,{p:p.read_bytes() for p in directory.rglob('*.json')})

    def test_新合集读取不依赖原表和旧重建(self):
        from 附注更新.A合集 import validate_collection
        with self.assertRaisesRegex(ValueError,'封存'):
            validate_collection({'facts':[]})

    def test_普通入口保存原回答明确修正缺项及离线回读(self):
        from openpyxl import Workbook
        from 测试.测试Codex语义调用 import test_directory
        from 测试.测试统一语义 import fact
        from 附注更新.统一识别 import UnifiedSemanticEngine
        from 附注更新.统一勾稽 import load_fact_mapping
        from 附注更新.A合集 import validate_collection,FIELDS
        from 附注更新.统一语义 import definition_hash
        from 附注更新.表格 import file_hash
        with test_directory() as directory:
            source=directory/'机械反例.xlsx';gold=directory/'合成标准.jsonl'
            gold.write_text('\n'.join(json.dumps(r,ensure_ascii=False) for r in records()),encoding='utf-8')
            book=Workbook();sheet=book.active;sheet.title='表'
            for ref,value in {'A1':'共同说明','B1':'金额','A2':'同文','B2':1,'A3':'同文','B3':2,'D1':'跨区域原文','E1':'金额','D2':'项目','E2':3}.items():sheet[ref]=value
            book.save(source)
            tables=[{'id':'甲','sheet':'表','range':'A1:B3','subject':'固定资产','meaning':'测试共同范围及例外','data_ranges':['B2:B3']},
                    {'id':'乙','sheet':'表','range':'D1:E2','subject':'固定资产','meaning':'测试同科目独立子表','data_ranges':['E2']}]
            relation={'content':'本例原答复明确的共同标签关系','record_ranges':[{'table_id':'甲','ranges':['B2:B3']}],
                      'exceptions':[],'evidence':['表!A1']}
            calls=[]
            def fake(system,payload):
                calls.append(payload['task'])
                if payload['task']=='a_semantic_read':return {'tables':tables,'notes':'测试替身，非真实语义识别'}
                if payload['task']=='a_semantic_select':return {'tables':[{'id':t['id'],'metric_ids':['固定资产原值'],'missing':''} for t in tables]}
                if payload['task']=='a_semantic_bind':
                    self.assertNotIn('不是本次单格门槛',system)
                    rows=[]
                    for t in tables:
                        dims=copy.deepcopy(fact()['dimensions']);dims.pop('report_scope');dims['unit_scale']=1
                        rows.append({'id':t['id'],'metrics':[{'id':'固定资产原值','ranges':t['data_ranges']}],
                                     'dimensions':[{'id':k,'ranges':['*'],'value':v,'evidence':['表!D1']} for k,v in dims.items()],
                                     'unresolved':[],'relations':[relation] if t['id']=='甲' else []})
                    return {'tables':rows}
                self.assertEqual(payload['task'],'a_semantic_review')
                self.assertNotIn('semantic_questions',payload)
                return {'tables':[{'id':'甲','checked_relations':[0],'corrections':{'patches':[
                    {'id':'entity','ranges':['B3'],'value':'明确修正主体','evidence':['表!A3']}],'relations':[]},'issues':[]},
                    {'id':'乙','checked_relations':[],'corrections':{},'issues':[]}]}
            engine=UnifiedSemanticEngine({'base_url':'codex://local','model':'test'},gold,directory/'运行')
            with patch.object(engine,'_http',side_effect=fake):
                result=engine.recognize_workbook(source,context=[],artifact_metadata={'carrier':'A','original_path':str(source),'original_hash':file_hash(source)})
            self.assertEqual(calls,['a_semantic_read','a_semantic_select','a_semantic_bind','a_semantic_review'])
            self.assertEqual(len(result['slot_bindings']),3)
            self.assertFalse(result['来源缺口'])
            self.assertEqual(result['完整性']['表!B3']['missing_dimensions'],['report_scope'])
            self.assertEqual(result['slot_bindings'][1]['partial_semantics']['dimensions']['entity'],'明确修正主体')
            self.assertEqual(len(result['已确认关系']),1)
            actual=Path.read_bytes;mapping_path=Path(result['mapping_path']).resolve()
            def only_collection(path):
                self.assertEqual(path.resolve(),mapping_path,'回读偷读外部材料')
                return actual(path)
            with patch.object(Path,'read_bytes',only_collection),patch('附注更新.A语义阅读.verify_mapping',side_effect=AssertionError('旧重建被调用')):
                loaded=load_fact_mapping(mapping_path,{})
            self.assertEqual(loaded['slot_bindings'],result['slot_bindings'])
            bad=copy.deepcopy(result);bad['slot_bindings'][0]['partial_semantics']['dimensions']['entity']='程序补造'
            bad['封存校验']=definition_hash({k:bad.get(k) for k in FIELDS})
            with self.assertRaisesRegex(ValueError,'原回答'):validate_collection(bad)
            bad=copy.deepcopy(result);bad['已确认关系'][0]['原有关系内容']['content']='无来源新关系'
            bad['封存校验']=definition_hash({k:bad.get(k) for k in FIELDS})
            with self.assertRaisesRegex(ValueError,'关系与原判断'):validate_collection(bad)
            bad=copy.deepcopy(result);bad['原文证据']['cells'].pop('表!D1')
            bad['封存校验']=definition_hash({k:bad.get(k) for k in FIELDS})
            with self.assertRaisesRegex(ValueError,'缺少原文证据'):validate_collection(bad)
            from 附注更新.分步流程 import run_step
            selections=[]
            def ordinary(system,payload):
                if payload['task']=='a_semantic_select':
                    selections.append(payload)
                    if len(selections)==1:return {'tables':[{'id':t['id'],'metric_ids':[],'missing':'需要扩大检索'} for t in tables]}
                return fake(system,payload)
            config={'a_path':str(source),'gold_path':str(gold),'output_dir':str(directory/'正式入口'),
                    'base_url':'codex://local','model':'gpt-6-luna','freeze_standard':True}
            with patch.object(UnifiedSemanticEngine,'_http',side_effect=ordinary):report=run_step('recognize_a',config)
            self.assertNotEqual(report['status'],'failed',report)
            self.assertEqual(len(selections),2)
            self.assertIn('definition',selections[1]['catalog'][0])
            formal=load_fact_mapping(report['files']['a_mapping'],{})
            self.assertEqual(formal['slot_bindings'],result['slot_bindings'])
            calls.clear()
            def invalid(system,payload):
                if payload['task']=='a_semantic_bind':calls.append(payload['task']);return {'不能解析的原回答':'仍应封存'}
                return fake(system,payload)
            config['output_dir']=str(directory/'解析失败')
            with patch.object(UnifiedSemanticEngine,'_http',side_effect=invalid):stopped=run_step('recognize_a',config)
            self.assertEqual(calls,['a_semantic_read','a_semantic_select','a_semantic_bind'])
            failed=load_fact_mapping(stopped['files']['a_mapping'],{})
            self.assertEqual(len(failed['未通过答复']),1)
            self.assertIn('仍应封存',failed['未通过答复'][0]['text'])
            self.assertTrue(failed['unresolved'])


if __name__=='__main__':unittest.main()
