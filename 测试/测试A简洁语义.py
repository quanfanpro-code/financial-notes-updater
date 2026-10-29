# coding: utf-8
"""只检验共享语义展开：行列交叉、空白目标、局部例外和越界拒绝。"""
import importlib.util
import unittest
import json
from pathlib import Path
import time
from unittest.mock import patch


class 简洁语义检查(unittest.TestCase):
    def test_局部核查保留范围内疑点及无定位问题(self):
        from copy import deepcopy
        from 附注更新.A语义阅读 import focused_review_payload
        tables=[{'id':key,'range':'A1:B4','data_ranges':['B2:B4']} for key in ('甲','乙')]
        payload={'reading':tables,'source':{'原表':'全部保留'},
                 'recognized':{'tables':[{'id':key,'metrics':[],'dimensions':[]} for key in ('甲','乙')]},
                 'program_issues':[{'id':'甲','reason':'返回维度无法解析'},
                                   {'id':'甲','cells':['B2','B3'],'reason':'缺原文依据'},
                                   {'id':'乙','reason':'其他表问题'}],
                 'semantic_questions':[{'id':'甲','ranges':['B2:B3'],'reason':'两行对象被合并'},
                                       {'id':'甲','ranges':['B4'],'reason':'范围外问题'}],
                 'previous_issues':[{'id':'甲','range':'B2:B3','reason':'旧含义未清','evidence':['A2']}]}
        before=deepcopy(payload)
        result=focused_review_payload(payload,{'甲':['B3'],'乙':[]},{'metric':{},'dimension':{}})
        self.assertEqual(result.get('semantic_questions'),[{'id':'甲','ranges':['B3'],'reason':'两行对象被合并'}])
        self.assertEqual(result.get('previous_issues'),[{'id':'甲','range':'B3','reason':'旧含义未清','evidence':['A2']}])
        self.assertEqual(result['program_issues'],[{'id':'甲','reason':'返回维度无法解析'},
                                                {'id':'甲','cells':['B3'],'reason':'缺原文依据'}])
        self.assertEqual(result['source'],payload['source'])
        self.assertEqual(payload,before)

    def test_旧全表核查不能被新局部核查缩小(self):
        from copy import deepcopy
        from 附注更新.A语义阅读 import carry_role_corrections,review_for_table
        table={'id':'甲','sheet':'表','range':'A1:B3','data_ranges':['B2:B3']}
        packet={'source_path':'原件.xlsx','sheet':'表','candidate_cells':['B2','B3'],
                'cells':{'A2':{'value':'甲单位'},'A3':{'value':'乙单位'}}}
        problem={'range':'B2','reason':'单位尚不明确','evidence':['A2']}
        old={'tables':[{'id':'甲','issues':[problem],'corrections':{}}]}
        new={'tables':[{'id':'甲','checked_ranges':['B3'],'issues':[],'corrections':{}}]}
        before=deepcopy((old,new))
        result=carry_role_corrections(new,[old])
        self.assertEqual(set(result['tables'][0].get('checked_ranges',['B2','B3'])),{'B2','B3'})
        checked=review_for_table(result['tables'][0],table,packet,{})
        self.assertEqual(checked['checked_cells'],['B2','B3'])
        self.assertEqual(checked['issues'][0]['cells'],['B2'])
        restored=carry_role_corrections({'tables':[]},[old])
        self.assertEqual(review_for_table(restored['tables'][0],table,packet,{})['checked_cells'],['B2','B3'])
        partial=deepcopy(old);partial['tables'][0]['checked_ranges']=['B2']
        self.assertEqual(carry_role_corrections({'tables':[]},[partial])['tables'][0]['checked_ranges'],['B2'])
        unconfirmed=carry_role_corrections(new,[{'tables':[{'id':'甲'}]}])
        self.assertEqual(unconfirmed['tables'][0]['checked_ranges'],['B3'])
        self.assertEqual((old,new),before)

    def test_已保存A原文身份在界面显示完成而非标准缺项(self):
        from types import SimpleNamespace
        from 附注更新.界面 import Application
        work=Path(r'C:\Users\27651\Documents\Codex\2026-09-29\c-users-27651-documents-codex-2026-2\work')/('A状态检查_'+str(time.time_ns()))
        work.mkdir();path=work/'A成果.json'
        record={'schema_version':2,'carrier':'A','source_path':'原件.xlsx','source_hash':'0'*64,
                'gold_path':'金标准.jsonl','gold_hash':'1'*64,'facts':[],'owners':{},'excluded':[],
                'unresolved':[{'sheet':'表','cell':'B2'},{'sheet':'表','cell':'B3'}],'complete':False,
                'source_identities':[{'source_reference':{'location':'表!B2'}},{'source_reference':{'location':'表!B3'}}],
                'a_identity_coverage':{'complete':True,'financial_cells':2,'unresolved_cells':0,'unresolved':[]}}
        with path.open('x',encoding='utf-8-sig') as handle:json.dump(record,handle,ensure_ascii=False)
        shown={};label=SimpleNamespace(winfo_exists=lambda:True,configure=lambda **kw:shown.update(kw))
        app=SimpleNamespace(mapping_summary_labels={'a_mapping':label},values={'a_mapping':SimpleNamespace(get=lambda:str(path))})
        Application.refresh_mapping_summary(app,'a_mapping')
        self.assertIn('A识别完成',shown['text'])
        self.assertIn('已识别 2 格',shown['text'])
        self.assertIn('语义待核实 0 格',shown['text'])
        self.assertIn('标准对应有差别 2 格',shown['text'])

    def test_新A拒绝原文身份代替标准绑定并保留旧协议审计(self):
        from copy import deepcopy
        from openpyxl import Workbook,load_workbook
        from 附注更新.统一识别 import UnifiedSemanticEngine
        from 附注更新.统一勾稽 import load_fact_mapping
        from 附注更新.分步流程 import _recognize_v2
        from 附注更新.统一语义 import read_standard_v2
        from 附注更新.表格 import file_hash
        base=Path(r'C:\Users\27651\Documents\Codex\2026-09-29\c-users-27651-documents-codex-2026-2')
        work=base/'work'/('A原文优先检查_'+str(time.time_ns()));work.mkdir(parents=True)
        source=work/'原文未提供公司年度.xlsx'
        book=Workbook();sheet=book.active;sheet.title='表'
        sheet.append(['项目','期末']);sheet.append(['预付款余额',None]);sheet.append(['一年内到期可转债本息','=0'])
        book.save(source);book.close()
        gold=Path(r'C:\Users\27651\Desktop\附注自动更新系统\金标准\版本\财务报表附注语义金标准_20260922_132731_236607_5547978f.jsonl')
        before=(file_hash(source),file_hash(gold));standard=read_standard_v2(gold)
        table={'id':'t','sheet':'表','range':'A1:B3','subject':'披露项目',
               'meaning':'各项目期末数据，公司和具体年度未提供','data_ranges':['B2:B3']}
        declarations=[{'ranges':['B2'],'meaning':'期末预付款余额，公司和具体年度未提供',
                       'evidence':['A2','B1'],'closest_metric_id':'metric.prepayment.gross_balance',
                       'standard_difference':'原文没有公司、具体日期等信息，保持未提供'},
                      {'ranges':['B3'],'meaning':'期末一年内到期的可转换公司债券本金与利息合计',
                       'evidence':['A3','B1'],'closest_metric_id':None,
                       'standard_difference':'未找到完全对应本息合计的指标'}]
        calls=[]
        def fake(engine,system,payload):
            calls.append(payload['task'])
            if payload['task']=='a_semantic_read':return {'tables':[table],'notes':'未提供公司及具体年度，不能补写'}
            if payload['task']=='a_semantic_select':return {'tables':[{'id':'t','metric_ids':['metric.prepayment.gross_balance'],'missing':'可转债本息无完全对应指标'}]}
            if payload['task']=='a_semantic_bind':
                return {'tables':[{'id':'t','metrics':[{'ranges':['B2'],'id':'metric.prepayment.gross_balance'}],
                                  'dimensions':[],'source_identities':deepcopy(declarations),'unresolved':[]}]}
            self.assertEqual(payload['task'],'a_semantic_review')
            self.assertFalse(payload.get('program_issues'),'不能因标准字段缺失要求补查')
            return {'tables':[{'id':'t','corrections':{},'issues':[]}]}
        settings={'base_url':'https://example.invalid','model':'mimo-v2.6-flash','thinking_mode':'default'}
        progress=[]
        with patch('附注更新.分步流程.validate_settings',return_value=settings),patch.object(UnifiedSemanticEngine,'_http',fake), \
             patch('附注更新.统一勾稽.calculate_mapping_formulas',side_effect=AssertionError('A识别不计算金额')):
            report=_recognize_v2({'a_path':str(source),'gold_path':str(gold),'output_dir':str(work)},'A',work,lambda _:None,None,progress.append)
        self.assertEqual(calls,['a_semantic_read','a_semantic_select','a_semantic_select','a_semantic_bind'])
        self.assertNotIn('a_standard_review',report['files'])
        loaded=load_fact_mapping(report['files']['a_mapping'],standard)
        self.assertFalse(loaded['a_identity_coverage']['complete'])
        self.assertFalse(loaded['complete'],'接近标准不能冒充完全一致的标准身份')
        self.assertEqual(loaded['source_identities'],[])
        self.assertTrue({'B2','B3'}<={row['cell'] for row in loaded['unresolved']})
        original=json.loads(loaded['未通过答复'][0]['text'].lstrip('\ufeff'))['result']
        self.assertEqual(original['tables'][0]['source_identities'],declarations)
        self.assertEqual((file_hash(source),file_hash(gold)),before)
        tampered=deepcopy(loaded);tampered.pop('_validated_standard_hash',None)
        tampered['source_identities']=[{'meaning':'编造了公司和年度'}]
        bad=work/'篡改副本.json';bad.write_text(json.dumps(tampered,ensure_ascii=False),encoding='utf-8-sig')
        with self.assertRaises(ValueError):load_fact_mapping(bad,standard)
        # 同一最小表检查真正的漏读、错误依据和矛盾仍不能冒充完成。
        from 附注更新.A语义阅读 import assemble,source_identity_records,source_material,table_packet,reviewed_binding,focused_review_payload,standard_material
        from 附注更新.表格 import read_workbook,materialize_candidate_blanks
        snapshot=read_workbook(source);materialize_candidate_blanks(snapshot)
        packet=table_packet(table,snapshot,[]);_,aliases=source_material(snapshot,[],[table])
        binding={'tables':[{'id':'t','metrics':[],'dimensions':[],'source_identities':deepcopy(declarations)}]}
        no_difference=deepcopy(binding['tables'][0]);no_difference['source_identities'][0].pop('standard_difference')
        self.assertEqual(len(source_identity_records(no_difference,table,packet,standard,aliases)),2)
        checked={'tables':[{'id':'t','issues':[{'range':'B3','reason':'本息范围尚未读懂','evidence':['A3']}]}]}
        partial=assemble(snapshot,[],standard,{'tables':[table]},[(binding,checked)])
        self.assertFalse(partial['a_identity_coverage']['complete'])
        self.assertEqual(partial['a_identity_coverage']['unresolved_cells'],1)
        checked['tables'][0]['issues']=[]
        missing=deepcopy(binding);missing['tables'][0]['source_identities'].pop()
        self.assertFalse(assemble(snapshot,[],standard,{'tables':[table]},[(missing,checked)])['a_identity_coverage']['complete'])
        for key,value in [('evidence',['Z99']),('ranges',['B4']),('meaning','')]:
            invalid=deepcopy(binding['tables'][0]);invalid['source_identities'][0][key]=value
            with self.assertRaises(ValueError):source_identity_records(invalid,table,packet,standard,aliases)
        correction={'tables':[{'id':'t','corrections':{'source_identities':[declarations[1]]}}]}
        restored=reviewed_binding(missing,correction,[table],snapshot)
        self.assertTrue(assemble(snapshot,[],standard,{'tables':[table]},[(restored,checked)])['a_identity_coverage']['complete'])
        focused=focused_review_payload({'reading':[table],'recognized':restored,'standards':standard_material([],standard)}, {'t':['B2']},standard)
        self.assertEqual(len(focused['recognized']['tables'][0]['source_identities']),1)
        self.assertEqual(focused['standards']['metrics'][0]['id'],'metric.prepayment.gross_balance')

    def test_一格引用错误不吞掉其他完整身份(self):
        from 附注更新.A语义阅读 import source_material,table_packet,checked_records
        from 附注更新.统一语义 import read_standard_v2
        standard=read_standard_v2(r'C:\Users\27651\Desktop\附注自动更新系统\金标准\版本\财务报表附注语义金标准_20260922_132731_236607_5547978f.jsonl')
        table={'id':'t','sheet':'表','range':'A1:B3','data_ranges':['B2:B3'],'meaning':'预付款按对象列示'}
        cells={a:{'row':r,'column':c,'value':v} for a,r,c,v in [('A1',1,1,'预付款'),('A2',2,1,'甲供应商'),('A3',3,1,'乙供应商'),('B2',2,2,None),('B3',3,2,None)]}
        snapshot={'path':'原件.xlsx','sha256':'0'*64,'sheets':[{'name':'表','cells':cells,'merges':[]}]}
        context=['甲公司单体报表，人民币元，2025年12月31日。']
        _,aliases=source_material(snapshot,context,[table])
        row={'metrics':[{'id':'metric.prepayment.gross_balance','ranges':['B2:B3']}],
             'dimensions':[{'id':key,'ranges':['*'],'value':value,'evidence':['@C0']} for key,value in
                [('entity','甲公司'),('report_scope','standalone'),('currency','CNY'),('unit_scale',1),('period',{'kind':'instant','date':'2025-12-31'})]]}
        row['dimensions'] += [{'id':'prepayment_selection','ranges':[a],'value':{'mode':'members','members':[name]},'evidence':[ref]}
            for a,name,ref in [('B2','甲供应商','不存在的原文'),('B3','乙供应商','A3')]]
        result=checked_records(row,table,table_packet(table,snapshot,context),standard,aliases)
        self.assertEqual(result['B2']['kind'],'unresolved')
        self.assertIn('不存在的原文',result['B2']['reason'])
        self.assertEqual(result['B3']['kind'],'metric_fact')
        self.assertEqual(result['B3']['fact']['dimensions']['prepayment_selection']['members'],['乙供应商'])
        from 附注更新.A语义阅读 import focused_review_payload,standard_material,validate_review,review_for_table,carry_role_corrections
        payload={'task':'a_semantic_review','reading':[table],'source':{'完整原表':'保留'},'recognized':{'tables':[dict(row,id='t')]},
                 'standards':standard_material(['metric.prepayment.gross_balance','metric.net_profit.amount'],standard)}
        focused=focused_review_payload(payload,{'t':['B3']},standard)
        self.assertEqual(focused['source'],payload['source'])
        self.assertEqual([m['id'] for m in focused['standards']['metrics']],['metric.prepayment.gross_balance'])
        payload['review_candidates']=['metric.prepayment.carrying_amount']
        extra=focused_review_payload(payload,{'t':['B3']},standard)
        self.assertEqual({m['id'] for m in extra['standards']['metrics']},{'metric.prepayment.gross_balance','metric.prepayment.carrying_amount'})
        payload.pop('review_candidates')
        self.assertTrue(all(d['ranges']==['B3'] for d in focused['recognized']['tables'][0]['dimensions']))
        answer={'tables':[{'id':'t','corrections':{},'issues':[]}]}
        reviewed=validate_review(answer,[table],snapshot,context,focused['review_focus'])
        self.assertNotIn('checked_ranges',answer['tables'][0])
        checked=review_for_table(reviewed['tables'][0],table,table_packet(table,snapshot,context),aliases)
        self.assertEqual(checked['checked_cells'],['B3'])
        from 附注更新.统一识别 import apply_identity_review
        adopted=apply_identity_review(result,checked,standard)
        self.assertEqual([f['id'] for f in adopted['facts']],['表!B3'])
        self.assertIn('B2',[u['cell'] for u in adopted['unresolved']])
        earlier={'tables':[{'id':'t','checked_ranges':['B2'],'corrections':{},'issues':[]}]}
        self.assertEqual(set(carry_role_corrections(reviewed,[earlier])['tables'][0]['checked_ranges']),{'B2','B3'})
        restored=carry_role_corrections({'tables':[]},[earlier,reviewed])
        self.assertEqual(set(restored['tables'][0]['checked_ranges']),{'B2','B3'})
        answer['tables'][0]['corrections']={'resolved_ranges':['B2']}
        with self.assertRaises(ValueError):validate_review(answer,[table],snapshot,context,focused['review_focus'])

    def test_共同范围跨过标题只落实已有财务目标(self):
        from 附注更新.A语义阅读 import expand_scopes,reviewed_binding,distribute_shared_dimensions
        table={'id':'t','sheet':'表','range':'A1:C5','data_ranges':['B2:C2','B4:C5']}
        cells={a:{} for a in ['B2','C2','B4','C4','B5','C5']}
        snapshot={'sheets':[{'name':'表','cells':cells}]}
        standard={'metric':{'金额':{}},'dimension':{'期间':{'value_type':'string'}}}
        row={'id':'t','metrics':[{'id':'金额','ranges':['B2:C5']}],
            'dimensions':[{'id':'期间','ranges':['B2:C5'],'value':'本期','evidence':['A1']}],
            'unresolved':[{'range':'B4:C5','reason':'待核'}]}
        review={'tables':[{'id':'t','corrections':{'patches':[{'id':'期间','ranges':['C2:C5'],'value':'上期','evidence':['C1']}],
            'resolved_ranges':['B3:C5']}}]}
        corrected=reviewed_binding({'tables':[row]},review,[table],snapshot)
        expanded=expand_scopes(corrected['tables'][0],table,cells,standard)
        self.assertEqual(set(expanded),set(cells))
        self.assertTrue(all(r['dimensions']['期间']['value']==('本期' if a.startswith('B') else '上期') for a,r in expanded.items()))
        self.assertEqual(corrected['tables'][0]['unresolved'],[])
        row['dimensions'][0]['ranges']=['B3:C3']
        with self.assertRaises(ValueError):distribute_shared_dimensions({'tables':[row]},[table],snapshot)
        row['dimensions'][0]['ranges']=['B2:C6']
        with self.assertRaises(ValueError):expand_scopes(row,table,cells,standard)

    def test_共用身份不能抹平不同行对象(self):
        from 附注更新.A语义阅读 import duplicate_identity_questions
        tables=[{'id':'t','sheet':'表','range':'A1:C3','data_ranges':['B2:C3']}]
        cells={a:{'row':r,'column':c,'value':v} for a,r,c,v in [('A2',2,1,'甲单位'),('A3',3,1,'乙单位'),
            ('B2',2,2,None),('C2',2,3,None),('B3',3,2,None),('C3',3,3,None)]}
        snapshot={'path':'原件.xlsx','sha256':'未改','sheets':[{'name':'表','cells':cells}]}
        standard={'metric':{'余额':{'dimension_refs':[{'id':'对象'}]}},'dimension':{'对象':{'value_type':'string'}}}
        row={'id':'t','metrics':[{'id':'余额','ranges':['B2:C3']}],'dimensions':[{'id':'对象','ranges':['*'],'value':'甲乙合集','evidence':['A2:A3']}]}
        questions=duplicate_identity_questions({'tables':[row]},tables,snapshot,standard)
        self.assertEqual({tuple(q['ranges']) for q in questions},{('B2','B3'),('C2','C3')})
        from 附注更新.A语义阅读 import expand_scopes
        expanded=expand_scopes(row,tables[0],cells,standard)
        with patch('附注更新.A语义阅读.expand_scopes',return_value=dict(reversed(list(expanded.items())))):
            self.assertEqual(duplicate_identity_questions({'tables':[row]},tables,snapshot,standard),questions)
        row['dimensions']=[{'id':'对象','ranges':['B2','C3'],'value':'相同业务','evidence':['A2']},
            {'id':'对象','ranges':['C2','B3'],'value':'另一业务','evidence':['A3']}]
        self.assertEqual(duplicate_identity_questions({'tables':[row]},tables,snapshot,standard),[])
        from 附注更新.A语义阅读 import validate_review
        snapshot['sheets'][0]['merges']=[]
        converted=validate_review({'t':{'id':'t','corrections':{},'issues':[]}},tables,snapshot,[])
        self.assertEqual(converted,{'tables':[{'id':'t','corrections':{},'issues':[]}]})
        raw={'tables':[{'id':'t','corrections':{'patches':[{'id':'metric','ranges':['B2'],'value':'metric.新指标','evidence':['甲单位']}]},'issues':[]}]}
        original=json.dumps(raw,ensure_ascii=False)
        converted=validate_review(raw,tables,snapshot,[])
        self.assertEqual(json.dumps(raw,ensure_ascii=False),original)
        self.assertEqual(converted['tables'][0]['corrections'],{'patches':[],'metrics':[{'ranges':['B2'],'id':'metric.新指标'}]})
        duplicate={'tables':[{'id':'t','corrections':{'metrics':[{'ranges':['B2'],'id':'metric.新指标'}],
            'patches':[{'id':'metric.新指标','ranges':['B2'],'evidence':['甲单位']}]},'issues':[]}]}
        self.assertEqual(validate_review(duplicate,tables,snapshot,[])['tables'][0]['corrections']['metrics'],[{'ranges':['B2'],'id':'metric.新指标'}])

    def test_标准检索保留概括未展开的原项目(self):
        from 附注更新.A语义阅读 import selection_payload
        reading={'tables':[{'id':'t','sheet':'表','range':'A1:B2','meaning':'现金流量调节表'}]}
        snapshot={'path':'原件.xlsx','sha256':'未改','sheets':[{'name':'表','cells':{
            'A1':{'value':'投资收益'},'B1':{'value':100},'A2':{'value':'财务费用'},'B2':{'value':None}}}]}
        payload=selection_payload(reading,snapshot,{'metric':{}})
        self.assertEqual(payload['原表文字'][0]['text'],{'A1':'投资收益','A2':'财务费用'})
        self.assertEqual(payload['reading'],reading)
        from 附注更新.A语义阅读 import metric_catalog
        standard={'metric':{'x':{'id':'x','status':'active','definition':{'economic_object':'营业外支出','measure':'发生额','time_kind':'duration',
            'measurement_basis':'包括有据的捐赠、罚款、毁损报废等事项。不是现金支付。'}}}}
        self.assertIn('毁损报废',metric_catalog(standard)[0]['basis'])
        self.assertNotIn('basis',metric_catalog(standard,include_basis=False)[0])

    def test_补查可增加现行标准但不能改定义(self):
        from 附注更新.A语义阅读 import standard_material,supplied_review_standards
        standard={'metric':{k:{'id':k,'status':'active','definition':{'economic_object':k},'dimension_refs':[]} for k in ['a','b']},'dimension':{}}
        original=standard_material(['a'],standard);extra=standard_material(['a','b'],standard)
        self.assertEqual(supplied_review_standards(original,extra,standard),extra)
        with self.assertRaises(ValueError):supplied_review_standards(original,standard_material(['b'],standard),standard)
        extra['metrics'][1]['definition']['economic_object']='篡改'
        with self.assertRaises(ValueError):supplied_review_standards(original,extra,standard)

    def test_明确范围的嵌套回执不重发模型(self):
        from 附注更新.A语义阅读 import validate_bindings
        tables=[{'id':'t','sheet':'表','range':'A1:B3','data_ranges':['B2:B3']}]
        snapshot={'sheets':[{'name':'表','cells':{'B2':{},'B3':{}}}]}
        dim={'id':'entity','ranges':['B2'],'value':'甲','evidence':['A1']}
        answer={'tables':[{'id':'表','metrics':[{'id':'金额','ranges':['B2'],'dimensions':[dim]}],'unresolved':[{'range':'B3','reason':'缺条件'}]}]}
        got=validate_bindings(answer,tables,snapshot,[],{},candidate=True)
        self.assertEqual(got['tables'][0]['id'],'t')
        self.assertEqual(got['tables'][0]['dimensions'],[dim])
        self.assertNotIn('dimensions',got['tables'][0]['metrics'][0])
        self.assertEqual(answer['tables'][0]['id'],'表')
        with self.assertRaises(ValueError):validate_bindings(answer,tables+[dict(tables[0],id='u')],snapshot,[],{},candidate=True)
        dim['ranges']=['B3']
        with self.assertRaises(ValueError):validate_bindings(answer,tables,snapshot,[],{},candidate=True)

    def test_共同原文只解析到当前格适用的原说明(self):
        from 附注更新.A语义阅读 import normalize_evidence
        aliases={'T0':[{'ref':'@context:a','scope':{'sheet':'表','bounds':[1,1,3,2]}},
                       {'ref':'@context:b','scope':{'sheet':'表','bounds':[5,1,7,2]}}]}
        self.assertEqual(normalize_evidence(['T0'],'表',aliases,'B2'),['@context:a'])
        self.assertEqual(normalize_evidence('T0','表',aliases,'B2'),['@context:a'])
        self.assertEqual(normalize_evidence(['@T0'],'表',aliases,'B2'),['@context:a'])
        self.assertEqual(normalize_evidence(['T0'],'表',aliases,'B6'),['@context:b'])
        self.assertEqual(normalize_evidence(['原文引用：T0'],'表',aliases,'B6'),['@context:b'])
        self.assertEqual(normalize_evidence(['表!C87（横线）'],'表',aliases),['C87'])
        self.assertEqual(normalize_evidence(['B110=固定资产'],'表',aliases),['B110'])
        self.assertEqual(normalize_evidence(['B110：固定资产，附带解释'],'表',aliases),['B110'])
        with self.assertRaises(ValueError):normalize_evidence(['大约在C87附近'],'表',aliases)
        with self.assertRaises(ValueError):normalize_evidence(['T0'],'表',aliases,'B4')
        from 附注更新.A语义阅读 import source_material
        cells={a:{'row':r,'column':c,'value':v} for a,r,c,v in [('A1',1,1,'类别'),('B1',1,2,'期末'),('A2',2,1,'其他'),('A3',3,1,'其他'),('B2',2,2,None),('B3',3,2,None)]}
        _,quoted=source_material({'path':'引用检查.xlsx','sheets':[{'name':'表','cells':cells,'merges':[]}]},[],[{'sheet':'表','range':'A1:B3'}])
        self.assertEqual(normalize_evidence(['期末'],'表',quoted,'B2'),['B1'])
        self.assertEqual(normalize_evidence(['其他'],'表',quoted,'B2'),['A2'])
        self.assertEqual(normalize_evidence(['其他'],'表',quoted,'B3'),['A3'])
        cells['C2']={'row':2,'column':3,'value':'其他'}
        _,ambiguous=source_material({'path':'引用检查.xlsx','sheets':[{'name':'表','cells':cells,'merges':[]}]},[],[{'sheet':'表','range':'A1:C3'}])
        with self.assertRaises(ValueError):normalize_evidence(['其他'],'表',ambiguous,'B2')
        cells['C1']={'row':1,'column':3,'value':'期末'}
        _,headers=source_material({'path':'引用检查.xlsx','sheets':[{'name':'表','cells':cells,'merges':[]}]},[],[{'sheet':'表','range':'A1:C3'}])
        self.assertEqual(normalize_evidence('期末','表',headers,'B2'),['B1','C1'])
        _,context_quotes=source_material({'path':'引用检查.xlsx','sheets':[{'name':'表','cells':cells,'merges':[]}]},
            ['本报表为单户报表，所有金额均以人民币元为货币单位。'],[{'sheet':'表','range':'A1:C3'}])
        self.assertEqual(normalize_evidence(['单户报表','人民币元'],'表',context_quotes,'B2'),[context_quotes['@C0']])
        with self.assertRaises(ValueError):normalize_evidence(['单户报告'],'表',context_quotes,'B2')
        cells['A2']['value']='固定资产报废损失（收益以负数列示）'
        _,quoted=source_material({'path':'引用检查.xlsx','sheets':[{'name':'表','cells':cells,'merges':[]}]},[],[{'sheet':'表','range':'A1:C3'}])
        self.assertEqual(normalize_evidence(['固定资产报废损失'],'表',quoted,'B2'),['A2'])
        with self.assertRaises(ValueError):normalize_evidence(['固定资产出售收益'],'表',quoted,'B2')

    def test_核查只改指定范围不抹掉其他行(self):
        from 附注更新.A语义阅读 import reviewed_binding
        binding={'tables':[{'id':'a','dimensions':[{'id':'对象','ranges':['*'],'value':'原对象','evidence':['A1']}],'unresolved':[{'range':'B2:B3','reason':'原缺项'}]}]}
        review={'tables':[{'id':'a','corrections':{'patches':[{'id':'对象','ranges':['B3'],'value':'合计对象','evidence':['A3']}],'resolved_ranges':['B3']}}]}
        table={'id':'a','sheet':'表','range':'A1:B3','data_ranges':['B2:B3']}
        got=reviewed_binding(binding,review,[table],{'sheets':[{'name':'表','cells':{'B2':{},'B3':{}}}]})
        self.assertEqual([(d['ranges'],d['value']) for d in got['tables'][0]['dimensions']],[(['B2'],'原对象'),(['B3'],'合计对象')])
        self.assertEqual(got['tables'][0]['unresolved'],[{'range':'B2','reason':'原缺项'}])
        binding['tables'][0]['metrics']=[{'id':'原指标','ranges':['B2:B3']}]
        review['tables'][0]['corrections']['metrics']=[{'id':'正确指标','ranges':['B3']}]
        got=reviewed_binding(binding,review,[table],{'sheets':[{'name':'表','cells':{'B2':{},'B3':{}}}]})
        self.assertEqual(got['tables'][0]['metrics'],[{'id':'原指标','ranges':['B2']},{'id':'正确指标','ranges':['B3']}])

    def test_跨表共同信息按明确范围分配且越界拒绝(self):
        from 附注更新.A语义阅读 import distribute_shared_dimensions
        tables=[{'id':'a','sheet':'表','range':'A1:B2','data_ranges':['B2']},
                {'id':'b','sheet':'表','range':'A3:B4','data_ranges':['B4']}]
        snapshot={'sheets':[{'name':'表','cells':{'B2':{},'B4':{}}}]}
        dimension={'id':'entity','value':'甲','ranges':['B2','B4'],'evidence':['@C0']}
        answer={'tables':[{'id':'a','dimensions':[dimension]},{'id':'b','dimensions':[]}]}
        got=distribute_shared_dimensions(answer,tables,snapshot)
        self.assertEqual([r['dimensions'][0]['ranges'] for r in got['tables']],[['B2'],['B4']])
        self.assertEqual(answer['tables'][1]['dimensions'],[])
        dimension['ranges'].append('B5')
        with self.assertRaises(ValueError):distribute_shared_dimensions(answer,tables,snapshot)

    def test_共同语义落实到空白格且例外不丢失(self):
        self.assertIsNotNone(importlib.util.find_spec('附注更新.A语义阅读'), '普通A流程尚无简洁语义实现')
        from 附注更新.A语义阅读 import expand_scopes
        source={'B5':{'value':None},'C5':{'value':100},'B6':{'value':10},'C6':{'value':None}}
        table={'data_ranges':['B5:C6'],'range':'A1:C6'}
        standard={'metric':{'金额':{},'比例':{}},'dimension':{'entity':{'value_type':'string'},'对象':{'value_type':'string'}}}
        answer={'metrics':[{'ranges':['B5:B6'],'id':'金额'},{'ranges':['C5:C6'],'id':'比例'}],
                'dimensions':[{'ranges':['*'],'id':'entity','value':'甲公司','evidence':['A1']},
                    {'ranges':['B5:C5'],'id':'对象','value':'一年以内','evidence':['A5']},
                    {'ranges':['B6:C6'],'id':'对象','value':'全体','evidence':['A6']}],
                'unresolved':[{'range':'C6','reason':'比例口径缺说明'}]}
        got=expand_scopes(answer,table,source,standard)
        self.assertEqual(set(got),{'B5','C5','B6','C6'})
        self.assertEqual(got['B5']['metric_id'],'金额')
        self.assertEqual(got['B5']['dimensions']['对象']['value'],'一年以内')
        self.assertEqual(got['B6']['dimensions']['对象']['value'],'全体')
        from 附注更新.A语义阅读 import shared_binding
        before={'tables':[{'id':'t',**answer}]}
        # 拆开原本共同的两格，确认合并只减少重复表达，不改变展开结果。
        before['tables'][0]['metrics']=[{'ranges':['B5'],'id':'金额'},{'ranges':['B6'],'id':'金额'},answer['metrics'][1]]
        merged=shared_binding(before)
        self.assertEqual(len(merged['tables'][0]['metrics']),2)
        self.assertEqual(expand_scopes(merged['tables'][0],table,source,standard),got)
        self.assertIsNone(got['C6']['metric_id'])
        self.assertEqual(source['B5']['value'],None)
        answer['dimensions'].append({'ranges':['B5'],'id':'entity','value':'甲公司','evidence':['A1']})
        self.assertEqual(expand_scopes(answer,table,source,standard)['B5']['dimensions']['entity']['value'],'甲公司')
        from 附注更新.A语义阅读 import normalize_evidence
        self.assertEqual(normalize_evidence(['表!$A$1:A2','@C0'],'表',{'@C0':'@context:abc'}),['A1','A2','@context:abc'])
        answer['dimensions'][-1]['value']='乙公司'
        with self.assertRaises(ValueError):expand_scopes(answer,table,source,standard)
        answer['dimensions'].pop()
        answer['metrics'][0]['ranges']=['B5:D6']
        with self.assertRaises(ValueError):expand_scopes(answer,table,source,standard)

    def test_普通A流程先读语义再查标准且正式回读拒绝篡改(self):
        from 附注更新.统一识别 import UnifiedSemanticEngine
        from 附注更新.统一勾稽 import load_fact_mapping
        from 附注更新.表格 import file_hash
        base=Path(r'C:\Users\27651\Documents\Codex\2026-09-29\c-users-27651-documents-codex-2026-2')
        scope=json.loads((base/'outputs/真实A小表_低思考强度对照_20260930_005423/实验范围.json').read_text(encoding='utf-8-sig'))
        work=base/'work'/('A正式回读定向检查_'+str(time.time_ns()))
        engine=UnifiedSemanticEngine({'base_url':'https://example.invalid','model':'mimo-v2.6-flash','thinking_mode':'default'},scope['金标准'],work)
        context=['甲公司单体报表，人民币元；期末2025年12月31日，期初2024年12月31日。']
        table={'id':'t1','sheet':'附注小表','range':'A1:G6','subject':'预付款项','meaning':'账龄为行，期初期末各有余额、占比、准备','data_ranges':['A5:G5','B6:G6']}
        state=[]
        def dim(value,ref='@C0'):return {'value':value,'evidence':[ref]}
        def fake(system,payload):
            task=payload['task'];state.append(task)
            self.assertEqual(engine.settings['thinking_mode'],'default')
            if task=='a_semantic_read':
                self.assertNotIn('standards',payload);self.assertNotIn('catalog',payload)
                return {'tables':[{**table,'data_ranges':['C5:G6']}],'notes':'本地代答故意漏掉金额列，检验有据局部补读'}
            if task=='a_semantic_read_check':
                self.assertEqual(payload['questions'][0]['数据范围外的数值'],['B5','B6'])
                return {'tables':[table]}
            if task=='a_semantic_select':
                self.assertEqual(state[0],'a_semantic_read')
                return {'tables':[{'id':'t1','metric_ids':['metric.prepayment.gross_balance','metric.prepayment.gross_balance.share_ratio','metric.prepayment.impairment_allowance.balance'],'missing':''}]}
            if task=='a_semantic_bind':
                self.assertEqual(payload['reading_notes'],'本地代答故意漏掉金额列，检验有据局部补读')
                self.assertIn('数值处理',payload['source'])
                self.assertIn('[待确认业务含义]',payload['source']['sheets'][0]['rows'])
                self.assertNotIn('35,574,767.76',payload['source']['sheets'][0]['rows'])
                self.assertEqual({'metric.prepayment.gross_balance','metric.prepayment.gross_balance.share_ratio','metric.prepayment.impairment_allowance.balance'}, {m['id'] for m in payload['standards']['metrics']})
                metrics=[];dimensions=[{'ranges':['*'],'id':key,**dim(value)} for key,value in
                    [('entity','甲公司'),('report_scope','standalone'),('currency','CNY'),('unit_scale',1)]]
                for col,kind in [('B','gross_balance'),('C','gross_balance.share_ratio'),('D','impairment_allowance.balance'),('E','gross_balance'),('F','gross_balance.share_ratio'),('G','impairment_allowance.balance')]:
                    metrics.append({'ranges':[col+'5:'+col+'6'],'id':'metric.prepayment.'+kind})
                    dimensions.append({'ranges':[col+'5:'+col+'6'],'id':'period',**dim({'kind':'instant','date':'2025-12-31' if col in 'BCD' else '2024-12-31'})})
                next(d for d in dimensions if d['id']=='currency')['ranges']=['B5:B6','D5:E6','G5:G6']
                next(d for d in dimensions if d['id']=='unit_scale')['ranges']=['B5:B6','D5:E6','G5:G6']
                dimensions.extend([{'ranges':['C5:C6','F5:F6'],'id':'prepayment_denominator_selection',**dim({'domain':'预付款项','mode':'all','completeness':'complete'},'A6')},
                    {'ranges':['B5:G5'],'id':'prepayment_selection',**dim({'domain':'预付款项','mode':'predicate','predicate':'账龄一年以内含一年','completeness':'complete'},'A5')},
                    {'ranges':['B6:G6'],'id':'prepayment_selection',**dim({'domain':'预付款项','mode':'all','completeness':'complete'},'A6')}])
                dimensions.append({'ranges':['B5'],'id':'临时备注','value':'错误新增字段','evidence':['A5']})
                return {'tables':[{'id':'t1','metrics':metrics,'dimensions':dimensions,'unresolved':[]}]}
            self.assertEqual(task,'a_semantic_review')
            self.assertEqual(payload['reading_notes'],'本地代答故意漏掉金额列，检验有据局部补读')
            self.assertTrue(payload['program_issues'])
            return {'tables':[{'id':'t1','corrections':{
                'remove_dimensions':['临时备注'],
                'dimensions':[{'ranges':['*'],'id':'currency',**dim('CNY')}],
                'non_business':[{'range':'A5','category':'label','reason':'账龄是解释金额属于哪类的身份标签','evidence':['A2','A5']}]},'issues':[]}]}
        with patch.object(engine,'_http',side_effect=fake):
            result=engine.recognize_workbook(scope['原表'],context=context,artifact_metadata={'carrier':'A','original_path':scope['原表'],'original_hash':file_hash(scope['原表'])})
        self.assertTrue(result['complete'],[json.loads(r['text'].lstrip('\ufeff'))['error'] for r in result.get('未通过答复',[])])
        loaded=load_fact_mapping(result['mapping_path'],engine.standard)
        self.assertEqual(len(loaded['facts']),12)
        self.assertEqual(sum(f['raw_value_state']=='blank' for f in loaded['facts']),6)
        self.assertEqual(loaded['facts'],result['facts'])
        self.assertEqual(len(state),5)
        from copy import deepcopy
        with patch('附注更新.A语义阅读._stage_receipt',side_effect=AssertionError('新合集不能重读外部历史回执')):
            self.assertEqual(load_fact_mapping(result['mapping_path'],engine.standard)['facts'],result['facts'])
        from 附注更新.A合集 import validate_collection,FIELDS
        from 附注更新.统一语义 import definition_hash
        changed=deepcopy(loaded)
        doc=next(d for d in changed['原判断与核查'].values() if json.loads(d['task_text'].lstrip('\ufeff'))['payload']['task']=='a_semantic_bind')
        doc['task_text']=doc['task_text'].replace('本地代答故意漏掉金额列','并非整体阅读所得的内容')
        changed['封存校验']=definition_hash({k:changed.get(k) for k in FIELDS})
        with self.assertRaisesRegex(ValueError,'原任务或回执'):validate_collection(changed)
        with patch.object(engine,'_http',side_effect=AssertionError('已通过结果不应再请求模型')):
            reused=engine.recognize_workbook(scope['原表'],context=context,previous_result=result['mapping_path'],artifact_metadata={'carrier':'A','original_path':scope['原表'],'original_hash':file_hash(scope['原表'])})
        self.assertEqual(reused['facts'],result['facts'])
        self.assertEqual(reused['usage']['reused_confirmed_tables'],1)
        # 新A不再接受程序按标题或同文生成后续语义轮；未决及停止重答另由封存边界反例覆盖。
        loaded.pop('_validated_standard_hash',None)
        loaded['facts'][0]['dimensions']['entity']='伪造主体'
        bad=work/'篡改副本.json';bad.write_text(json.dumps(loaded,ensure_ascii=False),encoding='utf-8-sig')
        with self.assertRaises(ValueError):load_fact_mapping(bad,engine.standard)


if __name__=='__main__':unittest.main()
