# coding: utf-8
"""只验证漏找补查和统一身份证的实际对应行为。"""
import copy
import unittest
from unittest.mock import patch

from 测试 import 测试B范围定位 as 原定位测试
from 测试.测试统一语义 import records, fact
from 附注更新.统一语义 import validate_records
from 附注更新.统一识别 import UnifiedSemanticEngine


class 两项缺陷测试(unittest.TestCase):
    def test_共用识别保留公式引用且不暴露数值作匹配(self):
        from 附注更新.A语义阅读 import source_material
        snapshot={'path':'原件.xlsx','semantic_formula_sources':True,'sheets':[{'name':'附注','merges':[],
            'cells':{'B3':{'row':3,'column':2,'value':"='明细'!C8",'formula':"='明细'!C8",'cached_value':123.45}}}]}
        tables=[{'sheet':'附注','range':'B3','data_ranges':['B3']}]
        material,_=source_material(snapshot,[],tables,semantic_only=True)
        self.assertEqual(material.get('公式原文'),{'附注!B3':"='明细'!C8"})
        self.assertNotIn('123.45',str(material))
        snapshot.pop('semantic_formula_sources')
        legacy,_=source_material(snapshot,[],tables,semantic_only=True)
        self.assertNotIn('公式原文',legacy)

    def test_已找到同槽位时不同年份不触发补查(self):
        from openpyxl import load_workbook
        from 附注更新.统一勾稽 import load_fact_mapping
        case=原定位测试.B范围定位测试();case.setUp()
        book=load_workbook(case.case.source)
        book['综合']['B2']='2024年原值';book['综合']['C2']='2025年原值'
        book['综合']['C3']=126
        case.case.source=case.case.folder/'先找到上年候选.xlsx'
        book.save(case.case.source);book.close()
        def reply(engine,system,payload):
            task=payload['task']
            if task=='b_candidate_followup':
                raise AssertionError('年份取值不同不应当成缺少金标准槽位')
            if task=='a_semantic_read':
                return {'tables':[{'id':'原值','sheet':'综合','range':'A1:C3','subject':'固定资产',
                    'meaning':'B列2024年原值，C列2025年原值','data_ranges':[p.rpartition('!')[2] for p in payload['candidate_locations']]}],
                    'notes':'分别保留两列实际年份'}
            if task=='a_semantic_bind':
                cells=payload['reading'][0]['data_ranges'];dimensions=[]
                for cell in cells:
                    dims=copy.deepcopy(fact()['dimensions'])
                    if cell=='B3':dims['period']={'kind':'instant','date':'2024-12-31'}
                    dimensions.extend({'ranges':[cell],'id':key,'value':value,'evidence':['A1',cell[0]+'2','A3']} for key,value in dims.items())
                return {'tables':[{'id':'原值','metrics':[{'ranges':cells,'id':'固定资产原值'}],'dimensions':dimensions,'unresolved':[]}]}
            return case.response(engine,system,payload)
        with patch.object(UnifiedSemanticEngine,'_http',reply):
            engine=case.case.engine();result=engine.recognize_workbook(case.case.source,target_mapping=case.case.target())
        self.assertEqual(result['recognition_scope']['coverage']['targets'][0]['standard_sources'],['综合!B3'])
        self.assertEqual(result['recognition_scope']['coverage']['targets'][0]['exact_sources'],[])
        saved=load_fact_mapping(result['mapping_path'],engine.standard)
        self.assertEqual({f['source_reference']['location'] for f in saved['facts']},{'综合!B3'})
        self.assertEqual(saved['facts'][0]['dimensions']['period']['date'],'2024-12-31')

    def test_报表口径不阻断格子身份及对应(self):
        from 附注更新.统一语义 import validate_fact, semantic_key, validate_facts
        standard=validate_records(records())
        left=fact(); right=copy.deepcopy(left)
        left['dimensions'].pop('report_scope');left['dimension_evidence'].pop('report_scope')
        right['dimensions']['report_scope']='consolidated'
        validate_fact(left,standard)
        self.assertEqual(semantic_key(left,standard),semantic_key(right,standard))
        left.update(id='格子',owner_reference='对象')
        validate_facts([left],standard,{'对象':copy.deepcopy(left['dimensions'])})
        left['dimensions'].pop('asset_selection');left['dimension_evidence'].pop('asset_selection')
        with self.assertRaisesRegex(ValueError,'缺少必需维度'):
            validate_fact(left,standard)

    def test_公式尚未取数仍展示身份证对应(self):
        from 测试.测试统一勾稽 import mapping
        from 附注更新.统一勾稽 import build_fact_update_plan
        standard=validate_records(records());standard['sha256']='a'*64
        target=mapping('综合!B3',100);source=mapping('来源!C8',125)
        source['facts'][0].update(raw_value_state='formula',raw_value='=1+1')
        plan=build_fact_update_plan(target,source,standard)
        self.assertEqual(plan['identity_matches'][0]['source_locations'],['来源!C8'])
        self.assertEqual(plan['updates'],[])

    def test_Word正文编制基础进入共用识别材料(self):
        from tempfile import mkdtemp
        from pathlib import Path
        from zipfile import ZipFile
        from 附注更新.分步流程 import _recognition_context
        from 附注更新.表格 import file_hash
        word=Path(mkdtemp(prefix='AB正文定向检查_'))/'附注.docx'
        # 第一张是章节外的会计政策表；第二张已单独提取，避免重复进入全文上下文。
        with ZipFile(word,'w') as package:
            package.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>应收款项会计政策</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>信用风险组合</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>按账龄确定损失率</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:tbl><w:tr><w:tc><w:p><w:r><w:t>已提取财务数据</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>')
        table={'title':'货币资金','word_body_index':2}
        result=_recognition_context({'word_path':str(word),'word_hash':file_hash(word),'context':{'tables':[table]}})
        self.assertEqual(result[0],table)
        self.assertIn('应收款项会计政策',result[1]['text'])
        self.assertIn('信用风险组合 | 按账龄确定损失率',result[1]['text'])
        self.assertNotIn('已提取财务数据',result[1]['text'])
        self.assertEqual(result[1]['source_hash'],file_hash(word))

    def test_金标准身份证直接对应且期间不同不对应(self):
        from 附注更新.表格编排 import target_requirements
        from 附注更新.统一勾稽 import build_fact_update_plan
        standard = validate_records(records()); standard['sha256'] = 'a' * 64
        from 测试.测试统一勾稽 import mapping
        target = mapping('综合!B3', 100)
        source = mapping('综合!C8', 125)
        self.assertIsNotNone(target_requirements(target, standard)[0]['semantic_key'])
        plan = build_fact_update_plan(target, source, standard)
        self.assertEqual([(r['cell'], r['value']) for r in plan['updates']], [('B3', 125)])
        self.assertEqual(plan['issues'], [])
        source['facts'][0]['dimensions']['period'] = {
            'kind': 'instant', 'date': '2024-12-31'}
        source['owners']['对象']['period'] = source['facts'][0]['dimensions']['period']
        changed = build_fact_update_plan(target, source, standard)
        self.assertEqual(changed['updates'], [])
        self.assertTrue(any(r['kind'] == 'missing_source' for r in changed['issues']))

    def test_只有文字和最近标准不能作为成功绑定(self):
        from 附注更新.A语义阅读 import validate_bindings
        with self.assertRaisesRegex(ValueError, '识别成功必须完整对应金标准'):
            validate_bindings({'tables': [{'source_identities': [{'meaning': '原值'}]}]},
                              [], {}, [], {}, candidate=True, require_identity=True)
        from 附注更新.统一勾稽 import build_fact_update_plan
        from 测试.测试统一勾稽 import mapping
        standard=validate_records(records());standard['sha256']='a'*64
        legacy=mapping('综合!B3');legacy['source_identities']=[{'meaning':'固定资产原值','closest_metric_id':'固定资产原值'}]
        with self.assertRaisesRegex(ValueError, '先准确对应金标准再比较'):
            build_fact_update_plan(legacy, mapping('来源!C2'), standard)
        from types import SimpleNamespace
        from 附注更新.界面 import Application
        legacy.update(facts=[],excluded=[],unresolved=[{'sheet':'综合','cell':'B3'}],complete=False,
                      coverage={'complete':False},a_identity_coverage={'complete':True,'financial_cells':1,'unresolved_cells':0})
        shown={};label=SimpleNamespace(winfo_exists=lambda:True,configure=lambda **kw:shown.update(kw))
        app=SimpleNamespace(mapping_summary_labels={'a_mapping':label},values={'a_mapping':SimpleNamespace(get=lambda:'旧成果')})
        with patch('附注更新.界面._read_record', return_value=legacy):
            Application.refresh_mapping_summary(app,'a_mapping')
        self.assertIn('A识别未完成',shown['text'])
        self.assertIn('准确对应金标准 0 格',shown['text'])

    def test_全部区域已选仍补回遗漏候选并正式回读(self):
        from 附注更新.统一勾稽 import load_fact_mapping
        case = 原定位测试.B范围定位测试(); case.setUp()
        def reply(engine, system, payload):
            if payload['task'] == 'b_candidate_followup':
                return {'read_ranges': [], 'ranges': [{'sheet': '综合', 'range': 'A1:C3'}],
                        'structure': '复查原值列，B3是本次所需候选', 'evidence': ['综合!A2'],
                        'targets': [{'id': r['id'], 'candidates': ['综合!B3'],
                                     'reason': '原值列与全部固定资产行交叉，首轮漏列'} for r in payload['targets']]}
            answer = case.response(engine, system, payload)
            if payload['task']=='a_semantic_bind':
                self.assertEqual(payload['已确认维度取值'],[{'metric_id':'固定资产原值','dimension_id':'asset_selection',
                    'value':fact()['dimensions']['asset_selection']}])
                self.assertFalse(any(r['dimension_id'] in {'entity','period','report_scope','currency','unit_scale'} for r in payload['已确认维度取值']))
            if payload['task'] == 'subtable_detail':
                for row in answer['targets']: row.update(candidates=[], reason='首轮没有列入候选')
            return answer
        with patch.object(UnifiedSemanticEngine, '_http', reply):
            target = case.case.target(); engine = case.case.engine()
            result = engine.recognize_workbook(case.case.source, target_mapping=target)
        self.assertEqual(result['recognition_scope']['candidate_locations'], ['综合!B3'])
        self.assertEqual(result['recognition_scope']['coverage']['targets'][0]['exact_sources'], ['综合!B3'])
        saved = load_fact_mapping(result['mapping_path'], engine.standard)
        self.assertEqual(saved['dimension_vocabulary_version'],1)
        self.assertEqual(saved['source_identities'], [])
        self.assertEqual(saved['facts'][0]['metric_id'], '固定资产原值')
        from 附注更新.分步流程 import run_step
        import json
        with patch.object(UnifiedSemanticEngine, '_http', side_effect=AssertionError('对应阶段不得再调用模型')):
            matched = run_step('match', {'a_mapping':target['mapping_path'], 'b_mapping':result['mapping_path'],
                               'gold_path':str(case.case.gold),'output_dir':str(case.case.folder/'实际对应入口')})
        self.assertEqual(matched['status'], 'complete', matched['message'])
        plan=json.loads(__import__('pathlib').Path(matched['files']['update_plan']).read_text(encoding='utf-8-sig'))
        self.assertEqual([(row['cell'],row['source_group']) for row in plan['targets']], [('B3','固定资产原值')])
        self.assertNotIn('updates',plan)


if __name__ == '__main__':
    unittest.main()
