# coding: utf-8
"""位置、金额和历史指定来源都不能替代身份证。"""
from copy import deepcopy
import unittest
from unittest.mock import patch

from 附注更新.模板更新编排 import direct_update_plan


class 唯一身份证匹配测试(unittest.TestCase):
    def test_候选字段差异注明未交的一方且不决定补交范围(self):
        a=self.mapping('A',[('C3','占位',{'period':'2025-12-31','类别':'某一业务类别'},'期末余额')])
        b=self.mapping('B',[('D7','占位',{'period':'2025-12-31'},'期末余额')])
        for mapping in (a,b):
            mapping['逐格字段版本']=1
            row=mapping['facts'][0]
            row.pop('metric_id')
            row.update(record_type='dimension_binding',dimension_id='period')
        for left,right,side in ((a,b,'b'),(b,a,'a')):
            plan=self.plan(left,right)
            issue=next(row for row in plan['issues'] if row.get('kind')=='missing_dependency')
            candidate,=issue['candidates']
            self.assertEqual(candidate['missing_dimensions'],['类别'])
            self.assertEqual(candidate['missing_in_'+side],['类别'])
            self.assertEqual(candidate['missing_in_'+('a' if side=='b' else 'b')],[])
            self.assertIn('候选差异不能直接作为本格补交要求',issue['reason'])
            self.assertEqual(plan['identity_matches'],[])
            self.assertEqual(plan['updates'],[])

    def test_维度格公式沿用已有重算值且不改身份证(self):
        a=self.mapping('A',[('C3','占位',{'类别':'甲'},'原名称')])
        b=self.mapping('B',[('D7','占位',{'类别':'甲'},"='B'!Z1")])
        for mapping in (a,b):
            mapping['逐格字段版本']=1
            row=mapping['facts'][0]
            row.pop('metric_id')
            row.update(record_type='dimension_binding',dimension_id='类别')
        formula="='B'!Z1"
        b['facts'][0].update(raw_value_state='formula',raw_value=formula)
        b['机械记录']['sheets'][0]['cells']['D7'].update(formula=formula,cached_value='原缓存')
        b['calculation_evidence']={'path':'已有重算依据.json'}
        calculated={'B!D7':{'formula':formula,'state':'value','value':'已有重算原名称'}}
        with patch('pathlib.Path.read_text',return_value='{}'), \
             patch('附注更新.统一勾稽._calculated_values',return_value=calculated) as reader:
            result=self.plan(a,b)
        reader.assert_called_once()
        self.assertEqual(result['identity_matches'][0]['source_locations'],['B!D7'])
        self.assertEqual(result['updates'][0]['value'],'已有重算原名称')
        self.assertEqual(b['facts'][0]['raw_value'],formula)

    def test_正常入口不回落旧匹配或直接执行旧清单(self):
        from pathlib import Path
        from 附注更新.分步流程 import _match, _prepare_update_v2, _write_v2
        with patch('附注更新.分步流程._path',return_value=Path('旧映射')), \
             patch('附注更新.分步流程._read',return_value={'schema_version':1}), \
             patch('附注更新.表格.build_update_plan') as legacy:
            with self.assertRaisesRegex(ValueError,'已保存身份证合集'):
                _match({},None,None,None)
            legacy.assert_not_called()
        with patch('附注更新.分步流程._read',return_value={'collection_version':2,'gold_path':'金标准','gold_hash':'g'}), \
             patch('附注更新.A合集._standard',return_value={}), \
             patch('附注更新.统一勾稽.load_fact_mapping',return_value={'collection_version':1}), \
             patch('附注更新.统一勾稽.build_fact_update_plan') as legacy:
            with self.assertRaisesRegex(ValueError,'旧位置取数流程'):
                _prepare_update_v2({},None,None,None,Path('A'),Path('B'))
            legacy.assert_not_called()
        with patch('附注更新.分步流程._write_direct') as writer:
            with self.assertRaisesRegex(ValueError,'识别答复无需重做'):
                _write_v2({},None,None,None,None,{'schema':'附注更新清单-v2','updates':[]})
            writer.assert_not_called()

    def test_旧匹配标记不能直接进入写回而当前匹配可以(self):
        from pathlib import Path
        from 附注更新.分步流程 import _write_v2
        a=self.mapping('A',[('C3','余额',{'类别':'甲'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲'},99)])
        plan=self.plan(a,b)
        paths=[str(Path(name).resolve()) for name in ('合成A.json','合成B.json')]
        plan.update(a_mapping=paths[0],b_mapping=paths[1],inputs=dict.fromkeys(paths,'hash'))
        with patch('附注更新.分步流程._read',return_value={}), \
             patch('附注更新.分步流程._write_direct',return_value={'已进入当前写回':True}) as writer:
            self.assertEqual(_write_v2({},None,None,None,None,plan),{'已进入当前写回':True})
            writer.assert_called_once()
            for marker in ('unique_exact_identity','unique_exact_identity_v2','unique_exact_identity_v3'):
                legacy=deepcopy(plan);legacy['matching_method']=marker
                writer.reset_mock()
                with self.assertRaisesRegex(ValueError,'识别答复无需重做'):
                    _write_v2({},None,None,None,None,legacy)
                writer.assert_not_called()

    def mapping(self, carrier, rows):
        result = {'carrier': carrier, 'source_hash': carrier, 'gold_hash': 'gold',
                  'facts': [], 'slot_bindings': [], 'outline': {'subtables': []},
                  '机械记录': {'names': {}, 'sheets': [{'name': carrier, 'cells': {}, 'merges': []}]}}
        for cell, metric, dimensions, value in rows:
            result['facts'].append({'record_type': 'metric_fact', 'metric_id': metric, 'dimensions': dimensions,
                'source_reference': {'location': carrier + '!' + cell}})
            result['机械记录']['sheets'][0]['cells'][cell] = {'value': value}
        return result

    def plan(self, a, b, purpose=None, scope=None, comparison_requirements=None):
        before = deepcopy((a, b))
        with patch('附注更新.A合集.validate_direct_collection'), \
             patch('附注更新.A合集._standard', return_value={'dimension': {'类别': {}, 'period': {}, 'report_scope': {}}}):
            result = direct_update_plan(a, b, purpose, scope,comparison_requirements=comparison_requirements)
        self.assertEqual((a, b), before)
        return result

    def relation(self, bindings):
        return {'kind': '取数对应', 'source_a_hash': 'A', 'source_b_hash': 'B',
                'gold_hash': 'gold', 'bindings': bindings}

    def test_旧模型名单答复不再生成结构和原值搬运(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲'},99)])
        a['机械记录']['names']={'名单':[{'sheet':'A','range':'B2:C3'}]}
        expected=self.plan(a,b)
        for source_a_hash in ('A','旧原件'):
            b['已确认关系']=[{'kind':'名单编排','source_a_hash':source_a_hash,
                'source_b_hash':'B','gold_hash':'gold','dimension':'类别','tables':[
                    {'range_name':'名单','first_row_offset':1,'old_detail_count':1,
                     'source_sheet':'B','source_rows':[7,8],'columns':{'C':'D'},'labels':{},
                     'retained_members':[{'index':1,'cells':{'C':{'source':'A!C3'}}}]}]}]
            actual=self.plan(a,b)
            for field in ('updates','labels','operations','identity_matches','template_layout','update_purpose'):
                self.assertEqual(actual[field],expected[field])
            self.assertEqual(actual['list_relations'],[])

    def test_旧模型额外文字和保留指令不覆盖身份证匹配(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲','period':'2024'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲','period':'2025'},99)])
        for mapping,cell,label in ((a,'C3','C2'),(b,'D7','D6')):
            mapping['outline']['subtables']=[{'sheet':mapping['carrier'],'range':'A1:H10',
                'metrics':[{'id':'余额','ranges':[cell],'template_period_role':'closing'}]}]
            mapping['facts'][0]['label_references']={'period':[mapping['carrier']+'!'+label]}
            mapping['机械记录']['sheets'][0]['cells'][label]={'value':mapping['facts'][0]['dimensions']['period']}
        a['机械记录']['names']={'名单':[{'sheet':'A','range':'B2:C3'}]}
        expected=self.plan(a,b)
        self.assertEqual([row['value'] for row in expected['updates']],[99])
        self.assertEqual(expected['labels'],[])
        b['已确认关系']=[{**self.relation([]),'retained_targets':['A!C3'],
            'label_actions':[{'target':'A!C2','action':'retain'},
                             {'target':'A!C3','action':'set','value':'额外覆盖'}],
            'caption_updates':[{'range_name':'名单','old_text':'原题','text':'另拟标题'}]}]
        actual=self.plan(a,b)
        for field in ('updates','labels','caption_updates','identity_matches','retained_label_targets'):
            self.assertEqual(actual[field],expected[field])
        b['facts']=[]
        actual=self.plan(a,b)
        self.assertEqual(actual['updates'],[])
        self.assertEqual(actual['labels'],[])
        self.assertEqual(actual['caption_updates'],[])

    def test_明确机械原值恢复不依赖旧模型安排(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲'},99)])
        a['机械记录']['names']={'名单':[{'sheet':'A','range':'B2:E3'}]}
        a['机械记录']['sheets'][0]['cells']['E3']={'value':42}
        purpose={'tables':[{'range_name':'名单','first_row_offset':1,'operations':[],
            'source_rows':[],'columns':{},'labels':{},'retained_members':[
                {'index':0,'cells':{'D':{'value':None},'E':{'source':'A!E3'}}}]}]}
        actual=self.plan(a,b,purpose)
        self.assertEqual({row['cell']:row['value'] for row in actual['updates']},
                         {'C3':99,'D3':None,'E3':42})
        self.assertEqual(actual['operations'],[])

    def test_旧模型指令经副本接续后仍不能屏蔽对应和标签(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲','period':'2024'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲','period':'2025'},99)])
        for mapping,cell,label in ((a,'C3','C2'),(b,'D7','D6')):
            mapping['outline']['subtables']=[{'sheet':mapping['carrier'],'range':'A1:H10',
                'metrics':[{'id':'余额','ranges':[cell],'template_period_role':'closing'}]}]
            mapping['facts'][0]['label_references']={'period':[mapping['carrier']+'!'+label]}
            mapping['机械记录']['sheets'][0]['cells'][label]={'value':mapping['facts'][0]['dimensions']['period']}
        expected=self.plan(a,b)
        a['副本接续']={'已执行安排':[{'source_b_hash':'B','purpose':{}}],
            '保留原值':['A!C3'],'保留标签':['A!C2']}
        actual=self.plan(a,b)
        for field in ('updates','labels','identity_matches','retained_label_targets'):
            self.assertEqual(actual[field],expected[field])

    def test_维度原文引用不生成另格搬运(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲'},99)])
        a['facts'][0]['label_references']={'类别':['A!B3']}
        b['facts'][0]['label_references']={'类别':['B!C7']}
        a['机械记录']['sheets'][0]['cells']['B3']={'value':'原名称'}
        b['机械记录']['sheets'][0]['cells']['C7']={'value':'同一身份的新名称'}
        actual=self.plan(a,b)
        self.assertEqual(actual['labels'],[])
        self.assertEqual([(row['cell'],row['value']) for row in actual['updates']],[('C3',99)])

    def test_标签不能因引用唯一先后顺序或文字相同获得对应(self):
        for texts in [('新甲','新乙'),('原标签','新乙'),('相同文字','相同文字')]:
            a=self.mapping('A',[('C3','余额',{'类别':'甲'},10),
                                ('C4','余额',{'类别':'乙'},20),('C5','余额',{'类别':'丙'},30)])
            b=self.mapping('B',[('D7','余额',{'类别':'甲'},99),
                                ('D8','余额',{'类别':'乙'},88),('D9','余额',{'类别':'丙'},77)])
            for fact,ref in zip(a['facts'],['A!B2','A!B2','A!B5']):
                fact['label_references']={'类别':[ref]}
            for fact,ref in zip(b['facts'],['B!C7','B!C8','B!C9']):
                fact['label_references']={'类别':[ref]}
            a['机械记录']['sheets'][0]['cells'].update({'B2':{'value':'原标签'},'B5':{'value':'旧丙'}})
            b['机械记录']['sheets'][0]['cells'].update({
                'C7':{'value':texts[0]},'C8':{'value':texts[1]},'C9':{'value':'新丙'}})
            for reverse in (False,True):
                with self.subTest(texts=texts,reverse=reverse):
                    if reverse:a['facts'].reverse()
                    plan=self.plan(a,b)
                    self.assertEqual({r['cell']:r['value'] for r in plan['updates']},{'C3':99,'C4':88,'C5':77})
                    self.assertEqual(plan['labels'],[])
            b['facts'][0]['label_references']={'类别':['B!C7','B!C8']}
            self.assertFalse(any(r['cell']=='B2' for r in self.plan(a,b)['labels']))
            b['facts'][0]['label_references']={'类别':['B!C7']}
            b['facts'][1]['label_references']={'类别':['B!C7']}
            plan=self.plan(a,b)
            self.assertFalse(any(r.get('action')=='label' for r in plan['issues']))
            self.assertEqual(plan['labels'],[])

    def test_文字格自身唯一同一身份证才搬运原文(self):
        a=self.mapping('A',[('B3','文本指标',{'类别':'甲'},'原名称')])
        b=self.mapping('B',[('C7','文本指标',{'类别':'甲'},'新名称')])
        plan=self.plan(a,b)
        self.assertEqual([(r['cell'],r['value'],r['sources'][0]['cell']) for r in plan['updates']],
                         [('B3','新名称','C7')])
        self.assertEqual(plan['labels'],[])
        b['facts'].append({**deepcopy(b['facts'][0]),'source_reference':{'location':'B!C8'}})
        b['机械记录']['sheets'][0]['cells']['C8']={'value':'新名称'}
        self.assertEqual(self.plan(a,b)['updates'],[])

    def test_分别保存的期间角色同样参与唯一身份且不能用同日期跨角色(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲','period':'2028-12-31'},10),
                            ('D3','余额',{'类别':'甲','period':'2027-12-31'},20)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲','period':'2027-12-31'},90),
                            ('E7','余额',{'类别':'甲','period':'2026-12-31'},80)])
        a['outline']['subtables']=[{'sheet':'A','range':'C3:D3','metrics':[],
            'dimensions':[{'id':'template_period_role','ranges':[cell],'value':role}
                          for cell,role in [('C3','closing'),('D3','opening')]]}]
        b['outline']['subtables']=[{'sheet':'B','range':'D7:E7','metrics':[
            {'id':'余额','ranges':[cell],'template_period_role':role}
            for cell,role in [('D7','closing'),('E7','opening')]]}]
        plan=self.plan(a,b)
        self.assertEqual({row['cell']:row['sources'][0]['cell'] for row in plan['updates']},
                         {'C3':'D7','D3':'E7'})
        a['outline']['subtables']=[]
        self.assertEqual(self.plan(a,b)['updates'],[])

    def test_期间角色只比较本格保存原值(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲','period':'2024'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲','period':'2025'},99)])
        for mapping,cell in ((a,'C3'),(b,'D7')):
            mapping['outline']['subtables']=[{'sheet':mapping['carrier'],'range':'A1:H10',
                'metrics':[{'id':'余额','ranges':[cell],'template_period_role':'closing'}]}]
        for value in ('期末','期末余额',{'C':'closing'},{'C3':'closing'}):
            with self.subTest(value=value):
                a['outline']['subtables'][0]['metrics'][0]['template_period_role']=value
                self.assertEqual(self.plan(a,b)['identity_matches'],[])
        a['outline']['subtables'][0]['metrics'][0]['template_period_role']='closing'
        for ranges in (['*'],['C3:C4'],['C']):
            with self.subTest(ranges=ranges):
                a['outline']['subtables'][0]['metrics'][0]['ranges']=ranges
                self.assertEqual(self.plan(a,b)['identity_matches'],[])
        a['outline']['subtables'][0]['metrics'][0]['ranges']=['C3']
        self.assertEqual([r['value'] for r in self.plan(a,b)['updates']],[99])

    def test_程序不改写已保存的维度取值再对应(self):
        for left,right in [('甲',' 甲'),(1,1.0),
                           ({'mode':'members','members':['甲','乙']},{'mode':'members','members':['乙','甲']}),
                           ({'mode':'all','completeness':'complete'},{'mode':'all'})]:
            with self.subTest(left=left,right=right):
                a=self.mapping('A',[('C3','余额',{'类别':left},10)])
                b=self.mapping('B',[('D7','余额',{'类别':right},99)])
                self.assertEqual(self.plan(a,b)['identity_matches'],[])

    def test_定向范围不扩大到历史名单和范围外标签(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲'},99)])
        a['机械记录']['names']={'原名单':[{'sheet':'A','range':'A1:C5'}]}
        b['已确认关系']=[{
            'kind':'名单编排','source_a_hash':'A','source_b_hash':'B','gold_hash':'gold',
            'tables':[{'range_name':'原名单'}]},
            {**self.relation([]),'label_actions':[{'target':'A!A1','action':'set','value':'不得写入'}]}]
        purpose={'operations':[{'range_name':'原名单'}],
                 'labels':[{'target':'A!A2','source':'B!D7'}]}
        plan=self.plan(a,b,purpose,{'target':[{'sheet':'A','range':'C3'}],
                                 'source':[{'sheet':'B','range':'D7'}]})
        self.assertEqual(plan['operations'],[])
        self.assertEqual(plan['labels'],[])
        self.assertEqual([(row['cell'],row['value']) for row in plan['updates']],[('C3',99)])

    def test_名单安排不为相同身份证另加对应条件(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲'},99),('D8','余额',{'类别':'乙'},88)])
        a['机械记录']['names']={'原名单':[{'sheet':'A','range':'B2:C4'}]}
        purpose={'tables':[{'range_name':'原名单','first_row_offset':1,'old_detail_count':1,
            'source_sheet':'B','source_rows':[7],'columns':{'C':'D'},'labels':{},'selection_dimension':None}]}
        plan=self.plan(a,b,purpose)
        self.assertEqual([(row['cell'],row['sources'][0]['cell'],row['value']) for row in plan['updates']],
                         [('C3','D7',99)])

    def test_名单名称不能借金额身份证和标签引用绕过对应(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲'},99),('D8','余额',{'类别':'乙'},88)])
        a['facts'][0]['label_references']={'类别':['A!B3']}
        b['facts'][0]['label_references']={'类别':['B!C7']}
        a['机械记录']['sheets'][0]['cells']['B3']={'value':'原名称'}
        b['机械记录']['sheets'][0]['cells'].update({'C7':{'value':'同一身份的新名称'},'C8':{'value':'其他成员'}})
        a['机械记录']['names']={'原名单':[{'sheet':'A','range':'B2:C4'}]}
        purpose={'tables':[{'range_name':'原名单','first_row_offset':1,'old_detail_count':1,
            'source_sheet':'B','source_rows':[8],'columns':{'C':'D'},'labels':{'B':'C'},'selection_dimension':'类别'}],
            'labels':[{'target':'A!B3','source':'B!C8'}]}
        plan=self.plan(a,b,purpose)
        self.assertEqual([(row['cell'],row['value']) for row in plan['updates']],[('C3',99)])
        self.assertEqual(plan['labels'],[])

    def test_位置指定不能覆盖不同身份证(self):
        a = self.mapping('A', [('C3', '余额', {'类别': '甲'}, 10)])
        b = self.mapping('B', [('D7', '余额', {'类别': '乙'}, 10)])
        binding = {'target': 'A!C3', 'source': 'B!D7'}
        for purpose in (None, {'bindings': [binding]}):
            b['已确认关系'] = [self.relation([binding])]
            plan = self.plan(a, b, purpose)
            self.assertEqual(plan['updates'], [])
            self.assertEqual(plan['identity_matches'], [])

    def test_原值类型及名称列安排不改变身份证对应(self):
        purpose={'tables':[{'range_name':'原名单','first_row_offset':1,'old_detail_count':1,
            'source_rows':[7],'source_sheet':'B','columns':{'C':'D'},'labels':{'B':'C'}}]}
        for value in (0,-13.5,'新文字',None,'','  '):
            with self.subTest(value=value):
                a=self.mapping('A',[('B3','同一标准',{'类别':'甲'},'原文字')])
                b=self.mapping('B',[('C7','同一标准',{'类别':'甲'},value)])
                a['机械记录']['names']={'原名单':[{'sheet':'A','range':'B2:C3'}]}
                plain=self.plan(a,b)
                arranged=self.plan(a,b,purpose)
                expected=[{'sheet':'A','cell':'B3','metric_id':'同一标准','source_locations':['B!C7']}]
                self.assertEqual(plain['identity_matches'],expected)
                self.assertEqual(arranged['identity_matches'],expected)
                self.assertEqual(arranged['updates'],plain['updates'])
                self.assertEqual(arranged['retained_blanks'],plain['retained_blanks'])
                self.assertEqual(arranged['labels'],[])

    def test_错误位置安排不改变唯一身份证对应(self):
        a = self.mapping('A', [('C3', '余额', {'类别': '甲'}, 10)])
        b = self.mapping('B', [('D7', '余额', {'类别': '乙'}, 10), ('F9', '余额', {'类别': '甲'}, 99)])
        b['已确认关系'] = [self.relation([{'target': 'A!C3', 'source': 'B!D7'}])]
        plan = self.plan(a, b, {'labels':[{'target':'A!C3','source':'B!D7'}]})
        self.assertEqual([(row['value'], row['sources'][0]['cell']) for row in plan['updates']], [(99, 'F9')])

    def test_缺少身份证不能用位置搬运(self):
        a = self.mapping('A', [('C3', '余额', {'类别': '甲'}, 10)])
        b = self.mapping('B', [('D7', '余额', {'类别': '甲'}, 99)])
        b['已确认关系'] = [self.relation([{'target': 'A!C3', 'source': 'B!D7'}])]
        for side in ('A', 'B'):
            pair = deepcopy((a, b))
            pair[0 if side == 'A' else 1]['facts'] = []
            self.assertEqual(self.plan(*pair)['updates'], [])

    def test_相同身份证不另加维度填满条件(self):
        for dimensions in ({}, {'类别': {'mode':'unknown'}}):
            a = self.mapping('A', [('C3', '余额', dimensions, 10)])
            b = self.mapping('B', [('D7', '余额', dimensions, 99)])
            for mapping in (a,b):mapping['facts'][0]['missing_dimensions']=['类别']
            plan=self.plan(a,b)
            self.assertEqual([(row['value'],row['sources'][0]['cell']) for row in plan['updates']],[(99,'D7')])

    def test_未识别值不能作为通配符(self):
        a = self.mapping('A', [('C3', '余额', {'类别': {'mode':'unknown'}}, 10)])
        b = self.mapping('B', [('D7', '余额', {'类别':'甲'}, 99)])
        self.assertEqual(self.plan(a,b)['updates'],[])

    def test_同一身份证多来源不能靠位置指定选一个(self):
        a = self.mapping('A', [('C3', '余额', {'类别': '甲'}, 10)])
        b = self.mapping('B', [('D7', '余额', {'类别': '甲'}, 99), ('F9', '余额', {'类别': '甲'}, 99)])
        b['已确认关系'] = [self.relation([{'target': 'A!C3', 'source': 'B!D7'}])]
        plan = self.plan(a, b)
        self.assertEqual(plan['updates'], [])
        self.assertEqual(plan['identity_matches'], [])
        self.assertTrue(any(row.get('kind') == 'ambiguous_source' for row in plan['issues']))

    def test_同一身份证多个A位置也不能按位置分配来源(self):
        for dimensions in ({}, {'类别':'甲'}):
            a=self.mapping('A',[('C3','余额',dimensions,10),('C4','余额',dimensions,20)])
            b=self.mapping('B',[('D7','余额',dimensions,99)])
            plan=self.plan(a,b,{'bindings':[{'target':'A!C3','source':'B!D7'}]})
            self.assertEqual(plan['identity_matches'],[])
            self.assertEqual(plan['updates'],[])
            self.assertEqual({r['cell'] for r in plan['issues'] if r.get('kind')=='ambiguous_target'},{'C3','C4'})
            limited=self.plan(a,b,scope={'target':[{'sheet':'A','range':'C3'}],'source':[{'sheet':'B','range':'D7'}]})
            self.assertEqual([(r['cell'],r['value']) for r in limited['updates']],[('C3',99)])

    def test_A唯一性使用已明确的比较排除和期间角色(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲','report_scope':'合并','period':'2023'},10),
                            ('C4','余额',{'类别':'甲','report_scope':'母公司','period':'2024'},20)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲','report_scope':'合并','period':'2025'},99)])
        for mapping,cells in [(a,['C3','C4']),(b,['D7'])]:
            mapping['outline']['subtables']=[{'sheet':mapping['carrier'],'range':'A1:H10','metrics':[
                {'id':'余额','ranges':[cell],'template_period_role':'closing'} for cell in cells]}]
        requirements=[{'维度ID':'report_scope'}]
        self.assertEqual(self.plan(a,b,comparison_requirements=requirements)['updates'],[])
        a['outline']['subtables'][0]['metrics'][1]['template_period_role']='opening'
        self.assertEqual([(r['cell'],r['value']) for r in self.plan(a,b,comparison_requirements=requirements)['updates']],[('C3',99)])

    def test_额外同义关系不能把不同身份证变成相同(self):
        a = self.mapping('A', [('C3', '余额', {'类别': '甲'}, 10)])
        b = self.mapping('B', [('D7', '余额', {'类别': '乙'}, 99)])
        b['已确认关系'] = [{'dimension': '类别', 'record_ranges': ['B!D7'], 'value': '乙',
                         'A需要': {'source_hash': 'A', 'gold_hash': 'gold', 'value': '甲'}}]
        self.assertEqual(self.plan(a, b)['updates'], [])

    def test_一项未识别不影响另一项相同身份证(self):
        a = self.mapping('A', [('C3', '余额', {'类别': '甲'}, 10), ('C4', '余额', {'类别': '乙'}, 20)])
        b = self.mapping('B', [('D7', '余额', {'类别': '乙'}, 99)])
        plan = self.plan(a, b)
        self.assertEqual([(row['cell'], row['value']) for row in plan['updates']], [('C4', 99)])

    def test_名单行列安排不能供应金额身份证(self):
        a = self.mapping('A', [('C3', '余额', {'类别': '甲'}, 10)])
        b = self.mapping('B', [('D7', '余额', {'类别': '乙'}, 99)])
        a['机械记录']['names'] = {'名单': [{'sheet': 'A', 'range': 'B2:C3'}]}
        purpose = {'tables': [{'range_name': '名单', 'first_row_offset': 1,
            'old_detail_count': 1, 'source_rows': [7], 'source_sheet': 'B',
            'columns': {'C': 'D'}, 'labels': {}}]}
        self.assertEqual(self.plan(a, b, purpose)['updates'], [])

    def test_名单原值搬运不能绕过实际身份证对应(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲'},10),('C4','余额',{'类别':'乙'},20)])
        a['机械记录']['names']={'名单':[{'sheet':'A','range':'B2:E4'}]}
        a['机械记录']['sheets'][0]['cells'].update(E3={'value':31},E4={'value':42})
        table={'range_name':'名单','first_row_offset':1,'old_detail_count':2,
            'source_rows':[7,8],'source_sheet':'B','columns':{'C':'D'},'labels':{},
            'blank_values':{'7':{'E':{'source':'A!E4'}},'8':{'E':{'value':None}}}}
        b=self.mapping('B',[('D7','余额',{'类别':'乙'},99),('D8','余额',{'类别':'甲'},88)])
        plan=self.plan(a,b,{'tables':[table]})
        self.assertEqual([(r['cell'],r['value']) for r in plan['updates']],[('C3',88),('C4',99)])
        b['facts'][0]['dimensions']={'类别':'甲'}
        b['facts'][1]['dimensions']={'类别':'乙'}
        plan=self.plan(a,b,{'tables':[table]})
        self.assertEqual({r['cell']:r['value'] for r in plan['updates']},{'C3':99,'C4':88,'E3':42,'E4':None})

    def test_B空白不能借错误名单行搬运其他A原值(self):
        a=self.mapping('A',[('C3','余额',{'类别':'甲'},10),('C4','余额',{'类别':'乙'},20)])
        a['机械记录']['names']={'名单':[{'sheet':'A','range':'B2:C4'}]}
        b=self.mapping('B',[('D7','余额',{'类别':'乙'},None),('D8','余额',{'类别':'甲'},None)])
        table={'range_name':'名单','first_row_offset':1,'old_detail_count':2,
            'source_rows':[7,8],'source_sheet':'B','columns':{'C':'D'},'labels':{},
            'blank_values':{'7':{'C':{'source':'A!C4'}},'8':{'C':{'source':'A!C3'}}}}
        plan=self.plan(a,b,{'tables':[table]})
        self.assertEqual(plan['updates'],[])
        self.assertEqual({r['cell']:r['source_locations'] for r in plan['identity_matches']},
                         {'C3':['B!D8'],'C4':['B!D7']})
        b['facts'][0]['dimensions']={'类别':'甲'}
        b['facts'][1]['dimensions']={'类别':'乙'}
        plan=self.plan(a,b,{'tables':[table]})
        self.assertEqual({r['cell']:r['value'] for r in plan['updates']},{'C3':20,'C4':10})

    def test_已连通成员原值搬运不要求其他列填满(self):
        a=self.mapping('A',[('C3','余额',{},10)])
        a['机械记录']['names']={'名单':[{'sheet':'A','range':'B2:E3'}]}
        a['机械记录']['sheets'][0]['cells']['E3']={'value':31}
        b=self.mapping('B',[('D7','余额',{},99)])
        table={'range_name':'名单','first_row_offset':1,'old_detail_count':1,
            'source_rows':[7],'source_sheet':'B','columns':{'C':'D','D':'F'},'labels':{},
            'blank_values':{'7':{'E':{'source':'A!E3'}}}}
        plan=self.plan(a,b,{'tables':[table]})
        self.assertEqual({r['cell']:r['value'] for r in plan['updates']},{'C3':99,'E3':31})

    def test_相同身份证缺A期间角色只报告依赖缺口不误称维度不同(self):
        a = self.mapping('A', [('C3', '余额', {'类别': '甲', 'period': '2025-12-31'}, 10)])
        b = self.mapping('B', [('D7', '余额', {'类别': '甲', 'period': '2025-12-31'}, 99)])
        for mapping in (a, b):
            mapping['逐格字段版本'] = 1
            mapping['outline']['subtables'] = [{'sheet': mapping['carrier'], 'range': 'A1:H10',
                'metrics': [], 'dimensions': []}]
        role = {'id': 'template_period_role', 'ranges': ['D7'], 'value': 'current', '逐格原字段': True}
        b['outline']['subtables'][0]['dimensions'] = [role]
        plan = self.plan(a, b)
        self.assertEqual(plan['identity_matches'], [])
        self.assertEqual(plan['updates'], [])
        issue = next(row for row in plan['issues'] if row.get('cell') == 'C3')
        self.assertEqual(issue['kind'], 'missing_dependency')
        self.assertIn('A', issue['reason'])
        self.assertEqual(issue['candidates'][0]['different_dimensions'], [])
        self.assertEqual(issue['candidates'][0]['missing_dimensions'], ['template_period_role'])
        a['outline']['subtables'][0]['dimensions'] = [{**role, 'ranges': ['C3']}]
        self.assertEqual(self.plan(a, b)['updates'][0]['value'], 99)
        b['outline']['subtables'][0]['dimensions'] = []
        plan = self.plan(a, b)
        self.assertTrue(any(row.get('kind') == 'missing_dependency' for row in plan['issues']))
        self.assertEqual(plan['updates'], [])
        b['outline']['subtables'][0]['dimensions'] = [{**role, 'value': 'prior'}]
        self.assertEqual(self.plan(a, b)['updates'], [])
        a['outline']['subtables'][0]['dimensions'] = []
        b['outline']['subtables'][0]['dimensions'] = []
        self.assertEqual(self.plan(a, b)['updates'][0]['value'], 99)

    def test_缺已比较维度不是已有值不同且不改变对应条件(self):
        for missing in ('A','B'):
            with self.subTest(missing=missing):
                a=self.mapping('A',[('C3','余额',{'类别':'甲'},10)])
                b=self.mapping('B',[('D7','余额',{'类别':'甲'},99)])
                (b if missing=='A' else a)['facts'][0]['dimensions']['currency']='CNY'
                plan=self.plan(a,b)
                self.assertEqual(plan['identity_matches'],[])
                issue=next(x for x in plan['issues'] if x.get('cell')=='C3')
                self.assertEqual(issue['kind'],'missing_dependency')
                self.assertEqual(issue['candidates'][0]['different_dimensions'],[])
                self.assertEqual(issue['candidates'][0]['missing_dimensions'],['currency'])
        a=self.mapping('A',[('C3','余额',{'类别':'甲','currency':'人民币'},10)])
        b=self.mapping('B',[('D7','余额',{'类别':'甲','currency':'CNY'},99)])
        issue=next(x for x in self.plan(a,b)['issues'] if x.get('cell')=='C3')
        self.assertEqual(issue['kind'],'semantic_difference')
        self.assertEqual(issue['candidates'][0]['different_dimensions'],['currency'])
        self.assertEqual(issue['candidates'][0]['missing_dimensions'],[])
        for mapping in (a,b):mapping['facts'][0]['dimensions'].pop('currency')
        self.assertEqual(self.plan(a,b)['updates'][0]['value'],99)
        b['facts'][0]['dimensions'].update(类别='乙',currency='CNY')
        issue=next(x for x in self.plan(a,b)['issues'] if x.get('cell')=='C3')
        self.assertEqual(issue['kind'],'semantic_difference')
        self.assertEqual(issue['candidates'][0]['different_dimensions'],['类别'])
        self.assertEqual(issue['candidates'][0]['missing_dimensions'],['currency'])

    def test_A缺角色时旧日期差异不能遮住角色缺口(self):
        periods = [('2025-12-31', '2024-12-31'),
            ({'kind': 'duration', 'start': '2025-01-01', 'end': '2025-12-31'},
             {'kind': 'duration', 'start': '2024-01-01', 'end': '2024-12-31'})]
        for original, source in periods:
            with self.subTest(period=original):
                a = self.mapping('A', [('C3', '余额', {'类别': '甲', 'period': original}, 10)])
                b = self.mapping('B', [('D7', '余额', {'类别': '甲', 'period': source}, 99)])
                for mapping in (a, b):
                    mapping['逐格字段版本'] = 1
                    mapping['outline']['subtables'] = [{'sheet': mapping['carrier'], 'range': 'A1:H10',
                        'metrics': [], 'dimensions': []}]
                role = {'id': 'template_period_role', 'ranges': ['D7'], 'value': 'current', '逐格原字段': True}
                b['outline']['subtables'][0]['dimensions'] = [role]
                plan = self.plan(a, b)
                issue = next(row for row in plan['issues'] if row.get('cell') == 'C3')
                self.assertEqual(issue['kind'], 'missing_dependency')
                self.assertEqual(issue['candidates'][0]['different_dimensions'], ['period'])
                self.assertEqual(issue['candidates'][0]['missing_dimensions'], ['template_period_role'])
                self.assertEqual(plan['identity_matches'], [])
                self.assertEqual(plan['updates'], [])
                a['outline']['subtables'][0]['dimensions'] = [{**role, 'ranges': ['C3']}]
                self.assertEqual(self.plan(a, b)['updates'][0]['value'], 99)
                b['outline']['subtables'][0]['dimensions'] = [{**role, 'value': 'prior'}]
                known = self.plan(a, b)
                self.assertEqual(known['updates'], [])
                self.assertEqual(next(row for row in known['issues'] if row.get('cell') == 'C3')['kind'], 'semantic_difference')
                a['outline']['subtables'][0]['dimensions'] = []
                b['outline']['subtables'][0]['dimensions'] = [role]
                b['facts'][0]['dimensions']['类别'] = '乙'
                other = self.plan(a, b)
                issue = next(row for row in other['issues'] if row.get('cell') == 'C3')
                self.assertEqual(issue['kind'], 'semantic_difference')
                self.assertEqual(set(issue['candidates'][0]['different_dimensions']), {'period', '类别'})
                self.assertEqual(issue['candidates'][0]['missing_dimensions'], ['template_period_role'])
                self.assertEqual(other['updates'], [])

    def test_已识别期间角色机械换期不按旧日期选择来源(self):
        a = self.mapping('A', [('C3', '余额', {'类别': '甲', 'period': '2024-12-31'}, 10)])
        b = self.mapping('B', [('D7', '余额', {'类别': '甲', 'period': '2024-12-31'}, 11),
                               ('F9', '余额', {'类别': '甲', 'period': '2025-12-31'}, 99)])
        for mapping, metrics in [(a, [('C3', 'closing')]), (b, [('D7', 'opening'), ('F9', 'closing')])]:
            mapping['outline']['subtables'] = [{'sheet': mapping['carrier'], 'range': 'A1:H10',
                'metrics': [{'id': '余额', 'ranges': [cell], 'template_period_role': role} for cell, role in metrics]}]
        self.assertEqual(self.plan(a, b)['updates'][0]['value'], 99)
        for record in [a['facts'][0], b['facts'][1]]:
            record['dimensions'].pop('period')
            record['missing_dimensions'] = ['period']
        self.assertEqual(self.plan(a, b)['updates'][0]['value'], 99)

    def test_字段为准是引用措辞不能生成维度值(self):
        from 附注更新.A合集 import direct_answer
        standard = {'metric': {'余额': {'id': '余额', 'dimension_refs': [{'id': 'selection'}]}},
                    'dimension': {'selection': {'id': 'selection', 'value_type': 'selection', 'domain': '测试范围'}}}
        for sheet, cell in [('原表', 'D8'), ('另一工作表', 'G21')]:
            reply = f'{sheet}!{cell}：selection=all。\n每格的范围以各项所列的selection为准。其他说明保留。'
            rows = direct_answer(reply, standard, [sheet])['dimensions']
            self.assertEqual([row['value'] for row in rows], ['all'])


if __name__ == '__main__':
    unittest.main()
