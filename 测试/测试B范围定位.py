# coding: utf-8
"""仅检查B定位的实际输入、范围去向、补读和保存回读，不调用外部模型。"""
import copy
import json
import unittest
from pathlib import Path
from unittest.mock import patch

from 测试 import 测试识别流程分层 as 分层测试
from 附注更新.统一识别 import UnifiedSemanticEngine


class B范围定位测试(unittest.TestCase):
    def test_共用阅读保留定位预览说明但不扩大数据范围(self):
        from 附注更新.A语义阅读 import semantic_read_input, validate_scoped_reading
        from 测试.测试引用来源语义 import 原格
        snapshot={'path':'另一材料.xlsx','sha256':'原件哈希','semantic_formula_sources':4,'sheets':[
            {'name':'明细','max_row':30,'max_column':3,'merges':[], 'cells':{
                'A1':原格(1,1,'说明：本期期末为2023年12月31日'),
                'A20':原格(20,1,'实际业务'), 'A21':原格(21,1,'项目'),
                'B21':原格(21,2,'期末余额'), 'A22':原格(22,1,'分类甲'),
                'B22':原格(22,2,12), 'A30':原格(30,1,'尚未阅读区域')}}]}
        # 预览取前八个内容行；补足正文行，确保表尾并未预读。
        snapshot['sheets'][0]['cells'].update({f'A{r}':原格(r,1,'正文') for r in range(2,10)})
        scope={'candidate_locations':['明细!B22'],'targets':[], 'details':[{'subtable_id':'原表','structure':'期末分类'}]}
        outline={'subtables':[{'id':'原表','ranges':[{'sheet':'明细','range':'A20:C22'}]}]}
        payload=semantic_read_input(snapshot,[],scope,outline)
        self.assertIn('本期期末为2023年12月31日',payload['source']['sheets'][0]['rows'])
        self.assertEqual(payload['candidate_locations'],['明细!B22'])
        self.assertEqual(payload['reading_ranges'],[{'sheet':'明细','range':'A20:C22','context_ranges':['A1:C8']}])
        self.assertIn('不并入或扩大',payload['range_rule'])
        table={'id':'业务','sheet':'明细','range':'A20:C22','subject':'业务','meaning':'分类余额',
               'data_ranges':['B22'],'context_ranges':['A1']}
        validate_scoped_reading({'tables':[table]},snapshot,scope,outline)
        scope['candidate_locations'].append('明细!A22')
        semantic_read_input(snapshot,[],scope,outline)
        validate_scoped_reading({'tables':[table]},snapshot,scope,outline)
        from 附注更新.A语义阅读 import table_packet, reading_questions
        self.assertIn('A22',table_packet(table,snapshot,[])['candidate_cells'])
        self.assertNotIn('A22',table['data_ranges'])
        scope['candidate_locations'].append('明细!A30')
        with self.assertRaisesRegex(ValueError,'遗漏候选=.*明细!A30'):
            validate_scoped_reading({'tables':[table]},snapshot,scope,outline)
        scope['candidate_locations'].remove('明细!A30')
        table['data_ranges'].append('B21')
        with self.assertRaisesRegex(ValueError,'多出数据=.*明细!B21'):
            validate_scoped_reading({'tables':[table]},snapshot,scope,outline)
        table['data_ranges'].remove('B21')
        table['data_ranges']=[]
        self.assertEqual(reading_questions({'tables':[table]},snapshot)[0]['数据范围外的数值'],['B22'])
        table['data_ranges']=['B22']
        scope['candidate_locations'].remove('明细!A22')
        table['context_ranges']=['A30']
        with self.assertRaisesRegex(ValueError,'公共说明超出已展示范围'):
            validate_scoped_reading({'tables':[table]},snapshot,scope,outline)
        table['context_ranges']=['A1'];table['range']='A1:C22'
        with self.assertRaisesRegex(ValueError,'B阅读超出已定位范围：明细!A1:C22'):
            validate_scoped_reading({'tables':[table]},snapshot,scope,outline)
        snapshot['semantic_formula_sources']=3
        legacy=semantic_read_input(snapshot,[],scope,outline)
        self.assertNotIn('reading_ranges',legacy)
        self.assertNotIn('本期期末为2023年12月31日',legacy['source']['sheets'][0]['rows'])

    def test_定位概要保留跨表槽位含义且不同成员不拆组(self):
        from 附注更新.B范围定位 import requirement_groups, location_payload
        from 附注更新.表格 import read_workbook
        from 测试.测试统一语义 import fact
        from 附注更新.A语义阅读 import metric_catalog
        first, second = fact(), fact()
        first['source_reference']['location'] = '甲表!B3'
        second['source_reference']['location'] = '乙表!C7'
        for row in (first, second):row['id'] = row['source_reference']['location']
        second['dimensions']['asset_selection'] = {
            'mode':'members', 'domain':'固定资产', 'members':['另一企业的设备'], 'completeness':'complete'}
        target = {'facts':[first, second]}
        standard = self.case.engine().standard
        groups = requirement_groups(target, standard, slot_correspondence=True)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0]['target_count'], 2)
        self.assertEqual(groups[0]['standard_slots'], metric_catalog(standard))
        payload = location_payload(read_workbook(self.case.source), [], groups)
        self.assertIn('槽位候选规则', payload)
        self.assertIn('跨表', payload['槽位候选规则'])
        self.assertNotIn('另一企业的设备', json.dumps(payload, ensure_ascii=False))
        self.assertEqual(len(requirement_groups(target, standard)), 2, '旧回执仍按旧分组复核')

    def setUp(self):
        self.case = 分层测试.分层识别测试()
        self.case.setUp()
        from openpyxl import load_workbook
        book = load_workbook(self.case.source)
        for row in range(20, 40): book['综合'][f'A{row}'] = '辅助说明' + str(row)
        self.case.source = self.case.folder / 'B含待补读内容.xlsx'
        book.save(self.case.source); book.close()

    def response(self, engine, system, payload):
        if payload['task'] == 'subtable_detail' and 'mechanical_comparison' in payload:
            answer = self.case.response(engine, system, payload)
            answer.pop('allowed_metrics')
            return answer
        if payload['task'].startswith('a_semantic_'):
            self.case.requests.append(copy.deepcopy(payload))
            if payload['task']=='a_semantic_read':
                return {'tables':[{'id':'原值','sheet':'综合','range':'A1:C3','subject':'固定资产',
                    'meaning':'合成企业甲，期末原值与累计折旧分别列示，本次只需原值。','data_ranges':['B3']}], 'notes':'人民币元，2025年12月31日。'}
            if payload['task']=='a_semantic_select':
                return {'tables':[{'id':'原值','metric_ids':['固定资产原值'],'missing':''}]}
            if payload['task']=='a_semantic_review':
                return {'tables':[{'id':'原值','corrections':{},'issues':[]}]}
            self.assertEqual(payload['task'],'a_semantic_bind')
            from 测试.测试统一语义 import fact
            return {'tables':[{'id':'原值','metrics':[{'ranges':['B3'],'id':'固定资产原值'}],
                'dimensions':[{'ranges':['B3'],'id':key,'value':value,'evidence':['A1','A2','A3']}
                              for key,value in fact()['dimensions'].items()], 'unresolved':[]}]}
        if payload['task'] != 'b_locate':
            return self.case.response(engine, system, payload)
        self.case.requests.append(copy.deepcopy(payload))
        groups = [g['id'] for g in payload['requirements']]
        if not payload['read_ranges']:
            return {'areas': [], 'unresolved_groups': [],
                    'read_ranges': [{'sheet': '综合', 'range': 'A1:C39'}]}
        return {'areas': [{'id': '原值', 'subject': '固定资产', 'purpose': '固定资产原值合计',
                           'ranges': [{'sheet': '综合', 'range': 'A1:C3'}], 'evidence': ['综合!A2'],
                           'group_ids': groups, 'reason': '期末原值披露在该区域'}],
                'unresolved_groups': [], 'read_ranges': []}

    def test_候选行名必须共同分类复核并正式回读(self):
        def respond(engine,system,payload):
            answer=self.response(engine,system,payload)
            if payload['task']=='subtable_detail':
                for row in answer['targets']:row['candidates'].append('综合!A3')
            return answer
        with patch.object(UnifiedSemanticEngine,'_http',respond):
            target=self.case.target();self.case.requests.clear()
            engine=self.case.engine()
            result=engine.recognize_workbook(self.case.source,target_mapping=target)
        self.assertEqual(result['recognition_scope']['candidate_locations'],['综合!A3','综合!B3'])
        self.assertEqual(len(result['facts']),1)
        self.assertIn('A3',[r['cell'] for r in result['excluded']])
        self.assertTrue(result['recognition_scope']['coverage']['complete'])
        self.assertEqual(result['recognition_scope']['coverage']['confirmed_candidate_count'],2)
        self.assertIn('a_semantic_review',[p['task'] for p in self.case.requests])
        from 附注更新.统一勾稽 import load_fact_mapping
        self.assertEqual(load_fact_mapping(result['mapping_path'],engine.standard)['excluded'],result['excluded'])

    def test_身份核查中断不当成身份不同补查(self):
        from 附注更新.语义 import SemanticError
        def reply(engine,system,payload):
            if payload['task']=='a_semantic_review':raise SemanticError('本轮核查中断')
            if payload['task']=='b_candidate_followup':raise AssertionError('未核查不能补查')
            return self.response(engine,system,payload)
        with patch.object(UnifiedSemanticEngine,'_http',reply):
            target=self.case.target()
            engine=self.case.engine()
            result=engine.recognize_workbook(self.case.source,target_mapping=target)
        self.assertFalse(result['b_identity_coverage']['complete'])
        self.assertFalse(result['recognition_scope']['coverage']['complete'])
        from 附注更新.统一勾稽 import load_fact_mapping
        saved=load_fact_mapping(result['mapping_path'],engine.standard)
        self.assertEqual(saved['unresolved'],result['unresolved'])

    def test_普通入口按概要定位且保存可回读(self):
        from 附注更新.表格编排 import validate_saved_scope
        from 附注更新.表格 import read_workbook, materialize_candidate_blanks
        with patch.object(UnifiedSemanticEngine, '_http', lambda e, s, p: self.response(e, s, p)):
            target = self.case.target()
            self.case.requests.clear()
            engine = self.case.engine()
            result = engine.recognize_workbook(self.case.source, target_mapping=target)
        tasks = [p['task'] for p in self.case.requests]
        self.assertIn('a_semantic_review', tasks, 'B必须通过A已有的身份核查流程')
        self.assertNotIn('workbook_outline', tasks, '有A时不再先登记整B所有子表')
        self.assertNotIn('select_subtables', tasks, '定位后不重复发送全部A身份选表')
        detail = next(p for p in self.case.requests if p['task'] == 'subtable_detail')
        self.assertIn('mechanical_comparison', detail, '详看前必须先利用五样式机械比较')
        comparison = result['recognition_scope']['mechanical_comparison']
        self.assertTrue(Path(comparison['path']).is_file())
        self.assertEqual(result['source_identities'], [])
        self.assertEqual(len(result['facts']), 1, '识别成功须准确对应金标准')
        self.assertTrue(result['recognition_scope']['coverage']['complete'])
        self.assertTrue(result['recognition_scope']['coverage']['satisfied'])
        locating = [p for p in self.case.requests if p['task'] == 'b_locate']
        self.assertEqual(len(locating), 2)
        self.assertNotIn('targets', locating[0])
        self.assertEqual({s['sheet'] for s in locating[0]['sheets']}, {'综合', '续页'})
        self.assertEqual(result['recognition_scope']['candidate_locations'], ['综合!B3'])
        self.assertEqual(result['recognition_scope']['selection']['subtables'][0]['target_ids'], [target['facts'][0]['id']])
        snapshot = read_workbook(self.case.source); materialize_candidate_blanks(snapshot)
        saved = json.loads(Path(result['mapping_path']).read_text(encoding='utf-8-sig'))
        validate_saved_scope(saved, snapshot, engine.standard, target)
        from 附注更新.统一勾稽 import load_fact_mapping
        from 附注更新.A语义阅读 import verify_mapping
        self.assertEqual(load_fact_mapping(result['mapping_path'],engine.standard)['source_identities'],result['source_identities'])
        changed=copy.deepcopy(saved)
        changed['semantic_reading']['batches'][0]['review']=[]
        with self.assertRaisesRegex(ValueError, '保存结果与原语义回执不符'):
            verify_mapping(changed,engine.standard)
        changed = copy.deepcopy(target); changed['facts'][0]['dimensions']['期间'] = '改变需求'
        with self.assertRaises(ValueError):
            validate_saved_scope(saved, snapshot, engine.standard, changed)
        changed = copy.deepcopy(saved)
        changed['recognition_scope']['mechanical_comparison']['sha256'] = '0' * 64
        with self.assertRaisesRegex(ValueError, '机械比较文件发生变化'):
            validate_saved_scope(changed, snapshot, engine.standard, target)
        changed['recognition_scope'].pop('mechanical_comparison')
        with self.assertRaisesRegex(ValueError, '缺少定位后的机械比较'):
            validate_saved_scope(changed, snapshot, engine.standard, target)
        from 附注更新.A语义阅读 import validate_scoped_reading
        table={'id':'原值','sheet':'综合','range':'A1:C3','subject':'固定资产','meaning':'原值','data_ranges':['B3'],
               'context_ranges':['A20']}
        outline=copy.deepcopy(result['outline'])
        outline['subtables'].append({'id':'尚未详看','ranges':[{'sheet':'综合','range':'A20:C39'}]})
        with self.assertRaisesRegex(ValueError, '公共说明超出已展示范围'):
            validate_scoped_reading({'tables':[table]},snapshot,result['recognition_scope'],outline)

    def test_概要合并但不丢无标准目标及拒绝遗漏越界(self):
        from 附注更新.B范围定位 import requirement_groups, validate_location, location_payload
        from 附注更新.表格编排 import target_requirements, check_evidence
        from 附注更新.表格 import read_workbook
        with patch.object(UnifiedSemanticEngine, '_http', lambda e, s, p: self.response(e, s, p)):
            target = self.case.target()
        target['facts'] = []
        target['outline'] = {'subtables': [{'id': 'T1', 'subject': '固定资产', 'meaning': '原值期末金额及备注'}]}
        target['source_identities'] = [{'table_id': 'T1', 'source_reference': {'location': '综合!' + a},
                                       'meaning': meaning, 'closest_metric_id': metric}
                                      for a, meaning, metric in [('B3', '原值金额', '固定资产原值'), ('C3', '空白备注', None)]]
        targets = target_requirements(target, self.case.engine().standard)
        groups = requirement_groups(target, self.case.engine().standard)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0]['target_ids'], [t['id'] for t in targets])
        self.assertEqual(groups[0]['target_count'], 2)
        self.assertIn('备注', groups[0]['purpose'])
        snapshot = read_workbook(self.case.source)
        payload = location_payload(snapshot, [], groups)
        answer = self.response(None, '', {**payload, 'read_ranges': [{'sheet': '综合', 'range': 'A1:C11'}]})
        checked = validate_location(answer, snapshot, groups, payload['shown_ranges'])
        self.assertEqual(checked['areas'][0]['group_ids'], [groups[0]['id']])
        snapshot['sheets'][0]['cells']['A2'].update(value='=其他表!A2', formula='=其他表!A2', cached_value='固定资产原值合计')
        cached = location_payload(snapshot, [], groups)
        self.assertIn('固定资产原值合计', json.dumps(cached, ensure_ascii=False))
        validate_location(answer, snapshot, groups, cached['shown_ranges'])
        check_evidence(['综合!A2'], snapshot)
        missing = copy.deepcopy(answer); missing['areas'][0]['group_ids'] = []
        with self.assertRaises(ValueError): validate_location(missing, snapshot, groups, payload['shown_ranges'])
        outside = copy.deepcopy(answer); outside['areas'][0]['ranges'][0]['range'] = 'A1:ZZ999'
        with self.assertRaises(ValueError): validate_location(outside, snapshot, groups, payload['shown_ranges'])

    def test_没有候选仍保存全部需要且不启动身份请求(self):
        from 附注更新.统一勾稽 import load_fact_mapping
        def reply(engine,system,payload):
            if payload['task']=='b_candidate_followup':
                return {'read_ranges':[], 'ranges':[{'sheet':'综合','range':'A1:C3'}],
                        'structure':'无对应来源的合成情形','evidence':['综合!A2'],
                        'targets':[{'id':row['id'],'candidates':[], 'reason':'合成无来源情形已补查所需区域'} for row in payload['targets']]}
            answer=self.response(engine,system,payload)
            if payload['task']=='subtable_detail' and 'mechanical_comparison' in payload:
                for row in answer['targets']:
                    row.update(candidates=[],reason='本例原文没有对应候选，保留A需要。')
            return answer
        with patch.object(UnifiedSemanticEngine,'_http',reply):
            target=self.case.target()
            self.case.requests.clear()
            engine=self.case.engine()
            result=engine.recognize_workbook(self.case.source,target_mapping=target)
        self.assertFalse(any(p['task'].startswith('a_semantic_') for p in self.case.requests))
        self.assertEqual(result['source_identities'],[])
        self.assertEqual(len(result['recognition_scope']['coverage']['targets']),1)
        self.assertFalse(result['recognition_scope']['coverage']['satisfied'])
        self.assertEqual(load_fact_mapping(result['mapping_path'],engine.standard)['source_identities'],[])


if __name__ == '__main__':
    unittest.main()
