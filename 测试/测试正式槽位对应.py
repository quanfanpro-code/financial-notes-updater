# coding: utf-8
"""正式对应入口只比较标准槽位，具体取值与取数分开。"""
import copy
import json
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import load_workbook
from 测试.测试统一语义 import records
from 测试 import 测试统一勾稽 as 勾稽夹具
from 附注更新 import 统一勾稽 as api
from 附注更新.统一语义 import validate_records


class 正式槽位对应测试(unittest.TestCase):
    def test_已发证的B槽位不再被候选遗漏遮蔽且旧版本可回读(self):
        from 附注更新.表格编排 import target_requirements,required_coverage
        standard=validate_records(records());standard['sha256']='a'*64
        a=勾稽夹具.mapping('甲表!D8');b=勾稽夹具.mapping('异构!J3')
        b.update(excluded=[],outline={'complete':True},recognition_scope={
            'identity_coverage_version':3,'targets':target_requirements(a,standard,slot_correspondence=True),
            'details':[{'subtable_id':'子表一','targets':[{'id':'甲表!D8','candidates':[]}]}],
            'candidate_locations':[],'selection':{'subtables':[]}})
        result=required_coverage(b,standard)
        self.assertTrue(result['satisfied']);self.assertEqual(result['targets'][0]['standard_sources'],['异构!J3'])
        b['recognition_scope']['identity_coverage_version']=2
        self.assertFalse(required_coverage(b,standard)['satisfied'])

    def test_槽位不受姓名金额旧候选遗漏影响且不同计量不混配(self):
        rows=records(); other=copy.deepcopy(next(r for r in rows if r['record_type']=='metric'))
        other['id']='累计折旧';other['definition']['measurement_basis']='累计折旧';rows.append(other)
        standard=validate_records(rows);standard['sha256']='a'*64
        a=勾稽夹具.mapping('甲表!D8');b=勾稽夹具.mapping('异构!J3',None)
        b['facts'][0]['dimensions']['asset_selection']={'domain':'固定资产','mode':'members','members':['另一成员'],'completeness':'complete'}
        b['owners']['对象']=copy.deepcopy(b['facts'][0]['dimensions'])
        additional=copy.deepcopy(b['facts'][0]);additional['id']='异构!J4';additional['source_reference']['location']='异构!J4'
        additional['raw_value_state']='value';additional['raw_value']=999;b['facts'].append(additional)
        different=copy.deepcopy(additional);different['id']='异构!K4';different['source_reference']['location']='异构!K4';different['metric_id']='累计折旧';b['facts'].append(different)
        b['recognition_scope']={'coverage':{'targets':[{'id':'甲表!D8','candidates':[]}]}}
        saved=copy.deepcopy((a,b));plan=api.compare_semantic_slots(a,b,standard)
        self.assertTrue(plan['complete']);self.assertEqual(plan['matched_targets'],1)
        self.assertEqual([r['cell'] for r in plan['source_groups']['固定资产原值']['cells']],['J3','J4'])
        self.assertEqual(plan['targets'][0]['source_group'],'固定资产原值')
        self.assertEqual((a,b),saved);self.assertNotIn('updates',plan)
        b['facts']=b['facts'][-1:];missing=api.compare_semantic_slots(a,b,standard)
        self.assertFalse(missing['complete']);self.assertIsNone(missing['targets'][0]['source_group'])
        known=copy.deepcopy(a['facts'][0]);known['dimensions'].pop('asset_selection');known['dimension_evidence'].pop('asset_selection')
        b.update(slot_binding_version=1,semantic_reading={'synthetic':True},slot_bindings=[{'sheet':'异构','cell':'J8','partial_semantics':{
            **{k:known[k] for k in ('metric_id','dimensions','dimension_evidence')},
            'missing_dimensions':['asset_selection'],'evidence_cells':['A2']}}])
        partial=api.compare_semantic_slots(a,b,standard)
        self.assertTrue(partial['complete']);self.assertEqual(partial['source_groups']['固定资产原值']['cells'][0]['cell'],'J8')
        self.assertNotIn('asset_selection',partial['source_groups']['固定资产原值']['cells'][0]['dimensions'])
        b['coverage']={'expected_cells':['异构!K4','异构!J8'],'complete':True}
        scoped=api.select_mapping_scope(b,[{'sheet':'异构','range':'J8:J8'}])
        self.assertEqual(scoped['scope_coverage']['financial_cells'],1)
        self.assertEqual(scoped['facts'],[])
        self.assertEqual(len(scoped['slot_bindings']),1)
        broken=copy.deepcopy(b);broken['slot_bindings'][0]['partial_semantics']['metric_id']='不存在'
        with self.assertRaises(ValueError):api.compare_semantic_slots(a,broken,standard)
        a['source_identities']=[{'source_reference':{'location':'甲表!E8'},'meaning':'旧原文尚未发证'}]
        unbound=api.compare_semantic_slots(a,b,standard)
        self.assertFalse(unbound['complete']);self.assertEqual(unbound['issues'][0]['kind'],'unbound_identity')

    def test_旧候选记录可查看但写入不重启旧流程(self):
        case=勾稽夹具.MatchTests();case.setUp();root,gold,paths=case.file_fixture()
        for carrier,path in zip(('A','B'),paths):
            record=case.complete_fixture_mapping(path);record['carrier']=carrier
            path.write_text(json.dumps(record,ensure_ascii=False),encoding='utf-8-sig')
        from 附注更新.分步流程 import run_step
        config={'output_dir':str(root/'槽位入口'),'gold_path':str(gold),'a_mapping':str(paths[0]),'b_mapping':str(paths[1])}
        with patch.object(api,'calculate_mapping_formulas',side_effect=AssertionError('不重新计算原表')):
            result=run_step('match',config)
        self.assertEqual(result['status'],'complete',result['message'])
        plan=json.loads(Path(result['files']['update_plan']).read_text(encoding='utf-8-sig'))
        self.assertEqual(plan['schema'],'附注更新清单-v2');self.assertEqual(plan['updates'][0]['value'],900)
        with patch('附注更新.分步流程._prepare_update_v2',side_effect=AssertionError('写入不重做匹配')):
            written=run_step('write',{**config,**result['files']})
        self.assertEqual(written['status'],'complete',written['message'])
        book=load_workbook(written['files']['a_prime']);self.assertEqual(book['附注']['D8'].value,900);book.close()
        from 附注更新.流程 import save_json
        from 附注更新.表格 import file_hash
        from 附注更新.统一语义 import read_standard_v2
        standard=read_standard_v2(gold)
        historical=api.compare_semantic_slots(*(api.load_fact_mapping(path,standard) for path in paths),standard)
        historical.update({key:plan[key] for key in ('a_mapping','b_mapping','inputs','gold_path','a_link')})
        historical_path=Path(save_json(root/'历史候选记录'/'单元格语义对应.json',historical))
        save_json(historical_path.with_name('更新清单校验.json'),{'path':str(historical_path),'sha256':file_hash(historical_path)})
        from 附注更新.界面 import file_selection_updates
        self.assertEqual(file_selection_updates('update_plan',str(historical_path),{})['a_mapping'],str(paths[0]))
        with patch('附注更新.分步流程._prepare_update_v2',side_effect=AssertionError('写入不重启旧流程')), \
             patch('附注更新.分步流程.write_updated_workbook',side_effect=AssertionError('候选关系不是搬运动作')):
            written=run_step('write',{**config,'update_plan':str(historical_path)})
        self.assertEqual(written['status'],'failed',written['message'])
        self.assertIn('历史候选记录只供查看',written['message'])
        self.assertNotIn('a_prime',written['files'])


if __name__=='__main__':unittest.main()
