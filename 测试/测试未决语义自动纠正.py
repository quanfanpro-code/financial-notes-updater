# coding: utf-8
"""原文独立识别后仅纠正真实未决，不能重读健康表或无进展循环。"""
from collections import Counter
from copy import deepcopy
from pathlib import Path
import json
import unittest
import uuid
from unittest.mock import patch
from openpyxl import Workbook

from 附注更新.设置 import load_settings
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一勾稽 import load_fact_mapping
from 附注更新.表格 import file_hash


class 未决语义纠正测试(unittest.TestCase):
    def test_共同核查要求改变时复核未决表已有的财务位置(self):
        import 附注更新.A语义阅读 as reading
        engine,result,calls=self.run_case(False,healthy_same_table=True);sent=[]
        def respond(system,payload):
            if payload['task']=='a_semantic_select':
                return {'tables':[{'id':'甲','metric_ids':['metric.inventory.gross_balance'],'missing':''}]}
            sent.append(deepcopy(payload))
            return {'tables':[{'id':'甲','corrections':{},'issues':[{'range':'B2','reason':'仍待核实','evidence':['A2']}]}]}
        with patch.object(reading,'REVIEW_PROMPT',reading.REVIEW_PROMPT+'\n按适用的共同说明复核已有实际维度。'),patch.object(engine,'_http',side_effect=respond):
            current=engine.recognize_workbook(result['source_path'],previous_result=result['mapping_path'],
                artifact_metadata={'carrier':'A','original_path':result['source_path'],'original_hash':result['source_hash']})
        self.assertEqual(sent[0]['review_focus'],{'甲':['B2','B3']})
        self.assertEqual([u['cell'] for u in current['unresolved']],['B2'])
    def test_纠正未决时同时保留新诊断的具体范围(self):
        import 附注更新.A语义阅读 as reading
        engine,result,calls=self.run_case(False,preserve_label=True)
        def questions(binding,tables,snapshot,standard):
            return [{'id':'甲','ranges':['B3'],'check':'new_source_difference','reason':'新发现的原文疑点'}] if any(t['id']=='甲' for t in tables) else []
        def respond(system,payload):
            if payload['task']=='a_semantic_select':
                return {'tables':[{'id':'甲','metric_ids':['metric.inventory.gross_balance'],'missing':''}]}
            self.assertEqual(payload['review_focus'],{'甲':['B2','B3']})
            return {'tables':[{'id':'甲','corrections':{},'issues':[{'range':'B2','reason':'仍待核对计量','evidence':['A2']}]}]}
        with patch.object(reading,'SELECT_PROMPT',reading.SELECT_PROMPT+'\n新的检查要求。'), \
             patch.object(reading,'duplicate_identity_questions',side_effect=questions),patch.object(engine,'_http',side_effect=respond):
            current=engine.recognize_workbook(result['source_path'],previous_result=result['mapping_path'],
                artifact_metadata={'carrier':'A','original_path':result['source_path'],'original_hash':result['source_hash']})
        self.assertEqual([u['cell'] for u in current['unresolved']],['B2'])

    def test_接续新发现的疑点仅复核疑点范围不取旧局部补集(self):
        import 附注更新.A语义阅读 as reading
        engine,result,calls=self.run_case(True,preserve_label=True)
        def questions(binding,tables,snapshot,standard):
            return [{'id':'甲','ranges':['B2'],'check':'new_source_difference','reason':'新检查需要核对本格'}] if any(t['id']=='甲' for t in tables) else []
        def respond(system,payload):
            self.assertEqual(payload['task'],'a_semantic_review')
            self.assertEqual(payload['review_focus'],{'甲':['B2']})
            return {'tables':[{'id':'甲','corrections':{},'issues':[]}]}
        with patch.object(reading,'duplicate_identity_questions',side_effect=questions),patch.object(engine,'_http',side_effect=respond) as request:
            current=engine.recognize_workbook(result['source_path'],previous_result=result['mapping_path'],
                artifact_metadata={'carrier':'A','original_path':result['source_path'],'original_hash':result['source_hash']})
        self.assertEqual(request.call_count,1)
        self.assertTrue(current['complete'])
        load_fact_mapping(current['mapping_path'],engine.standard)

    def test_再次纠正时明确未决范围不能被旧核查范围覆盖(self):
        import 附注更新.A语义阅读 as reading
        engine,result,calls=self.run_case(False)
        def respond(system,payload):
            if payload['task']=='a_semantic_select':
                return {'tables':[{'id':'甲','metric_ids':['metric.inventory.gross_balance'],'missing':''}]}
            self.assertEqual(payload['task'],'a_semantic_review')
            self.assertEqual(payload['review_focus'],{'甲':['B2']})
            return {'tables':[{'id':'甲','corrections':{},'issues':[
                {'range':'B2','reason':'待继续核对计量含义','evidence':['A2']}]}]}
        with patch.object(reading,'SELECT_PROMPT',reading.SELECT_PROMPT+'\n新的候选召回要求。'),patch.object(engine,'_http',side_effect=respond):
            current=engine.recognize_workbook(result['source_path'],previous_result=result['mapping_path'],
                artifact_metadata={'carrier':'A','original_path':result['source_path'],'original_hash':result['source_hash']})
        self.assertFalse(current['complete'])
        self.assertEqual([(u['sheet'],u['cell']) for u in current['unresolved']],[('资料','B2')])

    def test_实际指令变更才允许重查尚存未决(self):
        import 附注更新.A语义阅读 as reading
        engine,result,calls=self.run_case(False)
        self.assertEqual(reading.correction_scope(result),{})
        with patch.object(reading,'SELECT_PROMPT',reading.SELECT_PROMPT+'\n改进候选召回规则。'):
            self.assertEqual(reading.correction_scope(result),{'甲':['B2']})

    def test_格式未完整的核查不能触发扩大请求(self):
        from 附注更新.A语义阅读 import correction_scope,_stage_receipt
        engine,result,calls=self.run_case(False)
        candidate=deepcopy(result);candidate['a_semantic_reading'].pop('correction_pass')
        def incomplete(ref,mapping,task):
            request,answer=_stage_receipt(ref,mapping,task)
            if task=='a_semantic_review':answer={**answer,'tables':answer['tables']+['{']}
            return request,answer
        with patch('附注更新.A语义阅读._stage_receipt',side_effect=incomplete):
            self.assertEqual(correction_scope(candidate),{})

    def run_case(self,resolve,interrupted=False,legacy=False,invalid_partial=False,preserve_label=False,healthy_same_table=False):
        import 附注更新.A语义阅读 as reading_module
        initial_prompt=reading_module.BIND_PROMPT
        directory=Path(__file__).resolve().parents[1]/'测试结果'/('未决纠正_'+uuid.uuid4().hex)
        directory.mkdir();source=directory/'不同编排.xlsx'
        book=Workbook();sheet=book.active;sheet.title='资料'
        for a,v in {'A1':'预付款项','B1':'期末','A2':'供应商甲','A5':'预付款项','B5':'期末','A6':'供应商乙'}.items():sheet[a]=v
        for a in ('B2','B6'):sheet[a].number_format='0.00'
        if healthy_same_table:sheet['A3']='供应商丙';sheet['B3'].number_format='0.00'
        book.save(source);book.close()
        gold=load_settings()['gold_path'];before=(file_hash(source),file_hash(gold))
        engine=UnifiedSemanticEngine({'base_url':'codex://local','model':'gpt-6-luna','model_reasoning_effort':'high'},gold,directory,lambda _:None,None)
        tables=[{'id':key,'sheet':'资料','range':span,'subject':subject,'meaning':'预付款项余额','data_ranges':[cell]}
                for key,span,subject,cell in [('甲','A1:B2','科目一','B2'),('乙','A5:B6','科目二','B6')]]
        if preserve_label or healthy_same_table:
            tables[0].update(range='A1:B3',data_ranges=['B2:B3'])
        calls=[]
        def respond(system,payload):
            task=payload['task'];calls.append(deepcopy(payload))
            if task=='a_semantic_read':return {'tables':tables,'notes':'具体主体和日期未提供，保留已知槽位'}
            if task=='a_semantic_select':
                retry=bool(payload.get('selection_questions'))
                return {'tables':[{'id':t['id'],'metric_ids':[
                    'metric.prepayment.gross_balance' if t['id']=='乙' or retry and resolve else 'metric.inventory.gross_balance'],
                    'missing':''} for t in payload['reading']['tables']]}
            key=payload['reading'][0]['id'];cell='B2' if key=='甲' else 'B6'
            if task=='a_semantic_bind':
                correct=key=='乙' or any(x['id']=='metric.prepayment.gross_balance' for x in payload['standards']['metrics'])
                metrics=([{'ranges':[cell],'id':'metric.prepayment.gross_balance'}] if correct else [])+(
                    [{'ranges':['B3'],'id':'metric.prepayment.gross_balance'}] if healthy_same_table and key=='甲' else [])
                return {'tables':[{'id':key,'metrics':metrics,
                    'dimensions':[],'unresolved':[] if correct else [{'range':cell,'reason':'候选对象不符，须重新检索原文所表达的指标'}]}]}
            self.assertEqual(task,'a_semantic_review')
            if interrupted and key=='乙':reading_module.BIND_PROMPT=initial_prompt+'\n新版本格式要求，不改变业务语义。'
            bad=not payload['recognized']['tables'][0]['metrics']
            corrections={}
            if bad and resolve and any(x['id']=='metric.prepayment.gross_balance' for x in payload['standards']['metrics']):
                corrections={'metrics':[{'ranges':[cell],'id':'metric.prepayment.gross_balance'}],'resolved_ranges':[cell]}
                bad=False
            if preserve_label and key=='甲':
                if 'review_focus' in payload:self.assertEqual(payload['review_focus'],{'甲':['B2']})
                else:corrections['non_business']=[{'range':'B3','category':'label','reason':'共同阅读确认是分节标签位置','evidence':['A1']}]
            return {'tables':[{'id':key,'corrections':corrections,'issues':[
                {'range':cell,'reason':'当前候选没有表达原文对象','evidence':['A2' if key=='甲' else 'A6']}] if bad else []}]}
        metadata={'carrier':'A','original_path':str(source),'original_hash':before[0]};saved=[]
        def progress(record):
            if record['a_semantic_reading'].get('correction_pass') and len(record['a_semantic_reading']['batches'])==1:
                saved.append(record['mapping_path'])
                from 附注更新.语义 import SemanticCancelled
                raise SemanticCancelled('模拟纠正途中中断')
        with patch.object(engine,'_http',side_effect=respond),patch.object(reading_module,'BIND_PROMPT',initial_prompt):
            if interrupted:
                from 附注更新.语义 import SemanticCancelled
                with self.assertRaises(SemanticCancelled):engine.recognize_workbook(source,progress=progress,artifact_metadata=metadata)
                path=Path(saved[-1])
                if legacy:
                    raw=json.loads(path.read_text(encoding='utf-8-sig'));raw['a_semantic_reading'].pop('correction_source',None)
                    path=directory/'旧版中断进度.json';path.write_text(json.dumps(raw,ensure_ascii=False),encoding='utf-8-sig')
                if invalid_partial:
                    raw=json.loads(path.read_text(encoding='utf-8-sig'))
                    raw['slot_bindings'][0]['partial_semantics']['metric_id']='metric.inventory.gross_balance'
                    path=directory/'无法正式回读的中断进度.json';path.write_text(json.dumps(raw,ensure_ascii=False),encoding='utf-8-sig')
                    with self.assertRaises(ValueError):load_fact_mapping(path,engine.standard)
                result=engine.recognize_workbook(source,previous_result=str(path),artifact_metadata=metadata)
            else:result=engine.recognize_workbook(source,artifact_metadata=metadata)
        self.assertEqual((file_hash(source),file_hash(gold)),before)
        loaded=load_fact_mapping(result['mapping_path'],engine.standard)
        self.assertEqual(loaded['unresolved'],result['unresolved'])
        return engine,result,calls

    def test_漏选只纠正问题表且回读两格确认槽位(self):
        engine,result,calls=self.run_case(True)
        self.assertTrue(result['complete'])
        self.assertEqual(len(result['slot_bindings']),2)
        bindings=Counter(x['reading'][0]['id'] for x in calls if x['task']=='a_semantic_bind')
        self.assertEqual(bindings,{'甲':1,'乙':1})
        corrections=[x for x in calls if x['task']=='a_semantic_select' and x.get('selection_questions')]
        self.assertEqual(len(corrections),1)
        self.assertEqual([t['id'] for t in corrections[0]['reading']['tables']],['甲'])
        self.assertTrue(result['a_semantic_reading']['correction_pass'])

    def test_增加候选后沿用已核查绑定及非业务结论(self):
        engine,result,calls=self.run_case(True,preserve_label=True)
        self.assertTrue(result['complete'])
        self.assertEqual(sum(x['task']=='a_semantic_bind' and x['reading'][0]['id']=='甲' for x in calls),1)
        self.assertTrue(any(x['cell']=='B3' and x['category']=='label' for x in result['excluded']))

    def test_无法解决只纠正一轮并保留未决(self):
        engine,result,calls=self.run_case(False)
        self.assertFalse(result['complete'])
        self.assertEqual([(x['sheet'],x['cell']) for x in result['unresolved']],[('资料','B2')])
        self.assertEqual(sum(x['task']=='a_semantic_select' for x in calls),2)
        self.assertEqual(sum(x['task']=='a_semantic_review' and x['reading'][0]['id']=='甲' for x in calls),2)
        with patch.object(engine,'_http',side_effect=AssertionError('相同问题不能无进展重复调用模型')):
            carried=engine.recognize_workbook(result['source_path'],previous_result=result['mapping_path'],
                artifact_metadata={'carrier':'A','original_path':result['source_path'],'original_hash':result['source_hash']})
        self.assertFalse(carried['complete'])
        self.assertEqual(carried['unresolved'],result['unresolved'])

    def test_纠正中断后从完整起点和原回执恢复健康表(self):
        engine,result,calls=self.run_case(True,interrupted=True)
        self.assertTrue(result['complete'])
        self.assertEqual(Counter(x['reading'][0]['id'] for x in calls if x['task']=='a_semantic_bind'),{'甲':1,'乙':1})
        self.assertTrue(result['a_semantic_reading']['correction_source']['sha256'])

    def test_旧中断进度缺起点字段时只从本作业缓存恢复(self):
        engine,result,calls=self.run_case(True,interrupted=True,legacy=True)
        self.assertTrue(result['complete'])
        self.assertEqual(sum(x['task']=='a_semantic_bind' and x['reading'][0]['id']=='乙' for x in calls),1)

    def test_不可回读的中间结果不采纳并从完整起点重建(self):
        engine,result,calls=self.run_case(True,interrupted=True,invalid_partial=True)
        self.assertTrue(result['complete'])
        self.assertEqual({x['partial_semantics']['metric_id'] for x in result['slot_bindings']},{'metric.prepayment.gross_balance'})
        self.assertEqual(sum(x['task']=='a_semantic_bind' and x['reading'][0]['id']=='乙' for x in calls),1)
