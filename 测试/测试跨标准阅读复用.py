# coding: utf-8
"""原文阅读可跨标准复用，身份绑定仍严格绑定实际标准版本。"""
from copy import deepcopy
from pathlib import Path
import json
import uuid
import unittest
from unittest.mock import patch


class 跨标准阅读复用测试(unittest.TestCase):
    def test_跨标准中断可回到旧标准完整起点且不重读健康表(self):
        from 测试.测试未决语义自动纠正 import 未决语义纠正测试
        from 附注更新.统一识别 import UnifiedSemanticEngine
        from 附注更新.统一勾稽 import load_fact_mapping
        from 附注更新.语义 import SemanticCancelled
        from 附注更新.表格编排 import _stage_receipt
        old,original,_=未决语义纠正测试().run_case(False)
        folder=Path(old.work_dir)/('跨标准_'+uuid.uuid4().hex);folder.mkdir()
        rows=[json.loads(line) for line in Path(old.gold_path).read_text(encoding='utf-8-sig').splitlines() if line.strip()]
        extra=deepcopy(old.standard['metric']['metric.prepayment.gross_balance'])
        extra['id']='metric.test.other_measure';extra['definition']['measurement_basis']='仅用于受控检查的不同计量基础'
        gold=folder/'新增定义标准.jsonl'
        gold.write_text('\n'.join(json.dumps(r,ensure_ascii=False) for r in rows+[extra]),encoding='utf-8')
        engine=UnifiedSemanticEngine(old.settings,gold,folder,lambda _:None,None);saved=[];calls=[]
        def respond(system,payload):
            task=payload['task'];calls.append(task)
            if task=='a_semantic_select':
                return {'tables':[{'id':t['id'],'metric_ids':['metric.prepayment.gross_balance'],'missing':''} for t in payload['reading']['tables']]}
            if task=='a_semantic_review':return {'tables':[{'id':'甲','corrections':{
                'metrics':[{'id':'metric.prepayment.gross_balance','ranges':['B2']}],'resolved_ranges':['B2']},'issues':[]}]}
            raise AssertionError('原文未变，不应重新阅读：'+task)
        def interrupted(system,payload):
            if payload['task']=='a_semantic_review':raise SemanticCancelled('模拟标准更新后核查中断')
            return respond(system,payload)
        metadata={'carrier':'A','original_path':original['source_path'],'original_hash':original['source_hash']}
        with patch.object(engine,'_http',side_effect=interrupted),self.assertRaises(SemanticCancelled):
            engine.recognize_workbook(original['source_path'],previous_result=original['mapping_path'],
                                      artifact_metadata=metadata,progress=lambda m:saved.append(m['mapping_path']))
        with patch.object(engine,'_http',side_effect=respond):
            result=engine.recognize_workbook(original['source_path'],previous_result=saved[-1],artifact_metadata=metadata)
        self.assertTrue(result['complete']);load_fact_mapping(result['mapping_path'],engine.standard)
        self.assertNotIn('a_semantic_read',calls)
        ref=result['a_semantic_reading']['reading'][0]
        with self.assertRaises(ValueError):_stage_receipt(ref,{**result,'source_hash':'0'*64},'a_semantic_read')
        bound=next(b for b in result['a_semantic_reading']['batches'] if not b.get('reused_from'))['binding'][0]
        with self.assertRaises(ValueError):_stage_receipt(bound,{**result,'gold_hash':'0'*64},'a_semantic_bind')
        from 附注更新.统一语义 import read_standard_v2
        changed=deepcopy(rows+[extra])
        next(r for r in changed if r['id']=='period')['meaning']+='；已修订的不同完整定义'
        revised=folder/'同名定义修订.jsonl';revised.write_text('\n'.join(json.dumps(r,ensure_ascii=False) for r in changed),encoding='utf-8')
        revised_standard=read_standard_v2(revised)
        # 空候选绑定没有使用此维度；用真实已提供完整定义的旧复核核验定义变化仍拒绝。
        reference=original['a_semantic_reading']['batches'][1]['review'][0]
        with self.assertRaises(ValueError):_stage_receipt(reference,{**result,'gold_path':str(revised),
            'gold_hash':revised_standard['sha256']},'a_semantic_review')
