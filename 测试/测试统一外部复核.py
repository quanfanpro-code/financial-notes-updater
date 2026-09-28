# coding: utf-8
"""通过现有界面调用的入口导入v2复核意见；受控资料不代表模型准确率。"""
import copy
import json
import unittest
import uuid
from pathlib import Path
from openpyxl import Workbook
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一语义 import read_standard_v2, definition_hash
from 附注更新.统一勾稽 import load_fact_mapping, build_fact_update_plan
from 附注更新.表格 import file_hash
from 附注更新.外部复核 import export_review_template, preview_external_review, apply_external_review
from 附注更新.界面 import format_external_review_preview
from 附注更新.分步流程 import run_step
from 测试.测试统一识别 import answer
from 测试.测试统一语义 import records


class ExternalReviewTests(unittest.TestCase):
    def test_导入待决格后可勾稽及更正且不能冒改原值(self):
        root=Path(__file__).resolve().parents[1]/'测试结果'/('新版外部复核_'+uuid.uuid4().hex);root.mkdir()
        gold=root/'标准.jsonl';gold.write_text('\n'.join(json.dumps(r,ensure_ascii=False) for r in records()),encoding='utf-8')
        source=root/'人工值.xlsx';book=Workbook();book.active['A1']='合成企业甲，单户人民币元，2025年12月31日固定资产原值合计'
        book.active['B2']='人工保留原文';book.save(source);book.close()
        class PendingEngine(UnifiedSemanticEngine):
            def _request(self,system,payload,validator):
                return validator({'cells':{a:({'kind':'unresolved','reason':'留待实际复核','evidence_cells':['A1']} if a=='B2' else
                    {'kind':'non_business','category':'title' if a=='A1' else 'empty_padding','reason':'合成原文','evidence_cells':['A1']}) for a in payload['candidate_cells']}})
        original=PendingEngine({},gold,root/'识别').recognize_workbook(source)
        original.update(carrier='A',original_path=str(source),original_hash=file_hash(source))
        mapping=root/'原映射.json';mapping.write_text(json.dumps(original,ensure_ascii=False),encoding='utf-8-sig')
        frozen={str(p):file_hash(p) for p in (source,gold,mapping)}
        template=root/'意见格式.json';export_review_template(mapping,template,gold_path=gold,include_confirmed=True)
        opinion=json.loads(template.read_text(encoding='utf-8-sig'))
        self.assertEqual(opinion['schema'],'附注外部语义复核-v2')
        item=copy.deepcopy(answer()['cells']['D7']);item.pop('kind');item.update(sheet='Sheet',cell='B2')
        opinion.update(reviewer={'name':'本地行为测试审核者','method':'human','reviewed_at':'2026-09-21T15:00:00+08:00'},items=[item])
        opinion_path=root/'实际意见.json';opinion_path.write_text(json.dumps(opinion,ensure_ascii=False),encoding='utf-8-sig')
        preview=preview_external_review(mapping,opinion_path,gold_path=gold)
        self.assertIn('固定资产原值',format_external_review_preview(preview))
        self.assertNotIn('实际口径：',format_external_review_preview(preview))
        saved=apply_external_review(preview,root,confirmed=True,confirm_method='测试明确采纳')
        loaded=load_fact_mapping(saved['files']['a_mapping'],read_standard_v2(gold))
        self.assertTrue(loaded['complete']);self.assertEqual(loaded['facts'][0]['raw_value'],'人工保留原文')
        self.assertEqual(loaded['excluded'],original['excluded'])
        plan=build_fact_update_plan(loaded,loaded,read_standard_v2(gold))
        self.assertEqual(len(plan['updates']),1)
        config={'gold_path':str(gold),'a_mapping':saved['files']['a_mapping'],'b_mapping':saved['files']['a_mapping'],'output_dir':str(root/'正式流程')}
        matched=run_step('match',config)
        written=run_step('write',{**config,'update_plan':matched['files']['update_plan']})
        derived=load_fact_mapping(written['files']['a_prime_mapping'],read_standard_v2(gold))
        self.assertEqual(derived['facts'][0]['raw_value'],'人工保留原文')
        self.assertNotIn('external_review_v2',derived)
        self.assertIn('derivation',derived)
        self.assertTrue(all(file_hash(p)==h for p,h in frozen.items()))
        # 已确认格只能明确绑定原事实后更正；未绑定及无原文证据的意见均不能采纳。
        for suffix,change in [('原值注入',{'raw_value':99}),('数值作证据',{'dimension_evidence':{k:['B2'] for k in item['dimensions']}})]:
            bad=copy.deepcopy(opinion);bad['items'][0].update(change)
            path=root/(suffix+'.json');path.write_text(json.dumps(bad,ensure_ascii=False),encoding='utf-8-sig')
            with self.assertRaises(ValueError):preview_external_review(mapping,path,gold_path=gold)
        correction=copy.deepcopy(opinion);correction['mapping_sha256']=file_hash(saved['files']['a_mapping'])
        path=root/'未绑定更正.json';path.write_text(json.dumps(correction,ensure_ascii=False),encoding='utf-8-sig')
        with self.assertRaises(ValueError):preview_external_review(saved['files']['a_mapping'],path,gold_path=gold)
        correction['items'][0]['replace_fact_hash']=definition_hash(loaded['facts'][0])
        path=root/'绑定更正.json';path.write_text(json.dumps(correction,ensure_ascii=False),encoding='utf-8-sig')
        second=apply_external_review(preview_external_review(saved['files']['a_mapping'],path,gold_path=gold),root,confirmed=True,confirm_method='再次明确采纳')
        self.assertEqual(load_fact_mapping(second['files']['a_mapping'],read_standard_v2(gold))['facts'][0]['raw_value'],'人工保留原文')
        tampered=json.loads(Path(second['files']['a_mapping']).read_text(encoding='utf-8-sig'));tampered['facts'][0]['dimensions']['period']['date']='2024-12-31'
        path=root/'伪改语义.json';path.write_text(json.dumps(tampered,ensure_ascii=False),encoding='utf-8-sig')
        with self.assertRaises(ValueError):load_fact_mapping(path,read_standard_v2(gold))


if __name__=='__main__':unittest.main()
