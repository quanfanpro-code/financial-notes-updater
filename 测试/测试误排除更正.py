"""误排除可通过绑定原记录的正式复核纠正，值保持原样。"""
import copy
import json
import unittest
import uuid
from pathlib import Path
from openpyxl import Workbook
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一语义 import read_standard_v2, definition_hash
from 附注更新.统一勾稽 import load_fact_mapping
from 附注更新.表格 import file_hash
from 附注更新.外部复核 import export_review_template, preview_external_review, apply_external_review
from 测试.测试统一识别 import answer
from 测试.测试统一语义 import records


class ExclusionCorrectionTests(unittest.TestCase):
    def test_明确绑定后更正且原值与其他排除保留(self):
        root=Path(__file__).resolve().parents[1]/'测试结果'/('误排除更正_'+uuid.uuid4().hex)
        root.mkdir()
        def save(name,value):
            p=root/name;p.write_text(json.dumps(value,ensure_ascii=False),encoding='utf-8-sig');return p
        gold=root/'标准.jsonl';gold.write_text('\n'.join(json.dumps(r,ensure_ascii=False) for r in records()),encoding='utf-8')
        source=root/'原值.xlsx';book=Workbook();book.active['A1']='合成企业甲，单户人民币元，2025年12月31日固定资产原值合计'
        book.active['B2']='——';book.save(source);book.close()
        class ExcludedEngine(UnifiedSemanticEngine):
            def _request(self,system,payload,validator):
                return validator({'cells':{a:{'kind':'non_business','category':'annotation' if a=='B2' else 'title',
                    'reason':'模拟错误排除' if a=='B2' else '说明','evidence_cells':['A1']} for a in payload['candidate_cells']}})
        original=ExcludedEngine({},gold,root/'识别').recognize_workbook(source)
        original.update(carrier='B',original_path=str(source),original_hash=file_hash(source))
        mapping=save('原映射.json',original);source_hash=file_hash(source)
        template=root/'意见格式.json';export_review_template(mapping,template,gold_path=gold,include_confirmed=True)
        opinion=json.loads(template.read_text(encoding='utf-8-sig'))
        option=next(r for r in opinion['excluded_mapping_options'] if r['cell']=='B2')
        previous=next(r for r in original['excluded'] if r['cell']=='B2')
        self.assertEqual(option['replace_exclusion_hash'],definition_hash(previous))
        item=copy.deepcopy(answer()['cells']['D7']);item.pop('kind');item.update(sheet='Sheet',cell='B2')
        opinion.update(reviewer={'name':'本地行为验证','method':'human','reviewed_at':'2026-09-22T07:00:00+08:00'},items=[item])
        with self.assertRaises(ValueError):preview_external_review(mapping,save('无绑定.json',opinion),gold_path=gold)
        item['replace_exclusion_hash']='错误校验值'
        with self.assertRaises(ValueError):preview_external_review(mapping,save('错误绑定.json',opinion),gold_path=gold)
        item['replace_exclusion_hash']=definition_hash(previous)
        preview=preview_external_review(mapping,save('实际意见.json',opinion),gold_path=gold)
        saved=apply_external_review(preview,root,confirmed=True,confirm_method='测试明确采纳')
        loaded=load_fact_mapping(saved['files']['b_mapping'],read_standard_v2(gold))
        self.assertEqual(loaded['facts'][0]['raw_value'],'——')
        self.assertEqual(loaded['excluded'],[r for r in original['excluded'] if r['cell']!='B2'])
        self.assertEqual(file_hash(source),source_hash)
        self.assertTrue(loaded['complete'])


if __name__=='__main__':unittest.main()
