"""直接指标参照须经固定快照校验，未转换的旧槽位不影响读取。"""
import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from copy import deepcopy
from 测试统一语义 import records
from 附注更新.统一语义 import validate_records
from 附注更新.原版金标准衔接 import recognition_expressions


class 指标参照读取(unittest.TestCase):
    def test_无旧转换仍读取直接指标并拒绝定义漂移(self):
        root = Path(tempfile.mkdtemp(prefix='指标参照核验_'))
        snapshot = root/'语义金标准/当前指标标准.jsonl'
        snapshot.parent.mkdir()
        snapshot.write_text('\n'.join(json.dumps(r,ensure_ascii=False) for r in records()),encoding='utf-8-sig')
        digest = hashlib.sha256(snapshot.read_bytes()).hexdigest()
        example = {'style':1,'metric_id':'固定资产原值','standard_sha256':digest,
            'dimensions':{'asset_selection':{'domain':'固定资产','mode':'members',
                'members':['甲设备','乙设备'],'completeness':'complete'}},
            'dimension_evidence':{'asset_selection':['原表!A3']},
            'missing_dimensions':['entity','period','scope'], 'evidence_cells':['原表!C3'],
            'template_period_role':'closing','confirmation':{'依据':'原文完整类别'},
            'row_semantics':'甲设备及乙设备','column_semantics':'期末原值','table_semantics':'固定资产',
            'origin':{'mapping_id':'m1','sheet':'原表','cell':'C3'}}
        corpus = root/'语义示例索引.json'
        corpus.write_text(json.dumps({'source_gold_sha256':'a'*64,'metric_standard_sha256':digest,
            'legacy_definitions':{},'examples':[example]},ensure_ascii=False),encoding='utf-8-sig')
        standard = validate_records(records())
        # 必需缺项由当前定义推导，避免测试写错协议。
        example['missing_dimensions'] = sorted({r['id'] for r in standard['metric']['固定资产原值']['dimension_refs']
            if r['required']} - set(example['dimensions']))
        # 仅首次构造写入完成后才交读取器，试验文件保留。
        corpus2 = root/'完整示例.json'
        corpus2.write_text(json.dumps({'source_gold_sha256':'a'*64,'metric_standard_sha256':digest,
            'legacy_definitions':{},'examples':[example]},ensure_ascii=False),encoding='utf-8-sig')
        result = recognition_expressions(standard, corpus_path=corpus2)
        self.assertEqual(result['metric_templates'][0]['dimensions'],example['dimensions'])
        self.assertEqual(result['metric_templates'][0]['template_period_role'],'closing')
        changed = deepcopy(standard)
        changed['metric']['固定资产原值']['definition']['quantity']='不同的含义'
        with self.assertRaisesRegex(ValueError,'定义'):
            recognition_expressions(changed, corpus_path=corpus2)


if __name__=='__main__':unittest.main(verbosity=2)
