"""用真实旧定义检验跨表式归一；实际主体日期及金额是明确的合成测试输入。"""
import json
from pathlib import Path
import unittest
from 附注更新.设置 import load_settings
from 附注更新.表格 import file_hash
from 附注更新.统一语义 import read_standard_v2, semantic_key
from 附注更新.统一迁移 import migrate_fact


class 货币资金原版衔接测试(unittest.TestCase):
    def test_同义旧槽位归一且期初期末不能混同(self):
        root = Path(__file__).resolve().parents[1]
        old_path = root/'金标准/财务报表附注语义金标准-v1.jsonl'
        old = {x['id']:x for x in map(json.loads,old_path.read_text(encoding='utf-8-sig').splitlines())}
        standard = read_standard_v2(load_settings()['gold_path'])
        context = {key:{'value':value,'evidence_refs':['合成测试：显式业务上下文']} for key,value in {
            'entity':'合成测试主体','report_scope':'consolidated','currency':'CNY','unit_scale':1,
            'period.closing':{'kind':'instant','date':'2025-12-31'},
            'period.opening':{'kind':'instant','date':'2024-12-31'},
            'denomination_currency_selection':{'domain':'原始计价币种','mode':'all','completeness':'complete'},
        }.items()}
        facts=[]
        for identifier in ['C-N087-T001-S003','C-N087-T001-S004','C-N087-T002-S017','C-N087-T002-S018']:
            original={'legacy_id':identifier,'legacy_mapping_hash':'c'*64,'dimensions':{},'dimension_evidence':{},
                      'source_reference':{'file_hash':'d'*64,'location':'合成测试!B2'},
                      'raw_value_state':'value','raw_value':'123.45',
                      'recognition_provenance':{'method':'human_review','record_id':'合成测试'}}
            output=migrate_fact(original,old[identifier],file_hash(old_path),standard,context)['fact']
            self.assertEqual(output['metric_id'],'metric.monetary_funds.balance')
            self.assertEqual(output['dimensions']['cash_selection']['members'],['银行存款'])
            self.assertEqual(output['raw_value'],'123.45')
            facts.append(output)
        self.assertEqual(semantic_key(facts[0],standard),semantic_key(facts[2],standard))
        self.assertEqual(semantic_key(facts[1],standard),semantic_key(facts[3],standard))
        self.assertNotEqual(semantic_key(facts[0],standard),semantic_key(facts[1],standard))


if __name__=='__main__': unittest.main()
