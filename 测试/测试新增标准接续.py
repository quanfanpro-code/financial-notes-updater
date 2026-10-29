# coding: utf-8
"""新增标准只重识别原未决表；已有身份保留真实旧回执并可正式回读。"""
import copy
import json
import unittest
import uuid
from pathlib import Path
from openpyxl import Workbook
from 附注更新.统一识别 import UnifiedSemanticEngine
from 附注更新.统一勾稽 import load_fact_mapping
from 测试.测试统一语义 import records, fact


class 新增标准接续测试(unittest.TestCase):
    def test_新增标准后只绑定原缺项表且拒绝篡改复用依据(self):
        folder = Path(__file__).resolve().parents[1] / '测试输出' / ('新增标准接续_' + uuid.uuid4().hex)
        folder.mkdir()
        source = folder / '两表.xlsx'
        book = Workbook(); sheet = book.active; sheet.title = '表'
        for row in [('合成企业甲 人民币元 2025年12月31日', '期末'), ('固定资产原值', 125),
                    ('合成企业甲 人民币元 2025年12月31日', '期末'), ('固定资产减值准备', 2)]:
            sheet.append(row)
        book.save(source); book.close()
        old = folder / '原标准.jsonl'; new = folder / '补充后标准.jsonl'
        rows = records()
        old.write_text('\n'.join(json.dumps(r, ensure_ascii=False) for r in rows), encoding='utf-8-sig')
        extra = copy.deepcopy(rows[-1]); extra['id'] = '固定资产减值准备'
        extra['definition']['measurement_basis'] = '已确认减值准备余额'
        new.write_text('\n'.join(json.dumps(r, ensure_ascii=False) for r in rows + [extra]), encoding='utf-8-sig')
        tables = [{'id': key, 'sheet': '表', 'range': area, 'subject': subject, 'meaning': subject+'期末余额', 'data_ranges': [cell]}
                  for key, area, subject, cell in [('原值', 'A1:B2', '固定资产原值', 'B2'), ('准备', 'A3:B4', '固定资产减值准备', 'B4')]]
        calls = []
        def run(gold, previous=None):
            engine = UnifiedSemanticEngine({'model': '合成检查', 'base_url': 'http://unused.local'}, gold, folder / uuid.uuid4().hex)
            def reply(system, payload):
                task = payload['task']
                if task == 'a_semantic_read':
                    self.assertNotEqual(gold,new,'相同原件及阅读材料不能因更新标准而重新阅读并改写表号')
                    return {'tables': tables, 'notes': '两张表分别列示，无相互包含。'}
                if task == 'a_semantic_select':
                    return {'tables': [{'id': t['id'], 'metric_ids': ([t['subject']]+(['固定资产减值准备'] if t['id']=='原值' and gold==new else [])) if t['id']=='原值' or gold==new else [],
                                        'missing': '' if t['id']=='原值' or gold==new else '标准缺少减值准备余额'} for t in payload['reading']['tables']]}
                if task == 'a_semantic_review':
                    reviewed=[]
                    for t in payload['reading']:
                        corrections={}
                        if gold==new and t['id']=='准备' and not next(r for r in payload['recognized']['tables'] if r['id']==t['id'])['metrics']:
                            cell=t['data_ranges'][0]
                            corrections={'metrics':[{'ranges':[cell],'id':t['subject']}],'resolved_ranges':[cell],
                                         'patches':[{'ranges':[cell],'id':d,'value':v,'evidence':['A3','A4']} for d,v in fact()['dimensions'].items()]}
                        reviewed.append({'id':t['id'],'corrections':corrections,'issues':[]})
                    return {'tables':reviewed}
                self.assertEqual(task, 'a_semantic_bind')
                result = []
                for t in payload['reading']:
                    calls.append((gold.name, t['id']))
                    cell=t['data_ranges'][0]; known=t['id']=='原值' or gold==new
                    result.append({'id':t['id'], 'metrics':[{'ranges':[cell], 'id':t['subject']}] if known else [],
                                   'dimensions':[{'ranges':[cell], 'id':d, 'value':v, 'evidence':[t['range'].split(':')[0], 'A'+cell[1:]]}
                                                 for d,v in fact()['dimensions'].items()] if known else [],
                                   'unresolved':[] if known else [{'range':cell, 'reason':'标准缺少减值准备余额'}]})
                return {'tables': result}
            engine._http = reply
            result = engine.recognize_workbook(source, artifact_metadata={'carrier':'A'}, previous_result=previous)
            return result, engine.standard
        first, _ = run(old)
        second, standard = run(new, first['mapping_path'])
        self.assertEqual([t for g,t in calls if g==new.name], [])
        loaded = load_fact_mapping(second['mapping_path'], standard)
        self.assertEqual(len(loaded['facts']), 2)
        self.assertFalse(loaded['unresolved'])
        self.assertEqual(loaded['usage']['reused_confirmed_tables'], 1)
        from 附注更新.A语义阅读 import verify_mapping
        bad=copy.deepcopy(second)
        reused=next(b for b in bad['a_semantic_reading']['batches'] if b.get('reused_from'))
        reused['reused_from']['sha256']='0'*64
        with self.assertRaises(ValueError): verify_mapping(bad,standard)
        before=len(calls)
        third,standard=run(new,second['mapping_path'])
        self.assertEqual(len(calls),before)
        self.assertEqual(len(load_fact_mapping(third['mapping_path'],standard)['facts']),2)
        from 附注更新.分步流程 import run_step
        from unittest.mock import patch
        config={'a_mapping':third['mapping_path'],'b_mapping':second['mapping_path'],
                'gold_path':str(new),'output_dir':str(folder/'实际更新入口')}
        with patch.object(UnifiedSemanticEngine,'_http',side_effect=AssertionError('写回不得重新识别')):
            matched=run_step('match',config)
            self.assertEqual(matched['status'],'complete',matched)
            written=run_step('write',{**config,**matched['files']})
        self.assertEqual(written['status'],'complete',written)
        saved=json.loads(Path(written['files']['a_prime_mapping']).read_text(encoding='utf-8-sig'))
        self.assertNotIn('a_semantic_reading',saved)
        self.assertNotIn('facts',saved)
        self.assertEqual(saved['collections']['A']['mapping_path'],third['mapping_path'])
        self.assertEqual(saved['collections']['B']['mapping_path'],second['mapping_path'])


if __name__ == '__main__': unittest.main()
