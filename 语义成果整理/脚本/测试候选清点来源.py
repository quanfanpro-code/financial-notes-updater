"""候选清单必须携带实际工作簿来源，供后续真实建库。"""
import json
from pathlib import Path
import tempfile
import unittest
from 清点资产 import 清点候选,sha256


class 候选来源(unittest.TestCase):
    def test_保留五样式工作簿及原版本(self):
        root=Path(tempfile.mkdtemp(prefix='清点来源核验_'));old=root/'旧成果';source=root/'来源'
        def write(rel,value):
            p=root/rel;p.parent.mkdir(parents=True,exist_ok=True)
            p.write_text(value,encoding='utf-8-sig');return p
        def entry(p):return {'路径':str(p),'sha256':sha256(p)}
        fixed={k:entry(write(k+'.json','{}')) for k in ('v1','v2当前版本','A端映射','A端外部复核记录')}
        assets=[];styles={}
        for i in range(1,6):
            p=write(f'来源/附注样式示例/附注样式{i}.xlsx','合成只测清单')
            assets.append({**entry(p),'角色':'原始工作簿'});styles[str(i)]={'workbook_sha256':sha256(p)}
        write('旧成果/对应/维度规范对应.json','{}');write('旧成果/对应/表现身份迁移.jsonl','')
        write('旧成果/数据/表现记录.jsonl','')
        write('旧成果/资产清单.json',json.dumps({'固定输入':fixed,'资产':assets,'统计口径':{'五样式':{}}}))
        index=write('候选索引.json',json.dumps({'files':{},'styles':styles,'examples':[],'legacy_definitions':{}}))
        proof=write('纠错.jsonl','')
        out=清点候选(source,index,root/'清单.json',proof,old,fixed['v2当前版本']['路径'])
        actual=[e for e in out['资产'] if e['角色']=='原始工作簿']
        self.assertEqual(len(actual),5)
        self.assertEqual({e['sha256'] for e in actual},{e['sha256'] for e in assets})
        self.assertEqual(out['固定输入']['v2当前版本'],fixed['v2当前版本'])


if __name__=='__main__':unittest.main(verbosity=2)
