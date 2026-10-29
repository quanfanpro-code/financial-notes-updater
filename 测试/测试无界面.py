# coding: utf-8
"""检查正式无界面入口的编排；模拟步骤不充当真实业务验收。"""
import json
import unittest
import uuid
from pathlib import Path
from unittest.mock import patch


class 无界面测试(unittest.TestCase):
    def setUp(self):
        self.folder=Path(__file__).resolve().parents[1]/'测试结果'/('无界面_'+uuid.uuid4().hex)
        self.folder.mkdir()
        self.config={'output_dir':str(self.folder),'model':'gpt-6-luna','base_url':'codex://local','model_reasoning_effort':'high'}
        for key,name in [('word_path','原附注.docx'),('b_path','原试算.xlsx'),('gold_path','标准.jsonl')]:
            p=self.folder/name;p.write_bytes(b'original');self.config[key]=str(p)

    def test_原始材料依次经过相同五步且旧成果不进入作业(self):
        from 附注更新.无界面 import run_background
        calls=[]
        def step(name,config,log,cancel,progress):
            calls.append(name)
            if name=='extract_a':
                self.assertTrue(all(not config.get(k) for k in ['a_mapping','b_mapping','update_plan']))
            self.assertEqual(config['update_purpose'], {'reason':'日期与名单随B，原值搬运'})
            files={k:str(self.folder/(name+'_'+k+'.json')) for k in {'extract_a':['a_path','a_link'],
                'recognize_a':['a_mapping'],'recognize_b':['b_mapping'],'match':['update_plan'],'write':['word','a_prime']}[name]}
            for p in files.values():Path(p).write_text('{}',encoding='utf-8')
            return {'status':'complete','step':name,'files':files,'message':name,'usage':{'requests':1}}
        with patch('附注更新.无界面.run_step',side_effect=step):
            result=run_background({**self.config,'a_mapping':'旧A.json','b_mapping':'旧B.json','update_plan':'旧清单.json',
                                   'update_purpose':{'reason':'日期与名单随B，原值搬运'}})
        self.assertEqual(calls,['extract_a','recognize_a','recognize_b','match','write'])
        self.assertEqual(result['status'],'complete')
        self.assertEqual(result['usage']['requests'],5)
        self.assertTrue(Path(result['files']['word']).is_file())

    def test_识别失败停止且正常接续只沿用本作业正式进度(self):
        from 附注更新.无界面 import run_background
        calls=[]
        mapping=self.folder/'本次进度.json';mapping.write_text('{}',encoding='utf-8')
        def step(name,config,log,cancel,progress):
            calls.append(name)
            if name=='extract_a':return {'status':'complete','step':name,'files':{'a_path':self.config['b_path']},'message':'已提取'}
            progress({'files':{'a_mapping':str(mapping)},'checkpoint':True})
            return {'status':'failed','step':name,'files':{'a_mapping':str(mapping)},'message':'受控失败'}
        with patch('附注更新.无界面.run_step',side_effect=step):first=run_background(self.config)
        self.assertEqual(calls,['extract_a','recognize_a'])
        self.assertEqual(first['status'],'failed')
        calls.clear()
        def continued(name,config,log,cancel,progress):
            calls.append(name)
            if name=='recognize_a':self.assertEqual(config['a_mapping'],str(mapping))
            return {'status':'complete','step':name,'files':{},'message':'本例接续'}
        with patch('附注更新.无界面.run_step',side_effect=continued):
            result=run_background(self.config,resume=first['output_dir'])
        self.assertEqual(calls,['recognize_a','recognize_b','match','write'])
        self.assertEqual(result['status'],'complete')

    def test_部分识别成果继续传递且接续不重跑(self):
        from 附注更新.无界面 import run_background
        calls=[]
        def step(name,config,log,cancel,progress):
            calls.append(name)
            fields={'extract_a':['a_path'],'recognize_a':['a_mapping'],'recognize_b':['b_mapping'],
                    'match':['update_plan'],'write':['a_prime']}
            files={key:str(self.folder/(name+'_'+key+'.json')) for key in fields[name]}
            for path in files.values():Path(path).write_text('{}',encoding='utf-8')
            if name=='recognize_b':self.assertTrue(config.get('a_mapping'))
            if name=='match':self.assertTrue(config.get('a_mapping') and config.get('b_mapping'))
            return {'status':'complete' if name=='extract_a' else 'partial','files':files,'message':'模拟部分成果'}
        with patch('附注更新.无界面.run_step',side_effect=step):result=run_background(self.config)
        self.assertEqual(calls,['extract_a','recognize_a','recognize_b','match','write'])
        self.assertEqual(result['status'],'partial')
        with patch('附注更新.无界面.run_step',side_effect=AssertionError('不能重跑已有成果')):
            resumed=run_background(self.config,resume=result['output_dir'])
        self.assertEqual(resumed['status'],'partial')

    def test_原件变化拒绝接续且不启动任何步骤(self):
        from 附注更新.无界面 import run_background
        with patch('附注更新.无界面.run_step',return_value={'status':'failed','files':{},'message':'已停'}):
            first=run_background(self.config)
        Path(self.config['b_path']).write_bytes(b'changed')
        with patch('附注更新.无界面.run_step',side_effect=AssertionError('不能启动')):
            with self.assertRaisesRegex(ValueError,'原件|输入'):
                run_background(self.config,resume=first['output_dir'])

    def test_冻结标准传递到全部步骤并写入接续约束(self):
        from 附注更新.无界面 import run_background
        def failed(name,config,log,cancel,progress):
            self.assertTrue(config['freeze_standard'])
            return {'status':'failed','files':{},'message':'受控读取失败'}
        frozen={**self.config,'freeze_standard':True}
        with patch('附注更新.无界面.run_step',side_effect=failed):first=run_background(frozen)
        manifest=json.loads((Path(first['output_dir'])/'输入记录.json').read_text(encoding='utf-8-sig'))
        self.assertTrue(manifest['settings']['freeze_standard'])
        with patch('附注更新.无界面.run_step',side_effect=AssertionError('不能取消冻结')):
            with self.assertRaisesRegex(ValueError,'输入设置'):run_background(self.config,resume=first['output_dir'])
        calls=[]
        def complete(name,config,log,cancel,progress):
            self.assertTrue(config['freeze_standard']);calls.append(name)
            return {'status':'complete','files':{},'message':'本例完成'}
        with patch('附注更新.无界面.run_step',side_effect=complete):result=run_background(frozen,resume=first['output_dir'])
        self.assertEqual(result['status'],'complete')
        self.assertEqual(calls,['extract_a','recognize_a','recognize_b','match','write'])

    def test_正常命令行可以明确冻结标准(self):
        from 附注更新.无界面 import main
        with patch('附注更新.无界面.load_settings',return_value=self.config),patch('附注更新.无界面.run_background',return_value={'status':'partial'}) as run:
            self.assertEqual(main(['--无界面','--冻结标准']),2)
        self.assertTrue(run.call_args.args[0]['freeze_standard'])


if __name__=='__main__':unittest.main()
