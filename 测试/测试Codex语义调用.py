# coding: utf-8
"""只替代外部进程回执；实际设置持久化、传输解析和语义入口仍运行。"""
import json
import subprocess
import uuid
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch

from 附注更新.设置 import load_settings, save_settings, validate_settings


@contextmanager
def test_directory():
    directory=Path(__file__).resolve().parents[1]/'测试结果'/('Codex调用_'+uuid.uuid4().hex)
    directory.mkdir()
    yield directory


class Codex调用测试(unittest.TestCase):
    def test_正常请求超时由原生流程终止进程并保存诊断(self):
        import collections
        from 附注更新.语义 import SemanticEngine,SemanticError
        engine=object.__new__(SemanticEngine)
        engine.settings={'base_url':'codex://local','model':'gpt-6-luna','timeout':5}
        engine.cancel=None;engine.log=lambda message:None;engine.usage=collections.Counter()
        engine.source_hash='a'*64;engine.cache_source_hash='';engine.gold_hash='b'*64;engine._review_receipts=[]
        class Process:
            active=True
            def communicate(self,input=None,timeout=None):
                if self.active:raise subprocess.TimeoutExpired('codex',timeout)
                return b'{"type":"thread.started"}\n',b'actual timeout diagnostic'
            def terminate(self):self.active=False
        process=Process()
        with test_directory() as directory:
            engine.work_dir=directory
            with patch('附注更新.语义._call_with_deadline',side_effect=AssertionError('不能抢先结束原生进程清理')),patch('附注更新.Codex语义调用.shutil.which',return_value=r'C:\本机\codex.exe'),patch('附注更新.Codex语义调用.subprocess.Popen',return_value=process),patch('附注更新.Codex语义调用.time.monotonic',side_effect=[0,0,6]):
                with self.assertRaisesRegex(SemanticError,'连接失败或超时'):
                    engine._request('按原文阅读',{'task':'a_semantic_read'},lambda x:x)
            self.assertFalse(process.active)
            self.assertEqual(next(directory.rglob('调用诊断.txt')).read_bytes(),b'actual timeout diagnostic')
            self.assertIn('thread.started',next(directory.rglob('事件.jsonl')).read_text(encoding='utf-8'))

    def test_公开准则查询保留结构约束并继续禁止本机工具(self):
        from 附注更新.Codex语义调用 import call_codex,CODEX_INSTRUCTIONS
        with test_directory() as root:
            def spawn(command,**options):
                self.assertIn('web_search="live"',command)
                self.assertIn('tools.web_search.allowed_domains=["mof.gov.cn"]',command)
                self.assertIn('shell_tool',command)
                self.assertIn('memories',command)
                self.assertIn('multi_agent',command)
                self.assertIn('只搜索抽象术语',CODEX_INSTRUCTIONS)
                self.assertEqual(json.loads(Path(command[command.index('--output-schema')+1]).read_text(encoding='utf-8')),{'type':'object'})
                Path(command[command.index('--output-last-message')+1]).write_text('{}',encoding='utf-8')
                class Process:
                    returncode=0
                    def communicate(self,**kwargs):
                        return b'{"type":"item.completed","item":{"type":"web_search","query":"abstract term"}}\n{"type":"turn.completed","usage":{}}\n',b''
                return Process()
            form={'type':'json_schema','json_schema':{'schema':{'type':'object'}}}
            with patch('附注更新.Codex语义调用.shutil.which',return_value=r'C:\本机\codex.exe'),patch('附注更新.Codex语义调用.subprocess.Popen',side_effect=spawn):
                data,_=call_codex({'model':'gpt-6-luna'},root,{'messages':[],'response_format':form},lambda:None)
            self.assertIn('web_search',next(root.rglob('事件.jsonl')).read_text(encoding='utf-8'))
            request=json.loads(next(root.rglob('请求.json')).read_text(encoding='utf-8'))
            self.assertEqual(request['allowed_domains'],['mof.gov.cn'])

    def test_公开规范查询策略变化不复用旧缓存(self):
        import collections
        from 附注更新.语义 import SemanticEngine
        engine=object.__new__(SemanticEngine)
        engine.settings={'base_url':'codex://local','model':'gpt-6-luna','model_reasoning_effort':'high'}
        engine.cancel=None;engine.log=lambda message:None;engine.usage=collections.Counter()
        engine.source_hash='a'*64;engine.cache_source_hash='';engine.gold_hash='b'*64;engine._review_receipts=[]
        replies=iter([{'判断':'原策略'},{'判断':'新策略'}]);engine._http=lambda system,payload:next(replies)
        with test_directory() as directory:
            engine.work_dir=directory
            first=engine._request('解释原文术语',{'task':'缓存隔离检查'},lambda x:x)
            with patch('附注更新.Codex语义调用.CODEX_INSTRUCTIONS','新公开规范查询策略'):
                second=engine._request('解释原文术语',{'task':'缓存隔离检查'},lambda x:x)
            again=engine._request('解释原文术语',{'task':'缓存隔离检查'},lambda x:x)
            self.assertEqual((first,second,again),({'判断':'原策略'},{'判断':'新策略'},first))
            self.assertEqual(engine.usage['cached_requests'],1)

    def test_深层证据目录不充当Windows子进程当前目录(self):
        from 附注更新.Codex语义调用 import call_codex
        with test_directory() as root:
            directory=root/('深层记录'*50)/'本次作业证据'
            self.assertGreater(len(str(directory)),260)
            directory.mkdir(parents=True)
            def spawn(command,**options):
                self.assertLess(len(str(options['cwd'])),260)
                self.assertLess(len(command[command.index('--cd')+1]),260)
                Path(command[command.index('--output-last-message')+1]).write_text('{}',encoding='utf-8')
                class Process:
                    returncode=0
                    def communicate(self,**kwargs):return b'{"type":"turn.completed","usage":{}}\n',b''
                return Process()
            with patch('附注更新.Codex语义调用.shutil.which',return_value=r'C:\本机\codex.exe'),patch('附注更新.Codex语义调用.subprocess.Popen',side_effect=spawn):
                call_codex({'model':'gpt-6-luna'},directory,{'messages':[]},lambda:None)

    def test_本机Codex设置保存恢复且推理强度校验(self):
        config={'base_url':'codex://local','model':'gpt-6-luna','model_reasoning_effort':'high','api_key':''}
        self.assertEqual(validate_settings(config)['model_reasoning_effort'],'high')
        with test_directory() as directory:
            path=Path(directory)/'用户配置.json'
            save_settings(config,path)
            saved=load_settings(path)
            self.assertEqual(saved['base_url'],'codex://local')
            self.assertEqual(saved['model_reasoning_effort'],'high')
        with self.assertRaises(ValueError):validate_settings({**config,'model_reasoning_effort':'随便'})

    def test_统一模型入口接收真实进程格式并保留中文及用量(self):
        from 附注更新.语义 import SemanticEngine
        import collections
        engine=object.__new__(SemanticEngine)
        engine.settings={'base_url':'codex://local','model':'gpt-6-luna','model_reasoning_effort':'high','timeout':5}
        engine.cancel=None;engine.log=lambda message:None;engine.usage=collections.Counter()
        with test_directory() as directory:
            engine.work_dir=Path(directory)
            def spawn(command,**options):
                self.assertEqual(command[command.index('--model')+1],'gpt-6-luna')
                self.assertIn('model_reasoning_effort="high"',command)
                self.assertNotIn('resume',command)
                class Process:
                    returncode=0
                    def communicate(self,input=None,timeout=None):
                        self_test.assertIn('根据原文判断',input.decode('utf-8'))
                        Path(command[command.index('--output-last-message')+1]).write_text('{"结论":"仅来自本次原文"}',encoding='utf-8')
                        return (json.dumps({'type':'turn.completed','usage':{'input_tokens':11,'output_tokens':3}}).encode(),b'')
                self_test=self
                return Process()
            with patch('附注更新.Codex语义调用.shutil.which',return_value=r'C:\本机\codex.exe'),patch('附注更新.Codex语义调用.subprocess.Popen',side_effect=spawn):
                result=engine._http('根据原文判断',{'task':'连接检查','text':'中文原文'})
            self.assertEqual(result,{'结论':'仅来自本次原文'})
            self.assertEqual(engine.usage['requests'],1)
            self.assertEqual(engine.usage['prompt_tokens'],11)
            self.assertEqual(engine.usage['completion_tokens'],3)

    def test_失败进程即使留下结果文件也不能当作成功回执(self):
        from 附注更新.Codex语义调用 import call_codex
        with test_directory() as directory:
            def spawn(command,**options):
                Path(command[command.index('--output-last-message')+1]).write_text('{"伪成功":true}',encoding='utf-8')
                class Process:
                    returncode=1
                    def communicate(self,**kwargs):
                        return b'{"type":"turn.failed","error":{"message":"model unavailable"}}\n',b''
                return Process()
            with patch('附注更新.Codex语义调用.shutil.which',return_value=r'C:\本机\codex.exe'),patch('附注更新.Codex语义调用.subprocess.Popen',side_effect=spawn):
                with self.assertRaisesRegex(RuntimeError,'model unavailable'):
                    call_codex({'model':'gpt-6-luna','timeout':5},Path(directory),{'messages':[]},lambda:None)

    def test_超时停止本次进程并留下诊断(self):
        from 附注更新.Codex语义调用 import call_codex
        class Process:
            active=True
            def communicate(self,input=None,timeout=None):
                if self.active:raise subprocess.TimeoutExpired('codex',timeout)
                return b'',b'partial diagnostic'
            def terminate(self):self.active=False
        process=Process()
        with test_directory() as directory:
            with patch('附注更新.Codex语义调用.shutil.which',return_value=r'C:\本机\codex.exe'),patch('附注更新.Codex语义调用.subprocess.Popen',return_value=process),patch('附注更新.Codex语义调用.time.monotonic',side_effect=[0,0,6]):
                with self.assertRaises(TimeoutError):
                    call_codex({'model':'gpt-6-luna','timeout':5},Path(directory),{'messages':[]},lambda:None)
            self.assertFalse(process.active)
            self.assertEqual(next(Path(directory).rglob('调用诊断.txt')).read_bytes(),b'partial diagnostic')

    def test_取消传播且停止正在调用的本次进程(self):
        from 附注更新.Codex语义调用 import call_codex
        from 附注更新.语义 import SemanticCancelled
        class Process:
            active=True
            def communicate(self,timeout=None):return b'',b''
            def terminate(self):self.active=False
        process=Process();checks=[]
        def cancel_after_start():
            checks.append(True)
            if len(checks)>1:raise SemanticCancelled('已取消')
        with test_directory() as directory:
            with patch('附注更新.Codex语义调用.shutil.which',return_value=r'C:\本机\codex.exe'),patch('附注更新.Codex语义调用.subprocess.Popen',return_value=process):
                with self.assertRaises(SemanticCancelled):
                    call_codex({'model':'gpt-6-luna'},Path(directory),{'messages':[]},cancel_after_start)
            self.assertFalse(process.active)

    def test_改变推理强度不复用另一强度的回执(self):
        import collections
        from 附注更新.语义 import SemanticEngine
        engine=object.__new__(SemanticEngine)
        engine.settings={'base_url':'codex://local','model':'gpt-6-luna','model_reasoning_effort':'high'}
        engine.cancel=None;engine.log=lambda message:None;engine.usage=collections.Counter()
        engine.source_hash='a'*64;engine.cache_source_hash='';engine.gold_hash='b'*64;engine._review_receipts=[]
        replies=iter([{'判断':'第一次结果'},{'判断':'第二次结果'}])
        engine._http=lambda system,payload:next(replies)
        with test_directory() as directory:
            engine.work_dir=directory
            first=engine._request('按本次资料判断',{'task':'缓存隔离检查'},lambda x:x)
            engine.settings['model_reasoning_effort']='medium'
            second=engine._request('按本次资料判断',{'task':'缓存隔离检查'},lambda x:x)
            engine.settings['model_reasoning_effort']='high'
            again=engine._request('按本次资料判断',{'task':'缓存隔离检查'},lambda x:x)
            self.assertEqual(first,{'判断':'第一次结果'})
            self.assertEqual(second,{'判断':'第二次结果'})
            self.assertEqual(again,first)
            self.assertEqual(engine.usage['cached_requests'],1)


if __name__=='__main__':unittest.main()
