# coding: utf-8
"""模拟外部传输，检查原表图片真正进入A/B共用入口，不调用模型。"""
import base64
import json
from pathlib import Path
import unittest
from unittest.mock import patch

from PIL import Image
from 附注更新.A合集 import _records
from 附注更新.Codex语义调用 import call_codex
from 附注更新.分步流程 import run_step
from 附注更新.表格 import file_hash
from 附注更新.统一勾稽 import load_fact_mapping
from 测试 import 测试直接识别 as fixture


class 原表图片共用识别测试(unittest.TestCase):
    def setUp(self):
        work = Path(__file__).resolve().parents[1] / '输出' / '主线接续_20261027_213431' / '合成离线验证'
        self.case = fixture.直接识别测试()
        with patch.object(fixture, 'WORK', work):
            self.case.setUp()
        self.image = self.case.folder / '合成原表图片.png'
        Image.new('RGB', (8, 8), 'yellow').save(self.image)
        self.proof = {'path': str(self.image), 'sha256': file_hash(self.image),
                      'source_hash': file_hash(self.case.source), 'locations': ['合成表!B3']}
        self.raw = json.dumps({'location': '合成表!B3', 'metric_id': '固定资产原值',
                               'dimensions': {'asset_selection': {'mode': 'members', 'members': ['设备甲']}}},
                              ensure_ascii=False)

    def test_AB普通入口真正附图且不改写本格身份(self):
        previous = None
        for side in ('a', 'b'):
            config = {**self.case.config, side + '_source_images': [self.proof],
                      side + '_target_table': {'id': '单格', 'sheet': '合成表',
                                               'range': 'A2:C8', 'data_ranges': ['B3']}}
            if side == 'b':
                config['a_mapping'] = previous
            commands = []
            def spawn(command, **options):
                commands.append(command)
                self.assertEqual(command[command.index('--image') + 1], str(self.image.resolve()))
                Path(command[command.index('--output-last-message') + 1]).write_text(self.raw, encoding='utf-8')
                class Process:
                    returncode = 0
                    def communicate(inner, input=None, timeout=None):
                        text = input.decode('utf-8')
                        self.assertIn('本次作答位置.json', text)
                        self.assertIn('合成表!B3', text)
                        self.assertIn('原表图片.json', text)
                        return b'{"type":"turn.completed","usage":{}}\n', b''
                return Process()
            with patch('附注更新.Codex语义调用.shutil.which', return_value=r'C:\本机\codex.exe'), \
                 patch('附注更新.Codex语义调用.subprocess.Popen', side_effect=spawn):
                result = run_step('recognize_' + side, config)
            self.assertEqual(result['status'], 'complete', result)
            self.assertEqual(len(commands), 1)
            saved = load_fact_mapping(result['files'][side + '_mapping'], {})
            previous = saved['mapping_path'] if side == 'a' else previous
            self.assertEqual(saved['本次逐格原答复'], self.raw)
            row = dict(_records(saved))['合成表!B3']
            self.assertEqual(row['metric_id'], '固定资产原值')
            self.assertEqual(row['dimensions'], {'asset_selection': {'mode': 'members', 'members': ['设备甲']}})
            requests = [json.loads(f.read_text(encoding='utf-8-sig')) for f in self.case.folder.rglob('请求.json')]
            self.assertTrue(any(r.get('images') == [self.proof] for r in requests))

    def test_图片变化不能复用已缓存答复(self):
        engine = self.case.engine()
        materials = {'原表图片.json': [self.proof]}
        with patch.object(engine, '_http', return_value=self.raw) as send:
            self.assertEqual(engine.request_text('读取原表', materials), self.raw)
            self.assertEqual(engine.request_text('读取原表', materials), self.raw)
            original = Path.read_bytes
            def changed(path):
                return b'changed-image' if path.resolve() == self.image.resolve() else original(path)
            with patch.object(Path, 'read_bytes', changed):
                with self.assertRaisesRegex(ValueError, '图片已变化'):
                    engine.request_text('读取原表', materials)
            self.assertEqual(send.call_count, 1)

    def test_图片与原件不符在发送前停止(self):
        config = {**self.case.config, 'a_source_images': [{**self.proof, 'source_hash': '0' * 64}]}
        with patch('附注更新.统一识别.UnifiedSemanticEngine.request_text',
                   side_effect=AssertionError('错误图片不能发送')) as send:
            result = run_step('recognize_a', config)
        self.assertEqual(result['status'], 'failed')
        self.assertIn('图片不属于本次原件', result['message'])
        send.assert_not_called()

    def test_原生调用记录图片且坏哈希不启动进程(self):
        with patch('附注更新.Codex语义调用.subprocess.Popen',
                   side_effect=AssertionError('坏哈希不能启动模型')) as process:
            with self.assertRaisesRegex(ValueError, '图片已变化'):
                call_codex({'model': 'offline-test'}, self.case.folder,
                           {'messages': [], 'images': [{**self.proof, 'sha256': '0' * 64}]}, lambda: None)
        process.assert_not_called()

    def test_HTTP入口发送图片内容而非路径(self):
        engine = self.case.engine('https://example.invalid')
        def receive(timeout, reader, request, *args):
            body = json.loads(request.data)
            parts = body['messages'][1]['content']
            self.assertEqual(parts[0]['type'], 'text')
            url = parts[1]['image_url']['url']
            self.assertTrue(url.startswith('data:image/png;base64,'))
            self.assertEqual(base64.b64decode(url.split(',', 1)[1]), self.image.read_bytes())
            data = {'choices': [{'finish_reason': 'stop', 'message': {'content': self.raw}}], 'usage': {}}
            return data, json.dumps(data).encode('utf-8')
        with patch('附注更新.语义._call_with_deadline', side_effect=receive):
            answer = engine.request_text('读取原图', {'原表图片.json': [self.proof]})
        self.assertEqual(answer, self.raw)


if __name__ == '__main__':
    unittest.main()
