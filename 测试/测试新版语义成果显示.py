# coding: utf-8
"""新版独立语义成果应能在实际界面方法中显示数量和处理说明。"""
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from 附注更新 import 界面


class 新版成果显示(unittest.TestCase):
    def test_业务格统计不把维度绑定当成另一格(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'映射.json'
            path.write_text(json.dumps({'schema_version': 2, 'carrier': 'A',
                'facts': [{'record_type': 'metric_fact'}, {'record_type': 'dimension_binding'}],
                'owners': {}, 'source_path': '来源.xlsx', 'source_hash': '哈希',
                'gold_path': '金标准.jsonl', 'gold_hash': '哈希',
                'excluded': [], 'unresolved': [{'reason': '待核实'}], 'complete': False}), encoding='utf-8')
            shown = {}
            label = SimpleNamespace(winfo_exists=lambda: True, configure=lambda **kw: shown.update(kw))
            app = SimpleNamespace(mapping_summary_labels={'a_mapping': label},
                                  values={'a_mapping': SimpleNamespace(get=lambda: str(path))})
            界面.Application.refresh_mapping_summary(app, 'a_mapping')
            self.assertIn('已确认业务 1 格', shown['text'])
            self.assertIn('待核实 1 格', shown['text'])

    def test_新版补充材料可直接显示处理说明(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'补充材料.json'
            path.write_text(json.dumps({'schema': '统一金标准补充材料-v2',
                'review_instructions': ['沿用已有槽位。', '真实缺项再补充金标准。']}), encoding='utf-8')
            app = SimpleNamespace(root=None, values={'a_standard_review': SimpleNamespace(get=lambda: str(path))})
            with patch.object(界面.messagebox, 'showinfo') as info, patch.object(界面.messagebox, 'showerror') as error:
                界面.Application.open_pending_instructions(app, 'A')
            error.assert_not_called()
            self.assertIn('真实缺项再补充金标准。', info.call_args.args[1])


if __name__ == '__main__':
    unittest.main()
