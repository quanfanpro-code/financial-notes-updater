"""离线行为检查：不请求收费模型；临时样本保留供回查。"""
import sys
import shutil
import tempfile
import unittest
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path
from datetime import datetime

根 = Path(__file__).resolve().parent
sys.path.insert(0, str(根))
sys.path.insert(0, str(根 / '结构初筛'))
目录 = 根 / '测试结果' / ('公开版离线检查_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
目录.mkdir(parents=True, exist_ok=False)
tempfile.tempdir = str(目录)
# 测试样本留存；不以自动化清理绕过用户的删除确认。
shutil.rmtree = lambda *args, **kwargs: None
模块 = ['测试收尾', '测试.测试设置', '测试.测试语义', '测试.测试模型传输收尾', '测试.测试表格', '测试.测试统一语义']
套件 = unittest.defaultTestLoader.loadTestsFromNames(模块)
日志 = 目录 / '测试日志.txt'
with 日志.open('x', encoding='utf-8-sig') as f:
    with redirect_stdout(f), redirect_stderr(f):
        结果 = unittest.TextTestRunner(stream=f, verbosity=2).run(套件)
print(f'运行{结果.testsRun}项，通过{结果.testsRun-len(结果.failures)-len(结果.errors)-len(结果.skipped)}项，失败{len(结果.failures)}项，错误{len(结果.errors)}项，跳过{len(结果.skipped)}项。')
print('日志：' + str(日志))
print('本检查不代表模型识别准确率或完整财务报告业务验收。')
sys.exit(0 if 结果.wasSuccessful() else 1)
