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
sys.path.insert(0, str(根 / '测试'))
目录 = 根 / '测试结果' / ('公开版离线检查_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
目录.mkdir(parents=True, exist_ok=False)
tempfile.tempdir = str(目录)
# 测试样本留存；不以自动化清理绕过用户的删除确认。
shutil.rmtree = lambda *args, **kwargs: None
# 只运行当前四步路线及必要历史读取器的检查，不执行已停用的双轮生成测试。
模块=['测试.测试设置', '测试.测试表格', '测试.测试Word桥接', '测试.测试金标准', '测试.测试统一语义', '测试.测试统一迁移', '测试.测试识别输入与保存', '测试.测试逐格识别入口', '测试.测试唯一身份证匹配', '测试.测试正常入口贯通', '测试.测试表格文字结构', '测试.测试Word原文材料', '测试.测试维度格身份证入口', '测试.测试章节与既有身份参照', '测试.测试格式错误格补正']
套件 = unittest.defaultTestLoader.loadTestsFromNames(模块)
日志 = 目录 / '测试日志.txt'
with 日志.open('x', encoding='utf-8-sig') as f:
    with redirect_stdout(f), redirect_stderr(f):
        结果 = unittest.TextTestRunner(stream=f, verbosity=2).run(套件)
print(f'运行{结果.testsRun}项，通过{结果.testsRun-len(结果.failures)-len(结果.errors)-len(结果.skipped)}项，失败{len(结果.failures)}项，错误{len(结果.errors)}项，跳过{len(结果.skipped)}项。')
print('日志：' + str(日志))
print('本检查不代表模型识别准确率或完整财务报告业务验收。')
sys.exit(0 if 结果.wasSuccessful() else 1)
