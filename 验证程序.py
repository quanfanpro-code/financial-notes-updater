"""运行本项目行为测试；不调用收费模型，不代表语义识别准确率。"""
import sys
import shutil
import tempfile
import unittest
from pathlib import Path
from datetime import datetime

ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
# 只运行当前四步路线及必要历史读取器的检查，不执行已停用的双轮生成测试。
modules=['测试.测试设置', '测试.测试表格', '测试.测试Word桥接', '测试.测试金标准', '测试.测试统一语义', '测试.测试统一迁移', '测试.测试识别输入与保存', '测试.测试逐格识别入口', '测试.测试唯一身份证匹配', '测试.测试正常入口贯通', '测试.测试表格文字结构', '测试.测试Word原文材料', '测试.测试维度格身份证入口', '测试.测试章节与既有身份参照', '测试.测试格式错误格补正']
folder=ROOT/"测试结果"/("行为验证_"+datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
folder.mkdir(parents=True,exist_ok=False)
tempfile.tempdir=str(folder)
# 和公开版入口一致，合成样本保留，不自动删除。
shutil.rmtree=lambda *args,**kwargs:None
suite=unittest.defaultTestLoader.loadTestsFromNames(modules)
path=folder/("行为验证_"+datetime.now().strftime("%Y%m%d_%H%M%S_%f")+".txt")
with path.open("x",encoding="utf-8-sig") as stream:
    result=unittest.TextTestRunner(stream=stream,verbosity=2).run(suite)
print(f"运行 {result.testsRun} 项；失败 {len(result.failures)} 项；错误 {len(result.errors)} 项。")
print("完整记录："+str(path))
print("本记录只证明程序行为，不证明模型识别准确率。")
raise SystemExit(0 if result.wasSuccessful() else 1)
