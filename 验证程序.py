"""运行本项目行为测试；不调用收费模型，不代表语义识别准确率。"""
import sys
import unittest
from pathlib import Path
from datetime import datetime

ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
modules=["测试.测试语义","测试.测试设置","测试.测试表格","测试.测试流程","测试.测试结构","测试.测试评测","测试.测试Word桥接","测试.测试分步流程","测试.测试金标准","测试.测试语义补充","测试.测试识别范围","测试.测试回执格式","测试.测试统一语义","测试.测试统一迁移","测试.测试统一识别","测试.测试统一勾稽"]
modules.append('测试.测试同义复核')
modules.append('测试.测试统一补充')
modules.append('测试.测试统一上下文')
modules.append('测试.测试统一结构')
modules.append('测试.测试逐格回执隔离')
modules.append('测试.测试金标准重复读取')
suite=unittest.defaultTestLoader.loadTestsFromNames(modules)
folder=ROOT/"测试结果";folder.mkdir(exist_ok=True)
path=folder/("行为验证_"+datetime.now().strftime("%Y%m%d_%H%M%S_%f")+".txt")
with path.open("x",encoding="utf-8-sig") as stream:
    result=unittest.TextTestRunner(stream=stream,verbosity=2).run(suite)
print(f"运行 {result.testsRun} 项；失败 {len(result.failures)} 项；错误 {len(result.errors)} 项。")
print("完整记录："+str(path))
print("本记录只证明程序行为，不证明模型识别准确率。")
raise SystemExit(0 if result.wasSuccessful() else 1)
