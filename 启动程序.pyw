"""双击启动中文桌面程序；启动失败时显示明确错误。"""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from 附注更新.界面 import main
    main()
except Exception as exc:
    import ctypes
    ctypes.windll.user32.MessageBoxW(None, "程序未能启动：\n" + str(exc) + "\n\n请将此提示提供给维护人员。", "附注自动更新系统", 0x10)