"""通过 Windows 窗口选择已整理的语义成果和五样式材料，建立新的索引目录。"""
import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
from 建立结构索引 import 清点, 建立索引


def main():
    根 = tk.Tk(); 根.withdraw()
    try:
        成果库 = filedialog.askdirectory(title='选择包含资产清单.json的统一语义成果目录', parent=根)
        if not 成果库: return 1
        样式 = filedialog.askdirectory(title='选择包含附注样式1至5的Excel目录', parent=根)
        if not 样式: return 1
        输出根 = filedialog.askdirectory(title='选择保存新索引的文件夹（会在其内新建目录）', parent=根)
        if not 输出根: return 1
    finally:
        根.destroy()
    目录 = Path(输出根) / ('索引_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    目录.mkdir(exist_ok=False)
    清单 = 目录 / '输入清单.json'
    清点(成果库, 样式, 清单)
    结果 = 建立索引(清单, 目录 / '结构索引')
    print(f"已建立{结果['小表数']}张小表的结构索引。")
    print('下次初筛选择这个目录：' + str(目录 / '结构索引'))
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print('索引未完成：' + str(e), file=sys.stderr)
        sys.exit(1)
