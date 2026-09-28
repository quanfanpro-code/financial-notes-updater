"""生成完全虚构的资金小表，演示离线初筛；每次保存到新目录。"""
import json
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from 测试收尾 import 虚构参照, 内存索引
from 运行初筛 import 执行比较


def main():
    根 = Path(__file__).resolve().parent
    输出 = 根 / '运行结果' / ('虚构示例_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    索引目录 = 输出 / '演示索引'
    索引目录.mkdir(parents=True, exist_ok=False)
    表 = 虚构参照()
    表['上层语义'] = {'完整': True, '业务路径': [{'附注': '虚构货币资金', '表义': '虚构余额构成',
        '映射依据': ['虚构演示依据，不是正式金标准']}],
        '名称依据': [{'原文': '演示表', '来源': '虚构工作表名称'},
                    {'原文': '演示货币资金', '来源': '虚构小表名称'}]}
    索引 = 内存索引(表)
    (索引目录 / '小表记录.jsonl').write_text(json.dumps(表, ensure_ascii=False)+'\n', encoding='utf-8-sig')
    (索引目录 / '文字查找.json').write_text(json.dumps({'正文': 索引['正文'], '别名': {}}, ensure_ascii=False), encoding='utf-8-sig')
    (索引目录 / '索引清单.json').write_text(json.dumps({'规则版本': '虚构演示', '说明': '演示只有一张虚构小表，不是正式714张索引'}, ensure_ascii=False), encoding='utf-8-sig')
    w = Workbook(); s = w.active; s.title = '演示表'
    for n in 表['节点']:
        s.cell(n['行'], n['列'], n['原文'])
    s['B3'] = 123
    s['B4'] = 456
    s['B5'] = 579
    新表 = 输出 / '虚构新表.xlsx'
    w.save(新表); w.close()
    码 = 执行比较(新表, 索引目录, 输出 / '初筛材料', True)
    print('以上金额完全虚构。业务数值变动不会改变结构文字的对应；初筛没有批准语义。')
    print('请打开结果目录中的“初筛报告.md”和“LLM交接说明.md”查看。')
    return 码


if __name__ == '__main__':
    sys.exit(main())
