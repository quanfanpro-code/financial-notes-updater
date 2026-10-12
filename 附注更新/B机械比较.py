# coding: utf-8
"""把已经定位的B范围交给已有五样式比较器，不在此确认语义。"""
from pathlib import Path
import json
from openpyxl.utils.cell import range_boundaries

from .表格编排 import in_ranges, check_ranges
from .表格 import file_hash
from .统一语义 import _require
from 结构初筛.比较新表 import 载入索引, 比较工作簿, 完整性检查, 生成版本


def compare_located(snapshot, outline, index_path=None, log=lambda text: None):
    index_path = Path(index_path or Path(__file__).resolve().parents[1] / '结构初筛' / '索引' / '工作表语义版_正式')
    ranges = [r for t in outline['subtables'] for r in t['ranges']] + outline['context_ranges']
    check_ranges(ranges, snapshot)
    try:
        index = 载入索引(index_path, 允许读取器版本变化=True)
    except ValueError as error:
        reason='五样式索引不可复用，继续按已定位范围正常识别：'+str(error)
        log(reason)
        return {'输入版本':{'sha256':snapshot['sha256']},'本次定位范围':ranges,'候选':[],
                '区域候选':{t['id']:[] for t in outline['subtables']},'说明':reason,
                '索引校验说明':[str(error)],'复用可用':False}
    sheets = []
    for sheet in snapshot['sheets']:
        areas = [r for r in ranges if r['sheet'] == sheet['name']]
        if not areas: continue
        boxes = [range_boundaries(r['range']) for r in areas]
        merges = []
        for merged in sheet.get('merges', []):
            left, top, right, bottom = range_boundaries(merged)
            if any(left <= b[2] and right >= b[0] and top <= b[3] and bottom >= b[1] for b in boxes):
                merges.append(merged)
        sheets.append({**sheet, 'cells': {a: c for a, c in sheet['cells'].items() if in_ranges(sheet['name'], a, areas)},
                       'merges': merges})
    scoped = {**snapshot, 'sheets': sheets}
    log(f'五样式机械比较：{len(outline["subtables"])}个已定位区域，{len(sheets)}张相关工作表')
    result = 比较工作簿(scoped, index)
    完整性检查(scoped, result)
    result['本次定位范围'] = ranges
    result['生成版本'] = 生成版本(snapshot['path'], index_path)
    result['区域候选'] = {t['id']: [i for i, c in enumerate(result['候选'])
                                if any(in_ranges(c['新工作表'], d['新坐标'], t['ranges']) for d in c['结构对应'])]
                        for t in outline['subtables']}
    # 比较器可能推算空白落点；范围外的推算仅供查看，不扩大本次识别范围。
    for candidate in result['候选']:
        candidate['范围外推算'] = [d['新坐标'] for d in candidate['结构对应']
                                  if not in_ranges(candidate['新工作表'], d['新坐标'], ranges)]
    result['说明'] = '仅比较本次A所需已定位范围；结构对应及旧语义引用仍须核对本次实际含义、年度、主体与单位。'
    return result


def comparison_material(result, table):
    """详看只带表式摘要和范围内业务对应；完整差异另存，避免重复发送整本结果。"""
    candidates = []
    for number in result['区域候选'][table['id']]:
        c = result['候选'][number]
        business = [d for d in c['结构对应'] if d['角色'] == '业务数值'
                    and in_ranges(d['新工作表'], d['新坐标'], table['ranges'])]
        candidates.append({'编号': number, '旧表': c['小表名称'], '样式': c['样式来源'],
                           'B工作表': c['新工作表'], 'B范围': c['新表范围'],
                           '旧表文字覆盖': {k: c['覆盖计算']['旧表文字覆盖'][k] for k in ('分子', '分母')},
                           '差异类型': sorted({d['类型'] for d in c['差异']}),
                           '冲突数': len(c['冲突']), '歧义数': len(c['歧义']),
                           '业务对应': business})
    return {'说明': result['说明'], '候选': candidates}


def read_comparison(reference, snapshot, outline):
    _require(file_hash(reference['path']) == reference['sha256'], 'B机械比较文件发生变化')
    result = json.loads(Path(reference['path']).read_text(encoding='utf-8-sig'))
    ranges = [r for t in outline['subtables'] for r in t['ranges']] + outline['context_ranges']
    _require(result['输入版本']['sha256'] == snapshot['sha256'] and result['本次定位范围'] == ranges,
             'B机械比较与原件或定位范围不符')
    return result
