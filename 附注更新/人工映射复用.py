"""复用已人工确认的五种表式作为语义参考，不按坐标或数值推定新表身份。"""
from collections import Counter, defaultdict
from functools import lru_cache
import hashlib
import json
from pathlib import Path
import re
import shutil

ROOT = Path(__file__).resolve().parents[1]
CORPUS = ROOT / '金标准' / '五种样式人工映射' / '语义示例索引.json'
INSTRUCTIONS = (
    '以下是五种历史表式中人工确认的语义对应实例，仅供理解不同表达形式。'
    '检索按原文相似度挑选参考，不是语义等价结论；未检索到也不代表旧标准缺项。'
    'legacy_definitions保留旧槽位完整含义，legacy_dimensions是该实例的旧维度取值；'
    '其中字段名、期间角色等历史表达须重新区分指标、维度定义和实际取值，不能直接作为当前协议。'
    '当前输出只能使用本次metrics或existing_metrics及其声明的维度。'
    '若旧定义已有该含义，应先核查当前指标连同维度是否可表达，不能因表式不同重复建槽位。'
    '示例中的主体、期间、类别及origin只属于历史实例，不能用作当前表格的事实或证据。'
    '当前语义仍须由当前原表文字确认；财务值只从当前文件读取，不复制历史值。'
)


def install(source_root):
    """首次复制正式映射和金标准，校验副本并生成可重建索引；不覆盖已有成果。"""
    from .评测 import _paths, _jsonl, _hash
    source_root = Path(source_root)
    if CORPUS.parent.exists():
        raise ValueError('人工映射目录已存在，更新前须先备份，不能覆盖')
    gold_path = source_root / '语义金标准' / '财务报表附注语义金标准-v1.jsonl'
    gold = {row['id']: row for row in _jsonl(gold_path)}
    examples, files, counts, referenced = [], {}, {}, set()
    prepared = []
    for style in range(1, 6):
        workbook, _, paths = _paths(style, source_root)
        seen = set()
        for path in paths:
            relative = path.relative_to(source_root).as_posix()
            prepared.append((path, relative))
            digest = _hash(path)
            for row in _jsonl(path):
                slot = row.get('target_slot_id')
                if not slot:
                    continue
                if gold.get(slot, {}).get('status') != 'active':
                    raise ValueError('人工映射引用非active定义：' + slot)
                location = (row['worksheet'], row['cell'])
                if location in seen:
                    raise ValueError('人工映射出现重复来源：' + str(location))
                seen.add(location); referenced.add(slot)
                examples.append({'style': style, 'legacy_slot_id': slot,
                    'row_semantics': row.get('row_semantics', ''),
                    'column_semantics': row.get('column_semantics', ''),
                    'table_semantics': row.get('subtable', row.get('target_slot', {}).get('table', '')),
                    'legacy_dimensions': row.get('dimensions', {}),
                    'origin': {'mapping_file': relative, 'mapping_sha256': digest,
                               'mapping_id': row['mapping_id'], 'sheet': location[0], 'cell': location[1]}})
        counts[str(style)] = {'business_cells': len(seen), 'workbook_sha256': _hash(workbook)}
    prepared.append((gold_path, gold_path.relative_to(source_root).as_posix()))
    for path, relative in prepared:
        target = CORPUS.parent / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(path, target)
        digest = _hash(path)
        if _hash(target) != digest:
            raise ValueError('副本校验失败：' + relative)
        files[relative] = digest
    result = {'schema': '五种样式人工语义示例-v1', 'source_root': str(source_root),
              'source_gold_sha256': _hash(gold_path), 'files': files, 'styles': counts,
              'examples': examples, 'legacy_definitions': {key: gold[key] for key in sorted(referenced)}}
    with CORPUS.open('x', encoding='utf-8-sig') as handle:
        json.dump(result, handle, ensure_ascii=False, separators=(',', ':'))
    return {'path': str(CORPUS), 'business_cells': len(examples), 'copied_files': len(files), 'styles': counts}


def _terms(text):
    words = re.findall(r'[\u4e00-\u9fff]+|[a-zA-Z]+', text.lower())
    return {part[i:i+2] for part in words for i in range(max(1, len(part)-1))}


@lru_cache(maxsize=2)
def _load(path, stamp, size):
    data = Path(path).read_bytes()
    corpus = json.loads(data.decode('utf-8-sig'))
    index = defaultdict(list)
    for number, row in enumerate(corpus['examples']):
        definition = corpus['legacy_definitions'][row['legacy_slot_id']]
        # 原文检索只选参考材料，所有当前指标仍完整交给模型，不裁剪候选标准。
        text = json.dumps([row['row_semantics'], row['column_semantics'], row['table_semantics'],
                           definition.get('note'), definition.get('table')], ensure_ascii=False)
        for term in _terms(text):
            index[term].append(number)
    return corpus, index, hashlib.sha256(data).hexdigest()


def reviewed_examples(packet):
    """返回可追溯的少量人工实例；当前财务值与坐标不参与参考材料检索。"""
    info = CORPUS.stat()
    corpus, index, digest = _load(str(CORPUS), info.st_mtime_ns, info.st_size)
    candidates = set(packet.get('candidate_cells', []))
    text = str(packet.get('sheet', '')) + ' ' + ' '.join(
        cell['value'] for address, cell in packet.get('cells', {}).items()
        if address not in candidates and isinstance(cell.get('value'), str)
        and not cell.get('formula') and not cell['value'].startswith('='))
    scores = Counter()
    for term in sorted(_terms(text)):
        found = index.get(term, [])
        for number in found:
            scores[number] += 1 / len(found)
    selected, seen = [], set()
    # ponytail: 最多16例控制请求体积；召回不足时改进参考检索，不能把相似度变成业务匹配。
    for number, _ in scores.most_common():
        row = corpus['examples'][number]
        key = (row['style'], row['legacy_slot_id'])
        if key in seen:
            continue
        seen.add(key); selected.append(row)
        if len(selected) == 16:
            break
    return {'instructions': INSTRUCTIONS, 'corpus_sha256': digest,
            'corpus_business_cells': len(corpus['examples']), 'source_gold_sha256': corpus['source_gold_sha256'],
            'examples': selected, 'legacy_definitions': {
                row['legacy_slot_id']: corpus['legacy_definitions'][row['legacy_slot_id']] for row in selected}}


if __name__ == '__main__':
    import sys
    print(json.dumps(install(sys.argv[1]), ensure_ascii=False, indent=2))
