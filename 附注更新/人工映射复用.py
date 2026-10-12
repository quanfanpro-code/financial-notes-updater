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


def metric_template(row, standard, digest):
    """用现有部分语义校验模板；不补造实际主体、日期或范围。"""
    from .统一语义 import validate_partial_semantics
    if row.get('target_slot_id') or row.get('target_standard_sha256') != digest:
        raise ValueError('指标模板目标冲突或标准版本不符')
    if not row.get('confirmation'):
        raise ValueError('指标模板缺少本次确认依据')
    metric = standard['metric'].get(row.get('target_metric_id'))
    if metric is None:
        raise ValueError('指标模板目标不存在')
    dims = row.get('dimensions', {})
    result = {'metric_id': row['target_metric_id'], 'dimensions': dims,
              'dimension_evidence': row.get('dimension_evidence', {}),
              'missing_dimensions': sorted({d['id'] for d in metric['dimension_refs'] if d['required']} - set(dims)),
              'evidence_cells': [row['worksheet'] + '!' + row['cell']]}
    validate_partial_semantics(result, standard)
    if row.get('template_period_role') not in ('opening', 'closing', 'current', 'prior'):
        raise ValueError('模板须保留有原文依据的期间角色')
    return result


def install(source_root, output_dir=None, standard_path=None):
    """首次复制正式映射和金标准，校验副本并生成可重建索引；不覆盖已有成果。"""
    from .评测 import _paths, _jsonl, _hash
    source_root = Path(source_root)
    corpus = Path(output_dir) / CORPUS.name if output_dir is not None else CORPUS
    if corpus.parent.exists():
        raise ValueError('人工映射目录已存在，更新前须先备份，不能覆盖')
    gold_path = source_root / '语义金标准' / '财务报表附注语义金标准-v1.jsonl'
    gold = {row['id']: row for row in _jsonl(gold_path)}
    standard, standard_digest = None, None
    if standard_path:
        from .统一语义 import validate_records
        standard = validate_records(list(_jsonl(standard_path)))
        standard_digest = _hash(standard_path)
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
                metric = row.get('target_metric_id')
                if not slot and not metric:
                    if row.get('status') != '结构格排除':
                        raise ValueError('业务映射缺少目标：' + row['mapping_id'])
                    continue
                if metric and standard is None:
                    raise ValueError('直接指标引用须提供固定当前标准')
                partial = metric_template(row, standard, standard_digest) if metric else None
                if not metric and gold.get(slot, {}).get('status') != 'active':
                    raise ValueError('人工映射引用非active定义：' + slot)
                location = (row['worksheet'], row['cell'])
                if location in seen:
                    raise ValueError('人工映射出现重复来源：' + str(location))
                seen.add(location)
                if slot: referenced.add(slot)
                examples.append({'style': style, **({'legacy_slot_id': slot} if not metric else
                    {**partial, 'standard_sha256': standard_digest,
                     'template_period_role': row['template_period_role'], 'confirmation': row['confirmation']}),
                    'row_semantics': row.get('row_semantics', ''),
                    'column_semantics': row.get('column_semantics', ''),
                    'table_semantics': row.get('subtable', row.get('target_slot', {}).get('table', '')),
                    **({'legacy_dimensions': row.get('dimensions', {})} if not metric else {}),
                    **({'confirmation': row['confirmation']} if row.get('confirmation') else {}),
                    'origin': {'mapping_file': relative, 'mapping_sha256': digest,
                               'mapping_id': row['mapping_id'], 'sheet': location[0], 'cell': location[1]}})
        counts[str(style)] = {'business_cells': len(seen), 'workbook_sha256': _hash(workbook)}
    prepared.append((gold_path, gold_path.relative_to(source_root).as_posix()))
    if standard_path:
        prepared.append((Path(standard_path), '语义金标准/当前指标标准.jsonl'))
    for path, relative in prepared:
        target = corpus.parent / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(path, target)
        digest = _hash(path)
        if _hash(target) != digest:
            raise ValueError('副本校验失败：' + relative)
        files[relative] = digest
    result = {'schema': '五种样式人工语义示例-v1', 'source_root': str(source_root),
              'source_gold_sha256': _hash(gold_path), 'files': files, 'styles': counts,
              'examples': examples, 'legacy_definitions': {key: gold[key] for key in sorted(referenced)}}
    if standard_path:
        result['metric_standard_sha256'] = standard_digest
    with corpus.open('x', encoding='utf-8-sig') as handle:
        json.dump(result, handle, ensure_ascii=False, separators=(',', ':'))
    return {'path': str(corpus), 'business_cells': len(examples), 'copied_files': len(files), 'styles': counts}


def _terms(text):
    words = re.findall(r'[\u4e00-\u9fff]+|[a-zA-Z]+', text.lower())
    return {part[i:i+2] for part in words for i in range(max(1, len(part)-1))}


@lru_cache(maxsize=2)
def _standard_snapshot(path, stamp, size):
    from .统一语义 import validate_records
    raw = Path(path).read_bytes()
    return hashlib.sha256(raw).hexdigest(), validate_records(
        [json.loads(l) for l in raw.decode('utf-8-sig').splitlines() if l.strip()])


@lru_cache(maxsize=2)
def _load(path, stamp, size, snapshot_stamp=None):
    data = Path(path).read_bytes()
    corpus = json.loads(data.decode('utf-8-sig'))
    standard = None
    if any(r.get('metric_id') for r in corpus['examples']):
        snapshot = Path(path).parent / '语义金标准/当前指标标准.jsonl'
        info = snapshot.stat()
        digest,standard = _standard_snapshot(str(snapshot),info.st_mtime_ns,info.st_size)
        if digest != corpus.get('metric_standard_sha256'):
            raise ValueError('人工指标示例的固定标准快照不符')
    index = defaultdict(list)
    for number, row in enumerate(corpus['examples']):
        if row.get('metric_id'):
            candidate = {'target_metric_id':row['metric_id'],
                'target_standard_sha256':row.get('standard_sha256'),
                'dimensions':row['dimensions'],'dimension_evidence':row['dimension_evidence'],
                'worksheet':row['origin']['sheet'],'cell':row['origin']['cell'],
                'template_period_role':row['template_period_role'],'confirmation':row.get('confirmation')}
            expected = metric_template(candidate, standard, corpus['metric_standard_sha256'])
            if any(row.get(k) != v for k,v in expected.items()):
                raise ValueError('指标示例与完整模板校验不符')
        definition = (standard['metric'][row['metric_id']]['definition'] if row.get('metric_id') else
                      corpus['legacy_definitions'][row['legacy_slot_id']])
        # 原文检索只选参考材料，所有当前指标仍完整交给模型，不裁剪候选标准。
        text = json.dumps([row['row_semantics'], row['column_semantics'], row['table_semantics'],
                           definition.get('note'), definition.get('table')], ensure_ascii=False)
        for term in _terms(text):
            index[term].append(number)
    return corpus, index, hashlib.sha256(data).hexdigest()


def load_corpus(path=CORPUS):
    """快照变化同样使读取缓存失效。"""
    path = Path(path); info = path.stat()
    snapshot = path.parent/'语义金标准/当前指标标准.jsonl'
    stamp = snapshot.stat() if snapshot.exists() else None
    return _load(str(path), info.st_mtime_ns, info.st_size,
                 (stamp.st_mtime_ns,stamp.st_size) if stamp else None)


def reviewed_examples(packet):
    """返回可追溯的少量人工实例；当前财务值与坐标不参与参考材料检索。"""
    corpus, index, digest = load_corpus(CORPUS)
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
        if row.get('metric_id'):
            available = {m.get('id') for m in packet.get('metrics', packet.get('existing_metrics', [])) if isinstance(m, dict)}
            if row['metric_id'] not in available:
                continue
        key = (row['style'], row.get('legacy_slot_id') or row['metric_id'])
        if key in seen:
            continue
        seen.add(key); selected.append(row)
        if len(selected) == 16:
            break
    return {'instructions': INSTRUCTIONS, 'corpus_sha256': digest,
            'corpus_business_cells': len(corpus['examples']), 'source_gold_sha256': corpus['source_gold_sha256'],
            'examples': selected, 'legacy_definitions': {
                row['legacy_slot_id']: corpus['legacy_definitions'][row['legacy_slot_id']] for row in selected if row.get('legacy_slot_id')}}


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='生成五样式语义示例；输出目录必须尚不存在')
    parser.add_argument('来源根')
    parser.add_argument('--输出目录')
    parser.add_argument('--当前标准')
    args = parser.parse_args()
    print(json.dumps(install(args.来源根, args.输出目录, args.当前标准), ensure_ascii=False, indent=2))
