"""为完整定义未变的旧槽位补上来源版本关联，不重新批准语义、不改写旧证据。"""
import copy
import json
from functools import lru_cache
from pathlib import Path
from .统一语义 import definition_hash, validate_records, _hash


@lru_cache(maxsize=2)
def _manual_corpus(path, stamp, size):
    from .评测 import TRACE_KEYS
    corpus=json.loads(Path(path).read_text(encoding='utf-8-sig'))
    forms={}
    for item in corpus['examples']:
        form={key:item[key] for key in ('row_semantics','column_semantics','table_semantics','legacy_dimensions')}
        form['legacy_dimensions']={key:value for key,value in form['legacy_dimensions'].items() if key not in TRACE_KEYS}
        forms.setdefault(item['legacy_slot_id'],{})[definition_hash(form)]=form
    return corpus,forms


def recognition_expressions(standard):
    """编译已审定的新旧表达关系，不检索、不裁剪当前指标，也不生成当前表格事实。"""
    approved=[row for row in standard['legacy_conversion'].values() if row['status']=='approved'
              and row['target']['record_type']=='metric_fact']
    if not approved:
        return {'expressions':[], 'rules':{}}
    path=Path(__file__).resolve().parents[1]/'金标准/五种样式人工映射/语义示例索引.json'
    info=path.stat();corpus,forms=_manual_corpus(str(path),info.st_mtime_ns,info.st_size)
    expressions=[];rules={};rule_ids={}
    for conversion in approved:
        if conversion['source_gold_hash']!=corpus['source_gold_sha256']:
            continue
        identifier=conversion['legacy_id'];original=corpus['legacy_definitions'].get(identifier)
        if original is None:
            continue
        if definition_hash(original)!=conversion['legacy_definition_hash']:
            raise ValueError('已审表达与原版完整定义不一致：'+identifier)
        target=conversion['target'];metric=standard['metric'][target['metric_id']]
        if metric['status']!='active':
            continue
        signature=definition_hash(target)
        rule=rule_ids.setdefault(signature,'rule_'+str(len(rule_ids)+1))
        rules[rule]=copy.deepcopy(target)
        expressions.append({'legacy_id':identifier,'conversion_id':conversion['id'],
            'rule':rule,
            'original_meaning':{'note':original['note']['name'],'table':original['table']['name'],
                                'row_path':original['row_path'],'column_path':original['column_path']},
            'manual_forms':list(forms.get(identifier,{}).values())})
    shared={}
    for metric_id in {rule['metric_id'] for rule in rules.values()}:
        family=[rule for rule in rules.values() if rule['metric_id']==metric_id]
        common={key:value for key,value in family[0]['dimension_bindings'].items()
                if all(rule['dimension_bindings'].get(key)==value for rule in family)}
        if len(family)>1 and common:
            shared[metric_id]=common
            for rule in family:
                rule['shared_bindings']=metric_id
                rule['dimension_bindings']={key:value for key,value in rule['dimension_bindings'].items() if key not in common}
    return {'source_gold_sha256':corpus['source_gold_sha256'],'expressions':expressions,'rules':rules,'shared_bindings':shared,
        'instructions':'这是已审核的旧金标准表达及其到当前指标、维度的转换关系，manual_forms保留人工确认过的不同表式表达。'
        '依据完整上下文理解，不能只见同名词就套用。rule引用rules中的指标和维度取值规则；规则中的shared_bindings引用顶层同名目录，'
        '与该规则dimension_bindings合并才是完整维度规则，不能忽略共用部分。constant是旧定义明确的语义限定，'
        '仅当当前原文表达同一完整含义时适用；context和dimension必须从当前原文及其适用说明重新取证。'
        'manual_forms中的旧维度字段是历史表达，不是当前维度协议，不能照搬成新的槽位或维度。'
        '历史表式不提供本次财务值、主体或日期；不能用转换编号、历史实例作为当前格的原文证据。'
        '输出仍使用当前指标和实际维度，并引用当前标题、表头、项目及适用说明。'
        '未出现在本清单不表示不在金标准范围内，仍须核对全部当前指标和维度。'}


def bridge_definitions(originals, source_hash, rows):
    """返回完整候选标准及未衔接原因；调用方须从实际原版文件取得定义和哈希。"""
    _hash(source_hash)
    standard = validate_records(rows)
    result = copy.deepcopy(rows)
    linked, pending = [], []
    conversions = list(standard['legacy_conversion'].values())
    for identifier, original in originals.items():
        if original.get('id') != identifier:
            raise ValueError('旧定义编号与索引不一致')
        digest = definition_hash(original)
        related = [r for r in conversions if r['legacy_id'] == identifier]
        exact = [r for r in related if r['source_gold_hash'] == source_hash]
        if exact:
            if exact[0]['legacy_definition_hash'] != digest:
                raise ValueError('已登记版本的完整定义不一致：' + identifier)
            if exact[0]['status'] == 'approved':
                linked.append(identifier)
            else:
                pending.append({'legacy_id': identifier, 'reason': '该原版定义转换仍待复核'})
            continue
        usable = [r for r in related if r['status'] == 'approved' and r['legacy_definition_hash'] == digest]
        if not usable or len({definition_hash(r['target']) for r in usable}) != 1:
            pending.append({'legacy_id': identifier, 'reason': '缺少完整定义一致且去向唯一的已批准转换',
                            'existing_conversion_ids': [r['id'] for r in related]})
            continue
        approved = usable[0]
        suffix = source_hash[:16] + ':' + identifier
        source_id = 'source.original_bridge:' + suffix
        proof = {'record_type': 'source_example', 'id': source_id,
                 'origin': {'file_hash': source_hash, 'location': '旧定义ID:' + identifier},
                 'text_evidence': [
                     '原版完整定义与已批准转换绑定的完整定义逐字段一致，保留全部字段，不以名称相似代替。',
                     '完整定义哈希：' + digest,
                     '沿用转换：' + approved['id'] + '；转换哈希：' + definition_hash(approved)],
                 'review_refs': list(approved['evidence_refs'])}
        converted = copy.deepcopy(approved)
        converted.update(id='conversion.original_bridge:' + suffix, source_gold_hash=source_hash)
        converted['evidence_refs'] = [*approved['evidence_refs'], source_id]
        converted['review'] = {'reviewer': '程序：完整定义一致性核验；原业务审核：' + approved['review']['reviewer'],
                               'reason': '只补来源版本关联，转换目标及取值规则逐字段沿用。' + approved['review']['reason']}
        result.extend([proof, converted]); linked.append(identifier)
    validate_records(result)
    return {'records': result, 'linked': linked, 'pending': pending}
