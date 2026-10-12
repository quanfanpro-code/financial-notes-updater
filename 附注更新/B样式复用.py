# coding: utf-8
"""把已审定的旧表式转换送入共用绑定；当前上下文及差异由原文确认。"""
from copy import deepcopy
import json
from pathlib import Path
from .人工映射复用 import CORPUS, load_corpus
from .统一语义 import _require, definition_hash
from .表格 import file_hash


def reuse_templates(comparison, standard, corpus_path=CORPUS):
    corpus, _, corpus_hash=load_corpus(corpus_path)
    examples={e['origin']['mapping_id']:e for e in corpus['examples'] if e.get('legacy_slot_id')}
    conversions={r['legacy_id']:r for r in standard['legacy_conversion'].values()
                 if r['status']=='approved' and r['source_gold_hash']==corpus['source_gold_sha256']
                 and r['target']['record_type']=='metric_fact'}
    checked_files={};result={}
    for candidate in comparison['候选']:
        if candidate.get('冲突'):continue
        for pair in candidate['结构对应']:
            if pair['角色']!='业务数值':continue
            ref=pair.get('语义引用',{});example=examples.get(ref.get('mapping_id'))
            if not example or example['legacy_slot_id'] not in conversions:continue
            old=corpus['legacy_definitions'][example['legacy_slot_id']]
            conversion=conversions[example['legacy_slot_id']];target=conversion['target']
            _require(definition_hash(old)==conversion['legacy_definition_hash'],'表式复用的旧完整定义已变化')
            _require(example['origin']['sheet']==pair['旧工作表'] and example['origin']['cell']==pair['旧坐标'],
                     '表式复用的旧坐标与实际映射不一致')
            file=Path(corpus_path).parent/example['origin']['mapping_file']
            if file not in checked_files:
                _require(file_hash(file)==corpus['files'][example['origin']['mapping_file']]==example['origin']['mapping_sha256'],
                         '表式复用的人工映射原件已变化')
                checked_files[file]={r['mapping_id']:r for line in file.read_text(encoding='utf-8-sig').splitlines()
                                     if line.strip() for r in [json.loads(line)]}
            mapped=checked_files[file].get(ref['mapping_id'],{})
            _require(mapped.get('target_slot_id')==example['legacy_slot_id'] and
                     mapped.get('worksheet')==example['origin']['sheet'] and mapped.get('cell')==example['origin']['cell'],
                     '表式示例与人工映射目标不一致')
            metric=standard['metric'].get(target['metric_id'])
            if not metric or metric['status']!='active':continue
            constants={key:deepcopy(rule['value']) for key,rule in target['dimension_bindings'].items() if rule['op']=='constant'}
            # 主体、期间、口径、币种、单位始终取本次原文；旧实例不进入新表。
            for key in ('entity','period','report_scope','currency','unit_scale'):constants.pop(key,None)
            item={'location':pair['新工作表']+'!'+pair['新坐标'],'metric_id':target['metric_id'],
                  'constant_dimensions':constants,
                  'context_bindings':{key:deepcopy(rule) for key,rule in target['dimension_bindings'].items() if key not in constants},
                  'row_semantics':example['row_semantics'],'column_semantics':example['column_semantics'],
                  'table_semantics':example['table_semantics'],'differences':deepcopy(candidate['差异']),
                  'ambiguities':deepcopy(candidate.get('歧义',[])),
                  'origin':{'corpus_sha256':corpus_hash,'mapping_id':ref['mapping_id'],
                            'mapping_sha256':example['origin']['mapping_sha256'],'conversion_id':conversion['id'],
                            'legacy_definition_hash':conversion['legacy_definition_hash'],'old_table':candidate['旧表编号']}}
            item['id']='style:'+definition_hash(item)[:24]
            result[item['id']]=item
    return list(result.values())


def apply_reuse(row, table, templates):
    """只展开已明确选中的模板，不依据结构分数自动决定业务等价。"""
    from .A语义阅读 import addresses
    result=deepcopy(row);choices=result.pop('style_reuse',[])
    if not choices:return result
    index={t['id']:t for t in templates};owned={a for area in table['data_ranges'] for a in addresses(area)}
    seen=set()
    for choice in choices:
        template=index.get(choice.get('template_id'))
        _require(template is not None,'表式复用引用未知模板')
        sheet,_,cell=template['location'].rpartition('!')
        _require(sheet==table['sheet'] and cell in owned and cell not in seen,'表式复用越界或重复')
        seen.add(cell)
        _require(isinstance(choice.get('reason'),str) and choice['reason'].strip(),'表式复用须说明本次实际含义一致的依据')
        evidence=choice.get('evidence')
        _require(isinstance(evidence,list) and evidence and all(isinstance(e,str) for e in evidence),'表式复用缺少本次原文证据')
        _require(not any(a in owned for e in evidence for a in (addresses(e) if not e.startswith('@') else [])),
                 '表式复用不得以目标金额为依据')
        _require(not any(cell in (owned if area=='*' else set(addresses(area)))
                         for m in result.get('metrics',[]) for area in m['ranges']), '复用格不能重复生成指标')
        result.setdefault('metrics',[]).append({'ranges':[cell],'id':template['metric_id']})
        for key,value in template['constant_dimensions'].items():
            result.setdefault('dimensions',[]).append({'ranges':[cell],'id':key,'value':deepcopy(value),'evidence':deepcopy(evidence)})
    return result


def prepare_reuse(snapshot, scope, outline, standard):
    from .B机械比较 import read_comparison
    comparison=read_comparison(scope['mechanical_comparison'],snapshot,outline)
    wanted=set(scope['candidate_locations'])
    snapshot['semantic_style_templates']=[t for t in reuse_templates(comparison,standard) if t['location'] in wanted]
