# coding: utf-8
"""按统一业务身份生成更新计划；不以坐标、旧编号或金额相等寻找来源。"""
from collections import defaultdict
from copy import deepcopy
import re
import hashlib
import json
from pathlib import Path
from openpyxl import load_workbook

from .表格 import _plain, file_hash, recalculate_copy
from .金标准 import _pairs, _invalid_number
from .统一证据 import context_catalog, resolve_context
from .统一语义 import (_hash, _require, semantic_key, validate_facts, value_for_target,
                     ensure_standard_compatible, read_standard_v2, validate_partial_semantics, validate_fact)


def _location(fact):
    sheet, separator, cell = fact['source_reference']['location'].rpartition('!')
    _require(separator and sheet and re.fullmatch(r'[A-Z]{1,3}[1-9]\d*', cell), '事实缺少可回写的原格位置')
    return sheet, cell


def _check_mapping(mapping, standard):
    _require(mapping.get('schema_version') == 2 and
             (mapping.get('gold_hash') == standard.get('sha256') or mapping.get('_validated_standard_hash') == standard.get('sha256')),
             '映射与统一标准版本不符')
    _hash(mapping['source_hash'])
    validate_facts(mapping['facts'], standard, mapping['owners'])
    occupied = set()
    for fact in mapping['facts']:
        _require(fact['source_reference']['file_hash'] == mapping['source_hash'], '事实与载体源文件版本不符')
        if fact['record_type'] == 'metric_fact':
            position = _location(fact)
            _require(position not in occupied, '同一财务数据格有重复事实')
            occupied.add(position)


def _calculated_values(mapping, receipt):
    """重读Excel原生重算副本；从未把来源原有缓存当作已核实结果。"""
    _require(receipt.get('schema')=='统一语义公式计算-v1' and receipt.get('method')=='Excel.CalculateFullRebuild',
             '公式计算依据类型不符')
    _require(receipt['source_hash']==mapping['source_hash'] and file_hash(receipt['source_path'])==mapping['source_hash'],
             '公式计算依据与原件版本不符')
    _require(file_hash(receipt['workbook_path'])==receipt['workbook_hash'],'公式计算副本已变化')
    original=load_workbook(mapping['source_path'],data_only=False)
    calculated=load_workbook(receipt['workbook_path'],data_only=False)
    values=load_workbook(receipt['workbook_path'],data_only=True)
    result={}
    try:
        _require(original.sheetnames==calculated.sheetnames,'公式计算前后工作表不一致')
        for name in original.sheetnames:
            def contents(sheet):
                return {cell.coordinate:cell.value for cell in sheet._cells.values() if cell.value is not None}
            _require(contents(original[name])==contents(calculated[name]),'公式计算前后原值或公式发生变化')
        for fact in mapping['facts']:
            if fact['record_type']!='metric_fact' or fact['raw_value_state']!='formula':continue
            sheet,address=_location(fact);location=sheet+'!'+address
            _require(original[sheet][address].data_type=='f' and original[sheet][address].value==fact['raw_value'],
                     '公式记录与原件不符')
            cell=values[sheet][address];value=cell.value
            if cell.data_type=='e':state,reason='error','Excel计算错误：'+str(value)
            elif value is None and cell.data_type not in ('str','s'):
                state,reason='error','重算后仍未取得计算结果'
            elif value is None or value=='':state,reason='blank','公式结果为空，不解释成零'
            elif not isinstance(value,(str,int,float,bool)):
                state,reason='error','计算结果不是可搬运的数值或文字';value=str(value)
            else:state,reason='value','已由Excel原生重算并回读'
            result[location]={'formula':fact['raw_value'],'state':state,'value':value,'reason':reason}
    finally:
        original.close();calculated.close();values.close()
    _require(file_hash(mapping['source_path'])==mapping['source_hash'] and file_hash(receipt['workbook_path'])==receipt['workbook_hash'],
             '读取计算结果期间文件已变化')
    return result


def calculate_mapping_formulas(mapping, folder):
    """语义识别仍绑定原文件，计算副本与原公式分开保存。"""
    if not any(row['record_type']=='metric_fact' and row['raw_value_state']=='formula' for row in mapping['facts']):return
    from .流程 import save_json
    from .表格 import read_workbook
    folder=Path(folder);folder.mkdir(parents=True,exist_ok=True)
    source=Path(mapping['source_path']);calculated=folder/('公式重算副本'+source.suffix)
    _require(file_hash(source)==mapping['source_hash'],'公式计算前原件已变化')
    recalculate_copy(source,calculated)
    snapshot=read_workbook(calculated)
    _require(snapshot.get('recalculated') and snapshot['recalculation']['source_sha256']==mapping['source_hash'],
             '未取得本次Excel原生重算依据')
    receipt={'schema':'统一语义公式计算-v1','method':snapshot['recalculation']['method'],
             'source_path':str(source.resolve()),'source_hash':mapping['source_hash'],
             'workbook_path':str(calculated.resolve()),'workbook_hash':snapshot['sha256']}
    checked=_calculated_values(mapping,receipt)
    path=save_json(folder/'公式重算依据.json',receipt)
    mapping['calculation_evidence']={'path':path,'sha256':file_hash(path)}
    mapping['formula_values']=checked


def calculation_files(mapping):
    evidence=mapping.get('calculation_evidence')
    if not evidence:
        _require(not mapping.get('formula_values'),'公式计算结果缺少计算依据')
        return []
    _require(file_hash(evidence['path'])==evidence['sha256'],'公式计算依据已变化')
    receipt=json.loads(Path(evidence['path']).read_text(encoding='utf-8-sig'),object_pairs_hook=_pairs,parse_constant=_invalid_number)
    _require(_calculated_values(mapping,receipt)==mapping.get('formula_values'),'保存的公式计算结果与实际副本不符')
    return [evidence['path'],receipt['workbook_path'],receipt['source_path']]


def _effective_source(fact, mapping, standard):
    if fact['raw_value_state']!='formula':return fact
    calculated=mapping.get('formula_values',{}).get(fact['source_reference']['location'])
    if not calculated or not mapping.get('calculation_evidence') or calculated['state']=='error':return fact
    _require(calculated['formula']==fact['raw_value'],'计算结果与原公式不符')
    result=deepcopy(fact);result['raw_value_state']=calculated['state'];result['raw_value']=calculated['value']
    from .统一语义 import validate_fact
    validate_fact(result,standard)
    return result


def load_fact_mapping(path, standard, _ancestors=frozenset()):
    """独立映射读回时核对当前原件、逐格原值、证据位置及已保存回执版本。"""
    path=Path(path).resolve(); raw=path.read_bytes()
    _require(path not in _ancestors,'派生映射存在循环引用')
    ancestors=_ancestors|{path}
    mapping=json.loads(raw.decode('utf-8-sig'),object_pairs_hook=_pairs,parse_constant=_invalid_number)
    mapping.pop('_validated_standard_hash',None)
    ensure_standard_compatible(mapping,standard)
    mapping['_validated_standard_hash']=standard['sha256']
    _check_mapping(mapping,standard)
    def verify_classification(content):
        from .统一识别 import agree_rounds
        for row in content.get('unresolved',[]):
            if not row.get('partial_semantics'):continue
            validate_partial_semantics(row['partial_semantics'],standard)
            rounds=row.get('rounds',[])
            _require(len(rounds)==2 and all(item.get('kind') in {'partial_metric_fact','metric_fact'} for item in rounds),
                     '部分语义缺少两轮独立记录')
            for item in rounds:
                if item['kind']=='partial_metric_fact':validate_partial_semantics(item['partial_semantics'],standard)
                else:validate_fact(item['fact'],standard)
            if not content.get('equivalence_proofs') and 'packet_mappings' not in content:
                expected=agree_rounds({row['cell']:rounds[0]},{row['cell']:rounds[1]},standard)
                _require(expected['unresolved'][0].get('partial_semantics')==row['partial_semantics'],'保存的部分语义与两轮识别不符')
        if not content.get('equivalence_proofs'):
            if not content.get('derivation') and 'packet_mappings' not in content:
                _require(not any(row['recognition_provenance']['method']=='model_v2_two_rounds_equivalence' for row in content['facts']),
                         '同义识别事实缺少保存的复核依据')
            return
        historical=standard if content['gold_hash']==standard['sha256'] else read_standard_v2(content['gold_path'])
        _require(historical['sha256']==content['gold_hash'],'同义复核所用历史标准变化')
        _require(len(content.get('classification_rounds',[]))==2,'同义识别缺少原始两轮分类')
        for proof in content['equivalence_proofs']:
            for pair in proof['pairs']:
                for side in ('left','right'):
                    _require(pair.get(side+'_context',[])==content.get('source_context',[]),'同义识别复核说明与实际分类不符')
                    if pair[side].get('record_type')=='partial_metric_fact':
                        reference=pair[side]['source_reference'];address=reference['location'].rpartition('!')[2]
                        _require(reference=={'file_hash':content['source_hash'],'location':content['sheet']+'!'+address} and
                                 Path(pair[side+'_path']).resolve()==Path(content['source_path']).resolve(),'部分语义复核未绑定实际原格与原件')
        expected=agree_rounds(*content['classification_rounds'],historical,equivalence_proofs=content['equivalence_proofs'],
                              agreement_version=content.get('agreement_version',1))
        for key in ('facts','owners','unresolved','excluded'):
            _require(content[key]==expected[key],'同义识别结果与原分类和复核不符：'+key)
    if mapping.get('external_review_v2'):
        from .统一外部复核 import validate_external_mapping
        validate_external_mapping(mapping,standard,ancestors)
    else:
        _require(not any(row['recognition_provenance']['method']=='external_semantic_review_v2' for row in mapping['facts'])
                 or mapping.get('derivation') or mapping.get('reused_mapping'),'外部语义缺少实际采纳依据')
        verify_classification(mapping)
    calculation_files(mapping)
    if mapping.get('derivation'):
        derivation=mapping['derivation']; plan_path=Path(derivation['plan_path'])
        _require(hashlib.sha256(plan_path.read_bytes()).hexdigest()==derivation['plan_hash'],'派生映射的更新清单已变化')
        plan=json.loads(plan_path.read_text(encoding='utf-8-sig'),object_pairs_hook=_pairs,parse_constant=_invalid_number)
        for original,digest in plan['inputs'].items():
            _require(hashlib.sha256(Path(original).read_bytes()).hexdigest()==digest,'派生映射的依据文件已变化')
        historical=standard if plan['gold_hash']==standard['sha256'] else read_standard_v2(plan['gold_path'])
        _require(historical['sha256']==plan['gold_hash'],'派生依据所用历史标准已变化')
        parent=load_fact_mapping(plan['a_mapping'],historical,ancestors)
        source_mapping=load_fact_mapping(plan['b_mapping'],historical,ancestors)
        fresh=verify_fact_update_plan(plan,parent,source_mapping,historical)
        expected=derive_fact_mapping(fresh.pop('_working_target'),mapping['source_path'],fresh,plan_path,mapping.get('a_link'))
        for key in ('facts','owners','excluded','unresolved','coverage','complete','source_context','derivation'):
            _require(mapping.get(key)==expected.get(key),'派生映射与原语义及更新清单不符：'+key)
    source=Path(mapping['source_path']).resolve()
    _require(hashlib.sha256(source.read_bytes()).hexdigest()==mapping['source_hash'],'映射绑定的原件已变化')
    book=load_workbook(source,data_only=False,keep_vba=source.suffix.lower()=='.xlsm')
    try:
        catalog=context_catalog(mapping.get('source_context',[]),source)
        for fact in mapping['facts']:
            sheet,cell=_location(fact)
            _require(sheet in book,'事实引用的工作表不存在')
            original=book[sheet][cell].value
            if fact['record_type']=='metric_fact':
                _require(original==fact['raw_value'],'映射原值与实际工作簿不符：'+sheet+'!'+cell)
                references=[ref for values in fact['dimension_evidence'].values() for ref in values]
                for reference in references:
                    if reference.startswith('@context:'):
                        resolve_context(reference,catalog,fact['source_reference']['location'])
                        continue
                    evidence_sheet,separator,address=reference.rpartition('!')
                    _require(separator and evidence_sheet in book and re.fullmatch(r'[A-Z]{1,3}[1-9]\d*',address),'维度证据不是真实原格位置')
                    value=book[evidence_sheet][address].value
                    _require(value is not None and str(value).strip(),'维度证据原格为空')
        for row in mapping.get('unresolved',[]):
            known=row.get('partial_semantics')
            if not known:continue
            sheet_name=row.get('sheet',mapping.get('sheet'))
            _require(sheet_name in book and re.fullmatch(r'[A-Z]{1,3}[1-9]\d*',row['cell']),'部分语义原格不存在')
            references=known['evidence_cells']+[ref for values in known['dimension_evidence'].values() for ref in values]
            for reference in references:
                if reference.startswith('@context:'):
                    resolve_context(reference,catalog,sheet_name+'!'+row['cell']);continue
                sh,separator,address=reference.rpartition('!')
                _require(separator and sh in book and re.fullmatch(r'[A-Z]{1,3}[1-9]\d*',address),'部分语义证据位置无效')
                value=book[sh][address].value
                _require(value is not None and str(value).strip(),'部分语义证据原格为空')
    finally:book.close()
    for receipt in mapping.get('request_receipts',[]):
        _require(hashlib.sha256(Path(receipt['path']).read_bytes()).hexdigest()==receipt['sha256'],'识别回执已变化')
    packet_facts=[];packet_excluded=[];packet_owners={};packet_partial=[]
    if mapping.get('reused_mapping'):
        reused=mapping['reused_mapping']
        _require(file_hash(reused['path'])==reused['sha256'],'复用的原语义映射已变化')
        if reused.get('selection') == 'unchanged_definitions':
            from .统一语义 import compatible_classification
            original=json.loads(Path(reused['path']).read_text(encoding='utf-8-sig'),object_pairs_hook=_pairs,parse_constant=_invalid_number)
            previous=load_fact_mapping(reused['path'],read_standard_v2(original['gold_path']),ancestors)
            selected=compatible_classification(previous,standard)
        else:
            previous=load_fact_mapping(reused['path'],standard,ancestors)
            selected=previous
        _require(previous['source_hash']==mapping['source_hash'] and previous.get('source_context',[])==mapping.get('source_context',[]),
                 '复用映射的原件或上下文不一致')
        packet_facts.extend(selected['facts']);packet_excluded.extend(selected.get('excluded',[]));packet_owners.update(selected['owners'])
    for packet in mapping.get('packet_mappings',[]):
        _require(hashlib.sha256(Path(packet['path']).read_bytes()).hexdigest()==packet['sha256'],'识别任务包映射已变化')
        content=json.loads(Path(packet['path']).read_text(encoding='utf-8-sig'),object_pairs_hook=_pairs,parse_constant=_invalid_number)
        _require(content['source_hash']==mapping['source_hash'],'识别任务包的原件不符')
        # 历史识别包保留真实版本；只有引用的指标及维度未改义时才能接续。
        if content['gold_hash'] != mapping['gold_hash']:
            ensure_standard_compatible(content, standard)
        if 'source_context' in content:
            _require(content['source_context']==mapping.get('source_context',[]),'识别任务包与整份映射的语义说明不符')
        verify_classification(content)
        packet_facts.extend(content['facts']);packet_owners.update(content['owners'])
        packet_excluded.extend({'sheet':content['sheet'],**row} for row in content.get('excluded',[]))
        packet_partial.extend({'sheet':content['sheet'],**row} for row in content.get('unresolved',[]) if row.get('partial_semantics'))
    if 'packet_mappings' in mapping:
        _require(mapping['facts']==packet_facts and mapping['owners']==packet_owners and mapping.get('excluded',[])==packet_excluded,
                 '整份映射中的语义或覆盖分类与已保存识别任务包不符')
        _require([row for row in mapping.get('unresolved',[]) if row.get('partial_semantics')]==packet_partial,
                 '保存的部分语义与两轮识别不符：整份映射未保持已验证任务包的结果')
    if 'coverage' in mapping:
        from .表格 import read_workbook,materialize_candidate_blanks
        from .统一识别 import workbook_coverage
        snapshot=read_workbook(source,context=mapping.get('source_context',[]));materialize_candidate_blanks(snapshot)
        coverage=workbook_coverage(mapping,snapshot)
        _require(mapping['coverage']==coverage and mapping.get('complete')==coverage['complete'],'整表覆盖声明与原件不符')
    _require(hashlib.sha256(source.read_bytes()).hexdigest()==mapping['source_hash'],'核验期间原件已变化')
    _require(path.read_bytes()==raw,'核验期间映射已变化')
    mapping['mapping_path']=str(path); mapping['mapping_hash']=hashlib.sha256(raw).hexdigest()
    return mapping


def verify_fact_update_plan(plan, target, source, standard):
    """写入前重新按语义取数，并核对全部可执行内容。"""
    _require(plan.get('schema')=='附注更新清单-v2' and plan.get('complete') and not plan.get('issues'),
             '更新清单尚未确认完整')
    _require(plan.get('gold_hash')==standard['sha256'],'更新清单的标准版本不符')
    for mapping in (target,source):
        _require(plan['inputs'].get(mapping['mapping_path'])==mapping['mapping_hash'],'更新清单未绑定当前映射版本')
    original_target=target
    scope=plan.get('validation_scope')
    if scope:
        _require(isinstance(scope.get('label'),str) and scope['label'].strip(),'局部验证须明确名称')
        _require(not plan.get('structure'),'局部验证不隐式执行结构增减')
        target=select_mapping_scope(target,scope['target']);source=select_mapping_scope(source,scope['source'])
    for mapping in (target,source):
        _require(mapping.get('scope_coverage' if scope else 'coverage',{}).get('complete'),'所选范围语义尚未完整')
    if plan.get('structure'):
        from .统一结构 import load_structural_target
        target=load_structural_target(plan['structure'],target,source,standard)
        _require(plan['inputs'].get(target['mapping_path'])==target['mapping_hash'],'清单未绑定调整后独立语义映射')
    expected_operations=plan.get('structure',{}).get('layout',{}).get('operations',[])
    _require(plan.get('word_operations',[])==expected_operations,'Word结构变更与已核验语义布局不符')
    fresh=build_fact_update_plan(target,source,standard,equivalence_proofs=plan.get('equivalence_proofs',[]),update_purpose=plan.get('update_purpose'))
    from .统一未采用来源 import validate_unused_source_reviews
    unused=validate_unused_source_reviews(plan.get('unused_source_reviews',[]),target,source,standard,fresh)
    _require(plan.get('unused_sources',[])==unused,'未采用来源与原文及两轮复核依据不一致')
    fresh['unused_sources']=unused
    unused_ids={row['source_id'] for row in unused}
    fresh['unplaced_sources']=[row for row in fresh['unplaced_sources'] if row['id'] not in unused_ids]
    _require(not fresh['issues'] and not fresh['unplaced_sources'],'重新勾稽仍有未处理事项')
    for row in fresh['updates']:
        definition=standard['metric'][row['metric_id']]['definition']
        row['semantic_field']=' / '.join(definition[key] for key in ('economic_object','measure','measurement_basis'))
    for key in ('updates','retained_blanks','operations','labels','unplaced_sources','source_a_hash','source_b_hash'):
        _require(plan.get(key)==fresh.get(key),'清单与重新语义取数不一致：'+key)
    fresh['_working_target']=original_target if scope else target
    return fresh


def derive_fact_mapping(parent, source_path, plan, plan_path, word_link=None):
    """原格语义不变时沿用已识别含义，原值必须与实际更新结果逐格一致。"""
    from .表格 import read_workbook,materialize_candidate_blanks
    from .统一识别 import workbook_coverage
    _require(not plan.get('operations') and not plan.get('labels'),'布局变化须先迁移逐格语义，不能沿用原位置')
    path=Path(source_path).resolve(); digest=hashlib.sha256(path.read_bytes()).hexdigest()
    result=deepcopy(parent)
    for key in ('packet_mappings','request_receipts','mapping_hash','mapping_path','a_link','review_path',
                'calculation_evidence','formula_values','calculation_issues','_validated_standard_hash','reused_mapping',
                'classification_rounds','equivalence_proofs','external_review_v2'):
        result.pop(key,None)
    saved_plan=json.loads(Path(plan_path).read_text(encoding='utf-8-sig'),object_pairs_hook=_pairs,parse_constant=_invalid_number)
    if saved_plan.get('update_purpose'):
        selected=select_mapping_scope(parent,saved_plan['validation_scope']['target']) if saved_plan.get('validation_scope') else parent
        projected=target_for_purpose(selected,saved_plan['update_purpose'])
        dimensions={row['id']:row['dimensions'] for row in projected['facts'] if row['record_type']=='metric_fact'}
        from .统一语义 import definition_hash
        evidence='本次更新用途：'+json.dumps(saved_plan['update_purpose'],ensure_ascii=False,sort_keys=True)
        result.setdefault('source_context',[]).append(evidence)
        for row in result['facts']:
            if row['id'] in dimensions:
                if row['dimensions']['entity']!=dimensions[row['id']]['entity']:
                    row['dimension_evidence']['entity']=['@context:'+definition_hash(evidence)]
                row['dimensions']=deepcopy(dimensions[row['id']])
                result['owners'][row['owner_reference']]=deepcopy(row['dimensions'])
    result.update(source_path=str(path),source_hash=digest,original_path=str(path),original_hash=digest,carrier='A',usage={},
                  gold_path=saved_plan['gold_path'],gold_hash=plan['gold_hash'],
                  derivation={'plan_path':str(Path(plan_path).resolve()),'plan_hash':hashlib.sha256(Path(plan_path).read_bytes()).hexdigest()})
    if word_link:result['a_link']=str(Path(word_link).resolve())
    updates={(row['sheet'],row['cell']):row for row in plan['updates']}
    # openpyxl 保存会把空字符串格归一化为 None；两种表示同为空白，与写入回读校验一致。
    def _same_value(actual,wanted):return actual==wanted or (wanted=="" and actual is None)
    book=load_workbook(path,data_only=False)
    try:
        for row in result['facts']:
            if row['record_type']!='metric_fact':
                row['source_reference']['file_hash']=digest
                continue
            sheet,cell=_location(row); original=row['raw_value']
            expected=updates.get((sheet,cell),{}).get('value',original)
            value=book[sheet][cell].value
            _require(_same_value(value,expected),'派生映射对应的更新值与实际工作簿不符：'+sheet+'!'+cell)
            row['source_reference']['file_hash']=digest
            row['raw_value']=value
            # 空白语义与 validate_fact 一致：None 或纯空白字符串都算空白，不因全角空格误判为实际值。
            row['raw_value_state']='blank' if value is None or isinstance(value,str) and not value.strip() else ('formula' if book[sheet][cell].data_type=='f' else 'value')
    finally:book.close()
    snapshot=read_workbook(path,context=result.get('source_context',[]));materialize_candidate_blanks(snapshot)
    previous=read_workbook(parent['source_path'],context=parent.get('source_context',[]));materialize_candidate_blanks(previous)
    _require([s['name'] for s in snapshot['sheets']]==[s['name'] for s in previous['sheets']],'派生表格的工作表结构已变化')
    for before,after in zip(previous['sheets'],snapshot['sheets']):
        for key in ('merges','hidden','hidden_rows','hidden_columns'):
            _require(before[key]==after[key],'派生表格的布局已变化：'+key)
        _require(set(before['cells'])==set(after['cells']),'派生表格的单元格范围已变化')
        for address,cell in before['cells'].items():
            expected=updates.get((before['name'],address),{}).get('value',cell['value'])
            _require(_same_value(after['cells'][address]['value'],expected),'派生表格存在清单外的改动：'+before['name']+'!'+address)
    result['coverage']=workbook_coverage(result,snapshot);result['complete']=result['coverage']['complete']
    return result


def select_mapping_scope(mapping, ranges):
    """仅生成局部执行视图；原映射及其整份覆盖声明始终保留。"""
    from openpyxl.utils.cell import range_boundaries, coordinate_to_tuple
    _require(isinstance(ranges,list) and ranges, '局部验证须明确表格范围')
    bounds=[]
    for row in ranges:
        _require(set(row)=={'sheet','range'} and isinstance(row['sheet'],str), '局部验证范围格式不符')
        left,top,right,bottom=range_boundaries(row['range'])
        _require(all(type(v) is int and v>0 for v in (left,top,right,bottom)) and left<=right and top<=bottom, '局部范围无效')
        bounds.append((row['sheet'],left,top,right,bottom))
    def selected(location):
        sheet,_,address=location.rpartition('!');r,c=coordinate_to_tuple(address)
        return any(sheet==s and l<=c<=rr and t<=r<=b for s,l,t,rr,b in bounds)
    expected=mapping.get('coverage',{}).get('expected_cells')
    if expected is None:
        from .表格 import read_workbook,materialize_candidate_blanks
        snapshot=read_workbook(mapping['source_path'],context=mapping.get('source_context',[]));materialize_candidate_blanks(snapshot)
        expected=[sheet['name']+'!'+cell for sheet in snapshot['sheets'] for cell in sheet['cells']]
    expected={location for location in expected if selected(location)}
    _require(expected,'所选范围没有候选格')
    view=deepcopy(mapping)
    view['facts']=[row for row in view['facts'] if selected(row['source_reference']['location'])]
    assigned=[row['source_reference']['location'] for row in view['facts'] if row['record_type']=='metric_fact']
    for kind in ('excluded','unresolved'):
        view[kind]=[row for row in view.get(kind,[]) if selected(row.get('sheet',mapping.get('sheet',''))+'!'+row['cell'])]
        assigned.extend(row.get('sheet',mapping.get('sheet',''))+'!'+row['cell'] for row in view[kind])
    _require(len(assigned)==len(set(assigned)) and set(assigned)==expected,'局部验证范围存在遗漏或重复分类')
    view['scope_coverage']={'scope':'local_validation','ranges':deepcopy(ranges),'expected_count':len(expected),
        'financial_cells':len(view['facts']),'unresolved_cells':len(view['unresolved']),
        'outside_unresolved_count':len(mapping.get('unresolved',[]))-len(view['unresolved']),
        'complete':not view['unresolved']}
    return view


def target_for_purpose(mapping, purpose):
    """本次目标主体投影，与原始识别事实分开；不推断期间及金额。"""
    if not purpose:return mapping
    _require(set(purpose)=={'entity','reason'} and isinstance(purpose['reason'],str) and purpose['reason'].strip(), '本次更新用途须有明确依据')
    entity=purpose['entity']
    _require(isinstance(entity,dict) and set(entity)=={'from','to'} and all(isinstance(v,str) and v.strip() for v in entity.values()), '本次用途须明确原主体及目标主体')
    view=deepcopy(mapping)
    for row in view['facts']:
        if row['dimensions'].get('entity')==entity['from']:
            row['dimensions']['entity']=entity['to']
            if row.get('owner_reference'):view['owners'][row['owner_reference']]['entity']=entity['to']
    return view


def build_fact_update_plan(target_mapping, source_mapping, standard, equivalence_proofs=(), update_purpose=None):
    """输入为已识别事实。文件级执行者仍须校验原件和映射版本后再写入。"""
    for mapping in (target_mapping, source_mapping):
        _check_mapping(mapping, standard)
    original_dimensions={row['id']:deepcopy(row['dimensions']) for row in target_mapping['facts']}
    target_mapping=target_for_purpose(target_mapping,update_purpose)
    from .同义复核 import approved_mapping_pairs
    approved=approved_mapping_pairs(equivalence_proofs,target_mapping,source_mapping,standard)
    plan = {'schema_version': 2, 'gold_hash': standard['sha256'],
            'source_a_hash': target_mapping['source_hash'], 'source_b_hash': source_mapping['source_hash'],
            'updates': [], 'issues': [], 'unplaced_sources': [], 'retained_blanks': [], 'operations': [], 'labels': []}
    candidates = defaultdict(list); consumed = set()
    source_facts = [row for row in source_mapping['facts'] if row['record_type'] == 'metric_fact']
    for row in source_facts:
        try:
            candidates[semantic_key(row, standard)].append(_effective_source(row,source_mapping,standard))
        except ValueError as error:
            sheet, cell = _location(row)
            plan['issues'].append({'kind': 'unresolved_source', 'sheet': sheet, 'cell': cell, 'reason': str(error)})
    for mapping, carrier in ((target_mapping, '附注表格'), (source_mapping, '更新数据')):
        for row in mapping.get('unresolved', []):
            plan['issues'].append({'kind': 'unresolved_mapping', 'carrier': carrier, 'record': deepcopy(row)})
    for target in target_mapping['facts']:
        if target['record_type'] != 'metric_fact':
            continue
        sheet, cell = _location(target)
        issue = {'sheet': sheet, 'cell': cell, 'metric_id': target['metric_id']}
        try:
            matches = candidates.get(semantic_key(target, standard), [])
            matches=list(matches)
            matches.extend(_effective_source(row,source_mapping,standard) for row in source_facts
                           if (target['id'],row['id']) in approved and row['id'] not in {item['id'] for item in matches})
        except ValueError as error:
            plan['issues'].append({**issue, 'kind': 'unresolved_target', 'reason': str(error)})
            continue
        explicit = [row for row in matches if row['raw_value_state'] == 'value']
        if any(row['raw_value_state'] == 'formula' for row in matches):
            plan['issues'].append({**issue, 'kind': 'unevaluated_source_formula', 'reason': '来源含尚未取得实际计算值的公式，须先核验计算值'})
            continue
        if not explicit:
            if target['raw_value_state'] == 'blank':
                plan['retained_blanks'].append({**issue, 'reason': '来源没有明确新值，保留原空白，不计作已更新'})
                consumed.update(row['id'] for row in matches)
            else:
                plan['issues'].append({**issue, 'kind': 'source_unknown_blank' if matches else 'missing_source',
                                       'reason': '来源空白不能清空原金额' if matches else '没有相同指标和实际维度的来源'})
            continue
        def converted(row):
            if (target['id'],row['id']) not in approved:return value_for_target(row,target,standard)
            projected=deepcopy(row);projected['dimensions']=deepcopy(target['dimensions'])
            if 'unit_scale' in row['dimensions']:projected['dimensions']['unit_scale']=row['dimensions']['unit_scale']
            return value_for_target(projected,target,standard)
        try:
            values = [converted(row) for row in explicit]
        except ValueError as error:
            plan['issues'].append({**issue, 'kind': 'source_value_needs_review', 'reason': str(error),
                                   'source_locations': [row['source_reference']['location'] for row in explicit]})
            consumed.update(row['id'] for row in matches)
            continue
        if any(value != values[0] for value in values[1:]):
            plan['issues'].append({**issue, 'kind': 'conflicting_sources', 'reason': '相同业务身份的多个来源金额冲突，未选取或求和',
                                   'source_locations': [row['source_reference']['location'] for row in explicit]})
            continue
        sources = []
        for row in explicit:
            source_sheet, address = _location(row)
            source={'sheet': source_sheet, 'cell': address, 'value': row['raw_value'], 'fact_id': row['id']}
            original=next(item for item in source_facts if item['id']==row['id'])
            if original['raw_value_state']=='formula':source['formula']=original['raw_value']
            if (target['id'],row['id']) in approved:source['equivalence_pair']=approved[(target['id'],row['id'])]
            sources.append(source)
        plan['updates'].append({**issue, 'value': _plain(values[0]), 'old_value': target['raw_value'],
                                'dimensions': deepcopy(target['dimensions']), 'sources': sources,
                                **({'original_dimensions':original_dimensions[target['id']]} if update_purpose else {}),
                                'source_blank_locations': [row['source_reference']['location'] for row in matches if row not in explicit]})
        consumed.update(row['id'] for row in matches)
    plan['unplaced_sources'] = [deepcopy(row) for row in source_facts if row['id'] not in consumed]
    return plan
