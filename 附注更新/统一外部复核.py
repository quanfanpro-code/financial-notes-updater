# coding: utf-8
"""导入人工或高级AI的逐格语义结论；原值始终从原工作簿读取。"""
from copy import deepcopy
from datetime import datetime
import json
from pathlib import Path
import uuid
from openpyxl import Workbook
from .金标准 import _pairs, _invalid_number
from .表格 import file_hash, read_workbook, materialize_candidate_blanks
from .统一语义 import read_standard_v2, definition_hash, _require, validate_facts
from .统一识别 import validate_response, workbook_coverage, standard_review_material

SCHEMA='附注外部语义复核-v2'
PREVIEW_SCHEMA='附注外部复核预览-v2'
METHOD='external_semantic_review_v2'


def _read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'),object_pairs_hook=_pairs,parse_constant=_invalid_number)


def _save(path,value):
    with Path(path).open('x',encoding='utf-8-sig') as stream:json.dump(value,stream,ensure_ascii=False,indent=2)
    return str(path)


def _load(path,gold_path=None,ancestors=frozenset()):
    from .统一勾稽 import load_fact_mapping
    raw=_read(path);selected=Path(gold_path or raw['gold_path']).resolve();standard=read_standard_v2(selected)
    standard['path']=str(selected)
    record=load_fact_mapping(path,standard,ancestors)
    _require(record.get('carrier') in {'A','B'},'请选择整份附注表格或更新数据的独立语义映射')
    snapshot=read_workbook(record['source_path'],context=record.get('source_context',[]))
    materialize_candidate_blanks(snapshot)
    _require(snapshot['sha256']==record['source_hash'],'原工作簿已变化')
    workbook_coverage(record,snapshot)
    return record,snapshot,standard


def export_review_template(mapping_path,output_path,gold_path=None,*,include_confirmed=False):
    record,snapshot,standard=_load(mapping_path,gold_path)
    def item(row,confirmed=False):
        sheet,cell=row['source_reference']['location'].rpartition('!')[::2] if confirmed else (row['sheet'],row['cell'])
        known=row if confirmed else row.get('partial_semantics',{})
        evidence={k:[r[len(sheet)+1:] if r.startswith(sheet+'!') else r for r in refs] for k,refs in known.get('dimension_evidence',{}).items()}
        result={'sheet':sheet,'cell':cell,'metric_id':known.get('metric_id',''),
                'dimensions':deepcopy(known.get('dimensions',{})),'dimension_evidence':evidence,'reason':''}
        if confirmed:result['replace_fact_hash']=definition_hash(row)
        return result
    value={'schema':SCHEMA,'carrier':record['carrier'],'mapping_sha256':file_hash(mapping_path),
        'source_hash':record['source_hash'],'gold_hash':standard['sha256'],
        'reviewer':{'name':'','method':'','reviewed_at':''},'items':[item(row) for row in record['unresolved']],
        'instructions':['填写真实审核者、human或advanced_ai及审核时间。只在items保留本次实际核实的格，填写依据；未核实项从本次意见中移出，原映射会继续保留。',
            'metric_id为槽位，dimensions为各维度实际取值；每项dimension_evidence引用同表原文字格或适用的@context说明，不以财务值作证。',
            '不能填写或修改raw_value；程序回读原格。新槽位须先通过管理金标准保存新版本，再按该版本导出复核格式。',
            '更正已确认格时，将confirmed_mapping_options中的对应项复制到items，保留replace_fact_hash并填写实际意见。'],
        'reference':{'source_path':record['source_path'],'document_context':record.get('source_context',[]),
            'metrics':list(standard['metric'].values()),'dimension_definitions':list(standard['dimension'].values()),'sheets':snapshot['sheets']}}
    if include_confirmed:
        value['confirmed_mapping_options']=[item(row,True) for row in record['facts'] if row['record_type']=='metric_fact']
        value['excluded_mapping_options']=[{**item(row),'replace_exclusion_hash':definition_hash(row)} for row in record['excluded']]
        value['instructions'].append('纠正误排除格时，将excluded_mapping_options中的对应项复制到items，保留replace_exclusion_hash，按原文填写槽位、维度及依据；不能按原值或符号直接改判。')
    return _save(Path(output_path).resolve(),value)


def _prepare_review(mapping_path,opinion_path,gold_path=None,*,ancestors=frozenset()):
    mapping=Path(mapping_path).resolve();opinion_file=Path(opinion_path).resolve()
    mapping_hash=file_hash(mapping);opinion_hash=file_hash(opinion_file)
    record,snapshot,standard=_load(mapping,gold_path,ancestors);opinion=_read(opinion_file)
    _require(opinion.get('schema')==SCHEMA,'请选择新版外部语义复核意见')
    for key,expected in [('carrier',record['carrier']),('mapping_sha256',mapping_hash),('source_hash',record['source_hash']),('gold_hash',standard['sha256'])]:
        _require(opinion.get(key)==expected,'外部意见绑定不一致：'+key)
    reviewer=opinion.get('reviewer')
    _require(isinstance(reviewer,dict) and isinstance(reviewer.get('name'),str) and reviewer['name'].strip() and reviewer.get('method') in {'human','advanced_ai'},'请填写实际审核者及方式')
    try:datetime.fromisoformat(reviewer['reviewed_at'])
    except (ValueError,TypeError,KeyError):raise ValueError('请填写实际审核时间') from None
    items=opinion.get('items');_require(isinstance(items,list) and items,'没有实际复核意见')
    pending={(r['sheet'],r['cell']):r for r in record['unresolved']}
    confirmed={tuple(r['source_reference']['location'].rpartition('!')[::2]):r for r in record['facts'] if r['record_type']=='metric_fact'}
    excluded={(r['sheet'],r['cell']):r for r in record['excluded']}
    sheets={s['name']:s for s in snapshot['sheets']};changes=[];seen=set()
    fields={'sheet','cell','metric_id','dimensions','dimension_evidence','reason'}
    for item in items:
        _require(isinstance(item,dict) and fields<=set(item)<=fields|{'replace_fact_hash','replace_exclusion_hash'},'意见只能包含槽位、维度、原文依据及明确更正绑定，不能修改数值')
        pos=(item['sheet'],item['cell']);previous=confirmed.get(pos)
        previous_exclusion=excluded.get(pos)
        _require(pos not in seen and (pos in pending or previous is not None or previous_exclusion is not None),'意见引用重复或不属于原映射的格')
        seen.add(pos)
        if previous is not None:_require(item.get('replace_fact_hash')==definition_hash(previous),'更正已确认语义须绑定原事实')
        else:_require('replace_fact_hash' not in item,'此格没有可更正的原事实')
        if previous_exclusion is not None:_require(item.get('replace_exclusion_hash')==definition_hash(previous_exclusion),'更正排除结论须绑定原排除记录')
        else:_require('replace_exclusion_hash' not in item,'此格没有可更正的原排除记录')
        sheet=sheets[pos[0]]
        packet={'source_path':record['source_path'],'source_hash':record['source_hash'],'sheet':pos[0],
                'cells':sheet['cells'],'candidate_cells':[pos[1]],'document_context':record.get('source_context',[])}
        response={'kind':'metric_fact',**{k:deepcopy(item[k]) for k in ('metric_id','dimensions','dimension_evidence','reason')}}
        fact=validate_response({'cells':{pos[1]:response}},packet,standard)[pos[1]]['fact']
        fact.update(id=fact['source_reference']['location'],owner_reference=definition_hash(fact['dimensions']),
                    recognition_provenance={'method':METHOD,'record_id':definition_hash({'opinion_hash':opinion_hash,'reviewer':reviewer,'item':item})})
        changes.append({'sheet':pos[0],'cell':pos[1],'mapping':fact,'previous_mapping':deepcopy(previous),
                        **({'previous_exclusion':deepcopy(previous_exclusion)} if previous_exclusion is not None else {}),
                        'previous_reason':pending.get(pos,previous_exclusion or {}).get('reason','原已确认业务格'),'reason':item['reason']})
    _require(file_hash(mapping)==mapping_hash and file_hash(opinion_file)==opinion_hash,'预览期间输入发生变化')
    preview={'schema':PREVIEW_SCHEMA,'mapping_path':str(mapping),'mapping_sha256':mapping_hash,
        'opinion_path':str(opinion_file),'opinion_sha256':opinion_hash,'carrier':record['carrier'],
        'source_hash':record['source_hash'],'gold_path':standard['path'],'gold_hash':standard['sha256'],
        'reviewer':deepcopy(reviewer),'count':len(changes),'remaining_pending':len(pending)-sum(pos in pending for pos in seen),'changes':changes}
    return preview,record,snapshot,standard


def preview_external_review(mapping_path,opinion_path,gold_path=None):
    return _prepare_review(mapping_path,opinion_path,gold_path)[0]


def _combine(parent,preview):
    result=deepcopy(parent);positions={(r['sheet'],r['cell']) for r in preview['changes']}
    result['facts']=[r for r in result['facts'] if r['record_type']!='metric_fact' or tuple(r['source_reference']['location'].rpartition('!')[::2]) not in positions]
    result['facts'].extend(deepcopy(r['mapping']) for r in preview['changes'])
    result['owners'].update({r['mapping']['owner_reference']:deepcopy(r['mapping']['dimensions']) for r in preview['changes']})
    result['unresolved']=[r for r in result['unresolved'] if (r['sheet'],r['cell']) not in positions]
    result['excluded']=[r for r in result['excluded'] if (r['sheet'],r['cell']) not in positions]
    result.update(gold_path=preview['gold_path'],gold_hash=preview['gold_hash'])
    for key in ('packet_mappings','request_receipts','reused_mapping','equivalence_proofs','classification_rounds','derivation',
                'external_review_v2','_validated_standard_hash','mapping_path','mapping_hash','review_path','standard_review_path',
                'calculation_evidence','formula_values','calculation_issues'):
        result.pop(key,None)
    return result


def validate_external_mapping(mapping,standard,ancestors):
    proof=mapping['external_review_v2'];_require(file_hash(proof['path'])==proof['sha256'],'外部复核采纳依据已变化')
    receipt=_read(proof['path'])
    _require(receipt.get('schema')=='附注外部复核采纳-v2' and receipt.get('confirm_method') and receipt.get('confirmed_at'),'缺少实际采纳记录')
    _require(file_hash(receipt['gold_path'])==receipt['gold_hash'],'采纳所用金标准已变化')
    preview,parent,_,_=_prepare_review(receipt['mapping_path'],receipt['opinion_path'],receipt['gold_path'],ancestors=ancestors)
    for key in ('mapping_sha256','opinion_sha256','source_hash','gold_hash','reviewer','changes'):
        _require(receipt[key]==preview[key],'复核依据与实际原件及意见不一致：'+key)
    expected=_combine(parent,preview)
    for key in ('facts','owners','unresolved','excluded','source_context','source_path','source_hash','gold_hash'):
        _require(mapping.get(key)==expected.get(key),'复核后映射与实际采纳意见不一致：'+key)
    validate_facts(mapping['facts'],standard,mapping['owners'])


def apply_external_review(preview,output_dir,*,confirmed=False,confirm_method=''):
    _require(confirmed is True and isinstance(confirm_method,str) and confirm_method.strip(),'请先预览并明确采纳外部意见')
    _require(preview.get('schema')==PREVIEW_SCHEMA,'缺少新版复核预览')
    fresh,parent,snapshot,standard=_prepare_review(preview['mapping_path'],preview['opinion_path'],preview['gold_path'])
    _require(fresh==preview,'预览后输入变化，请重新预览')
    run=Path(output_dir).resolve()/('导入外部复核_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'_'+uuid.uuid4().hex[:6]);run.mkdir(parents=True)
    receipt={k:deepcopy(fresh[k]) for k in ('mapping_sha256','opinion_sha256','source_hash','gold_path','gold_hash','reviewer','changes')}
    for key,name in [('mapping','复核所依据的原映射.json'),('opinion','外部复核原始意见.json')]:
        dest=run/name
        with dest.open('xb') as stream:stream.write(Path(fresh[key+'_path']).read_bytes())
        _require(file_hash(dest)==fresh[key+'_sha256'],'复核输入副本校验失败');receipt[key+'_path']=str(dest)
    receipt.update(schema='附注外部复核采纳-v2',confirm_method=confirm_method,confirmed_at=datetime.now().astimezone().isoformat())
    receipt_path=_save(run/'外部复核确认记录.json',receipt)
    result=_combine(parent,fresh);result['external_review_v2']={'path':receipt_path,'sha256':file_hash(receipt_path)}
    result['coverage']=workbook_coverage(result,snapshot);result['complete']=result['coverage']['complete']
    from .统一勾稽 import load_fact_mapping,calculate_mapping_formulas
    try:calculate_mapping_formulas(result,run/'公式计算依据')
    except Exception as error:result['calculation_issues']=[{'reason':str(error)}]
    prefix={'A':'附注表格','B':'更新数据'}[result['carrier']];key=result['carrier'].lower()
    mapping_path=run/(prefix+'语义映射.json');result['mapping_path']=str(mapping_path)
    review_path=run/(prefix+'语义识别核对.xlsx');result['review_path']=str(review_path)
    files={key+'_mapping':str(mapping_path),key+'_review':str(review_path)}
    if result['unresolved']:
        pending=run/(prefix+'待处理语义与金标准补充材料.json');result['standard_review_path']=str(pending)
        files[key+'_standard_review']=str(pending)
    _save(mapping_path,result)
    load_fact_mapping(mapping_path,standard)
    if result['unresolved']:_save(pending,standard_review_material(result,standard))
    book=Workbook();sheet=book.active;sheet.title='逐格语义识别'
    sheet.append(['来源位置','识别结果','指标槽位','实际维度','原值或原公式','依据'])
    for row in result['facts']:
        if row['record_type']=='metric_fact':sheet.append([row['source_reference']['location'],'已识别',row['metric_id'],json.dumps(row['dimensions'],ensure_ascii=False),row['raw_value'],json.dumps(row['dimension_evidence'],ensure_ascii=False)])
    for kind,label in [('excluded','标题或说明等'),('unresolved','待核实')]:
        for row in result[kind]:sheet.append([row['sheet']+'!'+row['cell'],label,None,None,None,row['reason']])
    for row in sheet:
        for cell in row:
            if isinstance(cell.value,str):cell.data_type='s'
    sheet.freeze_panes='A2'
    with review_path.open('xb') as stream:book.save(stream)
    book.close()
    return {'status':'complete' if result['complete'] and not result.get('calculation_issues') else 'partial',
            'message':f"已采纳{fresh['count']}格语义复核意见，仍有{len(result['unresolved'])}格待核实；原始数值与金标准未修改。",
            'files':files,'accepted_count':fresh['count']}
