# coding: utf-8
"""把已识别子表的行列安排展开成既有更新协议，复用现有布局和取数校验。"""
from copy import deepcopy
import json
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter, range_boundaries
from .表格 import _name_range, read_workbook
from .统一语义 import _require, definition_hash
from .模板布局 import project_layout
from .统一勾稽 import _template_cell_records, build_fact_update_plan
from .统一证据 import context_catalog


def compile_template_purpose(a,b,standard,spec,*,target_locations=None):
    _require(isinstance(spec,dict) and isinstance(spec.get('reason'),str) and spec['reason'].strip(), '模板安排缺少依据')
    _require(isinstance(spec.get('tables'),list),'模板安排须列出子表')
    book=load_workbook(a['source_path']);areas={};operations=[]
    try:
        for table in spec['tables']:
            name=table['range_name'];_require(name not in areas,'同一表重复安排')
            offsets=table.get('column_row_offsets',{})
            _require(set(offsets)<=set(table['columns']) and all(type(v) is int for v in offsets.values()),'列区块偏移无效')
            _require(all(type(v) is int for v in table.get('label_row_offsets',[])),'名称区块偏移无效')
            row_offsets=table.get('target_row_offsets')
            _require(row_offsets is None or len(row_offsets)==len(table['source_rows'])==len(set(row_offsets))
                     and all(type(v) is int and v>=0 for v in row_offsets),'目标明细行偏移无效')
            _,sheet,box=_name_range(book,name);areas[name]=(sheet.title,box)
        for table in sorted(spec['tables'],key=lambda t:(areas[t['range_name']][0],areas[t['range_name']][1][1]),reverse=True):
            rows=table['source_rows'];_require(rows and len(rows)==len(set(rows)) and all(type(r) is int and r>0 for r in rows),'来源明细行无效')
            if table.get('operations'):
                for item in table['operations']:
                    operation={'range_name':table['range_name'],'count':1,'reason':spec['reason'],**item}
                    if operation['action']=='delete':operation.update(source_evidence=table['source_evidence'],coverage_reason=table['coverage_reason'])
                    operations.append(operation)
                continue
            old=table['old_detail_count'];new=len(rows);offset=table['first_row_offset']
            _require(type(old) is int and old>=0 and type(offset) is int and offset>=0,'原明细范围无效')
            for _ in range(abs(new-old)):
                operation={'range_name':table['range_name'],'axis':'row','action':'insert' if new>old else 'delete',
                           'index':offset+min(old,new),'count':1,'template_index':offset,'reason':spec['reason']}
                if new<old:operation.update(source_evidence=table['source_evidence'],coverage_reason='按原文完整来源名单替换明细')
                operations.append(operation)
    finally:book.close()
    selected_sources={item['source'] for item in spec.get('bindings',[])}
    for table in spec['tables']:
        selected_sources.update(f"{table['source_sheet']}!{column}{row+table.get('column_row_offsets',{}).get(target_column,0)}"
            for row in [*table['source_rows'],*[item['source_row'] for item in table.get('totals',[])]]
            for target_column,column in table['columns'].items())
    _require(selected_sources or spec.get('unprovided') and not operations,'模板安排没有已识别来源或逐项缺字段依据')
    view=project_layout(a,b,standard,operations,source_locations=selected_sources or None)
    positions=view['_template_layout']['positions'];targets=_template_cell_records(view)
    catalog=context_catalog(view.get('source_context',[]))
    purpose={'reason':spec['reason'],'operations':operations,'bindings':[],'labels':[],'unprovided':[]}

    def missing_entry(item,target):
        item=deepcopy(item);item['target']=target;sheet,cell=target.rsplit('!',1);row=coordinate_to_tuple(cell)[0]
        for key,refs in item.get('dimension_evidence',{}).items():
            moved=[f'{sheet}!{ref[5:]}{row}' if ref.startswith('@row:') else positions.get(ref,ref) for ref in refs]
            stale=[ref for ref in moved if ref.startswith('@context:') and ref not in catalog]
            if stale:
                known=targets[target]
                _require(item['dimensions'][key]==known['dimensions'][key],'旧说明对应维度值已改变，须重新核准')
                replacement=known['dimension_evidence'][key]
                _require(replacement and all(not ref.startswith('@context:') or ref in catalog for ref in replacement),'本轮说明依据缺失')
                moved=list(dict.fromkeys([ref for ref in moved if ref not in stale]+replacement))
            item['dimension_evidence'][key]=moved
        return item

    for table in spec['tables']:
        sheet,box=areas[table['range_name']];source_sheet=table['source_sheet'];first=len(purpose['bindings'])
        anchor=f'{sheet}!{get_column_letter(box[0])}{box[1]}'
        top=coordinate_to_tuple(positions[anchor].rsplit('!',1)[1])[0]
        for index,source_row in enumerate(table['source_rows']):
            row=top+(table['target_row_offsets'][index] if 'target_row_offsets' in table else table['first_row_offset']+index)
            for field,key in [('columns','bindings'),('labels','labels')]:
                for target_col,source_col in table[field].items():
                    offset=table.get('column_row_offsets',{}).get(target_col,0) if key=='bindings' else 0
                    item={'target':f'{sheet}!{target_col}{row}','source':f'{source_sheet}!{source_col}{source_row+offset}'}
                    if key=='labels' and table.get('column_row_offsets'):
                        item['source_aliases']=[f'{source_sheet}!{source_col}{source_row+v}' for v in
                            sorted(set(table.get('label_row_offsets',table['column_row_offsets'].values()))) if v]
                    purpose[key].append(item)
            for column,items in table.get('unprovided',{}).items():
                _require(len(items)==len(table['source_rows']),'每个缺字段须对应一个明细行')
                purpose['unprovided'].append(missing_entry(items[index],f'{sheet}!{column}{row}'))
        for total in table.get('totals',[]):
            for target_col,source_col in table['columns'].items():
                purpose['bindings'].append({'target':positions[f"{sheet}!{target_col}{total['target_row']}"],
                    'source':f"{source_sheet}!{source_col}{total['source_row']+table.get('column_row_offsets',{}).get(target_col,0)}"})
        for label in table.get('period_labels',[]):
            label=deepcopy(label);label['target']=positions[label['target']]
            if 'period_targets' in label:label['period_targets']=[positions[ref] for ref in label['period_targets']]
            elif 'period_source' in label or 'period_sources' in label:
                header_row,header_col=coordinate_to_tuple(label['target'].rsplit('!',1)[1])
                label['period_targets']=[x['target'] for x in purpose['bindings'][first:]
                    if coordinate_to_tuple(x['target'].rsplit('!',1)[1])[1]==header_col
                    and coordinate_to_tuple(x['target'].rsplit('!',1)[1])[0]>header_row]
                _require(label['period_targets'],'期间表头没有对应数据列')
            purpose['labels'].append(label)
    for kind in ('bindings','labels','unprovided'):
        for item in spec.get(kind,[]):
            item=deepcopy(item);item['target']=positions.get(item['target'],item['target'])
            _require(item['target'] is not None,'不能引用本次删除的旧行')
            if kind=='unprovided':item=missing_entry(item,item['target'])
            if kind=='labels' and 'period_targets' in item:item['period_targets']=[positions[ref] for ref in item['period_targets']]
            purpose[kind].append(item)
    _require(len({r['target'] for r in purpose['bindings']})==len(purpose['bindings']),'数据格重复绑定')
    if target_locations is not None:
        layout=view['_template_layout'];original=set(_template_cell_records(a))
        changed=set(layout['removed'])&original
        changed.update(item['template'] for item in layout['added'].values() if item['template'] in original)
        allowed={positions[ref] for ref in target_locations if positions.get(ref) is not None}
        allowed.update(ref for ref,item in layout['added'].items() if item['template'] in target_locations)
        _require(changed<=target_locations and all(row['target'] in allowed for kind in ('bindings','unprovided') for row in purpose[kind]),
                 '科目编排不能修改本科目之外的财务格')
    return purpose


def _layout_material(mapping):
    rows=_template_cell_records(mapping);snapshot=read_workbook(mapping['source_path']);values={};evidence_values={};indexes={};facts=[]
    def reference(value,registry):
        key=(id(registry),definition_hash(value))
        if key not in indexes:
            indexes[key]=str(len(registry));registry[indexes[key]]=deepcopy(value)
        return indexes[key]
    for location,row in rows.items():
        fields={}
        for key,value in row['dimensions'].items():
            fields[key]=reference(value,values)
        facts.append({'location':location,'metric_id':row['metric_id'],'dimensions':fields,
            'dimension_evidence':{key:reference(refs,evidence_values) for key,refs in row['dimension_evidence'].items()}})
    # 已核准目录负责语义路由，只传已定位的表区与身份依据，避免再次发送整份识别材料。
    outline={key:deepcopy(mapping['outline'][key]) for key in ('method','subjects','subtables','context_ranges','unlocated_ranges','complete')
             if key in mapping.get('outline',{})}
    areas=[area for table in outline.get('subtables',[]) for area in table.get('ranges',[])]
    areas.extend(outline.get('context_ranges',[]))
    if not areas:areas.extend(area for entries in snapshot['names'].values() for area in entries)
    boxes=[(area['sheet'],*range_boundaries(area['range'])) for area in areas if isinstance(area,dict) and 'sheet' in area and 'range' in area]
    boxes=[box for box in boxes if all(type(number) is int for number in box[1:])]
    evidence={ref for item in rows.values() for refs in item['dimension_evidence'].values() for ref in refs if not ref.startswith('@context:')}
    def included(sheet,cell):
        if sheet+'!'+cell in evidence:return True
        row,column=coordinate_to_tuple(cell)
        return any(sh==sheet and left<=column<=right and top<=row<=bottom for sh,left,top,right,bottom in boxes)
    return {'source_hash':mapping['source_hash'],'names':snapshot['names'],'outline':outline,
        'dimension_values':values,'evidence_values':evidence_values,'facts':facts,'text':{s['name']+'!'+cell:c['value'] for s in snapshot['sheets']
            for cell,c in s['cells'].items() if s['name']+'!'+cell not in rows and isinstance(c['value'],str)
            and included(s['name'],cell)},
        'merges':{s['name']:s['merges'] for s in snapshot['sheets'] if any(sh==s['name'] for sh,*_ in boxes)}}


TEMPLATE_PROMPT='''A、B身份已经分别核准，不重新解释或比较语义，不输出金额、指标或维度新答案。
dimension_values与evidence_values是已核准实际维度和原文依据的共享清单；facts只引用其编号，没有省略或改变原含义。
依据已有身份及编排，为A模板选择B的具体来源、名单与必要的行数变化。名称、日期与数据同时安排。
返回JSON对象reason、tables，可附bindings、labels、unprovided。tables每项含range_name、source_sheet、source_rows、old_detail_count、first_row_offset、columns、labels；columns及labels为A列字母到B列字母的对应，source_rows是B完整明细行，不含合计。
first_row_offset是A命名范围内首个明细行从0开始的偏移，old_detail_count是原明细行数。totals列表可含target_row及source_row（实际原件绝对行号）。跨分块列可给column_row_offsets和label_row_offsets；非连续目标行可给target_row_offsets。
必要的明确operations沿用现有range_name、axis、action、index、count、template_index协议。删行须有source_evidence原文范围及coverage_reason，不能删掉未解决目标来提高覆盖。
单独bindings为{target:"工作表!格",source:"工作表!格"}；labels同样给target、source，可给既有期间及范围字段。缺字段unprovided逐项保留目标、source_evidence和reason，名单或期间改变时附已确认的dimensions及dimension_evidence，不能把旧对象原值挂到新对象。
已有同槽位不等于任意来源可取数，整体与明细不可混配；利用已核准的实际维度和原文安排结构。来源选不到或有真实歧义则明确报错，不虚构位置或财务内容。'''


def plan_template_purpose(engine,a,b,standard,folder):
    from .统一补充 import _freeze_receipts
    from .A语义阅读 import semantic_groups
    from .表格编排 import in_ranges
    folder=Path(folder);folder.mkdir(parents=True,exist_ok=True)
    target=_layout_material(a);source=_layout_material(b)
    groups=semantic_groups(a.get('outline',{}).get('subtables',[]),{}) or [[]]
    parts=[];covered=set()
    start=len(engine._review_receipts);engine.source_hash=b['source_hash']
    def part(material,facts,subjects):
        value=deepcopy(material);value['facts']=facts
        for field,lookup in [('dimensions','dimension_values'),('dimension_evidence','evidence_values')]:
            wanted={ref for fact in facts for ref in fact[field].values()}
            value[lookup]={ref:item for ref,item in material[lookup].items() if ref in wanted}
        tables=[table for table in material['outline'].get('subtables',[]) if table['subject'] in subjects]
        ranges=[area for table in tables for area in table.get('ranges',[{'sheet':table.get('sheet'),'range':table.get('range')}]) if area.get('sheet') and area.get('range')]
        evidence={ref for refs in value['evidence_values'].values() for ref in refs}
        if facts:value['text']={ref:text for ref,text in material['text'].items() if ref in evidence or in_ranges(*ref.rsplit('!',1),ranges)}
        value['outline']={**value['outline'],'subtables':tables,'subjects':list(subjects)}
        return value
    for group in groups:
        ranges=[{'sheet':table['sheet'],'range':table['range']} for table in group]
        facts=[row for row in target['facts'] if not group or in_ranges(*row['location'].rsplit('!',1),ranges)]
        if not facts:continue
        locations={row['location'] for row in facts};_require(not covered&locations,'已核准科目分组包含重复目标格');covered.update(locations)
        subjects=list(dict.fromkeys(table['subject'] for table in group));metrics={row['metric_id'] for row in facts}
        candidates=[row for row in source['facts'] if row['metric_id'] in metrics]
        payload={'task':'template_update_layout','subjects':subjects,
                 'target':part(target,facts,subjects),'source':part(source,candidates,subjects)}
        def validate(answer):
            compile_template_purpose(a,b,standard,answer,target_locations=locations)
            return answer
        engine.log('沿用已确认科目分组编排：'+('、'.join(subjects) or '当前已识别表'))
        parts.append(engine._request(TEMPLATE_PROMPT,payload,validate))
    _require(covered=={row['location'] for row in target['facts']},'已核准科目目录未覆盖全部模板财务格')
    spec={'reason':'；'.join(dict.fromkeys(row['reason'] for row in parts)),
          **{kind:[item for row in parts for item in row.get(kind,[])] for kind in ('tables','bindings','labels','unprovided')}}
    checked=build_fact_update_plan(a,b,standard,update_purpose=compile_template_purpose(a,b,standard,spec))
    _require(not checked['issues'] and not checked['unplaced_sources'],'整份编排仍有未解决目标或来源')
    purpose=checked['update_purpose']
    report={'schema':'模板更新编排-v1','target_hash':a['source_hash'],'source_hash':b['source_hash'],
            'gold_hash':standard['sha256'],'spec':spec,'purpose':purpose,'evidence':_freeze_receipts(engine,start,folder)}
    path=folder/'模板更新编排.json'
    with path.open('x',encoding='utf-8') as handle:json.dump(report,handle,ensure_ascii=False,indent=2)
    return purpose,{'template_layout_plan':str(path)}
