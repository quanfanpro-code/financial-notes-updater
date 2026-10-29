# coding: utf-8
"""把已识别子表的行列安排展开成既有更新协议，复用现有布局和取数校验。"""
from copy import deepcopy
import json
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.cell.cell import ERROR_CODES
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
            if 'operations' in table:
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
            for target_column,column in table.get('row_columns',{}).get(str(row),table['columns']).items())
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
                columns=table.get('row_columns',{}).get(str(source_row),table[field]) if field=='columns' else table[field]
                for target_col,source_col in columns.items():
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
            for target_col,source_col in table.get('row_columns',{}).get(str(total['source_row']),table['columns']).items():
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






def direct_update_plan(a,b,purpose=None,scope=None,*,comparison_requirements=None):
    """只连接两个合集中的相同身份证；位置安排不得代替身份比较。"""
    from .A合集 import validate_direct_collection, _records, _standard
    from .A语义阅读 import addresses
    for mapping in (a,b):validate_direct_collection(mapping,for_matching=True)
    _require(a['gold_hash']==b['gold_hash'],'两合集标准版本不同')
    standard=_standard(a)
    comparison_requirements=deepcopy(comparison_requirements or [])
    _require(isinstance(comparison_requirements,list),'比较要求须逐项列明维度')
    for item in comparison_requirements:
        _require(isinstance(item,dict) and item.get('维度ID') in standard['dimension'],
                 '比较要求须使用已有维度编号')
    ignored={item['维度ID'] for item in comparison_requirements}
    def in_scope(location,kind):
        if not scope:return True
        sheet,cell=location.rsplit('!',1)
        return any(sheet==area['sheet'] and cell in addresses(area['range']) for area in scope[kind])
    def named_scope(name):
        if not scope:return True
        areas=a['机械记录'].get('names',{}).get(name,[])
        return bool(areas) and all(in_scope(area['sheet']+'!'+cell,'target')
            for area in areas for cell in addresses(area['range']))
    purpose=deepcopy(purpose or {})
    continued=a.get('副本接续',{})
    for previous in continued.get('已执行安排',[]):
        if previous['source_b_hash']!=b['source_hash']:continue
        for field in ('tables','operations','labels'):
            if field in purpose:
                done={definition_hash(row) for row in previous['purpose'].get(field,[])}
                purpose[field]=[row for row in purpose[field] if definition_hash(row) not in done]
    # 历史模型编排只保留在原答复中，不能自动成为用户的机械执行要求。
    if scope:
        # 局部范围不能带入未完整包含的名单改排，也不能扩大为整表写入。
        for field in ('tables','operations'):
            if field in purpose:purpose[field]=[row for row in purpose[field] if named_scope(row['range_name'])]
        if 'labels' in purpose:purpose['labels']=[row for row in purpose['labels'] if in_scope(row['target'],'target')]
    original_purpose=deepcopy(purpose)
    table_order=list(dict.fromkeys(table['range_name'] for table in purpose.get('tables',[])))
    for table in sorted(purpose.get('tables',[]),key=lambda row:(table_order.index(row['range_name']),-row['first_row_offset'])):
        if 'operations' in table:
            purpose.setdefault('operations',[]).extend({'range_name':table['range_name'],
                'reason':purpose.get('reason',''),**deepcopy(item)} for item in table['operations'])
            continue
        # 沿用已有明确名单安排，只把成员数差展开成物理行数变化。
        old,new=table['old_detail_count'],len(table['source_rows'])
        for _ in range(abs(new-old)):
            purpose.setdefault('operations',[]).append({'range_name':table['range_name'],'axis':'row',
                'action':'insert' if new>old else 'delete','index':table['first_row_offset']+min(old,new),
                'count':1,'template_index':table['first_row_offset'],'reason':purpose.get('reason',''),
                **({'source_evidence':table.get('source_evidence',[]),'coverage_reason':table.get('coverage_reason','')} if new<old else {})})
    plan={'schema_version':2,'matching_method':'unique_exact_identity_v4','gold_hash':a['gold_hash'],'source_a_hash':a['source_hash'],
          'source_b_hash':b['source_hash'],'updates':[],'labels':[],'issues':[],
          'retained_blanks':[],'unplaced_sources':[],'operations':[], 'identity_matches':[],
          'update_purpose':original_purpose,'comparison_requirements':comparison_requirements}
    plan['list_relations']=[]
    plan['list_relation_gaps']=[]
    left=dict(_records(a));right=dict(_records(b))
    # 合并标题的各成员是同一个物理格，只归并已保存的明确引用。
    for mapping,records in ((a,left),(b,right)):
        anchors={sheet['name']+'!'+cell:sheet['name']+'!'+area.split(':')[0]
                 for sheet in mapping['机械记录']['sheets'] for area in sheet.get('merges',[])
                 for cell in addresses(area)}
        for location,record in records.items():
            records[location]={**record,**{field:{
                key:list(dict.fromkeys(anchors.get(ref,ref) for ref in refs))
                for key,refs in record[field].items()} for field in ('dimension_evidence','label_references') if field in record}}
    original_left=left
    right={loc:row for loc,row in right.items() if in_scope(loc,'source')}
    original_right=right
    original_target_issues=[]
    for carrier,mapping,kind in [('A',a,'target'),('B',b,'source')]:
        for row in mapping.get('unresolved',[]):
            if not row.get('sheet') or not row.get('cell') or in_scope(row['sheet']+'!'+row['cell'],kind):
                item={'carrier':carrier,**deepcopy(row)};plan['issues'].append(item)
                if item.get('carrier')=='A':original_target_issues.append(item)
    physical=lambda m:{s['name']+'!'+c:v for s in m['机械记录']['sheets'] for c,v in s['cells'].items()}
    acells,bcells=physical(a),physical(b)
    calculated=b.get('formula_values',{})
    if b.get('calculation_evidence'):
        from .统一勾稽 import _calculation_formulas,_calculated_values
        if _calculation_formulas(b).keys()-calculated.keys():
            receipt=json.loads(Path(b['calculation_evidence']['path']).read_text(encoding='utf-8-sig'))
            calculated={**calculated,**_calculated_values(b,receipt)}
    _require(not calculated or b.get('calculation_evidence'),'公式计算结果缺少计算依据')
    for location,item in calculated.items():
        _require(location in bcells and bcells[location].get('formula')==item['formula'],'计算结果与原公式不符')
        if item['state'] in ('value','blank') or item.get('value') in ERROR_CODES:
            # 副本核值只替换此次搬运视图，不改合集中的原公式或身份。
            bcells[location]={**bcells[location],'value':item['value'],'formula':None,
                **({'excel_data_type':'e'} if item.get('value') in ERROR_CODES else {})}
    def value_key(value):
        return definition_hash(value)
    role_cache={}
    def role(mapping,location):
        cache_key=(id(mapping),location)
        if cache_key in role_cache:return role_cache[cache_key]
        sheet,cell=location.rsplit('!',1);found=set()
        for table in mapping['outline']['subtables']:
            if table['sheet']!=sheet:continue
            direct=mapping.get('逐格字段版本')==1
            values=[] if direct else [(item['ranges'],item.get('template_period_role')) for item in table['metrics']]
            values.extend((item['ranges'],item.get('value')) for item in table.get('dimensions',[])
                          if item.get('id')=='template_period_role' and (not direct or item.get('逐格原字段') is True))
            for ranges,value in values:
                if cell in ranges and isinstance(value,str) and value:
                    found.add(value)
        role_cache[cache_key]=next(iter(found)) if len(found)==1 else None
        return role_cache[cache_key]
    def key(record,omit=(),location=None):
        return tuple((key,record[key]) for key in ('metric_id','dimension_id') if record.get(key)),tuple(sorted((k,value_key(record['dimensions'][k]))
            for k in record['dimensions'] if k not in ignored and k not in omit))
    def source_value_keys(record,dim,location):
        return {value_key(record['dimensions'].get(dim))}
    def source_keys(record,omit=(),location=None):
        return [key(record,omit)]
    def issue(location,reason,**details):
        sheet,cell=location.rsplit('!',1)
        plan['issues'].append({'sheet':sheet,'cell':cell,'reason':reason,**details})
    # 新位置只复制版式，不因复制样式而获得旧行或B来源的身份证。
    original_positions=set(acells)
    original_positions.update(table['sheet']+'!'+cell for table in a['outline']['subtables'] for cell in addresses(table['range']))
    added={};positions={k:k for k in sorted(original_positions)};moved_left=left;layout_error=None;failed_layout=None
    if purpose.get('operations'):
        from .表格 import _apply_operation, _transform
        from openpyxl import Workbook
        from openpyxl.workbook.defined_name import DefinedName
        from openpyxl.utils.cell import quote_sheetname
        book=Workbook();book.remove(book.active)
        for sheet in a['机械记录']['sheets']:
            ws=book.create_sheet(sheet['name'])
            for cell,item in sheet['cells'].items():ws[cell]=item['value']
            for merge in sheet.get('merges',[]):ws.merge_cells(merge)
        for name,areas in a['机械记录'].get('names',{}).items():
            if len(areas)==1:
                area=areas[0];book.defined_names.add(DefinedName(name,attr_text=quote_sheetname(area['sheet'])+'!'+area['range']))
        try:
            for op in purpose['operations']:
                _,ws,box=_name_range(book,op['range_name']);sheet=ws.title
                axis=op['axis'];start=box[1] if axis=='row' else box[0];point=start+op['index'];template=start+op['template_index']
                before=dict(positions)
                for old,current in before.items():
                    if current is None:continue
                    sh,cell=current.rsplit('!',1)
                    row,col=coordinate_to_tuple(cell)
                    if op['action']=='delete' and sh==sheet and point<=(row if axis=='row' else col)<point+op.get('count',1) and (axis=='row' or box[1]<=row<=box[3]):
                        positions[old]=None;continue
                    positions[old]=sh+'!'+_transform(cell,sheet,sh,op,box)
                added={sh+'!'+_transform(cell,sheet,sh,op,box):origin for loc,origin in added.items() for sh,cell in [loc.rsplit('!',1)]}
                reverse={v:k for k,v in before.items()}
                for offset in range(op.get('count',1) if op['action']=='insert' else 0):
                    for other in range(box[0],box[2]+1) if axis=='row' else range(box[1],box[3]+1):
                        new=f'{sheet}!{get_column_letter(other if axis=="row" else point+offset)}{point+offset if axis=="row" else other}'
                        old=f'{sheet}!{get_column_letter(other if axis=="row" else template)}{template if axis=="row" else other}'
                        if reverse.get(old) in acells:added[new]=reverse[old]
                _apply_operation(book,op)
            moved_left={positions.get(loc,loc):row for loc,row in left.items() if positions.get(loc,loc) is not None}
        except ValueError as error:
            layout_error=str(error);failed_layout=op['range_name']
            positions={k:k for k in sorted(original_positions)};added={};moved_left=left
        finally:book.close()
    for item in original_target_issues:
        moved=positions.get(item.get('sheet','')+'!'+item.get('cell',''))
        if moved:item['sheet'],item['cell']=moved.rsplit('!',1)
    # 副本中的旧模型保留指令也不能屏蔽当前身份证；原有文字和数值仍在副本中保留。
    retained_labels=set();retained_targets=set()
    reverse={v:k for k,v in positions.items()};unassigned=set();blank_values={};retained_moves={}
    blank_sources={};retained_dependencies={}
    for table in purpose.get('tables',[]):
        if table['range_name']==failed_layout:continue
        area=a['机械记录']['names'][table['range_name']]
        _require(len(area)==1,'名单安排须引用唯一表区')
        area=area[0];l,t,r,bottom=range_boundaries(area['range']);sheet=area['sheet']
        retained_targets.update(positions.get(f'{sheet}!{column}{row}',f'{sheet}!{column}{row}')
                                for row in table.get('retained_rows',[]) for column in table['columns'])
        anchor=positions.get(f'{sheet}!{get_column_letter(l)}{t}',f'{sheet}!{get_column_letter(l)}{t}')
        _require(anchor is not None,'名单表区锚点被移除')
        top=coordinate_to_tuple(anchor.rsplit('!',1)[1])[0]
        first_original=f'{sheet}!{get_column_letter(l)}{t+table["first_row_offset"]}'
        first=coordinate_to_tuple(positions.get(first_original,first_original).rsplit('!',1)[1])[0]
        for member in table.get('retained_members',[]):
            retained_moves.update({f'{sheet}!{column}{first+member["index"]}':action for column,action in member['cells'].items()})
        for index,source_row in enumerate(table['source_rows']):
            target_row=top+table['target_row_offsets'][index] if table.get('target_row_offsets') else first+index
            columns=table.get('row_columns',{}).get(str(source_row),table['columns'])
            member_bindings={f'{sheet}!{column}{target_row}':
                f"{table['source_sheet']}!{source_column}{source_row+table.get('column_row_offsets',{}).get(column,0)}"
                for column,source_column in columns.items()}
            blank_sources.update(member_bindings)
            blank_values.update({f'{sheet}!{column}{target_row}':action
                for column,action in table.get('blank_values',{}).get(str(source_row),{}).items()})
            retained_dependencies.update({f'{sheet}!{column}{target_row}':member_bindings
                for column in table.get('blank_values',{}).get(str(source_row),{}) if column not in columns})
            retained_moves.update({f'{sheet}!{column}{target_row}':action
                for column,action in table.get('blank_values',{}).get(str(source_row),{}).items()
                if column not in table.get('row_columns',{}).get(str(source_row),table['columns'])})
        for total in table.get('totals',[]):
            columns=table.get('row_columns',{}).get(str(total['source_row']),table['columns'])
            member_bindings={positions[f"{sheet}!{column}{total['target_row']}"]:
                f"{table['source_sheet']}!{source_column}{total['source_row']+table.get('column_row_offsets',{}).get(column,0)}"
                for column,source_column in columns.items()}
            retained_dependencies.update({positions[f"{sheet}!{column}{total['target_row']}"]:member_bindings
                for column in table.get('blank_values',{}).get(str(total['source_row']),{}) if column not in columns})
            retained_moves.update({positions[f"{sheet}!{column}{total['target_row']}"]:action
                for column,action in table.get('blank_values',{}).get(str(total['source_row']),{}).items()
                if column not in table.get('row_columns',{}).get(str(total['source_row']),table['columns'])})
    unassigned.update(added)
    unassigned.update(retained_moves)
    retained_targets.difference_update(blank_values)
    same_identity={};same_period_role={}
    # 每格只凭自身身份证对应，原值类型和历史位置安排不参与比较。
    for location,record in right.items():
        if not record.get('metric_id') and not record.get('dimension_id'):continue
        period_role=role(b,location)
        if period_role:
            for match in source_keys(record,{'period'},location):
                same_period_role.setdefault((period_role,match),[]).append(location)
        else:
            for match in source_keys(record,location=location):same_identity.setdefault(match,[]).append(location)
    target_identities={}
    for destination,record in moved_left.items():
        original=added.get(destination,reverse.get(destination,destination))
        if destination in retained_targets or destination in unassigned or not in_scope(original,'target') or not (record.get('metric_id') or record.get('dimension_id')):continue
        period_role=role(a,original)
        identity=(period_role,key(record,{'period'})) if period_role else (None,key(record))
        target_identities.setdefault(identity,[]).append(destination)
    bound={};considered=set()
    for destination,old in moved_left.items():
        if destination in retained_targets:
            sheet,cell=destination.rsplit('!',1)
            plan.setdefault('retained_values',[]).append({'sheet':sheet,'cell':cell,
                'reason':'模型答复明确保留原内容，按已确定安排保留原值'})
            continue
        if destination in unassigned:continue
        original=added.get(destination,reverse.get(destination,destination))
        if not in_scope(original,'target'):continue
        period_role=role(a,original)
        if not old.get('metric_id') and not old.get('dimension_id'):
            issue(destination,'A尚无已识别的本格金标准身份证',kind='incomplete_identity');continue
        identity=(period_role,key(old,{'period'})) if period_role else (None,key(old))
        if len(target_identities[identity])>1:
            issue(destination,'同一身份证存在多个A位置，尚未在A识别阶段确定唯一对应',
                  kind='ambiguous_target',candidates=target_identities[identity])
            continue
        choices=same_identity.get(key(old),[])
        if period_role:
            # 换期以已有角色对应，不能把旧A期初与同日期的B期末混配。
            choices=same_period_role.get((period_role,key(old,{'period'})),[])
        if len(choices)>1:
            considered.update(choices)
            issue(destination,'同一身份证存在多个来源，尚未在B识别阶段确定唯一来源',kind='ambiguous_source',candidates=choices)
            continue
        if not choices:
            candidates=[];incomplete=[]
            for loc,record in original_right.items():
                # 未知不是通配符；未识别来源独立保留，不能据此推定与本项相关。
                if any(record.get(field)!=old.get(field) for field in ('metric_id','dimension_id')):continue
                compared=(set(old['dimensions'])|set(record['dimensions']))-ignored-({'period'} if period_role else set())
                missing=sorted(compared-(set(old['dimensions'])&set(record['dimensions'])))
                differences=sorted(dim for dim in compared if dim not in missing
                             and value_key(old['dimensions'][dim]) not in source_value_keys(record,dim,loc))
                candidate_role=role(b,loc)
                if period_role and candidate_role and candidate_role!=period_role:differences.append('template_period_role')
                candidate={'location':loc,'different_dimensions':differences,'missing_dimensions':missing,
                    'missing_in_a':sorted(compared-set(old['dimensions'])),
                    'missing_in_b':sorted(compared-set(record['dimensions']))}
                if period_role and not candidate_role:
                    candidate['missing_dimensions'].append('template_period_role')
                    candidate['missing_in_b'].append('template_period_role')
                if not period_role and candidate_role:
                    candidate['missing_dimensions'].append('template_period_role')
                    candidate['missing_in_a'].append('template_period_role')
                candidates.append(candidate)
                # A角色未交付时，旧日期差异也要保留为缺角色诊断。
                if candidate['missing_dimensions'] and (not differences or
                        not period_role and candidate_role and differences==['period']):
                    incomplete.append(candidate)
            if incomplete:
                issue(destination,'A本格尚未直接给出换期所需期间角色'
                      if not period_role and any('template_period_role' in item['missing_dimensions'] for item in incomplete)
                      else '同槽位候选之间存在未交字段；候选差异不能直接作为本格补交要求',kind='missing_dependency',candidates=incomplete)
            elif candidates:
                issue(destination,'已有同槽位候选，但已保存的维度或期间角色不对应',kind='semantic_difference',candidates=candidates)
            else:
                issue(destination,'当前B合集没有此槽位的完整已识别来源',kind='missing_source',candidates=[])
            continue
        origin,=choices
        considered.update(choices)
        if origin not in bcells:
            issue(destination,'指定来源位置尚无原始读取记录',kind='missing_source',candidates=[origin]);continue
        # 维度引用只作本格原文依据，不能给另一个标签格制造搬运对应。
        bound[destination]=origin
    areas={op['range_name']:{area['sheet']+'!'+cell
           for area in a['机械记录']['names'].get(op['range_name'],[]) for cell in addresses(area['range'])}
           for op in purpose.get('operations',[])}
    placed={name:{positions.get(loc,loc) for loc in cells}|{loc for loc,origin in added.items() if origin in cells}
            for name,cells in areas.items()}
    failed={failed_layout} if layout_error else set()
    dependencies={}
    if failed:
        # 重叠表区共用物理位置；无关名单按保留下来的原安排重新计算位置。
        while True:
            expanded=failed|{name for name,cells in areas.items() if any(cells&areas[bad] for bad in failed)}
            if expanded==failed:break
            failed=expanded
        affected=set().union(*(areas[name] for name in failed))
        affected_positions=set().union(*(placed[name] for name in failed))
        held=set()
        for name in failed:
            tables=[table for table in original_purpose.get('tables',[]) if table['range_name']==name]
            if (layout_error or len(tables)!=1 or tables[0].get('target_row_offsets') or
                    any(op['range_name']==name for op in original_purpose.get('operations',[]))):
                held.update(areas[name]);continue
            table=tables[0];area=a['机械记录']['names'][name][0]
            first=range_boundaries(area['range'])[1]+table['first_row_offset']
            # 名单停止只保留原明细区，区外各格仍按自身身份证匹配。
            held.update(loc for loc in areas[name]
                        if first<=coordinate_to_tuple(loc.rsplit('!',1)[1])[0]<first+table['old_detail_count'])
        remaining=deepcopy(original_purpose)
        for field in ('operations','tables'):
            remaining[field]=[r for r in original_purpose.get(field,[]) if r['range_name'] not in failed]
        direct={field:[r for r in remaining.pop(field,[]) if r['target'] not in affected_positions]
                for field in ('bindings','labels')}
        reduced_scope={'target':[{'sheet':loc.rsplit('!',1)[0],'range':loc.rsplit('!',1)[1]}
                                 for loc in original_left if in_scope(loc,'target') and loc not in held],
                       'source':scope['source'] if scope else [{'sheet':t['sheet'],'range':t['range']} for t in b['outline']['subtables']]}
        fallback=direct_update_plan(a,b,remaining,reduced_scope,comparison_requirements=comparison_requirements)
        if any(direct.values()):
            new_positions=fallback['template_layout']['positions']
            relocated={dest:new_positions.get(origin) for dest,origin in reverse.items()}
            for origin in set(added.values())-affected:
                old_added=sorted(loc for loc,ref in added.items() if ref==origin)
                new_added=sorted(loc for loc,ref in fallback['template_layout']['added'].items() if ref['template']==origin)
                if len(old_added)==len(new_added):relocated.update(zip(old_added,new_added))
            # 布局尚未成立时，不能把同一字面坐标当成同一旧位置。
            changed_sheets={loc.rsplit('!',1)[0] for cells in areas.values() for loc in cells} if layout_error else set()
            for field,items in direct.items():
                remaining[field]=[]
                for item in items:
                    target=relocated.get(item['target'])
                    if not target or item['target'].rsplit('!',1)[0] in changed_sheets:
                        dependencies.setdefault('坐标安排',[]).append({'reason':'原结构安排未成立，明确指定的位置尚不能重排','target':item['target']})
                        continue
                    remaining[field].append({**item,'target':target})
            fallback=direct_update_plan(a,b,remaining,reduced_scope,comparison_requirements=comparison_requirements)
        fallback['issues'].append({'reason':layout_error or '相关名单依赖未解决；停止结构及名单替换，其他原位置独立匹配',
                                   'range_names':sorted(failed),'dependency_issues':[r for name,rows in dependencies.items()
                                                                                if name in failed or name=='坐标安排' for r in rows]})
        fallback['update_purpose']=original_purpose
        return fallback
    else:plan['operations']=deepcopy(purpose.get('operations',[]))
    for destination,origin in bound.items():
        old=moved_left[destination];new=original_right.get(origin,{'dimensions':{}});sh,cell=destination.rsplit('!',1)
        origins=[origin]
        original=added.get(destination,reverse.get(destination,destination))
        plan['identity_matches'].append({'sheet':sh,'cell':cell,**{key:old[key] for key in ('metric_id','dimension_id') if old.get(key)},'source_locations':origins})
        if bcells[origin].get('formula'):
            result=calculated.get(origin)
            issue(destination,result['reason'] if result and result['state']=='error' else '来源公式尚无可搬运的计算值',
                  kind='unevaluated_source_formula',candidates=[origin])
            continue
        value=bcells[origin]['value']
        if value is None or isinstance(value,str) and not value.strip():
            if destination in blank_values and blank_sources.get(destination)==origin:
                action=blank_values[destination]
                kept=acells[action['source']]['value'] if 'source' in action else action['value']
                plan['updates'].append({'sheet':sh,'cell':cell,'value':kept,
                    'reason':'按明确要求保留或留空','blank_source':origin,
                    **({'retained_from':action['source']} if 'source' in action else {})})
            plan['retained_blanks'].append({'sheet':sh,'cell':cell,'reason':'B空白，按明确要求保留A原值；已确认身份及关系保留'})
        else:
            plan['updates'].append({'sheet':sh,'cell':cell,**{key:old[key] for key in ('metric_id','dimension_id') if old.get(key)},'value':value,
                **({'excel_data_type':'e'} if bcells[origin].get('excel_data_type')=='e' else {}),
                'old_value':None if destination in added else acells.get(original,{}).get('value'),
                'original_dimensions':deepcopy(old['dimensions']),
                'source_identity':{**{key:new[key] for key in ('metric_id','dimension_id') if new.get(key)},'dimensions':deepcopy(new['dimensions'])},
                'sources':[{'sheet':loc.rsplit('!',1)[0],'cell':loc.rsplit('!',1)[1],
                            'value':bcells[loc]['value'],'fact_id':loc} for loc in origins],
                'source_blank_locations':[]})
    for destination,action in retained_moves.items():
        required_bindings=retained_dependencies.get(destination)
        # 随B成员改排的A原值动作，须已有相同身份证连通，不能靠旧行序单独搬运或清空。
        if required_bindings is not None and (not any(bound.get(target)==source for target,source in required_bindings.items())
                or any(target in bound and bound[target]!=source for target,source in required_bindings.items())):
            continue
        sheet,cell=destination.rsplit('!',1)
        plan['updates'].append({'sheet':sheet,'cell':cell,
            'value':acells[action['source']]['value'] if 'source' in action else action['value'],
            'reason':'按明确要求搬运A原值或留空',
            **({'retained_from':action['source']} if 'source' in action else {})})
    plan['caption_updates']=[]
    for location,row in right.items():
        if location not in considered:
            plan['unplaced_sources'].append({'id':location,**{key:row[key] for key in ('metric_id','dimension_id') if row.get(key)},'source_reference':{'file_hash':b['source_hash'],'location':location},
            **({'missing_dimensions':deepcopy(row['missing_dimensions'])} if row.get('missing_dimensions') else {})})
            # B按需识别允许保留多个候选；未采用不等于需要新增一行或再次识别。
    plan['template_layout']={'positions':positions,'added':{k:{'template':v} for k,v in added.items()},'operations':plan['operations']}
    plan['retained_label_targets']=sorted(loc for loc in retained_labels if loc is not None)
    plan['list_dimensions']=[]
    return plan
