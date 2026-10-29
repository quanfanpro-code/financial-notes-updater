# coding: utf-8
"""正常入口按模型声明的范围直接识别，沿用既有合集与机械执行。"""
from copy import deepcopy
import json
from pathlib import Path
import re

from .A语义阅读 import (addresses, metric_catalog, scope_material, recognize_a,
                      direct_reference)
from .A合集 import (combine_direct_collections, direct_digest, save_direct_answer, direct_recognition_status,
                   source_relation, _declared_list_tables, _list_relation_conflicts, _records, append_declared_relation)
from .表格 import read_workbook, materialize_candidate_blanks
from .统一语义 import definition_hash


def read_scope_answer(reply, snapshot, standard, target=None):
    """只接收答复明说的工作表、范围和编号，不按科目或位置推定。"""
    names=[s['name'] for s in snapshot['sheets']]
    a_names=[s['name'] for s in target['机械记录']['sheets']] if target else ['sheet1','原表']
    aliases={}
    for name in names:
        code=re.match(r'^[A-Za-z]+\d+(?=\s)',name)
        if code:aliases.setdefault(code[0],[]).append(name)
    def sheet_name(value):
        if value in names:return value
        code=value.split(' ',1)[0]
        return aliases[code][0] if len(aliases.get(code,[]))==1 else value
    fields={'sheet':('工作表','表名'), 'range':('完整范围','表区范围','上下文范围','子表范围','B来源候选','范围'),
            'data':('数据范围','待识别范围','财务范围'), 'metrics':('相关标准','候选标准','标准编号','指标编号','候选编号'),
            'a':('A范围','对应A','A需要'), 'context':('共同原文','共同说明引用','上下文引用'),
            'lists':('名单维度',), 'subject':('科目','子表','用途','说明')}
    def field(label):
        label=label.replace('`','').replace('**','').strip()
        if '范围' in label and any(arrow in label for arrow in ('→','->','⇒')):return 'range_data'
        for key in ('data','a','context','lists','metrics','sheet','subject','range'):
            if any(word in label for word in fields[key]):return key
        return None
    raw=[];headers=None;block={};a_default=None;plain=[];shared_context=[]
    def collect(value):
        if isinstance(value,str):
            text=re.sub(r'^```(?:json)?\s*|\s*```$','',value.strip())
            try:parsed=json.loads(text)
            except ValueError:plain.append(value)
            else:
                if isinstance(parsed,(dict,list)):collect(parsed)
                else:plain.append(value)
        elif isinstance(value,list):
            for item in value:collect(item)
        elif isinstance(value,dict):
            if isinstance(value.get('引用'),str):shared_context.append(value['引用'])
            row={field(key):'、'.join(v) if isinstance(v,list) else v for key,v in value.items()
                 if field(key) and (isinstance(v,str) or isinstance(v,list) and all(isinstance(x,str) for x in v))}
            if 'range' in row or 'a' in row:raw.append(row)
            else:
                for item in value.values():collect(item)
    collect(reply)
    for line in '\n\n'.join(plain).replace('`','').replace('**','').splitlines():
        line=line.strip()
        if not line or line.startswith('```'):continue
        label,sep,value=line.lstrip('-* ').partition('：')
        if sep and any(name+'!' in label for name in a_names) and 'metric.' in value:
            raw.append({'a':label,'metrics':value});continue
        if line.startswith('|'):
            cells=[v.strip() for v in line.strip('|').split('|')]
            keys=[field(v) for v in cells]
            if ('range' in keys or 'range_data' in keys or 'a' in keys) and ('data' in keys or 'metrics' in keys):
                headers=keys
                for key,label in zip(keys,cells):
                    named=re.search(r'[（(]([^!（()）]+)!',label)
                    if key=='a' and named:a_default=named[1]
                continue
            if headers and len(cells)==len(headers) and not all(re.fullmatch(r'[-: ]*',v) for v in cells):
                raw.append({key:value for key,value in zip(headers,cells) if key})
            continue
        if line.startswith('#'):
            if block:raw.append(block);block={}
        for part in re.split(r'[；;]',line.lstrip('-* ')):
            label,sep,value=re.sub(r'^\d+[.、)）]\s*','',part).partition('：')
            if not sep:label,sep,value=part.partition(': ')
            key=field(label) if sep else None
            if key:
                if key=='range' and key in block:raw.append(block);block={}
                block[key]=value.strip()
    if block:raw.append(block)
    coordinate=r'\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?'
    def refs(text, default=None, sheet_names=names):
        found=[]
        declared='|'.join(re.escape(name) for name in sorted(sheet_names,key=len,reverse=True))
        pattern=(r"(?:'?("+declared+r")'?!|(?<![\w]))("+coordinate+r')') if declared else r'()('+coordinate+r')'
        for match in re.finditer(pattern,text):
            sheet=match[1] or default
            if sheet:found.append({'sheet':sheet,'range':match[2].replace('$','')});default=sheet
        return found
    expanded=[]
    arrow=re.compile(r"(?:(?P<sheet>"+'|'.join(re.escape(name) for name in sorted(names,key=len,reverse=True))+
                     r")!)?(?P<range>"+coordinate+r")\s*(?:→|->|⇒)\s*(?P<data>"+coordinate+
                     r"(?:\s*[、，,]\s*"+coordinate+r")*)")
    for row in raw:
        combined=row.get('range_data')
        if not combined:
            scopes=refs(row.get('range','')) if target is not None else []
            if scopes:
                expanded.extend({**row,**area,'context':'；'.join(filter(None,[row.get('context',''),*shared_context]))} for area in scopes)
            else:expanded.append(row)
            continue
        sheet=None
        for match in arrow.finditer(combined):
            sheet=match['sheet'] or sheet
            expanded.append({**row,'sheet':sheet or '', 'range':match['range'],'data':match['data'],
                             'subject':combined.split('（',1)[0].split('(',1)[0].strip()})
        if not any(arrow.finditer(combined)):expanded.append(row)
    result=[];unread=[]
    for row in expanded:
        if not row.get('range') and row.get('a') and row.get('metrics'):continue
        sheet=sheet_name(row.get('sheet','').strip(" '"))
        scope=refs(row.get('range',''),sheet if sheet in names else names[0] if len(names)==1 else None)
        if len(scope)!=1 or scope[0]['sheet'] not in names:
            unread.append(row);continue
        area=scope[0];data=refs(row.get('data',''),area['sheet'])
        a_ranges=refs(row.get('a',''),a_default or (a_names[0] if len(a_names)==1 else None),a_names)
        a_refs={(r['sheet'],r['range']) for r in a_ranges}
        candidates=[row.get('metrics','')]+[r.get('metrics','') for r in raw if not r.get('range') and
            a_refs & {(a['sheet'],a['range']) for a in refs(r.get('a',''),a_default,a_names)}]
        ids=[key for key in (*standard['metric'],*standard['dimension']) if re.search(r'(?<![\w.])'+re.escape(key)+r'(?![\w.])','；'.join(candidates))]
        same=re.search(r'与(.+?)相同',row.get('metrics',''))
        if not ids and same:
            matches={tuple(r['metric_ids']) for r in result if same[1].strip() in (r['subject'],r['subject']+'表')}
            if len(matches)==1:ids=list(next(iter(matches)))
        if not data and target is None:
            unread.append(row);continue
        result.append({'id':'范围'+str(len(result)+1),'sheet':area['sheet'],'range':area['range'],
            'data_ranges':[r['range'] for r in data if r['sheet']==area['sheet']],
            'context_references':refs(row.get('context',''),area['sheet']),
            'metric_ids':ids,'a_ranges':a_ranges,
            'list_dimensions':[key for key in standard['dimension'] if key in row.get('lists','')],
            'subject':row.get('subject') or area['sheet'],'原范围答复':row})
    return result,unread


def group_source_ranges(groups):
    """模型声明同一A需要的候选合并送入；不按科目名称或金额选择来源。"""
    merged=[]
    for group in groups:
        owned={r['sheet']+'!'+cell for r in group['a_ranges'] for cell in addresses(r['range'])}
        linked=[part for part in merged if owned & part['a_positions']]
        rows=[group]+[row for part in linked for row in part['groups']]
        for part in linked:merged.remove(part);owned.update(part['a_positions'])
        merged.append({'groups':rows,'a_positions':owned})
    result=[]
    for part in merged:
        rows=part['groups'];group=deepcopy(rows[0])
        group['subject']='、'.join(row['subject'] for row in rows)
        group['tables']=[{k:r[k] for k in ('id','sheet','range','data_ranges','context_references') if k in r} for r in rows]
        for key in ('metric_ids','a_ranges','list_dimensions'):
            group[key]=[]
            for row in rows:
                for value in row[key]:
                    if value not in group[key]:group[key].append(value)
        result.append(group)
    return result


def scope_key(tables):
    return json.dumps(sorted((t['sheet'],t['range'],sorted({cell for area in t['data_ranges'] for cell in addresses(area)}),
        sorted((r['sheet'],r['range']) for r in t.get('context_references',[])),t.get('read_full_rows',False))
        for t in (tables if isinstance(tables,list) else [tables]) if t),ensure_ascii=False)


def answer_relations(target, part, source_ranges, dimensions, standard, required_ranges, *, targets=None):
    """接入已识别的标签和版式安排；A/B金额对应只由身份证产生。"""
    reply=part.get('答复接入来源',{}).get('本次原答复',part['识别答复'])
    reference=direct_reference(target,compact=True,ranges=required_ranges)
    requested=targets
    targets=targets if targets is not None else list(dict.fromkeys([row['location'] for row in reference['需要']]+
        [row['source_reference']['location'] for row in reference.get('source_identities',[])]+
        [row['sheet']+'!'+row['cell'] for row in reference.get('unresolved',[]) if row.get('sheet') and row.get('cell')]))
    relation=source_relation(target,part,reply,targets)
    from .A合集 import declared_named_row_relations
    named,arrangements=declared_named_row_relations(target,part,reply)
    for row in named:
        if row.get('label'):
            label={'action':'copy','target':row['target'],'source':row['source']}
            if label not in relation['label_actions']:relation['label_actions'].append(label)
    relation['bindings']=[]
    relations=[relation] if relation['bindings'] or relation['label_actions'] or relation['caption_updates'] or relation.get('retained_targets') else []
    if arrangements:
        relations.append({'kind':'名单编排','dimension':None,'content':reply,'answer_hash':definition_hash(reply),
            'source_a_hash':target['source_hash'],'source_b_hash':part['source_hash'],'gold_hash':part['gold_hash'],
            'source_ranges':source_ranges,'tables':arrangements,'conflicts':[],'未采用的引用说明':[]})
    listed=re.findall(r'名单维度(?:编号)?\s*(?:为|是|[:：=])\s*([a-z_]+)',reply.replace('`','').replace('**',''))
    if re.search(r'按\s*A\s*原表样式',reply) and re.search(r'更新为上述[^。；;]*明细后接合计',reply):
        listed.extend(re.findall(r'(\w+)\s*[、，,]\s*mode\s*=\s*members',reply.replace('`','')))
    dimensions=list(dict.fromkeys([*dimensions,*[key for key in listed if key in standard['dimension']]]))
    declarations=[(dimension,reply) for dimension in dimensions]
    for block in re.split(r'\n(?=#{1,6}\s)',reply):
        if not re.search(r'名单[^。\n]*更新',block) and not ('排布' in block and ('名单' in block or '成员' in block)) and not ('新增' in block and '名单维度' in block) and not ('A沿用原样式' in block and '明细' in block and '合计' in block):continue
        keys=[key for key,row in standard['dimension'].items() if row.get('value_type')=='selection'
              and re.search(r'(?<![\w.])'+re.escape(key)+r'(?![\w.])',block)]
        if not keys and re.search(r'沿用\s*A\s*的名单维度',block):
            keys=list({item['id'] for table in target.get('outline',{}).get('subtables',[])
                if table['sheet']+'!'+table['range'] in block.replace('`','')
                for item in table.get('dimensions',[]) if standard['dimension'].get(item['id'],{}).get('value_type')=='selection'})
        if not keys:
            # 当前段落明说的指标已有维度引用，不借另一段的比例分母充当名单维度。
            metrics={key for key in standard.get('metric',{}) if re.search(r'(?<![A-Za-z0-9_.])'+re.escape(key)+r'(?![A-Za-z0-9_.])',block)}
            own_keys={ref['id'] for metric in metrics for ref in standard.get('metric',{}).get(metric,{}).get('dimension_refs',[])
                      if standard['dimension'].get(ref['id'],{}).get('value_type')=='selection'}
            stated=[key for key,row in standard['dimension'].items() if row.get('value_type')=='selection'
                    and re.search(r'(?<![\w.])'+re.escape(key)+r'(?![\w.])',reply)]
            if len(own_keys)==1:
                key=next(iter(own_keys))
                if key not in dimensions:declarations.append((key,reply if stated==[key] else block));dimensions.append(key)
                continue
            if len(stated)==1:keys=stated
            # 名单段落沿用同一答复表头已明示的维度；完整答复保留列角色与来源引用。
            headers='\n'.join(line for line in reply.splitlines() if line.strip().startswith('|'))
            declared=[key for key,row in standard['dimension'].items() if row.get('value_type')=='selection'
                      and re.search(r'\|\s*`?'+re.escape(key)+r'`?\s*\|',headers)]
            if len(declared)==1 and declared[0] not in dimensions:
                declarations.append((declared[0],reply));dimensions.append(declared[0]);continue
            title=block.splitlines()[0].strip(' #`')
            selected=re.findall(re.escape(title)+r'使用\s*`?(metric\.[\w.]+)',reply) if title else []
            if len(set(selected))==1:
                keys=[r['id'] for r in standard['metric'].get(selected[0],{}).get('dimension_refs',[])
                      if standard['dimension'][r['id']].get('value_type')=='selection']
        if len(keys)==1 and keys[0] not in dimensions:
            all_keys=[key for key,row in standard['dimension'].items() if row.get('value_type')=='selection'
                and re.search(r'(?<![\w.])'+re.escape(key)+r'(?![\w.])',reply)]
            declarations.append((keys[0],reply if len(all_keys)==1 or 'A沿用原样式' in block else block))
    for dimension,content in declarations:
        relation={'kind':'名单编排','dimension':dimension,'content':content,
            'answer_hash':definition_hash(content),'source_a_hash':target['source_hash'],
            'source_b_hash':part['source_hash'],'gold_hash':part['gold_hash'],'source_ranges':source_ranges,
            'tables':_declared_list_tables(target,content,dimension,source_ranges,part)}
        relation['conflicts'],relation['未采用的引用说明']=_list_relation_conflicts(part,content,source_ranges)
        relations.append(relation)
    allowed={r['sheet']+'!'+cell for r in required_ranges or [] for cell in addresses(r['range'])} if required_ranges else None
    label_scope=allowed
    if requested is not None:allowed=set(requested) if allowed is None else allowed&set(requested)
    if allowed is not None:
        from openpyxl.utils.cell import range_boundaries
        for relation in relations:
            for key in ('bindings','label_actions'):
                selected=label_scope if key=='label_actions' else allowed
                if key in relation and selected is not None:relation[key]=[row for row in relation[key] if row['target'] in selected]
            if 'retained_targets' in relation:relation['retained_targets']=[loc for loc in relation['retained_targets'] if loc in allowed]
            owned={name for name,areas in target.get('机械记录',{}).get('names',{}).items()
                if any(any(area['sheet']+'!'+cell in allowed for cell in addresses(area['range'])) for area in areas)}
            if 'caption_updates' in relation:relation['caption_updates']=[row for row in relation['caption_updates'] if row['range_name'] in owned]
            scoped=[]
            for table in relation.get('tables',[]):
                if table['range_name'] not in owned:continue
                table=deepcopy(table);area=target['机械记录']['names'][table['range_name']][0]
                top=range_boundaries(area['range'])[1];first=top+table['first_row_offset'];sheet=area['sheet']
                rows=[(row,first+min(i,table['old_detail_count'])) for i,row in enumerate(table['source_rows'])]
                rows.extend((row['source_row'],row['target_row']) for row in table.get('totals',[]))
                table['row_columns']={str(row):{col:value for col,value in table.get('row_columns',{}).get(str(row),table['columns']).items()
                    if sheet+'!'+col+str(original) in allowed} for row,original in rows}
                table['labels']={col:value for col,value in table.get('labels',{}).items()
                    if all(sheet+'!'+col+str(original) in allowed for _,original in rows)}
                if not any(table['row_columns'].values()) and not table['labels']:continue
                scoped.append(table)
            if 'tables' in relation:relation['tables']=scoped
    from openpyxl.utils.cell import range_boundaries
    for table in [t for relation in relations for t in relation.get('tables',[])]:
        area=target['机械记录']['names'][table['range_name']][0]
        top=range_boundaries(area['range'])[1]
        for index,source_row in enumerate(table['source_rows']):
            if index<table['old_detail_count']:continue
            row=top+(table['target_row_offsets'][index] if table.get('target_row_offsets') else table['first_row_offset']+index)
            for relation in relations:
                for binding in relation.get('bindings',[]):
                    sheet,cell=binding['target'].rsplit('!',1)
                    if sheet!=area['sheet'] or not re.fullmatch(r'[A-Z]+'+str(row),cell):continue
                    col=re.match(r'[A-Z]+',cell)[0]
                    source_col={**table.get('labels',{}),**table.get('row_columns',{}).get(str(source_row),table['columns'])}.get(col)
                    if binding['source']!=table['source_sheet']+'!'+str(source_col)+str(source_row):continue
                    binding['target_member']={key:deepcopy(table[key]) for key in ('range_name','first_row_offset','source_sheet')}
                    binding['target_member'].update(source_row=source_row,column=col)
    return relations


def recognize_document(engine, source, context, metadata, target, destination, log, progress=None, previous_result=None, scope_answers=None):
    """范围认识一次，每个明确范围直接识别一次；不核查、不重答。"""
    carrier=metadata['carrier'];destination=Path(destination)
    snapshot=read_workbook(source,context=context)
    materialize_candidate_blanks(snapshot)
    snapshot.update(semantic_identity_version=2,semantic_formula_sources=5)
    previous=None
    if previous_result:
        stored=json.loads(Path(previous_result).read_text(encoding='utf-8-sig'))
        if (stored.get('source_hash')==snapshot['sha256'] and stored.get('gold_hash')==engine.gold_hash
                and stored.get('carrier')==carrier and stored.get('source_context')==context
                and all(stored.get(key)==metadata.get(key) for key in ('update_purpose','comparison_requirements'))):
            previous=stored
    if previous and previous.get('mapping_path')==str(Path(previous_result).resolve()) and previous.get('正常入口范围认识') and not scope_answers:
        # A需要变化不撤销B已有身份；新缺口由明确范围识别，不能自动重答旧分组。
        from .A合集 import validate_direct_collection
        validate_direct_collection(previous,for_matching=True)
        files={carrier.lower()+'_mapping':str(Path(previous_result).resolve())}
        engine.usage['reused_answers']+=1
        status=direct_recognition_status(previous)
        return {'status':status,'message':carrier+('沿用已完整接入的逐格身份。' if status=='complete' else '沿用已保存的逐格身份；识别完整接入尚未确认，未决内容保留，未重答。'),
                'files':files,'usage':dict(engine.usage)}
    materials={'整份原文.json':scope_material(snapshot,context),
               '标准目录.json':metric_catalog(engine.standard),
               '维度标准目录.json':list(engine.standard['dimension'].values()),
               '名单维度目录.json':[{k:row[k] for k in ('id','domain')} for row in engine.standard['dimension'].values()
                                  if row.get('value_type')=='selection'],
               '明确更新要求.json':metadata.get('update_purpose',{}),
               '明确比较要求.json':metadata.get('comparison_requirements',[])}
    prompt=('先了解整份'+carrier+'的科目和子表，按含义划定后续需要识别的范围，并从目录选相关标准候选。'
        '每个子表说明工作表、含完整表头说明的完整范围、待识别数据范围（包括空白财务位置）、相关标准编号。'
        '数据范围按各子表分别列清，不重叠；子表外共同说明会随材料保留。相关标准可多选，下一步读取其完整定义后才确认身份。'
        '同一子表有余额、变动或比例时，相关计量的候选分别选齐；本次需要更新的名称、期间等格也按本格含义选择指标或维度标准。'
        '候选从本次完整标准目录选取，不能只沿用已有A编号；调节项目与其原始发生额、余额与流量分别按定义选候选。'
        '各子表同时列出需要共同阅读的原文引用（工作表!范围），包括分段表共用列头、各列原文含义，以及表区外的编制单位、日期、币种金额单位及适用说明；是否跨表共同适用由你依据整份材料判断并明确，程序只照搬引用。'
        '这里只认识范围，不逐格识别，不核查。用简洁文字、列表或表格即可，不要求JSON或固定模板。')
    if carrier=='A':
        prompt+=(' A材料保留原Word全部表格和正文；请按明确更新要求及原文含义确定实际参与更新的章节和子表，'
                 '章节标题不预设为某一种写法。提取出来不代表纳入更新，不参与本次更新的内容仍作为原文保留。')
    if target is not None:
        needs={}
        for location,row in _records(target):
            value={**{key:row[key] for key in ('metric_id','dimension_id') if row.get(key)},'dimensions':{key:deepcopy(value) for key,value in row['dimensions'].items()
                if engine.standard['dimension'].get(key,{}).get('value_type')=='selection'}}
            key=json.dumps(value,ensure_ascii=False,sort_keys=True)
            needs.setdefault(key,{**value,'locations':[]})['locations'].append(location)
        materials['A需要.json']={'载体':'A','范围认识':target.get('正常入口范围认识',{}).get('原答复',''),
            '已识别指标':sorted({row['metric_id'] for _,row in _records(target) if row.get('metric_id')}),
            '已识别维度标准':sorted({row['dimension_id'] for _,row in _records(target) if row.get('dimension_id')}),
            '已识别明细与范围':list(needs.values())}
        materials['A原表范围.json']=[{k:t[k] for k in ('id','sheet','range','data_ranges')}
            for t in target['outline']['subtables']]
        prompt+=(' B需了解整份后只选覆盖A需要的必要来源，各范围注明对应A范围（工作表!范围）；'
            '依据A已识别的明细与范围查找来源，汇总表只能提供合计时仍需查找其明细所在表。'
            '名单随B更新时包含新成员。注明要更新的名单维度编号，引用目录中指标的维度；'
            '没有来源的需要直接说明。多处可能同义的来源在此保留候选，完整定义和实际来源关系留给随后直接识别，不比较金额。')
    log(carrier+'：了解整份范围并收窄相关标准')
    reply=previous.get('正常入口范围认识',{}).get('原答复') if previous else None
    if scope_answers:
        reply=json.dumps([Path(path).read_text(encoding='utf-8-sig') for path in scope_answers],ensure_ascii=False)
    if reply:engine.usage['reused_answers']+=1
    else:reply=engine.request_text(prompt,materials)
    (destination/(carrier+'范围认识原答复.txt')).write_text(reply,encoding='utf-8-sig')
    groups,unread=read_scope_answer(reply,snapshot,engine.standard,target)
    if target is not None:groups=group_source_ranges(groups)
    (destination/(carrier+'识别范围.json')).write_text(json.dumps({'ranges':groups,'未接入范围原文':unread},ensure_ascii=False,indent=2),encoding='utf-8')
    carried=previous;corrections=[];seen=set()
    while previous and (correction:=previous.get('答复接入来源')) and correction.get('字段范围'):
        source=correction.get('原合集')
        if source in seen:break
        corrections.append(correction);seen.add(source)
        previous=json.loads(Path(source).read_text(encoding='utf-8-sig')) if source and Path(source).is_file() else None
    parts=[];saved={}
    indexed=[Path(part['mapping_path']) for part in carried.get('合集来源',[])] if carried else []
    for path in indexed or engine.work_dir.glob('直接识别_*/'+carrier+'识别结果.json'):
        previous=json.loads(path.read_text(encoding='utf-8-sig'))
        if (previous.get('source_hash')==snapshot['sha256'] and previous.get('gold_hash')==engine.gold_hash
                and previous.get('source_context')==context
                and (target is None or previous.get('A需要',{}).get('source_hash')==target['source_hash'])
                and all(previous.get(key)==metadata.get(key) for key in ('update_purpose','comparison_requirements'))):
            key=scope_key(previous.get('target_tables',previous.get('target_table')))
            saved.setdefault(key,{})[previous['答复指纹']]=previous
    for group in groups:
        engine._check_cancel()
        log(carrier+'：'+(group['subject'] or group['id'])+'，直接识别并保存')
        local={**metadata,'normal_entry':True,'a_required_ranges':group['a_ranges']}
        table={k:group[k] for k in ('id','sheet','range','data_ranges','context_references') if k in group}
        if target is not None:
            local['target_tables']=[{**row,'read_full_rows':True} for row in group['tables']]
            local['b_source_ranges']=[{k:row[k] for k in ('sheet','range','context_references') if k in row} for row in group['tables']]
        candidates=list(group['metric_ids'])
        if not candidates:
            unread.append({**group,'原因':'范围已保存，原答复未列明相关标准候选；未由程序补选。'})
            continue
        # 金标准已经写明的比例引用也带入材料；代码不据此选择任何身份。
        for key,row in engine.standard['metric'].items():
            refs={side['metric_id'] for side in row.get('definition',{}).get('ratio',{}).values()}
            if refs and refs<=set(group['metric_ids']) and key not in candidates:candidates.append(key)
        prior=saved.get(scope_key(local.get('target_tables',table)),{})
        if target is not None:
            prior={key:value for key,value in prior.items() if value.get('a_required_ranges')==group['a_ranges']}
        if len(prior)==1:
            previous=next(iter(prior.values()))
            from .统一勾稽 import load_fact_mapping
            part=load_fact_mapping(previous['mapping_path'],{})
            engine.usage['reused_answers']+=1
        else:
            part=recognize_a(engine,source,context,artifact_metadata=local,target_mapping=target,
                allowed_metrics=candidates,target_table=table if target is None else None,source_snapshot=snapshot)
        parts.append(part)
        if progress:progress({'status':'partial','message':carrier+'当前范围已保存','files':{},'usage':{}})
    if not parts:
        return {'status':'saved','message':carrier+'范围原答复已保存，尚无可接入的明确范围；未重答。',
                'files':{'recognition_answer':str(destination/(carrier+'范围认识原答复.txt'))},'usage':dict(engine.usage)}
    result=combine_direct_collections(parts)
    if carried:
        result['封存定义原文']=list(dict.fromkeys([*result['封存定义原文'],*carried['封存定义原文']]))
        result['金标准原文']='\n'.join(result['封存定义原文'])
    for index,correction in enumerate(reversed(corrections),1):
        if correction.get('来源范围'):result['b_source_ranges']=deepcopy(correction['来源范围'])
        result=save_direct_answer(result,correction['本次原答复'],destination/'已有纠正接入'/f'{index}.json',
            field_scope=correction['字段范围'],withdrawn=correction.get('撤下旧字段'),snapshot=snapshot,standard=engine.standard,
            target_mapping=target)
    if carried:
        for relation in carried.get('已确认关系',[]):
            append_declared_relation(result['已确认关系'],relation)
    result.update({k:deepcopy(metadata[k]) for k in ('a_link','update_purpose','comparison_requirements') if k in metadata})
    result['已确认关系']=list({(r['kind'],r['answer_hash'],r.get('dimension')):r for r in result['已确认关系']}.values())
    result['正常入口范围认识']={'原答复':reply,'ranges':groups,'未接入范围原文':unread}
    if target is not None:result['A需要指纹']=definition_hash(direct_reference(target,compact=True))
    result['usage']=dict(engine.usage)
    if unread or any(direct_recognition_status(part)!='complete' for part in parts):
        result['本次逐格接入']['complete']=False
    if carrier=='B':
        from .统一勾稽 import calculate_mapping_formulas
        calculate_mapping_formulas(result,destination/'来源取值副本')
    result['封存校验']=direct_digest(result)
    path=destination/(carrier+'识别结果.json');result['mapping_path']=str(path.resolve())
    with path.open('x',encoding='utf-8') as f:json.dump(result,f,ensure_ascii=False,indent=2)
    from .统一勾稽 import load_fact_mapping
    load_fact_mapping(path,{})
    files={carrier.lower()+'_mapping':str(path),'recognition_answer':str(destination/(carrier+'范围认识原答复.txt'))}
    if target is not None and result['已确认关系']:
        a={**target,'已确认关系':deepcopy(target.get('已确认关系',[]))}
        for relation in result['已确认关系']:append_declared_relation(a['已确认关系'],relation)
        a['封存校验']=direct_digest(a);a_path=destination/'A识别结果.json';a['mapping_path']=str(a_path.resolve())
        with a_path.open('x',encoding='utf-8') as f:json.dump(a,f,ensure_ascii=False,indent=2)
        load_fact_mapping(a_path,{})
        files['a_mapping']=str(a_path)
    status=direct_recognition_status(result)
    return {'status':status,'message':carrier+('各范围识别字段已完整接入并保存。' if status=='complete' else '各范围原答复和已接入字段已保存；识别尚未完整接入，缺原文和未接入内容保留，未重答。'),
            'files':files,'usage':dict(engine.usage)}
