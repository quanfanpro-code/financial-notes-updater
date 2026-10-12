# coding: utf-8
"""A先读原文语义，再找标准；共同含义按模型声明的范围展开。"""
from copy import deepcopy
from collections import Counter, deque
from pathlib import Path
import json
import re
import time

from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.formula.tokenizer import TokenizerError
from .统一语义 import _require, definition_hash, validate_facts, required_cell_dimensions
from .统一证据 import context_catalog, applicable_context
from .表格 import read_workbook, materialize_candidate_blanks, file_hash, Tokenizer, _REFERENCE
from .表格编排 import stage_proofs, _stage_receipt

READ_PROMPT = '''读懂整份文件的编排、全部科目和子表，再说明各表行列及说明的实际含义，共同信息只说一次。只识别语义，不推算或补填数值；空白财务格和业务说明也不能漏掉。
金额、比例等列与业务行交叉的数据位置，空白时仍是待更新数据；与纯排版空白区分。
返回JSON：{"tables":[{"id":"本文件内编号","sheet":"工作表","range":"含标题表头的范围","subject":"科目","meaning":"行列、上下级、期间及例外含义","data_ranges":["全部待更新数据范围"]}],"notes":"公共理解及不明确之处"}。
一表可列多个数据范围；范围内其余格表示标题、表头或标签。续表保留所属科目；不把整张混合工作表当一张表。'''

SELECT_PROMPT = '''根据A已经表达的语义，在金标准目录中查找准确对应的指标；候选必须继续核对完整定义。不能只按表名匹配，不能把最接近的指标当作已对应。
结合原表文字逐项查看行列，不能只看整表概括；上级汇总包含多种经济对象时也查找相应组合指标。
按每个项目自身的经济对象和计量含义选标准，同一业务事实换表展示不换标准。复合表的总项与子项分别选择适用指标；已有更直接的对象标准能完整表达子项时优先使用，不给子项追加原文没有的父项限定。
展开指标及各维度定义，核对原文对象与计量是否在允许范围内；维度的新取值不能扩大指标已明确排除的对象，也不能把费用、累计账项变动、单项支出与净损益相互替代。
返回JSON：{"tables":[{"id":"原表编号","metric_ids":["完整指标ID"],"missing":"未找到的业务含义，无则空字符串"}]}。每张表都登记。'''

READ_CHECK_PROMPT = '''沿用整份认识，只解决列出的范围疑点，返回这些表修正后的原reading格式。数字在数据范围外，确认它是表头序号还是漏掉的财务值；文字在数据范围内，区分独立披露的业务事实和用于说明对象/类别/账龄的身份标签。独立的款项内容、未偿还原因仍是数据，名称或分类标签保留为语义依据；不因没有金标准删掉真正的数据。空白财务位置仍保留，不能因范围调整漏掉。
只返回{"tables":[原表记录]}；保持原表编号、工作表和总范围，只修正meaning和data_ranges。不改数值。'''

BIND_PROMPT = '''依据A原文，把全部待更新格准确对应到金标准，保存标准指标及其维度。共同信息写一次并注明适用位置，空白数据位置同样有身份；不补算数值。
识别成功必须完整对应金标准。不能只保存文字含义、最近标准或改写标准定义作为另一种成功结果。对不上先检查阅读、共同说明、标准选择和维度解释；继续修正识别，不因一次未匹配就认定标准缺漏，也不凭空补条件。
公司、币种等可以共用；不同行的单位、项目、类别不能合并。例如甲单位行和乙单位行各有自己的对象，不能给两行都写甲乙合集。合计保留其总体，不能变成当前唯一明细；共同信息不得覆盖行列差异。
本任务确定“这里应填写什么”，不是判断现有金额是否为零或是否计提；已读出的data_ranges全部需要身份，不能因为没有数值而放入unresolved。依据引用解释含义的表头、行名或说明，不引用数据值本身。
业务范围必须写实际业务含义，不用“某表某行合计”或当前展示的几条明细代替。位置只写在ranges和evidence里。
返回JSON：{"tables":[{"id":"原编号","metrics":[{"ranges":["范围"],"id":"指标ID"}],"dimensions":[{"ranges":["范围或*表示本表全部数据"],"id":"维度ID","value":"实际取值","evidence":["逐字一致的原文短语或单独的@C编号"]}],"unresolved":[{"range":"范围","reason":"缺什么"}]}]}。
只用metrics和dimensions保存成功身份。仍未解决的识别错误列unresolved并指出具体位置和原因；只有列出原文证据、已查标准及无法表达的具体语义，才能报告金标准缺项。标准必填字段暂未识别出来，本身不是标准缺项的证据。
若材料含已确认维度取值，它们是对应标准维度中已有的规范表达，供本次识别复用。原文与其完整含义相同时使用该取值，避免同义内容另造写法；包含不同父项、条件、组成或边界时仍按本次原文表达，不为匹配抄用。每项维度依据必须来自正在识别的原文，取值目录本身不是原文证据。
各指标范围互不重叠；同一维度的不同取值范围互不重叠。期间用{"kind":"instant","date":"YYYY-MM-DD"}或{"kind":"duration","start":"YYYY-MM-DD","end":"YYYY-MM-DD"}。集合用{"mode":"all"}、{"mode":"members","members":["完整类别"]}或{"mode":"predicate","predicate":"完整限定条件"}，不确定就注明。程序补入定义中的domain，按该指标允许的维度保存。'''

REVIEW_PROMPT = '''核查已有身份证是否忠实表达A原文，重点看共同含义、行列差异、例外及全部目标覆盖。只改错误，不重新生成身份。
确认每个成功身份已准确对应金标准指标及维度；未对应的先检查并修正阅读、共同说明、标准选择和维度。不得改用source_identities或最近标准宣布成功。确实发现标准缺漏时，issues须提供具体原文、已查标准及无法表达的语义；不能把识别错误或未读到字段当成标准缺项。
有review_focus时，本次只检查该范围；原表其他部分用于理解上下文，不重写其身份。
这里不填金额、不查金额是否发生。每个data_ranges位置均须有身份，空白不是未决理由。父项和子项、不同类别分别保留；合计不能缩成当前明细名单。检查program_issues和semantic_questions中的具体疑点。
返回JSON：{"tables":[{"id":"原表编号","corrections":{"metrics":[{"ranges":["范围"],"id":"正确指标ID"}],"patches":[{"ranges":["范围"],"id":"维度ID","value":"按标准类型填写","evidence":["原文短语或单独的@C编号"]}],"resolved_ranges":["已经解决的旧未决范围"]},"issues":[{"range":"仍不明确的范围","reason":"缺少什么含义","evidence":["原文依据"]}]}]}。
每表登记；无更正或问题时相应数组为空。维度更正仅作用于声明范围；集合用all、members或predicate对象，期间遵守指标的时点/期间要求。evidence只引用逐字一致的原文短语或适用的@C编号，不把地址和编号拼在一起。确需排除身份标签时用corrections.non_business，标准外字段用remove_dimensions移除；财务目标不能排除。'''

# 用户指定的A/B格子对应不以整份报表口径为前提，标准原文件保持不变。
BIND_PROMPT += '\nreport_scope是报表共同信息，不作为格子绑定的必填维度，也不参与A/B格子对应；原文有明确说明时保留，没有时省略，不询问或猜测。其余业务维度仍须准确绑定。'
REVIEW_PROMPT += '\nreport_scope不参与格子对应，不因缺少该项提出未决；检查其余指标和业务维度。'
BIND_PROMPT += '\n槽位、维度定义、实际取值分别确认。原文未填写具体名称或其他实际取值时，仍在metrics保留已确认标准；dimensions只填有据部分，集合可用unknown或partial，不把空白成员改成all或合计。实际取值未提供不等于槽位未识别，不因此放入unresolved或要求补姓名。'
REVIEW_PROMPT += '\n核查槽位与维度定义是否符合原文，并检查已有取值是否有据。具体维度取值尚未填写但槽位已明确时，保留已知部分和未知状态，不仅因缺具体取值提出issues；错误指标、错误已有取值或无法确认的语义仍须指出。'
BIND_PROMPT += '\n本格明确表头及适用说明决定其含义；执行者的推断、文件名或金额公式引用另一表，不覆盖本格明确的期间和业务范围。公式可补充原文字段的引用线索，不用于按数字猜含义。'
REVIEW_PROMPT += '\n核对指标及维度定义允许的对象和计量边界，不能用维度取值容纳被指标排除的对象或改变计量基础。核查已有取值是否来自本格原文；执行者推断、文件名或金额公式引用另一表不能覆盖明确表头。'


def addresses(area):
    area=area.replace('、',',').replace('，',',')
    if ',' in area:
        return [address for part in area.split(',') for address in addresses(part.strip())]
    left, top, right, bottom = range_boundaries(area)
    _require(all(type(n) is int for n in (left,top,right,bottom)) and 1<=left<=right and 1<=top<=bottom,
             '范围须为有效的单格或矩形')
    _require((right-left+1)*(bottom-top+1)<=500000, '范围过大')
    return [f'{get_column_letter(c)}{r}' for r in range(top,bottom+1) for c in range(left,right+1)]


def data_cells(table, cells):
    selected=[]
    for area in table['data_ranges']:
        for address in addresses(area):
            if address in cells:
                _require(address not in selected, '财务范围重复')
                selected.append(address)
    _require(set(selected)<=set(addresses(table['range'])), '财务范围越出本表')
    return selected


def expand_scopes(answer, table, cells, standard):
    """把模型明确的指标及各项取值与范围组合，不推断业务含义。"""
    owned=set(data_cells(table,cells));result={a:{'dimensions':{}} for a in owned}
    def scope(ranges):
        selected=set()
        for area in ranges:
            current=owned if area=='*' else set(addresses(area))
            _require(current<=set(addresses(table['range'])),'取值范围越出本表')
            current &= owned
            _require(current and not current&selected,'取值范围没有财务目标或重复')
            selected.update(current)
        return selected
    for row in answer['metrics']:
        _require(row['id'] in standard['metric'],'不存在的指标')
        for address in scope(row['ranges']):
            _require('metric_id' not in result[address],'同一数据位置有两个指标')
            result[address]['metric_id']=row['id']
    for row in answer['dimensions']:
        key=row['id'];definition=standard['dimension'].get(key)
        refs=[row['evidence']] if isinstance(row['evidence'],str) else row['evidence']
        _require(definition is not None,'不存在的维度')
        value=deepcopy(row['value'])
        if definition['value_type']=='selection' and isinstance(value,dict):
            value.setdefault('domain',definition['domain'])
            value.setdefault('completeness','unknown' if value.get('mode')=='unknown' else 'complete')
        for address in scope(row['ranges']):
            previous=result[address]['dimensions'].get(key)
            _require(previous is None or previous['value']==value,'同一数据位置的维度取值冲突')
            result[address]['dimensions'][key]={'value':value,'evidence':list(dict.fromkeys((previous['evidence'] if previous else [])+refs))}
    for row in answer.get('unresolved',[]):
        for address in scope([row['range']]):result[address].update(metric_id=None,reason=row['reason'])
    return result


def formula_source_material(snapshot, selected):
    """追溯静态引用及其行列文字；不求值，也不据公式直接发证。"""
    sheets={s['name']:s for s in snapshot['sheets']};queue=deque();seen=set()
    for sheet in snapshot['sheets']:
        queue.extend((sheet['name'],a) for a,c in sheet['cells'].items()
                     if a in selected.get(sheet['name'],set()) and c.get('formula'))
    result={'引用':{},'原文':{},'未展开':[]};indexes={}
    def context_cells(name,address):
        sheet=sheets[name];cell=sheet['cells'][address];row,col=cell['row'],cell['column']
        if name not in indexes:
            rows={};columns={}
            for a,c in sheet['cells'].items():
                if isinstance(c.get('value'),str) or c.get('formula'):
                    rows.setdefault(c['row'],[]).append(a);columns.setdefault(c['column'],[]).append(a)
            indexes[name]=(rows,columns)
        rows,columns=indexes[name]
        owned=set(rows.get(row,[]))
        owned.update(a for a in columns.get(col,[]) if sheet['cells'][a]['row']<row and not sheet['cells'][a].get('formula'))
        for area in sheet['merges']:
            left,top,right,bottom=range_boundaries(area)
            if (top<=row<=bottom and left<=col) or (left<=col<=right and bottom<row):
                owned.add(f'{get_column_letter(left)}{top}')
                owned.update(a for a in rows.get(top,[]) if not sheet['cells'][a].get('formula'))
        text={}
        for a in sorted(owned):
            c=sheet['cells'].get(a,{})
            if c.get('formula'):
                if (c['row']==row and c['column']<col) or isinstance(c.get('cached_value'),str):
                    text[name+'!'+a]={'公式':c['formula']};queue.append((name,a))
            elif isinstance(c.get('value'),str) and c['value'].strip():text[name+'!'+a]=c['value']
        return text
    while queue:
        name,address=queue.popleft();location=name+'!'+address
        if location in seen:continue
        seen.add(location);cell=sheets[name]['cells'].get(address,{})
        if not cell.get('formula'):continue
        try:tokens=Tokenizer(cell['formula']).items
        except (TokenizerError,ValueError,IndexError):
            result['未展开'].append({'位置':location,'引用':cell['formula']});continue
        refs=[]
        for token in tokens:
            if token.type!='OPERAND' or token.subtype!='RANGE':continue
            match=_REFERENCE.fullmatch(token.value)
            target=(match['sheet'] or name).strip("'").replace("''","'") if match else None
            if not match or target not in sheets:
                result['未展开'].append({'位置':location,'引用':token.value});continue
            area=match['a']+(':'+match['b'] if match['b'] else '')
            left,top,right,bottom=range_boundaries(area.replace('$','').upper())
            # 只遍历真实读取网格，整列或越界引用不凭空生成单元格。
            cells=sheets[target]['cells'];single=match['a'].replace('$','').upper()
            sources=([(single,cells[single])] if single in cells else []) if not match['b'] else [
                (a,c) for a,c in cells.items() if left<=c['column']<=right and top<=c['row']<=bottom]
            if not sources:
                result['未展开'].append({'位置':location,'引用':token.value});continue
            for a,c in sources:
                ref=target+'!'+a;refs.append(ref)
                if ref not in result['原文']:
                    record={'行列文字':context_cells(target,a)}
                    if c.get('formula'):record['公式']=c['formula']
                    elif isinstance(c.get('value'),str):record['文字']=c['value']
                    else:record['数据位置']='只确认业务含义，不使用数值及空白状态'
                    result['原文'][ref]=record
                queue.append((target,a))
        result['引用'][location]=list(dict.fromkeys(refs))
    return result


def source_material(snapshot, context, tables=None, semantic_only=False):
    """按原行列呈现，不抹掉空白、合并关系或适用说明。"""
    catalog=context_catalog(context,snapshot['path'])
    aliases={f'@C{i}':ref for i,ref in enumerate(catalog)}
    selected={s['name']:set(s['cells']) for s in snapshot['sheets']}
    if tables is not None:
        selected={}
        for table in tables:
            selected.setdefault(table['sheet'],set()).update(addresses(table['range']))
            for area in table.get('context_ranges',[]):selected[table['sheet']].update(addresses(area))
    sheets=[]
    targets={}
    if semantic_only:
        for table in tables or []:targets.setdefault(table['sheet'],set()).update(a for area in table['data_ranges'] for a in addresses(area))
    for sheet in snapshot['sheets']:
        owned=selected.get(sheet['name'],set());lines={}
        for address,cell in sheet['cells'].items():
            if address not in owned:continue
            value=cell.get('cached_value') if cell.get('formula') else cell['value']
            if address in targets.get(sheet['name'],set()) and (value in (None,'') or isinstance(value,(int,float)) or cell.get('formula')):
                value='[待确认业务含义]'
            lines.setdefault(cell['row'],[]).append(f'{get_column_letter(cell["column"])}={"[空白]" if value is None or value=="" else value}')
        if lines:sheets.append({'sheet':sheet['name'],'rows':'\n'.join(str(r)+': '+' | '.join(parts) for r,parts in sorted(lines.items())),
                                'merges':[a for a in sheet['merges'] if set(addresses(a))&owned]})
    context_rows={}
    for alias,ref in aliases.items():
        record=catalog[ref];scope=record['scope']
        if tables is not None and scope and not any(t['sheet']==scope['sheet'] and
            ref in applicable_context(catalog,t['sheet'],[range_boundaries(t['range'])[1],range_boundaries(t['range'])[0],range_boundaries(t['range'])[3],range_boundaries(t['range'])[2]]) for t in tables):continue
        content=record['content']
        if isinstance(content,dict):content={k:content[k] for k in ('title','chapter_context','unit_text','text') if content.get(k)}
        context_rows[alias]={'scope':scope,'text':content}
    texts=[]
    def collect(value):
        if isinstance(value,str) and len(value)>40:texts.append(value)
        elif isinstance(value,dict):
            for item in value.values():collect(item)
        elif isinstance(value,list):
            for item in value:collect(item)
    collect(context_rows)
    common={text:'T'+str(i) for i,(text,count) in enumerate(Counter(texts).items()) if count>1}
    def compact(value):
        if isinstance(value,str) and value in common:return {'原文引用':common[value]}
        if isinstance(value,dict):return {k:compact(v) for k,v in value.items()}
        if isinstance(value,list):return [compact(v) for v in value]
        return value
    material={'sheets':sheets,'context':compact(context_rows)}
    if snapshot.get('semantic_formula_sources'):
        material['公式原文']={sheet['name']+'!'+address:cell['formula'] for sheet in snapshot['sheets']
            for address,cell in sheet['cells'].items() if address in selected.get(sheet['name'],set()) and cell.get('formula')}
        material['公式用途']='仅用于理解构成和引用来源。与文字说明不一致时核对原文及直接来源，保留冲突依据；不计算新金额，不按数值相等判断身份。'
        if snapshot['semantic_formula_sources']==2 and material['公式原文']:
            material['公式引用依据']=formula_source_material(snapshot,selected)
    if semantic_only:material['数值处理']='仅身份确认阶段隐藏已确定数据范围内的数值及空白状态；全部位置都要确认含义，原件及最终保存的原值不变。'
    if common:material['共同原文']={key:text for text,key in common.items()}
    def text_parts(value):
        if isinstance(value,str):yield value
        elif isinstance(value,dict):
            for child in value.values():yield from text_parts(child)
        elif isinstance(value,list):
            for child in value:yield from text_parts(child)
    aliases['_context_sources']=[{'text':text,'ref':aliases[key],'scope':record['scope']}
        for key,record in context_rows.items() for text in text_parts(record['text'])]
    def contains(value,text):
        if isinstance(value,str):return value==text
        if isinstance(value,dict):return any(contains(v,text) for v in value.values())
        if isinstance(value,list):return any(contains(v,text) for v in value)
        return False
    for text,key in common.items():
        aliases[key]=[{'ref':aliases[alias],'scope':record['scope']} for alias,record in context_rows.items() if contains(record['text'],text)]
    # 允许模型引用表内唯一原文；这里只定位逐字一致的文字，不推断其财务含义。
    for table in tables or []:
        cells=next(s for s in snapshot['sheets'] if s['name']==table['sheet'])['cells']
        literals={}
        for address in dict.fromkeys(a for area in [table['range'],*table.get('context_ranges',[])] for a in addresses(area)):
            cell=cells.get(address,{})
            value=cell.get('cached_value') if cell.get('formula') else cell.get('value')
            if isinstance(value,str) and value.strip() and ('semantic_candidates' in snapshot or not cell.get('formula')):
                literals.setdefault(value,[]).append(address)
        left,top,right,bottom=range_boundaries(table['range'])
        for value,refs in literals.items():
            if re.fullmatch(r'\$?[A-Z]+\$?[1-9]\d*',value) or value.startswith(('@','T')):continue
            row_counts=Counter(cells[ref]['row'] for ref in refs)
            for ref in refs:
                row=cells[ref]['row']
                if row_counts[row]!=1 and len(row_counts)!=1:continue
                # 同一行重复的原文保留全部出处，不猜其中哪个才是模型所指。
                bounds=[top,left,bottom,right] if len(row_counts)==1 else [row,left,row,right]
                aliases.setdefault(value,[]).append({'ref':ref,'scope':{'sheet':table['sheet'],'bounds':bounds}})
    return material,aliases


def validate_reading(answer, snapshot):
    _require(isinstance(answer.get('tables'),list) and answer['tables'],'没有返回子表')
    sheets={s['name']:s for s in snapshot['sheets']};seen=set();covered={}
    for table in answer['tables']:
        _require(table['id'] not in seen and table['sheet'] in sheets and table['meaning'] and table['subject'],'子表缺少唯一编号或含义')
        seen.add(table['id']);sheet=sheets[table['sheet']]
        area=set(addresses(table['range']));left,top,right,bottom=range_boundaries(table['range'])
        _require(right<=sheet['max_column'] and bottom<=sheet['max_row'],'子表越出原文件')
        _require(not area&covered.setdefault(table['sheet'],set()),'子表范围重叠')
        covered[table['sheet']].update(area)
        for target in table['data_ranges']:_require(set(addresses(target))<=area,'数据范围越出子表')
        data_cells(table,sheet['cells'])
    return deepcopy(answer)


def reading_questions(reading, snapshot):
    questions=[]
    for table in reading['tables']:
        packet=table_packet(table,snapshot,[]);owned=set(data_cells(table,packet['cells']))
        outside=[a for a,c in packet['cells'].items() if a not in owned and type(c['value']) in (int,float)
                 and ('semantic_candidates' not in snapshot or table['sheet']+'!'+a in snapshot['semantic_candidates'])]
        text=[a for a in owned if isinstance(packet['cells'][a]['value'],str) and
              packet['cells'][a]['value'].strip() not in {'','-','—','–','/'} and not packet['cells'][a].get('formula')]
        if outside or text:questions.append({'id':table['id'],'数据范围外的数值':outside,'须分清业务事实与身份标签的文字':sorted(text)})
    return questions


def reading_check_payload(reading, questions, snapshot, context):
    ids={q['id'] for q in questions};tables=[t for t in reading['tables'] if t['id'] in ids]
    source,_=source_material(snapshot,context,tables)
    return {'task':'a_semantic_read_check','reading':tables,'questions':questions,'source':source}


def apply_reading_check(reading, answer, questions, snapshot):
    wanted={q['id'] for q in questions};rows=answer['tables']
    _require(len(rows)==len(wanted) and {t['id'] for t in rows}==wanted,'局部阅读复查漏表或重复')
    index={t['id']:t for t in reading['tables']}
    for row in rows:
        original=index[row['id']]
        _require(all(row[k]==original[k] for k in ('sheet','range')),'局部复查改变了整份表格编排')
        index[row['id']]=row
    return validate_reading({**reading,'tables':[index[t['id']] for t in reading['tables']]},snapshot)


def metric_catalog(standard, include_basis=True):
    return [{'id':m['id'],'meaning':m['definition']['economic_object']+'：'+m['definition']['measure'],
             'time_kind':m['definition']['time_kind'],
             **({'basis':m['definition']['measurement_basis'].split('。',1)[0]} if include_basis else {})}
            for m in standard['metric'].values() if m['status']=='active']


def validate_selection(answer, reading, standard):
    wanted={t['id'] for t in reading['tables']};rows=answer['tables']
    _require(len(rows)==len(wanted) and {t['id'] for t in rows}==wanted,'标准检索遗漏或重复子表')
    for row in rows:
        _require(isinstance(row['metric_ids'],list),'指标候选须为列表')
        _require(all(m in standard['metric'] and standard['metric'][m]['status']=='active' for m in row['metric_ids']),'检索引用不存在的标准')
    return deepcopy(answer)


def selection_payload(reading,snapshot,standard,include_basis=True):
    labels=[]
    for table in reading['tables']:
        packet=table_packet(table,snapshot,[])
        labels.append({'id':table['id'],'sheet':table['sheet'],'text':{a:value for a,c in packet['cells'].items()
            if isinstance(value:=c.get('cached_value') if c.get('formula') else c.get('value'),str) and value.strip()
            and ('semantic_candidates' in snapshot or not c.get('formula'))}})
    return {'task':'a_semantic_select','reading':reading,'catalog':metric_catalog(standard,include_basis),'原表文字':labels}


def standard_material(ids, standard):
    metrics=[standard['metric'][m] for m in dict.fromkeys(ids)]
    dimensions={d['id'] for m in metrics for d in m['dimension_refs']}
    repeated=Counter(part for d in standard['dimension'].values() for part in d['meaning'].split(' ') if len(part)>100)
    definitions=[]
    for key in sorted(dimensions):
        original=standard['dimension'][key]
        row={k:v for k,v in original.items() if k in {'id','value_type','domain','allowed_values'}}
        # 各维度重复的通用识别说明统一表达，指标定义和该维度独有的业务职责保持原文。
        row['meaning']=' '.join(p for p in original['meaning'].split(' ') if repeated[p]<10)
        definitions.append(row)
    return {'metrics':[{'id':m['id'],'definition':{k:v for k,v in m['definition'].items() if k!='ratio'},
                        'dimension_refs':m['dimension_refs']} for m in metrics], 'dimensions':definitions,
            '共同识别原则':'范围以指标经济对象、计量基础和维度职责为总体。原文覆盖完整业务总体且无附加筛选时用all，不要求列出所有类别证明全体；不能把合计改成当前展示的明细名单。存在父项限定时保留父项，明确账龄或类别本身已是确定范围。保留其中项、交集、并集及排除条件，不添加原文没有的限制；不凭金额、当前行数、位置或常见做法推断。只记已知含义，真正缺少的条件才列未决。'}


def checked_records(answer, table, packet, standard, aliases):
    from .统一识别 import validate_response
    expanded=expand_scopes(answer,table,packet['cells'],standard);records={}
    label_refs=[a for a,c in packet['cells'].items() if isinstance(c['value'],str) and c['value'].strip() and not c.get('formula')]
    _require(label_refs,'表格缺少文字依据')
    for address in packet['candidate_cells']:
        if address not in expanded:
            records[address]={'kind':'non_business','category':'label','reason':'模型整体阅读明确的数据范围以外的标题、表头或标签', 'evidence_cells':label_refs[:1]}
            continue
        row=expanded[address];metric=row.get('metric_id')
        if metric is None:
            records[address]={'kind':'unresolved','reason':row.get('reason','缺少对应标准'),'evidence_cells':label_refs[:1]};continue
        _require(metric in standard['metric'], '未找到指标：'+str(metric))
        allowed={d['id'] for d in standard['metric'][metric]['dimension_refs']}
        required=required_cell_dimensions(standard['metric'][metric],standard)
        _require(set(row.get('dimensions',{}))<=set(standard['dimension']), '模型使用了不存在的维度ID')
        dims={k:v['value'] for k,v in row.get('dimensions',{}).items() if k in allowed}
        try:
            evidence={k:normalize_evidence(v['evidence'],table['sheet'],aliases,address) for k,v in row.get('dimensions',{}).items() if k in allowed}
        except (ValueError,KeyError,TypeError,IndexError) as error:
            records[address]={'kind':'unresolved','reason':'本格原文依据不能解析：'+str(error),'evidence_cells':label_refs[:1]}
            continue
        record={'kind':'metric_fact','metric_id':metric,'dimensions':dims,'dimension_evidence':evidence,'reason':row.get('reason',table['meaning'])}
        if required-set(dims) or any(isinstance(v,dict) and v.get('completeness','complete')!='complete' for v in dims.values()):
            record.update(kind='partial_metric_fact',missing_dimensions=sorted(required-set(dims)),evidence_cells=label_refs[:1])
        records[address]=record
    records=validate_response({'cells':records},packet,standard,isolate_errors=True)
    return apply_role_corrections(records,{'corrections':{'non_business':answer.get('non_business',[])}},table,packet,standard,aliases)


def normalize_evidence(refs, sheet, aliases, address=None):
    if isinstance(refs,str):refs=[refs]
    result=[]
    for ref in refs:
        literal=ref.removeprefix('原文引用：').removeprefix('原文引用:').strip()
        if re.fullmatch(r'@T\d+',literal) and literal[1:] in aliases:literal=literal[1:]
        if literal in aliases:ref=literal
        items=aliases.get(ref)
        if items is None and literal and not re.fullmatch(r'\$?[A-Z]+\$?[1-9]\d*(?::\$?[A-Z]+\$?[1-9]\d*)?',literal) and not literal.startswith('@'):
            items=[item for item in aliases.get('_context_sources',[]) if literal in item['text']]
            items.extend(item for text,refs in aliases.items() if isinstance(refs,list) and not text.startswith(('@','_'))
                and not re.fullmatch(r'T\d+',text) and literal in text for item in refs)
        if isinstance(items,list) and items:
            _require(address is not None,'共同原文引用需要明确适用位置')
            column,row,_,_=range_boundaries(address)
            matching=[]
            for item in items:
                scope=item['scope']
                if not scope or (scope['sheet']==sheet and scope['bounds'][0]<=row<=scope['bounds'][2] and scope['bounds'][1]<=column<=scope['bounds'][3]):matching.append(item['ref'])
            _require(matching,'共同原文说明不适用于当前位置：'+ref)
            result.extend(matching);continue
        if ref.startswith('@'):result.append(aliases.get(ref,ref));continue
        if '!' in ref:
            prefix,ref=ref.rsplit('!',1)
            _require(prefix.strip("'")==sheet,'依据引用了其他工作表')
        # 证据地址可带括注；只读取明确的地址，括注不作为财务判断或原文证据。
        annotated=re.fullmatch(r'(\$?[A-Z]+\$?[1-9]\d*(?::\$?[A-Z]+\$?[1-9]\d*)?)(?:[（(][^\r\n]*[）)]|[=：].+)',ref)
        if annotated:ref=annotated.group(1)
        result.extend(addresses(ref.replace('$','')))
    return list(dict.fromkeys(result))


def table_packet(table, snapshot, context):
    sheet=next(s for s in snapshot['sheets'] if s['name']==table['sheet'])
    area=set(addresses(table['range']))
    shown=area|{a for span in table.get('context_ranges',[]) for a in addresses(span)}
    return {'source_path':snapshot['path'],'source_hash':snapshot['sha256'],'sheet':sheet['name'],
            'cells':{a:c for a,c in sheet['cells'].items() if a in shown},
            'candidate_cells':[a for a in sheet['cells'] if a in area and
                ('semantic_candidates' not in snapshot or sheet['name']+'!'+a in snapshot['semantic_candidates'])], 'document_context':context}


def binding_payload(tables, selection, snapshot, context, standard, semantic_only=True, reading_notes=None):
    material,aliases=source_material(snapshot,context,tables,semantic_only=semantic_only)
    selected={r['id']:r for r in selection['tables']}
    ids=[m for t in tables for m in selected[t['id']]['metric_ids']]
    # 检索已读取完整行名，只提供实际选中的定义，不再次扩成整个同类指标。
    payload={'task':'a_semantic_bind','reading':tables,'source':material,'standards':standard_material(ids,standard)}
    if snapshot.get('semantic_slot_bindings'):
        payload['槽位与取值分层']='先确认金标准槽位及维度定义，再记录原文实际取值。名称或金额未填写不否定已知槽位；缺少具体取值时保留已有维度，集合用unknown或partial如实记录，不填假名称，不改成all或合计，不放进unresolved删除已知指标。核查仍须确认指标及已知维度都有原文依据。'
    vocabulary={}
    for target in snapshot.get('semantic_known_targets',[]):
        metric=target.get('metric_id')
        if target.get('semantic_key') is None or metric not in ids:continue
        allowed={r['id'] for r in standard['metric'][metric]['dimension_refs']}
        for key,value in target.get('dimensions',{}).items():
            definition=standard['dimension'].get(key,{})
            if key not in allowed or key in {'entity','report_scope','period','currency','unit_scale'}:continue
            if definition.get('value_type') not in {'selection','text'} or definition.get('allowed_values'):continue
            item={'metric_id':metric,'dimension_id':key,'value':deepcopy(value)}
            vocabulary[definition_hash(item)]=item
    if vocabulary:payload['已确认维度取值']=list(vocabulary.values())
    locations={t['sheet']+'!'+a for t in tables for span in t['data_ranges'] for a in addresses(span)}
    templates=[r for r in snapshot.get('semantic_style_templates',[]) if r['location'] in locations and r['metric_id'] in ids]
    if templates:
        payload['style_templates']=deepcopy(templates)
        payload['style_reuse_instruction']='核对本表上下文和所列差异。一致部分用style_reuse:[{template_id,evidence:[本次原文位置],reason:一致依据}]沿用；不要重写这些格的指标和constant_dimensions。只在dimensions补当前主体、实际日期、币种、单位及context_bindings所需业务范围。其他格及不一致部分按现有metrics、dimensions识别。模板不是当前原文证据，禁止沿用历史年份或主体。'
    if reading_notes is not None:payload['reading_notes']=deepcopy(reading_notes)
    return payload,aliases


def source_identity_records(row, table, packet, standard, aliases):
    """保存模型明确的原文含义，不以标准必填项是否齐全判定原文身份。"""
    from .统一识别 import validate_identity_review
    result={};owned=set(data_cells(table,packet['cells']))
    for declaration in row.get('source_identities',[]):
        meaning=declaration.get('meaning');difference=declaration.get('standard_difference','')
        _require(isinstance(meaning,str) and meaning.strip(),'原文身份证缺少实际含义')
        _require(isinstance(difference,str),'标准差别说明须为文字')
        candidate=declaration.get('closest_metric_id')
        _require(candidate is None or candidate in standard['metric'] and standard['metric'][candidate]['status']=='active','最近标准不存在或未启用')
        selected=[]
        for span in declaration['ranges']:
            area=owned if span=='*' else set(addresses(span))
            _require(area<=set(addresses(table['range'])),'原文身份证范围越表')
            selected.extend(sorted(area&owned))
        _require(selected and len(selected)==len(set(selected)) and not set(selected)&set(result),'原文身份证没有目标或范围重复')
        for address in selected:
            evidence=normalize_evidence(declaration['evidence'],table['sheet'],aliases,address)
            _require(evidence and not set(evidence)&owned,'原文身份证须引用解释含义的文字，不能引用目标数据')
            validate_identity_review({'checked_cells':[address],'issues':[{'cells':[address],'reason':meaning,'evidence_cells':evidence}]},packet)
            result[address]={'source_reference':{'file_hash':packet['source_hash'],'location':table['sheet']+'!'+address},
                'table_id':table['id'],'meaning':meaning,'evidence_cells':evidence,
                'closest_metric_id':candidate,'standard_difference':difference}
    return result


def validate_bindings(answer, tables, snapshot, context, standard, candidate=False, require_identity=False):
    _require(isinstance(answer.get('tables'),list),'身份结果须按表返回')
    answer=deepcopy(answer)
    if require_identity:
        _require(not any(t.get('source_identities') for t in answer['tables']),
                 '识别成功必须完整对应金标准，请通过metrics和dimensions修正绑定')
    index={t['id']:t for t in tables}
    if any(row.get('style_reuse') for row in answer['tables']):
        from .B样式复用 import apply_reuse
        answer['tables']=[apply_reuse(row,index[row['id']],snapshot.get('semantic_style_templates',[])) for row in answer['tables']]
    for row in answer['tables']:
        if row.get('id') not in index:
            scopes=[a for m in row.get('metrics',[]) for a in m['ranges']]+[u['range'] for u in row.get('unresolved',[])]
            _require(scopes and '*' not in scopes,'无法从明确范围定位返回的原表')
            declared=set().union(*(set(addresses(a)) for a in scopes))
            matches=[t for t in tables if row.get('id')==t['sheet'] and declared<=set(addresses(t['range']))]
            _require(len(matches)==1,'返回表编号无唯一对应原表')
            row['id']=matches[0]['id']
        row.setdefault('dimensions',[])
        row.setdefault('metrics',[])
        for metric in row.get('metrics',[]):
            for dimension in metric.pop('dimensions',[]):
                _require('*' not in metric['ranges'] and '*' not in dimension['ranges'],'嵌套维度须声明明确范围')
                metric_scope=set().union(*(set(addresses(a)) for a in metric['ranges']))
                dimension_scope=set().union(*(set(addresses(a)) for a in dimension['ranges']))
                _require(dimension_scope<=metric_scope,'嵌套维度范围超出对应指标')
                row['dimensions'].append(dimension)
    _require(len(answer['tables'])==len(index) and {t['id'] for t in answer['tables']}==set(index),'身份结果漏表或重复')
    if any('source_identities' in row for row in answer['tables']):
        _,aliases=source_material(snapshot,context,tables)
        for row in answer['tables']:
            table=index[row['id']]
            source_identity_records(row,table,table_packet(table,snapshot,context),standard,aliases)
    if candidate:
        try:answer=distribute_shared_dimensions(answer,tables,snapshot)
        except (ValueError,KeyError,TypeError):return deepcopy(answer)
        return answer
    answer=distribute_shared_dimensions(answer,tables,snapshot)
    _,aliases=source_material(snapshot,context,tables)
    for row in answer['tables']:
        table=index[row['id']]
        checked_records(row,table,table_packet(table,snapshot,context),standard,aliases)
    return deepcopy(answer)


def distribute_shared_dimensions(answer,tables,snapshot):
    """只分配模型明确声明的跨表范围；星号仍只代表所在表。"""
    result=deepcopy(answer);index={t['id']:t for t in tables}
    owned={t['id']:set(data_cells(t,next(s for s in snapshot['sheets'] if s['name']==t['sheet'])['cells'])) for t in tables}
    rows={r['id']:r for r in result['tables']}
    for row in result['tables']:row['dimensions']=[]
    for row in answer['tables']:
        origin=index[row['id']]
        siblings=[t for t in tables if t['sheet']==origin['sheet']]
        permitted=set().union(*(set(addresses(t['range'])) for t in siblings))
        for dimension in row['dimensions']:
            scope=set().union(*(owned[row['id']] if area=='*' else set(addresses(area)) for area in dimension['ranges']))
            _require(scope<=permitted,'共同信息范围超出本次已明确的数据范围')
            _require(any(scope&owned[t['id']] for t in siblings),'共同信息没有财务目标')
            for target in siblings:
                selected=scope&owned[target['id']]
                if selected:
                    ranges=dimension['ranges'] if scope<=set(addresses(origin['range'])) else sorted(selected)
                    rows[target['id']]['dimensions'].append({**deepcopy(dimension),'ranges':ranges})
    return result


def review_for_table(row, table, packet, aliases):
    from .统一识别 import validate_identity_review
    checked=packet['candidate_cells']
    if 'checked_ranges' in row:
        area={a for span in row['checked_ranges'] for a in addresses(span)}
        _require(area<=set(addresses(table['range'])),'核查声明越出本表')
        checked=[a for a in checked if a in area]
    issues=[]
    for issue in row.get('issues',[]):
        area=set(addresses(issue['range']))
        _require(area<=set(addresses(table['range'])),'核查范围越出本表')
        cells=[a for a in packet['candidate_cells'] if a in area]
        _require(cells,'核查问题没有对应原格')
        issues.append({'cells':cells,'reason':issue['reason'],'evidence_cells':normalize_evidence(issue['evidence'],table['sheet'],aliases,cells[0])})
    return validate_identity_review({'checked_cells':checked,'issues':issues},packet)


def validate_review(answer, tables, snapshot, context, focus=None, require_identity=False):
    index={t['id']:t for t in tables}
    if answer and 'tables' not in answer and all(k in index and isinstance(v,dict) and v.get('id',k)==k for k,v in answer.items()):
        answer={'tables':[dict(v,id=k) for k,v in answer.items()]}
    _require(isinstance(answer.get('tables'),list),'核查须按表返回')
    if require_identity:
        _require(not any(row.get('corrections',{}).get('source_identities') for row in answer['tables'] if isinstance(row,dict)),
                 '识别成功必须完整对应金标准，核查不能改回文字身份')
    valid=[t for t in answer['tables'] if isinstance(t,dict) and isinstance(t.get('id'),str)]
    _require(valid,'没有可恢复的完整子表核查记录')
    incomplete=len(valid)!=len(answer['tables']) or {t['id'] for t in valid}!=set(index)
    answer={**deepcopy(answer),'tables':deepcopy(valid)}
    if incomplete:answer['incomplete']=True
    _require(len(answer['tables'])==len({t['id'] for t in answer['tables']}) and
             {t['id'] for t in answer['tables']}<=set(index),'核查表重复或越界')
    _,aliases=source_material(snapshot,context,tables)
    for row in answer['tables']:
        if focus is not None:
            wanted={a for span in focus[row['id']] for a in addresses(span)}
            if 'checked_ranges' in row:
                _require({a for span in row['checked_ranges'] for a in addresses(span)}==wanted,'核查回复范围与请求不符')
            row['checked_ranges']=deepcopy(focus[row['id']])
            corrections=row.get('corrections',{})
            _require(not corrections.get('dimensions') and not corrections.get('remove_dimensions'),'局部核查须用有范围的修正')
            for key in ('metrics','patches','non_business','source_identities'):
                for change in corrections.get(key,[]):
                    spans=[change['range']] if key=='non_business' else change['ranges']
                    _require({a for span in spans for a in addresses(span)}<=wanted,'更正越出本次核查范围')
            _require({a for span in corrections.get('resolved_ranges',[]) for a in addresses(span)}<=wanted,'解除未决越出本次核查范围')
        corrections=row.get('corrections',{});patches=corrections.get('patches',[])
        metric_patches=[p for p in patches if (p.get('id')=='metric' and isinstance(p.get('value'),str) and p['value'].startswith('metric.'))
                        or (isinstance(p.get('id'),str) and p['id'].startswith('metric.') and 'value' not in p)]
        if metric_patches:
            metrics=corrections.get('metrics',[])+[{'ranges':p['ranges'],'id':p['value'] if p['id']=='metric' else p['id']} for p in metric_patches]
            corrections['metrics']=list({json.dumps(m,sort_keys=True):m for m in metrics}.values())
            corrections['patches']=[p for p in patches if p not in metric_patches]
        table=index[row['id']];review_for_table(row,table,table_packet(table,snapshot,context),aliases)
    return deepcopy(answer)


def reviewed_binding(binding, review, tables=None, snapshot=None):
    """核查沿用原结果，只采纳有明确取值及范围的维度更正。"""
    result=deepcopy(binding);index={r['id']:r for r in result['tables']}
    for row in review.get('tables',[]):
        corrections=row.get('corrections',{})
        _require(set(corrections)<= {'dimensions','metrics','non_business','remove_dimensions','patches','resolved_ranges','source_identities'},'核查更正字段不支持')
        replacements=corrections.get('dimensions',[])
        ids={d['id'] for d in replacements}|set(corrections.get('remove_dimensions',[]))
        target=index[row['id']]
        for change in corrections.get('source_identities',[]):
            table=next(t for t in tables if t['id']==row['id'])
            owned=set(data_cells(table,next(s for s in snapshot['sheets'] if s['name']==table['sheet'])['cells']))
            def declared(item):
                area=set().union(*(owned if span=='*' else set(addresses(span)) for span in item['ranges']))
                _require(area<=set(addresses(table['range'])),'原文身份更正越表')
                return area&owned
            changed=declared(change)
            _require(changed,'原文身份更正没有财务目标')
            kept=[{**old,'ranges':sorted(remaining)} for old in target.get('source_identities',[]) if (remaining:=declared(old)-changed)]
            target['source_identities']=kept+[deepcopy(change)]
        if corrections.get('non_business'):
            target.setdefault('non_business',[]).extend(deepcopy(corrections['non_business']))
        target['dimensions']=[d for d in target['dimensions'] if d['id'] not in ids]+deepcopy(replacements)
        for change in corrections.get('metrics',[]):
            table=next(t for t in tables if t['id']==row['id'])
            owned=set(data_cells(table,next(s for s in snapshot['sheets'] if s['name']==table['sheet'])['cells']))
            changed=set(a for area in change['ranges'] for a in addresses(area))
            _require(changed and changed<=set(addresses(table['range'])),'指标修正越出本表')
            changed &= owned
            _require(changed,'指标修正没有财务目标')
            kept=[]
            for original in target['metrics']:
                remaining=(set(a for area in original['ranges'] for a in addresses(area))&owned)-changed
                if remaining:kept.append({**original,'ranges':sorted(remaining)})
            target['metrics']=kept+[deepcopy(change)]
        for change in corrections.get('patches',[]):
            table=next(t for t in tables if t['id']==row['id'])
            owned=set(data_cells(table,next(s for s in snapshot['sheets'] if s['name']==table['sheet'])['cells']))
            def scope(item):
                declared=set().union(*(owned if a=='*' else set(addresses(a)) for a in item['ranges']))
                _require(declared<=set(addresses(table['range'])),'局部修正越出本表')
                return declared&owned
            changed=scope(change)
            _require(changed,'局部修正没有财务目标')
            kept=[]
            for original in target['dimensions']:
                if original['id']!=change['id']:kept.append(original);continue
                remaining=scope(original)-changed
                if remaining:kept.append({**original,'ranges':sorted(remaining)})
            target['dimensions']=kept+[deepcopy(change)]
        resolved=set(a for area in corrections.get('resolved_ranges',[]) for a in addresses(area))
        if resolved:
            table=next(t for t in tables if t['id']==row['id'])
            owned=set(data_cells(table,next(s for s in snapshot['sheets'] if s['name']==table['sheet'])['cells']))
            _require(resolved<=set(addresses(table['range'])),'解除未决范围越出本表')
            resolved &= owned
            _require(resolved,'解除未决范围没有财务目标')
            target['unresolved']=[{**item,'range':a} for item in target.get('unresolved',[]) for a in addresses(item['range']) if a not in resolved]
    return result


def duplicate_identity_questions(binding,tables,snapshot,standard):
    """同列不同原文却有相同身份时请模型查明，不自动判错或猜测财务范围。"""
    questions=[];index={t['id']:t for t in tables}
    for row in binding['tables']:
        table=index[row['id']];cells=table_packet(table,snapshot,[])['cells']
        try:expanded=expand_scopes(row,table,cells,standard)
        except (ValueError,KeyError,TypeError):continue
        source_only={a for item in row.get('source_identities',[]) for span in item['ranges']
                     for a in (expanded if span=='*' else addresses(span))}
        groups={}
        for address,item in sorted(expanded.items()):
            if address in source_only:continue
            metric=item.get('metric_id')
            if not metric:continue
            allowed={d['id'] for d in standard['metric'][metric]['dimension_refs']}
            values={k:v['value'] for k,v in item['dimensions'].items() if k in allowed}
            key=(cells[address]['column'],metric,json.dumps(values,sort_keys=True,ensure_ascii=False))
            groups.setdefault(key,[]).append(address)
        for group in groups.values():
            if len(group)<2:continue
            labels={a:'；'.join(c['value'] for c in cells.values() if c['row']==cells[a]['row'] and c['column']<cells[a]['column']
                and isinstance(c['value'],str) and c['value'].strip() and not c.get('formula')) for a in group}
            if len(set(labels.values())-{''})<2:continue
            questions.append({'id':row['id'],'ranges':sorted(group),'check':'same_identity_different_rows',
                'reason':'这些同列位置目前得到完全相同的身份，但行原文不同：'+json.dumps(labels,ensure_ascii=False)+
                '。请逐行保留各自对象/项目/类别及合计的总体；不能把多个明细的合集共同赋给每条明细。共同公司、期间等不能覆盖这些差异。只有原文确实表达相同范围才可沿用，并说明依据。'})
    return questions


def semantic_questions(binding,tables,snapshot,standard):
    """指出已发生过的合计范围疑点；只提问，财务结论仍由模型依据原文给出。"""
    questions=duplicate_identity_questions(binding,tables,snapshot,standard);index={t['id']:t for t in tables}
    for row in binding['tables']:
        table=index[row['id']];cells=next(s for s in snapshot['sheets'] if s['name']==table['sheet'])['cells']
        parent_checked=set()
        owned=set(data_cells(table,cells))
        objects=' '.join(standard['metric'][m['id']]['definition']['economic_object'] for m in row['metrics'] if m['id'] in standard['metric'])
        total_rows={c['row'] for a,c in cells.items() if a in addresses(table['range']) and isinstance(c['value'],str) and ('合计' in c['value'] or '小计' in c['value'])}
        for item in row['dimensions']:
            definition=standard['dimension'].get(item['id'],{});domain=definition.get('domain','');value=item['value']
            if not domain or domain not in objects or not isinstance(value,dict):continue
            reason=None;question_ranges=item['ranges']
            if value.get('mode') in {'all','members'} and table['subject']!=domain and item['id'] not in parent_checked:
                parent_checked.add(item['id']);question_ranges=table['data_ranges']
                examples=list(dict.fromkeys(str(name) for d in row['dimensions'] if d['id']==item['id'] and isinstance(d['value'],dict) for name in d['value'].get('members',[])))[:3]
                reason=f'本维度总体为指标计量基础下的“{domain}”，当前表为“{table["subject"]}”。请核对本表全部该维度取值是否保留实际科目/父项：all不能省掉局部父项，成员名称脱离表题也须完整（当前示例：{examples}）。指标本身已经保证的限定不必重复，不添加原文没有的条件。'
            elif value.get('mode')=='members':
                area=set(a for r in item['ranges'] for a in (owned if r=='*' else addresses(r)))
                if any(cells[a]['row'] in total_rows for a in area&owned):
                    reason='这些成员名单覆盖了原表合计/小计行。请核对合计代表的父项或排名条件，不能用目前列出的公司或明细名单代替；只有原文确实限定列举成员才保留名单。'
            if reason:questions.append({'id':row['id'],'dimension':item['id'],'ranges':question_ranges,'reason':reason})
    return questions


def apply_role_corrections(first, review, table, packet, standard, aliases):
    from .统一识别 import validate_response
    for row in review.get('corrections',{}).get('non_business',[]):
        area=addresses(row['range'])
        _require(set(area)<=set(packet['candidate_cells']),'业务角色修正超出原表')
        _require(row['category'] in {'label','annotation'},'只允许有据的标签或注释修正')
        record={'kind':'non_business','category':row['category'],'reason':row['reason'],
                'evidence_cells':normalize_evidence(row['evidence'],table['sheet'],aliases,area[0])}
        first.update(validate_response({'cells':{a:record for a in area}},
                     {**packet,'candidate_cells':area},standard))
    return first


def shared_binding(binding):
    """只合并完全相同且范围不重叠的声明，原文措辞不做同义改写。"""
    result=deepcopy(binding)
    for row in result['tables']:
        for key in ('metrics','dimensions'):
            grouped={}
            for item in row[key]:
                signature=json.dumps({k:v for k,v in item.items() if k!='ranges'},sort_keys=True,ensure_ascii=False)
                grouped.setdefault(signature,[]).append(item)
            merged=[]
            for group in grouped.values():
                ranges=[a for item in group for a in ([item['ranges']] if isinstance(item['ranges'],str) else item['ranges'])]
                try:flat=[a for area in ranges for a in addresses(area)]
                except (ValueError,TypeError):merged.extend(group);continue
                if len(flat)!=len(set(flat)):merged.extend(group);continue
                merged.append({**group[0],'ranges':ranges})
            row[key]=merged
    return result


def review_material(payload, binding, tables, snapshot, context, standard, prior_reviews=()):
    _,aliases=source_material(snapshot,context,tables);index={t['id']:t for t in tables};issues=[]
    catalog=context_catalog(context,snapshot['path'])
    for row in binding['tables']:
        table=index[row['id']]
        source_known=source_identity_records(row,table,table_packet(table,snapshot,context),standard,aliases)
        try:records=checked_records(row,table,table_packet(table,snapshot,context),standard,aliases)
        except (ValueError,KeyError,TypeError) as error:
            reason=str(error)
            if reason=='范围须为有效的单格或矩形':
                reason='某个身份条目的ranges或evidence包含不能解析的位置引用；不是说reading中的表范围有错。详见下面各条维度问题。'
            issues.append({'id':row['id'],'reason':reason})
            # 首个冲突不能遮住其他已知错误；逐项检查仅补报无效项，不重复列必需维度缺项。
            for dimension in row['dimensions']:
                try:
                    partial=checked_records({**row,'dimensions':[dimension]},table,table_packet(table,snapshot,context),standard,aliases)
                    failures={r['reason'] for r in partial.values() if r['kind']=='invalid_response'}
                except (ValueError,KeyError,TypeError) as detail:failures={str(detail)}
                for reason in failures:
                    if '@context:' in reason:reason='所引用的@C说明不适用于这些位置，请依据当前表的表头、行名或全局说明重新引用。'
                    issues.append({'id':row['id'],'dimension':dimension['id'],'ranges':dimension['ranges'],'reason':reason})
            continue
        reasons={}
        for address,record in records.items():
            if address in source_known and record['kind'] in {'partial_metric_fact','unresolved'}:continue
            if record['kind'] in {'invalid_response','unresolved'}:
                reason=record['reason']
                if reason=='缺少对应标准':reason='请核对原文及已查标准，修正指标选择；报告标准缺项须提供原文与标准对照证据。'
            elif record['kind']=='partial_metric_fact':
                if snapshot.get('semantic_slot_bindings'):continue
                reason='金标准绑定未完成，请结合行列和共同说明检查这些维度：'+','.join(record.get('missing_dimensions',[]))+'；不得用文字身份替代，也不得补造条件。'
            else:continue
            for alias,ref in aliases.items():
                if not isinstance(ref,str):continue
                if ref in reason:
                    scope=catalog[ref]['scope']
                    if scope:
                        top,left,bottom,right=scope['bounds']
                        area=f'{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}'
                        reason=f'所引说明{alias}只适用于{scope["sheet"]}!{area}，不适用于当前问题位置；请换用适用的原文依据。'
                    else:reason=reason.replace(ref,alias)
            reasons.setdefault(reason,[]).append(address)
        issues.extend({'id':row['id'],'cells':cells,'reason':reason} for reason,cells in reasons.items())
    grouped={}
    for issue in issues:
        key=json.dumps({k:v for k,v in issue.items() if k not in {'ranges','cells'}},sort_keys=True,ensure_ascii=False)
        target=grouped.setdefault(key,{k:v for k,v in issue.items() if k not in {'ranges','cells'}})
        for field in ('ranges','cells'):
            if field in issue:target[field]=list(dict.fromkeys(target.get(field,[])+issue[field]))
    issues=list(grouped.values())
    result={**payload,'task':'a_semantic_review','recognized':shared_binding(binding),'共同声明合并':True,
            '返回骨架':{'tables':[{'id':t['id'],'corrections':{'patches':[],'resolved_ranges':[]},'issues':[]} for t in tables]}}
    if issues:result['program_issues']=issues
    questions=semantic_questions(binding,tables,snapshot,standard)
    if questions:result['semantic_questions']=questions
    previous=[{'id':row['id'],**issue} for answer in prior_reviews for row in answer.get('tables',[]) for issue in row.get('issues',[])]
    if previous:result['previous_issues']=previous
    return result


def assemble(snapshot, context, standard, reading, batches, b_scope=None, b_outline=None, require_identity=False):
    from .统一识别 import apply_identity_review, workbook_coverage
    result={'facts':[],'owners':{},'excluded':[],'unresolved':[],'source_identities':[]}
    confirm_slots=bool(snapshot.get('semantic_slot_bindings'))
    if confirm_slots:result['slot_bindings']=[]
    tables={t['id']:t for t in reading['tables']};processed=set()
    for batch in batches:
        binding,review=batch[:2]
        review=carry_role_corrections(review,batch[2] if len(batch)>2 else [])
        _,aliases=source_material(snapshot,context,[tables[r['id']] for r in binding['tables']])
        binding=reviewed_binding(binding,review,list(tables.values()),snapshot)
        if require_identity:
            _require(not any(row.get('source_identities') for row in binding['tables']),
                     '识别成功必须完整对应金标准，核查不能改回文字身份')
        reviews={r['id']:r for r in review.get('tables',[])}
        for row in binding['tables']:
            table=tables[row['id']];packet=table_packet(table,snapshot,context)
            try:first=checked_records(row,table,packet,standard,aliases)
            except (ValueError,KeyError,TypeError) as error:
                first={a:{'kind':'unresolved','reason':'身份声明尚未修正：'+str(error),'evidence_cells':[]} for a in packet['candidate_cells']}
            first=apply_role_corrections(first,reviews.get(row['id'],{}),table,packet,standard,aliases)
            checked=review_for_table(reviews[row['id']],table,packet,aliases) if row['id'] in reviews else {'checked_cells':[],'issues':[]}
            part=apply_identity_review(first,checked,standard,confirm_slots=confirm_slots)
            source_known=source_identity_records(row,table,packet,standard,aliases)
            # 明确存在标准差别时，不能因旧metrics恰好填齐字段就保留为精确事实。
            retained=[]
            for fact in part['facts']:
                address=fact['source_reference']['location'].rpartition('!')[2]
                if address in source_known:
                    part['unresolved'].append({'cell':address,'reason':source_known[address]['standard_difference']})
                else:retained.append(fact)
            part['facts']=retained
            blocked={a for issue in checked['issues'] for a in issue['cells']}
            blocked.update(a for issue in row.get('unresolved',[]) for a in addresses(issue['range']))
            excluded={item['cell'] for item in part['excluded']}
            _require(not excluded&set(source_known),'原文身份证与非业务分类冲突')
            for address,identity in source_known.items():
                if address in checked['checked_cells'] and address not in blocked:
                    result['source_identities'].append(identity)
            _require(not {(table['sheet'],a) for a in packet['candidate_cells']}&processed,'重复采纳子表')
            processed.update((table['sheet'],a) for a in packet['candidate_cells'])
            result['facts'].extend(part['facts']);result['owners'].update(part['owners'])
            if confirm_slots:result['slot_bindings'].extend({'sheet':table['sheet'],**r} for r in part['slot_bindings'])
            for key in ('excluded','unresolved'):
                result[key].extend({'sheet':table['sheet'],**r} for r in part[key])
    table_areas={}
    for table in tables.values():table_areas.setdefault(table['sheet'],set()).update(addresses(table['range']))
    for sheet in snapshot['sheets']:
        physical=set(a for t in sheet.get('word_table_ranges',[]) for a in addresses(t['range']))
        for address,cell in sheet['cells'].items():
            if b_scope is not None and sheet['name']+'!'+address not in snapshot['semantic_candidates']:continue
            if (sheet['name'],address) in processed:continue
            if address not in table_areas.get(sheet['name'],set()) and address not in physical and (cell['value'] is None or not str(cell['value']).strip()):
                result['excluded'].append({'sheet':sheet['name'],'cell':address,'category':'empty_padding',
                    'reason':'整体阅读所列各表及原Word物理表以外的空白间隔','evidence_cells':[],
                    'recognition_provenance':{'method':'a_semantic_padding','record_id':definition_hash(reading)}})
            else:result['unresolved'].append({'sheet':sheet['name'],'cell':address,'reason':'此处尚未完成语义识别及核查'})
    result['outline']={'complete':all(any(t['sheet']==s['name'] and a in addresses(t['range']) for t in tables.values())
                                    for s in snapshot['sheets'] for a,c in s['cells'].items() if c['value'] is not None and str(c['value']).strip()),
                       'subjects':list(dict.fromkeys(t['subject'] for t in tables.values())),'subtables':reading['tables']}
    if b_scope is not None:
        result['outline']=deepcopy(b_outline)
        result['recognition_scope']=deepcopy(b_scope)
    result['coverage']=workbook_coverage(result,snapshot);result['complete']=result['coverage']['complete']
    validate_facts(result['facts'],standard,result['owners'])
    # 旧回执可重读；新识别的完成依据只有准确对应金标准的事实。
    facts={f['source_reference']['location'] for f in result['facts']}
    facts.update(r['sheet']+'!'+r['cell'] for r in result.get('slot_bindings',[]))
    source_known={f['source_reference']['location'] for f in result['source_identities']}
    _require(not facts&source_known,'同一格不能同时声明完全对应和存在标准差别')
    remaining=[u for u in result['unresolved'] if u['sheet']+'!'+u['cell'] not in source_known]
    if require_identity:remaining=result['unresolved']
    result['b_identity_coverage' if b_scope is not None else 'a_identity_coverage']={'financial_cells':len(facts|source_known),'unresolved_cells':len(remaining),
        'unresolved':remaining,'complete':not remaining and result['outline']['complete']}
    if b_scope is not None:
        from .表格编排 import required_coverage
        result['recognition_scope']['coverage']=required_coverage(result,standard)
    return result


def carry_role_corrections(review, earlier):
    result=deepcopy(review)
    for answer in earlier:
        for old in answer.get('tables',[]):
            if isinstance(old.get('issues'),list) and not any(r['id']==old['id'] for r in result.get('tables',[])):
                result.setdefault('tables',[]).append({'id':old['id'],'checked_ranges':[],'corrections':{},'issues':[]})
    for row in result.get('tables',[]):
        history=[old for answer in [*earlier,result] for old in answer.get('tables',[]) if old['id']==row['id']]
        # 旧回执未限定范围且明确返回问题数组，表示全表核查；后续局部核查不能缩小它。
        if any('checked_ranges' not in old and isinstance(old.get('issues'),list) for old in history):
            row.pop('checked_ranges',None)
        else:
            row['checked_ranges']=list(dict.fromkeys(span for old in history for span in old.get('checked_ranges',[])))
        previous=[fix for answer in earlier for old in answer.get('tables',[]) if old['id']==row['id'] for fix in old.get('corrections',{}).get('non_business',[])]
        if previous:
            corrections=row.setdefault('corrections',{})
            corrections['non_business']=previous+corrections.get('non_business',[])
        # 旧问题须明确解除，不能因后轮漏报而消失；只保留尚未解除的原范围。
        issues=[]
        for answer in [*earlier,result]:
            for old in answer.get('tables',[]):
                if old['id']!=row['id']:continue
                cleared={a for area in old.get('corrections',{}).get('resolved_ranges',[]) for a in addresses(area)}
                issues=[{**issue,'range':a} for issue in issues for a in addresses(issue['range']) if a not in cleared]
                issues.extend(deepcopy(old.get('issues',[])))
        row['issues']=list({json.dumps(issue,sort_keys=True,ensure_ascii=False):issue for issue in issues}.values())
    return result


def focused_review_payload(payload, focus, standard):
    """保留整表原文，只发送本次明确核查范围的已有身份及实际定义。"""
    result=deepcopy(payload);tables={t['id']:t for t in result['reading']}
    _require(set(focus)==set(tables),'局部核查须明确每张表的范围')
    _require(any(focus.values()),'本次核查范围为空')
    metrics=[]
    for row in result['recognized']['tables']:
        wanted={a for span in focus[row['id']] for a in addresses(span)}
        _require(wanted<=set(addresses(tables[row['id']]['range'])),'本次核查范围越表')
        data={a for span in tables[row['id']]['data_ranges'] for a in addresses(span)}
        def clip(spans, financial=True):
            declared=set().union(*(data if span=='*' else set(addresses(span)) for span in spans))
            return sorted(declared&wanted&(data if financial else wanted))
        for key in ('metrics','dimensions','source_identities'):
            if key not in row:continue
            row[key]=[{**item,'ranges':selected} for item in row.get(key,[]) if (selected:=clip(item['ranges']))]
        for key in ('unresolved','non_business'):
            if key in row:row[key]=[{**item,'range':a} for item in row[key] for a in clip([item['range']],financial=key!='non_business')]
        metrics.extend(item['id'] for item in row['metrics'])
        metrics.extend(item['closest_metric_id'] for item in row.get('source_identities',[]) if item.get('closest_metric_id'))
    metrics.extend(result.get('review_candidates',[]))
    _require(all(m in standard['metric'] and standard['metric'][m]['status']=='active' for m in metrics),'核查候选指标无效')
    result['standards']=standard_material(metrics,standard)
    result['review_focus']=deepcopy(focus)
    for key in ('program_issues','semantic_questions','previous_issues'):
        relevant=[]
        for issue in result.pop(key,[]):
            # 没有位置的具体问题仍须保留；不能把无法解析引用本身的错误也筛掉。
            if not isinstance(issue,dict) or 'id' not in issue:
                relevant.append(issue);continue
            wanted={a for span in focus.get(issue['id'],[]) for a in addresses(span)}
            if not wanted:continue
            field=next((name for name in ('cells','ranges','range') if name in issue),None)
            if field is None:
                relevant.append(issue);continue
            spans=[issue[field]] if field=='range' else issue[field]
            declared={a for span in spans for a in (wanted if span=='*' else addresses(span))}
            selected=sorted(declared&wanted)
            if field=='range':relevant.extend({**issue,'range':a} for a in selected)
            elif selected:relevant.append({**issue,field:selected})
        if relevant:result[key]=relevant
    return result


def semantic_read_input(snapshot, context, b_scope=None, b_outline=None):
    """A读整份，B读已定位范围；后续选标准、绑定和核查共用。"""
    tables=None
    if b_scope is not None:
        from .表格编排 import scoped_subtables
        snapshot['semantic_candidates']=set(b_scope['candidate_locations'])
        snapshot['semantic_known_targets']=deepcopy(b_scope['targets'])
        detailed={d['subtable_id'] for d in b_scope['details']}
        tables=[r for t in scoped_subtables(b_outline,b_scope) if t['id'] in detailed for r in t['ranges']]
    material,_=source_material(snapshot,context,tables)
    payload={'task':'a_semantic_read','source':material}
    if b_scope is not None:
        payload['candidate_locations']=b_scope['candidate_locations']
        payload['located_reading']=[{'id':d['subtable_id'],'structure':d['structure']} for d in b_scope['details']]
    return payload


def validate_scoped_reading(answer, snapshot, b_scope=None, b_outline=None):
    if b_scope is not None and not b_scope['candidate_locations']:
        _require(answer=={'tables':[],'notes':'未找到必要候选，原A需要继续保留。'},'无候选不能保存身份阅读结果')
        return deepcopy(answer)
    reading=validate_reading(answer,snapshot)
    if b_scope is not None:
        from .表格编排 import in_ranges, scoped_subtables
        detailed={d['subtable_id'] for d in b_scope['details']}
        allowed=[r for t in scoped_subtables(b_outline,b_scope) if t['id'] in detailed for r in t['ranges']]
        for table in reading['tables']:
            _require(all(in_ranges(table['sheet'],a,allowed) for a in addresses(table['range'])), 'B阅读超出已定位范围')
            _require(all(in_ranges(table['sheet'],a,allowed) for span in table.get('context_ranges',[]) for a in addresses(span)), 'B公共说明超出已展示范围')
        actual={t['sheet']+'!'+a for t in reading['tables'] for r in t['data_ranges'] for a in addresses(r)}
        _require(actual==set(b_scope['candidate_locations']), 'B阅读遗漏必要候选或增加了无关数据格')
    return reading


def batch_origin(batch, mapping, cache):
    """跨版本复用保留旧回执来源；业务材料仍由当前标准逐项重建核对。"""
    reference=batch.get('reused_from')
    if not reference:return mapping
    key=(reference['path'],reference['sha256'])
    if key not in cache:
        _require(file_hash(reference['path'])==reference['sha256'],'复用的旧语义映射发生变化')
        old=json.loads(Path(reference['path']).read_text(encoding='utf-8-sig'))
        _require(old['source_hash']==mapping['source_hash'] and old.get('source_context',[])==mapping.get('source_context',[]),
                 '复用的原件或上下文不一致')
        _require(file_hash(old['gold_path'])==old['gold_hash'],'复用的历史标准发生变化')
        cache[key]=old
    old=cache[key]
    fields=('binding','prior_reviews','review')
    proof=old.get('semantic_reading') or old.get('a_semantic_reading',{})
    _require(any(all(candidate.get(k,[])==batch.get(k,[]) for k in fields) for candidate in proof.get('batches',[])),
             '复用回执不属于所引用的旧映射')
    return batch_origin(next(candidate for candidate in proof['batches']
                             if all(candidate.get(k,[])==batch.get(k,[]) for k in fields)),old,cache)


def recognize_a(engine, source, context, progress=None, artifact_metadata=None, previous_result=None, *, b_scope=None, b_outline=None):
    """A/B共用：读原文→按语义检索→按范围确认→核查已有结果。保留旧A协议兼容。"""
    from .语义 import SemanticCancelled
    snapshot=read_workbook(source,context=context);materialize_candidate_blanks(snapshot)
    snapshot['semantic_formula_sources']=2
    snapshot['semantic_slot_bindings']=True
    engine.source_hash=snapshot['sha256']
    previous=None;origin_cache={}
    carrier='B' if b_scope is not None else 'A'
    proof_key='semantic_reading' if b_scope is not None else 'a_semantic_reading'
    if previous_result:
        from .统一勾稽 import load_fact_mapping
        candidate=json.loads(Path(previous_result).read_text(encoding='utf-8-sig'))
        if candidate.get(proof_key) and candidate.get('source_hash')==snapshot['sha256'] and candidate.get('source_context')==context:
            previous=load_fact_mapping(previous_result,engine.standard)
            if previous.get('source_identities'):
                # 只沿用已验真的原文阅读；旧“最近标准”检索和文字绑定均须修正。
                for field in ('selection','selection_additions','batches'):previous[proof_key][field]=[]
    read_input=semantic_read_input(snapshot,context,b_scope,b_outline)
    if b_scope is not None:
        from .B样式复用 import prepare_reuse
        prepare_reuse(snapshot,b_scope,b_outline,engine.standard)
    proof={'reading':[],'reading_check':[],'selection':[],'selection_additions':[],'batches':[]};timings=[];batches=[]
    def request(prompt,payload,validator):
        start=len(engine._review_receipts);before=dict(engine.usage);clock=time.perf_counter()
        previous_mode=engine.settings.get('thinking_mode','default')
        actual_mode=previous_mode
        engine.settings['thinking_mode']=actual_mode
        try:return engine._request(prompt.replace('A','B') if b_scope is not None else prompt,payload,validator)
        finally:
            engine.settings['thinking_mode']=previous_mode
            timings.append({'task':payload['task'],'thinking_mode':actual_mode,'seconds':round(time.perf_counter()-clock,3),
                            'usage':{k:v-before.get(k,0) for k,v in engine.usage.items()}})
    engine.log(carrier+'：先读原文及各表语义，本步不发送金标准')
    start=len(engine._review_receipts)
    prompt=READ_PROMPT
    if b_scope is not None:prompt+='\n已按待更新文件需要定位并比较；本次只读所给范围。data_ranges必须且只能覆盖candidate_locations，其他格用于理解上下文；不扩大为整本来源文件发证，不把范围外金额当作漏识别。每张表可增加context_ranges:[同一工作表内的公共表头或说明范围]，须来自已展示原文；不同业务表不因共用表头合并。实际含义以本次原文为准。'
    no_candidates=b_scope is not None and not b_scope['candidate_locations']
    reading=({'tables':[],'notes':'未找到必要候选，原A需要继续保留。'} if no_candidates else
             request(prompt,read_input,lambda x:validate_scoped_reading(x,snapshot,b_scope,b_outline)))
    proof['reading']=stage_proofs(engine,start)
    questions=reading_questions(reading,snapshot)
    if questions:
        engine.log(carrier+'：只补读数字范围或文字业务角色尚有疑点的子表：'+str(len(questions))+'个')
        start=len(engine._review_receipts)
        answer=request(READ_CHECK_PROMPT,reading_check_payload(reading,questions,snapshot,context),
                       lambda x: (apply_reading_check(reading,x,questions,snapshot),x)[1])
        proof['reading_check']=stage_proofs(engine,start)
        reading=apply_reading_check(reading,answer,questions,snapshot)
        validate_scoped_reading(reading,snapshot,b_scope,b_outline)
    destination=engine.work_dir/(carrier+'简洁语义进度')/str(time.time_ns())
    destination.mkdir(parents=True)
    def publish(final=False):
        _require(file_hash(source)==snapshot['sha256'] and file_hash(engine.gold_path)==engine.gold_hash,'识别期间原件或标准发生变化')
        result=assemble(snapshot,context,engine.standard,reading,batches,b_scope,b_outline,require_identity=True)
        result.update(schema_version=2,gold_path=engine.gold_path,gold_hash=engine.gold_hash,source_path=snapshot['path'],
                      source_hash=snapshot['sha256'],source_context=context,
                      stage_timings=deepcopy(timings),usage=dict(engine.usage),
                      request_receipts=proof['reading']+proof['reading_check']+proof['selection']+proof['selection_additions']+[p for b in proof['batches'] for k in ('binding','prior_reviews','review') for p in b.get(k,[])])
        result[proof_key]=deepcopy(proof)
        result['identity_version']=1
        result['cell_identity_version']=1
        result['formula_source_version']=2
        result['slot_binding_version']=1
        if b_scope is not None:
            result['dimension_vocabulary_version']=1
            result['style_reuse_version']=1
        if b_scope is None:result['recognition_scope']={'mode':'a_semantics'}
        result.update(deepcopy(artifact_metadata or {}))
        path=destination/(('B语义映射.json' if b_scope is not None else '整份A语义映射.json') if final else f'进度_{time.time_ns()}.json')
        result['mapping_path']=str(path);engine._save(path,result)
        if progress:progress(result)
        return result
    publish()
    if no_candidates:return publish(final=True)
    engine.log(carrier+'：按已经读出的语义，在金标准简明目录中查找对应定义')
    start=len(engine._review_receipts)
    select_input=selection_payload(reading,snapshot,engine.standard);select_prompt=SELECT_PROMPT
    same_selection=(previous or {}).get('gold_hash')==engine.gold_hash and (b_scope is None or
        all(previous.get('recognition_scope',{}).get(k)==v for k,v in b_scope.items() if k!='coverage'))
    if same_selection and previous.get(proof_key,{}).get('selection'):
        old_task,_=_stage_receipt(previous[proof_key]['selection'][0],previous,'a_semantic_select')
        select_input=old_task['payload'];select_prompt=old_task['system']
    selection=request(select_prompt,select_input,lambda x:validate_selection(x,reading,engine.standard))
    proof['selection']=stage_proofs(engine,start)
    selected={r['id']:r for r in selection['tables']}
    detailed=set(selected) if '原表文字' in select_input else set()
    for reference in ((previous or {}).get(proof_key,{}).get('selection_additions',[]) if same_selection else []):
        task,answer=_stage_receipt(reference,previous,'a_semantic_select')
        addition=validate_selection(answer,task['payload']['reading'],engine.standard)
        selected.update((r['id'],r) for r in addition['tables']);detailed.update(r['id'] for r in addition['tables'])
        proof['selection_additions'].append(deepcopy(reference))
    reviewed=set()
    for batch in (previous or {}).get(proof_key,{}).get('batches',[]):
        origin=batch_origin(batch,previous,origin_cache)
        for reference in batch['review']:
            _,answer=_stage_receipt(reference,origin,'a_semantic_review')
            reviewed.update(r['id'] for r in answer['tables'] if isinstance(r,dict) and isinstance(r.get('issues'),list))
    pending=[t for t in reading['tables'] if t['id'] not in detailed|reviewed]
    if pending:
        partial={'tables':pending}
        engine.log(carrier+'：仅给尚未核查的子表补充原表项目名称，避免概括遗漏相关标准：'+str(len(pending))+'张')
        start=len(engine._review_receipts)
        addition=request(SELECT_PROMPT,selection_payload(partial,snapshot,engine.standard),lambda x:validate_selection(x,partial,engine.standard))
        proof['selection_additions'].extend(stage_proofs(engine,start))
        selected.update((r['id'],r) for r in addition['tables'])
    selection={'tables':[selected[t['id']] for t in reading['tables']]}
    # 新请求逐张完整表处理；旧已核查分组保持原状，不为分组变化重跑。
    saved_groups={}
    for saved in (previous or {}).get(proof_key,{}).get('batches',[]):
        if not saved['review']:continue
        origin=batch_origin(saved,previous,origin_cache)
        task,_=_stage_receipt(saved['binding'][0],origin,'a_semantic_bind')
        _,answer=_stage_receipt(saved['review'][0],origin,'a_semantic_review')
        tables=task['payload']['reading']
        if (all(t in reading['tables'] for t in tables) and
            {t['id'] for t in answer.get('tables',[]) if isinstance(t,dict) and isinstance(t.get('issues'),list)}=={t['id'] for t in tables}):
            saved_groups[tables[0]['id']]=tables
    groups=[];grouped=set()
    for table in reading['tables']:
        if table['id'] in grouped:continue
        group=saved_groups.get(table['id'],[table]);groups.append(group);grouped.update(t['id'] for t in group)
    for tables in groups:
        engine._check_cancel()
        payload,aliases=binding_payload(tables,selection,snapshot,context,engine.standard,reading_notes=reading.get('notes',''))
        reused=False;saved_first=None
        for saved in (previous or {}).get(proof_key,{}).get('batches',[]):
            origin=batch_origin(saved,previous,origin_cache)
            task,binding=_stage_receipt(saved['binding'][0],origin,'a_semantic_bind')
            if task['payload']['reading']!=tables:continue
            # 已验真的旧请求可能包含较宽的候选；候选精简不使已有成果失效。
            old_payload=task['payload']
            old_source,_=source_material(snapshot,context,tables,semantic_only='数值处理' in old_payload['source'])
            supplied=[m['id'] for m in old_payload['standards']['metrics']]
            if old_payload['standards']!=standard_material(supplied,engine.standard):continue
            if (old_payload.get('style_templates')==payload.get('style_templates') and
                old_payload['source']==old_source and {m['id'] for m in payload['standards']['metrics']}<=set(supplied)):
                payload=deepcopy(old_payload)
            binding=validate_bindings(binding,tables,snapshot,context,engine.standard,candidate=True)
            earlier=[];known_questions=[]
            for reference in saved.get('prior_reviews',[]):
                prior_task,old_review=_stage_receipt(reference,origin,'a_semantic_review')
                known_questions.extend(prior_task['payload'].get('semantic_questions',[]))
                old_review=validate_review(old_review,tables,snapshot,context,prior_task['payload'].get('review_focus'))
                binding=reviewed_binding(binding,old_review,tables,snapshot);earlier.append(old_review)
            current_review=None
            if saved['review']:
                old_review_task,current_review=_stage_receipt(saved['review'][0],origin,'a_semantic_review')
                known_questions.extend(old_review_task['payload'].get('semantic_questions',[]))
                current_review=validate_review(current_review,tables,snapshot,context,old_review_task['payload'].get('review_focus'))
            if task['payload']==payload and origin['gold_hash']==engine.gold_hash:
                saved_first=(reviewed_binding(binding,current_review,tables,snapshot) if current_review else binding,
                             deepcopy(saved['binding']),deepcopy(saved.get('prior_reviews',[])+saved['review']),earlier+([current_review] if current_review else []))
            if not saved['review']:continue
            if task['payload']!=payload or {r['id'] for r in current_review['tables']}!={t['id'] for t in tables}:continue
            checked=carry_role_corrections(current_review,earlier)
            if any(set(table_packet(t,snapshot,context)['candidate_cells'])-
                   set(review_for_table(next(r for r in checked['tables'] if r['id']==t['id']),t,table_packet(t,snapshot,context),aliases)['checked_cells']) for t in tables):continue
            # 全范围身份已核查；父项维度的通用提醒不构成新错误，不能因补查使用自然语言就重查整表。
            questions=[q for q in semantic_questions(binding,tables,snapshot,engine.standard)
                       if q.get('check')!='same_identity_different_rows' and not (q.get('dimension') and not q.get('check'))]
            questions+=duplicate_identity_questions(reviewed_binding(binding,current_review,tables,snapshot),tables,snapshot,engine.standard)
            if any(q not in known_questions for q in questions):continue
            retained=deepcopy(saved)
            if origin['gold_hash']!=engine.gold_hash:
                retained['reused_from']={'path':previous['mapping_path'],'sha256':previous['mapping_hash']}
            proof['batches'].append(retained);batches.append((binding,current_review,earlier))
            engine.usage['reused_confirmed_tables']=engine.usage.get('reused_confirmed_tables',0)+len(tables)
            engine.log(carrier+'：沿用已完成核查的子表，已有未决原样保留：'+'、'.join(t['subject'] for t in tables))
            reused=True;break
        if reused:continue
        engine.log(carrier+'：确认这些完整子表的共同语义及行列差异：'+'、'.join(t['subject'] for t in tables))
        start=len(engine._review_receipts)
        try:
            if saved_first:
                binding,binding_proofs,prior_proofs,earlier=saved_first
                engine.usage['reused_first_tables']=engine.usage.get('reused_first_tables',0)+len(tables)
                engine.log(carrier+'：沿用已有第一轮，只处理其核查问题')
            else:
                prior_proofs=[];earlier=[]
                binding=request(BIND_PROMPT,payload,lambda x:validate_bindings(x,tables,snapshot,context,engine.standard,candidate=True,require_identity=True))
                binding_proofs=stage_proofs(engine,start)
                _,raw_binding=_stage_receipt(binding_proofs[0],{'source_hash':snapshot['sha256'],'gold_hash':engine.gold_hash},'a_semantic_bind')
                reused_cells=sum(len(row.get('style_reuse',[])) for row in raw_binding['tables'])
                if reused_cells:
                    engine.usage['style_reused_cells']=engine.usage.get('style_reused_cells',0)+reused_cells
                    engine.log(carrier+'：已沿用五样式中的指标及固定业务范围：'+str(reused_cells)+'格；本次上下文另行绑定')
        except SemanticCancelled:raise
        except Exception as error:
            engine.log('本组请求失败，停止扩量并保留已完成结果：'+engine._redact(str(error)));publish();break
        review_payload=review_material(payload,binding,tables,snapshot,context,engine.standard,earlier)
        review_payload.setdefault('reading_notes',deepcopy(reading.get('notes','')))
        explicit=[r for answer in earlier for r in answer.get('tables',[]) if 'checked_ranges' in r]
        if explicit:
            focus={t['id']:sorted(set(table_packet(t,snapshot,context)['candidate_cells'])-
                {a for r in explicit if r['id']==t['id'] for span in r['checked_ranges'] for a in addresses(span)}) for t in tables}
            if any(focus.values()):review_payload=focused_review_payload(review_payload,focus,engine.standard)
        engine.log(carrier+'：核查这些表已确定的含义及目标覆盖，不重新生成身份')
        start=len(engine._review_receipts)
        def check_review(answer):
            answer=validate_review(answer,tables,snapshot,context,review_payload.get('review_focus'),require_identity=True)
            _require(all(isinstance(row.get('issues'),list) for row in answer['tables']),'核查必须明确返回issues数组，只有表号不等于已经核查')
            validate_bindings(reviewed_binding(binding,answer,tables,snapshot),tables,snapshot,context,engine.standard)
            index={t['id']:t for t in tables}
            for row in answer['tables']:
                table=index[row['id']]
                apply_role_corrections({},row,table,table_packet(table,snapshot,context),engine.standard,aliases)
            return deepcopy(answer)
        failed_review=False
        try:
            review=request(REVIEW_PROMPT,review_payload,check_review)
            if review.get('incomplete'):
                engine.log('核查仅返回部分有效子表，保留实际记录；缺少部分仍未核查，停止续发')
                failed_review=True
        except SemanticCancelled:raise
        except Exception as error:
            engine.log('本组核查未完成，保留原结果及未决：'+engine._redact(str(error)));review={'tables':[]};failed_review=True
        proof['batches'].append({'binding':binding_proofs,'prior_reviews':prior_proofs,'review':stage_proofs(engine,start)})
        batches.append((binding,review,earlier));publish()
        if failed_review:break
    return publish(final=True)


def supplied_review_standards(original, supplied, standard):
    """补查可增加相关定义，实际发送的每项仍须与当前金标准逐字一致。"""
    ids=[m['id'] for m in supplied['metrics']]
    _require({m['id'] for m in original['metrics']}<=set(ids),'补查遗漏原有候选标准')
    _require(all(m in standard['metric'] and standard['metric'][m]['status']=='active' for m in ids),'补查引用无效标准')
    expected=standard_material(ids,standard)
    _require(expected==supplied,'补查定义与正式金标准不符')
    return expected


def verify_mapping(mapping, standard, *, rebuild=False):
    """从保存的原始请求和回执重建结果，不能只信任最终JSON中的完成声明。"""
    from .语义 import _stable_content
    from .统一语义 import read_standard_v2
    historical=deepcopy(standard) if standard['sha256']==mapping['gold_hash'] else read_standard_v2(mapping['gold_path'])
    historical['_historical_report_scope']=mapping.get('cell_identity_version')!=1
    _require(historical['sha256']==mapping['gold_hash'],'原识别标准已经改变')
    context=mapping.get('source_context',[])
    snapshot=read_workbook(mapping['source_path'],context=context);materialize_candidate_blanks(snapshot)
    snapshot['semantic_formula_sources']=mapping.get('formula_source_version',0)
    snapshot['semantic_slot_bindings']=mapping.get('slot_binding_version')==1
    _require(snapshot['sha256']==mapping['source_hash'],'识别原件发生变化')
    scoped=bool(mapping.get('semantic_reading'))
    b_scope=mapping['recognition_scope'] if scoped else None
    b_outline=mapping['outline'] if scoped else None
    proof=mapping['semantic_reading' if scoped else 'a_semantic_reading']
    read_input=semantic_read_input(snapshot,context,b_scope,b_outline)
    if scoped and mapping.get('style_reuse_version')==1:
        from .B样式复用 import prepare_reuse
        prepare_reuse(snapshot,b_scope,b_outline,historical)
    if mapping.get('dimension_vocabulary_version')!=1:snapshot.pop('semantic_known_targets',None)
    receipt_mapping=mapping;origin_cache={}
    def read(proofs,task_name,prompt,payload):
        _require(len(proofs)==1,'每个阶段只采纳一个完整回执')
        task,answer=_stage_receipt(proofs[0],receipt_mapping,task_name)
        # 提示词原文已包含在实际任务哈希中；以后改进措辞不能使已保存的合法成果失效。
        _require(isinstance(task.get('system'),str) and bool(task['system'].strip()) and
                 task['payload']==_stable_content(payload),'语义结果与实际请求材料不符')
        return answer
    no_candidates=scoped and not b_scope['candidate_locations']
    if no_candidates:
        _require(not any(proof.values()),'无候选不能伪造语义回执')
        reading={'tables':[],'notes':'未找到必要候选，原A需要继续保留。'}
    else:reading=validate_scoped_reading(read(proof['reading'],'a_semantic_read',READ_PROMPT,read_input),snapshot,b_scope,b_outline)
    if proof.get('reading_check'):
        questions=reading_questions(reading,snapshot)
        answer=read(proof['reading_check'],'a_semantic_read_check',READ_CHECK_PROMPT,
                    reading_check_payload(reading,questions,snapshot,context))
        reading=apply_reading_check(reading,answer,questions,snapshot)
        validate_scoped_reading(reading,snapshot,b_scope,b_outline)
    batches=[]
    if proof['selection']:
        old_task,_=_stage_receipt(proof['selection'][0],mapping,'a_semantic_select')
        select_input=selection_payload(reading,snapshot,historical,include_basis=any('basis' in m for m in old_task['payload']['catalog']))
        if '原表文字' not in old_task['payload']:select_input.pop('原表文字')
        selection=validate_selection(read(proof['selection'],'a_semantic_select',SELECT_PROMPT,select_input),reading,historical)
        index={t['id']:t for t in reading['tables']}
        selected={r['id']:r for r in selection['tables']}
        for reference in proof.get('selection_additions',[]):
            task,_=_stage_receipt(reference,mapping,'a_semantic_select')
            partial=task['payload']['reading']
            _require(all(t==index.get(t['id']) for t in partial['tables']),'补检索改变了原表理解')
            addition=validate_selection(read([reference],'a_semantic_select',SELECT_PROMPT,
                selection_payload(partial,snapshot,historical,include_basis=any('basis' in m for m in task['payload']['catalog']))),partial,historical)
            selected.update((r['id'],r) for r in addition['tables'])
        selection={'tables':[selected[t['id']] for t in reading['tables']]}
        for item in proof['batches']:
            receipt_mapping=batch_origin(item,mapping,origin_cache)
            _require(len(item['binding'])==1,'缺少唯一身份确认回执')
            task,_=_stage_receipt(item['binding'][0],receipt_mapping,'a_semantic_bind')
            tables=task['payload']['reading']
            _require(all(t==index.get(t['id']) for t in tables),'身份确认没有沿用原表理解')
            payload,_=binding_payload(tables,selection,snapshot,context,historical,semantic_only='数值处理' in task['payload']['source'],
                                      reading_notes=reading.get('notes','') if 'reading_notes' in task['payload'] else None)
            # 按该次回执实际提供的候选重建，补充同对象候选后不能使已合法保存的旧结果失效。
            supplied=[m['id'] for m in task['payload']['standards']['metrics']]
            required={m for r in selection['tables'] if r['id'] in {t['id'] for t in tables} for m in r['metric_ids']}
            _require(required<=set(supplied),'实际请求遗漏了所选标准')
            payload['standards']=standard_material(supplied,historical)
            binding=validate_bindings(read(item['binding'],'a_semantic_bind',BIND_PROMPT,payload),tables,snapshot,context,historical,candidate=True)
            earlier=[]
            for reference in item.get('prior_reviews',[]):
                old_task,_=_stage_receipt(reference,receipt_mapping,'a_semantic_review')
                prior_payload={**payload,'task':'a_semantic_review','recognized':binding}
                if 'reading_notes' in old_task['payload']:prior_payload['reading_notes']=deepcopy(reading.get('notes',''))
                if old_task['payload'].get('共同声明合并'):
                    prior_payload.update(recognized=shared_binding(binding),共同声明合并=True)
                focus=old_task['payload'].get('review_focus')
                if 'review_candidates' in old_task['payload']:prior_payload['review_candidates']=old_task['payload']['review_candidates']
                if focus is not None:prior_payload=focused_review_payload(prior_payload,focus,historical)
                prior_payload['standards']=supplied_review_standards(prior_payload['standards'],old_task['payload']['standards'],historical)
                for key in ('program_issues','semantic_questions','previous_issues','返回骨架'):
                    if key in old_task['payload']:prior_payload[key]=old_task['payload'][key]
                old_review=validate_review(read([reference],'a_semantic_review',REVIEW_PROMPT,prior_payload),tables,snapshot,context,focus)
                binding=reviewed_binding(binding,old_review,tables,snapshot);earlier.append(old_review)
            review_payload={**payload,'task':'a_semantic_review','recognized':binding}
            if item['review']:
                old_task,_=_stage_receipt(item['review'][0],receipt_mapping,'a_semantic_review')
                if 'reading_notes' in old_task['payload']:review_payload['reading_notes']=deepcopy(reading.get('notes',''))
                if old_task['payload'].get('共同声明合并'):
                    review_payload.update(recognized=shared_binding(binding),共同声明合并=True)
                focus=old_task['payload'].get('review_focus')
                if 'review_candidates' in old_task['payload']:review_payload['review_candidates']=old_task['payload']['review_candidates']
                if focus is not None:review_payload=focused_review_payload(review_payload,focus,historical)
                review_payload['standards']=supplied_review_standards(review_payload['standards'],old_task['payload']['standards'],historical)
                if 'program_issues' in old_task['payload']:
                    # 诊断文案只是辅助提示；原始文案由任务哈希保护，业务原文、标准、身份及结果仍重新核验。
                    review_payload['program_issues']=old_task['payload']['program_issues']
                if 'semantic_questions' in old_task['payload']:review_payload['semantic_questions']=old_task['payload']['semantic_questions']
                if '返回骨架' in old_task['payload']:review_payload['返回骨架']=old_task['payload']['返回骨架']
                if 'previous_issues' in old_task['payload']:review_payload['previous_issues']=old_task['payload']['previous_issues']
            review=validate_review(read(item['review'],'a_semantic_review',REVIEW_PROMPT,
                review_payload),tables,snapshot,context,review_payload.get('review_focus')) if item['review'] else {'tables':[]}
            batches.append((binding,review,earlier))
    else:_require(not proof['batches'],'尚未检索标准不能已有身份结果')
    expected=assemble(snapshot,context,historical,reading,batches,b_scope,b_outline,require_identity=mapping.get('identity_version')==1)
    if rebuild:return expected
    for key in ('facts','owners','excluded','unresolved','outline','coverage','complete'):
        _require(mapping.get(key)==expected[key],'保存结果与原语义回执不符：'+key)
    _require(mapping.get('slot_bindings',[])==expected.get('slot_bindings',[]),'保存结果与原语义回执不符：slot_bindings')
    for key in ('source_identities','b_identity_coverage' if scoped else 'a_identity_coverage'):
        if key in mapping or expected['source_identities']:
            _require(mapping.get(key)==expected[key],'保存结果与原语义回执不符：'+key)
