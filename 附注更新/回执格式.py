"""生成HTTP回执格式约束；不修改任务、缓存身份或财务语义校验。"""
from __future__ import annotations

import copy


def _strings(values):
    return list(dict.fromkeys(value for value in values if isinstance(value,str) and value))


def _text():
    return {"type":"string","minLength":1}


def _enum(values):
    values=_strings(values)
    return {"type":"string","enum":values} if values else {"type":"string"}


def _array(items, *, count=None):
    result={"type":"array","items":items}
    if count is not None:result.update(minItems=count,maxItems=count)
    return result


def _ids(values, *, count=None):
    values=_strings(values);result=_array(_enum(values),count=count)
    # JSON Schema的enum不能是空数组；没有可选ID时只允许返回空清单。
    if not values:result["maxItems"]=0
    return result


def _object(properties,required=None):
    return {"type":"object","properties":properties,
            "required":list(properties) if required is None else list(required),"additionalProperties":False}


def _evidence(payload):
    cells={**payload.get("context_cells",{}),**payload.get("cells",{})}
    return [address for address,item in cells.items() if isinstance(item,dict) and item.get("kind")=="text"]


def _areas(evidence, refinement=False):
    fields={"range":_text(),"decision":_enum(["include","out_of_scope","uncertain"]),
            "reason":_text(),"evidence_cells":_ids(evidence)}
    required=list(fields)
    if refinement:
        # 仅未决范围的再审可解释同词不同义；原文引用及双轮判断仍由业务校验器核实。
        target_evidence=_array(_text());target_evidence["minItems"]=1
        fields["semantic_comparison"]=_object({"source_business":_text(),"target_business":_text(),
            "difference_kind":_enum(["business_population","indicator","value_type"]),
            "reason":_text(),"target_evidence":target_evidence})
    return _array(_object(fields,required))


def _scope_cells(payload):
    return _object({"cells":_object({address:{"$ref":"#/$defs/scope_record"}
        for address in _strings(payload.get("candidate_cells",[]))})})


def _fixed_scope_schema(payload):
    tasks=payload.get("tasks",[]) if payload.get("task")=="source_scope_batch" else [payload]
    if not tasks or any(task.get("scope_response_mode")!="cells_v1" for task in tasks):
        raise ValueError("固定格范围任务不能与旧范围任务混用")
    record=_areas([a for task in tasks for a in _evidence(task)],
                  refinement=any("refinement" in task for task in tasks))["items"]
    record["properties"].pop("range");record["required"].remove("range")
    if payload.get("task")=="source_scope_batch":
        identifiers=[task.get("packet_id") for task in tasks]
        if any(not isinstance(i,str) or not i for i in identifiers) or len(set(identifiers))!=len(identifiers):
            raise ValueError("固定格范围任务包编号无效或重复")
        schema=_object({"packets":_object({task["packet_id"]:_scope_cells(task) for task in tasks})})
    else:schema=_scope_cells(payload)
    schema["$defs"]={"scope_record":record}
    return schema


def _normalize_scope_cells(payload,result):
    if not isinstance(result,dict) or set(result)!={"cells"} or not isinstance(result["cells"],dict):
        raise ValueError("固定格范围回执只允许cells对象，不生成坐标范围")
    owned=payload.get("candidate_cells",[]);records=result["cells"]
    if len(set(owned))!=len(owned) or set(records)!=set(owned):
        raise ValueError("固定格范围回执存在漏格、包外格或重复候选")
    evidence=set(_evidence(payload));areas=[]
    for address in owned:
        item=records[address]
        required={"decision","reason","evidence_cells"}
        optional={"semantic_comparison"} if "refinement" in payload else set()
        if not isinstance(item,dict) or not required<=set(item) or set(item)-(required|optional):
            raise ValueError(address+"：固定格范围判断字段无效")
        if item["decision"] not in {"include","out_of_scope","uncertain"}:
            raise ValueError(address+"：范围分类无效")
        if not isinstance(item["reason"],str) or not item["reason"].strip():
            raise ValueError(address+"：范围判断缺少业务依据")
        cited=item["evidence_cells"]
        if not isinstance(cited,list) or any(not isinstance(a,str) or a not in evidence for a in cited):
            raise ValueError(address+"：范围判断引用本包之外或非文字证据")
        areas.append({"range":address,**copy.deepcopy(item)})
    return {"areas":areas}


def _fixed_scope_mode(payload):
    return payload.get("scope_response_mode")=="cells_v1" or (
        payload.get("task")=="source_scope_batch" and any(
            task.get("scope_response_mode")=="cells_v1" for task in payload.get("tasks",[])))


def _layout(payload):
    shown=_strings(payload.get("cells",{}))
    notes=_strings(item.get("id") for item in payload.get("notes",[]) if isinstance(item,dict))
    note_ids=_ids(notes)
    if notes:note_ids.update(minItems=1,maxItems=1)
    tables=_array(_object({"range":_text(),"note_ids":note_ids,
        "scope":_enum(["consolidated","parent","standalone"]),"table_semantic":_text(),"header_cells":_ids(shown)}))
    if not notes:tables["maxItems"]=0
    evidence=_ids(shown)
    if shown:evidence["minItems"]=1
    nonbusiness=_array(_object({"range":_text(),"category":_enum(["header","label","title","unit","empty_padding","annotation"]),
        "reason":_text(),"evidence_cells":evidence}))
    if not shown:nonbusiness["maxItems"]=0
    unresolved=_array(_object({"range":_text(),"reason":_text(),
                               "evidence_cells":{**evidence,"minItems":1}}))
    if not shown:unresolved["maxItems"]=0
    # 三类区域的真实范围、覆盖互斥及证据含义仍由原布局校验器判断。
    return _object({"tables":tables,"non_business":nonbusiness,"unresolved":unresolved,
                    "carry_context":{"type":"string"}},["tables","non_business","unresolved"])


def _layout_owned(payload):
    owned=payload.get("candidate_cells",[])
    if (not isinstance(owned,list) or any(not isinstance(a,str) or not a for a in owned)
            or len(set(owned))!=len(owned) or any(a not in payload.get("cells",{}) for a in owned)):
        raise ValueError("固定格布局候选地址无效、重复或未提供原格")
    return owned


def _layout_cell_record(payload):
    shown=_strings([*payload.get("cells",{}),*payload.get("context",{}).get("preceding_cells",{})])
    notes=_strings(n.get("id") for n in payload.get("notes",[]) if isinstance(n,dict))
    evidence={**_ids(shown),"minItems":1}
    business=_object({"kind":_enum(["business"]),"reason":_text(),"note_id":_enum(notes),
        "scope":_enum(["consolidated","parent","standalone"]),"header_cells":evidence,"table_semantic":_text()})
    nonbusiness=_object({"kind":_enum(["non_business"]),"reason":_text(),
        "category":_enum(["header","label","title","unit","empty_padding","annotation"]),"evidence_cells":evidence})
    unresolved=_object({"kind":_enum(["unresolved"]),"reason":_text(),"evidence_cells":evidence})
    # 无科目目录时不能生成业务归属；有目录时共享三分支，不为每格重复展开。
    return {"oneOf":([business] if notes else [])+[nonbusiness,unresolved]}


def _layout_cells(payload):
    schema=_object({"cells":_object({a:{"$ref":"#/$defs/layout_record"} for a in _layout_owned(payload)})})
    schema["$defs"]={"layout_record":_layout_cell_record(payload)}
    return schema


def _normalize_layout_cells(payload,result):
    owned=_layout_owned(payload)
    if not isinstance(result,dict) or set(result)!={"cells"} or not isinstance(result["cells"],dict):
        raise ValueError("固定格布局回执只允许cells对象，不生成矩形或区域清单")
    records=result["cells"]
    if set(records)!=set(owned):raise ValueError("固定格布局回执存在漏格或包外格")
    branches={b["properties"]["kind"]["enum"][0]:b for b in _layout_cell_record(payload)["oneOf"]}
    normalized={"tables":[],"non_business":[],"unresolved":[]}
    for address in owned:
        item=records[address]
        if not isinstance(item,dict) or not isinstance(item.get("kind"),str) or item["kind"] not in branches:
            raise ValueError(address+"：固定格布局kind无效或无可用科目")
        branch=branches[item["kind"]]
        if set(item)!=set(branch["required"]):raise ValueError(address+"：固定格布局字段缺失或多余")
        for name,rule in branch["properties"].items():
            value=item[name]
            if rule["type"]=="string":
                if not isinstance(value,str) or not value.strip() or ("enum" in rule and value not in rule["enum"]):
                    raise ValueError(address+"：固定格布局字段无效："+name)
            else:
                allowed=rule["items"].get("enum",[])
                if not isinstance(value,list) or not value or any(not isinstance(a,str) or a not in allowed for a in value):
                    raise ValueError(address+"：固定格布局证据须引用已展示的真实原格："+name)
        converted={"range":address,**copy.deepcopy({k:v for k,v in item.items() if k!="kind"})}
        if item["kind"]=="business":
            converted["note_ids"]=[converted.pop("note_id")]
            normalized["tables"].append(converted)
        else:normalized[item["kind"]].append(converted)
    return normalized


def _classify_record(payload):
    slots=_strings(item.get("id") for item in payload.get("gold_slots",[]) if isinstance(item,dict))
    shown=_strings(payload.get("cells",{}))
    mapping=_object({"kind":_enum(["mapping"]),"reason":_text(),"slot_id":_enum(slots),
        "scope":_enum(["consolidated","parent","standalone"]),
        "dimensions":{"type":"object","additionalProperties":True},"semantic_field":_text(),
        "value_type":_enum(["monetary","number","percentage","text","date","boolean","enum"]),
        "row_label":{"type":"string"},"column_label":{"type":"string"},
        "metric_evidence":_ids(shown),"evidence_cells":_ids(shown)},
        ["kind","reason","slot_id","scope","dimensions","semantic_field","value_type","row_label","column_label"])
    excluded=_object({"kind":_enum(["excluded"]),"reason":_text(),
        "category":_enum(["header","label","title","unit","empty_padding","annotation"]),
        "evidence_cells":_ids(shown)},["kind","reason","category"])
    unresolved=_object({"kind":_enum(["unresolved"]),"reason":_text()},["kind","reason"])
    # 三种状态只在共享定义中展开一次；各格引用同一结构，不能省略所属状态的必填内容。
    return {"oneOf":[mapping,excluded,unresolved]}


def _classify(payload):
    owned=_strings(payload.get("candidate_cells",[]))
    schema=_object({"cells":_object({address:{"$ref":"#/$defs/cell_record"} for address in owned})})
    schema["$defs"]={"cell_record":_classify_record(payload)}
    return schema


def _unified_dimension(definition):
    """由维度定义生成取值结构，不填入任何样例或客户的实际维度值。"""
    kind=definition['value_type']
    if kind=='selection':
        branches=[]
        for mode, extras in (('all',{}),('members',{'members':{**_array(_text()),'minItems':1}}),
                             ('predicate',{'predicate':_text()}),
                             ('complement',{'base':_text(),'excluded':{**_array(_text()),'minItems':1}}),('unknown',{})):
            branches.append(_object({'domain':_enum([definition['domain']]),'mode':_enum([mode]),
                'completeness':_enum(['unknown'] if mode=='unknown' else ['complete','partial','unknown']),**extras}))
        result={'anyOf':branches}
    elif kind=='period':
        date={'type':'string','pattern':r'^\d{4}-\d{2}-\d{2}$'}
        result={'anyOf':[_object({'kind':_enum(['instant']),'date':date}),
                         _object({'kind':_enum(['duration']),'start':date,'end':date})]}
    elif kind=='number':result={'anyOf':[{'type':'number'},{'type':'string','pattern':r'^-?\d+(\.\d+)?$'}]}
    elif kind=='boolean':result={'type':'boolean'}
    else:result=_text()
    if 'allowed_values' in definition:result={**result,'enum':definition['allowed_values']}
    return result


def _unified_classify(payload):
    structured=payload.get('response_contract')=='unified_structure_v1'
    compact=structured or payload.get('response_contract')=='unified_dimension_refs_v1'
    definitions={row['id']:row for row in payload['dimensions']}
    shown=[address for address,item in payload['cells'].items() if item.get('value') is not None and str(item['value']).strip()]
    shared={'evidence':{**_ids(shown+list(payload.get('context_evidence',{}))),'minItems':1}}
    dimension_refs={}
    # 相同取值格式和证据目录只声明一次，展开后与原逐指标格式完全相同。
    for index,(identifier,definition) in enumerate(definitions.items()):
        name=f'dimension_{index}'
        shared[name]=_unified_dimension(definition)
        dimension_refs[identifier]={'$ref':'#/$defs/'+name}
    evidence={'$ref':'#/$defs/evidence'}
    branches=[]
    for metric in payload['metrics']:
        fields={ref['id']:dimension_refs[ref['id']] for ref in metric['dimension_refs']}
        required=[ref['id'] for ref in metric['dimension_refs'] if ref['required']]
        if compact:
            refs={'type':'array','items':{'type':'integer','minimum':0},'uniqueItems':True}
            branches.append(_object({'kind':_enum(['metric_fact']),'metric_id':_enum([metric['id']]),
                'dimension_refs':refs,'reason':_text()}))
            branches.append(_object({'kind':_enum(['partial_metric_fact']),'metric_id':_enum([metric['id']]),
                'dimension_refs':refs,'missing_dimensions':_ids(required),'evidence_cells':evidence,'reason':_text()}))
            continue
        branches.append(_object({'kind':_enum(['metric_fact']),'metric_id':_enum([metric['id']]),
            'dimensions':_object(fields,required),'dimension_evidence':_object({key:evidence for key in fields},required),
            'reason':_text()}))
        branches.append(_object({'kind':_enum(['partial_metric_fact']),'metric_id':_enum([metric['id']]),
            'dimensions':_object(fields,[]),'dimension_evidence':_object({key:evidence for key in fields},[]),
            'missing_dimensions':_ids(required),'evidence_cells':evidence,'reason':_text()}))
    branches.append(_object({'kind':_enum(['unresolved']),'reason':_text(),'evidence_cells':evidence}))
    branches.append(_object({'kind':_enum(['non_business']),'category':_enum(['title','header','label','unit','annotation','empty_padding']),
                            'reason':_text(),'evidence_cells':evidence}))
    root={'cells':_object({address:{'$ref':'#/$defs/unified_record'} for address in _layout_owned(payload)})}
    if compact:
        root['dimension_values']=_array({'anyOf':[_object({'dimension_id':_enum([identifier]),
            'value':reference,'evidence':evidence}) for identifier,reference in dimension_refs.items()]})
    if structured:
        root['cells']['required']=[]
        patch=_object({'dimension_refs':{'type':'array','items':{'type':'integer','minimum':0},'uniqueItems':True},
                       'missing_dimensions':_ids(definitions)},required=[])
        root['regions']=_array(_object({'range':_text(),'record':{'$ref':'#/$defs/unified_record'},
            'rows':_array(_object({'indices':{'type':'array','items':{'type':'integer','minimum':1},'minItems':1,'uniqueItems':True},'record':patch})),
            'columns':_array(_object({'indices':{'type':'array','items':{'type':'string','pattern':'^[A-Z]{1,3}$'},'minItems':1,'uniqueItems':True},'record':patch}))}))
    schema=_object(root)
    schema['$defs']={'unified_record':{'anyOf':branches},**shared}
    return schema


def transport_instruction(payload):
    """只补充HTTP回执外壳说明，不改变任务和缓存身份中的原始提示词。"""
    if not isinstance(payload,dict):return ""
    if payload.get('task')=='unified_classify':
        if payload.get('response_contract')=='unified_structure_v1':
            shared=transport_instruction({**payload,'response_contract':'unified_dimension_refs_v1'})
            shared=shared.replace('返回cells与dimension_values两个根字段','返回cells、dimension_values、regions三个根字段')
            return shared+('本次用regions明确表达已理解的共同语义：每项恰含range、record、rows、columns。'
                'range是仅含候选格的矩形范围，不得覆盖包外格、合并空位或与其他region重叠。'
                'record是该区共同的完整记录外壳，业务区必须明确kind、metric_id、dimension_refs、reason；'
                'partial_metric_fact还需missing_dimensions与evidence_cells。指标不同请分区，不在行列里覆盖指标。'
                'rows每项含indices行号整数数组及record；columns每项含indices列字母数组及record。'
                '行列record只补充dimension_refs及必要的missing_dimensions，与区共同记录的对应数组连接，'
                '同一格同一维度不得在区、行、列中重复声明。行列证据须适用于展开后的每格。'
                'cells只填写单格例外或尚未被区覆盖的候选格，例外记录整体替换该格的区域记录。'
                'regions加cells须覆盖全部候选格；不能压成共同语义的格可以全部逐格写在cells，regions留空。'
                '程序只是展开你的明确判断，不根据数值、位置或邻格默认生成语义。')
        if payload.get('response_contract')=='unified_dimension_refs_v1':
            return ('本次仅改变回执传输格式：返回cells与dimension_values两个根字段。'
                    'dimension_values是共享数组，每项明确包含dimension_id、value、evidence；仅当这三项完全相同才可复用。'
                    '每个metric_fact或partial_metric_fact用dimension_refs明确列出自己的数组下标（从0开始），'
                    '不再输出dimensions或dimension_evidence，不允许默认继承、省略必要维度、覆盖或同一维度多次引用。'
                    '程序只按引用还原完整逐格语义，不替你推断。所有原有指标定义、维度类型、证据及业务要求均适用于还原后的每格。'
                    'partial_metric_fact仍须填写missing_dimensions、evidence_cells和reason；non_business、unresolved格式不变。'
                    '各集合维度的value仍须完整包含domain、mode、completeness及对应范围字段。'
                    '不输出财务数值，不增减候选格，原文不能确定的语义仍保留未决。')
        return ('按本次统一指标定义返回完整JSON对象。各实际维度使用其声明的值类型，集合维度必填domain、mode、completeness及对应成员或范围字段。'
                '不输出财务数值，不增减候选地址；已知指标但缺实际维度用partial_metric_fact保存已知部分，指标尚不明确用unresolved。格式合法仍须通过独立业务校验。')
    if payload.get("task") in {"source_scope","source_scope_batch"} and _fixed_scope_mode(payload):
        wrapper='单包返回 {"cells":{地址:判断}}' if payload["task"]=="source_scope" else '批量返回 {"packets":{packet_id:{"cells":{地址:判断}}}}'
        return (wrapper+'。cells的键必须恰好是该包candidate_cells，不生成range，不返回areas；不得遗漏、增加或重复格。'
            '每格单独填写decision、reason、evidence_cells，可按原规则填写semantic_comparison。'
            '同一行左右可属不同业务，不能共用行结论；只判断本格业务是否属于目标范围。'
            '主表、附表、统计表等载体不同不决定业务不同；所有原业务证据和双轮校验规则继续有效。')
    if payload.get("task")=="layout_cells":
        return ('本次布局回执只填写 {"cells":{原地址:单格记录}}，cells的键必须恰好是candidate_cells，每格一次，不遗漏、不增加。'
            '不生成矩形或range，不输出金额、tables或其他根字段。每格kind只能为business、non_business、unresolved。'
            'business必填kind、reason、note_id、scope、header_cells、table_semantic；note_id限完整notes目录中的一个科目。'
            'non_business必填kind、reason、category、evidence_cells；unresolved必填kind、reason、evidence_cells。'
            'reason必须说明原格判断依据；header_cells和evidence_cells至少引用一个payload.cells或context.preceding_cells中的真实原地址。'
            '只填对应kind允许字段；相邻格可属不同科目，未知归属保留unresolved，不以格式完整替代原表证据、实际口径或财务判断。')
    if payload.get("task")=="layout":
        return ('本次布局HTTP回执必须提供 tables、non_business、unresolved 三个数组，均可为空但不能省略；'
            'candidate_cells 中每格必须恰好由一种区域覆盖，未知不能丢格。'
            '已确定业务科目放入 tables，每个区段的 note_ids 必须恰好一个科目；有原格证据的非业务区域放入 non_business；'
            '确属业务但科目、口径或必要上下文未确定的区域放入 unresolved，不能为覆盖候选格编造科目或排为非业务。'
            '每个 unresolved 区域只含 range、reason、evidence_cells 三个必填字段；'
            'range 必须是原表真实坐标范围，reason 必须说明待核实原因，evidence_cells 至少引用一个已提供的真实原格。'
            '没有某类区域时返回该类空数组；保留原有表头、证据和财务语义规则。')
    if payload.get("task")!="classify":return ""
    return ('本次HTTP回执只使用 {"cells":{地址:单格记录}} 对象，不输出 mappings、excluded、unresolved 数组。'
        'cells 的键必须且只能是 candidate_cells 的全部地址，每个地址恰好一次。'
        '每个单格记录都必须有 kind 和非空 reason；kind 只能是 mapping、excluded、unresolved 中的一种。'
        'mapping 必填 kind、reason、slot_id、scope、dimensions、semantic_field、value_type、row_label、column_label；'
        'scope 和 dimensions 必须依据当前原表及金标准明确填写，不能省略或由程序猜测。'
        'excluded 必填 kind、reason、category，并保留原业务规则要求的 evidence_cells；'
        'unresolved 必填 kind、reason，说明缺少何种证据。只输出对应 kind 允许的字段。'
        '地址由 cells 的键表示，记录内不再输出 cell 字段。所有原有业务识别、证据和维度要求继续有效。')


def _normalize_unified_refs(payload,result):
    """逐格展开模型明确给出的引用；没有默认值或隐含继承。"""
    if not isinstance(result,dict):raise ValueError('统一识别回执必须是对象')
    cells=result.get('cells')
    owned=_layout_owned(payload)
    if not isinstance(cells,dict) or set(cells)-set(owned):
        raise ValueError('统一识别回执不是单格对象或含包外格')
    # 漏格仅放空记录，交既有逐格校验标记失败；不补含义、证据或数值。
    # 原始回执由传输层另存，其他已返回格仍须经过双轮业务校验。
    def complete_addresses(records):
        return {'cells':{address:copy.deepcopy(records.get(address,{})) for address in owned}}
    if set(result)=={'cells'}:
        if any(isinstance(row,dict) and 'dimension_refs' in row for row in cells.values()):
            raise ValueError('维度引用缺少共享数组')
        # 完整旧格式继续交给既有逐格业务校验器。
        return complete_addresses(cells)
    if set(result)!={'cells','dimension_values'} or not isinstance(result['dimension_values'],list):
        raise ValueError('共享维度回执根字段无效')
    definitions={row['id'] for row in payload['dimensions']}
    evidence={address for address,item in payload['cells'].items() if item.get('value') is not None and str(item['value']).strip()}
    evidence.update(payload.get('context_evidence',{}))
    pool=result['dimension_values']
    for entry in pool:
        if not isinstance(entry,dict) or set(entry)!={'dimension_id','value','evidence'}:
            raise ValueError('共享维度记录字段不完整或含多余字段')
        identifier=entry['dimension_id'];refs=entry['evidence']
        if not isinstance(identifier,str) or identifier not in definitions:
            raise ValueError('共享维度引用未知定义')
        if (not isinstance(refs,list) or not refs or any(not isinstance(ref,str) or ref not in evidence for ref in refs)
                or len(set(refs))!=len(refs)):
            raise ValueError('共享维度证据无效')
    expanded=copy.deepcopy(cells)
    for address,row in expanded.items():
        if not isinstance(row,dict):raise ValueError(address+'：单格回执必须是对象')
        if row.get('kind') not in {'metric_fact','partial_metric_fact'}:
            if 'dimension_refs' in row:raise ValueError(address+'：非指标格不能引用维度')
            continue
        if 'dimensions' in row or 'dimension_evidence' in row:
            raise ValueError(address+'：共享引用不能与完整维度混用')
        refs=row.pop('dimension_refs',None)
        if isinstance(refs,dict):
            # 名称和数组项须完全一致；只转换明确引用，不猜测缺失维度。
            for identifier,ref in refs.items():
                if type(ref) is not int or not 0<=ref<len(pool):raise ValueError(address+'：维度下标无效')
                if identifier!=pool[ref]['dimension_id']:raise ValueError(address+'：维度名称与引用不一致')
            refs=list(refs.values())
        if not isinstance(refs,list):raise ValueError(address+'：缺少逐格维度引用')
        dims={};proof={}
        for ref in refs:
            if type(ref) is not int or not 0<=ref<len(pool):raise ValueError(address+'：维度下标无效')
            entry=pool[ref];identifier=entry['dimension_id']
            if identifier in dims:raise ValueError(address+'：同一维度重复引用')
            dims[identifier]=copy.deepcopy(entry['value']);proof[identifier]=copy.deepcopy(entry['evidence'])
        row.update(dimensions=dims,dimension_evidence=proof)
    return complete_addresses(expanded)


def normalize_response(payload,result):
    """将单格对象回执转成原业务协议；不合并状态，不推测数据。"""
    if not isinstance(payload,dict):return result
    if payload.get('task')=='unified_classify' and payload.get('response_contract')=='unified_structure_v1':
        from .结构语义 import expand_structure
        return _normalize_unified_refs(payload,expand_structure(payload,result))
    if payload.get('task')=='unified_classify' and payload.get('response_contract')=='unified_dimension_refs_v1':
        return _normalize_unified_refs(payload,result)
    if payload.get("task")=="layout_cells":return _normalize_layout_cells(payload,result)
    if payload.get("task") in {"source_scope","source_scope_batch"} and _fixed_scope_mode(payload):
        if payload["task"]=="source_scope":return _normalize_scope_cells(payload,result)
        tasks=payload.get("tasks",[]);identifiers=[task.get("packet_id") for task in tasks]
        if any(task.get("scope_response_mode")!="cells_v1" for task in tasks):
            raise ValueError("固定格范围任务不能与旧范围任务混用")
        if not isinstance(result,dict) or set(result)!={"packets"} or not isinstance(result["packets"],dict):
            raise ValueError("固定格批量范围回执必须是packets对象")
        if len(set(identifiers))!=len(identifiers) or set(result["packets"])!=set(identifiers):
            raise ValueError("固定格批量范围回执漏包、包外或任务编号重复")
        return {"results":[{"packet_id":task["packet_id"],
                **_normalize_scope_cells(task,result["packets"][task["packet_id"]])} for task in tasks]}
    if payload.get("task")!="classify":return result
    if not isinstance(result,dict):raise ValueError("分类回执必须是对象")
    collections=("mappings","excluded","unresolved")
    if "cells" not in result:
        # 旧成功缓存和旧三清单回执仍交原校验器，重复记录不得在此去重或吞掉。
        if not all(isinstance(result.get(name),list) for name in collections):
            raise ValueError("分类回执缺少 cells 对象或完整三种分类清单")
        return copy.deepcopy(result)
    if set(result)!={"cells"}:raise ValueError("cells 回执不能与三种分类清单或其他根字段混用")
    cells=result["cells"]
    if not isinstance(cells,dict):raise ValueError("分类回执 cells 必须是对象")
    owned=_strings(payload.get("candidate_cells",[]))
    missing=set(owned)-set(cells);outside=set(cells)-set(owned)
    if missing or outside:
        raise ValueError("分类回执候选格不完整或含包外格：缺少="+"、".join(sorted(missing))+"；包外="+"、".join(sorted(map(str,outside))))
    branches={branch["properties"]["kind"]["enum"][0]:branch for branch in _classify_record(payload)["oneOf"]}
    destinations={"mapping":"mappings","excluded":"excluded","unresolved":"unresolved"}
    normalized={name:[] for name in collections}
    for address in owned:
        record=cells[address]
        if not isinstance(record,dict):
            raise ValueError(address+"：单格回执无效或包含多余字段")
        kind=record.get("kind");reason=record.get("reason")
        if not isinstance(kind,str) or kind not in destinations:
            raise ValueError(address+"：单格回执 kind 无效")
        branch=branches[kind]
        if set(record)-set(branch["properties"]):
            raise ValueError(address+"：单格回执包含不属于"+kind+"的多余字段")
        missing=[name for name in branch["required"] if name not in record]
        if missing:raise ValueError(address+"："+kind+"单格回执缺少必填字段："+ "、".join(missing))
        if not isinstance(reason,str) or not reason.strip():raise ValueError(address+"：单格回执缺少依据")
        item=copy.deepcopy({key:value for key,value in record.items() if key!="kind"})
        item["cell"]=address
        normalized[destinations[kind]].append(item)
    return normalized


def response_format(payload):
    """只返回HTTP格式参数；输入对象和现有成功任务的身份保持不变。"""
    task=payload.get("task") if isinstance(payload,dict) else None
    if task in {"source_scope","source_scope_batch"} and _fixed_scope_mode(payload):
        schema=_fixed_scope_schema(payload)
    elif task=="source_scope":
        schema=_object({"areas":_areas(_evidence(payload),refinement="refinement" in payload)})
    elif task=="source_scope_batch":
        tasks=payload.get("tasks",[])
        identifiers=[item.get("packet_id") for item in tasks if isinstance(item,dict)]
        evidence=[address for item in tasks if isinstance(item,dict) for address in _evidence(item)]
        # 使用本批文字地址并集，避免为每个包展开大型分支；原校验器仍逐包拒绝跨包证据。
        records=_object({"packet_id":_enum(identifiers),"areas":_areas(evidence,
            refinement=any(isinstance(item,dict) and "refinement" in item for item in tasks))})
        schema=_object({"results":_array(records,count=len(tasks))})
    elif task=="tables":
        identifiers=[item.get("id") for item in payload.get("catalogue",[]) if isinstance(item,dict)]
        schema=_object({"table_ids":_ids(identifiers)})
    elif task=="missing_semantics_proposal_review":
        owned=_strings(payload.get("candidate_cells",[]))
        item=_object({"cell":_enum(owned),"accepted":{"type":"boolean"},"reason":_text(),
                      "evidence_cells":_ids(_evidence(payload))})
        schema=_object({"decisions":_array(item,count=len(owned))})
    elif task=="missing_semantics_existing_check":
        identifiers=_strings(item.get("id") for item in payload.get("existing_slots",[]) if isinstance(item,dict))
        schema=_object({"checked_ids":_ids(identifiers,count=len(identifiers)),"existing_ids":_ids(identifiers),
                        "open_dimension_ids":_ids(identifiers),"uncertain":{"type":"boolean"},"reason":_text()})
    elif task=="layout_cells":
        schema=_layout_cells(payload)
    elif task=="layout":
        schema=_layout(payload)
    elif task=="classify":
        schema=_classify(payload)
    elif task=="unified_classify":
        schema=_unified_classify(payload)
    else:
        # 完整新定义、结构操作和披露限制等复杂内容继续由现有校验器解释。
        return {"type":"json_object"}
    return {"type":"json_schema","json_schema":{"name":"reply_"+task,"strict":True,"schema":schema}}
