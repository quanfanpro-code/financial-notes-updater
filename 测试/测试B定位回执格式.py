# coding: utf-8
"""只限定定位回执能引用的已展示文字，不替模型选择业务来源。"""
from copy import deepcopy
import unittest
from jsonschema import Draft202012Validator,ValidationError
from 附注更新.回执格式 import response_format
from 附注更新.表格编排 import check_evidence


class B定位回执格式测试(unittest.TestCase):
    def test_核查只引用当前材料文字而不是数字格或其他表内容(self):
        payload={'task':'a_semantic_review','standards':{'metrics':[],'dimensions':[]},
            'reading':[{'id':'T1'}], 'source':{'context':{'@C0':{}}},
            'text_evidence':{'T1':{'当前表!A1':'实际标题','当前表!A3':'类别甲'}}}
        validator=Draft202012Validator(response_format(payload)['json_schema']['schema'])
        corrections={'metrics':[],'patches':[],'resolved_ranges':[],'non_business':[],'remove_dimensions':[]}
        answer={'tables':[{'id':'T1','corrections':corrections,'issues':[
            {'range':'B3','reason':'原文未说明主体','evidence':['当前表!A1']}]}]}
        validator.validate(answer)
        for refs in (['当前表!B3'],['其他表!A1'],['其他表的主体名称'],[]):
            wrong=deepcopy(answer);wrong['tables'][0]['issues'][0]['evidence']=refs
            with self.assertRaises(ValidationError):validator.validate(wrong)
        from 附注更新.A语义阅读 import normalize_evidence
        with self.assertRaises(ValueError):
            normalize_evidence(['其他表的主体名称'],'当前表',{},'B3')

    def test_共同阅读限制几何范围不替模型判定数据角色(self):
        payload={'task':'a_semantic_read',
            'reading_ranges':[{'sheet':'另一工作表','range':'A2:C4','context_ranges':['A1:C1']}],
            'candidate_locations':['另一工作表!A3','另一工作表!C3']}
        before=deepcopy(payload);form=response_format(payload)
        validator=Draft202012Validator(form['json_schema']['schema'])
        Draft202012Validator.check_schema(form['json_schema']['schema'])
        table={'id':'分类','sheet':'另一工作表','range':'A3:C3','subject':'原文业务','meaning':'行名与金额',
               'data_ranges':['C3'],'context_ranges':['A1:C1']}
        validator.validate({'tables':[table],'notes':''})
        validator.validate({'tables':[{**table,'data_ranges':['A3','C3']}],'notes':''})
        for change in ({'range':'A1:C3'},{'range':'A2:D4'},{'sheet':'未知表'},
                       {'data_ranges':['D3']},{'data_ranges':['B3']},{'context_ranges':['A9']}):
            with self.assertRaises(ValidationError):
                validator.validate({'tables':[{**table,**change}],'notes':''})
        self.assertEqual(payload,before)
        self.assertEqual(response_format({'task':'a_semantic_read'}),{'type':'json_object'})

    def test_千个以上候选保留原生结构上限且不漏地址(self):
        payload={'task':'a_semantic_read','reading_ranges':[{'sheet':'大表','range':'A1:C1200','context_ranges':[]}],
                 'candidate_locations':[f'大表!C{r}' for r in range(1,1201)]}
        schema=response_format(payload)['json_schema']['schema']
        data=schema['properties']['tables']['items']['anyOf'][0]['properties']['data_ranges']
        self.assertNotIn('enum',data['items'])
        validator=Draft202012Validator(schema)
        table={'id':'1','sheet':'大表','range':'A1:C1200','subject':'分类','meaning':'实际含义',
               'data_ranges':['C1','C1200'],'context_ranges':[]}
        validator.validate({'tables':[table],'notes':''})
        with self.assertRaises(ValidationError):
            validator.validate({'tables':[{**table,'data_ranges':['C1201']}],'notes':''})

    def test_同表范围共用结构但不同工作表不混用(self):
        payload={'task':'a_semantic_read','reading_ranges':[
            {'sheet':'甲','range':'A2:C4','context_ranges':['A1:C1']},
            {'sheet':'甲','range':'A10:C11','context_ranges':['A1:C1']},
            {'sheet':'乙','range':'G20:H22','context_ranges':[]}],
            'candidate_locations':['甲!C3','甲!C10','乙!H21']}
        schema=response_format(payload)['json_schema']['schema']
        self.assertEqual(len(schema['properties']['tables']['items']['anyOf']),2)
        validator=Draft202012Validator(schema)
        table={'id':'1','sheet':'甲','range':'A10:C11','subject':'原文','meaning':'实际含义',
               'data_ranges':['C10'],'context_ranges':['A1:C1']}
        validator.validate({'tables':[table],'notes':''})
        for change in ({'range':'A4:C10'},{'range':'G20:H22'},{'data_ranges':['H21']},{'range':'A1:C11'}):
            with self.assertRaises(ValidationError):
                validator.validate({'tables':[{**table,**change}],'notes':''})

    def setUp(self):
        self.payload={'task':'b_locate','requirements':[{'id':'G1'}],
            'sheets':[{'sheet':'原表','preview_text':[{'row':1,'cells':[['A1','总标题']]}]}],
            'local_text':[{'sheet':'原表','range':'A4:C8','text':[{'row':4,'cells':[['A4','合并标题'],['B4',12]]}]}],
            'previous_result':{'areas':[{'evidence':['原表!A3']}]}}
        self.answer={'areas':[{'id':'1','subject':'原文科目','purpose':'候选用途',
            'ranges':[{'sheet':'原表','range':'A4:C8'}],'evidence':['原表!A4'],
            'group_ids':['G1'],'reason':'已读原文'}],'unresolved_groups':[],'read_ranges':[]}

    def validator(self):
        form=response_format(self.payload);self.assertEqual(form['type'],'json_schema')
        schema=form['json_schema']['schema'];Draft202012Validator.check_schema(schema)
        return Draft202012Validator(schema)

    def test_仅允许实际已展示文字及既有核实引用(self):
        validator=self.validator();validator.validate(self.answer)
        for ref in ['原表!A1','原表!A3']:
            answer=deepcopy(self.answer);answer['areas'][0]['evidence']=[ref];validator.validate(answer)
        for ref in ['原表!A5','原表!B4','原表!Z99']:
            answer=deepcopy(self.answer);answer['areas'][0]['evidence']=[ref]
            with self.assertRaises(ValidationError):validator.validate(answer)
        answer=deepcopy(self.answer);answer['areas'][0]['evidence']=[]
        with self.assertRaises(ValidationError):validator.validate(answer)

    def test_拒绝未知需求组和工作表且不修改请求(self):
        before=deepcopy(self.payload);validator=self.validator()
        for field,value in [('group_ids',['G99']),('ranges',[{'sheet':'未给的表','range':'A1:B2'}])]:
            answer=deepcopy(self.answer);answer['areas'][0][field]=value
            with self.assertRaises(ValidationError):validator.validate(answer)
        self.assertEqual(self.payload,before)

    def test_原件校验继续拒绝合并延续空格(self):
        snapshot={'sheets':[{'name':'原表','cells':{'A4':{'value':'合并标题'},'A5':{'value':None}},'merges':['A4:A5']}]}
        check_evidence(['原表!A4'],snapshot)
        with self.assertRaisesRegex(ValueError,'不是已有文字'):check_evidence(['原表!A5'],snapshot)

    def test_补查结构完整且原校验器继续拒绝漏项与重复目标(self):
        from 附注更新.表格编排 import validate_candidate_search
        payload={**deepcopy(self.payload),'task':'b_candidate_followup',
            'targets':[{'id':'目标甲'},{'id':'目标乙'}],
            'shown_ranges':[{'sheet':'原表','range':'A4:B4'}],
            'previous_result':{'evidence':['原表!A3']}}
        form=response_format(payload);self.assertEqual(form['type'],'json_schema')
        validator=Draft202012Validator(form['json_schema']['schema'])
        answer={'read_ranges':[],'ranges':[{'sheet':'原表','range':'A4:B4'}],
            'structure':'原文查找结论','evidence':['原表!A4'],
            'targets':[{'id':'目标甲','candidates':['原表!B4'],'reason':'原文项目'},
                       {'id':'目标乙','candidates':[],'reason':'未列所需明细'}]}
        validator.validate(answer)
        snapshot={'sheets':[{'name':'原表','max_row':4,'max_column':2,'cells':{
            'A4':{'row':4,'column':1,'value':'项目'},'B4':{'row':4,'column':2,'value':0}}}]}
        self.assertEqual(validate_candidate_search(answer,payload,snapshot,{'metric':{}}),answer)
        for change in ({'targets':answer['targets'][:1]},{'targets':[{**answer['targets'][0],'id':'未知'},answer['targets'][1]]},
                       {'evidence':['原表!A5']}):
            with self.assertRaises(ValidationError):validator.validate({**answer,**change})
        repeated={**answer,'targets':[answer['targets'][0],answer['targets'][0]]}
        with self.assertRaisesRegex(ValueError,'未逐项覆盖'):validate_candidate_search(repeated,payload,snapshot,{'metric':{}})


if __name__=='__main__':unittest.main()
