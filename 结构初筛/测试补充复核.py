"""真实核验暴露的多落点遗漏和父项层级冲突。"""
import unittest
import hashlib
import json
from copy import deepcopy
from 比较新表 import 比较工作簿
from 测试独立复核比较 import 参照, 工作簿, 索引
import 建立结构索引 as 库

class 补充复核(unittest.TestCase):
    def test_排除纠错保留小表身份但退出业务节点(self):
        vals={'A1':'项目','B1':'准备','C1':'核对','A2':'原材料','A3':'合计'}
        cells={a:{'row':库.range_boundaries(a)[1],'column':库.range_boundaries(a)[0],
                  'value':v,'formula':None,'cached_value':v} for a,v in vals.items()}
        cells['C2']={'row':2,'column':3,'value':'=B2-B3','formula':'=B2-B3','cached_value':0}
        maps=[{'mapping_id':a,'cell':a,'worksheet':'表','target_slot_id':a,
               'row_semantics':vals['A'+a[1:]],'column_semantics':vals[a[0]+'1'],'__文件':'虚构'}
              for a in ('B2','C2','B3')]
        reps={a:{'representation_id':a,'semantic_id':a,'规范维度值':{},'确认依据':{'说明':'虚构'}}
              for a in ('B2','C2','B3')}
        def table(records):
            return 库._提取小表(1,'虚构.xlsx','表','表',records,(1,1,3,3),'固定范围',
                                {'cells':cells,'merges':[]},reps,'A',[],[])
        original=table(maps);new=deepcopy(maps);old=deepcopy(new[1])
        digest=hashlib.sha256(json.dumps(old,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest()
        new[1].update(target_slot_id=None,status='结构格排除',confirmation={'纠错ID':'差值'},
            correction_history={'原映射完整记录':old,'原记录摘要':digest})
        corrected=table(new)
        self.assertEqual(corrected['小表编号'],original['小表编号'])
        self.assertNotIn('C2',{n['坐标'] for n in corrected['节点'] if n['角色']=='业务数值'})
        self.assertEqual(next(n for n in corrected['节点'] if n['坐标']=='C2')['公式'],'=B2-B3')
        for field in ('cell','mapping_id'):
            damaged=deepcopy(new);damaged[1]['correction_history']['原映射完整记录'][field]='其他'
            with self.subTest(field=field),self.assertRaises(ValueError):table(damaged)

    def 共用表头样本(self):
        old,data=参照('cash','货币资金',['库存现金','银行存款'])
        old['节点']=[n for n in old['节点'] if n['坐标'] not in ('C2','D2')]
        for n in old['节点']:
            n['行']+=10;n['坐标']=库.get_column_letter(n['列'])+str(n['行'])
        old['表头']=[{'坐标':p,'查找文字':data[p],'覆盖列':p[0],'共用':True} for p in ('C2','D2')]
        values={库.get_column_letter(库.range_boundaries(p)[0])+str(库.range_boundaries(p)[1]+10):v
                for p,v in data.items() if p not in ('C2','D2')}
        values.update({p:data[p] for p in ('C2','D2')})
        return old,values

    def test_共用表头支持业务列且不混入正文范围(self):
        for mode in ('原样','期间互换','整表平移'):
            with self.subTest(mode=mode):
                old,values=self.共用表头样本()
                if mode=='期间互换':values['C2'],values['D2']=values['D2'],values['C2']
                if mode=='整表平移':
                    values={库.get_column_letter(库.range_boundaries(p)[0]+1)+str(库.range_boundaries(p)[1]+2):v for p,v in values.items()}
                result=比较工作簿(工作簿('货币资金',values),索引([old]))
                actual={(d['旧坐标'],d['新坐标']) for c in result['候选'] for d in c['结构对应'] if d['角色']=='业务数值'}
                target='D13' if mode=='期间互换' else 'D15' if mode=='整表平移' else 'C13'
                self.assertIn(('C13',target),actual)
                self.assertTrue(all(库.range_boundaries(c['新表范围'])[1]>4 for c in result['候选']))

    def test_共用表头无缓存或多列同名不能推定唯一列(self):
        for mode in ('无缓存','多列同名'):
            with self.subTest(mode=mode):
                old,values=self.共用表头样本()
                if mode=='多列同名':values['E2']=values['C2']
                book=工作簿('货币资金',values)
                if mode=='无缓存':book['sheets'][0]['cells']['C2'].update(value='="期末余额"',formula='="期末余额"',cached_value=None)
                result=比较工作簿(book,索引([old]))
                points={d['旧坐标'] for c in result['候选'] for d in c['结构对应'] if d['角色']=='业务数值'}
                self.assertNotIn('C13',points)
                self.assertIn('D13',points)

    def test_不同工作表的同一旧格也必须标明歧义(self):
        old,data=参照('cash','货币资金',['库存现金','银行存款'])
        book=工作簿('货币资金',data)
        book['sheets'].append({**deepcopy(book['sheets'][0]),'name':'货币资金副表'})
        result=比较工作簿(book,索引([old]))
        self.assertEqual(len(result['候选']),2)
        self.assertTrue(all(any(any(d['旧坐标']=='C3' for d in a.get('不同对应',[]))
                               for a in c['歧义']) for c in result['候选']))

    def test_总表祖先不能冒充原表中的直接父项(self):
        vals={'A1':'项目','B1':'期末','A2':'债务工具投资','A3':'其中：债券','A4':'合计'}
        cells={a:{'row':库.range_boundaries(a)[1],'column':库.range_boundaries(a)[0],
                  'value':v,'formula':None,'cached_value':v} for a,v in vals.items()}
        maps=[{'mapping_id':a,'cell':a,'target_slot_id':a,'row_semantics':vals['A'+a[1:]],
               'column_semantics':'期末','__文件':'虚构'} for a in ('B2','B3','B4')]
        reps={a:{'representation_id':a,'semantic_id':a,'规范维度值':{},'确认依据':{'说明':'虚构'},
                 '父项语义ID':'B4' if a=='B3' else None} for a in ('B2','B3','B4')}
        gaps=[]
        table=库._提取小表(1,'虚构.xlsx','表','表',maps,(1,1,2,4),'固定范围',
                           {'cells':cells,'merges':[]},reps,'A',gaps,[])
        self.assertIsNone(next(p for p in table['项目顺序'] if p['坐标']=='A3')['上级'])
        self.assertTrue(any(g.get('坐标')=='A3' for g in gaps))

if __name__=='__main__':unittest.main(verbosity=2)
