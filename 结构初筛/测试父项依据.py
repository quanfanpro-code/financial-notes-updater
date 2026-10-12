"""只有固定定义明确声明且本表唯一落地的父项可以承接。"""
import unittest
import 建立结构索引 as 库


class 父项依据(unittest.TestCase):
    def test_明确父项与不完整证据分开(self):
        cells={a:{'row':库.range_boundaries(a)[1],'column':库.range_boundaries(a)[0],
                  'value':v,'formula':None,'cached_value':v}
               for a,v in {'A1':'项目','B1':'期末','C1':'期初','A2':'合计','A3':'其中：境外款项'}.items()}
        maps=[{'mapping_id':a,'cell':a,'target_slot_id':a,'row_semantics':cells['A'+a[1:]]['value'],
               'column_semantics':'期末' if a[0]=='B' else '期初','__文件':'虚构'} for a in ('B2','C2','B3','C3')]
        reps={a:{'representation_id':a,'semantic_id':a,'规范维度值':{},'确认依据':{'说明':'虚构'},
                 '父项语义ID':a[0]+'2' if a[1]=='3' else None} for a in ('B2','C2','B3','C3')}
        def extract():
            gaps=[]
            table=库._提取小表(1,'虚构.xlsx','表','表',maps,(1,1,3,3),'固定范围',
                {'cells':cells,'merges':[]},reps,'A',gaps,[])
            return table,gaps
        table,gaps=extract()
        child=next(p for p in table['项目顺序'] if p['坐标']=='A3')
        self.assertEqual(child['上级'],'A2')
        self.assertFalse(any(g.get('坐标')=='A3' for g in gaps))
        reps['C3']['父项语义ID']=None
        table,gaps=extract()
        self.assertIsNone(next(p for p in table['项目顺序'] if p['坐标']=='A3')['上级'])
        self.assertTrue(any(g.get('坐标')=='A3' for g in gaps))

if __name__=='__main__':unittest.main(verbosity=2)
