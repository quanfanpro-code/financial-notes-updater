"""无缓存的中途表头不能当作数值公式放过。"""
import unittest
import 建立结构索引 as 库


class 公式依赖(unittest.TestCase):
    def test_公式已缓存空字符串仍保留原式但不报缺缓存(self):
        cells={a:{'row':库.range_boundaries(a)[1],'column':库.range_boundaries(a)[0],
                  'value':v,'formula':None,'cached_value':v}
               for a,v in {'A1':'项目','B1':'本期','A2':'甲'}.items()}
        cells['B3']={'row':3,'column':2,'value':'=IFERROR(B2/0,"")',
                     'formula':'=IFERROR(B2/0,"")','cached_value':''}
        maps=[{'mapping_id':'B2','cell':'B2','target_slot_id':'B2','row_semantics':'甲',
               'column_semantics':'本期','__文件':'虚构'}]
        reps={'B2':{'representation_id':'B2','semantic_id':'B2','规范维度值':{},'确认依据':{'说明':'虚构'}}}
        gaps=[];table=库._提取小表(1,'虚构.xlsx','表','表',maps,(1,1,2,3),'明确范围',
                   {'cells':cells,'merges':[]},reps,'A',gaps,[])
        node=next((n for n in table['节点'] if n['坐标']=='B3'),None)
        self.assertIsNotNone(node,'空字符串结果不能使原公式从内容中消失')
        self.assertEqual(node['公式'],cells['B3']['formula'])
        self.assertEqual(node['缓存'],'')
        self.assertNotEqual(node['角色'],'业务数值')
        self.assertFalse(any(g.get('坐标')=='B3' and g['问题']=='公式结构格缺少缓存显示' for g in gaps))

    def test_业务区文本公式仍保留对后续格的影响(self):
        cells={a:{'row':库.range_boundaries(a)[1],'column':库.range_boundaries(a)[0],
                  'value':v,'formula':None,'cached_value':v}
               for a,v in {'A1':'项目','B1':'本期','A2':'甲','A4':'乙'}.items()}
        cells['B3']={'row':3,'column':2,'value':'="上期"','formula':'="上期"','cached_value':None}
        maps=[{'mapping_id':a,'cell':a,'target_slot_id':a,'row_semantics':n,
               'column_semantics':p,'__文件':'虚构'} for a,n,p in [('B2','甲','本期'),('B4','乙','上期')]]
        reps={a:{'representation_id':a,'semantic_id':a,'规范维度值':{},'确认依据':{'说明':'虚构'}} for a in ('B2','B4')}
        gaps=[]
        库._提取小表(1,'虚构.xlsx','表','中途期间',maps,(1,1,2,4),'明确范围',
            {'cells':cells,'merges':[]},reps,'A',gaps,[])
        gap=next(g for g in gaps if g.get('坐标')=='B3')
        self.assertEqual(gap['影响业务格'],['B4'])


if __name__=='__main__':unittest.main(verbosity=2)
