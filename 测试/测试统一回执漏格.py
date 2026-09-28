"""漏格保持未决，不能连带丢弃其他格，也不能补出业务事实。"""
import copy
import unittest
from 附注更新.回执格式 import normalize_response
from 附注更新.统一识别 import build_payload, validate_response, agree_rounds
from 附注更新.统一语义 import validate_records
from 测试.测试统一识别 import packet, answer
from 测试.测试统一语义 import records


class 漏格隔离测试(unittest.TestCase):
    def test_漏格两轮不能形成事实其他格保留(self):
        source=packet();standard=validate_records(records())
        payload=build_payload(source,standard);payload['response_contract']='unified_dimension_refs_v1'
        full=answer();missing=copy.deepcopy(full);address=next(iter(missing['cells']))
        missing['cells'].pop(address);before=copy.deepcopy(missing)
        normalized=normalize_response(payload,missing)
        self.assertEqual(missing,before)
        self.assertEqual(normalized['cells'][address],{})
        with self.assertRaises(ValueError):validate_response(normalized,source,standard)
        checked=validate_response(normalized,source,standard,isolate_errors=True)
        complete=validate_response(full,source,standard,isolate_errors=True)
        self.assertEqual(checked[address]['kind'],'invalid_response')
        for key in missing['cells']:self.assertEqual(checked[key],complete[key])
        merged=agree_rounds(checked,complete,standard)
        self.assertIn(address,{r['cell'] for r in merged['unresolved']})
        self.assertNotIn(source['sheet']+'!'+address,{r['source_reference']['location'] for r in merged['facts']})
        self.assertEqual(len(merged['facts']),len(full['cells'])-1)
        foreign=copy.deepcopy(missing);foreign['cells']['Z999']={}
        with self.assertRaises(ValueError):normalize_response(payload,foreign)
        empty=normalize_response(payload,{'cells':{},'dimension_values':[]})
        self.assertTrue(all(r['kind']=='invalid_response' for r in validate_response(empty,source,standard,isolate_errors=True).values()))


if __name__=='__main__':unittest.main()
