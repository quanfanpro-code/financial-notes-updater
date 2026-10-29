"""同时检查完整响应、流式响应和不完整返回；不请求外部服务。"""
import io
import json
import unittest
from 附注更新.语义 import _read_sse_stream


def 响应(内容, 类型):
    r = io.BytesIO(内容)
    r.headers = {'Content-Type': 类型}
    return r


class 模型传输收尾(unittest.TestCase):
    def test_完整对象后多余结束符可留痕恢复且不接受第二个内容(self):
        from 附注更新.语义 import _repair_complete_response,SemanticError
        raw='{"tables":[{"id":"任意表","data_ranges":["D8:E9"]}]}'
        for tail in ('}',']}', '}}', ']\n }'):
            with self.subTest(tail=tail):
                result,repairs=_repair_complete_response(raw+tail)
                self.assertEqual(result,{'tables':[{'id':'任意表','data_ranges':['D8:E9']}]})
                self.assertEqual(repairs,[{'kind':'extra_root_object_closer','removed':tail}])
        for tail in ('[', '}{"another":true}', '{"another":true}','改成其他来源', ']false'):
            with self.subTest(tail=tail),self.assertRaises(SemanticError):_repair_complete_response(raw+tail)
        with self.assertRaises(ValueError):_repair_complete_response(raw[:-1])

    def test_完整返回原样读取(self):
        数据 = {'choices': [{'message': {'content': '{"完成":true}'}, 'finish_reason': 'stop'}]}
        原文 = json.dumps(数据, ensure_ascii=False).encode('utf-8')
        结果, 证据 = _read_sse_stream(响应(原文, 'application/json; charset=utf-8'))
        self.assertEqual(结果, 数据)
        self.assertEqual(证据, 原文)

    def test_流式分段拼接(self):
        原文 = b'data: {"choices":[{"delta":{"content":"{\\"ok\\":"}}]}\n\ndata: {"choices":[{"delta":{"content":"true}"},"finish_reason":"stop"}]}\n\ndata: [DONE]\n'
        结果, 证据 = _read_sse_stream(响应(原文, 'text/event-stream'))
        self.assertEqual(结果['choices'][0]['message']['content'], '{"ok":true}')
        self.assertNotIn('transport_error', 结果)
        self.assertEqual(证据, 原文)

    def test_流式缺结束状态不能冒充完整(self):
        原文 = b'data: {"choices":[{"delta":{"content":"{}"}}]}\n'
        结果, _ = _read_sse_stream(响应(原文, 'text/event-stream'))
        self.assertIn('transport_error', 结果)

    def test_坏数据不能被静默跳过(self):
        原文 = b'data: broken\n\ndata: {"choices":[{"delta":{"content":"{}"},"finish_reason":"stop"}]}\n\ndata: [DONE]\n'
        结果, 证据 = _read_sse_stream(响应(原文, 'text/event-stream'))
        self.assertIn('transport_error', 结果)
        self.assertEqual(证据, 原文)


if __name__ == '__main__':
    unittest.main()
