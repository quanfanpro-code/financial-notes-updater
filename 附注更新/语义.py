"""按金标准理解新工作簿的语义；仅返回位置和含义，金额始终由调用方重读来源。

模型调用仅依赖标准库；业务范围补格复用本程序表格模块。人工样式映射不参与识别。HTTP 协议依据：
https://developers.openai.com/api/reference/resources/chat
"""
from __future__ import annotations

import collections
import base64
import mimetypes
import copy
import datetime
import hashlib
import json
import math
import pathlib
import re
import shutil
import threading
import urllib.error
import urllib.parse
import urllib.request
import uuid
from .设置 import DEFAULT_TIMEOUT, validate_thinking_mode
from .金标准 import allowed_scopes, validate_mapping_scope, normalize_mapping_metric

PROMPT_VERSION = "2026-09-19.6"
VALUE_TYPES = {"monetary", "number", "percentage", "text", "date", "boolean", "enum"}
EXCLUDED_TYPES = {"header", "label", "title", "unit", "empty_padding", "annotation"}



class SemanticError(ValueError):
    """输入或模型回执不满足可核验要求。"""


class SemanticProgressError(RuntimeError):
    """已确认进度未能发布，不能当作模型分类失败继续运行。"""


def recognition_progress(snapshot, result):
    """生成覆盖全表的独立进度；未双轮确认及尚未处理的格子均保持未决。"""
    partial = copy.deepcopy(result)
    partial.pop("evidence_path", None)
    partial.pop("evidence_hash", None)
    expected = {(sheet["name"], address) for sheet in snapshot.get("sheets", []) for address in sheet.get("cells", {})}
    seen = set()
    pending = []
    for category in ("mappings", "excluded", "unresolved", "out_of_scope"):
        kept = []
        for item in partial.get(category, []):
            position = (item.get("sheet"), item.get("cell"))
            if position not in expected or position in seen:
                raise SemanticError("进度成果存在重复或包外单元格，不能保存")
            seen.add(position)
            if category in {"mappings", "excluded"} and item.get("reviewed") is not True:
                pending.append({"sheet":position[0], "cell":position[1], "reason":"尚未经过独立双轮确认，等待接续复核", "candidate":item})
            else:
                kept.append(item)
        if category in partial or category != "out_of_scope":partial[category] = kept
    partial["unresolved"].extend(pending)
    partial["unresolved"].extend({"sheet":name,"cell":address,"reason":"尚未识别，等待接续"}
                                 for name,address in sorted(expected-seen,key=lambda p:(p[0],_address(p[1]))))
    observed = [(item["sheet"],item["cell"]) for key in ("mappings","excluded","unresolved","out_of_scope") for item in partial.get(key, [])]
    if len(observed) != len(set(observed)) or set(observed) != expected:
        raise SemanticError("进度成果未完整覆盖实际单元格")
    partial["confirmed_business_ranges"] = {sheet["name"]:copy.deepcopy(sheet.get("confirmed_business_ranges", [])) for sheet in snapshot.get("sheets", [])}
    partial["complete"] = False
    return partial


class SemanticResponseError(SemanticError):
    """有原HTTP证据的格式错误；仅此类HTTP回执错误可进入一次格式纠正。"""
    def __init__(self,evidence_path,previous_response,diagnostic):
        super().__init__("模型回执格式不完整，原始证据已保存；该包仍须严格核验")
        self.evidence_path=str(evidence_path)
        self.previous_response=previous_response
        self.diagnostic=diagnostic




class SameOriginRedirect(urllib.request.HTTPRedirectHandler):
    """只允许同一来源跳转，禁止把 Bearer 凭据转交另一来源或降级传输。"""
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        def origin(url):
            parsed = urllib.parse.urlsplit(url)
            return parsed.scheme.lower(), (parsed.hostname or "").lower(), parsed.port or (443 if parsed.scheme == "https" else 80)
        if origin(req.full_url) != origin(newurl):
            raise SemanticError("模型服务要求跨来源或降级重定向，已停止以保护 API Key")
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _stable_content(value):
    """去掉临时文件位置；真实章节文字、业务坐标、数值及公式缓存仍参与身份。"""
    if isinstance(value, list):
        return [_stable_content(item) for item in value]
    if isinstance(value, dict):
        result = {}
        for key, item in value.items():
            if key == "blank_evidence" and isinstance(item, dict):
                # 临时Excel压缩包hash不属于业务含义，真实来源及全部原格已另行参与缓存身份。
                result[key] = _stable_content({k: v for k, v in item.items() if k != "source_hash"})
                continue
            if key in {"source", "path", "file_path", "source_path", "word", "excel", "linked_word", "original_path"}:
                if isinstance(item, str) and re.match(r"^(?:[A-Za-z]:[\\/]|\\\\)", item):
                    continue
            result[key] = _stable_content(item)
        return result
    return value


class SemanticCancelled(RuntimeError):
    """用户取消识别；已成功包保留以便续作。"""


def _json(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)


def _hash(value):
    return hashlib.sha256(_json(value).encode("utf-8")).hexdigest()


def _unique_json_object(pairs):
    """拒绝重复JSON键，不能让后一个单元格结果静默覆盖前一个。"""
    result={}
    for key,value in pairs:
        if key in result:
            raise SemanticError("模型JSON对象存在重复字段："+str(key)[:100])
        result[key]=value
    return result


def _repair_complete_response(text, *, review_commentary=False):
    """恢复完整JSON外的排版尾文及范围复核中完全相同的重复维度清单。"""
    repairs = []
    def scope_object(pairs):
        result = {}; duplicates = []
        for key, value in pairs:
            if key in result:
                if key != 'dimension_ids' or _json(result[key]) != _json(value):
                    raise SemanticError('模型JSON对象存在冲突或不支持的重复字段：'+str(key)[:100])
                duplicates.append(key)
            result[key] = value
        if duplicates:
            required = {'metric_id', 'can_partition', 'dimension_ids', 'reason'}
            if not required <= set(result) <= required | {'evidence_cells'}:
                raise SemanticError('重复维度清单不属于完整范围复核记录')
            repairs.append({'kind': 'identical_scope_dimension_ids', 'metric_id': result['metric_id'],
                            'duplicate_count': len(duplicates)})
        return result
    decoded, end = json.JSONDecoder(object_pairs_hook=scope_object).raw_decode(text)
    tail = text[end:].strip()
    if tail:
        commentary=tail.removeprefix('```').lstrip()
        if re.fullmatch(r'[\]}\s]+',tail) and isinstance(decoded,dict):
            repairs.append({'kind':'extra_root_object_closer','removed':text[end:]})
        elif (review_commentary and isinstance(decoded,dict) and set(decoded)=={'checked_cells','issues'}
                and commentary and not any(c in commentary for c in ('{','}','```'))
                and not re.search(r'(?m)^\s*\[',commentary)):
            # 只取已完整结束的结构化核查；尾文原样存证，不据尾文补充或改写判断。
            repairs.append({'kind':'trailing_review_commentary','removed':text[end:]})
        elif tail not in {'``', '```'}:
            raise SemanticError('完整JSON后仍有其他内容，不能恢复')
        else:repairs.append({'kind': 'trailing_code_fence', 'removed': text[end:]})
    if not isinstance(decoded, dict) or not repairs:
        raise SemanticError('没有可无歧义恢复的完整回执')
    return decoded, repairs


def _address(value):
    match = re.fullmatch(r"([A-Z]{1,3})([1-9]\d*)", str(value).replace("$", "").upper())
    if not match:
        raise SemanticError("无效单元格地址")
    column = 0
    for char in match[1]:
        column = column * 26 + ord(char) - 64
    row = int(match[2])
    if column > 16384 or row > 1048576:
        raise SemanticError("单元格地址超出工作表范围")
    return row, column


def _range(value):
    ends = str(value).split(":")
    if len(ends) not in (1, 2):
        raise SemanticError("业务表范围无效")
    r1, c1 = _address(ends[0])
    r2, c2 = _address(ends[-1])
    if r1 > r2 or c1 > c2:
        raise SemanticError("业务表范围方向无效")
    return r1, c1, r2, c2


def _inside(address, bounds):
    row, col = _address(address)
    return bounds[0] <= row <= bounds[2] and bounds[1] <= col <= bounds[3]


def _document_context(document, sheet_name, bounds):
    """只裁掉明确属于其他表的Word说明，保留继承段落和未定位的全局事实。"""
    result = []
    for item in document:
        if isinstance(item, dict) and isinstance(item.get("sheet"), str) and all(
            type(item.get(key)) is int and item[key] > 0 for key in ("first_row", "row_count")):
            if item["sheet"] != sheet_name or item["first_row"] > bounds[2] or item["first_row"] + item["row_count"] - 1 < bounds[0]:
                continue
            if all(type(item.get(key)) is int and item[key] > 0 for key in ("first_column", "column_count")):
                if item["first_column"] > bounds[3] or item["first_column"] + item["column_count"] - 1 < bounds[1]:
                    continue
        result.append(item)
    return result


def _source_text_rows(cells):
    """将已展示的原文按实际行列保序传递，不扩大候选或业务范围。"""
    rows=collections.defaultdict(list)
    for address in sorted(cells,key=_address):
        cell=cells[address];value=cell.get("value")
        if isinstance(value,str) and value.strip() and not cell.get("formula"):
            rows[_address(address)[0]].append([address,value])
    return [{"row":row,"cells":items} for row,items in rows.items()]


def _layout_packets(sheet, ordered):
    """先隔离真实Word物理表与表外原格，再限制每包候选数；不推断业务。"""
    physical=[_range(item["range"]) for item in sheet.get("word_table_ranges",[])]
    packet=[];current=None
    for address in ordered:
        owners=[box for box in physical if _inside(address,box)]
        if len(owners)>1:raise SemanticError("Word物理表范围重叠，不能确定原格来源")
        owner=owners[0] if owners else None
        # 包上限从240改为30：包越大模型思考越容易失控膨胀直至断流；小包配合流式与重试更稳。
        if packet and (owner!=current or len(packet)==30):
            yield packet,current
            packet=[]
        current=owner;packet.append(address)
    if packet:yield packet,current


def _normal_dimensions(dimensions, slot=None):
    if not isinstance(dimensions, dict):
        raise SemanticError("维度必须为对象")
    result = copy.deepcopy(dimensions)
    for key, value in result.items():
        if not re.fullmatch(r"[a-z][a-z0-9_]*", key):
            raise SemanticError("维度键须采用金标准英文语义名称")
        if key in {"raw_row_path", "raw_column_path", "row_path", "column_path", "sheet", "cell", "address"}:
            raise SemanticError("行列路径和坐标不能作为语义维度")
        if isinstance(value, str):
            result[key] = value.strip()
    currency = result.get("currency")
    if isinstance(currency, str):
        aliases = {"人民币": "CNY", "RMB": "CNY", "人民币元": "CNY", "美元": "USD", "港币": "HKD", "港元": "HKD", "欧元": "EUR"}
        result["currency"] = aliases.get(currency, currency.upper())
    if "scale" in result:
        try:
            scale = float(result["scale"])
            if not 0 < scale < 1e15:
                raise ValueError
            result["scale"] = int(scale) if scale.is_integer() else scale
        except (TypeError, ValueError, OverflowError):
            raise SemanticError("倍率必须是正数") from None
    age = result.get("age_bucket")
    if age is not None:
        if not isinstance(age,dict) or set(age)!={"lower","upper","lower_inclusive","upper_inclusive"}:
            raise SemanticError("账龄维度须为规范的四边界对象")
        for bound in ("lower","upper"):
            value=age[bound];included=age[bound+"_inclusive"]
            if type(included) is not bool:
                raise SemanticError("账龄边界是否包含须为布尔值")
            if value is None:
                if included:raise SemanticError("无穷账龄边界不能声明包含")
            elif type(value) not in (int,float) or not math.isfinite(value) or value<0:
                raise SemanticError("账龄边界须为非负有限数值或null")
            else:
                age[bound]=int(value) if float(value).is_integer() else value
        low,high=age["lower"],age["upper"]
        if low is None and high is None:
            raise SemanticError("账龄区间不能缺少两端界限")
        if low is not None and high is not None and (low>high or low==high and not(age["lower_inclusive"] and age["upper_inclusive"])):
            raise SemanticError("账龄区间边界倒置或为空")
    period = result.get("period")
    if isinstance(period, dict) and "role" in period:
        declared = next((d for d in (slot or {}).get("dimensions", []) if d.get("name") == "period"), {})
        role = period["role"]
        if not isinstance(role, str) or not role or role != declared.get("role"):
            raise SemanticError("期间角色必须与金标准明确声明一致；不能替代实际日期或自行改变期间含义")
        instant = "date" in period and "start" not in period and "end" not in period
        duration = "start" in period and "end" in period and "date" not in period
        if (declared.get("type") == "instant" and not instant) or (declared.get("type") == "duration" and not duration):
            raise SemanticError("实际期间类型与金标准时点或期间定义不一致")
        # role属于槽位维度的描述；仅去除与定义完全一致的冗余属性，原回执保持不变。
        del period["role"]
    if period is not None:
        def date_ok(text):
            try:
                return isinstance(text, str) and bool(datetime.date.fromisoformat(text))
            except ValueError:
                return False
        if isinstance(period, str):
            if not date_ok(period):
                raise SemanticError("期间必须有可核验的实际日期")
        elif isinstance(period, dict):
            keys = set(period)
            if keys == {"date"} or keys == {"type", "date"} and period["type"] == "instant":
                if not date_ok(period["date"]):
                    raise SemanticError("期间必须有可核验的实际日期")
                result["period"] = period["date"]
            elif keys == {"start", "end"} or keys == {"type", "start", "end"} and period["type"] == "duration":
                if not date_ok(period["start"]) or not date_ok(period["end"]):
                    raise SemanticError("期间必须有可核验的实际日期")
                if datetime.date.fromisoformat(period["start"]) > datetime.date.fromisoformat(period["end"]):
                    raise SemanticError("实际期间的开始日期不能晚于结束日期")
                result["period"] = {"start": period["start"], "end": period["end"]}
            else:
                raise SemanticError("期间对象的类型、日期结构不一致或含未知字段")
        else:
            raise SemanticError("期间必须有可核验的实际日期")
    return result


REQUEST_TOTAL_TIMEOUT = 1800.0  # 单次模型请求总时长上限（秒）；服务端挂而不断时强制断开重试。


def _call_with_deadline(timeout, function, *args):
    """限时执行一次调用；超时抛 TimeoutError，挂起的守护线程随进程退出。"""
    outcome = {}
    def _runner():
        try:
            outcome["value"] = function(*args)
        except BaseException as error:  # 跨线程传回任意异常，包括 SemanticError。
            outcome["error"] = error
    worker = threading.Thread(target=_runner, daemon=True)
    worker.start()
    worker.join(timeout)
    if worker.is_alive():
        raise TimeoutError(f"请求超过 {timeout:.0f} 秒总时长上限")
    if "error" in outcome:
        raise outcome["error"]
    return outcome["value"]


def _read_sse_stream(response):
    """逐块读取流式回执并组装成与非流式一致的结构；两块之间超过套接字超时即断开。

    同时返回原始报文字节，供失败存证使用。
    """
    if 'application/json' in str(getattr(response, 'headers', {}).get('Content-Type', 'application/json')).lower():
        raw = response.read()
        try:
            return json.loads(raw, object_pairs_hook=_unique_json_object), raw
        except (ValueError, UnicodeError) as error:
            return {"transport_error": "完整回执无法解析：" + str(error)}, raw
    content_parts, reasoning_parts, raw_parts = [], [], []
    usage, finish = {}, None
    transport_error = None
    for raw_line in response:
        raw_parts.append(raw_line)
        line = raw_line.decode("utf-8", errors="replace").strip()
        if not line.startswith("data:"):
            continue
        payload = line[5:].strip()
        if payload == "[DONE]":
            break
        try:
            chunk = json.loads(payload, object_pairs_hook=_unique_json_object)
        except (ValueError, TypeError) as error:
            transport_error = "流式回执包含无效数据：" + str(error)
            continue
        if not isinstance(chunk, dict) or chunk.get('error'):
            transport_error = "流式回执包含错误响应"
            continue
        choices = chunk.get("choices") or []
        if choices:
            delta = choices[0].get("delta") or {}
            piece = delta.get("content")
            if isinstance(piece, str):
                content_parts.append(piece)
            thinking = delta.get("reasoning_content")
            if isinstance(thinking, str):
                reasoning_parts.append(thinking)
            if choices[0].get("finish_reason"):
                finish = choices[0]["finish_reason"]
        if chunk.get("usage"):
            usage = chunk["usage"]
    data = {"choices": [{"message": {"content": "".join(content_parts), "reasoning_content": "".join(reasoning_parts)},
                          "finish_reason": finish}],
            "usage": usage}
    if transport_error or finish is None:
        data['transport_error'] = transport_error or '流式回执缺少结束状态，不能认定为完整结果'
    return data, b"".join(raw_parts)


class SemanticEngine:
    def __init__(self, settings: dict, gold_path: str, work_dir: str, log=None, cancel=None):
        self.settings = dict(settings)
        self.log = log or (lambda message: None)
        self.cancel = cancel
        self.work_dir = pathlib.Path(work_dir)
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.gold_path = str(pathlib.Path(gold_path).resolve())
        raw = pathlib.Path(self.gold_path).read_bytes()
        self.gold_hash = hashlib.sha256(raw).hexdigest()
        self.definition_hashes = {}
        self.slots = {}
        self.notes = {}
        self.tables = {}
        for line_number, line in enumerate(raw.decode("utf-8-sig").splitlines(), 1):
            if not line.strip():
                continue
            item = json.loads(line)
            if item.get("status") != "active":
                continue
            if item["id"] in self.slots:
                raise SemanticError(f"金标准第{line_number}行存在重复ID")
            # 只发送标准定义，不发送来源文件、人工样式映射和冗长历史证据。
            slim = {key: item.get(key) for key in ("id", "scope", "note", "table", "slot", "row_path", "column_path", "value_type", "dimensions", "aliases", "blank_policy", "applicability")}
            slim["applicability"] = item.get("applicability", {})
            slim["aggregation"] = copy.deepcopy(item.get("aggregation"))
            if isinstance(slim["aggregation"],dict):
                for key in ("parent_variants","component_variants"):
                    variants=slim["aggregation"].get(key)
                    if isinstance(variants,list):
                        # 原来源留在完整标准中；发送相同业务关系一次，保留不同总体、期间及全部限定。
                        unique={}
                        for variant in variants:
                            relation={k:v for k,v in variant.items() if k!="source_table_ids"} if isinstance(variant,dict) else variant
                            unique.setdefault(_json(relation),relation)
                        slim["aggregation"][key]=list(unique.values())
            slim["calculation"] = copy.deepcopy(item.get("calculation"))
            self.definition_hashes[item["id"]] = _hash(item)
            self.slots[item["id"]] = slim
            self.notes[item["note"]["id"]] = {"id": item["note"]["id"], "name": item["note"]["name"], "meaning": item["note"].get("meaning", ""), "scope": item.get("scope")}
            self.tables[item["table"]["id"]] = {"id": item["table"]["id"], "name": item["table"]["name"], "meaning": item["table"].get("meaning", ""), "note_id": item["note"]["id"], "scope": item.get("scope")}
        for note in self.notes.values():
            note["allowed_scopes"] = sorted({scope for slot in self.slots.values() if slot["note"]["id"] == note["id"] for scope in allowed_scopes(slot)})
        for table in self.tables.values():
            table["definition_scope"] = table["scope"]
            table["allowed_scopes"] = sorted({scope for slot in self.slots.values() if slot["table"]["id"] == table["id"] for scope in allowed_scopes(slot)})
        if not self.slots:
            raise SemanticError("金标准没有 active 槽位")
        self.usage = collections.Counter()
        self.report_scope = str(self.settings.get("report_scope") or "")
        if self.report_scope and self.report_scope not in {"consolidated", "parent", "standalone"}:
            raise SemanticError("指定的实际报表口径无效")
        self.source_hash = ""
        self.cache_source_hash = ""
        self._review_receipts = []

    def _check_cancel(self):
        if self.cancel is not None:
            cancelled = self.cancel.is_set() if hasattr(self.cancel, "is_set") else self.cancel()
            if cancelled:
                raise SemanticCancelled("已取消；成功任务包已保存，可继续识别")

    def _redact(self,value):
        secret=str(self.settings.get("api_key") or "")
        if isinstance(value,dict):return {self._redact(str(key)):self._redact(item) for key,item in value.items()}
        if isinstance(value,(list,tuple)):return [self._redact(item) for item in value]
        if not isinstance(value,str) or not secret:return value
        for variant in {secret,json.dumps(secret,ensure_ascii=False)[1:-1],json.dumps(secret,ensure_ascii=True)[1:-1]}:
            value=value.replace(variant,"[密钥]")
        if "\\u" in value:
            pattern="".join("(?:"+re.escape(char)+r"|\\u"+format(ord(char),"04x")+")" for char in secret)
            value=re.sub(pattern,"[密钥]",value,flags=re.I)
        return value

    def _response_failure(self,raw,data,error,system,payload):
        response=self._redact(data)
        choices=response.get("choices") if isinstance(response,dict) else None
        choice=choices[0] if isinstance(choices,list) and choices and isinstance(choices[0],dict) else {}
        message=choice.get("message") if isinstance(choice.get("message"),dict) else {}
        text=_json(response) if data is not None else self._redact(raw.decode("utf-8",errors="backslashreplace"))
        diagnostic={"type":type(error).__name__,"message":self._redact(str(error))}
        for source,target in (("pos","position"),("lineno","line"),("colno","column")):
            if hasattr(error,source):diagnostic[target]=getattr(error,source)
        evidence={"time":datetime.datetime.now().isoformat(),"request_hash":_hash({"system":system,"payload":payload}),
            "raw_response_sha256":hashlib.sha256(raw).hexdigest(),"raw_response_bytes":len(raw),
            "raw_response":text,"raw_response_redacted":True,"finish_reason":choice.get("finish_reason"),
            "usage":response.get("usage") if isinstance(response,dict) else None,"parse_error":diagnostic}
        # 无法解码字节及原响应已有替换字符用可见转义保留，不制造中文乱码。
        evidence=json.loads(_json(evidence).replace("\ufffd",r"\\ufffd"))
        path=self.work_dir/"失败响应"/("响应失败_"+uuid.uuid4().hex+".json")
        self._save(path,evidence)
        previous=message.get("content") if isinstance(message.get("content"),str) else text
        return SemanticResponseError(path,self._redact(previous),diagnostic)

    def _save(self, path, value):
        from .统一语义 import json_date
        data = json.dumps(self._redact(value), ensure_ascii=False, indent=2,
                          default=lambda item:json_date(item) if isinstance(item,datetime.date) else str(item))
        if "\ufffd" in data:
            raise SemanticError("写入内容包含异常替换字符")
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            relative = path.relative_to(self.work_dir)
            backup = pathlib.Path.home() / "BackUp" / ("附注自动更新系统_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")) / "语义任务" / relative
            backup.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, backup)
            if hashlib.sha256(path.read_bytes()).digest() != hashlib.sha256(backup.read_bytes()).digest():
                raise SemanticError("覆盖任务回执前备份校验失败")
        path.write_text(data, encoding="utf-8-sig")
        if json.loads(path.read_text(encoding="utf-8-sig")) != json.loads(data):
            raise SemanticError("任务回执回读校验失败")

    def _endpoint(self):
        base = str(self.settings.get("base_url", "")).strip().rstrip("/")
        parsed = urllib.parse.urlsplit(base)
        if parsed.scheme not in {"http", "https"} or not parsed.hostname:
            raise SemanticError("请在设置中填写有效的模型服务地址")
        if not self.settings.get("model"):
            raise SemanticError("请在设置中填写模型名称")
        if parsed.username or parsed.password or parsed.query or parsed.fragment:
            raise SemanticError("服务地址只填写接口地址；API Key 填在独立输入框")
        if base.endswith("/chat/completions"):
            return base
        return base + ("/v1" if not parsed.path or parsed.path == "/" else "") + "/chat/completions"

    def _read_http_response(self,request,simple_a,payload):
        with urllib.request.build_opener(SameOriginRedirect()).open(request, timeout=float(self.settings.get("timeout", DEFAULT_TIMEOUT))) as response:
            if simple_a:
                stream_dir=self.work_dir/'A实时回执';stream_dir.mkdir(parents=True,exist_ok=True)
                with (stream_dir/(payload['task']+'_'+uuid.uuid4().hex+'.txt')).open('xb') as stream_file:
                    class RecordedResponse:
                        headers=response.headers
                        def __iter__(self):
                            for line in response:
                                stream_file.write(line);stream_file.flush()
                                yield line
                        def read(self):
                            content=response.read();stream_file.write(content);stream_file.flush()
                            return content
                    return _read_sse_stream(RecordedResponse())
            return _read_sse_stream(response)

    def _http(self, system, payload):
        self._check_cancel()
        codex = self.settings.get('base_url') == 'codex://local'
        materials = payload['materials']
        plain = {name: value for name, value in materials.items() if name.endswith('.txt') and isinstance(value, str)}
        if plain:
            content = _json({name: value for name, value in materials.items()
                             if name not in plain and name != '本次作答位置.json'})
            for name, value in plain.items():
                content += '\n\n【' + name + '】\n' + value
            if '本次作答位置.json' in materials:
                content += '\n\n【本次作答位置.json】\n' + _json(materials['本次作答位置.json'])
        else:
            content = _json(materials)
        body = {'model': self.settings['model'], 'messages': [{'role': 'system', 'content': system},
                {'role': 'user', 'content': content}], 'stream': True, 'stream_options': {'include_usage': True}}
        images = payload['materials'].get('原表图片.json', [])
        if images:
            if codex:
                body['images'] = images
            else:
                from .Codex语义调用 import image_files
                files = image_files(images)
                body['messages'][1]['content'] = [{'type': 'text', 'text': content}] + [
                    {'type': 'image_url', 'image_url': {'url': 'data:' + (mimetypes.guess_type(file)[0] or 'image/png') +
                     ';base64,' + base64.b64encode(pathlib.Path(file).read_bytes()).decode('ascii')}} for file in files]
        thinking = validate_thinking_mode(self.settings.get('thinking_mode', 'default'))
        if thinking != 'default':
            if str(self.settings['model']).startswith('mimo-'):
                body['thinking'] = {'type': thinking}
            else:
                body['chat_template_kwargs'] = {'enable_thinking': thinking == 'enabled'}
        if codex:
            from .Codex语义调用 import call_codex
            data, raw = call_codex(self.settings, self.work_dir, body, self._check_cancel)
        else:
            headers = {'Content-Type': 'application/json'}
            if self.settings.get('api_key'):
                headers['Authorization'] = 'Bearer ' + str(self.settings['api_key'])
            request = urllib.request.Request(self._endpoint(), data=json.dumps(body, ensure_ascii=False).encode('utf-8'),
                                             headers=headers, method='POST')
            data, raw = _call_with_deadline(float(self.settings.get('timeout', DEFAULT_TIMEOUT)),
                self._read_http_response, request, True, payload)
        self._save(self.work_dir / ('原始响应_' + uuid.uuid4().hex + '.json'), data)
        self.usage['requests'] += 1
        for key in ('prompt_tokens', 'completion_tokens', 'total_tokens'):
            self.usage[key] += int((data.get('usage') or {}).get(key, 0) or 0)
        if data.get('transport_error'):
            raise SemanticError(data['transport_error'])
        choice = data['choices'][0]
        if choice.get('finish_reason') not in (None, 'stop'):
            raise SemanticError('答复未完整结束，已保存原始响应；未自动重答')
        answer = choice['message'].get('content')
        if not isinstance(answer, str) or not answer.strip():
            raise SemanticError('模型没有返回文字答复；未自动重答')
        return answer

    def _request(self, system, payload, validator):
        raise SemanticError('旧的分轮识别入口已停用，请使用A/B直接识别入口；不会调用模型。')

    def request_text(self, prompt, materials):
        """保存一次实际答复；不解析格式、不补问、不调用旧核查。"""
        self._check_cancel()
        from .Codex语义调用 import image_files
        images = materials.get('原表图片.json', [])
        image_files(images)
        identity = {'prompt': prompt, 'materials': materials, 'model': self.settings.get('model'),
                    'endpoint': self.settings.get('base_url'), 'thinking': self.settings.get('thinking_mode'),
                    'effort': self.settings.get('model_reasoning_effort')}
        directory = self.work_dir / '直接识别请求' / _hash(identity)
        receipt = directory / '答复.json'
        if receipt.is_file():
            saved = json.loads(receipt.read_text(encoding='utf-8-sig'))
            if isinstance(saved.get('text'), str) and saved.get('sha256') == _hash(saved['text']):
                self.usage['cached_requests'] += 1
                return saved['text']
            raise SemanticError('已保存答复损坏；未重新请求模型')
        failure = directory / '失败.txt'
        if failure.is_file():
            raise SemanticError('同一请求已有失败记录；未自动重答：' + failure.read_text(encoding='utf-8'))
        directory.mkdir(parents=True, exist_ok=True)
        material_dir = directory / '材料'
        if material_dir.exists():
            material_dir = directory / ('接续材料_' + uuid.uuid4().hex)
        material_dir.mkdir()
        for name, value in materials.items():
            if isinstance(value, str):
                (material_dir / name).write_text(value, encoding='utf-8')
            else:
                self._save(material_dir / name, value)
        if not (directory / '任务.json').exists():
            self._save(directory / '任务.json', {'prompt': prompt, 'model': self.settings.get('model'),
                       'materials': list(materials), 'input_hash': _hash(identity)})
        payload = {'task': 'direct_recognition', 'materials': materials, 'material_directory': str(material_dir.resolve())}
        try:
            answer = self._http(prompt, payload)
            image_files(images)
            if not isinstance(answer, str) or not answer.strip():
                raise SemanticError('没有收到非空答复；未自动重答')
            self._save(receipt, {'text': answer, 'sha256': _hash(answer), 'usage': dict(self.usage)})
            return answer
        except BaseException as error:
            failure.write_text(self._redact(str(error)), encoding='utf-8')
            raise

    def _retry_receipts(self, start, end=None):
        # 冲突回执需要重试；保留文件和备份，不通过删除缓存实现恢复。
        for path in dict.fromkeys(self._review_receipts[start:end]):
            saved = json.loads(path.read_text(encoding="utf-8-sig"))
            if not saved.get("retry_required"):
                saved["retry_required"] = True
                self._save(path, saved)

    @staticmethod
    def _check_exclusion(address, raw, sheet, table=None, *, allow_period_header=True):
        source = sheet["cells"][address]
        value = source.get("cached_value") if source.get("formula") else source.get("value")
        category = raw.get("category")
        if category not in EXCLUDED_TYPES:
            raise SemanticError(f"{sheet['name']}!{address}：排除单元格必须说明有效类别")
        if source.get("formula") and value is None:
            raise SemanticError(f"{sheet['name']}!{address}：公式没有可核验结果，不能作为非业务格排除")
        if category == "empty_padding" and value not in (None, ""):
            raise SemanticError(f"{sheet['name']}!{address}：非空格不能当作空白填充排除")
        text = str(value).strip().replace(",", "").replace("，", "").replace(" ", "")
        text = text.replace("％", "%").removesuffix("%")
        if text.startswith("(") and text.endswith(")"):
            text = "-" + text[1:-1]
        numeric = not isinstance(value, bool) and bool(re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", text))
        if numeric:
            # 表头年份可以是数字；必须处于已识别表头且有明确期间文字证据。
            evidence = raw.get("evidence_cells") or []
            year = bool(re.fullmatch(r"(?:19|20|21)\d{2}", text))
            period_evidence = any(a in sheet["cells"] and isinstance(sheet["cells"][a].get("value"), str)
                                  and (any(word in sheet["cells"][a]["value"] for word in ("年度", "年份", "期末", "期初", "本期", "上期"))
                                       or re.search(r"(?:19|20|21)\d{2}年", sheet["cells"][a]["value"]))
                                  for a in evidence)
            administrative=False
            if category in {"label","annotation"} and not source.get("formula") and re.fullmatch(r"[0-9]+",str(value).strip()):
                row,column=_address(address)
                for header in evidence:
                    cell=sheet["cells"].get(header,{})
                    label=re.sub(r"\s+","",str(cell.get("value",""))).replace("（","(").replace("）",")").rstrip(":：")
                    if cell.get("formula") or label not in {"邮政编码","邮编","电话号码(区号)","电话区号","电话号码","联系电话","手机号码","设立年份","成立年份"}:
                        continue
                    if label in {"设立年份","成立年份"} and not (len(str(value).strip())==4 and 1800<=int(value)<=datetime.date.today().year):
                        continue
                    area=next((_range(merged) for merged in sheet.get("merges",[]) if _inside(header,_range(merged))),_range(header))
                    if not (area[0]<=row<=area[2] and area[3]+1==column):continue
                    if "word_table_ranges" in sheet and not any(
                            box[0]<=area[0]<=area[2]<=box[2] and box[1]<=area[1]<=area[3]<=box[3] and _inside(address,box)
                            for item in sheet["word_table_ranges"] for box in [_range(item["range"])]):
                        continue
                    administrative=True
                    break
            numbering = False
            if (category in {"header", "label", "annotation"} and not source.get("formula") and table and table.get("range")
                    and re.fullmatch(r"[+]?\d+(?:\.0+)?", str(value).strip())):
                bounds = _range(table.get("range"))
                row, column = _address(address)
                merges = [_range(area) for area in sheet.get("merges", [])]
                def span(cell):
                    r, c = _address(cell)
                    return next((area for area in merges if _inside(cell, area)), (r, c, r, c))
                if _inside(address, bounds):
                    for header in evidence:
                        cell = sheet["cells"].get(header, {})
                        label = re.sub(r"\s+", "", str(cell.get("value", "")))
                        if header not in table.get("header_cells", []) or cell.get("formula") or label not in {"行次", "序号"}:
                            continue
                        top, left, bottom, right = span(header)
                        if not (bounds[1] <= left <= right <= bounds[3] and bottom < row and left <= column <= right):
                            continue
                        if "word_table_ranges" in sheet:
                            physical = [_range(item["range"]) for item in sheet["word_table_ranges"]]
                            if not any(box[0] <= top <= bottom <= box[2] and box[1] <= left <= right <= box[3]
                                       and _inside(address, box) for box in physical):
                                continue
                        elif top < bounds[0]:
                            # Excel共享表头须沿同列连续回读，不能跨空行借用前一张表的编号头。
                            covered = set()
                            for other, content in sheet["cells"].items():
                                if content.get("value") in (None, ""):
                                    continue
                                area = span(other)
                                if area[1] <= column <= area[3]:
                                    covered.update(range(max(bottom + 1, area[0]), min(row, area[2] + 1)))
                            if set(range(bottom + 1, row)) - covered:
                                continue
                        # 回读中间实际文字，不能靠漏报更近的金额表头把金额当作编号。
                        conflict = False
                        for other, content in sheet["cells"].items():
                            text_value = content.get("value")
                            if not isinstance(text_value, str) or not text_value.strip() or re.fullmatch(r"[+]?\d+(?:\.0+)?", text_value.strip()):
                                continue
                            near = span(other)
                            if bottom < near[0] < row and near[1] <= column <= near[3] and re.sub(r"\s+", "", text_value) not in {"行次", "序号"}:
                                conflict = True
                                break
                        if not conflict:
                            numbering = True
                            break
            if not administrative and not numbering and not (allow_period_header and category == "header" and year and table and address in table.get("header_cells", []) and period_evidence):
                raise SemanticError(f"{sheet['name']}!{address}：数值格不能仅凭非业务分类排除，须核验其财务含义")

    @staticmethod
    def _exclusion_kind(item, sheet):
        # 仅统一经排除验证的真实文字标签称呼；数值、公式和业务空白仍受原防护约束。
        category = item.get("category")
        cell = sheet["cells"].get(item.get("cell"), {})
        value = cell.get("value")
        if category in {"header", "label", "annotation"} and not cell.get("formula") and re.fullmatch(r"[+]?\d+(?:\.0+)?", str(value).strip()):
            address = item["cell"]
            # 以header重新走原数值保护，只可能凭真实编号列通过；关闭年份表头和行政号码例外。
            numbering_context = {"range": address + ":" + address, "header_cells": item.get("evidence_cells", [])}
            try:
                SemanticEngine._check_exclusion(address, {**item, "category": "header"}, sheet,
                                                numbering_context, allow_period_header=False)
            except SemanticError:
                pass
            else:
                return "header"
        if category in {"header", "title", "label"} and not cell.get("formula") and isinstance(value, str) and value.strip() not in {"", "-", "—", "--", "－"}:
            return "header"
        return category

    def _nonbusiness_cells(self, response, sheet, owned, *, errors=None):
        result = {}; seen=set(); issues=[] if errors is None else errors
        regions = response.get("non_business", [])
        if not isinstance(regions, list):
            raise SemanticError("非业务区域必须是数组")
        for region in regions:
            try:
                if not isinstance(region,dict):raise SemanticError("非业务区域须为对象")
                bounds = _range(region.get("range"))
                if bounds[2] > sheet["max_row"] or bounds[3] > sheet["max_column"]:
                    raise SemanticError("非业务区域超出工作表")
                evidence = region.get("evidence_cells")
                if not isinstance(evidence,list) or not evidence or any(not isinstance(a,str) or a not in sheet["cells"] for a in evidence):
                    raise SemanticError("非业务区域缺少可回读的证据单元格")
                if not isinstance(region.get("reason"), str) or not region["reason"].strip():
                    raise SemanticError("非业务区域缺少判断依据")
            except (ValueError,TypeError,KeyError) as error:
                issues.append(str(error));continue
            # 区域只提供编号列的边界和原证据，不冒充已独立识别的数字期间表头。
            numbering_context={"range":region["range"],"header_cells":evidence}
            rejected=collections.defaultdict(list)
            for address in owned:
                if _inside(address, bounds):
                    if address in seen:
                        rejected["非业务区域重复覆盖"].append(address);continue
                    seen.add(address)
                    try:
                        self._check_exclusion(address,region,sheet,numbering_context,allow_period_header=False)
                    except SemanticError as error:
                        rejected[str(error).split("：",1)[-1]].append(address);continue
                    result[address] = {"sheet": sheet["name"], "cell": address, "category": region["category"],
                                       "reason": region["reason"], "evidence_cells": evidence}
            for reason,addresses in rejected.items():
                issues.append(f"{sheet['name']} {region['range']}：{reason}（"+"、".join(addresses)+"）")
        if errors is None and issues:raise SemanticError("\n".join(issues))
        if response.get("layout_recovery"):
            for item in result.values():item["layout_recovery"]=copy.deepcopy(response["layout_recovery"])
        return result

    def _layout_unresolved_cells(self, response, sheet, owned):
        """保留真实业务但尚无确定科目归属的原格，不能借非业务分类丢弃。"""
        regions=response.get("unresolved",[])
        if not isinstance(regions,list):raise SemanticError("布局待核实区域须为列表")
        result={};occupied=[_range(item["range"]) for key in ("tables","non_business") for item in response.get(key,[])]
        for region in regions:
            if not isinstance(region,dict):raise SemanticError("布局待核实区域须为对象")
            bounds=_range(region.get("range"))
            if bounds[2]>sheet["max_row"] or bounds[3]>sheet["max_column"]:
                raise SemanticError("布局待核实区域超出实际工作表")
            if any(not (bounds[2]<old[0] or bounds[0]>old[2] or bounds[3]<old[1] or bounds[1]>old[3]) for old in occupied):
                raise SemanticError("布局待核实区域与业务、非业务或另一待核实区域重叠")
            occupied.append(bounds)
            evidence=region.get("evidence_cells")
            if not isinstance(evidence,list) or not evidence or any(not isinstance(a,str) or a not in sheet["cells"] for a in evidence):
                raise SemanticError("布局待核实区域缺少可回读的真实证据格")
            reason=region.get("reason")
            if not isinstance(reason,str) or not reason.strip():raise SemanticError("布局待核实区域缺少具体原因")
            for address in owned:
                if _inside(address,bounds):
                    result[address]={"sheet":sheet["name"],"cell":address,"reason":"布局待核实："+reason.strip(),
                        "evidence_cells":list(evidence),"table_range":region["range"],"layout_unresolved":True}
        if response.get("layout_recovery"):
            for item in result.values():item["layout_recovery"]=copy.deepcopy(response["layout_recovery"])
        return result

    def _validate_layout(self, response, sheet, owned, *, single_note=False):
        tables = response.get("tables")
        if not isinstance(tables, list):
            raise SemanticError("布局回执缺少业务表清单")
        used = set()
        table_bounds = []
        for table in tables:
            bounds = _range(table.get("range"))
            if any(not (bounds[2] < old[0] or bounds[0] > old[2] or bounds[3] < old[1] or bounds[1] > old[3]) for old in table_bounds):
                raise SemanticError("业务表范围重叠，不能确定缺失格归属")
            table_bounds.append(bounds)
            if bounds[2] > sheet["max_row"] or bounds[3] > sheet["max_column"]:
                raise SemanticError("识别的业务表范围超出实际工作表")
            physical = [_range(item["range"]) for item in sheet.get("word_table_ranges", [])]
            if physical and not any(box[0] <= bounds[0] <= bounds[2] <= box[2] and
                                    box[1] <= bounds[1] <= bounds[3] <= box[3] for box in physical):
                available="、".join(item["range"] for item in sheet["word_table_ranges"][:50])
                raise SemanticError("业务范围"+str(table.get("range"))+"越出Word实际表格边界；范围须落在单张实际表内："+available)
            note_ids = table.get("note_ids")
            if not isinstance(note_ids, list) or not note_ids or any(n not in self.notes for n in note_ids):
                raise SemanticError("识别的科目不在金标准中")
            if single_note and len(note_ids)!=1:
                raise SemanticError("每个新业务区段只能对应一个金标准科目；不同科目请按实际项目和下级明细拆分，未知归属列为unresolved")
            scope = table.get("scope")
            if self.report_scope and scope != self.report_scope:
                raise SemanticError("模型识别的口径与本次已确认报告口径不一致")
            if scope not in {"consolidated", "parent", "standalone"}:
                raise SemanticError("业务表实际报表口径无效")
            coverage = {a for a in owned if _inside(a, bounds)}
            if used & coverage:
                raise SemanticError("布局回执对同一格分配了多个业务表")
            used |= coverage
            if not isinstance(table.get("table_semantic"), str) or not table["table_semantic"].strip():
                raise SemanticError("业务表含义为空")
            if any(a not in sheet["cells"] for a in table.get("header_cells", [])):
                raise SemanticError("布局引用了不存在的表头")
        issues=[]
        try:self._nonbusiness_cells(response,sheet,owned,errors=issues)
        except (ValueError,TypeError,KeyError) as error:issues.append(str(error))
        declared_nonbusiness=set()
        for region in response.get("non_business",[]) if isinstance(response.get("non_business",[]),list) else []:
            try:bounds=_range(region.get("range"))
            except (ValueError,TypeError,AttributeError):continue
            declared_nonbusiness.update(a for a in owned if _inside(a,bounds))
        overlap=sorted(used & declared_nonbusiness,key=_address)
        if overlap:issues.append("业务表和非业务区域重叠："+"、".join(overlap))
        try:unresolved=self._layout_unresolved_cells(response,sheet,owned)
        except (ValueError,TypeError,KeyError) as error:
            issues.append(str(error));unresolved={}
        missing=sorted(set(owned)-used-declared_nonbusiness-set(unresolved),key=_address)
        if missing:
            issues.append(f"布局遗漏 {len(missing)} 个候选格："+"、".join(missing[:20])+"。须逐一归入有依据的业务区段、非业务区域或有具体原因和原格证据的unresolved区域，不能无说明遗漏。")
        if issues:raise SemanticError("\n".join(issues))
        return response



    @staticmethod
    def _layout_cell_partition(response, candidates):
        """固定格协议的整体归属必须完整唯一；局部恢复不能接受矩形或包外格。"""
        checked=copy.deepcopy(response);seen=set()
        for kind in ("tables","non_business","unresolved"):
            items=checked.get(kind,[])
            if not isinstance(items,list):raise SemanticError("逐格布局分类须为列表")
            for item in items:
                if not isinstance(item,dict):raise SemanticError("逐格布局记录须为对象")
                area=str(item.get("range",""));address=area.split(":")[0]
                if area not in {address,address+":"+address} or address not in candidates or address in seen:
                    raise SemanticError("逐格布局含重复、包外或非单格范围")
                item["range"]=address+":"+address;seen.add(address)
        if len(set(candidates))!=len(candidates) or seen!=set(candidates):raise SemanticError("逐格布局遗漏或重复候选格")
        return checked


    @staticmethod
    def _check_cell_table_evidence(table, sheet, address):
        if table.get("range")!=address+":"+address or address not in sheet["cells"]:
            raise SemanticError("逐格业务证据须为实际已有原格")
        if (not isinstance(table.get("note_ids"),list) or len(table["note_ids"])!=1
                or not isinstance(table["note_ids"][0],str) or not table["note_ids"][0].strip()
                or table.get("scope") not in {"consolidated","parent","standalone"}
                or not isinstance(table.get("table_semantic"),str) or not table["table_semantic"].strip()):
            raise SemanticError("逐格业务证明缺少科目、口径或业务含义")
        headers=table.get("header_cells",[])
        if not isinstance(headers,list) or any(not isinstance(a,str) or a not in sheet["cells"] for a in headers):
            raise SemanticError("逐格业务证明引用不存在的原表头")
        row,column=_address(address)
        if row>sheet["max_row"] or column>sheet["max_column"]:
            raise SemanticError("逐格业务证明超出原表实际边界")
        if "word_table_ranges" not in sheet:return
        physical=[_range(item["range"]) for item in sheet["word_table_ranges"]]
        containing=[box for box in physical if _inside(address,box)]
        if len(containing)!=1:raise SemanticError("逐格业务无法唯一定位原Word物理表")
        for header in headers:
            if any(box!=containing[0] and _inside(header,box) for box in physical):
                raise SemanticError("逐格业务不能借用另一张Word物理表的表头")
        # 未落入任何表内的真实标题仍可引用；其业务关联由原文上下文说明。

    def _validate_cell_table(self, table, sheet, address):
        self._validate_layout({"tables":[table]},sheet,[address],single_note=True)
        self._check_cell_table_evidence(table,sheet,address)

    def _cell_table_groups(self, sheet, cell_tables, owned, layout_batch_id):
        """只合并请求候选；外接范围用于原文上下文，不能登记为业务范围。"""
        from openpyxl.utils.cell import get_column_letter
        groups={}
        for address in sorted(owned,key=_address):
            table=copy.deepcopy(cell_tables[address])
            self._validate_cell_table(table,sheet,address)
            physical=[item["range"] for item in sheet.get("word_table_ranges",[]) if _inside(address,_range(item["range"]))]
            if "word_table_ranges" in sheet and len(physical)!=1:
                raise SemanticError("逐格合批无法唯一定位原Word物理表")
            key=(tuple(table["note_ids"]),table["scope"],physical[0] if physical else None)
            group=groups.setdefault(key,{"note_ids":list(table["note_ids"]),"scope":table["scope"],
                "cell_tables":{},"candidate_cells":[],"layout_batch_id":layout_batch_id,"header_cells":[]})
            table["layout_batch_id"]=layout_batch_id
            group["cell_tables"][address]=table;group["candidate_cells"].append(address)
            group["header_cells"]=list(dict.fromkeys(group["header_cells"]+table.get("header_cells",[])))
        for group in groups.values():
            coordinates=[_address(a) for a in group["candidate_cells"]]
            top=min(r for r,c in coordinates);bottom=max(r for r,c in coordinates)
            left=min(c for r,c in coordinates);right=max(c for r,c in coordinates)
            group["range"]=f"{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}"
            group["table_semantic"]="；".join(f"{a}：{t['table_semantic']}" for a,t in group["cell_tables"].items())
        return list(groups.values())

    @staticmethod
    def _cell_table(table, address):
        if "cell_tables" not in table:return table
        raw=table["cell_tables"].get(address)
        if not isinstance(raw,dict) or raw.get("range")!=address+":"+address:
            raise SemanticError("分类缺少该格独立业务范围证据")
        if raw.get("note_ids")!=table.get("note_ids") or raw.get("scope")!=table.get("scope"):
            raise SemanticError("分类合批混入不同科目或口径")
        return raw

    @staticmethod
    def _restore_cell_table(snapshot, sheet, table):
        address=table["range"].split(":")[0]
        SemanticEngine._check_cell_table_evidence(table,sheet,address)
        proof={"source_hash":snapshot["sha256"],"table_range":table["range"],
               "table_semantic":table["table_semantic"],"note_ids":list(table["note_ids"]),
               "scope":table["scope"],"header_cells":list(table.get("header_cells",[]))}
        for key in ("layout_batch_id","layout_review"):
            if key in table:proof[key]=copy.deepcopy(table[key])
        # 当前快照每个原格仅一份生效证明；旧轮原证据仍保存在原映射和回执文件中。
        records=sheet.setdefault("confirmed_business_ranges",[])
        records[:]=[old for old in records if not (old.get("source_hash")==snapshot["sha256"]
                    and old.get("table_range")==table["range"])]
        records.append(proof)



    def _canonical_metric(self, raw, dimensions, slot, sheet):
        default = str((slot.get("slot") or {}).get("name") or "").strip()
        if not default:
            raise SemanticError("金标准缺少规范指标定义")
        declared = next((d for d in slot.get("dimensions") or [] if d.get("name") == "metric"), None)
        if declared is None:
            try:
                normalized = normalize_mapping_metric(slot, {"dimensions": dimensions})
            except ValueError as error:
                raise SemanticError(str(error)) from error
            dimensions.clear()
            dimensions.update(normalized["dimensions"])
            return default
        supplied = dimensions.get("metric")
        aliases = {default, " / ".join(slot.get("column_path") or [])}
        aliases.update((slot.get("aliases") or {}).get("column_labels") or [])
        if supplied is None or (isinstance(supplied, str) and supplied in aliases):
            dimensions["metric"] = default
            return default
        if not isinstance(supplied, str) or not supplied.strip():
            raise SemanticError("实际指标 metric 必须为非空文字")
        if declared.get("open") is not True:
            raise SemanticError("该金标准未声明开放metric维度，不能用其他指标替代原定义；请核实或补充金标准")
        evidence = raw.get("metric_evidence")
        if not isinstance(evidence, list) or not evidence or any(a not in sheet["cells"] for a in evidence):
            raise SemanticError("额外实际指标缺少可回读的 metric_evidence")
        if not any(isinstance(sheet["cells"][a].get("value"), str) and sheet["cells"][a]["value"].strip()
                   and not sheet["cells"][a].get("formula") for a in evidence):
            raise SemanticError("实际指标证据必须包含原始文字表头或项目")
        dimensions["metric"] = supplied.strip()
        return dimensions["metric"]

    def _validate_classification(self, response, candidates, allowed_slots, sheet, table):
        results = {}
        for category in ("mappings", "excluded", "unresolved"):
            items = response.get(category)
            if not isinstance(items, list):
                raise SemanticError("每个任务包必须包含三种分类清单")
            for raw in items:
                if not isinstance(raw, dict):
                    raise SemanticError("单元格回执格式无效")
                address = raw.get("cell")
                if address not in candidates or address not in sheet["cells"]:
                    raise SemanticError(f"{sheet['name']}!{address}：回执引用了包外或不存在的单元格")
                if address in results:
                    raise SemanticError(f"{sheet['name']}!{address}：同一单元格出现多个分类结果")
                reason = raw.get("reason")
                if not isinstance(reason, str) or not reason.strip():
                    raise SemanticError(f"{sheet['name']}!{address}：单元格识别缺少依据")
                cell_table = self._cell_table(table,address)
                if "cell_tables" in table or cell_table.get("layout_batch_id"):
                    self._validate_cell_table(cell_table,sheet,address)
                    original_evidence={address,*cell_table.get("header_cells",[])}
                    for key in ("evidence_cells","metric_evidence"):
                        if key in raw and (not isinstance(raw[key],list)
                                or any(not isinstance(a,str) or a not in original_evidence for a in raw[key])):
                            raise SemanticError("逐格分类证据不属于该原格或其独立表头："+key)
                item = {"sheet": sheet["name"], "cell": address, "reason": reason, "table_range": cell_table["range"], "table_semantic": cell_table["table_semantic"]}
                for key in ("layout_batch_id","layout_review"):
                    if key in cell_table:item[key]=copy.deepcopy(cell_table[key])
                if category == "mappings":
                    slot_id = raw.get("slot_id")
                    if slot_id not in allowed_slots or slot_id not in self.slots:
                        raise SemanticError(f"{sheet['name']}!{address}：回执含未知、非 active 或当前包外的金标准ID")
                    slot = self.slots[slot_id]
                    if raw.get("scope") != cell_table["scope"]:
                        raise SemanticError(f"{sheet['name']}!{address}：单元格与业务表的实际报表口径不一致")
                    validate_mapping_scope(slot, {"scope": raw.get("scope"), "definition_scope": slot["scope"]})
                    dimensions = _normal_dimensions(raw.get("dimensions"), slot)
                    for dimension in slot.get("dimensions") or []:
                        name = dimension["name"]
                        value = dimensions.get(name)
                        if dimension.get("required") and value in (None, "", [], {}):
                            raise SemanticError(f"{sheet['name']}!{address}（{slot_id}）遗漏金标准必填维度：" + name)
                        if value is None:
                            continue
                        fixed = {k: dimension[k] for k in ("lower", "upper", "lower_inclusive", "upper_inclusive") if k in dimension}
                        if fixed and dimension.get("open") is not True:
                            if not isinstance(value, dict) or set(value) != set(fixed) or any(
                                value[k] != expected or (isinstance(value[k], bool) != isinstance(expected, bool))
                                for k, expected in fixed.items()):
                                raise SemanticError(f"{sheet['name']}!{address}（{slot_id}）实际维度与金标准固定边界不一致：" + name + "；请核实或补充金标准")
                            dimensions[name] = copy.deepcopy(fixed)
                        if name == "period":
                            instant = isinstance(value, str) or isinstance(value, dict) and "date" in value and "start" not in value and "end" not in value
                            duration = isinstance(value, dict) and "start" in value and "end" in value and "date" not in value
                            if (dimension.get("type") == "instant" and not instant) or (dimension.get("type") == "duration" and not duration):
                                raise SemanticError(f"{sheet['name']}!{address}：实际期间类型与金标准时点或期间定义不一致")
                    kind = raw.get("value_type")
                    if kind not in VALUE_TYPES:
                        raise SemanticError(f"{sheet['name']}!{address}：单元格实际值类型无效")
                    if kind != slot["value_type"]:
                        raise SemanticError(f"{sheet['name']}!{address}：实际值类型与金标准不一致，不能通过指标或单位改写；请核实或补充金标准")
                    metric = self._canonical_metric(raw, dimensions, slot, sheet)
                    source = sheet["cells"][address]
                    value = source.get("cached_value") if source.get("formula") else source.get("value")
                    # 语义识别与取数分开：公式缺缓存仍可确定身份，执行层必须先取得可靠计算值。
                    if kind in {"monetary", "number", "percentage"} and value not in (None, "", "-", "—", "--", "－"):
                        if isinstance(value, bool):
                            raise SemanticError(f"{sheet['name']}!{address}：布尔值不能作为金额或数量")
                        numeric = str(value).strip().replace(",", "").replace("，", "").replace(" ", "")
                        numeric = numeric.replace("％", "%").removesuffix("%")
                        if numeric.startswith("(") and numeric.endswith(")"):
                            numeric = "-" + numeric[1:-1]
                        if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", numeric):
                            raise SemanticError(f"{sheet['name']}!{address}：源单元格内容不是可核验的数值")
                    field = raw.get("semantic_field")
                    if not isinstance(field, str) or not field.strip():
                        raise SemanticError(f"{sheet['name']}!{address}：缺少实际业务指标含义")
                    item.update({"slot_id": slot_id, "scope": raw["scope"], "definition_scope": slot["scope"], "dimensions": dimensions, "semantic_field": metric, "value_type": kind, "row_label": str(raw.get("row_label", "")), "column_label": str(raw.get("column_label", "")), "aggregation": (slot.get("aggregation") or {}).get("role")})
                elif category == "excluded":
                    self._check_exclusion(address, raw, sheet, cell_table)
                    item["category"] = raw["category"]
                for evidence_key in ("evidence_cells", "metric_evidence"):
                    if evidence_key in raw:
                        item[evidence_key] = copy.deepcopy(raw[evidence_key])
                results[address] = (category, item)
        if set(results) != set(candidates):
            raise SemanticError("回执漏掉候选单元格，不能视为完成："+"、".join(sorted(set(candidates)-set(results),key=_address)[:20]))
        return results




    @staticmethod
    def _meaning(item):
        return {key: item.get(key) for key in ("slot_id", "scope", "dimensions", "value_type")}

    def _candidate_slots(self, table):
        """同名且完整科目含义一致时，按逐槽获准的实际口径补充候选。"""
        def identity(note):
            name, meaning = note.get("name"), note.get("meaning")
            return (name, meaning) if isinstance(name, str) and name.strip() and isinstance(meaning, str) and meaning.strip() else None
        note_ids = set(table["note_ids"])
        identities = {identity(self.notes[n]) for n in note_ids} - {None}
        return [s for s in self.slots.values()
                if (s["note"]["id"] in note_ids or identity(s["note"]) in identities)
                and table["scope"] in allowed_scopes(s)]


    @staticmethod
    def _source_identity(snapshot):
        # 忽略可由原表范围重建的派生空白；其余真实内容、实体和期间上下文全部绑定。
        sheets = copy.deepcopy(snapshot.get("sheets", []))
        for sheet in sheets:
            sheet.pop("confirmed_business_ranges", None)
            sheet["cells"] = {a:c for a,c in sheet.get("cells", {}).items() if not c.get("derived_blank")}
        return _hash({"original": snapshot.get("original_hash") or snapshot.get("sha256"),
                      "sheets": _stable_content(sheets), "context": _stable_content(snapshot.get("context", []))})

    @staticmethod
    def _selection_identity(selection):
        # 范围重验会增加恢复统计和新金标准哈希；复用依据是同一目标、分区及章外证据。
        if selection is None:return _hash(None)
        omitted = [{"sheet":a["sheet"],"cells":sorted(a["cells"]),"reason":a.get("reason"),
                    "evidence_cells":sorted(a.get("evidence_cells",[]))} for a in selection.get("out_of_scope",[])]
        return _hash({"source_hash":selection.get("source_hash"),"target_hash":selection.get("target_hash"),
                      "selected_cells":{s:sorted(c) for s,c in selection.get("selected_cells",{}).items()},
                      "out_of_scope":sorted(omitted,key=_json)})


    def recognize(self, snapshot: dict, reference: list[dict] | None = None, target_scope=None, scope_selection=None, previous_result=None, progress=None) -> dict:
        raise SemanticError("旧分轮识别入口已停用，请使用分步流程中的A/B直接识别。")
