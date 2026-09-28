"""仅管理本程序设置；API key 使用当前 Windows 用户的 DPAPI 保护。"""
from __future__ import annotations

import base64
import ctypes
import hashlib
import json
import math
import shutil
from ctypes import wintypes
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit

PROJECT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_SETTINGS_PATH = PROJECT_DIR / "用户配置.json"
DEFAULT_TIMEOUT = 600
LEGACY_GOLD_PATH = r"D:\BaiduSyncdisk\数字资产\审计数据自动搬运系统\语义金标准\财务报表附注语义金标准-v1.jsonl"
DEFAULT_GOLD_PATH = str(PROJECT_DIR / "金标准" / "财务报表附注语义金标准-v2.jsonl")
if not Path(DEFAULT_GOLD_PATH).is_file():
    DEFAULT_GOLD_PATH = str(PROJECT_DIR / "金标准" / "公开语义金标准-v2.jsonl")
_FIELDS = ("source_a", "source_b", "output_dir", "gold_path", "base_url", "model", "timeout", "thinking_mode", "mode", "review", "scope_prefilter", "a_context_text", "b_context_text",
           "word_path", "a_path", "b_path", "a_link", "a_mapping", "b_mapping", "update_plan",
           "a_prime", "a_prime_mapping", "word", "word_link", "trace", "a_review", "b_review", "a_standard_review", "b_standard_review", "b_scope", "a_gold_changes", "b_gold_changes", "report_scope", "current_step",
           "structure_review", "adjusted_a", "adjusted_a_mapping", "unused_source_review")


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_ubyte))]


def _crypt(data: bytes, decrypt: bool) -> bytes:
    """调用 Windows 用户级保护，不允许系统弹出凭据窗口。"""
    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    buffer = ctypes.create_string_buffer(data)
    source = _Blob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)))
    output = _Blob()
    function = crypt32.CryptUnprotectData if decrypt else crypt32.CryptProtectData
    second_type = ctypes.POINTER(wintypes.LPWSTR) if decrypt else wintypes.LPCWSTR
    function.argtypes = [ctypes.POINTER(_Blob), second_type, ctypes.POINTER(_Blob), ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_Blob)]
    function.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    description = None if decrypt else "附注自动更新系统 API key"
    if not function(ctypes.byref(source), description, None, None, None, 1, ctypes.byref(output)):
        raise OSError(ctypes.get_last_error(), "Windows 密钥保护操作失败")
    try:
        return ctypes.string_at(output.pbData, output.cbData)
    finally:
        kernel32.LocalFree(output.pbData)


def protect_secret(secret: str) -> str:
    return base64.b64encode(_crypt(secret.encode("utf-8"), False)).decode("ascii") if secret else ""


def unprotect_secret(encrypted: str) -> str:
    if not encrypted:
        return ""
    try:
        return _crypt(base64.b64decode(encrypted, validate=True), True).decode("utf-8")
    except (ValueError, OSError, UnicodeError) as exc:
        raise ValueError("无法解密本程序保存的 API key，请使用原 Windows 用户，或在界面重新填写并保存。") from exc


def _validate_timeout(value) -> float:
    try:
        timeout = float(value)
    except (ValueError, TypeError, OverflowError):
        raise ValueError("请求超时（秒）须为大于 0 的有限数字。") from None
    if isinstance(value, bool) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError("请求超时（秒）须为大于 0 的有限数字。")
    return timeout


def validate_thinking_mode(value="default") -> str:
    if not isinstance(value, str) or value not in ("default", "enabled", "disabled"):
        raise ValueError("模型思考模式须为服务默认、开启或关闭。")
    return value


def validate_settings(config: dict, require_model: bool = True) -> dict:
    result = dict(config)
    result.setdefault("scope_prefilter", False)
    if not isinstance(result["scope_prefilter"], bool):
        raise ValueError("来源筛选设置须为布尔值。")
    result["thinking_mode"] = validate_thinking_mode(result.get("thinking_mode", "default"))
    if result.get("report_scope", "") not in {"", "consolidated", "parent", "standalone"}:
        raise ValueError("报告口径须为按文件识别、单户、合并或母公司。")
    address = str(result.get("base_url") or "").strip().rstrip("/")
    try:
        parts = urlsplit(address)
        valid = parts.scheme in ("http", "https") and bool(parts.hostname) and not parts.username and not parts.password and not parts.query and not parts.fragment
        _ = parts.port
    except ValueError:
        valid = False
    if (address or require_model) and (not valid or any(char.isspace() for char in address)):
        raise ValueError("模型地址须为有效的 http:// 或 https:// 地址，不得包含账号密码、查询参数或空格。")
    model = str(result.get("model") or "").strip()
    if require_model and not model:
        raise ValueError("请填写服务商提供的模型名称。")
    secret = str(result.get("api_key") or "").strip()
    if "\n" in secret or "\r" in secret:
        raise ValueError("API key 不能包含换行，请重新粘贴。")
    result.update(base_url=address, model=model, api_key=secret, timeout=_validate_timeout(result.get("timeout", DEFAULT_TIMEOUT)))
    return result


def load_settings(path: str | Path = DEFAULT_SETTINGS_PATH) -> dict:
    result = {"source_a": "", "source_b": "", "output_dir": str(PROJECT_DIR / "输出"), "gold_path": DEFAULT_GOLD_PATH, "base_url": "", "model": "", "api_key": "", "timeout": DEFAULT_TIMEOUT, "thinking_mode": "default", "review": True, "scope_prefilter": False, "mode": "update"}
    result.update({key: "" for key in ("word_path", "a_path", "b_path", "a_link", "a_mapping", "b_mapping",
                                      "update_plan", "a_prime", "a_prime_mapping", "word", "word_link", "a_context_text", "b_context_text",
                                      "a_standard_review", "b_standard_review")})
    result["current_step"] = "extract_a"
    path = Path(path)
    if not path.exists():
        return result
    try:
        saved = json.loads(path.read_text(encoding="utf-8-sig"))
        if not isinstance(saved, dict):
            raise ValueError("配置应为对象")
        for key in _FIELDS:
            if key in saved:
                if key == "thinking_mode":
                    result[key] = validate_thinking_mode(saved[key])
                    continue
                if key == "timeout":
                    result[key] = _validate_timeout(saved[key])
                    continue
                if key in {"review", "scope_prefilter"}:
                    if not isinstance(saved[key], bool):
                        raise ValueError("来源筛选设置须为布尔值。" if key == "scope_prefilter" else "复核设置应为布尔值")
                elif not isinstance(saved[key], str):
                    raise ValueError("配置项目应为文本")
                result[key] = saved[key]
        old_context = saved.get("context_text", "")
        if not isinstance(old_context, str):
            raise ValueError("旧补充信息应为文本")
        for key in ("a_context_text", "b_context_text"):
            if key not in saved:
                result[key] = old_context
        if Path(result.get("gold_path") or "") == Path(LEGACY_GOLD_PATH):
            result["gold_path"] = DEFAULT_GOLD_PATH
        result["api_key"] = unprotect_secret(saved.get("encrypted_api_key", ""))
        old_a = result.get("source_a", "")
        if old_a:
            target = "word_path" if Path(old_a).suffix.lower() == ".docx" else "a_path"
            if not result[target]:result[target] = old_a
        if not result["b_path"]:result["b_path"] = result.get("source_b", "")
    except (OSError, ValueError, TypeError) as exc:
        raise ValueError(f"无法读取本程序用户配置：{exc}") from exc
    return result


def save_settings(config: dict, path: str | Path = DEFAULT_SETTINGS_PATH) -> None:
    config = validate_settings(config, require_model=False)
    path = Path(path)
    payload = {key: config[key] for key in _FIELDS if key in config}
    payload["encrypted_api_key"] = protect_secret(config["api_key"])
    payload["key_protection"] = "Windows DPAPI / 当前用户"
    content = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    if path.exists():
        backup_dir = Path.home() / "BackUp" / ("附注自动更新系统_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
        try:
            relative = path.resolve().relative_to(PROJECT_DIR)
        except ValueError:
            relative = Path(path.name)
        backup = backup_dir / relative
        backup.parent.mkdir(parents=True, exist_ok=False)
        shutil.copy2(path, backup)
        if hashlib.sha256(path.read_bytes()).digest() != hashlib.sha256(backup.read_bytes()).digest():
            raise OSError("用户配置备份校验失败，已停止保存。")
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + "." + datetime.now().strftime("%Y%m%d%H%M%S%f") + ".tmp")
    with temporary.open("x", encoding="utf-8-sig") as handle:
        handle.write(content)
    json.loads(temporary.read_text(encoding="utf-8-sig"))
    temporary.replace(path)
