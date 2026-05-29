"""Session state persistence."""

import datetime
import json
import os
import time
from typing import Any, Dict

import streamlit as streamlit

from config import PERSIST_DIR, _PERSIST_KEYS


def _json_serializable(v: Any) -> Any:
    """将不可直接 JSON 序列化的值转为可序列化形式。"""
    if isinstance(v, (datetime.date, datetime.datetime)):
        return {"__date__": v.isoformat()}
    return v


def _json_deserialize(v: Any) -> Any:
    """还原 _json_serializable 转换的对象。"""
    if isinstance(v, dict) and "__date__" in v:
        try:
            return datetime.date.fromisoformat(v["__date__"])
        except (ValueError, TypeError):
            return None
    return v

def _get_session_persist_path() -> str:
    """返回持久化文件路径（固定路径，确保刷新页面后仍可恢复登录态）。"""
    return os.path.join(PERSIST_DIR, ".session_persist.json")


def _cleanup_stale_sessions(max_age_hours: int = 24) -> None:
    """删除超过 max_age_hours 的旧 session 持久化文件。"""
    try:
        cutoff = time.time() - max_age_hours * 3600
        for fname in os.listdir(PERSIST_DIR):
            if fname.startswith(".session_persist_") and fname.endswith(".json"):
                fpath = os.path.join(PERSIST_DIR, fname)
                if os.path.getmtime(fpath) < cutoff:
                    os.remove(fpath)
        # 同时清理主持久化文件中的过期 token
        path = _get_session_persist_path()
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            expires_at = data.get("azure_token_expires_at", 0)
            if expires_at and time.time() > expires_at:
                os.remove(path)
    except Exception:
        pass


def persist_session_state() -> None:
    data: Dict[str, Any] = {}
    for k in _PERSIST_KEYS:
        if k in streamlit.session_state:
            data[k] = _json_serializable(streamlit.session_state[k])
    if not data:
        return
    try:
        path = _get_session_persist_path()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception:
        pass


def restore_session_state() -> None:
    _cleanup_stale_sessions()
    path = _get_session_persist_path()
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return
    for k, v in data.items():
        if k not in streamlit.session_state:
            streamlit.session_state[k] = _json_deserialize(v)


def clear_session_persist() -> None:
    try:
        os.remove(_get_session_persist_path())
    except OSError:
        pass
