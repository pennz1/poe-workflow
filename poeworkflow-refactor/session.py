"""Session state persistence."""

import hashlib
import json
import os
import time
from typing import Any, Dict

import streamlit as streamlit

from config import PERSIST_DIR, _PERSIST_KEYS

def _get_session_persist_path() -> str:
    """返回当前 Streamlit session 对应的持久化文件路径（per-session 隔离）。"""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        ctx = get_script_run_ctx()
        if ctx and ctx.session_id:
            sid = hashlib.md5(ctx.session_id.encode()).hexdigest()[:8]
        else:
            sid = "default"
    except Exception:
        sid = "default"
    return os.path.join(PERSIST_DIR, f".session_persist_{sid}.json")


def _cleanup_stale_sessions(max_age_hours: int = 24) -> None:
    """删除超过 max_age_hours 的旧 session 持久化文件。"""
    try:
        cutoff = time.time() - max_age_hours * 3600
        for fname in os.listdir(PERSIST_DIR):
            if fname.startswith(".session_persist_") and fname.endswith(".json"):
                fpath = os.path.join(PERSIST_DIR, fname)
                if os.path.getmtime(fpath) < cutoff:
                    os.remove(fpath)
    except Exception:
        pass


def persist_session_state() -> None:
    data: Dict[str, Any] = {}
    for k in _PERSIST_KEYS:
        if k in streamlit.session_state:
            data[k] = streamlit.session_state[k]
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
            streamlit.session_state[k] = v


def clear_session_persist() -> None:
    try:
        os.remove(_get_session_persist_path())
    except OSError:
        pass
