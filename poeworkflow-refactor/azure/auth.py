"""Azure authentication and resource selection helpers."""

import time
from typing import Any, Dict, List

import streamlit as st

from azure.arm import azure_arm_list
from config import (
    AZURE_ARM_SCOPE,
    AZURE_AUTHORITY,
    AZURE_RESOURCE_API_VERSION,
    MSAL_CLIENT_ID_DEFAULT,
    get_secret,
)
from frontend.ui import render_device_code_login
from session import clear_session_persist, persist_session_state

try:
    import msal
except ImportError:
    msal = None

def _get_msal_client_id() -> str:
    """返回 MSAL Public Client ID。Client ID 不是密钥，可使用默认值。"""
    return get_secret("MSAL_CLIENT_ID", MSAL_CLIENT_ID_DEFAULT)


def is_azure_token_valid() -> bool:
    expires_at = st.session_state.get("azure_token_expires_at", 0)
    return bool(st.session_state.get("azure_token")) and time.time() < expires_at


def msal_device_code_login() -> None:
    """通过 MSAL Device Code Flow 登录 Azure，并将 access token 放入 session state。"""
    if msal is None:
        raise RuntimeError("缺少 msal 依赖，请确认 requirements.txt 已包含 msal。")

    app = msal.PublicClientApplication(
        _get_msal_client_id(),
        authority=AZURE_AUTHORITY,
    )
    flow = app.initiate_device_flow(scopes=AZURE_ARM_SCOPE)
    if "user_code" not in flow:
        raise RuntimeError(f"无法启动 Microsoft 登录流程：{flow}")

    user_code = flow["user_code"]
    verify_url = flow.get("verification_uri", "https://microsoft.com/devicelogin")
    render_device_code_login(user_code, verify_url)

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or "Microsoft 登录失败。"
        if "7000218" in str(err):
            err += "\n\n请在 Azure Portal → 应用注册 → 身份验证 → 高级设置中，将「允许公共客户端流」设为「是」。"
        raise RuntimeError(err)

    account = result.get("id_token_claims", {}) or result.get("account", {}) or {}
    username = account.get("preferred_username") or account.get("email") or account.get("username") or "Azure 用户"
    expires_in = int(result.get("expires_in", 3600))
    st.session_state["azure_token"] = result["access_token"]
    st.session_state["azure_user"] = username
    st.session_state["azure_token_expires_at"] = time.time() + max(expires_in - 300, 300)
    persist_session_state()


def clear_azure_login() -> None:
    for key in [
        "azure_token",
        "azure_user",
        "azure_token_expires_at",
        "azure_subscription_id",
        "azure_subscription_name",
        "azure_resource_group",
        "_cached_subscription",
        "_cached_resource_group",
    ]:
        st.session_state.pop(key, None)
    clear_session_persist()

def list_azure_subscriptions(token: str) -> List[Dict[str, Any]]:
    subscriptions = azure_arm_list(f"/subscriptions?api-version={AZURE_RESOURCE_API_VERSION}", token)
    return sorted(subscriptions, key=lambda item: item.get("displayName", ""))


def list_azure_resource_groups(subscription_id: str, token: str) -> List[Dict[str, Any]]:
    groups = azure_arm_list(
        f"/subscriptions/{subscription_id}/resourceGroups?api-version={AZURE_RESOURCE_API_VERSION}",
        token,
    )
    return sorted(groups, key=lambda item: item.get("name", ""))


def _subscription_label(subscription: Dict[str, Any]) -> str:
    display_name = subscription.get("displayName") or subscription.get("subscriptionId")
    state = subscription.get("state", "Unknown")
    return f"{display_name} ({state})"


def _resource_group_label(resource_group: Dict[str, Any]) -> str:
    name = resource_group.get("name", "")
    location = resource_group.get("location", "")
    return f"{name} ({location})" if location else name
