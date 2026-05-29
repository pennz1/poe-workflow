"""Azure ARM REST helpers."""

import time
from typing import Any, Dict, List, Optional

import requests

from config import AZURE_MANAGEMENT_ENDPOINT, AZURE_PROVIDER_API_VERSION

def _format_arm_error(response: requests.Response) -> str:
    try:
        payload = response.json()
        error = payload.get("error", payload)
        message = error.get("message") or str(error)
    except Exception:
        message = response.text
    return f"Azure API {response.status_code}: {message[:1200]}"


def _retry_after_seconds(response: requests.Response, default: int = 10) -> int:
    retry_after = response.headers.get("Retry-After")
    if not retry_after:
        return default
    try:
        return max(1, min(int(retry_after), 60))
    except ValueError:
        return default


def _poll_azure_lro(
    operation_url: str,
    token: str,
    initial_delay: int = 10,
    timeout_seconds: int = 900,
) -> Dict[str, Any]:
    """轮询 ARM long-running operation，直到 Succeeded/Failed/Canceled。"""
    if operation_url.startswith("/"):
        operation_url = f"{AZURE_MANAGEMENT_ENDPOINT}{operation_url}"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    deadline = time.time() + timeout_seconds
    delay = max(1, min(initial_delay, 60))

    while time.time() < deadline:
        time.sleep(delay)
        response = requests.get(operation_url, headers=headers, timeout=90)
        if response.status_code >= 400:
            raise RuntimeError(_format_arm_error(response))

        payload: Dict[str, Any] = {}
        if response.content:
            try:
                payload = response.json()
            except ValueError:
                payload = {}

        status = str(
            payload.get("status")
            or payload.get("properties", {}).get("provisioningState")
            or ""
        ).strip()
        status_lower = status.lower()
        if status_lower in {"succeeded", "completed"}:
            return payload
        if status_lower in {"failed", "canceled", "cancelled"}:
            raise RuntimeError(f"Azure 长操作失败：{payload or status}")
        if response.status_code in {200, 204} and not status:
            return payload

        delay = _retry_after_seconds(response, default=10)

    raise TimeoutError("等待 Azure 长操作完成超时。")


def azure_arm_request(
    method: str,
    path_or_url: str,
    token: str,
    body: Optional[Dict[str, Any]] = None,
    timeout: int = 90,
    poll_lro: bool = True,
    lro_timeout: int = 900,
    max_retries: int = 3,
) -> Dict[str, Any]:
    """调用 Azure ARM REST API。path_or_url 可传完整 URL 或 ARM 相对路径。"""
    url = path_or_url if path_or_url.startswith("http") else f"{AZURE_MANAGEMENT_ENDPOINT}{path_or_url}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    last_exc: Optional[Exception] = None
    for attempt in range(max_retries):
        try:
            response = requests.request(method, url, headers=headers, json=body, timeout=timeout)
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as exc:
            last_exc = exc
            if attempt < max_retries - 1:
                time.sleep(5 * (attempt + 1))
                continue
            raise RuntimeError(
                f"HTTPSConnectionPool(host='management.azure.com', port=443): "
                f"Max retries exceeded with url: {path_or_url.split('?')[0]} — {exc}"
            ) from exc
        # 对可重试的服务端/限流错误自动重试
        if response.status_code in {429, 500, 502, 503, 504} and attempt < max_retries - 1:
            delay = _retry_after_seconds(response, default=10 * (attempt + 1))
            time.sleep(delay)
            continue
        break
    if response.status_code >= 400:
        raise RuntimeError(_format_arm_error(response))
    payload: Dict[str, Any] = {}
    if response.content:
        try:
            payload = response.json()
        except ValueError:
            payload = {}

    lro_url = response.headers.get("Azure-AsyncOperation") or response.headers.get("Location")
    if poll_lro and response.status_code in {201, 202} and lro_url:
        final_payload = _poll_azure_lro(
            lro_url,
            token,
            initial_delay=_retry_after_seconds(response, default=10),
            timeout_seconds=lro_timeout,
        )
        return payload or final_payload

    return payload


def azure_arm_list(path: str, token: str) -> List[Dict[str, Any]]:
    """读取 ARM 列表接口并处理 nextLink 分页。"""
    items: List[Dict[str, Any]] = []
    next_url = f"{AZURE_MANAGEMENT_ENDPOINT}{path}"
    while next_url:
        payload = azure_arm_request("GET", next_url, token)
        items.extend(payload.get("value", []))
        next_url = payload.get("nextLink")
    return items



def register_azure_provider(subscription_id: str, namespace: str, token: str) -> None:
    azure_arm_request(
        "POST",
        f"/subscriptions/{subscription_id}/providers/{namespace}/register?api-version={AZURE_PROVIDER_API_VERSION}",
        token,
    )
