"""OpenAI-compatible client wrapper."""

from openai import OpenAI

from config import get_secret


def get_openai_client() -> OpenAI:
    """创建 OpenAI 兼容客户端实例（支持 NewAPI 等网关）。"""
    endpoint = get_secret("AZURE_OPENAI_ENDPOINT").rstrip("/")
    base_url = endpoint if endpoint.endswith("/v1") else endpoint + "/v1"
    return OpenAI(
        api_key=get_secret("AZURE_OPENAI_KEY"),
        base_url=base_url,
    )


def call_azure_openai(system_prompt: str, user_prompt: str) -> str:
    """调用 OpenAI 兼容 Chat Completions API 并返回文本结果。"""
    client = get_openai_client()
    response = client.chat.completions.create(
        model=get_secret("AZURE_OPENAI_DEPLOYMENT"),
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.7,
        max_tokens=16384,
    )
    # openai SDK 在收到非标准 JSON 响应时可能返回原始字符串而非 ChatCompletion 对象
    if isinstance(response, str):
        raise RuntimeError(
            f"API 返回了非预期的原始字符串响应。响应内容: {response[:500]}"
        )
    if not hasattr(response, "choices") or not response.choices:
        raise RuntimeError(
            f"API 返回了无效响应结构: {type(response).__name__}。"
            f"请检查 API 密钥、端点和模型名称是否正确。"
        )
    content = response.choices[0].message.content
    if not content or not content.strip():
        raise ValueError(
            f"API 返回了空内容。finish_reason={response.choices[0].finish_reason}"
        )
    return content
