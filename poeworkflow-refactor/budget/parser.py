"""Budget parsing helpers."""

import re
from typing import Optional

def parse_annual_budget_usd(raw_budget: Optional[str]) -> Optional[float]:
    """把 UI 里的年预估消耗解析成 USD 数字；无法可靠解析时返回 None。"""
    if raw_budget is None:
        return None
    text = str(raw_budget).strip()
    if not text:
        return None

    normalized = (
        text.lower()
        .replace(",", "")
        .replace("$", "")
        .replace("usd", "")
        .replace("美元", "")
        .replace("美金", "")
        .replace("预估", "")
        .replace("年消耗", "")
        .replace("年度", "")
        .replace("每年", "")
        .replace("+", "")
        .strip()
    )
    if not normalized or normalized in {"na", "n/a", "none", "null", "未填写", "无"}:
        return None

    multiplier = 1.0
    if "万" in normalized or re.search(r"\d\s*w\b", normalized):
        multiplier = 10_000.0
        normalized = normalized.replace("万", "").replace("w", "")
    elif "million" in normalized:
        multiplier = 1_000_000.0
        normalized = normalized.replace("million", "")
    elif re.search(r"\d\s*m\b", normalized):
        multiplier = 1_000_000.0
        normalized = re.sub(r"\bm\b", "", normalized)
    elif re.search(r"\d\s*k\b", normalized):
        multiplier = 1_000.0
        normalized = re.sub(r"\bk\b", "", normalized)

    match = re.search(r"(\d+(?:\.\d+)?)", normalized)
    if not match:
        return None
    return float(match.group(1)) * multiplier



def _format_usd(value: Optional[float]) -> str:
    if value is None:
        return "未填写"
    return f"${value:,.2f}"
