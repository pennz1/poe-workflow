"""Budget tier selection and CSV template helpers."""

import datetime
import hashlib
import json
import os
import re
from typing import Any, Callable, Dict, List

from budget.parser import _format_usd
from config import BUDGET_TIERS, BUILTIN_CSV_PATH, TIER_CACHE_PATH

def load_builtin_csv_template() -> str:
    with open(BUILTIN_CSV_PATH, "r", encoding="utf-8-sig") as f:
        return f.read()


def _get_template_machine_names() -> List[str]:
    csv_text = load_builtin_csv_template()
    names: List[str] = []
    for line in csv_text.strip().split("\n")[1:]:
        if not line.strip():
            continue
        name = line.split(",", 1)[0].strip()
        if name:
            names.append(name)
    return names


def _safe_csv_prefix(account_name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9\-]", "-", (account_name or "customer").strip()).strip("-") or "customer"


def prefix_csv_server_names(csv_text: str, prefix: str) -> str:
    """给 CSV 中的服务器名添加客户前缀，并将序号随机化以避免多人导入时序号冲突。"""
    import random as _rng
    lines = csv_text.strip().split("\n")
    if not lines:
        return csv_text

    # 使用 prefix 作为随机种子，确保同一客户每次生成的序号相同（tier cache 可匹配）
    seed = int(hashlib.md5(prefix.encode()).hexdigest(), 16) % (2**32)
    rng = _rng.Random(seed)

    # 收集所有数字后缀并生成随机映射
    num_pattern = re.compile(r"(\d+)")
    # 找出所有行中使用的数字，生成一个不重复的随机序号池
    all_numbers = set()
    for line in lines[1:]:
        if not line.strip():
            continue
        name = line.split(",", 1)[0].strip()
        for m in num_pattern.finditer(name):
            all_numbers.add(int(m.group()))

    # 生成随机映射：原始序号 → 随机序号（范围扩大避免冲突）
    max_num = max(all_numbers) if all_numbers else 50
    pool_start = rng.randint(100, 800)
    number_map: Dict[int, int] = {}
    used_numbers = set()
    for orig_num in sorted(all_numbers):
        new_num = pool_start + rng.randint(1, 5)
        while new_num in used_numbers:
            new_num += rng.randint(1, 3)
        number_map[orig_num] = new_num
        used_numbers.add(new_num)
        pool_start = new_num

    result = [lines[0]]
    for line in lines[1:]:
        if not line.strip():
            continue
        parts = line.split(",", 1)
        if len(parts) >= 2:
            original_name = parts[0].strip()
            # 替换名称中的数字为随机化后的数字
            def _replace_num(m):
                orig = int(m.group())
                return str(number_map.get(orig, orig))
            randomized_name = num_pattern.sub(_replace_num, original_name)
            result.append(f"{prefix}-{randomized_name},{parts[1]}")
        else:
            result.append(line)
    return "\n".join(result)


def _csv_template_hash() -> str:
    try:
        with open(BUILTIN_CSV_PATH, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()[:12]
    except Exception:
        return "unknown"


def snap_budget_to_tier(annual_budget: float) -> int:
    if annual_budget is None or annual_budget <= 0:
        return BUDGET_TIERS[-1]
    for tier in BUDGET_TIERS:
        if annual_budget <= tier * 1.15:
            return tier
    return BUDGET_TIERS[-1]


def load_tier_cache() -> Dict[str, Any]:
    if not os.path.exists(TIER_CACHE_PATH):
        return {}
    try:
        with open(TIER_CACHE_PATH, "r", encoding="utf-8") as f:
            cache = json.load(f)
        if cache.get("template_hash") != _csv_template_hash():
            return {}
        created = cache.get("created_at", "")
        if created:
            created_dt = datetime.datetime.fromisoformat(created)
            if (datetime.datetime.now() - created_dt).days > 7:
                return {}
        return cache
    except Exception:
        return {}


def save_tier_cache(cache: Dict[str, Any]) -> None:
    cache["created_at"] = datetime.datetime.now().isoformat()
    cache["template_hash"] = _csv_template_hash()
    with open(TIER_CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2, ensure_ascii=False)



def _assessed_machine_monthly_cost(machine: Dict[str, Any]) -> float:
    props = machine.get("properties", {})
    cost = 0.0
    for key in ("monthlyComputeCostForRecommendedSize", "monthlyStorageCost", "monthlyBandwidthCost"):
        try:
            cost += float(props.get(key, 0) or 0)
        except (TypeError, ValueError):
            pass
    for comp in props.get("costComponents") or []:
        if str(comp.get("name", "")).lower() == "monthlysecuritycost":
            try:
                cost += float(comp.get("value", 0) or 0)
            except (TypeError, ValueError):
                pass
    return cost


def _strip_account_prefix(display_name: str, prefix: str) -> str:
    if prefix and display_name.lower().startswith(prefix.lower() + "-"):
        return display_name[len(prefix) + 1:]
    return display_name


def learn_tier_machine_selections(
    assessed_machines: List[Dict[str, Any]],
    account_prefix: str,
    progress: Callable[[str], None],
) -> Dict[str, Any]:
    machine_costs: List[Dict[str, Any]] = []
    for m in assessed_machines:
        display_name = m.get("properties", {}).get("displayName", "")
        monthly_cost = _assessed_machine_monthly_cost(m)
        original_name = _strip_account_prefix(display_name, account_prefix)
        machine_costs.append({
            "template_name": original_name,
            "monthly_cost": monthly_cost,
        })

    total_monthly = sum(mc["monthly_cost"] for mc in machine_costs)
    total_annual = total_monthly * 12
    progress(
        f"全量评估学习基准：{len(machine_costs)} 台服务器，"
        f"年化 {_format_usd(total_annual)}"
    )

    machine_costs.sort(key=lambda x: x["monthly_cost"])

    cache: Dict[str, Any] = {
        "total_monthly": total_monthly,
        "total_annual": total_annual,
        "machine_count": len(machine_costs),
        "tiers": {},
    }

    for tier in BUDGET_TIERS:
        target_monthly = tier / 12
        target_max_monthly = tier * 1.2 / 12

        if total_annual <= tier * 1.2:
            selected = list(machine_costs)
        else:
            selected: List[Dict[str, Any]] = []
            running = 0.0
            for mc in machine_costs:
                if running >= target_monthly:
                    break
                if running + mc["monthly_cost"] <= target_max_monthly:
                    selected.append(mc)
                    running += mc["monthly_cost"]

            if sum(s["monthly_cost"] for s in selected) < target_monthly:
                remaining = [mc for mc in machine_costs if mc not in selected]
                running = sum(s["monthly_cost"] for s in selected)
                for mc in remaining:
                    if running + mc["monthly_cost"] > target_max_monthly:
                        break
                    selected.append(mc)
                    running += mc["monthly_cost"]
                    if running >= target_monthly:
                        break

        sel_monthly = sum(s["monthly_cost"] for s in selected)
        sel_names = [s["template_name"] for s in selected]

        cache["tiers"][str(tier)] = {
            "machine_names": sel_names,
            "machine_count": len(sel_names),
            "expected_monthly": round(sel_monthly, 2),
            "expected_annual": round(sel_monthly * 12, 2),
        }
        progress(
            f"  规模 {_format_usd(float(tier))}：选择 {len(sel_names)}/{len(machine_costs)} 台，"
            f"预期年化 {_format_usd(sel_monthly * 12)}"
        )

    return cache


def get_machine_ids_for_tier(
    tier: int,
    machines: List[Dict[str, Any]],
    account_prefix: str,
    cache: Dict[str, Any],
) -> List[str]:
    tier_data = cache.get("tiers", {}).get(str(tier))
    if not tier_data:
        return [m.get("id") for m in machines if m.get("id")]

    selected_template_names = {n.lower() for n in tier_data["machine_names"]}

    # 先尝试按名称匹配
    selected_ids: List[str] = []
    for m in machines:
        display_name = m.get("properties", {}).get("displayName", "")
        original_name = _strip_account_prefix(display_name, account_prefix)
        if original_name.lower() in selected_template_names:
            mid = m.get("id")
            if mid:
                selected_ids.append(mid)

    # 如果名称匹配失败（例如缓存来自旧命名方案），按缓存的机器数量选取
    expected_count = tier_data.get("machine_count", len(selected_template_names))
    if len(selected_ids) < expected_count:
        all_ids = [m.get("id") for m in machines if m.get("id")]
        selected_ids = all_ids[:expected_count]

    return selected_ids
