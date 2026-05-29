"""Budget tier selection and CSV template helpers."""

import csv
import datetime
import hashlib
import io
import json
import os
import re
from typing import Any, Callable, Dict, List, Optional

from budget.parser import _format_usd
from config import BUDGET_TIERS, BUILTIN_CSV_PATH, TIER_CACHE_PATH

TIER_CACHE_SCHEMA_VERSION = 3

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


def ensure_csv_disk_columns(csv_text: str) -> str:
    """Ensure every imported server row has at least one assessable disk."""
    if not csv_text.strip():
        return csv_text

    rows = list(csv.reader(io.StringIO(csv_text)))
    if not rows:
        return csv_text

    header = rows[0]
    normalized_header = [str(col).lstrip("\ufeff").strip().lower() for col in header]

    def _index(column_name: str) -> Optional[int]:
        target = column_name.strip().lower()
        try:
            return normalized_header.index(target)
        except ValueError:
            return None

    defaults = {
        "Boot type": "BIOS",
        "Number of disks": "1",
        "Storage in use (In GB)": "64",
        "Disk 1 size (In GB)": "128",
        "Disk 1 read throughput (MB per second)": "200",
        "Disk 1 write throughput (MB per second)": "200",
        "Disk 1 read ops (operations per second)": "200",
        "Disk 1 write ops (operations per second)": "299",
    }
    indexes = {name: _index(name) for name in defaults}
    number_of_disks_idx = indexes.get("Number of disks")
    disk_size_idx = indexes.get("Disk 1 size (In GB)")
    if number_of_disks_idx is None or disk_size_idx is None:
        return csv_text

    def _positive_number(value: str) -> bool:
        try:
            return float(str(value or "").strip()) > 0
        except ValueError:
            return False

    for row in rows[1:]:
        if not any(str(cell).strip() for cell in row):
            continue
        while len(row) < len(header):
            row.append("")
        has_assessable_disk = _positive_number(row[number_of_disks_idx]) and _positive_number(row[disk_size_idx])
        if has_assessable_disk:
            continue
        for column_name, default_value in defaults.items():
            idx = indexes.get(column_name)
            if idx is None:
                continue
            current_value = str(row[idx]).strip() if idx < len(row) else ""
            should_fill = not current_value
            if column_name != "Boot type" and not _positive_number(current_value):
                should_fill = True
            if should_fill:
                while len(row) <= idx:
                    row.append("")
                row[idx] = default_value

    out = io.StringIO()
    csv.writer(out, lineterminator="\n").writerows(rows)
    return out.getvalue().rstrip("\n")


def _csv_template_hash() -> str:
    try:
        with open(BUILTIN_CSV_PATH, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()[:12]
    except Exception:
        return "unknown"


def snap_budget_to_tier(annual_budget: float) -> int:
    """将用户年预算匹配到不超过预算的最大档位。预算低于最低档位时使用最低档。"""
    if annual_budget is None or annual_budget <= 0:
        return BUDGET_TIERS[0]
    matched = None
    for tier in BUDGET_TIERS:
        if annual_budget >= tier:
            matched = tier
    if matched is None:
        return BUDGET_TIERS[0]
    return matched


def budget_target_range(annual_budget: Optional[float]) -> Dict[str, Any]:
    """返回预算目标区间：用户输入值 <= 年化估算 < 下一档；最高档允许 +20%。"""
    if annual_budget is None or annual_budget <= 0:
        return {
            "tier": BUDGET_TIERS[0],
            "target_min": None,
            "target_max": None,
            "target_mid": None,
            "target_max_exclusive": False,
        }

    tier = snap_budget_to_tier(annual_budget)
    tier_idx = BUDGET_TIERS.index(tier)
    target_min = float(annual_budget)
    if tier_idx < len(BUDGET_TIERS) - 1:
        target_max = float(BUDGET_TIERS[tier_idx + 1])
        target_max_exclusive = True
    else:
        target_max = target_min * 1.2
        target_max_exclusive = False

    return {
        "tier": tier,
        "target_min": target_min,
        "target_max": target_max,
        "target_mid": (target_min + target_max) / 2,
        "target_max_exclusive": target_max_exclusive,
    }


def tier_target_range(tier: int) -> Dict[str, Any]:
    """返回某个规模档位自身的学习目标区间。"""
    tier_idx = BUDGET_TIERS.index(tier) if tier in BUDGET_TIERS else 0
    target_min = float(tier)
    if tier_idx < len(BUDGET_TIERS) - 1:
        target_max = float(BUDGET_TIERS[tier_idx + 1])
        target_max_exclusive = True
    else:
        target_max = target_min * 1.2
        target_max_exclusive = False
    return {
        "tier": tier,
        "target_min": target_min,
        "target_max": target_max,
        "target_mid": (target_min + target_max) / 2,
        "target_max_exclusive": target_max_exclusive,
    }


def format_budget_target_range(range_info: Dict[str, Any]) -> str:
    target_min = range_info.get("target_min")
    target_max = range_info.get("target_max")
    if target_min is None or target_max is None:
        return "未填写"
    max_label = _format_usd(float(target_max))
    if range_info.get("target_max_exclusive"):
        max_label = f"<{max_label}"
    return f"{_format_usd(float(target_min))} - {max_label}"


def load_tier_cache() -> Dict[str, Any]:
    if not os.path.exists(TIER_CACHE_PATH):
        return {}
    try:
        with open(TIER_CACHE_PATH, "r", encoding="utf-8") as f:
            cache = json.load(f)
        if cache.get("schema_version") != TIER_CACHE_SCHEMA_VERSION:
            return {}
        if cache.get("budget_tiers") != BUDGET_TIERS:
            return {}
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
    cache["schema_version"] = TIER_CACHE_SCHEMA_VERSION
    cache["budget_tiers"] = BUDGET_TIERS
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
        range_info = tier_target_range(tier)
        target_monthly = float(range_info["target_min"]) / 12
        target_max_monthly = float(range_info["target_max"]) / 12
        target_max_exclusive = bool(range_info["target_max_exclusive"])

        def _within_upper(value: float) -> bool:
            return value < target_max_monthly if target_max_exclusive else value <= target_max_monthly

        if total_annual >= tier and _within_upper(total_annual / 12):
            selected = list(machine_costs)
        else:
            selected: List[Dict[str, Any]] = []
            running = 0.0
            for mc in machine_costs:
                if running >= target_monthly:
                    break
                if _within_upper(running + mc["monthly_cost"]):
                    selected.append(mc)
                    running += mc["monthly_cost"]

            if sum(s["monthly_cost"] for s in selected) < target_monthly:
                remaining = [mc for mc in machine_costs if mc not in selected]
                running = sum(s["monthly_cost"] for s in selected)
                for mc in remaining:
                    if _within_upper(running + mc["monthly_cost"]):
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
            f"目标区间 {format_budget_target_range(range_info)}，"
            f"预期年化 {_format_usd(sel_monthly * 12)}"
        )

    return cache


def get_machine_ids_for_tier(
    tier: int,
    machines: List[Dict[str, Any]],
    account_prefix: str,
    cache: Dict[str, Any],
    customer_name: str = "",
) -> List[str]:
    tier_data = cache.get("tiers", {}).get(str(tier))
    if not tier_data:
        return [m.get("id") for m in machines if m.get("id")]

    selected_template_names = {n.lower() for n in tier_data["machine_names"]}

    # 先尝试按名称匹配
    selected_ids: List[str] = []
    unselected_ids: List[str] = []
    for m in machines:
        display_name = m.get("properties", {}).get("displayName", "")
        original_name = _strip_account_prefix(display_name, account_prefix)
        mid = m.get("id")
        if not mid:
            continue
        if original_name.lower() in selected_template_names:
            selected_ids.append(mid)
        else:
            unselected_ids.append(mid)

    # 如果名称匹配失败（例如缓存来自旧命名方案），按缓存的机器数量选取
    expected_count = tier_data.get("machine_count", len(selected_template_names))
    if len(selected_ids) < expected_count:
        all_ids = [m.get("id") for m in machines if m.get("id")]
        selected_ids = all_ids[:expected_count]
        unselected_ids = all_ids[expected_count:]

    # ── 基于客户名的服务器数量变化（±1~3 台） ──
    # 旧方案只替换同数量服务器；若模板规格相同，价格不会变化。
    # 新方案直接增减数量，使总价按台数产生确定性差异。
    if customer_name and len(selected_ids) >= 4:
        import random as _rng
        seed = int(hashlib.md5(customer_name.encode()).hexdigest(), 16) % (2**32)
        rng = _rng.Random(seed)
        max_delta = min(3, max(1, len(selected_ids) // 8))
        candidates = [d for d in range(-max_delta, max_delta + 1) if d != 0]
        delta = rng.choice(candidates)

        if delta > 0 and unselected_ids:
            add_count = min(delta, len(unselected_ids))
            selected_ids.extend(rng.sample(unselected_ids, add_count))
        elif delta < 0:
            remove_count = min(abs(delta), len(selected_ids) - 1)
            for idx in sorted(rng.sample(range(len(selected_ids)), remove_count), reverse=True):
                selected_ids.pop(idx)

    return selected_ids


def jitter_csv_performance(csv_text: str, customer_name: str) -> str:
    """基于客户名对 CSV 做确定性差异化。

    v2 差异化策略：
    - 少量 VM 的核心数/内存成比例变化，跨越 Azure VM SKU 推荐边界；
    - 部分磁盘强制跳到相邻 Azure 磁盘定价层；
    - 性能指标提升到 ±25% 抖动。
    同一客户多次调用结果一致，不同客户尽量产生不同评估金额。
    """
    import random as _rng

    if not customer_name or not csv_text.strip():
        return csv_text

    seed = int(hashlib.md5(customer_name.encode()).hexdigest(), 16) % (2**32)
    rng = _rng.Random(seed)

    rows = list(csv.reader(io.StringIO(csv_text)))
    if len(rows) < 2:
        return csv_text

    header = rows[0]
    normalized_header = [str(col).lstrip("\ufeff").strip().lower() for col in header]

    def _find_index(*needles: str) -> Optional[int]:
        lowered = [n.lower() for n in needles]
        for idx, col_name in enumerate(normalized_header):
            if all(n in col_name for n in lowered):
                return idx
        return None

    cores_idx = _find_index("cores")
    memory_idx = _find_index("memory")
    disk_size_indices = [
        idx for idx, col_name in enumerate(normalized_header)
        if "disk" in col_name and "size" in col_name and "gb" in col_name
    ]
    storage_idx = _find_index("storage in use")

    data_rows = [row for row in rows[1:] if any(str(cell).strip() for cell in row)]

    # ── 规格强变异：只改少量 VM，避免看起来不合理，但足以跨 SKU 边界 ──
    if data_rows and cores_idx is not None and memory_idx is not None:
        spec_change_count = min(len(data_rows), rng.randint(2, 4))
        for row in rng.sample(data_rows, spec_change_count):
            while len(row) < len(header):
                row.append("")
            try:
                cores = int(float(str(row[cores_idx]).strip() or "0"))
                memory = int(float(str(row[memory_idx]).strip() or "0"))
            except ValueError:
                continue
            if cores <= 0 or memory <= 0:
                continue

            # 在常见规格间跳跃，强制跨越 D2/D4/D8 推荐边界。
            if cores <= 2:
                new_cores = 4 if rng.random() < 0.75 else 1
            elif cores <= 4:
                new_cores = 8 if rng.random() < 0.55 else 2
            else:
                new_cores = 4 if rng.random() < 0.45 else min(16, cores * 2)
            new_cores = max(1, min(16, new_cores))
            row[cores_idx] = str(new_cores)
            # 保持原内存/核心比例，最低 1GB/核，并对齐到 1024MB。
            mb_per_core = max(1024, memory // max(cores, 1))
            row[memory_idx] = str(max(1024, new_cores * mb_per_core))

    # ── 磁盘跨层：推动部分磁盘跨 P6/P10/P15/P20 等定价边界 ──
    disk_tiers = [32, 64, 128, 256, 512, 1024, 2048, 4096]

    def _neighbor_disk_size(value: float) -> int:
        nearest_idx = min(range(len(disk_tiers)), key=lambda i: abs(disk_tiers[i] - value))
        step = rng.choice([-1, 1])
        target_idx = max(0, min(len(disk_tiers) - 1, nearest_idx + step))
        if target_idx == nearest_idx and len(disk_tiers) > 1:
            target_idx = max(0, nearest_idx - 1)
        return disk_tiers[target_idx]

    if data_rows and disk_size_indices:
        disk_change_count = min(len(data_rows), rng.randint(3, 6))
        for row in rng.sample(data_rows, disk_change_count):
            while len(row) < len(header):
                row.append("")
            for idx in disk_size_indices:
                raw = str(row[idx]).strip()
                if not raw:
                    continue
                try:
                    val = float(raw)
                except ValueError:
                    continue
                if val <= 0:
                    continue
                new_disk_size = _neighbor_disk_size(val)
                row[idx] = str(new_disk_size)
                if storage_idx is not None and storage_idx < len(row):
                    try:
                        used = float(str(row[storage_idx]).strip() or "0")
                    except ValueError:
                        used = 0
                    if used > 0:
                        row[storage_idx] = str(int(max(16, min(new_disk_size, round(new_disk_size * rng.uniform(0.45, 0.8))))))

    # col_index -> (min_val, max_val, jitter_pct, snap_to_multiple)
    perf_columns: Dict[int, tuple] = {}
    for ci, col_name in enumerate(normalized_header):
        if "cpu utilization" in col_name:
            perf_columns[ci] = (10, 90, 0.25, None)
        elif "memory utilization" in col_name:
            perf_columns[ci] = (10, 90, 0.25, None)
        elif "disk" in col_name and "size" in col_name and "gb" in col_name:
            # 磁盘大小已在上方跨层处理；这里只对未被抽中的行做轻微兜底抖动。
            perf_columns[ci] = (32, 4096, 0.10, 8)
        elif "storage in use" in col_name:
            perf_columns[ci] = (16, 4096, 0.20, 8)
        elif "read throughput" in col_name or "write throughput" in col_name:
            perf_columns[ci] = (10, 9999, 0.25, None)
        elif "read ops" in col_name or "write ops" in col_name:
            perf_columns[ci] = (10, 9999, 0.25, None)
        elif "network in throughput" in col_name or "network out throughput" in col_name:
            perf_columns[ci] = (5, 9999, 0.25, None)

    if not perf_columns and cores_idx is None and not disk_size_indices:
        return csv_text

    for row in rows[1:]:
        if not any(str(cell).strip() for cell in row):
            continue
        for ci, (min_val, max_val, jitter_pct, snap) in perf_columns.items():
            if ci >= len(row):
                continue
            raw = str(row[ci]).strip()
            if not raw:
                continue
            try:
                val = float(raw)
            except ValueError:
                continue
            if val <= 0:
                continue
            factor = 1.0 + rng.uniform(-jitter_pct, jitter_pct)
            new_val = val * factor
            new_val = max(min_val, min(max_val, new_val))
            if snap:
                new_val = max(snap, round(new_val / snap) * snap)
            if "." not in raw:
                row[ci] = str(int(round(new_val)))
            else:
                row[ci] = f"{new_val:.2f}"

    out = io.StringIO()
    csv.writer(out, lineterminator="\n").writerows(rows)
    return out.getvalue().rstrip("\n")
