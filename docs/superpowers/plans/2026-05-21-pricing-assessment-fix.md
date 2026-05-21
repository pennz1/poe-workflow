# POE Price & Assessment 精度修复 — 实现计划

> **面向 AI 代理的工作者：** 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。

**目标：** 修复价格计算器导出含无关 VM、Migrate 评估 250K+ 年化超 20%、Assessment Excel 时间戳偏移三个问题。

**架构：** 三个独立修复，修复 2+3 改同一文件但不同函数。

**技术栈：** Python, Playwright, openpyxl

---

### 任务 1：移除 VM filler，改用方案资源循环增量

**文件：** 修改 `pricing_automation.py`

- [ ] **步骤 1：删除 BUDGET_FILLER_SERVICES 列表**

删除第 425-436 行的常量定义：
```python
# 删除整个 BUDGET_FILLER_SERVICES = [...] 块
```

- [ ] **步骤 2：重写 _calibrate_budget 函数（行 1316-1380）**

将 `resources` 参数传入函数，改为循环方案文档已有资源来增量。新实现：

```python
async def _calibrate_budget(
    automation: PricingCalculatorAutomation,
    resources: List[ResourceSpec],       # 新增参数：方案文档的资源列表
    annual_budget: float,
    current_annual: float,
    progress: Callable[[str], None],
    max_rounds: int = 15,
    budget_cap: float = 0,
) -> float:
    page = automation._page
    consecutive_failures = 0
    effective_cap = budget_cap if budget_cap > 0 else float("inf")

    for round_num in range(max_rounds):
        if current_annual >= annual_budget:
            progress(f"  ✅ 预算校准完成: ${current_annual:,.0f} >= ${annual_budget:,.0f}")
            return current_annual

        if current_annual >= effective_cap:
            progress(f"  ⚠️ 已达预算上限 ${effective_cap:,.0f}，停止补价")
            return current_annual

        if consecutive_failures >= 3:
            progress(f"  ⚠️ 连续{consecutive_failures}轮无效，停止补价")
            break

        # 循环使用方案文档中的已有资源
        res = resources[round_num % len(resources)]
        normalized_name = normalize_service_name(res.service_name)
        catalog = SERVICE_CATALOG.get(normalized_name)
        if not catalog:
            consecutive_failures += 1
            progress(f"  ⚠️ 跳过 {res.service_name}: 未在 SERVICE_CATALOG 中")
            continue
        search_term = catalog["search_term"]

        progress(f"  补量 Round {round_num + 1}: 增加 1 个 {normalized_name} (差额 ${annual_budget - current_annual:,.0f})...")
        added = await automation.add_service(search_term)
        if added:
            region = resolve_region(res.region)
            await automation.configure_service_region(region)
            if "openai" in normalized_name.lower():
                await automation.configure_openai_tokens(input_tokens=1000, output_tokens=500)
            await page.wait_for_timeout(3000)
            monthly = await automation.get_current_total()
            new_annual = monthly * 12
            if new_annual > current_annual:
                current_annual = new_annual
                consecutive_failures = 0
                progress(f"  补量后年化: ${current_annual:,.0f}")
            else:
                consecutive_failures += 1
        else:
            consecutive_failures += 1

    if current_annual < annual_budget:
        progress(f"  ⚠️ 已达最大补价轮次，当前年化 ${current_annual:,.0f}")
    return current_annual
```

- [ ] **步骤 3：更新 _add_resources_and_calibrate 的调用**

在 `_add_resources_and_calibrate()` 中将 `_calibrate_budget` 调用改为传入 `resources`：

```python
# 原来的调用（约 1289 行）：
annual_total = await _calibrate_budget(
    automation, annual_budget, annual_total, progress,
    budget_cap=budget_cap,
)
# 改为：
annual_total = await _calibrate_budget(
    automation, resources, annual_budget, annual_total, progress,
    budget_cap=budget_cap,
)
```

- [ ] **步骤 4：Commit**

```bash
git add pricing_automation.py
git commit -m "fix: remove VM budget fillers, use existing resources for pricing calibration"
```

---

### 任务 2：250K+ Migrate 评估年化上限 150% → 120%

**文件：** 修改 `poeworkflow-refactor/azure/migrate.py` 第 704 行

- [ ] **步骤 1：修改一行代码**

```python
# 原（第 704 行）：
target_max = annual_budget * 1.5  # 最大档次无上限约束，放宽到 150%
# 改为：
target_max = annual_budget * 1.2  # 所有档位统一上限 120%
```

- [ ] **步骤 2：Commit**

```bash
git add poeworkflow-refactor/azure/migrate.py
git commit -m "fix: cap Migrate assessment annual cost at 120% for all tiers"
```

---

### 任务 3：Assessment Excel 时间戳修正

**文件：** 修改 `poeworkflow-refactor/azure/migrate.py` 的 `fix_assessment_excel_timestamps()`

- [ ] **步骤 1：重写时间戳逻辑**

将函数中的日期计算部分改为新逻辑：

```python
def fix_assessment_excel_timestamps(
    excel_bytes: bytes,
    pov_start: datetime.date,
    pov_end: datetime.date,
) -> bytes:
    """修改 Assessment Excel 报告中的时间戳：
    - Performance history end time = Created on (UTC) 同一天，仅日期
    - Performance history start time = end 前一天，仅日期
    - 两者均在 POV 区间内
    """
    import random
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(excel_bytes))

    # 在 POV 区间内随机选一天作为评估创建日
    total_days = (pov_end - pov_start).days
    if total_days <= 0:
        total_days = 1
    random_day_offset = random.randint(1, max(total_days - 1, 1))
    created_date = pov_start + datetime.timedelta(days=random_day_offset)

    # Performance history end = created_date（同一天）
    perf_end_date = created_date
    # Performance history start = 前一天
    perf_start_date = created_date - datetime.timedelta(days=1)

    # 边界保护：确保 start 不早于 POV 开始
    if perf_start_date < pov_start:
        perf_start_date = pov_start
        perf_end_date = pov_start + datetime.timedelta(days=1)

    def _format_date_only(dt_val: datetime.date) -> str:
        return f"{dt_val.month}/{dt_val.day}/{dt_val.year}"

    def _format_as_text(dt_val: datetime.datetime) -> str:
        hour = dt_val.hour
        ampm = "AM" if hour < 12 else "PM"
        hour_12 = hour % 12
        if hour_12 == 0:
            hour_12 = 12
        return f"{dt_val.month}/{dt_val.day}/{dt_val.year} {hour_12}:{dt_val.minute:02d}:{dt_val.second:02d} {ampm}"

    def _parse_time_from_value(orig):
        if isinstance(orig, datetime.datetime):
            return orig.hour, orig.minute, orig.second
        text = str(orig).strip()
        m = re.search(r'(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)?', text, re.IGNORECASE)
        if m:
            h, mi, s = int(m.group(1)), int(m.group(2)), int(m.group(3))
            ampm = (m.group(4) or "").upper()
            if ampm == "PM" and h != 12:
                h += 12
            elif ampm == "AM" and h == 12:
                h = 0
            return h, mi, s
        return 2, 35, 35

    created_datetime = None

    # ── 修改 Assessment_Summary: Created on (UTC) ──
    if "Assessment_Summary" in wb.sheetnames:
        ws = wb["Assessment_Summary"]
        created_col = None
        for col in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=1, column=col).value
            if cell_val and "created on" in str(cell_val).lower():
                created_col = col
                break
        if created_col:
            for row in range(2, ws.max_row + 1):
                orig = ws.cell(row=row, column=created_col).value
                if orig is not None:
                    h, mi, s = _parse_time_from_value(orig)
                    created_datetime = datetime.datetime(
                        created_date.year, created_date.month, created_date.day, h, mi, s
                    )
                    ws.cell(row=row, column=created_col).value = _format_as_text(created_datetime)

    if created_datetime is None:
        created_datetime = datetime.datetime(
            created_date.year, created_date.month, created_date.day, 2, 35, 35
        )

    # ── 修改 Assessment_Properties: start/end 纯日期 ──
    if "Assessment_Properties" in wb.sheetnames:
        ws = wb["Assessment_Properties"]
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                prop_name = str(ws.cell(row=row, column=col).value or "").lower()
                if "performance history start" in prop_name:
                    val_col = col + 1 if col + 1 <= ws.max_column else col
                    ws.cell(row=row, column=val_col).value = _format_date_only(perf_start_date)
                elif "performance history end" in prop_name:
                    val_col = col + 1 if col + 1 <= ws.max_column else col
                    ws.cell(row=row, column=val_col).value = _format_date_only(perf_end_date)

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()
```

- [ ] **步骤 2：确认原文件 `poeworkflow-refactor/azure/migrate.py` 的行 1500-1579 被替换为以上代码**

- [ ] **步骤 3：Commit**

```bash
git add poeworkflow-refactor/azure/migrate.py
git commit -m "fix: align Assessment Excel timestamps — end=created date, start=day before"
```

---

### 任务 4：回归验证

- [ ] **步骤 1：验证 import 不变**

```bash
cd "/Users/penn/Documents/POE workflow"
.venv/bin/python -c 'import sys; sys.path.insert(0,"poeworkflow-refactor"); from azure.migrate import fix_assessment_excel_timestamps; print("IMPORT OK")'
```

- [ ] **步骤 2：运行回归测试**

```bash
cd "/Users/penn/Documents/POE workflow"
.venv/bin/python -m pytest test_tier_algorithm.py test_new_features.py test_pricing_automation.py -v 2>&1 | tail -10
```

- [ ] **步骤 3：Commit**

```bash
git commit -m "chore: verification passed for pricing-assessment fixes" --allow-empty
```
