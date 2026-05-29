实现 POE Price & Assessment 三个精度修复。计划文档位于 docs/superpowers/plans/2026-05-21-pricing-assessment-fix.md。

## 修复 1：移除 BUDGET_FILLER_SERVICES，改用方案资源循环增量

文件：pricing_automation.py

步骤 1：删除第 421-436 行的 BUDGET_FILLER_SERVICES 列表（包括注释行 421-423）。

步骤 2：重写 _calibrate_budget 函数（第 1316-1380 行）。新函数签名增加 resources 参数，循环使用方案文档已有资源替代硬编码 VM filler。替换为以下代码：

```python
async def _calibrate_budget(
    automation: PricingCalculatorAutomation,
    resources: List[ResourceSpec],
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

步骤 3：在 _add_resources_and_calibrate 中（约第 1289 行），更新 _calibrate_budget 调用，传入 resources 参数：
```python
annual_total = await _calibrate_budget(
    automation, resources, annual_budget, annual_total, progress,
    budget_cap=budget_cap,
)
```

步骤 4：Commit
```bash
git add pricing_automation.py
git commit -m "fix: remove VM budget fillers, use existing resources for pricing calibration"
```

## 修复 2：250K+ Migrate 评估年化上限 150% → 120%

文件：poeworkflow-refactor/azure/migrate.py 第 704 行

将 `target_max = annual_budget * 1.5` 改为 `target_max = annual_budget * 1.2`

Commit：
```bash
git add poeworkflow-refactor/azure/migrate.py
git commit -m "fix: cap Migrate assessment annual cost at 120% for all tiers"
```

## 修复 3：Assessment Excel 时间戳修正

文件：poeworkflow-refactor/azure/migrate.py 的 fix_assessment_excel_timestamps() 函数（第 1500-1628 行）

完全替换该函数为新逻辑：
- Performance history end time 的日期（不含时分秒）与 Created on (UTC) 同一天
- Performance history start time = end 前一天（仅日期）
- 两者均在 POV 区间内

替换为以下代码：

```python
def fix_assessment_excel_timestamps(
    excel_bytes: bytes,
    pov_start: datetime.date,
    pov_end: datetime.date,
) -> bytes:
    import random
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(excel_bytes))

    total_days = (pov_end - pov_start).days
    if total_days <= 0:
        total_days = 1
    random_day_offset = random.randint(1, max(total_days - 1, 1))
    created_date = pov_start + datetime.timedelta(days=random_day_offset)

    perf_end_date = created_date
    perf_start_date = created_date - datetime.timedelta(days=1)

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

Commit：
```bash
git add poeworkflow-refactor/azure/migrate.py
git commit -m "fix: align Assessment Excel timestamps — end=created date, start=day before"
```

## 实施顺序
先做修复 2，再做修复 3（同文件避免冲突）。修复 1 可独立进行。全部完成后 report 结果。
