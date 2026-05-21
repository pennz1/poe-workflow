# POE Price & Assessment 精度修复

## 目标

修复全自动 POE 流程中三个数据精度问题，使导出的价格估算表和评估报告与解决方案文档保持一致。

---

## 修复 1：价格计算器导出含无关 VM

**问题：** 导出的 xlsx 中出现 Virtual Machines，但方案文档的资源需求表里没有 VM。

**根因：** `_calibrate_budget()` 在 Calculator 总价低于用户预算时，用硬编码的 `BUDGET_FILLER_SERVICES`（10 个 VM 规格）补差价。

**修复：**
- 移除 `BUDGET_FILLER_SERVICES` 列表
- `_calibrate_budget()` 改为循环方案文档中已有的资源来增量
- 每轮从资源列表中取一个资源，重复 `add_service(search_term)` 加 1 个实例
- 直到 `annual_total >= annual_budget` 或达 budget_cap

**影响范围：** `pricing_automation.py`
- 删除 `BUDGET_FILLER_SERVICES`（行 425-436）
- 重写 `_calibrate_budget()`（行 1316-1380）：移除 filler 逻辑，改用已有 resources 循环

**伪代码：**
```python
async def _calibrate_budget(automation, resources, annual_budget, ...):
    for round_num in range(max_rounds):
        if current_annual >= annual_budget or current_annual >= effective_cap:
            break
        # 循环使用方案文档中的资源
        res = resources[round_num % len(resources)]
        search_term = SERVICE_CATALOG[normalize(res.service_name)]["search_term"]
        await automation.add_service(search_term)  # 加 1 个实例
        # ... 更新 current_annual
```

---

## 修复 2：250K+ 客户 Migrate 评估年化不超 20%

**问题：** 最高档位（≥250K）的目标上限是 `annual_budget * 1.5`（150%），过于宽松。

**修复：** 所有档位统一上限 120%。

**影响范围：** `poeworkflow-refactor/azure/migrate.py` 的 `tune_assessment_to_budget()`
- 第 704 行：`annual_budget * 1.5` → `annual_budget * 1.2`

---

## 修复 3：Assessment Excel 时间戳修正

**问题：** 当前逻辑 `end = start + 1天`，日期关系不对。

**要求：**
- Performance history **end** time：与 Created on (UTC) **同一天**，只保留日期不保留时分秒
- Performance history **start** time：end 往前推 1 天
- 两者都在 POV 项目时间段内

**修复：** `fix_assessment_excel_timestamps()` 逻辑改为：
1. 在 POV 区间内随机选一天作为评估创建日
2. Created on (UTC)：该日期 + 原时分秒
3. Performance history end：该日期，`0:00:00 AM`
4. Performance history start：该日期 - 1 天，`0:00:00 AM`

**影响范围：** `poeworkflow-refactor/azure/migrate.py` 的 `fix_assessment_excel_timestamps()`

**伪代码：**
```python
created_date = pov_start + timedelta(days=random_day_offset)  # 随机选一天
created_datetime = datetime(created_date.year, created_date.month, created_date.day, h, mi, s)

# Performance history end = 同一天，日期格式无时分秒
perf_end_date = created_date  # 同一天
# Performance history start = 前一天
perf_start_date = created_date - timedelta(days=1)

# 两者都要在 POV 区间内：如果 start < pov_start，用 pov_start
if perf_start_date < pov_start:
    perf_start_date = pov_start
    perf_end_date = pov_start + timedelta(days=1)
```

---

## 实施顺序

三个修复互不依赖，可并行：
- 修复 1：`pricing_automation.py`（项目根目录，非 refactor 目录）
- 修复 2：`poeworkflow-refactor/azure/migrate.py`
- 修复 3：`poeworkflow-refactor/azure/migrate.py`

修复 2 和 3 改同一文件，先做修复 2 再做修复 3 避免冲突。
