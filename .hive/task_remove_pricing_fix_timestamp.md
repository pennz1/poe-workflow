两个独立改动，请按顺序完成。

---

## 任务 1：彻底移除 Azure Pricing Calculator 自动化功能

**目标**：删除整个 pricing_automation 模块及其所有引用，不影响其他功能模块。

### 1.1 修改 poeworkflow-refactor/pipeline.py

删除以下内容：

**删除 import 块（第 24-35 行）**：
```python
try:
    from pricing_automation import (
        PricingExportResult,
        run_pricing_export,
    )
    HAS_PRICING_AUTOMATION = True
except ImportError:
    HAS_PRICING_AUTOMATION = False

    class PricingExportResult:
        def __init__(self):
            self.fallbacks = []
```

**删除 Step 2.5 整段（第 297-326 行）**：
从 "# ── Step 2.5: Pricing Calculator 自动导出 ──" 注释开始，到其对应的所有 pricing_artifact / pricing_result 逻辑结束为止。

**删除 zip 打包中的 pricing 引用（第 363-364 行）**：
```python
if pricing_artifact:
    all_artifacts.append(pricing_artifact)
```

**删除 result dict 中的 pricing 字段（第 373-374 行）**：
```python
"pricing": pricing_artifact,
"pricing_result": pricing_result,
```

**调整 run_full_auto_poe 函数中 Step 编号**：原 "四、Azure Migrate 评估报告" 改为 "三、Azure Migrate 评估报告"。

### 1.2 修改 poeworkflow-refactor/ui/auto_poe.py

搜索并删除所有 pricing 相关代码：
- 删除 `HAS_PRICING_AUTOMATION` 和 `PricingExportResult` 的 import（第 25-26 行）
- 删除 `from pricing_automation import is_browser_profile_ready`（第 35 行）
- 删除"初始化浏览器"按钮相关代码块（第 184-196 行附近）
- 删除结果展示中的 `pricing_file_name`、`pricing_fallbacks` 引用
- 删除 fallback 警告显示（第 350-352 行）

### 1.3 删除文件

```bash
git rm pricing_automation.py
git rm test_pricing_automation.py
```

### 1.4 验证

```bash
.venv/bin/python -m py_compile poeworkflow-refactor/pipeline.py
.venv/bin/python -m py_compile poeworkflow-refactor/ui/auto_poe.py
.venv/bin/python -m py_compile poeworkflow-refactor/main.py
grep -rn "pricing" poeworkflow-refactor/ --include="*.py" | grep -v "azurePricingTier" || echo "Pricing references cleaned"
.venv/bin/python -m pytest test_tier_algorithm.py test_new_features.py -v 2>&1 | tail -10
```

### 1.5 Commit

```bash
git add -A
git commit -m "chore: remove Azure Pricing Calculator automation entirely"
```

---

## 任务 2：修正 Assessment Excel 时间戳

**目标**：Performance history end time 必须等于 Created on (UTC) 的完整 datetime（含时分秒），start = end - 1 天。

**背景**：当前 Fix 3 的代码将 Performance history start/end 写为纯日期（_format_date_only），去掉了时分秒。实际需求是：end time = Created on (UTC) 完整 datetime（同一天同时刻），start time = end datetime - 1 天。

**文件**：poeworkflow-refactor/azure/migrate.py，函数 fix_assessment_excel_timestamps（第 1500-1588 行）

### 2.1 修改 Performance history 写入逻辑

将第 1574-1584 行替换为：

```python
    if "Assessment_Properties" in wb.sheetnames:
        ws = wb["Assessment_Properties"]
        # end datetime = Created on (UTC) 完整 datetime
        perf_end_datetime = created_datetime
        # start datetime = end datetime - 1 天
        perf_start_datetime = created_datetime - datetime.timedelta(days=1)
        # 边界保护：start 不能早于 POV 开始
        pov_start_datetime = datetime.datetime(
            pov_start.year, pov_start.month, pov_start.day, 0, 0, 0
        )
        if perf_start_datetime < pov_start_datetime:
            perf_start_datetime = pov_start_datetime
            perf_end_datetime = pov_start_datetime + datetime.timedelta(days=1)

        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                prop_name = str(ws.cell(row=row, column=col).value or "").lower()
                if "performance history start" in prop_name:
                    val_col = col + 1 if col + 1 <= ws.max_column else col
                    ws.cell(row=row, column=val_col).value = _format_as_text(perf_start_datetime)
                elif "performance history end" in prop_name:
                    val_col = col + 1 if col + 1 <= ws.max_column else col
                    ws.cell(row=row, column=val_col).value = _format_as_text(perf_end_datetime)
```

### 2.2 删除不再使用的 _format_date_only 函数

删除第 1523-1524 行的 `_format_date_only` 函数定义。

### 2.3 验证

```bash
.venv/bin/python -m py_compile poeworkflow-refactor/azure/migrate.py
.venv/bin/python -c "
import sys; sys.path.insert(0,'poeworkflow-refactor')
from azure.migrate import fix_assessment_excel_timestamps
import datetime, io
# 构造一个最小 Excel 练习验证
from openpyxl import Workbook
wb = Workbook()
# Assessment_Summary sheet
ws1 = wb.active
ws1.title = 'Assessment_Summary'
ws1.cell(row=1, column=1).value = 'Created on (UTC)'
ws1.cell(row=2, column=1).value = '5/15/2025 3:45:30 PM'
# Assessment_Properties sheet
ws2 = wb.create_sheet('Assessment_Properties')
ws2.cell(row=1, column=1).value = 'Property'
ws2.cell(row=2, column=1).value = 'Performance history start time'
ws2.cell(row=3, column=1).value = 'Performance history end time'
buf = io.BytesIO()
wb.save(buf)
excel_bytes = buf.getvalue()
result = fix_assessment_excel_timestamps(excel_bytes, datetime.date(2025, 3, 1), datetime.date(2025, 6, 30))
from openpyxl import load_workbook
wb2 = load_workbook(io.BytesIO(result))
created_val = wb2['Assessment_Summary'].cell(row=2, column=1).value
print(f'Created on (UTC): {created_val}')
# 读 Properties：用行扫描找 property/value pair
ws_props = wb2['Assessment_Properties']
for row in range(1, ws_props.max_row + 1):
    prop = str(ws_props.cell(row=row, column=1).value or '')
    val = str(ws_props.cell(row=row, column=2).value or '')
    if 'start' in prop.lower():
        print(f'Start: {val}')
    elif 'end' in prop.lower():
        print(f'End: {val}')
print('BEHAVIOR CHECK PASSED')
"
```

预期输出：Created on (UTC) 和 End 的值完全一致，Start = End - 1 天。

### 2.4 Commit

```bash
git add poeworkflow-refactor/azure/migrate.py
git commit -m "fix: Assessment Performance history end equals Created on UTC full datetime"
```

---

## 实施顺序
先做任务 1（删除功能，影响面大但简单），再做任务 2（精确修复一行逻辑）。

全部完成后 team report 汇报结果。
