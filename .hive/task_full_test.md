全量回归测试。验证以下所有改动（本次会话累计 7 个 commit）。

## 改动清单

| Commit | 内容 |
|--------|------|
| 9e74488 | 动态 SOLUTION_SYSTEM_PROMPT（大客户100K+ 详细模式） |
| 6300a4d | 删除 SVG 架构图生成逻辑 |
| 7b0d80f | 彻底移除 Azure Pricing Calculator 自动化 |
| 1e7e5dd | 删除 BUDGET_FILLER_SERVICES，_calibrate_budget 改用方案资源循环 |
| 9990da3 | Migrate 评估上限 120% + 修正 Assessment Excel 时间戳 |
| e30450e | Performance history end 等于 Created on (UTC) 完整 datetime |

## 验证步骤

### 1. 全文件语法编译
```bash
cd "/Users/penn/Documents/POE workflow"
for f in $(find poeworkflow-refactor -name "*.py" -not -path "*__pycache__*" -not -name "test_integration_auto_poe.py"); do
  .venv/bin/python -m py_compile "$f" || echo "FAIL: $f"
done
echo "SYNTAX CHECK DONE"
```

### 2. Import 验证
```bash
.venv/bin/python -c "
import sys; sys.path.insert(0,'poeworkflow-refactor')
from azure.migrate import fix_assessment_excel_timestamps
from llm.prompts import build_solution_system_prompt, build_infra_system_prompt
from pipeline import run_full_auto_poe, generate_solution_artifact
print('ALL IMPORTS OK')
"
```

### 3. 验证 Pricing 已完全移除
```bash
grep -rn "pricing" poeworkflow-refactor/ --include="*.py" | grep -v "azurePricingTier" || echo "PASS: No pricing references"
grep -rn "PRICING_AUTOMATION\|PricingExportResult\|run_pricing_export\|pricing_automation" poeworkflow-refactor/ --include="*.py" && echo "FAIL: pricing still referenced" || echo "PASS: All pricing imports removed"
```

### 4. 验证 SVG 已完全移除
```bash
grep -rn "svg\|SVG\|generate_svg\|_extract_svg\|SVG_SYSTEM" poeworkflow-refactor/ --include="*.py" | grep -iv "svg_code\|svg_architecture" || echo "PASS: No SVG references"
```

### 5. 验证大客户 prompt 函数可用
```bash
.venv/bin/python -c "
import sys; sys.path.insert(0,'poeworkflow-refactor')
from llm.prompts import build_solution_system_prompt
p_large = build_solution_system_prompt(True)
p_small = build_solution_system_prompt(False)
assert len(p_large) > len(p_small), 'large should be longer'
print(f'Large prompt: {len(p_large)} chars, Small prompt: {len(p_small)} chars - OK')
"
```

### 6. 验证时间戳行为
```bash
.venv/bin/python -c "
import sys; sys.path.insert(0,'poeworkflow-refactor')
from azure.migrate import fix_assessment_excel_timestamps
import datetime, io
from openpyxl import Workbook
wb = Workbook()
ws1 = wb.active
ws1.title = 'Assessment_Summary'
ws1.cell(row=1, column=1).value = 'Created on (UTC)'
ws1.cell(row=2, column=1).value = '5/15/2025 3:45:30 PM'
ws2 = wb.create_sheet('Assessment_Properties')
ws2.cell(row=1, column=1).value = 'Property'
ws2.cell(row=1, column=2).value = 'Value'
ws2.cell(row=2, column=1).value = 'Performance history start time'
ws2.cell(row=3, column=1).value = 'Performance history end time'
buf = io.BytesIO()
wb.save(buf)
excel_bytes = buf.getvalue()
result = fix_assessment_excel_timestamps(excel_bytes, datetime.date(2025, 3, 1), datetime.date(2025, 6, 30))
from openpyxl import load_workbook
wb2 = load_workbook(io.BytesIO(result))
created_val = wb2['Assessment_Summary'].cell(row=2, column=1).value
ws_props = wb2['Assessment_Properties']
for row in range(1, ws_props.max_row + 1):
    prop = str(ws_props.cell(row=row, column=1).value or '')
    val = str(ws_props.cell(row=row, column=2).value or '')
    if 'start' in prop.lower():
        start_val = val
    elif 'end' in prop.lower():
        end_val = val
print(f'Created: {created_val}')
print(f'Start: {start_val}')
print(f'End: {end_val}')
assert created_val == end_val, f'End ({end_val}) should equal Created ({created_val})'
print('TIMESTAMP CHECK PASSED: End = Created on (UTC)')
"
```

### 7. 回归测试
```bash
.venv/bin/python -m pytest test_tier_algorithm.py test_new_features.py -v 2>&1 | tail -20
```

### 8. Streamlit 启动检查
```bash
.venv/bin/streamlit run poeworkflow-refactor/main.py --server.port 8503 --server.headless true &
sleep 5
curl -s -o /dev/null -w "%{http_code}" http://localhost:8503
# 应返回 200
kill %1 2>/dev/null
```

## 输出要求
每项标注 PASS / FAIL。全部 PASS 后 team report 汇报结果。
