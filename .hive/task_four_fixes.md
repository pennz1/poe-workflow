四个改动，请按顺序完成。

---

## 修复 1：Azure 登录 — 去掉 window.open，保留自动复制 + 优雅手动 UI

**问题**：Chrome 拦截 window.open 弹窗。

**方案**：保留自动复制到剪贴板（不触发弹窗拦截），去掉 window.open。改为紧凑优雅的 UI：验证码旁边一个小复制按钮 + 一个明显的 Microsoft 登录链接。

**文件**：frontend/ui.py，render_device_code_login 函数（第 138-200 行）

**替换为**：

```python
def render_device_code_login(user_code: str, verify_url: str) -> None:
    code = _html(user_code)
    url = _html(verify_url)
    components.html(
        f"""
        <style>
        .poe-device-login {{
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 10px 14px;
            border-radius: 8px;
            border: 1px solid oklch(86.5% 0.014 248);
            background: oklch(99.2% 0.003 248);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei UI", system-ui, sans-serif;
            font-size: 14px;
        }}
        .poe-device-code {{
            font-weight: 700;
            font-size: 15px;
            letter-spacing: 0.5px;
            color: oklch(39% 0.15 250);
            background: oklch(94% 0.01 250);
            padding: 3px 10px;
            border-radius: 4px;
            user-select: all;
        }}
        .poe-device-copy-btn {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 28px;
            height: 28px;
            border: 1px solid oklch(80% 0.05 250);
            border-radius: 6px;
            background: oklch(100% 0 0);
            cursor: pointer;
            font-size: 13px;
            padding: 0;
            flex-shrink: 0;
        }}
        .poe-device-copy-btn:hover {{
            background: oklch(94% 0.01 250);
        }}
        .poe-device-copy-btn.copied {{
            background: oklch(48% 0.105 152);
            border-color: oklch(48% 0.105 152);
        }}
        .poe-device-link-btn {{
            display: inline-flex;
            align-items: center;
            padding: 5px 14px;
            border-radius: 6px;
            border: 1px solid oklch(58% 0.14 252);
            background: oklch(58% 0.14 252);
            color: #fff;
            font-weight: 600;
            font-size: 13px;
            text-decoration: none;
            cursor: pointer;
            margin-left: auto;
            flex-shrink: 0;
        }}
        .poe-device-link-btn:hover {{
            background: oklch(48% 0.12 252);
            border-color: oklch(48% 0.12 252);
        }}
        .poe-device-hint {{
            color: oklch(60% 0.02 250);
            font-size: 12px;
            flex-shrink: 1;
            min-width: 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        </style>
        <div class="poe-device-login">
            <span class="poe-device-code">{code}</span>
            <button type="button" class="poe-device-copy-btn" id="poe-copy-btn" title="复制验证码">📋</button>
            <span class="poe-device-hint" id="poe-hint">已自动复制</span>
            <a class="poe-device-link-btn" href="{url}" target="_blank" rel="noopener noreferrer">打开 Microsoft 登录 →</a>
        </div>
        <script>
        (function() {{
            const code = "{code}";
            const btn = document.getElementById("poe-copy-btn");
            const hint = document.getElementById("poe-hint");

            const copyToClipboard = async function(text) {{
                try {{
                    await navigator.clipboard.writeText(text);
                    return true;
                }} catch (e) {{
                    const ta = document.createElement("textarea");
                    ta.value = text;
                    ta.setAttribute("readonly", "");
                    ta.style.position = "fixed";
                    ta.style.opacity = "0";
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand("copy");
                    ta.remove();
                    return true;
                }}
            }};

            // 自动复制
            (async function() {{
                const ok = await copyToClipboard(code);
                if (ok) {{
                    hint.textContent = "已自动复制";
                }} else {{
                    hint.textContent = "请手动复制";
                }}
            }})();

            // 手动复制按钮
            btn.addEventListener("click", async function() {{
                await copyToClipboard(code);
                btn.textContent = "✓";
                btn.classList.add("copied");
                hint.textContent = "已复制";
                setTimeout(function() {{
                    btn.textContent = "📋";
                    btn.classList.remove("copied");
                }}, 1500);
            }});
        }})();
        </script>
        """,
        height=52,
    )
```

---

## 修复 2：任务完成时间后增加耗时显示

**文件**：poeworkflow-refactor/ui/auto_poe.py

**改动 1**：在生成开始时记录起始时间（约第 218 行，`st.session_state["auto_poe_running"] = True` 之后）：
```python
st.session_state["auto_poe_start_time"] = time.time()
```
并在文件顶部 import time（如果尚未导入）。

**改动 2**：修改完成时间显示逻辑（约第 314-316 行），增加耗时计算：
```python
finish_time = st.session_state.get("auto_poe_finish_time")
if finish_time:
    start_ts = st.session_state.get("auto_poe_start_time")
    if start_ts:
        elapsed_sec = int(time.time() - start_ts)
        if elapsed_sec < 60:
            elapsed_str = f"{elapsed_sec} 秒"
        elif elapsed_sec < 3600:
            minutes = elapsed_sec // 60
            seconds = elapsed_sec % 60
            elapsed_str = f"{minutes} 分 {seconds} 秒"
        else:
            hours = elapsed_sec // 3600
            minutes = (elapsed_sec % 3600) // 60
            elapsed_str = f"{hours} 时 {minutes} 分"
        st.info(f"✅ 任务完成时间：{finish_time}（耗时 {elapsed_str}）")
    else:
        st.info(f"✅ 任务完成时间：{finish_time}")
```

---

## 修复 3：确保 Azure Migrate 评估包含 Storage/Disk 评估

**文件**：poeworkflow-refactor/azure/migrate.py

**改动 1**：在 `_build_assessment_body` 函数中（约第 651 行），确认并强化 disk 评估配置。将 `azureDiskTypes` 从单行改为包含显式注释：
```python
"azureDiskTypes": ["PremiumSSD", "StandardSSD", "StandardHDD"],
```
（注意：将 "Premium" 改为标准名称 "PremiumSSD"，将 "Standard" 改为 "StandardHDD"）

**改动 2**：在 `tune_assessment_to_budget` 函数中（约第 710 行 `_record` 定义之前），增加一轮尝试调整 `azureStorageRedundancy` 的逻辑。在调整 AHUB/RI 的同时（第 730-790 行附近，具体在 scalingFactor 调整逻辑前后），添加：
```python
# 如有必要，调整存储冗余级别以微调成本
for redundancy in ["LocallyRedundant", "GeoRedundant"]:
    if redundancy == current_props.get("azureStorageRedundancy", ""):
        continue
    test_body = {**assessment_body, "properties": {**assessment_body["properties"], "azureStorageRedundancy": redundancy}}
    test_result = azure_arm_request("PUT", assessment_path, token, test_body)
    test_monthly = assessment_monthly_total_cost(test_result)
    test_annual = test_monthly * 12
    if target_min and test_annual < target_min:
        continue
    if target_max and test_annual > target_max:
        continue
    assessment_body = test_body
    assessment = test_result
    progress(f"  调整为存储冗余 {redundancy}: 年化 {_format_usd(test_annual)} ✓")
    break
```

## 修复 4：确认 Created on (UTC) 日期已修改

**说明**：当前 `fix_assessment_excel_timestamps` 逻辑已正确：随机选 POV 区间日期、保留原时分秒、Created on (UTC) 和 Performance history end 设为同一完整 datetime。无需代码改动。

**验证**：在验证步骤中加入专项检查确认。

---

## 清理：删除残余无用文件

```bash
cd "/Users/penn/Documents/POE workflow"
# 删除已被 git rm 的 1.svg（工作区残留）
rm -f 1.svg
# 删除浏览器 profile（属于 pricing_automation 的残留）
rm -rf .browser_profile/
# 删除 .codex/ 配置文件
rm -rf .codex/
```

---

## 实施顺序
修复 1 → 修复 2 → 修复 3 → 清理

## Commit
- 修复 1: `fix: revert window.open in device code login, use elegant manual UI`
- 修复 2: `feat: show elapsed time after POE generation completes`
- 修复 3: `fix: ensure Azure Migrate assessment evaluates storage disks`
- 清理: `chore: remove residual unused files`

---

## 验证
```bash
.venv/bin/python -m py_compile frontend/ui.py
.venv/bin/python -m py_compile poeworkflow-refactor/ui/auto_poe.py
.venv/bin/python -m py_compile poeworkflow-refactor/azure/migrate.py
.venv/bin/python -m pytest test_tier_algorithm.py test_new_features.py -v 2>&1 | tail -10
# Created on (UTC) 验证（已在修复 4 中说明逻辑正确，此处确认 import 可用）
.venv/bin/python -c "
import sys; sys.path.insert(0,'poeworkflow-refactor')
from azure.migrate import fix_assessment_excel_timestamps
import datetime, io
from openpyxl import Workbook
wb = Workbook()
ws1 = wb.active
ws1.title = 'Assessment_Summary'
ws1.cell(row=1, column=1).value = 'Created on (UTC)'
ws1.cell(row=2, column=1).value = '5/15/2026 3:45:30 PM'
ws2 = wb.create_sheet('Assessment_Properties')
ws2.cell(row=1, column=1).value = 'Property'
ws2.cell(row=1, column=2).value = 'Value'
ws2.cell(row=2, column=1).value = 'Performance history start time'
ws2.cell(row=3, column=1).value = 'Performance history end time'
buf = io.BytesIO()
wb.save(buf)
excel_bytes = buf.getvalue()
result = fix_assessment_excel_timestamps(excel_bytes, datetime.date(2026, 3, 1), datetime.date(2026, 6, 30))
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
assert created_val == end_val, f'MISMATCH: Created={created_val} vs End={end_val}'
# 验证日期在 POV 区间内
created_dt = datetime.datetime.strptime(created_val, '%m/%d/%Y %I:%M:%S %p')
assert datetime.date(2026, 3, 1) <= created_dt.date() <= datetime.date(2026, 6, 30), 'Date not in POV range'
# 验证时间保留
assert created_dt.hour == 15 and created_dt.minute == 45 and created_dt.second == 30, 'Time not preserved'
print('CREATED ON UTC CHECK PASSED: date modified in POV range, time preserved, end=created')
"
```

全部验证通过后 team report 汇报结果。
