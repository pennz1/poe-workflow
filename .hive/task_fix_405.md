修复 Azure Migrate 评估 405 错误：assessment is under computation, cannot be updated。

## 问题
全自动 POE 流程中 `tune_assessment_to_budget` 函数在调整存储冗余时，直接 PUT 更新评估设置，但 Azure Migrate API 返回 405："assessment is under computation. Cannot be updated." 原因是上次计算刚完成但内部状态尚未完全释放，或 PUT 后触发了新的计算但未等待完成就读取成本。

## 根因定位
文件：poeworkflow-refactor/azure/migrate.py，tune_assessment_to_budget 函数（第 682 行起）
- 第 802 行：`test_result = azure_arm_request("PUT", assessment_path, token, test_body)` — 没有 wait_for_assessment_complete，且没有捕获 405
- 第 825 行：`azure_arm_request("PUT", assessment_path, token, assessment_body)` — 有 wait_for_assessment_complete 但也没有 405 重试

## 修复

### 改动 1：新增 _put_assessment_and_wait 辅助函数

在 tune_assessment_to_budget 函数之前（约第 681 行）添加辅助函数：

```python
def _put_assessment_and_wait(
    assessment_path: str,
    assessment_body: Dict[str, Any],
    token: str,
    subscription_id: str,
    resource_group: str,
    project_name: str,
    group_name: str,
    assessment_name: str,
    progress: Callable[[str], None],
    timeout_seconds: int = 600,
) -> Dict[str, Any]:
    """PUT 评估更新并等待计算完成，自动处理 405 冲突。"""
    deadline = time.time() + timeout_seconds
    attempt = 0
    while time.time() < deadline:
        attempt += 1
        try:
            azure_arm_request("PUT", assessment_path, token, assessment_body)
            break
        except Exception as e:
            msg = str(e)
            if "405" in msg or "under computation" in msg.lower() or "Cannot be updated" in msg:
                if attempt == 1:
                    progress(f"  ⏳ 评估仍在计算中，等待完成后再更新...")
                time.sleep(15)
                continue
            raise

    return wait_for_assessment_complete(
        subscription_id, resource_group, project_name, group_name, assessment_name, token, progress
    )
```

### 改动 2：重写存储冗余调优段（第 797-814 行）

替换为使用 _put_assessment_and_wait：

```python
    current_props = assessment_body["properties"]
    # 如有必要，调整存储冗余级别以微调成本
    for redundancy in ["LocallyRedundant", "GeoRedundant"]:
        if redundancy == current_props.get("azureStorageRedundancy", ""):
            continue
        test_body = {**assessment_body, "properties": {**assessment_body["properties"], "azureStorageRedundancy": redundancy}}
        try:
            test_result = _put_assessment_and_wait(
                assessment_path, test_body, token,
                subscription_id, resource_group, project_name, group_name, assessment_name,
                progress,
            )
        except Exception as e:
            progress(f"  ⚠️ 存储冗余调整失败：{e}")
            continue
        test_monthly = assessment_monthly_total_cost(test_result)
        test_annual = test_monthly * 12
        if target_min and test_annual < target_min:
            continue
        if target_max and test_annual > target_max:
            continue
        assessment_body = test_body
        assessment = test_result
        annual_total = test_annual
        progress(f"  调整为存储冗余 {redundancy}: 年化 {_format_usd(test_annual)} ✓")
        _record("storage-redundancy", f"调整为存储冗余 {redundancy}", assessment, True)
        return assessment, history, True
```

### 改动 3：调优主循环 PUT 也改用 _put_assessment_and_wait（第 823-828 行）

```python
        assessment_body["properties"].update(patch)
        assessment_body["properties"]["stage"] = "InProgress"
        assessment = _put_assessment_and_wait(
            assessment_path, assessment_body, token,
            subscription_id, resource_group, project_name, group_name, assessment_name,
            progress,
        )
```

## 验证

```bash
.venv/bin/python -m py_compile poeworkflow-refactor/azure/migrate.py
.venv/bin/python -m pytest test_tier_algorithm.py test_new_features.py -v 2>&1 | tail -10
.venv/bin/python -c "
import sys; sys.path.insert(0,'poeworkflow-refactor')
from azure.migrate import tune_assessment_to_budget, _put_assessment_and_wait
print('IMPORT OK')
"
grep -n "_put_assessment_and_wait" poeworkflow-refactor/azure/migrate.py
# 应显示 3 处：函数定义 + 2 处调用
```

## Commit

```bash
git add poeworkflow-refactor/azure/migrate.py
git commit -m "fix: handle 405 under-computation error in assessment tuning with retry"
```

完成后 team report 汇报结果。
