修改 SOLUTION_SYSTEM_PROMPT：根据客户年预算自动调整输出详细度和架构严谨性。

## 背景
当前 SOLUTION_SYSTEM_PROMPT 是一个固定常量，不对大客户做区分。需要改为根据年预算动态调整：年100K以上大客户输出更详细、架构更严谨。

## 改动 1：将 SOLUTION_SYSTEM_PROMPT 改为函数

文件：poeworkflow-refactor/llm/prompts.py

删除当前第31-94行的 `SOLUTION_SYSTEM_PROMPT = (...)` 常量定义，替换为函数 `build_solution_system_prompt(is_large_customer: bool = False) -> str`。

新函数核心逻辑：
- is_large_customer=True（年100K+）时：各章节要求更详细、文字量1500-2500字、资源表8-12行、严禁编造不存在的Azure服务/SKU、架构必须严格合理（数据流有明确方向、服务间集成有明确协议/接口）
- is_large_customer=False 时：保持原有简洁风格不变

## 改动 2：同步更新 INFRA_SYSTEM_PROMPT

文件：poeworkflow-refactor/llm/prompts.py

将 INFRA_SYSTEM_PROMPT 也改为函数 `build_infra_system_prompt(is_large_customer: bool = False) -> str`，保持内容不变，仅包一层函数。

## 改动 3：更新 pipeline.py 的 generate_solution_artifact

文件：poeworkflow-refactor/pipeline.py

1. 修改 import：`SOLUTION_SYSTEM_PROMPT, INFRA_SYSTEM_PROMPT` → `build_solution_system_prompt, build_infra_system_prompt`
2. 函数签名增加 `annual_budget: float = 0` 参数
3. 函数内根据 annual_budget >= 100_000 判断 is_large
4. user_ctx 中加提示（大客户时追加"该客户为年度预算超10万美元的大客户，请提供更详尽、更严谨的架构方案"，不显示具体金额）
5. 调用 build_solution_system_prompt(is_large_customer=is_large) 或 build_infra_system_prompt(is_large_customer=is_large)

## 改动 4：更新 pipeline.py 的 run_full_auto_poe

在调用 generate_solution_artifact 处传入 annual_budget=parse_annual_budget_usd(annual_budget_text) or 0.0

## 验证步骤
1. `.venv/bin/python -m py_compile poeworkflow-refactor/llm/prompts.py`
2. `.venv/bin/python -m py_compile poeworkflow-refactor/pipeline.py`
3. `.venv/bin/python -c "import sys; sys.path.insert(0,'poeworkflow-refactor'); from llm.prompts import build_solution_system_prompt, build_infra_system_prompt; p=build_solution_system_prompt(True); print('LARGE OK', len(p)); p2=build_solution_system_prompt(False); print('SMALL OK', len(p2)); assert len(p) > len(p2), 'large should be longer'; print('ALL OK')"`
4. `.venv/bin/python -m pytest test_new_features.py -v 2>&1 | tail -15`

全部通过后 git add && git commit -m "feat: dynamic SOLUTION_SYSTEM_PROMPT with large-customer detailed mode"

完成后 team report 汇报结果。
