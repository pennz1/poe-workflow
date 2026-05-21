# app.py 重构设计

## 目标

将 4797 行的单体 `app.py` 拆分为多文件架构，解决模块边界不清的问题，为未来扩展打好基础。

## 方案选择

方案 B：按领域分包，中粒度拆分（~18 个文件，3 个子包）。

## 最终目录结构

```
项目根目录/
├── main.py                    # 新入口，~80 行
├── app.py                     # 兼容壳：from main import main
├── config.py                  # 常量、secrets、page_config、字体
├── session.py                 # Session 持久化
├── pipeline.py                # run_full_auto_poe 编排器
│
├── llm/
│   ├── __init__.py
│   ├── client.py              # OpenAI 客户端 + call_azure_openai
│   └── prompts.py             # 5 套 system prompt + extract_template_text
│
├── documents/
│   ├── __init__.py
│   ├── docx_utils.py          # 底层 docx 工具（字体、段落、MD→docx、表格、封面/目录）
│   ├── solution.py            # create_solution_docx + SVG 嵌入
│   ├── infra.py               # create_infra_docx
│   └── pov.py                 # create_pov_docx + build_pov_prompt
│
├── azure/
│   ├── __init__.py
│   ├── auth.py                # MSAL 登录、token、订阅/RG
│   ├── arm.py                 # ARM API 请求、分页、LRO 轮询
│   └── migrate.py             # Migrate 评估全管线（~1400 行）
│
├── budget/
│   ├── __init__.py
│   ├── parser.py              # parse_annual_budget_usd
│   └── tier.py                # snap、学习、缓存、机器选择
│
└── ui/
    ├── __init__.py
    ├── auto_poe.py            # 全自动 POE Tab 1
    └── tabs.py                # 解决方案/POV/CSV/年度价格 Tab 2-5
```

## 模块职责边界

### config.py
- 所有常量：路径（TEMPLATE_DIR 等）、API 版本号、endpoint、BUDGET_TIERS、CN_FONT、_PERSIST_KEYS
- `get_secret(key, default)` — 环境变量优先，fallback 到 st.secrets
- `check_secrets()` — 验证必需密钥
- `st.set_page_config()` 调用

### session.py
- `persist_session_state()` — 将 _PERSIST_KEYS 中的 session_state 写入 JSON 文件
- `restore_session_state()` — 从 JSON 文件恢复，清理过期 session
- `clear_session_persist()` — 删除持久化文件
- 依赖：st.session_state + 文件系统，不依赖其他项目模块

### llm/client.py
- `get_openai_client()` → OpenAI 实例
- `call_azure_openai(system_prompt, user_prompt)` → str
- 依赖：config.get_secret

### llm/prompts.py
- 5 个 system prompt 常量：SOLUTION_SYSTEM_PROMPT, INFRA_SYSTEM_PROMPT, POV_SYSTEM_PROMPT, SVG_SYSTEM_PROMPT, CSV_SYSTEM_PROMPT
- `extract_template_text(path)` — 从 .docx 提取文本，保留 `@st.cache_data`
- `_extract_svg_from_response(text)` — 从 AI 响应中提取 SVG 代码块

### documents/docx_utils.py
底层工具函数（从原 `_` 前缀函数重命名为公开 API）：
- `set_run_font(run, font_name, size_pt, bold, color_rgb)`
- `add_styled_paragraph(doc, text, ...)`
- `add_styled_heading(doc, text, level)`
- `markdown_to_docx(doc, markdown_text, body_size=9)` — Markdown 解析 → Word 元素
- `add_word_table(doc, table_data)`
- `parse_markdown_table(lines)` → Optional[List[List[str]]]
- `load_template(path)` → Document — 加载 .docx 模板，清空占位内容
- `extract_title(content, fallback)` → str — 从 Markdown 提取首个 # 标题
- `strip_first_heading(content)` → str — 移除首个 # 标题
- `add_page_break(doc)` — 插入分页符
- `add_toc(doc)` — 插入 Word TOC 域代码
- `add_svg_image_to_doc(doc, svg_code, width_cm)` → bool
- `svg_to_png_bytes(svg_code)` → Optional[bytes] — SVG→PNG（cairosvg→svglib→Edge）

### documents/solution.py
- `create_solution_docx(content, customer_name, svg_code=None)` → bytes
- `generate_svg_architecture(solution_text, customer_name)` → Optional[str] — 调 LLM 生成 SVG
- 依赖：docx_utils, llm.client, llm.prompts

### documents/infra.py
- `create_infra_docx(content, customer_name)` → bytes
- 依赖：docx_utils

### documents/pov.py
- `create_pov_docx(content, customer_name)` → bytes
- `build_pov_prompt(solution_text, customer_name, pov_start, pov_end, vendor_team, pov_ref)` → str
- `_workday_info(start, end)` → tuple[list, list]（内部函数）
- `has_meaningful_pov_team(vendor_team)` → bool

### azure/auth.py
- `is_azure_token_valid()` → bool
- `msal_device_code_login()` — 执行 Device Code Flow，写 st.session_state
- `clear_azure_login()` — 清除 session_state 中的 Azure 凭证
- `list_azure_subscriptions(token)` → list
- `list_azure_resource_groups(subscription_id, token)` → list
- `_subscription_label(sub)` → str
- `_resource_group_label(rg)` → str

### azure/arm.py
- `azure_arm_request(method, path_or_url, token, body, timeout, poll_lro, ...)` → dict
- `azure_arm_list(path, token)` → list（处理 nextLink 分页）
- `_poll_azure_lro(url, token, ...)` → dict
- `_retry_after_seconds(response)` → int
- `_format_arm_error(response)` → str

### azure/migrate.py（最大模块，~1400 行）
- `run_azure_migrate_assessment(token, subscription_id, resource_group, account_name, assessment_name, annual_budget_text, progress, target_location)` → dict
- 内部函数保持 `_` 前缀：
  - `_safe_azure_name`, `_safe_csv_prefix`, `_csv_template_hash`
  - `_extract_sas_url`, `_arm_path_with_api_version`, `_resource_id`
  - `_migrate_project_path`, `_migrate_solution_path`, `_migrate_solution_id`
  - `_servers_solution_details`, `_try_get_existing_resource`
  - `_server_summary_count`, `_format_import_job_error`, `_dedupe_machines_by_discovery_arm_id`
  - 各种 wait 函数（`wait_for_import_site_import`, `wait_for_project_machines`, 等）
  - 评估相关（`_build_assessment_body`, `_assessment_cost_component`, `_assessment_settings_snapshot`, `_extract_dominant_region`, `_REGION_NAME_TO_LOCATION`）
  - `register_azure_provider`, `register_migrate_tool`, `put_migrate_solution`, `ensure_migrate_solution`, `ensure_portal_menu_solutions`
  - `refresh_migrate_project_summary`, `resolve_migrate_project_location`
  - `prefix_csv_server_names`, `load_builtin_csv_template`, `_get_template_machine_names`
  - `fix_assessment_excel_timestamps`
  - `download_assessment_report`

### budget/parser.py
- `parse_annual_budget_usd(raw_budget)` → Optional[float]
- `snap_budget_to_tier(annual_budget)` → int
- `_format_usd(value)` → str

### budget/tier.py
- `load_tier_cache()` → dict
- `save_tier_cache(cache)` → None
- `learn_tier_machine_selections(assessed_machines, account_prefix, progress)` → dict
- `get_machine_ids_for_tier(tier, machines, account_prefix, cache)` → list
- `_strip_account_prefix(display_name, prefix)` → str

### pipeline.py
编排层，串起全自动 POE 流程：
- `run_full_auto_poe(...)` → dict — 全流程编排
- `generate_solution_artifact(doc_type, customer_name, account_name, customer_bg, solution_ref, infra_ref)` → dict
- `generate_pov_artifact(solution_text, customer_name, account_name, pov_ref, pov_start, pov_end, vendor_team)` → dict
- `create_solution_artifact_from_text(...)` → dict
- `create_pov_artifact_from_text(...)` → dict
- `get_existing_solution_text(current_doc_type)` → Optional[str]
- `get_generated_migrate_csv_text()` → Optional[str]
- `resolve_auto_inventory_csv(uploaded_inventory)` → tuple
- `format_auto_poe_log(message)` → str
- `should_display_auto_poe_log(message)` → bool
- `date_prefix()` → str
- `create_poe_zip(artifacts)` → bytes

### ui/auto_poe.py
- `render_full_auto_poe_area(current_doc_type, account_name, customer_name, budget, customer_bg, solution_ref, infra_ref, pov_ref)` → None
- 依赖：pipeline, azure.auth, budget.parser, budget.tier, st

### ui/tabs.py
- Tab 2：解决方案文档（AI 生成 + 手动导入 + 下载）
- Tab 3：POV 部署计划（生成 + 下载 + 预览）
- Tab 4：Azure Migrate CSV（上传 Excel → AI 生成 CSV → 下载预览）
- Tab 5：年度价格表（上传 Excel → 添加 yearly cost 列 → 下载）

### main.py（~80 行）
```python
def main():
    restore_session_state()
    render_app_header()
    if not check_secrets():
        st.stop()
    # 侧边栏：清除按钮、模板状态
    # 公共输入区：客户信息、POV 日期、人员
    # Tab 布局 → 调 ui.auto_poe / ui.tabs

if __name__ == "__main__":
    main()
```

### app.py（兼容壳）
```python
from main import main
if __name__ == "__main__":
    main()
```

## 依赖规则

1. `st.`（Streamlit session_state/UI）只出现在 `ui/`、`main.py`、`pipeline.py`、`azure/auth.py` 中
2. `st.cache_data` 只出现在 `llm/prompts.py`（extract_template_text）
3. 纯逻辑模块（config、llm、documents、budget、azure/arm、azure/migrate）不 import streamlit
4. 依赖方向严格单向：config → llm → documents → pipeline → ui → main
5. azure/ 分支：auth → arm → migrate，migrate 也依赖 budget/tier

## 现有依赖保持不变

- `frontend/ui.py` — 不改动，继续 import
- `pricing_automation.py` — 不改动，继续 import（HAS_PRICING_AUTOMATION 标记保留在 pipeline.py）
- `cli_generate.py` — 不改动

## 不在范围内

- 不添加新功能
- 不修改业务逻辑
- 不修改 API 调用方式
- 不修改 Word 文档样式
- 不添加测试（可在后续 PR 中补充）
- 不改动 `frontend/ui.py`、`pricing_automation.py`、`cli_generate.py`
