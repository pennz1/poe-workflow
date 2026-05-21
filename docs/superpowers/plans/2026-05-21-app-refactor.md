# app.py 重构实现计划

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推荐）或 superpowers:executing-plans 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。

**目标：** 将 4797 行的单体 `app.py` 拆分为 ~18 个模块文件，放入新目录 `poeworkflow-refactor/`，不改动原项目。

**架构：** 按领域分包（llm/、documents/、azure/、budget/、ui/）+ 顶层扁平模块（config.py、session.py、pipeline.py、main.py）。纯逻辑模块不依赖 Streamlit，UI 模块集中在 ui/。

**技术栈：** Streamlit、OpenAI SDK、python-docx、MSAL、requests

---

## 文件结构总览

新建目录 `poeworkflow-refactor/`，所有新文件放在其中。

```
poeworkflow-refactor/
├── __init__.py
├── config.py
├── session.py
├── pipeline.py
├── main.py
├── app.py                     # 兼容壳，同原项目根目录的 app.py
├── llm/
│   ├── __init__.py
│   ├── client.py
│   └── prompts.py
├── documents/
│   ├── __init__.py
│   ├── docx_utils.py
│   ├── solution.py
│   ├── infra.py
│   └── pov.py
├── azure/
│   ├── __init__.py
│   ├── auth.py
│   ├── arm.py
│   └── migrate.py
├── budget/
│   ├── __init__.py
│   ├── parser.py
│   └── tier.py
└── ui/
    ├── __init__.py
    ├── auto_poe.py
    └── tabs.py
```

---

### 任务 1：创建目录结构和 config.py

**文件：**
- 创建：`poeworkflow-refactor/__init__.py`
- 创建：`poeworkflow-refactor/config.py`
- 创建：`poeworkflow-refactor/llm/__init__.py`
- 创建：`poeworkflow-refactor/documents/__init__.py`
- 创建：`poeworkflow-refactor/azure/__init__.py`
- 创建：`poeworkflow-refactor/budget/__init__.py`
- 创建：`poeworkflow-refactor/ui/__init__.py`

- [ ] **步骤 1：创建目录结构**

```bash
mkdir -p poeworkflow-refactor/{llm,documents,azure,budget,ui}
touch poeworkflow-refactor/__init__.py
touch poeworkflow-refactor/llm/__init__.py
touch poeworkflow-refactor/documents/__init__.py
touch poeworkflow-refactor/azure/__init__.py
touch poeworkflow-refactor/budget/__init__.py
touch poeworkflow-refactor/ui/__init__.py
```

- [ ] **步骤 2：创建 config.py**

从 app.py 提取以下内容（行 8-97）：
- import 语句（io, json, os, re, copy, csv, datetime, hashlib, time, zipfile, typing）
- APP_DIR, TEMPLATE_DIR, SOLUTION_TEMPLATE_PATH, INFRA_TEMPLATE_PATH, POV_TEMPLATE_PATH, MIGRATE_TEMPLATE_PATH
- MSAL_CLIENT_ID_DEFAULT, AZURE_AUTHORITY, AZURE_ARM_SCOPE, AZURE_MANAGEMENT_ENDPOINT
- 所有 AZURE_*_API_VERSION 常量
- AZURE_MIGRATE_DEFAULT_TARGET_LOCATION
- BUILTIN_CSV_PATH
- PERSIST_DIR, TIER_CACHE_PATH, BUDGET_TIERS
- _PERSIST_KEYS
- CN_FONT, CN_FONT_ALT
- `get_secret(key, default=None)` 函数
- `check_secrets()` 函数
- `st.set_page_config(...)` 调用
- `load_desktop_theme(APP_DIR)` 调用（import 自 frontend.ui）

注意：`APP_DIR` 改为指向原项目目录（因为模板文件仍在原项目），所以用绝对路径：

```python
import os
APP_DIR = os.path.dirname(os.path.abspath(__file__))
# 模板和依赖文件仍在原项目目录
ORIGINAL_APP_DIR = os.path.join(os.path.dirname(APP_DIR))  # 实际上直接用 APP_DIR 的父目录
# 不，更简单的方式：
ORIGINAL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # poeworkflow-refactor 的父目录即原项目
TEMPLATE_DIR = os.path.join(ORIGINAL_DIR, "templates")
BUILTIN_CSV_PATH = os.path.join(ORIGINAL_DIR, "Azurecsvtemplate.csv")
PERSIST_DIR = os.environ.get("PERSIST_DIR", ORIGINAL_DIR)
# 等等，PERSIST_DIR 应该保持一致用原来的
```

实际上更简洁：直接用 `ORIGINAL_DIR` 指向项目根目录，所有路径基于它。

```python
"""
POE workflow 配置常量
"""
import os

# 项目根目录（poeworkflow-refactor 的父目录）
PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

TEMPLATE_DIR = os.path.join(PROJECT_DIR, "templates")
SOLUTION_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "solution_template.docx.docx")
INFRA_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "Infra_template.docx")
POV_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "pov_template.docx.docx")
MIGRATE_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "AzureMigrateimporttemplate.csv")

MSAL_CLIENT_ID_DEFAULT = "d999b252-3977-4eae-bdae-a0f58c5336b7"
AZURE_AUTHORITY = "https://login.microsoftonline.com/common"
AZURE_ARM_SCOPE = ["https://management.azure.com/.default"]
AZURE_MANAGEMENT_ENDPOINT = "https://management.azure.com"
AZURE_RESOURCE_API_VERSION = "2022-12-01"
AZURE_PROVIDER_API_VERSION = "2021-04-01"
AZURE_MIGRATE_API_VERSION = "2024-01-15"
AZURE_MIGRATE_REPORT_API_VERSION = "2019-10-01"
AZURE_OFFAZURE_API_VERSION = "2023-06-06"
AZURE_MIGRATE_PROJECTS_API_VERSION = "2018-09-01-preview"
AZURE_MIGRATE_DEFAULT_TARGET_LOCATION = "WestUs2"

BUILTIN_CSV_PATH = os.path.join(PROJECT_DIR, "Azurecsvtemplate.csv")

PERSIST_DIR = os.environ.get("PERSIST_DIR", PROJECT_DIR)
TIER_CACHE_PATH = os.path.join(PERSIST_DIR, ".tier_cache.json")
BUDGET_TIERS = [15_000, 50_000, 100_000, 250_000]

_PERSIST_KEYS = [
    "azure_token", "azure_user", "azure_token_expires_at",
    "azure_subscription_id", "azure_subscription_name", "azure_resource_group",
    "_cached_subscription", "_cached_resource_group",
    "solution_text", "infra_text", "pov_text", "csv_code",
    "customer_name", "account_name", "budget", "doc_type",
    "pov_source_doc_type", "pov_vendor_team",
    "auto_poe_result",
]

CN_FONT = "微软雅黑"
CN_FONT_ALT = "Microsoft YaHei UI"


def get_secret(key, default=None):
    """优先从环境变量读取，fallback 到 Streamlit secrets，再 fallback 到 default。"""
    val = os.environ.get(key)
    if val:
        return val
    try:
        import streamlit as st
        return st.secrets.get(key, default)
    except Exception:
        return default


def check_secrets():
    """检查是否已配置所需的 OpenAI 兼容 API 凭据。"""
    import streamlit as st
    required_keys = ["AZURE_OPENAI_KEY", "AZURE_OPENAI_ENDPOINT", "AZURE_OPENAI_DEPLOYMENT"]
    missing = [k for k in required_keys if not get_secret(k)]
    if missing:
        st.error("⚠️ **OpenAI API 配置缺失**")
        st.info(
            "请通过环境变量或 `.streamlit/secrets.toml` 配置以下密钥：\n\n"
            "```toml\n"
            'AZURE_OPENAI_KEY = "your-api-key"\n'
            'AZURE_OPENAI_ENDPOINT = "https://your-gateway.example.com/"\n'
            'AZURE_OPENAI_DEPLOYMENT = "gpt-4o"  # 模型名称\n'
            "```"
        )
        return False
    return True
```

- [ ] **步骤 3：Commit**

```bash
git add poeworkflow-refactor/
git commit -m "feat: init poeworkflow-refactor directory structure and config module"
```

---

### 任务 2：创建 llm/ 模块

**文件：**
- 创建：`poeworkflow-refactor/llm/client.py`
- 创建：`poeworkflow-refactor/llm/prompts.py`

- [ ] **步骤 1：创建 llm/client.py**

从 app.py 提取 `get_openai_client()` 和 `call_azure_openai()`（行 152-192）：

```python
"""OpenAI 兼容客户端封装"""
from openai import OpenAI
from ..config import get_secret


def get_openai_client():
    """创建 OpenAI 兼容客户端实例（支持 NewAPI 等网关）。"""
    endpoint = get_secret("AZURE_OPENAI_ENDPOINT").rstrip("/")
    base_url = endpoint if endpoint.endswith("/v1") else endpoint + "/v1"
    return OpenAI(
        api_key=get_secret("AZURE_OPENAI_KEY"),
        base_url=base_url,
    )


def call_azure_openai(system_prompt, user_prompt):
    """调用 OpenAI 兼容 Chat Completions API 并返回文本结果。"""
    client = get_openai_client()
    response = client.chat.completions.create(
        model=get_secret("AZURE_OPENAI_DEPLOYMENT"),
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.7,
        max_tokens=16384,
    )
    if isinstance(response, str):
        raise RuntimeError(
            f"API 返回了非预期的原始字符串响应。响应内容: {response[:500]}"
        )
    if not hasattr(response, "choices") or not response.choices:
        raise RuntimeError(
            f"API 返回了无效响应结构: {type(response).__name__}。"
            f"请检查 API 密钥、端点和模型名称是否正确。"
        )
    content = response.choices[0].message.content
    if not content or not content.strip():
        raise ValueError(
            f"API 返回了空内容。finish_reason={response.choices[0].finish_reason}"
        )
    return content
```

- [ ] **步骤 2：创建 llm/prompts.py**

从 app.py 提取（行 198-748）：
- `extract_template_text(path)` — 保留 `@st.cache_data`
- SOLUTION_SYSTEM_PROMPT
- INFRA_SYSTEM_PROMPT
- POV_SYSTEM_PROMPT
- SVG_SYSTEM_PROMPT
- CSV_SYSTEM_PROMPT
- `_extract_svg_from_response(text)`

```python
"""Prompt 模板和模板文本提取"""
import re
import streamlit as st
from docx import Document


@st.cache_data
def extract_template_text(path):
    """从 .docx 模板文件中提取所有文本内容（含表格），用于注入 AI prompt。"""
    doc = Document(path)
    lines = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            lines.append(text)
    for table in doc.tables:
        header_cells = [cell.text.strip() for cell in table.rows[0].cells]
        lines.append("| " + " | ".join(header_cells) + " |")
        lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |")
        for row in table.rows[1:]:
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            lines.append("| " + " | ".join(cells) + " |")
        lines.append("")
    return "\n".join(lines)


def _extract_svg_from_response(text):
    """从 AI 响应中提取 SVG 代码块。"""
    m = re.search(r"```(?:svg|xml)?\s*\n(.*?)```", text, re.DOTALL)
    if m:
        return m.group(1).strip()
    m = re.search(r"(<svg[\s\S]*?</svg>)", text, re.DOTALL)
    if m:
        return m.group(1).strip()
    return text.strip()


SOLUTION_SYSTEM_PROMPT = (
    # ... 完整内容（从 app.py 行 221-284 复制）
)
# 注意：此处省略完整 prompt 文本，实际创建文件时必须完整复制
# INFRA_SYSTEM_PROMPT = (...)  # 行 286-341
# POV_SYSTEM_PROMPT = (...)    # 行 343-378
# SVG_SYSTEM_PROMPT = (...)    # 行 383-432
# CSV_SYSTEM_PROMPT = (...)    # 行 724-748
```

**重要：** 实际创建时必须完整复制所有 prompt 文本，不能截断。

- [ ] **步骤 3：Commit**

```bash
git add poeworkflow-refactor/llm/
git commit -m "feat: add llm client and prompts modules"
```

---

### 任务 3：创建 documents/ 模块

**文件：**
- 创建：`poeworkflow-refactor/documents/docx_utils.py`
- 创建：`poeworkflow-refactor/documents/solution.py`
- 创建：`poeworkflow-refactor/documents/infra.py`
- 创建：`poeworkflow-refactor/documents/pov.py`

- [ ] **步骤 1：创建 documents/docx_utils.py**

从 app.py 提取（行 750-993）：
- `_set_run_font` → 重命名为 `set_run_font`
- `_add_styled_paragraph` → `add_styled_paragraph`
- `_add_styled_heading` → `add_styled_heading`
- `_parse_markdown_table` → `parse_markdown_table`
- `_add_word_table` → `add_word_table`
- `_markdown_to_docx` → `markdown_to_docx`
- `_load_template` → `load_template`
- `_extract_title` → `extract_title`
- `_strip_first_heading` → `strip_first_heading`
- `_add_page_break` → `add_page_break`
- `_add_toc` → `add_toc`
- `_svg_to_png_bytes` → `svg_to_png_bytes`
- `_svg_to_png_via_edge` → 保持为内部函数
- `_add_svg_image_to_doc` → `add_svg_image_to_doc`

所有函数签名去掉 `_` 前缀，import 使用：
```python
from ..config import CN_FONT, CN_FONT_ALT
```

- [ ] **步骤 2：创建 documents/solution.py**

从 app.py 提取（行 1038-1109）：
- `create_solution_docx(content, customer_name, svg_code=None)` → bytes
- 内联导入 `load_template`, `extract_title`, `strip_first_heading`, `add_page_break`, `add_toc`, `markdown_to_docx`, `add_svg_image_to_doc` 从 `.docx_utils`

- [ ] **步骤 3：创建 documents/infra.py**

从 app.py 提取（行 1144-1179）：
- `create_infra_docx(content, customer_name)` → bytes

- [ ] **步骤 4：创建 documents/pov.py**

从 app.py 提取：
- `create_pov_docx(content, customer_name)` → bytes（行 1112-1141）
- `build_pov_prompt(...)`（行 1487-1513）
- `_workday_info(start, end)`（行 1453-1464）
- `has_meaningful_pov_team(vendor_team)`（行 1467-1484）

- [ ] **步骤 5：Commit**

```bash
git add poeworkflow-refactor/documents/
git commit -m "feat: add documents module (docx utils, solution, infra, pov)"
```

---

### 任务 4：创建 budget/ 模块

**文件：**
- 创建：`poeworkflow-refactor/budget/parser.py`
- 创建：`poeworkflow-refactor/budget/tier.py`

- [ ] **步骤 1：创建 budget/parser.py**

从 app.py 提取（行 2050-2092, 2436-2439）：
- `parse_annual_budget_usd(raw_budget)` → Optional[float]
- `_format_usd(value)` → str

```python
"""预算解析"""
import re
from typing import Optional


def parse_annual_budget_usd(raw_budget):
    """把 UI 里的年预估消耗解析成 USD 数字；无法可靠解析时返回 None。"""
    if raw_budget is None:
        return None
    text = str(raw_budget).strip()
    if not text:
        return None
    normalized = (
        text.lower()
        .replace(",", "")
        .replace("$", "")
        .replace("usd", "")
        .replace("美元", "")
        .replace("美金", "")
        .replace("预估", "")
        .replace("年消耗", "")
        .replace("年度", "")
        .replace("每年", "")
        .replace("+", "")
        .strip()
    )
    if not normalized or normalized in {"na", "n/a", "none", "null", "未填写", "无"}:
        return None
    multiplier = 1.0
    if "万" in normalized or re.search(r"\d\s*w\b", normalized):
        multiplier = 10_000.0
        normalized = normalized.replace("万", "").replace("w", "")
    elif "million" in normalized:
        multiplier = 1_000_000.0
        normalized = normalized.replace("million", "")
    elif re.search(r"\d\s*m\b", normalized):
        multiplier = 1_000_000.0
        normalized = re.sub(r"\bm\b", "", normalized)
    elif re.search(r"\d\s*k\b", normalized):
        multiplier = 1_000.0
        normalized = re.sub(r"\bk\b", "", normalized)
    match = re.search(r"(\d+(?:\.\d+)?)", normalized)
    if not match:
        return None
    return float(match.group(1)) * multiplier


def _format_usd(value):
    if value is None:
        return "未填写"
    return f"${value:,.2f}"
```

- [ ] **步骤 2：创建 budget/tier.py**

从 app.py 提取（行 2098-2406）：
- `load_builtin_csv_template()` → 依赖 config.BUILTIN_CSV_PATH
- `_get_template_machine_names()` 
- `_safe_csv_prefix(account_name)`
- `prefix_csv_server_names(csv_text, prefix)`
- `_csv_template_hash()`
- `snap_budget_to_tier(annual_budget)` → int
- `load_tier_cache()` → dict
- `save_tier_cache(cache)` → None
- `learn_tier_machine_selections(assessed_machines, account_prefix, progress)` → dict
- `get_machine_ids_for_tier(tier, machines, account_prefix, cache)` → list
- `_strip_account_prefix(display_name, prefix)`
- `_assessed_machine_monthly_cost(machine)`

- [ ] **步骤 3：Commit**

```bash
git add poeworkflow-refactor/budget/
git commit -m "feat: add budget parser and tier modules"
```

---

### 任务 5：创建 azure/ 模块

**文件：**
- 创建：`poeworkflow-refactor/azure/auth.py`
- 创建：`poeworkflow-refactor/azure/arm.py`
- 创建：`poeworkflow-refactor/azure/migrate.py`

- [ ] **步骤 1：创建 azure/auth.py**

从 app.py 提取（行 1190-1403）：
- `_get_msal_client_id()` → 内部
- `_is_azure_token_valid()` → `is_azure_token_valid`
- `msal_device_code_login()`
- `clear_azure_login()`
- `list_azure_subscriptions(token)`
- `list_azure_resource_groups(subscription_id, token)`
- `_subscription_label(sub)` / `_resource_group_label(rg)`

import streamlit as st，因为 msal_device_code_login 需要写 session_state。

- [ ] **步骤 2：创建 azure/arm.py**

从 app.py 提取（行 1249-1397）：
- `_format_arm_error(response)`
- `_retry_after_seconds(response, default=10)`
- `_poll_azure_lro(operation_url, token, ...)`
- `azure_arm_request(method, path_or_url, token, ...)`
- `azure_arm_list(path, token)`
- `register_azure_provider(subscription_id, namespace, token)`

此模块不 import streamlit。

- [ ] **步骤 3：创建 azure/migrate.py**

从 app.py 提取（行 1405-3344），这是最大的模块：
- 所有 `_safe_azure_name`, `_migrate_project_path` 等辅助函数
- `register_migrate_tool`, `put_migrate_solution`, `ensure_migrate_solution`, `ensure_portal_menu_solutions`
- `resolve_migrate_project_location`
- `wait_for_import_site_import`, `wait_for_project_machines`, `wait_for_assessment_complete`, `wait_for_group_machine_membership`, `wait_for_portal_inventory_summary`
- `fix_assessment_excel_timestamps`
- `download_assessment_report`
- `_build_assessment_body`, `_extract_dominant_region`, `_assessment_cost_component`, `assessment_monthly_total_cost`, `_assessment_settings_snapshot`
- `tune_assessment_to_budget`
- `list_imported_machines`
- `run_azure_migrate_assessment(...)` → dict — 主入口

import 使用：
```python
from ..config import (
    AZURE_MANAGEMENT_ENDPOINT, AZURE_MIGRATE_API_VERSION,
    AZURE_MIGRATE_PROJECTS_API_VERSION, AZURE_OFFAZURE_API_VERSION,
    # ... 等其他常量
    BUILTIN_CSV_PATH, BUDGET_TIERS,
)
from .arm import azure_arm_request, azure_arm_list
from ..budget.parser import parse_annual_budget_usd, _format_usd
from ..budget.tier import (
    load_tier_cache, save_tier_cache, learn_tier_machine_selections,
    get_machine_ids_for_tier, snap_budget_to_tier,
    load_builtin_csv_template, prefix_csv_server_names,
    _get_template_machine_names, _safe_csv_prefix,
)
```

- [ ] **步骤 4：Commit**

```bash
git add poeworkflow-refactor/azure/
git commit -m "feat: add azure auth, arm, and migrate modules"
```

---

### 任务 6：创建 pipeline.py

**文件：**
- 创建：`poeworkflow-refactor/pipeline.py`

- [ ] **步骤 1：创建 pipeline.py**

从 app.py 提取编排层函数（行 1516-3742 中的非 UI 部分）：
- `_date_prefix()` → `date_prefix()`
- `generate_solution_artifact(doc_type, customer_name, account_name, customer_bg, solution_ref, infra_ref)` → dict（行 1516-1549）
- `generate_pov_artifact(...)` → dict（行 1552-1568）
- `create_solution_artifact_from_text(...)` → dict（行 3494-3506）
- `create_pov_artifact_from_text(...)` → dict（行 3509-3519）
- `get_existing_solution_text(current_doc_type)` → Optional[str]（行 3486-3491）
- `get_generated_migrate_csv_text()` → Optional[str]（行 3522-3526）
- `resolve_auto_inventory_csv(uploaded_inventory)` → tuple（行 3529-3535）
- `format_auto_poe_log(message)` → str（行 3538-3552）
- `should_display_auto_poe_log(message)` → bool（行 3555-3592）
- `run_full_auto_poe(...)` → dict（行 3595-3742）— 主编排函数
- `create_poe_zip(artifacts)` → bytes（行 3478-3483）

import 使用：
```python
from .llm.client import call_azure_openai
from .llm.prompts import SOLUTION_SYSTEM_PROMPT, INFRA_SYSTEM_PROMPT, POV_SYSTEM_PROMPT
from .documents.solution import create_solution_docx, generate_svg_architecture
from .documents.infra import create_infra_docx
from .documents.pov import create_pov_docx, build_pov_prompt, has_meaningful_pov_team
from .budget.parser import parse_annual_budget_usd
from .budget.tier import snap_budget_to_tier
from .config import BUDGET_TIERS
```

- [ ] **步骤 2：Commit**

```bash
git add poeworkflow-refactor/pipeline.py
git commit -m "feat: add pipeline orchestrator module"
```

---

### 任务 7：创建 session.py

**文件：**
- 创建：`poeworkflow-refactor/session.py`

- [ ] **步骤 1：创建 session.py**

从 app.py 提取（行 2215-2276）：
- `_get_session_persist_path()`
- `_cleanup_stale_sessions(max_age_hours=24)`
- `persist_session_state()`
- `restore_session_state()`
- `clear_session_persist()`

```python
"""Session 状态持久化"""
import json
import os
import hashlib
import time
import streamlit as st
from .config import PERSIST_DIR, _PERSIST_KEYS


def _get_session_persist_path():
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        ctx = get_script_run_ctx()
        if ctx and ctx.session_id:
            sid = hashlib.md5(ctx.session_id.encode()).hexdigest()[:8]
        else:
            sid = "default"
    except Exception:
        sid = "default"
    return os.path.join(PERSIST_DIR, f".session_persist_{sid}.json")


def _cleanup_stale_sessions(max_age_hours=24):
    try:
        cutoff = time.time() - max_age_hours * 3600
        for fname in os.listdir(PERSIST_DIR):
            if fname.startswith(".session_persist_") and fname.endswith(".json"):
                fpath = os.path.join(PERSIST_DIR, fname)
                if os.path.getmtime(fpath) < cutoff:
                    os.remove(fpath)
    except Exception:
        pass


def persist_session_state():
    data = {}
    for k in _PERSIST_KEYS:
        if k in st.session_state:
            data[k] = st.session_state[k]
    if not data:
        return
    try:
        path = _get_session_persist_path()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception:
        pass


def restore_session_state():
    _cleanup_stale_sessions()
    path = _get_session_persist_path()
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return
    for k, v in data.items():
        if k not in st.session_state:
            st.session_state[k] = v


def clear_session_persist():
    try:
        os.remove(_get_session_persist_path())
    except OSError:
        pass
```

- [ ] **步骤 2：Commit**

```bash
git add poeworkflow-refactor/session.py
git commit -m "feat: add session persistence module"
```

---

### 任务 8：创建 ui/ 模块

**文件：**
- 创建：`poeworkflow-refactor/ui/auto_poe.py`
- 创建：`poeworkflow-refactor/ui/tabs.py`

- [ ] **步骤 1：创建 ui/auto_poe.py**

从 app.py 提取 `render_full_auto_poe_area(...)`（行 3745-4070）：
- 全自动 POE Tab 的完整 UI 渲染
- 包括 Azure 登录按钮、订阅/RG 选择器、评估名称、浏览器 Profile 状态、生成按钮、结果展示

import 使用：
```python
import streamlit as st
from ..azure.auth import is_azure_token_valid, msal_device_code_login, clear_azure_login, list_azure_subscriptions, list_azure_resource_groups, _subscription_label, _resource_group_label
from ..pipeline import run_full_auto_poe, format_auto_poe_log, should_display_auto_poe_log, get_existing_solution_text
from ..budget.parser import parse_annual_budget_usd, _format_usd
from ..budget.tier import snap_budget_to_tier, load_tier_cache, _get_template_machine_names
from ..config import BUILTIN_CSV_PATH
```

- [ ] **步骤 2：创建 ui/tabs.py**

从 app.py 提取 Tab 2-5 内容（行 4183-4790）：
- Tab 2（解决方案文档）：AI 生成 + 手动导入 + 下载 + 预览
- Tab 3（POV 部署计划）：生成 + 下载 + 预览
- Tab 4（Azure Migrate CSV）：上传 Excel → 生成 CSV → 下载预览
- Tab 5（年度价格表）：上传 → 添加 yearly cost → 下载

注意：此文件仅包含 Tab 内的渲染逻辑（`with tab_*:` 块内的内容），Tab 容器布局仍在 main.py。

或者，为保持简单，可以让 `ui/tabs.py` 暴露 4 个渲染函数：
```python
def render_solution_tab(customer_name, account_name, budget, customer_bg, solution_ref, infra_ref):
    """渲染解决方案文档 Tab 内容"""
    ...

def render_pov_tab(pov_ref):
    """渲染 POV Tab 内容"""
    ...

def render_csv_tab():
    """渲染 CSV Tab 内容"""
    ...

def render_yearly_tab():
    """渲染年度价格表 Tab 内容"""
    ...
```

- [ ] **步骤 3：Commit**

```bash
git add poeworkflow-refactor/ui/
git commit -m "feat: add ui modules (auto poe, tabs)"
```

---

### 任务 9：创建 main.py 和兼容壳

**文件：**
- 创建：`poeworkflow-refactor/main.py`
- 创建：`poeworkflow-refactor/app.py`（兼容壳）

- [ ] **步骤 1：创建 main.py**

新入口文件，~80 行：

```python
"""POE 自动生成工作流 — 主入口"""
import os
import streamlit as st
from frontend.ui import load_desktop_theme, render_app_header, render_pill, render_section_head, render_template_status

from .config import (
    PROJECT_DIR, SOLUTION_TEMPLATE_PATH, INFRA_TEMPLATE_PATH,
    POV_TEMPLATE_PATH, MIGRATE_TEMPLATE_PATH, check_secrets,
)
from .session import restore_session_state, persist_session_state, clear_session_persist
from .llm.prompts import extract_template_text
from .ui.auto_poe import render_full_auto_poe_area
from .ui.tabs import render_solution_tab, render_pov_tab, render_csv_tab, render_yearly_tab


def main():
    restore_session_state()

    # 页面配置
    st.set_page_config(
        page_title="POE 自动生成工作流",
        page_icon="P",
        layout="wide",
    )
    load_desktop_theme(PROJECT_DIR)
    render_app_header()

    if not check_secrets():
        st.stop()

    # 侧边栏
    with st.sidebar:
        st.markdown("### 操作")
        if st.button("清除所有结果", use_container_width=True):
            for key in [
                "solution_text", "infra_text", "pov_text", "customer_name",
                "account_name", "csv_code", "budget", "doc_type",
                "pov_source_doc_type", "yearly_excel_bytes", "yearly_excel_name",
                "yearly_messages", "auto_poe_zip_bytes", "auto_poe_zip_name",
                "auto_poe_result", "auto_poe_error", "pov_vendor_team",
            ]:
                st.session_state.pop(key, None)
            clear_session_persist()
            st.rerun()

        st.markdown("---")
        st.markdown("### 模板状态")
        sol_ok = os.path.exists(SOLUTION_TEMPLATE_PATH)
        infra_ok = os.path.exists(INFRA_TEMPLATE_PATH)
        pov_ok = os.path.exists(POV_TEMPLATE_PATH)
        csv_ok = os.path.exists(MIGRATE_TEMPLATE_PATH)
        render_template_status([
            ("AI Solution", sol_ok),
            ("Infra", infra_ok),
            ("POV", pov_ok),
            ("CSV", csv_ok),
        ])

    # 加载模板文本（用于 prompt 参考）
    solution_ref = extract_template_text(SOLUTION_TEMPLATE_PATH) if os.path.exists(SOLUTION_TEMPLATE_PATH) else ""
    infra_ref = extract_template_text(INFRA_TEMPLATE_PATH) if os.path.exists(INFRA_TEMPLATE_PATH) else ""
    pov_ref = extract_template_text(POV_TEMPLATE_PATH) if os.path.exists(POV_TEMPLATE_PATH) else ""

    # 公共输入区域
    render_section_head(
        "客户信息",
        "这些输入会贯穿方案文档、POV 计划、CSV 推导和 Azure Migrate 评估。",
        render_pill("必填项优先", "accent"),
    )
    c0, c1, c2 = st.columns([1.5, 2, 1])
    with c0:
        account_name = st.text_input("账户名", placeholder="例如：Tetherflow")
    with c1:
        customer_name = st.text_input("客户名称", placeholder="例如：宇宙无敌科技有限公司")
    with c2:
        budget = st.text_input("预估年消耗 (USD)", placeholder="例如：500k+")

    customer_bg = st.text_area(
        "客户背景信息",
        placeholder="粘贴客户背景资料...",
        height=125,
    )

    pov_dc1, pov_dc2, pov_dc3, pov_dc4 = st.columns([1, 1, 1, 1])
    with pov_dc1:
        pov_start = st.date_input("POV 开始日期", value=None, key="pov_start_date")
    with pov_dc2:
        pov_end = st.date_input("POV 结束日期", value=None, key="pov_end_date")
    with pov_dc3:
        pov_tech_lead = st.text_input("技术负责人", key="_pov_tech_lead")
    with pov_dc4:
        pov_architect = st.text_input("架构师", key="_pov_architect")
    _parts = []
    if (pov_tech_lead or "").strip():
        _parts.append(f"技术负责人: {pov_tech_lead.strip()}")
    if (pov_architect or "").strip():
        _parts.append(f"架构师: {pov_architect.strip()}")
    st.session_state["pov_vendor_team"] = ", ".join(_parts)

    st.divider()

    # Tab 布局
    tab_auto, tab_sol, tab_pov, tab_csv, tab_yearly = st.tabs([
        "全自动POE生成", "解决方案文档", "POV 部署计划", "Azure Migrate CSV", "年度价格表"
    ])

    with tab_auto:
        auto_doc_type_label = st.radio(
            "生成文档类型",
            ["AI 解决方案", "Infra 基础设施"],
            horizontal=True,
            key="auto_doc_type_radio",
        )
        auto_doc_type = "AI" if auto_doc_type_label == "AI 解决方案" else "Infra"
        render_full_auto_poe_area(
            current_doc_type=auto_doc_type,
            account_name=account_name,
            customer_name=customer_name,
            budget=budget,
            customer_bg=customer_bg,
            solution_ref=solution_ref,
            infra_ref=infra_ref,
            pov_ref=pov_ref,
        )

    with tab_sol:
        render_solution_tab(customer_name, account_name, budget, customer_bg, solution_ref, infra_ref)

    with tab_pov:
        render_pov_tab(pov_ref)

    with tab_csv:
        render_csv_tab()

    with tab_yearly:
        render_yearly_tab()


if __name__ == "__main__":
    main()
```

- [ ] **步骤 2：创建 poeworkflow-refactor/app.py（兼容壳）**

```python
"""兼容壳 — 保留原 app.py 入口路径"""
from .main import main

if __name__ == "__main__":
    main()
```

- [ ] **步骤 3：添加 requirements.txt 到 poeworkflow-refactor/**

更新 `poeworkflow-refactor/requirements.txt`（如果存在），确保依赖与主项目一致。

- [ ] **步骤 4：Commit**

```bash
git add poeworkflow-refactor/main.py poeworkflow-refactor/app.py
git commit -m "feat: add main entry point and compatibility shell"
```

---

### 任务 10：集成验证

- [ ] **步骤 1：检查所有 import 路径**

```bash
cd poeworkflow-refactor
python -c "
from config import get_secret, check_secrets, BUDGET_TIERS
from llm.client import call_azure_openai
from llm.prompts import SOLUTION_SYSTEM_PROMPT, extract_template_text
from documents.docx_utils import set_run_font, markdown_to_docx
from documents.solution import create_solution_docx
from documents.infra import create_infra_docx
from documents.pov import create_pov_docx, build_pov_prompt
from budget.parser import parse_annual_budget_usd
from budget.tier import snap_budget_to_tier, load_tier_cache
from session import persist_session_state, restore_session_state
print('All imports OK (except streamlit-dependent modules)')
"
```

- [ ] **步骤 2：运行 Streamlit 应用验证**

```bash
cd "/Users/penn/Documents/POE workflow"
streamlit run poeworkflow-refactor/main.py --server.headless true --server.port 8502 &
sleep 5
curl -s http://localhost:8502 | head -20
kill %1
```

- [ ] **步骤 3：运行现有测试确认无回归**

```bash
cd "/Users/penn/Documents/POE workflow"
python -m pytest test_*.py -v --timeout=30 2>&1 | tail -30
```

- [ ] **步骤 4：最终 Commit**

```bash
git add -A
git commit -m "chore: integration verification, all imports working"
```

---

## 注意事项

1. **不要修改原 app.py** — 所有新文件在 `poeworkflow-refactor/` 中
2. **完整复制代码** — 从 app.py 复制时不改变任何业务逻辑，只更新 import 路径
3. **保留原 `_` 前缀函数的逻辑** — 仅在变为公开 API 时去掉前缀
4. **模板文件和 .streamlit/ 仍在原项目目录** — config.py 通过 PROJECT_DIR 引用
5. **测试文件不需要移动** — 验证时直接运行原目录中的测试
