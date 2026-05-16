#!/usr/bin/env python3
"""
POE CLI Generator — 无需 Streamlit 即可生成全套 POE 文档

用法:
    python cli_generate.py --config config.json
    python cli_generate.py --config config.json --output ./output

config.json 示例:
{
    "customer_name": "深圳跃瓦创新科技",
    "account_name": "Tetherflow",
    "doc_type": "AI",
    "customer_background": "...",
    "annual_budget_usd": "50000",
    "pov_start": "2026-06-01",
    "pov_end": "2026-06-15",
    "vendor_team": "技术负责人：张三\\n架构师：李四",
    "azure_subscription_id": "",
    "azure_resource_group": "",
    "skip_azure_migrate": false
}

环境变量或 .streamlit/secrets.toml 中需配置:
    AZURE_OPENAI_KEY, AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_DEPLOYMENT
"""

import argparse
import datetime
import json
import os
import sys
import time

# 将项目根目录加入路径
APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

# ──────────────────────────────────────────────
# Mock Streamlit 以便直接 import app.py 中的函数
# ──────────────────────────────────────────────
import types

_mock_st = types.ModuleType("streamlit")
_mock_st.session_state = {}
_mock_st.set_page_config = lambda **kw: None
_mock_st.markdown = lambda *a, **kw: None
_mock_st.cache_data = lambda func=None, **kw: func if func else (lambda f: f)
_mock_st.error = lambda *a, **kw: None
_mock_st.info = lambda *a, **kw: None
_mock_st.warning = lambda *a, **kw: None

# mock secrets
class _MockSecrets:
    def __init__(self):
        self._data = {}
        self._load()

    def _load(self):
        secrets_path = os.path.join(APP_DIR, ".streamlit", "secrets.toml")
        if os.path.exists(secrets_path):
            import re
            with open(secrets_path, "r", encoding="utf-8") as f:
                for line in f:
                    m = re.match(r'^(\w+)\s*=\s*"(.+)"', line.strip())
                    if m:
                        self._data[m.group(1)] = m.group(2)

    def get(self, key, default=None):
        return self._data.get(key, default)

    def __getattr__(self, key):
        if key.startswith("_"):
            return object.__getattribute__(self, key)
        return self._data.get(key)

_mock_st.secrets = _MockSecrets()

# mock components
_mock_components = types.ModuleType("streamlit.components")
_mock_components_v1 = types.ModuleType("streamlit.components.v1")
_mock_components_v1.html = lambda *a, **kw: None
_mock_components.v1 = _mock_components_v1

sys.modules["streamlit"] = _mock_st
sys.modules["streamlit.components"] = _mock_components
sys.modules["streamlit.components.v1"] = _mock_components_v1

# Mock frontend.ui 模块
_mock_ui = types.ModuleType("frontend.ui")
_mock_ui.load_desktop_theme = lambda *a, **kw: None
_mock_ui.render_app_header = lambda *a, **kw: None
_mock_ui.render_pill = lambda *a, **kw: None
_mock_ui.render_readiness = lambda *a, **kw: None
_mock_ui.render_auto_poe_result = lambda *a, **kw: None
_mock_ui.render_device_code_login = lambda *a, **kw: None
_mock_ui.render_section_head = lambda *a, **kw: None
_mock_ui.render_template_status = lambda *a, **kw: None
_mock_ui.render_workflow_steps = lambda *a, **kw: None
sys.modules["frontend"] = types.ModuleType("frontend")
sys.modules["frontend.ui"] = _mock_ui

# 现在可以安全 import app.py 中的功能
import app  # noqa: E402


# ──────────────────────────────────────────────
# CLI 主逻辑
# ──────────────────────────────────────────────
def progress_print(msg: str):
    """进度回调：打印到终端。"""
    print(f"  → {msg}")


def do_azure_login() -> str:
    """通过 MSAL Device Code Flow 登录 Azure，返回 access_token。"""
    try:
        import msal
    except ImportError:
        raise RuntimeError("缺少 msal 依赖，请 pip install msal")

    client_id = app.get_secret("MSAL_CLIENT_ID", app.MSAL_CLIENT_ID_DEFAULT)
    msal_app = msal.PublicClientApplication(client_id, authority=app.AZURE_AUTHORITY)
    flow = msal_app.initiate_device_flow(scopes=app.AZURE_ARM_SCOPE)

    if "user_code" not in flow:
        raise RuntimeError(f"无法启动 Microsoft 登录流程: {flow}")

    print("\n" + "=" * 60)
    print(f"  请在浏览器中打开: {flow.get('verification_uri', 'https://microsoft.com/devicelogin')}")
    print(f"  输入代码: {flow['user_code']}")
    print("=" * 60 + "\n")
    print("等待登录完成...")

    result = msal_app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or "登录失败"
        raise RuntimeError(f"Azure 登录失败: {err}")

    account = result.get("id_token_claims", {}) or {}
    username = account.get("preferred_username") or account.get("email") or "Azure 用户"
    print(f"  ✅ 登录成功: {username}")
    return result["access_token"]


def list_subscriptions(token: str) -> list:
    """列出用户的 Azure 订阅。"""
    result = app.azure_arm_request("GET", "/subscriptions?api-version=2022-12-01", token)
    return result.get("value", [])


def list_resource_groups(token: str, subscription_id: str) -> list:
    """列出指定订阅下的资源组。"""
    result = app.azure_arm_request(
        "GET",
        f"/subscriptions/{subscription_id}/resourcegroups?api-version=2021-04-01",
        token,
    )
    return result.get("value", [])


def interactive_select(items: list, label_key: str, prompt_text: str) -> dict:
    """交互式选择列表项。"""
    print(f"\n{prompt_text}:")
    for i, item in enumerate(items):
        print(f"  [{i + 1}] {item.get(label_key, item.get('name', str(i)))}")
    while True:
        choice = input(f"\n请输入编号 (1-{len(items)}): ").strip()
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(items):
                return items[idx]
        except ValueError:
            pass
        print("  无效选择，请重试。")


def run_cli(config_path: str, output_dir: str):
    """执行完整 POE 生成流程。"""
    # 加载配置
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    customer_name = config["customer_name"]
    account_name = config["account_name"]
    doc_type = config.get("doc_type", "AI")
    customer_bg = config["customer_background"]
    annual_budget_text = config.get("annual_budget_usd", "")
    pov_start_str = config.get("pov_start", "")
    pov_end_str = config.get("pov_end", "")
    vendor_team = config.get("vendor_team", "")
    skip_migrate = config.get("skip_azure_migrate", False)
    subscription_id = config.get("azure_subscription_id", "")
    resource_group = config.get("azure_resource_group", "")

    # 解析日期
    pov_start = datetime.date.fromisoformat(pov_start_str) if pov_start_str else None
    pov_end = datetime.date.fromisoformat(pov_end_str) if pov_end_str else None

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    # 加载模板参考文本
    print("\n📄 加载模板参考文本...")
    solution_ref = ""
    infra_ref = ""
    pov_ref = ""
    try:
        if os.path.exists(app.SOLUTION_TEMPLATE_PATH):
            solution_ref = app.extract_template_text.__wrapped__(app.SOLUTION_TEMPLATE_PATH) if hasattr(app.extract_template_text, '__wrapped__') else app.extract_template_text(app.SOLUTION_TEMPLATE_PATH)
    except Exception:
        pass
    try:
        if os.path.exists(app.INFRA_TEMPLATE_PATH):
            infra_ref = app.extract_template_text.__wrapped__(app.INFRA_TEMPLATE_PATH) if hasattr(app.extract_template_text, '__wrapped__') else app.extract_template_text(app.INFRA_TEMPLATE_PATH)
    except Exception:
        pass
    try:
        if os.path.exists(app.POV_TEMPLATE_PATH):
            pov_ref = app.extract_template_text.__wrapped__(app.POV_TEMPLATE_PATH) if hasattr(app.extract_template_text, '__wrapped__') else app.extract_template_text(app.POV_TEMPLATE_PATH)
    except Exception:
        pass

    # ── 步骤 1: 生成解决方案架构文档 ──
    print("\n📝 步骤 1/4: 生成解决方案架构文档...")
    solution_artifact = app.generate_solution_artifact(
        doc_type, customer_name, account_name, customer_bg, solution_ref, infra_ref
    )
    print(f"  ✅ 生成完成: {solution_artifact['file_name']}")

    # 保存 .docx
    sol_path = os.path.join(output_dir, solution_artifact["file_name"])
    with open(sol_path, "wb") as f:
        f.write(solution_artifact["bytes"])
    # 保存 Markdown
    sol_md_path = os.path.join(output_dir, solution_artifact["file_name"].replace(".docx", ".md"))
    with open(sol_md_path, "w", encoding="utf-8") as f:
        f.write(solution_artifact["content"])
    print(f"  📁 已保存: {sol_path}")
    print(f"  📁 已保存: {sol_md_path}")

    # ── 步骤 2: 生成 POV 部署计划 ──
    print("\n📝 步骤 2/4: 生成 POV 部署计划...")
    if not pov_start or not pov_end:
        print("  ⚠️ 未提供 POV 日期，跳过 POV 生成。")
        pov_artifact = None
    elif not vendor_team.strip():
        print("  ⚠️ 未提供乙方项目人员，跳过 POV 生成。")
        pov_artifact = None
    else:
        pov_artifact = app.generate_pov_artifact(
            solution_artifact["content"],
            customer_name,
            account_name,
            pov_ref,
            pov_start,
            pov_end,
            vendor_team,
        )
        print(f"  ✅ 生成完成: {pov_artifact['file_name']}")

        pov_path = os.path.join(output_dir, pov_artifact["file_name"])
        with open(pov_path, "wb") as f:
            f.write(pov_artifact["bytes"])
        pov_md_path = os.path.join(output_dir, pov_artifact["file_name"].replace(".docx", ".md"))
        with open(pov_md_path, "w", encoding="utf-8") as f:
            f.write(pov_artifact["content"])
        print(f"  📁 已保存: {pov_path}")
        print(f"  📁 已保存: {pov_md_path}")

    # ── 步骤 3: 生成 Azure Migrate CSV ──
    print("\n📝 步骤 3/4: 生成 Azure Migrate CSV...")
    csv_text_raw = app.load_builtin_csv_template()
    safe_prefix = app._safe_csv_prefix(account_name)
    csv_text = app.prefix_csv_server_names(csv_text_raw, safe_prefix)
    csv_path = os.path.join(output_dir, f"{account_name}-Azure migrate report.csv")
    with open(csv_path, "w", encoding="utf-8-sig") as f:
        f.write(csv_text)
    print(f"  ✅ 已保存: {csv_path}")

    # ── 步骤 4: Azure Migrate 评估 ──
    if skip_migrate:
        print("\n📝 步骤 4/4: Azure Migrate 评估 — 已跳过 (skip_azure_migrate=true)")
        assessment_artifact = None
    else:
        print("\n📝 步骤 4/4: Azure Migrate 评估...")
        print("  需要登录 Azure 账号以创建 Migrate 项目并运行评估。")

        # Azure 登录
        token = do_azure_login()

        # 选择订阅
        if not subscription_id:
            subs = list_subscriptions(token)
            if not subs:
                print("  ❌ 未找到可用 Azure 订阅。")
                assessment_artifact = None
            else:
                selected_sub = interactive_select(subs, "displayName", "选择 Azure 订阅")
                subscription_id = selected_sub["subscriptionId"]
                print(f"  已选择订阅: {selected_sub['displayName']} ({subscription_id})")

        # 选择资源组
        if subscription_id and not resource_group:
            rgs = list_resource_groups(token, subscription_id)
            if not rgs:
                print("  ❌ 未找到可用资源组。")
                assessment_artifact = None
            else:
                selected_rg = interactive_select(rgs, "name", "选择资源组")
                resource_group = selected_rg["name"]
                print(f"  已选择资源组: {resource_group}")

        if subscription_id and resource_group:
            assessment_name = f"poe-{account_name}-assess"
            try:
                migrate_result = app.run_azure_migrate_assessment(
                    token=token,
                    subscription_id=subscription_id,
                    resource_group=resource_group,
                    account_name=account_name,
                    assessment_name=assessment_name,
                    annual_budget_text=annual_budget_text,
                    progress=progress_print,
                )
                excel_bytes = migrate_result["excel_bytes"]

                # 修正时间戳
                if pov_start and pov_end:
                    try:
                        excel_bytes = app.fix_assessment_excel_timestamps(excel_bytes, pov_start, pov_end)
                        print("  ✅ 已修正评估报告时间戳至 POV 区间内")
                    except Exception:
                        pass

                assess_path = os.path.join(output_dir, f"{account_name}-Azure Migrate Assessment.xlsx")
                with open(assess_path, "wb") as f:
                    f.write(excel_bytes)
                print(f"  ✅ 已保存: {assess_path}")
                assessment_artifact = {"file_name": os.path.basename(assess_path), "bytes": excel_bytes}
            except Exception as e:
                print(f"  ❌ Azure Migrate 评估失败: {e}")
                assessment_artifact = None
        else:
            assessment_artifact = None

    # ── 打包 ZIP ──
    print("\n📦 打包 POE 套件...")
    zip_artifacts = [solution_artifact]
    if pov_artifact:
        zip_artifacts.append(pov_artifact)
    if assessment_artifact:
        zip_artifacts.append(assessment_artifact)

    zip_bytes = app.create_poe_zip(zip_artifacts)
    zip_path = os.path.join(output_dir, f"{account_name}-POE-Complete.zip")
    with open(zip_path, "wb") as f:
        f.write(zip_bytes)
    print(f"  ✅ ZIP 已保存: {zip_path}")

    # 总结
    print("\n" + "=" * 60)
    print("🎉 POE 文档生成完成！")
    print(f"   输出目录: {output_dir}")
    print("   生成文件:")
    for fname in os.listdir(output_dir):
        fpath = os.path.join(output_dir, fname)
        size_kb = os.path.getsize(fpath) / 1024
        print(f"     • {fname} ({size_kb:.1f} KB)")
    print("=" * 60)


def main():
    parser = argparse.ArgumentParser(description="POE 文档 CLI 生成工具")
    parser.add_argument("--config", required=True, help="JSON 配置文件路径")
    parser.add_argument("--output", default="./poe-output", help="输出目录（默认 ./poe-output）")
    args = parser.parse_args()

    if not os.path.exists(args.config):
        print(f"❌ 配置文件不存在: {args.config}")
        sys.exit(1)

    run_cli(args.config, args.output)


if __name__ == "__main__":
    main()
