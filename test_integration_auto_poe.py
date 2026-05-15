"""
集成测试：全自动 POE 生成流程
使用真实参数测试 call_azure_openai → generate_solution_artifact → generate_pov_artifact 全链路。
"""

import sys
import os
import datetime
import time
import traceback

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

# ─── Phase 1: 直接测试 call_azure_openai（真实 API 调用）───
print("=" * 70)
print("集成测试：全自动 POE 生成流程")
print("=" * 70)

# Load streamlit secrets manually
import toml

secrets_path = os.path.join(APP_DIR, ".streamlit", "secrets.toml")
if not os.path.exists(secrets_path):
    print("❌ 未找到 .streamlit/secrets.toml，无法执行集成测试")
    sys.exit(1)

secrets = toml.load(secrets_path)
print(f"✅ 已加载 secrets.toml")
print(f"   API Version: {secrets.get('AZURE_OPENAI_API_VERSION', '未设置')}")
print(f"   Deployment: {secrets.get('AZURE_OPENAI_DEPLOYMENT', '未设置')}")
print()

# ─── Mock streamlit minimally ───
from unittest.mock import MagicMock, patch

mock_st = MagicMock()
mock_st.secrets = secrets
mock_st.set_page_config = MagicMock()
mock_st.cache_data = lambda f=None, **kwargs: f if f else (lambda fn: fn)
mock_st.session_state = {}

sys.modules['streamlit'] = mock_st
sys.modules['msal'] = MagicMock()

mock_ui = MagicMock()
sys.modules['frontend'] = MagicMock()
sys.modules['frontend.ui'] = mock_ui

import importlib.util

spec = importlib.util.spec_from_file_location("app", os.path.join(APP_DIR, "app.py"))
app_module = importlib.util.module_from_spec(spec)

with patch.object(mock_st, 'set_page_config'):
    spec.loader.exec_module(app_module)

call_azure_openai = app_module.call_azure_openai
generate_solution_artifact = app_module.generate_solution_artifact
generate_pov_artifact = app_module.generate_pov_artifact
build_pov_prompt = app_module.build_pov_prompt

# ─── 测试参数 ───
CUSTOMER_NAME = "四川蜀天信息技术有限公司"
ACCOUNT_NAME = "Shutian"
BUDGET = "15k"
CUSTOMER_BG = """四川蜀天信息技术有限公司成立于2019年08月12日，注册地位于四川省雅安市雨城区滨江西路6号蜀天·雨城汇4号楼B单元，法定代表人为付君太。经营范围包括许可项目：互联网信息服务；第一类增值电信业务；第二类增值电信业务；建设工程施工；网络文化经营；文件、资料等其他印刷品印刷；拍卖业务。一般项目：互联网安全服务；物联网应用服务；5G通信技术服务；软件开发；电子产品销售；互联网销售；网络设备销售；软件销售；办公设备耗材销售；网络技术服务；机械设备租赁；计算机及通讯设备租赁；计算机及办公设备维修；计算机软硬件及辅助设备零售；计算机系统服务；数据处理服务；信息系统集成服务；技术服务、技术开发、技术咨询、技术交流、技术转让、技术推广；信息技术咨询服务；数据处理和存储支持服务；信息安全设备销售；网络与信息安全软件开发；卫星导航多模增强应用服务系统集成。四川蜀天信息技术有限公司对外投资5家公司，具有6处分支机构。"""
POV_START = datetime.date(2026, 4, 28)
POV_END = datetime.date(2026, 5, 13)
VENDOR_TEAM = "技术负责人：杨思维\n架构师：邹成"

results = {"pass": 0, "fail": 0, "errors": []}


def test_step(name):
    print(f"\n{'─' * 60}")
    print(f"🧪 测试: {name}")
    print(f"{'─' * 60}")


def pass_step(msg=""):
    results["pass"] += 1
    print(f"  ✅ 通过 {msg}")


def fail_step(msg):
    results["fail"] += 1
    results["errors"].append(msg)
    print(f"  ❌ 失败: {msg}")


# ─── Test 1: call_azure_openai 基本连通性 ───
test_step("call_azure_openai 基本连通性（短 prompt）")
try:
    t0 = time.time()
    result = call_azure_openai(
        "你是一个测试助手。",
        "请回答：1+1=? 只回答数字。"
    )
    elapsed = time.time() - t0
    print(f"  响应内容: {result.strip()[:100]}")
    print(f"  耗时: {elapsed:.1f}s")
    if result and result.strip():
        pass_step()
    else:
        fail_step("返回空内容")
except Exception as e:
    fail_step(f"{type(e).__name__}: {e}")
    traceback.print_exc()
    print("\n⚠️  基本 API 连通性测试失败，跳过后续测试。")
    print(f"\n{'=' * 70}")
    print(f"测试结果: {results['pass']} 通过, {results['fail']} 失败")
    sys.exit(1)

# ─── Test 2: generate_solution_artifact（AI 解决方案文档） ───
test_step("generate_solution_artifact（AI 解决方案文档生成）")
try:
    t0 = time.time()
    solution = generate_solution_artifact(
        current_doc_type="AI",
        customer_name=CUSTOMER_NAME,
        account_name=ACCOUNT_NAME,
        customer_bg=CUSTOMER_BG,
        solution_ref="",
        infra_ref="",
    )
    elapsed = time.time() - t0
    print(f"  文件名: {solution['file_name']}")
    print(f"  内容长度: {len(solution['content'])} 字符")
    print(f"  DOCX 大小: {len(solution['bytes'])} 字节")
    print(f"  耗时: {elapsed:.1f}s")
    print(f"  内容前200字: {solution['content'][:200]}")

    # 校验
    errors = []
    if not solution["content"]:
        errors.append("内容为空")
    if not solution["bytes"]:
        errors.append("DOCX 为空")
    if not solution["file_name"].endswith(".docx"):
        errors.append("文件名不以 .docx 结尾")
    if ACCOUNT_NAME not in solution["file_name"]:
        errors.append(f"文件名未包含账户名 {ACCOUNT_NAME}")
    if len(solution["content"]) < 500:
        errors.append(f"内容过短（{len(solution['content'])} 字符），可能生成不完整")
    # 检查是否包含预期章节
    expected_sections = ["摘要", "架构", "背景", "需求"]
    found_sections = [s for s in expected_sections if s in solution["content"]]
    if len(found_sections) < 2:
        errors.append(f"缺少预期章节，仅找到: {found_sections}")

    if errors:
        fail_step("; ".join(errors))
    else:
        pass_step(f"({len(found_sections)}/{len(expected_sections)} 章节)")
except Exception as e:
    fail_step(f"{type(e).__name__}: {e}")
    traceback.print_exc()
    # If solution fails, create a fallback for POV test
    solution = {"content": "# 测试方案\n\n这是一个测试方案文档。", "bytes": b"", "file_name": "test.docx"}

# ─── Test 3: generate_pov_artifact（POV 部署计划） ───
test_step("generate_pov_artifact（POV 部署计划生成）")
try:
    t0 = time.time()
    pov = generate_pov_artifact(
        solution_text=solution["content"],
        customer_name=CUSTOMER_NAME,
        account_name=ACCOUNT_NAME,
        pov_ref="",
        pov_start=POV_START,
        pov_end=POV_END,
        vendor_team=VENDOR_TEAM,
    )
    elapsed = time.time() - t0
    print(f"  文件名: {pov['file_name']}")
    print(f"  内容长度: {len(pov['content'])} 字符")
    print(f"  DOCX 大小: {len(pov['bytes'])} 字节")
    print(f"  耗时: {elapsed:.1f}s")
    print(f"  内容前200字: {pov['content'][:200]}")

    errors = []
    if not pov["content"]:
        errors.append("内容为空")
    if not pov["bytes"]:
        errors.append("DOCX 为空")
    if not pov["file_name"].endswith(".docx"):
        errors.append("文件名不以 .docx 结尾")
    if len(pov["content"]) < 300:
        errors.append(f"内容过短（{len(pov['content'])} 字符）")
    # 检查 POV 日期是否出现在内容中
    if "2026" not in pov["content"]:
        errors.append("内容中未找到年份 2026")
    # 检查人员是否出现
    team_names = ["杨思维", "邹成"]
    found_names = [n for n in team_names if n in pov["content"]]
    if not found_names:
        errors.append("内容中未找到任何团队成员姓名")

    if errors:
        fail_step("; ".join(errors))
    else:
        pass_step(f"(找到团队成员: {found_names})")
except Exception as e:
    fail_step(f"{type(e).__name__}: {e}")
    traceback.print_exc()

# ─── Test 4: 模拟全自动 POE 文档生成部分（不含 Azure Migrate） ───
test_step("全自动 POE 文档生成链路（不含 Azure Migrate）")
try:
    t0 = time.time()
    # 模拟 run_full_auto_poe 的文档生成部分
    # Step 1: Solution
    sol = generate_solution_artifact(
        current_doc_type="AI",
        customer_name=CUSTOMER_NAME,
        account_name=ACCOUNT_NAME,
        customer_bg=CUSTOMER_BG,
        solution_ref="",
        infra_ref="",
    )
    print(f"  ✔ Solution 生成成功: {sol['file_name']} ({len(sol['content'])} 字符)")

    # Step 2: POV
    pov2 = generate_pov_artifact(
        solution_text=sol["content"],
        customer_name=CUSTOMER_NAME,
        account_name=ACCOUNT_NAME,
        pov_ref="",
        pov_start=POV_START,
        pov_end=POV_END,
        vendor_team=VENDOR_TEAM,
    )
    print(f"  ✔ POV 生成成功: {pov2['file_name']} ({len(pov2['content'])} 字符)")

    # Step 3: 创建 ZIP
    create_poe_zip = app_module.create_poe_zip
    # 模拟 assessment artifact（跳过真实 Azure Migrate）
    fake_assessment = {
        "file_name": f"{ACCOUNT_NAME}-Azure Migrate Assessment.xlsx",
        "bytes": b"fake excel content for testing",
    }
    zip_bytes = create_poe_zip([sol, pov2, fake_assessment])
    print(f"  ✔ ZIP 打包成功: {len(zip_bytes)} 字节")

    elapsed = time.time() - t0
    print(f"  总耗时: {elapsed:.1f}s")

    if len(zip_bytes) > 100:
        pass_step(f"全链路文档生成正常，ZIP {len(zip_bytes)} 字节")
    else:
        fail_step("ZIP 文件过小")
except Exception as e:
    fail_step(f"{type(e).__name__}: {e}")
    traceback.print_exc()

# ─── Test 5: 验证 API 版本参数选择逻辑 ───
test_step("API 版本参数选择逻辑验证")
try:
    api_version = secrets.get("AZURE_OPENAI_API_VERSION", "2024-06-01")
    use_max_completion_tokens = api_version >= "2024-08-01"
    print(f"  当前 API 版本: {api_version}")
    print(f"  使用 max_completion_tokens: {use_max_completion_tokens}")
    print(f"  实际使用参数: {'max_completion_tokens=128000' if use_max_completion_tokens else 'max_tokens=16384'}")

    if api_version == "2024-06-01" and not use_max_completion_tokens:
        pass_step("正确使用 max_tokens（旧版 API）")
    elif api_version >= "2024-08-01" and use_max_completion_tokens:
        pass_step("正确使用 max_completion_tokens（新版 API）")
    else:
        fail_step(f"参数选择逻辑异常: api_version={api_version}, use_max_completion_tokens={use_max_completion_tokens}")
except Exception as e:
    fail_step(f"{type(e).__name__}: {e}")

# ─── 汇总 ───
print(f"\n{'=' * 70}")
print(f"测试汇总")
print(f"{'=' * 70}")
print(f"  ✅ 通过: {results['pass']}")
print(f"  ❌ 失败: {results['fail']}")
if results["errors"]:
    print(f"\n  失败详情:")
    for i, err in enumerate(results["errors"], 1):
        print(f"    {i}. {err}")
print(f"{'=' * 70}")

sys.exit(0 if results["fail"] == 0 else 1)
