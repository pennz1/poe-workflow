"""Workflow orchestration for POE generation."""

import datetime
import io
import re
import zipfile
from typing import Any, Callable, Dict, List, Optional

import streamlit as st

from azure.migrate import (
    _extract_dominant_region,
    fix_assessment_excel_timestamps,
    run_azure_migrate_assessment,
)
from budget.parser import parse_annual_budget_usd
from documents.infra import create_infra_docx
from documents.pov import build_pov_prompt, create_pov_docx, has_meaningful_pov_team
from documents.solution import create_solution_docx
from llm.client import call_azure_openai
from llm.prompts import POV_SYSTEM_PROMPT, build_infra_system_prompt, build_solution_system_prompt

def date_prefix():
    """返回当前日期前缀，如 0225"""
    return datetime.date.today().strftime("%m%d")


# ──────────────────────────────────────────────
# 全自动 POE：MSAL 登录与 Azure ARM API
# ──────────────────────────────────────────────

def generate_solution_artifact(
    current_doc_type: str,
    customer_name: str,
    account_name: str,
    customer_bg: str,
    solution_ref: str,
    infra_ref: str,
    annual_budget: float = 0,
) -> Dict[str, Any]:
    is_large_customer = annual_budget >= 100_000
    system_prompt = (
        build_solution_system_prompt(is_large_customer=is_large_customer)
        if current_doc_type == "AI"
        else build_infra_system_prompt(is_large_customer=is_large_customer)
    )
    ref_text = solution_ref if current_doc_type == "AI" else infra_ref
    user_ctx = (
        f"## 客户信息\n- **客户名称**：{customer_name}\n\n"
        f"## 客户背景\n{customer_bg}"
    )
    if is_large_customer:
        user_ctx += "\n\n该客户为年度预算超10万美元的大客户，请提供更详尽、更严谨的架构方案。"
    if ref_text:
        user_ctx += (
            "\n\n---\n\n## 【参考模板文档 —— 请学习其风格和结构，不要照抄具体数据】\n\n"
            f"{ref_text}"
        )

    content = call_azure_openai(system_prompt, user_ctx)
    if current_doc_type == "AI":
        docx_bytes = create_solution_docx(content=content, customer_name=customer_name)
        file_name = f"{account_name}-Solution Architecture.docx"
    else:
        docx_bytes = create_infra_docx(content=content, customer_name=customer_name)
        file_name = f"{account_name}-Infra Solution Architecture.docx"

    result = {"content": content, "bytes": docx_bytes, "file_name": file_name}
    return result


def generate_pov_artifact(
    solution_text: str,
    customer_name: str,
    account_name: str,
    pov_ref: str,
    pov_start: datetime.date,
    pov_end: datetime.date,
    vendor_team: str,
) -> Dict[str, Any]:
    pov_prompt = build_pov_prompt(solution_text, customer_name, pov_start, pov_end, vendor_team, pov_ref)
    content = call_azure_openai(POV_SYSTEM_PROMPT, pov_prompt)
    docx_bytes = create_pov_docx(content=content, customer_name=customer_name)
    return {
        "content": content,
        "bytes": docx_bytes,
        "file_name": f"{account_name}-PostAssessment POVdeployment.docx",
    }



def create_poe_zip(artifacts: List[Dict[str, Any]]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for artifact in artifacts:
            zip_file.writestr(artifact["file_name"], artifact["bytes"])
    return buffer.getvalue()


def get_existing_solution_text(current_doc_type: str) -> Optional[str]:
    key = "solution_text" if current_doc_type == "AI" else "infra_text"
    text = st.session_state.get(key)
    if isinstance(text, str) and text.strip():
        return text.strip()
    return None


def create_solution_artifact_from_text(
    current_doc_type: str,
    content: str,
    customer_name: str,
    account_name: str,
) -> Dict[str, Any]:
    if current_doc_type == "AI":
        docx_bytes = create_solution_docx(content=content, customer_name=customer_name)
        file_name = f"{account_name}-Solution Architecture.docx"
    else:
        docx_bytes = create_infra_docx(content=content, customer_name=customer_name)
        file_name = f"{account_name}-Infra Solution Architecture.docx"
    return {"content": content, "bytes": docx_bytes, "file_name": file_name}


def create_pov_artifact_from_text(
    content: str,
    customer_name: str,
    account_name: str,
) -> Dict[str, Any]:
    docx_bytes = create_pov_docx(content=content, customer_name=customer_name)
    return {
        "content": content,
        "bytes": docx_bytes,
        "file_name": f"{account_name}-PostAssessment POVdeployment.docx",
    }


def get_generated_migrate_csv_text() -> Optional[str]:
    csv_text = st.session_state.get("csv_code")
    if isinstance(csv_text, str) and csv_text.strip():
        return csv_text.strip()
    return None


def resolve_auto_inventory_csv(uploaded_inventory: Any) -> tuple[Optional[bytes], str]:
    if uploaded_inventory is not None:
        return uploaded_inventory.getvalue(), "手动上传 CSV"
    generated_csv = get_generated_migrate_csv_text()
    if generated_csv:
        return generated_csv.encode("utf-8-sig"), "Azure Migrate CSV 标签页生成"
    return None, "未提供"


def format_auto_poe_log(message: str) -> str:
    text = str(message or "").strip()
    if not text:
        return ""
    # 保留 Markdown 格式行（标题、颜色标注等）原样输出
    if text.startswith("**") or text.startswith("---") or ":green[" in text or ":red[" in text or ":orange[" in text:
        return text
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^[^\w\u4e00-\u9fff]+", "", text).strip()
    text = text.removesuffix("...").strip()
    text = re.sub(r"https?://\S+", "[url]", text)
    text = re.sub(r"/subscriptions/[^\s；;，,]+", lambda match: match.group(0).rstrip("/").split("/")[-1], text, flags=re.I)
    if not text:
        return ""
    return text


def should_display_auto_poe_log(message: str) -> bool:
    text = str(message or "").strip()
    if not text:
        return False
    # Markdown 格式行（分模块标题、带颜色的结果行）始终显示
    if text.startswith("**") or text.startswith("---") or ":green[" in text or ":red[" in text or ":orange[" in text:
        return True
    interim_markers = [
        "正在",
        "开始",
        "等待",
        "仍在",
        "继续",
        "第 ",
        "注册 ",
        "读取评估结果",
    ]
    result_markers = [
        "成功",
        "完成",
        "已",
        "复用",
        "解析",
        "目标区间",
        "不在目标区间",
        "未填写",
        "未解析",
        "请到 Azure Portal",
        "学习",
        "规模",
        "选定",
        "命中",
        "缓存",
        "前缀",
    ]
    if any(marker in text for marker in interim_markers) and not any(marker in text for marker in result_markers):
        return False
    return any(marker in text for marker in result_markers)


def run_full_auto_poe(
    current_doc_type: str,
    customer_name: str,
    account_name: str,
    customer_bg: str,
    solution_ref: str,
    infra_ref: str,
    pov_ref: str,
    token: str,
    subscription_id: str,
    resource_group: str,
    assessment_name: str,
    annual_budget_text: Optional[str],
    pov_start: Optional[datetime.date],
    pov_end: Optional[datetime.date],
    vendor_team: str,
    existing_solution_text: Optional[str],
    existing_pov_text: Optional[str],
    progress: Callable[[str], None],
) -> Dict[str, Any]:
    progress("---")
    progress("**一、AI 解决方案架构文档**")
    if existing_solution_text and existing_solution_text.strip():
        solution_artifact = create_solution_artifact_from_text(
            current_doc_type,
            existing_solution_text.strip(),
            customer_name,
            account_name,
        )
        progress(f"  :green[✅ 已复用] `{solution_artifact['file_name']}`")
    else:
        solution_artifact = generate_solution_artifact(
            current_doc_type,
            customer_name,
            account_name,
            customer_bg,
            solution_ref,
            infra_ref,
            annual_budget=parse_annual_budget_usd(annual_budget_text) or 0.0,
        )
        progress(f"  :green[✅ 已生成] `{solution_artifact['file_name']}`")

    target_key = "solution_text" if current_doc_type == "AI" else "infra_text"
    st.session_state[target_key] = solution_artifact["content"]
    st.session_state["customer_name"] = customer_name
    st.session_state["account_name"] = account_name

    progress("")
    progress("**二、POV 部署计划文档**")
    if existing_pov_text and existing_pov_text.strip():
        pov_artifact = create_pov_artifact_from_text(
            existing_pov_text.strip(),
            customer_name,
            account_name,
        )
        progress(f"  :green[✅ 已复用] `{pov_artifact['file_name']}`")
    else:
        if not pov_start or not pov_end:
            raise RuntimeError("缺少 POV 开始日期或结束日期，无法生成 POV 部署计划。")
        if not has_meaningful_pov_team(vendor_team):
            raise RuntimeError("缺少 POV 项目人员，无法生成 POV 部署计划。")
        pov_artifact = generate_pov_artifact(
            solution_artifact["content"],
            customer_name,
            account_name,
            pov_ref,
            pov_start,
            pov_end,
            vendor_team,
        )
        progress(f"  :green[✅ 已生成] `{pov_artifact['file_name']}`")
    st.session_state["pov_text"] = pov_artifact["content"]
    st.session_state["pov_source_doc_type"] = current_doc_type

    # 从解决方案文档中提取主要区域，用于评估目标地区
    target_location = _extract_dominant_region(solution_artifact["content"])

    progress("")
    progress("**三、Azure Migrate 评估报告**")
    migrate_result = run_azure_migrate_assessment(
        token, subscription_id, resource_group, account_name, assessment_name, annual_budget_text, progress,
        target_location=target_location,
    )
    # 修改 Excel 中的时间戳，使其位于用户 POV 时间区间内
    excel_bytes = migrate_result["excel_bytes"]
    if pov_start and pov_end:
        try:
            excel_bytes = fix_assessment_excel_timestamps(excel_bytes, pov_start, pov_end)
            progress("  ✅ 已修正评估报告时间戳至 POV 区间内")
        except Exception:
            pass  # 时间戳修改失败不阻塞主流程
    assessment_artifact = {
        "file_name": f"{account_name}-Azure Migrate Assessment.xlsx",
        "bytes": excel_bytes,
    }
    progress(f"  :green[✅ 已生成] `{assessment_artifact['file_name']}`")

    # 打包所有交付物
    all_artifacts = [solution_artifact, pov_artifact, assessment_artifact]
    zip_bytes = create_poe_zip(all_artifacts)
    progress(f"成功生成 POE 套件：{account_name}-POE-Complete.zip（{len(all_artifacts)} 个文件）")
    return {
        "zip_bytes": zip_bytes,
        "zip_name": f"{account_name}-POE-Complete.zip",
        "solution": solution_artifact,
        "pov": pov_artifact,
        "assessment": assessment_artifact,
        "migrate": migrate_result,
    }
