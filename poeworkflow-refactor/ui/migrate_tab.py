"""Standalone Azure Migrate assessment tab."""

import datetime
import os
import time
from typing import Optional

import streamlit as st

from azure.auth import (
    _resource_group_label,
    _subscription_label,
    clear_azure_login,
    is_azure_token_valid,
    list_azure_resource_groups,
    list_azure_subscriptions,
    msal_device_code_login,
    set_device_code_container,
)
from azure.migrate import (
    _extract_dominant_region,
    _safe_azure_name,
    fix_assessment_excel_timestamps,
    run_azure_migrate_assessment,
)
from budget.parser import _format_usd, parse_annual_budget_usd
from budget.tier import _get_template_machine_names, budget_target_range, format_budget_target_range, load_tier_cache
from config import BUILTIN_CSV_PATH
from frontend.ui import render_auto_poe_result, render_readiness, render_workflow_steps
from pipeline import format_auto_poe_log, should_display_auto_poe_log
from session import persist_session_state


def _pick_default_index(labels, current_label: Optional[str]) -> int:
    if current_label and current_label in labels:
        return labels.index(current_label)
    return 0


def render_migrate_tab(account_name: str, customer_name: str, budget: str) -> None:
    azure_logged_in = is_azure_token_valid()
    token = st.session_state.get("azure_token")
    selected_subscription = st.session_state.get("_cached_subscription")
    selected_resource_group = st.session_state.get("_cached_resource_group")
    resolved_account_name = (
        account_name.strip()
        or str(st.session_state.get("account_name") or "").strip()
        or customer_name.strip()
        or str(st.session_state.get("customer_name") or "").strip()
    )
    resolved_customer_name = customer_name.strip() or str(st.session_state.get("customer_name") or "").strip()
    resolved_budget_text = budget.strip() or str(st.session_state.get("budget") or "").strip()
    budget_value = parse_annual_budget_usd(resolved_budget_text)
    matched_range = budget_target_range(budget_value)
    matched_tier = matched_range["tier"] if budget_value and budget_value > 0 else None
    builtin_csv_ok = os.path.exists(BUILTIN_CSV_PATH)
    tier_cache = load_tier_cache()
    tier_learned = bool(tier_cache.get("tiers", {}).get(str(matched_tier))) if matched_tier else False

    default_assessment_name = _safe_azure_name(resolved_account_name or resolved_customer_name, "poe", "", 55)
    previous_assessment_default = st.session_state.get("_migrate_assessment_default")
    current_assessment_name = str(st.session_state.get("migrate_assessment_name") or "").strip()
    if not current_assessment_name or current_assessment_name in {"poeassessment", previous_assessment_default}:
        st.session_state["migrate_assessment_name"] = default_assessment_name
    st.session_state["_migrate_assessment_default"] = default_assessment_name
    assessment_name = st.session_state.get("migrate_assessment_name", default_assessment_name)

    render_workflow_steps([
        {"title": "登录 Azure", "state": "done" if azure_logged_in else "ready"},
        {"title": "选择目标", "state": "done" if selected_subscription and selected_resource_group else ("ready" if azure_logged_in else "blocked")},
        {"title": "校准规模", "state": "done" if budget_value and budget_value > 0 else "blocked"},
        {"title": "生成评估", "state": "ready" if False else "blocked"},
    ])
    render_readiness([
        ("账户名", bool(resolved_account_name), "用于 Azure Migrate 项目和下载文件名"),
        ("预估年消耗", budget_value is not None and budget_value > 0, "用于匹配评估规模与目标区间"),
        ("Azure 登录", azure_logged_in, "用于创建 Azure Migrate 项目和评估"),
        ("订阅与资源组", bool(selected_subscription and selected_resource_group), "评估资源会创建到这里"),
        ("内置服务器模板", builtin_csv_ok, "上传前会自动补齐每台服务器的磁盘字段"),
        ("评估名称", bool(str(assessment_name).strip()), "用于 Azure Migrate 评估资源"),
    ])

    device_code_slot = st.container()
    az_c1, az_c2, az_c3 = st.columns([1, 1.5, 1.5])
    with az_c1:
        if azure_logged_in:
            st.success(f"已登录：{st.session_state.get('azure_user', 'Azure 用户')}", icon="✅")
            if st.button("退出", use_container_width=True, key="btn_migrate_azure_logout"):
                clear_azure_login()
                st.rerun()
        else:
            if st.button("登录 Azure", type="primary", use_container_width=True, key="btn_migrate_azure_login"):
                try:
                    set_device_code_container(device_code_slot)
                    msal_device_code_login()
                    st.rerun()
                except Exception as exc:
                    st.error(f"登录失败：{exc}")
    with az_c2:
        if azure_logged_in:
            try:
                subscriptions = list_azure_subscriptions(token)
                if subscriptions:
                    subscription_labels = [_subscription_label(sub) for sub in subscriptions]
                    cached_label = _subscription_label(selected_subscription) if selected_subscription else None
                    sub_label = st.selectbox(
                        "订阅",
                        subscription_labels,
                        index=_pick_default_index(subscription_labels, cached_label),
                        key="migrate_subscription_select",
                    )
                    selected_subscription = subscriptions[subscription_labels.index(sub_label)]
                    subscription_id = selected_subscription.get("subscriptionId")
                    st.session_state["azure_subscription_id"] = subscription_id
                    st.session_state["azure_subscription_name"] = selected_subscription.get("displayName", "")
                    st.session_state["_cached_subscription"] = selected_subscription
                    persist_session_state()
            except Exception as exc:
                st.error(f"读取订阅失败：{exc}")
        else:
            st.selectbox("订阅", ["登录后选择"], disabled=True, key="_migrate_sub_placeholder")
    with az_c3:
        if azure_logged_in and selected_subscription:
            try:
                groups = list_azure_resource_groups(selected_subscription.get("subscriptionId"), token)
                if groups:
                    group_labels = [_resource_group_label(group) for group in groups]
                    cached_label = _resource_group_label(selected_resource_group) if selected_resource_group else None
                    rg_label = st.selectbox(
                        "资源组",
                        group_labels,
                        index=_pick_default_index(group_labels, cached_label),
                        key="migrate_resource_group_select",
                    )
                    selected_resource_group = groups[group_labels.index(rg_label)]
                    st.session_state["azure_resource_group"] = selected_resource_group.get("name")
                    st.session_state["_cached_resource_group"] = selected_resource_group
                    persist_session_state()
                else:
                    st.warning("当前订阅下没有资源组。")
            except Exception as exc:
                st.error(f"读取资源组失败：{exc}")
        else:
            st.selectbox("资源组", ["登录后选择"], disabled=True, key="_migrate_rg_placeholder")

    name_c1, name_c2 = st.columns([1, 2])
    with name_c1:
        assessment_name = st.text_input("评估名称", value=assessment_name, key="migrate_assessment_name")
    with name_c2:
        if builtin_csv_ok:
            template_names = _get_template_machine_names()
            if matched_tier:
                tier_label = _format_usd(float(matched_tier)).replace("$", "\\$")
                range_label = format_budget_target_range(matched_range).replace("$", "\\$")
                if tier_learned:
                    learned_count = tier_cache["tiers"][str(matched_tier)]["machine_count"]
                    st.caption(f"内置模板 {len(template_names)} 台 | 规模 {tier_label} | 目标 {range_label} | 已学习 → {learned_count} 台 | 磁盘字段自动补齐")
                else:
                    st.caption(f"内置模板 {len(template_names)} 台 | 规模 {tier_label} | 目标 {range_label} | 首次运行将自动学习 | 磁盘字段自动补齐")
            else:
                st.caption(f"内置模板 {len(template_names)} 台 | 请填写预估年消耗匹配规模 | 磁盘字段自动补齐")
        else:
            st.error("Azurecsvtemplate.csv 未找到")

    ready = all([
        resolved_account_name,
        budget_value is not None and budget_value > 0,
        azure_logged_in,
        selected_subscription,
        selected_resource_group,
        builtin_csv_ok,
        str(assessment_name).strip(),
    ])

    if st.button(
        "生成 Azure Migrate Assessment",
        type="primary",
        use_container_width=True,
        key="btn_standalone_migrate_assessment",
        disabled=not ready,
    ):
        if not ready:
            st.warning("请先完成账户名、预估年消耗、Azure 登录、订阅/资源组和评估名称。")
            return

        log_placeholder = st.empty()
        log_lines = []
        st.session_state["migrate_assessment_running"] = True
        st.session_state.pop("migrate_assessment_error", None)

        def progress(message: str) -> None:
            formatted = format_auto_poe_log(message)
            if should_display_auto_poe_log(formatted):
                log_lines.append(formatted)
                log_placeholder.markdown("\n\n".join(log_lines[-120:]))

        try:
            with st.status("正在生成 Azure Migrate Assessment...", expanded=True) as status:
                solution_text = st.session_state.get("solution_text") or st.session_state.get("infra_text") or ""
                target_location = _extract_dominant_region(solution_text) if isinstance(solution_text, str) else None
                migrate_result = run_azure_migrate_assessment(
                    token=token,
                    subscription_id=selected_subscription["subscriptionId"],
                    resource_group=selected_resource_group["name"],
                    account_name=resolved_account_name,
                    assessment_name=str(assessment_name).strip(),
                    annual_budget_text=resolved_budget_text,
                    progress=progress,
                    target_location=target_location,
                )
                excel_bytes = migrate_result["excel_bytes"]
                pov_start = st.session_state.get("pov_start_date")
                pov_end = st.session_state.get("pov_end_date")
                if pov_start and pov_end:
                    try:
                        excel_bytes = fix_assessment_excel_timestamps(excel_bytes, pov_start, pov_end)
                        progress("  ✅ 已修正评估报告时间戳至 POV 区间内")
                    except Exception:
                        pass

                file_name = f"{resolved_account_name}-Azure Migrate Assessment.xlsx"
                st.session_state["migrate_assessment_excel_bytes"] = excel_bytes
                st.session_state["migrate_assessment_excel_name"] = file_name
                st.session_state["migrate_assessment_result"] = {
                    "customer_name": resolved_customer_name or resolved_account_name,
                    "file_name": file_name,
                    "project_name": migrate_result.get("project_name"),
                    "assessment_name": migrate_result.get("assessment_name"),
                    "machine_count": len(migrate_result.get("assessed_machines", [])),
                    "portal_inventory_count": migrate_result.get("portal_inventory_count", 0),
                    "annualized_cost": migrate_result.get("annualized_cost"),
                    "budget_target": migrate_result.get("budget_target"),
                    "budget_target_range_label": migrate_result.get("budget_target_range_label"),
                    "budget_target_met": migrate_result.get("budget_target_met", True),
                    "tier": migrate_result.get("tier"),
                    "selected_machine_count": migrate_result.get("selected_machine_count"),
                    "total_machine_count": migrate_result.get("total_machine_count"),
                }
                st.session_state["migrate_assessment_finish_time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.session_state["migrate_assessment_start_time"] = time.time()
                st.session_state.pop("migrate_assessment_running", None)
                status.update(label="Azure Migrate Assessment 生成完成", state="complete")
                persist_session_state()
        except Exception as exc:
            st.session_state.pop("migrate_assessment_running", None)
            st.session_state["migrate_assessment_error"] = str(exc)
            st.error(f"Azure Migrate Assessment 生成失败：{exc}")

    if st.session_state.get("migrate_assessment_error"):
        st.error(st.session_state["migrate_assessment_error"])

    if "migrate_assessment_excel_bytes" in st.session_state:
        result = st.session_state.get("migrate_assessment_result", {})
        render_auto_poe_result(
            customer_name=result.get("customer_name") or resolved_customer_name or resolved_account_name or "该客户",
            generated_items=[("迁移评估文档", result.get("file_name") or "-")],
            migrate_items=[
                ("Azure Migrate 项目", result.get("project_name") or "-"),
                ("迁移评估名称", result.get("assessment_name") or "-"),
                ("CSV 来源", "内置模板（Azurecsvtemplate.csv，磁盘字段自动补齐）"),
                ("规模档位", _format_usd(float(result["tier"])) if result.get("tier") else "-"),
                ("选定服务器", f"{result.get('selected_machine_count', 0)}/{result.get('total_machine_count', 0)} 台"),
                ("Portal 库存", f"{result.get('portal_inventory_count', 0)} 台"),
                ("评估服务器数", f"{result.get('machine_count', 0)} 台"),
                ("年化估算", _format_usd(result.get("annualized_cost"))),
                ("用户预估", _format_usd(result.get("budget_target"))),
                ("目标区间", result.get("budget_target_range_label") or "-"),
            ],
        )
        if result.get("budget_target") and not result.get("budget_target_met", True):
            st.warning("自动调整 3 轮后仍未落入用户预估年消耗的档次区间，请在 Azure Portal 的评估设置中手动调整。")
        finish_time = st.session_state.get("migrate_assessment_finish_time")
        if finish_time:
            st.info(f"✅ 任务完成时间：{finish_time}")
        st.download_button(
            label="下载 Azure Migrate Assessment (.xlsx)",
            data=st.session_state["migrate_assessment_excel_bytes"],
            file_name=st.session_state["migrate_assessment_excel_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_standalone_migrate_assessment",
        )