"""Full auto POE Streamlit UI."""

import datetime
import os
import time
from typing import List

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
from azure.migrate import _safe_azure_name
from budget.parser import _format_usd, parse_annual_budget_usd
from budget.tier import _get_template_machine_names, budget_target_range, format_budget_target_range, load_tier_cache
from config import BUILTIN_CSV_PATH
from documents.pov import has_meaningful_pov_team
from frontend.ui import render_auto_poe_result, render_readiness, render_workflow_steps
from pipeline import (
    format_auto_poe_log,
    get_existing_solution_text,
    run_full_auto_poe,
    should_display_auto_poe_log,
)
from session import persist_session_state

def render_full_auto_poe_area(
    current_doc_type: str,
    account_name: str,
    customer_name: str,
    budget: str,
    customer_bg: str,
    solution_ref: str,
    infra_ref: str,
    pov_ref: str,
) -> None:
    azure_logged_in = is_azure_token_valid()
    selected_subscription = st.session_state.get("_cached_subscription")
    selected_resource_group = st.session_state.get("_cached_resource_group")
    resolved_customer_name = customer_name.strip() or str(st.session_state.get("customer_name") or "").strip()
    resolved_account_name = (
        account_name.strip()
        or str(st.session_state.get("account_name") or "").strip()
        or resolved_customer_name
    )
    resolved_budget_text = budget.strip() or str(st.session_state.get("budget") or "").strip()
    default_assessment_name = _safe_azure_name(resolved_account_name or resolved_customer_name, "poe", "", 55)
    previous_assessment_default = st.session_state.get("_auto_assessment_default")
    current_assessment_name = str(st.session_state.get("auto_assessment_name") or "").strip()
    if not current_assessment_name or current_assessment_name in {"poeassessment", previous_assessment_default}:
        st.session_state["auto_assessment_name"] = default_assessment_name
        current_assessment_name = default_assessment_name
    st.session_state["_auto_assessment_default"] = default_assessment_name
    assessment_name = st.session_state.get("auto_assessment_name", default_assessment_name)
    budget_value = parse_annual_budget_usd(resolved_budget_text)
    pov_start = st.session_state.get("pov_start_date")
    pov_end = st.session_state.get("pov_end_date")
    vendor_team = st.session_state.get("pov_vendor_team", "")
    existing_solution_text = get_existing_solution_text(current_doc_type)
    existing_pov_text = st.session_state.get("pov_text")
    existing_pov_text = existing_pov_text.strip() if isinstance(existing_pov_text, str) else None
    pov_source_doc_type = st.session_state.get("pov_source_doc_type")
    if existing_pov_text and pov_source_doc_type and pov_source_doc_type != current_doc_type:
        existing_pov_text = None
    builtin_csv_ok = os.path.exists(BUILTIN_CSV_PATH)
    matched_range = budget_target_range(budget_value)
    matched_tier = matched_range["tier"] if budget_value and budget_value > 0 else None
    tier_cache = load_tier_cache()
    tier_learned = bool(tier_cache.get("tiers", {}).get(str(matched_tier))) if matched_tier else False
    has_existing_solution = bool(existing_solution_text)
    has_existing_pov = bool(existing_pov_text)
    pov_dates_filled = bool(pov_start and pov_end)
    pov_team_ready = has_meaningful_pov_team(vendor_team)
    solution_ready = has_existing_solution or (bool(resolved_customer_name) and bool(customer_bg.strip()))
    pov_ready = has_existing_pov or (pov_dates_filled and pov_team_ready)
    status_slot = st.empty()

    token = st.session_state.get("azure_token")

    # ── 就绪状态 + 步骤 ──
    with status_slot.container():
        render_workflow_steps([
            {"title": "登录 Azure", "state": "done" if azure_logged_in else "ready"},
            {"title": "选择目标", "state": "done" if selected_subscription and selected_resource_group else ("ready" if azure_logged_in else "blocked")},
            {"title": "内置模板", "state": "done" if builtin_csv_ok else "blocked"},
            {"title": "生成套件", "state": "ready" if False else "blocked"},
        ])
        readiness_items = [
            ("客户名称", bool(resolved_customer_name), "用于文档标题和输出文件名"),
            ("方案文档", solution_ready, "已生成则复用；未生成时会按客户背景生成"),
            ("预估年消耗", budget_value is not None and budget_value > 0, "用于校准迁移评估年化估算和规模匹配"),
            ("POV 文档/输入", pov_ready, "已生成则复用；未生成时需填写时间区间和项目人员"),
            ("Azure 登录", azure_logged_in, "用于创建 Azure Migrate 项目和评估"),
            ("订阅与资源组", bool(selected_subscription and selected_resource_group), "评估资源会创建到这里"),
            ("内置服务器模板", builtin_csv_ok, "内置 Azurecsvtemplate.csv 模板"),
            ("评估名称", bool(assessment_name.strip()), "用于 Azure Migrate 评估资源"),
        ]
        render_readiness(readiness_items)

    # ── Azure 登录 + 订阅/资源组 — 单行 ──
    device_code_slot = st.container()
    az_c1, az_c2, az_c3 = st.columns([1, 1.5, 1.5])
    with az_c1:
        if azure_logged_in:
            st.success(f"已登录：{st.session_state.get('azure_user', 'Azure 用户')}", icon="✅")
            if st.button("退出", use_container_width=True, key="btn_azure_logout"):
                clear_azure_login()
                st.rerun()
        else:
            if st.button("登录 Azure", type="primary", use_container_width=True, key="btn_azure_login"):
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
                    default_index = next(
                        (
                            idx for idx, sub in enumerate(subscriptions)
                            if "Microsoft" in (sub.get("displayName") or "")
                            and ("合作伙伴" in (sub.get("displayName") or "") or "Partner" in (sub.get("displayName") or ""))
                        ),
                        0,
                    )
                    sub_label = st.selectbox("订阅", subscription_labels, index=default_index, key="auto_subscription_select")
                    selected_subscription = subscriptions[subscription_labels.index(sub_label)]
                    subscription_id = selected_subscription.get("subscriptionId")
                    st.session_state["azure_subscription_id"] = subscription_id
                    st.session_state["azure_subscription_name"] = selected_subscription.get("displayName", "")
                    st.session_state["_cached_subscription"] = selected_subscription
                    persist_session_state()
            except Exception as exc:
                st.error(f"读取订阅失败：{exc}")
        else:
            st.selectbox("订阅", ["登录后选择"], disabled=True, key="_sub_placeholder")
    with az_c3:
        if azure_logged_in and selected_subscription:
            try:
                groups = list_azure_resource_groups(subscription_id, token)
                if groups:
                    group_labels = [_resource_group_label(group) for group in groups]
                    rg_label = st.selectbox("资源组", group_labels, key="auto_resource_group_select")
                    selected_resource_group = groups[group_labels.index(rg_label)]
                    st.session_state["azure_resource_group"] = selected_resource_group.get("name")
                    st.session_state["_cached_resource_group"] = selected_resource_group
                    persist_session_state()
                else:
                    st.warning("当前订阅下没有资源组。")
            except Exception as exc:
                st.error(f"读取资源组失败：{exc}")
        else:
            st.selectbox("资源组", ["登录后选择"], disabled=True, key="_rg_placeholder")

    # ── 评估名称 + 模板状态 — 单行 ──
    name_c1, name_c2 = st.columns([1, 2])
    with name_c1:
        assessment_name = st.text_input("评估名称", value=assessment_name, key="auto_assessment_name")
    with name_c2:
        if builtin_csv_ok:
            template_names = _get_template_machine_names()
            if matched_tier:
                tier_label = _format_usd(float(matched_tier))
                range_label = format_budget_target_range(matched_range)
                tier_label_md = tier_label.replace("$", "\\$")
                range_label_md = range_label.replace("$", "\\$")
                if tier_learned:
                    learned_count = tier_cache["tiers"][str(matched_tier)]["machine_count"]
                    st.caption(f"内置模板 {len(template_names)} 台 | 规模 {tier_label_md} | 目标 {range_label_md} | 已学习 → {learned_count} 台")
                else:
                    st.caption(f"内置模板 {len(template_names)} 台 | 规模 {tier_label_md} | 目标 {range_label_md} | 首次运行将自动学习")
            else:
                st.caption(f"内置模板 {len(template_names)} 台 | 请填写预估年消耗匹配规模")
        else:
            st.error("Azurecsvtemplate.csv 未找到")

    workflow_ready = all(item[1] for item in readiness_items)

    if st.button(
        "生成完整 POE 套件",
        type="primary",
        use_container_width=True,
        key="btn_full_auto_poe",
        disabled=not workflow_ready,
    ):
        cust = resolved_customer_name
        acct = resolved_account_name or cust
        if not cust:
            st.warning("请输入客户名称。")
            return
        if not solution_ready:
            st.warning("请先在「解决方案文档」标签页生成/导入方案文档，或填写客户背景以便自动生成。")
            return
        if budget_value is None or budget_value <= 0:
            st.warning("请输入可解析的预估年消耗，例如 500k、50万 或 500000。")
            return
        if not has_existing_pov and not pov_dates_filled:
            st.warning("请在上方公共输入区域填写 POV 开始日期和结束日期。")
            return
        if not has_existing_pov and pov_dates_filled and pov_end < pov_start:
            st.warning("POV 结束日期不能早于开始日期，请修正。")
            return
        if not has_existing_pov and not pov_team_ready:
            st.warning("请在上方公共输入区域填写乙方项目人员。")
            return
        if not is_azure_token_valid():
            st.warning("请先登录 Microsoft Azure 账户。")
            return
        if not selected_subscription or not selected_resource_group:
            st.warning("请选择订阅和资源组。")
            return
        if not builtin_csv_ok:
            st.warning("内置服务器清单模板 Azurecsvtemplate.csv 未找到。")
            return
        if not assessment_name.strip():
            st.warning("请输入评估名称。")
            return

        try:
            st.session_state["auto_poe_running"] = True
            st.session_state["auto_poe_start_time"] = time.time()
            with st.status("正在生成完整 POE 套件", expanded=True) as status:
                stop_placeholder = st.empty()
                log_placeholder = st.empty()
                log_lines: List[str] = []
                stop_placeholder.button(
                    "停止生成", type="secondary", use_container_width=True, key="btn_stop_auto_poe",
                    on_click=lambda: st.session_state.update({"auto_poe_stop": True}),
                )

                def progress(message: str) -> None:
                    if st.session_state.get("auto_poe_stop"):
                        st.session_state.pop("auto_poe_stop", None)
                        st.session_state.pop("auto_poe_running", None)
                        raise RuntimeError("用户已手动停止生成。")
                    formatted = format_auto_poe_log(message)
                    if formatted and should_display_auto_poe_log(formatted):
                        log_lines.append(formatted)
                        # 使用 markdown 渲染，支持加粗、颜色等格式
                        log_placeholder.markdown("\n\n".join(log_lines[-120:]))

                result = run_full_auto_poe(
                    current_doc_type=current_doc_type,
                    customer_name=cust,
                    account_name=acct,
                    customer_bg=customer_bg.strip(),
                    solution_ref=solution_ref,
                    infra_ref=infra_ref,
                    pov_ref=pov_ref,
                    token=st.session_state["azure_token"],
                    subscription_id=selected_subscription["subscriptionId"],
                    resource_group=selected_resource_group["name"],
                    assessment_name=assessment_name.strip(),
                    annual_budget_text=resolved_budget_text,
                    pov_start=pov_start,
                    pov_end=pov_end,
                    vendor_team=vendor_team,
                    existing_solution_text=existing_solution_text,
                    existing_pov_text=existing_pov_text,
                    progress=progress,
                )
                stop_placeholder.empty()
                st.session_state.pop("auto_poe_running", None)
                st.session_state["auto_poe_zip_bytes"] = result["zip_bytes"]
                st.session_state["auto_poe_zip_name"] = result["zip_name"]
                st.session_state["auto_poe_result"] = {
                    "customer_name": cust,
                    "zip_name": result["zip_name"],
                    "solution_file_name": result["solution"]["file_name"],
                    "pov_file_name": result["pov"]["file_name"],
                    "assessment_file_name": result["assessment"]["file_name"],
                    "project_name": result["migrate"]["project_name"],
                    "assessment_name": result["migrate"]["assessment_name"],
                    "machine_count": len(result["migrate"].get("assessed_machines", [])),
                    "portal_inventory_count": result["migrate"].get("portal_inventory_count", 0),
                    "annualized_cost": result["migrate"].get("annualized_cost"),
                    "budget_target": result["migrate"].get("budget_target"),
                    "budget_target_range_label": result["migrate"].get("budget_target_range_label"),
                    "budget_target_met": result["migrate"].get("budget_target_met", True),
                    "csv_source": "内置模板（Azurecsvtemplate.csv）",
                    "tier": result["migrate"].get("tier"),
                    "selected_machine_count": result["migrate"].get("selected_machine_count"),
                    "total_machine_count": result["migrate"].get("total_machine_count"),
                }
                status.update(label="完整 POE 套件生成完成", state="complete")
                st.session_state["auto_poe_finish_time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except Exception as exc:
            st.session_state.pop("auto_poe_running", None)
            st.session_state["auto_poe_error"] = str(exc)
            st.error(f"全自动生成失败：{exc}")

    if "auto_poe_zip_bytes" in st.session_state:
        result = st.session_state.get("auto_poe_result", {})
        annualized_cost = result.get("annualized_cost")
        budget_target = result.get("budget_target")
        generated_items = [
            ("解决方案架构文档", result.get("solution_file_name") or "-"),
            ("POV 文档", result.get("pov_file_name") or "-"),
            ("迁移评估文档", result.get("assessment_file_name") or "-"),
        ]
        render_auto_poe_result(
            customer_name=result.get("customer_name") or customer_name.strip() or account_name.strip() or "该客户",
            generated_items=generated_items,
            migrate_items=[
                ("Azure Migrate 项目", result.get("project_name") or "-"),
                ("迁移评估名称", result.get("assessment_name") or "-"),
                ("CSV 来源", result.get("csv_source") or "-"),
                ("规模档位", _format_usd(float(result["tier"])) if result.get("tier") else "-"),
                ("选定服务器", f"{result.get('selected_machine_count', 0)}/{result.get('total_machine_count', 0)} 台"),
                ("Portal 库存", f"{result.get('portal_inventory_count', 0)} 台"),
                ("评估服务器数", f"{result.get('machine_count', 0)} 台"),
                ("年化估算", _format_usd(annualized_cost)),
                ("用户预估", _format_usd(budget_target)),
                ("目标区间", result.get("budget_target_range_label") or "-"),
            ],
        )
        if budget_target and not result.get("budget_target_met", True):
            st.warning("自动调整 3 轮后仍未落入用户预估年消耗的档次区间，请在 Azure Portal 的评估设置中手动调整。")
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
        st.download_button(
            label="下载全部 POE 文档 (.zip)",
            data=st.session_state["auto_poe_zip_bytes"],
            file_name=st.session_state["auto_poe_zip_name"],
            mime="application/zip",
            use_container_width=True,
            key="dl_auto_poe_zip",
        )


# ──────────────────────────────────────────────
# 主界面
# ──────────────────────────────────────────────
