"""POE 自动生成工作流主入口."""

import os

import streamlit as st

VERSION = "1.1.1"
from config import (
    INFRA_TEMPLATE_PATH,
    MIGRATE_TEMPLATE_PATH,
    POV_TEMPLATE_PATH,
    PROJECT_DIR,
    SOLUTION_TEMPLATE_PATH,
    check_secrets,
)
from frontend.ui import (
    inject_theme_toggle_js,
    load_desktop_theme,
    render_app_header,
    render_pill,
    render_section_head,
    render_template_status,
)
from llm.prompts import extract_template_text
from pipeline import date_prefix
from session import clear_session_persist, persist_session_state, restore_session_state
from ui.auto_poe import render_full_auto_poe_area
from ui.migrate_tab import render_migrate_tab
from ui.tabs import render_csv_tab, render_postsales_tab, render_pov_tab, render_solution_tab, render_yearly_tab


def main():
    st.set_page_config(
        page_title="POE 自动生成工作流",
        page_icon="P",
        layout="wide",
    )
    load_desktop_theme(PROJECT_DIR)
    inject_theme_toggle_js()
    restore_session_state()
    render_app_header(version=VERSION)

    if not check_secrets():
        st.stop()

    # 侧边栏
    with st.sidebar:
        st.markdown("### 操作")
        if st.button("清除所有结果", use_container_width=True):
            for key in [
                "solution_text", "infra_text", "pov_text", "customer_name", "account_name", "csv_code",
                "budget", "doc_type", "pov_source_doc_type", "yearly_excel_bytes", "yearly_excel_name", "yearly_messages",
                "auto_poe_zip_bytes", "auto_poe_zip_name", "auto_poe_result", "auto_poe_error",
                "migrate_assessment_excel_bytes", "migrate_assessment_excel_name", "migrate_assessment_result",
                "migrate_assessment_error", "migrate_assessment_finish_time",
                "pov_vendor_team",
                "postsales_outputs",
            ]:
                st.session_state.pop(key, None)
            clear_session_persist()
            st.rerun()

        st.markdown("---")
        st.markdown("### 主题")
        theme_choice = st.selectbox(
            "外观模式",
            ["跟随系统", "浅色", "深色"],
            key="theme_mode",
            label_visibility="collapsed",
        )
        theme_map = {"跟随系统": "auto", "浅色": "light", "深色": "dark"}
        theme_val = theme_map[theme_choice]
        st.markdown(
            f'<script>window.setPoeTheme && window.setPoeTheme("{theme_val}");</script>',
            unsafe_allow_html=True,
        )

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

    solution_ref = extract_template_text(SOLUTION_TEMPLATE_PATH) if sol_ok else ""
    infra_ref = extract_template_text(INFRA_TEMPLATE_PATH) if infra_ok else ""
    pov_ref = extract_template_text(POV_TEMPLATE_PATH) if pov_ok else ""

    # ════════════════════════════════════════════════════
    # 公共输入区域
    # ════════════════════════════════════════════════════
    render_section_head(
        "客户信息",
        "这些输入会贯穿方案文档、POV 计划、CSV 推导和 Azure Migrate 评估。",
        render_pill("必填项优先", "accent"),
    )
    c0, c1, c2 = st.columns([1.5, 2, 1])
    with c0:
        account_name = st.text_input("账户名", placeholder="例如：Tetherflow", help="用于生成下载文件名的前缀", key="account_name")
    with c1:
        customer_name = st.text_input("客户名称", placeholder="例如：宇宙无敌科技有限公司", key="customer_name")
    with c2:
        budget = st.text_input("预估年消耗 (USD)", placeholder="例如：500k+", key="budget")

    customer_bg = st.text_area(
        "客户背景信息",
        placeholder="粘贴客户背景资料，包括行业、规模、现有 IT 环境、核心需求和已知约束。",
        height=125,
        key="customer_bg",
    )

    pov_dc1, pov_dc2, pov_dc3, pov_dc4 = st.columns([1, 1, 1, 1])
    with pov_dc1:
        st.date_input("POV 开始日期", value=None, key="pov_start_date")
    with pov_dc2:
        st.date_input("POV 结束日期", value=None, key="pov_end_date")
    with pov_dc3:
        pov_tech_lead = st.text_input("技术负责人", key="_pov_tech_lead", help="填写我方技术负责人姓名")
    with pov_dc4:
        pov_architect = st.text_input("架构师", key="_pov_architect", help="填写我方架构师姓名")
    _parts = []
    if (pov_tech_lead or "").strip():
        _parts.append(f"技术负责人: {pov_tech_lead.strip()}")
    if (pov_architect or "").strip():
        _parts.append(f"架构师: {pov_architect.strip()}")
    st.session_state["pov_vendor_team"] = ", ".join(_parts)

    st.divider()

    tab_auto, tab_sol, tab_pov, tab_csv, tab_migrate, tab_yearly, tab_postsales = st.tabs([
        "全自动POE生成", "解决方案文档", "POV 部署计划", "Azure Migrate CSV", "Azure Migrate 评估", "年度价格表", "售后价格表"
    ])

    dp = date_prefix()  # 日期前缀
    _ = dp

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
        render_pov_tab(pov_ref, account_name)

    with tab_csv:
        render_csv_tab(budget, account_name)

    with tab_migrate:
        render_migrate_tab(account_name, customer_name, budget)

    with tab_yearly:
        render_yearly_tab(budget, account_name)

    with tab_postsales:
        render_postsales_tab(budget, account_name)

    persist_session_state()


if __name__ == "__main__":
    main()
