import html
from pathlib import Path
from typing import Any

import streamlit as st


def _html(text: Any) -> str:
    return html.escape(str(text or ""), quote=True)


def load_desktop_theme(app_dir: str) -> None:
    css_path = Path(app_dir) / "frontend" / "desktop_theme.css"
    st.markdown(f"<style>{css_path.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


def inject_theme_toggle_js() -> None:
    """Inject JS to sync the data-theme attribute on <html> based on session toggle."""
    st.markdown(
        """
        <script>
        (function() {
            const root = document.documentElement;
            // Check for theme override stored in sessionStorage
            const stored = window.sessionStorage.getItem('poe-theme');
            if (stored) {
                root.setAttribute('data-theme', stored);
            }
            // Expose a function for Streamlit to call
            window.setPoeTheme = function(theme) {
                if (theme === 'auto') {
                    root.removeAttribute('data-theme');
                    window.sessionStorage.removeItem('poe-theme');
                } else {
                    root.setAttribute('data-theme', theme);
                    window.sessionStorage.setItem('poe-theme', theme);
                }
            };
        })();
        </script>
        """,
        unsafe_allow_html=True,
    )


def render_app_header(version: str = "") -> None:
    version_badge = f'<span class="poe-badge" data-tone="version">v{version}</span>' if version else ""
    st.markdown(
        f"""
        <section class="poe-shell-header" aria-label="POE workflow header">
            <div>
                <p class="poe-kicker">POE Workflow Console {version_badge}</p>
                <h1 class="poe-title">微软客户 POE 工作台</h1>
                <p class="poe-subtitle">
                    输入客户背景，生成方案文档、POV 计划、Azure Migrate 评估和价格表交付物。
                </p>
            </div>
            <div class="poe-header-badges" aria-label="workflow capabilities">
                <span class="poe-badge" data-tone="accent">文档</span>
                <span class="poe-badge">POV</span>
                <span class="poe-badge">Migrate</span>
                <span class="poe-badge">Excel</span>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_section_head(title: str, desc: str = "", right: str = "") -> None:
    right_html = f'<div class="poe-status-row">{right}</div>' if right else ""
    desc_html = f'<p class="poe-section-desc">{_html(desc)}</p>' if desc else ""
    st.markdown(
        f'<div class="poe-section-head"><div><h2 class="poe-section-title">{_html(title)}</h2>'
        f'{desc_html}</div>{right_html}</div>',
        unsafe_allow_html=True,
    )


def render_pill(label: str, tone: str = "muted") -> str:
    return f'<span class="poe-pill" data-tone="{_html(tone)}">{_html(label)}</span>'


def render_workflow_steps(steps: list[dict[str, str]]) -> None:
    parts = []
    for idx, step in enumerate(steps, 1):
        state = step.get("state", "blocked")
        parts.append(
            f'<div class="poe-step" data-state="{_html(state)}">'
            f'<div class="poe-step-num">{idx}</div>'
            f'<p class="poe-step-title">{_html(step.get("title", ""))}</p>'
            f'</div>'
        )
    st.markdown(f'<div class="poe-steps">{"".join(parts)}</div>', unsafe_allow_html=True)


def render_readiness(items: list[tuple[str, bool, str]]) -> None:
    rows = []
    for label, is_ready, detail in items:
        state = "done" if is_ready else "todo"
        state_label = "就绪" if is_ready else "缺失"
        rows.append(
            f'<div class="poe-ready-item" data-state="{state}" title="{_html(detail)}">'
            f'<span class="poe-ready-dot"></span>'
            f'<strong>{_html(label)}</strong>'
            f'<span class="poe-ready-state">{state_label}</span>'
            f'</div>'
        )
    st.markdown(f'<div class="poe-readiness">{"".join(rows)}</div>', unsafe_allow_html=True)


def render_template_status(statuses: list[tuple[str, bool]]) -> None:
    rows = []
    for label, ok in statuses:
        tone = "success" if ok else "warning"
        state = "OK" if ok else "缺失"
        rows.append(
            f'<div class="poe-sidebar-row"><strong>{_html(label)}</strong>'
            f'<span class="poe-pill" data-tone="{tone}">{state}</span></div>'
        )
    st.markdown(f'<div class="poe-sidebar-status">{"".join(rows)}</div>', unsafe_allow_html=True)


def render_auto_poe_result(
    customer_name: str,
    generated_items: list[tuple[str, str]],
    migrate_items: list[tuple[str, str]],
) -> None:
    customer_label = _html(customer_name).strip()
    summary_text = (
        f"已生成{customer_label}的解决方案架构文档、POV 文档和迁移评估。"
        if customer_label and customer_label != "该客户"
        else "已生成该客户的解决方案架构文档、POV 文档和迁移评估。"
    )
    docs_html = "".join(
        f'<div class="poe-result-row"><span>{_html(label)}</span><strong>{_html(value)}</strong></div>'
        for label, value in generated_items
    )
    migrate_html = "".join(
        f'<div class="poe-result-metric"><span>{_html(label)}</span><strong>{_html(value)}</strong></div>'
        for label, value in migrate_items
    )
    st.markdown(
        f"""
        <section class="poe-result-panel" aria-label="POE generation result">
            <div class="poe-result-title">
                <span class="poe-result-mark">✓</span>
                <div>
                    <h3>已完成 POE 交付物生成</h3>
                    <p>{summary_text}</p>
                </div>
            </div>
            <div class="poe-result-section">
                <h4>产出文档</h4>
                <div class="poe-result-list">{docs_html}</div>
            </div>
            <div class="poe-result-section">
                <h4>迁移评估项目明细</h4>
                <div class="poe-result-grid">{migrate_html}</div>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_device_code_login(user_code: str, verify_url: str, container=None) -> None:
    """使用原生 Streamlit 组件渲染设备码登录区域（解决 iframe 剪贴板限制和窄列宽问题）。"""
    target = container if container is not None else st
    target.info(f"请复制验证码并在浏览器中完成登录，等待自动跳转。", icon="🔐")
    col_code, col_btn = target.columns([2, 1])
    with col_code:
        st.code(user_code, language=None)
    with col_btn:
        st.link_button("打开 Microsoft 登录", verify_url, use_container_width=True, type="primary")
