import html
from pathlib import Path
from typing import Any

import streamlit as st
import streamlit.components.v1 as components


def _html(text: Any) -> str:
    return html.escape(str(text or ""), quote=True)


def load_desktop_theme(app_dir: str) -> None:
    css_path = Path(app_dir) / "frontend" / "desktop_theme.css"
    st.markdown(f"<style>{css_path.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


def render_app_header() -> None:
    st.markdown(
        """
        <section class="poe-shell-header" aria-label="POE workflow header">
            <div>
                <p class="poe-kicker">POE Workflow Console</p>
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


def render_device_code_login(user_code: str, verify_url: str) -> None:
    code = _html(user_code)
    url = _html(verify_url)
    components.html(
        f"""
        <style>
        .poe-device-login {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            padding: 12px 16px;
            border-radius: 8px;
            border: 1px solid oklch(86.5% 0.014 248);
            background: oklch(99.2% 0.003 248);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei UI", system-ui, sans-serif;
            font-size: 14px;
            color: oklch(39% 0.15 250);
        }}
        </style>
        <div class="poe-device-login" id="poe-login-status">
            正在准备登录...
        </div>
        <script>
        (function() {{
            const code = "{code}";
            const url = "{url}";
            const el = document.getElementById("poe-login-status");

            const copyToClipboard = async function(text) {{
                try {{
                    await navigator.clipboard.writeText(text);
                    return true;
                }} catch (e) {{
                    const ta = document.createElement("textarea");
                    ta.value = text;
                    ta.setAttribute("readonly", "");
                    ta.style.position = "fixed";
                    ta.style.opacity = "0";
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand("copy");
                    ta.remove();
                    return true;
                }}
            }};

            (async function() {{
                const ok = await copyToClipboard(code);
                if (ok) {{
                    el.textContent = "验证码已复制 " + code + "，正在打开登录页面...";
                    setTimeout(function() {{
                        window.open(url, "_blank", "noopener,noreferrer");
                    }}, 800);
                }} else {{
                    el.innerHTML = "请手动复制验证码 <strong>" + code + "</strong> 并前往 <a href=\\"" + url + "\\" target=\\"_blank\\">Microsoft 登录</a>";
                }}
            }})();
        }})();
        </script>
        """,
        height=50,
    )
