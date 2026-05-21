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
            gap: 10px;
            padding: 10px 14px;
            border-radius: 8px;
            border: 1px solid oklch(86.5% 0.014 248);
            background: oklch(99.2% 0.003 248);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei UI", system-ui, sans-serif;
            font-size: 14px;
        }}
        .poe-device-code {{
            font-weight: 700;
            font-size: 15px;
            letter-spacing: 0.5px;
            color: oklch(39% 0.15 250);
            background: oklch(94% 0.01 250);
            padding: 3px 10px;
            border-radius: 4px;
            user-select: all;
        }}
        .poe-device-copy-btn {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 28px;
            height: 28px;
            border: 1px solid oklch(80% 0.05 250);
            border-radius: 6px;
            background: oklch(100% 0 0);
            cursor: pointer;
            font-size: 13px;
            padding: 0;
            flex-shrink: 0;
        }}
        .poe-device-copy-btn:hover {{
            background: oklch(94% 0.01 250);
        }}
        .poe-device-copy-btn.copied {{
            background: oklch(48% 0.105 152);
            border-color: oklch(48% 0.105 152);
        }}
        .poe-device-link-btn {{
            display: inline-flex;
            align-items: center;
            padding: 5px 14px;
            border-radius: 6px;
            border: 1px solid oklch(58% 0.14 252);
            background: oklch(58% 0.14 252);
            color: #fff;
            font-weight: 600;
            font-size: 13px;
            text-decoration: none;
            cursor: pointer;
            margin-left: auto;
            flex-shrink: 0;
        }}
        .poe-device-link-btn:hover {{
            background: oklch(48% 0.12 252);
            border-color: oklch(48% 0.12 252);
        }}
        .poe-device-hint {{
            color: oklch(60% 0.02 250);
            font-size: 12px;
            flex-shrink: 1;
            min-width: 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        </style>
        <div class="poe-device-login">
            <span class="poe-device-code">{code}</span>
            <button type="button" class="poe-device-copy-btn" id="poe-copy-btn" title="复制验证码">📋</button>
            <span class="poe-device-hint" id="poe-hint">已自动复制</span>
            <a class="poe-device-link-btn" href="{url}" target="_blank" rel="noopener noreferrer">打开 Microsoft 登录 →</a>
        </div>
        <script>
        (function() {{
            const code = "{code}";
            const btn = document.getElementById("poe-copy-btn");
            const hint = document.getElementById("poe-hint");

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

            // 自动复制
            (async function() {{
                const ok = await copyToClipboard(code);
                if (ok) {{
                    hint.textContent = "已自动复制";
                }} else {{
                    hint.textContent = "请手动复制";
                }}
            }})();

            // 手动复制按钮
            btn.addEventListener("click", async function() {{
                await copyToClipboard(code);
                btn.textContent = "✓";
                btn.classList.add("copied");
                hint.textContent = "已复制";
                setTimeout(function() {{
                    btn.textContent = "📋";
                    btn.classList.remove("copied");
                }}, 1500);
            }});
        }})();
        </script>
        """,
        height=52,
    )
