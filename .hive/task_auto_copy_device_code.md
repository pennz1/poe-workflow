修改 Azure 登录流程：点击"登录 Azure"后自动复制 device code 到剪贴板，自动打开浏览器新标签页。

## 文件
frontend/ui.py，render_device_code_login 函数（第 138 行起）

## 当前行为
- 显示"复制验证码"按钮（用户需手动点）
- 显示"打开 Microsoft 登录"链接（用户需手动点）

## 目标行为
- JS 自动执行：复制验证码到剪贴板 + 打开 verify_url 新标签页
- 保留简短 UI 提示："验证码已复制，正在打开 Microsoft 登录页面..."
- 保留原有 fallback（execCommand 兜底复制）

## 改动

将 render_device_code_login 函数体（约第 139-217 行）替换为：

```python
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
```

## 验证
```bash
.venv/bin/python -m py_compile frontend/ui.py
grep -n "render_device_code_login" frontend/ui.py
.venv/bin/python -c "from frontend.ui import render_device_code_login; print('IMPORT OK')"
```

## Commit
```bash
git add frontend/ui.py
git commit -m "feat: auto-copy device code and open login page on Azure login"
```

完成后 team report 汇报结果。
