"""
POE 自动生成工作流 (POE Workflow Automator)
==========================================
一个基于 Streamlit 的 Web 应用，用于自动生成售前解决方案架构文档和 POV 部署计划。
通过 Azure OpenAI 服务驱动内容生成，使用客户提供的 .docx 模板控制输出格式。
"""

import io
import os
import re
import copy
import datetime
from typing import List, Optional
import streamlit as st
from openai import AzureOpenAI
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ──────────────────────────────────────────────
# 常量
# ──────────────────────────────────────────────
APP_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(APP_DIR, "templates")
SOLUTION_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "solution_template.docx.docx")
POV_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "pov_template.docx.docx")

# 中文字体名称
CN_FONT = "微软雅黑"
CN_FONT_ALT = "Microsoft YaHei UI"

# ──────────────────────────────────────────────
# 页面配置
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="POE 自动生成工作流",
    page_icon="🚀",
    layout="wide",
)

# ──────────────────────────────────────────────
# 自定义样式
# ──────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="st-"] { font-family: 'Inter', sans-serif; }

    .main-title {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.4rem;
        font-weight: 700;
        text-align: center;
        padding: 0.5rem 0 0.2rem 0;
    }
    .sub-title {
        text-align: center;
        color: #888;
        font-size: 1rem;
        margin-bottom: 1.5rem;
    }
    div[data-testid="stForm"] {
        border: 1px solid rgba(102, 126, 234, 0.25);
        border-radius: 16px;
        padding: 1.5rem;
        background: linear-gradient(145deg, rgba(102,126,234,0.04), rgba(118,75,162,0.04));
    }
    .stFormSubmitButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.6rem 2rem !important;
        font-weight: 600 !important;
        font-size: 1.05rem !important;
        width: 100% !important;
        transition: transform 0.15s, box-shadow 0.15s !important;
    }
    .stFormSubmitButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.45) !important;
    }
    .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 10px 20px;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ──────────────────────────────────────────────
# 检查 Secrets 配置
# ──────────────────────────────────────────────
def check_secrets() -> bool:
    """检查 st.secrets 中是否已配置所需的 Azure OpenAI 凭据。"""
    required_keys = ["AZURE_OPENAI_KEY", "AZURE_OPENAI_ENDPOINT", "AZURE_OPENAI_DEPLOYMENT"]
    missing = [k for k in required_keys if k not in st.secrets]
    if missing:
        st.error("⚠️ **Azure OpenAI 配置缺失**")
        st.info(
            "请在 `.streamlit/secrets.toml` 中配置以下密钥：\n\n"
            "```toml\n"
            'AZURE_OPENAI_KEY = "your-api-key"\n'
            'AZURE_OPENAI_ENDPOINT = "https://your-resource.openai.azure.com/"\n'
            'AZURE_OPENAI_DEPLOYMENT = "your-deployment-name"\n'
            'AZURE_OPENAI_API_VERSION = "2024-06-01"  # 可选，默认 2024-06-01\n'
            "```"
        )
        return False
    return True


# ──────────────────────────────────────────────
# Azure OpenAI 客户端
# ──────────────────────────────────────────────
def get_openai_client() -> AzureOpenAI:
    """创建 Azure OpenAI 客户端实例。"""
    return AzureOpenAI(
        api_key=st.secrets["AZURE_OPENAI_KEY"],
        azure_endpoint=st.secrets["AZURE_OPENAI_ENDPOINT"],
        api_version=st.secrets.get("AZURE_OPENAI_API_VERSION", "2024-06-01"),
    )


# ──────────────────────────────────────────────
# LLM 调用封装
# ──────────────────────────────────────────────
def call_azure_openai(system_prompt: str, user_prompt: str) -> str:
    """调用 Azure OpenAI Chat Completions API 并返回文本结果。"""
    client = get_openai_client()
    response = client.chat.completions.create(
        model=st.secrets["AZURE_OPENAI_DEPLOYMENT"],
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.7,
        max_completion_tokens=16384,
    )
    content = response.choices[0].message.content
    if not content or not content.strip():
        raise ValueError(
            f"API 返回了空内容。finish_reason={response.choices[0].finish_reason}"
        )
    return content


# ──────────────────────────────────────────────
# 模板文本提取（用于注入 AI Prompt）
# ──────────────────────────────────────────────
@st.cache_data
def extract_template_text(path: str) -> str:
    """从 .docx 模板文件中提取所有文本内容（含表格），用于注入 AI prompt。"""
    doc = Document(path)
    lines = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            lines.append(text)
    for table in doc.tables:
        header_cells = [cell.text.strip() for cell in table.rows[0].cells]
        lines.append("| " + " | ".join(header_cells) + " |")
        lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |")
        for row in table.rows[1:]:
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            lines.append("| " + " | ".join(cells) + " |")
        lines.append("")
    return "\n".join(lines)


# ──────────────────────────────────────────────
# Prompt 模板
# ──────────────────────────────────────────────
SOLUTION_SYSTEM_PROMPT = (
    "你是一位顶级的 Microsoft Azure AI 解决方案架构师。"
    "请根据用户提供的【客户名称】、【背景信息】和【预估年消耗】，生成一份完整、专业的 AI 售前解决方案架构文档。\n\n"
    "**标题要求（极其重要）：** 你的输出的第一行必须是一个 `#` 标题，格式为: `# [客户名称] - [具体方案名称]`。"
    "方案名称必须具体且针对客户业务，例如：\n"
    "- `# 深圳跃瓦创新科技 - Azure AI 中台与多场景助手解决方案`\n"
    "- `# 京华数码 - 智能外贸供应链 AI 平台方案`\n"
    "绝对不要使用笼统的'AI 解决方案架构文档'作为标题。\n\n"
    "**章节结构要求（必须严格遵循以下 8 个章节，使用中文数字编号 一、二、三...）：**\n\n"
    "## 一、摘要\n"
    "2-3 句话概述方案核心思路和预期价值。保持简洁。\n\n"
    "## 二、解决方案架构概览\n"
    "用 2-3 段话概述整体架构设计理念，描述架构逻辑（如共享推理池、租户隔离层、行业模型路由等核心概念）。用段落叙述，不要用列表。\n"
    "**在本章节末尾必须添加一行：** `[此处插入架构总览图]`，作为架构图的占位符。\n\n"
    "## 三、业务背景\n"
    "用段落叙述客户的行业定位、痛点和机遇。不要用列表。\n\n"
    "## 四、需求摘要\n"
    "以 Markdown 表格形式列出需求，表头为：`| 类别 | 需求描述 |`。\n"
    "**严格要求：表格只有 3 行数据（业务需求、功能需求、技术需求各 1 行），同一类别的多条需求合并到同一个单元格中。**\n\n"
    "## 五、详细解决方案设计\n"
    "这是最核心的章节。**严格禁止使用项目符号列表（-、*、• 等）。**\n"
    "必须使用 ### 子标题分节组织内容，参考以下结构：\n"
    "### 5.1 控制平面设计\n"
    "用段落叙述控制平面的架构设计，包括资源组规划、Hub 和 Project 的划分。\n"
    "### 5.2 数据与知识平面设计\n"
    "用段落叙述数据存储、AI Search、知识库索引等设计。\n"
    "### 5.3 算力与模型部署设计\n"
    "用段落叙述模型部署、负载均衡等设计。\n"
    "每个子节用段落叙述，加粗关键词引导要点（如 **资源组:** xxx），不要用列表。每个要点 2-3 句即可。\n\n"
    "## 六、安全架构\n"
    "用段落叙述数据隔离、身份认证等安全设计。不要用列表。每个要点用加粗关键词引导。\n\n"
    "## 七、集成架构\n"
    "用段落叙述集成方案。不要用列表。每个要点用加粗关键词引导。\n\n"
    "## 八、资源架构\n"
    "### Azure 资源需求\n"
    "以 Markdown 表格形式列出所有 Azure 资源，表头为：`| 资源名称 | 区域 | 规模与用途 |`。资源数量控制在 5-7 行。\n\n"
    "**全局格式要求（极其重要）：**\n"
    "- 章节标题使用 `## 一、摘要` 格式（## 开头 + 中文数字编号）\n"
    "- **严格禁止使用项目符号列表（-、*、• 开头的行）。** 全文必须使用段落叙述，用加粗关键词引导要点\n"
    "- 内容要精炼简洁，每个章节不超过模板文档的篇幅\n"
    "- 表格必须使用 Markdown 表格语法\n\n"
    "**重要：** 下方会提供一份【参考模板文档】，你必须严格学习它的写作风格（段落叙述，非列表）、内容篇幅和表格格式。以完全相同的结构和风格为新客户生成内容。"
)

POV_SYSTEM_PROMPT = (
    "你是一位经验丰富的 Microsoft 技术方案交付专家。"
    "请根据用户提供的【解决方案架构文档】、【客户名称】、【POV周期】以及【甲乙方项目人员名单】，生成一份 POV Deployment Plan。\n\n"
    "**标题要求（极其重要）：** 你的输出的第一行必须是一个 `#` 标题，格式为: "
    "`# [客户名称] - \"[项目代号]\" [方案核心描述] POV 部署计划`。例如：\n"
    "- `# 深圳跃瓦创新科技 - \" Azure AI 中台与多场景助手 POV 部署计划`\n"
    "- `# 京华数码 - \"JH-SmartTrade\" 智能外贸供应链 AI 平台 POV 部署计划`\n"
    "绝对不要使用笼统的'POV 部署计划'作为标题，必须包含具体的项目名称。\n\n"
    "**强相关要求：** POV 部署计划必须与解决方案架构文档强相关：\n"
    "部署的服务必须来自方案文档，步骤顺序符合架构依赖关系，验证场景对应核心功能。\n\n"
    "**章节结构要求（必须严格遵循以下结构）：**\n\n"
    "## 一、执行周期\n"
    "直接写出起止日期，如：2026年2月25日 - 2026年3月11日\n\n"
    "## 二、项目目标\n"
    "先用一句话概括总体目标和工作日天数，然后列出 3 个可衡量的目标。\n"
    "**每个目标必须简洁，使用数字编号格式，例如：**\n"
    "1. **知识检索准确率:** 验证 Azure AI Search 对产品手册的检索准确率，杜绝技术参数幻觉。\n"
    "2. **双模型分流:** 验证常规问答走 GPT-4o-mini 与复杂方案生成走 GPT-4o 的路由机制。\n"
    "3. **成本与生产规划:** 基于压测数据证明该架构能在预算内稳定运行。\n\n"
    "## 三、核心团队成员与职责\n"
    "以 Markdown 表格形式输出，表头必须为：`| 角色 | 所属方 | 姓名 | 角色职责 |`\n"
    "根据用户提供的人员名单填充，每人用 1-2 句描述职责。\n\n"
    "## 四、分阶段详细部署计划\n"
    "由你自己智能来划分阶段，每个阶段包含：\n"
    "1. **阶段标题**（加粗）：`**阶段 N: [阶段主题] ([M月D日] - [M月D日])**`\n"
    "2. **目标描述**：一句话说明本阶段核心目标\n"
    "3. **任务表格**：Markdown 表格，表头必须为：`| 日期 | 核心任务 | 主要负责人 | 里程碑与交付物 |`\n\n"
    "**日期要求（极其重要）：**\n"
    "- 任务表格中的日期必须是具体的日历日期（如 2月25日、2月26日）\n"
    "- **必须跳过周六和周日，只安排工作日**\n"
    "- 日期格式统一为：M月D日\n\n"
    
    "每天的任务必须具体、可操作。里程碑与交付物是具体产出（例如 '部署日志'、'准确率报告'、'UAT 签字单'）。\n\n"
    "**重要：** 下方会提供一份【参考模板文档】，你必须严格学习它的章节结构、分阶段格式、表格详细度和交付物命名规范。内容风格要精炼简洁，与模板保持一致。"
)


# ──────────────────────────────────────────────
# Word 文档生成 —— 通用工具函数
# ──────────────────────────────────────────────
def _set_run_font(run, font_name=CN_FONT, size_pt=None, bold=None, color_rgb=None):
    """为 run 设置字体（含中文 eastAsia 字体）。"""
    run.font.name = font_name
    # python-docx 需要同时设置 eastAsia 字体才能在 Word 中正确显示中文
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if color_rgb is not None:
        run.font.color.rgb = color_rgb


def _add_styled_paragraph(doc, text, font_name=CN_FONT, size_pt=9, bold=False,
                          color_rgb=None, alignment=None):
    """添加一个带完整样式的段落。"""
    p = doc.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    # 处理 **加粗** 和普通文字的混合
    parts = text.split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run(part)
        is_bold = bold or (i % 2 == 1)
        _set_run_font(run, font_name=font_name, size_pt=size_pt, bold=is_bold,
                       color_rgb=color_rgb)
    return p


def _add_styled_heading(doc, text, level=1):
    """添加一个使用中文字体的标题。"""
    heading = doc.add_heading("", level=level)
    run = heading.add_run(text)
    size_map = {1: 18, 2: 14, 3: 12}
    _set_run_font(run, font_name=CN_FONT, size_pt=size_map.get(level, 12), bold=True)
    return heading


def _parse_markdown_table(lines: List[str]) -> Optional[List[List[str]]]:
    """
    尝试从 Markdown 行列表中解析表格。
    返回二维数组 (包含表头)，如果不是表格则返回 None。
    """
    if len(lines) < 2:
        return None
    # 检查是否是 Markdown 表格（至少有 | 分隔符和分隔行 ---）
    if "|" not in lines[0]:
        return None

    rows = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        # 跳过分隔行 |---|---|
        if re.match(r"^\|[\s\-:|]+\|$", stripped):
            continue
        # 解析单元格
        cells = [c.strip() for c in stripped.split("|")]
        # 去掉首尾空元素（因为 | 在开头和结尾会产生空字符串）
        if cells and cells[0] == "":
            cells = cells[1:]
        if cells and cells[-1] == "":
            cells = cells[:-1]
        if cells:
            rows.append(cells)
    return rows if len(rows) >= 2 else None


def _add_word_table(doc, table_data: list[list[str]]):
    """将二维数组写入 Word 表格，应用专业样式。"""
    if not table_data:
        return

    num_cols = max(len(row) for row in table_data)
    table = doc.add_table(rows=len(table_data), cols=num_cols)
    table.style = "Table Grid"

    for ri, row_data in enumerate(table_data):
        for ci, cell_text in enumerate(row_data):
            if ci >= num_cols:
                break
            cell = table.cell(ri, ci)
            cell.text = ""  # 清空默认段落文本
            p = cell.paragraphs[0]
            run = p.add_run(cell_text)
            is_header = (ri == 0)
            _set_run_font(
                run,
                font_name=CN_FONT,
                size_pt=9,
                bold=is_header,
            )
            # 表头行背景色
            if is_header:
                shading = cell._element.get_or_add_tcPr()
                shading_elem = shading.makeelement(
                    qn("w:shd"),
                    {qn("w:fill"): "156082", qn("w:val"): "clear"},
                )
                shading.append(shading_elem)
                run.font.color.rgb = RGBColor(255, 255, 255)


def _markdown_to_docx(doc, markdown_text: str, body_size=9):
    """
    将 AI 返回的 Markdown 文本解析并写入 Word 文档。
    支持: 标题 (#/##/###)、列表 (-/*)、Markdown 表格、加粗 (**)、普通段落。
    """
    lines = markdown_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # 空行跳过
        if not stripped:
            i += 1
            continue

        # ── 标题 ──
        if stripped.startswith("### "):
            _add_styled_heading(doc, stripped[4:], level=3)
            i += 1
            continue
        if stripped.startswith("## "):
            _add_styled_heading(doc, stripped[3:], level=2)
            i += 1
            continue
        if stripped.startswith("# "):
            _add_styled_heading(doc, stripped[2:], level=1)
            i += 1
            continue

        # ── Markdown 表格 ──
        if "|" in stripped and not stripped.startswith("-"):
            table_lines = []
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i])
                i += 1
            table_data = _parse_markdown_table(table_lines)
            if table_data:
                _add_word_table(doc, table_data)
                doc.add_paragraph()  # 表格后空行
            else:
                # 不是表格，作为普通文本处理
                for tl in table_lines:
                    _add_styled_paragraph(doc, tl.strip(), size_pt=body_size)
            continue

        # ── 无序列表 ──
        if stripped.startswith("- ") or stripped.startswith("* "):
            text = stripped[2:]
            _add_styled_paragraph(doc, f"•  {text}", size_pt=body_size)
            i += 1
            continue

        # ── 有序列表 ──
        if stripped[0].isdigit() and ". " in stripped[:5]:
            _add_styled_paragraph(doc, stripped, size_pt=body_size)
            i += 1
            continue

        # ── 普通段落 ──
        _add_styled_paragraph(doc, stripped, size_pt=body_size)
        i += 1


# ──────────────────────────────────────────────
# Word 文档生成 —— 基于模板
# ──────────────────────────────────────────────
def _load_template(template_path: str) -> Document:
    """
    加载 .docx 模板文件作为基础文档。
    如果模板不存在，则返回一个空白 Document。
    """
    if os.path.exists(template_path):
        doc = Document(template_path)
        # 清空模板中的所有正文段落（保留样式定义、页面设置、页眉页脚）
        for p in doc.paragraphs:
            p._element.getparent().remove(p._element)
        # 清空表格
        for t in doc.tables:
            t._element.getparent().remove(t._element)
        return doc
    else:
        return Document()


def _extract_title(content: str, fallback: str = "") -> str:
    """从 AI 生成的 Markdown 内容中提取第一个 # 标题作为文档标题。"""
    for line in content.split("\n"):
        stripped = line.strip()
        if stripped.startswith("# ") and not stripped.startswith("## "):
            return stripped[2:].strip()
    return fallback


def _strip_first_heading(content: str) -> str:
    """去掉 Markdown 内容中的第一个 # 标题行（因为封面已经显示了标题）。"""
    lines = content.split("\n")
    result = []
    found = False
    for line in lines:
        stripped = line.strip()
        if not found and stripped.startswith("# ") and not stripped.startswith("## "):
            found = True
            continue  # 跳过第一个 # 标题
        result.append(line)
    return "\n".join(result)


def _add_page_break(doc):
    """在文档中添加分页符。"""
    from docx.oxml.ns import qn as _qn
    p = doc.add_paragraph()
    run = p.add_run()
    br = run._element.makeelement(_qn("w:br"), {_qn("w:type"): "page"})
    run._element.append(br)


def _add_toc(doc):
    """插入 Word 目录域（用户打开文档后按 Ctrl+A → F9 即可更新）。"""
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement

    # 目录标题
    toc_title = doc.add_paragraph()
    toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = toc_title.add_run("目录")
    _set_run_font(run, font_name=CN_FONT, size_pt=16, bold=True)

    doc.add_paragraph()  # 空行

    # 插入 TOC 域代码
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fldChar_begin = OxmlElement("w:fldChar")
    fldChar_begin.set(_qn("w:fldCharType"), "begin")
    run._element.append(fldChar_begin)

    instrText = OxmlElement("w:instrText")
    instrText.set(_qn("xml:space"), "preserve")
    instrText.text = ' TOC \\o "1-3" \\h \\z \\u '
    run._element.append(instrText)

    fldChar_separate = OxmlElement("w:fldChar")
    fldChar_separate.set(_qn("w:fldCharType"), "separate")
    run._element.append(fldChar_separate)

    # 占位文本（打开 Word 后会自动替换）
    placeholder = OxmlElement("w:r")
    placeholder_text = OxmlElement("w:t")
    placeholder_text.text = "（请右键点击此处 → 更新域，生成目录）"
    placeholder.append(placeholder_text)
    run._element.append(placeholder)

    fldChar_end = OxmlElement("w:fldChar")
    fldChar_end.set(_qn("w:fldCharType"), "end")
    run._element.append(fldChar_end)


def create_solution_docx(content: str, customer_name: str) -> bytes:
    """
    基于 solution 模板生成解决方案架构 Word 文档。
    布局: 封面标题（独占一页） → 目录（独占一页） → 正文
    """
    doc = _load_template(SOLUTION_TEMPLATE_PATH)
    title = _extract_title(content, f"{customer_name} - AI 解决方案架构文档")
    body_content = _strip_first_heading(content)

    # ---- 第 1 页：封面标题 ----
    for _ in range(8):
        doc.add_paragraph()

    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cover.add_run(title)
    # 与模板一致: 18pt #4874CB
    _set_run_font(run, font_name=CN_FONT_ALT, size_pt=18,
                   bold=True, color_rgb=RGBColor(0x48, 0x74, 0xCB))

    # 封面分页
    _add_page_break(doc)

    # ---- 第 2 页：目录 ----
    _add_toc(doc)

    # 目录分页
    _add_page_break(doc)

    # ---- 第 3 页起：正文内容（已去掉第一个 # 标题） ----
    _markdown_to_docx(doc, body_content, body_size=9)

    # 导出
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def create_pov_docx(content: str, customer_name: str) -> bytes:
    """
    基于 POV 模板生成 POV 部署计划 Word 文档。
    布局: 封面标题（独占一页） → 正文
    """
    doc = _load_template(POV_TEMPLATE_PATH)
    title = _extract_title(content, f"{customer_name} - POV 部署计划")
    body_content = _strip_first_heading(content)

    # ---- 第 1 页：封面标题 ----
    for _ in range(8):
        doc.add_paragraph()

    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cover.add_run(title)
    # 与模板一致: 22pt #156082
    _set_run_font(run, font_name=CN_FONT_ALT, size_pt=22,
                   bold=True, color_rgb=RGBColor(0x15, 0x60, 0x82))

    # 封面分页
    _add_page_break(doc)

    # ---- 第 2 页起：正文内容（已去掉第一个 # 标题） ----
    _markdown_to_docx(doc, body_content, body_size=9)

    # 导出
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ──────────────────────────────────────────────
# 主界面
# ──────────────────────────────────────────────
def main():
    # 标题
    st.markdown('<div class="main-title">🚀 POE 自动生成工作流</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">自动生成售前解决方案架构文档 & POV 部署计划  ·  参照模板  ·  一键导出 Word</div>',
        unsafe_allow_html=True,
    )

    # 检查配置
    if not check_secrets():
        st.stop()

    # 侧边栏
    with st.sidebar:
        st.markdown("### ⚙️ 操作")
        if st.button("🗑️ 清除已生成的结果", use_container_width=True):
            for key in ["solution_text", "pov_text", "customer_name"]:
                st.session_state.pop(key, None)
            st.rerun()

        # 模板状态指示
        st.markdown("---")
        st.markdown("### 📄 模板状态")
        sol_ok = os.path.exists(SOLUTION_TEMPLATE_PATH)
        pov_ok = os.path.exists(POV_TEMPLATE_PATH)
        st.markdown(f"- Solution 模板: {'✅ 已加载' if sol_ok else '⚠️ 未找到'}")
        st.markdown(f"- POV 模板: {'✅ 已加载' if pov_ok else '⚠️ 未找到'}")

    # 提取模板参考文本（用于注入 prompt）
    solution_ref = extract_template_text(SOLUTION_TEMPLATE_PATH) if os.path.exists(SOLUTION_TEMPLATE_PATH) else ""
    pov_ref = extract_template_text(POV_TEMPLATE_PATH) if os.path.exists(POV_TEMPLATE_PATH) else ""

    # ---- 输入表单 ----
    with st.form("poe_form", clear_on_submit=False):
        st.markdown("### 1. 客户与方案信息")

        col1, col2 = st.columns([2, 1])
        with col1:
            customer_name = st.text_input(
                "🏢 客户名称",
                placeholder="例如：Contoso Ltd.",
            )
        with col2:
            budget = st.text_input(
                "💰 预估年消耗 (USD)",
                placeholder="例如：50k+",
            )

        customer_bg = st.text_area(
            "📋 客户背景信息",
            placeholder="请粘贴从 Web 搜索到的客户背景资料，包括行业、规模、现有 IT 环境、核心需求等...",
            height=180,
        )

        # ---- 2. POV 计划信息 ----
        st.markdown("### 2. POV 计划信息")

        date_col1, date_col2 = st.columns(2)
        with date_col1:
            pov_start_date = st.date_input(
                "📅 POV 开始日期",
                value=datetime.date.today(),
            )
        with date_col2:
            pov_end_date = st.date_input(
                "📅 POV 结束日期",
                value=datetime.date.today() + datetime.timedelta(days=14),
            )

        team_members = st.text_area(
            "👥 甲乙方项目人员",
            value=(
                "技术负责人: 吕兴安 (领驭科技)\n"
                "Azure架构师: alex (领驭科技)\n"
                "海外业务总监: 王海峰 (京华数码)\n"
                "供应链IT主管: 刘丽 (京华数码)\n"
                "资深外贸业务员: 张伟 (京华数码)"
            ),
            height=160,
        )

        submitted = st.form_submit_button("🎯 生成全套 POE 文档")

    # ---- 工作流执行 ----
    if submitted:
        # 输入校验
        if not customer_name.strip():
            st.warning("请输入客户名称。")
            st.stop()
        if not customer_bg.strip():
            st.warning("请输入客户背景信息。")
            st.stop()

        # 构建 Solution 用户 Prompt（注入模板参考）
        user_context = (
            f"## 客户信息\n"
            f"- **客户名称**：{customer_name}\n"
            f"- **预估年消耗 (USD)**：{budget}\n\n"
            f"## 客户背景\n{customer_bg}"
        )
        if solution_ref:
            user_context += (
                f"\n\n---\n\n"
                f"## 【参考模板文档 —— 请学习其风格和结构，不要照抄具体数据】\n\n"
                f"{solution_ref}"
            )

        try:
            with st.spinner("🔄 第 1/2 步：正在生成解决方案架构文档，请稍候..."):
                solution_text = call_azure_openai(SOLUTION_SYSTEM_PROMPT, user_context)
                st.session_state["solution_text"] = solution_text
                st.session_state["customer_name"] = customer_name

            # 构建 POV 用户 Prompt（注入模板参考 + 日期 + 人员）
            pov_period = f"{pov_start_date.strftime('%Y/%m/%d')} - {pov_end_date.strftime('%Y/%m/%d')}"
            pov_user_prompt = (
                f"以下是已生成的解决方案架构文档，请据此生成 POV 部署计划：\n\n"
                f"{solution_text}\n\n"
                f"## 补充信息\n"
                f"- **客户名称**：{customer_name}\n"
                f"- **POV 周期**：{pov_period}\n\n"
                f"## 甲乙方项目人员\n{team_members}"
            )
            if pov_ref:
                pov_user_prompt += (
                    f"\n\n---\n\n"
                    f"## 【参考模板文档 —— 请学习其风格和结构，不要照抄具体数据】\n\n"
                    f"{pov_ref}"
                )

            with st.spinner("🔄 第 2/2 步：正在生成 POV 部署计划，请稍候..."):
                pov_text = call_azure_openai(POV_SYSTEM_PROMPT, pov_user_prompt)
                st.session_state["pov_text"] = pov_text

            st.success("✅ 文档生成完成！请在下方查看并下载。")
        except Exception as e:
            st.error(f"❌ 生成失败：{e}")
            st.stop()

    # ---- 展示结果（从 session_state 读取，避免刷新丢失） ----
    if "solution_text" in st.session_state and "pov_text" in st.session_state:
        customer = st.session_state.get("customer_name", "Customer")
        solution = st.session_state["solution_text"]
        pov = st.session_state["pov_text"]

        st.divider()

        # ---- 下载按钮区（置顶） ----
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            docx_solution = create_solution_docx(
                content=solution,
                customer_name=customer,
            )
            st.download_button(
                label="⬇️ 下载解决方案架构文档 (.docx)",
                data=docx_solution,
                file_name=f"{customer}_Solution_Architecture.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with dl_col2:
            docx_pov = create_pov_docx(
                content=pov,
                customer_name=customer,
            )
            st.download_button(
                label="⬇️ 下载 POV 部署计划 (.docx)",
                data=docx_pov,
                file_name=f"{customer}_POV_Deployment_Plan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        st.divider()

        # ---- 内容预览 ----
        tab1, tab2 = st.tabs(["📄 解决方案架构文档", "📋 POV 部署计划"])

        with tab1:
            st.markdown(solution)

        with tab2:
            st.markdown(pov)


# ──────────────────────────────────────────────
# 入口
# ──────────────────────────────────────────────
if __name__ == "__main__":
    main()
