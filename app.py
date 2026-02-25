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
MIGRATE_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "AzureMigrateimporttemplate.csv")

# 中文字体名称
CN_FONT = "微软雅黑"
CN_FONT_ALT = "Microsoft YaHei UI"

# ──────────────────────────────────────────────
# 页面配置
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="POE 自动生成工作流",
    page_icon="P",
    layout="wide",
)

# ──────────────────────────────────────────────
# 自定义样式
# ──────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    @import url('https://fonts.googleapis.com/icon?family=Material+Icons');
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
        max_completion_tokens=128000,
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
    "2-3 段话概述整体架构设计理念。用段落叙述，不要用列表。\n\n"
    "## 三、业务背景\n"
    "用段落叙述客户的行业定位、痛点和机遇。不要用列表。\n\n"
    "## 四、需求摘要\n"
    "以 Markdown 表格形式列出需求，表头为：`| 类别 | 需求描述 |`。\n"
    "**严格要求：表格只有 3 行数据（业务需求、功能需求、技术需求各 1 行），同一类别的多条需求合并到同一个单元格中，用换行分隔。**\n"
    "示例：\n"
    "| 类别 | 需求描述 |\n"
    "| --- | --- |\n"
    "| 业务需求 | 多租户隔离: 必须在逻辑和物理层面上隔离不同客户的数据。弹性吞吐: 需支持突发高并发 |\n"
    "| 功能需求 | 一键开通: 支持快速为新客户配置全套 AI 资源。全模型支持: 集成 GPT-5, GPT-4.1 等全系模型 |\n"
    "| 技术需求 | 高可用性: 单个实例故障不影响整体服务。年消耗目标: 维持预算内的消耗水平 |\n\n"
    "## 五、详细解决方案设计\n"
    "这是最核心的章节。**严格禁止使用项目符号列表（-、*、• 等）。**\n"
    "必须使用段落叙述的方式，用加粗的关键词引导每个设计要点，例如：\n"
    "**控制平面:** 位于 East US 2 的 admin 系列资源组。这是为客户开设的'管理账户'...\n"
    "**数据平面:** 位于 East US 的 test 系列资源组。这是算力发电厂...\n"
    "每个设计要点用一个完整的段落描述，不要拆成多个子列表。内容简洁精炼，每个要点 2-3 句话即可。\n\n"
    "## 六、安全架构\n"
    "用段落叙述数据隔离、身份认证等安全设计。不要用列表。每个要点用加粗关键词引导。\n\n"
    "## 七、集成架构\n"
    "用段落叙述集成方案。不要用列表。每个要点用加粗关键词引导。\n\n"
    "## 八、资源架构\n"
    "### Azure 资源需求\n"
    "以 Markdown 表格形式列出所有 Azure 资源，表头必须为：`| 服务名称 | 配置规格 (SKU) | 区域 | 核心用途 |`。资源数量控制在 5-8 行。\n"
    "示例：\n"
    "| 服务名称 | 配置规格 (SKU) | 区域 | 核心用途 |\n"
    "| --- | --- | --- | --- |\n"
    "| Azure AI Vision | Standard | East US 2 | 核心功能。假设每月有 100 万次动作识别请求。 |\n"
    "| Azure OpenAI | GPT-4o-mini | East US 2 | 支撑每月数亿 Token 的高频陌伴对话。 |\n"
    "| Azure AI Speech | Standard (S0) | East US 2 | TTS 语音播报。 |\n\n"
    "**全局格式要求（极其重要）：**\n"
    "- 章节标题使用 `## 一、摘要` 格式（## 开头 + 中文数字编号）\n"
    "- **严格禁止使用项目符号列表（-、*、• 开头的行）。** 全文必须使用段落叙述\n"
    "- **加粗关键词引导的要点必须每个单独成段（单独一行），不要把多个加粗要点拼在同一个段落中**\n"
    "- **严禁对专业术语缩写进行括号解释。** 例如：写 RAG，不要写 RAG（检索增强生成）；写 APIM，不要写 APIM（API Management）\n"
    "- **严禁使用模棱两可的表述。** 不要写'建议'、'例如'、'或'、'可以选择'。你必须根据客户预算和场景直接做出决策\n"
    "- **必须使用具体的模型名称和 SKU**，如 GPT-4o、GPT-4o-mini，不要写 Azure OpenAI Service\n"
    "- **必须选择一个确定的 Azure 区域**，不要写 East Asia 或 Southeast Asia 这种二选一表述\n"
    "- 内容要精炼简洁，每个章节不超过模板文档的篇幅\n"
    "- 表格必须使用 Markdown 表格语法\n\n"
    "**重要：** 下方会提供一份【参考模板文档】，你必须严格学习它的写作风格（段落叙述，非列表）、内容篇幅和表格格式。以完全相同的结构和风格为新客户生成内容。"
)

POV_SYSTEM_PROMPT = (
    "你是一位经验丰富的 Microsoft 技术方案交付专家。"
    "请根据用户提供的【解决方案架构文档】、【客户名称】、【POV周期】以及【甲乙方项目人员名单】，生成一份 POV Deployment Plan。\n\n"
    "**标题要求（极其重要）：** 你的输出的第一行必须是一个 `#` 标题，格式为: "
    "`# [客户名称] - [方案核心描述] POV 部署计划`。例如：\n"
    "- `# 深圳跃瓦创新科技 - \   Azure AI 中台与多场景助手 POV 部署计划`\n"
    "- `# 京华数码 - 智能外贸供应链 AI 平台 POV 部署计划`\n"
    "绝对不要使用笼统的'POV 部署计划'作为标题，必须包含具体的项目名称。\n\n"
    "**强相关要求：** POV 部署计划必须与解决方案架构文档强相关：\n"
    "部署的服务必须来自方案文档，步骤顺序符合架构依赖关系，验证场景对应核心功能。\n\n"
    "**章节结构要求（必须严格遵循以下结构）：**\n\n"
    "## 一、执行周期\n"
    "直接写出起止日期，如：2026年2月25日 - 2026年3月11日，周末不算工作日，不包含周末\n\n"
    "## 二、项目目标\n"
    "先用一句话概括总体目标和工作日天数，然后列出 3 个可衡量的目标。\n"
    "**每个目标必须简洁：** 只需要一个加粗标题和 1-2 句描述即可，不要使用子列表展开。参考格式：\n"
    "**知识检索准确率:** 验证 Azure AI Search 对产品手册的检索准确率，杜绝技术参数幻觉。\n"
    "**双模型分流:** 验证常规问答走 GPT-4o-mini 与复杂方案生成走 GPT-4o 的路由机制。\n"
    "**成本与生产规划 :** 基于压测数据证明该架构能在预算内稳定运行。\n\n"
    "## 三、核心团队成员与职责\n"
    "以 Markdown 表格形式输出，表头必须为：`| 角色 | 所属方 | 姓名 | 角色职责 |`\n"
    "根据用户提供的人员名单填充，每人用 1-2 句描述职责。\n\n"
    "## 四、分阶段详细部署计划\n"
    "由你自己智能来划分阶段，每个阶段包含：\n"
    "1. **阶段标题**：使用 `### 阶段 N: [阶段主题] ([M月D日] - [M月D日])` 格式，用 ### 标记，不要用 ** 包裹\n"
    "2. **目标描述**：紧跟标题，一句话说明核心目标\n"
    "3. **任务表格**：Markdown 表格，表头必须为：`| 日期 | 核心任务 | 主要负责人 | 里程碑与交付物 |`\n"
    "**严禁**在阶段内添加 `#### 阶段 N 任务安排` 之类的子标题。阶段标题后直接跟目标描述和表格。\n\n"
    "**日期要求（极其重要）：**\n"
    "- 任务表格中的日期必须是具体的日历日期（如 2月25日、2月26日）\n"
    "- **必须跳过周六和周日，只安排工作日**\n"
    "- 日期格式统一为：M月D日\n\n"
    
    "每天的任务必须具体、可操作。里程碑与交付物是具体产出（例如 '部署日志'、'准确率报告'、'UAT 签字单'）。\n\n"
    "**重要：** 下方会提供一份【参考模板文档】，你必须严格学习它的章节结构、分阶段格式、表格详细度和交付物命名规范。内容风格要精炼简洁，与模板保持一致。"
)

# -----------------------------------------------------------------
# Prompt 模板：SVG 现代化架构图生成器
# -----------------------------------------------------------------

SVG_SYSTEM_PROMPT = (
    "你是一位顶尖的云计算架构师和资深 UI/UX 视觉设计师，精通编写直接可渲染的、具有极高现代审美的 SVG 代码。\n"
    "请根据我提供的【架构描述】和【组件列表】，为我绘制一份企业级的 Azure 解决方案逻辑架构图。\n\n"
    "你的输出必须且只能是一段完整的 <svg> 代码，画布大小设置为 <svg width=\"1350\" height=\"900\" ...>。\n\n"
    "**【强制性视觉与排版规范（极度重要，违反将导致渲染失败！）】**：\n\n"
    "1. **绝对网格布局系统 (Strict Grid System - 防重叠核心规则)**：\n"
    "   你必须在脑海中建立一个严格的网格来放置组件，**绝对禁止任何两个卡片重叠或覆盖！**\n"
    "   - 标准卡片尺寸强制为：`width=\"240\"`，`height=\"80\"`。\n"
    "   - **X 轴 (列) 固定锚点：** \n"
    "     列 1 (接入层) X = 80\n"
    "     列 2 (网关层) X = 380\n"
    "     列 3 (AI/核心层) X = 680\n"
    "     列 4 (数据层) X = 980\n"
    "   - **Y 轴 (行) 固定步长：** \n"
    "     每个卡片高度 80，垂直间距必须至少为 40。因此 Y 轴步长必须是 120！\n"
    "     行 1 Y = 160\n"
    "     行 2 Y = 280\n"
    "     行 3 Y = 400\n"
    "     行 4 Y = 520\n"
    "     行 5 Y = 640\n"
    "     行 6 Y = 760\n"
    "   - **坐标分配警告：** 在同一个列中，每放置一个新组件，其 `translate(x, y)` 中的 `y` 值必须比上一个组件增加至少 `120`。严禁将两个组件放在如 `y=280` 和 `y=300` 这样接近的位置！\n\n"
    "2. **图层 Z-Index 隔离策略 (绝对强制)**：\n"
    "   你的 SVG 结构必须严格按照以下顺序编写，以防连线遮挡文字：\n"
    '   - 第 1 层：全局背景 <rect width="100%" height="100%" fill="url(#bgGradient)" />\n'
    "   - 第 2 层：区域划分大背景框 (Zone Backgrounds)\n"
    '   - 第 3 层：所有连线 <g id="connectors"> ... </g> (连线必须在组件之前绘制！)\n'
    '   - 第 4 层：所有服务卡片 <g id="components"> ... </g> (卡片必须有 fill="#FFF"，这样连线误差会被白色卡片背景完美遮挡)\n\n'
    "3. **SVG 定义 (Defs)**：必须在 SVG 开头包含以下 <defs>：\n"
    '<defs>\n'
    '  <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">\n'
    '    <feGaussianBlur in="SourceAlpha" stdDeviation="4"/>\n'
    '    <feOffset dx="2" dy="4" result="offsetblur"/>\n'
    '    <feComponentTransfer><feFuncA type="linear" slope="0.15"/></feComponentTransfer>\n'
    '    <feMerge><feMergeNode/><feMergeNode in="SourceGraphic"/></feMerge>\n'
    '  </filter>\n'
    '  <linearGradient id="bgGradient" x1="0%" y1="0%" x2="0%" y2="100%">\n'
    '    <stop offset="0%" style="stop-color:#F4F7FB;stop-opacity:1" />\n'
    '    <stop offset="100%" style="stop-color:#FFFFFF;stop-opacity:1" />\n'
    '  </linearGradient>\n'
    '  <marker id="arrow" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto">\n'
    '    <path d="M0,0 L0,6 L9,3 z" fill="#0078D4" />\n'
    '  </marker>\n'
    '</defs>\n\n'
    "4. **连线与路由规范 (Orthogonal Routing)**：\n"
    "   - 严禁乱穿交叉！严禁一笔画斜线！\n"
    '   - 必须使用正交折线，格式为 <path d="M x1 y1 L x2 y1 L x2 y2 L x3 y2" fill="none" stroke="#0078D4" stroke-width="2" marker-end="url(#arrow)"/>\n'
    "   - 连线起点：`x + width` (例如 80+240=320)；连线终点：目标卡片的 `x` (例如 380)。\n\n"
    "5. **卡片内容规范**：\n"
    '   - 必须使用 <g transform="translate(x, y)"> 组合服务。\n'
    "   - 示例如下：\n"
    '   <g transform="translate(680, 280)"> <!-- 注意 Y 值的网格化 -->\n'
    '     <rect width="240" height="80" rx="8" fill="#FFF" stroke="#5C2D91" stroke-width="2" filter="url(#shadow)"/>\n'
    '     <text x="120" y="35" font-family="\'Segoe UI\', sans-serif" font-size="15" fill="#333" font-weight="bold" text-anchor="middle">Azure OpenAI</text>\n'
    '     <text x="120" y="55" font-family="\'Segoe UI\', sans-serif" font-size="12" fill="#666" text-anchor="middle">GPT-4o-mini (主力引擎)</text>\n'
    '   </g>\n\n'
    "6. **Azure 品牌配色参考**：AI服务(紫色 #5C2D91), 网关/基础计算(蓝色 #0078D4), 数据/存储(青色 #008272), 安全(红色 #D13438)。\n\n"
    "**【输出要求】**：\n"
    "开始编写前，请在脑海中严格为每个组件分配 (列X, 行Y) 的坐标，确保同一个列中的 Y 坐标以 120 的倍数递增（160, 280, 400...）。\n"
    "直接输出完整的 XML/SVG 代码，禁止输出任何解释性文字。\n"
    'SVG 标签中务必包含 xmlns="http://www.w3.org/2000/svg"。确保闭合所有标签。'
)

CSV_SYSTEM_PROMPT = (
    "你是一位 Azure 迁移专家。用户会提供一份 Azure 价格估算表（包含资源名称、SKU、估算金额等）和一份 Azure Migrate 导入 CSV 模板。\n\n"
    "你的任务是：\n"
    "1. 分析价格估算表中的资源列表和金额\n"
    "2. 倒推客户在本地环境可能使用什么配置的 VM\n"
    "3. 按照 CSV 模板格式填充数据\n\n"
    "**CSV 填充规则：**\n"
    "- **必填列**：*Server name, *Cores, *Memory (In MB), *OS name\n"
    "- **Server name 格式**：用描述性命名，格式为 `服务类型-区域-规模-序号`，例如：\n"
    "  TTS-EastAsia-37M-01, LLM-4o-EUS2-01, Search-S1-EastAsia-01, CosmosDB-Serverless-01, Storage-Blob-2TB-01\n"
    "- **Cores/Memory**：根据价格表中 Azure VM/服务规格倒推，常规值：4/8192, 8/16384, 8/32768\n"
    "- **OS name**：大多数写 Linux，数据库类可写 Linux\n"
    "- **磁盘字段也必须填写**：\n"
    "  - Number of disks: 根据服务类型填 1 或 2\n"
    "  - Disk 1 size (In GB): 64/128/256/512/2048 等\n"
    "  - Disk 1 read throughput (MB per second): 根据时期情况填写\n"
    "  - Disk 1 write throughput (MB per second):  根据时期情况填写\n"
    "  - Disk 1 read ops (operations per second): 根据时期情况填写\n"
    "  - Disk 1 write ops (operations per second): 根据时期情况填写\n"
    "  - 如果有第二块盘（数据库、搜索等服务），填写 Disk 2 的相同字段\n"
    "- 其他列留空\n"
    "- 预估 VM 总数量和配置要合理，使得迁移后的 Azure 费用大约等于用户提供的年消耗预算\n\n"
    "**输出格式：**\n"
    "只输出纯粹的 CSV 内容（包含表头行），不要输出 Markdown 代码块标记、解释性文字或其他内容。"
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
                          color_rgb=None, alignment=None, indent=True):
    """添加一个带完整样式的段落。indent=True 时添加首行缩进。"""
    p = doc.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    # 首行缩进（约 1 个 Tab = 0.74cm）
    if indent and alignment is None:
        p.paragraph_format.first_line_indent = Cm(0.74)
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

        # ── 跳过 --- 分隔线 ──
        if stripped == '---' or stripped == '***' or stripped == '___':
            i += 1
            continue

        # ── 标题 ──
        if stripped.startswith("#### "):
            _add_styled_heading(doc, stripped[5:], level=4)
            i += 1
            continue
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

        # ── 独立的 **加粗行**（如阶段标题），转为三级标题 ──
        if stripped.startswith("**") and stripped.endswith("**") and len(stripped) > 4:
            title_text = stripped[2:-2]
            _add_styled_heading(doc, title_text, level=3)
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
# 辅助：日期前缀文件名
# ──────────────────────────────────────────────
def _date_prefix():
    """返回当前日期前缀，如 0225"""
    return datetime.date.today().strftime("%m%d")


# ──────────────────────────────────────────────
# 主界面
# ──────────────────────────────────────────────
def main():
    st.markdown('<div class="main-title">微软客户POE文档生成器</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">🤖 拒绝 Ctrl+CV，摸鱼工程师的终极救星！一键生成让老板狂喜的文档！</div>',
        unsafe_allow_html=True,
    )

    if not check_secrets():
        st.stop()

    # 侧边栏
    with st.sidebar:
        st.markdown("### 操作")
        if st.button("清除所有结果", use_container_width=True):
            for key in ["solution_text", "pov_text", "customer_name", "svg_code", "csv_code", "budget"]:
                st.session_state.pop(key, None)
            st.rerun()

        st.markdown("---")
        st.markdown("### 模板状态")
        sol_ok = os.path.exists(SOLUTION_TEMPLATE_PATH)
        pov_ok = os.path.exists(POV_TEMPLATE_PATH)
        csv_ok = os.path.exists(MIGRATE_TEMPLATE_PATH)
        st.markdown(f"- Solution: {'OK' if sol_ok else 'Missing'}")
        st.markdown(f"- POV: {'OK' if pov_ok else 'Missing'}")
        st.markdown(f"- CSV: {'OK' if csv_ok else 'Missing'}")

    solution_ref = extract_template_text(SOLUTION_TEMPLATE_PATH) if sol_ok else ""
    pov_ref = extract_template_text(POV_TEMPLATE_PATH) if pov_ok else ""

    # ════════════════════════════════════════════════════
    # 公共输入区域
    # ════════════════════════════════════════════════════
    st.markdown("### 客户信息")
    c1, c2 = st.columns([2, 1])
    with c1:
        customer_name = st.text_input("客户名称 (必填)", placeholder="例如：宇宙无敌科技有限公司")
    with c2:
        budget = st.text_input("预估年消耗 (USD)", placeholder="越多越好，例如：500k+")

    customer_bg = st.text_area(
        "客户背景信息",
        placeholder="请粘贴从 Web 搜索到的客户背景资料，包括行业、规模、现有 IT 环境、核心需求等...",
        height=150,
    )

    st.divider()

    # ════════════════════════════════════════════════════
    # Tab 布局
    # ════════════════════════════════════════════════════
    tab_sol, tab_pov, tab_svg, tab_csv = st.tabs([
        "AI 解决方案文档", "POV 部署计划", "架构图 (SVG)", "Azure Migrate CSV"
    ])

    dp = _date_prefix()  # 日期前缀

    # ─────────── Tab 1: 解决方案文档 ───────────
    with tab_sol:
        left, right = st.columns([1, 1])
        with left:
            has_solution = "solution_text" in st.session_state
            sol_label = "重新生成" if has_solution else "生成解决方案架构文档"
            if st.button(sol_label, type="primary", use_container_width=True, key="btn_sol"):
                if not customer_name.strip():
                    st.warning("请输入客户名称。")
                    st.stop()
                if not customer_bg.strip():
                    st.warning("请输入客户背景信息。")
                    st.stop()
                try:
                    with st.spinner("正在生成解决方案架构文档..."):
                        user_ctx = (
                            f"## 客户信息\n- **客户名称**：{customer_name}\n"
                            f"- **预估年消耗 (USD)**：{budget}\n\n"
                            f"## 客户背景\n{customer_bg}"
                        )
                        if solution_ref:
                            user_ctx += (
                                f"\n\n---\n\n## 【参考模板文档 —— 请学习其风格和结构，不要照抄具体数据】\n\n"
                                f"{solution_ref}"
                            )
                        sol_text = call_azure_openai(SOLUTION_SYSTEM_PROMPT, user_ctx)
                        st.session_state["solution_text"] = sol_text
                        st.session_state["customer_name"] = customer_name
                        st.session_state["budget"] = budget
                        st.session_state.pop("pov_text", None)
                        st.session_state.pop("svg_code", None)
                    st.rerun()
                except Exception as e:
                    st.error(f"生成失败：{e}")

            if "solution_text" in st.session_state:
                customer = st.session_state["customer_name"]
                docx_sol = create_solution_docx(
                    content=st.session_state["solution_text"], customer_name=customer
                )
                st.download_button(
                    label="下载解决方案架构文档 (.docx)",
                    data=docx_sol,
                    file_name=f"{dp}-{customer}-AI解决方案架构文档.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        with right:
            if "solution_text" in st.session_state:
                st.markdown("**文档预览**")
                st.markdown(st.session_state["solution_text"], unsafe_allow_html=True)
            else:
                st.info("请先生成解决方案文档")

    # ─────────── Tab 2: POV 部署计划 ───────────
    with tab_pov:
        if "solution_text" not in st.session_state:
            st.info("请先在「AI 解决方案文档」标签页中生成解决方案文档")
        else:
            customer = st.session_state["customer_name"]
            solution = st.session_state["solution_text"]
            left, right = st.columns([1, 1])
            with left:
                dc1, dc2 = st.columns(2)
                with dc1:
                    pov_start = st.date_input("POV 开始日期", value=datetime.date.today())
                with dc2:
                    pov_end = st.date_input(
                        "POV 结束日期",
                        value=datetime.date.today() + datetime.timedelta(days=14),
                    )

                vendor_team = st.text_area(
                    "乙方项目人员（我方团队）",
                    value=(
                        "技术负责人: 吕兴安\n"
                        "Azure架构师: alex\n"
                    ),
                    height=120,
                    help="只需填写乙方（我方）人员，甲方人员由 AI 根据客户背景自动生成",
                )

                has_pov = "pov_text" in st.session_state
                pov_label = "重新生成" if has_pov else "生成 POV 部署计划"
                if st.button(pov_label, type="primary", use_container_width=True, key="btn_pov"):
                    try:
                        pov_period = f"{pov_start.strftime('%Y/%m/%d')} - {pov_end.strftime('%Y/%m/%d')}"
                        pov_prompt = (
                            f"以下是已生成的解决方案架构文档，请据此生成 POV 部署计划：\n\n"
                            f"{solution}\n\n"
                            f"## 补充信息\n- **客户名称**：{customer}\n"
                            f"- **POV 周期**：{pov_period}\n\n"
                            f"## 乙方项目人员\n{vendor_team}\n\n"
                            f"请根据客户背景信息自动生成合理的甲方人员（2-3人，包含项目负责人和技术对接人）。"
                        )
                        if pov_ref:
                            pov_prompt += (
                                f"\n\n---\n\n## 【参考模板文档 —— 请学习其风格和结构，不要照抄具体数据】\n\n"
                                f"{pov_ref}"
                            )
                        with st.spinner("正在生成 POV 部署计划..."):
                            pov_text = call_azure_openai(POV_SYSTEM_PROMPT, pov_prompt)
                            st.session_state["pov_text"] = pov_text
                        st.rerun()
                    except Exception as e:
                        st.error(f"生成失败：{e}")

                if "pov_text" in st.session_state:
                    docx_pov = create_pov_docx(
                        content=st.session_state["pov_text"], customer_name=customer
                    )
                    st.download_button(
                        label="下载 POV 部署计划 (.docx)",
                        data=docx_pov,
                        file_name=f"{dp}-{customer}-POV部署计划.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

            with right:
                if "pov_text" in st.session_state:
                    st.markdown("**文档预览**")
                    st.markdown(st.session_state["pov_text"], unsafe_allow_html=True)
                else:
                    st.info("请填写信息后点击生成")

    # ─────────── Tab 3: SVG 架构图 ───────────
    with tab_svg:
        if "solution_text" not in st.session_state:
            st.info("请先在「AI 解决方案文档」标签页中生成解决方案文档")
        else:
            customer = st.session_state["customer_name"]
            solution = st.session_state["solution_text"]

            has_svg = "svg_code" in st.session_state
            svg_label = "重新生成架构图" if has_svg else "生成 SVG 架构图"
            if st.button(svg_label, type="primary", use_container_width=True, key="btn_svg"):
                try:
                    svg_prompt = (
                        f"请根据以下解决方案架构文档生成一张完整的 SVG 架构图：\n\n"
                        f"{solution}"
                    )
                    with st.spinner("正在生成 SVG 架构图..."):
                        svg_raw = call_azure_openai(SVG_SYSTEM_PROMPT, svg_prompt)
                        import re as _re
                        svg_match = _re.search(r'(<svg[\s\S]*?</svg>)', svg_raw, _re.IGNORECASE)
                        svg_code = svg_match.group(1) if svg_match else svg_raw
                        st.session_state["svg_code"] = svg_code
                    st.rerun()
                except Exception as e:
                    st.error(f"生成失败：{e}")

            if "svg_code" in st.session_state:
                svg = st.session_state["svg_code"]

                import streamlit.components.v1 as components
                preview_html = f'''
                <div style="background: #fff; border: 1px solid #e0e0e0; border-radius: 8px;
                            padding: 20px; overflow-x: auto;">
                    {svg}
                </div>
                '''
                components.html(preview_html, height=700, scrolling=True)

                dl_col, code_col = st.columns(2)
                with dl_col:
                    st.download_button(
                        label="下载 SVG 架构图",
                        data=svg,
                        file_name=f"{dp}-{customer}-架构图.svg",
                        mime="image/svg+xml",
                        use_container_width=True,
                    )
                with code_col:
                    with st.expander("查看 SVG 代码"):
                        # 一键复制按钮
                        import base64 as _b64
                        svg_b64 = _b64.b64encode(svg.encode("utf-8")).decode("ascii")
                        copy_html = f'''
                        <textarea id="svgSrc" style="position:absolute;left:-9999px">{svg_b64}</textarea>
                        <button id="copyBtn" onclick="
                            var b64=document.getElementById('svgSrc').value;
                            var txt=decodeURIComponent(escape(atob(b64)));
                            navigator.clipboard.writeText(txt).then(function(){{
                                document.getElementById('copyBtn').textContent='已复制';
                                setTimeout(function(){{document.getElementById('copyBtn').textContent='复制 SVG 代码'}},2000);
                            }});
                        " style="background:#0078D4;color:#fff;border:none;border-radius:6px;padding:6px 18px;cursor:pointer;font-size:14px;">
                            复制 SVG 代码
                        </button>
                        '''
                        import streamlit.components.v1 as _comp
                        _comp.html(copy_html, height=45)
                        st.code(svg, language="xml")
            else:
                st.info("点击上方按钮生成架构图")

    # ─────────── Tab 4: Azure Migrate CSV ───────────
    with tab_csv:
        if "solution_text" not in st.session_state:
            st.info("请先在「AI 解决方案文档」标签页中生成解决方案文档")
        else:
            customer = st.session_state["customer_name"]
            bdgt = st.session_state.get("budget", budget)
            left, right = st.columns([1, 1])
            with left:
                migrate_csv_header = ""
                if os.path.exists(MIGRATE_TEMPLATE_PATH):
                    with open(MIGRATE_TEMPLATE_PATH, "r", encoding="utf-8-sig") as f:
                        migrate_csv_header = f.readline().strip()

                uploaded_excel = st.file_uploader(
                    "上传价格估算表 (.xlsx)",
                    type=["xlsx"],
                    help="上传包含 Azure 资源估算金额的 Excel 文件",
                )

                has_csv = "csv_code" in st.session_state
                csv_label = "重新生成 CSV" if has_csv else "生成 Azure Migrate CSV"
                if st.button(csv_label, type="primary", use_container_width=True, key="btn_csv"):
                    if not uploaded_excel:
                        st.warning("请先上传价格估算表 Excel 文件。")
                        st.stop()
                    if not migrate_csv_header:
                        st.warning("Azure Migrate CSV 模板未找到。")
                        st.stop()
                    try:
                        import openpyxl
                        wb = openpyxl.load_workbook(uploaded_excel, data_only=True)
                        excel_text_parts = []
                        for sheet_name in wb.sheetnames:
                            ws = wb[sheet_name]
                            rows = list(ws.iter_rows(values_only=True))
                            if not rows:
                                continue
                            excel_text_parts.append(f"### Sheet: {sheet_name}")
                            headers = [str(c) if c is not None else "" for c in rows[0]]
                            excel_text_parts.append("| " + " | ".join(headers) + " |")
                            excel_text_parts.append("| " + " | ".join(["---"] * len(headers)) + " |")
                            for row in rows[1:]:
                                cells = [str(c) if c is not None else "" for c in row]
                                excel_text_parts.append("| " + " | ".join(cells) + " |")
                        excel_text = "\n".join(excel_text_parts)

                        csv_prompt = (
                            f"以下是客户的 Azure 价格估算表内容：\n\n{excel_text}\n\n"
                            f"客户预估年消耗：{bdgt}\n\n"
                            f"Azure Migrate CSV 模板表头：\n{migrate_csv_header}\n\n"
                            f"请根据价格估算表倒推本地 VM 配置，按模板格式生成 CSV。"
                        )

                        with st.spinner("正在生成 Azure Migrate CSV..."):
                            csv_raw = call_azure_openai(CSV_SYSTEM_PROMPT, csv_prompt)
                            csv_clean = csv_raw.strip()
                            if csv_clean.startswith("```"):
                                csv_clean = csv_clean.split("\n", 1)[1] if "\n" in csv_clean else csv_clean
                            if csv_clean.endswith("```"):
                                csv_clean = csv_clean[:-3].strip()
                            st.session_state["csv_code"] = csv_clean
                        st.rerun()
                    except Exception as e:
                        st.error(f"生成失败：{e}")

                if "csv_code" in st.session_state:
                    csv_data = st.session_state["csv_code"]
                    st.download_button(
                        label="下载 Azure Migrate CSV",
                        data=csv_data.encode("utf-8-sig"),
                        file_name=f"{dp}-{customer}-AzureMigrate.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

            with right:
                if "csv_code" in st.session_state:
                    csv_data = st.session_state["csv_code"]
                    st.markdown("**CSV 预览**")
                    try:
                        import csv as csv_mod
                        csv_lines = csv_data.strip().split("\n")
                        reader = csv_mod.reader(csv_lines)
                        all_rows = list(reader)
                        if len(all_rows) > 1:
                            header = all_rows[0]
                            num_cols = len(header)
                            # 对齐列数：补齐或截断
                            data_rows = []
                            for row in all_rows[1:]:
                                if len(row) < num_cols:
                                    row = row + [""] * (num_cols - len(row))
                                elif len(row) > num_cols:
                                    row = row[:num_cols]
                                data_rows.append(row)
                            import pandas as pd
                            df = pd.DataFrame(data_rows, columns=header)
                            st.dataframe(df, use_container_width=True)
                    except Exception as e:
                        st.warning(f"预览失败，请使用下载查看: {e}")
                        st.code(csv_data, language="csv")
                else:
                    st.info("上传 Excel 后点击生成")


# ──────────────────────────────────────────────
# 入口
# ──────────────────────────────────────────────
if __name__ == "__main__":
    main()

