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
INFRA_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "Infra_template.docx")
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
        color: #1F4E79;
        font-size: 2.2rem;
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
        border: 1px solid #E0E0E0;
        border-radius: 12px;
        padding: 1.5rem;
        background-color: #F8F9FA;
    }
    .stFormSubmitButton > button {
        background-color: #1F4E79 !important;
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
    "请根据用户提供的【客户名称】和【背景信息】，生成一份完整、专业的 AI 售前解决方案架构文档。\n\n"
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
    "**严格要求：表格只有 3 行数据（业务需求、功能需求、技术需求各 1 行），每格仅写 1-2 个关键需求点，用分号分隔，不展开解释。**\n"
    "示例：\n"
    "| 类别 | 需求描述 |\n"
    "| --- | --- |\n"
    "| 业务需求 | 多租户数据隔离；高并发弹性吞吐 |\n"
    "| 功能需求 | 一键开通 AI 资源；支持全系模型接入 |\n"
    "| 技术需求 | 跨实例高可用；私有网络访问 |\n\n"
    "## 五、详细解决方案设计\n"
    "本节分为两部分，格式严格如下：\n\n"
    "**第一部分（解决方案预览）：** 用 1-2 句纯文字段落，简要描述整体方案的核心部署思路和区域选择。不使用列表，不加粗，无符号，无表情，无卡片。\n\n"
    "**第二部分（详细资源用途）：** 紧接第一部分，直接列出每个 Azure 资源的详细用途，格式严格为：资源名称: 详细用途描述（1-2句）。每个资源单独占一行，资源名称与正文之间用冒号加空格分隔，不加粗资源名称，不使用项目符号（-、*、•），不使用任何表情或卡片。控制在 4-6 个资源行。\n"
    "示例（严格照此格式，不照抄内容）：\n"
    "Azure OpenAI (GPT-4o): 作为核心推理引擎，处理用户自然语言查询，生成个性化推荐和客服回复。\n"
    "Azure AI Speech: 提供语音识别与语音合成能力，支撑语音交互入口和呼叫中心坐席辅助场景。\n"
    "Azure AI Search: 构建向量检索索引，对接产品知识库，为模型提供精准的 RAG 上下文。\n"
    "Azure API Management: 统一管理所有 AI 服务调用入口，实现限流、鉴权及 Token 消耗监控。\n"
    "绝对禁止在第二部分再拆分子功能列表或多个换行子句，每个资源描述必须是单独一行。\n\n"
    "## 六、安全架构\n"
    "格式与详细设计完全相同：每个要点 `关键词: 正文` 在同一行，不加粗关键词。控制在 2-3 个要点。例如：\n"
    "数据沙箱: 利用 AI Foundry 项目隔离机制，确保各租户数据在物理存储上完全隔离。\n"
    "托管标识: 所有资源间调用通过 Azure AD 托管标识认证，杜绝 API Key 泄露风险。\n\n"
    "## 七、集成架构\n"
    "格式与安全架构完全相同：每个要点 `关键词: 正文` 在同一行，不加粗关键词。控制在 2-3 个要点。例如：\n"
    "RedSteed SDK: 向客户提供封装好的 SDK，传入 Tenant-ID 即可路由至对应资源池。\n"
    "智能网关: 部署在 APIM 上，解析租户订阅等级并路由请求，实时记录 Token 消耗。\n\n"
    "## 八、资源架构\n"
    "### Azure 资源需求\n"
    "以 Markdown 表格形式列出所有 Azure 资源，表头必须为：`| 服务名称 | 配置规格 (SKU) | 区域 | 核心用途 |`。资源数量控制在 5-8 行。\n"
    "**严格禁止将 Azure AI Foundry（含 AI Studio、AI Foundry Hub、AI Foundry Project 等）列入此表。** AI Foundry 是开发门户，不计入客户部署资源清单。\n"
    "示例：\n"
    "| 服务名称 | 配置规格 (SKU) | 区域 | 核心用途 |\n"
    "| --- | --- | --- | --- |\n"
    "| Azure OpenAI | GPT-4o | East US 2 | 核心 AI 推理，支撑主要业务场景。 |\n"
    "| Azure AI Speech | Standard (S0) | East US 2 | TTS 语音播报。 |\n\n"
    "**全局格式要求（极其重要）：**\n"
    "- 章节标题使用 `## 一、摘要` 格式（## 开头 + 中文数字编号）\n"
    "- **严格禁止使用项目符号列表（-、*、• 开头的行）。** 全文必须使用段落叙述\n"
    "- **每个要点（`关键词: 正文` 格式）必须单独成段（单独一行），绝不加粗关键词，不要把多个要点拼在同一段落中**\n"
    "- **严禁对专业术语缩写进行括号解释。** 例如：写 RAG，不要写 RAG（检索增强生成）\n"
    "- **严禁在文档任何位置提及预算金额或年消耗数字**\n"
    "- **根据客户业务场景选择合适的大模型**（GPT-5/GPT-4o 适合通用场景，o1/o3-mini 适合推理场景，GPT-4o-mini 适合高并发轻量场景），必须写具体模型名称\n"
    "- **必须选择全球 Azure 区域**（East US、East US 2、West Europe、Southeast Asia、East Asia、Japan East 等），**严禁使用中国区域（China East、China North 等）**\n"
    "- 内容要精炼简洁，严格对齐参考模板的篇幅，不要更长\n"
    "- 表格必须使用 Markdown 表格语法\n\n"
    "**重要：** 下方会提供一份【参考模板文档】，你必须严格学习它的写作风格（段落叙述，非列表）、内容篇幅和表格格式。以完全相同的结构和风格为新客户生成内容。"
)

INFRA_SYSTEM_PROMPT = (
    "你是一位顶级的 Microsoft Azure 基础设施解决方案架构师。"
    "请根据用户提供的【客户名称】和【背景信息】，生成一份完整、专业的 Azure 基础设施解决方案架构文档。\n\n"
    "**标题要求（极其重要）：** 你的输出的第一行必须是一个 `#` 标题，格式为: `# [客户名称] - [具体方案名称]`。"
    "方案名称必须具体且针对客户业务，例如：\n"
    "- `# 新疆云基智能科技 - 全球智能制造与可穿戴物联网云平台解决方案`\n"
    "- `# 深圳跃瓦创新科技 - 混合云智慧工厂 IaaS 底座方案`\n"
    "绝对不要使用笼统的'基础设施解决方案架构文档'作为标题。\n\n"
    "**章节结构要求（必须严格遵循以下 10 个章节）：**\n\n"
    "## 一、执行摘要\n"
    "2-3 句话概述方案核心思路和预期价值。保持简洁。\n\n"
    "## 二、解决方案架构概览\n"
    "2-3 段话概述整体架构设计理念（如全球分布式接入、混合云底座等）。用段落叙述，不要用列表。\n\n"
    "## 三、业务背景\n"
    "用段落叙述客户的行业定位、痛点和机遇。可使用加粗关键词引导要点（如 **跨国网络延迟高:**、**数据合规风险:** 等）。\n\n"
    "## 四、需求概述\n"
    "分三类叙述：业务需求、功能需求、技术需求。每类仅 1-2 个关键需求点，用段落叙述，不要用列表，不展开解释。\n\n"
    "## 五、解决方案设计\n"
    "先写 1 句总体部署概述（说明 Azure 全球区域和核心策略）。然后每个要点写成 `关键词: 正文描述` 的形式，关键词和正文在同一行，不加粗关键词，控制在 3-5 个要点。例如：\n"
    "统一网络出口: 利用 Azure Front Door 提供全球应用入口负载均衡、CDN 加速及 WAF 防护。\n"
    "微服务计算中枢: 部署 AKS 生产集群，配置多节点池，支持早晚高峰弹性扩缩容。\n"
    "高可用 IaaS 集群: 使用 E 系列内存优化型 VM 搭建高可用集群，支撑 ERP/MES 等信息系统。\n\n"
    "## 六、详细解决方案架构\n"
    "针对每个核心 Azure 服务，写成 `服务名称: 配置和用途描述` 的形式，服务名称和正文在同一行，不加粗服务名称，每个服务仅 1-2 句话。例如：\n"
    "Azure IoT Hub: Standard S2，承担每天千万条级别的设备心跳与健康数据双向通信。\n"
    "Azure Kubernetes Service: Standard 规模集，开启自动扩缩容，运行所有应用层微服务逻辑。\n"
    "严禁展开为子列表，每个服务单独成段。\n\n"
    "## 七、集成架构\n"
    "每个要点写成 `关键词: 正文` 在同一行，不加粗关键词，仅 1 句话描述，不展开。例如：\n"
    "消息流转: IoT Hub 数据流通过 Event Hubs 路由分发，报警数据送入 AKS 实时处理，历史数据投递 Blob 归档。\n"
    "混合云互联: 通过 VPN Gateway 建立总部机房与 Azure VNet 的安全 IPsec 隧道。\n\n"
    "## 八、数据架构\n"
    "按热/温/冷三层各写 1 句话：热数据存储介质与用途；温数据；冷数据。不展开。\n\n"
    "## 九、安全架构\n"
    "每个要点写成 `关键词: 正文` 在同一行，不加粗关键词，仅 1 句话描述，不展开。例如：\n"
    "网络隔离: 所有 PaaS 数据库通过 Private Endpoint 注入 VNet，完全关闭公网访问端点。\n"
    "零信任访问: 部署 Azure Bastion，运维人员通过浏览器加密会话管理 VM，杜绝 RDP/SSH 公网暴露。\n\n"
    "## 十、基础设施需求\n"
    "以 Markdown 表格形式列出所有 Azure 资源，表头必须为：`| 服务名称 | 配置规格 | 区域 | 核心用途 |`。\n"
    "资源数量控制在 6-10 行，涵盖计算、存储、网络、数据库等核心组件。\n"
    "示例：\n"
    "| 服务名称 | 配置规格 | 区域 | 核心用途 |\n"
    "| --- | --- | --- | --- |\n"
    "| AKS | Standard + 6x D4s_v5 Node Pool | Southeast Asia | 弹性承载微服务、应用后端 |\n"
    "| Virtual Machines | 4x E8s_v5 + P15 SSD | Southeast Asia | 支撑 ERP/MES/传统单体系统 |\n\n"
    "**全局格式要求（极其重要）：**\n"
    "- 章节标题使用 `## 一、执行摘要` 格式（## 开头 + 中文数字编号）\n"
    "- **严格禁止使用项目符号列表（-、*、• 开头的行）。** 全文必须使用段落叙述\n"
    "- **每个要点（`关键词: 正文` 格式）必须单独成段（单独一行），绝不加粗关键词**\n"
    "- **严禁对专业术语缩写进行括号解释**\n"
    "- **必须选择全球 Azure 区域**（East US、East US 2、West Europe、Southeast Asia、East Asia、Japan East 等），**严禁使用中国区域（China East、China North 等）**\n"
    "- **严禁在文档任何位置提及预算金额或年消耗数字**\n"
    "- 内容要精炼简洁，严格对齐参考模板的篇幅，不要更长\n"
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
    "直接写出起止日期，格式如：2026年2月25日 - 2026年3月11日。不要在文档中提及周末或工作日相关文字。\n\n"
    "## 二、项目目标\n"
    "先用一句话概括总体目标，然后列出 3 个可衡量的目标（不要提工作日天数）。\n"
    "**每个目标必须简洁：** 只需要一个加粗标题和 1-2 句描述即可，不要使用子列表展开。参考格式：\n"
    "**知识检索准确率:** 验证 Azure AI Search 对产品手册的检索准确率，杜绝技术参数幻觉。\n"
    "**双模型分流:** 验证常规问答走 GPT-4o-mini 与复杂方案生成走 GPT-4o 的路由机制。\n"
    "**成本与生产规划 :** 基于压测数据证明该架构能在预算内稳定运行。\n\n"
    "## 三、核心团队成员与职责\n"
    "以 Markdown 表格形式输出，表头必须为：`| 角色 | 所属方 | 姓名 | 角色职责 |`\n"
    "根据用户提供的人员名单填充，每人用 1-2 句描述职责。\n\n"
    "## 四、分阶段详细部署计划\n"
    "**严格限制：阶段总数最多 3 个，禁止超过 3 个阶段。** 由你自己智能划分，每个阶段包含：\n"
    "1. **阶段标题**：使用 `### 阶段 N: [阶段主题] ([M月D日] - [M月D日])` 格式，用 ### 标记，不要用 ** 包裹\n"
    "2. **目标描述**：紧跟标题，一句话说明核心目标\n"
    "3. **任务表格**：Markdown 表格，表头必须为：`| 日期 | 核心任务 | 主要负责人 | 里程碑与交付物 |`\n"
    "**严禁**在阶段内添加 `#### 阶段 N 任务安排` 之类的子标题。阶段标题后直接跟目标描述和表格。\n\n"
    "**日期要求（内部规则，不得出现在文档正文描述中）：**\n"
    "- 任务表格中的日期必须是具体的日历日期（如 2月25日、2月26日）\n"
    "- 周六、周日不排任务，只使用工作日日期，但文档正文中绝对不要出现'周末'、'工作日'等字样\n"
    "- 日期格式统一为：M月D日\n\n"
    
    "每天的任务必须具体、可操作。里程碑与交付物是具体产出（例如 '部署日志'、'准确率报告'、'UAT 签字单'）。\n\n"
    "**重要：** 下方会提供一份【参考模板文档】，你必须严格学习它的章节结构、分阶段格式、表格详细度和交付物命名规范。内容风格要精炼简洁，与模板保持一致。"
)

# -----------------------------------------------------------------
# (SVG 架构图功能已移除)
# -----------------------------------------------------------------

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


def create_infra_docx(content: str, customer_name: str) -> bytes:
    """
    基于 Infra 模板生成基础设施解决方案 Word 文档。
    布局: 封面标题（独占一页） → 目录（独占一页） → 正文
    """
    doc = _load_template(INFRA_TEMPLATE_PATH)
    title = _extract_title(content, f"{customer_name} - 基础设施解决方案架构文档")
    body_content = _strip_first_heading(content)

    # ---- 第 1 页：封面标题 ----
    for _ in range(8):
        doc.add_paragraph()

    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cover.add_run(title)
    # 与 AI 解决方案一致: 18pt #4874CB
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
    st.markdown('<div class="main-title">微软客户 POE 文档生成器</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">专注、高效、专业的文档自动化工作流解决方案。</div>',
        unsafe_allow_html=True,
    )

    if not check_secrets():
        st.stop()

    # 侧边栏
    with st.sidebar:
        st.markdown("### 操作")
        if st.button("清除所有结果", use_container_width=True):
            for key in ["solution_text", "infra_text", "pov_text", "customer_name", "account_name", "csv_code", "budget", "doc_type", "yearly_excel_bytes", "yearly_excel_name", "yearly_messages"]:
                st.session_state.pop(key, None)
            st.rerun()

        st.markdown("---")
        st.markdown("### 模板状态")
        sol_ok = os.path.exists(SOLUTION_TEMPLATE_PATH)
        infra_ok = os.path.exists(INFRA_TEMPLATE_PATH)
        pov_ok = os.path.exists(POV_TEMPLATE_PATH)
        csv_ok = os.path.exists(MIGRATE_TEMPLATE_PATH)
        st.markdown(f"- AI Solution: {'OK' if sol_ok else 'Missing'}")
        st.markdown(f"- Infra: {'OK' if infra_ok else 'Missing'}")
        st.markdown(f"- POV: {'OK' if pov_ok else 'Missing'}")
        st.markdown(f"- CSV: {'OK' if csv_ok else 'Missing'}")

    solution_ref = extract_template_text(SOLUTION_TEMPLATE_PATH) if sol_ok else ""
    infra_ref = extract_template_text(INFRA_TEMPLATE_PATH) if infra_ok else ""
    pov_ref = extract_template_text(POV_TEMPLATE_PATH) if pov_ok else ""

    # ════════════════════════════════════════════════════
    # 公共输入区域
    # ════════════════════════════════════════════════════
    st.markdown("### 客户信息")
    c0, c1, c2 = st.columns([1.5, 2, 1])
    with c0:
        account_name = st.text_input("账户名 (必填)", placeholder="例如：Tetherflow", help="用于生成下载文件名的前缀，例如：Tetherflow")
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
    tab_sol, tab_pov, tab_csv, tab_yearly = st.tabs([
        "解决方案文档", "POV 部署计划", "Azure Migrate CSV", "年度价格表"
    ])

    dp = _date_prefix()  # 日期前缀

    # ─────────── Tab 1: 解决方案文档 ───────────
    with tab_sol:
        # 文档类型切换
        doc_type = st.radio(
            "选择文档类型",
            ["AI 解决方案", "Infra 基础设施"],
            horizontal=True,
            key="doc_type_radio",
            index=0 if st.session_state.get("doc_type", "AI") == "AI" else 1,
        )
        current_doc_type = "AI" if doc_type == "AI 解决方案" else "Infra"
        st.session_state["doc_type"] = current_doc_type

        # 文档来源切换
        doc_source = st.radio(
            "文档来源",
            ["AI 生成", "手动导入"],
            horizontal=True,
            key="doc_source_radio",
        )

        left, right = st.columns([1, 1])
        with left:
            if doc_source == "手动导入":
                # ── 手动导入：两步流程 ──
                # Step 1：上传 / 粘贴，确认后暂存原文
                # Step 2：AI 按模板格式重新生成
                if "imported_doc_text" not in st.session_state:
                    # ── Step 1：上传或粘贴 ──
                    uploaded_doc = st.file_uploader(
                        "上传已有的 .docx 文档",
                        type=["docx"],
                        key="upload_existing_doc",
                        help="上传后将自动提取文档文本内容",
                    )
                    manual_text = st.text_area(
                        "或直接粘贴文本内容",
                        height=200,
                        key="manual_doc_text",
                        placeholder="将已有的解决方案文档内容粘贴到此处...",
                    )

                    if st.button("确认导入", type="primary", use_container_width=True, key="btn_import"):
                        imported_text = ""
                        if uploaded_doc is not None:
                            doc = Document(uploaded_doc)
                            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
                            for table in doc.tables:
                                for row in table.rows:
                                    cells = [cell.text.strip() for cell in row.cells]
                                    paragraphs.append(" | ".join(cells))
                            imported_text = "\n\n".join(paragraphs)
                        elif manual_text.strip():
                            imported_text = manual_text.strip()
                        else:
                            st.warning("请上传文档或粘贴文本。")
                            st.stop()

                        st.session_state["imported_doc_text"] = imported_text
                        # 同步写入 solution_text / infra_text，使 POV 等后续步骤可立即识别到文档
                        target_key = "solution_text" if current_doc_type == "AI" else "infra_text"
                        st.session_state[target_key] = imported_text
                        st.session_state["customer_name"] = customer_name.strip() if customer_name.strip() else "未命名客户"
                        st.session_state["account_name"] = account_name.strip() if account_name.strip() else (customer_name.strip() or "未命名客户")
                        st.session_state["budget"] = budget
                        st.session_state.pop("pov_text", None)
                        st.rerun()

                else:
                    # ── Step 2：确认内容 + AI 重新生成 ──
                    imported_text = st.session_state["imported_doc_text"]
                    st.success(f"文档已导入（共 {len(imported_text)} 字符）")
                    st.text_area(
                        "导入内容预览",
                        value=imported_text[:600] + "\n\n..." if len(imported_text) > 600 else imported_text,
                        height=160,
                        disabled=True,
                        key="preview_imported",
                    )

                    c_reimport, c_regen = st.columns(2)
                    with c_reimport:
                        if st.button("重新上传", use_container_width=True, key="btn_reimport"):
                            st.session_state.pop("imported_doc_text", None)
                            st.rerun()
                    with c_regen:
                        if st.button("AI 重新生成", type="primary", use_container_width=True, key="btn_regen_import"):
                            cust = customer_name.strip() or st.session_state.get("customer_name", "未命名客户")
                            system_prompt = SOLUTION_SYSTEM_PROMPT if current_doc_type == "AI" else INFRA_SYSTEM_PROMPT
                            ref_text = solution_ref if current_doc_type == "AI" else infra_ref
                            user_ctx = (
                                f"## 客户信息\n- **客户名称**：{cust}\n\n"
                                f"## 已有解决方案文档（请基于以下内容，按照要求的章节格式重新整理生成，不要照抄原文）\n\n"
                                f"{imported_text}"
                            )
                            if ref_text:
                                user_ctx += (
                                    f"\n\n---\n\n## 【参考模板文档 —— 请学习其风格和结构，不要照抄具体数据】\n\n"
                                    f"{ref_text}"
                                )
                            try:
                                with st.spinner("正在基于导入内容 AI 重新生成..."):
                                    result_text = call_azure_openai(system_prompt, user_ctx)
                                    target_key = "solution_text" if current_doc_type == "AI" else "infra_text"
                                    st.session_state[target_key] = result_text
                                    st.session_state["customer_name"] = cust
                                    st.session_state["account_name"] = account_name.strip() if account_name.strip() else cust
                                    st.session_state["budget"] = budget
                                    st.session_state.pop("pov_text", None)
                                    st.session_state.pop("imported_doc_text", None)
                                st.rerun()
                            except Exception as e:
                                st.error(f"生成失败：{e}")

                    # 若已生成，显示下载按钮
                    target_key = "solution_text" if current_doc_type == "AI" else "infra_text"
                    if target_key in st.session_state:
                        customer = st.session_state["customer_name"]
                        acct = st.session_state.get("account_name") or account_name.strip() or customer
                        if current_doc_type == "AI":
                            docx_bytes = create_solution_docx(
                                content=st.session_state["solution_text"], customer_name=customer
                            )
                            st.download_button(
                                label="下载 AI 解决方案架构文档 (.docx)",
                                data=docx_bytes,
                                file_name=f"{acct}-Solution Architecture.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                                key="dl_sol_import",
                            )
                        else:
                            docx_bytes = create_infra_docx(
                                content=st.session_state["infra_text"], customer_name=customer
                            )
                            st.download_button(
                                label="下载 Infra 基础设施架构文档 (.docx)",
                                data=docx_bytes,
                                file_name=f"{acct}-Infra Solution Architecture.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                                key="dl_infra_import",
                            )

            else:
                # ── AI 生成文档 ──
                if current_doc_type == "AI":
                    # AI 解决方案文档逻辑
                    has_solution = "solution_text" in st.session_state
                    sol_label = "重新生成" if has_solution else "生成 AI 解决方案架构文档"
                    if st.button(sol_label, type="primary", use_container_width=True, key="btn_sol"):
                        if not customer_name.strip():
                            st.warning("请输入客户名称。")
                            st.stop()
                        if not customer_bg.strip():
                            st.warning("请输入客户背景信息。")
                            st.stop()
                        try:
                            with st.spinner("正在生成 AI 解决方案架构文档..."):
                                user_ctx = (
                                    f"## 客户信息\n- **客户名称**：{customer_name}\n\n"
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
                                st.session_state["account_name"] = account_name.strip() if account_name.strip() else customer_name
                                st.session_state["budget"] = budget
                                st.session_state.pop("pov_text", None)
                                st.session_state.pop("svg_code", None)
                            st.rerun()
                        except Exception as e:
                            st.error(f"生成失败：{e}")

                    if "solution_text" in st.session_state:
                        customer = st.session_state["customer_name"]
                        acct = st.session_state.get("account_name") or account_name.strip() or customer
                        docx_sol = create_solution_docx(
                            content=st.session_state["solution_text"], customer_name=customer
                        )
                        st.download_button(
                            label="下载 AI 解决方案架构文档 (.docx)",
                            data=docx_sol,
                            file_name=f"{acct}-Solution Architecture.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )
                else:
                    # Infra 基础设施文档逻辑
                    has_infra = "infra_text" in st.session_state
                    infra_label = "重新生成" if has_infra else "生成 Infra 基础设施架构文档"
                    if st.button(infra_label, type="primary", use_container_width=True, key="btn_infra"):
                        if not customer_name.strip():
                            st.warning("请输入客户名称。")
                            st.stop()
                        if not customer_bg.strip():
                            st.warning("请输入客户背景信息。")
                            st.stop()
                        try:
                            with st.spinner("正在生成 Infra 基础设施架构文档..."):
                                user_ctx = (
                                    f"## 客户信息\n- **客户名称**：{customer_name}\n\n"
                                    f"## 客户背景\n{customer_bg}"
                                )
                                if infra_ref:
                                    user_ctx += (
                                        f"\n\n---\n\n## 【参考模板文档 —— 请学习其风格和结构，不要照抄具体数据】\n\n"
                                        f"{infra_ref}"
                                    )
                                infra_text = call_azure_openai(INFRA_SYSTEM_PROMPT, user_ctx)
                                st.session_state["infra_text"] = infra_text
                                st.session_state["customer_name"] = customer_name
                                st.session_state["account_name"] = account_name.strip() if account_name.strip() else customer_name
                                st.session_state["budget"] = budget
                                st.session_state.pop("pov_text", None)
                                st.session_state.pop("svg_code", None)
                            st.rerun()
                        except Exception as e:
                            st.error(f"生成失败：{e}")

                    if "infra_text" in st.session_state:
                        customer = st.session_state["customer_name"]
                        acct = st.session_state.get("account_name") or account_name.strip() or customer
                        docx_infra = create_infra_docx(
                            content=st.session_state["infra_text"], customer_name=customer
                        )
                        st.download_button(
                            label="下载 Infra 基础设施架构文档 (.docx)",
                            data=docx_infra,
                            file_name=f"{acct}-Infra Solution Architecture.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )

        with right:
            if current_doc_type == "AI":
                if "solution_text" in st.session_state:
                    st.markdown("**AI 解决方案文档预览**")
                    st.markdown(st.session_state["solution_text"], unsafe_allow_html=True)
                else:
                    st.info("请先生成或导入 AI 解决方案文档")
            else:
                if "infra_text" in st.session_state:
                    st.markdown("**Infra 基础设施文档预览**")
                    st.markdown(st.session_state["infra_text"], unsafe_allow_html=True)
                else:
                    st.info("请先生成或导入 Infra 基础设施文档")

    # ─────────── Tab 2: POV 部署计划 ───────────
    with tab_pov:
        # 根据当前文档类型确定使用哪个解决方案文档
        current_doc_type = st.session_state.get("doc_type", "AI")
        has_base_doc = ("solution_text" in st.session_state) if current_doc_type == "AI" else ("infra_text" in st.session_state)
        
        if not has_base_doc:
            doc_type_name = "AI 解决方案" if current_doc_type == "AI" else "Infra 基础设施"
            st.info(f"请先在「解决方案文档」标签页中生成或导入 {doc_type_name} 文档")
        else:
            customer = st.session_state["customer_name"]
            solution = st.session_state["solution_text"] if current_doc_type == "AI" else st.session_state["infra_text"]
            left, right = st.columns([1, 1])
            with left:
                st.caption(f"📄 当前基于: **{current_doc_type}** 解决方案文档")
                dc1, dc2 = st.columns(2)
                with dc1:
                    pov_start = st.date_input("POV 开始日期", value=None)
                with dc2:
                    pov_end = st.date_input(
                        "POV 结束日期",
                        value=None,
                    )

                vendor_team = st.text_area(
                    "乙方项目人员（我方团队）",
                    value=(
                        "技术负责人: \n"
                        "Azure架构师: \n"
                    ),
                    height=120,
                    help="只需填写乙方（我方）人员，甲方人员由 AI 根据客户背景自动生成",
                )

                has_pov = "pov_text" in st.session_state
                pov_label = "重新生成" if has_pov else "生成 POV 部署计划"
                if st.button(pov_label, type="primary", use_container_width=True, key="btn_pov"):
                    if not pov_start or not pov_end:
                        st.warning("请先选择 POV 开始日期和结束日期。")
                        st.stop()
                    try:
                        pov_period = f"{pov_start.strftime('%Y/%m/%d')} - {pov_end.strftime('%Y/%m/%d')}"

                        # ── 预计算工作日与周末 ──
                        all_days = []
                        workdays = []
                        weekends = []
                        d = pov_start
                        while d <= pov_end:
                            if d.weekday() < 5:  # 0=Mon .. 4=Fri
                                workdays.append(f"{d.month}月{d.day}日")
                            else:
                                weekends.append(f"{d.month}月{d.day}日")
                            d += datetime.timedelta(days=1)

                        workday_list_str = "、".join(workdays)
                        weekend_list_str = "、".join(weekends) if weekends else "无"

                        pov_prompt = (
                            f"以下是已生成的解决方案架构文档，请据此生成 POV 部署计划：\n\n"
                            f"{solution}\n\n"
                            f"## 补充信息\n- **客户名称**：{customer}\n"
                            f"- **POV 周期**：{pov_period}\n\n"
                            f"## 可用工作日清单（共 {len(workdays)} 天，必须且只能使用这些日期）\n"
                            f"{workday_list_str}\n\n"
                            f"## 禁用日期（周末，严禁安排任何任务）\n"
                            f"{weekend_list_str}\n\n"
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
                    acct = st.session_state.get("account_name") or account_name.strip() or customer
                    docx_pov = create_pov_docx(
                        content=st.session_state["pov_text"], customer_name=customer
                    )
                    st.download_button(
                        label="下载 POV 部署计划 (.docx)",
                        data=docx_pov,
                        file_name=f"{acct}-PostAssessment POVdeployment.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

            with right:
                if "pov_text" in st.session_state:
                    st.markdown("**文档预览**")
                    st.markdown(st.session_state["pov_text"], unsafe_allow_html=True)
                else:
                    st.info("请填写信息后点击生成")

    # ─────────── Tab 3: Azure Migrate CSV ───────────
    with tab_csv:
        current_doc_type = st.session_state.get("doc_type", "AI")
        has_base_doc = ("solution_text" in st.session_state) if current_doc_type == "AI" else ("infra_text" in st.session_state)
        
        if not has_base_doc:
            doc_type_name = "AI 解决方案" if current_doc_type == "AI" else "Infra 基础设施"
            st.info(f"请先在「解决方案文档」标签页中生成或导入 {doc_type_name} 文档")
        else:
            customer = st.session_state["customer_name"]
            bdgt = st.session_state.get("budget", budget)
            left, right = st.columns([1, 1])
            with left:
                st.caption(f"📄 当前基于: **{current_doc_type}** 解决方案文档")

        current_doc_type = st.session_state.get("doc_type", "AI")
        has_base_doc = ("solution_text" in st.session_state) if current_doc_type == "AI" else ("infra_text" in st.session_state)
        
        if not has_base_doc:
            doc_type_name = "AI 解决方案" if current_doc_type == "AI" else "Infra 基础设施"
            st.info(f"请先在「解决方案文档」标签页中生成或导入 {doc_type_name} 文档")
        else:
            customer = st.session_state["customer_name"]
            bdgt = st.session_state.get("budget", budget)
            left, right = st.columns([1, 1])
            with left:
                st.caption(f"📄 当前基于: **{current_doc_type}** 解决方案文档")
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
                    acct = st.session_state.get("account_name") or account_name.strip() or customer
                    csv_data = st.session_state["csv_code"]
                    st.download_button(
                        label="下载 Azure Migrate CSV",
                        data=csv_data.encode("utf-8-sig"),
                        file_name=f"{acct}-Azure migrate report.csv",
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


    # ─────────── Tab 4: 年度价格表 ───────────
    with tab_yearly:
        st.markdown(
            "上传从 Azure 定价计算器导出的原始 Excel，自动新增 **Estimated yearly cost** 列（月费用 × 12）并在 Total 行汇总。"
        )

        st.divider()

        uploaded_price = st.file_uploader(
            "上传原始价格表 (.xlsx)",
            type=["xlsx"],
            key="upload_price_excel",
            help="支持标准 Azure 定价计算器导出格式",
        )

        if uploaded_price is not None:
            if st.button("生成年度价格表", type="primary", use_container_width=True, key="btn_gen_yearly"):
                import openpyxl
                from copy import copy as _copy
                from openpyxl.styles import Font as _Font

                def _col_letter(n):
                    result = ""
                    while n:
                        n, rem = divmod(n - 1, 26)
                        result = chr(65 + rem) + result
                    return result

                def _copy_cell_style(src, dst):
                    if src.has_style:
                        dst.font      = _copy(src.font)
                        dst.fill      = _copy(src.fill)
                        dst.border    = _copy(src.border)
                        dst.alignment = _copy(src.alignment)
                        dst.number_format = src.number_format

                def _find_header_row(ws):
                    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
                        if row and "Estimated monthly cost" in row:
                            return i
                    return None

                def _find_total_row(ws, hrow):
                    for i, row in enumerate(ws.iter_rows(min_row=hrow + 1, values_only=True), hrow + 1):
                        if row and "Total" in row:
                            return i
                    return None

                def _get_account_name(ws):
                    """从 Sheet 第 2 行前 5 列取账号名（非空的第一个值）。"""
                    for col in range(1, 6):
                        v = ws.cell(2, col).value
                        if v and str(v).strip():
                            return str(v).strip().rstrip("\t").strip()
                    return None

                def _process_sheet(ws):
                    hrow = _find_header_row(ws)
                    if hrow is None:
                        return False, "未找到标题行（含 'Estimated monthly cost'）", None
                    trow = _find_total_row(ws, hrow)
                    if trow is None:
                        return False, "未找到 Total 行", None

                    header_vals = [ws.cell(hrow, c).value for c in range(1, ws.max_column + 1)]
                    try:
                        monthly_col = header_vals.index("Estimated monthly cost") + 1
                        upfront_col = header_vals.index("Estimated upfront cost") + 1
                    except ValueError:
                        return False, "未找到必要列名", None

                    yearly_col     = upfront_col + 1
                    ws.insert_cols(yearly_col)
                    monthly_letter = _col_letter(monthly_col)
                    yearly_letter  = _col_letter(yearly_col)

                    # 标题行：复制 upfront 列样式
                    hcell = ws.cell(hrow, yearly_col, "Estimated yearly cost")
                    _copy_cell_style(ws.cell(hrow, upfront_col), hcell)
                    src_hdr = ws.cell(hrow, upfront_col)
                    hcell.font = _Font(
                        name=src_hdr.font.name or "Calibri",
                        bold=True,
                        size=src_hdr.font.size or 11,
                    )

                    data_start = hrow + 1
                    data_end   = trow - 1

                    # 数据行：写公式，复制样式并特别保留 number_format（用于显示 $）
                    for r in range(data_start, data_end + 1):
                        mv = ws.cell(r, monthly_col).value
                        if mv is not None and (isinstance(mv, (int, float)) or (isinstance(mv, str) and mv.startswith("="))):
                            cell = ws.cell(r, yearly_col)
                            cell.value = f"={monthly_letter}{r}*12"
                            src_cell = ws.cell(r, monthly_col)
                            _copy_cell_style(src_cell, cell)
                            # 显式保留原始单元格的 number_format，以带上 $ 符号
                            if src_cell.number_format and src_cell.number_format != 'General':
                                cell.number_format = src_cell.number_format
                            else:
                                cell.number_format = '"$"#,##0.00'
                        else:
                            ws.cell(r, yearly_col).value = None

                    # Total 行
                    tcell = ws.cell(trow, yearly_col)
                    tcell.value = f"=SUM({yearly_letter}{data_start}:{yearly_letter}{data_end})"
                    src_total = ws.cell(trow, monthly_col)
                    _copy_cell_style(src_total, tcell)
                    if src_total.number_format and src_total.number_format != 'General':
                        tcell.number_format = src_total.number_format
                    else:
                        tcell.number_format = '"$"#,##0.00'
                    tcell.font = _Font(bold=True, name="Calibri", size=11)

                    ws.column_dimensions[yearly_letter].width = 22
                    
                    account = _get_account_name(ws)
                    return True, "处理成功", account

                try:
                    with st.spinner("正在处理 Excel..."):
                        wb = openpyxl.load_workbook(uploaded_price)
                        messages = []
                        account_name = None
                        for sname in wb.sheetnames:
                            ok, msg, acct = _process_sheet(wb[sname])
                            messages.append(f"**{sname}**: {msg}")
                            if acct and not account_name:
                                account_name = acct

                        # 优先使用用户输入的账户名，其次使用 Excel 中提取的名称
                        _budget = st.session_state.get("budget", budget) or "未填写"
                        _acct_from_input = st.session_state.get("account_name") or account_name.strip()
                        _acct_final = _acct_from_input or account_name or uploaded_price.name.replace(".xlsx", "")
                        new_dl_name = f"{_acct_final}-Azure calculator.xlsx"

                        out_buf = io.BytesIO()
                        wb.save(out_buf)
                        out_buf.seek(0)
                        st.session_state["yearly_excel_bytes"] = out_buf.getvalue()
                        st.session_state["yearly_excel_name"]  = new_dl_name
                        st.session_state["yearly_messages"]    = messages

                    st.rerun()
                except Exception as e:
                    st.error(f"处理失败：{e}")
        else:
            st.info("请先上传 Excel 文件")

        # 处理结果与下载
        if "yearly_excel_bytes" in st.session_state:
            st.divider()
            for msg in st.session_state.get("yearly_messages", []):
                st.markdown(msg)
            st.download_button(
                label="下载任务年度价格表 (.xlsx)",
                data=st.session_state["yearly_excel_bytes"],
                file_name=st.session_state["yearly_excel_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_yearly",
            )


# ──────────────────────────────────────────────
# 入口
# ──────────────────────────────────────────────
if __name__ == "__main__":
    main()
