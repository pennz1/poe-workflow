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
import time
import zipfile
from typing import Any, Callable, Dict, List, Optional
import streamlit as st
import requests
from openai import AzureOpenAI
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from frontend.ui import (
    load_desktop_theme,
    render_app_header,
    render_pill,
    render_readiness,
    render_auto_poe_result,
    render_device_code_login,
    render_section_head,
    render_template_status,
    render_workflow_steps,
)

try:
    import msal
except ImportError:
    msal = None

# ──────────────────────────────────────────────
# 常量
# ──────────────────────────────────────────────
APP_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(APP_DIR, "templates")
SOLUTION_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "solution_template.docx.docx")
INFRA_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "Infra_template.docx")
POV_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "pov_template.docx.docx")
MIGRATE_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "AzureMigrateimporttemplate.csv")

MSAL_CLIENT_ID_DEFAULT = "d999b252-3977-4eae-bdae-a0f58c5336b7"
AZURE_AUTHORITY = "https://login.microsoftonline.com/common"
AZURE_ARM_SCOPE = ["https://management.azure.com/.default"]
AZURE_MANAGEMENT_ENDPOINT = "https://management.azure.com"
AZURE_RESOURCE_API_VERSION = "2022-12-01"
AZURE_PROVIDER_API_VERSION = "2021-04-01"
AZURE_MIGRATE_API_VERSION = "2024-01-15"
AZURE_MIGRATE_REPORT_API_VERSION = "2019-10-01"
AZURE_OFFAZURE_API_VERSION = "2023-06-06"
AZURE_MIGRATE_PROJECTS_API_VERSION = "2018-09-01-preview"
AZURE_MIGRATE_DEFAULT_TARGET_LOCATION = "WestUs2"

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
load_desktop_theme(APP_DIR)


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
    "- `# 深圳跃瓦创新科技 - Azure AI 中台与多场景助手 POV 部署计划`\n"
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
# 全自动 POE：MSAL 登录与 Azure ARM API
# ──────────────────────────────────────────────
def _get_msal_client_id() -> str:
    """返回 MSAL Public Client ID。Client ID 不是密钥，可使用默认值。"""
    return st.secrets.get("MSAL_CLIENT_ID", MSAL_CLIENT_ID_DEFAULT)


def _is_azure_token_valid() -> bool:
    expires_at = st.session_state.get("azure_token_expires_at", 0)
    return bool(st.session_state.get("azure_token")) and time.time() < expires_at


def msal_device_code_login() -> None:
    """通过 MSAL Device Code Flow 登录 Azure，并将 access token 放入 session state。"""
    if msal is None:
        raise RuntimeError("缺少 msal 依赖，请确认 requirements.txt 已包含 msal。")

    app = msal.PublicClientApplication(
        _get_msal_client_id(),
        authority=AZURE_AUTHORITY,
    )
    flow = app.initiate_device_flow(scopes=AZURE_ARM_SCOPE)
    if "user_code" not in flow:
        raise RuntimeError(f"无法启动 Microsoft 登录流程：{flow}")

    user_code = flow["user_code"]
    verify_url = flow.get("verification_uri", "https://microsoft.com/devicelogin")
    render_device_code_login(user_code, verify_url)

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or "Microsoft 登录失败。"
        if "7000218" in str(err):
            err += "\n\n请在 Azure Portal → 应用注册 → 身份验证 → 高级设置中，将「允许公共客户端流」设为「是」。"
        raise RuntimeError(err)

    account = result.get("id_token_claims", {}) or result.get("account", {}) or {}
    username = account.get("preferred_username") or account.get("email") or account.get("username") or "Azure 用户"
    expires_in = int(result.get("expires_in", 3600))
    st.session_state["azure_token"] = result["access_token"]
    st.session_state["azure_user"] = username
    st.session_state["azure_token_expires_at"] = time.time() + max(expires_in - 300, 300)


def clear_azure_login() -> None:
    for key in [
        "azure_token",
        "azure_user",
        "azure_token_expires_at",
        "azure_subscription_id",
        "azure_subscription_name",
        "azure_resource_group",
    ]:
        st.session_state.pop(key, None)


def _format_arm_error(response: requests.Response) -> str:
    try:
        payload = response.json()
        error = payload.get("error", payload)
        message = error.get("message") or str(error)
    except Exception:
        message = response.text
    return f"Azure API {response.status_code}: {message[:1200]}"


def _retry_after_seconds(response: requests.Response, default: int = 10) -> int:
    retry_after = response.headers.get("Retry-After")
    if not retry_after:
        return default
    try:
        return max(1, min(int(retry_after), 60))
    except ValueError:
        return default


def _poll_azure_lro(
    operation_url: str,
    token: str,
    initial_delay: int = 10,
    timeout_seconds: int = 900,
) -> Dict[str, Any]:
    """轮询 ARM long-running operation，直到 Succeeded/Failed/Canceled。"""
    if operation_url.startswith("/"):
        operation_url = f"{AZURE_MANAGEMENT_ENDPOINT}{operation_url}"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    deadline = time.time() + timeout_seconds
    delay = max(1, min(initial_delay, 60))

    while time.time() < deadline:
        time.sleep(delay)
        response = requests.get(operation_url, headers=headers, timeout=90)
        if response.status_code >= 400:
            raise RuntimeError(_format_arm_error(response))

        payload: Dict[str, Any] = {}
        if response.content:
            try:
                payload = response.json()
            except ValueError:
                payload = {}

        status = str(
            payload.get("status")
            or payload.get("properties", {}).get("provisioningState")
            or ""
        ).strip()
        status_lower = status.lower()
        if status_lower in {"succeeded", "completed"}:
            return payload
        if status_lower in {"failed", "canceled", "cancelled"}:
            raise RuntimeError(f"Azure 长操作失败：{payload or status}")
        if response.status_code in {200, 204} and not status:
            return payload

        delay = _retry_after_seconds(response, default=10)

    raise TimeoutError("等待 Azure 长操作完成超时。")


def azure_arm_request(
    method: str,
    path_or_url: str,
    token: str,
    body: Optional[Dict[str, Any]] = None,
    timeout: int = 90,
    poll_lro: bool = True,
    lro_timeout: int = 900,
) -> Dict[str, Any]:
    """调用 Azure ARM REST API。path_or_url 可传完整 URL 或 ARM 相对路径。"""
    url = path_or_url if path_or_url.startswith("http") else f"{AZURE_MANAGEMENT_ENDPOINT}{path_or_url}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    response = requests.request(method, url, headers=headers, json=body, timeout=timeout)
    if response.status_code >= 400:
        raise RuntimeError(_format_arm_error(response))
    payload: Dict[str, Any] = {}
    if response.content:
        try:
            payload = response.json()
        except ValueError:
            payload = {}

    lro_url = response.headers.get("Azure-AsyncOperation") or response.headers.get("Location")
    if poll_lro and response.status_code in {201, 202} and lro_url:
        final_payload = _poll_azure_lro(
            lro_url,
            token,
            initial_delay=_retry_after_seconds(response, default=10),
            timeout_seconds=lro_timeout,
        )
        return payload or final_payload

    return payload


def azure_arm_list(path: str, token: str) -> List[Dict[str, Any]]:
    """读取 ARM 列表接口并处理 nextLink 分页。"""
    items: List[Dict[str, Any]] = []
    next_url = f"{AZURE_MANAGEMENT_ENDPOINT}{path}"
    while next_url:
        payload = azure_arm_request("GET", next_url, token)
        items.extend(payload.get("value", []))
        next_url = payload.get("nextLink")
    return items


def list_azure_subscriptions(token: str) -> List[Dict[str, Any]]:
    subscriptions = azure_arm_list(f"/subscriptions?api-version={AZURE_RESOURCE_API_VERSION}", token)
    return sorted(subscriptions, key=lambda item: item.get("displayName", ""))


def list_azure_resource_groups(subscription_id: str, token: str) -> List[Dict[str, Any]]:
    groups = azure_arm_list(
        f"/subscriptions/{subscription_id}/resourceGroups?api-version={AZURE_RESOURCE_API_VERSION}",
        token,
    )
    return sorted(groups, key=lambda item: item.get("name", ""))


def _subscription_label(subscription: Dict[str, Any]) -> str:
    display_name = subscription.get("displayName") or subscription.get("subscriptionId")
    state = subscription.get("state", "Unknown")
    return f"{display_name} ({state})"


def _resource_group_label(resource_group: Dict[str, Any]) -> str:
    name = resource_group.get("name", "")
    location = resource_group.get("location", "")
    return f"{name} ({location})" if location else name


# ──────────────────────────────────────────────
# 全自动 POE：文档、Migrate、打包
# ──────────────────────────────────────────────
def _safe_azure_name(value: str, fallback: str, suffix: str = "", max_len: int = 60) -> str:
    base = re.sub(r"[^A-Za-z0-9-]+", "-", value or "").strip("-")
    base = re.sub(r"-+", "-", base)
    if not base:
        base = fallback
    reserve = len(suffix)
    return f"{base[: max_len - reserve].strip('-')}{suffix}".strip("-")


AZURE_MIGRATE_PROJECT_LOCATIONS = {
    "centralus", "westeurope", "uksouth", "ukwest", "northeurope", "westus2",
    "southeastasia", "eastasia", "centralindia", "southindia", "canadacentral",
    "australiasoutheast", "japanwest", "japaneast", "brazilsouth", "koreacentral",
    "koreasouth", "francecentral", "switzerlandnorth", "australiaeast", "uaenorth",
    "southafricanorth", "germanywestcentral", "norwayeast", "jioindiawest",
    "swedencentral", "qatarcentral", "polandcentral", "italynorth", "israelcentral",
    "spaincentral", "mexicocentral", "newzealandnorth", "indonesiacentral",
    "malaysiawest", "chilecentral", "austriaeast", "belgiumcentral", "denmarkeast",
}

AZURE_MIGRATE_LOCATION_ALIASES = {
    "westus": "westus2",
    "west us": "westus2",
    "east us 2": "eastus2",
    "eastus 2": "eastus2",
}


def _normalize_azure_location(location: str) -> str:
    return re.sub(r"\s+", "", (location or "").strip().lower())


def resolve_migrate_project_location(subscription_id: str, resource_group: str, token: str) -> str:
    resource_group_payload = azure_arm_request(
        "GET",
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}?api-version={AZURE_RESOURCE_API_VERSION}",
        token,
    )
    normalized = _normalize_azure_location(resource_group_payload.get("location", ""))
    normalized = AZURE_MIGRATE_LOCATION_ALIASES.get(normalized, normalized)
    if normalized in AZURE_MIGRATE_PROJECT_LOCATIONS:
        return normalized
    return "westus2"


def _workday_info(start_date: datetime.date, end_date: datetime.date) -> tuple[list[str], list[str]]:
    workdays = []
    weekends = []
    current = start_date
    while current <= end_date:
        target = f"{current.month}月{current.day}日"
        if current.weekday() < 5:
            workdays.append(target)
        else:
            weekends.append(target)
        current += datetime.timedelta(days=1)
    return workdays, weekends


def has_meaningful_pov_team(vendor_team: Optional[str]) -> bool:
    text = str(vendor_team or "").strip()
    if not text:
        return False
    placeholders = {"技术负责人", "Azure架构师", "Azure 架构师", "项目经理", "负责人"}
    for raw_line in text.splitlines():
        line = raw_line.strip().strip("：:")
        if not line:
            continue
        if ":" in raw_line or "：" in raw_line:
            _, value = re.split(r"[:：]", raw_line, maxsplit=1)
            if value.strip():
                return True
            if line in placeholders:
                continue
        elif line not in placeholders:
            return True
    return False


def build_pov_prompt(
    solution_text: str,
    customer_name: str,
    pov_start: datetime.date,
    pov_end: datetime.date,
    vendor_team: str,
    pov_ref: str,
) -> str:
    workdays, weekends = _workday_info(pov_start, pov_end)
    pov_prompt = (
        f"以下是已生成的解决方案架构文档，请据此生成 POV 部署计划：\n\n"
        f"{solution_text}\n\n"
        f"## 补充信息\n- **客户名称**：{customer_name}\n"
        f"- **POV 周期**：{pov_start.strftime('%Y/%m/%d')} - {pov_end.strftime('%Y/%m/%d')}\n\n"
        f"## 可用工作日清单（共 {len(workdays)} 天，必须且只能使用这些日期）\n"
        f"{'、'.join(workdays)}\n\n"
        f"## 禁用日期（周末，严禁安排任何任务）\n"
        f"{'、'.join(weekends) if weekends else '无'}\n\n"
        f"## 乙方项目人员\n{vendor_team.strip()}\n\n"
        f"请根据客户背景信息自动生成合理的甲方人员（2-3人，包含项目负责人和技术对接人，一定要中文名！）。"
    )
    if pov_ref:
        pov_prompt += (
            "\n\n---\n\n## 【参考模板文档 —— 请学习其风格和结构，不要照抄具体数据】\n\n"
            f"{pov_ref}"
        )
    return pov_prompt


def generate_solution_artifact(
    current_doc_type: str,
    customer_name: str,
    account_name: str,
    customer_bg: str,
    solution_ref: str,
    infra_ref: str,
) -> Dict[str, Any]:
    system_prompt = SOLUTION_SYSTEM_PROMPT if current_doc_type == "AI" else INFRA_SYSTEM_PROMPT
    ref_text = solution_ref if current_doc_type == "AI" else infra_ref
    user_ctx = (
        f"## 客户信息\n- **客户名称**：{customer_name}\n\n"
        f"## 客户背景\n{customer_bg}"
    )
    if ref_text:
        user_ctx += (
            "\n\n---\n\n## 【参考模板文档 —— 请学习其风格和结构，不要照抄具体数据】\n\n"
            f"{ref_text}"
        )

    content = call_azure_openai(system_prompt, user_ctx)
    if current_doc_type == "AI":
        docx_bytes = create_solution_docx(content=content, customer_name=customer_name)
        file_name = f"{account_name}-Solution Architecture.docx"
    else:
        docx_bytes = create_infra_docx(content=content, customer_name=customer_name)
        file_name = f"{account_name}-Infra Solution Architecture.docx"

    return {"content": content, "bytes": docx_bytes, "file_name": file_name}


def generate_pov_artifact(
    solution_text: str,
    customer_name: str,
    account_name: str,
    pov_ref: str,
    pov_start: datetime.date,
    pov_end: datetime.date,
    vendor_team: str,
) -> Dict[str, Any]:
    pov_prompt = build_pov_prompt(solution_text, customer_name, pov_start, pov_end, vendor_team, pov_ref)
    content = call_azure_openai(POV_SYSTEM_PROMPT, pov_prompt)
    docx_bytes = create_pov_docx(content=content, customer_name=customer_name)
    return {
        "content": content,
        "bytes": docx_bytes,
        "file_name": f"{account_name}-PostAssessment POVdeployment.docx",
    }


def register_azure_provider(subscription_id: str, namespace: str, token: str) -> None:
    azure_arm_request(
        "POST",
        f"/subscriptions/{subscription_id}/providers/{namespace}/register?api-version={AZURE_PROVIDER_API_VERSION}",
        token,
    )


def _extract_sas_url(payload: Any) -> Optional[str]:
    if isinstance(payload, str):
        if payload.startswith("http") and "sig=" in payload:
            return payload
        return None
    if isinstance(payload, dict):
        for value in payload.values():
            found = _extract_sas_url(value)
            if found:
                return found
    if isinstance(payload, list):
        for value in payload:
            found = _extract_sas_url(value)
            if found:
                return found
    return None


def _arm_path_with_api_version(path_or_id: str, api_version: str) -> str:
    separator = "&" if "?" in path_or_id else "?"
    return f"{path_or_id}{separator}api-version={api_version}"


def _resource_id(subscription_id: str, resource_group: str, provider_path: str) -> str:
    return f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}/{provider_path}"


def _migrate_project_path(subscription_id: str, resource_group: str, project_name: str) -> str:
    return (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/migrateProjects/{project_name}"
        f"?api-version={AZURE_MIGRATE_PROJECTS_API_VERSION}"
    )


def _migrate_solution_path(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    solution_name: str,
) -> str:
    return (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/migrateProjects/{project_name}"
        f"/solutions/{solution_name}?api-version={AZURE_MIGRATE_PROJECTS_API_VERSION}"
    )


def _migrate_solution_id(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    solution_name: str,
) -> str:
    return _resource_id(
        subscription_id,
        resource_group,
        f"providers/Microsoft.Migrate/migrateProjects/{project_name}/solutions/{solution_name}",
    )


def _servers_solution_details(extended_details: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    details = {
        "dependencyEnabledMachines": "0",
        "machinesHavingSqlServers": "0",
        "machinesHavingWebServers": "0",
        "serversOnLinux": "0",
        "serversOnWindows": "0",
        "serversOnOther": "0",
    }
    if extended_details:
        details.update(extended_details)
    return {
        "assessmentCount": 0,
        "groupCount": 0,
        "extendedDetails": details,
    }


def register_migrate_tool(subscription_id: str, resource_group: str, project_name: str, tool: str, token: str) -> None:
    azure_arm_request(
        "POST",
        (
            f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
            f"/providers/Microsoft.Migrate/migrateProjects/{project_name}"
            f"/registerTool?api-version={AZURE_MIGRATE_PROJECTS_API_VERSION}"
        ),
        token,
        {"tool": tool},
    )


def put_migrate_solution(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    solution_name: str,
    properties: Dict[str, Any],
    token: str,
) -> Dict[str, Any]:
    return azure_arm_request(
        "PUT",
        _migrate_solution_path(subscription_id, resource_group, project_name, solution_name),
        token,
        {"properties": properties},
    )


def ensure_migrate_solution(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    solution_name: str,
    properties: Dict[str, Any],
    token: str,
    progress: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    path = _migrate_solution_path(subscription_id, resource_group, project_name, solution_name)
    existing = _try_get_existing_resource(path, token)
    if existing and existing.get("properties", {}).get("details"):
        if progress:
            progress(f"  ✅ Solution 已存在且 details 完整，复用: {solution_name}")
        return existing
    result = put_migrate_solution(subscription_id, resource_group, project_name, solution_name, properties, token)
    if progress:
        progress(f"  ✅ Solution 已补齐: {solution_name}")
    return result


def ensure_portal_menu_solutions(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    master_site_id: str,
    token: str,
    progress: Callable[[str], None],
) -> None:
    """补齐 Azure Portal 项目菜单 blade 会枚举的默认 solution，避免前端读取 undefined.details。"""
    default_solutions = [
        (
            "Servers-Discovery-ServerDiscovery",
            {
                "tool": "ServerDiscovery",
                "purpose": "Discovery",
                "goal": "Servers",
                "status": "Inactive",
                "details": _servers_solution_details({"masterSiteId": master_site_id}),
            },
        ),
        (
            "Servers-Migration-ServerMigration",
            {
                "tool": "ServerMigration",
                "purpose": "Migration",
                "goal": "Servers",
                "status": "Active",
                "details": _servers_solution_details(),
            },
        ),
        (
            "Servers-Migration-ServerMigration_DataReplication",
            {
                "tool": "ServerMigration_DataReplication",
                "purpose": "Migration",
                "goal": "Servers",
                "status": "Inactive",
                "details": _servers_solution_details(),
            },
        ),
    ]
    for solution_name, properties in default_solutions:
        ensure_migrate_solution(
            subscription_id,
            resource_group,
            project_name,
            solution_name,
            properties,
            token,
            progress,
        )


def refresh_migrate_project_summary(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    token: str,
) -> None:
    azure_arm_request(
        "POST",
        (
            f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
            f"/providers/Microsoft.Migrate/migrateProjects/{project_name}"
            f"/refreshSummary?api-version={AZURE_MIGRATE_PROJECTS_API_VERSION}"
        ),
        token,
        {"goal": "Servers"},
        poll_lro=False,
    )


def _server_summary_count(migrate_project: Dict[str, Any]) -> int:
    servers = migrate_project.get("properties", {}).get("summary", {}).get("servers", {})
    direct_count = int(servers.get("discoveredCount") or 0)
    extended = servers.get("extendedSummary") or {}
    microsoft_count = int(extended.get("microsoftMachinesCount") or 0)
    return max(direct_count, microsoft_count)


def wait_for_portal_inventory_summary(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    expected_machine_count: int,
    token: str,
    progress: Callable[[str], None],
    timeout_seconds: int = 300,
) -> int:
    """等待 migrateProject summary 反映 Import CSV 库存，这是 Portal 全部库存 blade 使用的项目汇总层。"""
    project_path = _migrate_project_path(subscription_id, resource_group, project_name)
    deadline = time.time() + timeout_seconds
    attempt = 0
    last_count = 0

    while time.time() < deadline:
        attempt += 1
        try:
            refresh_migrate_project_summary(subscription_id, resource_group, project_name, token)
        except Exception:
            pass

        project = azure_arm_request("GET", project_path, token, poll_lro=False)
        last_count = _server_summary_count(project)
        if last_count >= expected_machine_count:
            return last_count

        progress(
            f"等待 Azure Portal 全部库存汇总刷新，当前 {last_count}/{expected_machine_count} 台"
            f"（第 {attempt} 次检查）"
        )
        time.sleep(20)

    raise TimeoutError(
        "Azure Portal 全部库存汇总未刷新到预期数量。"
        f"当前 {last_count}/{expected_machine_count} 台；"
        "请检查 ServerDiscovery_Import solution 与 Import Site 关联。"
    )


def _format_import_job_error(job: Dict[str, Any]) -> str:
    props = job.get("properties", {}) if isinstance(job, dict) else {}
    summary = props.get("errorSummary") if isinstance(props, dict) else {}
    parts = []
    if isinstance(summary, dict):
        error_count = summary.get("errorCount")
        warning_count = summary.get("warningCount")
        if error_count is not None:
            parts.append(f"errors={error_count}")
        if warning_count is not None:
            parts.append(f"warnings={warning_count}")
        errors = summary.get("errors")
        if isinstance(errors, list) and errors:
            preview = "; ".join(str(item) for item in errors[:5])
            parts.append(f"details={preview}")
    result = props.get("jobResult") or job.get("status")
    if result:
        parts.insert(0, f"jobResult={result}")
    return "；".join(parts) if parts else str(job)[:800]


def _dedupe_machines_by_discovery_arm_id(machines: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    deduped: List[Dict[str, Any]] = []
    seen = set()
    for machine in machines:
        props = machine.get("properties", {})
        key = str(props.get("discoveryMachineArmId") or machine.get("id") or machine.get("name") or "").lower()
        if not key or key in seen:
            continue
        seen.add(key)
        deduped.append(machine)
    return deduped


def wait_for_import_site_import(
    subscription_id: str,
    resource_group: str,
    site_name: str,
    token: str,
    progress: Callable[[str], None],
    job_arm_id: Optional[str] = None,
    timeout_seconds: int = 900,
) -> List[Dict[str, Any]]:
    """等待 OffAzure import site 完成 CSV 解析，并返回导入到 import site 的机器。"""
    machines_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.OffAzure/importSites/{site_name}/machines"
        f"?api-version={AZURE_OFFAZURE_API_VERSION}"
    )
    jobs_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.OffAzure/importSites/{site_name}/importJobs"
        f"?api-version={AZURE_OFFAZURE_API_VERSION}"
    )
    job_path = (
        _arm_path_with_api_version(job_arm_id, AZURE_OFFAZURE_API_VERSION)
        if job_arm_id and (job_arm_id.startswith("/") or job_arm_id.startswith("http"))
        else None
    )
    deadline = time.time() + timeout_seconds
    attempt = 0
    last_job: Dict[str, Any] = {}

    while time.time() < deadline:
        attempt += 1
        try:
            machines = azure_arm_list(machines_path, token)
            if machines:
                return machines
        except Exception:
            pass

        job: Dict[str, Any] = {}
        if job_path:
            try:
                job = azure_arm_request("GET", job_path, token, poll_lro=False)
            except Exception:
                job = {}
        if not job:
            try:
                jobs = azure_arm_list(jobs_path, token)
                if jobs:
                    job = jobs[-1]
            except Exception:
                job = {}
        if job:
            last_job = job
            props = job.get("properties", {})
            result = str(props.get("jobResult") or job.get("status") or "Unknown").strip()
            imported_count = props.get("numberOfMachinesImported")
            if result in {"Completed", "CompletedWithWarnings"}:
                machines = azure_arm_list(machines_path, token)
                if machines:
                    return machines
            if result in {"Failed", "CompletedWithErrors"}:
                raise RuntimeError(f"Azure Migrate CSV 导入失败：{_format_import_job_error(job)}")
            suffix = f"，已导入 {imported_count} 台" if imported_count is not None else ""
            progress(f"服务器清单导入任务状态：{result}{suffix}（第 {attempt} 次检查）")
        else:
            progress(f"等待 Azure Migrate 创建服务器清单导入任务...（第 {attempt} 次检查）")
        time.sleep(15)

    detail = f"最后一次任务状态：{_format_import_job_error(last_job)}" if last_job else "未查询到导入任务。"
    raise TimeoutError(f"等待 Azure Migrate 导入服务器清单超时（15 分钟）。{detail}")


def wait_for_project_machines(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    collector_name: str,
    token: str,
    progress: Callable[[str], None],
    site_name: Optional[str] = None,
    timeout_seconds: int = 600,
) -> List[Dict[str, Any]]:
    """等待 Azure Migrate 导入完成，通过直接查询 machines 列表来判断。"""
    machines_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"/machines?api-version={AZURE_MIGRATE_API_VERSION}"
    )
    project_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"?api-version={AZURE_MIGRATE_API_VERSION}"
    )
    deadline = time.time() + timeout_seconds
    attempt = 0

    def _filter_current_import(machines: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        if not site_name:
            return _dedupe_machines_by_discovery_arm_id(machines)
        marker = f"/importsites/{site_name}".lower()
        return _dedupe_machines_by_discovery_arm_id([
            machine for machine in machines
            if marker in str(machine.get("properties", {}).get("discoveryMachineArmId", "")).lower()
        ])

    while time.time() < deadline:
        attempt += 1
        # 先尝试直接列出 machines
        try:
            all_machines = azure_arm_list(machines_path, token)
            machines = _filter_current_import(all_machines)
            if machines:
                return machines
        except Exception:
            pass
        # 也检查 project 的 numberOfMachines
        try:
            project = azure_arm_request("GET", project_path, token)
            machine_count = project.get("properties", {}).get("numberOfMachines", 0) or 0
            if machine_count > 0:
                machines = _filter_current_import(azure_arm_list(machines_path, token))
                if machines:
                    return machines
        except Exception:
            pass
        progress(f"服务器清单仍在导入中，继续等待 Azure Migrate 发现结果...（第 {attempt} 次检查）")
        time.sleep(15)
    raise TimeoutError("等待 Azure Migrate 导入服务器清单超时（10 分钟），请在 Azure Portal 检查导入状态。")


def list_imported_machines(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    collector_name: str,
    token: str,
) -> List[Dict[str, Any]]:
    candidate_paths = [
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}/providers/Microsoft.Migrate/assessmentProjects/{project_name}/machines?api-version={AZURE_MIGRATE_API_VERSION}",
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}/providers/Microsoft.Migrate/assessmentProjects/{project_name}/importcollectors/{collector_name}/machines?api-version={AZURE_MIGRATE_API_VERSION}",
    ]
    last_error = None
    for path in candidate_paths:
        try:
            machines = azure_arm_list(path, token)
            if machines:
                return machines
        except Exception as exc:
            last_error = exc
    if last_error:
        raise RuntimeError(f"无法读取 Azure Migrate 已导入服务器列表：{last_error}")
    return []


def wait_for_assessment_complete(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    group_name: str,
    assessment_name: str,
    token: str,
    progress: Callable[[str], None],
    timeout_seconds: int = 600,
) -> Dict[str, Any]:
    path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"/groups/{group_name}/assessments/{assessment_name}"
        f"?api-version={AZURE_MIGRATE_API_VERSION}"
    )
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        assessment = azure_arm_request("GET", path, token)
        props = assessment.get("properties", {})
        status = props.get("status", "Unknown")
        stage = props.get("stage", "Unknown")
        provisioning_state = props.get("provisioningState", "Unknown")
        if status == "Completed":
            return assessment
        if status in {"Invalid", "OutOfSync", "OutDated", "Deleted"}:
            raise RuntimeError(f"Azure Migrate 评估状态异常：{status}")
        if provisioning_state in {"Failed", "Canceled"}:
            raise RuntimeError(f"Azure Migrate 评估资源状态异常：{provisioning_state}")
        progress(f"评估仍在计算中，当前状态：{status}，阶段：{stage}，资源状态：{provisioning_state}")
        time.sleep(15)
    raise TimeoutError("等待 Azure Migrate 评估完成超时，请稍后在 Azure Portal 检查评估结果。")


def parse_annual_budget_usd(raw_budget: Optional[str]) -> Optional[float]:
    """把 UI 里的年预估消耗解析成 USD 数字；无法可靠解析时返回 None。"""
    if raw_budget is None:
        return None
    text = str(raw_budget).strip()
    if not text:
        return None

    normalized = (
        text.lower()
        .replace(",", "")
        .replace("$", "")
        .replace("usd", "")
        .replace("美元", "")
        .replace("美金", "")
        .replace("预估", "")
        .replace("年消耗", "")
        .replace("年度", "")
        .replace("每年", "")
        .replace("+", "")
        .strip()
    )
    if not normalized or normalized in {"na", "n/a", "none", "null", "未填写", "无"}:
        return None

    multiplier = 1.0
    if "万" in normalized or re.search(r"\d\s*w\b", normalized):
        multiplier = 10_000.0
        normalized = normalized.replace("万", "").replace("w", "")
    elif "million" in normalized:
        multiplier = 1_000_000.0
        normalized = normalized.replace("million", "")
    elif re.search(r"\d\s*m\b", normalized):
        multiplier = 1_000_000.0
        normalized = re.sub(r"\bm\b", "", normalized)
    elif re.search(r"\d\s*k\b", normalized):
        multiplier = 1_000.0
        normalized = re.sub(r"\bk\b", "", normalized)

    match = re.search(r"(\d+(?:\.\d+)?)", normalized)
    if not match:
        return None
    return float(match.group(1)) * multiplier


def _assessment_cost_component(assessment: Dict[str, Any], component_name: str) -> float:
    props = assessment.get("properties", {})
    for component in props.get("costComponents") or []:
        if str(component.get("name") or "").lower() == component_name.lower():
            try:
                return float(component.get("value") or 0)
            except (TypeError, ValueError):
                return 0.0
    return 0.0


def assessment_monthly_total_cost(assessment: Dict[str, Any]) -> float:
    props = assessment.get("properties", {})

    def _num(name: str) -> float:
        try:
            return float(props.get(name) or 0)
        except (TypeError, ValueError):
            return 0.0

    return (
        _num("monthlyComputeCost")
        + _num("monthlyStorageCost")
        + _num("monthlyBandwidthCost")
        + _assessment_cost_component(assessment, "MonthlySecurityCost")
    )


def _format_usd(value: Optional[float]) -> str:
    if value is None:
        return "未填写"
    return f"${value:,.2f}"


def _assessment_settings_snapshot(settings: Dict[str, Any]) -> Dict[str, Any]:
    keys = [
        "azureLocation",
        "sizingCriterion",
        "reservedInstance",
        "azureHybridUseBenefit",
        "linuxAzureHybridUseBenefit",
        "azureSecurityOfferingType",
        "scalingFactor",
        "discountPercentage",
    ]
    return {key: settings.get(key) for key in keys}


def _build_assessment_body() -> Dict[str, Any]:
    return {
        "properties": {
            "groupType": "Import",
            "assessmentType": "MachineAssessment",
            "azureLocation": AZURE_MIGRATE_DEFAULT_TARGET_LOCATION,
            "azureOfferCode": "MSAZR0003P",
            "azurePricingTier": "Standard",
            "azureStorageRedundancy": "LocallyRedundant",
            "scalingFactor": 1.3,
            "percentile": "Percentile95",
            "timeRange": "Day",
            "currency": "USD",
            "azureHybridUseBenefit": "Yes",
            "linuxAzureHybridUseBenefit": "Yes",
            "azureSecurityOfferingType": "MDC",
            "discountPercentage": 0,
            "sizingCriterion": "PerformanceBased",
            "azureDiskTypes": ["Premium", "StandardSSD", "Standard"],
            "azureVmFamilies": [
                "Dv2_series", "Dv3_series", "DSv2_series", "Dsv3_series", "Ev3_series",
                "Esv3_series", "F_series", "Fs_series", "Fsv2_series", "M_series", "D_series",
                "DS_series", "H_series", "Lsv2_series",
            ],
            "vmUptime": {"daysPerMonth": 31, "hoursPerDay": 24},
            "reservedInstance": "RI3Year",
            "stage": "InProgress",
        },
    }


def tune_assessment_to_budget(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    group_name: str,
    assessment_name: str,
    assessment_path: str,
    assessment_body: Dict[str, Any],
    assessment: Dict[str, Any],
    annual_budget: Optional[float],
    token: str,
    progress: Callable[[str], None],
) -> tuple[Dict[str, Any], List[Dict[str, Any]], bool]:
    history: List[Dict[str, Any]] = []
    target_min = annual_budget if annual_budget and annual_budget > 0 else None
    target_max = annual_budget * 1.2 if annual_budget and annual_budget > 0 else None
    target_mid = annual_budget * 1.1 if annual_budget and annual_budget > 0 else None

    def _record(round_name: str, action: str, current: Dict[str, Any], met: bool) -> None:
        monthly_total = assessment_monthly_total_cost(current)
        annual_total = monthly_total * 12
        history.append({
            "round": round_name,
            "action": action,
            "monthly_total": monthly_total,
            "annual_total": annual_total,
            "target_annual": annual_budget,
            "target_min": target_min,
            "target_max": target_max,
            "met_target": met,
            "settings": _assessment_settings_snapshot(assessment_body["properties"]),
        })

    def _in_target_range(current: Dict[str, Any]) -> bool:
        if target_min is None or target_max is None:
            return True
        annual_total = assessment_monthly_total_cost(current) * 12
        return target_min <= annual_total <= target_max

    def _next_patch(annual_total: float) -> tuple[str, Dict[str, Any]]:
        settings = assessment_body["properties"]
        if target_max is not None and target_mid is not None and annual_total > target_max:
            current_discount = float(settings.get("discountPercentage") or 0)
            current_discount_factor = max(1 - current_discount / 100, 0.01)
            undiscounted_annual = annual_total / current_discount_factor
            required_discount = 100 * (1 - (target_mid / max(undiscounted_annual, 1)))
            next_discount = min(max(current_discount, required_discount), 99.0)
            if next_discount > current_discount + 0.1:
                return (
                    f"设置折扣为 {next_discount:.2f}%，控制年化估算不超过预估值 20%",
                    {"discountPercentage": round(next_discount, 2)},
                )
            current_factor = float(settings.get("scalingFactor") or 1.3)
            if current_factor > 1.0:
                next_factor = max(1.0, current_factor * 0.8)
                return (
                    f"降低舒适因子到 {next_factor:.2f}，控制年化估算不超过预估值 20%",
                    {"scalingFactor": round(next_factor, 2)},
                )
            if settings.get("azureSecurityOfferingType") != "NO":
                return ("关闭安全成本估算，控制年化估算不超过预估值 20%", {"azureSecurityOfferingType": "NO"})
            return ("已达到自动降价边界", {})

        if target_mid is None:
            return ("未设置预算，不调整", {})

        current_discount = float(settings.get("discountPercentage") or 0)
        if current_discount > 0:
            return ("取消折扣以提高年化估算", {"discountPercentage": 0})
        if settings.get("azureHybridUseBenefit") != "No" or settings.get("linuxAzureHybridUseBenefit") != "No":
            return (
                "关闭 Azure Hybrid Benefit，把 OS 许可成本计入估算",
                {"azureHybridUseBenefit": "No", "linuxAzureHybridUseBenefit": "No"},
            )
        current_factor = float(settings.get("scalingFactor") or 1.3)
        factor_ratio = min(max(target_mid / max(annual_total, 1), 1.05), 1.35)
        next_factor = min(current_factor * factor_ratio, 5.0)
        if next_factor > current_factor + 0.01:
            return (f"提高舒适因子到 {next_factor:.2f}", {"scalingFactor": round(next_factor, 2)})
        if settings.get("reservedInstance") != "None":
            return ("切换为按量计费以提高年化估算", {"reservedInstance": "None"})
        return ("已达到自动提价边界", {})

    if annual_budget is None or annual_budget <= 0:
        monthly_total = assessment_monthly_total_cost(assessment)
        progress(
            "未填写可解析的预估年消耗，跳过价格校准。"
            f"当前 Azure Migrate 年化估算：{_format_usd(monthly_total * 12)}"
        )
        _record("initial", "未填写可解析预算，未调整评估设置", assessment, True)
        return assessment, history, True

    monthly_total = assessment_monthly_total_cost(assessment)
    annual_total = monthly_total * 12
    progress(
        "Azure Migrate 当前年化估算："
        f"{_format_usd(annual_total)}；目标区间：{_format_usd(target_min)} - {_format_usd(target_max)}"
    )
    if _in_target_range(assessment):
        _record("initial", "初始 Portal 默认评估已在目标区间内", assessment, True)
        return assessment, history, True
    direction = "高于" if target_max is not None and annual_total > target_max else "低于"
    _record("initial", f"初始 Portal 默认评估{direction}目标区间", assessment, False)

    for round_index in range(1, 4):
        round_name = f"round-{round_index}"
        action, patch = _next_patch(annual_total)
        if not patch:
            progress(f"评估年化估算仍不在目标区间内，{action}。")
            break
        progress(f"评估年化估算不在目标区间，开始自动调整（{round_name}）：{action}")
        assessment_body["properties"].update(patch)
        assessment_body["properties"]["stage"] = "InProgress"
        azure_arm_request("PUT", assessment_path, token, assessment_body)
        assessment = wait_for_assessment_complete(
            subscription_id, resource_group, project_name, group_name, assessment_name, token, progress
        )
        monthly_total = assessment_monthly_total_cost(assessment)
        annual_total = monthly_total * 12
        met = _in_target_range(assessment)
        progress(
            f"{round_name} 重新计算完成：月估算 {_format_usd(monthly_total)}，"
            f"年化 {_format_usd(annual_total)}，目标区间 {_format_usd(target_min)} - {_format_usd(target_max)}"
        )
        _record(round_name, action, assessment, met)
        if met:
            return assessment, history, True

    progress(
        "已自动调整 3 轮，但 Azure Migrate 年化估算仍未落入用户预估年消耗的 100%-120% 区间；"
        "请到 Azure Portal 的评估设置中手动调整后重新导出。"
    )
    return assessment, history, False


def download_assessment_report(
    subscription_id: str,
    resource_group: str,
    project_name: str,
    group_name: str,
    assessment_name: str,
    token: str,
) -> bytes:
    download_url_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"/groups/{group_name}/assessments/{assessment_name}/downloadUrl"
        f"?api-version={AZURE_MIGRATE_REPORT_API_VERSION}"
    )
    payload = azure_arm_request("POST", download_url_path, token, poll_lro=False)
    report_url = payload.get("assessmentReportUrl")
    if not report_url:
        raise RuntimeError(f"Azure Migrate 未返回评估报告下载地址：{payload}")

    response = requests.get(report_url, timeout=180)
    if response.status_code >= 400:
        raise RuntimeError(f"下载 Azure Migrate 评估报告失败：{response.status_code} {response.text[:800]}")
    if not response.content:
        raise RuntimeError("Azure Migrate 评估报告为空。")
    return response.content


def _try_get_existing_resource(path: str, token: str) -> Optional[Dict[str, Any]]:
    """尝试 GET 某资源，如果存在返回 dict，不存在返回 None。"""
    try:
        result = azure_arm_request("GET", path, token)
        if result and result.get("id"):
            return result
    except Exception:
        pass
    return None


def wait_for_group_machine_membership(
    group_path: str,
    token: str,
    expected_machine_count: int,
    progress: Callable[[str], None],
    timeout_seconds: int = 600,
) -> Dict[str, Any]:
    """等待 updateMachines 完成，并返回最新评估组信息。"""
    deadline = time.time() + timeout_seconds
    attempt = 0
    last_group: Dict[str, Any] = {}

    while time.time() < deadline:
        attempt += 1
        last_group = azure_arm_request("GET", group_path, token, poll_lro=False)
        props = last_group.get("properties", {})
        machine_count = int(props.get("machineCount") or 0)
        provisioning_state = props.get("provisioningState", "Unknown")
        supported_types = sorted(
            str(item).strip()
            for item in (props.get("supportedAssessmentTypes") or [])
            if str(item).strip()
        )
        supported_label = ", ".join(supported_types) if supported_types else "未返回"

        if machine_count >= expected_machine_count and provisioning_state not in {"Failed", "Canceled"}:
            progress(f"  ℹ️ 评估组已包含 {machine_count} 台服务器；支持类型: {supported_label}")
            return last_group
        if provisioning_state in {"Failed", "Canceled"}:
            raise RuntimeError(f"评估组更新失败，资源状态：{provisioning_state}")

        progress(
            f"评估组仍在关联服务器，当前 {machine_count}/{expected_machine_count} 台；"
            f"支持类型: {supported_label}（第 {attempt} 次检查）"
        )
        time.sleep(10)

    props = last_group.get("properties", {})
    raise TimeoutError(
        "等待评估组关联服务器超时。"
        f"当前 machineCount={props.get('machineCount')}，"
        f"supportedAssessmentTypes={props.get('supportedAssessmentTypes')}"
    )


def run_azure_migrate_assessment(
    token: str,
    subscription_id: str,
    resource_group: str,
    account_name: str,
    assessment_name: str,
    csv_bytes: bytes,
    annual_budget_text: Optional[str],
    progress: Callable[[str], None],
) -> Dict[str, Any]:
    safe_base = _safe_azure_name(account_name, f"poe-{_date_prefix()}", max_len=36).lower()
    run_suffix = str(int(time.time()))
    short_run_suffix = run_suffix[-6:]
    project_name = _safe_azure_name(safe_base, "poe", "project", 55)
    site_name = _safe_azure_name(safe_base, "poe", f"site{short_run_suffix}", 24)
    master_site_name = _safe_azure_name(safe_base, "poe", "masterSite", 55)
    collector_name = _safe_azure_name(safe_base, "poe", f"collector-{short_run_suffix}", 55)
    group_name = _safe_azure_name(safe_base, "poe", f"group-{run_suffix}", 55)
    assessment_resource_name = _safe_azure_name(assessment_name, "poe-assessment", max_len=55)
    project_location = resolve_migrate_project_location(subscription_id, resource_group, token)
    annual_budget = parse_annual_budget_usd(annual_budget_text)
    migrate_project_id = _resource_id(
        subscription_id,
        resource_group,
        f"providers/Microsoft.Migrate/migrateProjects/{project_name}",
    )
    assessment_project_id = _resource_id(
        subscription_id,
        resource_group,
        f"providers/Microsoft.Migrate/assessmentProjects/{project_name}",
    )
    master_site_id = _resource_id(
        subscription_id,
        resource_group,
        f"providers/Microsoft.OffAzure/masterSites/{master_site_name}",
    )
    import_site_id = _resource_id(
        subscription_id,
        resource_group,
        f"providers/Microsoft.OffAzure/importSites/{site_name}",
    )

    progress("注册 Microsoft.Migrate 与 Microsoft.OffAzure 资源提供程序...")
    register_azure_provider(subscription_id, "Microsoft.Migrate", token)
    register_azure_provider(subscription_id, "Microsoft.OffAzure", token)

    # ── Step 1: 创建 migrateProject（Portal 可见） ──
    progress(f"创建 Azure Migrate 项目（区域：{project_location}）...")
    migrate_project_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/migrateProjects/{project_name}"
        f"?api-version={AZURE_MIGRATE_PROJECTS_API_VERSION}"
    )
    existing_mp = _try_get_existing_resource(migrate_project_path, token)
    if existing_mp:
        mp_id = existing_mp.get("id", project_name)
        progress(f"成功复用 Azure Migrate 项目：{project_name}")
    else:
        migrate_project_body = {
            "properties": {},
            "location": project_location,
            "tags": {"Migrate Project": project_name, "createdBy": "POE Workflow"},
            "identity": {"type": "SystemAssigned"},
        }
        try:
            mp_result = azure_arm_request("PUT", migrate_project_path, token, migrate_project_body)
        except Exception:
            migrate_project_body.pop("identity", None)
            mp_result = azure_arm_request("PUT", migrate_project_path, token, migrate_project_body)
        mp_id = mp_result.get("id", project_name)
        progress(f"成功创建 Azure Migrate 项目：{project_name}")

    # ── Step 2: 注册 Portal 同款 Discovery Import 与 Assessment 工具 ──
    progress("注册 ServerDiscovery_Import 与 ServerAssessment 工具...")
    for tool in ("ServerDiscovery_Import", "ServerAssessment"):
        try:
            register_migrate_tool(subscription_id, resource_group, project_name, tool, token)
            progress(f"  ✅ 工具已注册: {tool}")
        except Exception:
            progress(f"  ℹ️ 工具可能已注册: {tool}")

    # ── Step 3: 创建 ServerAssessment Solution ──
    assessment_solution_name = "Servers-Assessment-ServerAssessment"
    assessment_solution_path = _migrate_solution_path(
        subscription_id, resource_group, project_name, assessment_solution_name
    )
    existing_sol = _try_get_existing_resource(assessment_solution_path, token)
    if existing_sol:
        assessment_solution_id = existing_sol.get("id", "")
        progress(f"成功复用评估 Solution：{assessment_solution_name}")
    else:
        sol_result = put_migrate_solution(
            subscription_id,
            resource_group,
            project_name,
            assessment_solution_name,
            {
                "tool": "ServerAssessment",
                "purpose": "Assessment",
                "goal": "Servers",
                "status": "Active",
                "details": _servers_solution_details({
                    "projectId": assessment_project_id,
                    "avsAssessment": "0",
                    "azureSqlAssessment": "0",
                    "azureVmAssessment": "0",
                    "azureWebAppAssessment": "0",
                    "businessCaseCount": "0",
                }),
            },
            token,
        )
        assessment_solution_id = sol_result.get("id", "")
        progress(f"成功创建评估 Solution：{assessment_solution_name}")

    # ── Step 4: 创建 assessmentProject 并关联 Solution ──
    progress("创建 Assessment Project...")
    ap_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"?api-version={AZURE_MIGRATE_API_VERSION}"
    )
    existing_ap = _try_get_existing_resource(ap_path, token)
    ap_body = {
        "kind": "Migrate",
        "properties": {
            "projectStatus": "Active",
            "assessmentSolutionId": assessment_solution_id,
            "publicNetworkAccess": "Enabled",
        },
        "location": project_location,
        "tags": {"createdBy": "POE Workflow"},
    }
    if existing_ap and str(existing_ap.get("kind") or "").lower() == "migrate":
        ap_id = existing_ap.get("id", project_name)
        progress(f"成功复用 Assessment Project：{project_name}")
    else:
        ap_result = azure_arm_request("PUT", ap_path, token, ap_body)
        ap_id = ap_result.get("id", project_name)
        progress(f"成功创建 Assessment Project：{project_name}")

    # Portal 的评估 blade 通过 Assessment Solution 的 projectId 找到 assessmentProject。
    assessment_solution = put_migrate_solution(
        subscription_id,
        resource_group,
        project_name,
        assessment_solution_name,
        {
            "tool": "ServerAssessment",
            "purpose": "Assessment",
            "goal": "Servers",
            "status": "Active",
            "details": _servers_solution_details({
                "projectId": assessment_project_id,
                "avsAssessment": "0",
                "azureSqlAssessment": "0",
                "azureVmAssessment": "0",
                "azureWebAppAssessment": "0",
                "businessCaseCount": "0",
            }),
        },
        token,
    )
    assessment_solution_id = assessment_solution.get("id", assessment_solution_id)
    progress("  ✅ Assessment Solution 已关联 assessmentProject")

    # ── Step 5: 创建 Portal Discovery Import 链路（Master Site + Discovery Solution + Import Site） ──
    progress("创建 Portal 可识别的 Discovery Import 链路...")
    discovery_solution_name = "Servers-Discovery-ServerDiscovery_Import"
    discovery_solution_id = _migrate_solution_id(
        subscription_id, resource_group, project_name, discovery_solution_name
    )
    master_site_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.OffAzure/masterSites/{master_site_name}"
        f"?api-version={AZURE_OFFAZURE_API_VERSION}"
    )
    existing_master_site = _try_get_existing_resource(master_site_path, token)
    existing_sites = []
    if existing_master_site:
        existing_sites = existing_master_site.get("properties", {}).get("sites") or []
    master_site_result = azure_arm_request("PUT", master_site_path, token, {
        "kind": "Migrate",
        "location": project_location,
        "tags": {"Migrate Project": project_name, "createdBy": "POE Workflow"},
        "properties": {
            "allowMultipleSites": True,
            "publicNetworkAccess": "Enabled",
            "sites": existing_sites,
        },
    })
    master_site_id = master_site_result.get("id", master_site_id)
    progress(f"成功创建 Master Site：{master_site_name}")
    ensure_portal_menu_solutions(
        subscription_id,
        resource_group,
        project_name,
        master_site_id,
        token,
        progress,
    )

    put_migrate_solution(
        subscription_id,
        resource_group,
        project_name,
        discovery_solution_name,
        {
            "tool": "ServerDiscovery_Import",
            "purpose": "Discovery",
            "goal": "Servers",
            "status": "Inactive",
            "details": _servers_solution_details({
                "importSiteId": import_site_id,
            }),
        },
        token,
    )
    progress("  ✅ Discovery Import Solution 已关联 importSite")

    # ── Step 6: 创建 Import Site ──
    progress("创建 Import Site...")
    site_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.OffAzure/importSites/{site_name}"
        f"?api-version={AZURE_OFFAZURE_API_VERSION}"
    )
    existing_site = _try_get_existing_resource(site_path, token)
    if existing_site:
        site_id = existing_site.get("id", site_name)
        progress(f"成功复用 Import Site：{site_name}")
    else:
        site_result = azure_arm_request("PUT", site_path, token, {
            "location": project_location,
            "properties": {
                "masterSiteId": master_site_id,
                "discoverySolutionId": discovery_solution_id,
            },
        })
        site_id = site_result.get("id", site_name)
        progress(f"成功创建 Import Site：{site_name}")

    normalized_sites = {str(site).lower(): site for site in existing_sites}
    normalized_sites.setdefault(import_site_id.lower(), import_site_id)
    azure_arm_request("PUT", master_site_path, token, {
        "kind": "Migrate",
        "location": project_location,
        "tags": {"Migrate Project": project_name, "createdBy": "POE Workflow"},
        "properties": {
            "allowMultipleSites": True,
            "publicNetworkAccess": "Enabled",
            "sites": list(normalized_sites.values()),
        },
    })
    progress("  ✅ Master Site 已关联 Import Site")

    # ── Step 7: 创建 Import Collector，关联 Import Site 到 Assessment Project ──
    progress("创建 Import Collector 关联 Import Site 到 Assessment Project...")
    collector_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"/importcollectors/{collector_name}?api-version={AZURE_MIGRATE_API_VERSION}"
    )
    discovery_site_id = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.OffAzure/importSites/{site_name}"
    )
    coll_result = azure_arm_request("PUT", collector_path, token, {
        "properties": {"discoverySiteId": discovery_site_id}
    })
    progress(f"成功创建 Import Collector：{collector_name}")

    # ── Step 8: 获取 SAS URL 并上传 CSV ──
    progress("获取 CSV 上传地址并上传服务器清单...")
    import_uri_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.OffAzure/importSites/{site_name}/importUri"
        f"?api-version={AZURE_OFFAZURE_API_VERSION}"
    )
    import_uri_payload = azure_arm_request("POST", import_uri_path, token, {})
    sas_url = _extract_sas_url(import_uri_payload)
    if not sas_url:
        raise RuntimeError(f"Azure 未返回可用的 CSV 上传 SAS URL：{import_uri_payload}")
    import_job_arm_id = import_uri_payload.get("jobArmId") if isinstance(import_uri_payload, dict) else None

    upload_response = requests.put(
        sas_url,
        data=csv_bytes,
        headers={"x-ms-blob-type": "BlockBlob", "Content-Type": "text/csv"},
        timeout=180,
    )
    if upload_response.status_code >= 400:
        raise RuntimeError(f"上传 CSV 到 Azure Migrate 失败：{upload_response.status_code} {upload_response.text[:800]}")
    progress(f"  ✅ CSV 已上传（{len(csv_bytes)} 字节）")

    # ── Step 9: 触发 Import Job（回传 importUri 返回的 SasUriResponse） ──
    progress("触发 Import Job 导入服务器清单...")
    import_trigger_body = dict(import_uri_payload) if isinstance(import_uri_payload, dict) else {}
    import_trigger_body["uri"] = sas_url
    if import_job_arm_id:
        import_trigger_body["jobArmId"] = import_job_arm_id
    job_result = azure_arm_request("POST", import_uri_path, token, import_trigger_body)
    import_job_arm_id = (
        job_result.get("jobArmId")
        or import_job_arm_id
        or job_result.get("id")
        if isinstance(job_result, dict)
        else import_job_arm_id
    )
    progress("成功触发 Import Job")

    # ── Step 10: 等待 OffAzure Import Site 完成 CSV 解析 ──
    imported_site_machines = wait_for_import_site_import(
        subscription_id,
        resource_group,
        site_name,
        token,
        progress,
        job_arm_id=import_job_arm_id,
    )
    progress(f"  ✅ Import Site 已导入 {len(imported_site_machines)} 台服务器")

    # Import Collector 在导入完成后再 PUT 一次，触发 assessmentProject 拉取刚导入的机器。
    azure_arm_request("PUT", collector_path, token, {
        "properties": {"discoverySiteId": discovery_site_id}
    })
    progress("  ✅ Import Collector 已刷新同步")

    portal_inventory_count = wait_for_portal_inventory_summary(
        subscription_id,
        resource_group,
        project_name,
        len(imported_site_machines),
        token,
        progress,
    )
    progress(f"  ✅ Azure Portal 全部库存汇总已刷新: {portal_inventory_count} 台服务器")

    # ── Step 11: 等待 Assessment Project 可读取机器 ──
    machines = wait_for_project_machines(
        subscription_id,
        resource_group,
        project_name,
        collector_name,
        token,
        progress,
        site_name=site_name,
    )
    machine_ids = [machine.get("id") for machine in machines if machine.get("id")]
    if not machine_ids:
        raise RuntimeError("Azure Migrate 未返回可加入评估的服务器，请检查 CSV 导入结果。")
    progress(f"  ✅ 已发现 {len(machine_ids)} 台服务器")

    # ── Step 12: 创建评估组并通过 updateMachines 加入服务器 ──
    progress("创建评估组并添加全部服务器 workload...")
    group_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"/groups/{group_name}?api-version={AZURE_MIGRATE_API_VERSION}"
    )
    existing_group = _try_get_existing_resource(group_path, token)
    if existing_group:
        existing_group_type = str(existing_group.get("properties", {}).get("groupType", "")).strip().lower()
        if existing_group_type and existing_group_type != "import":
            # 同名组如果不是 Import 类型，groupType 无法修改，只能改名新建。
            group_name = _safe_azure_name(group_name, "poe-group", f"-imp-{int(time.time())}", 55)
            group_path = (
                f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
                f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
                f"/groups/{group_name}?api-version={AZURE_MIGRATE_API_VERSION}"
            )
            existing_group = None
            progress(f"  ⚠️ 发现同名评估组类型为 {existing_group_type}，改用新组名: {group_name}")

    if not existing_group:
        azure_arm_request("PUT", group_path, token, {
            "properties": {"groupType": "Import"},
            "eTag": "",
        })
        progress(f"  ✅ 已创建 Import 评估组: {group_name}")
    else:
        progress(f"  ✅ 复用评估组: {group_name}")

    update_machines_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"/groups/{group_name}/updateMachines?api-version={AZURE_MIGRATE_API_VERSION}"
    )
    azure_arm_request("POST", update_machines_path, token, {
        "eTag": "*",
        "properties": {
            "operationType": "Add",
            "machines": machine_ids,
        },
    })
    progress(f"  ✅ 已向评估组添加服务器: {len(machine_ids)} 台")

    group_payload = wait_for_group_machine_membership(
        group_path,
        token,
        expected_machine_count=len(machine_ids),
        progress=progress,
    )
    supported_types = {
        str(item).strip()
        for item in (group_payload.get("properties", {}).get("supportedAssessmentTypes") or [])
        if str(item).strip()
    }
    if supported_types and "MachineAssessment" not in supported_types:
        progress(
            "  ℹ️ 评估组尚未显式返回 MachineAssessment；"
            "继续按服务器评估类型创建 Azure VM 评估。"
        )

    # ── Step 13: 创建评估 ──
    progress("创建 Azure Migrate 评估...")
    assessment_path = (
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"/groups/{group_name}/assessments/{assessment_resource_name}"
        f"?api-version={AZURE_MIGRATE_API_VERSION}"
    )
    assessment_body = _build_assessment_body()
    if annual_budget is not None:
        progress(f"已解析用户预估年消耗：{_format_usd(annual_budget)}")
    else:
        progress("未解析到有效预估年消耗；评估会按 Portal 默认设置创建。")
    azure_arm_request("PUT", assessment_path, token, assessment_body)
    assessment = wait_for_assessment_complete(
        subscription_id, resource_group, project_name, group_name, assessment_resource_name, token, progress
    )
    assessment, tuning_history, budget_target_met = tune_assessment_to_budget(
        subscription_id=subscription_id,
        resource_group=resource_group,
        project_name=project_name,
        group_name=group_name,
        assessment_name=assessment_resource_name,
        assessment_path=assessment_path,
        assessment_body=assessment_body,
        assessment=assessment,
        annual_budget=annual_budget,
        token=token,
        progress=progress,
    )
    try:
        refresh_migrate_project_summary(subscription_id, resource_group, project_name, token)
    except Exception:
        pass

    progress("读取评估结果并下载 Azure Migrate Portal 同源 Excel 报告...")
    assessed_machines = azure_arm_list(
        f"/subscriptions/{subscription_id}/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Migrate/assessmentProjects/{project_name}"
        f"/groups/{group_name}/assessments/{assessment_resource_name}/assessedMachines"
        f"?api-version={AZURE_MIGRATE_API_VERSION}",
        token,
    )
    excel_bytes = download_assessment_report(
        subscription_id, resource_group, project_name, group_name, assessment_resource_name, token
    )
    progress(f"  ✅ Azure Migrate 导出报告已下载（{len(excel_bytes)} 字节）")
    return {
        "project_name": project_name,
        "site_name": site_name,
        "collector_name": collector_name,
        "group_name": group_name,
        "assessment_name": assessment_resource_name,
        "portal_inventory_count": portal_inventory_count,
        "migrate_project_id": migrate_project_id,
        "assessment_project_id": assessment_project_id,
        "import_site_id": import_site_id,
        "assessment": assessment,
        "assessed_machines": assessed_machines,
        "excel_bytes": excel_bytes,
        "budget_target": annual_budget,
        "monthly_cost": assessment_monthly_total_cost(assessment),
        "annualized_cost": assessment_monthly_total_cost(assessment) * 12,
        "budget_target_met": budget_target_met,
        "tuning_history": tuning_history,
    }


def create_poe_zip(artifacts: List[Dict[str, Any]]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for artifact in artifacts:
            zip_file.writestr(artifact["file_name"], artifact["bytes"])
    return buffer.getvalue()


def get_existing_solution_text(current_doc_type: str) -> Optional[str]:
    key = "solution_text" if current_doc_type == "AI" else "infra_text"
    text = st.session_state.get(key)
    if isinstance(text, str) and text.strip():
        return text.strip()
    return None


def create_solution_artifact_from_text(
    current_doc_type: str,
    content: str,
    customer_name: str,
    account_name: str,
) -> Dict[str, Any]:
    if current_doc_type == "AI":
        docx_bytes = create_solution_docx(content=content, customer_name=customer_name)
        file_name = f"{account_name}-Solution Architecture.docx"
    else:
        docx_bytes = create_infra_docx(content=content, customer_name=customer_name)
        file_name = f"{account_name}-Infra Solution Architecture.docx"
    return {"content": content, "bytes": docx_bytes, "file_name": file_name}


def create_pov_artifact_from_text(
    content: str,
    customer_name: str,
    account_name: str,
) -> Dict[str, Any]:
    docx_bytes = create_pov_docx(content=content, customer_name=customer_name)
    return {
        "content": content,
        "bytes": docx_bytes,
        "file_name": f"{account_name}-PostAssessment POVdeployment.docx",
    }


def get_generated_migrate_csv_text() -> Optional[str]:
    csv_text = st.session_state.get("csv_code")
    if isinstance(csv_text, str) and csv_text.strip():
        return csv_text.strip()
    return None


def resolve_auto_inventory_csv(uploaded_inventory: Any) -> tuple[Optional[bytes], str]:
    if uploaded_inventory is not None:
        return uploaded_inventory.getvalue(), "手动上传 CSV"
    generated_csv = get_generated_migrate_csv_text()
    if generated_csv:
        return generated_csv.encode("utf-8-sig"), "Azure Migrate CSV 标签页生成"
    return None, "未提供"


def format_auto_poe_log(message: str) -> str:
    text = re.sub(r"\s+", " ", str(message or "")).strip()
    text = re.sub(r"^[^\w\u4e00-\u9fff]+", "", text).strip()
    text = text.removesuffix("...").strip()
    text = re.sub(r"https?://\S+", "[url]", text)
    text = re.sub(r"/subscriptions/[^\s；;，,]+", lambda match: match.group(0).rstrip("/").split("/")[-1], text, flags=re.I)
    text = text.replace("✅", "").replace("⚠️", "").replace("ℹ️", "").strip()
    if not text:
        return ""
    return text


def should_display_auto_poe_log(message: str) -> bool:
    text = str(message or "").strip()
    if not text:
        return False
    interim_markers = [
        "正在",
        "开始",
        "等待",
        "仍在",
        "继续",
        "第 ",
        "注册 ",
        "读取评估结果",
    ]
    result_markers = [
        "成功",
        "完成",
        "已",
        "复用",
        "解析",
        "目标区间",
        "不在目标区间",
        "未填写",
        "未解析",
        "请到 Azure Portal",
    ]
    if any(marker in text for marker in interim_markers) and not any(marker in text for marker in result_markers):
        return False
    return any(marker in text for marker in result_markers)


def run_full_auto_poe(
    current_doc_type: str,
    customer_name: str,
    account_name: str,
    customer_bg: str,
    solution_ref: str,
    infra_ref: str,
    pov_ref: str,
    token: str,
    subscription_id: str,
    resource_group: str,
    assessment_name: str,
    csv_bytes: bytes,
    annual_budget_text: Optional[str],
    pov_start: Optional[datetime.date],
    pov_end: Optional[datetime.date],
    vendor_team: str,
    existing_solution_text: Optional[str],
    existing_pov_text: Optional[str],
    progress: Callable[[str], None],
) -> Dict[str, Any]:
    if existing_solution_text and existing_solution_text.strip():
        solution_artifact = create_solution_artifact_from_text(
            current_doc_type,
            existing_solution_text.strip(),
            customer_name,
            account_name,
        )
        progress(f"成功复用文档：{solution_artifact['file_name']}")
    else:
        solution_artifact = generate_solution_artifact(
            current_doc_type, customer_name, account_name, customer_bg, solution_ref, infra_ref
        )
        progress(f"成功生成文档：{solution_artifact['file_name']}")

    target_key = "solution_text" if current_doc_type == "AI" else "infra_text"
    st.session_state[target_key] = solution_artifact["content"]
    st.session_state["customer_name"] = customer_name
    st.session_state["account_name"] = account_name

    if existing_pov_text and existing_pov_text.strip():
        pov_artifact = create_pov_artifact_from_text(
            existing_pov_text.strip(),
            customer_name,
            account_name,
        )
        progress(f"成功复用文档：{pov_artifact['file_name']}")
    else:
        if not pov_start or not pov_end:
            raise RuntimeError("缺少 POV 开始日期或结束日期，无法生成 POV 部署计划。")
        if not has_meaningful_pov_team(vendor_team):
            raise RuntimeError("缺少 POV 项目人员，无法生成 POV 部署计划。")
        pov_artifact = generate_pov_artifact(
            solution_artifact["content"],
            customer_name,
            account_name,
            pov_ref,
            pov_start,
            pov_end,
            vendor_team,
        )
        progress(f"成功生成文档：{pov_artifact['file_name']}")
    st.session_state["pov_text"] = pov_artifact["content"]
    st.session_state["pov_source_doc_type"] = current_doc_type

    migrate_result = run_azure_migrate_assessment(
        token, subscription_id, resource_group, account_name, assessment_name, csv_bytes, annual_budget_text, progress
    )
    assessment_artifact = {
        "file_name": f"{account_name}-Azure Migrate Assessment.xlsx",
        "bytes": migrate_result["excel_bytes"],
    }
    progress(f"成功下载评估报告：{assessment_artifact['file_name']}")

    zip_bytes = create_poe_zip([solution_artifact, pov_artifact, assessment_artifact])
    progress(f"成功生成 POE 套件：{account_name}-POE-Complete.zip")
    return {
        "zip_bytes": zip_bytes,
        "zip_name": f"{account_name}-POE-Complete.zip",
        "solution": solution_artifact,
        "pov": pov_artifact,
        "assessment": assessment_artifact,
        "migrate": migrate_result,
    }


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
    render_section_head(
        "全自动 POE 套件",
        "完成下方状态后启动。",
        render_pill("长任务", "accent"),
    )
    azure_logged_in = _is_azure_token_valid()
    selected_subscription = None
    selected_resource_group = None
    uploaded_inventory = None
    resolved_customer_name = customer_name.strip() or str(st.session_state.get("customer_name") or "").strip()
    resolved_account_name = (
        account_name.strip()
        or str(st.session_state.get("account_name") or "").strip()
        or resolved_customer_name
    )
    resolved_budget_text = budget.strip() or str(st.session_state.get("budget") or "").strip()
    default_assessment_name = _safe_azure_name(resolved_account_name or resolved_customer_name, "poe", "assessment", 55)
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
    generated_csv_text = get_generated_migrate_csv_text()
    has_existing_solution = bool(existing_solution_text)
    has_existing_pov = bool(existing_pov_text)
    pov_dates_ready = bool(pov_start and pov_end and pov_end >= pov_start)
    pov_team_ready = has_meaningful_pov_team(vendor_team)
    solution_ready = has_existing_solution or (bool(resolved_customer_name) and bool(customer_bg.strip()))
    pov_ready = has_existing_pov or (pov_dates_ready and pov_team_ready)
    status_slot = st.empty()

    token = st.session_state.get("azure_token")
    with st.container(border=True):
        login_col, azure_col = st.columns([0.95, 2.05])
        with login_col:
            st.markdown("**1. Azure 登录**")
            if azure_logged_in:
                st.success(f"当前账户：{st.session_state.get('azure_user', 'Azure 用户')}")
                if st.button("退出 Azure 登录", use_container_width=True, key="btn_azure_logout"):
                    clear_azure_login()
                    st.rerun()
            else:
                if st.button("登录 Microsoft Azure 账户", type="primary", use_container_width=True, key="btn_azure_login"):
                    try:
                        msal_device_code_login()
                        st.rerun()
                    except Exception as exc:
                        st.error(f"登录失败：{exc}")

        with azure_col:
            st.markdown("**2. Azure 目标位置**")
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
                        sub_label = st.selectbox(
                            "订阅",
                            subscription_labels,
                            index=default_index,
                            key="auto_subscription_select",
                        )
                        selected_subscription = subscriptions[subscription_labels.index(sub_label)]
                        subscription_id = selected_subscription.get("subscriptionId")
                        st.session_state["azure_subscription_id"] = subscription_id
                        st.session_state["azure_subscription_name"] = selected_subscription.get("displayName", "")

                        groups = list_azure_resource_groups(subscription_id, token)
                        if groups:
                            group_labels = [_resource_group_label(group) for group in groups]
                            rg_label = st.selectbox(
                                "资源组",
                                group_labels,
                                key="auto_resource_group_select",
                            )
                            selected_resource_group = groups[group_labels.index(rg_label)]
                            st.session_state["azure_resource_group"] = selected_resource_group.get("name")
                        else:
                            st.warning("当前订阅下没有可读取的资源组。")
                    else:
                        st.warning("当前账户没有可读取的 Azure 订阅。")
                except Exception as exc:
                    st.error(f"读取 Azure 订阅或资源组失败：{exc}")
            else:
                st.info("登录后会在这里选择订阅和资源组。")

        input_col, action_col = st.columns([2, 1])
        with input_col:
            st.markdown("**3. 输入材料**")
            uploaded_inventory = st.file_uploader(
                "上传服务器清单 CSV",
                type=["csv"],
                key="auto_server_inventory_csv",
                help="手动上传优先；未上传时会自动复用 Azure Migrate CSV 标签页生成的 CSV。",
            )
            if uploaded_inventory is not None and generated_csv_text:
                st.info("已检测到手动上传和已生成 CSV，本次会优先使用手动上传的 CSV。")
            elif uploaded_inventory is not None:
                st.success("本次会使用手动上传的 CSV。")
            elif generated_csv_text:
                st.success("已检测到 Azure Migrate CSV 标签页生成的 CSV，本次会自动复用。")
            else:
                st.warning("请上传 CSV，或先到「Azure Migrate CSV」标签页生成 CSV。")
        with action_col:
            st.markdown("**4. 评估命名**")
            assessment_name = st.text_input("评估名称", value=default_assessment_name, key="auto_assessment_name")

        inventory_ready = uploaded_inventory is not None or bool(generated_csv_text)
        readiness_items = [
            ("客户名称", bool(resolved_customer_name), "用于文档标题和输出文件名"),
            ("方案文档", solution_ready, "已生成则复用；未生成时会按客户背景生成"),
            ("预估年消耗", budget_value is not None and budget_value > 0, "用于校准迁移评估年化估算"),
            ("POV 文档/输入", pov_ready, "已生成则复用；未生成时需填写时间区间和项目人员"),
            ("Azure 登录", azure_logged_in, "用于创建 Azure Migrate 项目和评估"),
            ("订阅与资源组", bool(selected_subscription and selected_resource_group), "评估资源会创建到这里"),
            ("服务器清单 CSV", inventory_ready, "手动上传优先，否则复用 Azure Migrate CSV 标签页生成的 CSV"),
            ("评估名称", bool(assessment_name.strip()), "用于 Azure Migrate 评估资源"),
        ]
        workflow_ready = all(item[1] for item in readiness_items)
        with status_slot.container():
            render_workflow_steps([
                {
                    "title": "登录 Azure",
                    "state": "done" if azure_logged_in else "ready",
                },
                {
                    "title": "选择目标",
                    "state": "done" if selected_subscription and selected_resource_group else ("ready" if azure_logged_in else "blocked"),
                },
                {
                    "title": "准备 CSV",
                    "state": "done" if inventory_ready else ("ready" if selected_subscription and selected_resource_group else "blocked"),
                },
                {
                    "title": "生成套件",
                    "state": "ready" if workflow_ready else "blocked",
                },
            ])
            render_readiness(readiness_items)

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
            if not has_existing_pov and not pov_dates_ready:
                st.warning("请先到「POV 部署计划」标签页填写 POV 开始日期和结束日期。")
                return
            if not has_existing_pov and not pov_team_ready:
                st.warning("请先到「POV 部署计划」标签页填写乙方项目人员。")
                return
            if not _is_azure_token_valid():
                st.warning("请先登录 Microsoft Azure 账户。")
                return
            if not selected_subscription or not selected_resource_group:
                st.warning("请选择订阅和资源组。")
                return
            csv_bytes, csv_source = resolve_auto_inventory_csv(uploaded_inventory)
            if not csv_bytes:
                st.warning("请上传服务器清单 CSV，或先到「Azure Migrate CSV」标签页生成 CSV。")
                return
            if not assessment_name.strip():
                st.warning("请输入评估名称。")
                return

            try:
                st.session_state["auto_poe_running"] = True
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
                            log_placeholder.code("\n".join(log_lines[-120:]), language="text")

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
                        csv_bytes=csv_bytes,
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
                        "budget_target_met": result["migrate"].get("budget_target_met", True),
                        "csv_source": csv_source,
                    }
                    status.update(label="完整 POE 套件生成完成", state="complete")
            except Exception as exc:
                st.session_state.pop("auto_poe_running", None)
                st.session_state["auto_poe_error"] = str(exc)
                st.error(f"全自动生成失败：{exc}")

        if "auto_poe_zip_bytes" in st.session_state:
            result = st.session_state.get("auto_poe_result", {})
            annualized_cost = result.get("annualized_cost")
            budget_target = result.get("budget_target")
            render_auto_poe_result(
                customer_name=result.get("customer_name") or customer_name.strip() or account_name.strip() or "该客户",
                generated_items=[
                    ("解决方案架构文档", result.get("solution_file_name") or "-"),
                    ("POV 文档", result.get("pov_file_name") or "-"),
                    ("迁移评估文档", result.get("assessment_file_name") or "-"),
                ],
                migrate_items=[
                    ("Azure Migrate 项目", result.get("project_name") or "-"),
                    ("迁移评估名称", result.get("assessment_name") or "-"),
                    ("CSV 来源", result.get("csv_source") or "-"),
                    ("Portal 库存", f"{result.get('portal_inventory_count', 0)} 台"),
                    ("评估服务器数", f"{result.get('machine_count', 0)} 台"),
                    ("年化估算", _format_usd(annualized_cost)),
                    ("用户预估", _format_usd(budget_target)),
                ],
            )
            if budget_target and not result.get("budget_target_met", True):
                st.warning("自动调整 3 轮后仍未落入用户预估年消耗的 100%-120% 区间，请在 Azure Portal 的评估设置中手动调整。")
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
def main():
    render_app_header()

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
            ]:
                st.session_state.pop(key, None)
            st.rerun()

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
        account_name = st.text_input("账户名", placeholder="例如：Tetherflow", help="用于生成下载文件名的前缀")
    with c1:
        customer_name = st.text_input("客户名称", placeholder="例如：宇宙无敌科技有限公司")
    with c2:
        budget = st.text_input("预估年消耗 (USD)", placeholder="例如：500k+")

    customer_bg = st.text_area(
        "客户背景信息",
        placeholder="粘贴客户背景资料，包括行业、规模、现有 IT 环境、核心需求和已知约束。",
        height=125,
    )

    st.divider()

    # ════════════════════════════════════════════════════
    # Tab 布局
    # ════════════════════════════════════════════════════
    tab_auto, tab_sol, tab_pov, tab_csv, tab_yearly = st.tabs([
        "全自动POE生成", "解决方案文档", "POV 部署计划", "Azure Migrate CSV", "年度价格表"
    ])

    dp = _date_prefix()  # 日期前缀

    # ─────────── Tab 1: 全自动 POE 生成 ───────────
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

    # ─────────── Tab 2: 解决方案文档 ───────────
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
                            )
                            if customer_bg.strip():
                                user_ctx += (
                                    f"## 客户背景信息\n{customer_bg.strip()}\n\n"
                                )
                            user_ctx += (
                                f"## 已有解决方案文档（请基于以上客户信息和以下已有文档，按照要求的章节格式重新整理生成，不要照抄原文）\n\n"
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
                    pov_start = st.date_input("POV 开始日期", value=None, key="pov_start_date")
                with dc2:
                    pov_end = st.date_input(
                        "POV 结束日期",
                        value=None,
                        key="pov_end_date",
                    )

                vendor_team = st.text_area(
                    "乙方项目人员（我方团队）",
                    value=(
                        "技术负责人: \n"
                        "Azure架构师: \n"
                    ),
                    height=120,
                    help="只需填写乙方（我方）人员，甲方人员由 AI 根据客户背景自动生成",
                    key="pov_vendor_team",
                )

                has_pov = "pov_text" in st.session_state
                pov_label = "重新生成" if has_pov else "生成 POV 部署计划"
                if st.button(pov_label, type="primary", use_container_width=True, key="btn_pov"):
                    if not pov_start or not pov_end:
                        st.warning("请先选择 POV 开始日期和结束日期。")
                        st.stop()
                    if pov_end < pov_start:
                        st.warning("POV 结束日期不能早于开始日期。")
                        st.stop()
                    if not has_meaningful_pov_team(vendor_team):
                        st.warning("请填写乙方项目人员，不能只保留默认空模板。")
                        st.stop()
                    try:
                        pov_prompt = build_pov_prompt(solution, customer, pov_start, pov_end, vendor_team, pov_ref)
                        with st.spinner("正在生成 POV 部署计划..."):
                            pov_text = call_azure_openai(POV_SYSTEM_PROMPT, pov_prompt)
                            st.session_state["pov_text"] = pov_text
                            st.session_state["pov_source_doc_type"] = current_doc_type
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
