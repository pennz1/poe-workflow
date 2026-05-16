---
name: poe-generation
description: '生成微软售前 POE 交付文档套件。Use when: 需要为客户生成方案文档、POV部署计划、Azure Migrate CSV；输入客户名称/背景/预算等信息后一键生成全套售前交付物。'
argument-hint: '客户名称、背景信息、预算等'
---

# 微软客户 POE 文档生成

为微软售前团队生成完整的 POE（Proof of Engagement）交付文档套件，包含解决方案架构文档、POV 部署计划和 Azure Migrate CSV。

## 触发场景

- 需要为新客户生成全套 POE 文档
- 输入客户信息后自动生成方案文档
- 生成 POV 部署计划
- 生成 Azure Migrate 导入 CSV

## 所需输入

向用户收集以下信息（如未提供则逐一询问）：

| 字段 | 必填 | 说明 |
|------|------|------|
| 客户名称 | ✅ | 用于文档标题（如"深圳跃瓦创新科技"） |
| 账户名 | ✅ | 用于文件命名前缀（如"Tetherflow"），英文 |
| 文档类型 | ✅ | `AI 解决方案` 或 `Infra 基础设施` |
| 客户背景信息 | ✅ | 行业/规模/痛点/需求的详细描述 |
| 预估年消耗 (USD) | ✅ | Azure 年消耗预算（如 50000） |
| POV 开始日期 | ✅ | 部署计划起始日期 |
| POV 结束日期 | ✅ | 部署计划结束日期 |
| 乙方技术负责人 | ✅ | 项目人员姓名 |
| 乙方架构师 | ✅ | 项目人员姓名 |

## 输出文档

| 文档 | 格式 | 文件名模式 |
|------|------|------|
| 解决方案架构文档 | Markdown → .docx | `{账户名}-Solution Architecture.docx` |
| POV 部署计划 | Markdown → .docx | `{账户名}-PostAssessment POVdeployment.docx` |
| Azure Migrate CSV | .csv | `{账户名}-Azure migrate report.csv` |

## 执行流程

### 步骤 1：收集客户信息

确保所有必填字段已收集。如果用户一次性提供了所有信息，直接进入下一步。

### 步骤 2：生成解决方案架构文档

使用 [解决方案 Prompt](./references/prompts.md#解决方案文档-prompt) 调用 LLM 生成 Markdown 格式的方案文档。

**关键规则：**
- 第一行必须是 `# [客户名称] - [具体方案名称]`
- AI 类型使用 8 章节结构（摘要→架构概览→业务背景→需求摘要→详细设计→安全→集成→资源）
- Infra 类型使用 10 章节结构（执行摘要→架构概览→业务背景→需求概述→方案设计→详细架构→集成→数据→安全→基础设施需求）
- 严格禁止使用项目符号列表（-、*、•）
- 严禁使用中国区域（China East/North），只用全球区域
- 严禁提及预算金额
- 必须包含多种 Azure AI 服务，禁止只用 Azure OpenAI

### 步骤 3：生成 POV 部署计划

基于步骤 2 的方案文档，使用 [POV Prompt](./references/prompts.md#pov-部署计划-prompt) 生成部署计划。

**关键规则：**
- 标题格式：`# [客户名称] - [方案核心描述] POV 部署计划`
- 最多 3 个阶段
- 任务表格只安排工作日（跳过周六日），但文档中不提"周末""工作日"
- 日期格式统一为 `M月D日`
- 自动生成甲方人员（中文名，2-3人）

### 步骤 4：生成 Azure Migrate CSV

使用 [CSV Prompt](./references/prompts.md#azure-migrate-csv-prompt) 根据预算倒推服务器配置。

**关键规则：**
- 必填列：Server name, Cores, Memory (In MB), OS name
- Server name 格式：`服务类型-区域-规模-序号`
- 预算档位：15k / 50k / 100k / 250k USD
- VM 配置要合理，总成本接近年消耗预算

### 步骤 5：输出交付

将生成的 Markdown 内容分别输出：
1. 解决方案文档（完整 Markdown）
2. POV 部署计划（完整 Markdown）
3. Azure Migrate CSV（纯 CSV 格式）

如果项目中有 `app.py` 可运行，也可指导用户通过 Streamlit UI 生成 .docx 文件。

## 格式规范

所有文档遵循以下全局格式：

- 章节标题使用 `## 一、摘要` 格式（## + 中文数字编号）
- 全文使用段落叙述，严禁项目符号列表
- 要点格式：`关键词: 正文描述`（同一行，不加粗关键词）
- 表格使用 Markdown 表格语法
- 严禁对术语缩写加括号解释（写 RAG 不写 RAG（检索增强生成））
- 内容精炼简洁，不要冗长

## 模型选择指南

根据客户业务场景选择合适模型：
- **GPT-5.5**：百万上下文旗舰推理
- **GPT-5.4 / GPT-5.4-mini**：通用推理场景
- **o4-mini / o3**：深度推理场景
- **GPT-4.1**：性价比长上下文场景

禁止使用已过时的 GPT-4o。

## 参考资源

- [完整 Prompt 模板](./references/prompts.md) — 所有文档生成的 System Prompt
- [工作流详情](./references/workflow.md) — 端到端流程和技术细节
