# 端到端工作流详情

本文件描述 POE 文档生成的完整技术流程和实现细节。

---

## 完整流程图

```
用户输入（客户名称 + 背景 + 预算 + POV 日期 + 团队）
        │
        ▼
┌─────────────────────────────┐
│  步骤 1: 生成解决方案架构文档   │
│  System: SOLUTION/INFRA_PROMPT │
│  Input: 客户名称 + 背景        │
│  Output: Markdown 方案文档     │
└─────────────────────────────┘
        │
        ▼
┌─────────────────────────────┐
│  步骤 2: 生成 POV 部署计划     │
│  System: POV_SYSTEM_PROMPT    │
│  Input: 方案文档 + 日期 + 人员  │
│  Output: Markdown POV 文档    │
└─────────────────────────────┘
        │
        ▼
┌─────────────────────────────┐
│  步骤 3: 生成 Azure Migrate CSV│
│  System: CSV_SYSTEM_PROMPT    │
│  Input: 预算 + 方案资源列表     │
│  Output: CSV 文件             │
└─────────────────────────────┘
        │
        ▼
┌─────────────────────────────┐
│  步骤 4: 打包交付              │
│  方案.docx + POV.docx + CSV   │
│  → ZIP 压缩包                 │
└─────────────────────────────┘
```

---

## 步骤 1：解决方案架构文档生成

### 输入构造

```python
user_prompt = f"""## 客户信息
- **客户名称**：{customer_name}

## 客户背景
{customer_background}
"""
```

### 文档类型选择

| 类型 | System Prompt | Word 模板 | 章节数 |
|------|------|------|------|
| AI 解决方案 | SOLUTION_SYSTEM_PROMPT | solution_template.docx.docx | 8 章 |
| Infra 基础设施 | INFRA_SYSTEM_PROMPT | Infra_template.docx | 10 章 |

### AI 方案额外步骤：SVG 架构图

当文档类型为 AI 时，额外调用 LLM 生成 SVG 架构图：
- 提取第 2、5、6、7、8 章节内容作为输入
- 生成企业级 Azure 解决方案逻辑架构图
- 嵌入到 Word 文档中

### Markdown → Word 转换规则

| Markdown 元素 | Word 格式 |
|------|------|
| `# 标题` | Heading Level 1, 微软雅黑 18pt |
| `## 章节` | Heading Level 2, 微软雅黑 14pt |
| `### 子标题` | Heading Level 3, 微软雅黑 12pt |
| 段落 | 正文, 微软雅黑 9pt, 首行缩进 0.74cm |
| `**加粗**` | Bold run |
| Markdown 表格 | Word Table Grid, 表头蓝色背景白字 |
| `关键词: 正文` | 单段落，关键词不加粗 |

### Word 模板继承

加载 .docx 模板时保留：
- 页眉页脚（品牌 Logo、公司信息）
- 默认样式和字体设置
- 页面布局（A4、页边距）

生成内容写入模板已有正文段落之后。

---

## 步骤 2：POV 部署计划生成

### 工作日计算

```python
def _workday_info(start: date, end: date):
    """计算 POV 周期内的工作日和周末。"""
    workdays = []  # 格式: "M月D日"
    weekends = []
    current = start
    while current <= end:
        if current.weekday() < 5:  # 周一到周五
            workdays.append(f"{current.month}月{current.day}日")
        else:
            weekends.append(f"{current.month}月{current.day}日")
        current += timedelta(days=1)
    return workdays, weekends
```

### 输入构造

将解决方案文档全文 + 补充信息（客户名称、POV 周期、工作日清单、禁用日期、乙方人员）组合为 User Prompt。

### 甲方人员自动生成

Prompt 中要求 AI 自动生成 2-3 名甲方人员（中文名），包含：
- 项目负责人
- 技术对接人
- （可选）业务对接人

---

## 步骤 3：Azure Migrate CSV 生成

### 预算档位匹配

| 输入预算范围 | 匹配档位 |
|------|------|
| ≤ 15,000 USD | 15k |
| 15,001 - 50,000 | 50k |
| 50,001 - 100,000 | 100k |
| 100,001 - 250,000 | 250k |
| > 250,000 | 250k |

### CSV 输出格式

必须严格遵循 Azure Migrate Import Template 格式：
- 第一行为表头（见 prompts.md 中的模板表头参考）
- 后续每行为一台服务器
- 必填列不能为空
- 非必填列可留空

### Server Name 命名规则

格式：`服务类型-区域缩写-规模-序号`

示例：
```
LLM-GPT54-EUS2-01
Search-S1-EAsia-01
Speech-S0-EUS2-01
APIM-Standard-EUS2-01
CosmosDB-Serverless-01
```

---

## 步骤 4：打包交付

### ZIP 内容

全自动模式最终生成 ZIP 包，包含：
```
{账户名}-POE-Complete.zip
├── {账户名}-Solution Architecture.docx     (或 Infra 版本)
├── {账户名}-PostAssessment POVdeployment.docx
└── {账户名}-Azure Migrate Assessment.xlsx   (如有 Azure 登录)
```

---

## 技术配置

### Azure OpenAI 配置

在 `.streamlit/secrets.toml` 中配置：
```toml
AZURE_OPENAI_KEY = "your-api-key"
AZURE_OPENAI_ENDPOINT = "https://your-gateway.example.com/"
AZURE_OPENAI_DEPLOYMENT = "gpt-5.5"
AZURE_OPENAI_API_VERSION = "2024-06-01"
```

### LLM 调用参数

| 参数 | 值 |
|------|------|
| temperature | 0.7 |
| max_tokens | 16384 |
| API 格式 | OpenAI-compatible (base_url + /v1) |

---

## 运行方式

### 通过 Streamlit Web UI

```bash
pip install -r requirements.txt
streamlit run app.py
```

### 通过 Copilot Chat（本 Skill）

在 VS Code Chat 中输入 `/poe-generation`，按提示输入客户信息，Copilot 直接生成 Markdown 格式的文档内容。
