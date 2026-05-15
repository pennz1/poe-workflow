# 项目简报：微软客户 POE 文档生成器

## 项目定位

面向微软售前团队的一站式 POE（Proof of Engagement）交付工具，通过 Azure OpenAI 驱动内容生成，一键产出完整售前文档套件。

## 技术栈

| 层次 | 技术 |
|------|------|
| 前端框架 | Streamlit (Web UI) |
| AI 引擎 | Azure OpenAI (GPT-4o / GPT-5) |
| 文档生成 | python-docx (Word)、openpyxl (Excel) |
| 数据处理 | pandas |
| 云服务交互 | Azure ARM REST API、MSAL (Device Code Flow) |
| 网络层 | requests、httpx (含 SOCKS 代理支持) |

## 项目结构

```
poe-workflow/
├── app.py                  # 主应用 (~3300 行)，包含全部业务逻辑
├── add_yearly_cost.py      # 独立工具：为价格表 Excel 添加年度成本列
├── requirements.txt        # Python 依赖
├── frontend/
│   ├── __init__.py
│   ├── ui.py               # 自定义 Streamlit UI 组件（HTML 渲染）
│   └── desktop_theme.css   # 桌面端自定义样式
├── templates/
│   ├── solution_template.docx.docx   # AI 解决方案 Word 模板
│   ├── Infra_template.docx           # 基础设施解决方案 Word 模板
│   ├── pov_template.docx.docx        # POV 部署计划 Word 模板
│   └── AzureMigrateimporttemplate.csv # Azure Migrate 导入 CSV 模板
├── docs/images/            # 文档截图
└── .streamlit/
    └── secrets.toml        # Azure OpenAI 密钥配置（不入库）
```

## 功能模块（5 个标签页）

### 1. 全自动 POE 生成
- 一键完成：方案文档 → POV 计划 → Azure Migrate 评估 → 打包 ZIP
- 需要 Azure 登录（MSAL Device Code Flow）
- 自动校准评估成本到用户预算的 100%-120% 区间

### 2. 解决方案文档
- 支持 **AI 解决方案** 和 **Infra 基础设施** 两种文档类型
- AI 生成或手动导入（上传 .docx / 粘贴文本）
- 输出：带品牌样式的 Word 文档（模板继承页眉页脚）

### 3. POV 部署计划
- 基于解决方案文档自动生成分阶段部署计划
- 输入：POV 周期、乙方项目人员
- 自动生成甲方人员、分阶段任务表（只排工作日）

### 4. Azure Migrate CSV
- 根据预算倒推客户本地服务器配置
- 生成符合 Azure Migrate 导入格式的 CSV

### 5. 年度价格表
- 为 Azure 原始价格表 Excel 添加 "Estimated yearly cost" 列

## 核心流程

```
用户输入（客户名称 + 背景 + 预算）
        ↓
   Azure OpenAI 生成 Markdown
        ↓
   Markdown → Word (.docx) 转换（python-docx）
        ↓
   全自动模式额外步骤：
     MSAL 登录 → ARM API 创建 Migrate 项目
     → 上传 CSV → 运行评估 → 下载报告 → 打包 ZIP
```

## 配置要求

在 `.streamlit/secrets.toml` 中配置：
```toml
AZURE_OPENAI_KEY = "your-api-key"
AZURE_OPENAI_ENDPOINT = "https://your-resource.openai.azure.com/"
AZURE_OPENAI_DEPLOYMENT = "your-deployment-name"
AZURE_OPENAI_API_VERSION = "2024-06-01"
```

## 启动方式

```bash
pip install -r requirements.txt
streamlit run app.py
```

浏览器自动打开 `http://localhost:8501`
