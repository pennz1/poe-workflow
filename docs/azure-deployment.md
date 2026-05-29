# Azure App Service 部署手册

本文档用于将本地最新 POE Workflow 代码部署到 Azure App Service。

## 固定环境

- 全量项目目录：`/Users/penn/Documents/POE workflow`
- 精简部署副本：`/Users/penn/Development/poe-workflow`
- Azure Web App：`poe-workflow-auto`
- Azure 资源组：`newapi-rg`
- 生产 URL：`http://poe-workflow-auto.azurewebsites.net`
- Python Runtime：`PYTHON|3.11`
- Startup Command：`bash startup.sh`

## 前置检查

确认 Azure CLI 已登录到正确账号：

```bash
az account show --query '{name:name,user:user.name,tenantId:tenantId}' -o json
```

确认 App Service 存在并处于 Running：

```bash
az webapp show \
  --resource-group newapi-rg \
  --name poe-workflow-auto \
  --query '{state:state,defaultHostName:defaultHostName,enabled:enabled}' \
  -o json
```

## 同步精简部署副本

部署必须从精简副本 `/Users/penn/Development/poe-workflow` 发起。该副本不包含 `.venv`、`.git`、测试文件和本地临时文件。

```bash
SRC="/Users/penn/Documents/POE workflow"
DEPLOY="/Users/penn/Development/poe-workflow"

mkdir -p "$DEPLOY"

rsync -a --delete \
  --exclude '__pycache__/' \
  --exclude '*.pyc' \
  "$SRC/poeworkflow-refactor/" \
  "$DEPLOY/poeworkflow-refactor/"

rsync -a --delete \
  --exclude '__pycache__/' \
  --exclude '*.pyc' \
  "$SRC/frontend/" \
  "$DEPLOY/frontend/"

rsync -a --delete "$SRC/templates/" "$DEPLOY/templates/"
rsync -a "$SRC/requirements.txt" "$DEPLOY/requirements.txt"
rsync -a "$SRC/Azurecsvtemplate.csv" "$DEPLOY/Azurecsvtemplate.csv"
rsync -a "$SRC/startup.sh" "$DEPLOY/startup.sh"
```

`startup.sh` 应保持如下入口：

```bash
streamlit run poeworkflow-refactor/main.py \
    --server.port "${PORT:-8000}" \
    --server.address 0.0.0.0 \
    --server.headless true \
    --browser.gatherUsageStats false
```

## 本地部署副本验证

在部署副本中做最小导入验证：

```bash
cd /Users/penn/Development/poe-workflow
PYTHONPATH="poeworkflow-refactor:." "/Users/penn/Documents/POE workflow/.venv/bin/python" - <<'PY'
import main
import pipeline
from budget.tier import budget_target_range, ensure_csv_disk_columns, load_builtin_csv_template

csv_text = ensure_csv_disk_columns(load_builtin_csv_template())
assert "Disk 1 size (In GB)" in csv_text
assert budget_target_range(20000)["target_max"] == 60000
print("deploy copy imports ok")
PY
```

## 打包

```bash
cd /Users/penn/Development/poe-workflow
rm -f /tmp/poe-workflow-auto.zip

zip -r /tmp/poe-workflow-auto.zip \
  poeworkflow-refactor \
  frontend \
  templates \
  requirements.txt \
  Azurecsvtemplate.csv \
  startup.sh \
  -x '*/__pycache__/*' '*.pyc' '*.DS_Store' '.git/*' '.venv/*'
```

## 部署

### 推荐：使用 deploy.sh 脚本（位于部署副本根目录）

```bash
cd /Users/penn/Development/poe-workflow

# 快速部署 — 仅代码变更，不触发 pip install (~10秒)
./deploy.sh fast

# 完整部署 — requirements.txt 变更时使用 (~90秒)
./deploy.sh full

# 单文件热修复 (~5秒)
./deploy.sh file poeworkflow-refactor/pipeline.py

# 仅同步源码，不部署
./deploy.sh sync
```

### 何时使用 fast vs full

| 场景 | 命令 | 耗时 |
|------|------|------|
| 改了 Python 代码 / 模板 / 前端 | `./deploy.sh fast` | ~10秒 |
| 修了 1-2 个文件想快速验证 | `./deploy.sh file <path>` | ~5秒 |
| `requirements.txt` 新增依赖 | `./deploy.sh full` | ~90秒 |
| 首次部署 / fast 部署后异常 | `./deploy.sh full` | ~90秒 |

### 手动部署（不使用脚本时）

完整构建：

```bash
az webapp deployment source config-zip \
  --resource-group newapi-rg \
  --name poe-workflow-auto \
  --src /tmp/poe-workflow-auto.zip
```

快速覆盖（Kudu ZIP API）：

```bash
CREDS=$(az webapp deployment list-publishing-credentials \
  -g newapi-rg -n poe-workflow-auto \
  --query "[publishingUserName,publishingPassword]" -o tsv)
USER=$(echo "$CREDS" | head -1); PASS=$(echo "$CREDS" | tail -1)

curl -X PUT -u "$USER:$PASS" \
  -H "Content-Type: application/zip" \
  --data-binary @/tmp/poe-workflow-auto.zip \
  "https://poe-workflow-auto.scm.azurewebsites.net/api/zip/site/wwwroot/"

az webapp restart -g newapi-rg -n poe-workflow-auto
```

如果需要确认运行配置：

```bash
az webapp config show \
  --resource-group newapi-rg \
  --name poe-workflow-auto \
  --query '{linuxFxVersion:linuxFxVersion,appCommandLine:appCommandLine}' \
  -o json
```

## 生产验证

HTTP 健康检查：

```bash
curl -I --max-time 60 http://poe-workflow-auto.azurewebsites.net
```

页面验证要点：

- 页面标题显示 `POE 自动生成工作流`
- 首页没有 OpenAI API 配置缺失提示
- Tab 列表包含 `Azure Migrate 评估`
- 在预算输入 `20K+` 后，目标区间显示 `$20,000.00 - <$60,000.00`
- Azure Migrate 评估区域显示 `磁盘字段自动补齐`

## 清理

部署完成后删除临时 zip 和 Python 缓存：

```bash
rm -f /tmp/poe-workflow-auto.zip
find /Users/penn/Development/poe-workflow \
  /Users/penn/Documents/POE\ workflow/poeworkflow-refactor \
  /Users/penn/Documents/POE\ workflow/frontend \
  \( -name "__pycache__" -o -name "*.pyc" \) \
  -print -exec rm -rf {} +
```

## 注意事项

- 不要从 `/Users/penn/Documents/POE workflow/poeworkflow-refactor` 子目录直接部署或启动生产包；生产入口依赖项目根目录结构。
- 不要把 `.venv`、`.git`、测试文件、历史导出的 Excel 或本地 session 文件打进 zip。
- 不要在日志或文档中输出 Azure OpenAI Key、App Settings 明文或任何 token。
- 如果 Azure Migrate 评估结果异常，优先使用 `Azure Migrate 评估` 独立 tab 单独重跑，不需要重跑完整 POE 流程。
- **快速部署 (Kudu API) 需要 SCM Basic Auth 已启用**。如果遇到 401 错误，运行：
  ```bash
  az resource update \
    --ids "/subscriptions/$(az account show --query id -o tsv)/resourceGroups/newapi-rg/providers/Microsoft.Web/sites/poe-workflow-auto/basicPublishingCredentialsPolicies/scm" \
    --set properties.allow=true
  ```