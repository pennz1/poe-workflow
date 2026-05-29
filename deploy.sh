#!/bin/bash
set -euo pipefail

# ======================================================================
# POE Workflow Azure 部署脚本
# 用法:
#   ./deploy.sh fast   — 仅覆盖代码文件，不触发 pip install (~10秒)
#   ./deploy.sh full   — 完整构建，触发 Oryx pip install (~90秒)
#   ./deploy.sh sync   — 仅同步源码到部署副本，不部署
#   ./deploy.sh file <相对路径>  — 单文件热修复 (~5秒)
# ======================================================================

APP_NAME="poe-workflow-auto"
RESOURCE_GROUP="newapi-rg"
SRC="/Users/penn/Documents/POE workflow"
DEPLOY="/Users/penn/Development/poe-workflow"
ZIP_PATH="/tmp/poe-workflow-auto.zip"
KUDU_URL="https://${APP_NAME}.scm.azurewebsites.net"

# ---- 颜色输出 ----
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

info()  { echo -e "${GREEN}[INFO]${NC} $*"; }
warn()  { echo -e "${YELLOW}[WARN]${NC} $*"; }
error() { echo -e "${RED}[ERROR]${NC} $*" >&2; exit 1; }

# ---- 同步源码 ----
sync_source() {
    info "同步源码: $SRC → $DEPLOY"
    mkdir -p "$DEPLOY"

    rsync -a --delete \
        --exclude '__pycache__/' --exclude '*.pyc' --exclude '.DS_Store' \
        --exclude '.venv/' --exclude '.codegraph/' --exclude '*.xlsx' \
        --exclude '*.svg' --exclude 'errordoc/' \
        "$SRC/poeworkflow-refactor/" "$DEPLOY/poeworkflow-refactor/"

    rsync -a --delete \
        --exclude '__pycache__/' --exclude '*.pyc' --exclude '.DS_Store' \
        "$SRC/frontend/" "$DEPLOY/frontend/"

    rsync -a --delete "$SRC/templates/" "$DEPLOY/templates/"
    cp "$SRC/requirements.txt" "$SRC/Azurecsvtemplate.csv" "$SRC/startup.sh" "$DEPLOY/"

    info "同步完成 ✓"
}

# ---- 打包 ----
pack_zip() {
    info "打包: $ZIP_PATH"
    cd "$DEPLOY"
    rm -f "$ZIP_PATH"
    zip -r "$ZIP_PATH" \
        poeworkflow-refactor frontend templates \
        requirements.txt Azurecsvtemplate.csv startup.sh \
        -x '*/__pycache__/*' '*.pyc' '*.DS_Store' '.git/*' '.venv/*' > /dev/null
    local size
    size=$(du -h "$ZIP_PATH" | cut -f1)
    info "打包完成: $size"
}

# ---- 获取 Kudu 凭据 ----
get_kudu_creds() {
    KUDU_USER=$(az webapp deployment list-publishing-credentials \
        -g "$RESOURCE_GROUP" -n "$APP_NAME" \
        --query "publishingUserName" -o tsv)
    KUDU_PASS=$(az webapp deployment list-publishing-credentials \
        -g "$RESOURCE_GROUP" -n "$APP_NAME" \
        --query "publishingPassword" -o tsv)
    if [[ -z "$KUDU_USER" || -z "$KUDU_PASS" ]]; then
        error "无法获取 Kudu 凭据，请确认 az login 状态"
    fi
}

# ---- 快速部署 (az webapp deploy, 跳过 pip install) ----
deploy_fast() {
    info "快速部署 (az webapp deploy, 跳过 pip install)..."

    # 临时关闭构建，部署后恢复
    az webapp config appsettings set \
        -g "$RESOURCE_GROUP" -n "$APP_NAME" \
        --settings SCM_DO_BUILD_DURING_DEPLOYMENT=false \
        --output none 2>/dev/null

    az webapp deploy \
        -g "$RESOURCE_GROUP" -n "$APP_NAME" \
        --src-path "$ZIP_PATH" \
        --type zip \
        --clean false \
        --restart true \
        --output none

    # 恢复构建设置（下次 full deploy 需要）
    az webapp config appsettings set \
        -g "$RESOURCE_GROUP" -n "$APP_NAME" \
        --settings SCM_DO_BUILD_DURING_DEPLOYMENT=true \
        --output none 2>/dev/null

    info "快速部署完成 ✓"
}

# ---- 完整部署 (config-zip, 触发 Oryx build) ----
deploy_full() {
    info "完整部署 (Oryx build，触发 pip install)..."
    az webapp deployment source config-zip \
        --resource-group "$RESOURCE_GROUP" \
        --name "$APP_NAME" \
        --src "$ZIP_PATH" \
        --output none
    info "完整部署完成 ✓"
}

# ---- 单文件部署 ----
deploy_file() {
    local file_path="$1"
    if [[ ! -f "$DEPLOY/$file_path" ]]; then
        error "文件不存在: $DEPLOY/$file_path"
    fi

    info "单文件部署: $file_path"
    get_kudu_creds

    local http_code
    http_code=$(curl -s -o /dev/null -w "%{http_code}" \
        -X PUT \
        -u "$KUDU_USER:$KUDU_PASS" \
        -H "If-Match: *" \
        --data-binary @"$DEPLOY/$file_path" \
        "${KUDU_URL}/api/vfs/site/wwwroot/${file_path}")

    if [[ "$http_code" == "200" || "$http_code" == "201" || "$http_code" == "204" ]]; then
        info "文件上传成功 (HTTP ${http_code})"
    else
        error "VFS API 返回 HTTP ${http_code}, 上传失败"
    fi

    info "重启应用..."
    az webapp restart -g "$RESOURCE_GROUP" -n "$APP_NAME" --output none
    info "单文件部署完成 ✓"
}

# ---- 生产验证 ----
verify() {
    info "验证生产环境..."
    sleep 5
    local http_code
    http_code=$(curl -s -o /dev/null -w "%{http_code}" \
        --max-time 60 \
        "http://${APP_NAME}.azurewebsites.net")
    if [[ "$http_code" == "200" ]]; then
        info "生产验证通过: HTTP $http_code ✓"
    else
        warn "生产返回 HTTP $http_code，可能需要等待冷启动 (约30秒后重试)"
    fi
}

# ---- 清理 ----
cleanup() {
    rm -f "$ZIP_PATH"
    find "$DEPLOY" \( -name "__pycache__" -o -name "*.pyc" \) -exec rm -rf {} + 2>/dev/null || true
}

# ---- 主逻辑 ----
MODE="${1:-}"

case "$MODE" in
    fast)
        sync_source
        pack_zip
        deploy_fast
        verify
        cleanup
        ;;
    full)
        sync_source
        pack_zip
        deploy_full
        verify
        cleanup
        ;;
    sync)
        sync_source
        ;;
    file)
        FILE_PATH="${2:-}"
        [[ -z "$FILE_PATH" ]] && error "用法: ./deploy.sh file <相对路径>\n示例: ./deploy.sh file poeworkflow-refactor/pipeline.py"
        sync_source
        deploy_file "$FILE_PATH"
        verify
        ;;
    *)
        echo "用法: $0 {fast|full|sync|file <path>}"
        echo ""
        echo "  fast  — 快速部署：Kudu ZIP 覆盖，不触发 pip install (~10秒)"
        echo "  full  — 完整部署：触发 Oryx 构建和 pip install (~90秒)"
        echo "  sync  — 仅同步源码到部署副本，不部署"
        echo "  file  — 单文件热修复：VFS API 上传单个文件 (~5秒)"
        echo ""
        echo "何时使用 full："
        echo "  - requirements.txt 新增/变更依赖"
        echo "  - 首次部署或 Kudu API 报错后恢复"
        echo ""
        echo "何时使用 fast："
        echo "  - 代码改动不涉及新依赖（绝大多数情况）"
        exit 1
        ;;
esac

info "全部完成!"
