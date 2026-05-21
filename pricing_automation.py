"""
Azure Pricing Calculator 自动化模块
====================================
从 AI/Infra 解决方案文档的资源架构表中提取资源清单，
通过 Playwright 在 Azure Pricing Calculator 中逐一添加并配置，
校准到用户年预算后导出 xlsx。

集成点：在 run_full_auto_poe() 中 Step 2 (POV) 之后、Step 3 (Migrate) 之前调用。
"""

from __future__ import annotations

import asyncio
import logging
import os
import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# ─── 配置 ─────────────────────────────────────────────────────────────────────

CALCULATOR_URL = "https://azure.microsoft.com/en-us/pricing/calculator/"
PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".browser_profile")
DEFAULT_REGION = "East US 2"

# ─── 数据模型 ──────────────────────────────────────────────────────────────────


@dataclass
class ResourceSpec:
    """从解决方案文档中提取的单个 Azure 资源规格。"""
    service_name: str
    sku: str
    region: str
    purpose: str


@dataclass
class FallbackRecord:
    """记录资源降级/跳过的情况。"""
    original_service: str
    original_sku: str
    resolved_service: str
    resolved_sku: str
    reason: str


@dataclass
class PricingExportResult:
    """定价导出的最终结果。"""
    xlsx_bytes: Optional[bytes] = None
    monthly_total: float = 0.0
    annual_total: float = 0.0
    fallbacks: List[FallbackRecord] = field(default_factory=list)
    error: Optional[str] = None
    services_added: int = 0
    services_skipped: int = 0


# ─── SERVICE_CATALOG: 服务名 → 计算器搜索词 + SKU 降级映射 ────────────────────

SERVICE_CATALOG: Dict[str, Dict[str, Any]] = {
    "Azure OpenAI": {
        "search_term": "Azure OpenAI Service",
        "category": "AI + Machine Learning",
        "known_skus": ["GPT-4o", "GPT-4", "GPT-4o-mini", "GPT-3.5-Turbo"],
        "default_sku": "GPT-4o",
        "sku_fallback": {
            # GPT-5 系列 → 逐级降级（5.4→5.3→5.2→5.1→5→4o）
            "GPT-5.5": "GPT-5.4",
            "GPT-5.4": "GPT-5.3",
            "GPT-5.3": "GPT-5.2",
            "GPT-5.2": "GPT-5.1",
            "GPT-5.1": "GPT-5",
            "GPT-5": "GPT-4o",
            "GPT-5-turbo": "GPT-5",
            "GPT-5o": "GPT-5",
            # GPT-5 mini 系列 → 逐级降级
            "GPT-5.5-mini": "GPT-5.4-mini",
            "GPT-5.4-mini": "GPT-5.3-mini",
            "GPT-5.3-mini": "GPT-5.2-mini",
            "GPT-5.2-mini": "GPT-5.1-mini",
            "GPT-5.1-mini": "GPT-5-mini",
            "GPT-5-mini": "GPT-4o-mini",
            "GPT-5o-mini": "GPT-5-mini",
            # GPT-4.1 系列 → 同级 GPT-4o 系列
            "GPT-4.1": "GPT-4o",
            "GPT-4.1-mini": "GPT-4o-mini",
            "GPT-4.1-nano": "GPT-4o-mini",
            # o 系列推理模型 → 同级 GPT-4o
            "o4-mini": "GPT-4o-mini",
            "o3": "GPT-4o",
            "o3-mini": "GPT-4o-mini",
            "o1": "GPT-4o",
            "o1-mini": "GPT-4o-mini",
            "o1-preview": "GPT-4o",
        },
        # 权重系数：预算校准时此服务的价格权重（AI类权重高）
        "budget_weight": 5,
    },
    "Azure AI Search": {
        "search_term": "Azure AI Search",
        "category": "AI + Machine Learning",
        "known_skus": ["Free", "Basic", "Standard", "Standard S2", "Standard S3", "Storage Optimized"],
        "default_sku": "Standard",
        "sku_fallback": {
            "S0": "Free",
            "S1": "Standard",
            "S2": "Standard S2",
            "S3": "Standard S3",
        },
        "budget_weight": 3,
    },
    "Azure AI Speech": {
        "search_term": "Azure AI Speech",
        "category": "AI + Machine Learning",
        "known_skus": ["Free", "Standard"],
        "default_sku": "Standard",
        "sku_fallback": {
            "S0": "Standard",
            "Standard (S0)": "Standard",
            "F0": "Free",
        },
    },
    "Azure AI Language": {
        "search_term": "Azure AI Language",
        "category": "AI + Machine Learning",
        "known_skus": ["Free", "Standard"],
        "default_sku": "Standard",
        "sku_fallback": {"S0": "Standard", "S": "Standard", "F0": "Free"},
    },
    "Azure AI Document Intelligence": {
        "search_term": "Azure AI Document Intelligence",
        "category": "AI + Machine Learning",
        "known_skus": ["Free", "Standard"],
        "default_sku": "Standard",
        "sku_fallback": {"S0": "Standard", "F0": "Free"},
    },
    "Azure AI Content Safety": {
        "search_term": "Azure AI Content Safety",
        "category": "AI + Machine Learning",
        "known_skus": ["Free", "Standard"],
        "default_sku": "Standard",
        "sku_fallback": {"S0": "Standard", "F0": "Free"},
    },
    "Azure API Management": {
        "search_term": "API Management",
        "category": "Integration",
        "known_skus": ["Consumption", "Developer", "Basic", "Standard", "Premium"],
        "default_sku": "Standard",
        "sku_fallback": {},
        "budget_weight": 1,
    },
    "Azure Cosmos DB": {
        "search_term": "Azure Cosmos DB",
        "category": "Databases",
        "known_skus": ["Serverless", "Provisioned throughput", "Autoscale"],
        "default_sku": "Serverless",
        "sku_fallback": {
            "Standard": "Provisioned throughput",
        },
    },
    "Azure SQL Database": {
        "search_term": "Azure SQL Database",
        "category": "Databases",
        "known_skus": ["Basic", "Standard", "Premium", "General Purpose", "Business Critical", "Hyperscale"],
        "default_sku": "General Purpose",
        "sku_fallback": {
            "S0": "Standard",
            "S1": "Standard",
            "P1": "Premium",
        },
    },
    "Azure Kubernetes Service": {
        "search_term": "Azure Kubernetes Service (AKS)",
        "category": "Compute",
        "known_skus": ["Free", "Standard", "Premium"],
        "default_sku": "Standard",
        "sku_fallback": {},
    },
    "Virtual Machines": {
        "search_term": "Virtual Machines",
        "category": "Compute",
        "known_skus": ["D2s_v5", "D4s_v5", "D8s_v5", "E4s_v5", "E8s_v5", "B2s", "B4ms"],
        "default_sku": "D4s_v5",
        "sku_fallback": {
            "D2_v3": "D2s_v5",
            "D4_v3": "D4s_v5",
            "D2s_v3": "D2s_v5",
            "D4s_v3": "D4s_v5",
            "E4s_v3": "E4s_v5",
            "E8s_v3": "E8s_v5",
        },
    },
    "Azure App Service": {
        "search_term": "App Service",
        "category": "Compute",
        "known_skus": ["Free", "Basic", "Standard", "Premium v3"],
        "default_sku": "Standard",
        "sku_fallback": {
            "B1": "Basic",
            "S1": "Standard",
            "P1v3": "Premium v3",
            "P1v2": "Premium v3",
        },
    },
    "Azure Functions": {
        "search_term": "Azure Functions",
        "category": "Compute",
        "known_skus": ["Consumption", "Premium", "Dedicated"],
        "default_sku": "Consumption",
        "sku_fallback": {},
    },
    "Azure Storage": {
        "search_term": "Storage Accounts",
        "category": "Storage",
        "known_skus": ["Standard LRS", "Standard GRS", "Premium LRS"],
        "default_sku": "Standard LRS",
        "sku_fallback": {
            "Hot": "Standard LRS",
            "Cool": "Standard LRS",
            "Archive": "Standard LRS",
        },
        "budget_weight": 1,
    },
    "Azure Key Vault": {
        "search_term": "Key Vault",
        "category": "Security",
        "known_skus": ["Standard", "Premium"],
        "default_sku": "Standard",
        "sku_fallback": {},
    },
    "Azure Monitor": {
        "search_term": "Azure Monitor",
        "category": "Management and Governance",
        "known_skus": ["Log Analytics"],
        "default_sku": "Log Analytics",
        "sku_fallback": {},
    },
    "Azure Virtual Network": {
        "search_term": "Virtual Network",
        "category": "Networking",
        "known_skus": ["Standard"],
        "default_sku": "Standard",
        "sku_fallback": {},
    },
    "Azure Application Gateway": {
        "search_term": "Application Gateway",
        "category": "Networking",
        "known_skus": ["Standard v2", "WAF v2"],
        "default_sku": "Standard v2",
        "sku_fallback": {},
    },
    "Azure Front Door": {
        "search_term": "Azure Front Door",
        "category": "Networking",
        "known_skus": ["Standard", "Premium"],
        "default_sku": "Standard",
        "sku_fallback": {},
    },
    "Azure Container Apps": {
        "search_term": "Azure Container Apps",
        "category": "Compute",
        "known_skus": ["Consumption", "Dedicated"],
        "default_sku": "Consumption",
        "sku_fallback": {},
    },
    "Azure Redis Cache": {
        "search_term": "Azure Cache for Redis",
        "category": "Databases",
        "known_skus": ["Basic", "Standard", "Premium", "Enterprise"],
        "default_sku": "Standard",
        "sku_fallback": {
            "C0": "Basic",
            "C1": "Standard",
            "P1": "Premium",
        },
    },
    "Azure Service Bus": {
        "search_term": "Service Bus",
        "category": "Integration",
        "known_skus": ["Basic", "Standard", "Premium"],
        "default_sku": "Standard",
        "sku_fallback": {},
    },
    "Azure Event Hub": {
        "search_term": "Event Hubs",
        "category": "Analytics",
        "known_skus": ["Basic", "Standard", "Premium", "Dedicated"],
        "default_sku": "Standard",
        "sku_fallback": {},
    },
    "Azure Cognitive Services": {
        "search_term": "Azure AI services multi-service account",
        "category": "AI + Machine Learning",
        "known_skus": ["Standard"],
        "default_sku": "Standard",
        "sku_fallback": {"S0": "Standard"},
    },
    "Azure Backup": {
        "search_term": "Azure Backup",
        "category": "Management and Governance",
        "known_skus": ["Standard"],
        "default_sku": "Standard",
        "sku_fallback": {},
    },
    "Azure Bot Service": {
        "search_term": "Azure Bot Service",
        "category": "AI + Machine Learning",
        "known_skus": ["Free", "Standard"],
        "default_sku": "Standard",
        "sku_fallback": {"S1": "Standard", "F0": "Free"},
    },
    "Azure AI Translator": {
        "search_term": "Translator",
        "category": "AI + Machine Learning",
        "known_skus": ["Free", "Standard"],
        "default_sku": "Standard",
        "sku_fallback": {"S0": "Standard", "S1": "Standard"},
    },
    "Azure AI Vision": {
        "search_term": "Azure AI Vision",
        "category": "AI + Machine Learning",
        "known_skus": ["Free", "Standard"],
        "default_sku": "Standard",
        "sku_fallback": {"S0": "Standard", "S1": "Standard"},
    },
}

# ─── 服务名规范化映射 ──────────────────────────────────────────────────────────

_SERVICE_NAME_ALIASES: Dict[str, str] = {
    # AI services
    "azure openai": "Azure OpenAI",
    "azure openai service": "Azure OpenAI",
    "openai": "Azure OpenAI",
    "azure ai search": "Azure AI Search",
    "cognitive search": "Azure AI Search",
    "azure cognitive search": "Azure AI Search",
    "azure ai speech": "Azure AI Speech",
    "speech services": "Azure AI Speech",
    "azure speech": "Azure AI Speech",
    "azure ai language": "Azure AI Language",
    "language service": "Azure AI Language",
    "azure ai document intelligence": "Azure AI Document Intelligence",
    "form recognizer": "Azure AI Document Intelligence",
    "azure form recognizer": "Azure AI Document Intelligence",
    "document intelligence": "Azure AI Document Intelligence",
    "azure ai content safety": "Azure AI Content Safety",
    "content safety": "Azure AI Content Safety",
    "azure ai translator": "Azure AI Translator",
    "translator": "Azure AI Translator",
    "azure ai vision": "Azure AI Vision",
    "computer vision": "Azure AI Vision",
    "azure cognitive services": "Azure Cognitive Services",
    "cognitive services": "Azure Cognitive Services",
    "azure bot service": "Azure Bot Service",
    "bot service": "Azure Bot Service",
    # Compute
    "virtual machines": "Virtual Machines",
    "azure virtual machines": "Virtual Machines",
    "vm": "Virtual Machines",
    "azure vm": "Virtual Machines",
    "azure kubernetes service": "Azure Kubernetes Service",
    "aks": "Azure Kubernetes Service",
    "azure app service": "Azure App Service",
    "app service": "Azure App Service",
    "azure functions": "Azure Functions",
    "functions": "Azure Functions",
    "azure container apps": "Azure Container Apps",
    "container apps": "Azure Container Apps",
    # Databases
    "azure cosmos db": "Azure Cosmos DB",
    "cosmosdb": "Azure Cosmos DB",
    "cosmos db": "Azure Cosmos DB",
    "azure sql database": "Azure SQL Database",
    "azure sql": "Azure SQL Database",
    "sql database": "Azure SQL Database",
    "azure cache for redis": "Azure Redis Cache",
    "azure redis cache": "Azure Redis Cache",
    "redis cache": "Azure Redis Cache",
    "redis": "Azure Redis Cache",
    # Integration
    "azure api management": "Azure API Management",
    "api management": "Azure API Management",
    "apim": "Azure API Management",
    "azure service bus": "Azure Service Bus",
    "service bus": "Azure Service Bus",
    "azure event hub": "Azure Event Hub",
    "azure event hubs": "Azure Event Hub",
    "event hubs": "Azure Event Hub",
    "event hub": "Azure Event Hub",
    # Storage
    "azure storage": "Azure Storage",
    "storage accounts": "Azure Storage",
    "storage account": "Azure Storage",
    "blob storage": "Azure Storage",
    "azure blob storage": "Azure Storage",
    # Networking
    "azure virtual network": "Azure Virtual Network",
    "virtual network": "Azure Virtual Network",
    "vnet": "Azure Virtual Network",
    "azure application gateway": "Azure Application Gateway",
    "application gateway": "Azure Application Gateway",
    "azure front door": "Azure Front Door",
    "front door": "Azure Front Door",
    # Security & Management
    "azure key vault": "Azure Key Vault",
    "key vault": "Azure Key Vault",
    "azure monitor": "Azure Monitor",
    "monitor": "Azure Monitor",
    "log analytics": "Azure Monitor",
    "azure backup": "Azure Backup",
    "backup": "Azure Backup",
}

# ─── Phase 1: 资源表解析 ──────────────────────────────────────────────────────


def extract_resource_table(solution_text: str) -> List[ResourceSpec]:
    """
    从解决方案 Markdown 文本中提取 Chapter 8 资源架构表。
    查找包含 "服务名称" 列头的表格并解析。
    """
    lines = solution_text.split("\n")
    resources: List[ResourceSpec] = []

    # 定位资源表区域：找到"资源架构"章节或含"服务名称"的表头行
    table_start = -1
    for i, line in enumerate(lines):
        if "服务名称" in line and "|" in line and "配置规格" in line:
            table_start = i
            break
        # 备选：找到章节标题后的第一个表格
        if re.search(r"(资源架构|Azure\s*资源需求)", line):
            # 从这里开始向下找表格
            for j in range(i + 1, min(i + 10, len(lines))):
                if "|" in lines[j] and "服务名称" in lines[j]:
                    table_start = j
                    break
            if table_start >= 0:
                break

    if table_start < 0:
        logger.warning("未在解决方案文档中找到资源架构表")
        return resources

    # 收集连续的表格行
    table_lines: List[str] = []
    for i in range(table_start, len(lines)):
        line = lines[i].strip()
        if not line:
            # 空行可能是表格结束
            if table_lines:
                break
            continue
        if "|" in line:
            table_lines.append(line)
        elif table_lines:
            # 非表格行且已经有内容 → 结束
            break

    if len(table_lines) < 3:  # header + separator + at least 1 data row
        return resources

    # 解析表格（复用 _parse_markdown_table 的逻辑）
    rows = _parse_table_lines(table_lines)
    if not rows or len(rows) < 2:
        return resources

    # 第一行是表头，找到列索引
    header = [h.lower().strip() for h in rows[0]]
    name_idx = _find_col_index(header, ["服务名称", "service name", "服务"])
    sku_idx = _find_col_index(header, ["配置规格", "sku", "规格", "配置规格 (sku)"])
    region_idx = _find_col_index(header, ["区域", "region", "地区"])
    purpose_idx = _find_col_index(header, ["核心用途", "purpose", "用途"])

    if name_idx < 0:
        logger.warning("资源表缺少'服务名称'列")
        return resources

    # 解析数据行
    for row in rows[1:]:
        if len(row) <= name_idx:
            continue
        service_name = row[name_idx].strip()
        if not service_name or service_name == "---":
            continue
        sku = row[sku_idx].strip() if sku_idx >= 0 and sku_idx < len(row) else ""
        region = row[region_idx].strip() if region_idx >= 0 and region_idx < len(row) else DEFAULT_REGION
        purpose = row[purpose_idx].strip() if purpose_idx >= 0 and purpose_idx < len(row) else ""

        resources.append(ResourceSpec(
            service_name=service_name,
            sku=sku,
            region=region or DEFAULT_REGION,
            purpose=purpose,
        ))

    return resources


def _parse_table_lines(lines: List[str]) -> List[List[str]]:
    """解析 Markdown 表格行，返回二维数组。"""
    rows = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        # 跳过分隔行
        if re.match(r"^\|[\s\-:|]+\|$", stripped):
            continue
        cells = [c.strip() for c in stripped.split("|")]
        if cells and cells[0] == "":
            cells = cells[1:]
        if cells and cells[-1] == "":
            cells = cells[:-1]
        if cells:
            rows.append(cells)
    return rows


def _find_col_index(header: List[str], candidates: List[str]) -> int:
    """在表头中查找匹配的列索引。"""
    for i, h in enumerate(header):
        for c in candidates:
            if c in h:
                return i
    return -1


# ─── Phase 2: 服务名规范化与 SKU 降级 ─────────────────────────────────────────


def normalize_service_name(raw_name: str) -> str:
    """将文档中的服务名规范化为 SERVICE_CATALOG 中的标准名称。"""
    # 直接匹配
    if raw_name in SERVICE_CATALOG:
        return raw_name

    # 小写查 alias
    key = raw_name.lower().strip()
    if key in _SERVICE_NAME_ALIASES:
        return _SERVICE_NAME_ALIASES[key]

    # 去掉括号内容再匹配: "Azure AI Speech (S0)" → "Azure AI Speech"
    clean = re.sub(r"\s*\([^)]*\)\s*", "", raw_name).strip()
    if clean in SERVICE_CATALOG:
        return clean
    if clean.lower() in _SERVICE_NAME_ALIASES:
        return _SERVICE_NAME_ALIASES[clean.lower()]

    # 去掉 "Azure" 前缀
    no_azure = re.sub(r"^Azure\s+", "", raw_name, flags=re.IGNORECASE).strip()
    if no_azure.lower() in _SERVICE_NAME_ALIASES:
        return _SERVICE_NAME_ALIASES[no_azure.lower()]

    # 无法匹配 → 返回原名（后续 add_service 会用搜索词直接尝试）
    return raw_name


def resolve_sku(service_name: str, requested_sku: str) -> Tuple[str, Optional[str]]:
    """
    解析 SKU，返回 (resolved_sku, fallback_reason)。
    fallback_reason 为 None 表示精确匹配成功。
    """
    catalog = SERVICE_CATALOG.get(service_name)
    if not catalog:
        return requested_sku, None  # 不在目录中，直接用原始值

    known = catalog.get("known_skus", [])
    fallback_map = catalog.get("sku_fallback", {})
    default = catalog.get("default_sku", "")

    # 精确匹配
    if requested_sku in known:
        return requested_sku, None

    # SKU fallback 映射（支持链式降级：GPT-5.4→5.3→5.2→…→4o）
    if requested_sku in fallback_map:
        current = requested_sku
        for _ in range(10):  # 最多追踪 10 跳
            next_sku = fallback_map.get(current)
            if not next_sku:
                break
            if next_sku in known:
                return next_sku, f"{requested_sku} 不可用，已降级为 {next_sku}"
            current = next_sku
        # 链路耗尽未命中 known_skus，使用链末端值
        return current, f"{requested_sku} 不可用，已降级为 {current}"

    # 去掉括号再试: "Standard (S0)" → "Standard"
    clean_sku = re.sub(r"\s*\([^)]*\)\s*", "", requested_sku).strip()
    if clean_sku in known:
        return clean_sku, None
    if clean_sku in fallback_map:
        current = clean_sku
        for _ in range(10):
            next_sku = fallback_map.get(current)
            if not next_sku:
                break
            if next_sku in known:
                return next_sku, f"{requested_sku} 不可用，已降级为 {next_sku}"
            current = next_sku
        return current, f"{requested_sku} 不可用，已降级为 {current}"

    # 模糊前缀匹配（取版本号前的基础名）
    base = re.split(r"[-_\s]", requested_sku)[0] if requested_sku else ""
    for k, v in fallback_map.items():
        if k.startswith(base) or base.startswith(k.split("-")[0]):
            return v, f"{requested_sku} 不可用，已降级为 {v}"

    # 全部失败 → 用 default
    if default:
        return default, f"{requested_sku} 不可用，已使用默认 {default}"
    return requested_sku, None


def resolve_region(requested_region: str) -> str:
    """规范化区域名称，返回计算器可用的区域字符串。"""
    if not requested_region:
        return DEFAULT_REGION

    # 常见区域规范映射
    region_map = {
        "east us": "East US",
        "east us 2": "East US 2",
        "west us": "West US",
        "west us 2": "West US 2",
        "west us 3": "West US 3",
        "central us": "Central US",
        "west europe": "West Europe",
        "north europe": "North Europe",
        "southeast asia": "Southeast Asia",
        "east asia": "East Asia",
        "japan east": "Japan East",
        "japan west": "Japan West",
        "australia east": "Australia East",
        "uk south": "UK South",
        "korea central": "Korea Central",
        "canada central": "Canada Central",
    }

    key = requested_region.lower().strip()
    if key in region_map:
        return region_map[key]

    # 已经是标准格式
    for v in region_map.values():
        if v.lower() == key:
            return v

    # 无法识别 → 退回默认
    return DEFAULT_REGION


# ─── Phase 3: Playwright 自动化引擎 ───────────────────────────────────────────


class PricingCalculatorAutomation:
    """Playwright 驱动的 Azure Pricing Calculator 自动化类。"""

    def __init__(
        self,
        headed: bool = False,
        profile_dir: str = PROFILE_DIR,
        timeout_ms: int = 60000,
        progress: Optional[Callable[[str], None]] = None,
    ):
        self.headed = headed
        self.profile_dir = profile_dir
        self.timeout_ms = timeout_ms
        self._progress = progress or (lambda msg: None)
        self._browser = None
        self._context = None
        self._page = None
        self._playwright = None
        self._download_dir = tempfile.mkdtemp(prefix="pricing_export_")
        self.fallbacks: List[FallbackRecord] = []

    def _log(self, msg: str):
        logger.info(msg)
        self._progress(msg)

    async def open_calculator(self, use_profile: bool = False) -> bool:
        """打开计算器页面并处理初始化。
        
        use_profile=False（默认）使用全新的浏览器上下文，不需要清空已有估算。
        use_profile=True 使用持久化 profile（保留登录态）。
        """
        from playwright.async_api import async_playwright

        self._playwright = await async_playwright().start()

        if use_profile and self.profile_dir:
            os.makedirs(self.profile_dir, exist_ok=True)
            self._context = await self._playwright.chromium.launch_persistent_context(
                user_data_dir=self.profile_dir,
                headless=not self.headed,
                accept_downloads=True,
                viewport={"width": 1440, "height": 900},
                locale="en-US",
            )
            self._page = self._context.pages[0] if self._context.pages else await self._context.new_page()
        else:
            # 全新上下文 — 空白估算，无需 clear_estimate
            self._browser = await self._playwright.chromium.launch(
                headless=not self.headed,
            )
            self._context = await self._browser.new_context(
                accept_downloads=True,
                viewport={"width": 1440, "height": 900},
                locale="en-US",
            )
            self._page = await self._context.new_page()

        browser_mode = "可见浏览器" if self.headed else "后台浏览器"
        self._log(f"已启动{browser_mode}，正在打开 Azure Pricing Calculator...")
        await self._page.goto(CALCULATOR_URL, wait_until="domcontentloaded", timeout=self.timeout_ms)
        # 等待关键 UI 元素出现（而非 networkidle，因为页面有持续的遥测请求）
        try:
            await self._page.wait_for_selector(
                "input[placeholder='Search products']",
                state="visible",
                timeout=30000,
            )
        except Exception:
            # 备选：等待一段固定时间
            await self._page.wait_for_timeout(5000)

        # 处理 cookie consent
        try:
            cookie_btn = self._page.locator("button:has-text('Accept')")
            if await cookie_btn.count() > 0:
                await cookie_btn.first.click()
                await self._page.wait_for_timeout(500)
        except Exception:
            pass

        # 检查登录态
        try:
            login_btn = self._page.locator("text=Log in")
            if await login_btn.count() > 0:
                self._log("⚠️ 未检测到登录态，使用公开定价")
            else:
                self._log("已检测到登录态")
        except Exception:
            pass

        # 关闭聊天小组件（可能遮挡按钮）
        try:
            no_thanks = self._page.locator("button:has-text('No thanks')")
            if await no_thanks.count() > 0 and await no_thanks.first.is_visible():
                await no_thanks.first.click()
                await self._page.wait_for_timeout(500)
        except Exception:
            pass

        return True

    async def add_service(self, search_term: str) -> bool:
        """通过搜索添加一个服务到估算。返回是否成功。"""
        page = self._page
        if not page:
            return False

        # 确保在 Products 标签页
        try:
            products_tab = page.locator("[role='tab']:has-text('Products')")
            if await products_tab.count() > 0:
                await products_tab.first.click()
                await page.wait_for_timeout(1000)
        except Exception:
            pass

        # 使用计算器的产品搜索框 (placeholder="Search products")
        search_input = page.locator("input[placeholder='Search products']")
        if await search_input.count() == 0:
            # 备选：aria-label
            search_input = page.locator("input[aria-label='Search products']")
        
        if await search_input.count() == 0 or not await search_input.first.is_visible():
            logger.debug(f"未找到可见的产品搜索框")
            return False

        # 搜索并添加
        result = await self._search_and_add(page, search_input.first, search_term)
        if result:
            return True

        # 简化搜索词重试
        simplified = re.sub(r"^Azure\s+", "", search_term).strip()
        if simplified != search_term:
            result = await self._search_and_add(page, search_input.first, simplified)
            if result:
                return True

        # 再简化
        simplified2 = re.sub(r"\s+(Service|Services)$", "", simplified).strip()
        if simplified2 != simplified:
            result = await self._search_and_add(page, search_input.first, simplified2)
            if result:
                return True

        return False

    async def _search_and_add(self, page, search_input, term: str) -> bool:
        """在搜索框中输入关键词，找到并点击 Add to estimate。"""
        # 清空并输入
        await search_input.click()
        await page.wait_for_timeout(200)
        await search_input.fill("")
        await page.wait_for_timeout(300)
        await search_input.fill(term)
        await page.wait_for_timeout(2000)

        # 找到 "Add to estimate" 按钮
        # 搜索结果中每个产品卡片都有一个 "Add to estimate" 按钮
        add_btn = page.locator("button:has-text('Add to estimate')")
        count = await add_btn.count()
        
        if count == 0:
            return False

        # 尝试点击第一个可见的
        for i in range(count):
            btn = add_btn.nth(i)
            try:
                if await btn.is_visible(timeout=2000):
                    await btn.scroll_into_view_if_needed(timeout=3000)
                    await page.wait_for_timeout(300)
                    await btn.click(timeout=5000)
                    await page.wait_for_timeout(2000)
                    # 清空搜索框（恢复产品列表）
                    try:
                        await search_input.fill("")
                        await page.wait_for_timeout(500)
                    except Exception:
                        pass
                    return True
            except Exception:
                continue

        # 按钮存在但不可见 — 可能需要先点击产品卡片展开
        # 尝试点击包含搜索词文本的卡片区域
        try:
            # 找到产品名称链接/标题
            product_link = page.locator(f"a:has-text('{term}'), h3:has-text('{term}'), [class*='product-name']:has-text('{term}')")
            if await product_link.count() > 0:
                await product_link.first.click()
                await page.wait_for_timeout(1000)
                # 再次尝试 Add to estimate
                if await add_btn.count() > 0:
                    for i in range(await add_btn.count()):
                        btn = add_btn.nth(i)
                        try:
                            if await btn.is_visible(timeout=2000):
                                await btn.click(timeout=5000)
                                await page.wait_for_timeout(2000)
                                try:
                                    await search_input.fill("")
                                except Exception:
                                    pass
                                return True
                        except Exception:
                            continue
        except Exception:
            pass

        # 最后尝试 force click
        try:
            await add_btn.first.click(force=True, timeout=5000)
            await page.wait_for_timeout(2000)
            try:
                await search_input.fill("")
            except Exception:
                pass
            return True
        except Exception:
            pass

        return False

    async def configure_service_region(self, region: str) -> bool:
        """
        为最近添加的服务配置区域。
        非阻塞：如果区域不可选则跳过（使用默认区域），不影响后续流程。
        """
        page = self._page
        if not page:
            return False

        try:
            # 找到 region dropdown（通常是最后一个配置面板中的）
            region_selectors = [
                "select[aria-label*='Region']",
                "select[aria-label*='region']",
                "[aria-label*='Region'] select",
                "select:has(option:has-text('East US'))",
            ]

            for selector in region_selectors:
                el = page.locator(selector)
                if await el.count() > 0:
                    last_el = el.last
                    # 尝试多种区域名称格式（Calculator 可能用不同命名）
                    region_variants = [
                        region,                          # "East Asia"
                        f"(Asia Pacific) {region}",     # "(Asia Pacific) East Asia"
                        region.replace(" ", ""),          # "EastAsia"
                    ]
                    for variant in region_variants:
                        try:
                            await last_el.select_option(label=variant, timeout=3000)
                            await page.wait_for_timeout(500)
                            return True
                        except Exception:
                            continue
                    # 所有变体都失败 → 尝试用 value 属性（一些 select 用 value 而非 label）
                    try:
                        value_key = region.lower().replace(" ", "")
                        await last_el.select_option(value=value_key, timeout=3000)
                        await page.wait_for_timeout(500)
                        return True
                    except Exception:
                        pass
                    # 最终兜底：选择 East US 2（确保不会因区域问题阻塞）
                    try:
                        await last_el.select_option(label="East US 2", timeout=3000)
                        await page.wait_for_timeout(500)
                        self._log(f"  ⚠️ 区域 {region} 不可用，已使用 East US 2")
                        return True
                    except Exception:
                        pass
                    break  # 找到了 select 但选不了 → 不再尝试其他 selector

        except Exception as e:
            logger.debug(f"配置区域失败: {e}")

        return False

    async def get_current_total(self) -> float:
        """读取当前估算面板的月度总价（从页面顶栏 'Estimated monthly cost' 旁读取）。"""
        page = self._page
        if not page:
            return 0.0

        try:
            # 方式1：直接从顶栏文本中提取 — 页面顶部有 "$X.XX / Estimated monthly cost"
            header_text = await page.evaluate("""() => {
                // 查找包含 "Estimated monthly cost" 的元素的前一个兄弟元素
                const all = document.querySelectorAll('*');
                for (const el of all) {
                    const text = (el.textContent || '').trim();
                    if (text.includes('Estimated monthly cost') && text.includes('$')) {
                        const match = text.match(/\\$([\\ \\d,]+\\.\\d{2})\\s*\\/\\s*Estimated monthly cost/);
                        if (match) return match[1].replace(/[, ]/g, '');
                    }
                }
                return null;
            }""")

            if header_text:
                return float(header_text)

            # 方式2：尝试通用价格选择器
            total_selectors = [
                "[class*='total'] [class*='price']",
                "[class*='estimate-total']",
                "[aria-label*='total']",
                "[class*='summary'] [class*='cost']",
            ]

            for selector in total_selectors:
                el = page.locator(selector)
                if await el.count() > 0:
                    text = await el.first.text_content()
                    if text:
                        match = re.search(r"\$?([\d,]+\.?\d*)", text.replace(",", ""))
                        if match:
                            return float(match.group(1))

            # 方式3：扫描所有含 $ 的元素找最大值
            all_prices = await page.evaluate("""() => {
                const results = [];
                const els = document.querySelectorAll('*');
                for (const el of els) {
                    const text = (el.textContent || '').trim();
                    if (text.match(/^\\$[\\d,]+\\.\\d{2}$/) && !text.includes('$0.00')) {
                        results.push(text.replace('$', '').replace(',', ''));
                    }
                }
                return results;
            }""")

            if all_prices:
                prices = [float(p) for p in all_prices if p]
                if prices:
                    return max(prices)

        except Exception as e:
            logger.debug(f"读取总价失败: {e}")

        return 0.0

    async def set_estimate_name(self, name: str) -> bool:
        """在 Calculator 页面中设置 Estimate 名称（导出前调用）。"""
        page = self._page
        if not page or not name:
            return False

        try:
            # 估算名称输入框 — 通常在 "Your Estimate" 区域最上方
            name_input = page.locator("input[aria-label*='name' i], input[placeholder*='name' i], input[aria-label*='Estimate' i]")
            if await name_input.count() > 0:
                await name_input.first.click()
                await page.wait_for_timeout(300)
                await name_input.first.fill("")
                await page.wait_for_timeout(200)
                await name_input.first.fill(name)
                await page.wait_for_timeout(500)
                # 点击其他区域使其保存
                await page.locator("body").click(position={"x": 10, "y": 10})
                await page.wait_for_timeout(500)
                self._log(f"  ✅ 已设置估算名称为 \"{name}\"")
                return True

            # 备选：直接在 estimate 标签页中找到可编辑名称
            editable = page.locator("[contenteditable='true']")
            if await editable.count() > 0:
                await editable.first.click()
                await page.keyboard.press("Control+a")
                await page.keyboard.type(name)
                await page.wait_for_timeout(500)
                return True

        except Exception as e:
            logger.debug(f"设置 Estimate 名称失败: {e}")

        return False

    async def configure_openai_tokens(self, input_tokens: int = 1000, output_tokens: int = 500) -> bool:
        """为最近添加的 Azure OpenAI 服务配置 token 数量（使描述不为 0）。"""
        page = self._page
        if not page:
            return False

        try:
            # Azure OpenAI 配置面板中有 input/output token 的数字输入框
            # 典型 aria-label: "Input tokens" / "Output tokens" 或类似
            token_inputs = page.locator(
                "input[aria-label*='input token' i], "
                "input[aria-label*='Input Tokens' i], "
                "input[aria-label*='1,000 input' i]"
            )
            if await token_inputs.count() > 0:
                await token_inputs.last.click()
                await token_inputs.last.fill(str(input_tokens))
                await page.wait_for_timeout(500)

            output_inputs = page.locator(
                "input[aria-label*='output token' i], "
                "input[aria-label*='Output Tokens' i], "
                "input[aria-label*='1,000 output' i]"
            )
            if await output_inputs.count() > 0:
                await output_inputs.last.click()
                await output_inputs.last.fill(str(output_tokens))
                await page.wait_for_timeout(500)

            # 通用方式：找到数字输入框带有 "token" 标签
            if await token_inputs.count() == 0 and await output_inputs.count() == 0:
                # 尝试找 estimate 面板中最新服务卡的数字输入
                all_inputs = page.locator(".service-card:last-child input[type='number'], [class*='estimate'] input[type='number']")
                count = await all_inputs.count()
                if count >= 2:
                    await all_inputs.nth(-2).fill(str(input_tokens))
                    await page.wait_for_timeout(300)
                    await all_inputs.nth(-1).fill(str(output_tokens))
                    await page.wait_for_timeout(500)
                    return True

            return await token_inputs.count() > 0 or await output_inputs.count() > 0

        except Exception as e:
            logger.debug(f"配置 OpenAI token 数量失败: {e}")
            return False

    async def export_estimate(self) -> Optional[bytes]:
        """点击 Export 按钮并获取下载的 xlsx 文件内容。"""
        page = self._page
        if not page:
            return None

        export_selectors = [
            "button:has-text('Export')",
            "[aria-label*='Export']",
            "a:has-text('Export')",
            "button:has-text('Download')",
            "text=Export",
        ]

        for selector in export_selectors:
            try:
                el = page.locator(selector)
                if await el.count() > 0:
                    await el.first.scroll_into_view_if_needed()
                    await page.wait_for_timeout(500)

                    async with page.expect_download(timeout=30000) as download_info:
                        await el.first.click()

                    download = await download_info.value
                    save_path = os.path.join(self._download_dir, download.suggested_filename)
                    await download.save_as(save_path)

                    with open(save_path, "rb") as f:
                        return f.read()
            except Exception:
                continue

        return None

    async def clear_estimate(self):
        """清空当前估算面板中的所有服务（通过逐个删除）。"""
        page = self._page
        if not page:
            return

        try:
            # 方式1：找到 estimate 面板中的删除按钮（每个服务卡片上的 X/trash 按钮）
            # 从 "Your Estimate" 下拉菜单中选择 "Delete estimate"
            estimate_dropdown = page.locator("button:has-text('Your Estimate'), [class*='estimate'] button[aria-haspopup]")
            if await estimate_dropdown.count() > 0:
                # 点击三点菜单
                menu_btn = page.locator("[class*='estimate'] button:has([class*='more']), button[aria-label*='more'], button:has-text('...')")
                if await menu_btn.count() > 0:
                    await menu_btn.first.click()
                    await page.wait_for_timeout(500)

            # 方式2：直接找 "Delete all" 类按钮
            delete_triggers = [
                "button:has-text('Delete all')",
                "button:has-text('Clear all')",
                "button:has-text('Remove all')",
                "[aria-label*='Delete all']",
                "[aria-label*='delete estimate']",
            ]
            for selector in delete_triggers:
                el = page.locator(selector)
                if await el.count() > 0:
                    await el.first.click()
                    await page.wait_for_timeout(500)
                    break

            # 处理确认对话框 — "Delete Estimate" 弹窗有 "Delete" 和 "Cancel" 按钮
            confirm_selectors = [
                "button:has-text('Delete'):not(:has-text('Cancel'))",
                "button:has-text('Yes')",
                "button:has-text('OK')",
                "button:has-text('Confirm')",
            ]
            await page.wait_for_timeout(500)
            for selector in confirm_selectors:
                el = page.locator(selector)
                if await el.count() > 0:
                    # 确保点击的是对话框中的确认按钮，不是其他删除按钮
                    for i in range(await el.count()):
                        btn = el.nth(i)
                        if await btn.is_visible():
                            await btn.click()
                            await page.wait_for_timeout(1000)
                            return
                            
        except Exception as e:
            logger.debug(f"清空估算失败: {e}")
            # 如果有对话框残留，尝试关掉它
            try:
                cancel = page.locator("button:has-text('Cancel')")
                if await cancel.count() > 0 and await cancel.first.is_visible():
                    await cancel.first.click()
                    await page.wait_for_timeout(500)
            except Exception:
                pass

    async def close(self):
        """关闭浏览器。"""
        if self._context:
            await self._context.close()
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()


# ─── Phase 4: 预算校准 ────────────────────────────────────────────────────────


async def _add_resources_and_calibrate(
    automation: PricingCalculatorAutomation,
    resources: List[ResourceSpec],
    annual_budget: float,
    progress: Callable[[str], None],
    budget_cap: float = 0,
) -> PricingExportResult:
    """
    核心流程：逐一添加资源 → 校准预算 → 导出。
    """
    result = PricingExportResult()

    # Step 1: 添加所有资源
    for i, res in enumerate(resources, 1):
        normalized_name = normalize_service_name(res.service_name)
        catalog = SERVICE_CATALOG.get(normalized_name)

        # 解析 SKU
        resolved_sku, fallback_reason = resolve_sku(normalized_name, res.sku)
        if fallback_reason:
            record = FallbackRecord(
                original_service=res.service_name,
                original_sku=res.sku,
                resolved_service=normalized_name,
                resolved_sku=resolved_sku,
                reason=fallback_reason,
            )
            result.fallbacks.append(record)
            automation.fallbacks.append(record)
            progress(f"  ⚠️ {res.service_name}: {fallback_reason}")

        # 确定搜索词
        search_term = catalog["search_term"] if catalog else res.service_name

        # 添加到估算
        progress(f"  正在添加 [{i}/{len(resources)}] {normalized_name}...")
        added = await automation.add_service(search_term)

        if added:
            result.services_added += 1
            # 配置区域
            region = resolve_region(res.region)
            await automation.configure_service_region(region)
            # 如果是 Azure OpenAI，配置 token 数量（避免描述显示 0）
            if "openai" in normalized_name.lower():
                await automation.configure_openai_tokens(input_tokens=1000, output_tokens=500)
            progress(f"  ✅ 已添加 {normalized_name} ({region})")
        else:
            result.services_skipped += 1
            record = FallbackRecord(
                original_service=res.service_name,
                original_sku=res.sku,
                resolved_service="(跳过)",
                resolved_sku="",
                reason=f"在计算器中未找到 {search_term}",
            )
            result.fallbacks.append(record)
            progress(f"  ⚠️ 跳过 {normalized_name}: 计算器中未找到")

    # Step 2: 读取当前总价
    await automation._page.wait_for_timeout(3000)  # 等待价格更新
    monthly_total = await automation.get_current_total()
    annual_total = monthly_total * 12
    progress(f"  当前月总价: ${monthly_total:,.2f}，年化: ${annual_total:,.2f}")

    # Step 3: 预算校准
    if annual_budget > 0 and annual_total < annual_budget:
        progress(f"  年化 ${annual_total:,.0f} < 预算 ${annual_budget:,.0f}，开始补价...")
        annual_total = await _calibrate_budget(
            automation, resources, annual_budget, annual_total, progress,
            budget_cap=budget_cap,
        )

    result.monthly_total = annual_total / 12 if annual_total > 0 else monthly_total
    result.annual_total = annual_total if annual_total > 0 else monthly_total * 12

    # Step 4: 导出
    progress("  正在导出价格估算表...")
    xlsx_bytes = await automation.export_estimate()
    if xlsx_bytes:
        result.xlsx_bytes = xlsx_bytes
        progress(f"  ✅ 成功导出价格估算表 ({len(xlsx_bytes):,} bytes)")
        progress("  ✅ 价格表已从 Azure Pricing Calculator Export 下载，未进行 Excel 改价")
        if annual_budget > 0 and result.annual_total < annual_budget:
            progress(
                f"  ⚠️ Calculator 原始年化 ${result.annual_total:,.0f} 低于预算 "
                f"${annual_budget:,.0f}；已保留网站 Export 原始文件，未改写 Excel 价格"
            )
    else:
        result.error = "Export 按钮点击失败或未触发下载"
        progress("  ❌ 导出失败")

    return result


async def _calibrate_budget(
    automation: PricingCalculatorAutomation,
    resources: List[ResourceSpec],
    annual_budget: float,
    current_annual: float,
    progress: Callable[[str], None],
    max_rounds: int = 15,
    budget_cap: float = 0,
) -> float:
    page = automation._page
    consecutive_failures = 0
    effective_cap = budget_cap if budget_cap > 0 else float("inf")
    for round_num in range(max_rounds):
        if current_annual >= annual_budget:
            return current_annual
        if current_annual >= effective_cap:
            return current_annual
        if consecutive_failures >= 3:
            break
        res = resources[round_num % len(resources)]
        normalized_name = normalize_service_name(res.service_name)
        catalog = SERVICE_CATALOG.get(normalized_name)
        if not catalog:
            consecutive_failures += 1
            continue
        search_term = catalog["search_term"]
        added = await automation.add_service(search_term)
        if added:
            region = resolve_region(res.region)
            await automation.configure_service_region(region)
            if "openai" in normalized_name.lower():
                await automation.configure_openai_tokens(input_tokens=1000, output_tokens=500)
            await page.wait_for_timeout(3000)
            monthly = await automation.get_current_total()
            new_annual = monthly * 12
            if new_annual > current_annual:
                current_annual = new_annual
                consecutive_failures = 0
            else:
                consecutive_failures += 1
        else:
            consecutive_failures += 1
    return current_annual


# ─── Phase 5: 顶层接口（供 app.py 调用） ──────────────────────────────────────


def run_pricing_export(
    solution_text: str,
    annual_budget: float,
    progress: Callable[[str], None],
    account_name: str = "",
    headed: bool = False,
    budget_cap: float = 0,
) -> PricingExportResult:
    """
    同步入口：从解决方案文本中提取资源 → 自动化 Calculator → 导出 xlsx。
    供 run_full_auto_poe() 调用。

    参数:
        solution_text: 生成的解决方案 Markdown 文本
        annual_budget: 用户年预算 (USD)
        progress: 进度回调
        account_name: 客户账户名（用于 Calculator 中 Estimate Name）
        headed: 是否显示浏览器（调试用）
        budget_cap: 预算上限（不超过下一档位，0=不限）

    返回:
        PricingExportResult 对象
    """
    # 提取资源表
    resources = extract_resource_table(solution_text)
    if not resources:
        return PricingExportResult(error="未从解决方案文档中提取到资源表")

    progress(f"从文档中提取到 {len(resources)} 个 Azure 资源")
    for r in resources:
        progress(f"  · {r.service_name} ({r.sku}) @ {r.region}")

    # 运行 async 自动化
    try:
        return asyncio.run(_run_pricing_export_async(
            resources, annual_budget, progress, account_name, headed, budget_cap
        ))
    except Exception as e:
        logger.exception("Pricing export failed")
        return PricingExportResult(error=f"浏览器自动化失败: {e}")


async def _run_pricing_export_async(
    resources: List[ResourceSpec],
    annual_budget: float,
    progress: Callable[[str], None],
    account_name: str,
    headed: bool,
    budget_cap: float = 0,
) -> PricingExportResult:
    """异步核心流程。"""
    automation = PricingCalculatorAutomation(headed=headed, progress=progress)

    try:
        # 打开计算器（使用全新上下文，空白估算）
        await automation.open_calculator()

        # 设置估算名称（在 Calculator UI 中填写，导出时自动带入 Excel 第二行）
        if account_name:
            await automation.set_estimate_name(account_name)

        # 添加资源 + 页面预算校准 + 网站导出。导出后不得改写 Excel 价格。
        result = await _add_resources_and_calibrate(automation, resources, annual_budget, progress, budget_cap)

        return result

    except Exception as e:
        return PricingExportResult(error=f"自动化过程异常: {e}")
    finally:
        await automation.close()


# ─── 辅助函数：初始化浏览器 Profile（手动登录用）────────────────────────────────


def initialize_browser_profile(progress: Optional[Callable[[str], None]] = None):
    """
    有头模式打开浏览器让用户手动登录，登录完成后关闭保存 profile。
    供 UI 中的"初始化浏览器"按钮调用。
    """
    _progress = progress or print
    asyncio.run(_initialize_browser_profile_async(_progress))


async def _initialize_browser_profile_async(progress: Callable[[str], None]):
    from playwright.async_api import async_playwright

    os.makedirs(PROFILE_DIR, exist_ok=True)
    progress("正在打开浏览器，请手动登录 Azure...")

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=False,
            accept_downloads=True,
            viewport={"width": 1440, "height": 900},
            locale="en-US",
        )
        page = context.pages[0] if context.pages else await context.new_page()
        await page.goto(CALCULATOR_URL, wait_until="domcontentloaded", timeout=60000)

        progress("浏览器已打开。请在浏览器中完成 Azure 登录。")
        progress("登录完成后，请关闭浏览器窗口或等待 3 分钟自动关闭。")

        # 等待登录完成（Log in 消失）或超时
        try:
            login_btn = page.locator("text=Log in")
            if await login_btn.count() > 0:
                await login_btn.first.click()
                await page.wait_for_selector("text=Log in", state="hidden", timeout=180000)
                progress("✅ 登录成功，浏览器 Profile 已保存")
            else:
                progress("✅ 已检测到登录态，Profile 已保存")
        except Exception:
            progress("⚠️ 登录超时，Profile 已保存当前状态")

        await context.close()


def is_browser_profile_ready() -> bool:
    """检查浏览器 profile 是否已初始化。"""
    return os.path.isdir(PROFILE_DIR) and any(
        f for f in os.listdir(PROFILE_DIR) if not f.startswith(".")
    )
