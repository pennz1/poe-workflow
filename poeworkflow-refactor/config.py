"""POE workflow configuration constants."""

import os
from typing import Optional


# Project root: parent of poeworkflow-refactor.
PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP_DIR = PROJECT_DIR

TEMPLATE_DIR = os.path.join(PROJECT_DIR, "templates")
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

BUILTIN_CSV_PATH = os.path.join(PROJECT_DIR, "Azurecsvtemplate.csv")

# App Service can set PERSIST_DIR to a persistent mount; local runs use the project root.
PERSIST_DIR = os.environ.get("PERSIST_DIR", PROJECT_DIR)
os.makedirs(PERSIST_DIR, exist_ok=True)
TIER_CACHE_PATH = os.path.join(PERSIST_DIR, ".tier_cache.json")
BUDGET_TIERS = [15_000, 50_000, 100_000, 250_000]

_PERSIST_KEYS = [
    "azure_token",
    "azure_user",
    "azure_token_expires_at",
    "azure_subscription_id",
    "azure_subscription_name",
    "azure_resource_group",
    "_cached_subscription",
    "_cached_resource_group",
    "solution_text",
    "infra_text",
    "pov_text",
    "csv_code",
    "customer_name",
    "account_name",
    "budget",
    "doc_type",
    "pov_source_doc_type",
    "pov_vendor_team",
    "auto_poe_result",
]

CN_FONT = "微软雅黑"
CN_FONT_ALT = "Microsoft YaHei UI"


def get_secret(key: str, default: Optional[str] = None) -> Optional[str]:
    """Read from environment first, then Streamlit secrets, then default."""
    val = os.environ.get(key)
    if val:
        return val
    try:
        streamlit = __import__("streamlit")
        return streamlit.secrets.get(key, default)
    except Exception:
        return default


def check_secrets() -> bool:
    """Check whether required OpenAI-compatible API credentials are configured."""
    required_keys = ["AZURE_OPENAI_KEY", "AZURE_OPENAI_ENDPOINT", "AZURE_OPENAI_DEPLOYMENT"]
    missing = [key for key in required_keys if not get_secret(key)]
    if missing:
        streamlit = __import__("streamlit")
        streamlit.error("⚠️ **OpenAI API 配置缺失**")
        streamlit.info(
            "请通过环境变量或 `.streamlit/secrets.toml` 配置以下密钥：\n\n"
            "```toml\n"
            'AZURE_OPENAI_KEY = "your-api-key"\n'
            'AZURE_OPENAI_ENDPOINT = "https://your-gateway.example.com/"\n'
            'AZURE_OPENAI_DEPLOYMENT = "gpt-4o"  # 模型名称\n'
            "```"
        )
        return False
    return True
