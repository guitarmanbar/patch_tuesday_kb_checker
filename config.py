"""Конфигурация системы Patch Tuesday."""

from pathlib import Path

# ──────────────────────────────────────────────
# Пути
# ──────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
DB_PATH = DATA_DIR / "msrc_cache.db"
EOL_PATH = DATA_DIR / "eol_products.json"

DATA_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ──────────────────────────────────────────────
# MSRC CVRF API v3.0
# Ключ не требуется: https://api.msrc.microsoft.com/cvrf/v3.0
# ──────────────────────────────────────────────
MSRC_CVRF_BASE = "https://api.msrc.microsoft.com/cvrf/v3.0"
MSRC_UPDATES_URL = f"{MSRC_CVRF_BASE}/updates"
MSRC_CVRF_URL = f"{MSRC_CVRF_BASE}/cvrf/{{release}}"
MSRC_HEADERS = {"Accept": "application/json"}

# ──────────────────────────────────────────────
# Продукты, включаемые в анализ
# ──────────────────────────────────────────────
TARGET_PRODUCT_GROUPS = {
    "Windows":    r"Windows\s+(10|11|Server)",
    "Office":     r"(Microsoft\s+)?(Office|Word|Excel|Outlook|PowerPoint|Access|Publisher|OneNote|Visio|Project)",
    "Exchange":   r"Exchange\s+Server",
    "SharePoint": r"SharePoint\s+Server",
    "SQL_Server": r"SQL\s+Server",
    "Visual_Studio": r"Visual\s+Studio",
    "Edge":       r"Microsoft\s+Edge",
    "DotNet":     r"\.NET\s+(Framework|Core|Runtime)",
}

# ──────────────────────────────────────────────
# Регулярные выражения
# ──────────────────────────────────────────────
import re

KB_PATTERN = re.compile(r"(?:KB)?(\d{6,7})", re.IGNORECASE)
WINDOWS_VERSION_PATTERN = re.compile(
    r"Windows\s+(10|11|Server)\s*(?:Version\s+)?([\d]{2}H\d|[\d]{4})?",
    re.IGNORECASE,
)

# ──────────────────────────────────────────────
# Формат вывода
# ──────────────────────────────────────────────
BUILD_MAJOR_ONLY = True  # False — выводить полную сборку (например, 19045.5247)