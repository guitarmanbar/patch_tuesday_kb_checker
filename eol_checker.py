"""Проверка статуса поддержки продуктов."""

import json
from datetime import datetime
from pathlib import Path
from typing import Optional

from config import EOL_PATH


# Встроенный справочник EOL (дополняется вручную или из Microsoft Lifecycle)
DEFAULT_EOL = {
    "Windows 10|1507":  {"eol": "2017-05-09", "status": "EOL"},
    "Windows 10|1511":  {"eol": "2017-10-10", "status": "EOL"},
    "Windows 10|1607":  {"eol": "2019-04-09", "status": "EOL"},
    "Windows 10|1703":  {"eol": "2018-10-09", "status": "EOL"},
    "Windows 10|1709":  {"eol": "2019-04-09", "status": "EOL"},
    "Windows 10|1803":  {"eol": "2019-11-12", "status": "EOL"},
    "Windows 10|1809":  {"eol": "2020-11-10", "status": "EOL"},
    "Windows 10|1903":  {"eol": "2020-12-08", "status": "EOL"},
    "Windows 10|1909":  {"eol": "2021-05-11", "status": "EOL"},
    "Windows 10|2004":  {"eol": "2021-12-14", "status": "EOL"},
    "Windows 10|20H2":  {"eol": "2022-05-10", "status": "EOL"},
    "Windows 10|21H1":  {"eol": "2022-12-13", "status": "EOL"},
    "Windows 10|21H2":  {"eol": "2023-06-13", "status": "EOL"},
    "Windows 10|22H2":  {"eol": "2025-10-14", "status": "Supported"},
    "Windows 11|21H2":  {"eol": "2023-10-10", "status": "EOL"},
    "Windows 11|22H2":  {"eol": "2024-10-08", "status": "EOL"},
    "Windows 11|23H2":  {"eol": "2025-11-11", "status": "Supported"},
    "Windows 11|24H2":  {"eol": "2026-10-13", "status": "Supported"},
    "Windows Server 2016|1607": {"eol": "2027-01-12", "status": "Supported"},
    "Windows Server 2019|1809": {"eol": "2029-01-09", "status": "Supported"},
    "Windows Server 2022|21H2": {"eol": "2031-10-14", "status": "Supported"},
    "Windows Server 2025|24H2": {"eol": "2034-10-10", "status": "Supported"},
    "Office 2016":     {"eol": "2025-10-14", "status": "Supported"},
    "Office 2019":     {"eol": "2025-10-14", "status": "Supported"},
    "Office LTSC 2021": {"eol": "2026-10-13", "status": "Supported"},
    "Office LTSC 2024": {"eol": "2029-10-09", "status": "Supported"},
}


def load_eol() -> dict:
    """Загружает справочник EOL из файла или возвращает встроенный."""
    if EOL_PATH.exists():
        with EOL_PATH.open("r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_EOL


def save_eol(data: dict) -> None:
    """Сохраняет справочник EOL в файл."""
    with EOL_PATH.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def check_eol(product_name: str, version: str = "") -> Optional[dict]:
    """
    Проверяет, снят ли продукт с поддержки.
    Возвращает {'eol': 'YYYY-MM-DD', 'status': 'EOL'/'Supported'} или None.
    """
    eol_data = load_eol()

    # Пробуем точное совпадение с версией
    if version:
        key = f"{product_name}|{version}"
        if key in eol_data:
            return eol_data[key]

    # Пробуем без версии
    if product_name in eol_data:
        return eol_data[product_name]

    # Пробуем частичное совпадение
    for key, value in eol_data.items():
        if product_name.lower() in key.lower():
            return value

    return None