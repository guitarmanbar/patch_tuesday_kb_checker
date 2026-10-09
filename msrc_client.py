"""Клиент MSRC CVRF API v3.0."""

import time
import requests
from typing import Optional

from config import MSRC_UPDATES_URL, MSRC_CVRF_URL, MSRC_HEADERS


def fetch_updates_index() -> list[dict]:
    """Скачивает индекс всех релизов MSRC."""
    resp = requests.get(MSRC_UPDATES_URL, headers=MSRC_HEADERS, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data.get("value", [])


def fetch_cvrf_document(release_id: str) -> Optional[dict]:
    """
    Скачивает CVRF-документ за указанный месяц.
    release_id: '2024-Jan', '2025-Jun' и т.д.
    """
    url = MSRC_CVRF_URL.format(release=release_id)
    for attempt in range(3):
        try:
            resp = requests.get(url, headers=MSRC_HEADERS, timeout=90)
            resp.raise_for_status()
            return resp.json()
        except requests.RequestException as e:
            if attempt == 2:
                raise
            time.sleep(2 ** attempt)
    return None


def fetch_all_releases_since(start_year: int = 2024) -> list[str]:
    """
    Возвращает список release_id, начиная с указанного года.
    Формат: ['2024-Jan', '2024-Feb', ..., '2025-Dec']
    """
    index = fetch_updates_index()
    releases = []
    for entry in index:
        rid = entry.get("ID", "")
        if not rid:
            continue
        try:
            year = int(rid.split("-")[0])
        except (ValueError, IndexError):
            continue
        if year >= start_year:
            releases.append(rid)
    return sorted(releases)