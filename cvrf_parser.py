"""Парсер CVRF-документов MSRC."""

import re
from typing import Any

from config import KB_PATTERN, TARGET_PRODUCT_GROUPS

# Паттерны для определения группы продукта
import re as _re
_GROUP_PATTERNS = {name: _re.compile(pattern, _re.IGNORECASE)
                   for name, pattern in TARGET_PRODUCT_GROUPS.items()}


def classify_product(product_name: str) -> str:
    """Определяет группу продукта (Windows, Office, Exchange и т.д.)."""
    for group, pattern in _GROUP_PATTERNS.items():
        if pattern.search(product_name):
            return group
    return "Other"


def parse_product_tree(doc: dict) -> dict[str, str]:
    """
    Строит карту ProductID → название продукта.
    Возвращает {product_id: product_name}.
    """
    mapping = {}
    tree = doc.get("ProductTree", {})
    for item in tree.get("FullProductName", []):
        pid = item.get("ProductID") or item.get("@ProductID")
        name = item.get("Value") or item.get("#text", "")
        if pid and name:
            mapping[str(pid)] = name.strip()
    return mapping


def extract_kb_from_remediation(rem: dict) -> str:
    """
    Извлекает номер KB из Remediation.
    Приоритет: Description (Type 2) → URL → Description (любой).
    """
    desc = rem.get("Description", {})
    desc_text = ""
    if isinstance(desc, dict):
        desc_text = desc.get("Value", "") or desc.get("#text", "")
    elif isinstance(desc, str):
        desc_text = desc

    # Приоритет 1: Description содержит "KB" или число 6-7 знаков
    match = KB_PATTERN.search(desc_text)
    if match:
        return f"KB{match.group(1)}"

    # Приоритет 2: URL
    url = rem.get("URL", "") or rem.get("@URL", "")
    if url:
        match = KB_PATTERN.search(url)
        if match:
            return f"KB{match.group(1)}"

    return ""


def parse_cvrf_document(doc: dict, release_id: str) -> dict:
    """
    Разбирает CVRF-документ.
    Возвращает словарь с ключами:
      - products:    {product_id: (name, group)}
      - cves:        [(cve_id, release_id, severity, cvss, description)]
      - cve_product_kb: [(cve_id, product_id, kb, release_date, rem_type, fixed_build)]
      - builds:      [(product_name, version, build_major, build_full, kb, date, type)]
    """
    product_map = parse_product_tree(doc)

    # Определяем дату релиза документа
    tracking = doc.get("DocumentTracking", {})
    release_date = (
        tracking.get("CurrentReleaseDate", "")[:10]
        or tracking.get("InitialReleaseDate", "")[:10]
    )

    products = {}
    for pid, name in product_map.items():
        group = classify_product(name)
        products[pid] = (name, group)

    cves = []
    cve_product_kb = []
    builds = []

    for vuln in doc.get("Vulnerability", []):
        cve_id = vuln.get("CVE", "")
        if not cve_id or not cve_id.startswith("CVE-"):
            continue

        # Severity и CVSS
        severity = ""
        cvss_score = 0.0
        for threat in vuln.get("Threats", []):
            if threat.get("Type") == 3:  # Severity
                severity = threat.get("Description", {}).get("Value", "")
                break
        for score_set in vuln.get("CVSSScoreSets", []):
            cvss_score = float(score_set.get("BaseScore", 0) or 0)
            break

        # Описание
        description = ""
        for note in vuln.get("Notes", []):
            if note.get("Type") == 2:
                description = note.get("Value", "")
                break

        cves.append((cve_id, release_id, severity, cvss_score, description[:500]))

        # Remediations
        for rem in vuln.get("Remediations", []):
            rem_type = rem.get("Type", "")
            kb = extract_kb_from_remediation(rem)
            fixed_build = rem.get("FixedBuild", "") or ""
            product_ids = rem.get("ProductID", [])
            if isinstance(product_ids, str):
                product_ids = [product_ids]

            for pid in product_ids:
                pid_str = str(pid)
                if pid_str not in product_map:
                    continue

                name = product_map[pid_str]
                group = classify_product(name)

                if kb:
                    cve_product_kb.append(
                        (cve_id, pid_str, kb, release_date, rem_type, fixed_build)
                    )

                    # Извлекаем версию и сборку из fixed_build
                    build_major, build_full = _extract_build(fixed_build)
                    version = _extract_version(name)
                    if version:
                        builds.append(
                            (name, version, build_major, build_full, kb, release_date, rem_type)
                        )

    return {
        "products": products,
        "cves": cves,
        "cve_product_kb": cve_product_kb,
        "builds": builds,
    }


def _extract_build(fixed_build: str) -> tuple[str, str]:
    """Извлекает major и full сборку из строки FixedBuild."""
    if not fixed_build:
        return "", ""
    parts = fixed_build.split(".")
    major = parts[0] if parts else ""
    full = fixed_build
    return major, full


def _extract_version(product_name: str) -> str:
    """Пытается извлечь версию (22H2, 1607, 24H2) из названия продукта."""
    match = re.search(r"(\d{2}H\d|Version\s+\d{4}|\d{4})", product_name)
    if match:
        return match.group(1).replace("Version ", "")
    return ""