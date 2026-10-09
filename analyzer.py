"""Логика анализа: CVE → KB, Host → KB."""

from collections import defaultdict
from datetime import datetime
from typing import Optional

import db
from config import BUILD_MAJOR_ONLY
from eol_checker import check_eol


# ──────────────────────────────────────────────
# Режим 1: CVE → KB
# ──────────────────────────────────────────────
def analyze_cves(cve_list: list[str]) -> list[dict]:
    """
    Для каждого CVE возвращает список продуктов и KB-цепочек.

    Возвращает список словарей:
    {
        'cve': 'CVE-2024-49112',
        'product': 'Windows 10',
        'version': '22H2',
        'build': '19045',
        'kb_chain': ['KB5046613', 'KB5048652'],
        'first_fix': 'KB5046613',
        'lcu': 'KB5048652',
        'status': 'Supported' / 'EOL (...)',
        'note': '' / 'No_KB_Available'
    }
    """
    results = []

    for cve in cve_list:
        cve = cve.strip().upper()
        if not cve.startswith("CVE-"):
            cve = f"CVE-{cve}"

        records = db.get_kbs_for_cve(cve)

        if not records:
            results.append({
                "cve": cve,
                "product": "—",
                "version": "—",
                "build": "—",
                "kb_chain": "",
                "first_fix": "",
                "lcu": "",
                "status": "No_KB_Available",
                "note": "KB не найден в MSRC",
            })
            continue

        # Группируем по продукту
        by_product: dict[str, list[dict]] = defaultdict(list)
        for rec in records:
            key = f"{rec['product_name']}|{rec.get('version', '')}"
            by_product[key].append(rec)

        for product_key, recs in by_product.items():
            product_name = recs[0]["product_name"]
            version = _extract_version(product_name)
            build_major = _extract_build_major(recs)

            # Сортируем по дате
            sorted_recs = sorted(recs, key=lambda r: r.get("release_date", ""))
            kb_chain = [r["kb_number"] for r in sorted_recs if r["kb_number"]]

            # Убираем дубликаты, сохраняя порядок
            seen = set()
            kb_chain = [k for k in kb_chain if not (k in seen or seen.add(k))]

            # Определяем EOL
            eol = check_eol(product_name, version)
            if eol:
                if eol["status"] == "EOL":
                    status = f"EOL ({eol['eol']})"
                else:
                    status = "Supported"
            else:
                status = "Unknown"

            results.append({
                "cve": cve,
                "product": product_name,
                "version": version,
                "build": build_major if BUILD_MAJOR_ONLY else _extract_build_full(recs),
                "kb_chain": ", ".join(kb_chain),
                "first_fix": kb_chain[0] if kb_chain else "",
                "lcu": kb_chain[-1] if kb_chain else "",
                "status": status,
                "note": "",
            })

    return results


# ──────────────────────────────────────────────
# Режим 2: Host → рекомендуемые KB
# ──────────────────────────────────────────────
def analyze_hosts(hosts: list[dict]) -> list[dict]:
    """
    Для каждого хоста (ОС + версия + сборка) возвращает все CU от первого до LCU.

    hosts: [{'os': 'Windows Server 2022', 'version': '21H2', 'build': '20348'}]
    """
    results = []

    for host in hosts:
        os_name = host.get("os", "").strip()
        version = host.get("version", "").strip()
        build_major = host.get("build", "").strip()

        builds = db.get_builds_for_host(os_name, version, build_major)

        if not builds:
            results.append({
                "os": os_name,
                "version": version,
                "build": build_major,
                "kb": "",
                "release_date": "",
                "type": "Not_Found",
                "note": "Комбинация не найдена в базе",
            })
            continue

        # Определяем EOL
        eol = check_eol(os_name, version)
        eol_note = ""
        if eol and eol["status"] == "EOL":
            eol_note = f"EOL ({eol['eol']})"

        for i, b in enumerate(builds):
            is_lcu = (i == len(builds) - 1)
            results.append({
                "os": os_name,
                "version": version,
                "build": build_major,
                "kb": b["kb_number"],
                "release_date": b["release_date"],
                "type": "LCU" if is_lcu else "Cumulative",
                "note": eol_note if is_lcu else "",
            })

    return results


# ──────────────────────────────────────────────
# Вспомогательные функции
# ──────────────────────────────────────────────
def _extract_version(product_name: str) -> str:
    import re
    match = re.search(r"(\d{2}H\d|Version\s+\d{4}|\d{4})", product_name)
    return match.group(1).replace("Version ", "") if match else ""


def _extract_build_major(recs: list[dict]) -> str:
    """Извлекает major-часть сборки из FixedBuild."""
    for r in recs:
        fb = r.get("fixed_build", "")
        if fb:
            return fb.split(".")[0]
    return ""


def _extract_build_full(recs: list[dict]) -> str:
    """Извлекает полную сборку."""
    for r in recs:
        fb = r.get("fixed_build", "")
        if fb:
            return fb
    return ""