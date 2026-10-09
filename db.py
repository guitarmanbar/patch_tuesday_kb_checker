"""SQLite-кэш и запросы к нему."""

import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from config import DB_PATH


def get_conn() -> sqlite3.Connection:
    """Возвращает соединение с БД с включёнными foreign keys и row_factory."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    """Создаёт схему БД, если её нет."""
    conn = get_conn()
    cur = conn.cursor()

    # Кэш CVRF-документов
    cur.execute("""
        CREATE TABLE IF NOT EXISTS cvrf_cache (
            release_id  TEXT PRIMARY KEY,
            fetched_at  TEXT NOT NULL,
            raw_json    TEXT NOT NULL
        )
    """)

    # Продукты из ProductTree
    cur.execute("""
        CREATE TABLE IF NOT EXISTS products (
            product_id    TEXT PRIMARY KEY,
            product_name  TEXT NOT NULL,
            product_group TEXT,
            is_eol        INTEGER DEFAULT 0,
            eol_date      TEXT
        )
    """)

    # CVE
    cur.execute("""
        CREATE TABLE IF NOT EXISTS cves (
            cve_id      TEXT PRIMARY KEY,
            release_id  TEXT,
            severity    TEXT,
            cvss_score  REAL,
            description TEXT
        )
    """)

    # Связь CVE ↔ Продукт ↔ KB
    cur.execute("""
        CREATE TABLE IF NOT EXISTS cve_product_kb (
            cve_id           TEXT NOT NULL,
            product_id       TEXT NOT NULL,
            kb_number        TEXT,
            release_date     TEXT,
            remediation_type TEXT,
            fixed_build      TEXT,
            PRIMARY KEY (cve_id, product_id, kb_number)
        )
    """)

    # Индекс сборок (для режима host)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS builds (
            product_name  TEXT NOT NULL,
            version       TEXT NOT NULL,
            build_major   TEXT NOT NULL,
            build_full    TEXT,
            kb_number     TEXT NOT NULL,
            release_date  TEXT,
            update_type   TEXT,
            PRIMARY KEY (product_name, version, build_major, kb_number)
        )
    """)

    # Индексы для ускорения поиска
    cur.execute("CREATE INDEX IF NOT EXISTS idx_cpk_cve ON cve_product_kb(cve_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_cpk_kb ON cve_product_kb(kb_number)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_builds_lookup ON builds(product_name, version, build_major)")

    conn.commit()
    conn.close()


# ──────────────────────────────────────────────
# Кэш CVRF
# ──────────────────────────────────────────────
def is_cached(release_id: str) -> bool:
    """Проверяет, есть ли документ в кэше."""
    conn = get_conn()
    row = conn.execute(
        "SELECT 1 FROM cvrf_cache WHERE release_id = ?", (release_id,)
    ).fetchone()
    conn.close()
    return row is not None


def save_cvrf(release_id: str, raw_json: dict) -> None:
    """Сохраняет сырой CVRF-документ в кэш."""
    conn = get_conn()
    conn.execute(
        """INSERT OR REPLACE INTO cvrf_cache (release_id, fetched_at, raw_json)
           VALUES (?, ?, ?)""",
        (release_id, datetime.utcnow().isoformat(), json.dumps(raw_json)),
    )
    conn.commit()
    conn.close()


def load_cvrf(release_id: str) -> Optional[dict]:
    """Загружает сырой CVRF-документ из кэша."""
    conn = get_conn()
    row = conn.execute(
        "SELECT raw_json FROM cvrf_cache WHERE release_id = ?", (release_id,)
    ).fetchone()
    conn.close()
    return json.loads(row["raw_json"]) if row else None


def get_cached_releases() -> list[str]:
    """Возвращает список всех release_id в кэше."""
    conn = get_conn()
    rows = conn.execute(
        "SELECT release_id FROM cvrf_cache ORDER BY release_id"
    ).fetchall()
    conn.close()
    return [r["release_id"] for r in rows]


# ──────────────────────────────────────────────
# Запись распарсенных данных
# ──────────────────────────────────────────────
def upsert_product(product_id: str, name: str, group: str = None) -> None:
    conn = get_conn()
    conn.execute(
        """INSERT OR REPLACE INTO products (product_id, product_name, product_group)
           VALUES (?, ?, ?)""",
        (product_id, name, group),
    )
    conn.commit()
    conn.close()


def upsert_cve(cve_id: str, release_id: str, severity: str, cvss: float, desc: str) -> None:
    conn = get_conn()
    conn.execute(
        """INSERT OR REPLACE INTO cves (cve_id, release_id, severity, cvss_score, description)
           VALUES (?, ?, ?, ?, ?)""",
        (cve_id, release_id, severity, cvss, desc),
    )
    conn.commit()
    conn.close()


def insert_cve_product_kb(
    cve_id: str, product_id: str, kb: str,
    release_date: str, rem_type: str, fixed_build: str
) -> None:
    conn = get_conn()
    conn.execute(
        """INSERT OR IGNORE INTO cve_product_kb
           (cve_id, product_id, kb_number, release_date, remediation_type, fixed_build)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (cve_id, product_id, kb, release_date, rem_type, fixed_build),
    )
    conn.commit()
    conn.close()


def insert_build(
    product_name: str, version: str, build_major: str, build_full: str,
    kb: str, release_date: str, update_type: str
) -> None:
    conn = get_conn()
    conn.execute(
        """INSERT OR IGNORE INTO builds
           (product_name, version, build_major, build_full, kb_number, release_date, update_type)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (product_name, version, build_major, build_full, kb, release_date, update_type),
    )
    conn.commit()
    conn.close()


# ──────────────────────────────────────────────
# Запросы для анализа
# ──────────────────────────────────────────────
def get_kbs_for_cve(cve_id: str) -> list[dict]:
    """Возвращает все записи CVE ↔ Продукт ↔ KB для указанной CVE."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT c.cve_id, p.product_name, p.product_group,
               c.kb_number, c.release_date, c.remediation_type, c.fixed_build
        FROM cve_product_kb c
        JOIN products p ON p.product_id = c.product_id
        WHERE c.cve_id = ?
        ORDER BY p.product_name, c.release_date
    """, (cve_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_builds_for_host(product_name: str, version: str, build_major: str) -> list[dict]:
    """Возвращает все KB для указанной комбинации ОС + версия + сборка."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT kb_number, release_date, update_type, build_full
        FROM builds
        WHERE product_name LIKE ? AND version = ? AND build_major = ?
        ORDER BY release_date
    """, (f"%{product_name}%", version, build_major)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_all_builds() -> list[dict]:
    """Возвращает все записи из таблицы builds (для генерации шаблона)."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT DISTINCT product_name, version, build_major
        FROM builds
        ORDER BY product_name, version, build_major
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]