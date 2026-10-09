"""Чтение и запись Excel-файлов."""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from config import OUTPUT_DIR, BUILD_MAJOR_ONLY


# ──────────────────────────────────────────────
# Стили
# ──────────────────────────────────────────────
HEADER_FONT = Font(bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
CELL_ALIGN = Alignment(vertical="top", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def _style_sheet(ws, df: pd.DataFrame) -> None:
    """Применяет единый стиль к листу Excel."""
    # Заголовки
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER

    # Ширина столбцов
    for col_idx, col_name in enumerate(df.columns, 1):
        max_len = max(
            len(str(col_name)),
            df[col_name].astype(str).str.len().max() if len(df) > 0 else 0,
        )
        ws.column_dimensions[
            ws.cell(row=1, column=col_idx).column_letter
        ].width = min(max_len + 4, 60)

    # Границы для данных
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(df.columns)):
        for cell in row:
            cell.alignment = CELL_ALIGN
            cell.border = THIN_BORDER

    # Закрепление заголовка
    ws.freeze_panes = "A2"


# ──────────────────────────────────────────────
# Запись результатов CVE-анализа
# ──────────────────────────────────────────────
def write_cve_results(results: list[dict], output_path: str = None) -> str:
    """Сохраняет результаты анализа CVE в Excel (два листа)."""
    if output_path is None:
        output_path = OUTPUT_DIR / "cve_results.xlsx"
    output_path = Path(output_path)

    df_results = pd.DataFrame(results)
    df_results.columns = [
        "CVE", "Product", "Version", "Build",
        "KB_Chain", "First_Fix", "LCU", "Status", "Note",
    ]

    # Лист с CVE, у которых нет KB
    df_no_kb = df_results[df_results["Status"] == "No_KB_Available"]
    df_main = df_results[df_results["Status"] != "No_KB_Available"]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_main.to_excel(writer, sheet_name="Results", index=False)
        if not df_no_kb.empty:
            df_no_kb.to_excel(writer, sheet_name="No_KB_Available", index=False)

        # Лист Summary
        summary = df_results.groupby("CVE").agg(
            Products=("Product", lambda x: ", ".join(sorted(set(x)))),
            KB_Count=("KB_Chain", "count"),
        ).reset_index()
        summary.to_excel(writer, sheet_name="Summary", index=False)

        # Стили
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            df = pd.read_excel(output_path, sheet_name=sheet_name)
            _style_sheet(ws, df)

    return str(output_path)


# ──────────────────────────────────────────────
# Запись результатов Host-анализа
# ──────────────────────────────────────────────
def write_host_results(results: list[dict], output_path: str = None) -> str:
    """Сохраняет результаты анализа хостов в Excel."""
    if output_path is None:
        output_path = OUTPUT_DIR / "host_results.xlsx"
    output_path = Path(output_path)

    df = pd.DataFrame(results)
    df.columns = ["OS", "Version", "Build", "KB", "Release_Date", "Type", "Note"]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Recommendations", index=False)
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            _style_sheet(ws, df)

    return str(output_path)


# ──────────────────────────────────────────────
# Генерация шаблона
# ──────────────────────────────────────────────
def generate_template(output_path: str = None) -> str:
    """
    Генерирует предзаполненный шаблон Excel для режима host.
    Заполнен комбинациями OS + Version + Build из кэша.
    """
    if output_path is None:
        output_path = OUTPUT_DIR / "host_template.xlsx"
    output_path = Path(output_path)

    all_builds = db.get_all_builds()

    rows = []
    for b in all_builds:
        rows.append({
            "OS": b["product_name"],
            "Version": b["version"],
            "Build": b["build_major"],
        })

    if not rows:
        rows = [{"OS": "", "Version": "", "Build": ""}]

    df = pd.DataFrame(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Hosts"

    # Заголовки
    headers = ["OS", "Version", "Build"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER

    # Данные
    for row_idx, row in enumerate(rows, 2):
        for col_idx, key in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row.get(key, ""))
            cell.alignment = CELL_ALIGN
            cell.border = THIN_BORDER

    # Ширина
    for col_idx, h in enumerate(headers, 1):
        ws.column_dimensions[
            ws.cell(row=1, column=col_idx).column_letter
        ].width = 30

    ws.freeze_panes = "A2"
    wb.save(output_path)
    return str(output_path)


# ──────────────────────────────────────────────
# Чтение входных файлов
# ──────────────────────────────────────────────
def read_cve_excel(path: str) -> list[str]:
    """Читает список CVE из Excel (первый столбец)."""
    df = pd.read_excel(path, header=None)
    return [str(v).strip() for v in df.iloc[:, 0].dropna() if str(v).strip()]


def read_hosts_excel(path: str) -> list[dict]:
    """Читает список хостов из Excel (OS, Version, Build)."""
    df = pd.read_excel(path)
    df.columns = [c.strip().lower() for c in df.columns]

    hosts = []
    for _, row in df.iterrows():
        os_name = str(row.get("os", "")).strip()
        version = str(row.get("version", "")).strip()
        build = str(row.get("build", "")).strip()
        if os_name and version and build:
            hosts.append({"os": os_name, "version": version, "build": build})
    return hosts


# ──────────────────────────────────────────────
# Импорт для generate_template
# ──────────────────────────────────────────────
import db