#!/usr/bin/env python3
"""
Скрипт для автоматического сбора данных Patch Tuesday из MSRC CVRF API
и генерации Excel-таблицы с уязвимостями для Windows 10, Windows Server и Office perpetual.

Требования:
    pip install requests pandas openpyxl

Использование:
    python patch_tuesday_updater.py

Выходные файлы:
    - patch_tuesday_2024_2026.xlsx: Excel-файл с таблицей обновлений
    - patch_tuesday_raw.csv: сырые данные в CSV для отладки
"""

import json
import requests
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
import sys

# === КОНФИГУРАЦИЯ ===
CVRF_BASE = "https://api.msrc.microsoft.com/cvrf/v3.0"
START_YEAR = 2024
END_YEAR = 2026
END_MONTH = 9  # Сентябрь
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Маппинг ProductID → (Продукт, Версия, Сборка)
PRODUCT_MAP = {
    # Windows 10
    'Windows 10 Version 1507 for x64-based Systems': ('Windows 10', '1507', '10240'),
    'Windows 10 Version 1507 for 32-bit Systems': ('Windows 10', '1507', '10240'),
    'Windows 10 Version 1607 for x64-based Systems': ('Windows 10', '1607', '14393'),
    'Windows 10 Version 1607 for 32-bit Systems': ('Windows 10', '1607', '14393'),
    'Windows 10 Version 1709 for x64-based Systems': ('Windows 10', '1709', '16299'),
    'Windows 10 Version 1709 for 32-bit Systems': ('Windows 10', '1709', '16299'),
    'Windows 10 Version 1803 for x64-based Systems': ('Windows 10', '1803', '17134'),
    'Windows 10 Version 1803 for 32-bit Systems': ('Windows 10', '1803', '17134'),
    'Windows 10 Version 1809 for x64-based Systems': ('Windows 10', '1809', '17763'),
    'Windows 10 Version 1809 for 32-bit Systems': ('Windows 10', '1809', '17763'),
    'Windows 10 Version 1903 for x64-based Systems': ('Windows 10', '1903', '18362'),
    'Windows 10 Version 1903 for 32-bit Systems': ('Windows 10', '1903', '18362'),
    'Windows 10 Version 1909 for x64-based Systems': ('Windows 10', '1909', '18363'),
    'Windows 10 Version 1909 for 32-bit Systems': ('Windows 10', '1909', '18363'),
    'Windows 10 Version 2004 for x64-based Systems': ('Windows 10', '2004', '19041'),
    'Windows 10 Version 2004 for 32-bit Systems': ('Windows 10', '2004', '19041'),
    'Windows 10 Version 20H2 for x64-based Systems': ('Windows 10', '20H2', '19042'),
    'Windows 10 Version 20H2 for 32-bit Systems': ('Windows 10', '20H2', '19042'),
    'Windows 10 Version 21H1 for x64-based Systems': ('Windows 10', '21H1', '19043'),
    'Windows 10 Version 21H1 for 32-bit Systems': ('Windows 10', '21H1', '19043'),
    'Windows 10 Version 21H2 for x64-based Systems': ('Windows 10', '21H2', '19044'),
    'Windows 10 Version 21H2 for 32-bit Systems': ('Windows 10', '21H2', '19044'),
    'Windows 10 Version 21H2 for ARM64-based Systems': ('Windows 10', '21H2', '19044'),
    'Windows 10 Version 22H2 for x64-based Systems': ('Windows 10', '22H2', '19045'),
    'Windows 10 Version 22H2 for 32-bit Systems': ('Windows 10', '22H2', '19045'),
    'Windows 10 Version 22H2 for ARM64-based Systems': ('Windows 10', '22H2', '19045'),
    
    # Windows Server
    'Windows Server 2016': ('Windows Server', '2016', '14393'),
    'Windows Server 2016 (Server Core installation)': ('Windows Server', '2016', '14393'),
    'Windows Server 2019': ('Windows Server', '2019', '17763'),
    'Windows Server 2019 (Server Core installation)': ('Windows Server', '2019', '17763'),
    'Windows Server 2022': ('Windows Server', '2022', '20348'),
    'Windows Server 2022 (Server Core installation)': ('Windows Server', '2022', '20348'),
    'Windows Server 2022 23H2': ('Windows Server', '2022 23H2', '20348'),
    'Windows Server 2022 23H2 (Server Core installation)': ('Windows Server', '2022 23H2', '20348'),
    'Windows Server 2025': ('Windows Server', '2025', '26100'),
    'Windows Server 2025 (Server Core installation)': ('Windows Server', '2025', '26100'),
}

# Office компоненты
OFFICE_COMPONENT_MAP = {
    'Office 2016 (ace)': ('Office', '2016', 'Access'),
    'Office 2016 (oart)': ('Office', '2016', 'Office Core'),
    'Office 2016 (word)': ('Office', '2016', 'Word'),
    'Office 2016 (excel)': ('Office', '2016', 'Excel'),
    'Office 2016 (outlook)': ('Office', '2016', 'Outlook'),
    'Office 2016 (powerpoint)': ('Office', '2016', 'PowerPoint'),
    'Office 2016 (publisher)': ('Office', '2016', 'Publisher'),
    'Office 2016 (visio)': ('Office', '2016', 'Visio'),
    'Office 2016 (project)': ('Office', '2016', 'Project'),
    'Office 2019 (ace)': ('Office', '2019', 'Access'),
    'Office 2019 (oart)': ('Office', '2019', 'Office Core'),
    'Office 2019 (word)': ('Office', '2019', 'Word'),
    'Office 2019 (excel)': ('Office', '2019', 'Excel'),
    'Office 2019 (outlook)': ('Office', '2019', 'Outlook'),
    'Office 2019 (powerpoint)': ('Office', '2019', 'PowerPoint'),
    'Office 2019 (publisher)': ('Office', '2019', 'Publisher'),
    'Office 2019 (visio)': ('Office', '2019', 'Visio'),
    'Office 2019 (project)': ('Office', '2019', 'Project'),
    'Office 2021 (ace)': ('Office', '2021', 'Access'),
    'Office 2021 (oart)': ('Office', '2021', 'Office Core'),
    'Office 2021 (word)': ('Office', '2021', 'Word'),
    'Office 2021 (excel)': ('Office', '2021', 'Excel'),
    'Office 2021 (outlook)': ('Office', '2021', 'Outlook'),
    'Office 2021 (powerpoint)': ('Office', '2021', 'PowerPoint'),
    'Office 2021 (publisher)': ('Office', '2021', 'Publisher'),
    'Office 2021 (visio)': ('Office', '2021', 'Visio'),
    'Office 2021 (project)': ('Office', '2021', 'Project'),
    'Office LTSC 2024 (ace)': ('Office', '2024', 'Access'),
    'Office LTSC 2024 (oart)': ('Office', '2024', 'Office Core'),
    'Office LTSC 2024 (word)': ('Office', '2024', 'Word'),
    'Office LTSC 2024 (excel)': ('Office', '2024', 'Excel'),
    'Office LTSC 2024 (outlook)': ('Office', '2024', 'Outlook'),
    'Office LTSC 2024 (powerpoint)': ('Office', '2024', 'PowerPoint'),
    'Office LTSC 2024 (publisher)': ('Office', '2024', 'Publisher'),
    'Office LTSC 2024 (visio)': ('Office', '2024', 'Visio'),
    'Office LTSC 2024 (project)': ('Office', '2024', 'Project'),
}


def get_monthly_updates(year, month):
    """Загружает CVRF XML для конкретного месяца"""
    url = f"{CVRF_BASE}/cvrf/{year}-{month}"
    try:
        print(f"Загрузка {url}...")
        response = requests.get(url, timeout=30)
        if response.status_code == 200:
            return response.text
        else:
            print(f"Warning: Failed to fetch {url}, status {response.status_code}")
            return None
    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return None


def parse_cvrf_xml(xml_content, year, month):
    """Парсит CVRF XML и извлекает данные об уязвимостях, продуктах и KB"""
    if not xml_content:
        return []
    
    ns = {
        'cvrf': 'http://docs.oasis-open.org/csaf/csaf/v2.0/cvrf/v1.2',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'ns0': 'http://docs.oasis-open.org/csaf/csaf/v2.0/cvrf/v1.2'
    }
    
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        print(f"XML parse error for {year}-{month}: {e}")
        return []
    
    records = []
    
    for vuln in root.findall('.//ns0:Vulnerability', ns):
        cve_id_elem = vuln.find('ns0:CVE', ns)
        cve_id = cve_id_elem.text if cve_id_elem is not None else "N/A"
        
        product_statuses = vuln.findall('ns0:ProductStatuses/ns0:Status', ns)
        affected_products = []
        for status in product_statuses:
            status_type = status.get('Type')
            if status_type == 'Affected':
                for product_id in status.findall('ns0:ProductID', ns):
                    affected_products.append(product_id.text)
        
        remediations = vuln.findall('ns0:Remediations/ns0:Remediation', ns)
        kb_numbers = []
        for rem in remediations:
            rem_type = rem.get('Type')
            if rem_type == 'Vendor Fix':
                description_elem = rem.find('ns0:Description', ns)
                if description_elem is not None:
                    desc_text = description_elem.text
                    if desc_text:
                        kb_match = re.search(r'KB?(\d{6,7})', desc_text, re.IGNORECASE)
                        if kb_match:
                            kb_numbers.append(f"KB{kb_match.group(1)}")
        
        if not kb_numbers:
            notes = vuln.findall('ns0:Notes/ns0:Note', ns)
            for note in notes:
                note_text = note.text
                if note_text:
                    kb_matches = re.findall(r'KB?(\d{6,7})', note_text, re.IGNORECASE)
                    for kb_match in kb_matches:
                        kb_numbers.append(f"KB{kb_match}")
        
        for product_id in affected_products:
            records.append({
                'CVE': cve_id,
                'ProductID': product_id,
                'KB': ', '.join(kb_numbers) if kb_numbers else '',
                'Year': year,
                'Month': month
            })
    
    return records


def expand_product_id(product_id):
    """Расширяет ProductID до Продукт/Версия/Сборка"""
    if product_id in PRODUCT_MAP:
        return PRODUCT_MAP[product_id]
    
    for key, value in PRODUCT_MAP.items():
        if key in product_id:
            return value
    
    for key, value in OFFICE_COMPONENT_MAP.items():
        if key in product_id:
            return (value[0], value[1], value[2])
    
    return ('Unknown', product_id, '')


def main():
    print("Загрузка данных из MSRC CVRF API...")
    all_records = []
    
    for year in range(START_YEAR, END_YEAR + 1):
        for month in MONTHS:
            if year == END_YEAR and MONTHS.index(month) + 1 > END_MONTH:
                break
            
            xml_content = get_monthly_updates(year, month)
            if xml_content:
                records = parse_cvrf_xml(xml_content, year, month)
                all_records.extend(records)
    
    print(f"Всего записей: {len(all_records)}")
    
    processed_records = []
    for record in all_records:
        product_info = expand_product_id(record['ProductID'])
        product_type = product_info[0]
        
        if product_type == 'Office':
            office_version = product_info[1]
            component = product_info[2] if len(product_info) > 2 else 'Office Core'
            processed_records.append({
                'Product': product_type,
                'Version': office_version,
                'Component': component,
                'OS_Build': '',
                'KB': record['KB'],
                'Release_Date': f"{record['Year']}-{record['Month']}",
                'CVE': record['CVE']
            })
        elif product_type in ['Windows 10', 'Windows Server']:
            processed_records.append({
                'Product': product_type,
                'Version': product_info[1],
                'Component': '',
                'OS_Build': product_info[2],
                'KB': record['KB'],
                'Release_Date': f"{record['Year']}-{record['Month']}",
                'CVE': record['CVE']
            })
        else:
            processed_records.append({
                'Product': product_type,
                'Version': product_info[1] if len(product_info) > 1 else '',
                'Component': '',
                'OS_Build': product_info[2] if len(product_info) > 2 else '',
                'KB': record['KB'],
                'Release_Date': f"{record['Year']}-{record['Month']}",
                'CVE': record['CVE']
            })
    
    df = pd.DataFrame(processed_records)
    
    if len(df) > 0:
        grouped = df.groupby(['Product', 'Version', 'Component', 'OS_Build', 'KB', 'Release_Date'])['CVE'].apply(lambda x: ', '.join(sorted(set(x)))).reset_index()
    else:
        grouped = df
    
    print(f"После группировки: {len(grouped)} записей")
    
    # Создаём Excel
    wb = Workbook()
    
    # Windows 10
    ws_win10 = wb.active
    ws_win10.title = "Windows 10"
    headers_win10 = ['Продукт', 'Версия', 'Сборка', 'KB', 'Дата релиза', 'CVE']
    ws_win10.append(headers_win10)
    
    if len(grouped) > 0:
        win10_data = grouped[grouped['Product'] == 'Windows 10'][['Product', 'Version', 'OS_Build', 'KB', 'Release_Date', 'CVE']]
        for _, row in win10_data.iterrows():
            ws_win10.append(row.tolist())
    
    # Windows Server
    ws_server = wb.create_sheet("Windows Server")
    ws_server.append(headers_win10)
    
    if len(grouped) > 0:
        server_data = grouped[grouped['Product'] == 'Windows Server'][['Product', 'Version', 'OS_Build', 'KB', 'Release_Date', 'CVE']]
        for _, row in server_data.iterrows():
            ws_server.append(row.tolist())
    
    # Office
    ws_office = wb.create_sheet("Office Perpetual")
    headers_office = ['Продукт', 'Версия', 'Компонент', 'KB', 'Дата релиза', 'CVE']
    ws_office.append(headers_office)
    
    if len(grouped) > 0:
        office_data = grouped[grouped['Product'] == 'Office'][['Product', 'Version', 'Component', 'KB', 'Release_Date', 'CVE']]
        for _, row in office_data.iterrows():
            ws_office.append(row.tolist())
    
    # README
    ws_readme = wb.create_sheet("README")
    ws_readme.append(['Инструкция по использованию'])
    ws_readme.append([])
    ws_readme.append(['Этот файл содержит данные Patch Tuesday за период январь 2024 — сентябрь 2026.'])
    ws_readme.append(['Данные загружены автоматически из MSRC CVRF API.'])
    ws_readme.append([])
    ws_readme.append(['Листы:'])
    ws_readme.append(['- Windows 10: все версии Windows 10'])
    ws_readme.append(['- Windows Server: 2016, 2019, 2022, 2022 23H2, 2025'])
    ws_readme.append(['- Office Perpetual: Office 2016, 2019, 2021, LTSC 2024'])
    ws_readme.append([])
    ws_readme.append(['Колонки:'])
    ws_readme.append(['- Продукт, Версия, Сборка/Компонент, KB, Дата релиза, CVE'])
    ws_readme.append([])
    ws_readme.append(['Примечание: каждое следующее CU включает исправления всех предыдущих CVE.'])
    
    excel_path = 'patch_tuesday_2024_2026.xlsx'
    wb.save(excel_path)
    print(f"Excel сохранён: {excel_path}")
    
    csv_path = 'patch_tuesday_raw.csv'
    grouped.to_csv(csv_path, index=False, sep=';')
    print(f"CSV сохранён: {csv_path}")
    
    print("Готово!")


if __name__ == '__main__':
    main()
