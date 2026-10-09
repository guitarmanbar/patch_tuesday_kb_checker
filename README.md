# Patch Tuesday Analyzer

Инструмент для анализа обновлений безопасности Microsoft Patch Tuesday: сопоставление CVE с KB, поиск рекомендуемых кумулятивных обновлений для парка хостов, генерация шаблонов для актуализации.

---

## Содержание

- [Возможности](#возможности)
- [Архитектура](#архитектура)
- [Требования](#требования)
- [Установка](#установка)
- [Быстрый старт](#быстрый-старт)
- [Использование REPL](#использование-repl)
- [Batch-режим](#batch-режим)
- [Структура проекта](#структура-проекта)
- [Формат данных](#формат-данных)
- [Кэширование](#кэширование)
- [EOL-продукты](#eol-продукты)
- [Часто задаваемые вопросы](#часто-задаваемые-вопросы)
- [Ограничения](#ограничения)
- [Разработка](#разработка)

---

## Возможности

**Режим 1 — CVE → KB.** По списку идентификаторов CVE скрипт возвращает таблицу, где каждая CVE сопоставлена с продуктом (ОС + версия + сборка, Office, Exchange, SharePoint, SQL Server, Visual Studio, Edge, .NET), а также с полной цепочкой KB-обновлений, начиная от первого исправления и заканчивая последним LCU.

**Режим 2 — Host → рекомендуемые CU.** По списку хостов (ОС + версия + сборка) скрипт собирает все кумулятивные обновления, которые необходимо установить, чтобы закрыть все уязвимости, — от первого релиза до последнего LCU.

**Режим 3 — генерация шаблона.** Формирует предзаполненный Excel-файл со всеми известными комбинациями ОС + версия + сборка, который можно актуализировать вручную при выходе новых релизов.

**Дополнительно:**

- Поддержка всех версий Windows 10 (1507–22H2), Windows 11 (21H2–24H2), Windows Server 2016–2025.
- Поддержка всего пакета MS Office (2016, 2019, LTSC 2021, LTSC 2024, отдельные приложения).
- Пометка EOL-продуктов (снятых с поддержки) с датой окончания жизненного цикла.
- Пометка CVE без KB как `No_KB_Available`.
- Локальный SQLite-кэш CVRF-документов — повторные запросы не дёргают API.
- Работа без API-ключа (MSRC CVRF API v3.0 доступен анонимно).

---

## Архитектура

```
┌────────────────────────────────────────────────────────────┐
│                        patch.py                            │
│         (точка входа: REPL или batch-режим)                │
└──────┬────────────────────────────────────────────┬────────┘
       │                                            │
       ▼                                            ▼
┌─────────────┐                              ┌─────────────┐
│  repl.py    │                              │ batch mode  │
│  (cmd)      │                              │ (argparse)  │
└──────┬──────┘                              └──────┬──────┘
       │                                            │
       └────────────────┬───────────────────────────┘
                        ▼
                ┌───────────────┐
                │  analyzer.py  │  ◄── логика CVE→KB и Host→KB
                └───────┬───────┘
                        │
        ┌───────────────┼───────────────┐
        ▼               ▼               ▼
  ┌──────────┐   ┌────────────┐   ┌───────────┐
  │  db.py   │   │ excel_io   │   │eol_checker│
  │ (SQLite) │   │  (openpyxl)│   │  (JSON)   │
  └────┬─────┘   └────────────┘   └───────────┘
       │
       ▼
  ┌──────────────────┐      ┌───────────────────┐
  │msrc_client.py    │─────▶│  MSRC CVRF API    │
  │cvrf_parser.py    │      │  v3.0 (анонимно)  │
  └──────────────────┘      └───────────────────┘
```

**Поток данных:**

1. `msrc_client` скачивает CVRF-документ за месяц с `https://api.msrc.microsoft.com/cvrf/v3.0/cvrf/{release}`.
2. `cvrf_parser` разбирает документ в плоские структуры: продукты, CVE, связи CVE↔Продукт↔KB, сборки.
3. `db` сохраняет всё в SQLite (`data/msrc_cache.db`) — и сырой JSON, и нормализованные таблицы.
4. `analyzer` выполняет запросы к БД по заданным CVE или хостам и собирает итоговые цепочки.
5. `excel_io` записывает результат в Excel с готовым форматированием.

---

## Требования

- **Python** 3.10 или выше
- **ОС:** Windows / Linux / macOS
- **Зависимости** (устанавливаются через pip):
  - `requests`
  - `pandas`
  - `openpyxl`
  - `python-dateutil`

API-ключ MSRC **не требуется** — CVRF API v3.0 работает анонимно.

---

## Установка

```bash
# 1. Клонировать репозиторий
git clone <repo-url> patch_tuesday
cd patch_tuesday

# 2. Создать виртуальное окружение (рекомендуется)
python -m venv .venv
source .venv/bin/activate       # Linux/macOS
.venv\Scripts\activate          # Windows

# 3. Установить зависимости
pip install requests pandas openpyxl python-dateutil

# 4. Инициализировать БД (создаётся автоматически при первом запуске)
python patch.py
```

---

## Быстрый старт

```bash
# Запуск REPL
python patch.py
```

```
╔══════════════════════════════════════════════════╗
║   Patch Tuesday Analyzer — интерактивный режим  ║
║   Введите 'help' для списка команд              ║
╚══════════════════════════════════════════════════╝

pt> refresh
Загрузка 21 релизов...
  ✓ 2024-Jan
  ✓ 2024-Feb
  ...
  ✓ 2025-Dec
Обновление кэша завершено.

pt> cve CVE-2024-49112
Анализ 1 CVE...
Результат сохранён: output/cve_results.xlsx

pt> quit
```

---

## Использование REPL

Доступные команды (полный список — по `help`):

| Команда | Назначение |
|---|---|
| `cve <список>` | Анализ CVE → KB |
| `cve` | Интерактивный ввод CVE столбиком |
| `cve --file cves.xlsx` | Загрузка CVE из Excel |
| `host` | Интерактивный ввод хостов |
| `host --file hosts.xlsx` | Загрузка хостов из Excel |
| `template` | Генерация предзаполненного шаблона |
| `refresh` | Загрузка всех отсутствующих релизов CVRF |
| `refresh 2025-Jun` | Загрузка конкретного релиза |
| `cache` | Информация о состоянии кэша |
| `export <файл>` | Экспорт последнего результата |
| `help` | Справка |
| `quit` / `exit` | Выход (также Ctrl+D) |

### Пример: анализ CVE

```
pt> cve CVE-2024-49112, CVE-2024-43600
```

### Пример: анализ хостов

```
pt> host
Введите хосты в формате: OS Version Build
Пример: Windows Server 2022 21H2 20348
Пустая строка — завершение:
  > Windows Server 2022 21H2 20348
  > Windows 10 22H2 19045
  >
Анализ 2 хостов...
Результат сохранён: output/host_results.xlsx
```

### Пример: генерация шаблона

```
pt> template
Шаблон сохранён: output/host_template.xlsx
```

---

## Batch-режим

Для автоматизации по расписанию (cron / Task Scheduler):

```bash
# Анализ CVE через запятую
python patch.py --mode cve --cves "CVE-2024-49112,CVE-2024-43600" --output result.xlsx

# Анализ CVE из файла
python patch.py --mode cve --input cves.txt --output result.xlsx

# Анализ хостов
python patch.py --mode host --input hosts.xlsx --output recommendations.xlsx

# Генерация шаблона
python patch.py --mode template --output template.xlsx

# Обновление кэша
python patch.py --mode refresh
```

Batch-режим **не мешает** REPL: запуск без аргументов открывает интерактивную оболочку.

### Пример cron-задачи

```cron
# Обновлять кэш каждый второй вторник месяца в 08:00
0 8 * * 2 [ $(date +\%d) -ge 8 ] && [ $(date +\%d) -le 14 ] && \
    cd /opt/patch_tuesday && \
    .venv/bin/python patch.py --mode refresh >> logs/refresh.log 2>&1
```

---

## Структура проекта

```
patch_tuesday/
├── patch.py              # Точка входа (REPL + batch)
├── config.py             # Конфигурация (URL, пути, паттерны)
├── db.py                 # SQLite: схема, кэш, запросы
├── msrc_client.py        # Клиент MSRC CVRF API v3.0
├── cvrf_parser.py        # Парсер CVRF-документов
├── analyzer.py           # Логика: CVE → KB, Host → KB
├── excel_io.py           # Чтение/запись Excel
├── eol_checker.py        # Проверка EOL-статуса
├── repl.py               # REPL-интерфейс (cmd)
├── README.md
├── data/
│   ├── msrc_cache.db     # SQLite-кэш (создаётся автоматически)
│   └── eol_products.json # Справочник EOL (опционально)
└── output/
    ├── cve_results.xlsx
    ├── host_results.xlsx
    └── host_template.xlsx
```

---

## Формат данных

### Результат анализа CVE (`cve_results.xlsx`)

**Лист `Results`:**

| CVE | Product | Version | Build | KB_Chain | First_Fix | LCU | Status | Note |
|---|---|---|---|---|---|---|---|---|
| CVE-2024-49112 | Windows 10 | 22H2 | 19045 | KB5046613, KB5048652 | KB5046613 | KB5048652 | Supported | |
| CVE-2024-43600 | Office 2016 | — | — | KB2920716 | KB2920716 | KB2920716 | EOL (2025-10-14) | |

**Лист `No_KB_Available`** — CVE, для которых в MSRC нет KB (только workaround или mitigation).

**Лист `Summary`** — агрегированная сводка по CVE.

### Результат анализа хостов (`host_results.xlsx`)

| OS | Version | Build | KB | Release_Date | Type | Note |
|---|---|---|---|---|---|---|
| Windows Server 2022 | 21H2 | 20348 | KB5034129 | 2024-01-09 | Cumulative | |
| Windows Server 2022 | 21H2 | 20348 | KB5034770 | 2024-02-13 | Cumulative | |
| ... | ... | ... | ... | ... | ... | |
| Windows Server 2022 | 21H2 | 20348 | KB5048654 | 2024-12-10 | LCU | |

### Шаблон хостов (`host_template.xlsx`)

| OS | Version | Build |
|---|---|---|
| Windows 10 | 22H2 | 19045 |
| Windows Server 2019 | 1809 | 17763 |
| Windows Server 2022 | 21H2 | 20348 |
| Windows Server 2025 | 24H2 | 26100 |

Заполните шаблон своими хостами и подайте его в команду `host --file`.

### Входной файл CVE (`cves.xlsx`)

Один столбец без заголовка:

| |
|---|
| CVE-2024-49112 |
| CVE-2024-43600 |
| CVE-2024-49074 |

---

## Кэширование

Все скачанные CVRF-документы сохраняются в `data/msrc_cache.db` (SQLite). Повторный запрос тех же месяцев не обращается к API.

**Схема БД:**

| Таблица | Назначение |
|---|---|
| `cvrf_cache` | Сырые JSON-документы CVRF по месяцам |
| `products` | Продукты из ProductTree |
| `cves` | CVE с severity, CVSS, описанием |
| `cve_product_kb` | Связь CVE ↔ Продукт ↔ KB |
| `builds` | Индекс сборок (для режима host) |

**Состояние кэша:**

```
pt> cache

Кэшировано релизов: 21
  Первый: 2024-Jan
  Последний: 2025-Sep
```

**Обновление:**

```
pt> refresh              # дозагрузить всё отсутствующее
pt> refresh 2025-Oct     # дозагрузить конкретный месяц
```

---

## EOL-продукты

Продукты, снятые с поддержки, помечаются в колонке `Status`:

```
EOL (2025-10-14)
```

Справочник EOL хранится в `data/eol_products.json`. Если файла нет — используется встроенный список в `eol_checker.py`. Файл можно редактировать вручную:

```json
{
  "Windows 10|22H2": {"eol": "2025-10-14", "status": "Supported"},
  "Office 2019":     {"eol": "2025-10-14", "status": "Supported"},
  "Windows 11|21H2": {"eol": "2023-10-10", "status": "EOL"}
}
```

Актуальные даты — на [Microsoft Lifecycle Policy](https://learn.microsoft.com/lifecycle/).

---

## Часто задаваемые вопросы

**Q: Нужен ли API-ключ MSRC?**
A: Нет. CVRF API v3.0 доступен анонимно. Ключ требуется только для SUG API (другой сервис, не используется).

**Q: Сколько времени занимает первичная загрузка кэша?**
A: Около 2–3 минут для 21 релиза (2024-01 — 2025-09). Зависит от скорости сети.

**Q: Что делать, если CVE помечена как `No_KB_Available`?**
A: Проверьте MSRC Security Update Guide — возможно, исправление выпущено в закрытом виде, или доступен только workaround. CVE попадёт на отдельный лист Excel.

**Q: Как выводить полную сборку (`19045.5247`) вместо major (`19045`)?**
A: В `config.py` установите `BUILD_MAJOR_ONLY = False`.

**Q: Как добавить продукт, которого нет в списке?**
A: Отредактируйте `TARGET_PRODUCT_GROUPS` в `config.py`, добавив регулярное выражение для нового продукта.

**Q: Что делать, если в выводе нет данных для старой сборки?**
A: Проверьте, покрывает ли кэш нужный период. Запустите `refresh` или `refresh YYYY-MMM` для конкретного месяца.

**Q: Можно ли запустить на Windows Server Core?**
A: Да, скрипт не требует GUI. Все зависимости — чистый Python.

---

## Ограничения

1. **Только Windows-версии Office.** Office для Mac, Office Online и Office 365 не включены в анализ.
2. **Гранулярность CVRF.** В некоторых релизах MSRC указывает KB без привязки к конкретной сборке (`FixedBuild`), что приводит к менее точному сопоставлению с хостом.
3. **CVE без KB.** Часть уязвимостей исправляется только через workaround или закрытые обновления — они помечаются `No_KB_Available`.
4. **EOL-справочник статичен.** Даты окончания поддержки не обновляются автоматически; при необходимости правьте `eol_products.json` вручную.
5. **Office Release Notes не парсятся.** CVRF покрывает большинство Office-обновлений, но не все. Для полной картины по Office рекомендуется сверяться с [Office Release Notes](https://learn.microsoft.com/officeupdates/).

---

## Разработка

### Добавление нового продукта в анализ

1. В `config.py` добавьте паттерн в `TARGET_PRODUCT_GROUPS`:
   ```python
   "Teams": r"Microsoft\s+Teams",
   ```
2. Проверьте, что продукт присутствует в CVRF (`ProductTree.FullProductName`).
3. Запустите `refresh` для перезагрузки кэша.

### Добавление EOL-записи

Отредактируйте `data/eol_products.json` или `DEFAULT_EOL` в `eol_checker.py`:

```python
"Windows 11|25H2": {"eol": "2027-10-12", "status": "Supported"},
```

### Тестирование

```bash
# Проверка парсинга одного релиза
python -c "import db, msrc_client, cvrf_parser; \
  db.init_db(); \
  doc = msrc_client.fetch_cvrf_document('2025-Jan'); \
  parsed = cvrf_parser.parse_cvrf_document(doc, '2025-Jan'); \
  print(f'CVE: {len(parsed[\"cves\"])}, KB: {len(parsed[\"cve_product_kb\"])}')"
```

### Отладка

Установите уровень логирования в `config.py`:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

---

## Лицензия

Внутренний инструмент. Уточните условия использования у владельца репозитория.

---

## Ссылки

- [MSRC Security Update Guide](https://msrc.microsoft.com/update-guide)
- [MSRC CVRF API v3.0 (Swagger)](https://api.msrc.microsoft.com/cvrf/v3.0/swagger/v3/swagger.json)
- [Microsoft Update Catalog](https://www.catalog.update.microsoft.com/)
- [Windows Release Health](https://learn.microsoft.com/windows/release-health/)
- [Microsoft Lifecycle Policy](https://learn.microsoft.com/lifecycle/)
- [Office Release Notes](https://learn.microsoft.com/officeupdates/)
