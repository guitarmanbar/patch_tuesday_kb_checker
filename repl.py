"""REPL-интерфейс на базе cmd."""

import cmd
import sys
from pathlib import Path

import db
import msrc_client
import cvrf_parser
import analyzer
import excel_io
from config import OUTPUT_DIR


class PatchREPL(cmd.Cmd):
    intro = (
        "\n╔══════════════════════════════════════════════════╗\n"
        "║   Patch Tuesday Analyzer — интерактивный режим  ║\n"
        "║   Введите 'help' для списка команд              ║\n"
        "╚══════════════════════════════════════════════════╝\n"
    )
    prompt = "pt> "

    # ──────────────────────────────────────────────
    # Команда: cve
    # ──────────────────────────────────────────────
    def do_cve(self, arg: str) -> None:
        """
        Анализ CVE → KB.
        Использование:
          cve CVE-2024-49112, CVE-2024-43600
          cve                          (интерактивный ввод столбиком)
          cve --file cves.xlsx         (загрузка из Excel)
        """
        arg = arg.strip()

        if not arg:
            cve_list = self._input_cve_interactive()
        elif arg.startswith("--file"):
            path = arg.replace("--file", "").strip()
            cve_list = excel_io.read_cve_excel(path)
        else:
            cve_list = [c.strip() for c in arg.split(",") if c.strip()]

        if not cve_list:
            print("Список CVE пуст.")
            return

        print(f"\nАнализ {len(cve_list)} CVE...")
        self._ensure_cache_for_cves(cve_list)

        results = analyzer.analyze_cves(cve_list)
        self._last_results = results

        output_path = excel_io.write_cve_results(results)
        print(f"\nРезультат сохранён: {output_path}")
        print(f"  Всего строк: {len(results)}")

    # ──────────────────────────────────────────────
    # Команда: host
    # ──────────────────────────────────────────────
    def do_host(self, arg: str) -> None:
        """
        Анализ хостов → рекомендуемые CU.
        Использование:
          host                         (интерактивный ввод)
          host --file hosts.xlsx       (загрузка из Excel)
        """
        arg = arg.strip()

        if arg.startswith("--file"):
            path = arg.replace("--file", "").strip()
            hosts = excel_io.read_hosts_excel(path)
        else:
            hosts = self._input_hosts_interactive()

        if not hosts:
            print("Список хостов пуст.")
            return

        print(f"\nАнализ {len(hosts)} хостов...")
        results = analyzer.analyze_hosts(hosts)
        self._last_results = results

        output_path = excel_io.write_host_results(results)
        print(f"\nРезультат сохранён: {output_path}")
        print(f"  Всего строк: {len(results)}")

    # ──────────────────────────────────────────────
    # Команда: template
    # ──────────────────────────────────────────────
    def do_template(self, arg: str) -> None:
        """Генерация предзаполненного шаблона Excel."""
        output_path = excel_io.generate_template()
        print(f"Шаблон сохранён: {output_path}")

    # ──────────────────────────────────────────────
    # Команда: refresh
    # ──────────────────────────────────────────────
    def do_refresh(self, arg: str) -> None:
        """
        Обновление кэша CVRF.
        Использование:
          refresh              (загрузить все отсутствующие релизы)
          refresh 2025-Jun     (загрузить конкретный релиз)
        """
        arg = arg.strip()
        if arg:
            releases = [arg]
        else:
            releases = msrc_client.fetch_all_releases_since(2024)

        to_fetch = [r for r in releases if not db.is_cached(r)]
        if not to_fetch:
            print("Кэш актуален — все релизы уже загружены.")
            return

        print(f"Загрузка {len(to_fetch)} релизов...")
        for rid in to_fetch:
            try:
                doc = msrc_client.fetch_cvrf_document(rid)
                if doc:
                    db.save_cvrf(rid, doc)
                    parsed = cvrf_parser.parse_cvrf_document(doc, rid)
                    self._persist_parsed(rid, parsed)
                    print(f"  ✓ {rid}")
                else:
                    print(f"  ✗ {rid} — пустой ответ")
            except Exception as e:
                print(f"  ✗ {rid} — ошибка: {e}")

        print("Обновление кэша завершено.")

    # ──────────────────────────────────────────────
    # Команда: cache
    # ──────────────────────────────────────────────
    def do_cache(self, arg: str) -> None:
        """Информация о состоянии кэша."""
        releases = db.get_cached_releases()
        print(f"\nКэшировано релизов: {len(releases)}")
        if releases:
            print(f"  Первый: {releases[0]}")
            print(f"  Последний: {releases[-1]}")

    # ──────────────────────────────────────────────
    # Команда: export
    # ──────────────────────────────────────────────
    def do_export(self, arg: str) -> None:
        """Экспорт последнего результата в указанный файл."""
        if not hasattr(self, "_last_results") or not self._last_results:
            print("Нет результатов для экспорта. Сначала выполните 'cve' или 'host'.")
            return

        path = arg.strip() or str(OUTPUT_DIR / "export.xlsx")
        if isinstance(self._last_results[0], dict) and "cve" in self._last_results[0]:
            excel_io.write_cve_results(self._last_results, path)
        else:
            excel_io.write_host_results(self._last_results, path)
        print(f"Результат сохранён: {path}")

    # ──────────────────────────────────────────────
    # Команда: quit
    # ──────────────────────────────────────────────
    def do_quit(self, arg: str) -> None:
        """Выход из REPL."""
        print("До свидания!")
        return True

    def do_exit(self, arg: str) -> None:
        """Выход из REPL (синоним quit)."""
        return self.do_quit(arg)

    def do_EOF(self, arg: str) -> None:
        """Обработка Ctrl+D."""
        print()
        return self.do_quit(arg)

    # ──────────────────────────────────────────────
    # Внутренние методы
    # ──────────────────────────────────────────────
    def _input_cve_interactive(self) -> list[str]:
        """Интерактивный ввод CVE столбиком."""
        print("Введите CVE (по одному на строку, пустая строка — завершение):")
        lines = []
        while True:
            try:
                line = input("  > ").strip()
            except EOFError:
                break
            if not line:
                break
            lines.append(line)
        return lines

    def _input_hosts_interactive(self) -> list[dict]:
        """Интерактивный ввод хостов."""
        print("Введите хосты в формате: OS Version Build")
        print("Пример: Windows Server 2022 21H2 20348")
        print("Пустая строка — завершение:")
        hosts = []
        while True:
            try:
                line = input("  > ").strip()
            except EOFError:
                break
            if not line:
                break
            parts = line.split()
            if len(parts) >= 3:
                hosts.append({
                    "os": " ".join(parts[:-2]),
                    "version": parts[-2],
                    "build": parts[-1],
                })
        return hosts

    def _ensure_cache_for_cves(self, cve_list: list[str]) -> None:
        """Догружает CVRF, если нужных CVE нет в кэше."""
        cached = set(db.get_cached_releases())
        # Если кэш пуст — загружаем все релизы с 2024
        if not cached:
            print("Кэш пуст. Загрузка релизов с 2024 года...")
            self.do_refresh("")

    def _persist_parsed(self, release_id: str, parsed: dict) -> None:
        """Сохраняет распарсенные данные в БД."""
        for pid, (name, group) in parsed["products"].items():
            db.upsert_product(pid, name, group)

        for cve_row in parsed["cves"]:
            db.upsert_cve(*cve_row)

        for row in parsed["cve_product_kb"]:
            db.insert_cve_product_kb(*row)

        for row in parsed["builds"]:
            db.insert_build(*row)