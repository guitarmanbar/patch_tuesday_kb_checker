#!/usr/bin/env python3
"""
Patch Tuesday Analyzer.
Точка входа: REPL или batch-режим.
"""

import argparse
import sys

import db
import msrc_client
import cvrf_parser
import analyzer
import excel_io


def batch_mode(args) -> None:
    """Выполняет одну операцию без REPL."""
    if args.mode == "cve":
        if args.input:
            cve_list = excel_io.read_cve_excel(args.input) if args.input.endswith(".xlsx") \
                else [l.strip() for l in open(args.input) if l.strip()]
        else:
            cve_list = [c.strip() for c in args.cves.split(",")]
        results = analyzer.analyze_cves(cve_list)
        path = excel_io.write_cve_results(results, args.output)
        print(f"Результат: {path}")

    elif args.mode == "host":
        hosts = excel_io.read_hosts_excel(args.input)
        results = analyzer.analyze_hosts(hosts)
        path = excel_io.write_host_results(results, args.output)
        print(f"Результат: {path}")

    elif args.mode == "template":
        path = excel_io.generate_template(args.output)
        print(f"Шаблон: {path}")

    elif args.mode == "refresh":
        from repl import PatchREPL
        r = PatchREPL()
        r.do_refresh(args.release or "")


def main():
    parser = argparse.ArgumentParser(description="Patch Tuesday Analyzer")
    parser.add_argument("--mode", choices=["cve", "host", "template", "refresh"],
                        help="Режим работы (без REPL)")
    parser.add_argument("--input", help="Входной файл (Excel или txt)")
    parser.add_argument("--output", help="Выходной файл")
    parser.add_argument("--cves", help="Список CVE через запятую")
    parser.add_argument("--release", help="Релиз для refresh (YYYY-MMM)")
    args = parser.parse_args()

    db.init_db()

    if args.mode:
        batch_mode(args)
    else:
        from repl import PatchREPL
        PatchREPL().cmdloop()


if __name__ == "__main__":
    main()