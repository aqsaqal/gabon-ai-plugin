#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Прайс-лист компании.

    uv run prices.py template data/prices.xlsx          пустой прайс-лист
    uv run prices.py check data/prices.xlsx [--date 2026-11-03] [--max-age 90]
    uv run prices.py find data/prices.xlsx бетон В25      поиск по коду или словам
"""
from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from prices_core import (DEFAULT_MAX_AGE_DAYS, KINDS, age_warning,  # noqa: E402
                         load_price_list)
from kp_core import fmt_money  # noqa: E402

HEADERS = ["Код", "Наименование", "Ед. изм.", "Цена без НДС, тг", "Вид затрат",
           "Источник (прайс поставщика, КП, база)", "Дата цены", "Примечание"]


def make_template(out: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = Workbook()
    ws = wb.active
    ws.title = "Прайс"
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DCE3EE")
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for col, w in zip("ABCDEFGH", (12, 55, 10, 16, 14, 45, 12, 30)):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"
    for r in range(2, 1001):
        ws.cell(r, 4).number_format = "#,##0.00"
        ws.cell(r, 7).number_format = "dd.mm.yyyy"
    dv = DataValidation(type="list", formula1='"' + ",".join(KINDS) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E2:E1000")
    dd = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
    dd.error, dd.errorTitle = "Введите дату цены", "Дата"
    ws.add_data_validation(dd)
    dd.add("G2:G1000")
    hl = wb.create_sheet("Инструкция")
    hl.column_dimensions["A"].width = 110
    for line in [
        "ПРАЙС-ЛИСТ КОМПАНИИ ДЛЯ GABON AI",
        "1. Одна строка = одна расценка. Код уникален, например ЗР-001 (земляные), МН-010 (монолит).",
        "2. Цена указывается БЕЗ НДС за единицу измерения.",
        "3. Вид затрат: " + ", ".join(KINDS) + ".",
        "   «комплекс» — работа вместе с материалами одной ценой.",
        "4. Источник: откуда цена (прайс поставщика № и дата, КП, расчёт, экспорт АВС-4, решение руководства).",
        "5. Дата цены обязательна. Цены старше 90 дней отмечаются как требующие актуализации.",
        "6. Цены из чужих КП-образцов не являются рыночными расценками. Указывайте это в источнике.",
        "7. В КП позиция ссылается на код: цена, ед. изм., вид и источник подставятся автоматически.",
    ]:
        hl.append([line])
    hl["A1"].font = Font(bold=True, size=12)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


def main() -> int:
    ap = argparse.ArgumentParser(description="Прайс-лист компании")
    sub = ap.add_subparsers(dest="cmd", required=True)
    t = sub.add_parser("template", help="создать пустой прайс-лист")
    t.add_argument("out")
    c = sub.add_parser("check", help="проверить прайс-лист")
    c.add_argument("prices")
    c.add_argument("--date", help="на какую дату проверять актуальность (по умолчанию сегодня)")
    c.add_argument("--max-age", type=int, default=DEFAULT_MAX_AGE_DAYS)
    f = sub.add_parser("find", help="найти расценку")
    f.add_argument("prices")
    f.add_argument("query", nargs="+")
    f.add_argument("--limit", type=int, default=10)
    a = ap.parse_args()

    if a.cmd == "template":
        out = Path(a.out)
        if out.exists():
            print(f"Файл уже существует, не перезаписываю: {out}")
            return 1
        make_template(out)
        print(f"Пустой прайс-лист: {out.resolve()}")
        return 0

    pl = load_price_list(a.prices)
    if a.cmd == "check":
        on = date.fromisoformat(a.date) if a.date else date.today()
        stale = [f"код {r.code}: {w}" for r in pl.records.values()
                 if (w := age_warning(r, on, a.max_age))]
        print(f"ПРОВЕРКА ПРАЙС-ЛИСТА {pl.path}")
        print(f"  Расценок: {len(pl.records)}")
        for e in pl.errors:
            print(f"  [ОШИБКА] {e}")
        for w in pl.warnings:
            print(f"  [ЗАМЕЧАНИЕ] {w}")
        for w in stale:
            print(f"  [УСТАРЕЛО] {w}")
        print("РЕЗУЛЬТАТ:", "пройдена" if not pl.errors else f"ошибок: {len(pl.errors)}",
              f"(устаревших цен: {len(stale)})")
        return 0 if not pl.errors else 1

    if pl.errors:
        print("Прайс-лист содержит ошибки, сначала выполните check.")
    hits = pl.find(" ".join(a.query), a.limit)
    if not hits:
        print("Ничего не найдено. Цену нужно запросить или добавить в прайс-лист.")
        return 1
    for score, r in hits:
        print(f"{r.code:10} {r.name[:60]:60} {r.unit:7} {fmt_money(r.price):>16}  "
              f"{r.kind:9} от {r.date:%d.%m.%Y}  совпадение {score}%")
    return 0


if __name__ == "__main__":
    sys.exit(main())
