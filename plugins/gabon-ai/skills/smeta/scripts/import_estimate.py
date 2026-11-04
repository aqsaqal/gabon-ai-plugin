#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Смета или КП подрядчика из Excel -> contractor.json для сравнения.

    uv run import_estimate.py смета_подрядчика.xlsx --out contractor.json
        [--sheet Лист1] [--header-row 8] [--map name=B,unit=C,qty=D,price=E,amount=F,code=A]
        [--contractor "ТОО …"] [--prices-include-vat]

Колонки находятся по заголовкам (наименование, ед., кол-во, цена, сумма/стоимость),
при неудаче укажите --map. Строки «итого», «НДС», «всего» распознаются как итоги.
Исходный файл не изменяется. Смету в PDF перенесите в тот же формат JSON вручную.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path

HEAD = {"name": ("наименов",), "unit": ("ед",), "qty": ("кол",), "price": ("цена",),
        "amount": ("сумма", "стоимост", "всего"), "code": ("код", "шифр", "обоснов")}


def num(v):
    if v is None or v == "" or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    s = re.sub(r"[^\d,.\-]", "", str(v)).replace(",", ".")
    try:
        return Decimal(s) if s not in ("", ".", "-") else None
    except InvalidOperation:
        return None


def col_index(letter: str) -> int:
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter.strip().upper()) - 1


def main() -> int:
    ap = argparse.ArgumentParser(description="Импорт сметы подрядчика из Excel")
    ap.add_argument("xlsx")
    ap.add_argument("--out", required=True)
    ap.add_argument("--sheet")
    ap.add_argument("--header-row", type=int)
    ap.add_argument("--map", help="name=B,unit=C,qty=D,price=E,amount=F[,code=A]")
    ap.add_argument("--contractor", default="")
    ap.add_argument("--prices-include-vat", action="store_true")
    a = ap.parse_args()

    from openpyxl import load_workbook
    wb = load_workbook(a.xlsx, data_only=True, read_only=True)
    ws = wb[a.sheet] if a.sheet else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))

    cols, hdr = {}, a.header_row
    if a.map:
        for part in a.map.split(","):
            k, v = part.split("=")
            cols[k.strip()] = col_index(v)
        hdr = hdr or 0
    else:
        for i, row in enumerate(rows[:40]):
            found = {}
            for ci, v in enumerate(row):
                t = str(v or "").lower().replace("ё", "е")
                for k, keys in HEAD.items():
                    if k not in found and any(t.startswith(x) or f" {x}" in t for x in keys):
                        found[k] = ci
            if {"name", "qty", "price"} <= found.keys():
                cols, hdr = found, i + 1
                break
    need = {"name", "unit", "qty", "price"}
    if not need <= cols.keys():
        print("Не удалось найти колонки " + ", ".join(sorted(need - cols.keys()))
              + ". Укажите их явно: --map name=B,unit=C,qty=D,price=E,amount=F")
        return 1

    def cell(row, k):
        i = cols.get(k)
        return row[i] if i is not None and i < len(row) else None

    positions, totals, section = [], {}, ""
    for rn, row in enumerate(rows[hdr:], hdr + 1):
        name = str(cell(row, "name") or "").strip()
        if not name:
            continue
        low = name.lower()
        qty, price = num(cell(row, "qty")), num(cell(row, "price"))
        amount = num(cell(row, "amount"))
        if qty is None and price is None:
            figure = amount if amount is not None else next(
                (num(v) for v in reversed(row) if num(v) is not None), None)
            if re.search(r"\bндс\b", low) and "без" not in low and "с ндс" not in low \
                    and "с учет" not in low:
                totals["vat"] = str(figure) if figure is not None else None
            elif re.search(r"всего|с ндс|с учет", low):
                totals["total"] = str(figure) if figure is not None else None
            elif re.search(r"итого", low):
                totals["subtotal"] = str(figure) if figure is not None else None
            else:
                section = name
            continue
        positions.append({"row": rn, "n": str(row[0]) if row[0] is not None else "",
                          "section": section, "name": name,
                          "unit": str(cell(row, "unit") or "").strip(),
                          "qty": None if qty is None else str(qty),
                          "price": None if price is None else str(price),
                          "amount": None if amount is None else str(amount),
                          "code": str(cell(row, "code") or "").strip()})
    out = {"source_file": Path(a.xlsx).name, "sheet": ws.title,
           "contractor": a.contractor, "prices_include_vat": a.prices_include_vat,
           "positions": positions, "totals": totals}
    Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Импортировано позиций: {len(positions)}; итоги: "
          + (", ".join(f"{k} {v}" for k, v in totals.items()) or "не найдены"))
    print(f"Файл: {Path(a.out).resolve()}")
    print("Проверьте contractor.json: разделы, позиции и итоги распознаны автоматически.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
