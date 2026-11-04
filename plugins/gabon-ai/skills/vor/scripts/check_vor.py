#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "pypdf>=4.0", "reportlab>=4.0"]
# ///
"""Независимая проверка ВОР: формулы Excel пересчитываются и сверяются с vor.json, PDF сверяется
по каждой строке.

    uv run check_vor.py vor.json VOR_xxx.xlsx VOR_xxx.pdf
"""
from __future__ import annotations

import re
import sys
from decimal import Decimal
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1] / "kp" / "scripts"))
from check_kp import FormulaEvaluator  # noqa: E402
from vor_core import STATUSES, VorError, compute, fmt_q, load_vor, rows  # noqa: E402


def main() -> int:
    if len(sys.argv) != 4:
        print(__doc__)
        return 2
    inp, xlsx, pdf = (Path(a) for a in sys.argv[1:])
    try:
        data = load_vor(inp)
        values, _ = compute(data)
    except VorError as e:
        print("ВОР не проходит проверку:\n  - " + "\n  - ".join(e.errors))
        return 1
    places = data.get("qty_decimals", 2)
    ok, problems = [], []

    from openpyxl import load_workbook
    wb = load_workbook(xlsx)
    ws = wb["ВОР"]
    ev = FormulaEvaluator(wb)
    by_id = {it["_id"]: it for *_, it in rows(data)}
    seen = set()
    for r in range(1, ws.max_row + 1):
        rid, status = ws.cell(r, 10).value, ws.cell(r, 7).value
        if status not in STATUSES or rid is None:
            continue
        rid = str(rid)
        seen.add(rid)
        it = by_id.get(rid)
        if it is None:
            problems.append(f"Строка Excel {r}: ID «{rid}» нет в vor.json.")
            continue
        raw = ws.cell(r, 4).value
        if it["status"] == "нет данных":
            if raw not in (None, ""):
                problems.append(f"D{r}: у строки без данных указано количество.")
            continue
        if it.get("formula") and not (isinstance(raw, str) and raw.startswith("=")):
            problems.append(f"D{r}: количество записано числом, а должно считаться формулой.")
        got = ev.value("ВОР", f"D{r}")
        if not isinstance(got, Decimal) or got != values[rid]:
            problems.append(f"D{r} ({it['name']}): Excel даёт {got}, ожидалось {values[rid]}.")
    missing_rows = set(by_id) - seen
    if missing_rows:
        problems.append(f"В Excel нет строк: {', '.join(sorted(missing_rows))}.")
    n_items = len(by_id)
    if not problems:
        ok.append(f"Все {n_items} строк: количества и формулы совпадают с расчётом.")
    for st, name in zip(STATUSES, ("VOR_OK", "VOR_ASSUME", "VOR_MISSING")):
        dn = wb.defined_names.get(name)
        expected = sum(1 for it in by_id.values() if it["status"] == st)
        if dn is None:
            problems.append(f"Нет итоговой ячейки {name}.")
            continue
        sheet, coord = next(dn.destinations)
        # COUNTIF не входит в мини-вычислитель: считаем по колонке статусов
        got = sum(1 for r in range(1, ws.max_row + 1) if ws.cell(r, 7).value == st
                  and ws.cell(r, 10).value is not None)
        raw = wb[sheet][coord.replace("$", "")].value
        if not (isinstance(raw, str) and raw.startswith("=COUNTIF")):
            problems.append(f"Сводка «{st}»: должна быть формула COUNTIF.")
        elif got != expected:
            problems.append(f"Сводка «{st}»: в Excel {got} строк, в vor.json {expected}.")
    if not any("Сводка" in p for p in problems):
        ok.append("Сводка по статусам совпадает.")

    from pypdf import PdfReader
    if not pdf.exists() or pdf.stat().st_size == 0:
        problems.append(f"PDF отсутствует или пуст: {pdf}")
    else:
        reader = PdfReader(str(pdf))
        flat = re.sub(r"\s+", " ", "\n".join(p.extract_text() or "" for p in reader.pages)
                      .replace(" ", " "))
        ok.append(f"PDF открывается, страниц: {len(reader.pages)}.")
        if str(data["number"]) not in flat:
            problems.append("В PDF не найден номер ВОР.")
        miss = [fmt_q(v, places) for v in values.values() if fmt_q(v, places) not in flat]
        if miss:
            problems.append(f"В PDF не найдены количества: {', '.join(miss[:5])}.")
        else:
            ok.append("В PDF есть все количества.")

    print("ПРОВЕРКА ВОР")
    for x in ok:
        print(f"  [OK]  {x}")
    for x in problems:
        print(f"  [ОШИБКА] {x}")
    print("РЕЗУЛЬТАТ:", "пройдена" if not problems else f"найдено ошибок: {len(problems)}")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
