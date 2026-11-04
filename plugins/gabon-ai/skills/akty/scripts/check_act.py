#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "pypdf>=4.0", "reportlab>=4.0", "python-docx>=1.1"]
# ///
"""Независимая проверка акта: формулы Excel (акт, накопительная, справка), PDF и Word.

    uv run check_act.py act.json AKT_xxx.xlsx AKT_xxx.pdf [AKT_xxx.docx]
"""
from __future__ import annotations

import re
import sys
from decimal import Decimal
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from akty_core import ActError, load_act  # noqa: E402
from check_kp import FormulaEvaluator  # noqa: E402
from kp_core import fmt_money  # noqa: E402


def main() -> int:
    if len(sys.argv) not in (4, 5):
        print(__doc__)
        return 2
    inp, xlsx, pdf = (Path(x) for x in sys.argv[1:4])
    docx_path = Path(sys.argv[4]) if len(sys.argv) == 5 else None
    try:
        a = load_act(inp)
    except ActError as e:
        print("Акт не проходит проверку:\n  - " + "\n  - ".join(e.errors))
        return 1
    calc, c = a["_calc"], a["_contract"]
    ok, problems = [], []
    from openpyxl import load_workbook
    wb = load_workbook(xlsx)
    ev = FormulaEvaluator(wb)
    ws = wb["Акт"]
    rows = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 10).value in c["_by_n"]
            and isinstance(ws.cell(r, 6).value, (int, float))]
    if len(rows) != len(calc["rows"]):
        problems.append(f"Строк в акте Excel {len(rows)}, ожидалось {len(calc['rows'])}.")
    for r, exp in zip(rows, calc["rows"]):
        raw = ws.cell(r, 8).value
        if not (isinstance(raw, str) and raw.startswith("=")):
            problems.append(f"H{r}: стоимость записана числом, а не формулой.")
        got = ev.value("Акт", f"H{r}")
        if got != exp["amount"] or str(ws.cell(r, 10).value) != exp["n"]:
            problems.append(f"Строка {r} «{exp['name']}»: Excel {got}, ожидалось {exp['amount']}.")
    exp_names = {"ACT_SUB": calc["subtotal"], "ACT_VAT": calc["vat"], "ACT_TOTAL": calc["total"],
                 "ACT_PERIOD": calc["total"], "ACT_PAY": calc["pay"]}
    if calc["adv"]:
        exp_names["ACT_OFFSET"] = calc["offset"]
    if calc["ret"]:
        exp_names["ACT_RET"] = calc["retention"]
    for name, exp in exp_names.items():
        dn = wb.defined_names.get(name)
        if dn is None:
            problems.append(f"Нет ячейки {name}.")
            continue
        sheet, coord = next(dn.destinations)
        got = ev.value(sheet, coord.replace("$", ""))
        if got != exp:
            problems.append(f"{name}: Excel {got}, ожидалось {exp}.")
    ns = wb["Накопительная"]
    this = {r["n"]: r["qty"] for r in calc["rows"]}
    for r in range(4, ns.max_row + 1):
        nn = str(ns.cell(r, 1).value)
        if nn not in c["_by_n"]:
            continue
        exp_cum = a["_prev_qty"].get(nn, Decimal(0)) + this.get(nn, Decimal(0))
        got = ev.value("Накопительная", f"G{r}")
        if got != exp_cum:
            problems.append(f"Накопительная, позиция {nn}: с начала {got}, ожидалось {exp_cum}.")
    if not problems:
        ok.append(f"Excel: {len(rows)} строк акта, итоги, справка и накопительная ведомость верны.")

    from pypdf import PdfReader
    text = re.sub(r"\s+", " ", " ".join(p.extract_text() or "" for p in PdfReader(str(pdf)).pages))
    need = [fmt_money(r["amount"]) for r in calc["rows"]] + \
        [fmt_money(calc[k]) for k in ("subtotal", "vat", "total", "pay")]
    miss = [x for x in need if x not in text]
    if miss or f"№ {c['number']}" not in text:
        problems.append(f"PDF: не найдены {', '.join(miss[:5]) or 'номер договора'}.")
    else:
        ok.append("PDF: строки, итоги и сумма к оплате совпадают.")
    if docx_path:
        import docx
        d = docx.Document(str(docx_path))
        dt = " ".join(p.text for p in d.paragraphs) + " " + " ".join(
            cell.text for t in d.tables for row in t.rows for cell in row.cells)
        miss = [x for x in need[:-1] if x not in dt]
        if miss:
            problems.append(f"Word: не найдены {', '.join(miss[:5])}.")
        else:
            ok.append("Word: строки и итоги совпадают.")
    print("ПРОВЕРКА АКТА")
    for x in ok:
        print(f"  [OK]  {x}")
    for x in problems:
        print(f"  [ОШИБКА] {x}")
    print("РЕЗУЛЬТАТ:", "пройдена" if not problems else f"найдено ошибок: {len(problems)}")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
