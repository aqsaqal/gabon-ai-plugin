#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "pypdf>=4.0", "reportlab>=4.0"]
# ///
"""Независимая проверка сметы: пересчёт всех формул Excel и сверка PDF.

    uv run check_smeta.py smeta.json SMETA_xxx.xlsx SMETA_xxx.pdf
"""
from __future__ import annotations

import re
import sys
from decimal import Decimal
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from smeta_core import TYPES, SmetaError, compute, load_smeta  # noqa: E402
from check_kp import FormulaEvaluator  # noqa: E402
from kp_core import fmt_money  # noqa: E402


def main() -> int:
    if len(sys.argv) != 4:
        print(__doc__)
        return 2
    inp, xlsx, pdf = (Path(a) for a in sys.argv[1:])
    try:
        data = load_smeta(inp)
        calc, _ = compute(data)
    except SmetaError as e:
        print("Смета не проходит проверку:\n  - " + "\n  - ".join(e.errors))
        return 1
    ok, problems = [], []
    from openpyxl import load_workbook
    wb = load_workbook(xlsx)
    ws = wb["Смета"]
    ev = FormulaEvaluator(wb)

    exp_pos = [p for s in calc["sections"] for p in s["positions"]]
    exp_res = [r for p in exp_pos for r in p["resources"]]
    got_pos, got_res = [], []
    for r in range(1, ws.max_row + 1):
        t = ws.cell(r, 8).value
        if t == "позиция":
            got_pos.append(r)
        elif t in TYPES:
            got_res.append(r)
    if len(got_pos) != len(exp_pos) or len(got_res) != len(exp_res):
        problems.append(f"Строк в Excel: позиций {len(got_pos)}, ресурсов {len(got_res)}; "
                        f"ожидалось {len(exp_pos)} и {len(exp_res)}.")
    else:
        for r, e in zip(got_res, exp_res):
            for col in (5, 7):
                raw = ws.cell(r, col).value
                if not (isinstance(raw, str) and raw.startswith("=")):
                    problems.append(f"{'EG'[col == 7]}{r}: записано числом, а не формулой.")
            q, a = ev.value("Смета", f"E{r}"), ev.value("Смета", f"G{r}")
            if q != e["qty"] or a != e["amount"]:
                problems.append(f"Строка {r} «{e['name']}»: Excel {q} / {a}, ожидалось "
                                f"{e['qty']} / {e['amount']}.")
        for r, e in zip(got_pos, exp_pos):
            if ev.value("Смета", f"G{r}") != e["total"]:
                problems.append(f"G{r} «{e['name']}»: итог позиции не совпадает ({e['total']}).")
        if not problems:
            ok.append(f"Все {len(exp_pos)} позиций и {len(exp_res)} ресурсов: формулы верны.")

    checks = [("SM_" + n, label, calc["by_type"][t]) for t, n, label in
              zip(TYPES, ("LABOR", "MACH", "MAT", "TRANS"), TYPES.values())]
    checks += [("SM_DIRECT", "Прямые затраты", calc["direct"]),
               ("SM_OVERHEAD", "Накладные", calc["overhead"]),
               ("SM_PROFIT", "Прибыль", calc["profit"]),
               ("SM_SUBTOTAL", "Итого без НДС", calc["subtotal"]),
               ("SM_VAT", "НДС", calc["vat"]), ("SM_TOTAL", "Всего", calc["total"]),
               ("SM_CONTROL", "Контроль", Decimal(0)),
               ("SM_RES_TOTAL", "Итог ресурсной ведомости", calc["direct"])]
    for name, label, exp in checks:
        dn = wb.defined_names.get(name)
        if dn is None:
            problems.append(f"Нет ячейки {name} ({label}).")
            continue
        sheet, coord = next(dn.destinations)
        got = ev.value(sheet, coord.replace("$", ""))
        if got != exp:
            problems.append(f"{label}: Excel {got}, ожидалось {exp}.")
        else:
            ok.append(f"{label}: {fmt_money(got)}")

    from pypdf import PdfReader
    reader = PdfReader(str(pdf))
    flat = re.sub(r"\s+", " ", " ".join(p.extract_text() or "" for p in reader.pages)
                  .replace(" ", " "))
    need = [fmt_money(p["total"]) for p in exp_pos] + [fmt_money(calc[k]) for k in
                                                        ("direct", "subtotal", "vat", "total")]
    miss = [x for x in need if x not in flat]
    if miss or str(data["number"]) not in flat:
        problems.append(f"В PDF не найдены: {', '.join(miss[:5]) or 'номер сметы'}.")
    else:
        ok.append(f"PDF ({len(reader.pages)} стр.): номер, итоги позиций и сметы совпадают.")

    print("ПРОВЕРКА СМЕТЫ")
    for x in ok:
        print(f"  [OK]  {x}")
    for x in problems:
        print(f"  [ОШИБКА] {x}")
    print("РЕЗУЛЬТАТ:", "пройдена" if not problems else f"найдено ошибок: {len(problems)}")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
