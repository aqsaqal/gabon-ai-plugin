#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl==3.1.5", "pypdf==6.21.4", "reportlab==5.1.0", "python-docx==1.2.0"]
# ///
"""Независимая проверка готового КП.

1. Каждая сумма в Excel — формула, а не число.
2. Формулы пересчитываются этим скриптом и сверяются с расчётом по исходным данным.
3. Нет ошибок формул (#REF!, #DIV/0! и т. п.), контрольная строка равна 0.
4. PDF открывается, в нём есть номер КП и все итоговые суммы.

Запуск:
    uv run check_kp.py kp.json KP_xxx.xlsx KP_xxx.pdf [KP_xxx.docx]
"""
from __future__ import annotations

import re
import sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from kp_core import KINDS, InputError, compute, fmt_money, load_input, validate  # noqa: E402


import ast
import operator

_OPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
        ast.Div: operator.truediv}


def _safe_eval(expr: str, env: dict):
    """Вычисляет только арифметику над числами и ячейками. Никакого кода из файла."""
    try:
        tree = ast.parse(expr.strip(), mode="eval")
    except SyntaxError:
        raise ValueError(f"Неподдерживаемая формула: {expr}")

    def ev(n):
        if isinstance(n, ast.Expression):
            return ev(n.body)
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)) \
                and not isinstance(n.value, bool):
            return Decimal(str(n.value))
        if isinstance(n, ast.Name) and n.id in env and n.id not in ("_round", "_min", "_max"):
            return env[n.id]
        if isinstance(n, ast.BinOp) and type(n.op) in _OPS:
            return _OPS[type(n.op)](ev(n.left), ev(n.right))
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, (ast.USub, ast.UAdd)):
            v = ev(n.operand)
            return -v if isinstance(n.op, ast.USub) else v
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and n.func.id == "_round" \
                and len(n.args) == 2 and not n.keywords:
            return env["_round"](ev(n.args[0]), ev(n.args[1]))
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) \
                and n.func.id in ("_min", "_max") and n.args and not n.keywords:
            vals = [ev(a) for a in n.args]
            return min(vals) if n.func.id == "_min" else max(vals)
        raise ValueError(f"Неподдерживаемая формула: {expr}")
    return ev(tree)


class FormulaEvaluator:
    """Мини-вычислитель для формул, которые создаёт build_kp.py:
    ссылки на ячейки (в т.ч. на другой лист), + - * /, ROUND, SUM, SUMIF."""

    REF = r"\$?([A-Z]{1,3})\$?(\d+)"

    def __init__(self, wb):
        self.wb = wb
        self.cache: dict[tuple[str, str], Decimal | str] = {}

    @staticmethod
    def _coords(col: str, row: str) -> str:
        return f"{col}{row}"

    def value(self, sheet: str, coord: str):
        key = (sheet, coord)
        if key in self.cache:
            return self.cache[key]
        raw = self.wb[sheet][coord].value
        if isinstance(raw, str) and raw.startswith("="):
            try:
                val = self._eval(sheet, raw[1:])
            except (ValueError, TypeError, ArithmeticError) as e:
                val = f"#ОШИБКА: {e}"
        elif raw is None:
            val = Decimal("0")
        elif isinstance(raw, (int, float)):
            val = Decimal(str(raw))
        else:
            val = raw
        self.cache[key] = val
        return val

    def _range(self, sheet, c1, r1, c2, r2):
        from openpyxl.utils import column_index_from_string, get_column_letter
        out = []
        for ci in range(column_index_from_string(c1), column_index_from_string(c2) + 1):
            for ri in range(int(r1), int(r2) + 1):
                out.append(self.value(sheet, f"{get_column_letter(ci)}{ri}"))
        return out

    def _eval(self, sheet: str, f: str):
        ref = self.REF
        sh = r"(?:'([^']+)'|([A-Za-zА-Яа-яЁё_]+))!"
        calls = {}

        def stash(v):
            k = f"__v{len(calls)}"
            calls[k] = v
            return k

        # SUMIF(диапазон;"критерий";диапазон)
        def sumif(m):
            crit_cells = self._range(sheet, m[1], m[2], m[3], m[4])
            sum_cells = self._range(sheet, m[6], m[7], m[8], m[9])
            total = sum((Decimal(str(s)) if not isinstance(s, Decimal) else s)
                        for c, s in zip(crit_cells, sum_cells) if c == m[5] and isinstance(s, (Decimal, int, float)))
            return stash(Decimal(total))
        f = re.sub(rf"SUMIF\({ref}:{ref},\"([^\"]*)\",{ref}:{ref}\)", sumif, f)

        def sum_(m):
            vals = [v for v in self._range(sheet, m[1], m[2], m[3], m[4]) if isinstance(v, Decimal)]
            return stash(sum(vals, Decimal("0")))
        f = re.sub(rf"SUM\({ref}:{ref}\)", sum_, f)

        f = re.sub(rf"{sh}{ref}", lambda m: stash(self.value(m[1] or m[2], f"{m[3]}{m[4]}")), f)
        f = re.sub(rf"(?<![A-Za-z_]){ref}", lambda m: stash(self.value(sheet, f"{m[1]}{m[2]}")), f)
        f = f.replace("ROUND(", "_round(").replace("MIN(", "_min(").replace("MAX(", "_max(")
        if re.search(r"[A-Z]{2,}\(", f):
            raise ValueError(f"Неподдерживаемая функция в формуле: {f}")

        def _round(x, n):
            return x.quantize(Decimal(1).scaleb(-int(n)), rounding=ROUND_HALF_UP)
        env = {"_round": _round, **calls}
        return _safe_eval(f, env)


def main() -> int:
    if len(sys.argv) not in (4, 5):
        print(__doc__)
        return 2
    inp, xlsx, pdf = (Path(a) for a in sys.argv[1:4])
    docx_path = Path(sys.argv[4]) if len(sys.argv) == 5 else None
    problems: list[str] = []
    ok: list[str] = []

    try:
        data = load_input(inp)
        validate(data)
    except InputError as e:
        print("Исходные данные не проходят проверку:\n  - " + "\n  - ".join(e.errors))
        return 1
    calc = compute(data)

    # ---------- Excel ----------
    from openpyxl import load_workbook
    if not xlsx.exists() or xlsx.stat().st_size == 0:
        problems.append(f"Файл Excel отсутствует или пуст: {xlsx}")
    else:
        wb = load_workbook(xlsx)
        ws = wb["КП"]
        ev = FormulaEvaluator(wb)

        # Позиции: строки, где в колонке G указан вид затрат
        expected_items = [a for s in calc["sections"] for a in s["items"]]
        found = []
        for row in range(1, ws.max_row + 1):
            if ws.cell(row, 7).value in KINDS:
                cell = ws.cell(row, 6)
                if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                    problems.append(f"F{row}: сумма позиции записана числом, а не формулой.")
                found.append((row, ev.value("КП", f"F{row}")))
        if len(found) != len(expected_items):
            problems.append(f"Число позиций в Excel ({len(found)}) не совпадает с исходными "
                            f"данными ({len(expected_items)}).")
        else:
            bad = [(r, v, e) for (r, v), e in zip(found, expected_items) if v != e]
            for r, v, e in bad:
                problems.append(f"F{r}: формула даёт {v}, ожидалось {e}.")
            if not bad:
                ok.append(f"Все {len(found)} позиций: формулы верны, суммы совпадают.")

        checks = [("KP_DIRECT", "Прямые затраты", calc["direct"]),
                  ("KP_OVERHEAD", "Накладные расходы", calc["overhead"]),
                  ("KP_PROFIT", "Прибыль", calc["profit"]),
                  ("KP_SUBTOTAL", "Итого без НДС", calc["subtotal"]),
                  ("KP_VAT", "НДС", calc["vat"]),
                  ("KP_TOTAL", "Всего к оплате", calc["total"]),
                  ("KP_CONTROL", "Контроль по видам затрат", Decimal("0"))]
        for name, label, expected in checks:
            dn = wb.defined_names.get(name)
            if dn is None:
                problems.append(f"Нет именованной ячейки {name} ({label}).")
                continue
            sheet, coord = next(dn.destinations)
            raw = wb[sheet][coord.replace("$", "")].value
            if not (isinstance(raw, str) and raw.startswith("=")):
                problems.append(f"{label}: записано числом, а не формулой.")
            got = ev.value(sheet, coord.replace("$", ""))
            if got != expected:
                problems.append(f"{label}: Excel даёт {got}, ожидалось {expected}.")
            else:
                ok.append(f"{label}: {fmt_money(got)} (формула верна).")

        for wsx in wb.worksheets:
            for row in wsx.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("#"):
                        problems.append(f"{wsx.title}!{c.coordinate}: ошибка {c.value}")

    # ---------- PDF ----------
    from pypdf import PdfReader
    if not pdf.exists() or pdf.stat().st_size == 0:
        problems.append(f"Файл PDF отсутствует или пуст: {pdf}")
    else:
        reader = PdfReader(str(pdf))
        text = "\n".join(p.extract_text() or "" for p in reader.pages)
        flat = re.sub(r"\s+", " ", text.replace(" ", " "))
        ok.append(f"PDF открывается, страниц: {len(reader.pages)}.")
        if str(data["number"]) not in flat:
            problems.append("В PDF не найден номер КП.")
        sums = [("Итого без НДС", calc["subtotal"]), ("Всего", calc["total"])]
        if calc["vat"]:
            sums.append(("НДС", calc["vat"]))
        for label, v in sums:
            if fmt_money(v) not in flat:
                problems.append(f"В PDF не найдена сумма «{label}» {fmt_money(v)}.")
        missing = [fmt_money(a) for s in calc["sections"] for a in s["items"]
                   if fmt_money(a) not in flat]
        if missing:
            problems.append(f"В PDF не найдены суммы позиций: {', '.join(missing[:5])}"
                            + (" и др." if len(missing) > 5 else "") + ".")
        if not problems:
            ok.append("В PDF есть номер КП, суммы всех позиций и итоги, совпадающие с Excel.")

    if docx_path is not None:
        if not docx_path.exists() or docx_path.stat().st_size == 0:
            problems.append(f"Файл Word отсутствует или пуст: {docx_path}")
        else:
            import docx
            dd = docx.Document(str(docx_path))
            dt = " ".join(p.text for p in dd.paragraphs) + " " + " ".join(
                c.text for t in dd.tables for row in t.rows for c in row.cells)
            dt = re.sub(r"\s+", " ", dt.replace("\u00a0", " "))
            need = [fmt_money(a) for s in calc["sections"] for a in s["items"]] + \
                [fmt_money(calc["subtotal"]), fmt_money(calc["total"])] + \
                ([fmt_money(calc["vat"])] if calc["vat"] else [])
            miss = [x for x in need if x not in dt]
            if miss or str(data["number"]) not in dt:
                problems.append(f"В Word не найдены: {', '.join(miss[:5]) or 'номер КП'}.")
            else:
                ok.append("Word: номер КП, суммы всех позиций и итоги совпадают с Excel и PDF.")

    print("ПРОВЕРКА КП")
    for x in ok:
        print(f"  [OK]  {x}")
    for x in problems:
        print(f"  [ОШИБКА] {x}")
    print("РЕЗУЛЬТАТ:", "пройдена" if not problems else f"найдено ошибок: {len(problems)}")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
