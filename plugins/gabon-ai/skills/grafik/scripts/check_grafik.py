#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "pypdf>=4.0", "reportlab>=4.0"]
# ///
"""Независимая проверка графика: формулы Excel (WORKDAY.INTL, ROUNDUP, MAX) пересчитываются
и сверяются с расчётом, PDF сверяется по датам каждой работы.

    uv run check_grafik.py grafik.json GRAFIK_xxx.xlsx GRAFIK_xxx.pdf
"""
from __future__ import annotations

import ast
import math
import re
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from grafik_core import GrafikError, compute, load_grafik, workday  # noqa: E402
from kp_core import fmt_date  # noqa: E402

EPOCH = date(1899, 12, 30)


def serial(d: date) -> int:
    return (d - EPOCH).days


class Evaluator:
    """Вычисляет только формулы графика. Никакого кода из файла не исполняет."""

    REF = r"\$?([A-Z]{1,3})\$?(\d+)"
    SH = r"(?:'([^']+)'|([A-Za-zА-Яа-яЁё_]+))!"

    def __init__(self, wb):
        self.wb, self.cache = wb, {}

    def value(self, sheet, coord):
        key = (sheet, coord)
        if key not in self.cache:
            raw = self.wb[sheet][coord].value
            if isinstance(raw, str) and raw.startswith("="):
                v = self.eval(sheet, raw[1:])
            elif isinstance(raw, datetime):
                v = Decimal(serial(raw.date()))
            elif isinstance(raw, date):
                v = Decimal(serial(raw))
            elif isinstance(raw, (int, float)):
                v = Decimal(str(raw))
            else:
                v = raw
            self.cache[key] = v
        return self.cache[key]

    def eval(self, sheet, f):
        env, k = {}, [0]

        def stash(v):
            k[0] += 1
            env[f"_v{k[0]}"] = v
            return f"_v{k[0]}"
        f = f.replace("_xlfn.", "")
        f = re.sub(rf"{self.SH}{self.REF}:{self.REF}", lambda m: stash(
            [self.value(m[1] or m[2], f"{m[3]}{r}") for r in range(int(m[4]), int(m[6]) + 1)]), f)
        f = re.sub(rf"{self.SH}{self.REF}", lambda m: stash(self.value(m[1] or m[2], f"{m[3]}{m[4]}")), f)
        f = re.sub(rf"(?<![\w.]){self.REF}(?![\w(])", lambda m: stash(self.value(sheet, f"{m[1]}{m[2]}")), f)
        f = f.replace("WORKDAY.INTL(", "WD(")
        tree = ast.parse(f, mode="eval")
        return self._ev(tree.body, env)

    def _ev(self, n, env):
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float, str)) \
                and not isinstance(n.value, bool):
            return Decimal(str(n.value)) if not isinstance(n.value, str) else n.value
        if isinstance(n, ast.Name) and n.id in env:
            return env[n.id]
        if isinstance(n, ast.BinOp) and isinstance(n.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
            a, b = self._ev(n.left, env), self._ev(n.right, env)
            return {ast.Add: a + b, ast.Sub: a - b, ast.Mult: a * b}.get(type(n.op)) \
                if not isinstance(n.op, ast.Div) else a / b
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and not n.keywords:
            args = [self._ev(x, env) for x in n.args]
            fn = n.func.id
            if fn == "MAX":
                return max(args)
            if fn == "ROUND":
                return args[0].quantize(Decimal(1).scaleb(-int(args[1])), rounding=ROUND_HALF_UP)
            if fn == "ROUNDUP":
                q = Decimal(1).scaleb(-int(args[1]))
                return Decimal(math.ceil(args[0] / q)) * q
            if fn == "DATE":
                return Decimal(serial(date(int(args[0]), int(args[1]), int(args[2]))))
            if fn == "WD":
                start = EPOCH + timedelta(days=int(args[0]))
                hol = frozenset(EPOCH + timedelta(days=int(v)) for v in (args[3] if len(args) > 3 else [])
                                if isinstance(v, Decimal))
                return Decimal(serial(workday(start, int(args[1]), args[2], hol)))
        raise ValueError(f"неподдерживаемая формула: {ast.dump(n)[:80]}")


def main() -> int:
    if len(sys.argv) != 4:
        print(__doc__)
        return 2
    inp, xlsx, pdf = (Path(a) for a in sys.argv[1:])
    try:
        data = load_grafik(inp)
        res, _ = compute(data)
    except GrafikError as e:
        print("График не проходит проверку:\n  - " + "\n  - ".join(e.errors))
        return 1
    ok, problems = [], []
    from openpyxl import load_workbook
    wb = load_workbook(xlsx)
    ws = wb["График"]
    ev = Evaluator(wb)
    hr = next(r for r in range(1, 20) if ws.cell(r, 1).value == "№")
    tasks = data["tasks"]
    for i, t in enumerate(tasks):
        r = hr + 1 + i
        if str(ws.cell(r, 1).value) != t["_n"]:
            problems.append(f"Строка {r}: ожидалась работа № {t['_n']}.")
            continue
        for col in "JK":
            raw = ws[f"{col}{r}"].value
            if not (isinstance(raw, str) and raw.startswith("=")):
                problems.append(f"{col}{r}: дата записана значением, а не формулой.")
        if t["_mode"] in ("productivity", "labor") and not str(ws[f"I{r}"].value).startswith("="):
            problems.append(f"I{r}: длительность должна считаться формулой.")
        try:
            dur, es, ef = (ev.value("График", f"{c}{r}") for c in "IJK")
        except Exception as e:  # noqa: BLE001
            problems.append(f"Строка {r} «{t['name']}»: формула не вычисляется ({e}).")
            continue
        if (int(dur), int(es), int(ef)) != (t["_dur"], serial(t["_es"]), serial(t["_ef"])):
            problems.append(f"Строка {r} «{t['name']}»: Excel {int(dur)} дн., "
                            f"{fmt_date(EPOCH + timedelta(days=int(es)))}–"
                            f"{fmt_date(EPOCH + timedelta(days=int(ef)))}; ожидалось {t['_dur']} дн., "
                            f"{fmt_date(t['_es'])}–{fmt_date(t['_ef'])}.")
    if not problems:
        ok.append(f"Все {len(tasks)} работ: длительности и даты в Excel совпадают с расчётом.")

    from pypdf import PdfReader
    reader = PdfReader(str(pdf))
    text = re.sub(r"\s+", " ", " ".join(p.extract_text() or "" for p in reader.pages))
    miss = [t["_n"] for t in tasks if fmt_date(t["_es"]) not in text or fmt_date(t["_ef"]) not in text]
    if miss or str(data["number"]) not in text:
        problems.append(f"В PDF нет дат работ: {', '.join(miss) or 'номер графика'}.")
    else:
        ok.append(f"PDF ({len(reader.pages)} стр.): даты всех работ совпадают.")

    print("ПРОВЕРКА ГРАФИКА")
    for x in ok:
        print(f"  [OK]  {x}")
    for x in problems:
        print(f"  [ОШИБКА] {x}")
    print("РЕЗУЛЬТАТ:", "пройдена" if not problems else f"найдено ошибок: {len(problems)}")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
