#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Заполненный шаблон KP_template.xlsx -> kp.json (для build_kp.py).

Запуск:
    uv run xlsx_to_kp.py заполненный.xlsx [--out kp.json]
Затем:
    uv run build_kp.py kp.json && uv run check_kp.py ...
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_template import PAR, REQ, TERMS_N, TERMS_START  # noqa: E402


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    return str(v).replace(" ", "").replace(" ", "").replace(",", ".")


def convert(path: Path) -> dict:
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=False)
    ws, rq, pr = wb["КП"], wb["Реквизиты"], wb["Параметры"]

    company, bank = {}, {}
    for i, (_, key, bkey) in enumerate(REQ, 2):
        v = rq.cell(i, 2).value
        v = "" if v is None else str(v).strip()
        if key and v:
            company[key] = v
        if bkey and v:
            bank[bkey] = v
    if bank:
        company["bank"] = bank

    p = {key: pr.cell(i, 2).value for i, (_, key, _, _) in enumerate(PAR, 2)}
    d = p["date"]
    if isinstance(d, (datetime, date)):
        d = (d.date() if isinstance(d, datetime) else d).isoformat()
    elif isinstance(d, str) and "." in d:
        dd, mm, yy = d.strip().split(".")
        d = f"{yy}-{mm}-{dd}"

    def pct(v, label):  # в шаблоне 16 % хранится как 0.16
        v = num(v)
        if v is None:
            return None
        v = float(v)
        if not 0 <= v <= 1:
            raise SystemExit(f"Параметры: {label} = {v}. В шаблоне процент вводится со знаком %, "
                             "например 16%. Excel считает число 16 как 1600 %, исправьте ячейку.")
        return round(v * 100, 6)

    terms = [str(pr.cell(TERMS_START + i, 2).value).strip()
             for i in range(TERMS_N) if pr.cell(TERMS_START + i, 2).value]

    # строки позиций: после заголовка, где в G «Тип строки»
    hr = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 7).value == "Тип строки")
    sections, current = [], None
    for r in range(hr + 1, ws.max_row + 1):
        t = ws.cell(r, 7).value
        if t not in ("позиция", "раздел"):
            continue
        name = str(ws.cell(r, 2).value or "").strip()
        if t == "раздел":
            current = {"title": name, "items": []}
            sections.append(current)
            continue
        if current is None:
            current = {"title": "Работы", "items": []}
            sections.append(current)
        current["items"].append({
            "name": name, "unit": str(ws.cell(r, 3).value or "").strip(),
            "qty": num(ws.cell(r, 4).value), "price": num(ws.cell(r, 5).value),
            "kind": ws.cell(r, 8).value, "source": str(ws.cell(r, 9).value or "").strip(),
            "_row": r,
        })
    sections = [s for s in sections if s["items"]]

    data = {
        "sample": str(p.get("sample") or "").strip().lower() == "да",
        "number": str(p["number"] or "").strip(),
        "date": d,
        "valid_days": int(float(num(p["valid_days"]) or 0)) or None,
        "currency": "KZT",
        "numbering": "continuous",
        "company": company,
        "client": {"name": str(p["client"] or "").strip()},
        "object": {"name": str(p["object"] or "").strip(), "address": str(p["address"] or "").strip()},
        "intro": str(p["intro"] or "").strip(),
        "rates": {
            "overhead_pct": pct(p["overhead_pct"], "накладные") or 0,
            "profit_pct": pct(p["profit_pct"], "прибыль") or 0,
            "vat_pct": pct(p["vat_pct"], "НДС"),
            "vat_source": str(p["vat_source"] or "").strip(),
            "amount_decimals": int(num(p["amount_decimals"]) or 0),
        },
        "sections": sections,
        "terms": terms,
        "source_file": str(path.resolve()),
    }
    return data


def main() -> int:
    ap = argparse.ArgumentParser(description="Шаблон КП (xlsx) -> kp.json")
    ap.add_argument("xlsx")
    ap.add_argument("--out", help="куда сохранить kp.json (по умолчанию рядом с xlsx)")
    ap.add_argument("--brand", help="фирменный шаблон brand.json для PDF")
    ap.add_argument("--price-list", help="прайс-лист для позиций с кодом")
    a = ap.parse_args()
    src = Path(a.xlsx)
    data = convert(src)
    out = Path(a.out) if a.out else src.with_name("kp.json")
    import os
    for key, val in (("brand", a.brand), ("price_list", a.price_list)):
        if val:
            data[key] = os.path.relpath(Path(val).resolve(), out.resolve().parent)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    n = sum(len(s["items"]) for s in data["sections"])
    print(f"kp.json: {out.resolve()}  (разделов: {len(data['sections'])}, позиций: {n})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
