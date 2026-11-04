#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl==3.1.5", "reportlab==5.1.0", "pypdf==6.21.4", "pillow==13.0.0"]
# ///
"""Согласованное КП -> contract.json с замороженными договорными ценами.

    uv run make_contract.py kp.json --out contract.json --number 15/2026 --date 2026-11-10

Цены фиксируются на дату договора: дальнейшие изменения прайс-листа на акты не влияют.
Если в КП накладные и прибыль начислены отдельно, они включаются в единичные цены
(договорная цена «под ключ» без НДС, округление до тиын). Разница округления с итогом КП
выводится: её нужно согласовать с заказчиком.
"""
from __future__ import annotations

import argparse
import json
import sys
from decimal import Decimal
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1] / "kp" / "scripts"))
from kp_core import InputError, compute, fmt_money, item_numbers, load_input, money, validate  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description="КП -> договор (замороженные цены)")
    ap.add_argument("kp")
    ap.add_argument("--out", required=True)
    ap.add_argument("--number", required=True)
    ap.add_argument("--date", required=True)
    a = ap.parse_args()
    out = Path(a.out)
    if out.exists():
        print(f"Файл уже существует, не перезаписываю: {out}. Договор меняется только доп. соглашением.")
        return 1
    try:
        data = load_input(a.kp)
        warnings = validate(data)
    except InputError as e:
        print("КП с ошибками:\n  - " + "\n  - ".join(e.errors))
        return 2
    calc = compute(data)
    n = int(data["rates"].get("amount_decimals", 2))
    factor = calc["subtotal"] / calc["direct"] if calc["direct"] else Decimal(1)
    nums = item_numbers(data)
    positions, sub = [], Decimal(0)
    for si, (s, cs) in enumerate(zip(data["sections"], calc["sections"])):
        for ii, (it, amt) in enumerate(zip(s["items"], cs["items"])):
            price = money(Decimal(str(it["price"])) * factor, 2)
            q = Decimal(str(it["qty"]))
            positions.append({"n": nums[si][ii], "section": s["title"], "name": it["name"],
                              "unit": it["unit"], "qty": str(q), "price": str(price),
                              "kp_source": it.get("source", "")})
            sub += money(q * price, n)
    vat = money(sub * Decimal(str(data["rates"]["vat_pct"])) / 100, n)
    comp = data["company"]
    contract = {
        "number": a.number, "date": a.date,
        "subject": data["object"]["name"], "address": data["object"].get("address", ""),
        "basis_kp": f"КП № {data['number']} от {data['date']}",
        "customer": {"name": data.get("client", {}).get("name", ""), "bin": "", "address": "",
                     "phone": "", "signer_position": "", "signer_name": ""},
        "contractor": {"name": comp.get("name", ""), "bin": comp.get("bin", ""),
                       "address": comp.get("address", ""), "phone": comp.get("phone", ""),
                       "signer_position": comp.get("director_position", "Директор"),
                       "signer_name": comp.get("director", "")},
        "vat_pct": data["rates"]["vat_pct"], "vat_source": data["rates"]["vat_source"],
        "amount_decimals": n, "positions": positions,
        "totals": {"subtotal": str(sub), "vat": str(vat), "total": str(sub + vat)},
        "sample": bool(data.get("sample")),
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(contract, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Договор: {out.resolve()} (позиций: {len(positions)})")
    print(f"Договорная цена: без НДС {fmt_money(sub)}, НДС {fmt_money(vat)}, всего {fmt_money(sub + vat)}")
    diff = sub + vat - calc["total"]
    if diff:
        print(f"  - Разница округления с итогом КП ({fmt_money(calc['total'])}): {fmt_money(diff)}. "
              "Согласуйте договорную цену с заказчиком.")
    if not contract["customer"]["name"]:
        print("  - Заказчик не указан в КП: заполните customer.name.")
    print("  - Заполните реквизиты заказчика (БИН, адрес, телефон, подписант) в contract.json.")
    if "advance" not in contract:
        print("  - Если по договору есть аванс или гарантийное удержание, добавьте advance и retention "
              "с пунктом договора.")
    for w in warnings:
        print(f"  - КП: {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
