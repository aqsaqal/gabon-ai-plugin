#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "reportlab>=4.0", "pypdf>=4.0", "pillow>=10"]
# ///
"""ВОР -> черновик kp.json для навыка kp.

    uv run vor_to_kp.py vor.json --out kp.json --number КП-001
        [--price-list data/prices.xlsx] [--brand data/brand.json] [--date ГГГГ-ММ-ДД]

Переносятся разделы, наименования, единицы и количества. Цены берутся из прайс-листа
по коду строки (поле code). Строки без данных в КП не попадают: они перечисляются в
missing_data. Допущения ВОР переносятся в допущения КП. Ставки и цены не придумываются:
чего нет, то нужно заполнить, и сборка КП об этом сообщит.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from vor_core import VorError, compute, load_vor  # noqa: E402


def convert(vor_path: Path, out: Path, number: str, kp_date: str, price_list: str | None,
            brand: str | None) -> tuple[dict, list[str]]:
    data = load_vor(vor_path)
    values, _ = compute(data)
    notes = []
    rel = lambda p: os.path.relpath(Path(p).resolve(), out.resolve().parent)  # noqa: E731
    vor_dir = Path(data["_input_dir"])
    raw = json.loads(vor_path.read_text(encoding="utf-8"))

    sections, assumptions, missing, exclusions = [], [], [], []
    for si, s in enumerate(data["sections"], 1):
        items = []
        for ii, it in enumerate(s["items"], 1):
            num = f"{si}.{ii}"
            if it["status"] == "нет данных":
                missing.append(f"ВОР {num} «{it['name']}»: {it['missing']}")
                exclusions.append(f"{it['name']}: объём не определён (нет исходных данных), "
                                  "стоимость будет рассчитана дополнительно.")
                continue
            q = values[it["_id"]]
            item = {"name": it["name"], "unit": it["unit"], "qty": str(q),
                    "vor_ref": f"ВОР № {data['number']}, строка {num}"}
            if it.get("code"):
                item["code"] = it["code"]
            else:
                notes.append(f"ВОР {num} «{it['name']}»: нет кода прайса, укажите price, kind "
                             "и source вручную или добавьте code.")
            if it["status"] == "допущение":
                assumptions.append(f"ВОР {num} «{it['name']}»: {it['assumption']}")
            items.append(item)
        if items:
            sections.append({"title": s["title"], "items": items})

    kp = {
        "number": number,
        "date": kp_date,
        "basis": f"Ведомость объёмов работ № {data['number']} от "
                 f"{date.fromisoformat(str(data['date'])):%d.%m.%Y}",
        "object": data["object"],
        "client": raw.get("client") or {"name": ""},
        "rates": {},
        "sections": sections,
        "assumptions": assumptions,
        "missing_data": missing,
        "exclusions": exclusions,
    }
    comp = raw.get("company")
    if isinstance(comp, str):
        kp["company"] = rel(vor_dir / comp)
    elif comp:
        kp["company"] = comp
    if raw.get("sample"):
        kp["sample"] = True
    if price_list:
        kp["price_list"] = rel(price_list)
    if brand:
        kp["brand"] = rel(brand)
    elif raw.get("brand"):
        kp["brand"] = rel(vor_dir / raw["brand"])
    if not kp.get("company"):
        notes.append("Нет реквизитов компании: укажите company в kp.json.")
    notes.append("Ставки (rates): накладные, прибыль, НДС с источником берутся из фирменного "
                 "шаблона, если они там заданы, иначе заполните их в kp.json.")
    return kp, notes


def main() -> int:
    ap = argparse.ArgumentParser(description="ВОР -> черновик КП")
    ap.add_argument("vor")
    ap.add_argument("--out", required=True)
    ap.add_argument("--number", required=True, help="номер КП")
    ap.add_argument("--date", default=date.today().isoformat())
    ap.add_argument("--price-list")
    ap.add_argument("--brand")
    a = ap.parse_args()
    out = Path(a.out)
    if out.exists():
        print(f"Файл уже существует, не перезаписываю: {out}")
        return 1
    try:
        kp, notes = convert(Path(a.vor), out, a.number, a.date, a.price_list, a.brand)
    except VorError as e:
        print("ВОР содержит ошибки, сначала исправьте их:\n  - " + "\n  - ".join(e.errors))
        return 2
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(kp, ensure_ascii=False, indent=2), encoding="utf-8")
    n = sum(len(s["items"]) for s in kp["sections"])
    print(f"Черновик КП: {out.resolve()} (позиций: {n}, без данных: {len(kp['missing_data'])}, "
          f"с допущениями: {len(kp['assumptions'])})")
    for x in notes:
        print(f"  - {x}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
