#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl==3.1.5", "reportlab==5.1.0", "pypdf==6.21.4", "pillow==13.0.0"]
# ///
"""Сравнение сметы (КП) подрядчика с нашей сметой или КП.

    uv run compare.py --ours smeta.json|kp.json --contractor contractor.json --out папка
        [--basis all_in|direct] [--price-tol 10] [--qty-tol 2] [--vat 16]

Что делает:
- проверяет арифметику подрядчика: кол-во × цена, сумма строк, НДС, итог, признаки подгонки;
- сопоставляет позиции: по полю match, по коду, иначе по похожести наименования при той же ед.;
- считает отклонения по количеству, цене и сумме; перечисляет лишние и пропущенные позиции.
По умолчанию (basis all_in) наши цены приводятся к «под ключ» без НДС: накладные и прибыль
распределяются пропорционально прямым затратам, так как подрядчики обычно дают полные цены.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from decimal import Decimal
from difflib import SequenceMatcher
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from smeta_core import SmetaError, compute as smeta_compute, load_smeta  # noqa: E402
from brand_core import brand_for  # noqa: E402
from build_kp import register_fonts, safe_name  # noqa: E402
from check_kp import FormulaEvaluator  # noqa: E402
from kp_core import (D, InputError, compute as kp_compute, fmt_money, fmt_qty,  # noqa: E402
                     load_input, money, normalize_unit, validate)


def norm_name(s: str) -> str:
    s = str(s).lower().replace("ё", "е")
    s = re.sub(r"[«»\"'(),.;:]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def load_ours(path: Path, basis: str) -> tuple[dict, list[dict]]:
    raw = json.loads(path.read_text(encoding="utf-8"))
    rows = []
    if any(isinstance(s, dict) and "positions" in s for s in raw.get("sections") or []):
        data = load_smeta(path)
        calc, _ = smeta_compute(data)
        kind = "смета"
        for si, s in enumerate(calc["sections"], 1):
            for pi, p in enumerate(s["positions"], 1):
                rows.append({"n": f"{si}.{pi}", "name": p["name"], "unit": p["unit"],
                             "qty": p["qty"], "direct": p["total"], "code": p.get("code") or ""})
    else:
        data = load_input(path)
        validate(data)
        calc = kp_compute(data)
        kind = "КП"
        for si, (s, cs) in enumerate(zip(data["sections"], calc["sections"]), 1):
            for ii, (it, amt) in enumerate(zip(s["items"], cs["items"]), 1):
                rows.append({"n": f"{si}.{ii}", "name": it["name"], "unit": normalize_unit(it["unit"]),
                             "qty": D(it["qty"]), "direct": amt, "code": it.get("code") or ""})
    factor = calc["subtotal"] / calc["direct"] if basis == "all_in" and calc["direct"] else Decimal(1)
    for r in rows:
        r["amount"] = money(r["direct"] * factor)
    target = calc["subtotal"] if basis == "all_in" else calc["direct"]
    rest = target - sum((r["amount"] for r in rows), Decimal(0))
    if rows and rest:  # остаток округления — на самую крупную строку, чтобы итог совпал
        max(rows, key=lambda r: r["amount"])["amount"] += rest
    for r in rows:
        r["price"] = money(r["amount"] / r["qty"])
    meta = {"kind": kind, "number": data["number"], "subtotal": calc["subtotal"],
            "direct": calc["direct"], "factor": factor, "vat_pct": D(data["rates"]["vat_pct"]),
            "data": data}
    return meta, rows


def check_contractor(c: dict, vat_pct: Decimal) -> tuple[list[dict], list[str], dict]:
    """Арифметика подрядчика. Возвращает (позиции, замечания, итоги)."""
    findings, pos = [], []
    seen = {}
    for p in c["positions"]:
        q, pr = D(p["qty"]) if p.get("qty") else None, D(p["price"]) if p.get("price") else None
        am = D(p["amount"]) if p.get("amount") else None
        e = {**p, "qty": q, "price": pr, "amount": am, "unit": normalize_unit(p.get("unit")),
             "issues": []}
        if q is not None and pr is not None:
            calc_am = q * pr
            e["true_amount"] = money(calc_am)
            if am is None:
                e["amount"] = money(calc_am)
            elif abs(am - calc_am) > Decimal("1"):
                e["issues"].append(f"сумма {fmt_money(am)} не равна кол-во × цена "
                                   f"{fmt_money(calc_am)} (разница {fmt_money(am - calc_am)})")
        key = (norm_name(p["name"]), e["unit"])
        if key in seen:
            e["issues"].append(f"повтор позиции строки {seen[key]}")
        seen[key] = p.get("row", "")
        for x in e["issues"]:
            findings.append(f"Строка {p.get('row', '?')} «{p['name']}»: {x}.")
        pos.append(e)
    s_lines = sum((p["amount"] for p in pos if p["amount"] is not None), Decimal(0))
    t = {k: D(v) for k, v in (c.get("totals") or {}).items() if v not in (None, "")}
    if "subtotal" in t and abs(t["subtotal"] - s_lines) > Decimal("1"):
        findings.append(f"Итого без НДС {fmt_money(t['subtotal'])} не равно сумме строк "
                        f"{fmt_money(s_lines)} (разница {fmt_money(t['subtotal'] - s_lines)}).")
    sub = t.get("subtotal", s_lines)
    if "vat" in t and abs(t["vat"] - sub * vat_pct / 100) > Decimal("1"):
        findings.append(f"НДС {fmt_money(t['vat'])} не равен {vat_pct} % от итога без НДС "
                        f"({fmt_money(sub * vat_pct / 100)}).")
    if "total" in t:
        vat = t.get("vat", money(sub * vat_pct / 100))
        if abs(t["total"] - (sub + vat)) > Decimal("1"):
            findings.append(f"Всего {fmt_money(t['total'])} не равно итог + НДС "
                            f"{fmt_money(sub + vat)}.")
        true_lines = sum((p.get("true_amount", p["amount"]) for p in pos
                          if p["amount"] is not None), Decimal(0))
        honest = true_lines + money(true_lines * vat_pct / 100)
        if t["total"] % 100000 == 0 and abs(t["total"] - honest) > Decimal("1"):
            findings.append(f"Признак подгонки: итог круглый ({fmt_money(t['total'])}), а пересчёт "
                            f"строк (кол-во × цена) с НДС даёт {fmt_money(honest)} (разница "
                            f"{fmt_money(t['total'] - honest)}).")
    return pos, findings, {"lines": s_lines, **t}


def match(ours: list[dict], theirs: list[dict], threshold: float = 0.55) -> list[tuple]:
    """Пары (наша строка или None, их строка или None, оценка, способ)."""
    pairs, used_o, used_t = [], set(), set()
    by_n = {o["n"]: i for i, o in enumerate(ours)}
    for ti, t in enumerate(theirs):
        if t.get("match") and t["match"] in by_n and by_n[t["match"]] not in used_o:
            oi = by_n[t["match"]]
            pairs.append((oi, ti, 1.0, "вручную"))
            used_o.add(oi)
            used_t.add(ti)
    for ti, t in enumerate(theirs):
        if ti in used_t or not t.get("code"):
            continue
        for oi, o in enumerate(ours):
            if oi not in used_o and o["code"] and o["code"] == t["code"]:
                pairs.append((oi, ti, 1.0, "по коду"))
                used_o.add(oi)
                used_t.add(ti)
                break
    cands = []
    for ti, t in enumerate(theirs):
        if ti in used_t:
            continue
        for oi, o in enumerate(ours):
            if oi in used_o or (t["unit"] and o["unit"] != t["unit"]):
                continue
            sc = SequenceMatcher(None, norm_name(o["name"]), norm_name(t["name"])).ratio()
            if sc >= threshold:
                cands.append((sc, oi, ti))
    for sc, oi, ti in sorted(cands, reverse=True):
        if oi not in used_o and ti not in used_t:
            pairs.append((oi, ti, round(sc, 2), "по названию"))
            used_o.add(oi)
            used_t.add(ti)
    pairs += [(oi, None, 0, "") for oi in range(len(ours)) if oi not in used_o]
    pairs += [(None, ti, 0, "") for ti in range(len(theirs)) if ti not in used_t]
    return sorted(pairs, key=lambda p: (p[0] is None, p[0] if p[0] is not None else p[1]))


def pct(a, b):
    return None if a is None or b in (None, 0) else (a - b) / b * 100


def main() -> int:
    ap = argparse.ArgumentParser(description="Сравнение сметы подрядчика с нашей")
    ap.add_argument("--ours", required=True)
    ap.add_argument("--contractor", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--basis", choices=("all_in", "direct"), default="all_in")
    ap.add_argument("--price-tol", type=float, default=10.0, help="порог отклонения цены, %")
    ap.add_argument("--qty-tol", type=float, default=2.0, help="порог отклонения объёма, %")
    ap.add_argument("--vat", type=float, help="ставка НДС подрядчика, по умолчанию как у нас")
    a = ap.parse_args()
    try:
        meta, ours = load_ours(Path(a.ours), a.basis)
    except (SmetaError, InputError) as e:
        print("Наша смета/КП с ошибками:\n  - " + "\n  - ".join(e.errors))
        return 2
    c = json.loads(Path(a.contractor).read_text(encoding="utf-8"))
    vat_pct = D(a.vat) if a.vat is not None else meta["vat_pct"]
    theirs, findings, ctot = check_contractor(c, vat_pct)
    if c.get("prices_include_vat"):
        for t in theirs:
            for k in ("price", "amount"):
                if t[k] is not None:
                    t[k] = money(t[k] / (1 + vat_pct / 100))
        findings.append("Цены подрядчика указаны с НДС: для сравнения пересчитаны без НДС.")
    pairs = match(ours, theirs)

    rows = []
    for oi, ti, sc, how in pairs:
        o = ours[oi] if oi is not None else None
        t = theirs[ti] if ti is not None else None
        notes = []
        if o and t:
            dq, dp = pct(t["qty"], o["qty"]), pct(t["price"], o["price"])
            if dq is not None and abs(dq) > a.qty_tol:
                notes.append(f"объём {dq:+.1f} %")
            if dp is not None and abs(dp) > a.price_tol:
                notes.append(f"цена {dp:+.1f} %")
            if how == "по названию" and sc < 0.75:
                notes.append("проверьте сопоставление")
            notes += t["issues"]
        elif o:
            notes.append("нет у подрядчика")
        else:
            notes.append("нет у нас: обоснуйте или исключите")
            notes += t["issues"]
        rows.append({"o": o, "t": t, "score": sc, "how": how, "notes": notes,
                     "diff": (t["amount"] if t and t["amount"] is not None else Decimal(0))
                     - (o["amount"] if o else Decimal(0))})

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    base = out / f"COMPARE_{safe_name(str(meta['number']))}"
    xlsx, pdf = base.with_suffix(".xlsx"), base.with_suffix(".pdf")
    our_total = sum((o["amount"] for o in ours), Decimal(0))
    their_total = ctot.get("subtotal", ctot["lines"])
    write_xlsx(xlsx, rows, findings, meta, c, our_total, their_total, a)
    self_check(xlsx, rows)
    register_fonts(*(brand_for(meta["data"])[k] for k in ("font_regular", "font_bold")))
    write_pdf(pdf, rows, findings, meta, c, our_total, their_total, a)

    print(f"Excel: {xlsx}\nPDF:   {pdf}")
    print(f"Наш итог без НДС ({'с накладными и прибылью' if a.basis == 'all_in' else 'прямые затраты'}): "
          f"{fmt_money(our_total)}")
    print(f"Подрядчик, сумма строк без НДС: {fmt_money(ctot['lines'])}")
    print(f"Подрядчик, заявленный итог без НДС: {fmt_money(their_total)}")
    print(f"Разница: {fmt_money(their_total - our_total)} ({pct(their_total, our_total):+.1f} %)")
    print(f"Сопоставлено: {sum(1 for r in rows if r['o'] and r['t'])}, нет у подрядчика: "
          f"{sum(1 for r in rows if not r['t'])}, лишних у подрядчика: "
          f"{sum(1 for r in rows if not r['o'])}")
    if findings:
        print("Проверка документа подрядчика:")
        for f in findings:
            print(f"  - {f}")
    return 0


HEAD = ["№ наш", "Наименование (наше)", "Ед.", "Кол-во наше", "Цена наша", "Сумма наша",
        "№ подр.", "Наименование (подрядчик)", "Кол-во подр.", "Цена подр.", "Сумма подр.",
        "Откл. объёма, %", "Откл. цены, %", "Разница суммы", "Сопоставление", "Замечания"]


def write_xlsx(path, rows, findings, meta, c, our_total, their_total, a):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.workbook.defined_name import DefinedName
    wb = Workbook()
    ws = wb.active
    ws.title = "Сравнение"
    ws["A1"] = f"Сравнение: {meta['kind']} № {meta['number']} и документ подрядчика " \
               f"{c.get('contractor') or ''} ({Path(c.get('source_file', '')).name})"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = ("Наши цены приведены к полным без НДС: накладные и прибыль распределены "
                f"пропорционально (коэффициент {meta['factor']:.4f})." if a.basis == "all_in"
                else "Наши цены: прямые затраты без накладных и прибыли.")
    ws.append([])
    ws.append(HEAD)
    hr = ws.max_row
    for cell in ws[hr]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DCE3EE")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    first = hr + 1
    red, yel = PatternFill("solid", fgColor="F8CBAD"), PatternFill("solid", fgColor="FFF2CC")
    for i, r in enumerate(rows):
        n = first + i
        o, t = r["o"], r["t"]
        ws.append([o["n"] if o else "", o["name"] if o else "", (o or t)["unit"],
                   float(o["qty"]) if o else None, float(o["price"]) if o else None,
                   float(o["amount"]) if o else None,
                   (t.get("n") or t.get("row")) if t else "", t["name"] if t else "",
                   float(t["qty"]) if t and t["qty"] is not None else None,
                   float(t["price"]) if t and t["price"] is not None else None,
                   float(t["amount"]) if t and t["amount"] is not None else None,
                   f"=ROUND((I{n}-D{n})/D{n}*100,1)" if o and t and t["qty"] is not None else None,
                   f"=ROUND((J{n}-E{n})/E{n}*100,1)" if o and t and t["price"] is not None else None,
                   f"=K{n}-F{n}" if o and t else (f"=K{n}" if t else f"=-F{n}"),
                   f"{r['how']} {r['score']}" if r["how"] else "", "; ".join(r["notes"])])
        fill = red if not (o and t) or (t and t["issues"]) else (yel if r["notes"] else None)
        if fill:
            for cell in ws[n]:
                cell.fill = fill
        for col in "EFJKN":
            ws[f"{col}{n}"].number_format = "#,##0.00"
    last = ws.max_row
    ws.append([])
    ws.append(["", "ИТОГО без НДС", "", "", "", f"=SUM(F{first}:F{last})", "", "", "", "",
               f"=SUM(K{first}:K{last})", "", "", f"=K{last + 2}-F{last + 2}"])
    tr = ws.max_row
    for col in "FKN":
        ws[f"{col}{tr}"].number_format = "#,##0.00"
        ws[f"{col}{tr}"].font = Font(bold=True)
    wb.defined_names["CMP_OURS"] = DefinedName("CMP_OURS", attr_text=f"'Сравнение'!$F${tr}")
    wb.defined_names["CMP_THEIRS"] = DefinedName("CMP_THEIRS", attr_text=f"'Сравнение'!$K${tr}")
    for col, w in zip("ABCDEFGHIJKLMNOP", (7, 40, 7, 11, 13, 15, 7, 40, 11, 13, 15, 10, 10, 15, 16, 50)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = ws.cell(first, 1)
    fs = wb.create_sheet("Проверка подрядчика")
    fs.column_dimensions["A"].width = 130
    fs.append(["Замечания к документу подрядчика"])
    fs["A1"].font = Font(bold=True)
    for f in findings or ["Арифметических ошибок не найдено."]:
        fs.append([f"• {f}"])
    wb.save(path)


def self_check(path, rows):
    """Пересчитывает формулы отчёта и сверяет с расчётом скрипта."""
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ev = FormulaEvaluator(wb)
    ws = wb["Сравнение"]
    first = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "№ наш") + 1
    for i, r in enumerate(rows):
        got = ev.value("Сравнение", f"N{first + i}")
        exp = r["diff"]
        if not isinstance(got, Decimal) or abs(got - exp) > Decimal("0.01"):
            raise SystemExit(f"Самопроверка отчёта: строка {first + i}, разница {got} вместо {exp}.")


def write_pdf(path, rows, findings, meta, c, our_total, their_total, a):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    def esc(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    st = {"n": ParagraphStyle("n", fontName="KP", fontSize=8.5, leading=11),
          "s": ParagraphStyle("s", fontName="KP", fontSize=7, leading=8.8),
          "r": ParagraphStyle("r", fontName="KP", fontSize=7, leading=8.8, alignment=TA_RIGHT),
          "h": ParagraphStyle("h", fontName="KP-Bold", fontSize=7, leading=8.8, alignment=TA_CENTER,
                              textColor=colors.white),
          "t": ParagraphStyle("t", fontName="KP-Bold", fontSize=13, leading=17),
          "h2": ParagraphStyle("h2", fontName="KP-Bold", fontSize=10, leading=13,
                               spaceBefore=4 * mm, spaceAfter=1.5 * mm)}
    story = [Paragraph(f"Сравнение с документом подрядчика: {esc(meta['kind'])} № {esc(meta['number'])}",
                       st["t"]),
             Paragraph(esc(f"Подрядчик: {c.get('contractor') or 'не указан'}. Файл: "
                           f"{Path(c.get('source_file', '')).name}."), st["n"]),
             Paragraph(esc("Наши цены приведены к полным без НДС (накладные и прибыль распределены "
                           "пропорционально)." if a.basis == "all_in" else
                           "Наши цены: прямые затраты."), st["n"]), Spacer(1, 3 * mm)]
    d = their_total - our_total
    lines = sum((t["amount"] for r in rows if (t := r["t"]) and t["amount"] is not None), Decimal(0))
    summ = Table([["Наш итог без НДС", fmt_money(our_total)],
                  ["Подрядчик, сумма строк без НДС", fmt_money(lines)],
                  ["Подрядчик, заявленный итог без НДС", fmt_money(their_total)],
                  ["Разница", f"{fmt_money(d)} ({pct(their_total, our_total):+.1f} %)"]],
                 colWidths=[70 * mm, 50 * mm])
    summ.setStyle(TableStyle([("FONTNAME", (0, 0), (-1, -1), "KP"), ("FONTSIZE", (0, 0), (-1, -1), 9),
                              ("FONTNAME", (0, 3), (-1, 3), "KP-Bold"), ("ALIGN", (1, 0), (1, -1), "RIGHT")]))
    story.append(summ)
    if findings:
        story.append(Paragraph("Ошибки и замечания в документе подрядчика", st["h2"]))
        story += [Paragraph(f"• {esc(f)}", st["n"]) for f in findings]
    story.append(Paragraph("Позиции", st["h2"]))
    tbl = [[Paragraph(h, st["h"]) for h in ("№", "Наименование (наше / подрядчик)", "Ед.",
                                            "Кол-во наше", "Кол-во подр.", "Цена наша", "Цена подр.",
                                            "Разница суммы", "Замечания")]]
    style = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3864")),
             ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]
    for r in sorted(rows, key=lambda r: -abs(r["diff"])):
        o, t = r["o"], r["t"]
        name = (o["name"] if o else "—") + ("" if not t or (o and t["name"] == o["name"])
                                           else f" / {t['name']}")
        tbl.append([Paragraph(o["n"] if o else "+", st["s"]), Paragraph(esc(name), st["s"]),
                    Paragraph(esc((o or t)["unit"]), st["s"]),
                    Paragraph(fmt_qty(o["qty"]) if o else "—", st["r"]),
                    Paragraph(fmt_qty(t["qty"]) if t and t["qty"] is not None else "—", st["r"]),
                    Paragraph(fmt_money(o["price"]) if o else "—", st["r"]),
                    Paragraph(fmt_money(t["price"]) if t and t["price"] is not None else "—", st["r"]),
                    Paragraph(fmt_money(r["diff"]), st["r"]),
                    Paragraph(esc("; ".join(r["notes"])), st["s"])])
        if r["notes"]:
            bad = not (o and t) or (t and t["issues"])
            style.append(("BACKGROUND", (0, len(tbl) - 1), (-1, len(tbl) - 1),
                          colors.HexColor("#F8CBAD" if bad else "#FFF2CC")))
    T = Table(tbl, colWidths=[x * mm for x in (10, 80, 12, 20, 20, 24, 24, 26, 61)], repeatRows=1)
    T.setStyle(TableStyle(style))
    story += [T, Spacer(1, 4 * mm),
              Paragraph("Строки отсортированы по величине разницы. Отчёт подготовлен для проверки "
                        "ответственным специалистом.", st["s"])]
    SimpleDocTemplate(str(path), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                      topMargin=12 * mm, bottomMargin=12 * mm,
                      title=f"Сравнение {meta['number']}").build(story)


if __name__ == "__main__":
    sys.exit(main())
