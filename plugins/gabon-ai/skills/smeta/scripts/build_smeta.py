#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl==3.1.5", "reportlab==5.1.0", "pypdf==6.21.4", "pillow==13.0.0"]
# ///
"""Локальный сметный расчёт: Excel с формулами + PDF.

    uv run build_smeta.py projects/<объект>/input/smeta.json [--out папка]
"""
from __future__ import annotations

import argparse
import sys
from decimal import Decimal
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from smeta_core import (OVERHEAD_BASES, PROFIT_BASES, RES_PLACES, TYPES, SmetaError,  # noqa: E402
                        compute, load_smeta)
from brand_core import brand_for  # noqa: E402
from build_kp import merge_letterhead, register_fonts, safe_name  # noqa: E402
from kp_core import D, amount_in_words, fmt_date, fmt_money, fmt_pct, fmt_qty  # noqa: E402

TITLE = "ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЁТ"
SUBTITLE = "коммерческий расчёт ресурсным методом"
NAMES = {"труд": "SM_LABOR", "машины": "SM_MACH", "материал": "SM_MAT", "транспорт": "SM_TRANS"}
PARAMS = "'Параметры'"


def aggregate(calc: dict) -> list[dict]:
    """Ресурсная ведомость: одинаковые ресурсы (тип, наименование, ед., цена) суммируются."""
    agg: dict[tuple, dict] = {}
    for s in calc["sections"]:
        for p in s["positions"]:
            for r in p["resources"]:
                k = (r["type"], r["name"], r["unit"], r["price"])
                a = agg.setdefault(k, {"type": r["type"], "name": r["name"], "unit": r["unit"],
                                       "price": r["price"], "qty": Decimal(0),
                                       "amount": Decimal(0), "rows": []})
                a["qty"] += r["qty"]
                a["amount"] += r["amount"]
    order = list(TYPES)
    return sorted(agg.values(), key=lambda a: (order.index(a["type"]), a["name"]))


def build_xlsx(data: dict, calc: dict, path: Path, warnings: list[str]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.workbook.defined_name import DefinedName

    br = brand_for(data)
    C = {k: v.lstrip("#") for k, v in br["colors"].items()}
    n = calc["places"]
    rates = data["rates"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Смета"
    rs = wb.create_sheet("Ресурсы")
    ps = wb.create_sheet("Параметры")
    ns = wb.create_sheet("Замечания")
    bold = Font(bold=True)
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    MONEY, Q = "#,##0.00", "#,##0.000"

    ps.append(["Параметр", "Значение", "Основание / база"])
    ps.append(["Накладные расходы, %", float(D(rates["overhead_pct"]) / 100),
               f"{OVERHEAD_BASES[rates.get('overhead_base', 'direct')]}. "
               f"{rates.get('overhead_source', '')}"])
    ps.append(["Сметная прибыль, %", float(D(rates["profit_pct"]) / 100),
               f"{PROFIT_BASES[rates.get('profit_base', 'direct_overhead')]}. "
               f"{rates.get('profit_source', '')}"])
    ps.append(["НДС, %", float(D(rates["vat_pct"]) / 100), rates["vat_source"]])
    for c in ("B2", "B3", "B4"):
        ps[c].number_format = "0.00%"
    for c in ps[1]:
        c.font = bold
    for col, w in zip("ABC", (34, 12, 80)):
        ps.column_dimensions[col].width = w

    for col, w in zip("ABCDEFGHIJ", (7, 50, 9, 11, 12, 14, 16, 11, 34, 34)):
        ws.column_dimensions[col].width = w
    comp = data.get("company") or {}
    obj = data["object"]
    r = 1
    if comp.get("name"):
        ws.cell(r, 1, comp["name"]).font = Font(bold=True, size=12, color=C["primary"])
    if br["logo"]:
        from openpyxl.drawing.image import Image as XLImage
        img = XLImage(br["logo"])
        img.width, img.height = int(img.width * 50 / img.height), 50
        ws.add_image(img, "I1")
    r = 3
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c = ws.cell(r, 1, f"{TITLE} № {data['number']} от {fmt_date(data['date'])}")
    c.font = Font(bold=True, size=14, color=C["primary"])
    c.alignment = Alignment(horizontal="center")
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(r, 1, f"({SUBTITLE})").alignment = Alignment(horizontal="center")
    r += 1
    for label, val in (("Объект:", obj.get("name")), ("Адрес:", obj.get("address")),
                       ("Основание:", data.get("basis"))):
        if val:
            ws.cell(r, 1, label).font = bold
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            ws.cell(r, 2, val)
            r += 1
    if data.get("sample"):
        ws.cell(r, 1, "ОБРАЗЕЦ. Тестовые данные.").font = Font(bold=True, color="C00000")
        r += 1
    r += 1
    hdr = ["№", "Наименование работ и ресурсов", "Ед. изм.", "Норма на ед.", "Кол-во",
           "Цена без НДС, тг", "Сумма без НДС, тг", "Тип", "Источник нормы / количества",
           "Источник цены"]
    for ci, h in enumerate(hdr, 1):
        c = ws.cell(r, ci, h)
        c.font = Font(bold=True, color=C["header_text"])
        c.fill = PatternFill("solid", fgColor=C["primary"])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = box
    header_row = r
    r += 1
    first = r
    res_cells: dict[tuple, list[str]] = {}
    for si, s in enumerate(calc["sections"], 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        c = ws.cell(r, 1, f"Раздел {si}. {s['title']}")
        c.font = bold
        c.fill = PatternFill("solid", fgColor=C["section"])
        r += 1
        for pi, p in enumerate(s["positions"], 1):
            pr = r
            vals = [f"{si}.{pi}", p["name"], p["unit"], None, float(p["qty"]), None,
                    f"=SUM(G{pr + 1}:G{pr + len(p['resources'])})", "позиция",
                    p["qty_source"] + (f"; калькуляция {p['card']}" if p["card"] else ""), ""]
            for ci, v in enumerate(vals, 1):
                c = ws.cell(r, ci, v)
                c.font = bold
                c.border = box
                c.alignment = wrap
                c.fill = PatternFill("solid", fgColor=C["subtotal"])
            ws.cell(r, 5).number_format = Q
            ws.cell(r, 7).number_format = MONEY
            r += 1
            for res in p["resources"]:
                vals = ["", f"   {res['name']}", res["unit"], float(res["norm"]),
                        f"=ROUND(D{r}*$E${pr},{RES_PLACES})", float(res["price"]),
                        f"=ROUND(E{r}*F{r},{n})", res["type"], res["norm_source"],
                        res["price_source"]]
                for ci, v in enumerate(vals, 1):
                    c = ws.cell(r, ci, v)
                    c.border = box
                    c.alignment = wrap
                ws.cell(r, 4).number_format = "0.0000"
                ws.cell(r, 5).number_format = Q
                ws.cell(r, 6).number_format = MONEY
                ws.cell(r, 7).number_format = MONEY
                res_cells.setdefault((res["type"], res["name"], res["unit"], res["price"]),
                                     []).append(r)
                r += 1
    last = r - 1
    r += 1
    names = {}

    def line(label, formula, name=None, rate=None, strong=False):
        nonlocal r
        ws.cell(r, 2, label).font = Font(bold=True, size=12 if strong else 11)
        if rate:
            ws.cell(r, 6, f"={rate}").number_format = "0.00%"
        c = ws.cell(r, 7, formula)
        c.number_format = MONEY
        c.font = Font(bold=True, size=12 if strong else 11)
        if name:
            names[name] = f"'Смета'!$G${r}"
        r += 1
        return r - 1

    rng = lambda col: f"${col}${first}:${col}${last}"  # noqa: E731
    type_rows = {}
    for t, label in TYPES.items():
        type_rows[t] = line(label, f'=SUMIF({rng("H")},"{t}",{rng("G")})', NAMES[t])
    rd = line("Итого прямые затраты", "=" + "+".join(f"G{x}" for x in type_rows.values()),
              "SM_DIRECT")
    oh_base = f"G{type_rows['труд']}" if rates.get("overhead_base", "direct") == "fot" else f"G{rd}"
    ro = line(f"Накладные расходы ({OVERHEAD_BASES[rates.get('overhead_base', 'direct')]})",
              f"=ROUND({oh_base}*{PARAMS}!$B$2,{n})", "SM_OVERHEAD", f"{PARAMS}!$B$2")
    pr_base = (f"G{type_rows['труд']}" if rates.get("profit_base", "direct_overhead") == "fot"
               else f"(G{rd}+G{ro})")
    rp = line(f"Сметная прибыль ({PROFIT_BASES[rates.get('profit_base', 'direct_overhead')]})",
              f"=ROUND({pr_base}*{PARAMS}!$B$3,{n})", "SM_PROFIT", f"{PARAMS}!$B$3")
    rsub = line("Итого без НДС", f"=G{rd}+G{ro}+G{rp}", "SM_SUBTOTAL")
    rv = line("НДС", f"=ROUND(G{rsub}*{PARAMS}!$B$4,{n})", "SM_VAT", f"{PARAMS}!$B$4")
    rt = line("ВСЕГО по смете", f"=G{rsub}+G{rv}", "SM_TOTAL", strong=True)
    ws.cell(rt, 7).fill = PatternFill("solid", fgColor=C["total"])
    line("Контроль: прямые затраты минус сумма позиций (должно быть 0)",
         f'=G{rd}-SUMIF({rng("H")},"позиция",{rng("G")})', "SM_CONTROL")
    r += 1
    ws.cell(r, 2, f"Сумма прописью на дату расчёта: {amount_in_words(calc['total'])}. "
                  "Текст не пересчитывается формулами.").font = Font(italic=True)
    r += 2
    for role in ("Составил", "Проверил"):
        ws.cell(r, 2, f"{role}: ____________________ / ____________________ /")
        r += 2
    ws.cell(r, 2, "Документ подготовлен для проверки и утверждения ответственным специалистом.") \
        .font = Font(italic=True, color="666666")
    ws.freeze_panes = ws.cell(header_row + 1, 1)
    ws.print_title_rows = f"{header_row}:{header_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Ресурсная ведомость: количества и суммы — ссылки на строки сметы
    rs.append(["Тип", "Ресурс", "Ед. изм.", "Кол-во", "Цена без НДС, тг", "Сумма без НДС, тг"])
    for c in rs[1]:
        c.font = bold
    for a in aggregate(calc):
        cells = res_cells[(a["type"], a["name"], a["unit"], a["price"])]
        i = rs.max_row + 1
        rs.append([a["type"], a["name"], a["unit"], "=" + "+".join(f"'Смета'!E{x}" for x in cells),
                   float(a["price"]), "=" + "+".join(f"'Смета'!G{x}" for x in cells)])
        rs.cell(i, 4).number_format = Q
        rs.cell(i, 5).number_format = MONEY
        rs.cell(i, 6).number_format = MONEY
    i = rs.max_row + 1
    rs.cell(i, 2, "Итого по ресурсам").font = bold
    rs.cell(i, 6, f"=SUM(F2:F{i - 1})").number_format = MONEY
    names["SM_RES_TOTAL"] = f"'Ресурсы'!$F${i}"
    for col, w in zip("ABCDEF", (11, 50, 9, 13, 16, 18)):
        rs.column_dimensions[col].width = w

    ns.column_dimensions["A"].width = 120
    for w in warnings or ["нет"]:
        ns.append([f"• {w}"])
    for name, ref in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def build_pdf(data: dict, calc: dict, path: Path) -> None:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer,
                                    Table, TableStyle)

    br = brand_for(data)
    Cc = {k: colors.HexColor(v) for k, v in br["colors"].items()}
    M = br["margins_mm"]
    W = (A4[0] / mm - M["left"] - M["right"]) * mm
    rates = data["rates"]
    comp = data.get("company") or {}
    obj = data["object"]

    def esc(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    st = {
        "n": ParagraphStyle("n", fontName="KP", fontSize=8.5, leading=11),
        "s": ParagraphStyle("s", fontName="KP", fontSize=7, leading=8.8),
        "c": ParagraphStyle("c", fontName="KP", fontSize=7, leading=8.8, alignment=TA_CENTER),
        "r": ParagraphStyle("r", fontName="KP", fontSize=7, leading=8.8, alignment=TA_RIGHT),
        "rb": ParagraphStyle("rb", fontName="KP-Bold", fontSize=7, leading=8.8, alignment=TA_RIGHT),
        "b": ParagraphStyle("b", fontName="KP-Bold", fontSize=8, leading=10),
        "h": ParagraphStyle("h", fontName="KP-Bold", fontSize=7, leading=8.8, alignment=TA_CENTER,
                            textColor=Cc["header_text"]),
        "t": ParagraphStyle("t", fontName="KP-Bold", fontSize=13, leading=17, alignment=TA_CENTER,
                            textColor=Cc["primary"]),
        "h2": ParagraphStyle("h2", fontName="KP-Bold", fontSize=10, leading=13,
                             textColor=Cc["primary"], spaceBefore=4 * mm, spaceAfter=1.5 * mm),
        "tn": ParagraphStyle("tn", fontName="KP", fontSize=9, leading=12, alignment=TA_RIGHT),
        "tb": ParagraphStyle("tb", fontName="KP-Bold", fontSize=9, leading=12, alignment=TA_RIGHT),
    }
    sample = bool(data.get("sample"))

    def on_page(canvas, doc):
        canvas.saveState()
        if sample:
            canvas.setFont("KP-Bold", 70)
            canvas.setFillColor(colors.Color(0.85, 0.1, 0.1, alpha=0.10))
            canvas.translate(A4[0] / 2, A4[1] / 2)
            canvas.rotate(45)
            canvas.drawCentredString(0, 0, "ОБРАЗЕЦ")
            canvas.restoreState()
            canvas.saveState()
        canvas.setFont("KP", 7.5)
        canvas.setFillColor(colors.grey)
        y = max(6, M["bottom"] - 8) * mm
        canvas.drawString(M["left"] * mm, y, f"Смета № {data['number']} от {fmt_date(data['date'])}")
        if br["footer_text"]:
            canvas.drawCentredString(A4[0] / 2, y, br["footer_text"])
        canvas.drawRightString(A4[0] - M["right"] * mm, y, f"Стр. {doc.page}")
        canvas.restoreState()

    story = []
    if br["logo"] or (br["show_requisites"] and comp.get("name")):
        cells, widths = [], []
        if br["logo"]:
            img = Image(br["logo"])
            lw = min(br["logo_width_mm"], 35) * mm
            img.drawWidth, img.drawHeight = lw, lw * img.imageHeight / img.imageWidth
            cells, widths = [img], [lw + 5 * mm]
        cells.append(Paragraph(f"<font name='KP-Bold' size='11'>{esc(comp.get('name', ''))}</font>",
                               st["n"]) if br["show_requisites"] else "")
        widths.append(W - sum(widths))
        head = Table([cells], colWidths=widths)
        head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                  ("LINEBELOW", (0, 0), (-1, 0), 1, Cc["primary"]),
                                  ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                  ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * mm)]))
        story += [head, Spacer(1, 4 * mm)]
    story += [Paragraph(f"{TITLE} № {esc(data['number'])}", st["t"]),
              Paragraph(f"от {fmt_date(data['date'])}, {SUBTITLE}",
                        ParagraphStyle("d", parent=st["n"], alignment=TA_CENTER, spaceAfter=3 * mm))]
    info = [(a, b) for a, b in (("Объект", obj.get("name")), ("Адрес", obj.get("address")),
                                ("Основание", data.get("basis"))) if b]
    t = Table([[Paragraph(esc(a), st["b"]), Paragraph(esc(b), st["n"])] for a, b in info],
              colWidths=[28 * mm, W - 28 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    story += [t, Spacer(1, 3 * mm)]

    k = W / (180 * mm)
    rows_ = [[Paragraph(h, st["h"]) for h in ("№", "Наименование", "Ед.", "Кол-во", "ФОТ",
                                             "Машины", "Материалы", "Транспорт", "Всего")]]
    style = [("BACKGROUND", (0, 0), (-1, 0), Cc["primary"]),
             ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
             ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2)]
    for si, s in enumerate(calc["sections"], 1):
        rows_.append([Paragraph(f"Раздел {si}. {esc(s['title'])}", st["b"])] + [""] * 8)
        i = len(rows_) - 1
        style += [("SPAN", (0, i), (-1, i)), ("BACKGROUND", (0, i), (-1, i), Cc["section"])]
        for pi, p in enumerate(s["positions"], 1):
            bt = p["by_type"]
            rows_.append([Paragraph(f"{si}.{pi}", st["c"]), Paragraph(esc(p["name"]), st["s"]),
                          Paragraph(esc(p["unit"]), st["c"]), Paragraph(fmt_qty(p["qty"]), st["r"])]
                         + [Paragraph(fmt_money(bt[x]) if bt[x] else "—", st["r"]) for x in TYPES]
                         + [Paragraph(fmt_money(p["total"]), st["rb"])])
    bt = calc["by_type"]
    rows_.append(["", Paragraph("Итого прямые затраты", st["b"]), "", ""]
                 + [Paragraph(fmt_money(bt[x]), st["rb"]) for x in TYPES]
                 + [Paragraph(fmt_money(calc["direct"]), st["rb"])])
    i = len(rows_) - 1
    style += [("SPAN", (1, i), (3, i)), ("BACKGROUND", (0, i), (-1, i), Cc["subtotal"])]
    tbl = Table(rows_, colWidths=[x * k * mm for x in (9, 50, 10, 14, 19, 19, 21, 17, 21)],
                repeatRows=1)
    tbl.setStyle(TableStyle(style))
    story.append(tbl)
    story.append(Spacer(1, 3 * mm))
    oh_b = OVERHEAD_BASES[rates.get("overhead_base", "direct")]
    pr_b = PROFIT_BASES[rates.get("profit_base", "direct_overhead")]
    tot = [("Итого прямые затраты", calc["direct"], False),
           (f"Накладные расходы {fmt_pct(rates['overhead_pct'])} {oh_b}", calc["overhead"], False),
           (f"Сметная прибыль {fmt_pct(rates['profit_pct'])} {pr_b}", calc["profit"], False),
           ("Итого без НДС", calc["subtotal"], True),
           (f"НДС {fmt_pct(rates['vat_pct'])}", calc["vat"], False),
           ("ВСЕГО по смете, тенге", calc["total"], True)]
    tt = Table([[Paragraph(esc(a), st["tb"] if b else st["tn"]),
                 Paragraph(fmt_money(v), st["tb"] if b else st["tn"])] for a, v, b in tot],
               colWidths=[W * 0.7, W * 0.3])
    tt.setStyle(TableStyle([("LINEABOVE", (0, 3), (-1, 3), 0.6, colors.grey),
                            ("BACKGROUND", (0, 5), (-1, 5), Cc["total"]),
                            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
    story += [KeepTogether([tt]), Spacer(1, 2 * mm),
              Paragraph(f"<i>Сумма прописью: {esc(amount_in_words(calc['total']))}.</i>", st["n"]),
              Paragraph(f"Основание ставки НДС: {esc(rates['vat_source'])}.", st["s"])]

    story.append(Paragraph("Ресурсная ведомость", st["h2"]))
    rr = [[Paragraph(h, st["h"]) for h in ("Тип", "Ресурс", "Ед.", "Кол-во", "Цена, тг", "Сумма, тг")]]
    for a in aggregate(calc):
        rr.append([Paragraph(esc(a["type"]), st["c"]), Paragraph(esc(a["name"]), st["s"]),
                   Paragraph(esc(a["unit"]), st["c"]), Paragraph(fmt_qty(a["qty"]), st["r"]),
                   Paragraph(fmt_money(a["price"]), st["r"]), Paragraph(fmt_money(a["amount"]), st["r"])])
    rt = Table(rr, colWidths=[x * k * mm for x in (18, 76, 12, 22, 24, 28)], repeatRows=1)
    rt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), Cc["primary"]),
                            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
    story += [rt, Spacer(1, 8 * mm)]
    for role in ("Составил", "Проверил"):
        story += [Paragraph(f"{role}: ______________________ / ______________________ /", st["n"]),
                  Spacer(1, 5 * mm)]
    story.append(Paragraph("Документ подготовлен для проверки и утверждения ответственным специалистом.",
                           st["s"]))

    out_path = path.with_suffix(".content.pdf") if br["letterhead_pdf"] else path
    doc = SimpleDocTemplate(str(out_path), pagesize=A4, leftMargin=M["left"] * mm,
                            rightMargin=M["right"] * mm, topMargin=M["top"] * mm,
                            bottomMargin=M["bottom"] * mm, title=f"Смета № {data['number']}",
                            author=comp.get("name", ""), subject=obj.get("name", ""))
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    if br["letterhead_pdf"]:
        merge_letterhead(out_path, Path(br["letterhead_pdf"]), path, br["letterhead_pages"])
        out_path.unlink()


def main() -> int:
    ap = argparse.ArgumentParser(description="Сборка локальной сметы")
    ap.add_argument("input")
    ap.add_argument("--out")
    a = ap.parse_args()
    try:
        data = load_smeta(a.input)
        calc, warnings = compute(data)
    except SmetaError as e:
        print("ОШИБКИ В СМЕТЕ. Файлы не созданы:")
        for x in e.errors:
            print(f"  - {x}")
        return 2
    inp = Path(a.input).resolve()
    out = Path(a.out).resolve() if a.out else inp.parent.parent / "output"
    out.mkdir(parents=True, exist_ok=True)
    base = out / f"SMETA_{safe_name(str(data['number']))}"
    xlsx, pdf = base.with_suffix(".xlsx"), base.with_suffix(".pdf")
    build_xlsx(data, calc, xlsx, warnings)
    br = brand_for(data)
    register_fonts(br["font_regular"], br["font_bold"])
    build_pdf(data, calc, pdf)
    print(f"Excel: {xlsx}")
    print(f"PDF:   {pdf}")
    for t, label in TYPES.items():
        print(f"{label:28} {fmt_money(calc['by_type'][t]):>18}")
    print(f"{'Итого прямые затраты':28} {fmt_money(calc['direct']):>18}")
    print(f"{'Итого без НДС':28} {fmt_money(calc['subtotal']):>18}")
    print(f"{'ВСЕГО по смете':28} {fmt_money(calc['total']):>18}")
    if warnings:
        print("Замечания:")
        for w in warnings:
            print(f"  - {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
