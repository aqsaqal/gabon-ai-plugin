#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "reportlab>=4.0", "pypdf>=4.0", "pillow>=10", "python-docx>=1.1"]
# ///
"""Акт выполненных работ (форма Р-1) + накопительная ведомость + справка о стоимости.
Excel с формулами, PDF и редактируемый Word.

    uv run build_act.py projects/<объект>/input/act_1.json [--out папка]
"""
from __future__ import annotations

import argparse
import sys
from decimal import Decimal
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from akty_core import FORM_HEADER, ActError, load_act  # noqa: E402
from build_kp import register_fonts, safe_name  # noqa: E402
from kp_core import D, amount_in_words, fmt_date, fmt_money, fmt_qty  # noqa: E402

TITLE = "АКТ ВЫПОЛНЕННЫХ РАБОТ (ОКАЗАННЫХ УСЛУГ)"
COLS = ["Номер по порядку",
        "Наименование работ (услуг) (в разрезе их подвидов в соответствии с технической "
        "спецификацией, заданием, графиком выполнения работ (услуг) при их наличии)",
        "Дата выполнения работ (оказания услуг)",
        "Сведения об отчете о научных исследованиях, маркетинговых, консультационных и прочих "
        "услугах (дата, номер, количество страниц) (при их наличии)",
        "Единица измерения", "Выполнено работ (оказано услуг): количество", "цена за единицу",
        "стоимость"]


def party(p: dict) -> str:
    parts = [p.get("name", "")]
    for k, lab in (("address", ""), ("phone", "тел. ")):
        if p.get(k):
            parts.append(lab + p[k])
    return ", ".join(x for x in parts if x)


def line_date(a, ln):
    return fmt_date(ln["date"]) if ln.get("date") else \
        f"{fmt_date(a['_from'])}–{fmt_date(a['_to'])}"


def build_xlsx(a, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, Side
    from openpyxl.workbook.defined_name import DefinedName
    c, calc = a["_contract"], a["_calc"]
    n = c["_places"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Акт"
    thin = Side(style="thin", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    bold = Font(bold=True)
    MONEY = "#,##0.00"
    for col, w in zip("ABCDEFGHIJ", (8, 46, 14, 16, 10, 13, 14, 16, 40, 10)):
        ws.column_dimensions[col].width = w
    for i, t in enumerate(FORM_HEADER, 1):
        ws.cell(i, 8, t).alignment = Alignment(horizontal="right")
    r = 5
    for label, p in (("Заказчик", c["customer"]), ("Исполнитель", c["contractor"])):
        ws.cell(r, 1, label).font = bold
        ws.cell(r, 2, f"ИИН/БИН {p.get('bin') or '____________'}")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        ws.cell(r, 3, party(p)).alignment = wrap
        r += 1
    ws.cell(r, 1, "Договор (контракт)").font = bold
    ws.cell(r, 3, f"№ {c['number']} от {fmt_date(c['date'])}")
    r += 2
    ws.cell(r, 6, "Номер документа").font = bold
    ws.cell(r, 7, "Дата составления").font = bold
    ws.cell(r + 1, 6, str(a["number"]))
    ws.cell(r + 1, 7, fmt_date(a["date"]))
    r += 3
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(r, 1, TITLE).font = Font(bold=True, size=13)
    ws.cell(r, 1).alignment = Alignment(horizontal="center")
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(r, 1, f"Отчётный период: с {fmt_date(a['_from'])} по {fmt_date(a['_to'])}. "
                  f"Объект: {c.get('subject', '')}").alignment = Alignment(horizontal="center")
    r += 2
    for ci, h in enumerate(COLS + ["Подтверждение объёма (не печатается)", "Позиция договора"], 1):
        cell = ws.cell(r, ci, h)
        cell.font = Font(bold=True, size=8)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = box
    ws.row_dimensions[r].height = 90
    hr = r
    r += 1
    for ci in range(1, 9):
        cell = ws.cell(r, ci, ci)
        cell.alignment = Alignment(horizontal="center")
        cell.border = box
    r += 1
    first = r
    rows_of = {}
    for i, (ln, row) in enumerate(zip(a["lines"], calc["rows"]), 1):
        vals = [i, row["name"], line_date(a, ln), ln.get("report", ""), row["unit"],
                float(row["qty"]), float(row["price"]), f"=ROUND(F{r}*G{r},{n})",
                ln["evidence"], row["n"]]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(r, ci, v)
            cell.alignment = wrap
            if ci <= 8:
                cell.border = box
        ws.cell(r, 6).number_format = "#,##0.000"
        ws.cell(r, 7).number_format = MONEY
        ws.cell(r, 8).number_format = MONEY
        rows_of[row["n"]] = r
        r += 1
    last = r - 1
    names = {}

    def tot(label, formula, name, rate=None):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.cell(r, 1, label).font = bold
        ws.cell(r, 1).alignment = Alignment(horizontal="right")
        if rate is not None:
            ws.cell(r, 7, rate).number_format = "0.00%"
        cell = ws.cell(r, 8, formula)
        cell.number_format = MONEY
        cell.font = bold
        for ci in range(1, 9):
            ws.cell(r, ci).border = box
        names[name] = f"'Акт'!$H${r}"
        r += 1
        return r - 1
    rs = tot("Итого без НДС", f"=SUM(H{first}:H{last})", "ACT_SUB")
    rv = tot("НДС", f"=ROUND(H{rs}*G{r},{n})", "ACT_VAT", float(D(c["vat_pct"]) / 100))
    rt = tot("Итого с НДС", f"=H{rs}+H{rv}", "ACT_TOTAL")
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(r, 1, f"Сумма прописью: {amount_in_words(calc['total'])}, в том числе НДС "
                  f"{fmt_money(calc['vat'])} тенге.")
    r += 2
    for txt in (f"Сведения об использовании запасов, полученных от заказчика: "
                f"{a.get('customer_materials') or 'не использовались'}",
                f"Приложение: документы, подтверждающие выполненные объёмы, "
                f"{len({ln['evidence'] for ln in a['lines']})} наименований, на ____ страницах"):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.cell(r, 1, txt).alignment = wrap
        r += 1
    r += 1
    cu, co = c["customer"], c["contractor"]
    ws.cell(r, 1, "Сдал (Исполнитель)").font = bold
    ws.cell(r, 5, "Принял (Заказчик)").font = bold
    r += 1
    ws.cell(r, 1, f"{co.get('signer_position') or '__________'} ________ / "
                  f"{co.get('signer_name') or '__________'} /")
    ws.cell(r, 5, f"{cu.get('signer_position') or '__________'} ________ / "
                  f"{cu.get('signer_name') or '__________'} /")
    r += 1
    ws.cell(r, 1, "должность    подпись    расшифровка подписи").font = Font(size=7)
    ws.cell(r, 5, "должность    подпись    расшифровка подписи").font = Font(size=7)
    r += 1
    ws.cell(r, 1, "М.П.")
    ws.cell(r, 5, "Дата подписания (принятия) работ (услуг) ____________")
    r += 1
    ws.cell(r, 5, "М.П.")
    ws.print_area = f"A1:H{r}"
    ws.print_title_rows = f"{hr}:{hr}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Накопительная ведомость
    ns = wb.create_sheet("Накопительная")
    ns.append([f"Накопительная ведомость выполнения по договору № {c['number']} "
               f"на {fmt_date(a['_to'])} (акт № {a['number']})"])
    ns["A1"].font = bold
    ns.append([])
    ns.append(["Позиция договора", "Наименование", "Ед.", "По договору", "Выполнено ранее",
               "За период", "С начала работ", "Остаток", "% выполнения"])
    for cell in ns[3]:
        cell.font = bold
        cell.alignment = Alignment(wrap_text=True)
    for p in c["positions"]:
        i = ns.max_row + 1
        this = f"='Акт'!F{rows_of[str(p['n'])]}" if str(p["n"]) in rows_of else 0
        ns.append([p["n"], p["name"], p["unit"], float(p["_qty"]),
                   float(a["_prev_qty"].get(str(p["n"]), Decimal(0))), this,
                   f"=E{i}+F{i}", f"=D{i}-G{i}", f"=ROUND(G{i}/D{i}*100,1)"])
        for col in "DEFGH":
            ns[f"{col}{i}"].number_format = "#,##0.000"
    for col, w in zip("ABCDEFGHI", (10, 46, 8, 13, 13, 13, 13, 13, 11)):
        ns.column_dimensions[col].width = w

    # Справка о стоимости
    ss = wb.create_sheet("Справка")
    ss.append([f"Справка о стоимости выполненных работ по договору № {c['number']} "
               f"от {fmt_date(c['date'])}, акт № {a['number']}"])
    ss["A1"].font = bold
    adv, ret = calc["adv"], calc["ret"]
    rows = [("Стоимость работ по договору с НДС", float(D(c["totals"]["total"])), None),
            ("Выполнено ранее с НДС", float(calc["prev_total"]), None),
            ("Выполнено за отчётный период с НДС", f"='Акт'!H{rt}", "ACT_PERIOD"),
            ("Выполнено с начала работ с НДС", "=B3+B4", None),
            ("Остаток по договору с НДС", "=B2-B5", None)]
    if adv:
        rows.append((f"Остаток аванса к зачёту (получено {fmt_money(D(adv['paid']))}, "
                     f"зачтено ранее {fmt_money(calc['prev_offset'])})",
                     float(D(adv["paid"]) - calc["prev_offset"]), None))
        rows.append((f"Зачёт аванса {adv['offset_pct']} % ({adv['source']})",
                     f"=MIN(ROUND(B4*{float(D(adv['offset_pct']) / 100)},{n}),MAX(B7,0))", "ACT_OFFSET"))
    if ret:
        rows.append((f"Гарантийное удержание {ret['pct']} % ({ret['source']})",
                     f"=ROUND(B4*{float(D(ret['pct']) / 100)},{n})", "ACT_RET"))
    for label, v, name in rows:
        ss.append([label, v])
        if name:
            names[name] = f"'Справка'!$B${ss.max_row}"
    deduct = [f"B{i}" for i in range(2, ss.max_row + 1)
              if ss.cell(i, 1).value.startswith(("Зачёт аванса", "Гарантийное"))]
    ss.append(["К оплате за отчётный период", "=B4" + "".join(f"-{x}" for x in deduct)])
    names["ACT_PAY"] = f"'Справка'!$B${ss.max_row}"
    ss.cell(ss.max_row, 1).font = bold
    ss.cell(ss.max_row, 2).font = bold
    for i in range(2, ss.max_row + 1):
        ss.cell(i, 2).number_format = MONEY
    ss.column_dimensions["A"].width = 80
    ss.column_dimensions["B"].width = 20

    es = wb.create_sheet("Замечания")
    es.column_dimensions["A"].width = 120
    for w in a["_warnings"] or ["нет"]:
        es.append([f"• {w}"])
    for k, v in names.items():
        wb.defined_names[k] = DefinedName(k, attr_text=v)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def build_pdf(a, path):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    c, calc = a["_contract"], a["_calc"]
    st = {"n": ParagraphStyle("n", fontName="KP", fontSize=8.5, leading=11),
          "s": ParagraphStyle("s", fontName="KP", fontSize=7, leading=8.6),
          "sc": ParagraphStyle("sc", fontName="KP", fontSize=6.5, leading=8, alignment=TA_CENTER),
          "r": ParagraphStyle("r", fontName="KP", fontSize=7.5, leading=9, alignment=TA_RIGHT),
          "rb": ParagraphStyle("rb", fontName="KP-Bold", fontSize=7.5, leading=9, alignment=TA_RIGHT),
          "hr": ParagraphStyle("hr", fontName="KP", fontSize=7, leading=8.6, alignment=TA_RIGHT),
          "t": ParagraphStyle("t", fontName="KP-Bold", fontSize=12, leading=15, alignment=TA_CENTER),
          "b": ParagraphStyle("b", fontName="KP-Bold", fontSize=8.5, leading=11),
          "h2": ParagraphStyle("h2", fontName="KP-Bold", fontSize=10, leading=13, spaceBefore=3 * mm,
                               spaceAfter=1.5 * mm)}

    def esc(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    sample = bool(a.get("sample") or c.get("sample"))

    def on_page(cv, doc):
        cv.saveState()
        if sample:
            cv.setFont("KP-Bold", 70)
            cv.setFillColor(colors.Color(0.85, 0.1, 0.1, alpha=0.08))
            cv.translate(landscape(A4)[0] / 2, landscape(A4)[1] / 2)
            cv.rotate(25)
            cv.drawCentredString(0, 0, "ОБРАЗЕЦ")
        cv.restoreState()

    s = [Paragraph("<br/>".join(FORM_HEADER), st["hr"]), Spacer(1, 2 * mm)]
    hdr = Table([[Paragraph("<b>Заказчик</b>", st["n"]),
                  Paragraph(f"ИИН/БИН {esc(c['customer'].get('bin') or '____________')}", st["n"]),
                  Paragraph(esc(party(c["customer"])), st["n"])],
                 [Paragraph("<b>Исполнитель</b>", st["n"]),
                  Paragraph(f"ИИН/БИН {esc(c['contractor'].get('bin') or '____________')}", st["n"]),
                  Paragraph(esc(party(c["contractor"])), st["n"])],
                 [Paragraph("<b>Договор (контракт)</b>", st["n"]), "",
                  Paragraph(f"№ {esc(c['number'])} от {fmt_date(c['date'])}", st["n"])]],
                colWidths=[35 * mm, 45 * mm, 190 * mm])
    hdr.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    num = Table([["Номер документа", "Дата составления"], [str(a["number"]), fmt_date(a["date"])]],
                colWidths=[35 * mm, 35 * mm], hAlign="RIGHT")
    num.setStyle(TableStyle([("FONTNAME", (0, 0), (-1, -1), "KP"), ("FONTSIZE", (0, 0), (-1, -1), 8),
                             ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                             ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    s += [hdr, Spacer(1, 1 * mm), num, Spacer(1, 3 * mm), Paragraph(TITLE, st["t"]),
          Paragraph(f"Отчётный период: с {fmt_date(a['_from'])} по {fmt_date(a['_to'])}. "
                    f"Объект: {esc(c.get('subject', ''))}",
                    ParagraphStyle("p", parent=st["n"], alignment=TA_CENTER)), Spacer(1, 2 * mm)]
    rows = [[Paragraph(esc(h), st["sc"]) for h in COLS],
            [Paragraph(str(i), st["sc"]) for i in range(1, 9)]]
    for i, (ln, row) in enumerate(zip(a["lines"], calc["rows"]), 1):
        rows.append([Paragraph(str(i), st["sc"]), Paragraph(esc(row["name"]), st["s"]),
                     Paragraph(line_date(a, ln), st["sc"]), Paragraph(esc(ln.get("report", "")), st["s"]),
                     Paragraph(esc(row["unit"]), st["sc"]), Paragraph(fmt_qty(row["qty"]), st["r"]),
                     Paragraph(fmt_money(row["price"]), st["r"]), Paragraph(fmt_money(row["amount"]), st["r"])])
    for label, v in (("Итого без НДС", calc["subtotal"]), (f"НДС {c['vat_pct']} %", calc["vat"]),
                     ("Итого с НДС", calc["total"])):
        rows.append([Paragraph(label, st["rb"]), "", "", "", "", "", "", Paragraph(fmt_money(v), st["rb"])])
    t = Table(rows, colWidths=[x * mm for x in (14, 92, 26, 36, 18, 26, 26, 31)], repeatRows=2)
    k = len(rows)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE")] +
                          [("SPAN", (0, i), (6, i)) for i in range(k - 3, k)]))
    s += [t, Spacer(1, 2 * mm),
          Paragraph(f"Сумма прописью: {esc(amount_in_words(calc['total']))}, в том числе НДС "
                    f"{fmt_money(calc['vat'])} тенге.", st["n"]),
          Paragraph(f"Сведения об использовании запасов, полученных от заказчика: "
                    f"{esc(a.get('customer_materials') or 'не использовались')}", st["n"]),
          Paragraph(f"Приложение: документы, подтверждающие выполненные объёмы, "
                    f"{len({ln['evidence'] for ln in a['lines']})} наименований, на ____ страницах.",
                    st["n"]), Spacer(1, 6 * mm)]
    cu, co = c["customer"], c["contractor"]
    sig = Table([[Paragraph("<b>Сдал (Исполнитель)</b>", st["n"]), Paragraph("<b>Принял (Заказчик)</b>", st["n"])],
                 [Paragraph(f"{esc(co.get('signer_position') or '__________')} ____________ / "
                            f"{esc(co.get('signer_name') or '______________')} /", st["n"]),
                  Paragraph(f"{esc(cu.get('signer_position') or '__________')} ____________ / "
                            f"{esc(cu.get('signer_name') or '______________')} /", st["n"])],
                 [Paragraph("должность, подпись, расшифровка подписи", st["s"]),
                  Paragraph("должность, подпись, расшифровка подписи", st["s"])],
                 [Paragraph("М.П.", st["n"]),
                  Paragraph("Дата подписания (принятия) работ (услуг) ____________<br/>М.П.", st["n"])]],
                colWidths=[135 * mm, 135 * mm])
    s += [sig, PageBreak()]

    s.append(Paragraph(f"Справка о стоимости выполненных работ по договору № {esc(c['number'])}, "
                       f"акт № {esc(a['number'])}", st["h2"]))
    sp = [("Стоимость работ по договору с НДС", D(c["totals"]["total"])),
          ("Выполнено ранее с НДС", calc["prev_total"]),
          ("Выполнено за отчётный период с НДС", calc["total"]),
          ("Выполнено с начала работ с НДС", calc["prev_total"] + calc["total"]),
          ("Остаток по договору с НДС", D(c["totals"]["total"]) - calc["prev_total"] - calc["total"])]
    if calc["adv"]:
        sp.append((f"Зачёт аванса {calc['adv']['offset_pct']} % ({calc['adv']['source']})", calc["offset"]))
    if calc["ret"]:
        sp.append((f"Гарантийное удержание {calc['ret']['pct']} % ({calc['ret']['source']})", calc["retention"]))
    sp.append(("К оплате за отчётный период", calc["pay"]))
    tt = Table([[Paragraph(esc(x), st["n"]), Paragraph(fmt_money(v), st["r"])] for x, v in sp],
               colWidths=[150 * mm, 40 * mm], hAlign="LEFT")
    tt.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                            ("FONTNAME", (0, -1), (-1, -1), "KP-Bold")]))
    s.append(tt)
    s.append(Paragraph("Накопительная ведомость по договору", st["h2"]))
    acc = [[Paragraph(h, st["sc"]) for h in ("Поз.", "Наименование", "Ед.", "По договору",
                                              "Ранее", "За период", "С начала", "Остаток", "%")]]
    this = {r["n"]: r["qty"] for r in calc["rows"]}
    for p in c["positions"]:
        nn = str(p["n"])
        prev, cur = a["_prev_qty"].get(nn, Decimal(0)), this.get(nn, Decimal(0))
        cum = prev + cur
        acc.append([Paragraph(nn, st["sc"]), Paragraph(esc(p["name"]), st["s"]), Paragraph(esc(p["unit"]), st["sc"])]
                   + [Paragraph(fmt_qty(x), st["r"]) for x in (p["_qty"], prev, cur, cum, p["_qty"] - cum)]
                   + [Paragraph(f"{cum / p['_qty'] * 100:.1f}".replace(".", ","), st["r"])])
    at = Table(acc, colWidths=[x * mm for x in (14, 100, 14, 24, 24, 24, 24, 24, 16)], repeatRows=1)
    at.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    s.append(at)
    s.append(Paragraph("Подтверждение выполненных объёмов", st["h2"]))
    for i, (ln, row) in enumerate(zip(a["lines"], calc["rows"]), 1):
        s.append(Paragraph(f"{i}. {esc(row['name'])}: {esc(ln['evidence'])}", st["s"]))
    SimpleDocTemplate(str(path), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                      topMargin=10 * mm, bottomMargin=10 * mm,
                      title=f"Акт № {a['number']}").build(s, onFirstPage=on_page, onLaterPages=on_page)


def build_docx(a, path):
    import docx
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Mm, Pt
    c, calc = a["_contract"], a["_calc"]
    d = docx.Document()
    sec = d.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Mm(297), Mm(210)
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Mm(12))
    style = d.styles["Normal"]
    style.font.name, style.font.size = "Arial", Pt(9)
    for t in FORM_HEADER:
        p = d.add_paragraph(t)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(0)
    for label, pp in (("Заказчик", c["customer"]), ("Исполнитель", c["contractor"])):
        d.add_paragraph(f"{label}: ИИН/БИН {pp.get('bin') or '____________'}, {party(pp)}")
    d.add_paragraph(f"Договор (контракт): № {c['number']} от {fmt_date(c['date'])}")
    d.add_paragraph(f"Номер документа: {a['number']}    Дата составления: {fmt_date(a['date'])}")
    h = d.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.add_run(TITLE).bold = True
    p = d.add_paragraph(f"Отчётный период: с {fmt_date(a['_from'])} по {fmt_date(a['_to'])}. "
                        f"Объект: {c.get('subject', '')}")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t = d.add_table(rows=2, cols=8)
    t.style = "Table Grid"
    for i, txt in enumerate(COLS):
        t.cell(0, i).text = txt
        t.cell(1, i).text = str(i + 1)
    for i, (ln, row) in enumerate(zip(a["lines"], calc["rows"]), 1):
        cells = t.add_row().cells
        for j, v in enumerate([str(i), row["name"], line_date(a, ln), ln.get("report", ""), row["unit"],
                               fmt_qty(row["qty"]), fmt_money(row["price"]), fmt_money(row["amount"])]):
            cells[j].text = v
    for label, v in (("Итого без НДС", calc["subtotal"]), (f"НДС {c['vat_pct']} %", calc["vat"]),
                     ("Итого с НДС", calc["total"])):
        cells = t.add_row().cells
        m = cells[0].merge(cells[6])
        m.text = label
        cells[7].text = fmt_money(v)
    for row in t.rows:
        for cell in row.cells:
            for par in cell.paragraphs:
                for run in par.runs:
                    run.font.size = Pt(7.5)
    d.add_paragraph(f"Сумма прописью: {amount_in_words(calc['total'])}, в том числе НДС "
                    f"{fmt_money(calc['vat'])} тенге.")
    d.add_paragraph(f"Сведения об использовании запасов, полученных от заказчика: "
                    f"{a.get('customer_materials') or 'не использовались'}")
    d.add_paragraph(f"Приложение: документы, подтверждающие выполненные объёмы, "
                    f"{len({ln['evidence'] for ln in a['lines']})} наименований, на ____ страницах.")
    cu, co = c["customer"], c["contractor"]
    s = d.add_table(rows=3, cols=2)
    s.cell(0, 0).text, s.cell(0, 1).text = "Сдал (Исполнитель)", "Принял (Заказчик)"
    s.cell(1, 0).text = (f"{co.get('signer_position') or '__________'} ____________ / "
                         f"{co.get('signer_name') or '______________'} /")
    s.cell(1, 1).text = (f"{cu.get('signer_position') or '__________'} ____________ / "
                         f"{cu.get('signer_name') or '______________'} /")
    s.cell(2, 0).text = "М.П."
    s.cell(2, 1).text = "Дата подписания (принятия) работ (услуг) ____________   М.П."
    d.save(path)


def main() -> int:
    ap = argparse.ArgumentParser(description="Акт выполненных работ")
    ap.add_argument("input")
    ap.add_argument("--out")
    ap.add_argument("--no-docx", action="store_true")
    a_ = ap.parse_args()
    try:
        a = load_act(a_.input)
    except ActError as e:
        print("ОШИБКИ В АКТЕ. Файлы не созданы:")
        for x in e.errors:
            print(f"  - {x}")
        return 2
    inp = Path(a_.input).resolve()
    out = Path(a_.out).resolve() if a_.out else inp.parent.parent / "output"
    out.mkdir(parents=True, exist_ok=True)
    base = out / f"AKT_{safe_name(str(a['_contract']['number']))}_{safe_name(str(a['number']))}"
    build_xlsx(a, base.with_suffix(".xlsx"))
    register_fonts(None, None)
    build_pdf(a, base.with_suffix(".pdf"))
    if not a_.no_docx:
        build_docx(a, base.with_suffix(".docx"))
    calc = a["_calc"]
    print(f"Excel: {base.with_suffix('.xlsx')}\nPDF:   {base.with_suffix('.pdf')}")
    if not a_.no_docx:
        print(f"Word:  {base.with_suffix('.docx')}")
    print(f"За период: без НДС {fmt_money(calc['subtotal'])}, НДС {fmt_money(calc['vat'])}, "
          f"всего {fmt_money(calc['total'])}")
    if calc["adv"] or calc["ret"]:
        print(f"Зачёт аванса {fmt_money(calc['offset'])}, удержание {fmt_money(calc['retention'])}, "
              f"к оплате {fmt_money(calc['pay'])}")
    for w in a["_warnings"]:
        print(f"  - {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
