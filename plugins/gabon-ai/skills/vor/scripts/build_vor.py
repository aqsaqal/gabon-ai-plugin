#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl==3.1.5", "reportlab==5.1.0", "pypdf==6.21.4", "pillow==13.0.0"]
# ///
"""Сборка ведомости объёмов работ (ВОР): Excel с формулами подсчёта + PDF.

    uv run build_vor.py projects/<объект>/input/vor.json [--out папка]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1] / "kp" / "scripts"))
sys.path.insert(0, str(HERE.parents[1] / "brand" / "scripts"))
from brand_core import brand_for  # noqa: E402
from build_kp import merge_letterhead, register_fonts, safe_name  # noqa: E402
from kp_core import fmt_date  # noqa: E402
from vor_core import (STATUSES, VorError, compute, display_formula, fmt_q, load_vor,  # noqa: E402
                      row_numbers, rows, to_excel)

STATUS_FILL = {"подтверждено": None, "допущение": "FFF2CC", "нет данных": "F8CBAD"}
STATUS_SHORT = {"подтверждено": "подтв.", "допущение": "допущ.", "нет данных": "нет данных"}


def note_of(it: dict) -> str:
    parts = [it.get("assumption"), it.get("missing"), it.get("note")]
    return "; ".join(str(p).strip() for p in parts if p and str(p).strip())


def source_of(it: dict) -> str:
    ref = str(it.get("ref") or "").strip()
    src = str(it.get("source") or "").strip()
    loc = str(it.get("loc") or "").strip()
    return ", ".join(x for x in (ref, loc, src) if x)


def build_xlsx(data: dict, values: dict, path: Path, warnings: list[str]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.workbook.defined_name import DefinedName

    br = brand_for(data)
    C = {k: v.lstrip("#") for k, v in br["colors"].items()}
    places = data.get("qty_decimals", 2)
    comp = data.get("company") or {}
    obj = data["object"]
    wb = Workbook()
    ws = wb.active
    ws.title = "ВОР"
    ds = wb.create_sheet("Документы")
    bold = Font(bold=True)
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    QTY = "#,##0." + "0" * places if places else "#,##0"

    for col, w in zip("ABCDEFGHIJ", (7, 48, 9, 13, 30, 22, 13, 40, 11, 9)):
        ws.column_dimensions[col].width = w
    r = 1
    if comp.get("name"):
        ws.cell(r, 1, comp["name"]).font = Font(bold=True, size=12, color=C["primary"])
        r += 1
    if br["logo"]:
        from openpyxl.drawing.image import Image as XLImage
        img = XLImage(br["logo"])
        img.width, img.height = int(img.width * 50 / img.height), 50
        ws.add_image(img, "H1")
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c = ws.cell(r, 1, f"ВЕДОМОСТЬ ОБЪЁМОВ РАБОТ № {data['number']} от {fmt_date(data['date'])}")
    c.font = Font(bold=True, size=14, color=C["primary"])
    c.alignment = Alignment(horizontal="center")
    r += 1
    for label, val in (("Объект:", obj.get("name", "")), ("Адрес:", obj.get("address", "")),
                       ("Основание:", data.get("basis", ""))):
        if val:
            ws.cell(r, 1, label).font = bold
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            ws.cell(r, 2, val).alignment = wrap
            r += 1
    if data.get("sample"):
        ws.cell(r, 1, "ОБРАЗЕЦ. Тестовые данные.").font = Font(bold=True, color="C00000")
        r += 1
    r += 1
    hdr = ["№", "Наименование работ", "Ед. изм.", "Кол-во", "Формула подсчёта", "Источник",
           "Статус", "Примечание, допущение, недостающие данные", "Код прайса", "ID строки"]
    for ci, h in enumerate(hdr, 1):
        c = ws.cell(r, ci, h)
        c.font = Font(bold=True, color=C["header_text"])
        c.fill = PatternFill("solid", fgColor=C["primary"])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = box
    header_row = r
    r += 1

    # Сначала раскладываем строки, чтобы знать адреса для ссылок @id
    nums = row_numbers(data)
    layout, cell_of, rr = [], {}, r
    for si, s in enumerate(data["sections"], 1):
        layout.append(("section", si, s, rr))
        rr += 1
        for ii, it in enumerate(s["items"], 1):
            layout.append(("item", f"{si}.{ii}", it, rr))
            cell_of[it["_id"]] = f"D{rr}"
            rr += 1
    first = r
    for kind, num, obj_, row in layout:
        if kind == "section":
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
            c = ws.cell(row, 1, f"Раздел {num}. {obj_['title']}")
            c.font = bold
            c.fill = PatternFill("solid", fgColor=C["section"])
            continue
        it = obj_
        st = it["status"]
        if st == "нет данных":
            q = None
        elif it.get("formula"):
            q = f"=ROUND({to_excel(it['formula'], cell_of)},{places})"
        else:
            q = float(values[it["_id"]])
        vals = [num, it["name"], it["unit"], q, display_formula(it.get("formula"), nums),
                source_of(it),
                st, note_of(it), it.get("code", ""), it["_id"]]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row, ci, v)
            c.border = box
            c.alignment = wrap
            if STATUS_FILL[st]:
                c.fill = PatternFill("solid", fgColor=STATUS_FILL[st])
        ws.cell(row, 4).number_format = QTY
    last = rr - 1
    r = rr + 1
    ws.cell(r, 2, "Сводка по статусам строк").font = bold
    r += 1
    for st, name in zip(STATUSES, ("VOR_OK", "VOR_ASSUME", "VOR_MISSING")):
        ws.cell(r, 2, st)
        ws.cell(r, 4, f'=COUNTIF($G${first}:$G${last},"{st}")')
        wb.defined_names[name] = DefinedName(name, attr_text=f"'ВОР'!$D${r}")
        r += 1
    r += 1
    ws.cell(r, 2, f"Количества округлены до {places} знаков после запятой. Ячейки «Кол-во» "
                  "считаются формулами из колонки «Формула подсчёта».").font = Font(italic=True)
    r += 2
    for role in ("Составил", "Проверил"):
        ws.cell(r, 2, f"{role}: ____________________ / ____________________ /  «___» ________ 20__ г.")
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

    ds.append(["ID", "Документ", "Файл", "Примечание"])
    for c in ds[1]:
        c.font = bold
    for d in data.get("documents") or []:
        ds.append([d.get("id", ""), d.get("name", ""), d.get("file", ""), d.get("note", "")])
    ds.append([])
    ds.append(["", "Замечания проверки"])
    for w in warnings or ["нет"]:
        ds.append(["", w])
    for col, w in zip("ABCD", (12, 60, 40, 40)):
        ds.column_dimensions[col].width = w

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def build_pdf(data: dict, values: dict, path: Path) -> None:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    br = brand_for(data)
    C = {k: colors.HexColor(v) for k, v in br["colors"].items()}
    M = br["margins_mm"]
    W = (A4[0] / mm - M["left"] - M["right"]) * mm
    places = data.get("qty_decimals", 2)
    comp = data.get("company") or {}
    obj = data["object"]

    def esc(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    st = {
        "n": ParagraphStyle("n", fontName="KP", fontSize=9, leading=11.5),
        "s": ParagraphStyle("s", fontName="KP", fontSize=7.5, leading=9.5),
        "c": ParagraphStyle("c", fontName="KP", fontSize=7.5, leading=9.5, alignment=TA_CENTER),
        "r": ParagraphStyle("r", fontName="KP", fontSize=7.5, leading=9.5, alignment=TA_RIGHT),
        "b": ParagraphStyle("b", fontName="KP-Bold", fontSize=8.5, leading=11),
        "h": ParagraphStyle("h", fontName="KP-Bold", fontSize=7.5, leading=9.5,
                            alignment=TA_CENTER, textColor=C["header_text"]),
        "t": ParagraphStyle("t", fontName="KP-Bold", fontSize=13, leading=17, alignment=TA_CENTER,
                            textColor=C["primary"], spaceAfter=2 * mm),
        "h2": ParagraphStyle("h2", fontName="KP-Bold", fontSize=10, leading=13,
                             textColor=C["primary"], spaceBefore=4 * mm, spaceAfter=1.5 * mm),
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
        canvas.drawString(M["left"] * mm, y, f"ВОР № {data['number']} от {fmt_date(data['date'])}")
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
        name = Paragraph(f"<font name='KP-Bold' size='11'>{esc(comp.get('name', ''))}</font>",
                         st["n"]) if br["show_requisites"] else ""
        cells.append(name)
        widths.append(W - sum(widths))
        head = Table([cells], colWidths=widths)
        head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                  ("LINEBELOW", (0, 0), (-1, 0), 1, C["primary"]),
                                  ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                  ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * mm)]))
        story += [head, Spacer(1, 4 * mm)]
    story.append(Paragraph(f"ВЕДОМОСТЬ ОБЪЁМОВ РАБОТ № {esc(data['number'])}", st["t"]))
    story.append(Paragraph(f"от {fmt_date(data['date'])}",
                           ParagraphStyle("d", parent=st["n"], alignment=TA_CENTER,
                                          spaceAfter=3 * mm)))
    info = [(a, b) for a, b in (("Объект", obj.get("name")), ("Адрес", obj.get("address")),
                                ("Основание", data.get("basis"))) if b]
    t = Table([[Paragraph(esc(a), st["b"]), Paragraph(esc(b), st["n"])] for a, b in info],
              colWidths=[30 * mm, W - 30 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
    story += [t, Spacer(1, 3 * mm)]

    k = W / (180 * mm)
    rows_ = [[Paragraph(h, st["h"]) for h in
              ("№", "Наименование работ", "Ед.", "Кол-во", "Формула подсчёта", "Источник", "Статус")]]
    style = [("BACKGROUND", (0, 0), (-1, 0), C["primary"]),
             ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
    fills = {"допущение": colors.HexColor("#FFF2CC"), "нет данных": colors.HexColor("#F8CBAD")}
    nums = row_numbers(data)
    for si, s in enumerate(data["sections"], 1):
        rows_.append([Paragraph(f"Раздел {si}. {esc(s['title'])}", st["b"])] + [""] * 6)
        n = len(rows_) - 1
        style += [("SPAN", (0, n), (-1, n)), ("BACKGROUND", (0, n), (-1, n), C["section"])]
        for ii, it in enumerate(s["items"], 1):
            q = values.get(it["_id"])
            rows_.append([Paragraph(f"{si}.{ii}", st["c"]), Paragraph(esc(it["name"]), st["s"]),
                          Paragraph(esc(it["unit"]), st["c"]),
                          Paragraph("—" if q is None else fmt_q(q, places), st["r"]),
                          Paragraph(esc(display_formula(it.get("formula"), nums) or
                                        ("из документа" if q is not None else "")), st["s"]),
                          Paragraph(esc(source_of(it)), st["s"]),
                          Paragraph(STATUS_SHORT[it["status"]], st["c"])])
            if it["status"] in fills:
                n = len(rows_) - 1
                style.append(("BACKGROUND", (0, n), (-1, n), fills[it["status"]]))
    tbl = Table(rows_, colWidths=[x * k * mm for x in (10, 58, 12, 18, 38, 28, 16)], repeatRows=1)
    tbl.setStyle(TableStyle(style))
    story.append(tbl)

    notes = [(f"{si}.{ii}", it) for si, ii, it in rows(data) if note_of(it)]
    if notes:
        story.append(Paragraph("Примечания, допущения и недостающие данные", st["h2"]))
        for num, it in notes:
            story.append(Paragraph(f"<b>{num}</b> ({STATUS_SHORT[it['status']]}): "
                                   f"{esc(note_of(it))}", st["n"]))
    docs = data.get("documents") or []
    if docs:
        story.append(Paragraph("Исходные документы", st["h2"]))
        for d in docs:
            story.append(Paragraph(f"<b>{esc(d.get('id', ''))}</b>: {esc(d.get('name', ''))}",
                                   st["n"]))
    story.append(Spacer(1, 8 * mm))
    for role in ("Составил", "Проверил"):
        story.append(Paragraph(f"{role}: ______________________ / ______________________ /"
                               "&nbsp;&nbsp;«___» ____________ 20__ г.", st["n"]))
        story.append(Spacer(1, 5 * mm))

    out_path = path.with_suffix(".content.pdf") if br["letterhead_pdf"] else path
    doc = SimpleDocTemplate(str(out_path), pagesize=A4, leftMargin=M["left"] * mm,
                            rightMargin=M["right"] * mm, topMargin=M["top"] * mm,
                            bottomMargin=M["bottom"] * mm, title=f"ВОР № {data['number']}",
                            author=comp.get("name", ""), subject=obj.get("name", ""))
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    if br["letterhead_pdf"]:
        merge_letterhead(out_path, Path(br["letterhead_pdf"]), path, br["letterhead_pages"])
        out_path.unlink()


def main() -> int:
    ap = argparse.ArgumentParser(description="Сборка ВОР в Excel и PDF")
    ap.add_argument("input")
    ap.add_argument("--out")
    a = ap.parse_args()
    try:
        data = load_vor(a.input)
        values, warnings = compute(data)
    except VorError as e:
        print("ОШИБКИ В ВОР. Файлы не созданы:")
        for x in e.errors:
            print(f"  - {x}")
        return 2
    inp = Path(a.input).resolve()
    out = Path(a.out).resolve() if a.out else inp.parent.parent / "output"
    out.mkdir(parents=True, exist_ok=True)
    base = out / f"VOR_{safe_name(str(data['number']))}"
    xlsx, pdf = base.with_suffix(".xlsx"), base.with_suffix(".pdf")
    build_xlsx(data, values, xlsx, warnings)
    br = brand_for(data)
    register_fonts(br["font_regular"], br["font_bold"])
    build_pdf(data, values, pdf)
    places = data.get("qty_decimals", 2)
    print(f"Excel: {xlsx}")
    print(f"PDF:   {pdf}")
    counts = {s: 0 for s in STATUSES}
    for *_, it in rows(data):
        counts[it["status"]] += 1
    print("Строк: " + ", ".join(f"{k} {v}" for k, v in counts.items()))
    for si, ii, it in rows(data):
        q = values.get(it["_id"])
        print(f"  {si}.{ii:<3} {it['name'][:55]:55} {it['unit']:7} "
              f"{'—' if q is None else fmt_q(q, places):>12}  {it['status']}")
    if warnings:
        print("Замечания:")
        for w in warnings:
            print(f"  - {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
