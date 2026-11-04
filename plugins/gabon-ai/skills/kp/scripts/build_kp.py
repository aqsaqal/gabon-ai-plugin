#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "reportlab>=4.0", "pypdf>=4.0", "pillow>=10", "python-docx>=1.1"]
# ///
"""Сборка коммерческого предложения (КП): Excel с формулами + PDF + Word.

Запуск:
    uv run build_kp.py путь/к/kp.json [--out папка]
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "brand" / "scripts"))
from brand_core import brand_for  # noqa: E402
from kp_core import (KINDS, InputError, amount_in_words, compute, fmt_date, vat_exempt,  # noqa: E402
                     fmt_money, fmt_pct, fmt_qty, item_numbers, load_input, places_of,
                     section_title, subtotal_label, terms_lines, validate, valid_until, D)

PARAMS = "'Параметры'"


# ----------------------------------------------------------------- Excel ---
def build_xlsx(data: dict, calc: dict, path: Path, warnings: list[str]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.workbook.defined_name import DefinedName

    comp, obj, rates = data["company"], data["object"], data["rates"]
    client = data.get("client") or {}
    n = places_of(data)
    nums = item_numbers(data)
    wb = Workbook()
    ws = wb.active
    ws.title = "КП"
    ps = wb.create_sheet("Параметры")
    ns = wb.create_sheet("Допущения")

    br = brand_for(data)
    C = {k: v.lstrip("#") for k, v in br["colors"].items()}
    bold = Font(bold=True)
    head_fill = PatternFill("solid", fgColor=C["primary"])
    sec_fill = PatternFill("solid", fgColor=C["section"])
    sub_fill = PatternFill("solid", fgColor=C["subtotal"])
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    MONEY = '#,##0.00'
    QTY = '#,##0.000'

    # --- Параметры (входные ставки, на них ссылаются формулы) ---
    ps.append(["Параметр", "Значение", "Источник / примечание"])
    ps.append(["Накладные расходы, % от прямых затрат",
               float(D(rates["overhead_pct"]) / 100), rates.get("overhead_source", "")])
    ps.append(["Прибыль, % от (прямые затраты + накладные)",
               float(D(rates["profit_pct"]) / 100), rates.get("profit_source", "")])
    ps.append(["НДС, % от суммы без НДС", float(D(rates["vat_pct"]) / 100), rates["vat_source"]])
    ps.append([])
    ps.append(["Валюта", data.get("currency", "KZT"), "Все цены в позициях указаны без НДС"])
    for c in ("B2", "B3", "B4"):
        ps[c].number_format = "0.00%"
    for c in ps[1]:
        c.font = bold
    ps.column_dimensions["A"].width = 46
    ps.column_dimensions["B"].width = 14
    ps.column_dimensions["C"].width = 70

    # --- Лист КП ---
    widths = {"A": 7, "B": 52, "C": 9, "D": 12, "E": 16, "F": 18, "G": 13, "H": 42}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws["A1"] = comp["name"]
    ws["A1"].font = Font(bold=True, size=13, color=C["primary"])
    if br["logo"]:
        from openpyxl.drawing.image import Image as XLImage
        img = XLImage(br["logo"])
        h = 60
        img.width, img.height = int(img.width * h / img.height), h
        ws.add_image(img, "H1")
    req = [f"БИН {comp['bin']}" if comp.get("bin") else "", comp.get("address", "")]
    ws["A2"] = ", ".join(x for x in req if x)
    contacts = [comp.get("phone", ""), comp.get("email", "")]
    ws["A3"] = ", ".join(x for x in contacts if x)

    ws.merge_cells("A5:H5")
    ws["A5"] = f"{br['title']} № {data['number']} от {fmt_date(data['date'])}"
    ws["A5"].font = Font(bold=True, size=14, color=C["primary"])
    ws["A5"].alignment = Alignment(horizontal="center")

    info = [("Заказчик:", client.get("name") or "не указан"), ("Объект:", obj["name"]),
            ("Адрес объекта:", obj.get("address", "")), ("Основание:", data.get("basis", ""))]
    vu = valid_until(data)
    if vu:
        info.append(("Срок действия:", f"до {fmt_date(vu)}"))
    r = 6
    for label, val in info:
        ws.cell(r, 1, label).font = bold
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.cell(r, 2, val)
        r += 1
    if data.get("sample"):
        ws.cell(r, 1, "ОБРАЗЕЦ. Тестовые данные, не для отправки заказчику.").font = \
            Font(bold=True, color="C00000")
        r += 1

    r += 1
    header = ["№", "Наименование работ и затрат", "Ед. изм.", "Кол-во",
              "Цена без НДС, тг", "Сумма без НДС, тг", "Вид затрат", "Источник цены"]
    for ci, h in enumerate(header, 1):
        c = ws.cell(r, ci, h)
        c.font = Font(bold=True, color=C["header_text"])
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = box
    header_row = r
    r += 1
    first_item = r
    subtotal_cells = []

    for si, s in enumerate(data["sections"], 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c = ws.cell(r, 1, section_title(data, si, s["title"]))
        c.font = bold
        c.fill = sec_fill
        r += 1
        start = r
        for ii, it in enumerate(s["items"], 1):
            row = [nums[si - 1][ii - 1], it["name"], it["unit"], float(D(it["qty"])),
                   float(D(it["price"])), f"=ROUND(D{r}*E{r},{n})", it["kind"], it["source"]]
            for ci, v in enumerate(row, 1):
                c = ws.cell(r, ci, v)
                c.border = box
                c.alignment = wrap
            ws.cell(r, 4).number_format = QTY
            ws.cell(r, 5).number_format = MONEY
            ws.cell(r, 6).number_format = MONEY
            r += 1
        end = r - 1
        ws.cell(r, 2, subtotal_label(data, si, s["title"])).font = bold
        c = ws.cell(r, 6, f"=SUM(F{start}:F{end})")
        c.font = bold
        c.number_format = MONEY
        for ci in range(1, 9):
            ws.cell(r, ci).fill = sub_fill
        subtotal_cells.append(f"F{r}")
        r += 1
    last_item = r - 1

    r += 1
    names = {}

    def total_row(label, formula, name, rate_ref=None, strong=False):
        nonlocal r
        ws.cell(r, 2, label).font = Font(bold=True, size=12 if strong else 11)
        if rate_ref:
            c = ws.cell(r, 5, f"={rate_ref}")
            c.number_format = "0.00%"
        c = ws.cell(r, 6, formula)
        c.number_format = MONEY
        c.font = Font(bold=True, size=12 if strong else 11)
        names[name] = f"'КП'!$F${r}"
        row = r
        r += 1
        return row

    rd = total_row("Итого прямые затраты", "=" + "+".join(subtotal_cells), "KP_DIRECT")
    ro = total_row("Накладные расходы", f"=ROUND(F{rd}*{PARAMS}!$B$2,{n})", "KP_OVERHEAD",
                   f"{PARAMS}!$B$2")
    rp = total_row("Прибыль", f"=ROUND((F{rd}+F{ro})*{PARAMS}!$B$3,{n})", "KP_PROFIT",
                   f"{PARAMS}!$B$3")
    rs = total_row("Итого без НДС", f"=F{rd}+F{ro}+F{rp}", "KP_SUBTOTAL")
    no_vat = vat_exempt(data)
    rv = total_row("НДС (не облагается)" if no_vat else "НДС",
                   f"=ROUND(F{rs}*{PARAMS}!$B$4,{n})", "KP_VAT", f"{PARAMS}!$B$4")
    rt = total_row("ВСЕГО (без НДС)" if no_vat else "ВСЕГО с НДС", f"=F{rs}+F{rv}",
                   "KP_TOTAL", strong=True)
    ws.cell(rt, 6).fill = PatternFill("solid", fgColor=C["total"])
    ws.cell(r, 2, f"Сумма прописью: {amount_in_words(calc['total'])}, "
                  f"{'без НДС' if no_vat else 'с учётом НДС'}. Текст не пересчитывается "
                  "формулами: при изменении цифр пересоберите КП.").font = Font(italic=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    r += 1

    r += 1
    ws.cell(r, 2, "Структура прямых затрат").font = bold
    r += 1
    kind_rows = []
    for kind, label in KINDS.items():
        ws.cell(r, 2, label)
        c = ws.cell(r, 6, f'=SUMIF($G${first_item}:$G${last_item},"{kind}",'
                          f'$F${first_item}:$F${last_item})')
        c.number_format = MONEY
        kind_rows.append(r)
        r += 1
    ws.cell(r, 2, "Контроль: прямые затраты минус сумма по видам (должно быть 0)").font = \
        Font(italic=True)
    c = ws.cell(r, 6, f"=F{rd}-SUM(F{kind_rows[0]}:F{kind_rows[-1]})")
    c.number_format = MONEY
    names["KP_CONTROL"] = f"'КП'!$F${r}"
    r += 2
    ws.cell(r, 2, "Документ подготовлен для проверки и утверждения ответственным специалистом.") \
        .font = Font(italic=True, color="666666")

    for name, ref in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)

    ws.freeze_panes = ws.cell(header_row + 1, 1)
    ws.print_title_rows = f"{header_row}:{header_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # --- Допущения ---
    ns.column_dimensions["A"].width = 120
    blocks = [("Условия", terms_lines(data)),
              ("Допущения", data.get("assumptions", [])),
              ("Не включено в стоимость", data.get("exclusions", [])),
              ("Недостающие данные (требуют уточнения)", data.get("missing_data", [])),
              ("Замечания автоматической проверки", warnings)]
    row = 1
    for title, items in blocks:
        ns.cell(row, 1, title).font = bold
        row += 1
        for x in items or ["нет"]:
            ns.cell(row, 1, f"• {x}").alignment = Alignment(wrap_text=True)
            row += 1
        row += 1

    wb.calculation.fullCalcOnLoad = True  # Excel пересчитает формулы при открытии
    wb.save(path)


# ------------------------------------------------------------------- PDF ---
FONT_CANDIDATES = [
    ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ("/usr/share/fonts/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
]


def register_fonts(regular: str | None, bold: str | None) -> None:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    pairs = [(regular, bold)] if regular and bold else FONT_CANDIDATES
    for reg, bld in pairs:
        if Path(reg).exists() and Path(bld).exists():
            pdfmetrics.registerFont(TTFont("KP", reg))
            pdfmetrics.registerFont(TTFont("KP-Bold", bld))
            return
    raise SystemExit("Не найден шрифт с кириллицей. Укажите --font-regular и --font-bold "
                     "(файлы .ttf, например Arial или DejaVu Sans).")


def totals_lines(data: dict, calc: dict) -> list[tuple[str, object, bool]]:
    """Итоговый блок КП: (подпись, сумма или None, жирный). Один источник для PDF и Word."""
    rates = data["rates"]
    tot = []
    if D(rates["overhead_pct"]) or D(rates["profit_pct"]):
        tot.append(("Итого прямые затраты", calc["direct"], False))
        if D(rates["overhead_pct"]):
            tot.append((f"Накладные расходы ({fmt_pct(rates['overhead_pct'])})", calc["overhead"], False))
        if D(rates["profit_pct"]):
            tot.append((f"Прибыль ({fmt_pct(rates['profit_pct'])})", calc["profit"], False))
    if vat_exempt(data):
        tot += [("Итого без НДС", calc["subtotal"], True),
                ("НДС не облагается", None, False),
                ("ВСЕГО к оплате, тенге (без НДС)", calc["total"], True)]
    else:
        tot += [("Итого без НДС", calc["subtotal"], True),
                (f"НДС ({fmt_pct(rates['vat_pct'])})", calc["vat"], False),
                ("ВСЕГО с НДС, тенге", calc["total"], True)]
    return tot


def requisites_lines(comp: dict) -> list[str]:
    lines = []
    for label, key in (("БИН", "bin"), ("Адрес", "address"), ("Тел.", "phone"), ("E-mail", "email")):
        if comp.get(key):
            lines.append(f"{label}: {comp[key]}")
    bank = comp.get("bank") or {}
    if bank:
        parts = [f"{k} {bank[v]}" for k, v in (("ИИК", "iik"), ("БИК", "bik"), ("КБе", "kbe")) if bank.get(v)]
        if bank.get("name"):
            parts.insert(0, bank["name"])
        lines.append(", ".join(parts))
    return lines


def build_docx(data: dict, calc: dict, path: Path) -> None:
    """Редактируемое КП в Word: то же содержание и те же суммы, что в PDF."""
    import docx
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor

    br = brand_for(data)
    C = {k: v.lstrip("#") for k, v in br["colors"].items()}
    comp, obj, rates = data["company"], data["object"], data["rates"]
    client = data.get("client") or {}
    nums = item_numbers(data)
    rgb = lambda h: RGBColor.from_string(h.upper())  # noqa: E731

    def shade(cell, hex_):
        tc = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_)
        tc.append(shd)

    def merged(a, b):
        """Объединяет ячейки и убирает пустые абзацы, оставшиеся от объединения."""
        m = a.merge(b)
        for extra in m.paragraphs[1:]:
            extra._element.getparent().remove(extra._element)
        return m

    def fix_widths(table, widths, borders=True):
        """Ширины колонок и рамки задаются явно: так их одинаково показывают Word и просмотрщики."""
        table.autofit = False
        tbl = table._tbl
        grid = tbl.tblGrid
        for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
            gc.set(qn("w:w"), str(int(w.twips)))
        for tr in tbl.tr_lst:          # ширина объединённой ячейки = сумма колонок, которые она занимает
            col = 0
            for tc in tr.tc_lst:
                span = tc.grid_span
                tc.width = sum(widths[col:col + span], start=type(widths[0])(0))
                col += span
        pr = tbl.tblPr
        for old in pr.findall(qn("w:tblW")):
            pr.remove(old)
        tw = OxmlElement("w:tblW")
        tw.set(qn("w:type"), "dxa")
        tw.set(qn("w:w"), str(int(sum(w.twips for w in widths))))
        pr.append(tw)
        if borders:
            b = OxmlElement("w:tblBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                e = OxmlElement(f"w:{side}")
                e.set(qn("w:val"), "single")
                e.set(qn("w:sz"), "4")
                e.set(qn("w:color"), "999999")
                b.append(e)
            pr.append(b)

    def para(text="", bold=False, size=None, align=None, color=None, italic=False, after=2):
        p = d.add_paragraph()
        if text:
            r = p.add_run(text)
            r.bold, r.italic = bold, italic
            if size:
                r.font.size = Pt(size)
            if color:
                r.font.color.rgb = rgb(color)
        if align:
            p.alignment = align
        p.paragraph_format.space_after = Pt(after)
        return p

    def cell_text(cell, text, bold=False, align=None, size=8.5, color=None):
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(str(text))
        r.bold = bold
        r.font.size = Pt(size)
        if color:
            r.font.color.rgb = rgb(color)
        if align:
            p.alignment = align
        p.paragraph_format.space_after = Pt(0)

    R, CEN = WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER
    d = docx.Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    M = br["margins_mm"]
    sec.left_margin, sec.right_margin = Mm(M["left"]), Mm(M["right"])
    sec.top_margin, sec.bottom_margin = Mm(15), Mm(15)
    st = d.styles["Normal"]
    st.font.name, st.font.size = "Arial", Pt(9.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    st.paragraph_format.space_after = Pt(0)
    if data.get("sample"):
        para("ОБРАЗЕЦ. Тестовые данные, не для отправки заказчику.", bold=True, color="C00000", align=CEN)
    # шапка: логотип и реквизиты (фирменный бланк PDF в Word не переносится)
    head = d.add_table(rows=1, cols=2 if br["logo"] else 1)
    fix_widths(head, [Mm(50), Mm(130)] if br["logo"] else [Mm(180)], borders=False)
    cells = head.rows[0].cells
    req_cell = cells[-1]
    if br["logo"]:
        logo_cell, req_cell = (cells[1], cells[0]) if br["logo_position"] == "right" else (cells[0], cells[1])
        logo_cell.paragraphs[0].add_run().add_picture(br["logo"], width=Mm(min(br["logo_width_mm"], 45)))
        if br["logo_position"] == "right":
            logo_cell.paragraphs[0].alignment = R
    cell_text(req_cell, comp["name"], bold=True, size=12, color=C["primary"])
    for line in requisites_lines(comp):
        p = req_cell.add_paragraph()
        r = p.add_run(line)
        r.font.size = Pt(8)
        p.paragraph_format.space_after = Pt(0)
    para()
    para(f"{br['title']} № {data['number']}", bold=True, size=14, align=CEN, color=C["primary"])
    para(f"от {fmt_date(data['date'])}", align=CEN, after=8)

    info = [("Заказчик", client.get("name") or "____________________ (не указан)"),
            ("Объект", obj["name"]), ("Адрес объекта", obj.get("address", "")),
            ("Основание", data.get("basis", ""))]
    vu = valid_until(data)
    if vu:
        info.append(("Срок действия", f"до {fmt_date(vu)}"))
    it = d.add_table(rows=0, cols=2)
    for a, b in info:
        if b:
            row = it.add_row().cells
            cell_text(row[0], a, bold=True, size=9.5)
            cell_text(row[1], b, size=9.5)
    fix_widths(it, [Mm(35), Mm(145)], borders=False)
    para()
    if data.get("intro"):
        para(data["intro"])

    t = d.add_table(rows=1, cols=6)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Mm(11), Mm(79), Mm(14), Mm(20), Mm(27), Mm(29)]
    for i, h in enumerate(("№", "Наименование", "Ед.", "Кол-во", "Цена без НДС, тг",
                           "Сумма без НДС, тг")):
        c = t.rows[0].cells[i]
        cell_text(c, h, bold=True, align=CEN, size=8, color=C["header_text"])
        shade(c, C["primary"])
    for si, (s, cs) in enumerate(zip(data["sections"], calc["sections"]), 1):
        row = t.add_row().cells
        m = merged(row[0], row[5])
        cell_text(m, section_title(data, si, s["title"]), bold=True)
        shade(m, C["section"])
        for ii, (item, amt) in enumerate(zip(s["items"], cs["items"]), 1):
            row = t.add_row().cells
            for j, (v, al) in enumerate(((nums[si - 1][ii - 1], CEN), (item["name"], None),
                                         (item["unit"], CEN), (fmt_qty(item["qty"]), R),
                                         (fmt_money(D(item["price"])), R), (fmt_money(amt), R))):
                cell_text(row[j], v, align=al)
        row = t.add_row().cells
        m = merged(row[1], row[4])
        cell_text(m, subtotal_label(data, si, s["title"]), bold=True)
        cell_text(row[5], fmt_money(cs["subtotal"]), bold=True, align=R)
        for c in (row[0], m, row[5]):
            shade(c, C["subtotal"])
    # Итоги — последними строками той же таблицы: надёжно отображаются в любом редакторе
    tot = totals_lines(data, calc)
    for k, (label, v, b) in enumerate(tot):
        row = t.add_row().cells
        m = merged(row[0], row[4])
        cell_text(m, label, bold=b, align=R, size=9.5 if b else 8.5)
        cell_text(row[5], "—" if v is None else fmt_money(v), bold=b, align=R, size=9.5 if b else 8.5)
        if k == len(tot) - 1:
            shade(m, C["total"])
            shade(row[5], C["total"])
    fix_widths(t, widths)

    para()
    no_vat = vat_exempt(data)
    para(f"Сумма прописью: {amount_in_words(calc['total'])}, "
         f"{'без НДС' if no_vat else 'с учётом НДС'}.", italic=True)

    def block(title, lines):
        if lines:
            para(title, bold=True, size=11, color=C["primary"], after=3)
            for x in lines:
                p = d.add_paragraph(str(x), style="List Bullet")
                p.paragraph_format.space_after = Pt(1)
    block("Условия", terms_lines(data))
    block("Допущения", data.get("assumptions"))
    block("Не включено в стоимость", data.get("exclusions"))
    if rates.get("vat_source"):
        para(f"НДС не начисляется. Основание: {rates['vat_source']}." if no_vat else
             f"Ставка НДС: {fmt_pct(rates['vat_pct'])}. Основание: {rates['vat_source']}.", size=8)
    para()
    para()
    sig = d.add_table(rows=2, cols=3)
    cell_text(sig.rows[0].cells[0], comp.get("director_position", "Руководитель"), size=10)
    cell_text(sig.rows[0].cells[1], "______________________", align=CEN, size=10)
    cell_text(sig.rows[0].cells[2], comp.get("director") or "______________________", size=10)
    cell_text(sig.rows[1].cells[1], "подпись", align=CEN, size=7.5)
    fix_widths(sig, [Mm(55), Mm(60), Mm(65)], borders=False)
    para("М.П.", size=9)

    foot = sec.footer.paragraphs[0]
    foot.text = f"КП № {data['number']} от {fmt_date(data['date'])}" + \
        (f"    {br['footer_text']}" if br["footer_text"] else "")
    for r in foot.runs:
        r.font.size = Pt(7.5)
        r.font.color.rgb = rgb("808080")
    d.core_properties.title = f"КП № {data['number']}"
    d.core_properties.author = comp["name"]
    d.save(path)


def build_pdf(data: dict, calc: dict, path: Path) -> None:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer,
                                    Table, TableStyle)

    comp, obj, rates = data["company"], data["object"], data["rates"]
    client = data.get("client") or {}
    nums = item_numbers(data)
    br = brand_for(data)
    C = {k: colors.HexColor(v) for k, v in br["colors"].items()}
    navy = C["primary"]
    M = br["margins_mm"]
    st = {
        "n": ParagraphStyle("n", fontName="KP", fontSize=9, leading=11.5),
        "s": ParagraphStyle("s", fontName="KP", fontSize=8, leading=10),
        "c": ParagraphStyle("c", fontName="KP", fontSize=8, leading=10, alignment=TA_CENTER),
        "r": ParagraphStyle("r", fontName="KP", fontSize=8, leading=10, alignment=TA_RIGHT),
        "b": ParagraphStyle("b", fontName="KP-Bold", fontSize=9, leading=11.5),
        "h": ParagraphStyle("h", fontName="KP-Bold", fontSize=8, leading=10,
                            alignment=TA_CENTER, textColor=C["header_text"]),
        "title": ParagraphStyle("t", fontName="KP-Bold", fontSize=14, leading=18,
                                alignment=TA_CENTER, textColor=navy, spaceAfter=2 * mm),
        "h2": ParagraphStyle("h2", fontName="KP-Bold", fontSize=10.5, leading=13,
                             textColor=navy, spaceBefore=4 * mm, spaceAfter=1.5 * mm),
        "co": ParagraphStyle("co", fontName="KP-Bold", fontSize=12, leading=15, textColor=navy),
    }

    def esc(s) -> str:
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    sample = bool(data.get("sample"))

    def on_page(canvas, doc):
        canvas.saveState()
        if sample:
            canvas.setFont("KP-Bold", 70)
            canvas.setFillColor(colors.Color(0.85, 0.1, 0.1, alpha=0.10))
            canvas.translate(A4[0] / 2, A4[1] / 2)
            canvas.rotate(45)
            canvas.drawCentredString(0, 0, "ОБРАЗЕЦ")
            canvas.rotate(-45)
            canvas.translate(-A4[0] / 2, -A4[1] / 2)
        canvas.setFont("KP", 7.5)
        canvas.setFillColor(colors.grey)
        y = max(6, M["bottom"] - 8) * mm
        canvas.drawString(M["left"] * mm, y, f"КП № {data['number']} от {fmt_date(data['date'])}")
        if br["footer_text"]:
            canvas.drawCentredString(A4[0] / 2, y, br["footer_text"])
        canvas.drawRightString(A4[0] - M["right"] * mm, y, f"Стр. {doc.page}")
        canvas.restoreState()

    story = []

    # Шапка: логотип + реквизиты
    req_lines = [f"<font name='KP-Bold' size='12' color='#1F3864'>{esc(comp['name'])}</font>"]
    for label, key in (("БИН", "bin"), ("Адрес", "address"), ("Тел.", "phone"), ("E-mail", "email")):
        if comp.get(key):
            req_lines.append(f"{label}: {esc(comp[key])}")
    bank = comp.get("bank") or {}
    if bank:
        parts = [f"{k} {esc(bank[v])}" for k, v in (("ИИК", "iik"), ("БИК", "bik"), ("КБе", "kbe"))
                 if bank.get(v)]
        if bank.get("name"):
            parts.insert(0, esc(bank["name"]))
        req_lines.append(", ".join(parts))
    W = (A4[0] / mm - M["left"] - M["right"]) * mm  # ширина области текста
    req = Paragraph("<br/>".join(req_lines), st["s"]) if br["show_requisites"] else ""
    cells, widths = [], []
    if br["logo"]:
        img = Image(br["logo"])
        ratio = img.imageHeight / img.imageWidth
        lw = br["logo_width_mm"] * mm
        img.drawWidth, img.drawHeight = lw, lw * ratio
        cells, widths = [img, req], [lw + 5 * mm, W - lw - 5 * mm]
        if br["logo_position"] == "right":
            cells, widths = cells[::-1], widths[::-1]
            img.hAlign = "RIGHT"
    elif br["show_requisites"]:
        cells, widths = [req], [W]
    if cells:
        head = Table([cells], colWidths=widths)
        head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                  ("LINEBELOW", (0, 0), (-1, 0), 1.2, navy),
                                  ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                  ("BOTTOMPADDING", (0, 0), (-1, -1), 4 * mm)]))
        if br["logo_position"] == "right" and br["logo"]:
            head.setStyle(TableStyle([("ALIGN", (1, 0), (1, 0), "RIGHT")]))
        story += [head, Spacer(1, 5 * mm)]

    story.append(Paragraph(f"{esc(br['title'])} № {esc(data['number'])}", st["title"]))
    story.append(Paragraph(f"от {fmt_date(data['date'])}", ParagraphStyle(
        "d", parent=st["n"], alignment=TA_CENTER, spaceAfter=4 * mm)))

    info = [("Заказчик", client.get("name") or "____________________ (не указан)"),
            ("Объект", obj["name"]),
            ("Адрес объекта", obj.get("address", "")), ("Основание", data.get("basis", ""))]
    vu = valid_until(data)
    if vu:
        info.append(("Срок действия", f"до {fmt_date(vu)}"))
    t = Table([[Paragraph(esc(a), st["b"]), Paragraph(esc(b), st["n"])] for a, b in info if b],
              colWidths=[35 * mm, 145 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
    story += [t, Spacer(1, 4 * mm)]

    if data.get("intro"):
        story += [Paragraph(esc(data["intro"]), st["n"]), Spacer(1, 3 * mm)]

    # Таблица позиций
    rows = [[Paragraph(h, st["h"]) for h in
             ("№", "Наименование", "Ед.", "Кол-во", "Цена без НДС, тг", "Сумма без НДС, тг")]]
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    for si, (s, cs) in enumerate(zip(data["sections"], calc["sections"]), 1):
        rows.append([Paragraph(esc(section_title(data, si, s["title"])), st["b"]), "", "", "", "", ""])
        n = len(rows) - 1
        style += [("SPAN", (0, n), (-1, n)), ("BACKGROUND", (0, n), (-1, n), C["section"])]
        for ii, (it, amt) in enumerate(zip(s["items"], cs["items"]), 1):
            rows.append([Paragraph(nums[si - 1][ii - 1], st["c"]), Paragraph(esc(it["name"]), st["s"]),
                         Paragraph(esc(it["unit"]), st["c"]), Paragraph(fmt_qty(it["qty"]), st["r"]),
                         Paragraph(fmt_money(D(it["price"])), st["r"]),
                         Paragraph(fmt_money(amt), st["r"])])
        rows.append(["", Paragraph(esc(subtotal_label(data, si, s["title"])), st["b"]), "", "", "",
                     Paragraph(f"<b>{fmt_money(cs['subtotal'])}</b>",
                               ParagraphStyle("rb", parent=st["r"], fontName="KP-Bold"))])
        n = len(rows) - 1
        style += [("SPAN", (1, n), (4, n)), ("BACKGROUND", (0, n), (-1, n), C["subtotal"])]
    k = W / (180 * mm)
    items = Table(rows, colWidths=[x * k * mm for x in (11, 79, 14, 20, 27, 29)],
                  repeatRows=1)
    items.setStyle(TableStyle(style))
    story += [items, Spacer(1, 4 * mm)]

    # Итоги
    rb = ParagraphStyle("rb2", parent=st["n"], fontName="KP-Bold", alignment=TA_RIGHT)
    rn = ParagraphStyle("rn2", parent=st["n"], alignment=TA_RIGHT)
    tot = totals_lines(data, calc)
    i_sub = next(i for i, (lbl, _, _) in enumerate(tot) if lbl == "Итого без НДС")
    i_tot = len(tot) - 1
    no_vat = vat_exempt(data)
    tt = Table([[Paragraph(esc(a), rb if b_ else rn),
                 Paragraph("—" if v is None else fmt_money(v), rb if b_ else rn)]
                for a, v, b_ in tot], colWidths=[W * 2 / 3, W / 3])
    tt.setStyle(TableStyle([
        ("LINEABOVE", (0, i_sub), (-1, i_sub), 0.6, colors.grey),
        ("BACKGROUND", (0, i_tot), (-1, i_tot), C["total"]),
        ("LINEABOVE", (0, i_tot), (-1, i_tot), 1, navy),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(KeepTogether([tt]))
    story += [Spacer(1, 2 * mm), Paragraph(
        f"<i>Сумма прописью: {esc(amount_in_words(calc['total']))}, "
        f"{'без НДС' if no_vat else 'с учётом НДС'}.</i>", st["n"])]

    def bullet_block(title, items_):
        if not items_:
            return
        story.append(Paragraph(title, st["h2"]))
        for x in items_:
            story.append(Paragraph(f"• {esc(x)}", st["n"]))

    bullet_block("Условия", terms_lines(data))
    bullet_block("Допущения", data.get("assumptions"))
    bullet_block("Не включено в стоимость", data.get("exclusions"))
    if rates.get("vat_source"):
        story += [Spacer(1, 2 * mm), Paragraph(
            (f"НДС не начисляется. Основание: {esc(rates['vat_source'])}." if no_vat else
             f"Ставка НДС: {fmt_pct(rates['vat_pct'])}. Основание: {esc(rates['vat_source'])}."),
            st["s"])]

    # Подпись (без факсимиле: подписывает ответственное лицо)
    signer = comp.get("director") or "______________________"
    position = comp.get("director_position", "Руководитель")
    sig = Table([[Paragraph(esc(position), st["n"]), "", Paragraph(esc(signer), st["n"])],
                 ["", Paragraph("подпись", st["c"]), ""]],
                colWidths=[60 * mm, 50 * mm, 70 * mm])
    sig.setStyle(TableStyle([("LINEBELOW", (1, 0), (1, 0), 0.6, colors.black),
                             ("VALIGN", (0, 0), (-1, -1), "BOTTOM")]))
    story += [Spacer(1, 10 * mm), KeepTogether([sig]), Spacer(1, 3 * mm),
              Paragraph("М.П.", st["s"])]

    out_path = path.with_suffix(".content.pdf") if br["letterhead_pdf"] else path
    doc = SimpleDocTemplate(str(out_path), pagesize=A4, leftMargin=M["left"] * mm,
                            rightMargin=M["right"] * mm, topMargin=M["top"] * mm,
                            bottomMargin=M["bottom"] * mm,
                            title=f"КП № {data['number']}", author=comp["name"],
                            subject=obj["name"])
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    if br["letterhead_pdf"]:
        merge_letterhead(out_path, Path(br["letterhead_pdf"]), path, br["letterhead_pages"])
        out_path.unlink()


def merge_letterhead(content: Path, letterhead: Path, out: Path, pages: str) -> None:
    """Накладывает содержимое КП на фирменный бланк (первая страница или все)."""
    from pypdf import PageObject, PdfReader, PdfWriter
    lh = PdfReader(str(letterhead)).pages[0]
    writer = PdfWriter()
    for i, page in enumerate(PdfReader(str(content)).pages):
        if i == 0 or pages == "all":
            blank = PageObject.create_blank_page(width=page.mediabox.width,
                                                 height=page.mediabox.height)
            blank.merge_page(lh)
            blank.merge_page(page)
            page = blank
        writer.add_page(page)
    for k, v in (PdfReader(str(content)).metadata or {}).items():
        writer.add_metadata({k: v})
    with open(out, "wb") as f:
        writer.write(f)


# ------------------------------------------------------------------ main ---
def safe_name(s: str) -> str:
    return re.sub(r"[^\w\-.]+", "_", s, flags=re.UNICODE).strip("_")


def main() -> int:
    ap = argparse.ArgumentParser(description="Сборка КП в Excel и PDF")
    ap.add_argument("input", help="JSON с исходными данными КП")
    ap.add_argument("--out", help="папка для результата (по умолчанию ../output рядом с input)")
    ap.add_argument("--font-regular")
    ap.add_argument("--font-bold")
    ap.add_argument("--no-docx", action="store_true", help="не создавать Word")
    args = ap.parse_args()

    try:
        data = load_input(args.input)
        warnings = validate(data)
    except InputError as e:
        print("ОШИБКИ В ИСХОДНЫХ ДАННЫХ. Файлы не созданы:")
        for err in e.errors:
            print(f"  - {err}")
        return 2

    calc = compute(data)
    inp = Path(args.input).resolve()
    out = Path(args.out).resolve() if args.out else inp.parent.parent / "output"
    out.mkdir(parents=True, exist_ok=True)
    base = out / f"KP_{safe_name(str(data['number']))}"
    xlsx, pdf = base.with_suffix(".xlsx"), base.with_suffix(".pdf")

    build_xlsx(data, calc, xlsx, warnings)
    br = brand_for(data)
    register_fonts(args.font_regular or br["font_regular"], args.font_bold or br["font_bold"])
    build_pdf(data, calc, pdf)
    docx_path = base.with_suffix(".docx")
    if not args.no_docx:
        build_docx(data, calc, docx_path)

    if data.get("_brand"):
        used = ", ".join(data["_brand_used"]) or "только оформление"
        print(f"Фирменный шаблон: {data['_brand']['path']} (взято: {used})")
    print(f"Excel: {xlsx}")
    print(f"PDF:   {pdf}")
    if not args.no_docx:
        print(f"Word:  {docx_path}")
    print(f"Итого прямые затраты: {fmt_money(calc['direct'])}")
    print(f"Итого без НДС:        {fmt_money(calc['subtotal'])}")
    print(f"НДС:                  {fmt_money(calc['vat'])}")
    print(f"ВСЕГО с НДС:          {fmt_money(calc['total'])}")
    if warnings:
        print("Замечания:")
        for w in warnings:
            print(f"  - {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
