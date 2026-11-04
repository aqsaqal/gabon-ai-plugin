#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "reportlab>=4.0", "pypdf>=4.0", "pillow>=10"]
# ///
"""Календарный график производства работ: Excel (формулы WORKDAY.INTL + диаграмма Ганта) и PDF.

    uv run build_grafik.py projects/<объект>/input/grafik.json [--out папка]
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from grafik_core import ALL_DAYS, GrafikError, _deps, compute, is_workday, load_grafik  # noqa: E402
from brand_core import brand_for  # noqa: E402
from build_kp import register_fonts, safe_name  # noqa: E402
from kp_core import fmt_date, fmt_qty  # noqa: E402

P = "'Параметры'"
HOL = "'Праздники'!$A$2:$A$200"
FIRST_GANTT_COL = 17  # колонка Q


def dep_text(t, by_id):
    out = []
    for d in _deps(t):
        s = by_id[d["id"]]["_n"]
        s += "" if d["type"] == "FS" else " НН"
        s += f" +{d['lag']}" if d["lag"] else ""
        out.append(s)
    return ", ".join(out)


def note_of(t):
    parts = [t.get("assumption"), t.get("duration_source"), t.get("note")]
    return "; ".join(str(x) for x in parts if x)


def build_xlsx(data, res, path, warnings):
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L
    from openpyxl.workbook.defined_name import DefinedName

    br = brand_for(data)
    C = {k: v.lstrip("#") for k, v in br["colors"].items()}
    tasks = data["tasks"]
    by_id = {t["_id"]: t for t in tasks}
    wb = Workbook()
    ws = wb.active
    ws.title = "График"
    ps = wb.create_sheet("Параметры")
    hs = wb.create_sheet("Праздники")
    ns = wb.create_sheet("Замечания")
    bold = Font(bold=True)
    thin = Side(style="thin", color="AAAAAA")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    ps.append(["Параметр", "Значение", "Пояснение"])
    ps.append(["Начало работ", data_start(data), "Первый рабочий день не раньше этой даты"])
    ps.append(["Выходные (маска пн…вс)", res["weekend"], "1 — выходной. 0000011 — сб и вс"])
    ps.append(["Смен в день", res["shifts"], ""])
    ps.append(["Часов в смене", float(res["shift_hours"]), ""])
    ps["B2"].number_format = "dd.mm.yyyy"
    for col, w in zip("ABC", (28, 14, 50)):
        ps.column_dimensions[col].width = w
    hs.append(["Праздничные дни"])
    for h in res["holidays"]:
        hs.append([h])
        hs.cell(hs.max_row, 1).number_format = "dd.mm.yyyy"
    hs["C1"] = "Основание: " + str((data.get("calendar") or {}).get("holidays_source", "не задано"))
    hs.column_dimensions["A"].width = 14

    comp = data.get("company") or {}
    ws["A1"] = comp.get("name", "")
    ws["A1"].font = Font(bold=True, color=C["primary"], size=12)
    ws["A2"] = f"КАЛЕНДАРНЫЙ ГРАФИК ПРОИЗВОДСТВА РАБОТ № {data['number']} от {fmt_date(data['date'])}"
    ws["A2"].font = Font(bold=True, size=14, color=C["primary"])
    ws["A3"] = f"Объект: {data['object'].get('name', '')}"
    ws["A4"] = (f"Начало {fmt_date(res['start'])}, окончание {fmt_date(res['finish'])}, "
                f"{res['workdays']} раб. дн. ({res['calendar_days']} календ. дн.)")
    hr = 6
    hdr = ["№", "Работа", "Ед.", "Объём", "Выработка бригады в смену", "Бригад",
           "Трудозатраты, чел.-ч", "Рабочих", "Длит., дн.", "Начало", "Окончание",
           "Предшественники", "Резерв, раб. дн.", "Критич.", "Календарь", "Источник, допущение"]
    for ci, h in enumerate(hdr, 1):
        c = ws.cell(hr, ci, h)
        c.font = Font(bold=True, color=C["header_text"])
        c.fill = PatternFill("solid", fgColor=C["primary"])
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        c.border = box
    ws.row_dimensions[hr].height = 45
    row_of = {t["_id"]: hr + 1 + i for i, t in enumerate(tasks)}
    for t in tasks:
        r = row_of[t["_id"]]
        cal_txt = "календарные" if t["_cal"] == ALL_DAYS else "рабочие"
        wk = f'"{ALL_DAYS}"' if t["_cal"] == ALL_DAYS else f"{P}!$B$3"
        hol = "" if t["_cal"] == ALL_DAYS else f",{HOL}"
        dur = {"productivity": f"=ROUNDUP(ROUND(D{r}/(E{r}*F{r}*{P}!$B$4),6),0)",
               "labor": f"=ROUNDUP(ROUND(G{r}/(H{r}*{P}!$B$5*{P}!$B$4),6),0)",
               "duration": t["_dur"], "milestone": 0}[t["_mode"]]
        terms = [f"_xlfn.WORKDAY.INTL({P}!$B$2-1,1,{wk}{hol})"]
        if t.get("not_before"):
            nb = date.fromisoformat(str(t["not_before"]))
            terms.append(f"_xlfn.WORKDAY.INTL(DATE({nb.year},{nb.month},{nb.day})-1,1,{wk}{hol})")
        for d in _deps(t):
            pr = row_of[d["id"]]
            if d["type"] == "SS":
                terms.append(f"_xlfn.WORKDAY.INTL(J{pr},{d['lag']},{wk}{hol})")
            elif t["_dur"] == 0:
                terms.append(f"_xlfn.WORKDAY.INTL(K{pr},{d['lag']},{wk}{hol})")
            else:
                terms.append(f"_xlfn.WORKDAY.INTL(K{pr},{1 + d['lag']},{wk}{hol})")
        start_f = "=" + (terms[0] if len(terms) == 1 else f"MAX({','.join(terms)})")
        fin_f = f"=J{r}" if t["_mode"] == "milestone" else f"=_xlfn.WORKDAY.INTL(J{r},I{r}-1,{wk}{hol})"
        vals = [t["_n"], t["name"], t["_unit"] or "",
                float(t["_qty"]) if t["_qty"] is not None else None,
                float(t["_out"]) if t["_mode"] == "productivity" else None,
                t.get("_crews") if t["_mode"] == "productivity" else None,
                float(t["_hours"]) if t["_mode"] == "labor" else None,
                t.get("_workers") if t["_mode"] == "labor" else None,
                dur, start_f, fin_f, dep_text(t, by_id), t["_float"],
                "да" if t["_critical"] else "", cal_txt,
                "; ".join(x for x in (t["_qsrc"], t.get("_lsrc", ""),
                                      (t.get("productivity") or {}).get("source", ""), note_of(t)) if x)]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(r, ci, v)
            c.border = box
            c.alignment = Alignment(wrap_text=ci in (2, 12, 16), vertical="center")
        for col in "JK":
            ws[f"{col}{r}"].number_format = "dd.mm.yy"
        ws[f"D{r}"].number_format = "#,##0.00"
        ws[f"G{r}"].number_format = "#,##0.0"
        if t["_critical"]:
            ws[f"B{r}"].font = Font(bold=True, color="C00000")
    last = hr + len(tasks)

    # Диаграмма Ганта по дням (запас 25 % на случай изменения объёмов)
    span = (res["finish"] - res["start"]).days + 1
    days = span + max(10, span // 4)
    for i in range(days):
        d = res["start"] + timedelta(days=i)
        col = FIRST_GANTT_COL + i
        c = ws.cell(hr, col, d)
        c.number_format = "dd.mm"
        c.alignment = Alignment(text_rotation=90, horizontal="center")
        c.font = Font(size=8, color=C["header_text"])
        c.fill = PatternFill("solid", fgColor=C["primary"])
        ws.column_dimensions[L(col)].width = 2.6
        if not is_workday(d, res["weekend"], frozenset(res["holidays"])):
            for r in range(hr + 1, last + 1):
                ws.cell(r, col).fill = PatternFill("solid", fgColor="EDEDED")
    g1, g2 = L(FIRST_GANTT_COL), L(FIRST_GANTT_COL + days - 1)
    rng = f"{g1}{hr + 1}:{g2}{last}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND({g1}${hr}>=$J{hr + 1},{g1}${hr}<=$K{hr + 1},$N{hr + 1}="да")'],
        fill=PatternFill("solid", fgColor="C00000"), stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({g1}${hr}>=$J{hr + 1},{g1}${hr}<=$K{hr + 1})"],
        fill=PatternFill("solid", fgColor=C["primary"])))
    for col, w in zip("ABCDEFGHIJKLMNOP", (5, 42, 7, 10, 10, 7, 10, 8, 7, 10, 10, 13, 8, 7, 11, 45)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = ws.cell(hr + 1, 3)
    ws.print_title_rows = f"{hr}:{hr}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    r = last + 2
    ws.cell(r, 2, "Красным: критический путь (резерв 0). Серым: выходные и праздники. "
                  "Резерв и критичность рассчитаны при формировании; при изменении объёмов "
                  "даты пересчитываются формулами, а резерв обновится после пересборки.").font = \
        Font(italic=True, size=9)
    r += 2
    for role in ("Составил", "Проверил", "Согласовано (заказчик)"):
        ws.cell(r, 2, f"{role}: ____________________ / ____________________ /")
        r += 2
    wb.defined_names["GR_FINISH"] = DefinedName("GR_FINISH", attr_text=f"'График'!$K${hr + 1}:$K${last}")
    ns.column_dimensions["A"].width = 120
    for w in warnings or ["нет"]:
        ns.append([f"• {w}"])
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def data_start(data):
    return date.fromisoformat(str(data["start"]))


def build_pdf(data, res, path):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as cv

    br = brand_for(data)
    Cc = {k: colors.HexColor(v) for k, v in br["colors"].items()}
    red = colors.HexColor("#C00000")
    tasks = data["tasks"]
    by_id = {t["_id"]: t for t in tasks}
    W, H = landscape(A4)
    left, right, top, bottom = 12 * mm, 12 * mm, 14 * mm, 14 * mm
    cols = [("№", 8), ("Работа", 82), ("Длит.", 11), ("Начало", 17), ("Оконч.", 17), ("Резерв", 11)]
    table_w = sum(w for _, w in cols) * mm
    chart_x = left + table_w + 2 * mm
    chart_w = W - right - chart_x
    span = (res["finish"] - res["start"]).days + 1
    day_w = chart_w / span
    row_h = 6.2 * mm
    hol = frozenset(res["holidays"])
    c = cv.Canvas(str(path), pagesize=(W, H))
    c.setTitle(f"График № {data['number']}")
    c.setAuthor((data.get("company") or {}).get("name", ""))
    comp = data.get("company") or {}
    sample = bool(data.get("sample"))

    def header(page):
        if sample:
            c.saveState()
            c.setFont("KP-Bold", 70)
            c.setFillColor(colors.Color(0.85, 0.1, 0.1, alpha=0.08))
            c.translate(W / 2, H / 2)
            c.rotate(30)
            c.drawCentredString(0, 0, "ОБРАЗЕЦ")
            c.restoreState()
        y = H - top
        c.setFillColor(Cc["primary"])
        c.setFont("KP-Bold", 12)
        c.drawString(left, y, f"КАЛЕНДАРНЫЙ ГРАФИК ПРОИЗВОДСТВА РАБОТ № {data['number']}")
        c.setFont("KP", 8.5)
        c.setFillColor(colors.black)
        c.drawRightString(W - right, y, comp.get("name", ""))
        y -= 4.5 * mm
        c.drawString(left, y, f"Объект: {data['object'].get('name', '')}")
        y -= 4 * mm
        c.drawString(left, y, f"Начало {fmt_date(res['start'])}, окончание {fmt_date(res['finish'])}: "
                              f"{res['workdays']} рабочих дней ({res['calendar_days']} календарных). "
                              f"Красным: критический путь.")
        c.setFont("KP", 7)
        c.setFillColor(colors.grey)
        c.drawString(left, bottom - 6 * mm, f"График № {data['number']} от {fmt_date(data['date'])}")
        c.drawRightString(W - right, bottom - 6 * mm, f"Стр. {page}")
        return y - 4 * mm

    def grid_head(y):
        # шапка таблицы
        c.setFillColor(Cc["primary"])
        c.rect(left, y - 9 * mm, table_w + 2 * mm + chart_w, 9 * mm, fill=1, stroke=0)
        c.setFillColor(Cc["header_text"])
        c.setFont("KP-Bold", 7)
        x = left
        for name, w in cols:
            c.drawCentredString(x + w * mm / 2, y - 5.5 * mm, name)
            x += w * mm
        # шкала: месяцы/недели
        c.setFont("KP", 6)
        d = res["start"]
        for i in range(span):
            d = res["start"] + timedelta(days=i)
            x = chart_x + i * day_w
            if d.weekday() == 0 or i == 0:
                c.drawString(x + 0.5, y - 4 * mm, d.strftime("%d.%m"))
            if day_w > 2.2 * mm:
                c.drawCentredString(x + day_w / 2, y - 8 * mm, str(d.day))
        return y - 9 * mm

    page = 1
    y = grid_head(header(page))
    body_top = y
    for t in tasks:
        if y - row_h < bottom + 30 * mm and t is not tasks[-1] or y - row_h < bottom:
            draw_shading(c, body_top, y, chart_x, day_w, span, res, hol)
            c.showPage()
            page += 1
            y = grid_head(header(page))
            body_top = y
        c.setStrokeColor(colors.HexColor("#DDDDDD"))
        c.line(left, y - row_h, W - right, y - row_h)
        c.setFillColor(red if t["_critical"] else colors.black)
        c.setFont("KP-Bold" if t["_critical"] else "KP", 7)
        x = left
        name = fit(t["name"], "KP-Bold" if t["_critical"] else "KP", 7, cols[1][1] * mm - 2.5 * mm)
        vals = [t["_n"], name, "веха" if t["_mode"] == "milestone" else
                f"{t['_dur']}{' к' if t['_cal'] == ALL_DAYS else ''}",
                fmt_date(t["_es"]), fmt_date(t["_ef"]), str(t["_float"])]
        for (cname, w), v in zip(cols, vals):
            if cname == "Работа":
                c.drawString(x + 1.2 * mm, y - row_h + 2 * mm, v)
            else:
                c.drawCentredString(x + w * mm / 2, y - row_h + 2 * mm, v)
            x += w * mm
        bx = chart_x + (t["_es"] - res["start"]).days * day_w
        if t["_mode"] == "milestone":
            cx, cy, s = bx + day_w / 2, y - row_h / 2, 2.2 * mm
            p = c.beginPath()
            p.moveTo(cx, cy + s)
            p.lineTo(cx + s, cy)
            p.lineTo(cx, cy - s)
            p.lineTo(cx - s, cy)
            p.close()
            c.setFillColor(red if t["_critical"] else Cc["primary"])
            c.drawPath(p, fill=1, stroke=0)
        else:
            bw = ((t["_ef"] - t["_es"]).days + 1) * day_w
            c.setFillColor(red if t["_critical"] else Cc["primary"])
            c.roundRect(bx, y - row_h + 1.4 * mm, bw, row_h - 2.8 * mm, 1, fill=1, stroke=0)
            if t["_float"]:
                fx = chart_x + (t["_lf"] - res["start"]).days * day_w + day_w
                c.setStrokeColor(colors.grey)
                c.setDash(1, 1.5)
                c.line(bx + bw, y - row_h / 2, fx, y - row_h / 2)
                c.setDash()
        y -= row_h
    draw_shading(c, body_top, y, chart_x, day_w, span, res, hol)

    # примечания
    notes = [(t["_n"], note_of(t)) for t in tasks if note_of(t)]
    y -= 5 * mm
    c.setFillColor(colors.black)
    lines = []
    if notes:
        lines.append(("KP-Bold", "Допущения и обоснования длительностей:"))
        lines += [("KP", f"{n}. {txt}") for n, txt in notes]
    calx = data.get("calendar") or {}
    lines.append(("KP", f"Режим: {res['shifts']} смена по {res['shift_hours']} ч, выходные по маске "
                        f"{res['weekend']}. Праздники: "
                        + (", ".join(fmt_date(h) for h in res["holidays"]) or "не заданы")
                        + (f" ({calx.get('holidays_source')})" if calx.get("holidays_source") else "")))
    lines.append(("KP", "Составил: ____________________     Проверил: ____________________     "
                        "Согласовано (заказчик): ____________________"))
    from reportlab.lib.utils import simpleSplit
    for font, txt in lines:
        for part in simpleSplit(txt, font, 7.5, W - left - right):
            if y < bottom + 4 * mm:
                c.showPage()
                page += 1
                y = header(page) - 4 * mm
            c.setFont(font, 7.5)
            c.drawString(left, y, part)
            y -= 4.2 * mm
    c.save()


def fit(text: str, font: str, size: float, width: float) -> str:
    """Обрезает текст по фактической ширине, чтобы он не заходил на соседнюю колонку."""
    from reportlab.pdfbase.pdfmetrics import stringWidth
    if stringWidth(text, font, size) <= width:
        return text
    while text and stringWidth(text + "…", font, size) > width:
        text = text[:-1]
    return text.rstrip() + "…"


def draw_shading(c, top, bottom_y, chart_x, day_w, span, res, hol):
    from reportlab.lib import colors
    c.saveState()
    c.setFillColor(colors.Color(0, 0, 0, alpha=0.06))
    for i in range(span):
        d = res["start"] + timedelta(days=i)
        if not is_workday(d, res["weekend"], hol):
            c.rect(chart_x + i * day_w, bottom_y, day_w, top - bottom_y, fill=1, stroke=0)
    c.restoreState()


def main() -> int:
    ap = argparse.ArgumentParser(description="Календарный график")
    ap.add_argument("input")
    ap.add_argument("--out")
    a = ap.parse_args()
    try:
        data = load_grafik(a.input)
        res, warnings = compute(data)
    except GrafikError as e:
        print("ОШИБКИ В ГРАФИКЕ. Файлы не созданы:")
        for x in e.errors:
            print(f"  - {x}")
        return 2
    inp = Path(a.input).resolve()
    out = Path(a.out).resolve() if a.out else inp.parent.parent / "output"
    out.mkdir(parents=True, exist_ok=True)
    base = out / f"GRAFIK_{safe_name(str(data['number']))}"
    xlsx, pdf = base.with_suffix(".xlsx"), base.with_suffix(".pdf")
    build_xlsx(data, res, xlsx, warnings)
    br = brand_for(data)
    register_fonts(br["font_regular"], br["font_bold"])
    build_pdf(data, res, pdf)
    print(f"Excel: {xlsx}\nPDF:   {pdf}")
    print(f"Начало {fmt_date(res['start'])}, окончание {fmt_date(res['finish'])}: "
          f"{res['workdays']} раб. дн., {res['calendar_days']} календ. дн.")
    for t in data["tasks"]:
        print(f"  {t['_n']:>3} {t['name'][:52]:52} {t['_dur']:>3} дн. "
              f"{fmt_date(t['_es'])} – {fmt_date(t['_ef'])}  резерв {t['_float']:>2}"
              f"{'  КРИТИЧ.' if t['_critical'] else ''}")
    if warnings:
        print("Замечания:")
        for w in warnings:
            print(f"  - {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
