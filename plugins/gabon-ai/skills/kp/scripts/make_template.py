#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl==3.1.5", "pillow==13.0.0", "reportlab==5.1.0", "pypdf==6.21.4"]
# ///
"""Универсальный шаблон КП в Excel для ручного заполнения.

Пользователь заполняет листы «Реквизиты», «Параметры» и строки на листе «КП».
Все суммы считаются формулами Excel. Печать / «Сохранить как PDF» из Excel работает
сразу (область печати A4 задана). Фирменный PDF собирается командой xlsx_to_kp.py.

Запуск:
    uv run make_template.py путь/KP_template.xlsx [--rows 60] [--brand data/brand.json]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from kp_core import KINDS  # noqa: E402

REQ = [  # (подпись, ключ company.json, ключ bank)
    ("Наименование компании", "name", None), ("БИН", "bin", None), ("Адрес", "address", None),
    ("Телефон", "phone", None), ("E-mail", "email", None), ("Банк", None, "name"),
    ("БИК", None, "bik"), ("КБе", None, "kbe"), ("ИИК", None, "iik"),
    ("Должность подписанта", "director_position", None), ("ФИО подписанта", "director", None),
]
PAR = [  # (подпись, ключ, пример/по умолчанию, формат)
    ("Номер КП", "number", "", None),
    ("Дата КП", "date", None, "dd.mm.yyyy"),
    ("Срок действия, календарных дней", "valid_days", 14, None),
    ("Заказчик", "client", "", None),
    ("Объект", "object", "", None),
    ("Адрес объекта", "address", "", None),
    ("Вводный текст", "intro", "", None),
    ("Накладные расходы, % от прямых затрат (0, если цены полные)", "overhead_pct", 0, "0.00%"),
    ("Прибыль, % от (прямые + накладные) (0, если цены полные)", "profit_pct", 0, "0.00%"),
    ("НДС, %", "vat_pct", None, "0.00%"),
    ("Источник ставки НДС (норма закона и дата проверки)", "vat_source", "", None),
    ("Округление сумм: 0 — до тенге, 2 — до тиын", "amount_decimals", 0, None),
    ("Образец (да/нет)", "sample", "нет", None),
]
TERMS_START = len(PAR) + 3  # строка на листе «Параметры», с которой идут условия
TERMS_N = 8


def par_row(key: str) -> int:
    return next(i for i, p in enumerate(PAR, 2) if p[1] == key)


def build(path: Path, n_rows: int, brand: dict | None = None) -> None:
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    ws = wb.active
    ws.title = "КП"
    rq = wb.create_sheet("Реквизиты")
    pr = wb.create_sheet("Параметры")
    hl = wb.create_sheet("Инструкция")

    bold = Font(bold=True)
    inp = PatternFill("solid", fgColor="FFF9E5")  # поля для ввода
    C = {k: v.lstrip("#") for k, v in (brand or {}).get("colors", {}).items()}
    head = PatternFill("solid", fgColor=C.get("primary", "DCE3EE"))
    head_font = Font(bold=True, color=C["header_text"]) if brand else Font(bold=True)
    thin = Side(style="thin", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    MONEY = '#,##0.00'

    # --- Реквизиты ---
    rq.append(["Реквизит", "Значение (заполняется только по указанию руководства)"])
    for label, _, _ in REQ:
        rq.append([label, ""])
    rq["B11"] = "Директор"
    for r in range(2, len(REQ) + 2):
        rq.cell(r, 2).fill = inp
    rq.column_dimensions["A"].width = 26
    rq.column_dimensions["B"].width = 70
    for c in rq[1]:
        c.font = bold

    # --- Параметры ---
    pr.append(["Параметр", "Значение"])
    for label, key, default, fmt in PAR:
        pr.append([label, default])
        c = pr.cell(pr.max_row, 2)
        c.fill = inp
        if fmt:
            c.number_format = fmt
    pr.cell(TERMS_START - 1, 1, "Условия КП (каждое с новой строки)").font = bold
    for i in range(TERMS_N):
        pr.cell(TERMS_START + i, 1, f"Условие {i + 1}")
        pr.cell(TERMS_START + i, 2).fill = inp
    pr.column_dimensions["A"].width = 62
    pr.column_dimensions["B"].width = 80
    for c in pr[1]:
        c.font = bold
    dv = DataValidation(type="list", formula1='"да,нет"', allow_blank=True)
    pr.add_data_validation(dv)
    dv.add(f"B{par_row('sample')}")
    dvp = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                         allow_blank=True, showErrorMessage=True)
    dvp.errorTitle = "Процент"
    dvp.error = "Введите процент со знаком %, например 16%. Число 16 без знака % означает 1600 %."
    pr.add_data_validation(dvp)
    for key in ("overhead_pct", "profit_pct", "vat_pct"):
        dvp.add(f"B{par_row(key)}")
    dv0 = DataValidation(type="list", formula1='"0,2"', allow_blank=False)
    pr.add_data_validation(dv0)
    dv0.add(f"B{par_row('amount_decimals')}")

    P = lambda key: f"'Параметры'!$B${par_row(key)}"  # noqa: E731
    R = lambda i: f"'Реквизиты'!$B${i + 2}"  # noqa: E731

    # --- Лист КП ---
    for col, w in {"A": 6, "B": 50, "C": 8, "D": 12, "E": 15, "F": 18,
                   "G": 11, "H": 22, "I": 44}.items():
        ws.column_dimensions[col].width = w

    ws["A1"] = f"={R(0)}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f'="Адрес: "&{R(2)}'
    ws["A3"] = f'="БИН: "&{R(1)}'
    ws["A4"] = f'="Банк: "&{R(5)}&"    БИК: "&{R(6)}&"    КБе: "&{R(7)}'
    ws["A5"] = f'="ИИК: "&{R(8)}'
    for r in range(2, 6):
        ws.cell(r, 1).font = Font(size=9)
    for ci in range(1, 7):
        ws.cell(5, ci).border = Border(bottom=Side(style="medium"))

    def centered(row, value, font):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        c = ws.cell(row, 1, value)
        c.font = font
        c.alignment = Alignment(horizontal="center")
        return c

    title = (brand or {}).get("title") or "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ"
    centered(7, title, Font(bold=True, size=14, color=C.get("primary")))
    if brand and brand.get("logo"):
        from openpyxl.drawing.image import Image as XLImage
        img = XLImage(brand["logo"])
        img.width, img.height = int(img.width * 60 / img.height), 60
        ws.add_image(img, "E1")
    centered(8, f'="№ "&{P("number")}', Font(size=11))
    c = centered(9, f"={P('date')}", Font(size=11))
    c.number_format = '"от "dd.mm.yyyy'

    info = [("Заказчик:", f'=IF({P("client")}="","не указан",{P("client")})'),
            ("Объект:", f"={P('object')}"), ("Адрес:", f"={P('address')}"),
            ("", f"={P('intro')}")]
    r = 11
    for label, f in info:
        ws.cell(r, 1, label).font = bold
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.cell(r, 2, f).alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    hr = r + 1
    headers = ["№", "Наименование работ", "Ед. изм.", "Кол-во", "Цена за ед. без НДС, тг",
               "Стоимость без НДС, тг", "Тип строки", "Вид затрат", "Источник цены и дата"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(hr, ci, h)
        c.font = head_font
        c.fill = head
        c.border = box
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[hr].height = 32

    first, last = hr + 1, hr + n_rows
    for r in range(first, last + 1):
        ws.cell(r, 1, f'=IF($G{r}="позиция",COUNTIF($G${first}:$G{r},"позиция"),"")')
        ws.cell(r, 6, f'=IF($G{r}="позиция",ROUND($D{r}*$E{r},{P("amount_decimals")}),"")')
        for ci in range(1, 10):
            c = ws.cell(r, ci)
            c.border = box
            c.alignment = Alignment(wrap_text=True, vertical="center",
                                    horizontal="center" if ci in (1, 3) else None)
        for ci in (2, 3, 4, 5, 7, 8, 9):
            ws.cell(r, ci).fill = inp
        ws.cell(r, 4).number_format = '#,##0.00'
        ws.cell(r, 5).number_format = MONEY
        ws.cell(r, 6).number_format = MONEY

    rng = f"A{first}:F{last}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'$G{first}="раздел"'], font=Font(bold=True),
        fill=PatternFill("solid", fgColor="EFEFEF")))
    ws.conditional_formatting.add(f"I{first}:I{last}", FormulaRule(
        formula=[f'AND($G{first}="позиция",$I{first}="")'],
        fill=PatternFill("solid", fgColor="F8CBAD")))

    dv_t = DataValidation(type="list", formula1='"позиция,раздел"', allow_blank=True)
    dv_k = DataValidation(type="list", formula1='"' + ",".join(KINDS) + '"', allow_blank=True)
    ws.add_data_validation(dv_t)
    ws.add_data_validation(dv_k)
    dv_t.add(f"G{first}:G{last}")
    dv_k.add(f"H{first}:H{last}")

    # Итоги
    r = last + 2
    rows = {}

    def total(label, formula, key, rate=None, big=False):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(r, 2, label).font = Font(bold=True, size=13 if big else 11)
        ws.cell(r, 2).alignment = Alignment(horizontal="right")
        if rate:
            ws.cell(r, 5, f"={rate}").number_format = "0.00%"
        c = ws.cell(r, 6, formula)
        c.number_format = MONEY
        c.font = Font(bold=True, size=13 if big else 11)
        rows[key] = r
        r += 1

    d = P("amount_decimals")
    total("Итого прямые затраты:", f"=SUM(F{first}:F{last})", "direct")
    total("Накладные расходы:", f"=ROUND(F{rows['direct']}*{P('overhead_pct')},{d})",
          "overhead", P("overhead_pct"))
    total("Прибыль:", f"=ROUND((F{rows['direct']}+F{rows['overhead']})*{P('profit_pct')},{d})",
          "profit", P("profit_pct"))
    total("Итого (без НДС):", f"=F{rows['direct']}+F{rows['overhead']}+F{rows['profit']}",
          "subtotal")
    total("НДС:", f"=ROUND(F{rows['subtotal']}*{P('vat_pct')},{d})", "vat", P("vat_pct"))
    total("ВСЕГО С НДС:", f"=F{rows['subtotal']}+F{rows['vat']}", "total", big=True)
    ws.cell(rows["total"], 6).fill = PatternFill("solid", fgColor=C.get("total", "FFF2CC"))

    # Контроль (вне области печати)
    ctl = rows["direct"]
    ws.cell(ctl, 8, "Позиций без источника цены:")
    ws.cell(ctl, 9, f'=COUNTIFS(G{first}:G{last},"позиция",I{first}:I{last},"")')
    ws.cell(ctl + 1, 8, "Позиций без вида затрат:")
    ws.cell(ctl + 1, 9, f'=COUNTIFS(G{first}:G{last},"позиция",H{first}:H{last},"")')
    ws.cell(ctl + 2, 8, "Позиций без кол-ва или цены:")
    ws.cell(ctl + 2, 9, f'=COUNTIFS(G{first}:G{last},"позиция",D{first}:D{last},"")'
                        f'+COUNTIFS(G{first}:G{last},"позиция",E{first}:E{last},"")')
    for i in range(3):
        ws.cell(ctl + i, 8).font = Font(italic=True, color="C00000")

    r += 1
    ws.cell(r, 1, "Условия:").font = bold
    r += 1
    for i in range(TERMS_N):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        src = f"'Параметры'!$B${TERMS_START + i}"
        ws.cell(r, 1, f'=IF({src}="","",{src})').alignment = Alignment(wrap_text=True)
        r += 1
    r += 1
    ws.cell(r, 1, "От Исполнителя:")
    r += 1
    ws.cell(r, 1, f'={R(9)}&" "&{R(0)}').font = bold
    r += 2
    ws.cell(r, 1, f'="_______________________  /  "&{R(10)}')
    ws.cell(r, 5, "М.П.")
    r += 1
    ws.cell(r, 1, "(подпись)").font = Font(size=8)
    end = r

    ws.print_area = f"A1:F{end}"
    ws.print_title_rows = f"{hr}:{hr}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.right.text = "Стр. &P из &N"
    ws.freeze_panes = ws.cell(hr + 1, 1)

    # --- Инструкция ---
    hl.column_dimensions["A"].width = 110
    for line in [
        "КАК ЗАПОЛНИТЬ ШАБЛОН КП",
        "1. «Реквизиты»: данные компании-исполнителя. Заполняются только по указанию руководства.",
        "2. «Параметры»: номер, дата, заказчик, объект, ставки, условия. Ставку НДС подтвердите у бухгалтера.",
        "3. «КП»: жёлтые ячейки. В колонке G выберите «раздел» (заголовок группы) или «позиция».",
        "   Для позиции заполните наименование, ед. изм., кол-во, цену без НДС, вид затрат и источник цены.",
        "4. №, стоимость и итоги считаются формулами. Не вписывайте суммы вручную.",
        "5. Красная ячейка в колонке I означает, что у позиции нет источника цены.",
        "6. Колонки G–I служебные и не печатаются.",
        "7. PDF: Файл → Сохранить как PDF (область печати A4 уже задана),",
        "   или фирменный PDF через GABON AI: команда xlsx_to_kp.py.",
        "8. Цены из прошлых КП не являются рыночными расценками. Указывайте реальный источник и дату.",
        "9. КП — проект для проверки и утверждения ответственным специалистом.",
    ]:
        hl.append([line])
    hl["A1"].font = Font(bold=True, size=12)

    if brand:
        apply_brand_defaults(pr, brand)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def apply_brand_defaults(pr, brand: dict) -> None:
    """Предзаполняет лист «Параметры» текстами и ставками из фирменного шаблона."""
    d = brand.get("defaults") or {}
    vals = {k: d[k] for k in ("valid_days", "intro", "amount_decimals") if k in d}
    for k, v in (d.get("rates") or {}).items():
        if k.endswith("_pct"):
            vals[k] = float(v) / 100
        elif k == "vat_source":
            vals[k] = v
    for i, (_, key, _, _) in enumerate(PAR, 2):
        if key in vals:
            pr.cell(i, 2).value = vals[key]
    terms = d.get("terms") or []
    if isinstance(terms, dict):
        from kp_core import terms_lines
        terms = terms_lines({"terms": terms})
    for i, t in enumerate(terms[:TERMS_N]):
        pr.cell(TERMS_START + i, 2).value = t


def main() -> int:
    ap = argparse.ArgumentParser(description="Создать пустой шаблон КП в Excel")
    ap.add_argument("out")
    ap.add_argument("--rows", type=int, default=60, help="число строк для позиций")
    ap.add_argument("--brand", help="фирменный шаблон brand.json")
    a = ap.parse_args()
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    brand = None
    if a.brand:
        sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "brand" / "scripts"))
        from brand_core import BrandError, load_brand
        try:
            brand, _ = load_brand(a.brand)
        except BrandError as e:
            print("Ошибки фирменного шаблона:\n  - " + "\n  - ".join(e.errors))
            return 1
    build(out, a.rows, brand)
    print(f"Шаблон: {out.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
