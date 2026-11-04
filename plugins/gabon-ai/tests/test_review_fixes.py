"""Регрессионные тесты по итогам ревью 03.11.2026."""
import copy
import json
import subprocess
import sys
from datetime import date

import pytest

from conftest import ROOT
from kp_core import InputError, load_input, normalize_unit, validate

SCRIPTS = ROOT / "skills" / "kp" / "scripts"
BASE = {"number": "R-1", "date": "2026-11-03", "company": {"name": "Т"}, "client": {"name": "З"},
        "object": {"name": "О"},
        "rates": {"overhead_pct": 0, "profit_pct": 0, "vat_pct": 16, "vat_source": "т"},
        "sections": [{"title": "Р", "items": [
            {"name": "А", "unit": "м2", "qty": 10, "price": 1000, "kind": "работа", "source": "т"}]}]}


def _write(tmp_path, data):
    p = tmp_path / "kp.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return p


@pytest.mark.parametrize("mutate", [
    lambda d: d.__setitem__("object", "строка"),
    lambda d: d.__setitem__("rates", None),
    lambda d: d.__setitem__("client", "строка"),
    lambda d: d["sections"][0]["items"].append("строка"),
    lambda d: d["sections"][0]["items"][0].__setitem__("price", "NaN"),
    lambda d: d["sections"][0]["items"][0].__setitem__("price", "Infinity"),
    lambda d: d["sections"][0]["items"][0].__setitem__("qty", "1e30"),
    lambda d: d.__setitem__("valid_days", "abc"),
])
def test_bad_input_gives_message_not_crash(tmp_path, mutate):
    d = copy.deepcopy(BASE)
    mutate(d)
    with pytest.raises(InputError):
        validate(load_input(_write(tmp_path, d)))


def test_units_are_normalized(tmp_path):
    d = copy.deepcopy(BASE)
    d["sections"][0]["items"][0]["unit"] = "м²"
    data = load_input(_write(tmp_path, d))
    assert validate(data) == []
    assert data["sections"][0]["items"][0]["unit"] == "м2"
    assert normalize_unit("кв.м") == "м2" and normalize_unit("М.П") == "п.м"


def test_checker_refuses_code_in_formulas():
    from openpyxl import Workbook
    from check_kp import FormulaEvaluator
    wb = Workbook()
    ws = wb.active
    ws.title = "КП"
    ws["A1"] = "=(1).__class__.__name__"
    ws["A2"] = "=ROUND(2.5*3,0)"
    ev = FormulaEvaluator(wb)
    assert str(ev.value("КП", "A1")).startswith("#ОШИБКА")
    assert ev.value("КП", "A2") == 8


def test_price_override_keeps_unit_and_kind(make_prices, good_rows):
    from prices_core import apply_prices, load_price_list
    pl = load_price_list(make_prices(good_rows))
    data = {"sections": [{"title": "Р", "items": [
        {"code": "ЗР-001", "qty": 1, "price": 1700, "price_override_reason": "скидка"}]}]}
    errors, _ = apply_prices(data, pl, date(2026, 11, 3))
    it = data["sections"][0]["items"][0]
    assert errors == []
    assert it["unit"] == "м3" and it["kind"] == "механизм" and it["name"]
    assert "Цена прайса 1850" in it["source"]


def test_vat_exempt_wording(tmp_path):
    d = copy.deepcopy(BASE)
    d["rates"].update(vat_pct=0, vat_source="Компания не является плательщиком НДС")
    kp = _write(tmp_path, d)
    out = tmp_path / "out"
    r = subprocess.run([sys.executable, SCRIPTS / "build_kp.py", kp, "--out", out],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    r = subprocess.run([sys.executable, SCRIPTS / "check_kp.py", kp, out / "KP_R-1.xlsx",
                        out / "KP_R-1.pdf"], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    from pypdf import PdfReader
    text = PdfReader(str(out / "KP_R-1.pdf")).pages[0].extract_text()
    assert "НДС не облагается" in text and "с НДС" not in text and "без НДС." in text


def test_template_rejects_percent_typed_without_sign(tmp_path):
    from openpyxl import load_workbook
    from make_template import PAR, build
    t = tmp_path / "t.xlsx"
    build(t, 5)
    wb = load_workbook(t)
    row = next(i for i, p in enumerate(PAR, 2) if p[1] == "vat_pct")
    assert any(f"B{row}" in str(dv.sqref) and dv.type == "decimal"
               for dv in wb["Параметры"].data_validations.dataValidation)
    wb["Параметры"][f"B{row}"] = 16
    wb.save(t)
    r = subprocess.run([sys.executable, SCRIPTS / "xlsx_to_kp.py", t], capture_output=True, text=True)
    assert r.returncode != 0 and "16%" in (r.stdout + r.stderr)
