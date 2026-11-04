import copy
import json
import subprocess
import sys
from decimal import Decimal

import pytest

from akty_core import ActError, load_act
from conftest import ROOT

AK = ROOT / "skills" / "akty" / "scripts"
CONTRACT = {
    "number": "Д-1", "date": "2026-11-10", "subject": "Объект",
    "customer": {"name": "ТОО Заказчик", "bin": "1"}, "contractor": {"name": "ТОО Подрядчик", "bin": "2"},
    "vat_pct": 16, "vat_source": "НК РК", "amount_decimals": 2,
    "positions": [{"n": "1.1", "name": "Выемка", "unit": "м3", "qty": "100", "price": "2000"},
                  {"n": "1.2", "name": "Бетон", "unit": "м3", "qty": "50", "price": "40000"}],
    "totals": {"subtotal": "2200000", "vat": "352000", "total": "2552000"},
    "advance": {"paid": "765600", "offset_pct": 30, "source": "п. 4.2"},
    "retention": {"pct": 5, "source": "п. 7.1"},
}
ACT1 = {"contract": "contract.json", "number": "1", "date": "2026-11-30",
        "period": {"from": "2026-11-16", "to": "2026-11-30"},
        "lines": [{"n": "1.1", "qty": 60, "date": "2026-11-20", "evidence": "Исп. схема № 1"},
                  {"n": "1.2", "qty": 30, "evidence": "АОСР № 1"}]}
ACT2 = {"contract": "contract.json", "number": "2", "date": "2026-12-15", "previous": ["act_1.json"],
        "period": {"from": "2026-12-01", "to": "2026-12-15"},
        "lines": [{"n": "1.1", "qty": 40, "evidence": "Исп. схема № 2"},
                  {"n": "1.2", "qty": 20, "evidence": "АОСР № 2"}]}


@pytest.fixture
def proj(tmp_path):
    def write(name, data):
        (tmp_path / name).write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return tmp_path / name
    write("contract.json", CONTRACT)
    write("act_1.json", ACT1)
    write("act_2.json", ACT2)
    return tmp_path, write


def test_amounts_advance_and_retention(proj):
    tmp, _ = proj
    a = load_act(tmp / "act_1.json")
    c = a["_calc"]
    assert c["subtotal"] == Decimal("1320000.00")          # 60 × 2000 + 30 × 40000
    assert c["vat"] == Decimal("211200.00")
    assert c["offset"] == Decimal("459360.00")             # 30 % от 1 531 200
    assert c["retention"] == Decimal("76560.00")
    assert c["pay"] == c["total"] - c["offset"] - c["retention"]


def test_cumulative_and_advance_cap(proj):
    tmp, _ = proj
    a = load_act(tmp / "act_2.json")
    assert a["_prev_qty"] == {"1.1": Decimal(60), "1.2": Decimal(30)}
    assert a["_calc"]["offset"] == Decimal("306240.00")    # остаток аванса 765 600 − 459 360
    a1 = load_act(tmp / "act_1.json")
    assert a1["_calc"]["total"] + a["_calc"]["total"] == Decimal(CONTRACT["totals"]["total"])


@pytest.mark.parametrize("mut, text", [
    (lambda d: d["lines"][0].pop("evidence"), "нет подтверждения"),
    (lambda d: d["lines"][0].__setitem__("qty", 50), "больше договорного"),
    (lambda d: d["lines"][0].__setitem__("n", "9.9"), "нет в договоре"),
    (lambda d: d["lines"].append(dict(d["lines"][0])), "дважды"),
    (lambda d: d["lines"][0].__setitem__("date", "2026-11-20"), "вне отчётного периода"),
    (lambda d: d.__setitem__("date", "2026-12-10"), "раньше окончания"),
    (lambda d: d.__setitem__("number", "1"), "уже использован"),
    (lambda d: d["lines"][0].__setitem__("qty", 0), "больше нуля"),
])
def test_validation(proj, mut, text):
    tmp, write = proj
    d = copy.deepcopy(ACT2)
    mut(d)
    write("act_2.json", d)
    with pytest.raises(ActError) as e:
        load_act(tmp / "act_2.json")
    assert any(text in x for x in e.value.errors), e.value.errors


def test_overrun_needs_agreement(proj):
    tmp, write = proj
    d = copy.deepcopy(ACT2)
    d["lines"][0].update(qty=45, overrun_reason="Доп. соглашение № 1 от 10.12.2026")
    write("act_2.json", d)
    a = load_act(tmp / "act_2.json")
    assert any("превышение" in w for w in a["_warnings"])


def test_previous_must_list_all_earlier_acts(proj):
    tmp, write = proj
    d3 = copy.deepcopy(ACT2)
    d3.update(number="3", date="2026-12-20", period={"from": "2026-12-16", "to": "2026-12-20"},
              previous=["act_2.json"], lines=[{"n": "1.1", "qty": 1, "evidence": "x",
                                               "overrun_reason": "ДС № 2"}])
    write("act_3.json", d3)
    with pytest.raises(ActError) as e:
        load_act(tmp / "act_3.json")
    assert any("act_1.json" in x for x in e.value.errors)


def test_build_check_all_formats(proj):
    tmp, _ = proj
    out = tmp / "out"
    for i in (1, 2):
        r = subprocess.run([sys.executable, AK / "build_act.py", tmp / f"act_{i}.json", "--out", out],
                           capture_output=True, text=True)
        assert r.returncode == 0, r.stdout + r.stderr
        b = out / f"AKT_Д-1_{i}"
        r = subprocess.run([sys.executable, AK / "check_act.py", tmp / f"act_{i}.json",
                            b.with_suffix(".xlsx"), b.with_suffix(".pdf"), b.with_suffix(".docx")],
                           capture_output=True, text=True)
        assert r.returncode == 0 and "пройдена" in r.stdout, r.stdout + r.stderr


def test_check_detects_tampering(proj):
    tmp, _ = proj
    out = tmp / "out"
    subprocess.run([sys.executable, AK / "build_act.py", tmp / "act_1.json", "--out", out, "--no-docx"],
                   check=True, capture_output=True)
    from openpyxl import load_workbook
    x = out / "AKT_Д-1_1.xlsx"
    wb = load_workbook(x)
    ws = wb["Акт"]
    r = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 10).value == "1.1")
    ws.cell(r, 6).value = 70                                 # объём завысили вручную
    wb.save(x)
    r = subprocess.run([sys.executable, AK / "check_act.py", tmp / "act_1.json", x,
                        out / "AKT_Д-1_1.pdf"], capture_output=True, text=True)
    assert r.returncode == 1 and "ожидалось" in r.stdout


def test_make_contract_freezes_prices(tmp_path, make_prices, good_rows):
    from datetime import date
    prices = make_prices(good_rows)
    kp = {"number": "К-1", "date": "2026-11-03", "company": {"name": "ТОО П"}, "client": {"name": "ТОО З"},
          "object": {"name": "О"}, "price_list": prices.name,
          "rates": {"overhead_pct": 10, "profit_pct": 8, "vat_pct": 16, "vat_source": "т"},
          "sections": [{"title": "Р", "items": [{"code": "ЗР-001", "qty": 100}]}]}
    (tmp_path / "kp.json").write_text(json.dumps(kp, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "contract.json"
    r = subprocess.run([sys.executable, AK / "make_contract.py", tmp_path / "kp.json", "--out", out,
                        "--number", "Д-9", "--date", "2026-11-10"], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    c = json.loads(out.read_text(encoding="utf-8"))
    assert c["positions"][0]["price"] == "2197.80"            # 1850 × 1,1 × 1,08
    assert c["totals"]["total"] == "254944.80"
    r = subprocess.run([sys.executable, AK / "make_contract.py", tmp_path / "kp.json", "--out", out,
                        "--number", "Д-9", "--date", "2026-11-10"], capture_output=True, text=True)
    assert r.returncode == 1 and "не перезаписываю" in r.stdout
