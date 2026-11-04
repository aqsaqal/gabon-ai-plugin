import copy
import json
import subprocess
import sys
from datetime import date
from decimal import Decimal

import pytest

from conftest import ROOT
from smeta_core import SmetaError, compute, load_smeta

SM = ROOT / "skills" / "smeta" / "scripts"
RATES = {"overhead_pct": 10, "profit_pct": 8, "vat_pct": 16, "vat_source": "т"}


@pytest.fixture
def env(tmp_path, make_prices):
    make_prices([
        ("ТР-1", "Рабочий", "чел.-ч", 3500, "работа", "src", date(2026, 10, 1), ""),
        ("МШ-1", "Насос", "маш.-ч", 38000, "механизм", "src", date(2026, 10, 1), ""),
        ("МТ-1", "Бетон", "м3", 34500, "материал", "src", date(2026, 10, 1), ""),
    ])
    (tmp_path / "cards.json").write_text(json.dumps({"cards": {
        "К-1": {"name": "Бетонирование", "unit": "м3", "source": "Калькуляция № 1",
                "resources": [{"type": "материал", "code": "МТ-1", "norm": 1.015},
                              {"type": "труд", "code": "ТР-1", "norm": 0.6},
                              {"type": "машины", "code": "МШ-1", "norm": 0.125}]}}},
        ensure_ascii=False), encoding="utf-8")
    base = {"number": "С-1", "date": "2026-11-03", "object": {"name": "О"},
            "company": {"name": "ТОО Тест"}, "cards": "cards.json", "price_list": "prices.xlsx",
            "rates": dict(RATES),
            "sections": [{"title": "Р", "positions": [
                {"card": "К-1", "qty": 100, "qty_source": "ВОР стр. 1"},
                {"name": "Доставка", "unit": "рейс", "qty": 4, "qty_source": "расчёт",
                 "resources": [{"type": "транспорт", "name": "Самосвал", "unit": "рейс", "norm": 1,
                                "price": 25000, "price_source": "КП перевозчика № 5", "source": "1 рейс"}]},
            ]}]}
    return tmp_path, base


def _run(tmp_path, data):
    p = tmp_path / "smeta.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return compute(load_smeta(p)), p


def test_resource_breakdown(env):
    tmp, base = env
    (calc, _), _ = _run(tmp, base)
    bt = calc["by_type"]
    assert bt["материал"] == Decimal("3501750.00")   # 101,5 × 34 500
    assert bt["труд"] == Decimal("210000.00")         # 60 × 3 500
    assert bt["машины"] == Decimal("475000.00")       # 12,5 × 38 000
    assert bt["транспорт"] == Decimal("100000.00")
    assert calc["direct"] == Decimal("4286750.00")
    assert calc["overhead"] == Decimal("428675.00")


def test_overhead_and_profit_from_fot(env):
    tmp, base = env
    base["rates"].update(overhead_base="fot", overhead_pct=50, profit_base="fot", profit_pct=20)
    (calc, _), _ = _run(tmp, base)
    assert calc["overhead"] == Decimal("105000.00") and calc["profit"] == Decimal("42000.00")


@pytest.mark.parametrize("mut, text", [
    (lambda d: d["sections"][0]["positions"][0].__setitem__("card", "НЕТ"), "не найдена"),
    (lambda d: d["sections"][0]["positions"][0].pop("qty_source"), "qty_source"),
    (lambda d: d["sections"][0]["positions"][1]["resources"][0].pop("price_source"), "price_source"),
    (lambda d: d["sections"][0]["positions"][1]["resources"][0].pop("source"), "источника нормы"),
    (lambda d: d["sections"][0]["positions"][1]["resources"][0].__setitem__("type", "прочее"), "тип"),
    (lambda d: d["sections"][0]["positions"][0].__setitem__("unit", "м2"), "калькуляции"),
    (lambda d: d["rates"].pop("vat_source"), "vat_source"),
    (lambda d: d["rates"].__setitem__("overhead_base", "фот"), "overhead_base"),
    (lambda d: d["sections"][0]["positions"][0].__setitem__("vor_id", "x"), "vor"),
])
def test_validation(env, mut, text):
    tmp, base = env
    d = copy.deepcopy(base)
    mut(d)
    with pytest.raises(SmetaError) as e:
        _run(tmp, d)
    assert any(text in x for x in e.value.errors), e.value.errors


def test_price_kind_must_match_type(env):
    tmp, base = env
    base["sections"][0]["positions"][1]["resources"][0] = {
        "type": "транспорт", "code": "ТР-1", "norm": 1, "source": "1"}
    base["sections"][0]["positions"][1]["unit"] = "рейс"
    with pytest.raises(SmetaError) as e:
        _run(tmp, base)
    assert any("вид «работа»" in x for x in e.value.errors)


def test_build_and_check(env):
    tmp, base = env
    _, p = _run(tmp, base)
    out = tmp / "out"
    r = subprocess.run([sys.executable, SM / "build_smeta.py", p, "--out", out],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    r = subprocess.run([sys.executable, SM / "check_smeta.py", p, out / "SMETA_С-1.xlsx",
                        out / "SMETA_С-1.pdf"], capture_output=True, text=True)
    assert r.returncode == 0 and "пройдена" in r.stdout, r.stdout + r.stderr


def _contractor_xlsx(path, conc_amount=5170500, total=13_000_000):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["№", "Наименование работ", "Ед. изм.", "Кол-во", "Цена, тг", "Стоимость, тг"])
    ws.append(["", "Бетонные работы"])
    ws.append([1, "Бетонирование фундамента", "м3", 110, 47000, conc_amount])
    ws.append([2, "Непредвиденные работы", "компл", 1, 450000, 450000])
    ws.append(["", "Итого без НДС", "", "", "", 5620000])
    ws.append(["", "НДС 16%", "", "", "", 899200])
    ws.append(["", "Всего с НДС", "", "", "", total])
    wb.save(path)


def test_import_and_compare(env):
    tmp, base = env
    base["sections"][0]["positions"].pop(1)
    _, p = _run(tmp, base)
    _contractor_xlsx(tmp / "c.xlsx")
    r = subprocess.run([sys.executable, SM / "import_estimate.py", tmp / "c.xlsx", "--out",
                        tmp / "c.json"], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    c = json.loads((tmp / "c.json").read_text(encoding="utf-8"))
    assert len(c["positions"]) == 2 and c["positions"][0]["section"] == "Бетонные работы"
    assert c["source_file"] == "c.xlsx"                  # без пути с компьютера пользователя
    assert c["totals"] == {"subtotal": "5620000", "vat": "899200", "total": "13000000"}
    r = subprocess.run([sys.executable, SM / "compare.py", "--ours", p, "--contractor",
                        tmp / "c.json", "--out", tmp / "cmp"], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    out = r.stdout
    assert "разница 500,00" in out                       # ошибка в строке
    assert "Всего 13 000 000,00 не равно" in out          # итог не сходится
    assert "Признак подгонки" in out
    assert "лишних у подрядчика: 1" in out
    from openpyxl import load_workbook
    wb = load_workbook(tmp / "cmp" / "COMPARE_С-1.xlsx")
    ws = wb["Сравнение"]
    notes = " ".join(str(ws.cell(r, 16).value or "") for r in range(1, ws.max_row + 1))
    assert "объём +10.0 %" in notes and "нет у нас" in notes


def test_compare_all_in_total_matches_subtotal(env):
    from compare import load_ours
    tmp, base = env
    _, p = _run(tmp, base)
    meta, rows = load_ours(p, "all_in")
    assert sum(r["amount"] for r in rows) == meta["subtotal"]
    meta, rows = load_ours(p, "direct")
    assert sum(r["amount"] for r in rows) == meta["direct"]
