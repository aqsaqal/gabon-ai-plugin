import copy
import json
import subprocess
import sys
from datetime import date

import pytest

from conftest import ROOT
from grafik_core import GrafikError, compute, load_grafik, workday, workdays_between

GR = ROOT / "skills" / "grafik" / "scripts"
NOHOL = frozenset()
BASE = {
    "number": "Г-1", "date": "2026-11-03", "object": {"name": "О"}, "start": "2026-11-14",
    "calendar": {"holidays": ["2026-12-16"], "holidays_source": "закон"},
    "tasks": [
        {"id": "а", "name": "Земля", "qty": 500, "qty_source": "ВОР", "unit": "м3",
         "productivity": {"per_shift": 100, "unit": "м3", "source": "ППР"}, "crews": 2},
        {"id": "б", "name": "Опалубка", "labor": {"workers": 6, "hours": 311.04, "source": "смета"},
         "after": ["а"]},
        {"id": "в", "name": "Арматура", "duration": 4, "assumption": "оценка",
         "after": [{"id": "б", "type": "SS", "lag": 2}]},
        {"id": "г", "name": "Выдержка", "duration": 7, "calendar_days": True,
         "duration_source": "ППР", "after": ["б", "в"]},
        {"id": "д", "name": "Готово", "milestone": True, "after": ["г"]},
    ]}


def _calc(mut=None):
    d = copy.deepcopy(BASE)
    if mut:
        mut(d)
    res, w = compute(d)
    return res, {t["_id"]: t for t in d["tasks"]}, w


def test_workday_matches_excel_semantics():
    fri = date(2026, 11, 20)
    assert workday(fri, 1, "0000011", NOHOL) == date(2026, 11, 23)      # пт + 1 = пн
    assert workday(fri, 0, "0000011", NOHOL) == fri
    assert workday(date(2026, 11, 21), 0, "0000011", NOHOL) == date(2026, 11, 21)  # 0 дней: та же дата
    assert workday(date(2026, 12, 15), 1, "0000011", frozenset({date(2026, 12, 16)})) == date(2026, 12, 17)
    assert workday(date(2026, 11, 23), -1, "0000011", NOHOL) == fri
    assert workday(fri, 1, "0000001", NOHOL) == date(2026, 11, 21)      # шестидневка
    assert workday(fri, 2, "0000000", NOHOL) == date(2026, 11, 22)      # календарные дни
    assert workdays_between(date(2026, 11, 20), date(2026, 11, 27), "0000011", NOHOL) == 5


def test_schedule_dates_float_and_critical():
    res, t, _ = _calc()
    assert t["а"]["_es"] == date(2026, 11, 16)                 # старт в субботу -> понедельник
    assert t["а"]["_dur"] == 3                                  # 500 / (100 × 2) = 2,5 -> 3
    assert t["б"]["_dur"] == 7                                  # 311,04 / (6 × 8) = 6,48 -> 7
    assert t["б"]["_es"] == date(2026, 11, 19)
    assert t["б"]["_ef"] == date(2026, 11, 27)
    assert t["в"]["_es"] == date(2026, 11, 23)                 # НН +2 раб. дня
    assert t["г"]["_es"] == date(2026, 11, 28)                 # календарные: суббота
    assert t["г"]["_ef"] == date(2026, 12, 4)
    assert t["д"]["_es"] == t["д"]["_ef"] == date(2026, 12, 4)
    assert res["finish"] == date(2026, 12, 4)
    assert [k for k in "абвгд" if t[k]["_critical"]] == ["а", "б", "г", "д"]
    assert t["в"]["_float"] == 1


def test_exact_division_is_not_rounded_up():
    _, t, _ = _calc(lambda d: d["tasks"][1]["labor"].__setitem__("hours", 336))  # 336 / 48 = 7
    assert t["б"]["_dur"] == 7


def test_holiday_shifts_dates():
    res, t, _ = _calc(lambda d: d.__setitem__("start", "2026-12-14"))
    assert t["а"]["_ef"] == date(2026, 12, 17)                 # 14, 15, (16 праздник), 17


@pytest.mark.parametrize("mut, text", [
    (lambda d: d["tasks"][0]["productivity"].pop("source"), "productivity.source"),
    (lambda d: d["tasks"][2].pop("assumption"), "duration_source"),
    (lambda d: d["tasks"][1].__setitem__("duration", 3), "ровно один способ"),
    (lambda d: d["tasks"][1].__setitem__("after", ["нет"]), "не найден"),
    (lambda d: d["tasks"][0].__setitem__("after", ["д"]), "Циклическая"),
    (lambda d: d["calendar"].pop("holidays_source"), "holidays_source"),
    (lambda d: d["calendar"].__setitem__("weekend", "1111111"), "weekend"),
    (lambda d: d["tasks"][0]["productivity"].__setitem__("unit", "м2"), "выработка"),
    (lambda d: d["tasks"][2]["after"][0].__setitem__("type", "FF"), "FS или SS"),
    (lambda d: d["tasks"][0].__setitem__("vor_id", "x"), "vor"),
])
def test_validation(mut, text):
    with pytest.raises(GrafikError) as e:
        _calc(mut)
    assert any(text in x for x in e.value.errors), e.value.errors


def _build(tmp_path, data):
    p = tmp_path / "grafik.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "out"
    r = subprocess.run([sys.executable, GR / "build_grafik.py", p, "--out", out],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    return p, out / "GRAFIK_Г-1.xlsx", out / "GRAFIK_Г-1.pdf"


def test_build_check_and_live_excel(tmp_path):
    p, xlsx, pdf = _build(tmp_path, BASE)
    r = subprocess.run([sys.executable, GR / "check_grafik.py", p, xlsx, pdf],
                       capture_output=True, text=True)
    assert r.returncode == 0 and "пройдена" in r.stdout, r.stdout + r.stderr
    from openpyxl import load_workbook
    from check_grafik import EPOCH, Evaluator
    from datetime import timedelta
    wb = load_workbook(xlsx)
    assert "_xlfn.WORKDAY.INTL" in wb["График"]["J8"].value
    wb["График"]["D7"] = 900                                    # объём земли вырос: 900 / 200 = 4,5 -> 5 дн.
    ev = Evaluator(wb)
    fin = EPOCH + timedelta(days=int(ev.value("График", "K11")))
    assert fin == date(2026, 12, 8)                             # всё сдвинулось на 2 рабочих дня


def test_check_detects_tampering(tmp_path):
    p, xlsx, pdf = _build(tmp_path, BASE)
    from openpyxl import load_workbook
    wb = load_workbook(xlsx)
    wb["График"]["I8"] = 5
    wb.save(xlsx)
    r = subprocess.run([sys.executable, GR / "check_grafik.py", p, xlsx, pdf],
                       capture_output=True, text=True)
    assert r.returncode == 1 and "должна считаться формулой" in r.stdout


def test_many_tasks_paginate(tmp_path):
    d = copy.deepcopy(BASE)
    d["tasks"] = [{"id": f"t{i}", "name": f"Работа {i} с длинным наименованием для проверки обрезки текста в колонке",
                   "duration": 2, "duration_source": "тест", "after": [f"t{i - 1}"] if i else []}
                  for i in range(45)]
    p, xlsx, pdf = _build(tmp_path, d)
    from pypdf import PdfReader
    assert len(PdfReader(str(pdf)).pages) >= 2
    r = subprocess.run([sys.executable, GR / "check_grafik.py", p, xlsx, pdf],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout


def test_from_vor_and_smeta():
    data = load_grafik(ROOT.parents[1] / "projects" / "test-fundament" / "input" / "grafik.json") \
        if (ROOT.parents[1] / "projects" / "test-fundament" / "input" / "grafik.json").exists() else None
    if data is None:
        pytest.skip("нет тестового проекта")
    res, _ = compute(data)
    assert res["finish"] == date(2026, 12, 24) and res["workdays"] == 28
