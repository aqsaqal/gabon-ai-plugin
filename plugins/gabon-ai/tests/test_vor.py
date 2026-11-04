import copy
import json
import subprocess
import sys
from decimal import Decimal

import pytest

from conftest import ROOT
from vor_core import VorError, compute, display_formula, load_vor, normalize_formula, row_numbers

VOR = ROOT / "skills" / "vor" / "scripts"
KP = ROOT / "skills" / "kp" / "scripts"
BASE = {
    "number": "В-1", "date": "2026-11-03", "object": {"name": "Объект"},
    "company": {"name": "ТОО Тест"},
    "documents": [{"id": "КЖ-3", "name": "Лист КЖ-3"}],
    "sections": [{"title": "Фундамент", "items": [
        {"id": "бетон", "name": "Бетон", "unit": "м³", "formula": "144 × 0,6 х 1,2", "ref": "КЖ-3", "code": "МН-001"},
        {"id": "выемка", "name": "Выемка", "unit": "м3", "formula": "144*1,6*1,5", "ref": "КЖ-3", "code": "ЗР-001"},
        {"id": "засыпка", "name": "Засыпка", "unit": "м3", "formula": "@выемка-@бетон", "ref": "КЖ-3"},
        {"id": "арм", "name": "Арматура", "unit": "т", "status": "нет данных", "missing": "нет КЖ"},
        {"id": "мастика", "name": "Мастика", "unit": "кг", "formula": "@бетон*2",
         "status": "допущение", "assumption": "расход условный"},
    ]}],
}


def _calc(mut=None):
    d = copy.deepcopy(BASE)
    if mut:
        mut(d)
    return compute(d), d


def _items(d):
    return d["sections"][0]["items"]


def test_formulas_units_and_refs():
    (values, warnings), d = _calc()
    assert values["бетон"] == Decimal("103.68")
    assert values["засыпка"] == Decimal("345.60") - Decimal("103.68")
    assert values["мастика"] == Decimal("207.36")
    assert "арм" not in values
    assert _items(d)[0]["unit"] == "м3"
    assert any("допущениями: 1" in w for w in warnings)


def test_normalize_formula():
    assert normalize_formula("= 24 × 0,6 х 1,2 · 2") == "24*0.6*1.2*2"
    assert normalize_formula("@a-@b") == "@a-@b"


def test_display_formula_uses_row_numbers():
    (_, _), d = _calc()
    assert display_formula("@выемка-@бетон", row_numbers(d)) == "п.1.2-п.1.1"


@pytest.mark.parametrize("mut, text", [
    (lambda d: _items(d)[0].pop("ref"), "требует ref"),
    (lambda d: _items(d)[4].pop("assumption"), "assumption"),
    (lambda d: _items(d)[3].__setitem__("qty", 5), "нет данных"),
    (lambda d: _items(d)[3].pop("missing"), "missing"),
    (lambda d: _items(d)[0].__setitem__("qty", 100), "не равно результату"),
    (lambda d: _items(d)[0].__setitem__("ref", "АР-1"), "не описан"),
    (lambda d: _items(d)[1].__setitem__("id", "бетон"), "повторяется"),
    (lambda d: _items(d)[2].__setitem__("formula", "@арм*2"), "без данных"),
    (lambda d: _items(d)[2].__setitem__("formula", "@нет*2"), "несуществующую"),
    (lambda d: _items(d)[2].__setitem__("formula", "@засыпка+1"), "циклическая"),
    (lambda d: _items(d)[2].__setitem__("formula", "10/0"), "деление на ноль"),
    (lambda d: _items(d)[2].__setitem__("formula", "__import__('os')"), "допустимы только"),
    (lambda d: _items(d)[2].__setitem__("formula", "(1).__class__"), "допустимы только"),
    (lambda d: _items(d)[2].__setitem__("formula", "@бетон-@выемка"), "больше нуля"),
    (lambda d: _items(d)[1].pop("formula"), "нет ни formula"),
])
def test_validation_errors(mut, text):
    with pytest.raises(VorError) as e:
        _calc(mut)
    assert any(text in x for x in e.value.errors), e.value.errors


def test_qty_matching_formula_is_ok():
    (values, _), _ = _calc(lambda d: _items(d)[0].__setitem__("qty", "103,68"))
    assert values["бетон"] == Decimal("103.68")


def _write(tmp_path, data):
    p = tmp_path / "vor.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return p


def test_build_check_and_convert_to_kp(tmp_path, make_prices):
    from datetime import date
    vor = _write(tmp_path, BASE)
    out = tmp_path / "out"
    r = subprocess.run([sys.executable, VOR / "build_vor.py", vor, "--out", out],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    r = subprocess.run([sys.executable, VOR / "check_vor.py", vor, out / "VOR_В-1.xlsx",
                        out / "VOR_В-1.pdf"], capture_output=True, text=True)
    assert r.returncode == 0 and "пройдена" in r.stdout, r.stdout + r.stderr

    prices = make_prices([
        ("МН-001", "Бетон", "м3", 34500, "материал", "src", date(2026, 10, 1), ""),
        ("ЗР-001", "Выемка", "м3", 1850, "механизм", "src", date(2026, 10, 1), "")])
    kp = tmp_path / "kp.json"
    r = subprocess.run([sys.executable, VOR / "vor_to_kp.py", vor, "--out", kp, "--number", "К-1",
                        "--price-list", prices], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    k = json.loads(kp.read_text(encoding="utf-8"))
    names = [it["name"] for s in k["sections"] for it in s["items"]]
    assert "Арматура" not in names and len(names) == 4
    assert any("Арматура" in x for x in k["exclusions"])
    assert any("Мастика" in x for x in k["assumptions"])
    assert k["sections"][0]["items"][0] == {"name": "Бетон", "unit": "м3", "qty": "103.68",
                                            "vor_ref": "ВОР № В-1, строка 1.1", "code": "МН-001"}
    r = subprocess.run([sys.executable, VOR / "vor_to_kp.py", vor, "--out", kp, "--number", "К-1"],
                       capture_output=True, text=True)
    assert r.returncode == 1 and "не перезаписываю" in r.stdout


def test_check_detects_tampered_quantity(tmp_path):
    from openpyxl import load_workbook
    vor = _write(tmp_path, BASE)
    out = tmp_path / "out"
    subprocess.run([sys.executable, VOR / "build_vor.py", vor, "--out", out], check=True,
                   capture_output=True)
    x = out / "VOR_В-1.xlsx"
    wb = load_workbook(x)
    ws = wb["ВОР"]
    row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 10).value == "бетон")
    ws.cell(row, 4).value = 100
    wb.save(x)
    r = subprocess.run([sys.executable, VOR / "check_vor.py", vor, x, out / "VOR_В-1.pdf"],
                       capture_output=True, text=True)
    assert r.returncode == 1
    assert "записано числом" in r.stdout and "ожидалось 103.68" in r.stdout


def test_extract_docs_reads_docx_xlsx_pdf(tmp_path):
    pytest.importorskip("docx")
    import docx
    from openpyxl import Workbook
    from reportlab.pdfgen import canvas
    src = tmp_path / "src"
    src.mkdir()
    d = docx.Document()
    d.add_paragraph("Длина ленты 144 м")
    t = d.add_table(rows=1, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "Бетон", "103,68"
    d.save(src / "tz.docx")
    wb = Workbook()
    wb.active.append(["Ширина", 0.6])
    wb.save(src / "spec.xlsx")
    c = canvas.Canvas(str(src / "scan.pdf"))
    c.rect(10, 10, 100, 100)
    c.save()
    out = tmp_path / "txt"
    r = subprocess.run([sys.executable, VOR / "extract_docs.py", src, "--out", out],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    assert "144 м" in (out / "tz.docx.txt").read_text(encoding="utf-8")
    assert "Бетон | 103,68" in (out / "tz.docx.txt").read_text(encoding="utf-8")
    assert "Ширина | 0.6" in (out / "spec.xlsx.txt").read_text(encoding="utf-8")
    assert "без текста" in r.stdout


def test_load_vor_bad_json(tmp_path):
    p = tmp_path / "vor.json"
    p.write_text("{ плохой json", encoding="utf-8")
    with pytest.raises(VorError):
        load_vor(p)
