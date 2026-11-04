import copy
import json
import subprocess
import sys

import pytest

from conftest import ROOT

KP = ROOT / "skills" / "kp" / "scripts"
BASE = {"number": "W-1", "date": "2026-11-03", "valid_days": 14,
        "company": {"name": "ТОО «Тест»", "bin": "123", "director": "Иванов И. И."},
        "client": {"name": "ТОО «Заказчик»"}, "object": {"name": "Склад"},
        "rates": {"overhead_pct": 10, "profit_pct": 8, "vat_pct": 16, "vat_source": "НК РК"},
        "sections": [{"title": "Земляные работы", "items": [
            {"name": "Выемка грунта", "unit": "м3", "qty": 420, "price": 1850, "kind": "механизм", "source": "прайс"},
            {"name": "Засыпка", "unit": "м3", "qty": 260, "price": 2100, "kind": "работа", "source": "прайс"}]}],
        "terms": ["Оплата по актам"], "assumptions": ["Грунт II группы"], "exclusions": ["Геодезия"]}


def _build(tmp_path, data, *extra):
    p = tmp_path / "kp.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "out"
    r = subprocess.run([sys.executable, KP / "build_kp.py", p, "--out", out, *extra],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    return p, out / "KP_W-1"


def _text(path):
    import docx
    d = docx.Document(str(path))
    return " ".join(p.text for p in d.paragraphs) + " " + " ".join(
        c.text for t in d.tables for row in t.rows for c in row.cells)


def _check(p, b, docx=True):
    args = [p, b.with_suffix(".xlsx"), b.with_suffix(".pdf")] + ([b.with_suffix(".docx")] if docx else [])
    return subprocess.run([sys.executable, KP / "check_kp.py", *args], capture_output=True, text=True)


def test_word_has_same_content_and_passes_check(tmp_path):
    p, b = _build(tmp_path, BASE)
    t = _text(b.with_suffix(".docx"))
    for x in ("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ № W-1", "ТОО «Заказчик»", "777 000,00", "546 000,00",
              "Накладные расходы (10 %)", "ВСЕГО с НДС, тенге", "Сумма прописью", "Оплата по актам",
              "Грунт II группы", "Геодезия", "Иванов И. И.", "БИН: 123"):
        assert x in t, x
    r = _check(p, b)
    assert r.returncode == 0 and "Word: номер КП" in r.stdout, r.stdout


def test_word_table_has_fixed_full_width(tmp_path):
    import docx
    from docx.oxml.ns import qn
    _, b = _build(tmp_path, BASE)
    tbl = docx.Document(str(b.with_suffix(".docx"))).tables[2]._tbl
    assert tbl.tblPr.find(qn("w:tblLayout")).get(qn("w:type")) == "fixed"
    total = sum(int(g.get(qn("w:w"))) for g in tbl.tblGrid.findall(qn("w:gridCol")))
    assert abs(total / 56.7 - 180) < 1                      # 180 мм рабочей ширины A4


def test_vat_exempt_wording_in_word(tmp_path):
    d = copy.deepcopy(BASE)
    d["rates"].update(vat_pct=0, vat_source="Не плательщик НДС")
    p, b = _build(tmp_path, d)
    t = _text(b.with_suffix(".docx"))
    assert "НДС не облагается" in t and "с НДС" not in t
    assert _check(p, b).returncode == 0


def test_sample_mark_and_no_docx_flag(tmp_path):
    d = copy.deepcopy(BASE)
    d["sample"] = True
    _, b = _build(tmp_path, d)
    assert "ОБРАЗЕЦ" in _text(b.with_suffix(".docx"))
    (tmp_path / "x").mkdir()
    _, b2 = _build(tmp_path / "x", BASE, "--no-docx")
    assert not b2.with_suffix(".docx").exists()


def test_logo_right_position(tmp_path):
    from PIL import Image
    Image.new("RGB", (200, 80), (11, 93, 59)).save(tmp_path / "logo.png")
    (tmp_path / "brand.json").write_text(json.dumps({"logo": "logo.png", "logo_position": "right"}),
                                         encoding="utf-8")
    d = copy.deepcopy(BASE)
    d["brand"] = "brand.json"
    p, b = _build(tmp_path, d)
    import docx
    doc = docx.Document(str(b.with_suffix(".docx")))
    assert len(doc.inline_shapes) == 1
    head = doc.tables[0].rows[0].cells
    assert head[1]._tc.xml.count("<pic:pic") == 1 and "ТОО «Тест»" in head[0].text
    assert _check(p, b).returncode == 0


def test_check_detects_edited_word(tmp_path):
    import docx
    p, b = _build(tmp_path, BASE)
    path = b.with_suffix(".docx")
    doc = docx.Document(str(path))
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                if c.text == "777 000,00":
                    c.text = "877 000,00"
    doc.save(str(path))
    r = _check(p, b)
    assert r.returncode == 1 and "В Word не найдены" in r.stdout
