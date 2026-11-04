import json
import subprocess
import sys

import pytest

from brand_core import BrandError, apply_defaults, brand_for, load_brand
from conftest import ROOT

SCRIPTS = ROOT / "skills" / "kp" / "scripts"


def _brand(tmp_path, **kw):
    p = tmp_path / "brand.json"
    p.write_text(json.dumps(kw, ensure_ascii=False), encoding="utf-8")
    return p


def _png(path):
    from PIL import Image
    Image.new("RGB", (200, 80), (11, 93, 59)).save(path)
    return path


def _letterhead(path):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(str(path), pagesize=A4)
    c.rect(0, A4[1] - 80, A4[0], 80, fill=1)
    c.save()
    return path


def test_empty_brand_uses_defaults(tmp_path):
    b, w = load_brand(_brand(tmp_path))
    assert b["colors"]["primary"] == "#1F3864"
    assert b["show_requisites"] is True and w == []


def test_bad_color_and_missing_logo(tmp_path):
    with pytest.raises(BrandError) as e:
        load_brand(_brand(tmp_path, colors={"primary": "зелёный"}, logo="нет.png"))
    msg = " ".join(e.value.errors)
    assert "#RRGGBB" in msg and "не найден" in msg


def test_unknown_color_is_warning(tmp_path):
    _, w = load_brand(_brand(tmp_path, colors={"фон": "#FFFFFF"}))
    assert any("Неизвестный цвет" in x for x in w)


def test_fonts_need_both_files(tmp_path):
    with pytest.raises(BrandError) as e:
        load_brand(_brand(tmp_path, fonts={"regular": "a.ttf"}))
    assert any("оба файла" in x for x in e.value.errors)


def test_vat_default_needs_source(tmp_path):
    with pytest.raises(BrandError):
        load_brand(_brand(tmp_path, defaults={"rates": {"vat_pct": 16}}))


def test_letterhead_hides_requisites_by_default(tmp_path):
    _letterhead(tmp_path / "blank.pdf")
    b, _ = load_brand(_brand(tmp_path, letterhead_pdf="blank.pdf"))
    assert b["show_requisites"] is False


def test_defaults_do_not_override_kp(tmp_path):
    b, _ = load_brand(_brand(tmp_path, defaults={
        "terms": ["из шаблона"], "valid_days": 14, "intro": "вступление",
        "rates": {"vat_pct": 16, "vat_source": "шаблон", "overhead_pct": 5}}))
    data = {"terms": ["из КП"], "rates": {"vat_pct": 12, "vat_source": "КП"}}
    used = apply_defaults(data, b)
    assert data["terms"] == ["из КП"]                 # КП главнее
    assert data["rates"]["vat_pct"] == 12
    assert data["valid_days"] == 14 and data["intro"] == "вступление"
    assert data["rates"]["overhead_pct"] == 5
    assert set(used) == {"valid_days", "intro", "rates.overhead_pct"}


def test_brand_for_without_brand_keeps_company_logo(tmp_path):
    _png(tmp_path / "logo.png")
    b = brand_for({"company": {"logo": "logo.png"}, "_company_dir": str(tmp_path)})
    assert b["logo"].endswith("logo.png")


def _kp(tmp_path, **extra):
    data = {"number": "B-1", "date": "2026-11-03", "company": {"name": "ТОО Бренд"},
            "client": {"name": "З"}, "object": {"name": "О"}, "brand": "brand.json",
            "rates": {"overhead_pct": 0, "profit_pct": 0},
            "sections": [{"title": "Р", "items": [
                {"name": "Работа", "unit": "м2", "qty": 10, "price": 1000, "kind": "работа", "source": "тест"}]}],
            **extra}
    p = tmp_path / "kp.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return p


def _build_and_check(kp, out):
    r = subprocess.run([sys.executable, SCRIPTS / "build_kp.py", kp, "--out", out],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    r = subprocess.run([sys.executable, SCRIPTS / "check_kp.py", kp, out / "KP_B-1.xlsx",
                        out / "KP_B-1.pdf"], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    return out / "KP_B-1.pdf"


def test_build_with_logo_colors_and_vat_from_brand(tmp_path):
    _png(tmp_path / "logo.png")
    _brand(tmp_path, logo="logo.png", colors={"primary": "#0B5D3B"}, footer_text="Подвал",
           defaults={"rates": {"vat_pct": 16, "vat_source": "из шаблона"},
                     "terms": ["Условие из шаблона"]})
    pdf = _build_and_check(_kp(tmp_path), tmp_path / "out")
    from pypdf import PdfReader
    text = PdfReader(str(pdf)).pages[0].extract_text()
    assert "Условие из шаблона" in text and "Подвал" in text and "1 600,00" in text
    import zipfile
    assert any(n.startswith("xl/media/") for n in zipfile.ZipFile(tmp_path / "out" / "KP_B-1.xlsx").namelist())


def test_build_on_letterhead(tmp_path):
    _letterhead(tmp_path / "blank.pdf")
    _brand(tmp_path, letterhead_pdf="blank.pdf", margins_mm={"top": 40},
           defaults={"rates": {"vat_pct": 16, "vat_source": "т"}})
    pdf = _build_and_check(_kp(tmp_path), tmp_path / "out")
    from pypdf import PdfReader
    page = PdfReader(str(pdf)).pages[0]
    assert abs(float(page.mediabox.height) - 841.89) < 1
    assert not (tmp_path / "out" / "KP_B-1.content.pdf").exists()


def test_bad_brand_blocks_build(tmp_path):
    _brand(tmp_path, colors={"primary": "red"})
    r = subprocess.run([sys.executable, SCRIPTS / "build_kp.py", _kp(tmp_path), "--out",
                        tmp_path / "out"], capture_output=True, text=True)
    assert r.returncode == 2 and "Фирменный шаблон" in r.stdout
