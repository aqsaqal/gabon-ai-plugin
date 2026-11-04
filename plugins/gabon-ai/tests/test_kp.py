import json
import shutil
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

import pytest

from conftest import ROOT
from kp_core import InputError, amount_in_words, compute, load_input, validate

SCRIPTS = ROOT / "skills" / "kp" / "scripts"


def _write_kp(tmp_path, prices_path, items, **extra):
    data = {
        "number": "T-1", "date": "2026-11-03", "company": {"name": "ТОО Тест"},
        "client": {"name": "Заказчик"}, "object": {"name": "Объект"},
        "price_list": prices_path.name,
        "rates": {"overhead_pct": 0, "profit_pct": 0, "vat_pct": 16, "vat_source": "тест",
                  "amount_decimals": 0},
        "sections": [{"title": "Раздел", "items": items}], **extra,
    }
    p = tmp_path / "kp.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return p


def test_kp_prices_from_list(tmp_path, make_prices, good_rows):
    pp = make_prices(good_rows)
    kp = _write_kp(tmp_path, pp, [{"code": "ЗР-001", "qty": 100}, {"code": "БЛ-001", "qty": 2}])
    data = load_input(kp)
    warnings = validate(data)
    calc = compute(data)
    assert calc["direct"] == Decimal("185000") + Decimal("15660")
    assert calc["vat"] == Decimal("32106")          # 200660 × 16 % = 32105,6 -> 32106
    assert any("Актуализируйте" in w for w in warnings)


def test_kp_unknown_code_blocks_build(tmp_path, make_prices, good_rows):
    kp = _write_kp(tmp_path, make_prices(good_rows), [{"code": "НЕТ", "qty": 1}])
    with pytest.raises(InputError) as e:
        validate(load_input(kp))
    assert any("не найден" in x for x in e.value.errors)


def test_code_without_price_list_is_error(tmp_path):
    data = {"number": "1", "date": "2026-11-03", "company": {"name": "Т"}, "object": {"name": "О"},
            "rates": {"overhead_pct": 0, "profit_pct": 0, "vat_pct": 16, "vat_source": "т"},
            "sections": [{"title": "Р", "items": [{"code": "X", "qty": 1}]}]}
    p = tmp_path / "kp.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(InputError) as e:
        validate(load_input(p))
    assert any("price_list" in x for x in e.value.errors)


def test_rounding_to_whole_tenge_is_half_up():
    # 37,5 -> 38 и 200,5 -> 201 (как ОКРУГЛ в Excel), итог не подгоняется под круглую сумму
    data = load_input(Path(__file__).parent / "fixtures" / "kp_round_total.json")
    validate(data)
    calc = compute(data)
    assert calc["sections"][0]["items"][:2] == [Decimal("38"), Decimal("201")]
    assert calc["subtotal"] == Decimal("1000239")
    assert calc["vat"] == Decimal("160038")               # 160 038,24 -> 160 038
    assert calc["total"] == Decimal("1160277")


def test_amount_in_words():
    assert amount_in_words(Decimal("165000000")) == "сто шестьдесят пять миллионов тенге 00 тиын"
    assert amount_in_words(Decimal("1001.05")) == "одна тысяча один тенге 05 тиын"
    assert amount_in_words(Decimal("22")) == "двадцать два тенге 00 тиын"


@pytest.mark.skipif(shutil.which("uv") is None, reason="нужен uv")
def test_build_and_check_end_to_end(tmp_path, make_prices, good_rows):
    kp = _write_kp(tmp_path, make_prices(good_rows),
                   [{"code": "ЗР-001", "qty": 420}, {"code": "МН-001", "qty": 86}])
    out = tmp_path / "out"
    r = subprocess.run([sys.executable, SCRIPTS / "build_kp.py", kp, "--out", out],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    r = subprocess.run([sys.executable, SCRIPTS / "check_kp.py", kp, out / "KP_T-1.xlsx",
                        out / "KP_T-1.pdf"], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    assert "пройдена" in r.stdout
