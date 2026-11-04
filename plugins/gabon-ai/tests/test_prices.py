from datetime import date
from decimal import Decimal

from prices_core import apply_prices, load_price_list, normalize_unit


def test_load_good_list(make_prices, good_rows):
    pl = load_price_list(make_prices(good_rows))
    assert pl.errors == []
    assert len(pl.records) == 3
    assert pl.get("ЗР-001").unit == "м3"          # м³ нормализовано
    assert pl.get("МН-001").price == Decimal("34500.50")  # пробелы и запятая
    assert pl.get("МН-001").date == date(2026, 10, 15)    # дата строкой ДД.ММ.ГГГГ


def test_units_normalized():
    assert normalize_unit("м²") == "м2"
    assert normalize_unit(" М3 ") == "м3"
    assert normalize_unit("м.п") == "п.м"


def test_duplicate_code_is_error(make_prices, good_rows):
    pl = load_price_list(make_prices(good_rows + [good_rows[0]]))
    assert any("повторяется" in e for e in pl.errors)


def test_missing_source_and_date_are_errors(make_prices):
    pl = load_price_list(make_prices([("X-1", "Работа", "м2", 100, "работа", "", None, "")]))
    assert any("нет источника" in e for e in pl.errors)
    assert any("нет даты" in e for e in pl.errors)
    assert "X-1" not in pl.records


def test_bad_kind_and_price(make_prices):
    pl = load_price_list(make_prices([("X-1", "Работа", "м2", "дорого", "услуги", "src", date(2026, 1, 1), "")]))
    assert any("цена не число" in e for e in pl.errors)
    assert any("вид затрат" in e for e in pl.errors)


def test_price_header_must_say_without_vat(make_prices, good_rows):
    hdr = ["Код", "Наименование", "Ед. изм.", "Цена, тг", "Вид затрат", "Источник", "Дата цены"]
    pl = load_price_list(make_prices([r[:7] for r in good_rows], header=hdr))
    assert any("без НДС" in e for e in pl.errors)


def test_find_by_code_and_words(make_prices, good_rows):
    pl = load_price_list(make_prices(good_rows))
    assert pl.find("мн-001")[0][1].code == "МН-001"
    assert pl.find("бетон")[0][1].code == "МН-001"
    assert pl.find("кирпич") == []


def _kp(items):
    return {"date": "2026-11-03", "sections": [{"title": "Р", "items": items}]}


def test_apply_fills_price_source_unit_kind(make_prices, good_rows):
    pl = load_price_list(make_prices(good_rows))
    data = _kp([{"code": "ЗР-001", "qty": 10}])
    errors, warnings = apply_prices(data, pl, date(2026, 11, 3))
    it = data["sections"][0]["items"][0]
    assert errors == [] and warnings == []
    assert it["price"] == "1850" and it["unit"] == "м3" and it["kind"] == "механизм"
    assert it["name"] == "Разработка грунта экскаватором"
    assert "код ЗР-001" in it["source"] and "01.10.2026" in it["source"]


def test_apply_stale_price_warns(make_prices, good_rows):
    pl = load_price_list(make_prices(good_rows))
    errors, warnings = apply_prices(_kp([{"code": "БЛ-001", "qty": 1}]), pl, date(2026, 11, 3))
    assert errors == []
    assert any("186 дн." in w for w in warnings)


def test_apply_unknown_code_is_error(make_prices, good_rows):
    pl = load_price_list(make_prices(good_rows))
    errors, _ = apply_prices(_kp([{"code": "НЕТ-1", "qty": 1}]), pl, date(2026, 11, 3))
    assert any("не найден" in e for e in errors)


def test_apply_unit_mismatch_is_error(make_prices, good_rows):
    pl = load_price_list(make_prices(good_rows))
    errors, _ = apply_prices(_kp([{"code": "МН-001", "qty": 1, "unit": "т"}]), pl, date(2026, 11, 3))
    assert any("единица" in e for e in errors)


def test_apply_manual_price_needs_reason(make_prices, good_rows):
    pl = load_price_list(make_prices(good_rows))
    errors, _ = apply_prices(_kp([{"code": "ЗР-001", "qty": 1, "price": 2000}]), pl, date(2026, 11, 3))
    assert any("отличается от прайса" in e for e in errors)
    data = _kp([{"code": "ЗР-001", "qty": 1, "price": 2000, "price_override_reason": "скидка по письму № 7"}])
    errors, warnings = apply_prices(data, pl, date(2026, 11, 3))
    assert errors == [] and any("изменена вручную" in w for w in warnings)
    assert data["sections"][0]["items"][0]["price"] == 2000
