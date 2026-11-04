"""Общая логика КП: загрузка исходных данных, проверка и расчёт.

Используется и сборщиком (build_kp.py), и проверяющим скриптом (check_kp.py),
чтобы расчёт велся по одним и тем же правилам.
"""
from __future__ import annotations

import json
import sys
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path

# Вид затрат -> подпись в документах
KINDS = {
    "работа": "Работы (оплата труда)",
    "материал": "Материалы",
    "механизм": "Машины и механизмы",
    "транспорт": "Транспортные расходы",
    "комплекс": "Работы с материалами (комплексная цена)",
    "прочее": "Прочие затраты",
}

KNOWN_UNITS = {
    "м", "п.м", "м2", "м3", "т", "кг", "л", "шт", "компл", "к-т", "маш.-ч", "чел.-ч",
    "рейс", "км", "т·км", "смена", "мес", "сутки", "услуга", "объект",
}

REQUIRED_TOP = ("number", "date", "company", "object", "sections", "rates")


MAX_NUMBER = Decimal("1e15")  # больше любой реальной сметы; защищает от опечаток

_UNIT_ALIASES = {"м.п": "п.м", "м/п": "п.м", "пм": "п.м", "мп": "п.м", "пог.м": "п.м",
                 "тн": "т", "тонна": "т", "шт.": "шт", "штук": "шт", "маш-ч": "маш.-ч",
                 "маш.ч": "маш.-ч", "мч": "маш.-ч", "чел-ч": "чел.-ч", "чел.ч": "чел.-ч",
                 "кв.м": "м2", "куб.м": "м3", "комплект": "компл"}


def normalize_unit(u) -> str:
    """«м²» -> «м2», «М3 » -> «м3», «м.п» -> «п.м»."""
    s = str(u or "").strip().lower().replace("²", "2").replace("³", "3").replace(" ", "")
    return _UNIT_ALIASES.get(s, s)


class InputError(Exception):
    def __init__(self, errors: list[str]):
        super().__init__("\n".join(errors))
        self.errors = errors


def D(value) -> Decimal:
    """Число из JSON -> Decimal (без потерь двоичной арифметики)."""
    if isinstance(value, bool) or not isinstance(value, (int, float, str, Decimal)):
        raise InvalidOperation(value)
    d = value if isinstance(value, Decimal) else \
        Decimal(str(value).replace("\u00a0", "").replace(" ", "").replace(",", "."))
    if not d.is_finite() or abs(d) >= MAX_NUMBER:
        raise InvalidOperation(value)
    return d


def money(x: Decimal, places: int = 2) -> Decimal:
    """Округление «половина вверх», как ROUND в Excel. places=2 — до тиын, 0 — до тенге."""
    return x.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def places_of(data: dict) -> int:
    return int(data.get("rates", {}).get("amount_decimals", 2))


def load_input(path: str | Path) -> dict:
    p = Path(path).resolve()
    data = json.loads(p.read_text(encoding="utf-8"))
    data["_input_dir"] = str(p.parent)
    comp = data.get("company")
    if isinstance(comp, str):  # путь к файлу реквизитов
        cp = (p.parent / comp).resolve()
        if not cp.exists():
            raise InputError([f"Не найден файл реквизитов компании: {cp}"])
        data["company"] = json.loads(cp.read_text(encoding="utf-8"))
        data["_company_dir"] = str(cp.parent)
    else:
        data["_company_dir"] = str(p.parent)

    data["_price_errors"], data["_price_warnings"] = [], []
    data["_brand"], data["_brand_used"] = None, []
    if data.get("brand"):
        # Фирменный шаблон компании (навык brand): стиль и тексты по умолчанию
        sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "brand" / "scripts"))
        from brand_core import BrandError, apply_defaults, load_brand
        try:
            brand, bw = load_brand((p.parent / data["brand"]).resolve())
        except BrandError as e:
            raise InputError([f"Фирменный шаблон: {x}" for x in e.errors])
        data["_brand"] = brand
        data["_brand_used"] = apply_defaults(data, brand)
        data["_price_warnings"] += [f"Фирменный шаблон: {w}" for w in bw]
    if data.get("price_list"):
        # Цены по кодам из прайс-листа компании (навык prices)
        sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "prices" / "scripts"))
        from prices_core import DEFAULT_MAX_AGE_DAYS, apply_prices, load_price_list
        pl = load_price_list((p.parent / data["price_list"]).resolve())
        try:
            on = date.fromisoformat(str(data.get("date")))
        except ValueError:
            on = date.today()
        errs, warns = apply_prices(data, pl, on,
                                   int(data.get("price_max_age_days", DEFAULT_MAX_AGE_DAYS)))
        data["_price_errors"], data["_price_warnings"] = errs, warns

    for s in data.get("sections") or []:
        for it in (s.get("items") or []) if isinstance(s, dict) else []:
            if isinstance(it, dict) and it.get("unit"):
                it["unit"] = normalize_unit(it["unit"])
    return data


def _pct(rates: dict, key: str, label: str, errors: list[str]) -> None:
    v = rates.get(key)
    if v is None:
        errors.append(f"Не указан процент {label} (rates.{key}).")
        return
    try:
        d = D(v)
    except InvalidOperation:
        errors.append(f"Процент {label} не число: {v!r}.")
        return
    if d < 0 or d > 100:
        errors.append(f"Процент {label} вне диапазона 0–100: {v}.")


def validate(data: dict) -> list[str]:
    """Проверяет исходные данные. Ошибки -> InputError, замечания -> список."""
    errors: list[str] = list(data.get("_price_errors", []))
    warnings: list[str] = list(data.get("_price_warnings", []))

    for k in REQUIRED_TOP:
        if k not in data:
            errors.append(f"Нет обязательного поля «{k}».")
    for k, label in (("company", "реквизиты компании"), ("object", "объект"), ("rates", "ставки")):
        if k in data and not isinstance(data[k], dict):
            errors.append(f"Поле «{k}» ({label}) должно быть объектом {{...}}, а не {type(data[k]).__name__}.")
    if data.get("client") is not None and not isinstance(data["client"], dict):
        errors.append("Поле «client» должно быть объектом, например {\"name\": \"ТОО ...\"}.")
    if isinstance(data.get("sections"), list):
        for si, s in enumerate(data["sections"], 1):
            if not isinstance(s, dict):
                errors.append(f"Раздел {si} должен быть объектом с полями title и items.")
            elif not isinstance(s.get("items", []), list) or any(
                    not isinstance(it, dict) for it in s.get("items", [])):
                errors.append(f"Раздел {si}: items должен быть списком позиций-объектов.")
    if errors:
        raise InputError(errors)

    try:
        date.fromisoformat(str(data["date"]))
    except ValueError:
        errors.append(f"Дата КП должна быть в формате ГГГГ-ММ-ДД: {data['date']!r}.")

    if not str(data["company"].get("name", "")).strip():
        errors.append("В реквизитах компании нет названия (company.name).")
    if not str((data.get("client") or {}).get("name", "")).strip():
        warnings.append("Заказчик не указан: в документе оставлено поле для заполнения.")
    if not str(data["object"].get("name", "")).strip():
        errors.append("Не указан объект (object.name).")

    rates = data["rates"]
    _pct(rates, "overhead_pct", "накладных расходов", errors)
    _pct(rates, "profit_pct", "прибыли", errors)
    _pct(rates, "vat_pct", "НДС", errors)
    if str(rates.get("amount_decimals", 2)) not in ("0", "2"):
        errors.append("rates.amount_decimals может быть только 0 (до тенге) или 2 (до тиын).")
    vd = data.get("valid_days")
    if vd not in (None, ""):
        try:
            ok = 1 <= int(str(vd)) <= 365
        except ValueError:
            ok = False
        if not ok:
            errors.append(f"valid_days: целое число дней от 1 до 365, а не {vd!r}.")
    if data.get("numbering", "section") not in ("section", "continuous"):
        errors.append("numbering может быть «section» (1.1, 1.2) или «continuous» (1, 2, 3).")
    if not str(rates.get("vat_source", "")).strip():
        errors.append("Не указан источник ставки НДС (rates.vat_source): норма Налогового "
                      "кодекса РК и дата проверки.")

    sections = data["sections"]
    if not isinstance(sections, list) or not sections:
        errors.append("Нет ни одного раздела работ (sections).")
        sections = []
    for si, s in enumerate(sections, 1):
        title = str(s.get("title", "")).strip()
        if not title:
            errors.append(f"Раздел {si}: нет названия.")
        items = s.get("items") or []
        if not items:
            errors.append(f"Раздел {si} «{title}»: нет позиций.")
        for ii, it in enumerate(items, 1):
            where = f"Раздел {si}, позиция {ii} «{it.get('name') or it.get('code') or '?'}»"
            if it.get("code") and not data.get("price_list"):
                errors.append(f"{where}: указан код, но в КП нет price_list (путь к прайс-листу).")
            if not str(it.get("name", "")).strip():
                errors.append(f"{where}: нет наименования.")
            unit = str(it.get("unit", "")).strip()
            if not unit:
                errors.append(f"{where}: нет единицы измерения.")
            elif unit not in KNOWN_UNITS:
                warnings.append(f"{where}: нестандартная единица «{unit}», проверьте.")
            for field, label, allow_zero in (("qty", "количество", False), ("price", "цена", True)):
                v = it.get(field)
                try:
                    d = D(v)
                except InvalidOperation:
                    if field == "price" and it.get("code") and data.get("_price_errors"):
                        continue  # причина уже названа в ошибках прайс-листа
                    errors.append(f"{where}: {label} не число или вне допустимого диапазона ({v!r}).")
                    continue
                if d < 0 or (d == 0 and not allow_zero):
                    errors.append(f"{where}: недопустимое значение «{label}»: {v}.")
                elif d == 0:
                    warnings.append(f"{where}: цена равна нулю.")
            if it.get("kind") not in KINDS:
                errors.append(f"{where}: вид затрат «{it.get('kind')}» не из списка: "
                              f"{', '.join(KINDS)}.")
            if not str(it.get("source", "")).strip():
                errors.append(f"{where}: не указан источник цены (source). "
                              "Цены без источника не допускаются.")

    if data.get("sample"):
        warnings.append("Документ помечен как ОБРАЗЕЦ (sample: true), не для отправки заказчику.")
    if errors:
        raise InputError(errors)
    return warnings


def compute(data: dict) -> dict:
    """Расчёт КП. Порядок совпадает с формулами в Excel."""
    r = data["rates"]
    oh_rate = D(r["overhead_pct"]) / 100
    pr_rate = D(r["profit_pct"]) / 100
    vat_rate = D(r["vat_pct"]) / 100
    n = places_of(data)

    by_kind = {k: Decimal("0") for k in KINDS}
    sections = []
    for s in data["sections"]:
        amounts = []
        for it in s["items"]:
            amt = money(D(it["qty"]) * D(it["price"]), n)
            amounts.append(amt)
            by_kind[it["kind"]] += amt
        sections.append({"items": amounts, "subtotal": sum(amounts, Decimal("0"))})

    direct = sum((s["subtotal"] for s in sections), Decimal("0"))
    overhead = money(direct * oh_rate, n)
    profit = money((direct + overhead) * pr_rate, n)
    subtotal = direct + overhead + profit
    vat = money(subtotal * vat_rate, n)
    total = subtotal + vat
    return {
        "sections": sections, "by_kind": by_kind, "direct": direct, "overhead": overhead,
        "profit": profit, "subtotal": subtotal, "vat": vat, "total": total,
        "rates": {"overhead": oh_rate, "profit": pr_rate, "vat": vat_rate},
    }


def vat_exempt(data: dict) -> bool:
    """Ставка НДС 0: компания не плательщик НДС или работы не облагаются."""
    return D(data["rates"]["vat_pct"]) == 0


def valid_until(data: dict) -> date | None:
    days = data.get("valid_days")
    if not days:
        return None
    return date.fromisoformat(str(data["date"])) + timedelta(days=int(days))


def fmt_money(x: Decimal) -> str:
    """1234567.8 -> '1 234 567,80'"""
    return f"{money(x):,.2f}".replace(",", " ").replace(".", ",")


def fmt_qty(x) -> str:
    d = D(x).normalize()
    s = f"{d:,f}".replace(",", " ").replace(".", ",")
    return s


def fmt_pct(x) -> str:
    return f"{D(x).normalize():f}".replace(".", ",") + " %"


def fmt_date(s: str | date) -> str:
    d = s if isinstance(s, date) else date.fromisoformat(str(s))
    return d.strftime("%d.%m.%Y")


def item_numbers(data: dict) -> list[list[str]]:
    """Номера позиций: «1.1, 1.2…» по разделам или сквозные «1, 2, 3…»."""
    out, k = [], 0
    for si, s in enumerate(data["sections"], 1):
        row = []
        for ii, _ in enumerate(s["items"], 1):
            k += 1
            row.append(str(k) if data.get("numbering") == "continuous" else f"{si}.{ii}")
        out.append(row)
    return out


def section_title(data: dict, si: int, title: str) -> str:
    return title if data.get("numbering") == "continuous" else f"Раздел {si}. {title}"


def subtotal_label(data: dict, si: int, title: str) -> str:
    if data.get("numbering") == "continuous":
        return f"Итого: {title}"
    return f"Итого по разделу {si}"


def terms_lines(data: dict) -> list[str]:
    terms = data.get("terms") or {}
    if isinstance(terms, list):
        return [str(t) for t in terms if str(t).strip()]
    labels = {"duration": "Срок выполнения работ", "payment": "Условия оплаты",
              "warranty": "Гарантия", "start": "Начало работ"}
    return [f"{labels.get(k, k)}: {v}" for k, v in terms.items() if v]


# ------------------------------------------------- сумма прописью (RU) ---
_ONES_M = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
_ONES_F = ["", "одна", "две"] + _ONES_M[3:]
_TEENS = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать",
          "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"]
_TENS = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят",
         "восемьдесят", "девяносто"]
_HUNDREDS = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот",
             "восемьсот", "девятьсот"]
_GROUPS = [(("", "", ""), False), (("тысяча", "тысячи", "тысяч"), True),
           (("миллион", "миллиона", "миллионов"), False),
           (("миллиард", "миллиарда", "миллиардов"), False),
           (("триллион", "триллиона", "триллионов"), False)]


def _plural(n: int, forms: tuple[str, str, str]) -> str:
    if 11 <= n % 100 <= 19:
        return forms[2]
    return forms[0] if n % 10 == 1 else forms[1] if 2 <= n % 10 <= 4 else forms[2]


def _triad(n: int, female: bool) -> list[str]:
    words = [_HUNDREDS[n // 100]]
    t = n % 100
    if 10 <= t <= 19:
        words.append(_TEENS[t - 10])
    else:
        words += [_TENS[t // 10], (_ONES_F if female else _ONES_M)[t % 10]]
    return [w for w in words if w]


def int_to_words(n: int) -> str:
    if n == 0:
        return "ноль"
    parts, gi = [], 0
    while n:
        n, tri = divmod(n, 1000)
        if tri:
            forms, female = _GROUPS[gi]
            chunk = _triad(tri, female)
            if gi:
                chunk.append(_plural(tri, forms))
            parts = chunk + parts
        gi += 1
    return " ".join(parts)


def amount_in_words(x: Decimal) -> str:
    """19010544.70 -> 'девятнадцать миллионов десять тысяч пятьсот сорок четыре тенге 70 тиын'"""
    x = money(x)
    tenge = int(x)
    tiyn = int((x - tenge) * 100)
    return f"{int_to_words(tenge)} тенге {tiyn:02d} тиын"
