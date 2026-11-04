"""Локальный сметный расчёт (коммерческий, ресурсный метод).

Каждая позиция раскладывается на ресурсы: труд (ФОТ), машины (ЭМ), материалы, транспорт.
Норма расхода ресурса на единицу работы берётся из калькуляции (карточки) с источником,
цена ресурса — из прайс-листа по коду или с указанием источника.
Количество позиции — числом или из ВОР по id строки.
"""
from __future__ import annotations

import json
import sys
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

HERE = Path(__file__).resolve().parent
for d in ("kp", "prices", "vor", "brand"):
    sys.path.insert(0, str(HERE.parents[1] / d / "scripts"))
from kp_core import D, money, normalize_unit  # noqa: E402

TYPES = {"труд": "ФОТ (труд рабочих)", "машины": "Эксплуатация машин",
         "материал": "Материалы", "транспорт": "Транспорт"}
TYPE_TO_KIND = {"труд": "работа", "машины": "механизм", "материал": "материал",
                "транспорт": "транспорт"}
OVERHEAD_BASES = {"direct": "от прямых затрат", "fot": "от ФОТ"}
PROFIT_BASES = {"direct_overhead": "от (прямые затраты + накладные)", "fot": "от ФОТ"}
RES_PLACES = 3


class SmetaError(Exception):
    def __init__(self, errors: list[str]):
        super().__init__("\n".join(errors))
        self.errors = errors


def rq(x: Decimal) -> Decimal:
    return money(x, RES_PLACES)


def _read_json(p: Path, label: str) -> dict:
    if not p.exists():
        raise SmetaError([f"Не найден {label}: {p}"])
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise SmetaError([f"{p.name}: ошибка JSON в строке {e.lineno}: {e.msg}"])


def load_smeta(path: str | Path) -> dict:
    p = Path(path).resolve()
    data = _read_json(p, "файл сметы")
    base = p.parent
    data["_input_dir"] = str(base)
    comp = data.get("company")
    if isinstance(comp, str):
        cp = (base / comp).resolve()
        data["company"] = _read_json(cp, "файл реквизитов")
        data["_company_dir"] = str(cp.parent)
    else:
        data["_company_dir"] = str(base)
    data["_brand"] = None
    if data.get("brand"):
        from brand_core import BrandError, load_brand
        try:
            data["_brand"], _ = load_brand((base / data["brand"]).resolve())
        except BrandError as e:
            raise SmetaError([f"Фирменный шаблон: {x}" for x in e.errors])
    data["_cards"] = {}
    if data.get("cards"):
        c = _read_json((base / data["cards"]).resolve(), "файл калькуляций")
        data["_cards"] = c.get("cards", c)
        data["_cards_file"] = Path(data["cards"]).name
    data["_vor_values"], data["_vor_rows"] = None, {}
    if data.get("vor"):
        from vor_core import VorError, compute as vor_compute, load_vor, rows as vor_rows
        try:
            vor = load_vor((base / data["vor"]).resolve())
            data["_vor_values"], _ = vor_compute(vor)
        except VorError as e:
            raise SmetaError([f"ВОР: {x}" for x in e.errors])
        data["_vor_rows"] = {it["_id"]: (f"{si}.{ii}", it) for si, ii, it in vor_rows(vor)}
        data["_vor_number"] = vor["number"]
    data["_price_list"] = None
    if data.get("price_list"):
        from prices_core import load_price_list
        data["_price_list"] = load_price_list((base / data["price_list"]).resolve())
    return data


def positions(data: dict):
    for si, s in enumerate(data.get("sections") or [], 1):
        for pi, pos in enumerate(s.get("positions") or [], 1):
            yield si, pi, pos


def _check_rates(rates, errors):
    if not isinstance(rates, dict):
        errors.append("rates должен быть объектом.")
        return
    for k, label in (("overhead_pct", "накладных"), ("profit_pct", "прибыли"), ("vat_pct", "НДС")):
        try:
            v = D(rates.get(k))
            if not 0 <= v <= 100:
                raise InvalidOperation
        except InvalidOperation:
            errors.append(f"rates.{k}: процент {label} от 0 до 100.")
    if rates.get("overhead_base", "direct") not in OVERHEAD_BASES:
        errors.append(f"rates.overhead_base: {', '.join(OVERHEAD_BASES)}.")
    if rates.get("profit_base", "direct_overhead") not in PROFIT_BASES:
        errors.append(f"rates.profit_base: {', '.join(PROFIT_BASES)}.")
    if not str(rates.get("vat_source", "")).strip():
        errors.append("rates.vat_source: укажите основание ставки НДС.")


def compute(data: dict) -> tuple[dict, list[str]]:
    """Проверяет смету и считает её. Возвращает (расчёт, замечания)."""
    errors: list[str] = []
    warnings: list[str] = []
    for k in ("number", "date", "object", "sections", "rates"):
        if k not in data:
            errors.append(f"Нет обязательного поля «{k}».")
    if errors:
        raise SmetaError(errors)
    try:
        on = date.fromisoformat(str(data["date"]))
    except ValueError:
        errors.append(f"Дата должна быть ГГГГ-ММ-ДД: {data['date']!r}.")
        on = date.today()
    if not isinstance(data["object"], dict) or not data["object"].get("name"):
        errors.append("Не указан объект (object.name).")
    _check_rates(data["rates"], errors)
    pl = data.get("_price_list")
    if pl is not None and pl.errors:
        errors += [f"Прайс-лист: {e}" for e in pl.errors]
    if not isinstance(data["sections"], list) or not data["sections"] or any(
            not isinstance(s, dict) or not isinstance(s.get("positions"), list) or not s["positions"]
            for s in data["sections"]):
        errors.append("sections: нужен список разделов с непустыми positions.")
    if errors:
        raise SmetaError(errors)

    max_age = int(data.get("price_max_age_days", 90))
    places = int(data["rates"].get("amount_decimals", 2))
    totals = {t: Decimal(0) for t in TYPES}
    calc_sections = []
    for si, s in enumerate(data["sections"], 1):
        cs = []
        for pi, pos in enumerate(s["positions"], 1):
            if not isinstance(pos, dict):
                errors.append(f"Позиция {si}.{pi}: должна быть объектом.")
                continue
            vr0 = data["_vor_rows"].get(pos.get("vor_id")) if pos.get("vor_id") else None
            label = (pos.get("name") or (vr0[1]["name"] if vr0 else None)
                     or (data["_cards"].get(pos.get("card")) or {}).get("name") or "?")
            where = f"Позиция {si}.{pi} «{label}»"
            card = None
            if pos.get("card"):
                card = data["_cards"].get(pos["card"])
                if card is None:
                    errors.append(f"{where}: калькуляция «{pos['card']}» не найдена в файле калькуляций.")
                    continue
            # количество
            qty, qty_src = None, ""
            if pos.get("vor_id"):
                if data["_vor_values"] is None:
                    errors.append(f"{where}: указан vor_id, но в смете нет поля vor.")
                    continue
                vr = data["_vor_rows"].get(pos["vor_id"])
                if vr is None:
                    errors.append(f"{where}: строки «{pos['vor_id']}» нет в ВОР.")
                    continue
                num, vit = vr
                if vit["status"] == "нет данных":
                    errors.append(f"{where}: строка ВОР {num} без данных, количества нет.")
                    continue
                qty = data["_vor_values"][pos["vor_id"]]
                qty_src = f"ВОР № {data['_vor_number']}, строка {num}"
                pos.setdefault("name", vit["name"])
                pos.setdefault("unit", vit["unit"])
                if vit["status"] == "допущение":
                    warnings.append(f"{where}: количество из ВОР со статусом «допущение».")
                if pos.get("qty") not in (None, "") and D(pos["qty"]) != qty:
                    errors.append(f"{where}: qty {pos['qty']} не совпадает с ВОР ({qty}).")
                    continue
            else:
                try:
                    qty = D(pos.get("qty"))
                    if qty <= 0:
                        raise InvalidOperation
                except InvalidOperation:
                    errors.append(f"{where}: нужно количество больше нуля (qty) или vor_id.")
                    continue
                qty_src = str(pos.get("qty_source") or "").strip()
                if not qty_src:
                    errors.append(f"{where}: укажите qty_source (откуда количество) или vor_id.")
            if card and not pos.get("name"):
                pos["name"] = card.get("name", "")
            if not str(pos.get("name") or "").strip():
                errors.append(f"{where}: нет наименования.")
            unit = normalize_unit(pos.get("unit") or (card or {}).get("unit"))
            if not unit:
                errors.append(f"{where}: нет единицы измерения.")
            if card and normalize_unit(card.get("unit")) != unit:
                errors.append(f"{where}: единица позиции «{unit}», калькуляции «{card.get('unit')}».")
            pos["unit"] = unit
            resources = (card or {}).get("resources") if card else pos.get("resources")
            norm_src_default = (card or {}).get("source", "")
            if not resources:
                errors.append(f"{where}: нет ресурсов (ни card, ни resources).")
                continue
            calc_res = []
            for ri, r in enumerate(resources, 1):
                rw = f"{where}, ресурс {ri} «{r.get('name') or r.get('code') or '?'}»"
                t = r.get("type")
                if t not in TYPES:
                    errors.append(f"{rw}: тип «{t}» не из списка: {', '.join(TYPES)}.")
                    continue
                try:
                    norm = D(r.get("norm"))
                    if norm <= 0:
                        raise InvalidOperation
                except InvalidOperation:
                    errors.append(f"{rw}: норма расхода на единицу (norm) должна быть больше нуля.")
                    continue
                nsrc = str(r.get("source") or norm_src_default).strip()
                if not nsrc:
                    errors.append(f"{rw}: нет источника нормы (source).")
                runit = normalize_unit(r.get("unit"))
                name = r.get("name")
                price, psrc = None, ""
                if r.get("code"):
                    if pl is None:
                        errors.append(f"{rw}: указан code, но в смете нет price_list.")
                        continue
                    rec = pl.get(r["code"])
                    if rec is None:
                        errors.append(f"{rw}: код «{r['code']}» не найден в прайс-листе.")
                        continue
                    if runit and runit != rec.unit:
                        errors.append(f"{rw}: единица «{runit}», в прайсе «{rec.unit}».")
                        continue
                    if rec.kind != TYPE_TO_KIND[t] and rec.kind != "комплекс":
                        errors.append(f"{rw}: тип «{t}», а в прайсе вид «{rec.kind}».")
                        continue
                    runit, name = rec.unit, name or rec.name
                    price = rec.price
                    psrc = f"Прайс {pl.name}, код {rec.code}, от {rec.date:%d.%m.%Y}"
                    from prices_core import age_warning
                    w = age_warning(rec, on, max_age)
                    if w:
                        warnings.append(f"{rw}: {w}")
                else:
                    try:
                        price = D(r.get("price"))
                        if price < 0:
                            raise InvalidOperation
                    except InvalidOperation:
                        errors.append(f"{rw}: нет цены (price) или кода прайса (code).")
                        continue
                    psrc = str(r.get("price_source") or "").strip()
                    if not psrc:
                        errors.append(f"{rw}: укажите price_source (источник цены).")
                if not runit or not name:
                    errors.append(f"{rw}: нужны name и unit.")
                    continue
                res_qty = rq(norm * qty)
                amount = money(res_qty * price, places)
                totals[t] += amount
                calc_res.append({"type": t, "name": name, "unit": runit, "norm": norm,
                                 "qty": res_qty, "price": price, "amount": amount,
                                 "norm_source": nsrc, "price_source": psrc})
            ptot = sum((x["amount"] for x in calc_res), Decimal(0))
            by_type = {t: sum((x["amount"] for x in calc_res if x["type"] == t), Decimal(0))
                       for t in TYPES}
            cs.append({"name": pos.get("name"), "unit": unit, "qty": qty, "qty_source": qty_src,
                       "card": pos.get("card"), "resources": calc_res, "total": ptot,
                       "by_type": by_type, "code": pos.get("code")})
        calc_sections.append({"title": s.get("title", ""), "positions": cs})
    if errors:
        raise SmetaError(errors)

    r = data["rates"]
    direct = sum(totals.values(), Decimal(0))
    fot = totals["труд"]
    oh_base = fot if r.get("overhead_base", "direct") == "fot" else direct
    overhead = money(oh_base * D(r["overhead_pct"]) / 100, places)
    pr_base = fot if r.get("profit_base", "direct_overhead") == "fot" else direct + overhead
    profit = money(pr_base * D(r["profit_pct"]) / 100, places)
    subtotal = direct + overhead + profit
    vat = money(subtotal * D(r["vat_pct"]) / 100, places)
    if data.get("sample"):
        warnings.append("Документ помечен как ОБРАЗЕЦ (sample: true).")
    return {"sections": calc_sections, "by_type": totals, "direct": direct, "fot": fot,
            "overhead": overhead, "profit": profit, "subtotal": subtotal, "vat": vat,
            "total": subtotal + vat, "places": places}, warnings
