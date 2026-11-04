"""Акты выполненных работ по договору.

contract.json — замороженные договорные позиции и цены (создаётся из согласованного КП
командой make_contract.py, дальше не меняется без доп. соглашения).
act.json — один акт: период, строки «номер позиции договора — объём — подтверждение»,
ссылки на предыдущие акты для накопительного учёта, зачёт аванса и удержание по договору.
"""
from __future__ import annotations

import json
import sys
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

HERE = Path(__file__).resolve().parent
for d in ("kp", "brand"):
    sys.path.insert(0, str(HERE.parents[1] / d / "scripts"))
from kp_core import D, money  # noqa: E402

FORM_HEADER = ["Приложение 50 к приказу Министра финансов Республики Казахстан",
               "от 20 декабря 2012 года № 562", "Форма Р-1"]


class ActError(Exception):
    def __init__(self, errors: list[str]):
        super().__init__("\n".join(errors))
        self.errors = errors


def _read(p: Path, label: str) -> dict:
    if not p.exists():
        raise ActError([f"Не найден {label}: {p}"])
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise ActError([f"{p.name}: ошибка JSON в строке {e.lineno}: {e.msg}"])


def _date(v, label, errors):
    try:
        return date.fromisoformat(str(v))
    except ValueError:
        errors.append(f"{label}: дата в формате ГГГГ-ММ-ДД, а не {v!r}.")
        return None


def load_contract(path: Path) -> dict:
    c = _read(path, "файл договора")
    errors = []
    for k in ("number", "date", "customer", "contractor", "positions", "vat_pct", "vat_source"):
        if k not in c:
            errors.append(f"Договор: нет поля «{k}».")
    if errors:
        raise ActError(errors)
    c["_date"] = _date(c["date"], "Договор, date", errors)
    c["_places"] = int(c.get("amount_decimals", 2))
    by_n = {}
    for p in c["positions"]:
        try:
            p["_qty"], p["_price"] = D(p["qty"]), D(p["price"])
        except (InvalidOperation, KeyError):
            errors.append(f"Договор, позиция {p.get('n')}: нужны qty и price числами.")
            continue
        if p["n"] in by_n:
            errors.append(f"Договор: номер позиции {p['n']} повторяется.")
        by_n[str(p["n"])] = p
    c["_by_n"] = by_n
    for side in ("customer", "contractor"):
        if not str(c[side].get("name", "")).strip():
            errors.append(f"Договор: нет наименования ({side}.name).")
    if errors:
        raise ActError(errors)
    return c


def act_amounts(contract: dict, lines: list[dict]) -> dict:
    n = contract["_places"]
    rows = []
    for ln in lines:
        p = contract["_by_n"][str(ln["n"])]
        q = D(ln["qty"])
        rows.append({"n": str(ln["n"]), "name": p["name"], "unit": p["unit"], "qty": q,
                     "price": p["_price"], "amount": money(q * p["_price"], n)})
    sub = sum((r["amount"] for r in rows), Decimal(0))
    vat = money(sub * D(contract["vat_pct"]) / 100, n)
    return {"rows": rows, "subtotal": sub, "vat": vat, "total": sub + vat}


def load_act(path: str | Path, _seen=None) -> dict:
    """Загружает акт, договор и предыдущие акты. Проверяет всё; ошибки -> ActError."""
    p = Path(path).resolve()
    _seen = (_seen or set()) | {p}
    a = _read(p, "файл акта")
    a["_path"] = str(p)
    errors, warnings = [], []
    for k in ("contract", "number", "date", "period", "lines"):
        if k not in a:
            errors.append(f"Акт: нет поля «{k}».")
    if errors:
        raise ActError(errors)
    c = load_contract((p.parent / a["contract"]).resolve())
    a["_contract"] = c
    ad = _date(a["date"], "Акт, date", errors)
    per = a["period"] if isinstance(a["period"], dict) else {}
    pf = _date(per.get("from"), "period.from", errors)
    pt = _date(per.get("to"), "period.to", errors)
    if pf and pt and pf > pt:
        errors.append("Период: from позже to.")
    if ad and pt and ad < pt:
        errors.append("Дата акта раньше окончания отчётного периода.")
    if c["_date"] and pf and pf < c["_date"]:
        warnings.append("Период начинается раньше даты договора: проверьте.")
    a["_date"], a["_from"], a["_to"] = ad, pf, pt

    # предыдущие акты
    prev_qty: dict[str, Decimal] = {}
    prev_total = Decimal(0)
    prev_offset = Decimal(0)
    prev_list = []
    for rel in a.get("previous") or []:
        pp = (p.parent / rel).resolve()
        if pp in _seen:
            errors.append(f"Предыдущий акт {rel} ссылается сам на себя или по кругу.")
            continue
        try:
            pa = load_act(pp, _seen)
        except ActError as e:
            errors += [f"Предыдущий акт {rel}: {x}" for x in e.errors]
            continue
        missing = {str(Path(x).name) for x in pa.get("previous") or []} - \
            {str(Path(x).name) for x in a.get("previous") or []}
        if missing:
            errors.append(f"Предыдущий акт {rel} учитывает акты {', '.join(sorted(missing))}, "
                          "их тоже нужно указать в previous текущего акта.")
        if pa["_contract"]["number"] != c["number"]:
            errors.append(f"Предыдущий акт {rel} относится к другому договору.")
            continue
        if pa["_date"] and ad and pa["_date"] > ad:
            errors.append(f"Предыдущий акт № {pa['number']} датирован позже текущего.")
        if str(pa["number"]) == str(a["number"]):
            errors.append(f"Номер акта {a['number']} уже использован.")
        for ln in pa["lines"]:
            prev_qty[str(ln["n"])] = prev_qty.get(str(ln["n"]), Decimal(0)) + D(ln["qty"])
        prev_total += pa["_calc"]["total"]
        prev_offset += pa["_calc"]["offset"]
        prev_list.append({"number": pa["number"], "date": pa["_date"], "total": pa["_calc"]["total"]})
    # previous должен перечислять все более ранние акты: так накопление считается один раз
    a["_prev_qty"], a["_prev_list"] = prev_qty, prev_list

    # строки
    seen_n = set()
    lines = a["lines"] if isinstance(a["lines"], list) else []
    if not lines:
        errors.append("В акте нет строк.")
    for i, ln in enumerate(lines, 1):
        w = f"Строка {i} (позиция {ln.get('n', '?')})"
        n = str(ln.get("n", ""))
        if n not in c["_by_n"]:
            errors.append(f"{w}: позиции нет в договоре.")
            continue
        if n in seen_n:
            errors.append(f"{w}: позиция указана дважды.")
        seen_n.add(n)
        try:
            q = D(ln.get("qty"))
            if q <= 0:
                raise InvalidOperation
        except InvalidOperation:
            errors.append(f"{w}: объём должен быть числом больше нуля.")
            continue
        if not str(ln.get("evidence") or "").strip():
            errors.append(f"{w}: нет подтверждения объёма (evidence): исполнительная схема, акт "
                          "скрытых работ, запись журнала работ. Неподтверждённые объёмы в акт не включаются.")
        if ln.get("date"):
            dd = _date(ln["date"], f"{w}, date", errors)
            if dd and pf and pt and not pf <= dd <= pt:
                errors.append(f"{w}: дата выполнения {dd:%d.%m.%Y} вне отчётного периода.")
        pos = c["_by_n"][n]
        cum = prev_qty.get(n, Decimal(0)) + q
        if cum > pos["_qty"]:
            over = ln.get("overrun_reason") or ""
            if not str(over).strip():
                errors.append(f"{w}: с начала работ {cum} {pos['unit']} больше договорного объёма "
                              f"{pos['_qty']}. Нужно доп. соглашение (overrun_reason).")
            else:
                warnings.append(f"{w}: превышение договорного объёма на {cum - pos['_qty']} "
                                f"{pos['unit']}: {over}.")
    if errors:
        raise ActError(errors)

    calc = act_amounts(c, lines)
    nd = c["_places"]
    adv = a.get("advance") or c.get("advance") or {}
    offset = Decimal(0)
    if adv:
        try:
            paid, pct = D(adv["paid"]), D(adv["offset_pct"])
        except (InvalidOperation, KeyError):
            raise ActError(["advance: нужны paid (сумма полученного аванса) и offset_pct."])
        if not str(adv.get("source", "")).strip():
            raise ActError(["advance.source: пункт договора о зачёте аванса."])
        offset = min(money(calc["total"] * pct / 100, nd), max(paid - prev_offset, Decimal(0)))
    ret = a.get("retention") or c.get("retention") or {}
    retention = Decimal(0)
    if ret:
        if not str(ret.get("source", "")).strip():
            raise ActError(["retention.source: пункт договора о гарантийном удержании."])
        retention = money(calc["total"] * D(ret["pct"]) / 100, nd)
    calc.update(offset=offset, retention=retention, pay=calc["total"] - offset - retention,
                prev_total=prev_total, prev_offset=prev_offset, adv=adv, ret=ret)
    a["_calc"], a["_warnings"] = calc, warnings
    if a.get("sample"):
        warnings.append("Документ помечен как ОБРАЗЕЦ (sample: true).")
    return a
