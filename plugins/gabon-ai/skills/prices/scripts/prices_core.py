"""Прайс-лист компании: загрузка, проверка, поиск и подстановка цен в КП.

Формат: Excel, лист «Прайс» (или первый лист), строка заголовков с колонками
Код | Наименование | Ед. изм. | Цена без НДС, тг | Вид затрат | Источник | Дата цены | Примечание
"""
from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "kp" / "scripts"))
from kp_core import KINDS, KNOWN_UNITS, normalize_unit  # noqa: E402

COLUMNS = {  # ключ -> варианты заголовка (в нижнем регистре, начало строки)
    "code": ("код", "шифр", "артикул"),
    "name": ("наименование",),
    "unit": ("ед",),
    "price": ("цена",),
    "kind": ("вид",),
    "source": ("источник",),
    "date": ("дата",),
    "note": ("примеч",),
}
REQUIRED = ("code", "name", "unit", "price", "kind", "source", "date")
DEFAULT_MAX_AGE_DAYS = 90

def _norm_text(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").lower().replace("ё", "е")).strip()


@dataclass
class PriceRecord:
    code: str
    name: str
    unit: str
    price: Decimal
    kind: str
    source: str
    date: date
    note: str = ""
    row: int = 0


@dataclass
class PriceList:
    path: Path
    records: dict[str, PriceRecord] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def name(self) -> str:
        return self.path.name

    def get(self, code: str) -> PriceRecord | None:
        return self.records.get(str(code).strip())

    def find(self, query: str, limit: int = 10) -> list[tuple[int, PriceRecord]]:
        q = _norm_text(query)
        if q.upper() in (c.upper() for c in self.records):
            rec = next(r for c, r in self.records.items() if c.upper() == q.upper())
            return [(100, rec)]
        words = [w for w in re.split(r"[^\wа-я0-9.,×х]+", q) if len(w) >= 2]
        scored = []
        for rec in self.records.values():
            hay = _norm_text(f"{rec.code} {rec.name} {rec.note}")
            score = sum(1 for w in words if w in hay)
            if score:
                scored.append((score * 100 // max(len(words), 1), rec))
        scored.sort(key=lambda x: (-x[0], x[1].code))
        return scored[:limit]


def _to_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or "").strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _to_dec(v) -> Decimal | None:
    if isinstance(v, bool) or v is None or v == "":
        return None
    try:
        return Decimal(str(v).replace(" ", "").replace(" ", "").replace(",", "."))
    except InvalidOperation:
        return None


def load_price_list(path: str | Path) -> PriceList:
    from openpyxl import load_workbook
    p = Path(path).resolve()
    pl = PriceList(p)
    if not p.exists():
        pl.errors.append(f"Не найден прайс-лист: {p}")
        return pl
    wb = load_workbook(p, data_only=True, read_only=True)
    ws = wb["Прайс"] if "Прайс" in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))

    hdr_i, cols = None, {}
    for i, row in enumerate(rows[:20]):
        found = {}
        for ci, v in enumerate(row):
            t = _norm_text(v)
            for key, starts in COLUMNS.items():
                if key not in found and any(t.startswith(s) for s in starts):
                    found[key] = ci
        if "code" in found and "price" in found:
            hdr_i, cols = i, found
            break
    if hdr_i is None:
        pl.errors.append("Не найдена строка заголовков (нужны хотя бы колонки «Код» и «Цена»).")
        return pl
    for k in REQUIRED:
        if k not in cols:
            pl.errors.append(f"В прайсе нет колонки «{COLUMNS[k][0]}».")
    if pl.errors:
        return pl
    price_header = _norm_text(rows[hdr_i][cols["price"]])
    if "без ндс" not in price_header:
        pl.errors.append(f"Заголовок цены «{rows[hdr_i][cols['price']]}» должен явно говорить "
                         "«без НДС». Цены с НДС в прайс не вносятся.")

    for i, row in enumerate(rows[hdr_i + 1:], hdr_i + 2):
        def cell(k):
            ci = cols.get(k)
            return row[ci] if ci is not None and ci < len(row) else None
        if all(v in (None, "") for v in row):
            continue
        code = str(cell("code") or "").strip()
        where = f"Строка {i}" + (f" (код {code})" if code else "")
        if not code:
            pl.errors.append(f"{where}: нет кода.")
            continue
        if code in pl.records:
            pl.errors.append(f"{where}: код повторяется (уже в строке {pl.records[code].row}).")
            continue
        name = str(cell("name") or "").strip()
        unit = normalize_unit(cell("unit"))
        price = _to_dec(cell("price"))
        kind = str(cell("kind") or "").strip()
        source = str(cell("source") or "").strip()
        d = _to_date(cell("date"))
        bad = False
        if not name:
            pl.errors.append(f"{where}: нет наименования."); bad = True
        if not unit:
            pl.errors.append(f"{where}: нет единицы измерения."); bad = True
        elif unit not in KNOWN_UNITS:
            pl.warnings.append(f"{where}: нестандартная единица «{unit}».")
        if price is None:
            pl.errors.append(f"{where}: цена не число ({cell('price')!r})."); bad = True
        elif price < 0:
            pl.errors.append(f"{where}: отрицательная цена."); bad = True
        elif price == 0:
            pl.warnings.append(f"{where}: цена равна нулю.")
        if kind not in KINDS:
            pl.errors.append(f"{where}: вид затрат «{kind}» не из списка: {', '.join(KINDS)}."); bad = True
        if not source:
            pl.errors.append(f"{where}: нет источника цены."); bad = True
        if d is None:
            pl.errors.append(f"{where}: нет даты цены или дата не распознана ({cell('date')!r})."); bad = True
        if not bad:
            pl.records[code] = PriceRecord(code, name, unit, price, kind, source, d,
                                           str(cell("note") or "").strip(), i)
    if not pl.records and not pl.errors:
        pl.errors.append("Прайс-лист пуст.")
    return pl


def age_warning(rec: PriceRecord, on: date, max_age: int) -> str | None:
    age = (on - rec.date).days
    if age > max_age:
        return (f"цене {age} дн. (от {rec.date:%d.%m.%Y}), больше {max_age} дн. "
                "Актуализируйте цену.")
    if age < 0:
        return f"дата цены {rec.date:%d.%m.%Y} позже даты документа."
    return None


def apply_prices(data: dict, pl: PriceList, on: date, max_age: int = DEFAULT_MAX_AGE_DAYS
                 ) -> tuple[list[str], list[str]]:
    """Подставляет цены в позиции КП, у которых указан code. Меняет data на месте."""
    errors, warnings = list(pl.errors), [f"Прайс: {w}" for w in pl.warnings]
    if pl.errors:
        return errors, warnings
    for si, s in enumerate(data.get("sections", []), 1):
        for ii, it in enumerate(s.get("items", []), 1):
            code = str(it.get("code") or "").strip()
            if not code:
                continue
            where = f"Раздел {si}, позиция {ii} (код {code})"
            rec = pl.get(code)
            if rec is None:
                errors.append(f"{where}: код не найден в прайс-листе {pl.name}.")
                continue
            if it.get("unit") and normalize_unit(it["unit"]) != rec.unit:
                errors.append(f"{where}: единица в КП «{it['unit']}», в прайсе «{rec.unit}».")
                continue
            if it.get("kind") and it["kind"] != rec.kind:
                errors.append(f"{where}: вид затрат в КП «{it['kind']}», в прайсе «{rec.kind}».")
                continue
            it["unit"] = rec.unit
            it["kind"] = rec.kind
            if not str(it.get("name") or "").strip():
                it["name"] = rec.name
            list_src = (f"Прайс-лист {pl.name}, код {rec.code}, цена от "
                        f"{rec.date:%d.%m.%Y} ({rec.source})")
            manual = _to_dec(it.get("price")) if it.get("price") not in (None, "") else None
            if manual is not None and manual != rec.price:
                reason = str(it.get("price_override_reason", "")).strip()
                if not reason:
                    errors.append(f"{where}: цена в КП {it['price']} отличается от прайса "
                                  f"{rec.price}. Уберите цену из КП или укажите "
                                  "price_override_reason.")
                    continue
                it["source"] = (f"Цена изменена вручную: {reason}. Цена прайса {rec.price} "
                                f"({list_src})")
                warnings.append(f"{where}: цена изменена вручную ({it['price']} вместо "
                                f"{rec.price}): {reason}.")
                continue
            it["price"] = str(rec.price)
            it["source"] = list_src
            w = age_warning(rec, on, max_age)
            if w:
                warnings.append(f"{where}: {w}")
    return errors, warnings
