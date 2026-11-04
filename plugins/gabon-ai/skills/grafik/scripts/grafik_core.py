"""Календарный график производства работ.

Длительность работы считается одним из способов:
- по выработке: объём / (выработка бригады в смену × число бригад × смен в день);
- по трудозатратам: чел.-ч / (рабочих × часов в смене × смен в день), чел.-ч можно взять из сметы;
- фиксированно (duration) с обоснованием;
- веха (milestone): длительность 0.
Связи: FS (окончание — начало) и SS (начало — начало) с задержкой в рабочих днях.
Календарь: выходные по маске Excel («0000011» — сб, вс) и праздники с источником.
Работы с calendar_days: true идут по календарным дням (например, выдерживание бетона).
"""
from __future__ import annotations

import json
import math
import sys
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

HERE = Path(__file__).resolve().parent
for d in ("kp", "vor", "smeta", "brand"):
    sys.path.insert(0, str(HERE.parents[1] / d / "scripts"))
from kp_core import D, normalize_unit  # noqa: E402

ALL_DAYS = "0000000"


class GrafikError(Exception):
    def __init__(self, errors: list[str]):
        super().__init__("\n".join(errors))
        self.errors = errors


# ---------- рабочий календарь (та же логика, что WORKDAY.INTL в Excel) ----------
def is_workday(d: date, weekend: str, holidays: frozenset) -> bool:
    return weekend[d.weekday()] == "0" and d not in holidays


def workday(start: date, n: int, weekend: str, holidays: frozenset) -> date:
    """Как WORKDAY.INTL(start; n; weekend; holidays): n рабочих дней после start."""
    d, step = start, 1 if n >= 0 else -1
    left = abs(n)
    while left:
        d += timedelta(days=step)
        if is_workday(d, weekend, holidays):
            left -= 1
    return d


def workdays_between(a: date, b: date, weekend: str, holidays: frozenset) -> int:
    """Число рабочих дней в (a; b]. Отрицательное, если b < a."""
    if b == a:
        return 0
    sign = 1 if b > a else -1
    lo, hi = (a, b) if b > a else (b, a)
    n, d = 0, lo
    while d < hi:
        d += timedelta(days=1)
        if is_workday(d, weekend, holidays):
            n += 1
    return sign * n


def ceil6(x: Decimal) -> int:
    """ОКРУГЛВВЕРХ(ОКРУГЛ(x; 6); 0) — так же, как формула в Excel, без шума двоичной арифметики."""
    return math.ceil(x.quantize(Decimal("0.000001")))


def _read(p: Path, label: str) -> dict:
    if not p.exists():
        raise GrafikError([f"Не найден {label}: {p}"])
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise GrafikError([f"{p.name}: ошибка JSON в строке {e.lineno}: {e.msg}"])


def load_grafik(path: str | Path) -> dict:
    p = Path(path).resolve()
    data = _read(p, "файл графика")
    base = p.parent
    data["_input_dir"] = str(base)
    comp = data.get("company")
    if isinstance(comp, str):
        cp = (base / comp).resolve()
        data["company"] = _read(cp, "файл реквизитов")
        data["_company_dir"] = str(cp.parent)
    else:
        data["_company_dir"] = str(base)
    data["_brand"] = None
    if data.get("brand"):
        from brand_core import BrandError, load_brand
        try:
            data["_brand"], _ = load_brand((base / data["brand"]).resolve())
        except BrandError as e:
            raise GrafikError([f"Фирменный шаблон: {x}" for x in e.errors])
    data["_vor"], data["_smeta"] = None, None
    if data.get("vor"):
        from vor_core import VorError, compute as vc, load_vor, rows
        try:
            vor = load_vor((base / data["vor"]).resolve())
            vals, _ = vc(vor)
        except VorError as e:
            raise GrafikError([f"ВОР: {x}" for x in e.errors])
        data["_vor"] = {"number": vor["number"], "values": vals,
                        "rows": {it["_id"]: (f"{si}.{ii}", it) for si, ii, it in rows(vor)}}
    if data.get("smeta"):
        from smeta_core import SmetaError, compute as sc, load_smeta
        try:
            sm = load_smeta((base / data["smeta"]).resolve())
            calc, _ = sc(sm)
        except SmetaError as e:
            raise GrafikError([f"Смета: {x}" for x in e.errors])
        pos = {}
        for si, s in enumerate(calc["sections"], 1):
            for pi, ps in enumerate(s["positions"], 1):
                pos[f"{si}.{pi}"] = ps
        data["_smeta"] = {"number": sm["number"], "positions": pos}
    return data


def _deps(t: dict) -> list[dict]:
    out = []
    for d in t.get("after") or []:
        if isinstance(d, str):
            d = {"id": d}
        out.append({"id": str(d.get("id", "")), "type": str(d.get("type", "FS")).upper(),
                    "lag": d.get("lag", 0)})
    return out


def compute(data: dict) -> tuple[dict, list[str]]:
    errors, warnings = [], []
    for k in ("number", "date", "object", "start", "tasks"):
        if k not in data:
            errors.append(f"Нет обязательного поля «{k}».")
    if errors:
        raise GrafikError(errors)
    try:
        start = date.fromisoformat(str(data["start"]))
        date.fromisoformat(str(data["date"]))
    except ValueError:
        raise GrafikError(["Даты date и start в формате ГГГГ-ММ-ДД."])
    cal = data.get("calendar") or {}
    weekend = str(cal.get("weekend", "0000011"))
    if len(weekend) != 7 or set(weekend) - {"0", "1"} or weekend == "1111111":
        errors.append("calendar.weekend: 7 символов 0/1 с понедельника, например «0000011».")
        weekend = "0000011"
    hol = []
    for h in cal.get("holidays") or []:
        try:
            hol.append(date.fromisoformat(str(h)))
        except ValueError:
            errors.append(f"calendar.holidays: дата «{h}» не в формате ГГГГ-ММ-ДД.")
    if hol and not str(cal.get("holidays_source", "")).strip():
        errors.append("calendar.holidays_source: укажите основание списка праздников.")
    if not hol:
        warnings.append("Праздничные дни не заданы: проверьте календарь на период работ.")
    holidays = frozenset(hol)
    try:
        shifts = int(cal.get("shifts_per_day", 1))
        hours = D(cal.get("shift_hours", 8))
        if shifts < 1 or hours <= 0:
            raise ValueError
    except (ValueError, InvalidOperation):
        errors.append("calendar.shifts_per_day ≥ 1 и calendar.shift_hours > 0.")
        shifts, hours = 1, Decimal(8)
    tasks = data["tasks"]
    if not isinstance(tasks, list) or not tasks or any(not isinstance(t, dict) for t in tasks):
        raise GrafikError(errors + ["tasks: нужен непустой список работ."])

    by_id = {}
    for i, t in enumerate(tasks, 1):
        tid = str(t.get("id") or i)
        t["_id"], t["_n"] = tid, str(i)
        if tid in by_id:
            errors.append(f"Работа {i}: id «{tid}» повторяется.")
        by_id[tid] = t
    for i, t in enumerate(tasks, 1):
        w = f"Работа {i} «{t.get('name') or t.get('vor_id') or '?'}»"
        t["_where"] = w
        t["_cal"] = ALL_DAYS if t.get("calendar_days") else weekend
        t["_hol"] = frozenset() if t.get("calendar_days") else holidays
        for dep in _deps(t):
            if dep["id"] not in by_id:
                errors.append(f"{w}: предшественник «{dep['id']}» не найден.")
            if dep["type"] not in ("FS", "SS"):
                errors.append(f"{w}: тип связи «{dep['type']}»: FS или SS.")
            if not isinstance(dep["lag"], int) or dep["lag"] < 0:
                errors.append(f"{w}: задержка lag — целое число дней ≥ 0.")
        # объём и способ расчёта длительности
        qty, unit, qsrc = None, normalize_unit(t.get("unit")), ""
        if t.get("vor_id"):
            v = data.get("_vor")
            if v is None:
                errors.append(f"{w}: указан vor_id, но нет поля vor.")
                continue
            if t["vor_id"] not in v["rows"]:
                errors.append(f"{w}: строки «{t['vor_id']}» нет в ВОР.")
                continue
            num, it = v["rows"][t["vor_id"]]
            if it["status"] == "нет данных":
                errors.append(f"{w}: строка ВОР {num} без данных. Задайте duration с обоснованием.")
                continue
            qty, unit = v["values"][t["vor_id"]], it["unit"]
            qsrc = f"ВОР № {v['number']}, строка {num}"
            t.setdefault("name", it["name"])
            if it["status"] == "допущение":
                warnings.append(f"{w}: объём из ВОР со статусом «допущение».")
        elif t.get("qty") not in (None, ""):
            try:
                qty = D(t["qty"])
            except InvalidOperation:
                errors.append(f"{w}: qty не число.")
            qsrc = str(t.get("qty_source") or "")
        t["_qty"], t["_unit"], t["_qsrc"] = qty, unit, qsrc

        modes = [m for m in ("productivity", "labor", "duration", "milestone") if t.get(m) not in (None, False, "")]
        if t.get("smeta_pos") and "labor" not in modes:
            modes.append("labor")
        if len(modes) != 1:
            errors.append(f"{w}: укажите ровно один способ: productivity, labor (или smeta_pos), "
                          "duration или milestone.")
            continue
        mode = modes[0]
        t["_mode"] = mode
        if mode == "productivity":
            p = t["productivity"]
            try:
                out_, crews = D(p.get("per_shift")), int(t.get("crews", 1))
                if out_ <= 0 or crews < 1:
                    raise InvalidOperation
            except (InvalidOperation, ValueError, TypeError):
                errors.append(f"{w}: productivity.per_shift > 0 и crews ≥ 1.")
                continue
            if qty is None or qty <= 0:
                errors.append(f"{w}: для расчёта по выработке нужен объём (vor_id или qty).")
                continue
            if not str(p.get("source", "")).strip():
                errors.append(f"{w}: productivity.source — откуда выработка (ЕНиР, ППР, опыт бригады…).")
            if p.get("unit") and normalize_unit(p["unit"]) != unit:
                errors.append(f"{w}: выработка в «{p['unit']}», а объём в «{unit}».")
            t["_out"], t["_crews"] = out_, crews
            t["_dur"] = ceil6(qty / (out_ * crews * shifts))
        elif mode == "labor":
            lab = t.get("labor") or {}
            try:
                workers = int(lab.get("workers") or t.get("workers"))
                if workers < 1:
                    raise ValueError
            except (ValueError, TypeError):
                errors.append(f"{w}: labor.workers — число рабочих ≥ 1.")
                continue
            hrs = None
            if t.get("smeta_pos"):
                sm = data.get("_smeta")
                if sm is None or t["smeta_pos"] not in sm["positions"]:
                    errors.append(f"{w}: позиция сметы «{t['smeta_pos']}» не найдена (нужно поле smeta).")
                    continue
                ps = sm["positions"][t["smeta_pos"]]
                hrs = sum((r["qty"] for r in ps["resources"] if r["type"] == "труд"
                           and r["unit"] == "чел.-ч"), Decimal(0))
                if hrs <= 0:
                    errors.append(f"{w}: в позиции сметы {t['smeta_pos']} нет трудозатрат в чел.-ч.")
                    continue
                t["_lsrc"] = f"Смета № {sm['number']}, позиция {t['smeta_pos']}"
                t.setdefault("name", ps["name"])
                if qty is None:
                    qty, t["_qty"], t["_unit"] = ps["qty"], ps["qty"], ps["unit"]
            else:
                try:
                    hrs = D(lab.get("hours"))
                    if hrs <= 0:
                        raise InvalidOperation
                except InvalidOperation:
                    errors.append(f"{w}: labor.hours (чел.-ч) > 0 или smeta_pos.")
                    continue
                t["_lsrc"] = str(lab.get("source") or "")
                if not t["_lsrc"]:
                    errors.append(f"{w}: labor.source — откуда трудозатраты.")
            t["_hours"], t["_workers"] = hrs, workers
            t["_dur"] = ceil6(hrs / (workers * hours * shifts))
        elif mode == "duration":
            try:
                dd = int(t["duration"])
                if dd < 1:
                    raise ValueError
            except (ValueError, TypeError):
                errors.append(f"{w}: duration — целое число дней ≥ 1.")
                continue
            if not str(t.get("duration_source") or t.get("assumption") or "").strip():
                errors.append(f"{w}: фиксированная длительность требует duration_source или assumption.")
            t["_dur"] = dd
        else:
            t["_dur"] = 0
        if not str(t.get("name") or "").strip():
            errors.append(f"{w}: нет наименования.")
        if t["_dur"] > 3650:
            errors.append(f"{w}: длительность {t['_dur']} дн. нереальна, проверьте данные.")
    if errors:
        raise GrafikError(errors)

    # прямой проход (порядок с учётом связей, проверка циклов)
    order, state = [], {}

    def visit(tid, chain):
        if state.get(tid) == "done":
            return
        if state.get(tid) == "busy":
            raise GrafikError(["Циклическая связь: " + " -> ".join(chain + [tid])])
        state[tid] = "busy"
        for dep in _deps(by_id[tid]):
            visit(dep["id"], chain + [tid])
        state[tid] = "done"
        order.append(tid)
    for t in tasks:
        visit(t["_id"], [])

    for tid in order:
        t = by_id[tid]
        cal, hl = t["_cal"], t["_hol"]
        cands = [workday(start - timedelta(days=1), 1, cal, hl)]
        if t.get("not_before"):
            nb = date.fromisoformat(str(t["not_before"]))
            cands.append(workday(nb - timedelta(days=1), 1, cal, hl))
        for dep in _deps(t):
            p = by_id[dep["id"]]
            if dep["type"] == "SS":
                cands.append(workday(p["_es"], dep["lag"], cal, hl))
            elif t["_dur"] == 0:
                cands.append(workday(p["_ef"], dep["lag"], cal, hl))
            else:
                cands.append(workday(p["_ef"], 1 + dep["lag"], cal, hl))
        t["_es"] = max(cands)
        t["_ef"] = t["_es"] if t["_dur"] == 0 else workday(t["_es"], t["_dur"] - 1, cal, hl)

    finish = max(t["_ef"] for t in tasks)
    # обратный проход: поздние сроки и резерв
    succ = {t["_id"]: [] for t in tasks}
    for t in tasks:
        for dep in _deps(t):
            succ[dep["id"]].append((t, dep))
    for tid in reversed(order):
        t = by_id[tid]
        cal, hl = t["_cal"], t["_hol"]
        lims = [finish]
        for s, dep in succ[tid]:
            scal, shl = s["_cal"], s["_hol"]
            if dep["type"] == "SS":
                ls = workday(s["_ls"], -dep["lag"], scal, shl)
                lims.append(ls if t["_dur"] == 0 else workday(ls, t["_dur"] - 1, cal, hl))
            elif s["_dur"] == 0:
                lims.append(workday(s["_ls"], -dep["lag"], scal, shl))
            else:
                lims.append(workday(s["_ls"], -(1 + dep["lag"]), scal, shl))
        lf = min(lims)
        if not is_workday(lf, cal, hl) and t["_dur"] > 0:
            lf = workday(lf, -1, cal, hl)
        t["_lf"] = lf
        t["_ls"] = lf if t["_dur"] == 0 else workday(lf, -(t["_dur"] - 1), cal, hl)
        t["_float"] = max(0, workdays_between(t["_ef"], t["_lf"], cal, hl))
        t["_critical"] = t["_float"] == 0

    total_wd = workdays_between(workday(start - timedelta(days=1), 1, weekend, holidays), finish,
                                weekend, holidays) + 1
    n_assume = sum(1 for t in tasks if t.get("assumption"))
    if n_assume:
        warnings.append(f"Работ с допущениями: {n_assume}. Требуют подтверждения в ППР.")
    if data.get("sample"):
        warnings.append("Документ помечен как ОБРАЗЕЦ (sample: true).")
    return {"start": min(t["_es"] for t in tasks), "finish": finish, "workdays": total_wd,
            "calendar_days": (finish - min(t["_es"] for t in tasks)).days + 1,
            "weekend": weekend, "holidays": sorted(holidays), "shifts": shifts,
            "shift_hours": hours, "order": order}, warnings
