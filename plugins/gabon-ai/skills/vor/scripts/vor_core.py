"""Ведомость объёмов работ (ВОР): загрузка, проверка, расчёт количеств по формулам.

Количество задаётся формулой подсчёта из размеров («24 × 0,6 × 1,2»), числом из документа
заказчика или ссылкой на другие строки («@выемка - @бетон»). Каждая строка имеет статус:
подтверждено (есть источник), допущение (с обоснованием) или нет данных.
"""
from __future__ import annotations

import ast
import json
import operator
import re
import sys
from datetime import date
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "kp" / "scripts"))
from kp_core import KNOWN_UNITS, MAX_NUMBER, D, normalize_unit  # noqa: E402

STATUSES = ("подтверждено", "допущение", "нет данных")
REF = re.compile(r"@([\wА-Яа-яЁё]+(?:\.[\wА-Яа-яЁё]+)*)")
_OPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
        ast.Div: operator.truediv}


class VorError(Exception):
    def __init__(self, errors: list[str]):
        super().__init__("\n".join(errors))
        self.errors = errors


def normalize_formula(f: str) -> str:
    """«24 × 0,6 х 1,2» -> «24*0.6*1.2». Ссылки @id не трогает."""
    s = str(f).strip().lstrip("=")
    s = re.sub(r"(?<=[\d)\s])[×хХxX·](?=[\s\d(@])", "*", s)
    s = s.replace("×", "*").replace("·", "*").replace("−", "-").replace("–", "-")
    s = re.sub(r"(?<=\d),(?=\d)", ".", s)
    return re.sub(r"\s+", "", s)


def parse_formula(f: str) -> tuple[ast.AST, list[str]]:
    """Разбирает формулу в безопасное дерево. Возвращает (дерево, ссылки)."""
    norm = normalize_formula(f)
    refs = REF.findall(norm)
    py = REF.sub(lambda m: f"_r{refs.index(m[1])}", norm)
    try:
        tree = ast.parse(py, mode="eval")
    except SyntaxError:
        raise ValueError(f"формула не разобрана: «{f}»")
    for n in ast.walk(tree):
        if isinstance(n, (ast.Expression, ast.BinOp, ast.UnaryOp, ast.USub, ast.UAdd,
                          ast.Load) + tuple(_OPS)):
            continue
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)) \
                and not isinstance(n.value, bool):
            continue
        if isinstance(n, ast.Name) and re.fullmatch(r"_r\d+", n.id):
            continue
        raise ValueError(f"в формуле «{f}» допустимы только числа, + - × /, скобки и ссылки @id")
    return tree, refs


def eval_formula(tree: ast.AST, refs: list[str], values: dict[str, Decimal]) -> Decimal:
    def ev(n):
        if isinstance(n, ast.Expression):
            return ev(n.body)
        if isinstance(n, ast.Constant):
            return Decimal(str(n.value))
        if isinstance(n, ast.Name):
            return values[refs[int(n.id[2:])]]
        if isinstance(n, ast.UnaryOp):
            v = ev(n.operand)
            return -v if isinstance(n.op, ast.USub) else v
        if isinstance(n, ast.BinOp):
            right = ev(n.right)
            if isinstance(n.op, ast.Div) and right == 0:
                raise ValueError("деление на ноль")
            return _OPS[type(n.op)](ev(n.left), right)
        raise ValueError("недопустимое выражение")
    return ev(tree)


def to_excel(formula: str, cell_of: dict[str, str]) -> str:
    """Формула ВОР -> выражение Excel (ссылки заменяются адресами ячеек)."""
    norm = normalize_formula(formula)
    return REF.sub(lambda m: cell_of[m[1]], norm)


def display_formula(formula: str, num_of: dict[str, str]) -> str:
    """Формула для людей: ссылки @id показываются номерами строк, «@бетон» -> «п.2.4»."""
    return REF.sub(lambda m: f"п.{num_of.get(m[1], m[1])}", str(formula or ""))


def row_numbers(data: dict) -> dict[str, str]:
    return {it["_id"]: f"{si}.{ii}" for si, ii, it in rows(data)}


def qround(x: Decimal, places: int) -> Decimal:
    return x.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def load_vor(path: str | Path) -> dict:
    p = Path(path).resolve()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise VorError([f"{p.name}: ошибка JSON в строке {e.lineno}: {e.msg}"])
    data["_input_dir"] = str(p.parent)
    comp = data.get("company")
    if isinstance(comp, str):
        cp = (p.parent / comp).resolve()
        if not cp.exists():
            raise VorError([f"Не найден файл реквизитов компании: {cp}"])
        data["company"] = json.loads(cp.read_text(encoding="utf-8"))
        data["_company_dir"] = str(cp.parent)
    else:
        data["_company_dir"] = str(p.parent)
    data["_brand"] = None
    if data.get("brand"):
        sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "brand" / "scripts"))
        from brand_core import BrandError, load_brand
        try:
            data["_brand"], _ = load_brand((p.parent / data["brand"]).resolve())
        except BrandError as e:
            raise VorError([f"Фирменный шаблон: {x}" for x in e.errors])
    return data


def rows(data: dict):
    """Все строки ВОР по порядку: (номер раздела, номер строки, строка)."""
    for si, s in enumerate(data.get("sections") or [], 1):
        for ii, it in enumerate(s.get("items") or [], 1):
            yield si, ii, it


def compute(data: dict) -> tuple[dict, list[str]]:
    """Проверяет ВОР и считает количества. Возвращает ({id: кол-во}, замечания)."""
    errors: list[str] = []
    warnings: list[str] = []
    places = data.get("qty_decimals", 2)
    if places not in (0, 1, 2, 3):
        errors.append("qty_decimals: 0, 1, 2 или 3 знака после запятой.")
        places = 2

    for k in ("number", "date", "object", "sections"):
        if k not in data:
            errors.append(f"Нет обязательного поля «{k}».")
    if errors:
        raise VorError(errors)
    try:
        date.fromisoformat(str(data["date"]))
    except ValueError:
        errors.append(f"Дата должна быть в формате ГГГГ-ММ-ДД: {data['date']!r}.")
    if not isinstance(data["object"], dict) or not str(data["object"].get("name", "")).strip():
        errors.append("Не указан объект (object.name).")
    docs = data.get("documents") or []
    if not isinstance(docs, list) or any(not isinstance(d, dict) for d in docs):
        errors.append("documents должен быть списком объектов {id, name}.")
        docs = []
    doc_ids = {str(d.get("id", "")).strip() for d in docs}
    if "" in doc_ids:
        errors.append("У каждого документа в documents нужен id (например «КЖ-3»).")
    if not isinstance(data["sections"], list) or not data["sections"]:
        errors.append("Нет ни одного раздела (sections).")
    else:
        for si, s in enumerate(data["sections"], 1):
            if not isinstance(s, dict) or not isinstance(s.get("items"), list) or not s["items"] \
                    or any(not isinstance(it, dict) for it in s["items"]):
                errors.append(f"Раздел {si}: нужен объект с title и непустым списком items.")
            elif not str(s.get("title", "")).strip():
                errors.append(f"Раздел {si}: нет названия.")
    if errors:
        raise VorError(errors)

    # Первый проход: поля строк, ids
    by_id: dict[str, dict] = {}
    parsed: dict[str, tuple] = {}
    for si, ii, it in rows(data):
        rid = str(it.get("id") or f"{si}.{ii}").strip()
        it["_id"] = rid
        where = f"Строка {si}.{ii} «{it.get('name', '?')}»"
        it["_where"] = where
        if rid in by_id:
            errors.append(f"{where}: id «{rid}» повторяется.")
            continue
        by_id[rid] = it
        if not str(it.get("name", "")).strip():
            errors.append(f"{where}: нет наименования работы.")
        unit = normalize_unit(it.get("unit"))
        it["unit"] = unit
        if not unit:
            errors.append(f"{where}: нет единицы измерения.")
        elif unit not in KNOWN_UNITS:
            warnings.append(f"{where}: нестандартная единица «{unit}», проверьте.")
        st = it.setdefault("status", "подтверждено")
        if st not in STATUSES:
            errors.append(f"{where}: статус «{st}» не из списка: {', '.join(STATUSES)}.")
            continue
        has_q = it.get("qty") not in (None, "")
        has_f = str(it.get("formula") or "").strip() != ""
        ref = str(it.get("ref") or "").strip()
        if ref and ref not in doc_ids:
            errors.append(f"{where}: документ «{ref}» не описан в documents.")
        if st == "нет данных":
            if has_q or has_f:
                errors.append(f"{where}: статус «нет данных», но указано количество. "
                              "Либо уберите количество, либо смените статус.")
            if not str(it.get("missing") or "").strip():
                errors.append(f"{where}: укажите в missing, каких данных не хватает.")
            continue
        if not (has_q or has_f):
            errors.append(f"{where}: нет ни formula, ни qty. Если данных нет, "
                          "поставьте статус «нет данных».")
        if st == "подтверждено" and not ref and not str(it.get("source") or "").strip():
            errors.append(f"{where}: статус «подтверждено» требует ref (документ) или source.")
        if st == "допущение" and not str(it.get("assumption") or "").strip():
            errors.append(f"{where}: статус «допущение» требует пояснения в assumption.")
        if has_f:
            try:
                parsed[rid] = parse_formula(it["formula"])
            except ValueError as e:
                errors.append(f"{where}: {e}.")
    if errors:
        raise VorError(errors)

    # Второй проход: значения с учётом ссылок
    values: dict[str, Decimal] = {}
    state: dict[str, str] = {}

    def value(rid: str, chain: list[str]) -> Decimal:
        if rid in values:
            return values[rid]
        if state.get(rid) == "busy":
            raise ValueError("циклическая ссылка: " + " -> ".join(chain + [rid]))
        it = by_id[rid]
        state[rid] = "busy"
        if rid in parsed:
            tree, refs = parsed[rid]
            for r in refs:
                if r not in by_id:
                    raise ValueError(f"ссылка @{r} на несуществующую строку")
                if by_id[r]["status"] == "нет данных":
                    raise ValueError(f"ссылка @{r} на строку без данных")
            dep = {r: value(r, chain + [rid]) for r in refs}
            v = qround(eval_formula(tree, refs, dep), places)
            if it.get("qty") not in (None, ""):
                given = qround(D(it["qty"]), places)
                if given != v:
                    raise ValueError(f"qty {given} не равно результату формулы {v}")
        else:
            v = qround(D(it["qty"]), places)
        values[rid] = v
        state[rid] = "done"
        return v

    for si, ii, it in rows(data):
        if it["status"] == "нет данных":
            continue
        try:
            v = value(it["_id"], [])
            if v <= 0:
                errors.append(f"{it['_where']}: количество {v} должно быть больше нуля.")
            elif abs(v) >= MAX_NUMBER:
                errors.append(f"{it['_where']}: слишком большое количество.")
        except (ValueError, InvalidOperation) as e:
            errors.append(f"{it['_where']}: {e}.")
    if errors:
        raise VorError(errors)

    n_assume = sum(1 for *_, it in rows(data) if it["status"] == "допущение")
    n_missing = sum(1 for *_, it in rows(data) if it["status"] == "нет данных")
    if n_assume:
        warnings.append(f"Строк с допущениями: {n_assume}. Требуют подтверждения.")
    if n_missing:
        warnings.append(f"Строк без данных: {n_missing}. Нужны сведения от заказчика.")
    if data.get("sample"):
        warnings.append("Документ помечен как ОБРАЗЕЦ (sample: true).")
    return values, warnings


def fmt_q(x: Decimal, places: int) -> str:
    return f"{qround(x, places):,.{places}f}".replace(",", " ").replace(".", ",")
