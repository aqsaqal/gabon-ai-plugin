#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl==3.1.5", "reportlab==5.1.0", "pypdf==6.21.4", "pillow==13.0.0", "python-docx==1.2.0"]
# ///
"""Проверка всех документов проекта GABON AI перед выдачей.

    uv run review_project.py projects/<объект> [--report output/REVIEW.md]

Находит исходные данные в input/ (КП, ВОР, смета, график, договор, акты), для каждого:
- проверяет исходные данные и запускает независимый проверяющий скрипт навыка;
- проверяет, что выходные файлы есть и пересобраны после последнего изменения данных;
- ищет тестовые данные в документах без пометки «образец»;
- ищет изображения в актах (подписи и печати вставлять нельзя);
и сверяет документы между собой (ВОР ↔ КП, КП ↔ договор, акты ↔ договор).
Код выхода 1, если есть замечания, блокирующие выдачу.
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[2]
for d in ("kp", "prices", "brand", "vor", "smeta", "grafik", "akty"):
    sys.path.insert(0, str(SKILLS / d / "scripts"))
from build_kp import safe_name  # noqa: E402

BLOCK, WARN, INFO = "БЛОКИРУЕТ ВЫДАЧУ", "ЗАМЕЧАНИЕ", "К СВЕДЕНИЮ"
TEST_MARK = re.compile(r"\bтест", re.IGNORECASE)


class Report:
    def __init__(self):
        self.items: list[tuple[str, str, str]] = []

    def add(self, level, where, text):
        self.items.append((level, where, text))

    def count(self, level):
        return sum(1 for x in self.items if x[0] == level)


def classify(d: dict) -> str | None:
    if "tasks" in d:
        return "grafik"
    if "lines" in d and "contract" in d:
        return "act"
    if "positions" in d and "customer" in d:
        return "contract"
    if "positions" in d and "source_file" in d:
        return "contractor"
    secs = d.get("sections")
    if isinstance(secs, list) and secs and isinstance(secs[0], dict):
        if "positions" in secs[0]:
            return "smeta"
        if "items" in secs[0]:
            return "kp" if "rates" in d else "vor"
    return None


def deps(path: Path, d: dict) -> list[Path]:
    out = []
    for k in ("company", "brand", "price_list", "vor", "smeta", "cards", "contract"):
        if isinstance(d.get(k), str):
            out.append((path.parent / d[k]).resolve())
    for x in d.get("previous") or []:
        out.append((path.parent / x).resolve())
    return [p for p in out if p.exists()]


def outputs(kind: str, d: dict, out_dir: Path) -> tuple[list[Path], str, list[str]]:
    """Ожидаемые файлы, проверяющий скрипт и его аргументы-файлы."""
    n = safe_name(str(d.get("number", "")))
    if kind == "kp":
        b = out_dir / f"KP_{n}"
        return [b.with_suffix(".xlsx"), b.with_suffix(".pdf"), b.with_suffix(".docx")], \
            "kp/scripts/check_kp.py", []
    if kind == "vor":
        b = out_dir / f"VOR_{n}"
        return [b.with_suffix(".xlsx"), b.with_suffix(".pdf")], "vor/scripts/check_vor.py", []
    if kind == "smeta":
        b = out_dir / f"SMETA_{n}"
        return [b.with_suffix(".xlsx"), b.with_suffix(".pdf")], "smeta/scripts/check_smeta.py", []
    if kind == "grafik":
        b = out_dir / f"GRAFIK_{n}"
        return [b.with_suffix(".xlsx"), b.with_suffix(".pdf")], "grafik/scripts/check_grafik.py", []
    if kind == "act":
        return [], "akty/scripts/check_act.py", []
    return [], "", []


def text_of(obj) -> str:
    return json.dumps(obj, ensure_ascii=False)


def review(project: Path) -> tuple[Report, dict]:
    rep = Report()
    inp, out_dir = project / "input", project / "output"
    if not inp.is_dir():
        rep.add(BLOCK, str(project), "Нет папки input/ с исходными данными.")
        return rep, {}
    docs = {}
    for f in sorted(inp.glob("*.json")):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            rep.add(BLOCK, f.name, f"Ошибка JSON в строке {e.lineno}: {e.msg}.")
            continue
        kind = classify(d)
        if kind is None:
            rep.add(INFO, f.name, "Файл не распознан как документ GABON AI, пропущен.")
            continue
        docs[f] = (kind, d)

    contracts = {}
    for f, (kind, d) in docs.items():
        where = f"{f.name} ({ {'kp': 'КП', 'vor': 'ВОР', 'smeta': 'смета', 'grafik': 'график', 'contract': 'договор', 'act': 'акт', 'contractor': 'смета подрядчика'}[kind]} № {d.get('number', '—')})"
        if kind == "contractor":
            rep.add(INFO, f.name, "Импорт сметы подрядчика: проверяется сравнением (навык smeta).")
            continue
        sample = bool(d.get("sample"))
        if kind == "act":
            try:
                c = json.loads((f.parent / d["contract"]).read_text(encoding="utf-8"))
                sample = sample or bool(c.get("sample"))
            except (OSError, KeyError, json.JSONDecodeError):
                pass
        if sample:
            rep.add(INFO, where, "Помечен как ОБРАЗЕЦ: не для отправки заказчику.")
        elif TEST_MARK.search(text_of(d)):
            rep.add(BLOCK, where, "В документе без пометки «образец» есть тестовые данные "
                                  "(слово «тест» в источниках, реквизитах или подтверждениях).")

        # загрузка и проверка исходных данных средствами навыка
        try:
            if kind == "kp":
                from kp_core import load_input, validate
                data = load_input(f)
                for w in validate(data):
                    if "ОБРАЗЕЦ" not in w:
                        rep.add(WARN, where, w)
                if re.search(r"подтверд|уточн|проверить", str(d.get("rates", {}).get("vat_source", "")), re.I):
                    rep.add(WARN, where, "Ставка НДС помечена как неподтверждённая: подтвердите у бухгалтера.")
                for m in d.get("missing_data") or []:
                    rep.add(WARN, where, f"Недостающие данные: {m}")
            elif kind == "vor":
                from vor_core import compute as vc, load_vor, rows
                data = load_vor(f)
                vals, ws = vc(data)
                for w in ws:
                    if "ОБРАЗЕЦ" not in w:
                        rep.add(WARN, where, w)
                d["_values"] = vals
                d["_missing"] = [it["name"] for *_, it in rows(data) if it["status"] == "нет данных"]
            elif kind == "smeta":
                from smeta_core import compute as sc, load_smeta
                for w in sc(load_smeta(f))[1]:
                    if "ОБРАЗЕЦ" not in w:
                        rep.add(WARN, where, w)
            elif kind == "grafik":
                from grafik_core import compute as gc, load_grafik
                res, ws = gc(load_grafik(f))
                for w in ws:
                    if "ОБРАЗЕЦ" not in w:
                        rep.add(WARN, where, w)
                rep.add(INFO, where, f"Срок: {res['start']:%d.%m.%Y}–{res['finish']:%d.%m.%Y}, "
                                     f"{res['workdays']} раб. дн.")
            elif kind == "contract":
                from akty_core import load_contract
                c = load_contract(f)
                contracts[c["number"]] = (f, c)
                if not str(c["customer"].get("bin", "")).strip():
                    rep.add(BLOCK, where, "Нет БИН заказчика: обязателен в акте формы Р-1.")
                for side, label in (("customer", "заказчика"), ("contractor", "исполнителя")):
                    if not str(c[side].get("signer_name", "")).strip():
                        rep.add(WARN, where, f"Не указан подписант {label} (ФИО и должность).")
            elif kind == "act":
                from akty_core import load_act
                a = load_act(f)
                for w in a["_warnings"]:
                    if "ОБРАЗЕЦ" not in w:
                        rep.add(WARN, where, w)
                d["_act"] = a
        except Exception as e:  # noqa: BLE001 — ошибки навыков содержат список errors
            errs = getattr(e, "errors", None) or [str(e)]
            for x in errs:
                rep.add(BLOCK, where, f"Исходные данные: {x}")
            continue

        # выходные файлы
        files, checker, _ = outputs(kind, d, out_dir)
        if kind == "act":
            a = d["_act"]
            b = out_dir / f"AKT_{safe_name(str(a['_contract']['number']))}_{safe_name(str(a['number']))}"
            files = [b.with_suffix(".xlsx"), b.with_suffix(".pdf"), b.with_suffix(".docx")]
        if kind == "contract":
            continue
        missing = [p.name for p in files if not p.exists()]
        if kind in ("act", "kp") and missing == [files[2].name]:
            rep.add(WARN, where, f"Нет Word-версии {'акта' if kind == 'act' else 'КП'}: пересоберите документ.")
            files = files[:2]
        elif missing:
            rep.add(BLOCK, where, f"Нет выходных файлов: {', '.join(missing)}. Соберите документ.")
            continue
        newest_src = max([f.stat().st_mtime] + [p.stat().st_mtime for p in deps(f, d)])
        stale = [p.name for p in files if p.stat().st_mtime < newest_src]
        if stale:
            rep.add(BLOCK, where, f"Файлы {', '.join(stale)} старше исходных данных: документ "
                                  "не пересобран после изменений.")
        r = subprocess.run([sys.executable, str(SKILLS / checker), str(f)] + [str(p) for p in files],
                           capture_output=True, text=True)
        if r.returncode != 0:
            errs = [ln.strip() for ln in r.stdout.splitlines() if "[ОШИБКА]" in ln or ln.strip().startswith("- ")]
            rep.add(BLOCK, where, "Независимая проверка не пройдена: " + ("; ".join(errs[:5]) or r.stderr[-300:]))
        else:
            rep.add(INFO, where, "Независимая проверка пройдена.")
        if kind == "act":
            from pypdf import PdfReader
            imgs = sum(len(p.images) for p in PdfReader(str(files[1])).pages)
            if imgs:
                rep.add(BLOCK, where, f"В PDF акта {imgs} изображений: подписи и печати вставлять нельзя.")
            ev = text_of([ln.get("evidence") for ln in d["lines"]])
            if TEST_MARK.search(ev) and not d["_act"]["_contract"].get("sample") and not d.get("sample"):
                rep.add(BLOCK, where, "Подтверждения объёмов помечены как тестовые.")

    cross_check(docs, contracts, rep)
    return rep, docs


def cross_check(docs, contracts, rep):
    vors = {f.name: d for f, (k, d) in docs.items() if k == "vor" and "_values" in d}
    by_vor_number = {d["number"]: d for d in vors.values()}
    for f, (k, d) in docs.items():
        if k != "kp":
            continue
        where = f"{f.name} (КП № {d.get('number')})"
        refs = [it.get("vor_ref") for s in d.get("sections", []) for it in s.get("items", []) if it.get("vor_ref")]
        for s in d.get("sections", []):
            for it in s.get("items", []):
                m = re.match(r"ВОР № (.+), строка (\d+)\.(\d+)$", str(it.get("vor_ref", "")))
                if not m or m[1] not in by_vor_number:
                    continue
                v = by_vor_number[m[1]]
                try:
                    vit = v["sections"][int(m[2]) - 1]["items"][int(m[3]) - 1]
                except (IndexError, KeyError):
                    rep.add(BLOCK, where, f"Ссылка на {it['vor_ref']}: такой строки в ВОР нет.")
                    continue
                vq = v["_values"].get(vit.get("_id") or vit.get("id") or f"{m[2]}.{m[3]}")
                if vq is not None and Decimal(str(it["qty"])) != vq:
                    rep.add(BLOCK, where, f"«{it['name']}»: в КП {it['qty']}, в ВОР сейчас {vq}. "
                                          "ВОР изменён после КП: пересоздайте КП из ВОР.")
        if refs:
            vnum = re.match(r"ВОР № (.+), строка", refs[0])
            v = by_vor_number.get(vnum[1]) if vnum else None
            if v:
                excl = text_of(d.get("exclusions", []))
                for name in v.get("_missing", []):
                    if name not in excl:
                        rep.add(WARN, where, f"Строка ВОР без данных «{name}» не указана в разделе "
                                             "«Не включено в стоимость».")
    # КП ↔ договор
    kps = {d.get("number"): (f, d) for f, (k, d) in docs.items() if k == "kp"}
    for num, (cf, c) in contracts.items():
        m = re.match(r"КП № (.+) от", str(c.get("basis_kp", "")))
        if m and m[1] in kps:
            from kp_core import compute, load_input, validate
            try:
                kd = load_input(kps[m[1]][0])
                validate(kd)
                kt = compute(kd)["total"]
                if kt != Decimal(c["totals"]["total"]):
                    rep.add(INFO, f"{cf.name} (договор № {num})",
                            f"Итог КП сейчас {kt}, в договоре {c['totals']['total']}: КП менялось после "
                            "заключения договора. Акты считаются по договору.")
            except Exception:  # noqa: BLE001
                pass
    # акты: номера и полнота выполнения
    acts = [(f, d["_act"]) for f, (k, d) in docs.items() if k == "act" and "_act" in d]
    by_contract = {}
    for f, a in acts:
        by_contract.setdefault(a["_contract"]["number"], []).append((f, a))
    for num, lst in by_contract.items():
        lst.sort(key=lambda x: x[1]["_date"])
        nums = [str(a["number"]) for _, a in lst]
        if len(set(nums)) != len(nums):
            rep.add(BLOCK, f"Договор № {num}", f"Повторяются номера актов: {', '.join(nums)}.")
        last_f, last = lst[-1]
        names = {Path(x).name for x in last.get("previous") or []}
        missed = [f.name for f, _ in lst[:-1] if f.name not in names]
        if missed:
            rep.add(BLOCK, f"{last_f.name} (акт № {last['number']})",
                    f"Не учтены предыдущие акты: {', '.join(missed)}. Накопительный итог неверен.")
        c = last["_contract"]
        done = sum((a["_calc"]["total"] for _, a in lst), Decimal(0))
        total = Decimal(c["totals"]["total"])
        rep.add(INFO, f"Договор № {num}", f"Закрыто актами {done} из {total} "
                                         f"({done / total * 100:.1f} %), актов: {len(lst)}.")


def write_report(rep: Report, project: Path, path: Path):
    lines = [f"# Проверка документов проекта {project.name}", "",
             f"Дата проверки: {date.today():%d.%m.%Y}.", ""]
    verdict = ("**Нельзя выдавать** до исправления замечаний первой группы."
               if rep.count(BLOCK) else "**Блокирующих замечаний нет.** Остальное требует решения "
               "ответственного специалиста.")
    lines += [verdict, ""]
    for level in (BLOCK, WARN, INFO):
        items = [x for x in rep.items if x[0] == level]
        if not items:
            continue
        lines += [f"## {level} ({len(items)})", ""]
        for _, where, text in items:
            lines.append(f"- **{where}**: {text}")
        lines.append("")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser(description="Проверка документов проекта")
    ap.add_argument("project")
    ap.add_argument("--report", help="путь к отчёту .md (по умолчанию output/REVIEW_<дата>.md)")
    a = ap.parse_args()
    project = Path(a.project).resolve()
    rep, docs = review(project)
    path = Path(a.report) if a.report else project / "output" / f"REVIEW_{date.today():%Y-%m-%d}.md"
    write_report(rep, project, path)
    print(f"Документов: {len(docs)}. Блокирующих: {rep.count(BLOCK)}, замечаний: {rep.count(WARN)}, "
          f"к сведению: {rep.count(INFO)}.")
    for level in (BLOCK, WARN):
        for _, where, text in [x for x in rep.items if x[0] == level]:
            print(f"[{level}] {where}: {text}")
    print(f"Отчёт: {path}")
    return 1 if rep.count(BLOCK) else 0


if __name__ == "__main__":
    sys.exit(main())
