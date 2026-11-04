#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "reportlab>=4.0", "pypdf>=4.0", "pillow>=10"]
# ///
"""Фирменный шаблон компании.

    uv run brand.py init data/brand.json                 создать brand.json (не перезаписывает)
    uv run brand.py check data/brand.json                проверить цвета, логотип, шрифты, бланк
    uv run brand.py preview data/brand.json --company data/company.json --out папка
                                                         образец КП в фирменном стиле
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from brand_core import BrandError, load_brand  # noqa: E402

KP_SCRIPTS = HERE.parents[1] / "kp" / "scripts"


def cmd_init(out: Path) -> int:
    if out.exists():
        print(f"Файл уже существует, не перезаписываю: {out}")
        return 1
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(HERE.parent / "assets" / "brand.example.json", out)
    print(f"Создан фирменный шаблон: {out.resolve()}")
    print("Заполните цвета, путь к логотипу и тексты по умолчанию, затем выполните check.")
    return 0


def cmd_check(path: Path) -> int:
    print(f"ПРОВЕРКА ФИРМЕННОГО ШАБЛОНА {path}")
    try:
        b, warnings = load_brand(path)
    except BrandError as e:
        for x in e.errors:
            print(f"  [ОШИБКА] {x}")
        print(f"РЕЗУЛЬТАТ: ошибок: {len(e.errors)}")
        return 1
    print(f"  [OK]  Цвета: " + ", ".join(f"{k} {v}" for k, v in b["colors"].items()))
    print(f"  [OK]  Логотип: {b['logo'] or 'нет'}")
    print(f"  [OK]  Шрифты: {b['font_regular'] or 'системные (Arial / DejaVu Sans)'}")
    print(f"  [OK]  Бланк PDF: {b['letterhead_pdf'] or 'нет'}"
          + (f" ({'все страницы' if b['letterhead_pages'] == 'all' else 'первая страница'})"
             if b["letterhead_pdf"] else ""))
    print(f"  [OK]  Реквизиты в шапке: {'да' if b['show_requisites'] else 'нет'}")
    d = b["defaults"]
    print(f"  [OK]  Тексты по умолчанию: {', '.join(d) if d else 'нет'}")
    for w in warnings:
        print(f"  [ЗАМЕЧАНИЕ] {w}")
    print("РЕЗУЛЬТАТ: пройдена")
    return 0


def cmd_preview(path: Path, company: Path, out: Path) -> int:
    if cmd_check(path):
        return 1
    out.mkdir(parents=True, exist_ok=True)
    kp = json.loads((HERE.parent / "assets" / "preview_kp.json").read_text(encoding="utf-8"))
    kp["brand"] = str(path.resolve())
    kp["company"] = str(company.resolve())
    kp_path = out / "preview_kp.json"
    kp_path.write_text(json.dumps(kp, ensure_ascii=False, indent=2), encoding="utf-8")
    r = subprocess.run([sys.executable, str(KP_SCRIPTS / "build_kp.py"), str(kp_path),
                        "--out", str(out)], text=True)
    if r.returncode == 0:
        print("Это образец оформления с условными данными. Откройте PDF и Excel и проверьте вид.")
    return r.returncode


def main() -> int:
    ap = argparse.ArgumentParser(description="Фирменный шаблон компании")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("init").add_argument("out")
    sub.add_parser("check").add_argument("brand")
    pv = sub.add_parser("preview")
    pv.add_argument("brand")
    pv.add_argument("--company", required=True, help="путь к company.json")
    pv.add_argument("--out", required=True, help="папка для образца")
    a = ap.parse_args()
    if a.cmd == "init":
        return cmd_init(Path(a.out))
    if a.cmd == "check":
        return cmd_check(Path(a.brand))
    return cmd_preview(Path(a.brand), Path(a.company), Path(a.out))


if __name__ == "__main__":
    sys.exit(main())
