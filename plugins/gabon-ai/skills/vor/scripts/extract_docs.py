#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["pypdf>=4.0", "openpyxl>=3.1", "python-docx>=1.1"]
# ///
"""Текст из документов заказчика для составления ВОР.

    uv run extract_docs.py input/ --out input/_text

Для каждого PDF, DOCX и XLSX создаётся .txt с метками [стр. N], [лист «…»], [таблица N],
чтобы каждую цифру в ВОР можно было сослаться на место в документе.
Страницы PDF без текста (сканы, чертежи) перечисляются отдельно: их нужно смотреть
как изображения, а данные с них помечать для проверки инженером.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def pdf_text(p: Path) -> tuple[str, list[int]]:
    from pypdf import PdfReader
    out, empty = [], []
    for i, page in enumerate(PdfReader(str(p)).pages, 1):
        t = (page.extract_text() or "").strip()
        if len(t) < 20:
            empty.append(i)
        out.append(f"[стр. {i}]\n{t}\n")
    return "\n".join(out), empty


def docx_text(p: Path) -> tuple[str, list[int]]:
    import docx
    d = docx.Document(str(p))
    out = [para.text for para in d.paragraphs if para.text.strip()]
    for ti, table in enumerate(d.tables, 1):
        out.append(f"\n[таблица {ti}]")
        for row in table.rows:
            out.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(out), []


def xlsx_text(p: Path) -> tuple[str, list[int]]:
    from openpyxl import load_workbook
    wb = load_workbook(p, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"\n[лист «{ws.title}»]")
        for r, row in enumerate(ws.iter_rows(values_only=True), 1):
            vals = ["" if v is None else str(v) for v in row]
            if any(vals):
                out.append(f"{r}: " + " | ".join(vals).rstrip(" |"))
    return "\n".join(out), []


READERS = {".pdf": pdf_text, ".docx": docx_text, ".xlsx": xlsx_text, ".xlsm": xlsx_text}


def main() -> int:
    ap = argparse.ArgumentParser(description="Извлечь текст из документов заказчика")
    ap.add_argument("src", nargs="+", help="файлы или папки")
    ap.add_argument("--out", required=True, help="папка для .txt")
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    files = []
    for s in map(Path, a.src):
        files += sorted(f for f in s.rglob("*") if f.suffix.lower() in READERS) if s.is_dir() \
            else [s]
    files = [f for f in files if out.resolve() not in f.resolve().parents]
    if not files:
        print("Не найдено PDF, DOCX или XLSX.")
        return 1
    status = 0
    for f in files:
        reader = READERS.get(f.suffix.lower())
        if reader is None:
            print(f"[ПРОПУЩЕН] {f}: формат {f.suffix} не поддерживается (нужен PDF, DOCX, XLSX). "
                  "Файлы .doc и .xls пересохраните в новом формате.")
            continue
        try:
            text, empty = reader(f)
        except Exception as e:  # noqa: BLE001
            print(f"[ОШИБКА] {f}: не читается ({e})")
            status = 1
            continue
        t = out / f"{f.name}.txt"
        t.write_text(f"Источник: {f}\n\n{text}", encoding="utf-8")
        print(f"[OK] {f.name} -> {t.name} ({len(text)} символов)")
        if empty:
            print(f"     Страницы без текста (скан или чертёж): {', '.join(map(str, empty))}. "
                  "Смотрите их как изображения, данные отмечайте для проверки.")
    return status


if __name__ == "__main__":
    sys.exit(main())
