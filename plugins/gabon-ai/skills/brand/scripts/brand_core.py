"""Фирменный шаблон компании: цвета, логотип, шрифты, бланк, тексты по умолчанию.

Файл brand.json лежит в папке данных пользователя (рядом с company.json).
Пути внутри brand.json указываются относительно самого brand.json.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

DEFAULT_COLORS = {
    "primary": "#1F3864",      # заголовки, шапка таблицы, линии
    "header_text": "#FFFFFF",  # текст в шапке таблицы
    "section": "#D9E1F2",      # строки разделов
    "subtotal": "#F2F2F2",     # строки «Итого по разделу»
    "total": "#FFF2CC",        # строка «ВСЕГО с НДС»
}
DEFAULT_MARGINS = {"top": 15, "bottom": 18, "left": 15, "right": 15}
DEFAULT_KEYS = ("intro", "terms", "valid_days", "numbering", "amount_decimals",
                "assumptions", "exclusions")
RATE_KEYS = ("overhead_pct", "overhead_source", "profit_pct", "profit_source",
             "vat_pct", "vat_source")
HEX = re.compile(r"^#[0-9A-Fa-f]{6}$")
CYRILLIC_SAMPLE = "АБВЖЩЫЭЮЯабвжщыэюяЁёӘәҒғҚқҢңӨөҰұҮүҺһІі"


class BrandError(Exception):
    def __init__(self, errors: list[str]):
        super().__init__("\n".join(errors))
        self.errors = errors


def _font_missing_glyphs(path: Path) -> str:
    from reportlab.pdfbase.ttfonts import TTFontFile
    face = TTFontFile(str(path))
    return "".join(ch for ch in CYRILLIC_SAMPLE if ord(ch) not in face.charToGlyph)


def load_brand(path: str | Path, check_files: bool = True) -> tuple[dict, list[str]]:
    """Загружает и проверяет brand.json. Возвращает (brand, замечания) или BrandError."""
    p = Path(path).resolve()
    if not p.exists():
        raise BrandError([f"Не найден фирменный шаблон: {p}"])
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise BrandError([f"brand.json: ошибка JSON в строке {e.lineno}: {e.msg}"])
    errors: list[str] = []
    warnings: list[str] = []
    base = p.parent

    colors = dict(DEFAULT_COLORS)
    for k, v in (raw.get("colors") or {}).items():
        if k not in DEFAULT_COLORS:
            warnings.append(f"Неизвестный цвет «{k}» не используется. Доступны: {', '.join(DEFAULT_COLORS)}.")
        elif not isinstance(v, str) or not HEX.match(v):
            errors.append(f"Цвет «{k}» должен быть в виде #RRGGBB, а не {v!r}.")
        else:
            colors[k] = v.upper()

    def file(key, suffixes, label):
        v = raw.get(key)
        if not v:
            return None
        fp = (base / v).resolve()
        if check_files and not fp.exists():
            errors.append(f"{label}: файл не найден: {fp}")
            return None
        if fp.suffix.lower() not in suffixes:
            errors.append(f"{label}: нужен файл {', '.join(suffixes)}, а не {fp.suffix}.")
            return None
        return str(fp)

    logo = file("logo", (".png", ".jpg", ".jpeg"), "Логотип")
    letterhead = file("letterhead_pdf", (".pdf",), "Бланк PDF")

    fonts = raw.get("fonts") or {}
    font_reg = font_bold = None
    if fonts:
        if not (fonts.get("regular") and fonts.get("bold")):
            errors.append("Шрифты: укажите оба файла, fonts.regular и fonts.bold.")
        else:
            for key in ("regular", "bold"):
                fp = (base / fonts[key]).resolve()
                if not fp.exists():
                    errors.append(f"Шрифт {key}: файл не найден: {fp}")
                elif fp.suffix.lower() != ".ttf":
                    errors.append(f"Шрифт {key}: нужен файл .ttf, а не {fp.suffix}.")
                elif check_files:
                    try:
                        miss = _font_missing_glyphs(fp)
                    except Exception as e:  # noqa: BLE001
                        errors.append(f"Шрифт {key}: файл не читается ({e}).")
                        continue
                    if miss:
                        errors.append(f"Шрифт {key}: нет букв «{miss}». Нужен шрифт с кириллицей "
                                      "и казахскими буквами.")
            if not errors:
                font_reg = str((base / fonts["regular"]).resolve())
                font_bold = str((base / fonts["bold"]).resolve())

    if letterhead and check_files:
        from pypdf import PdfReader
        try:
            pg = PdfReader(letterhead).pages[0]
            w, h = float(pg.mediabox.width), float(pg.mediabox.height)
            if abs(w - 595.3) > 3 or abs(h - 841.9) > 3:
                warnings.append(f"Бланк PDF не формата A4 ({w:.0f}×{h:.0f} pt), "
                                "он будет растянут или обрезан.")
        except Exception as e:  # noqa: BLE001
            errors.append(f"Бланк PDF не читается: {e}")

    margins = dict(DEFAULT_MARGINS)
    for k, v in (raw.get("margins_mm") or {}).items():
        if k in margins and isinstance(v, (int, float)) and 5 <= v <= 120:
            margins[k] = v
        else:
            errors.append(f"Поле margins_mm.{k}: число от 5 до 120 мм.")

    pages = raw.get("letterhead_pages", "first")
    if pages not in ("first", "all"):
        errors.append("letterhead_pages: «first» (только первая страница) или «all».")
    pos = raw.get("logo_position", "left")
    if pos not in ("left", "right"):
        errors.append("logo_position: «left» или «right».")
    lw = raw.get("logo_width_mm", 35)
    if not isinstance(lw, (int, float)) or not 10 <= lw <= 90:
        errors.append("logo_width_mm: число от 10 до 90.")

    defaults = raw.get("defaults") or {}
    for k in defaults:
        if k not in DEFAULT_KEYS + ("rates",):
            warnings.append(f"defaults.{k} не используется. Доступны: {', '.join(DEFAULT_KEYS)}, rates.")
    for k in (defaults.get("rates") or {}):
        if k not in RATE_KEYS:
            warnings.append(f"defaults.rates.{k} не используется.")
    dr = defaults.get("rates") or {}
    if "vat_pct" in dr and not str(dr.get("vat_source", "")).strip():
        errors.append("defaults.rates: если задан vat_pct, обязателен vat_source.")

    if errors:
        raise BrandError(errors)

    show_req = raw.get("show_requisites")
    if show_req is None:
        show_req = not letterhead  # на фирменном бланке реквизиты уже напечатаны
    return {
        "path": str(p),
        "name": raw.get("name", ""),
        "colors": colors,
        "logo": logo,
        "logo_position": pos,
        "logo_width_mm": lw,
        "font_regular": font_reg,
        "font_bold": font_bold,
        "letterhead_pdf": letterhead,
        "letterhead_pages": pages,
        "margins_mm": margins,
        "show_requisites": bool(show_req),
        "title": raw.get("title") or "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ",
        "footer_text": raw.get("footer_text", ""),
        "defaults": defaults,
    }, warnings


def apply_defaults(data: dict, brand: dict) -> list[str]:
    """Дополняет КП текстами и ставками по умолчанию. Значения из kp.json не трогает.
    Возвращает список ключей, взятых из шаблона."""
    used = []
    d = brand.get("defaults") or {}
    for k in DEFAULT_KEYS:
        if k in d and data.get(k) in (None, "", [], {}):
            data[k] = d[k]
            used.append(k)
    rates = data.setdefault("rates", {})
    for k, v in (d.get("rates") or {}).items():
        if k in RATE_KEYS and rates.get(k) in (None, ""):
            rates[k] = v
            used.append(f"rates.{k}")
    return used


def brand_for(data: dict) -> dict:
    """Стиль для документа: из brand.json или стандартный GABON AI."""
    b = data.get("_brand")
    if b:
        return b
    comp_logo = (data.get("company") or {}).get("logo")
    logo = None
    if comp_logo:
        fp = Path(data.get("_company_dir", "."), comp_logo)
        logo = str(fp) if fp.exists() else None
    return {"colors": dict(DEFAULT_COLORS), "logo": logo, "logo_position": "left",
            "logo_width_mm": 35, "font_regular": None, "font_bold": None,
            "letterhead_pdf": None, "letterhead_pages": "first",
            "margins_mm": dict(DEFAULT_MARGINS), "show_requisites": True,
            "title": "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", "footer_text": "", "defaults": {}}
