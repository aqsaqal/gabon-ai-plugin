import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "skills" / "kp" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "prices" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "brand" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "vor" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "smeta" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "grafik" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "akty" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "proverka" / "scripts"))

HEADER = ["Код", "Наименование", "Ед. изм.", "Цена без НДС, тг", "Вид затрат", "Источник",
          "Дата цены", "Примечание"]


@pytest.fixture
def make_prices(tmp_path):
    """Создаёт прайс-лист из списка строк."""
    from openpyxl import Workbook

    def _make(rows, header=HEADER, name="prices.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Прайс"
        ws.append(header)
        for r in rows:
            ws.append(list(r))
        path = tmp_path / name
        wb.save(path)
        return path
    return _make


@pytest.fixture
def good_rows():
    return [
        ("ЗР-001", "Разработка грунта экскаватором", "м³", 1850, "механизм", "Прайс ТОО А № 5", date(2026, 10, 1), ""),
        ("МН-001", "Бетон В25, поставка", "м3", "34 500,50", "материал", "КП поставщика № 12", "15.10.2026", ""),
        ("БЛ-001", "Бортовой камень 300х150", "п.м", 7830, "комплекс", "Решение руководства", date(2026, 5, 1), ""),
    ]
