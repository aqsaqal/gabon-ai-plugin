import json
import os
import shutil
import time

import pytest

from conftest import ROOT
from review_project import BLOCK, review

PROJECT = ROOT.parents[1] / "projects" / "test-fundament"
pytestmark = pytest.mark.skipif(not (PROJECT / "output").exists(), reason="нет собранного тестового проекта")


@pytest.fixture
def proj(tmp_path):
    dst = tmp_path / "test-fundament"
    shutil.copytree(PROJECT, dst, copy_function=shutil.copy2)   # с сохранением времени файлов
    return dst


def blocks(rep):
    return [f"{w}: {t}" for lvl, w, t in rep.items if lvl == BLOCK]


def _edit(path, fn):
    d = json.loads(path.read_text(encoding="utf-8"))
    fn(d)
    time.sleep(0.01)
    path.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def test_clean_project_has_no_blockers(proj):
    rep, docs = review(proj)
    assert blocks(rep) == []
    kinds = sorted(k for k, _ in docs.values())
    assert kinds.count("act") == 2 and "vor" in kinds and "grafik" in kinds and "contract" in kinds


def test_vor_changed_after_kp(proj):
    _edit(proj / "input" / "vor.json",
          lambda d: d["sections"][1]["items"][3].__setitem__("formula", "(2*(24+36)+24) × 0,6 × 1,25"))
    b = " ".join(blocks(rep := review(proj)[0]))
    assert "в ВОР сейчас 108.00" in b           # КП расходится с новым ВОР
    assert "не пересобран" in b                  # выходные файлы ВОР, сметы и графика устарели


def test_test_data_without_sample_mark(proj):
    _edit(proj / "input" / "kp.json", lambda d: d.pop("sample"))
    b = " ".join(blocks(review(proj)[0]))
    assert "тестовые данные" in b


def test_image_in_act_pdf(proj):
    from reportlab.pdfgen import canvas
    pdf = proj / "output" / "AKT_ТЕСТ-Д-15_2026_1.pdf"
    c = canvas.Canvas(str(pdf))
    c.drawImage(str(proj / "data" / "logo.png"), 100, 600, 100, 40)
    c.save()
    b = " ".join(blocks(review(proj)[0]))
    assert "изображений" in b


def test_later_act_must_count_all_previous(proj):
    a3 = {"sample": True, "contract": "contract.json", "number": "3", "date": "2026-12-28",
          "period": {"from": "2026-12-25", "to": "2026-12-28"}, "previous": ["act_2.json"],
          "lines": [{"n": "1.1", "qty": 1, "evidence": "тест", "overrun_reason": "ДС № 1 (тест)"}]}
    (proj / "input" / "act_3.json").write_text(json.dumps(a3, ensure_ascii=False), encoding="utf-8")
    b = " ".join(blocks(review(proj)[0]))
    assert "act_1.json" in b


def test_missing_outputs(proj):
    os.remove(proj / "output" / "GRAFIK_ТЕСТ-ГР-001.pdf")
    assert any("Нет выходных файлов" in x for x in blocks(review(proj)[0]))
