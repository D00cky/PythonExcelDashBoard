from openpyxl import load_workbook

from app.core.templates.sabesp_pimentas import SabespPimentasTemplate
from tests.fixtures.sabesp_minimal import make_minimal_sabesp


def test_extract_periodo_strips_label_prefix(tmp_path):
    wb = load_workbook(make_minimal_sabesp(tmp_path))

    result = SabespPimentasTemplate().extract_periodo(wb)

    assert result == "01/03/2026 à 31/03/2026"


def test_extract_periodo_returns_none_when_cell_empty(tmp_path):
    wb = load_workbook(make_minimal_sabesp(tmp_path, with_periodo=False))

    assert SabespPimentasTemplate().extract_periodo(wb) is None
