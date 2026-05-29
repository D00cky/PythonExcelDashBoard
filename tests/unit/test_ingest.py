from openpyxl import Workbook

from app.core.ingest import read_workbook


def test_read_workbook_returns_dict_keyed_by_sheet_name(tmp_path):
    wb = Workbook()
    wb.active.title = "First"
    wb.create_sheet("Second")
    path = tmp_path / "wb.xlsx"
    wb.save(path)

    result = read_workbook(path)

    assert set(result.keys()) == {"First", "Second"}
