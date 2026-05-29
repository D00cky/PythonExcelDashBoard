from collections.abc import Iterable

from openpyxl.workbook import Workbook

DATA_SHEET = "DADOS - PIMENTAS"
PERIODO_CELL = "B4"
PERIODO_PREFIX = "Período: "


class SabespPimentasTemplate:
    REQUIRED_SHEETS = frozenset({"CAPA", DATA_SHEET})
    SERVICE_SHEETS = frozenset({"ÁGUA", "ESGOTO", "CAVALETE", "REPOSIÇÃO"})
    MIN_SERVICES = 2

    @classmethod
    def matches(cls, sheet_names: Iterable[str]) -> bool:
        names = {name.strip() for name in sheet_names}
        if not cls.REQUIRED_SHEETS.issubset(names):
            return False
        return len(cls.SERVICE_SHEETS & names) >= cls.MIN_SERVICES

    def extract_periodo(self, workbook: Workbook) -> str | None:
        if DATA_SHEET not in workbook.sheetnames:
            return None
        value = workbook[DATA_SHEET][PERIODO_CELL].value
        if not isinstance(value, str):
            return None
        return value.removeprefix(PERIODO_PREFIX).strip() or None
