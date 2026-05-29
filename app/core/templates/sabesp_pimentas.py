from collections.abc import Iterable
from dataclasses import dataclass

from openpyxl.workbook import Workbook

DATA_SHEET = "DADOS - PIMENTAS"
PERIODO_CELL = "B4"
PERIODO_PREFIX = "Período: "

IQS_SERVICE_ROWS = {
    "Água": 26,
    "Esgoto": 27,
    "Cavalete": 28,
    "Reposição": 29,
}
IQS_TOTAL_ROW = 30
IQS_OVERALL_CELL = f"G{IQS_TOTAL_ROW}"


@dataclass(frozen=True)
class ServiceIQS:
    name: str
    fotos_avaliadas: int
    fotos_nc: int
    fotos_conforme: int
    nc_pct: float
    conforme_pct: float


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

    def extract_iqs_by_service(self, workbook: Workbook) -> list[ServiceIQS]:
        if DATA_SHEET not in workbook.sheetnames:
            return []
        ws = workbook[DATA_SHEET]
        rows: list[ServiceIQS] = []
        for name, row in IQS_SERVICE_ROWS.items():
            if ws[f"B{row}"].value != name:
                continue
            rows.append(
                ServiceIQS(
                    name=name,
                    fotos_avaliadas=ws[f"C{row}"].value or 0,
                    fotos_nc=ws[f"D{row}"].value or 0,
                    fotos_conforme=ws[f"E{row}"].value or 0,
                    nc_pct=ws[f"F{row}"].value or 0.0,
                    conforme_pct=ws[f"G{row}"].value or 0.0,
                )
            )
        return rows

    def extract_iqs_overall(self, workbook: Workbook) -> float | None:
        if DATA_SHEET not in workbook.sheetnames:
            return None
        value = workbook[DATA_SHEET][IQS_OVERALL_CELL].value
        return value if isinstance(value, int | float) else None
