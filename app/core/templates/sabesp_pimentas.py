from collections.abc import Iterable
from dataclasses import dataclass

import plotly.graph_objects as go
from openpyxl.workbook import Workbook

DATA_SHEET_PREFIX = "DADOS - "
COVER_SHEET = "CAPA"
PERIODO_CELL = "B4"
PERIODO_PREFIX = "Período: "

IC_SERVICE_ROWS = {
    "Água": 11,
    "Esgoto": 12,
    "Reposição": 13,
}

IQS_SERVICE_ROWS = {
    "Água": 26,
    "Esgoto": 27,
    "Cavalete": 28,
    "Reposição": 29,
}
IQS_TOTAL_ROW = 30
IQS_OVERALL_CELL = f"G{IQS_TOTAL_ROW}"


@dataclass(frozen=True)
class ServiceIC:
    name: str
    ic_pct: float
    lvs: int


@dataclass(frozen=True)
class ServiceIQS:
    name: str
    fotos_avaliadas: int
    fotos_nc: int
    fotos_conforme: int
    nc_pct: float
    conforme_pct: float


class SabespPimentasTemplate:
    SERVICE_SHEETS = frozenset({"ÁGUA", "ESGOTO", "CAVALETE", "REPOSIÇÃO"})
    MIN_SERVICES = 2

    def __init__(self, data_sheet_name: str = "DADOS - PIMENTAS"):
        self.data_sheet_name = data_sheet_name
        self.polo_name = data_sheet_name.removeprefix(DATA_SHEET_PREFIX).strip()

    @classmethod
    def detect(cls, sheet_names: Iterable[str]) -> "SabespPimentasTemplate | None":
        names = {name.strip() for name in sheet_names}
        if COVER_SHEET not in names:
            return None
        data_sheet = next((n for n in names if n.startswith(DATA_SHEET_PREFIX)), None)
        if data_sheet is None:
            return None
        if len(cls.SERVICE_SHEETS & names) < cls.MIN_SERVICES:
            return None
        return cls(data_sheet_name=data_sheet)

    @classmethod
    def matches(cls, sheet_names: Iterable[str]) -> bool:
        return cls.detect(sheet_names) is not None

    def extract_periodo(self, workbook: Workbook) -> str | None:
        if self.data_sheet_name not in workbook.sheetnames:
            return None
        value = workbook[self.data_sheet_name][PERIODO_CELL].value
        if not isinstance(value, str):
            return None
        return value.removeprefix(PERIODO_PREFIX).strip() or None

    def extract_iqs_by_service(self, workbook: Workbook) -> list[ServiceIQS]:
        if self.data_sheet_name not in workbook.sheetnames:
            return []
        ws = workbook[self.data_sheet_name]
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
        if self.data_sheet_name not in workbook.sheetnames:
            return None
        value = workbook[self.data_sheet_name][IQS_OVERALL_CELL].value
        return value if isinstance(value, int | float) else None

    def extract_ic_by_service(self, workbook: Workbook) -> list[ServiceIC]:
        if self.data_sheet_name not in workbook.sheetnames:
            return []
        ws = workbook[self.data_sheet_name]
        rows: list[ServiceIC] = []
        for name, row in IC_SERVICE_ROWS.items():
            if ws[f"B{row}"].value != name:
                continue
            rows.append(
                ServiceIC(
                    name=name,
                    ic_pct=ws[f"C{row}"].value or 0.0,
                    lvs=ws[f"E{row}"].value or 0,
                )
            )
        return rows

    def build_service_iqs_bar(self, rows: list[ServiceIQS]) -> go.Figure:
        return go.Figure(
            data=[
                go.Bar(
                    x=[r.name for r in rows],
                    y=[r.conforme_pct for r in rows],
                    marker_color="#2a9d8f",
                    text=[f"{r.conforme_pct:.1%}" for r in rows],
                    textposition="outside",
                )
            ],
            layout=go.Layout(
                title="Índice de Qualidade SABESP por Serviço",
                yaxis={"title": "Conforme (%)", "tickformat": ".0%", "range": [0, 1]},
                xaxis={"title": "Serviço"},
                template="plotly_white",
            ),
        )

    def build_ic_bar(self, rows: list[ServiceIC]) -> go.Figure:
        return go.Figure(
            data=[
                go.Bar(
                    x=[r.name for r in rows],
                    y=[r.ic_pct for r in rows],
                    marker_color="#264653",
                    text=[f"{r.ic_pct:.0%}" for r in rows],
                    textposition="outside",
                )
            ],
            layout=go.Layout(
                title="Índice de Conformidade (IC) por Serviço",
                yaxis={"title": "IC (%)", "tickformat": ".0%", "range": [0, 1.1]},
                xaxis={"title": "Serviço"},
                template="plotly_white",
            ),
        )

    def build_photo_conformity_stacked(self, rows: list[ServiceIQS]) -> go.Figure:
        names = [r.name for r in rows]
        return go.Figure(
            data=[
                go.Bar(
                    name="Não Conforme",
                    x=names,
                    y=[r.fotos_nc for r in rows],
                    marker_color="#e76f51",
                ),
                go.Bar(
                    name="Conforme",
                    x=names,
                    y=[r.fotos_conforme for r in rows],
                    marker_color="#2a9d8f",
                ),
            ],
            layout=go.Layout(
                barmode="stack",
                title="Fotos Avaliadas por Serviço",
                yaxis={"title": "Quantidade de fotos"},
                xaxis={"title": "Serviço"},
                template="plotly_white",
            ),
        )
