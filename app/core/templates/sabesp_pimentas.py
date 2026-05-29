from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
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

    def extract_inspections(self, path: Path) -> pd.DataFrame:
        """Read the four service sheets, return one row per inspection.

        Columns: team, tss, service, conforme_count, nao_conforme_count.
        Stage columns (those whose only non-null values are in {C, NC, SF, NA})
        are auto-detected per sheet; conforme_count is the count of 'C' cells
        across stage columns per row, nao_conforme_count the count of 'NC'.
        """
        parts: list[pd.DataFrame] = []
        for service in sorted(self.SERVICE_SHEETS):
            try:
                df = pd.read_excel(path, sheet_name=service, engine="openpyxl")
            except (ValueError, KeyError):
                continue
            df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
            team_col = _ci_column(df, "EQUIPE")
            tss_col = _ci_column(df, "Descrição TSS")
            if team_col is None or tss_col is None:
                continue
            df = df.dropna(subset=[team_col, tss_col])

            stage_cols = _detect_stage_columns(df, exclude={team_col, tss_col})
            if stage_cols:
                stages = df[stage_cols].astype(str).apply(lambda s: s.str.strip())
                df["conforme_count"] = (stages == "C").sum(axis=1)
                df["nao_conforme_count"] = (stages == "NC").sum(axis=1)
            else:
                df["conforme_count"] = 0
                df["nao_conforme_count"] = 0

            df = df.rename(columns={team_col: "team", tss_col: "tss"})
            df = df[["team", "tss", "conforme_count", "nao_conforme_count"]].copy()
            df["service"] = service
            parts.append(df)
        if not parts:
            return pd.DataFrame(
                columns=["team", "tss", "service", "conforme_count", "nao_conforme_count"]
            )
        return pd.concat(parts, ignore_index=True)

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

    def build_tss_distribution(self, df: pd.DataFrame, top_n: int = 15) -> go.Figure:
        if df.empty:
            return _empty_figure("Sem inspeções registradas")

        counts = df.groupby("tss").size().sort_values(ascending=False).head(top_n)
        labels = list(reversed(counts.index))
        values = list(reversed(counts.values))

        return go.Figure(
            data=[
                go.Bar(
                    x=values,
                    y=labels,
                    orientation="h",
                    marker_color="#457b9d",
                    text=values,
                    textposition="outside",
                )
            ],
            layout=go.Layout(
                title=f"Tipos de Serviço (TSS) — Top {len(labels)}",
                xaxis={"title": "Quantidade de inspeções"},
                yaxis={"title": "Descrição TSS", "automargin": True},
                template="plotly_white",
                height=max(400, 30 * len(labels) + 100),
                margin={"l": 280},
            ),
        )

    def build_team_service_stacked(self, df: pd.DataFrame, top_n: int = 15) -> go.Figure:
        if df.empty:
            return _empty_figure("Sem inspeções registradas")

        team_totals = df.groupby("team").size().sort_values(ascending=False).head(top_n)
        teams = list(reversed(team_totals.index))  # plotly y starts at bottom

        pivot = (
            df[df["team"].isin(teams)]
            .groupby(["team", "service"])
            .size()
            .unstack(fill_value=0)
            .reindex(teams)
        )

        traces = []
        for service in sorted(self.SERVICE_SHEETS):
            if service not in pivot.columns:
                continue
            traces.append(
                go.Bar(
                    name=service,
                    y=teams,
                    x=pivot[service].tolist(),
                    orientation="h",
                    marker_color=_SERVICE_COLORS.get(service, "#888"),
                )
            )

        return go.Figure(
            data=traces,
            layout=go.Layout(
                barmode="stack",
                title=f"Inspeções por Equipe (Top {len(teams)})",
                xaxis={"title": "Quantidade de inspeções"},
                yaxis={"title": "Equipe", "automargin": True},
                template="plotly_white",
                height=max(400, 30 * len(teams) + 100),
                legend={"orientation": "h", "yanchor": "bottom", "y": 1.02},
            ),
        )


_SERVICE_COLORS = {
    "ÁGUA": "#2a9d8f",
    "ESGOTO": "#264653",
    "CAVALETE": "#e9c46a",
    "REPOSIÇÃO": "#e76f51",
}


def _empty_figure(message: str) -> go.Figure:
    return go.Figure(
        layout=go.Layout(
            template="plotly_white",
            annotations=[
                {
                    "text": message,
                    "showarrow": False,
                    "xref": "paper",
                    "yref": "paper",
                    "x": 0.5,
                    "y": 0.5,
                    "font": {"size": 14, "color": "#888"},
                }
            ],
        )
    )


def _ci_column(df: pd.DataFrame, target: str) -> str | None:
    target_lower = target.casefold()
    for col in df.columns:
        if isinstance(col, str) and col.casefold() == target_lower:
            return col
    return None


_STAGE_CODE_VALUES = frozenset({"C", "NC", "SF", "NA"})


def _detect_stage_columns(df: pd.DataFrame, exclude: set) -> list[str]:
    """Return columns whose every non-null value is one of {C, NC, SF, NA}."""
    cols: list[str] = []
    for col in df.columns:
        if col in exclude:
            continue
        non_null = df[col].dropna()
        if len(non_null) == 0:
            continue
        as_str = non_null.astype(str).str.strip()
        if as_str.isin(_STAGE_CODE_VALUES).all():
            cols.append(col)
    return cols
