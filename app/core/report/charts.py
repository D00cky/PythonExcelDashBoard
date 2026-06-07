"""Render Plotly figures to cached PNG images for reports."""

from __future__ import annotations

import base64
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from flask import current_app

from app.core.report.data import load_report_data
from app.core.report.scope import ReportScope


def render_charts_for_scope(
    uuid: str,
    scope: ReportScope,
    scope_name: str | None,
    theme: str,
    output_dir: Path,
    *,
    period: str | None = None,
    view: str = "weekly",
) -> dict[str, Path]:
    """Render report charts for a scope and return their PNG paths.

    Images are cached under the caller-supplied directory. If Kaleido/Chrome is
    unavailable, a small valid PNG placeholder is written so DOCX/PDF/PPTX
    generation degrades gracefully instead of failing the report.
    """
    data = load_report_data(uuid, scope=scope, scope_name=scope_name, view=view, period=period)
    output_dir.mkdir(parents=True, exist_ok=True)
    template = "polo_light" if theme == "light" else "polo_dark"
    figures = _figures(data.inspections, scope=scope, template=template)
    paths: dict[str, Path] = {}
    for name, fig in figures.items():
        path = output_dir / f"{name}.png"
        _write_png(fig, path)
        paths[name] = path
    return paths


def _figures(
    inspections: pd.DataFrame,
    *,
    scope: ReportScope,
    template: str,
) -> dict[str, go.Figure]:
    group_col = "zone" if scope == ReportScope.CITY else "polo"
    if scope == ReportScope.MUNICIPALITY:
        group_col = "service"
    grouped = _grouped(inspections, group_col)
    return {
        "chart_ic": _bar(grouped, group_col, "ic", "IC", template),
        "chart_iqs": _bar(grouped, group_col, "iqs", "IQS", template),
        "chart_volume": _bar(grouped, group_col, "inspecoes", "Inspeções", template),
    }


def _grouped(inspections: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if inspections.empty or group_col not in inspections.columns:
        return pd.DataFrame([{"label": "Sem dados", "ic": 0, "iqs": 0, "inspecoes": 0, "fotos": 0}])
    grouped = (
        inspections.groupby(group_col, dropna=False)
        .agg(
            inspecoes=("conforme_count", "size"),
            conforme=("conforme_count", "sum"),
            fotos=("photo_total", "sum"),
            fotos_conforme=("photo_conforme", "sum"),
        )
        .reset_index()
        .rename(columns={group_col: "label"})
    )
    grouped["label"] = grouped["label"].fillna("Sem classificação").astype(str)
    grouped["ic"] = grouped["conforme"] / grouped["inspecoes"]
    grouped["iqs"] = grouped["fotos_conforme"] / grouped["fotos"].where(grouped["fotos"] > 0)
    grouped["iqs"] = grouped["iqs"].fillna(0)
    return grouped.sort_values("iqs", ascending=True).tail(20)


def _bar(
    grouped: pd.DataFrame, group_col: str, metric: str, title: str, template: str
) -> go.Figure:
    values = grouped[metric].tolist()
    labels = grouped["label"].tolist()
    is_pct = metric in {"ic", "iqs"}
    return go.Figure(
        data=[
            go.Bar(
                x=values,
                y=labels,
                orientation="h",
                marker_color="#d45a0a" if metric == "ic" else "#16a34a",
                text=[f"{v:.1%}" if is_pct else str(int(v)) for v in values],
                textposition="outside",
            )
        ],
        layout=go.Layout(
            title=f"{title} por {group_col}",
            xaxis={"tickformat": ".0%" if is_pct else None},
            yaxis={"automargin": True},
            template=template,
            height=max(360, 34 * len(labels) + 120),
        ),
    )


def _write_png(fig: go.Figure, path: Path) -> None:
    if current_app and current_app.testing:
        path.write_bytes(_PLACEHOLDER_PNG)
        return
    try:
        fig.write_image(path, width=1000, height=500, scale=2, engine="kaleido")
    except Exception:  # noqa: BLE001 - report generation must degrade gracefully
        path.write_bytes(_PLACEHOLDER_PNG)


_PLACEHOLDER_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
)
