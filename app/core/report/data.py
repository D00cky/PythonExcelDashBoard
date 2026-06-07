"""Data selection for scoped reports."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from flask import current_app

from app.core import dashboard_data
from app.core.aggregator import PoloBatch, filter_batch, load_batch, polo_geography
from app.core.report.scope import ReportScope


@dataclass(frozen=True)
class ReportData:
    """Normalized data and metadata for one scoped report."""

    uuid: str
    scope: ReportScope
    scope_name: str | None
    label: str
    period: str
    view: str
    batch: PoloBatch
    scoped_batch: PoloBatch
    inspections: pd.DataFrame


def load_report_data(
    uuid: str,
    *,
    scope: ReportScope,
    scope_name: str | None,
    view: str = "weekly",
    period: str | None = None,
) -> ReportData:
    """Load normalized inspection rows for a report scope.

    The current app's ``instance/uploads/<uuid>/manifest.json`` is the source of
    truth. Rows are read from the parquet cache when warm, with the existing
    live-parse fallback preserved through ``dashboard_data``.
    """
    batch_dir = Path(current_app.instance_path) / "uploads" / uuid
    batch = load_batch(batch_dir)
    view = view if view in ("weekly", "monthly") else "weekly"
    period = period or _default_period(batch, view)
    polos = _polos_for_scope(batch, scope=scope, scope_name=scope_name)
    scoped_batch = filter_batch(batch, polos=polos, view=view, period_key=period)
    inspections = dashboard_data.combined_inspections(uuid, scoped_batch)
    return ReportData(
        uuid=uuid,
        scope=scope,
        scope_name=scope_name,
        label=_label(scope, scope_name, polos),
        period=period,
        view=view,
        batch=batch,
        scoped_batch=scoped_batch,
        inspections=inspections,
    )


def _default_period(batch: PoloBatch, view: str) -> str:
    values = batch.iso_weeks if view == "weekly" else batch.months
    return values[-1] if values else ""


def _polos_for_scope(
    batch: PoloBatch,
    *,
    scope: ReportScope,
    scope_name: str | None,
) -> tuple[str, ...]:
    if scope == ReportScope.CITY:
        return tuple(batch.polos)
    if scope == ReportScope.ZONE:
        return tuple(
            sorted(f.polo for f in batch.files if _same(polo_geography(f)[1], scope_name))
        ) or tuple(batch.polos)
    if scope == ReportScope.MUNICIPALITY:
        return tuple(sorted(f.polo for f in batch.files if _same(f.polo, scope_name))) or tuple(
            batch.polos
        )
    return tuple(batch.polos)


def _same(a: str | None, b: str | None) -> bool:
    return (a or "").strip().casefold() == (b or "").strip().casefold()


def _label(scope: ReportScope, scope_name: str | None, polos: tuple[str, ...]) -> str:
    if scope == ReportScope.CITY:
        return "São Paulo"
    if scope == ReportScope.ZONE:
        return scope_name or "Zona"
    if len(polos) == 1:
        return polos[0].title()
    return scope_name or "Município"
