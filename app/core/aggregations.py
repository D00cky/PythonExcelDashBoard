"""Multi-level aggregation engine for Polo *batch* audit data.

Aggregation levels mirror the dashboard drill-down:

    city          → all Polo files combined
    zone          → the Polo files in one zone (see app.core.geography)
    municipality  → a single Polo file

Every level is computed the same way: from the **concatenated raw inspection
rows** of that scope. Summing 0/1 conformity indicators and photo counts over
rows makes a parent level the inspection-count-weighted aggregate of its
children by construction — no separate "average of averages" step, and no
double-counting. The per-service IC/IQS math itself is reused from
:mod:`app.core.aggregator` so the cached numbers match the live dashboard.

Public API:
    SERVICES: tuple[str, ...]
    compute_scope(inspections) -> ScopeAgg        # pure
    scope_key_city() -> str
    scope_key_zone(zone) -> str
    scope_key_muni(polo) -> str
    save_scope(uuid, scope_key, scope) -> None
    load_scope(uuid, scope_key) -> ScopeAgg
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from app.core import cache
from app.core.aggregator import (
    ic_rows_from_inspections,
    iqs_overall_from_inspections,
    iqs_rows_from_inspections,
)
from app.core.templates.pimentas import PimentasTemplate

#: Canonical service names, matching the ``service`` column of inspection rows.
SERVICES: tuple[str, ...] = tuple(sorted(PimentasTemplate.SERVICE_SHEETS))

_IC_COLS = ["service", "ic_pct", "lvs"]
_IQS_COLS = ["service", "fotos_avaliadas", "fotos_nc", "fotos_conforme", "nc_pct", "conforme_pct"]
_TEAM_COLS = ["team", "inspecoes", "nao_conforme", "fail_pct"]
_SUMMARY_COLS = ["iqs_overall", "total_photos", "total_inspections"]
_METRICS = ("ic", "iqs", "teams", "summary")


@dataclass(frozen=True)
class ScopeAgg:
    """Pre-computed summary frames for one scope (city / zone / municipality)."""

    ic: pd.DataFrame
    iqs: pd.DataFrame
    teams: pd.DataFrame
    summary: pd.DataFrame


def compute_scope(inspections: pd.DataFrame) -> ScopeAgg:
    """Build all summary frames for one scope from its raw inspection rows.

    ``inspections`` is the concatenation of the relevant Polo files' normalized
    rows (see :func:`app.core.templates.pimentas.PimentasTemplate.extract_inspections`).
    An empty input yields empty frames with the correct schema and a zeroed
    summary row, so callers never special-case the no-data case.
    """
    services = list(SERVICES)
    ic = pd.DataFrame(
        [(r.name, r.ic_pct, r.lvs) for r in ic_rows_from_inspections(inspections, services)],
        columns=_IC_COLS,
    )
    iqs = pd.DataFrame(
        [
            (r.name, r.fotos_avaliadas, r.fotos_nc, r.fotos_conforme, r.nc_pct, r.conforme_pct)
            for r in iqs_rows_from_inspections(inspections, services)
        ],
        columns=_IQS_COLS,
    )
    summary = pd.DataFrame(
        [
            {
                "iqs_overall": iqs_overall_from_inspections(inspections),
                "total_photos": int(inspections["photo_total"].sum())
                if not inspections.empty
                else 0,
                "total_inspections": int(len(inspections)),
            }
        ],
        columns=_SUMMARY_COLS,
    )
    return ScopeAgg(ic=ic, iqs=iqs, teams=_team_ranking(inspections), summary=summary)


def _team_ranking(df: pd.DataFrame) -> pd.DataFrame:
    """Per-team non-conformity ranking, worst first."""
    if df.empty:
        return pd.DataFrame(columns=_TEAM_COLS)
    grouped = df.groupby("team", as_index=False).agg(
        inspecoes=("conforme_count", "size"),
        nao_conforme=("nao_conforme_count", "sum"),
    )
    grouped["inspecoes"] = grouped["inspecoes"].astype(int)
    grouped["nao_conforme"] = grouped["nao_conforme"].astype(int)
    grouped["fail_pct"] = grouped["nao_conforme"] / grouped["inspecoes"]
    return grouped[_TEAM_COLS].sort_values("fail_pct", ascending=False, ignore_index=True)


def scope_key_city() -> str:
    return "city"


def scope_key_zone(zone: str) -> str:
    return f"zone:{zone}"


def scope_key_muni(polo: str) -> str:
    return f"muni:{polo}"


def save_scope(uuid: str, scope_key: str, scope: ScopeAgg) -> None:
    """Persist a scope's four summary frames as parquet under the cache."""
    for metric in _METRICS:
        cache.save_aggregation(uuid, metric, scope_key, getattr(scope, metric))


def load_scope(uuid: str, scope_key: str) -> ScopeAgg:
    """Load a previously saved scope. Raises ``KeyError`` if not cached."""
    return ScopeAgg(**{m: cache.load_aggregation(uuid, m, scope_key) for m in _METRICS})
