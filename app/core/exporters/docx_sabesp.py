"""Render the Sabesp Mensal and Semanal reports via docxtpl skeletons.

The skeletons under ``app/core/templates/docx_skeletons/`` are real Sabesp report
templates with anchor strings replaced by Jinja placeholders by
``scripts/build_docx_skeletons.py``. This module loads them, binds extracted data,
and returns the rendered bytes.

- Mensal: data-rich. EQUIPE list + ÍNDICE TECNOLÓGICO POR EQUIPE table get bound
  to xlsx-derived data alongside the polo name and period.
- Semanal: cover-page-only. The structural body (1. ACOMPANHAMENTO, 2.x service
  blocks) is intentionally empty in the source — auditors paste their per-service
  charts/screenshots manually. Only polo + period bind.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate

from app.core.templates.pimentas import PimentasTemplate

_SKELETON_DIR = Path(__file__).resolve().parent.parent / "templates" / "docx_skeletons"
MENSAL_SKELETON_PATH = _SKELETON_DIR / "mensal_sabesp.docx"
SEMANAL_SKELETON_PATH = _SKELETON_DIR / "semanal_sabesp.docx"

_MONTH_PT = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}


@dataclass(frozen=True)
class EquipeMember:
    """One row in the auditing-team listing (Sondotécnica auditors, not field crews).

    The xlsx doesn't carry this org chart, so callers either supply it explicitly
    or accept an empty list (loop renders zero entries, ASSISTENTES stays static).
    """

    id: str  # e.g. "I", "III", "IV"
    role: str  # e.g. "Tecnólogo", "Engenheira"
    name: str  # e.g. "Lucas Jeremias"


@dataclass(frozen=True)
class IndiceRow:
    """One row of the ÍNDICE TECNOLÓGICO POR EQUIPE table.

    ``ic_pct_str`` is pre-formatted in pt-BR (decimal comma, two places) because the
    Sabesp standard uses that locale and docxtpl templates can't easily format inline.
    """

    equipe: str
    servico: str
    quantidade: int
    ic_pct_str: str


@dataclass(frozen=True)
class MensalContext:
    """Inputs the Sabesp Mensal skeleton expects bound. Frozen so callers can't
    silently extend the contract without updating the skeleton at the same time."""

    polo_label: str
    periodo_inicio: str  # dd/mm/yyyy
    periodo_fim: str  # dd/mm/yyyy
    mes_extenso: str  # e.g. "Abril"
    ano: str  # e.g. "2026"
    equipe: tuple[EquipeMember, ...] = ()
    indice_tecnologico: tuple[IndiceRow, ...] = ()

    @property
    def polo_label_upper(self) -> str:
        return self.polo_label.upper()

    def as_render_dict(self) -> dict[str, object]:
        # Skeleton uses ``indice`` as the loop variable, so the dict key must match.
        return {
            "polo_label": self.polo_label,
            "polo_label_upper": self.polo_label_upper,
            "periodo_inicio": self.periodo_inicio,
            "periodo_fim": self.periodo_fim,
            "mes_extenso": self.mes_extenso,
            "ano": self.ano,
            "equipe": list(self.equipe),
            "indice": list(self.indice_tecnologico),
        }


def render_mensal(context: MensalContext, *, skeleton: Path = MENSAL_SKELETON_PATH) -> bytes:
    tpl = DocxTemplate(str(skeleton))
    tpl.render(context.as_render_dict())
    buf = BytesIO()
    tpl.save(buf)
    return buf.getvalue()


@dataclass(frozen=True)
class SemanalContext:
    """Inputs the Sabesp Semanal skeleton expects bound. The semanal source only
    has cover-page fields that vary per report, so the binding surface is small —
    polo name + period — but ``polo_label_upper`` is derived rather than passed
    so callers don't have to think about Brazilian-Portuguese case rules.
    """

    polo_label: str
    periodo_inicio: str  # dd/mm/yyyy
    periodo_fim: str  # dd/mm/yyyy

    @property
    def polo_label_upper(self) -> str:
        return self.polo_label.upper()

    def as_render_dict(self) -> dict[str, str]:
        return {
            "polo_label": self.polo_label,
            "polo_label_upper": self.polo_label_upper,
            "periodo_inicio": self.periodo_inicio,
            "periodo_fim": self.periodo_fim,
        }


def render_semanal(context: SemanalContext, *, skeleton: Path = SEMANAL_SKELETON_PATH) -> bytes:
    tpl = DocxTemplate(str(skeleton))
    tpl.render(context.as_render_dict())
    buf = BytesIO()
    tpl.save(buf)
    return buf.getvalue()


def indice_rows_from_inspections(inspections: pd.DataFrame) -> tuple[IndiceRow, ...]:
    """Aggregate inspections into (team × service) IC% rows for the tecnológico table.

    A row is dropped when neither conforme nor não-conforme inspections exist for that
    (team, service) — those are zero-quantity entries that would clutter the table.
    """
    required = {"team", "service", "conforme_count", "nao_conforme_count"}
    if inspections.empty or not required.issubset(inspections.columns):
        return ()
    grouped = inspections.groupby(["team", "service"], dropna=True).agg(
        conforme=("conforme_count", "sum"),
        nao_conforme=("nao_conforme_count", "sum"),
    )
    rows: list[IndiceRow] = []
    for (team, service), agg in grouped.iterrows():
        conforme = int(agg["conforme"])
        nao_conforme = int(agg["nao_conforme"])
        total = conforme + nao_conforme
        if total == 0:
            continue
        ic_pct = (conforme / total) * 100
        rows.append(
            IndiceRow(
                equipe=str(team).title() if isinstance(team, str) else str(team),
                servico=str(service).title() if isinstance(service, str) else str(service),
                quantidade=total,
                ic_pct_str=f"{ic_pct:.2f}".replace(".", ","),
            )
        )
    return tuple(rows)


def _period_fields(inspections: pd.DataFrame) -> tuple[str, str, str, str]:
    """Return ``(periodo_inicio, periodo_fim, mes_extenso, ano)`` derived from the
    inspections' ``start_date`` column. Empty strings when no dated rows exist —
    the Sabesp templates expect *something* in those fields, blanking them looks
    broken to a reviewer, but a clearly-empty value is still better than 1970/1/1.
    """
    if "start_date" not in inspections.columns:
        return "", "", "", ""
    dates = inspections["start_date"].dropna()
    if dates.empty:
        return "", "", "", ""
    start = dates.min()
    end = dates.max()
    return (
        f"{start:%d/%m/%Y}",
        f"{end:%d/%m/%Y}",
        _MONTH_PT.get(int(end.month), ""),
        f"{end.year}",
    )


def context_from_template(
    template: PimentasTemplate,
    path: Path,
    *,
    equipe: tuple[EquipeMember, ...] = (),
) -> MensalContext:
    """Build a MensalContext from a PimentasTemplate + the uploaded xlsx path.

    ``equipe`` defaults to empty because the xlsx doesn't carry the audit-team org
    chart — callers can override it.
    """
    inspections = template.extract_inspections(path)
    periodo_inicio, periodo_fim, mes_extenso, ano = _period_fields(inspections)
    return MensalContext(
        polo_label=template.polo_name.title(),
        periodo_inicio=periodo_inicio,
        periodo_fim=periodo_fim,
        mes_extenso=mes_extenso,
        ano=ano,
        equipe=equipe,
        indice_tecnologico=indice_rows_from_inspections(inspections),
    )


def semanal_context_from_template(template: PimentasTemplate, path: Path) -> SemanalContext:
    """Build a SemanalContext from a PimentasTemplate + the uploaded xlsx path."""
    inspections = template.extract_inspections(path)
    periodo_inicio, periodo_fim, _mes_extenso, _ano = _period_fields(inspections)
    return SemanalContext(
        polo_label=template.polo_name.title(),
        periodo_inicio=periodo_inicio,
        periodo_fim=periodo_fim,
    )
