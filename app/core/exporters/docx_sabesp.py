"""Render the Sabesp Mensal and Semanal reports via docxtpl skeletons.

The skeletons under ``app/core/templates/docx_skeletons/`` are real Sabesp report
templates with anchor strings replaced by Jinja placeholders by
``scripts/build_docx_skeletons.py``. This module loads them, binds extracted data,
and returns the rendered bytes.

- Mensal: scoped to digital surveillance — cover, §1 INTRODUÇÃO, §6.4 with the
  ÍNDICE TECNOLÓGICO POR EQUIPE table (xlsx-derived), §8 ACOMPANHAMENTO – OLHAR
  DIGITAL with dashboard charts embedded as inline images, §9 CONCLUSÃO.
- Semanal: cover-page-only. The structural body is intentionally empty in the
  source — auditors paste per-service charts/screenshots manually. Only polo +
  period bind.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

import pandas as pd
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

from app.core.aggregator import date_bounds, ic_rows_from_inspections, iqs_rows_from_inspections
from app.core.templates.pimentas import PimentasTemplate

_SKELETON_DIR = Path(__file__).resolve().parent.parent / "templates" / "docx_skeletons"
MENSAL_SKELETON_PATH = _SKELETON_DIR / "mensal_sabesp.docx"
SEMANAL_SKELETON_PATH = _SKELETON_DIR / "semanal_sabesp.docx"

# Width applied to every chart image embedded in the Mensal report. 150 mm
# leaves ~3 cm of horizontal margin on A4 portrait, which matches the source
# template's chart sizes when the auditor pastes screenshots manually.
_CHART_WIDTH = Mm(150)

# (chart key, figure-builder name) for the §8.1 KPI block. The keys match the
# {{ name }} placeholders inserted by scripts/build_docx_skeletons.py.
_MENSAL_CHART_BUILDERS: tuple[tuple[str, str], ...] = (
    ("ic_bar", "build_ic_bar"),
    ("iqs_bar", "build_service_iqs_bar"),
    ("photo_conformity", "build_photo_conformity_stacked"),
)

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
    silently extend the contract without updating the skeleton at the same time.

    ``chart_pngs`` holds raw PNG bytes for each {{ chart_name }} placeholder in
    the skeleton (currently ``ic_bar``, ``iqs_bar``, ``photo_conformity`` under
    §8.1). ``render_mensal`` wraps each entry in a ``docxtpl.InlineImage`` at
    render time — the InlineImage needs a live DocxTemplate reference, so we
    keep raw bytes here and bind late.
    """

    polo_label: str
    periodo_inicio: str  # dd/mm/yyyy
    periodo_fim: str  # dd/mm/yyyy
    mes_extenso: str  # e.g. "Abril"
    ano: str  # e.g. "2026"
    indice_tecnologico: tuple[IndiceRow, ...] = ()
    chart_pngs: dict[str, bytes] = field(default_factory=dict)

    @property
    def polo_label_upper(self) -> str:
        return self.polo_label.upper()

    def as_render_dict(self) -> dict[str, object]:
        # Skeleton uses ``indice`` as the loop variable, so the dict key must match.
        # Chart placeholders bind separately in render_mensal so InlineImage gets
        # a live tpl reference; the names go through as-is here for forward compat.
        return {
            "polo_label": self.polo_label,
            "polo_label_upper": self.polo_label_upper,
            "periodo_inicio": self.periodo_inicio,
            "periodo_fim": self.periodo_fim,
            "mes_extenso": self.mes_extenso,
            "ano": self.ano,
            "indice": list(self.indice_tecnologico),
        }


def render_mensal(context: MensalContext, *, skeleton: Path = MENSAL_SKELETON_PATH) -> bytes:
    tpl = DocxTemplate(str(skeleton))
    render_dict = context.as_render_dict()
    for name, png in context.chart_pngs.items():
        # Empty bytes → leave the placeholder unrendered; docxtpl would otherwise
        # raise on a zero-byte image. Skipping lets the doc render cleanly when a
        # chart can't be built (e.g. no data for the period).
        if png:
            render_dict[name] = InlineImage(tpl, BytesIO(png), width=_CHART_WIDTH)
    tpl.render(render_dict)
    buf = BytesIO()
    tpl.save(buf)
    return buf.getvalue()


def render_mensal_chart_pngs(
    template: PimentasTemplate,
    iqs_rows: list,
    ic_rows: list,
) -> dict[str, bytes]:
    """Render the §8.1 dashboard charts as PNG bytes via kaleido.

    Returns a dict matching the placeholders inserted by
    ``scripts/build_docx_skeletons.py:insert_chart_placeholders``. Width/height
    target the source template's chart aspect ratio (≈ 9:4); kaleido scale=1.5
    keeps the rendered images crisp on a printed A4.
    """
    chart_pngs: dict[str, bytes] = {}
    builder_args = {
        "build_ic_bar": (ic_rows,),
        "build_service_iqs_bar": (iqs_rows,),
        "build_photo_conformity_stacked": (iqs_rows,),
    }
    for key, builder_name in _MENSAL_CHART_BUILDERS:
        fig = getattr(template, builder_name)(*builder_args[builder_name])
        chart_pngs[key] = fig.to_image(format="png", width=900, height=400, scale=1.5)
    return chart_pngs


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
    bounds = date_bounds(inspections)
    if bounds is None:
        return "", "", "", ""
    start, end = bounds
    return (
        f"{start:%d/%m/%Y}",
        f"{end:%d/%m/%Y}",
        _MONTH_PT.get(int(end.month), ""),
        f"{end.year}",
    )


def context_from_template(template: PimentasTemplate, path: Path) -> MensalContext:
    """Build a MensalContext from a PimentasTemplate + the uploaded xlsx path."""
    inspections = template.extract_inspections(path)
    periodo_inicio, periodo_fim, mes_extenso, ano = _period_fields(inspections)
    services = sorted(template.SERVICE_SHEETS)
    iqs_rows = iqs_rows_from_inspections(inspections, services)
    ic_rows = ic_rows_from_inspections(inspections, services)
    return MensalContext(
        polo_label=template.polo_name.title(),
        periodo_inicio=periodo_inicio,
        periodo_fim=periodo_fim,
        mes_extenso=mes_extenso,
        ano=ano,
        indice_tecnologico=indice_rows_from_inspections(inspections),
        chart_pngs=render_mensal_chart_pngs(template, iqs_rows, ic_rows),
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
