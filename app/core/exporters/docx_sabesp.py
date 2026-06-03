"""Render the Sabesp Mensal report via the docxtpl skeleton.

The skeleton at ``app/core/templates/docx_skeletons/mensal_sabesp.docx`` is a real
Sabesp monthly report with anchor strings replaced by Jinja placeholders by
``scripts/build_docx_skeletons.py``. This module loads it, binds extracted data,
and returns the rendered bytes.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

from docxtpl import DocxTemplate

from app.core.templates.pimentas import PimentasTemplate

SKELETON_PATH = (
    Path(__file__).resolve().parent.parent / "templates" / "docx_skeletons" / "mensal_sabesp.docx"
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
class MensalContext:
    """Inputs the Sabesp Mensal skeleton expects bound. Frozen so callers can't
    silently extend the contract without updating the skeleton at the same time."""

    polo_label: str
    periodo_inicio: str  # dd/mm/yyyy
    periodo_fim: str  # dd/mm/yyyy
    mes_extenso: str  # e.g. "Abril"
    ano: str  # e.g. "2026"

    @property
    def polo_label_upper(self) -> str:
        return self.polo_label.upper()

    def as_render_dict(self) -> dict[str, str]:
        return {
            "polo_label": self.polo_label,
            "polo_label_upper": self.polo_label_upper,
            "periodo_inicio": self.periodo_inicio,
            "periodo_fim": self.periodo_fim,
            "mes_extenso": self.mes_extenso,
            "ano": self.ano,
        }


def render_mensal(context: MensalContext, *, skeleton: Path = SKELETON_PATH) -> bytes:
    tpl = DocxTemplate(str(skeleton))
    tpl.render(context.as_render_dict())
    buf = BytesIO()
    tpl.save(buf)
    return buf.getvalue()


def context_from_template(template: PimentasTemplate, path: Path) -> MensalContext:
    """Build a MensalContext from a PimentasTemplate + the uploaded xlsx path.

    Falls back to safe placeholders when the workbook lacks dated inspections (the
    Sabesp report needs *something* in those fields, blanking them would look broken).
    """
    inspections = template.extract_inspections(path)
    dates = inspections["start_date"].dropna() if "start_date" in inspections.columns else None
    if dates is not None and not dates.empty:
        start = dates.min()
        end = dates.max()
        periodo_inicio = f"{start:%d/%m/%Y}"
        periodo_fim = f"{end:%d/%m/%Y}"
        mes_extenso = _MONTH_PT.get(int(end.month), "")
        ano = f"{end.year}"
    else:
        periodo_inicio = ""
        periodo_fim = ""
        mes_extenso = ""
        ano = ""
    return MensalContext(
        polo_label=template.polo_name.title(),
        periodo_inicio=periodo_inicio,
        periodo_fim=periodo_fim,
        mes_extenso=mes_extenso,
        ano=ano,
    )
