"""Build scoped PDF reports.

The production plan prefers LibreOffice conversion from DOCX. This builder uses
ReportLab directly as the reliable fallback available in the current Render
configuration.
"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from app.core.report.charts import render_charts_for_scope
from app.core.report.data import load_report_data
from app.core.report.scope import ReportScope


def build_pdf(
    uuid: str,
    period: str,
    scope: ReportScope,
    scope_name: str | None,
    output_path: Path,
) -> Path:
    """Build a scoped PDF report and return ``output_path``."""
    data = load_report_data(uuid, scope=scope, scope_name=scope_name, period=period)
    charts = render_charts_for_scope(
        uuid,
        scope,
        scope_name,
        "light",
        output_path.parent / "charts" / "light",
        period=period,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(output_path), pagesize=A4)
    styles = getSampleStyleSheet()
    story = [
        Paragraph(_title(scope, data.label), styles["Title"]),
        Paragraph(f"Período: {data.period}", styles["Normal"]),
        Paragraph(f"Inspeções: {len(data.inspections)}", styles["Normal"]),
        Spacer(1, 10),
        _summary_table(data.inspections),
        Spacer(1, 10),
    ]
    for path in charts.values():
        story.append(Image(str(path), width=16 * cm, height=8 * cm))
        story.append(Spacer(1, 8))
    doc.build(story)
    return output_path


def _summary_table(inspections) -> Table:
    rows = [["Grupo", "Inspeções", "Conforme", "Não Conforme"]]
    group_col = "polo" if "polo" in inspections.columns else "service"
    if not inspections.empty:
        grouped = inspections.groupby(group_col).agg(
            inspecoes=("conforme_count", "size"),
            conforme=("conforme_count", "sum"),
            nao_conforme=("nao_conforme_count", "sum"),
        )
        rows.extend(
            [
                [str(label), int(row["inspecoes"]), int(row["conforme"]), int(row["nao_conforme"])]
                for label, row in grouped.iterrows()
            ]
        )
    table = Table(rows, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#264653")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ]
        )
    )
    return table


def _title(scope: ReportScope, label: str) -> str:
    return (
        "Relatório de Auditoria — São Paulo"
        if scope == ReportScope.CITY
        else f"Relatório — {label}"
    )
