from io import BytesIO
from pathlib import Path

from openpyxl.workbook import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from app.core.templates.pimentas import PimentasTemplate


def render_pdf(template: PimentasTemplate, workbook: Workbook, path: Path) -> bytes:
    inspections = template.extract_inspections(path)
    iqs_rows = template.extract_iqs_by_service(workbook)
    ic_rows = template.extract_ic_by_service(workbook)
    iqs_overall = template.extract_iqs_overall(workbook)
    periodo = _periodo(inspections) or template.extract_periodo(workbook)

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
        title=f"Dashboard — Polo {template.polo_name.title()}",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title",
        parent=styles["Title"],
        textColor=colors.HexColor("#264653"),
        spaceAfter=12,
    )
    h2_style = ParagraphStyle(
        "h2",
        parent=styles["Heading2"],
        textColor=colors.HexColor("#264653"),
        spaceBefore=18,
        spaceAfter=8,
    )

    story = []
    story.append(Paragraph(f"Dashboard — Polo {template.polo_name.title()}", title_style))
    if periodo:
        story.append(Paragraph(f"<b>Período</b>: {periodo}", styles["Normal"]))
    if iqs_overall is not None:
        story.append(Paragraph(f"<b>IQS Geral</b>: {iqs_overall:.1%}", styles["Normal"]))
    if not inspections.empty:
        story.append(
            Paragraph(
                f"<b>Total de inspeções</b>: {len(inspections)} "
                f"&nbsp;&nbsp; <b>Equipes distintas</b>: {inspections['team'].nunique()}",
                styles["Normal"],
            )
        )
    story.append(Spacer(1, 8))

    if ic_rows:
        story.append(Paragraph("Índice de Conformidade por Serviço", h2_style))
        story.append(
            _table(
                ["Serviço", "IC (%)", "LVs"],
                [[r.name, f"{r.ic_pct:.1%}", str(r.lvs)] for r in ic_rows],
            )
        )

    if iqs_rows:
        story.append(Paragraph("Índice de Qualidade por Serviço", h2_style))
        story.append(
            _table(
                ["Serviço", "Avaliadas", "NC", "Conforme", "NC (%)", "Conforme (%)"],
                [
                    [
                        r.name,
                        str(r.fotos_avaliadas),
                        str(r.fotos_nc),
                        str(r.fotos_conforme),
                        f"{r.nc_pct:.1%}",
                        f"{r.conforme_pct:.1%}",
                    ]
                    for r in iqs_rows
                ],
            )
        )

    if not inspections.empty:
        for service in sorted(template.SERVICE_SHEETS):
            sub = inspections[inspections["service"] == service]
            if sub.empty:
                continue
            story.append(Paragraph(f"{service} — Top 10 Equipes", h2_style))
            agg = sub.groupby("team").agg(
                conforme=("conforme_count", "sum"),
                nao_conforme=("nao_conforme_count", "sum"),
            )
            agg["total"] = agg["conforme"] + agg["nao_conforme"]
            top = agg[agg["total"] > 0].sort_values("total", ascending=False).head(10)
            story.append(
                _table(
                    ["Equipe", "Conforme", "Não Conforme"],
                    [
                        [team, str(int(r["conforme"])), str(int(r["nao_conforme"]))]
                        for team, r in top.iterrows()
                    ],
                )
            )

    doc.build(story)
    return buf.getvalue()


def _table(headers: list[str], rows: list[list[str]]) -> Table:
    data = [headers] + rows
    t = Table(data, repeatRows=1, hAlign="LEFT")
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#264653")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                ("TOPPADDING", (0, 0), (-1, 0), 6),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ]
        )
    )
    return t


def _periodo(df) -> str | None:
    if df.empty or "start_date" not in df.columns:
        return None
    dates = df["start_date"].dropna()
    if dates.empty:
        return None
    return f"{dates.min():%d/%m/%Y} à {dates.max():%d/%m/%Y}"
