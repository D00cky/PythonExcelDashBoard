from io import BytesIO
from pathlib import Path

import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from openpyxl.workbook import Workbook

from app.core.aggregator import date_bounds, format_period_pt
from app.core.templates.pimentas import PimentasTemplate, ServiceIC, ServiceIQS


def render_docx(template: PimentasTemplate, workbook: Workbook, path: Path) -> bytes:
    inspections = template.extract_inspections(path)
    return render_docx_from_data(
        polo_label=template.polo_name.title(),
        periodo=format_period_pt(date_bounds(inspections)) or template.extract_periodo(workbook),
        iqs_overall=template.extract_iqs_overall(workbook),
        iqs_rows=template.extract_iqs_by_service(workbook),
        ic_rows=template.extract_ic_by_service(workbook),
        inspections=inspections,
        services=sorted(template.SERVICE_SHEETS),
    )


def render_docx_from_data(
    *,
    polo_label: str,
    periodo: str | None,
    iqs_overall: float | None,
    iqs_rows: list[ServiceIQS],
    ic_rows: list[ServiceIC],
    inspections: pd.DataFrame,
    services: list[str],
    polos_included: list[str] | None = None,
) -> bytes:
    doc = Document()
    title = doc.add_heading(f"Dashboard — Polo {polo_label}", level=0)
    for run in title.runs:
        run.font.color.rgb = RGBColor(0x26, 0x46, 0x53)

    if periodo:
        doc.add_paragraph().add_run(f"Período: {periodo}").bold = True
    if polos_included:
        doc.add_paragraph().add_run(f"Polos incluídos: {', '.join(polos_included)}").bold = True
    if iqs_overall is not None:
        doc.add_paragraph().add_run(f"IQS Geral: {iqs_overall:.1%}").bold = True
    if not inspections.empty:
        doc.add_paragraph(f"Total de inspeções: {len(inspections)}")
        doc.add_paragraph(f"Equipes distintas: {inspections['team'].nunique()}")

    # Per-Polo summary table when polos_included contains > 1 entry
    if polos_included and len(polos_included) > 1 and "polo" in inspections.columns:
        doc.add_heading("Resumo por Polo", level=1)
        per_polo_rows = []
        for polo in polos_included:
            sub = inspections[inspections["polo"] == polo]
            per_polo_rows.append(
                [
                    polo,
                    str(len(sub)),
                    str(int(sub["conforme_count"].sum())) if not sub.empty else "0",
                    str(int(sub["nao_conforme_count"].sum())) if not sub.empty else "0",
                ]
            )
        _table(doc, ["Polo", "Inspeções", "Conforme", "Não Conforme"], per_polo_rows)

    if ic_rows:
        doc.add_heading("Índice de Conformidade por Serviço", level=1)
        _table(
            doc,
            ["Serviço", "IC (%)", "LVs"],
            [[r.name, f"{r.ic_pct:.1%}", str(r.lvs)] for r in ic_rows],
        )

    if iqs_rows:
        doc.add_heading("Índice de Qualidade por Serviço", level=1)
        _table(
            doc,
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

    if not inspections.empty:
        for service in services:
            sub = inspections[inspections["service"] == service]
            if sub.empty:
                continue
            doc.add_heading(f"{service} — Top 10 Equipes", level=1)
            agg = sub.groupby("team").agg(
                conforme=("conforme_count", "sum"),
                nao_conforme=("nao_conforme_count", "sum"),
            )
            agg["total"] = agg["conforme"] + agg["nao_conforme"]
            top = agg[agg["total"] > 0].sort_values("total", ascending=False).head(10)
            _table(
                doc,
                ["Equipe", "Conforme", "Não Conforme"],
                [
                    [team, str(int(r["conforme"])), str(int(r["nao_conforme"]))]
                    for team, r in top.iterrows()
                ],
            )

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _table(doc, headers: list[str], rows: list[list[str]]) -> None:
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Light Grid Accent 1"
    for c, h in enumerate(headers):
        cell = t.rows[0].cells[c]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(10)
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            t.rows[r].cells[c].text = val
