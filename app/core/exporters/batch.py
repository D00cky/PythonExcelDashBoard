"""Multi-Polo (batch) report exporters.

Builds md / xlsx output from a combined inspections frame. Renders the same
HTML dashboard offline as a self-contained .html attachment. docx / pdf are
not yet supported for batches — callers should fall back to a 501.
"""

from __future__ import annotations

import io
from dataclasses import dataclass

import pandas as pd
from openpyxl import Workbook

from app.core.aggregator import (
    PoloBatch,
    combined_inspections,
    combined_stage_failures,
    filter_batch,
    ic_rows_from_inspections,
    iqs_overall_from_inspections,
    iqs_rows_from_inspections,
)
from app.core.templates.pimentas import PimentasTemplate

_BATCH_MIMETYPES = {
    "md": "text/markdown; charset=utf-8",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "html": "text/html; charset=utf-8",
}


@dataclass(frozen=True)
class BatchSelection:
    polos: tuple[str, ...]
    view: str
    period_key: str


def render_batch_export(fmt: str, batch: PoloBatch, selection: BatchSelection) -> tuple[bytes, str]:
    if fmt not in _BATCH_MIMETYPES:
        raise ValueError(f"unsupported batch format: {fmt}")

    filtered = filter_batch(
        batch, polos=selection.polos, view=selection.view, period_key=selection.period_key
    )
    inspections = combined_inspections(filtered)
    failures = combined_stage_failures(filtered)

    if fmt == "md":
        body = _render_markdown(filtered, selection, inspections, failures).encode("utf-8")
    elif fmt == "xlsx":
        body = _render_xlsx(filtered, selection, inspections, failures)
    elif fmt == "html":
        body = _render_html(filtered, selection, inspections)
    return body, _BATCH_MIMETYPES[fmt]


def _polo_label(selection: BatchSelection) -> str:
    if len(selection.polos) == 1:
        return selection.polos[0].title()
    return "Múltiplos Polos"


def _periodo_label(inspections: pd.DataFrame, selection: BatchSelection) -> str:
    if not inspections.empty and "start_date" in inspections.columns:
        dates = inspections["start_date"].dropna()
        if not dates.empty:
            return f"{dates.min():%d/%m/%Y} à {dates.max():%d/%m/%Y}"
    return selection.period_key


def _render_markdown(
    batch: PoloBatch,
    selection: BatchSelection,
    inspections: pd.DataFrame,
    failures: pd.DataFrame,
) -> str:
    template = PimentasTemplate()
    services = sorted(template.SERVICE_SHEETS)
    iqs_rows = iqs_rows_from_inspections(inspections, services)
    ic_rows = ic_rows_from_inspections(inspections, services)
    iqs_overall = iqs_overall_from_inspections(inspections)

    lines: list[str] = [
        f"# Dashboard — {_polo_label(selection)}",
        "",
        f"**Visão**: {'Semanal' if selection.view == 'weekly' else 'Mensal'}  ",
        f"**Período**: {_periodo_label(inspections, selection)}  ",
        f"**Polos incluídos**: {', '.join(p.title() for p in selection.polos)}  ",
    ]
    if iqs_overall is not None:
        lines.append(f"**IQS Geral**: {iqs_overall:.1%}  ")
    if not inspections.empty:
        lines.append(f"**Total de inspeções**: {len(inspections)}  ")
    lines.append("")

    if ic_rows:
        lines += [
            "## Índice de Conformidade (IC) por Serviço",
            "",
            "| Serviço | IC (%) | LVs |",
            "|---|---:|---:|",
        ]
        lines += [f"| {r.name} | {r.ic_pct:.1%} | {r.lvs} |" for r in ic_rows]
        lines.append("")

    if iqs_rows:
        lines += [
            "## Índice de Qualidade (IQS) por Serviço",
            "",
            "| Serviço | Fotos Avaliadas | NC | Conforme | NC (%) | Conforme (%) |",
            "|---|---:|---:|---:|---:|---:|",
        ]
        lines += [
            f"| {r.name} | {r.fotos_avaliadas} | {r.fotos_nc} | "
            f"{r.fotos_conforme} | {r.nc_pct:.1%} | {r.conforme_pct:.1%} |"
            for r in iqs_rows
        ]
        lines.append("")

    if not inspections.empty and "polo" in inspections.columns:
        per_polo = (
            inspections.groupby("polo")
            .agg(
                inspecoes=("polo", "size"),
                conforme=("conforme_count", "sum"),
                nao_conforme=("nao_conforme_count", "sum"),
            )
            .reset_index()
        )
        lines += [
            "## Resumo por Polo",
            "",
            "| Polo | Inspeções | Conforme | Não Conforme |",
            "|---|---:|---:|---:|",
        ]
        for _, row in per_polo.iterrows():
            lines.append(
                f"| {row['polo']} | {int(row['inspecoes'])} | "
                f"{int(row['conforme'])} | {int(row['nao_conforme'])} |"
            )
        lines.append("")

    if not failures.empty:
        top = failures.groupby("stage").size().sort_values(ascending=False).head(10)
        lines += [
            "## Top 10 Etapas com Não Conformidade",
            "",
            "| Etapa | Ocorrências |",
            "|---|---:|",
        ]
        lines += [f"| {stage} | {int(count)} |" for stage, count in top.items()]
        lines.append("")

    return "\n".join(lines)


def _render_xlsx(
    batch: PoloBatch,
    selection: BatchSelection,
    inspections: pd.DataFrame,
    failures: pd.DataFrame,
) -> bytes:
    wb = Workbook()
    summary = wb.active
    summary.title = "Resumo"
    summary["A1"] = "Polo"
    summary["B1"] = "Inspeções"
    summary["C1"] = "Conforme"
    summary["D1"] = "Não Conforme"
    summary["E1"] = "Período"
    if not inspections.empty and "polo" in inspections.columns:
        row = 2
        for polo in selection.polos:
            sub = inspections[inspections["polo"] == polo]
            if sub.empty:
                continue
            summary[f"A{row}"] = polo
            summary[f"B{row}"] = int(len(sub))
            summary[f"C{row}"] = int(sub["conforme_count"].sum())
            summary[f"D{row}"] = int(sub["nao_conforme_count"].sum())
            summary[f"E{row}"] = selection.period_key
            row += 1

    if not inspections.empty:
        ws = wb.create_sheet("Inspeções")
        ws.append(list(inspections.columns))
        for record in inspections.itertuples(index=False):
            ws.append([_xlsx_cell(v) for v in record])

    if not failures.empty:
        ws = wb.create_sheet("Não Conformidades")
        ws.append(list(failures.columns))
        for record in failures.itertuples(index=False):
            ws.append([_xlsx_cell(v) for v in record])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _xlsx_cell(value):
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if pd.isna(value):
        return None
    return value


def _render_html(
    batch: PoloBatch,
    selection: BatchSelection,
    inspections: pd.DataFrame,
) -> bytes:
    """Minimal self-contained HTML — title + KPIs + per-Polo summary table.

    The full live dashboard's Plotly charts require a JS runtime; this
    standalone export ships the headline KPIs and a per-Polo table so it
    can be opened offline without a browser hitting the CDN.
    """
    template = PimentasTemplate()
    services = sorted(template.SERVICE_SHEETS)
    ic_rows = ic_rows_from_inspections(inspections, services)
    iqs_rows = iqs_rows_from_inspections(inspections, services)
    iqs_overall = iqs_overall_from_inspections(inspections)

    polo_rows = ""
    if not inspections.empty and "polo" in inspections.columns:
        for polo in selection.polos:
            sub = inspections[inspections["polo"] == polo]
            polo_rows += (
                f"<tr><td>{polo.title()}</td><td>{len(sub)}</td>"
                f"<td>{int(sub['conforme_count'].sum())}</td>"
                f"<td>{int(sub['nao_conforme_count'].sum())}</td></tr>"
            )

    iqs_pct = f"{iqs_overall:.1%}" if iqs_overall is not None else "—"
    ic_rows_html = "".join(
        f"<tr><td>{r.name}</td><td>{r.ic_pct:.1%}</td><td>{r.lvs}</td></tr>" for r in ic_rows
    )
    iqs_rows_html = "".join(
        f"<tr><td>{r.name}</td><td>{r.fotos_avaliadas}</td>"
        f"<td>{r.fotos_nc}</td><td>{r.fotos_conforme}</td>"
        f"<td>{r.nc_pct:.1%}</td><td>{r.conforme_pct:.1%}</td></tr>"
        for r in iqs_rows
    )
    return (
        f"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<title>Dashboard — {_polo_label(selection)}</title>
<style>body{{font-family:sans-serif;margin:2rem}}table{{border-collapse:collapse;margin:1rem 0}}
th,td{{border:1px solid #ccc;padding:.35rem .7rem;text-align:left}}th{{background:#f0f4f3}}</style>
</head><body>
<h1>Dashboard — {_polo_label(selection)}</h1>
<p><strong>Visão:</strong> {"Semanal" if selection.view == "weekly" else "Mensal"} &nbsp;
<strong>Período:</strong> {_periodo_label(inspections, selection)} &nbsp;
<strong>Polos incluídos:</strong> {", ".join(p.title() for p in selection.polos)}</p>
<p><strong>IQS Geral:</strong> {iqs_pct} &nbsp; <strong>Inspeções:</strong> {len(inspections)}</p>
<h2>Resumo por Polo</h2>
<table><tr><th>Polo</th><th>Inspeções</th><th>Conforme</th><th>Não Conforme</th></tr>
{polo_rows}</table>
<h2>IC por Serviço</h2>
<table><tr><th>Serviço</th><th>IC (%)</th><th>LVs</th></tr>{ic_rows_html}</table>
<h2>IQS por Serviço</h2>
<table><tr><th>Serviço</th><th>Fotos Avaliadas</th><th>NC</th><th>Conforme</th>
<th>NC (%)</th><th>Conforme (%)</th></tr>{iqs_rows_html}</table>
</body></html>"""
    ).encode()
