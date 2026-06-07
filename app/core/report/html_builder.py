"""Build self-contained HTML reports for scoped exports."""

from __future__ import annotations

from html import escape
from pathlib import Path

from app.core.report.data import load_report_data
from app.core.report.scope import ReportScope


def build_html(
    uuid: str,
    period: str,
    scope: ReportScope,
    scope_name: str | None,
    output_path: Path,
) -> Path:
    """Build a scoped HTML report and return ``output_path``."""
    data = load_report_data(uuid, scope=scope, scope_name=scope_name, period=period)
    rows = ""
    if not data.inspections.empty:
        group_col = "polo" if "polo" in data.inspections.columns else "service"
        grouped = data.inspections.groupby(group_col).agg(
            inspecoes=("conforme_count", "size"),
            conforme=("conforme_count", "sum"),
            nao_conforme=("nao_conforme_count", "sum"),
        )
        for label, row in grouped.iterrows():
            rows += (
                f"<tr><td>{escape(str(label))}</td><td>{int(row['inspecoes'])}</td>"
                f"<td>{int(row['conforme'])}</td><td>{int(row['nao_conforme'])}</td></tr>"
            )
    title = (
        "Relatório de Auditoria — São Paulo"
        if scope == ReportScope.CITY
        else f"Relatório — {data.label}"
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        f"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>{escape(title)}</title>
<style>body{{font-family:sans-serif;margin:2rem;color:#1a1a1a}}button{{padding:.5rem 1rem}}
table{{border-collapse:collapse;margin-top:1rem}}th,td{{border:1px solid #ccc;padding:.4rem .7rem}}
@media print{{button{{display:none}}body{{background:white;color:#1a1a1a}}}}</style></head>
<body><button onclick="window.print()">Imprimir</button><h1>{escape(title)}</h1>
<p><strong>Período:</strong> {escape(data.period)} ·
<strong>Inspeções:</strong> {len(data.inspections)}</p>
<table><thead><tr><th>Grupo</th><th>Inspeções</th><th>Conforme</th>
<th>Não Conforme</th></tr></thead>
<tbody>{rows}</tbody></table></body></html>""",
        encoding="utf-8",
    )
    return output_path
