from pathlib import Path

from openpyxl.workbook import Workbook

from app.core.templates.pimentas import PimentasTemplate


def render_markdown(template: PimentasTemplate, workbook: Workbook, path: Path) -> str:
    inspections = template.extract_inspections(path)
    iqs_rows = template.extract_iqs_by_service(workbook)
    ic_rows = template.extract_ic_by_service(workbook)

    lines: list[str] = []
    lines.append(f"# Dashboard — Polo {template.polo_name.title()}")
    lines.append("")

    periodo = _periodo(inspections) or template.extract_periodo(workbook)
    if periodo:
        lines.append(f"**Período**: {periodo}")
    iqs_overall = template.extract_iqs_overall(workbook)
    if iqs_overall is not None:
        lines.append(f"**IQS Geral**: {iqs_overall:.1%}")
    if not inspections.empty:
        lines.append(f"**Total de inspeções**: {len(inspections)}")
        lines.append(f"**Total de equipes distintas**: {inspections['team'].nunique()}")
    lines.append("")

    if ic_rows:
        lines.append("## Índice de Conformidade (IC) por Serviço")
        lines.append("")
        lines.append("| Serviço | IC (%) | Quantidade de LVs |")
        lines.append("|---|---:|---:|")
        for r in ic_rows:
            lines.append(f"| {r.name} | {r.ic_pct:.1%} | {r.lvs} |")
        lines.append("")

    if iqs_rows:
        lines.append("## Índice de Qualidade (IQS) por Serviço")
        lines.append("")
        lines.append("| Serviço | Fotos Avaliadas | NC | Conforme | NC (%) | Conforme (%) |")
        lines.append("|---|---:|---:|---:|---:|---:|")
        for r in iqs_rows:
            lines.append(
                f"| {r.name} | {r.fotos_avaliadas} | {r.fotos_nc} | "
                f"{r.fotos_conforme} | {r.nc_pct:.1%} | {r.conforme_pct:.1%} |"
            )
        lines.append("")

    if not inspections.empty:
        for service in sorted(template.SERVICE_SHEETS):
            sub = inspections[inspections["service"] == service]
            if sub.empty:
                continue
            lines.append(f"## {service} — Top 10 Equipes")
            lines.append("")
            agg = sub.groupby("team").agg(
                conforme=("conforme_count", "sum"),
                nao_conforme=("nao_conforme_count", "sum"),
            )
            agg["total"] = agg["conforme"] + agg["nao_conforme"]
            top = agg[agg["total"] > 0].sort_values("total", ascending=False).head(10)
            lines.append("| Equipe | Conforme | Não Conforme |")
            lines.append("|---|---:|---:|")
            for team, row in top.iterrows():
                lines.append(f"| {team} | {int(row['conforme'])} | {int(row['nao_conforme'])} |")
            lines.append("")

    return "\n".join(lines)


def _periodo(df) -> str | None:
    if df.empty or "start_date" not in df.columns:
        return None
    dates = df["start_date"].dropna()
    if dates.empty:
        return None
    return f"{dates.min():%d/%m/%Y} à {dates.max():%d/%m/%Y}"
