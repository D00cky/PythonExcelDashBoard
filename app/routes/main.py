import uuid
from pathlib import Path
from typing import Any, Literal

from flask import (
    Blueprint,
    Response,
    abort,
    current_app,
    redirect,
    render_template,
    request,
    url_for,
)
from openpyxl import load_workbook

from app.core.templates import recognize
from app.core.templates.sabesp_pimentas import SabespPimentasTemplate

bp = Blueprint("main", __name__)


@bp.get("/")
def index() -> str:
    return render_template("index.html")


@bp.post("/upload")
def upload():
    if "file" not in request.files:
        abort(400)
    file = request.files["file"]
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        abort(400)

    upload_id = uuid.uuid4().hex
    uploads_dir = Path(current_app.instance_path) / "uploads"
    uploads_dir.mkdir(parents=True, exist_ok=True)
    file.save(uploads_dir / f"{upload_id}.xlsx")

    return redirect(url_for("main.dashboard", upload_id=upload_id), code=303)


@bp.get("/dashboard/<upload_id>")
def dashboard(upload_id: str) -> str:
    path = _upload_path(upload_id)
    workbook = load_workbook(path, data_only=True)
    template = recognize(workbook.sheetnames)
    if not isinstance(template, SabespPimentasTemplate):
        return render_template("dashboard_unknown.html", sheet_names=workbook.sheetnames)

    return render_template(
        "dashboard.html",
        download_action=url_for("main.download", upload_id=upload_id),
        **_build_sabesp_context(template, workbook, path, plotly_mode="cdn"),
    )


_SUPPORTED_FORMATS = {"md", "xlsx", "pdf", "docx"}


@bp.get("/download/<upload_id>")
def download(upload_id: str) -> Response:
    fmt = request.args.get("fmt", "md").lower()
    if fmt not in _SUPPORTED_FORMATS:
        abort(400)

    path = _upload_path(upload_id)
    workbook = load_workbook(path, data_only=True)
    template = recognize(workbook.sheetnames)
    if not isinstance(template, SabespPimentasTemplate):
        abort(404)

    from app.core.exporters import render_export

    body, mimetype = render_export(fmt, template, workbook, path)
    response = Response(body, mimetype=mimetype)
    response.headers["Content-Disposition"] = f'attachment; filename="dashboard-{upload_id}.{fmt}"'
    return response


def _upload_path(upload_id: str) -> Path:
    path = Path(current_app.instance_path) / "uploads" / f"{upload_id}.xlsx"
    if not path.exists():
        abort(404)
    return path


def _periodo_from_inspections(df) -> str | None:
    if df.empty or "start_date" not in df.columns:
        return None
    dates = df["start_date"].dropna()
    if dates.empty:
        return None
    return f"{dates.min():%d/%m/%Y} à {dates.max():%d/%m/%Y}"


def _build_sabesp_context(
    template: SabespPimentasTemplate,
    workbook,
    path: Path,
    plotly_mode: Literal["cdn", "inline"],
) -> dict[str, Any]:
    iqs_rows = template.extract_iqs_by_service(workbook)
    inspections = template.extract_inspections(path)
    periodo = _periodo_from_inspections(inspections) or template.extract_periodo(workbook)

    per_service_sections = []
    for idx, service in enumerate(sorted(template.SERVICE_SHEETS)):
        per_service_sections.append(
            {
                "service": service,
                "team_chart": template.build_team_conformity_for_service(
                    inspections, service
                ).to_html(include_plotlyjs=False, full_html=False, div_id=f"conf-team-{idx}"),
                "tss_chart": template.build_tss_conformity_for_service(
                    inspections, service
                ).to_html(include_plotlyjs=False, full_html=False, div_id=f"conf-tss-{idx}"),
            }
        )

    return {
        "polo_name": template.polo_name.title(),
        "periodo": periodo,
        "iqs_overall": template.extract_iqs_overall(workbook),
        "total_fotos": sum(r.fotos_avaliadas for r in iqs_rows),
        "total_inspections": len(inspections),
        "fig_ic_bar": template.build_ic_bar(template.extract_ic_by_service(workbook)).to_html(
            include_plotlyjs=plotly_mode, full_html=False, div_id="ic-bar"
        ),
        "fig_iqs_bar": template.build_service_iqs_bar(iqs_rows).to_html(
            include_plotlyjs=False, full_html=False, div_id="iqs-bar"
        ),
        "fig_photos": template.build_photo_conformity_stacked(iqs_rows).to_html(
            include_plotlyjs=False, full_html=False, div_id="photos"
        ),
        "fig_team_service": template.build_team_service_stacked(inspections).to_html(
            include_plotlyjs=False, full_html=False, div_id="team-service"
        ),
        "fig_tss": template.build_tss_distribution(inspections).to_html(
            include_plotlyjs=False, full_html=False, div_id="tss-distribution"
        ),
        "per_service_sections": per_service_sections,
    }
