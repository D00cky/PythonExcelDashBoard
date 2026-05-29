import uuid
from pathlib import Path

from flask import Blueprint, abort, current_app, redirect, render_template, request, url_for
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
    path = Path(current_app.instance_path) / "uploads" / f"{upload_id}.xlsx"
    if not path.exists():
        abort(404)

    workbook = load_workbook(path, data_only=True)
    template = recognize(workbook.sheetnames)
    if not isinstance(template, SabespPimentasTemplate):
        return render_template("dashboard_unknown.html", sheet_names=workbook.sheetnames)

    iqs_rows = template.extract_iqs_by_service(workbook)
    iqs_bar = template.build_service_iqs_bar(iqs_rows)
    photos = template.build_photo_conformity_stacked(iqs_rows)

    return render_template(
        "dashboard.html",
        periodo=template.extract_periodo(workbook),
        iqs_overall=template.extract_iqs_overall(workbook),
        fig_iqs_bar=iqs_bar.to_html(include_plotlyjs="cdn", full_html=False, div_id="iqs-bar"),
        fig_photos=photos.to_html(include_plotlyjs=False, full_html=False, div_id="photos"),
    )
