import uuid
from pathlib import Path

from flask import Blueprint, abort, current_app, redirect, render_template, request

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

    return redirect(f"/dashboard/{upload_id}", code=303)
