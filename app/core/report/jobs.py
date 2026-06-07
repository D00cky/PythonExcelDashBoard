"""Local persisted export jobs for heavy report generation."""

from __future__ import annotations

import json
import time
import uuid as uuidlib
from pathlib import Path

from flask import Flask

from app.core.report.docx_builder import build_docx
from app.core.report.html_builder import build_html
from app.core.report.pdf_builder import build_pdf
from app.core.report.pptx_builder import build_pptx
from app.core.report.scope import ReportScope

EXPORT_DONE = "done"
EXPORT_FAILED = "failed"

_MIMETYPES = {
    "html": "text/html; charset=utf-8",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pdf": "application/pdf",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}


def start_export_job(
    app: Flask,
    upload_id: str,
    *,
    fmt: str,
    scope: ReportScope,
    scope_name: str | None,
    period: str,
) -> str:
    """Generate a report synchronously and persist a status file.

    The API is intentionally job-shaped so it can be swapped to a queue later
    without changing routes or templates.
    """
    job_id = uuidlib.uuid4().hex
    try:
        path = _output_path(app, job_id, scope=scope, scope_name=scope_name, period=period, fmt=fmt)
        with app.app_context():
            builder = {
                "html": build_html,
                "docx": build_docx,
                "pdf": build_pdf,
                "pptx": build_pptx,
            }[fmt]
            builder(upload_id, period, scope, scope_name, path)
        _write_status(
            app,
            job_id,
            status=EXPORT_DONE,
            path=str(path),
            fmt=fmt,
            mimetype=_MIMETYPES[fmt],
            download_name=path.name,
            scope=scope.value,
            period=period,
        )
    except Exception as exc:  # noqa: BLE001 - persisted job failure for UI
        _write_status(app, job_id, status=EXPORT_FAILED, error=f"{type(exc).__name__}: {exc}")
    return job_id


def get_export_status(app: Flask, job_id: str) -> dict | None:
    """Return an export job status dictionary, or ``None`` if unknown."""
    path = _status_path(app, job_id)
    if not path.is_file():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def _reports_dir(app: Flask) -> Path:
    return Path(app.instance_path) / "reports"


def _status_path(app: Flask, job_id: str) -> Path:
    return _reports_dir(app) / "jobs" / f"{job_id}.json"


def _write_status(app: Flask, job_id: str, **fields) -> None:
    path = _status_path(app, job_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    fields["updated_at"] = time.time()
    path.write_text(json.dumps(fields, ensure_ascii=False), encoding="utf-8")


def _output_path(
    app: Flask,
    job_id: str,
    *,
    scope: ReportScope,
    scope_name: str | None,
    period: str,
    fmt: str,
) -> Path:
    safe_scope = scope.value if not scope_name else f"{scope.value}_{_slug(scope_name)}"
    safe_period = _slug(period or "periodo")
    return _reports_dir(app) / safe_period / f"relatorio_{safe_scope}_{job_id[:8]}.{fmt}"


def _slug(value: str) -> str:
    return "".join(ch if ch.isalnum() else "_" for ch in value.lower()).strip("_") or "relatorio"
