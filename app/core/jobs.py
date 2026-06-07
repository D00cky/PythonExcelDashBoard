"""Background ingestion jobs — threaded, no external queue.

Each batch upload's ingestion runs in a daemon thread; progress and final state
are persisted to ``<instance>/jobs/<uuid>.json`` so the ``/api/status`` endpoint
and the processing page can read it from any request handler. This targets the
single-process waitress deployment on Render's free tier — **no Redis required**
(the spec's RQ/Redis design needs a service tier Render no longer offers free).

Under ``TESTING`` (or ``INGEST_SYNC=True``) jobs run inline so request/response
flow stays deterministic in tests.

Status files live outside the parquet cache dir because ingest wipes that dir at
the start of every run.

Public API:
    start_ingest_job(uuid) -> None
    run_ingest(app, uuid) -> None          # the worker body (also runs inline)
    get_status(uuid) -> dict | None
    JOB_QUEUED / JOB_PROCESSING / JOB_DONE / JOB_FAILED
"""

from __future__ import annotations

import json
import threading
import time
from pathlib import Path

from flask import Flask, current_app

from app.core import cache
from app.core.aggregator import MANIFEST_FILENAME, load_batch
from app.core.ingest import ingest_upload

JOB_QUEUED = "queued"
JOB_PROCESSING = "processing"
JOB_DONE = "done"
JOB_FAILED = "failed"

_ACTIVE = (JOB_QUEUED, JOB_PROCESSING)


def _jobs_dir(app: Flask) -> Path:
    return Path(app.instance_path) / "jobs"


def _status_path(app: Flask, uuid: str) -> Path:
    return _jobs_dir(app) / f"{uuid}.json"


def _write_status(app: Flask, uuid: str, **fields) -> None:
    """Merge ``fields`` into the persisted status, stamping ``updated_at``."""
    path = _status_path(app, uuid)
    path.parent.mkdir(parents=True, exist_ok=True)
    current = {}
    if path.is_file():
        current = json.loads(path.read_text(encoding="utf-8"))
    current.update(fields)
    current["updated_at"] = time.time()
    path.write_text(json.dumps(current, ensure_ascii=False), encoding="utf-8")


def get_status(uuid: str) -> dict | None:
    """Current job status, or ``None`` if nothing is known about ``uuid``.

    A cache that exists without a status file (e.g. warmed directly, or by a
    pre-jobs code path) reports as ``done``. Processing statuses are enriched
    with live ``elapsed_seconds`` and ``estimated_remaining``.
    """
    app = current_app._get_current_object()
    path = _status_path(app, uuid)
    if not path.is_file():
        return {"status": JOB_DONE, "progress": 0, "total": 0} if cache.cache_exists(uuid) else None
    status = json.loads(path.read_text(encoding="utf-8"))
    return _enrich(status)


def _enrich(status: dict) -> dict:
    started = status.get("started_at")
    if status.get("status") == JOB_PROCESSING and started is not None:
        elapsed = max(0.0, time.time() - started)
        status["elapsed_seconds"] = round(elapsed, 1)
        progress, total = status.get("progress", 0), status.get("total", 0)
        if progress > 0 and total > 0:
            status["estimated_remaining"] = round(elapsed / progress * (total - progress), 1)
    elif status.get("status") == JOB_DONE and started is not None:
        status["elapsed_seconds"] = round(status.get("updated_at", started) - started, 1)
    return status


def _total_units(app: Flask, uuid: str) -> int:
    uploads = Path(app.instance_path) / "uploads"
    if (uploads / f"{uuid}.xlsx").is_file():
        return 1
    batch_dir = uploads / uuid
    if (batch_dir / MANIFEST_FILENAME).is_file():
        return len(load_batch(batch_dir).files)
    return 0


def run_ingest(app: Flask, uuid: str) -> None:
    """Ingest ``uuid`` and record progress/terminal state. Never raises."""
    with app.app_context():
        started = time.time()
        total = _total_units(app, uuid)
        _write_status(
            app,
            uuid,
            status=JOB_PROCESSING,
            progress=0,
            total=total,
            current_polo="",
            current_zone="",
            started_at=started,
            error="",
        )

        def progress(current: int, total_files: int, polo: str, zone: str) -> None:
            _write_status(
                app,
                uuid,
                status=JOB_PROCESSING,
                progress=current,
                total=total_files,
                current_polo=polo,
                current_zone=zone,
                started_at=started,
            )

        try:
            ingest_upload(uuid, progress_callback=progress)
        except Exception as exc:  # noqa: BLE001 — surface as failed status, not a crash
            current_app.logger.exception("ingest job failed for %s", uuid)
            _write_status(app, uuid, status=JOB_FAILED, error=f"{type(exc).__name__}: {exc}")
            return
        _write_status(app, uuid, status=JOB_DONE, progress=total or 0, total=total or 0)


def start_ingest_job(uuid: str) -> None:
    """Queue ingestion for ``uuid``. Runs inline under TESTING/INGEST_SYNC,
    otherwise in a daemon thread. Writes a ``queued`` status immediately so the
    dashboard shows the processing page on the very next request."""
    app = current_app._get_current_object()
    _write_status(app, uuid, status=JOB_QUEUED, progress=0, total=_total_units(app, uuid))
    if app.config.get("INGEST_SYNC", app.testing):
        run_ingest(app, uuid)
    else:
        threading.Thread(target=run_ingest, args=(app, uuid), daemon=True).start()
