"""Unit tests for threaded background ingestion jobs (``app.core.jobs``)."""

import time
from pathlib import Path

import pytest

from app import create_app
from app.core import cache, jobs
from app.core.aggregator import discover_polo_file, write_manifest
from tests.fixtures.pimentas_minimal import make_minimal_pimentas


@pytest.fixture
def app(tmp_path):
    application = create_app({"TESTING": True, "INSTANCE_PATH": str(tmp_path)})
    with application.app_context():
        yield application


def _make_batch(app, uuid="u1") -> str:
    base = Path(app.instance_path) / "uploads" / uuid
    base.mkdir(parents=True)
    pim = make_minimal_pimentas(
        base, polo="PIMENTAS", with_inspections=True, file_name="file_00.xlsx"
    )
    san = make_minimal_pimentas(
        base, polo="SANTANA", with_inspections=True, file_name="file_01.xlsx"
    )
    write_manifest(base, [discover_polo_file(pim), discover_polo_file(san)])
    return uuid


def test_get_status_none_when_unknown(app):
    assert jobs.get_status("nope") is None


def test_run_ingest_marks_done_and_builds_cache(app):
    uuid = _make_batch(app)
    jobs.run_ingest(app, uuid)
    status = jobs.get_status(uuid)
    assert status["status"] == jobs.JOB_DONE
    assert status["progress"] == status["total"] == 2
    assert status["elapsed_seconds"] >= 0
    assert cache.cache_exists(uuid)


def test_run_ingest_failure_sets_failed_with_error(app):
    # No files for this uuid → ingest_upload raises FileNotFoundError.
    jobs.run_ingest(app, "missing")
    status = jobs.get_status("missing")
    assert status["status"] == jobs.JOB_FAILED
    assert status["error"]


def test_get_status_done_when_cache_exists_without_status_file(app):
    from app.core.ingest import ingest_batch

    base = Path(app.instance_path) / "uploads" / "u2"
    _make_batch(app, "u2")
    ingest_batch("u2", base)  # warm cache directly, no job status written
    status = jobs.get_status("u2")
    assert status["status"] == jobs.JOB_DONE


def test_start_ingest_job_sync_under_testing(app):
    uuid = _make_batch(app)
    jobs.start_ingest_job(uuid)  # TESTING → runs inline
    assert jobs.get_status(uuid)["status"] == jobs.JOB_DONE


def test_start_ingest_job_threaded(app):
    app.config["INGEST_SYNC"] = False
    uuid = _make_batch(app)
    jobs.start_ingest_job(uuid)
    # Poll until the daemon thread finishes (tiny fixtures complete fast).
    deadline = time.time() + 10
    while time.time() < deadline and jobs.get_status(uuid)["status"] != jobs.JOB_DONE:
        time.sleep(0.05)
    assert jobs.get_status(uuid)["status"] == jobs.JOB_DONE
    assert cache.cache_exists(uuid)


def test_processing_status_reports_elapsed_and_eta(app):
    # Hand-write a processing status to exercise the live-computed fields.
    jobs._write_status(
        app,
        "u3",
        status=jobs.JOB_PROCESSING,
        progress=1,
        total=4,
        current_polo="PIMENTAS",
        current_zone="Zona Leste Metropolitana",
        started_at=time.time() - 10,
    )
    status = jobs.get_status("u3")
    assert status["status"] == jobs.JOB_PROCESSING
    assert status["elapsed_seconds"] >= 9
    # 10s for 1 of 4 → ~30s remaining for the other 3.
    assert status["estimated_remaining"] == pytest.approx(30, abs=5)


def test_run_ingest_single_file_upload(app):
    uploads = Path(app.instance_path) / "uploads"
    uploads.mkdir(parents=True, exist_ok=True)
    make_minimal_pimentas(uploads, polo="LAPA", with_inspections=True, file_name="solo.xlsx")
    jobs.run_ingest(app, "solo")
    status = jobs.get_status("solo")
    assert status["status"] == jobs.JOB_DONE
    assert status["total"] == 1


def test_processing_status_without_progress_has_no_eta(app):
    jobs._write_status(
        app, "u4", status=jobs.JOB_PROCESSING, progress=0, total=4, started_at=time.time() - 3
    )
    status = jobs.get_status("u4")
    assert "estimated_remaining" not in status
    assert status["elapsed_seconds"] >= 2


def test_queued_status_written_before_work(app):
    uuid = _make_batch(app)
    jobs._write_status(app, uuid, status=jobs.JOB_QUEUED, progress=0, total=2)
    assert jobs.get_status(uuid)["status"] == jobs.JOB_QUEUED
