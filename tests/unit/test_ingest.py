"""Unit tests for the ingestion orchestrator (``app.core.ingest``).

Batches are built from openpyxl fixtures (no real Model/ files) so the fast
suite stays fast. Two Polos in two different zones (PIMENTAS → Leste
Metropolitana, SANTANA → Norte) exercise the municipality → zone → city
rollup. The memory test guards the "never accumulate raw rows" invariant; the
calamine-vs-openpyxl benchmark is marked slow.
"""

import time
import tracemalloc
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

from app import create_app
from app.core import aggregations, cache, ingest
from app.core.aggregator import discover_polo_file, write_manifest
from tests.fixtures.pimentas_minimal import make_minimal_pimentas


@pytest.fixture
def app(tmp_path):
    application = create_app({"TESTING": True, "INSTANCE_PATH": str(tmp_path)})
    with application.app_context():
        yield application


def _uploads(app) -> Path:
    d = Path(app.instance_path) / "uploads"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _make_batch(base: Path) -> Path:
    base.mkdir(parents=True, exist_ok=True)
    pim = make_minimal_pimentas(
        base, polo="PIMENTAS", with_inspections=True, file_name="file_00.xlsx"
    )
    san = make_minimal_pimentas(
        base, polo="SANTANA", with_inspections=True, file_name="file_01.xlsx"
    )
    write_manifest(base, [discover_polo_file(pim), discover_polo_file(san)])
    return base


def _muni_inspections(uuid, polo):
    return int(
        aggregations.load_scope(uuid, aggregations.scope_key_muni(polo)).summary.iloc[0][
            "total_inspections"
        ]
    )


def test_ingest_batch_populates_cache(app):
    batch_dir = _make_batch(_uploads(app) / "u1")
    result = ingest.ingest_batch("u1", batch_dir)

    assert cache.cache_exists("u1")
    assert result.polos == ["PIMENTAS", "SANTANA"]
    assert result.n_files == 2
    assert result.elapsed_seconds >= 0
    meta = cache.load_meta("u1")
    assert set(meta["polo_to_zone"]) == {"PIMENTAS", "SANTANA"}
    assert meta["polo_to_zone"]["SANTANA"] == "Zona Norte"


def test_raw_parquet_per_file_loadable(app):
    batch_dir = _make_batch(_uploads(app) / "u1")
    ingest.ingest_batch("u1", batch_dir)
    raw = cache.load_raw("u1", "file_00")  # keyed by source file stem
    assert (raw["polo"] == "PIMENTAS").all()
    assert "zone" in raw.columns
    failures = cache.load_raw("u1", "file_00__fail")
    assert "polo" in failures.columns


def test_same_polo_across_weeks_does_not_collide(app):
    base = _uploads(app) / "u9"
    base.mkdir(parents=True)
    w1 = make_minimal_pimentas(
        base,
        polo="PIMENTAS",
        periodo="Período: 04/05/2026 à 10/05/2026",
        with_inspections=True,
        file_name="file_00.xlsx",
    )
    w2 = make_minimal_pimentas(
        base,
        polo="PIMENTAS",
        periodo="Período: 11/05/2026 à 17/05/2026",
        with_inspections=True,
        file_name="file_01.xlsx",
    )
    write_manifest(base, [discover_polo_file(w1), discover_polo_file(w2)])
    ingest.ingest_batch("u9", base)

    # Both weeks survive as distinct raw files...
    assert len(cache.load_raw("u9", "file_00")) > 0
    assert len(cache.load_raw("u9", "file_01")) > 0
    # ...and the per-Polo overview combines both weeks.
    muni = aggregations.load_scope("u9", aggregations.scope_key_muni("PIMENTAS"))
    per_week = len(cache.load_raw("u9", "file_00")) + len(cache.load_raw("u9", "file_01"))
    assert int(muni.summary.iloc[0]["total_inspections"]) == per_week


def test_city_inspections_equal_sum_of_municipalities(app):
    batch_dir = _make_batch(_uploads(app) / "u1")
    ingest.ingest_batch("u1", batch_dir)
    city = aggregations.load_scope("u1", aggregations.scope_key_city())
    total = int(city.summary.iloc[0]["total_inspections"])
    assert total == _muni_inspections("u1", "PIMENTAS") + _muni_inspections("u1", "SANTANA")
    assert total > 0


def test_zone_scope_saved_for_each_zone(app):
    batch_dir = _make_batch(_uploads(app) / "u1")
    ingest.ingest_batch("u1", batch_dir)
    norte = aggregations.load_scope("u1", aggregations.scope_key_zone("Zona Norte"))
    assert int(norte.summary.iloc[0]["total_inspections"]) == _muni_inspections("u1", "SANTANA")


def test_progress_callback_invoked_per_file(app):
    batch_dir = _make_batch(_uploads(app) / "u1")
    calls = []
    ingest.ingest_batch("u1", batch_dir, progress_callback=lambda *a: calls.append(a))
    assert len(calls) == 2
    assert calls[-1][0] == 2 and calls[-1][1] == 2  # (current, total, polo, zone)


def test_ingest_single_file(app):
    path = make_minimal_pimentas(
        _uploads(app), polo="PIRITUBA", with_inspections=True, file_name="x.xlsx"
    )
    result = ingest.ingest_single("u2", path)
    assert result.polos == ["PIRITUBA"]
    assert cache.cache_exists("u2")


def test_ingest_upload_resolves_single(app):
    make_minimal_pimentas(_uploads(app), polo="LAPA", with_inspections=True, file_name="u3.xlsx")
    result = ingest.ingest_upload("u3")
    assert result.polos == ["LAPA"]


def test_ingest_upload_resolves_batch(app):
    _make_batch(_uploads(app) / "u4")
    result = ingest.ingest_upload("u4")
    assert result.n_files == 2


def test_ingest_file_without_inspections(app):
    path = make_minimal_pimentas(
        _uploads(app), polo="SE", with_inspections=False, file_name="empty.xlsx"
    )
    result = ingest.ingest_single("u5", path)
    assert result.polos == ["SE"]
    city = aggregations.load_scope("u5", aggregations.scope_key_city())
    assert int(city.summary.iloc[0]["total_inspections"]) == 0


def test_ingest_upload_missing_raises(app):
    with pytest.raises(FileNotFoundError):
        ingest.ingest_upload("does-not-exist")


def test_reingest_replaces_stale_cache(app):
    batch_dir = _make_batch(_uploads(app) / "u1")
    ingest.ingest_batch("u1", batch_dir)
    first = cache.load_meta("u1")["polos"]
    ingest.ingest_batch("u1", batch_dir)  # idempotent re-run
    assert cache.load_meta("u1")["polos"] == first


def test_ingest_peak_memory_bounded(app):
    batch_dir = _make_batch(_uploads(app) / "u1")
    tracemalloc.start()
    ingest.ingest_batch("u1", batch_dir)
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    # Tiny fixtures; the point is we never hold all raw frames at once.
    assert peak < 100 * 1024 * 1024


def test_calamine_matches_openpyxl_values(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=True)
    cols = [
        "Unidade Executante",
        "Descrição TSS",
        "Município",
        "Data Início Execução",
        "EQUIPE",
        "FACHADA",
        "SINALIZAÇÃO",
    ]
    cal = pd.read_excel(path, sheet_name="ÁGUA", engine="calamine")
    opx = pd.read_excel(path, sheet_name="ÁGUA", engine="openpyxl")
    assert cal.shape == opx.shape
    pd.testing.assert_frame_equal(cal[cols], opx[cols], check_dtype=False)


@pytest.mark.slow
def test_calamine_faster_than_openpyxl_on_large_sheet(tmp_path):
    path = tmp_path / "big.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append([f"col{c}" for c in range(20)])
    for r in range(3000):
        ws.append([r + c for c in range(20)])
    wb.save(path)

    def median_read(engine):
        times = []
        for _ in range(3):
            t0 = time.perf_counter()
            pd.read_excel(path, sheet_name=0, engine=engine)
            times.append(time.perf_counter() - t0)
        return sorted(times)[1]

    calamine_t = median_read("calamine")
    openpyxl_t = median_read("openpyxl")
    assert calamine_t < openpyxl_t
