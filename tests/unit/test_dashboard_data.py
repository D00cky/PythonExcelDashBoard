"""Unit tests for the cache-aware dashboard data source.

Verifies the warm path reads from parquet and the cold / partially-warm paths
fall back to live parsing with an identical result.
"""

from pathlib import Path

import pytest

from app import create_app
from app.core import aggregator, cache, dashboard_data, ingest
from app.core.aggregator import PoloBatch, discover_polo_file, write_manifest
from tests.fixtures.pimentas_minimal import make_minimal_pimentas


@pytest.fixture
def app(tmp_path):
    application = create_app({"TESTING": True, "INSTANCE_PATH": str(tmp_path)})
    with application.app_context():
        yield application


def _batch(app) -> tuple[str, Path]:
    base = Path(app.instance_path) / "uploads" / "u1"
    base.mkdir(parents=True)
    pim = make_minimal_pimentas(
        base, polo="PIMENTAS", with_inspections=True, file_name="file_00.xlsx"
    )
    san = make_minimal_pimentas(
        base, polo="SANTANA", with_inspections=True, file_name="file_01.xlsx"
    )
    write_manifest(base, [discover_polo_file(pim), discover_polo_file(san)])
    return "u1", base


def test_warm_cache_reads_from_parquet(app):
    uuid, base = _batch(app)
    ingest.ingest_batch(uuid, base)
    batch = aggregator.load_batch(base)

    cached = dashboard_data.combined_inspections(uuid, batch)
    live = aggregator.combined_inspections(batch)
    assert len(cached) == len(live)
    assert set(cached["polo"]) == set(live["polo"]) == {"PIMENTAS", "SANTANA"}

    cached_fail = dashboard_data.combined_stage_failures(uuid, batch)
    assert "polo" in cached_fail.columns


def test_cold_cache_falls_back_to_live(app):
    _, base = _batch(app)
    batch = aggregator.load_batch(base)
    cached = dashboard_data.combined_inspections("never-warmed", batch)
    live = aggregator.combined_inspections(batch)
    assert len(cached) == len(live)


def test_partial_cache_falls_back_to_live(app):
    uuid, base = _batch(app)
    ingest.ingest_batch(uuid, base)
    # Simulate a file present in the batch but missing from the cache: drop one
    # inspections parquet (filenames are slugged, so target it by globbing).
    inspection_parquets = sorted(
        p for p in (cache.cache_dir(uuid) / "raw").glob("*.parquet") if "fail" not in p.name
    )
    inspection_parquets[1].unlink()
    batch = aggregator.load_batch(base)

    cached = dashboard_data.combined_inspections(uuid, batch)
    live = aggregator.combined_inspections(batch)
    assert len(cached) == len(live)  # fell back, not a truncated cache read


def test_warm_cache_empty_filtered_batch(app):
    uuid, base = _batch(app)
    ingest.ingest_batch(uuid, base)
    empty = PoloBatch(batch_dir=base, files=[])
    result = dashboard_data.combined_inspections(uuid, empty)
    assert result.empty
    assert list(result.columns) == ["polo", "municipality", "zone"]
