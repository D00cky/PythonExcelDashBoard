"""Unit tests for the parquet cache store (``app.core.cache``).

The cache is a thin, uuid-keyed parquet store; orchestration (parsing xlsx
into it) lives in ``app.core.ingest`` and is tested separately. These tests
exercise storage primitives in isolation, inside a Flask app context so the
cache can resolve ``instance_path``.
"""

import pandas as pd
import pytest

from app import create_app
from app.core import cache


@pytest.fixture
def ctx(tmp_path):
    app = create_app({"TESTING": True, "INSTANCE_PATH": str(tmp_path)})
    with app.app_context():
        yield


def _frame():
    return pd.DataFrame(
        {
            "team": ["A", "B"],
            "service": ["ÁGUA", "ESGOTO"],
            "conforme_count": [1, 0],
            "polo": ["PIMENTAS", "PIMENTAS"],
            "zone": ["Zona Leste Metropolitana"] * 2,
        }
    )


def test_cache_exists_false_when_absent(ctx):
    assert cache.cache_exists("nope") is False


def test_save_and_load_raw_roundtrip(ctx):
    df = _frame()
    cache.save_raw("u1", "PIMENTAS", df)
    loaded = cache.load_raw("u1", "PIMENTAS")
    pd.testing.assert_frame_equal(loaded, df)


def test_raw_handles_accented_polo_name(ctx):
    df = _frame()
    cache.save_raw("u1", "GOPOÚVA", df)
    pd.testing.assert_frame_equal(cache.load_raw("u1", "GOPOÚVA"), df)


def test_distinct_polos_do_not_collide(ctx):
    cache.save_raw("u1", "SANTANA", _frame().assign(polo="SANTANA"))
    cache.save_raw("u1", "SÃO MATEUS", _frame().assign(polo="SÃO MATEUS"))
    assert cache.load_raw("u1", "SANTANA")["polo"].iloc[0] == "SANTANA"
    assert cache.load_raw("u1", "SÃO MATEUS")["polo"].iloc[0] == "SÃO MATEUS"


def test_save_and_load_aggregation_roundtrip(ctx):
    agg = pd.DataFrame({"service": ["ÁGUA"], "ic_pct": [0.5], "lvs": [10]})
    cache.save_aggregation("u1", "zone", "Zona Norte", agg)
    pd.testing.assert_frame_equal(cache.load_aggregation("u1", "zone", "Zona Norte"), agg)


def test_city_aggregation_uses_empty_name(ctx):
    agg = pd.DataFrame({"iqs_overall": [0.66]})
    cache.save_aggregation("u1", "city", "", agg)
    pd.testing.assert_frame_equal(cache.load_aggregation("u1", "city", ""), agg)


def test_meta_roundtrip(ctx):
    meta = {"raw": {"PIMENTAS": "raw/pimentas.parquet"}, "parsed_at": "2026-06-07T00:00:00"}
    cache.write_meta("u1", meta)
    assert cache.load_meta("u1") == meta


def test_cache_exists_true_after_meta_written(ctx):
    cache.write_meta("u1", {"raw": {}})
    assert cache.cache_exists("u1") is True


def test_invalidate_removes_everything(ctx):
    cache.save_raw("u1", "PIMENTAS", _frame())
    cache.write_meta("u1", {"raw": {"PIMENTAS": "raw/pimentas.parquet"}})
    assert cache.cache_exists("u1") is True
    cache.invalidate("u1")
    assert cache.cache_exists("u1") is False
    assert not cache.cache_dir("u1").exists()


def test_invalidate_missing_uuid_is_noop(ctx):
    cache.invalidate("never-existed")  # must not raise


def test_get_cache_size_counts_bytes(ctx):
    assert cache.get_cache_size("u1") == 0
    cache.save_raw("u1", "PIMENTAS", _frame())
    assert cache.get_cache_size("u1") > 0


def test_load_raw_missing_raises_keyerror(ctx):
    cache.write_meta("u1", {"raw": {}})
    with pytest.raises(KeyError):
        cache.load_raw("u1", "ABSENT")


def test_load_aggregation_missing_raises_keyerror(ctx):
    cache.write_meta("u1", {"raw": {}})
    with pytest.raises(KeyError):
        cache.load_aggregation("u1", "zone", "Inexistente")
