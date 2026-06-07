"""Unit tests for the multi-level aggregation engine (``app.core.aggregations``).

These use a synthetic inspections DataFrame (no xlsx) so they stay fast and
intent-revealing. The central invariant: an aggregation computed from the
concatenated raw rows of a scope equals the inspection-count-weighted average
of its children — which is what city/zone rollups must satisfy.
"""

import pandas as pd
import pytest

from app import create_app
from app.core import aggregations as agg
from app.core import cache


def _rows(polo, zone, service, n_conforme, n_nc, photos_c, photos_nc):
    """Expand into per-inspection rows for one (polo, service) group."""
    out = []
    for _ in range(n_conforme):
        out.append((polo, zone, service, 1, 0, photos_c, 0, 0, photos_c))
    for _ in range(n_nc):
        out.append((polo, zone, service, 0, 1, 0, photos_nc, 0, photos_nc))
    return out


def _frame():
    cols = [
        "polo",
        "zone",
        "service",
        "conforme_count",
        "nao_conforme_count",
        "photo_conforme",
        "photo_nc",
        "photo_sf",
        "photo_total",
    ]
    data = []
    # Zona Norte: SANTANA + PIRITUBA ; Zona Sul: SANTO AMARO
    data += _rows("SANTANA", "Zona Norte", "ÁGUA", 8, 2, 3, 4)
    data += _rows("PIRITUBA", "Zona Norte", "ÁGUA", 5, 5, 3, 4)
    data += _rows("SANTO AMARO", "Zona Sul", "ÁGUA", 9, 1, 3, 4)
    data += _rows("SANTANA", "Zona Norte", "ESGOTO", 4, 1, 2, 5)
    df = pd.DataFrame(data, columns=cols)
    df["team"] = "EQUIPE"
    return df


@pytest.fixture
def ctx(tmp_path):
    app = create_app({"TESTING": True, "INSTANCE_PATH": str(tmp_path)})
    with app.app_context():
        yield


def _ic_for(scope, service):
    row = scope.ic[scope.ic["service"] == service]
    return float(row["ic_pct"].iloc[0]), int(row["lvs"].iloc[0])


def test_municipality_ic_is_direct_ratio():
    df = _frame()
    santana = agg.compute_scope(df[df["polo"] == "SANTANA"])
    pct, lvs = _ic_for(santana, "Água")
    assert lvs == 10  # 8 conforme + 2 nc for ÁGUA
    assert pct == pytest.approx(0.8)


def test_zone_ic_is_weighted_across_municipalities():
    df = _frame()
    norte = agg.compute_scope(df[df["zone"] == "Zona Norte"])
    pct, lvs = _ic_for(norte, "Água")
    # Santana 8/10 + Pirituba 5/10 -> 13/20
    assert lvs == 20
    assert pct == pytest.approx(13 / 20)


def test_city_ic_equals_inspection_weighted_average_of_zones():
    df = _frame()
    city = agg.compute_scope(df)
    norte = agg.compute_scope(df[df["zone"] == "Zona Norte"])
    sul = agg.compute_scope(df[df["zone"] == "Zona Sul"])

    cpct, clvs = _ic_for(city, "Água")
    npct, nlvs = _ic_for(norte, "Água")
    spct, slvs = _ic_for(sul, "Água")

    weighted = (npct * nlvs + spct * slvs) / (nlvs + slvs)
    assert clvs == nlvs + slvs
    assert cpct == pytest.approx(weighted)


def test_summary_totals_and_iqs_overall():
    df = _frame()
    city = agg.compute_scope(df)
    s = city.summary.iloc[0]
    assert int(s["total_inspections"]) == len(df)
    assert int(s["total_photos"]) == int(df["photo_total"].sum())
    expected_iqs = df["photo_conforme"].sum() / df["photo_total"].sum()
    assert float(s["iqs_overall"]) == pytest.approx(expected_iqs)


def test_iqs_by_service_present():
    df = _frame()
    city = agg.compute_scope(df)
    assert set(city.iqs["service"]) == {"Água", "Esgoto"}


def test_team_ranking_fail_pct():
    df = _frame()
    city = agg.compute_scope(df)
    teams = city.teams.set_index("team")
    total = int(df["nao_conforme_count"].sum())
    assert int(teams.loc["EQUIPE", "nao_conforme"]) == total
    assert int(teams.loc["EQUIPE", "inspecoes"]) == len(df)


def test_empty_scope_yields_empty_frames_with_schema():
    empty = agg.compute_scope(_frame().iloc[0:0])
    assert list(empty.ic.columns) == ["service", "ic_pct", "lvs"]
    assert empty.ic.empty
    assert empty.summary.iloc[0]["total_inspections"] == 0


def test_save_and_load_scope_roundtrip(ctx):
    df = _frame()
    scope = agg.compute_scope(df[df["zone"] == "Zona Norte"])
    agg.save_scope("u1", agg.scope_key_zone("Zona Norte"), scope)
    loaded = agg.load_scope("u1", agg.scope_key_zone("Zona Norte"))
    pd.testing.assert_frame_equal(loaded.ic, scope.ic)
    pd.testing.assert_frame_equal(loaded.iqs, scope.iqs)
    pd.testing.assert_frame_equal(loaded.teams, scope.teams)
    pd.testing.assert_frame_equal(loaded.summary, scope.summary)


def test_scope_keys_are_distinct(ctx):
    assert agg.scope_key_city() == "city"
    assert agg.scope_key_zone("Zona Norte") != agg.scope_key_muni("Zona Norte")
    # stored under distinct cache files
    df = _frame()
    agg.save_scope("u1", agg.scope_key_city(), agg.compute_scope(df))
    agg.save_scope("u1", agg.scope_key_muni("SANTANA"), agg.compute_scope(df))
    assert cache.get_cache_size("u1") > 0
