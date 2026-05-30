from datetime import datetime

import pandas as pd

from app.routes.main import (
    _ic_rows_from_inspections,
    _iqs_overall_from_inspections,
    _iqs_rows_from_inspections,
    _swap_day_month,
)


def _df(rows: list[dict]) -> pd.DataFrame:
    cols = [
        "team",
        "tss",
        "service",
        "conforme_count",
        "nao_conforme_count",
        "photo_conforme",
        "photo_nc",
        "photo_sf",
        "photo_total",
        "start_date",
    ]
    return pd.DataFrame(rows, columns=cols)


def test_iqs_rows_from_inspections_sums_photos_per_service():
    df = _df(
        [
            {
                "team": "A",
                "tss": "T1",
                "service": "ÁGUA",
                "conforme_count": 1,
                "nao_conforme_count": 0,
                "photo_conforme": 5,
                "photo_nc": 0,
                "photo_sf": 0,
                "photo_total": 5,
                "start_date": pd.NaT,
            },
            {
                "team": "B",
                "tss": "T2",
                "service": "ÁGUA",
                "conforme_count": 0,
                "nao_conforme_count": 1,
                "photo_conforme": 2,
                "photo_nc": 1,
                "photo_sf": 1,
                "photo_total": 4,
                "start_date": pd.NaT,
            },
        ]
    )

    rows = _iqs_rows_from_inspections(df, ["ÁGUA"])

    assert len(rows) == 1
    row = rows[0]
    assert row.name == "Água"
    assert row.fotos_avaliadas == 9
    assert row.fotos_conforme == 7
    # NC + SF lumped per CAPA convention (both are failures).
    assert row.fotos_nc == 2
    assert row.conforme_pct == 7 / 9


def test_iqs_rows_skips_services_with_no_photos():
    df = _df(
        [
            {
                "team": "A",
                "tss": "T",
                "service": "ÁGUA",
                "conforme_count": 0,
                "nao_conforme_count": 0,
                "photo_conforme": 0,
                "photo_nc": 0,
                "photo_sf": 0,
                "photo_total": 0,
                "start_date": pd.NaT,
            }
        ]
    )

    assert _iqs_rows_from_inspections(df, ["ÁGUA"]) == []


def test_ic_rows_from_inspections_counts_conforme_os():
    df = _df(
        [
            {
                "team": "A",
                "tss": "T",
                "service": "ÁGUA",
                "conforme_count": 1,
                "nao_conforme_count": 0,
                "photo_conforme": 3,
                "photo_nc": 0,
                "photo_sf": 0,
                "photo_total": 3,
                "start_date": pd.NaT,
            },
            {
                "team": "B",
                "tss": "T",
                "service": "ÁGUA",
                "conforme_count": 0,
                "nao_conforme_count": 1,
                "photo_conforme": 1,
                "photo_nc": 2,
                "photo_sf": 0,
                "photo_total": 3,
                "start_date": pd.NaT,
            },
        ]
    )

    rows = _ic_rows_from_inspections(df, ["ÁGUA"])

    assert len(rows) == 1
    assert rows[0].name == "Água"
    assert rows[0].ic_pct == 0.5
    assert rows[0].lvs == 2


def test_iqs_overall_returns_none_for_empty_df():
    assert _iqs_overall_from_inspections(_df([])) is None


def test_swap_day_month_only_targets_dominant_month():
    # Target = month 5 (unambiguous May 25 pins it). May 12 stays put;
    # Jan 5 swaps to May 1.
    df = _df(
        [
            {
                "team": "x",
                "tss": "x",
                "service": "ÁGUA",
                "conforme_count": 1,
                "nao_conforme_count": 0,
                "photo_conforme": 0,
                "photo_nc": 0,
                "photo_sf": 0,
                "photo_total": 0,
                "start_date": datetime(2026, 1, 5),
            },
            {
                "team": "x",
                "tss": "x",
                "service": "ÁGUA",
                "conforme_count": 1,
                "nao_conforme_count": 0,
                "photo_conforme": 0,
                "photo_nc": 0,
                "photo_sf": 0,
                "photo_total": 0,
                "start_date": datetime(2026, 5, 12),
            },
            {
                "team": "x",
                "tss": "x",
                "service": "ÁGUA",
                "conforme_count": 1,
                "nao_conforme_count": 0,
                "photo_conforme": 0,
                "photo_nc": 0,
                "photo_sf": 0,
                "photo_total": 0,
                "start_date": datetime(2026, 5, 25),
            },
        ]
    )

    out = _swap_day_month(df)

    dates = out["start_date"].tolist()
    assert dates[0] == pd.Timestamp("2026-05-01")  # swapped
    assert dates[1] == pd.Timestamp("2026-05-12")  # left alone
    assert dates[2] == pd.Timestamp("2026-05-25")  # unambiguous, unchanged


def test_swap_day_month_no_op_when_no_unambiguous_dates():
    df = _df(
        [
            {
                "team": "x",
                "tss": "x",
                "service": "ÁGUA",
                "conforme_count": 1,
                "nao_conforme_count": 0,
                "photo_conforme": 0,
                "photo_nc": 0,
                "photo_sf": 0,
                "photo_total": 0,
                "start_date": datetime(2026, 5, 5),
            },
        ]
    )

    out = _swap_day_month(df)
    assert out["start_date"].iloc[0] == pd.Timestamp("2026-05-05")
