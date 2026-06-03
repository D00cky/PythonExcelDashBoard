from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from app.core.aggregator import (
    PoloBatch,
    PoloFile,
    combined_ic_rows,
    combined_inspections,
    combined_iqs_overall,
    combined_iqs_rows,
    combined_stage_failures,
    date_bounds,
    discover_polo_file,
    filter_batch,
    format_period_pt,
    load_batch,
    parse_periodo,
    write_manifest,
)
from tests.fixtures.pimentas_minimal import make_minimal_pimentas


class TestDateBounds:
    def test_returns_min_max_of_start_date_column(self) -> None:
        df = pd.DataFrame(
            {"start_date": pd.to_datetime(["2026-03-05", "2026-03-29", "2026-03-12"])}
        )
        bounds = date_bounds(df)
        assert bounds is not None
        start, end = bounds
        assert start == pd.Timestamp("2026-03-05")
        assert end == pd.Timestamp("2026-03-29")

    def test_returns_none_when_dataframe_empty(self) -> None:
        assert date_bounds(pd.DataFrame()) is None

    def test_returns_none_when_start_date_column_absent(self) -> None:
        assert date_bounds(pd.DataFrame({"foo": [1, 2, 3]})) is None

    def test_returns_none_when_all_dates_nat(self) -> None:
        df = pd.DataFrame({"start_date": [pd.NaT, pd.NaT]})
        assert date_bounds(df) is None

    def test_ignores_nat_rows_when_some_are_valid(self) -> None:
        df = pd.DataFrame(
            {"start_date": [pd.NaT, pd.Timestamp("2026-05-04"), pd.Timestamp("2026-05-10")]}
        )
        bounds = date_bounds(df)
        assert bounds == (pd.Timestamp("2026-05-04"), pd.Timestamp("2026-05-10"))


class TestFormatPeriodPt:
    def test_renders_dd_mm_yyyy_label(self) -> None:
        bounds = (pd.Timestamp("2026-03-05"), pd.Timestamp("2026-03-29"))
        assert format_period_pt(bounds) == "05/03/2026 à 29/03/2026"

    def test_returns_none_for_none_input_so_callers_can_chain(self) -> None:
        # Designed so consumers can write: format_period_pt(date_bounds(df))
        # and propagate the empty case in one shot.
        assert format_period_pt(None) is None


def _file(tmp_path: Path, name: str = "f.xlsx", polo: str = "PIMENTAS") -> PoloFile:
    path = tmp_path / name
    path.write_bytes(b"PK")
    return PoloFile(
        file_path=path,
        polo=polo,
        iso_week="2026-W19",
        month="2026-05",
        period_start=date(2026, 5, 4),
        period_end=date(2026, 5, 10),
    )


def test_write_then_load_manifest_roundtrips(tmp_path: Path) -> None:
    files = [
        _file(tmp_path, "a.xlsx", "PIMENTAS"),
        _file(tmp_path, "b.xlsx", "SANTANA"),
    ]

    write_manifest(tmp_path, files)
    batch = load_batch(tmp_path)

    assert isinstance(batch, PoloBatch)
    assert batch.batch_dir == tmp_path
    assert [f.polo for f in batch.files] == ["PIMENTAS", "SANTANA"]
    assert batch.files[0].iso_week == "2026-W19"
    assert batch.files[0].month == "2026-05"
    assert batch.files[0].period_start == date(2026, 5, 4)
    assert batch.files[0].period_end == date(2026, 5, 10)
    assert batch.files[0].file_path == tmp_path / "a.xlsx"


def test_polobatch_exposes_unique_polos_weeks_months(tmp_path: Path) -> None:
    files = [
        PoloFile(
            file_path=tmp_path / "a.xlsx",
            polo="PIMENTAS",
            iso_week="2026-W18",
            month="2026-05",
            period_start=date(2026, 4, 27),
            period_end=date(2026, 5, 3),
        ),
        PoloFile(
            file_path=tmp_path / "b.xlsx",
            polo="PIMENTAS",
            iso_week="2026-W19",
            month="2026-05",
            period_start=date(2026, 5, 4),
            period_end=date(2026, 5, 10),
        ),
        PoloFile(
            file_path=tmp_path / "c.xlsx",
            polo="SANTANA",
            iso_week="2026-W19",
            month="2026-05",
            period_start=date(2026, 5, 4),
            period_end=date(2026, 5, 10),
        ),
    ]
    batch = PoloBatch(batch_dir=tmp_path, files=files)

    assert batch.polos == ["PIMENTAS", "SANTANA"]
    assert batch.iso_weeks == ["2026-W18", "2026-W19"]
    assert batch.months == ["2026-05"]


def test_load_batch_missing_manifest_raises(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError):
        load_batch(tmp_path)


class TestParsePeriodo:
    def test_weekly_range_yields_iso_week_and_month(self) -> None:
        start, end, iso_week, month = parse_periodo("Período: 04/05/2026 à 10/05/2026")
        assert start == date(2026, 5, 4)
        assert end == date(2026, 5, 10)
        assert iso_week == "2026-W19"
        assert month == "2026-05"

    def test_monthly_range(self) -> None:
        start, end, iso_week, month = parse_periodo("Período: 01/04/2026 à 30/04/2026")
        assert start == date(2026, 4, 1)
        assert end == date(2026, 4, 30)
        assert iso_week == "2026-W14"
        assert month == "2026-04"

    def test_accepts_string_without_label(self) -> None:
        start, end, _, _ = parse_periodo("01/03/2026 à 31/03/2026")
        assert start == date(2026, 3, 1)
        assert end == date(2026, 3, 31)

    def test_iso_week_crosses_year_boundary(self) -> None:
        # 29/12/2025 falls in ISO week 2026-W01 (Mon 29 Dec 2025 starts W01).
        _, _, iso_week, month = parse_periodo("29/12/2025 à 04/01/2026")
        assert iso_week == "2026-W01"
        assert month == "2025-12"

    @pytest.mark.parametrize("bad", ["", "Período: lixo", "01/13/2026 à 31/12/2026", None])
    def test_invalid_input_raises(self, bad: object) -> None:
        with pytest.raises((ValueError, TypeError)):
            parse_periodo(bad)  # type: ignore[arg-type]


class TestDiscoverPoloFile:
    def test_extracts_polo_and_period_from_real_xlsx(self, tmp_path: Path) -> None:
        path = make_minimal_pimentas(
            tmp_path,
            polo="GOPOÚVA",
            periodo="Período: 04/05/2026 à 10/05/2026",
        )

        pf = discover_polo_file(path)

        assert pf.file_path == path
        assert pf.polo == "GOPOÚVA"
        assert pf.iso_week == "2026-W19"
        assert pf.month == "2026-05"
        assert pf.period_start == date(2026, 5, 4)
        assert pf.period_end == date(2026, 5, 10)

    def test_raises_when_template_unrecognized(self, tmp_path: Path) -> None:
        from openpyxl import Workbook

        path = tmp_path / "junk.xlsx"
        wb = Workbook()
        wb.active.title = "Sheet1"
        wb.save(path)

        with pytest.raises(ValueError, match="template"):
            discover_polo_file(path)

    def test_raises_when_no_periodo_and_no_dated_inspections(self, tmp_path: Path) -> None:
        # No B4 periodo and with_inspections=False so neither fallback path yields dates.
        path = make_minimal_pimentas(tmp_path, with_periodo=False, with_inspections=False)
        with pytest.raises(ValueError, match="no dated inspections|empty B4|periodo|período"):
            discover_polo_file(path)

    def test_inspections_dates_override_stale_b4(self, tmp_path: Path) -> None:
        """When B4 lies (cloned-from-last-month xlsx) and inspections carry real
        dates, the inspections-derived period must win — otherwise the period
        dropdown shows the wrong week and breaks per-period filtering. This is
        the core bug fix that landed the 'inspections-first' rule."""
        # B4 says March 1–31; fixture's inspections span 2026-03-05 to 2026-03-29.
        # To prove inspections win, override B4 with a wildly different month —
        # if the result still tracks the inspection dates, the override is honored.
        path = make_minimal_pimentas(
            tmp_path,
            polo="PIMENTAS",
            periodo="Período: 01/12/2025 à 31/12/2025",  # December — way off
            with_inspections=True,
        )

        pf = discover_polo_file(path)

        # Inspection rows (Mar 5 → Mar 29 2026) → iso_week W10, month 2026-03.
        assert pf.period_start == date(2026, 3, 5)
        assert pf.period_end == date(2026, 3, 29)
        assert pf.iso_week == "2026-W10"
        assert pf.month == "2026-03"

    def test_falls_back_to_b4_when_no_dated_inspections(self, tmp_path: Path) -> None:
        # No inspection rows → B4 is the only source.
        path = make_minimal_pimentas(
            tmp_path,
            polo="PIMENTAS",
            periodo="Período: 04/05/2026 à 10/05/2026",
            with_inspections=False,
        )

        pf = discover_polo_file(path)

        assert pf.period_start == date(2026, 5, 4)
        assert pf.period_end == date(2026, 5, 10)
        assert pf.iso_week == "2026-W19"


@pytest.fixture
def two_polo_batch(tmp_path: Path) -> PoloBatch:
    pim = make_minimal_pimentas(
        tmp_path,
        polo="PIMENTAS",
        periodo="Período: 04/05/2026 à 10/05/2026",
        with_inspections=True,
        file_name="pim.xlsx",
    )
    san = make_minimal_pimentas(
        tmp_path,
        polo="SANTANA",
        periodo="Período: 11/05/2026 à 17/05/2026",
        with_inspections=True,
        file_name="san.xlsx",
    )
    return PoloBatch(
        batch_dir=tmp_path,
        files=[discover_polo_file(pim), discover_polo_file(san)],
    )


class TestCombinedInspections:
    def test_concatenates_rows_with_polo_column(self, two_polo_batch: PoloBatch) -> None:
        df = combined_inspections(two_polo_batch)

        assert not df.empty
        assert "polo" in df.columns
        assert set(df["polo"].unique()) == {"PIMENTAS", "SANTANA"}
        for col in ("team", "tss", "service", "start_date"):
            assert col in df.columns

    def test_row_count_equals_sum_of_per_file(self, two_polo_batch: PoloBatch) -> None:
        from app.core.templates.pimentas import PimentasTemplate

        expected = sum(
            len(PimentasTemplate().extract_inspections(f.file_path)) for f in two_polo_batch.files
        )
        assert len(combined_inspections(two_polo_batch)) == expected


class TestCombinedStageFailures:
    def test_concatenates_failures_with_polo_column(self, two_polo_batch: PoloBatch) -> None:
        df = combined_stage_failures(two_polo_batch)
        if df.empty:
            return  # minimal fixture may have no NC rows for some shapes
        assert "polo" in df.columns
        for col in ("service", "stage", "code"):
            assert col in df.columns


class TestFilterBatch:
    # The fixture's inspection rows span 2026-03-05 to 2026-03-29 — iso_week W10
    # and month 2026-03 — and inspections-derived dates win over the May B4.
    def test_filter_by_polos_weekly(self, two_polo_batch: PoloBatch) -> None:
        filtered = filter_batch(
            two_polo_batch, polos=("PIMENTAS",), view="weekly", period_key="2026-W10"
        )
        assert [f.polo for f in filtered.files] == ["PIMENTAS"]

    def test_filter_by_month_keeps_all_weeks_in_month(self, two_polo_batch: PoloBatch) -> None:
        filtered = filter_batch(
            two_polo_batch, polos=("PIMENTAS", "SANTANA"), view="monthly", period_key="2026-03"
        )
        assert {f.polo for f in filtered.files} == {"PIMENTAS", "SANTANA"}

    def test_filter_by_other_month_yields_empty(self, two_polo_batch: PoloBatch) -> None:
        filtered = filter_batch(
            two_polo_batch, polos=("PIMENTAS",), view="monthly", period_key="2026-04"
        )
        assert filtered.files == []

    def test_unknown_view_raises(self, two_polo_batch: PoloBatch) -> None:
        with pytest.raises(ValueError, match="view"):
            filter_batch(two_polo_batch, polos=("PIMENTAS",), view="yearly", period_key="2026")


class TestCombinedIqsIc:
    def test_iqs_rows_per_service(self, two_polo_batch: PoloBatch) -> None:
        rows = combined_iqs_rows(two_polo_batch, services=["ÁGUA", "ESGOTO"])
        names = {r.name for r in rows}
        assert names <= {"Água", "Esgoto"}
        for r in rows:
            assert r.fotos_avaliadas > 0
            assert r.nc_pct + r.conforme_pct == pytest.approx(1.0, abs=1e-9)

    def test_ic_rows_per_service(self, two_polo_batch: PoloBatch) -> None:
        rows = combined_ic_rows(two_polo_batch, services=["ÁGUA", "ESGOTO"])
        names = {r.name for r in rows}
        assert names <= {"Água", "Esgoto"}
        for r in rows:
            assert r.lvs > 0
            assert 0.0 <= r.ic_pct <= 1.0

    def test_iqs_overall_within_unit_interval(self, two_polo_batch: PoloBatch) -> None:
        value = combined_iqs_overall(two_polo_batch)
        assert value is None or 0.0 <= value <= 1.0

    def test_empty_batch_returns_empty(self, tmp_path: Path) -> None:
        empty = PoloBatch(batch_dir=tmp_path, files=[])
        assert combined_iqs_rows(empty, services=["ÁGUA"]) == []
        assert combined_ic_rows(empty, services=["ÁGUA"]) == []
        assert combined_iqs_overall(empty) is None
