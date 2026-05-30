from app.core.templates.pimentas import PimentasTemplate
from tests.fixtures.pimentas_minimal import make_minimal_pimentas


def test_extract_inspections_concatenates_service_sheets_with_team_and_tss(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=True)

    df = PimentasTemplate().extract_inspections(path)

    assert set(df["service"].unique()) == {"ÁGUA", "ESGOTO"}
    assert {"team", "tss", "service"} <= set(df.columns)
    # 5 ÁGUA + 3 ESGOTO inspection rows
    assert len(df) == 8


def test_extract_inspections_dedupes_team_via_groupby(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=True)

    df = PimentasTemplate().extract_inspections(path)
    counts = df.groupby("team").size().sort_values(ascending=False)

    # JOSIAS appears 2× in ÁGUA + 1× in ESGOTO = 3
    # LAIS appears 2× in ÁGUA + 1× in ESGOTO = 3
    # FERNANDO appears 1× in ÁGUA + 1× in ESGOTO = 2
    assert counts.loc["JOSIAS ALMEIDA FRANCISCO"] == 3
    assert counts.loc["LAIS RAMOS SOBRAL"] == 3
    assert counts.loc["FERNANDO PEREIRA ASSIS DE LIMA MARTINS"] == 2


def test_extract_inspections_returns_empty_frame_when_no_service_sheets(tmp_path):
    path = make_minimal_pimentas(tmp_path)  # no inspections rows

    df = PimentasTemplate().extract_inspections(path)

    assert df.empty
    assert set(df.columns) >= {"team", "tss", "service"}


def test_extract_inspections_counts_conforme_and_nc_per_inspection(tmp_path):
    """conforme_count / nao_conforme_count are 0/1 per OS row.

    An inspection is conforme iff every stage cell is 'C' (NA / blank do
    not count as failures). Any 'NC' or 'SF' makes the whole row not-
    conforme. Summed over a team, the totals match the inspection count.
    """
    path = make_minimal_pimentas(tmp_path, with_inspections=True)

    df = PimentasTemplate().extract_inspections(path)

    assert "conforme_count" in df.columns
    assert "nao_conforme_count" in df.columns

    # ÁGUA / JOSIAS: row (C,C) → conforme, row (C,NC) → not conforme.
    josias_agua = df[(df["service"] == "ÁGUA") & (df["team"] == "JOSIAS ALMEIDA FRANCISCO")]
    assert josias_agua["conforme_count"].sum() == 1
    assert josias_agua["nao_conforme_count"].sum() == 1

    # ÁGUA / LAIS: row (NC,NC) → not conforme, row (C,SF) → not conforme.
    lais_agua = df[(df["service"] == "ÁGUA") & (df["team"] == "LAIS RAMOS SOBRAL")]
    assert lais_agua["conforme_count"].sum() == 0
    assert lais_agua["nao_conforme_count"].sum() == 2

    # ESGOTO / FERNANDO: 1 row (NC,NC) → not conforme.
    fern_esg = df[(df["service"] == "ESGOTO") & (df["team"].str.startswith("FERNANDO"))]
    assert fern_esg["conforme_count"].sum() == 0
    assert fern_esg["nao_conforme_count"].sum() == 1


def test_extract_inspections_conforme_plus_nao_conforme_equals_inspections(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=True)

    df = PimentasTemplate().extract_inspections(path)

    # Per-row counts must split the row into exactly one bucket, so summed
    # per team-service they equal the number of inspections for that group.
    grouped = df.groupby(["service", "team"]).agg(
        inspecoes=("team", "size"),
        conforme=("conforme_count", "sum"),
        nao_conforme=("nao_conforme_count", "sum"),
    )
    assert ((grouped["conforme"] + grouped["nao_conforme"]) == grouped["inspecoes"]).all()


def test_extract_inspections_includes_start_date_per_row(tmp_path):
    import pandas as pd

    path = make_minimal_pimentas(tmp_path, with_inspections=True)

    df = PimentasTemplate().extract_inspections(path)

    assert "start_date" in df.columns
    assert pd.api.types.is_datetime64_any_dtype(df["start_date"])
    # Earliest in fixture: 2026-03-05 (ÁGUA row 1). Latest: 2026-03-29 (ESGOTO row 3).
    assert df["start_date"].min() == pd.Timestamp("2026-03-05")
    assert df["start_date"].max() == pd.Timestamp("2026-03-29")
