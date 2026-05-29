from app.core.templates.sabesp_pimentas import SabespPimentasTemplate
from tests.fixtures.sabesp_minimal import make_minimal_sabesp


def test_extract_inspections_concatenates_service_sheets_with_team_and_tss(tmp_path):
    path = make_minimal_sabesp(tmp_path, with_inspections=True)

    df = SabespPimentasTemplate().extract_inspections(path)

    assert set(df["service"].unique()) == {"ÁGUA", "ESGOTO"}
    assert {"team", "tss", "service"} <= set(df.columns)
    # 5 ÁGUA + 3 ESGOTO inspection rows
    assert len(df) == 8


def test_extract_inspections_dedupes_team_via_groupby(tmp_path):
    path = make_minimal_sabesp(tmp_path, with_inspections=True)

    df = SabespPimentasTemplate().extract_inspections(path)
    counts = df.groupby("team").size().sort_values(ascending=False)

    # JOSIAS appears 2× in ÁGUA + 1× in ESGOTO = 3
    # LAIS appears 2× in ÁGUA + 1× in ESGOTO = 3
    # FERNANDO appears 1× in ÁGUA + 1× in ESGOTO = 2
    assert counts.loc["JOSIAS ALMEIDA FRANCISCO"] == 3
    assert counts.loc["LAIS RAMOS SOBRAL"] == 3
    assert counts.loc["FERNANDO PEREIRA ASSIS DE LIMA MARTINS"] == 2


def test_extract_inspections_returns_empty_frame_when_no_service_sheets(tmp_path):
    path = make_minimal_sabesp(tmp_path)  # no inspections rows

    df = SabespPimentasTemplate().extract_inspections(path)

    assert df.empty
    assert set(df.columns) >= {"team", "tss", "service"}


def test_extract_inspections_counts_conforme_and_nc_per_row(tmp_path):
    path = make_minimal_sabesp(tmp_path, with_inspections=True)

    df = SabespPimentasTemplate().extract_inspections(path)

    assert "conforme_count" in df.columns
    assert "nao_conforme_count" in df.columns

    # ÁGUA / JOSIAS: row 1 (C,C) + row 2 (C,NC) = 3C, 1NC
    josias_agua = df[(df["service"] == "ÁGUA") & (df["team"] == "JOSIAS ALMEIDA FRANCISCO")]
    assert josias_agua["conforme_count"].sum() == 3
    assert josias_agua["nao_conforme_count"].sum() == 1

    # ÁGUA / LAIS: row 3 (NC,NC) + row 4 (C,SF) = 1C, 2NC
    lais_agua = df[(df["service"] == "ÁGUA") & (df["team"] == "LAIS RAMOS SOBRAL")]
    assert lais_agua["conforme_count"].sum() == 1
    assert lais_agua["nao_conforme_count"].sum() == 2

    # ESGOTO / FERNANDO: 1 row (NC,NC) = 0C, 2NC
    fern_esg = df[(df["service"] == "ESGOTO") & (df["team"].str.startswith("FERNANDO"))]
    assert fern_esg["conforme_count"].sum() == 0
    assert fern_esg["nao_conforme_count"].sum() == 2
