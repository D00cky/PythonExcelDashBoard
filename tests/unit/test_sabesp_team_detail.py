import pandas as pd
from openpyxl import Workbook

from app.core.templates.sabesp_pimentas import (
    SabespPimentasTemplate,
    _stage_code_counts,
    _top_reason,
    _top_tss,
)
from tests.fixtures.sabesp_minimal import make_minimal_sabesp


def test_extract_team_detail_returns_empty_dict_when_team_absent(tmp_path):
    path = make_minimal_sabesp(tmp_path, with_inspections=True)

    detail = SabespPimentasTemplate().extract_team_detail(path, "NOBODY")

    assert detail == {}


def test_extract_team_detail_aggregates_per_service_stage_counts(tmp_path):
    path = make_minimal_sabesp(tmp_path, with_inspections=True)

    detail = SabespPimentasTemplate().extract_team_detail(path, "JOSIAS ALMEIDA FRANCISCO")

    assert set(detail) == {"ÁGUA", "ESGOTO"}
    # JOSIAS in ÁGUA: 2 inspections, (C,C) + (C,NC) → 3C, 1NC, 0SF, 0NA
    agua = detail["ÁGUA"]
    assert agua["inspecoes"] == 2
    assert agua["conforme"] == 3
    assert agua["nao_conforme"] == 1
    assert agua["sem_foto"] == 0
    assert agua["nao_avaliado"] == 0
    # Fixture has no Observação columns → no reasons can be extracted
    assert agua["top_nc_reason"] is None
    assert agua["top_sf_reason"] is None


def test_extract_team_detail_tss_summary_lists_distinct_services(tmp_path):
    path = make_minimal_sabesp(tmp_path, with_inspections=True)

    detail = SabespPimentasTemplate().extract_team_detail(path, "JOSIAS ALMEIDA FRANCISCO")

    assert dict(detail["ÁGUA"]["tss_summary"]) == {
        "TROCAR RAMAL DE ÁGUA PREVENTIVA": 1,
        "VAZAMENTO DE ÁGUA NO PASSEIO": 1,
    }


def test_extract_team_detail_trims_whitespace_in_team_name(tmp_path):
    path = make_minimal_sabesp(tmp_path, with_inspections=True)

    detail = SabespPimentasTemplate().extract_team_detail(
        path, "  JOSIAS ALMEIDA FRANCISCO  "
    )

    assert "ÁGUA" in detail


def test_extract_team_detail_top_nc_reason_picks_most_common_observation(tmp_path):
    path = _write_observations_xlsx(
        tmp_path,
        [
            ("ALICE", "SERV1", "NC", "telhado quebrado"),
            ("ALICE", "SERV1", "NC", "telhado quebrado"),
            ("ALICE", "SERV1", "NC", "muro pichado"),
            ("ALICE", "SERV1", "C", None),
            ("ALICE", "SERV1", "SF", "sem foto da fachada"),
        ],
    )

    detail = SabespPimentasTemplate().extract_team_detail(path, "ALICE")

    assert detail["ÁGUA"]["top_nc_reason"] == ("telhado quebrado", 2)
    assert detail["ÁGUA"]["top_sf_reason"] == ("sem foto da fachada", 1)


def test_top_reason_returns_most_common_string():
    assert _top_reason(["a", "b", "a", "c", "a"]) == ("a", 3)
    assert _top_reason([]) is None


def test_top_tss_returns_empty_when_column_missing():
    assert _top_tss(pd.DataFrame({"x": [1]}), None) == []


def test_top_tss_limits_to_top_n():
    sub = pd.DataFrame({"tss": ["a", "a", "b", "b", "c"]})
    assert _top_tss(sub, "tss", n=2) == [("a", 2), ("b", 2)]


def test_stage_code_counts_returns_empty_when_no_stage_cols():
    assert _stage_code_counts(pd.DataFrame(), []) == {}


def _write_observations_xlsx(tmp_path, rows):
    wb = Workbook()
    wb.active.title = "CAPA"
    wb.create_sheet("DADOS - PIMENTAS")
    ws = wb.create_sheet("ÁGUA")
    ws["A1"] = "EQUIPE"
    ws["B1"] = "Descrição TSS"
    ws["C1"] = "FACHADA"
    ws["D1"] = "Observação FACHADA"
    for i, (team, tss, stage, observation) in enumerate(rows, start=2):
        ws[f"A{i}"] = team
        ws[f"B{i}"] = tss
        ws[f"C{i}"] = stage
        ws[f"D{i}"] = observation
    path = tmp_path / "team_obs.xlsx"
    wb.save(path)
    return path
