from datetime import datetime

import pandas as pd
from openpyxl import Workbook

from app.core.templates.sabesp_pimentas import (
    SabespPimentasTemplate,
    top_observations,
)


def _write_failures_xlsx(tmp_path):
    """ÁGUA sheet with 3 NC and 2 SF outcomes plus observation columns."""
    wb = Workbook()
    wb.active.title = "CAPA"
    wb.create_sheet("DADOS - PIMENTAS")
    ws = wb.create_sheet("ÁGUA")
    ws["A1"] = "EQUIPE"
    ws["B1"] = "Descrição TSS"
    ws["C1"] = "FACHADA"
    ws["D1"] = "Observação FACHADA"
    ws["E1"] = "SINALIZAÇÃO"
    ws["F1"] = "Observação SINALIZAÇÃO"
    ws["G1"] = "Data Início Execução"
    rows = [
        ("ALICE", "T1", "NC", "telhado quebrado", "C", None, datetime(2026, 3, 5)),
        (
            "ALICE",
            "T1",
            "NC",
            "telhado quebrado",
            "SF",
            "câmera descarregada",
            datetime(2026, 3, 6),
        ),
        ("BOB", "T2", "C", None, "NC", "placa caída", datetime(2026, 3, 7)),
        ("BOB", "T2", "SF", "celular sem espaço", "C", None, datetime(2026, 3, 8)),
        ("ALICE", "T1", "C", None, "C", None, datetime(2026, 3, 9)),
    ]
    for i, (team, tss, fa, fa_obs, sn, sn_obs, dt) in enumerate(rows, start=2):
        ws[f"A{i}"], ws[f"B{i}"], ws[f"C{i}"], ws[f"D{i}"] = team, tss, fa, fa_obs
        ws[f"E{i}"], ws[f"F{i}"], ws[f"G{i}"] = sn, sn_obs, dt
    path = tmp_path / "failures.xlsx"
    wb.save(path)
    return path


def test_extract_stage_failures_returns_one_row_per_failing_cell(tmp_path):
    path = _write_failures_xlsx(tmp_path)

    failures = SabespPimentasTemplate().extract_stage_failures(path)

    # 3 NC + 2 SF across the two stage columns = 5 failures total.
    assert len(failures) == 5
    assert set(failures["code"]) == {"NC", "SF"}
    assert set(failures["stage"]) == {"FACHADA", "SINALIZAÇÃO"}
    assert set(failures["team"]) == {"ALICE", "BOB"}


def test_extract_stage_failures_pairs_observation_with_correct_stage(tmp_path):
    path = _write_failures_xlsx(tmp_path)

    failures = SabespPimentasTemplate().extract_stage_failures(path)

    fa_nc = failures[(failures["stage"] == "FACHADA") & (failures["code"] == "NC")]
    assert sorted(fa_nc["observation"].tolist()) == ["telhado quebrado", "telhado quebrado"]
    sn_nc = failures[(failures["stage"] == "SINALIZAÇÃO") & (failures["code"] == "NC")]
    assert sn_nc["observation"].tolist() == ["placa caída"]


def test_top_observations_returns_most_common_for_code(tmp_path):
    path = _write_failures_xlsx(tmp_path)
    failures = SabespPimentasTemplate().extract_stage_failures(path)

    top_nc = top_observations(failures, "NC")
    assert top_nc[0] == {
        "observation": "telhado quebrado",
        "stages": "FACHADA",
        "count": 2,
    }

    top_sf = top_observations(failures, "SF")
    # SF observations: 'câmera descarregada' (FACHADA, but cell is SF) +
    # 'celular sem espaço' (paired with FACHADA SF).
    assert {r["observation"] for r in top_sf} == {
        "câmera descarregada",
        "celular sem espaço",
    }


def test_top_observations_skips_empty_and_returns_empty_for_no_failures():
    empty = pd.DataFrame(
        columns=["service", "stage", "code", "observation", "team", "tss", "start_date"]
    )
    assert top_observations(empty, "NC") == []


def test_build_top_failing_stages_handles_empty_df():
    fig = SabespPimentasTemplate().build_top_failing_stages(pd.DataFrame(columns=["stage", "code"]))
    assert fig.layout.annotations[0].text.startswith("Sem etapas")


def test_build_worst_teams_filters_small_samples():
    df = pd.DataFrame(
        {
            "team": ["A"] * 5 + ["B"] * 2,
            "nao_conforme_count": [1, 1, 1, 0, 0, 1, 1],
        }
    )

    fig = SabespPimentasTemplate().build_worst_teams(df, min_inspections=3)

    # Only team A meets the 3-inspection threshold.
    bar = fig.data[0]
    assert list(bar.y) == ["A"]
