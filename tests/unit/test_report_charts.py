from pathlib import Path

import plotly.graph_objects as go

from app.core.aggregator import discover_polo_file, write_manifest
from app.core.report.charts import render_charts_for_scope
from app.core.report.scope import ReportScope
from tests.fixtures.pimentas_minimal import make_minimal_pimentas


def _make_batch(app, tmp_path) -> str:
    uuid = "reportcharts"
    batch_dir = Path(app.instance_path) / "uploads" / uuid
    batch_dir.mkdir(parents=True)
    pim = make_minimal_pimentas(
        batch_dir, polo="PIMENTAS", with_inspections=True, file_name="file_00.xlsx"
    )
    san = make_minimal_pimentas(
        batch_dir, polo="SANTANA", with_inspections=True, file_name="file_01.xlsx"
    )
    write_manifest(batch_dir, [discover_polo_file(pim), discover_polo_file(san)])
    return uuid


def test_render_charts_for_scope_writes_png_fallback_when_kaleido_fails(app, tmp_path, monkeypatch):
    uuid = _make_batch(app, tmp_path)

    def boom(self, *args, **kwargs):
        raise RuntimeError("kaleido unavailable")

    monkeypatch.setattr(go.Figure, "write_image", boom)
    with app.app_context():
        charts = render_charts_for_scope(
            uuid,
            ReportScope.CITY,
            None,
            "light",
            tmp_path / "charts",
            period="2026-W10",
        )

    assert set(charts) == {"chart_ic", "chart_iqs", "chart_volume"}
    assert all(
        path.exists() and path.read_bytes().startswith(b"\x89PNG") for path in charts.values()
    )
