import plotly.graph_objects as go

from app.core.templates.sabesp_pimentas import (
    SabespPimentasTemplate,
    ServiceIC,
    ServiceIQS,
)

_ROWS = [
    ServiceIQS("Água", 182, 91, 91, 0.5, 0.5),
    ServiceIQS("Esgoto", 33, 16, 17, 0.4848, 0.5152),
    ServiceIQS("Cavalete", 73, 13, 60, 0.1780, 0.8220),
    ServiceIQS("Reposição", 31, 6, 25, 0.1935, 0.8065),
]


def test_build_service_iqs_bar_has_one_bar_per_service():
    fig = SabespPimentasTemplate().build_service_iqs_bar(_ROWS)

    assert isinstance(fig, go.Figure)
    assert len(fig.data) == 1
    assert tuple(fig.data[0].x) == ("Água", "Esgoto", "Cavalete", "Reposição")
    assert tuple(fig.data[0].y) == (0.5, 0.5152, 0.8220, 0.8065)


def test_build_photo_conformity_stacked_has_nc_and_conforme_traces():
    fig = SabespPimentasTemplate().build_photo_conformity_stacked(_ROWS)

    assert isinstance(fig, go.Figure)
    assert fig.layout.barmode == "stack"
    assert len(fig.data) == 2
    by_name = {trace.name: trace for trace in fig.data}
    assert tuple(by_name["Não Conforme"].y) == (91, 16, 13, 6)
    assert tuple(by_name["Conforme"].y) == (91, 17, 60, 25)
    assert tuple(by_name["Conforme"].x) == ("Água", "Esgoto", "Cavalete", "Reposição")


def test_build_ic_bar_has_one_bar_per_service():
    rows = [
        ServiceIC("Água", 1.0, 100),
        ServiceIC("Esgoto", 0.5, 50),
        ServiceIC("Reposição", 0.1, 1),
    ]

    fig = SabespPimentasTemplate().build_ic_bar(rows)

    assert isinstance(fig, go.Figure)
    assert len(fig.data) == 1
    assert tuple(fig.data[0].x) == ("Água", "Esgoto", "Reposição")
    assert tuple(fig.data[0].y) == (1.0, 0.5, 0.1)


def test_build_team_service_stacked_orders_by_total_descending_and_dedupes():
    import pandas as pd

    df = pd.DataFrame(
        {
            "team": ["A", "A", "A", "B", "B", "C"],
            "tss": ["t"] * 6,
            "service": ["ÁGUA", "ESGOTO", "ÁGUA", "ÁGUA", "ESGOTO", "ÁGUA"],
        }
    )

    fig = SabespPimentasTemplate().build_team_service_stacked(df, top_n=3)

    assert isinstance(fig, go.Figure)
    assert fig.layout.barmode == "stack"
    # Horizontal bar: y is team, x is count. Highest count first (top of chart),
    # so plotly's y-axis (which puts the first item at the bottom) gets reversed.
    first = fig.data[0]
    assert tuple(first.y) == ("C", "B", "A")
    # Same y order across all traces (one per service present)
    for trace in fig.data:
        assert tuple(trace.y) == ("C", "B", "A")


def test_build_tss_distribution_top_n_horizontal_bar():
    import pandas as pd

    df = pd.DataFrame(
        {
            "team": ["t"] * 10,
            "tss": ["A", "A", "A", "B", "B", "C", "D", "D", "D", "D"],
            "service": ["ÁGUA"] * 10,
        }
    )

    fig = SabespPimentasTemplate().build_tss_distribution(df, top_n=3)

    assert isinstance(fig, go.Figure)
    assert len(fig.data) == 1
    trace = fig.data[0]
    assert trace.orientation == "h"
    # Top 3 TSS by count: D (4), A (3), B (2). Reversed for y-axis display.
    assert tuple(trace.y) == ("B", "A", "D")
    assert tuple(trace.x) == (2, 3, 4)
