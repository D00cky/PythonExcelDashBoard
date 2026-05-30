import plotly.graph_objects as go

from app.core.templates.pimentas import (
    PimentasTemplate,
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
    fig = PimentasTemplate().build_service_iqs_bar(_ROWS)

    assert isinstance(fig, go.Figure)
    assert len(fig.data) == 1
    assert tuple(fig.data[0].x) == ("Água", "Esgoto", "Cavalete", "Reposição")
    assert tuple(fig.data[0].y) == (0.5, 0.5152, 0.8220, 0.8065)


def test_build_photo_conformity_stacked_has_nc_and_conforme_traces():
    fig = PimentasTemplate().build_photo_conformity_stacked(_ROWS)

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

    fig = PimentasTemplate().build_ic_bar(rows)

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

    fig = PimentasTemplate().build_team_service_stacked(df, top_n=3)

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

    fig = PimentasTemplate().build_tss_distribution(df, top_n=3)

    assert isinstance(fig, go.Figure)
    assert len(fig.data) == 1
    trace = fig.data[0]
    assert trace.orientation == "h"
    # Top 3 TSS by count: D (4), A (3), B (2). Reversed for y-axis display.
    assert tuple(trace.y) == ("B", "A", "D")
    assert tuple(trace.x) == (2, 3, 4)


def _conformity_df():
    import pandas as pd

    return pd.DataFrame(
        {
            "team": ["A", "A", "B", "B", "C", "D"],
            "tss": ["x", "x", "y", "y", "z", "w"],
            "service": ["ÁGUA"] * 5 + ["ESGOTO"],
            "conforme_count": [2, 3, 1, 1, 3, 9],
            "nao_conforme_count": [1, 0, 2, 3, 1, 0],
        }
    )


def test_build_team_conformity_for_service_stacks_conforme_and_nc():
    fig = PimentasTemplate().build_team_conformity_for_service(
        _conformity_df(), service="ÁGUA", top_n=3
    )

    assert fig.layout.barmode == "stack"
    by_name = {t.name: t for t in fig.data}
    assert set(by_name) == {"Conforme", "Não Conforme"}
    # ÁGUA totals: A=5C/1NC=6, B=2C/5NC=7, C=3C/1NC=4
    # Ranked desc: B(7), A(6), C(4). Reversed for plotly y: C, A, B
    assert tuple(by_name["Conforme"].y) == ("C", "A", "B")
    assert tuple(by_name["Conforme"].x) == (3, 5, 2)
    assert tuple(by_name["Não Conforme"].x) == (1, 1, 5)


def test_build_tss_conformity_for_service_stacks_by_tss():
    fig = PimentasTemplate().build_tss_conformity_for_service(
        _conformity_df(), service="ÁGUA", top_n=3
    )

    by_name = {t.name: t for t in fig.data}
    # ÁGUA TSS totals: x=5C/1NC=6, y=2C/5NC=7, z=3C/1NC=4. Ranked: y,x,z. Reversed: z,x,y
    assert tuple(by_name["Conforme"].y) == ("z", "x", "y")
    assert tuple(by_name["Conforme"].x) == (3, 5, 2)
