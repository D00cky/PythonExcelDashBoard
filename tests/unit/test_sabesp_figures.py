import plotly.graph_objects as go

from app.core.templates.sabesp_pimentas import SabespPimentasTemplate, ServiceIQS

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
