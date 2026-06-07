import plotly.io as pio

from app.core.chart_theme import register_templates


def test_chart_theme_registers_dark_and_light_templates():
    register_templates()

    assert "polo_dark" in pio.templates
    assert "polo_light" in pio.templates
    assert pio.templates.default == "polo_dark"
    assert pio.templates["polo_dark"].layout.paper_bgcolor == "rgba(0,0,0,0)"
    assert pio.templates["polo_light"].layout.paper_bgcolor == "#ffffff"
