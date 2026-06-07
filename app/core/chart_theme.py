"""Plotly chart themes for dashboard and report rendering.

Public API:
    register_templates() -> None

The dark template is used by the interactive dashboard; the light template is
reserved for print/report exports where transparent dark surfaces do not work.
"""

from __future__ import annotations

import plotly.graph_objects as go
import plotly.io as pio

POLO_TEMPLATE = go.layout.Template(
    layout={
        "font": {"family": "DM Sans, sans-serif", "size": 13, "color": "#b0b0b8"},
        "paper_bgcolor": "rgba(0,0,0,0)",
        "plot_bgcolor": "rgba(0,0,0,0)",
        "colorway": [
            "#F97316",
            "#0EA5E9",
            "#22c55e",
            "#A855F7",
            "#22D3EE",
            "#EAB308",
            "#EF4444",
        ],
        "xaxis": {
            "gridcolor": "rgba(255,255,255,0.06)",
            "zerolinecolor": "rgba(255,255,255,0.06)",
            "tickfont": {"size": 11},
        },
        "yaxis": {
            "gridcolor": "rgba(255,255,255,0.06)",
            "zerolinecolor": "rgba(255,255,255,0.06)",
            "tickfont": {"size": 11},
        },
        "margin": {"l": 50, "r": 20, "t": 50, "b": 40},
        "hoverlabel": {
            "bgcolor": "rgba(15,15,20,0.95)",
            "bordercolor": "rgba(249,115,22,0.3)",
            "font": {"family": "DM Sans", "size": 12, "color": "#e0e0e0"},
        },
        "bargap": 0.3,
        "legend": {
            "bgcolor": "rgba(0,0,0,0)",
            "font": {"size": 11},
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "right",
            "x": 1,
        },
    }
)

POLO_TEMPLATE_LIGHT = go.layout.Template(
    layout={
        "font": {"family": "DM Sans, sans-serif", "size": 13, "color": "#1a1a1a"},
        "paper_bgcolor": "#ffffff",
        "plot_bgcolor": "#ffffff",
        "colorway": [
            "#d45a0a",
            "#0284c7",
            "#16a34a",
            "#9333ea",
            "#0891b2",
            "#ca8a04",
            "#dc2626",
        ],
        "xaxis": {"gridcolor": "#e5e7eb", "tickfont": {"size": 11, "color": "#444"}},
        "yaxis": {"gridcolor": "#e5e7eb", "tickfont": {"size": 11, "color": "#444"}},
        "margin": {"l": 50, "r": 20, "t": 50, "b": 40},
    }
)


def register_templates() -> None:
    """Register Polo Plotly templates and set the dashboard default."""
    pio.templates["polo_dark"] = POLO_TEMPLATE
    pio.templates["polo_light"] = POLO_TEMPLATE_LIGHT
    pio.templates.default = "polo_dark"
