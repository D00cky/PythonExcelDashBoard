"""Report scope types shared by export routes and report builders."""

from __future__ import annotations

from enum import Enum


class ReportScope(Enum):
    """Supported report scopes."""

    CITY = "city"
    ZONE = "zone"
    MUNICIPALITY = "municipality"


def parse_scope(value: str | None) -> ReportScope:
    """Parse a form/query value into a ``ReportScope``.

    Invalid or missing values default to city scope, matching the dashboard's
    top-level view.
    """
    try:
        return ReportScope(value or ReportScope.CITY.value)
    except ValueError:
        return ReportScope.CITY
