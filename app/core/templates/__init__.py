from collections.abc import Iterable

from app.core.templates.sabesp_pimentas import SabespPimentasTemplate


def recognize(sheet_names: Iterable[str]) -> SabespPimentasTemplate | None:
    if SabespPimentasTemplate.matches(sheet_names):
        return SabespPimentasTemplate()
    return None
