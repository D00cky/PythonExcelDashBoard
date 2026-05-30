from collections.abc import Iterable

from app.core.templates.pimentas import PimentasTemplate


def recognize(sheet_names: Iterable[str]) -> PimentasTemplate | None:
    return PimentasTemplate.detect(sheet_names)
