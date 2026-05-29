from collections.abc import Iterable


class SabespPimentasTemplate:
    REQUIRED_SHEETS = frozenset({"CAPA", "DADOS - PIMENTAS"})
    SERVICE_SHEETS = frozenset({"ÁGUA", "ESGOTO", "CAVALETE", "REPOSIÇÃO"})
    MIN_SERVICES = 2

    @classmethod
    def matches(cls, sheet_names: Iterable[str]) -> bool:
        names = set(sheet_names)
        if not cls.REQUIRED_SHEETS.issubset(names):
            return False
        return len(cls.SERVICE_SHEETS & names) >= cls.MIN_SERVICES
