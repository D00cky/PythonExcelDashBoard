from app.core.templates import recognize
from app.core.templates.pimentas import PimentasTemplate


def test_recognize_returns_pimentas_template_when_signature_matches():
    sheets = ["CAPA", "DADOS - PIMENTAS", "ÁGUA", "ESGOTO", "Sheet17"]

    result = recognize(sheets)

    assert isinstance(result, PimentasTemplate)


def test_recognize_returns_none_for_unknown_signature():
    sheets = ["Sheet1", "Sheet2", "Sheet3"]

    assert recognize(sheets) is None


def test_recognize_returns_none_when_only_one_service_sheet_present():
    sheets = ["CAPA", "DADOS - PIMENTAS", "ÁGUA"]

    assert recognize(sheets) is None


def test_recognize_matches_despite_trailing_whitespace_in_sheet_names():
    # The real file has 'EQUIPES - CAVALETE ' with a trailing space;
    # the same kind of typo could appear on any sheet. Recognition must tolerate it.
    sheets = ["CAPA ", " DADOS - PIMENTAS", "ÁGUA ", "ESGOTO"]

    result = recognize(sheets)

    assert isinstance(result, PimentasTemplate)


def test_recognize_matches_any_polo_via_dados_prefix():
    # The same template ships per polo (region). The DADOS sheet name
    # carries the polo: 'DADOS - PIRITUBA', 'DADOS - SANTANA', etc. The
    # recognizer must accept all of them.
    for polo in ("PIRITUBA", "SANTANA", "GOPOÚVA", "FREGUESIA DO Ó", "EXTREMO NORTE"):
        sheets = ["CAPA", f"DADOS - {polo}", "ÁGUA", "ESGOTO", "CAVALETE", "REPOSIÇÃO"]
        result = recognize(sheets)
        assert isinstance(result, PimentasTemplate), f"failed for polo {polo!r}"
        assert result.data_sheet_name == f"DADOS - {polo}"
        assert result.polo_name == polo


def test_recognize_returns_none_when_no_dados_sheet_present():
    sheets = ["CAPA", "ÁGUA", "ESGOTO", "CAVALETE", "REPOSIÇÃO"]

    assert recognize(sheets) is None
