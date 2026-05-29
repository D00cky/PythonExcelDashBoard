from app.core.templates import recognize
from app.core.templates.sabesp_pimentas import SabespPimentasTemplate


def test_recognize_returns_sabesp_template_when_signature_matches():
    sheets = ["CAPA", "DADOS - PIMENTAS", "ÁGUA", "ESGOTO", "Sheet17"]

    result = recognize(sheets)

    assert isinstance(result, SabespPimentasTemplate)


def test_recognize_returns_none_for_unknown_signature():
    sheets = ["Sheet1", "Sheet2", "Sheet3"]

    assert recognize(sheets) is None


def test_recognize_returns_none_when_only_one_service_sheet_present():
    sheets = ["CAPA", "DADOS - PIMENTAS", "ÁGUA"]

    assert recognize(sheets) is None


def test_recognize_matches_despite_trailing_whitespace_in_sheet_names():
    # The real SABESP file has 'EQUIPES - CAVALETE ' with a trailing space;
    # the same kind of typo could appear on any sheet. Recognition must tolerate it.
    sheets = ["CAPA ", " DADOS - PIMENTAS", "ÁGUA ", "ESGOTO"]

    result = recognize(sheets)

    assert isinstance(result, SabespPimentasTemplate)
