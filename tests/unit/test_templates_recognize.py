from app.core.templates import recognize
from app.core.templates.sabesp_pimentas import SabespPimentasTemplate


def test_recognize_returns_sabesp_template_when_signature_matches():
    sheets = ["CAPA", "DADOS - PIMENTAS", "ÁGUA", "ESGOTO", "Sheet17"]

    result = recognize(sheets)

    assert isinstance(result, SabespPimentasTemplate)
