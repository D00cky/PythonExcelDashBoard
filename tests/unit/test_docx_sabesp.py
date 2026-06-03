from io import BytesIO

import pytest
from docx import Document
from openpyxl import load_workbook

from app.core.exporters.docx_sabesp import (
    MensalContext,
    context_from_template,
    render_mensal,
)
from app.core.templates.pimentas import PimentasTemplate
from tests.fixtures.pimentas_minimal import make_minimal_pimentas


def _all_text(doc):
    chunks = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    chunks.append(p.text)
    return "\n".join(chunks)


def test_render_mensal_substitutes_polo_and_period():
    ctx = MensalContext(
        polo_label="Pimentas",
        periodo_inicio="06/05/2026",
        periodo_fim="12/05/2026",
        mes_extenso="Maio",
        ano="2026",
    )

    body = render_mensal(ctx)

    assert body[:2] == b"PK"  # .docx is a zip
    doc = Document(BytesIO(body))
    text = _all_text(doc)
    assert "Pimentas" in text
    assert "PIMENTAS" in text
    assert "06/05/2026" in text
    assert "12/05/2026" in text
    assert "Maio de 2026" in text
    # No source polo name should leak through.
    assert "Gopoúva" not in text
    assert "Gopouva" not in text
    assert "GOPOÚVA" not in text
    # No unrendered Jinja markers.
    assert "{{" not in text


def test_render_mensal_polo_label_upper_derived_from_polo_label():
    ctx = MensalContext(
        polo_label="Santana",
        periodo_inicio="01/06/2026",
        periodo_fim="07/06/2026",
        mes_extenso="Junho",
        ano="2026",
    )
    assert ctx.polo_label_upper == "SANTANA"


def test_context_from_template_uses_inspection_date_bounds(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=True)
    wb = load_workbook(path, data_only=True, read_only=True)
    template = PimentasTemplate.detect(wb.sheetnames)

    ctx = context_from_template(template, path)

    # Fixture dates span 2026-03-05 to 2026-03-29 — period text and month should reflect that.
    assert ctx.periodo_inicio == "05/03/2026"
    assert ctx.periodo_fim == "29/03/2026"
    assert ctx.mes_extenso == "Março"
    assert ctx.ano == "2026"
    assert ctx.polo_label == "Pimentas"


def test_context_from_template_handles_workbook_without_inspections(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=False)
    wb = load_workbook(path, data_only=True, read_only=True)
    template = PimentasTemplate.detect(wb.sheetnames)

    ctx = context_from_template(template, path)

    # No dated rows → blank period, but polo label still binds.
    assert ctx.polo_label == "Pimentas"
    assert ctx.periodo_inicio == ""
    assert ctx.periodo_fim == ""
    assert ctx.mes_extenso == ""


@pytest.mark.parametrize(
    "field",
    ["polo_label", "periodo_inicio", "periodo_fim", "mes_extenso", "ano"],
)
def test_render_dict_exposes_every_skeleton_placeholder(field):
    ctx = MensalContext(
        polo_label="X",
        periodo_inicio="i",
        periodo_fim="f",
        mes_extenso="m",
        ano="a",
    )
    assert field in ctx.as_render_dict()
    assert "polo_label_upper" in ctx.as_render_dict()
