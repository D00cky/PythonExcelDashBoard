from io import BytesIO

import pytest
from docx import Document
from openpyxl import load_workbook

from app.core.exporters.docx_sabesp import (
    EquipeMember,
    IndiceRow,
    MensalContext,
    SemanalContext,
    context_from_template,
    indice_rows_from_inspections,
    render_mensal,
    render_semanal,
    semanal_context_from_template,
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
    ["polo_label", "periodo_inicio", "periodo_fim", "mes_extenso", "ano", "equipe", "indice"],
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


def _indice_table(doc):
    return next(t for t in doc.tables if "ÍNDICE TECNOLÓGICO" in t.rows[0].cells[0].text)


def test_render_mensal_expands_equipe_loop_to_supplied_members():
    ctx = MensalContext(
        polo_label="Pimentas",
        periodo_inicio="06/05/2026",
        periodo_fim="12/05/2026",
        mes_extenso="Maio",
        ano="2026",
        equipe=(
            EquipeMember(id="I", role="Tecnólogo", name="Foo Bar"),
            EquipeMember(id="II", role="Engenheiro", name="Baz Qux"),
        ),
    )

    doc = Document(BytesIO(render_mensal(ctx)))
    equipe_lines = [p.text for p in doc.paragraphs if p.text.startswith("EQUIPE ")]

    # Only the section-2 EQUIPE I/II members render — section-5/6/etc. mention
    # "EQUIPE I, IV e V" as part of a heading and don't start with "EQUIPE " (note
    # the trailing space), so they aren't counted here.
    equipe_section_lines = [line for line in equipe_lines if " – " in line]
    assert equipe_section_lines == [
        "EQUIPE I – Tecnólogo Foo Bar",
        "EQUIPE II – Engenheiro Baz Qux",
    ]


def test_render_mensal_with_empty_equipe_keeps_assistentes_static():
    ctx = MensalContext(
        polo_label="Pimentas",
        periodo_inicio="06/05/2026",
        periodo_fim="12/05/2026",
        mes_extenso="Maio",
        ano="2026",
        equipe=(),
    )

    doc = Document(BytesIO(render_mensal(ctx)))
    para_texts = [p.text for p in doc.paragraphs]

    # ASSISTENTES line is static — present regardless of equipe length.
    assert any(line.startswith("ASSISTENTES TÉCNICOS") for line in para_texts)
    # No named EQUIPE list members.
    section_equipe = [line for line in para_texts if line.startswith("EQUIPE ") and " – " in line]
    assert section_equipe == []


def test_render_mensal_indice_table_repeats_one_row_per_supplied_entry():
    ctx = MensalContext(
        polo_label="Pimentas",
        periodo_inicio="06/05/2026",
        periodo_fim="12/05/2026",
        mes_extenso="Maio",
        ano="2026",
        indice_tecnologico=(
            IndiceRow(equipe="Maria", servico="Água", quantidade=12, ic_pct_str="98,50"),
            IndiceRow(equipe="João", servico="Esgoto", quantidade=5, ic_pct_str="100,00"),
            IndiceRow(equipe="Ana", servico="Reposição", quantidade=8, ic_pct_str="95,25"),
        ),
    )

    doc = Document(BytesIO(render_mensal(ctx)))
    table = _indice_table(doc)
    # 3 header rows + 3 data rows.
    assert len(table.rows) == 6
    data_rows = [[c.text.strip() for c in row.cells] for row in table.rows[3:]]
    assert data_rows == [
        ["Maria", "Água", "12", "98,50"],
        ["João", "Esgoto", "5", "100,00"],
        ["Ana", "Reposição", "8", "95,25"],
    ]


def test_render_mensal_indice_table_collapses_to_zero_rows_when_empty():
    ctx = MensalContext(
        polo_label="Pimentas",
        periodo_inicio="06/05/2026",
        periodo_fim="12/05/2026",
        mes_extenso="Maio",
        ano="2026",
        indice_tecnologico=(),
    )

    doc = Document(BytesIO(render_mensal(ctx)))
    table = _indice_table(doc)
    # Just the 3 header rows survive.
    assert len(table.rows) == 3


def test_indice_rows_from_inspections_groups_by_team_and_service(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=True)
    wb = load_workbook(path, data_only=True, read_only=True)
    template = PimentasTemplate.detect(wb.sheetnames)

    rows = indice_rows_from_inspections(template.extract_inspections(path))

    assert rows  # non-empty for the fixture
    # Every row has positive quantidade and a pt-BR formatted IC%.
    for r in rows:
        assert r.quantidade > 0
        assert "," in r.ic_pct_str
        # IC% strictly in [0.00, 100.00] formatted with two decimals.
        whole, frac = r.ic_pct_str.split(",")
        assert 0 <= int(whole) <= 100
        assert len(frac) == 2


def test_indice_rows_from_inspections_returns_empty_for_workbook_without_inspections(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=False)
    wb = load_workbook(path, data_only=True, read_only=True)
    template = PimentasTemplate.detect(wb.sheetnames)

    rows = indice_rows_from_inspections(template.extract_inspections(path))
    assert rows == ()


def test_context_from_template_populates_indice_tecnologico_from_xlsx(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=True)
    wb = load_workbook(path, data_only=True, read_only=True)
    template = PimentasTemplate.detect(wb.sheetnames)

    ctx = context_from_template(template, path)

    assert ctx.indice_tecnologico  # populated automatically
    # Default equipe stays empty — xlsx doesn't carry that org chart.
    assert ctx.equipe == ()


def test_context_from_template_passes_through_explicit_equipe(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=True)
    wb = load_workbook(path, data_only=True, read_only=True)
    template = PimentasTemplate.detect(wb.sheetnames)
    overrides = (EquipeMember(id="I", role="Tecnólogo", name="Quem Bom"),)

    ctx = context_from_template(template, path, equipe=overrides)

    assert ctx.equipe == overrides


# Semanal tests — cycle 2 (cover-page-only skeleton, polo + period bindings).


def test_render_semanal_substitutes_polo_and_period_in_cover_table():
    ctx = SemanalContext(
        polo_label="Pimentas",
        periodo_inicio="06/05/2026",
        periodo_fim="12/05/2026",
    )

    body = render_semanal(ctx)

    assert body[:2] == b"PK"  # docx zip
    doc = Document(BytesIO(body))
    # The cover table is doc.tables[0]; row 20 cell 1 is the period line,
    # row 34 cell 1 is the POLO line. Other merged cells hold duplicates.
    text = _all_text(doc)
    assert "06/05/2026" in text
    assert "12/05/2026" in text
    assert "PIMENTAS" in text
    # Source-template values must not leak.
    assert "01/03/2026" not in text
    assert "25/03/2026" not in text
    assert "EXTREMO NORTE" not in text
    # No unrendered Jinja markers.
    assert "{{" not in text


def test_semanal_context_polo_label_upper_derived_from_polo_label():
    ctx = SemanalContext(
        polo_label="Santana",
        periodo_inicio="04/05/2026",
        periodo_fim="10/05/2026",
    )
    assert ctx.polo_label_upper == "SANTANA"
    assert "polo_label_upper" in ctx.as_render_dict()


def test_semanal_context_from_template_uses_inspection_date_bounds(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=True)
    wb = load_workbook(path, data_only=True, read_only=True)
    template = PimentasTemplate.detect(wb.sheetnames)

    ctx = semanal_context_from_template(template, path)

    # Fixture inspections span 2026-03-05 to 2026-03-29.
    assert ctx.periodo_inicio == "05/03/2026"
    assert ctx.periodo_fim == "29/03/2026"
    assert ctx.polo_label == "Pimentas"


def test_semanal_context_from_template_handles_workbook_without_inspections(tmp_path):
    path = make_minimal_pimentas(tmp_path, with_inspections=False)
    wb = load_workbook(path, data_only=True, read_only=True)
    template = PimentasTemplate.detect(wb.sheetnames)

    ctx = semanal_context_from_template(template, path)

    # No dated inspections → blank period strings, polo label still binds.
    assert ctx.polo_label == "Pimentas"
    assert ctx.periodo_inicio == ""
    assert ctx.periodo_fim == ""
