"""Convert a real Sabesp .docx report into a docxtpl-compatible Jinja skeleton.

Run after the source `.docx` in ``Model/RELATORIOS/`` changes — this script is the
single source of truth for which anchors get bound to which variables. Skeletons it
emits live under ``app/core/templates/docx_skeletons/`` and are committed to the repo
so production rendering does not depend on ``Model/`` being present.

Word splits running text across XML runs that do not align with our anchors — e.g.
``01/04/2026`` lands in five separate runs. The replacement helpers below find a
literal anchor across run boundaries and write the placeholder into the first
overlapping run, blanking the rest. This preserves the formatting of the first run.
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.document import Document as DocxDocument
from docx.table import _Cell
from docx.text.paragraph import Paragraph

REPO_ROOT = Path(__file__).resolve().parent.parent
MENSAL_SOURCE_DIR = REPO_ROOT / "Model" / "RELATORIOS"
SKELETON_DIR = REPO_ROOT / "app" / "core" / "templates" / "docx_skeletons"
OUTPUT_MENSAL = SKELETON_DIR / "mensal_sabesp.docx"


def find_mensal_source() -> Path:
    """Glob for the Sabesp monthly .docx so a mojibake filename doesn't break us.

    The committed file in ``Model/RELATORIOS/`` was unzipped with a CP437→UTF-8
    artifact, so a literal string match fails. The shape ``*Sabesp*Polo*.docx`` is
    stable across re-extracts.
    """
    candidates = sorted(MENSAL_SOURCE_DIR.glob("*Sabesp*Polo*.docx"))
    if not candidates:
        raise FileNotFoundError(
            f"no '*Sabesp*Polo*.docx' under {MENSAL_SOURCE_DIR}; "
            "did the Model/RELATORIOS layout change?"
        )
    return candidates[0]


# (anchor literal in source, placeholder to emit) — order matters: longer/more-specific
# anchors first, so partial overlaps with shorter ones don't corrupt the document.
MENSAL_REPLACEMENTS: list[tuple[str, str]] = [
    ("01/04/2026", "{{ periodo_inicio }}"),
    ("30/04/2026", "{{ periodo_fim }}"),
    ("Abril de 2026", "{{ mes_extenso }} de {{ ano }}"),
    ("GOPOÚVA", "{{ polo_label_upper }}"),
    ("Gopoúva", "{{ polo_label }}"),
    ("Gopouva", "{{ polo_label }}"),
]


def replace_in_paragraph(paragraph: Paragraph, anchor: str, placeholder: str) -> int:
    """Replace every occurrence of ``anchor`` with ``placeholder`` in ``paragraph``,
    crossing run boundaries. Returns the number of replacements made.

    Strategy: walk the runs, build the concatenated text, locate the anchor, then
    rewrite the runs that overlap the match. The first overlapping run absorbs the
    placeholder (preserving its formatting); subsequent overlapping runs are cleared.
    The tail of the last overlapping run (after the match) is preserved.
    """
    runs = paragraph.runs
    if not runs:
        return 0

    replacements = 0
    while True:
        offsets: list[tuple[int, int]] = []
        cursor = 0
        for run in runs:
            length = len(run.text)
            offsets.append((cursor, cursor + length))
            cursor += length
        concat = "".join(run.text for run in runs)

        match_start = concat.find(anchor)
        if match_start == -1:
            return replacements
        match_end = match_start + len(anchor)

        # Identify which runs the match spans, and the within-run boundaries.
        first_run_idx = next(
            i for i, (start, end) in enumerate(offsets) if start <= match_start < end
        )
        last_run_idx = next(i for i, (start, end) in enumerate(offsets) if start < match_end <= end)

        first_start, _first_end = offsets[first_run_idx]
        _last_start, last_end = offsets[last_run_idx]
        prefix = runs[first_run_idx].text[: match_start - first_start]
        suffix = runs[last_run_idx].text[len(runs[last_run_idx].text) - (last_end - match_end) :]

        runs[first_run_idx].text = prefix + placeholder
        for idx in range(first_run_idx + 1, last_run_idx + 1):
            runs[idx].text = ""
        # Append the trailing suffix back onto the last run (or first if same).
        if last_run_idx == first_run_idx:
            runs[first_run_idx].text += suffix
        else:
            runs[last_run_idx].text = suffix

        replacements += 1


def iter_paragraphs(doc: DocxDocument):
    """Yield every paragraph in the document, including those inside tables."""
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_cell_paragraphs(cell)


def _iter_cell_paragraphs(cell: _Cell):
    yield from cell.paragraphs
    for table in cell.tables:
        for row in table.rows:
            for inner_cell in row.cells:
                yield from _iter_cell_paragraphs(inner_cell)


def _find_paragraph_starting_with(paragraphs, prefix: str):
    return next((p for p in paragraphs if p.text.startswith(prefix)), None)


def wrap_equipe_block(doc: DocxDocument) -> int:
    """Turn the EQUIPE I/III/IV listing into a docxtpl paragraph loop.

    Source layout (after cycle-1 substitutions) is three List Paragraphs:
      EQUIPE I  – Tecnólogo Lucas Jeremias
      EQUIPE III – Engenheira Geovana ...
      EQUIPE IV – Engenheiro Evanuel ...
      ASSISTENTES TÉCNICOS – ...

    We keep the first as the loop body (preserving the bold "EQUIPE X" run),
    delete the other two, and wrap with ``{%p for m in equipe %}`` /
    ``{%p endfor %}`` directive paragraphs. ASSISTENTES stays as static text.
    Returns 1 on success, 0 if the anchor wasn't found (template already wrapped).
    """
    anchor = _find_paragraph_starting_with(doc.paragraphs, "EQUIPE I")
    if anchor is None:
        return 0

    anchor.insert_paragraph_before("{%p for m in equipe %}", style=anchor.style)

    # Body: keep run[0] bold "EQUIPE {{ m.id }}", run[1] normal
    # " – {{ m.role }} {{ m.name }}", blank the rest.
    runs = anchor.runs
    if runs:
        runs[0].text = "EQUIPE {{ m.id }}"
        if len(runs) > 1:
            runs[1].text = " – {{ m.role }} {{ m.name }}"
        for r in runs[2:]:
            r.text = ""

    # Delete EQUIPE III / EQUIPE IV — text-anchored so a reorder in the source
    # won't silently produce a malformed skeleton.
    for p in list(doc.paragraphs):
        if p.text.startswith(("EQUIPE III", "EQUIPE IV")):
            p._element.getparent().remove(p._element)

    assistentes = _find_paragraph_starting_with(doc.paragraphs, "ASSISTENTES TÉCNICOS")
    if assistentes is not None:
        assistentes.insert_paragraph_before("{%p endfor %}", style=anchor.style)
    return 1


def wrap_indice_tecnologico_table(doc: DocxDocument) -> int:
    """Turn the ÍNDICE TECNOLÓGICO POR EQUIPE table body into a docxtpl row loop.

    Per docxtpl's row-loop convention, the ``{%tr for %}`` / ``{%tr endfor %}``
    directives must sit in *separate* rows surrounding the body row — putting both
    on a single row makes docxtpl's `<w:tr>` boundary rewrite fail and Jinja then
    sees a dangling ``endfor``.

    Layout produced:
      row 0-2: existing header rows (unchanged)
      row 3:   {%tr for r in indice %}   (directive row — removed at render time)
      row 4:   body row with {{ r.* }} Jinja expressions (repeated for each item)
      row 5:   {%tr endfor %}            (directive row — removed at render time)
    Surplus original data rows are deleted. Returns 1 on success, 0 if the table
    wasn't found.
    """
    import copy

    HEADER_ROWS = 3
    target = next(
        (
            t
            for t in doc.tables
            if t.rows
            and t.rows[0].cells[0].text.strip().startswith("ÍNDICE TECNOLÓGICO POR EQUIPE")
        ),
        None,
    )
    if target is None or len(target.rows) <= HEADER_ROWS:
        return 0

    body_row = target.rows[HEADER_ROWS]

    # Drop surplus original data rows first so insertion indices stay simple.
    for row in list(target.rows[HEADER_ROWS + 1 :]):
        target._tbl.remove(row._tr)

    # Body row: Jinja expressions only, no directives.
    body = body_row.cells
    body[0].text = "{{ r.equipe }}"
    body[1].text = "{{ r.servico }}"
    body[2].text = "{{ r.quantidade }}"
    body[3].text = "{{ r.ic_pct_str }}"

    # Insert {%tr for %} directive row immediately before the body row.
    body_row._tr.addprevious(copy.deepcopy(body_row._tr))
    for_row = target.rows[HEADER_ROWS]
    for c in for_row.cells:
        c.text = ""
    for_row.cells[0].text = "{%tr for r in indice %}"

    # Insert {%tr endfor %} directive row immediately after the body row.
    body_row = target.rows[HEADER_ROWS + 1]  # body row shifted by one
    body_row._tr.addnext(copy.deepcopy(body_row._tr))
    endfor_row = target.rows[HEADER_ROWS + 2]
    for c in endfor_row.cells:
        c.text = ""
    endfor_row.cells[0].text = "{%tr endfor %}"
    return 1


def build_mensal_skeleton(source: Path, target: Path) -> dict[str, int]:
    if not source.exists():
        raise FileNotFoundError(f"source docx not found: {source}")
    doc = Document(str(source))
    counts: dict[str, int] = {}
    for anchor, placeholder in MENSAL_REPLACEMENTS:
        n = 0
        for paragraph in iter_paragraphs(doc):
            n += replace_in_paragraph(paragraph, anchor, placeholder)
        counts[anchor] = n
    counts["<equipe loop>"] = wrap_equipe_block(doc)
    counts["<indice table loop>"] = wrap_indice_tecnologico_table(doc)
    target.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(target))
    return counts


def main() -> int:
    source = find_mensal_source()
    print(f"source:  {source}")
    print(f"target:  {OUTPUT_MENSAL}")
    counts = build_mensal_skeleton(source, OUTPUT_MENSAL)
    width = max(len(a) for a in counts)
    for anchor, n in counts.items():
        marker = "  " if n else "!!"
        print(f"  {marker} {anchor:<{width}} → {n} replacement(s)")
    print(f"wrote {OUTPUT_MENSAL.stat().st_size:,} bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
