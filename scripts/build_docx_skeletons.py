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
