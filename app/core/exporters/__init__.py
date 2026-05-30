from pathlib import Path
from typing import TYPE_CHECKING

from openpyxl.workbook import Workbook

if TYPE_CHECKING:
    from app.core.templates.pimentas import PimentasTemplate

_MIMETYPES = {
    "md": "text/markdown; charset=utf-8",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


def render_export(
    fmt: str,
    template: "PimentasTemplate",
    workbook: Workbook,
    path: Path,
) -> tuple[bytes, str]:
    if fmt == "md":
        from app.core.exporters.markdown import render_markdown

        body = render_markdown(template, workbook, path).encode("utf-8")
    elif fmt == "xlsx":
        from app.core.exporters.xlsx import render_xlsx

        body = render_xlsx(template, workbook, path)
    elif fmt == "pdf":
        from app.core.exporters.pdf import render_pdf

        body = render_pdf(template, workbook, path)
    elif fmt == "docx":
        from app.core.exporters.docx import render_docx

        body = render_docx(template, workbook, path)
    else:
        raise ValueError(f"unsupported format: {fmt}")
    return body, _MIMETYPES[fmt]
