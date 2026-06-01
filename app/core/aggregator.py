"""Multi-file Polo report aggregation.

A *batch* is a directory holding 1..N weekly xlsx files (one Polo × one week per
file) plus a ``manifest.json`` sidecar that records the parsed period and Polo
for each file. This module owns the manifest format and the in-memory
representation; the route layer reads/writes batches through these primitives.
"""

from __future__ import annotations

import json
import re
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from app.core.templates import recognize
from app.core.templates.pimentas import PimentasTemplate

MANIFEST_FILENAME = "manifest.json"

_PERIODO_RE = re.compile(
    r"(?P<d1>\d{2})/(?P<m1>\d{2})/(?P<y1>\d{4})\s*[àaÀ]\s*"
    r"(?P<d2>\d{2})/(?P<m2>\d{2})/(?P<y2>\d{4})"
)


def parse_periodo(s: str) -> tuple[date, date, str, str]:
    """Parse a ``"Período: dd/mm/yyyy à dd/mm/yyyy"`` cell into structured fields.

    Returns ``(start, end, iso_week, month)`` where ``iso_week`` is ``"YYYY-Www"``
    derived from the start date and ``month`` is ``"YYYY-MM"`` (also from start).
    """
    if not isinstance(s, str):
        raise TypeError(f"expected str, got {type(s).__name__}")
    match = _PERIODO_RE.search(s)
    if match is None:
        raise ValueError(f"no dd/mm/yyyy date range in {s!r}")
    start = date(int(match["y1"]), int(match["m1"]), int(match["d1"]))
    end = date(int(match["y2"]), int(match["m2"]), int(match["d2"]))
    iso_year, iso_week, _ = start.isocalendar()
    return start, end, f"{iso_year:04d}-W{iso_week:02d}", f"{start.year:04d}-{start.month:02d}"


@dataclass(frozen=True)
class PoloFile:
    file_path: Path
    polo: str
    iso_week: str
    month: str
    period_start: date
    period_end: date


@dataclass
class PoloBatch:
    batch_dir: Path
    files: list[PoloFile] = field(default_factory=list)

    @property
    def polos(self) -> list[str]:
        return sorted({f.polo for f in self.files})

    @property
    def iso_weeks(self) -> list[str]:
        return sorted({f.iso_week for f in self.files})

    @property
    def months(self) -> list[str]:
        return sorted({f.month for f in self.files})


def discover_polo_file(xlsx_path: Path) -> PoloFile:
    """Open an xlsx, detect its Polo template, and return a manifest entry.

    Raises ``ValueError`` if the workbook does not match any known template or
    if the period cell is empty/unparseable.
    """
    workbook = load_workbook(xlsx_path, data_only=True, read_only=True)
    try:
        template = recognize(workbook.sheetnames)
        if template is None:
            raise ValueError(f"unknown template in {xlsx_path.name}")
        raw = template.extract_periodo(workbook)
        if not raw:
            raise ValueError(f"empty periodo cell in {xlsx_path.name}")
        start, end, iso_week, month = parse_periodo(raw)
        return PoloFile(
            file_path=xlsx_path,
            polo=template.polo_name,
            iso_week=iso_week,
            month=month,
            period_start=start,
            period_end=end,
        )
    finally:
        workbook.close()


_VALID_VIEWS = ("weekly", "monthly")


def filter_batch(
    batch: PoloBatch, *, polos: tuple[str, ...], view: str, period_key: str
) -> PoloBatch:
    """Return a sub-batch matching ``polos`` and the ``view``/``period_key`` bucket."""
    if view not in _VALID_VIEWS:
        raise ValueError(f"view must be one of {_VALID_VIEWS}, got {view!r}")
    polos_set = set(polos)
    key_attr = "iso_week" if view == "weekly" else "month"
    selected = [
        f for f in batch.files if f.polo in polos_set and getattr(f, key_attr) == period_key
    ]
    return PoloBatch(batch_dir=batch.batch_dir, files=selected)


def combined_inspections(batch: PoloBatch) -> pd.DataFrame:
    """Concat per-file inspections with an extra ``polo`` column."""
    frames = []
    for f in batch.files:
        df = PimentasTemplate().extract_inspections(f.file_path)
        if df.empty:
            continue
        df = df.assign(polo=f.polo)
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=["polo"])
    return pd.concat(frames, ignore_index=True)


def combined_stage_failures(batch: PoloBatch) -> pd.DataFrame:
    """Concat per-file stage-failures with an extra ``polo`` column."""
    frames = []
    for f in batch.files:
        df = PimentasTemplate().extract_stage_failures(f.file_path)
        if df.empty:
            continue
        df = df.assign(polo=f.polo)
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=["polo"])
    return pd.concat(frames, ignore_index=True)


def write_manifest(batch_dir: Path, files: Iterable[PoloFile]) -> Path:
    payload = {
        "files": [
            {
                "file_name": f.file_path.name,
                "polo": f.polo,
                "iso_week": f.iso_week,
                "month": f.month,
                "period_start": f.period_start.isoformat(),
                "period_end": f.period_end.isoformat(),
            }
            for f in files
        ],
    }
    path = batch_dir / MANIFEST_FILENAME
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def load_batch(batch_dir: Path) -> PoloBatch:
    manifest_path = batch_dir / MANIFEST_FILENAME
    if not manifest_path.exists():
        raise FileNotFoundError(manifest_path)
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    files = [
        PoloFile(
            file_path=batch_dir / entry["file_name"],
            polo=entry["polo"],
            iso_week=entry["iso_week"],
            month=entry["month"],
            period_start=date.fromisoformat(entry["period_start"]),
            period_end=date.fromisoformat(entry["period_end"]),
        )
        for entry in payload["files"]
    ]
    return PoloBatch(batch_dir=batch_dir, files=files)
