"""Ingestion pipeline: Polo xlsx batch → parquet cache → aggregations.

Orchestrates the read side of :mod:`app.core.cache`. Files are processed **one
at a time** — each Polo file is parsed (calamine, via the template), normalized,
written to ``raw/<polo>.parquet``, and reduced to its municipality summary;
only the tiny summary frames are kept around to roll up into zone and city
scopes (see :func:`app.core.aggregations.combine_scopes`). Raw inspection rows
are never accumulated across files, so peak memory stays flat regardless of
batch size.

Public API:
    ingest_batch(uuid, batch_dir, progress_callback=None) -> IngestResult
    ingest_single(uuid, xlsx_path, progress_callback=None) -> IngestResult
    ingest_upload(uuid, progress_callback=None) -> IngestResult
"""

from __future__ import annotations

import gc
import time
from collections import defaultdict
from collections.abc import Callable
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path

import pandas as pd
from flask import current_app

from app.core import aggregations, cache
from app.core.aggregator import MANIFEST_FILENAME, discover_polo_file, load_batch
from app.core.geography import zone_for
from app.core.templates.pimentas import PimentasTemplate

#: ``(current, total, polo, zone)`` — called once per file processed.
ProgressCallback = Callable[[int, int, str, str], None]


@dataclass(frozen=True)
class IngestResult:
    uuid: str
    polos: list[str]
    zones: list[str]
    n_files: int
    elapsed_seconds: float


def ingest_batch(
    uuid: str, batch_dir: Path, progress_callback: ProgressCallback | None = None
) -> IngestResult:
    """Ingest every Polo file referenced by a batch directory's manifest."""
    batch = load_batch(Path(batch_dir))
    items = [(f.polo, Path(f.file_path)) for f in batch.files]
    return _ingest(uuid, items, progress_callback)


def ingest_single(
    uuid: str, xlsx_path: Path, progress_callback: ProgressCallback | None = None
) -> IngestResult:
    """Ingest a single-Polo upload (the legacy one-file layout)."""
    pf = discover_polo_file(Path(xlsx_path))
    return _ingest(uuid, [(pf.polo, Path(pf.file_path))], progress_callback)


def ingest_upload(uuid: str, progress_callback: ProgressCallback | None = None) -> IngestResult:
    """Resolve an upload id under ``instance/uploads`` and ingest it.

    Mirrors the route layer's single-file vs. batch convention:
    ``<uuid>.xlsx`` is a single Polo, ``<uuid>/manifest.json`` is a batch.
    """
    uploads = Path(current_app.instance_path) / "uploads"
    single = uploads / f"{uuid}.xlsx"
    if single.is_file():
        return ingest_single(uuid, single, progress_callback)
    batch_dir = uploads / uuid
    if (batch_dir / MANIFEST_FILENAME).is_file():
        return ingest_batch(uuid, batch_dir, progress_callback)
    raise FileNotFoundError(uuid)


def _ingest(
    uuid: str,
    items: list[tuple[str, Path]],
    progress_callback: ProgressCallback | None,
) -> IngestResult:
    cache.invalidate(uuid)  # always start from a clean, consistent cache
    start = time.perf_counter()
    template = PimentasTemplate()
    total = len(items)

    raw_index: dict[str, str] = {}
    row_counts: dict[str, int] = {}
    polo_to_zone: dict[str, str] = {}
    per_zone: dict[str, list[aggregations.ScopeAgg]] = defaultdict(list)
    all_scopes: list[aggregations.ScopeAgg] = []

    for i, (polo, path) in enumerate(items, start=1):
        df = template.extract_inspections(path)
        zone = zone_for(municipality=_dominant_municipality(df), polo=polo)
        enriched = _enrich(df, polo, zone)
        raw_path = cache.save_raw(uuid, polo, enriched)
        raw_index[polo] = str(raw_path.relative_to(cache.cache_dir(uuid)))
        row_counts[polo] = int(len(enriched))

        scope = aggregations.compute_scope(enriched)
        aggregations.save_scope(uuid, aggregations.scope_key_muni(polo), scope)
        per_zone[zone].append(scope)
        all_scopes.append(scope)
        polo_to_zone[polo] = zone

        if progress_callback is not None:
            progress_callback(i, total, polo, zone)
        del df, enriched
        gc.collect()

    for zone, scopes in per_zone.items():
        aggregations.save_scope(
            uuid, aggregations.scope_key_zone(zone), aggregations.combine_scopes(scopes)
        )
    aggregations.save_scope(
        uuid, aggregations.scope_key_city(), aggregations.combine_scopes(all_scopes)
    )

    elapsed = time.perf_counter() - start
    meta = {
        "raw": raw_index,
        "polo_to_zone": polo_to_zone,
        "zones": sorted(per_zone),
        "polos": sorted(polo_to_zone),
        "row_counts": row_counts,
        "n_files": total,
        "elapsed_seconds": elapsed,
        "parsed_at": datetime.now(UTC).isoformat(),
    }
    cache.write_meta(uuid, meta)  # last: flips cache_exists() true atomically
    return IngestResult(
        uuid=uuid,
        polos=meta["polos"],
        zones=meta["zones"],
        n_files=total,
        elapsed_seconds=elapsed,
    )


def _dominant_municipality(df: pd.DataFrame) -> str:
    """Most-frequent non-empty Município in a file (matches aggregator.py)."""
    if df.empty or "municipality" not in df.columns:
        return ""
    munis = df["municipality"].dropna().astype(str).str.strip()
    munis = munis[munis != ""]
    return munis.value_counts().idxmax() if not munis.empty else ""


def _enrich(df: pd.DataFrame, polo: str, zone: str) -> pd.DataFrame:
    """Tag a Polo file's inspections with its ``polo`` and (dominant) ``zone``."""
    out = df.assign(polo=polo)
    if "municipality" not in out.columns:
        out["municipality"] = ""
    out["zone"] = zone
    return out
