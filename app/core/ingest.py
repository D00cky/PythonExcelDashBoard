"""Ingestion pipeline: Polo xlsx batch → parquet cache → aggregations.

Orchestrates the read side of :mod:`app.core.cache`. Files are processed **one
at a time** — each Polo file is parsed (calamine, via the template) and written
to the cache as two per-file parquets: its normalized inspections and its stage
failures. Keying by the source file (its stem) means the same Polo across
different weeks never collides, so the period-filtered dashboard can later
concatenate exactly the in-scope files' parquet instead of re-parsing xlsx.

Alongside the per-file raw layer, an **all-periods** overview is rolled up via
:func:`app.core.aggregations.combine_scopes` into municipality (per Polo, all
its weeks), zone, and city summaries. Only the small summary frames are kept
across files, so peak memory stays flat regardless of batch size.

Public API:
    ingest_batch(uuid, batch_dir, progress_callback=None) -> IngestResult
    ingest_single(uuid, xlsx_path, progress_callback=None) -> IngestResult
    ingest_upload(uuid, progress_callback=None) -> IngestResult
    raw_key_for(file_path) -> str
"""

from __future__ import annotations

import gc
import time
from collections import defaultdict
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path

import pandas as pd
from flask import current_app

from app.core import aggregations, cache
from app.core.aggregator import MANIFEST_FILENAME, PoloFile, discover_polo_file, load_batch
from app.core.geography import zone_for
from app.core.templates.pimentas import PimentasTemplate

#: ``(current, total, polo, zone)`` — called once per file processed.
ProgressCallback = Callable[[int, int, str, str], None]

_FAIL_SUFFIX = "__fail"


@dataclass(frozen=True)
class IngestResult:
    uuid: str
    polos: list[str]
    zones: list[str]
    n_files: int
    elapsed_seconds: float


def raw_key_for(file_path: Path) -> str:
    """Cache key for a source file's inspections (its stem, unique per batch)."""
    return Path(file_path).stem


def fail_key_for(file_path: Path) -> str:
    """Cache key for a source file's stage failures."""
    return raw_key_for(file_path) + _FAIL_SUFFIX


def ingest_batch(
    uuid: str, batch_dir: Path, progress_callback: ProgressCallback | None = None
) -> IngestResult:
    """Ingest every Polo file referenced by a batch directory's manifest."""
    return _ingest(uuid, load_batch(Path(batch_dir)).files, progress_callback)


def ingest_single(
    uuid: str, xlsx_path: Path, progress_callback: ProgressCallback | None = None
) -> IngestResult:
    """Ingest a single-Polo upload (the legacy one-file layout)."""
    return _ingest(uuid, [discover_polo_file(Path(xlsx_path))], progress_callback)


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
    files: Sequence[PoloFile],
    progress_callback: ProgressCallback | None,
) -> IngestResult:
    cache.invalidate(uuid)  # always start from a clean, consistent cache
    start = time.perf_counter()
    template = PimentasTemplate()
    total = len(files)

    files_meta: dict[str, dict] = {}
    polo_to_zone: dict[str, str] = {}
    per_polo: dict[str, list[aggregations.ScopeAgg]] = defaultdict(list)
    per_zone: dict[str, list[aggregations.ScopeAgg]] = defaultdict(list)
    all_scopes: list[aggregations.ScopeAgg] = []

    for i, f in enumerate(files, start=1):
        polo = f.polo
        path = Path(f.file_path)
        key = raw_key_for(path)

        inspections = template.extract_inspections(path)
        zone = zone_for(municipality=_dominant_municipality(inspections), polo=polo)
        inspections = _enrich(inspections, polo, zone)
        failures = template.extract_stage_failures(path).assign(polo=polo)

        insp_path = cache.save_raw(uuid, key, inspections)
        fail_path = cache.save_raw(uuid, fail_key_for(path), failures)
        files_meta[key] = {
            "polo": polo,
            "zone": zone,
            "iso_week": f.iso_week,
            "month": f.month,
            "inspections": str(insp_path.relative_to(cache.cache_dir(uuid))),
            "failures": str(fail_path.relative_to(cache.cache_dir(uuid))),
            "rows": int(len(inspections)),
        }

        scope = aggregations.compute_scope(inspections)
        per_polo[polo].append(scope)
        per_zone[zone].append(scope)
        all_scopes.append(scope)
        polo_to_zone[polo] = zone

        if progress_callback is not None:
            progress_callback(i, total, polo, zone)
        del inspections, failures
        gc.collect()

    for polo, scopes in per_polo.items():
        aggregations.save_scope(
            uuid, aggregations.scope_key_muni(polo), aggregations.combine_scopes(scopes)
        )
    for zone, scopes in per_zone.items():
        aggregations.save_scope(
            uuid, aggregations.scope_key_zone(zone), aggregations.combine_scopes(scopes)
        )
    aggregations.save_scope(
        uuid, aggregations.scope_key_city(), aggregations.combine_scopes(all_scopes)
    )

    elapsed = time.perf_counter() - start
    meta = {
        "files": files_meta,
        "polo_to_zone": polo_to_zone,
        "polos": sorted(polo_to_zone),
        "zones": sorted(per_zone),
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
