"""Cache-aware data source for the batch dashboard.

Reads each in-scope file's inspections / stage-failures from the parquet cache
when it's warm, concatenating them — avoiding the per-request xlsx re-parse that
:func:`app.core.aggregator.combined_inspections` and
:func:`app.core.aggregator.combined_stage_failures` perform. Falls back to live
parsing (the pre-cache behavior) when the cache is cold or any in-scope file is
missing from it, so the dashboard degrades gracefully and always returns a
consistent result.

Public API:
    combined_inspections(uuid, batch) -> pd.DataFrame
    combined_stage_failures(uuid, batch) -> pd.DataFrame
"""

from __future__ import annotations

from collections.abc import Callable

import pandas as pd

from app.core import aggregator, cache
from app.core.aggregator import PoloBatch
from app.core.ingest import fail_key_for, raw_key_for

_INSP_EMPTY_COLS = ["polo", "municipality", "zone"]
_FAIL_EMPTY_COLS = ["polo"]


def combined_inspections(uuid: str, batch: PoloBatch) -> pd.DataFrame:
    """In-scope inspections, from cache when warm else live parse."""
    return _combined(
        uuid,
        batch,
        key_for=raw_key_for,
        empty_cols=_INSP_EMPTY_COLS,
        live=aggregator.combined_inspections,
    )


def combined_stage_failures(uuid: str, batch: PoloBatch) -> pd.DataFrame:
    """In-scope stage failures, from cache when warm else live parse."""
    return _combined(
        uuid,
        batch,
        key_for=fail_key_for,
        empty_cols=_FAIL_EMPTY_COLS,
        live=aggregator.combined_stage_failures,
    )


def _combined(
    uuid: str,
    batch: PoloBatch,
    *,
    key_for: Callable[[object], str],
    empty_cols: list[str],
    live: Callable[[PoloBatch], pd.DataFrame],
) -> pd.DataFrame:
    if not cache.cache_exists(uuid):
        return live(batch)
    frames = []
    for f in batch.files:
        try:
            frames.append(cache.load_raw(uuid, key_for(f.file_path)))
        except KeyError:
            # Partially-warm cache (e.g. a file added after ingest): fall back
            # to live so inspections and failures stay drawn from one source.
            return live(batch)
    if not frames:
        return pd.DataFrame(columns=empty_cols)
    return pd.concat(frames, ignore_index=True)
