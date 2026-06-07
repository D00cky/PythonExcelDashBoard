"""Parquet cache for parsed Polo *batch* data.

A cache lives under ``<instance>/cache/<uuid>/`` and mirrors the batch model:
one normalized parquet per Polo file under ``raw/`` plus pre-computed
multi-level summaries under ``agg/``. The dashboard reads these instead of
re-parsing every xlsx on each request (see :mod:`app.core.ingest` for the
write side).

Layout::

    instance/cache/<uuid>/
    ├── meta.json                  # raw polo→file map, parse timestamp, hashes
    ├── raw/<slug>.parquet         # normalized inspections, one per Polo file
    └── agg/<level>__<slug>.parquet

This module is pure storage: it never opens an xlsx. The ``uuid`` is the
upload id; the cache root is resolved from the active Flask app's
``instance_path``, so all calls must run inside an app context.

Public API:
    cache_dir(uuid) -> Path
    cache_exists(uuid) -> bool
    load_meta(uuid) -> dict
    write_meta(uuid, meta) -> None
    save_raw(uuid, polo, df) -> Path
    load_raw(uuid, polo) -> pd.DataFrame
    save_aggregation(uuid, level, name, df) -> Path
    load_aggregation(uuid, level, name) -> pd.DataFrame
    invalidate(uuid) -> None
    get_cache_size(uuid) -> int
"""

from __future__ import annotations

import json
import re
import shutil
import unicodedata
from pathlib import Path

import pandas as pd
from flask import current_app

META_FILENAME = "meta.json"
_RAW_DIR = "raw"
_AGG_DIR = "agg"


def _cache_root() -> Path:
    return Path(current_app.instance_path) / "cache"


def cache_dir(uuid: str) -> Path:
    """Absolute path to this upload's cache directory (may not exist yet)."""
    return _cache_root() / uuid


def _slug(name: str) -> str:
    """Filesystem-safe, accent-folded slug for a Polo / scope name.

    ``"GOPOÚVA" -> "gopouva"``. Collisions across genuinely distinct names
    (e.g. after accent folding) are avoided by the per-uuid ``raw`` map in
    ``meta.json`` rather than relying on the slug being unique on its own.
    """
    folded = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    folded = re.sub(r"[^a-zA-Z0-9]+", "-", folded).strip("-").lower()
    return folded or "unnamed"


def cache_exists(uuid: str) -> bool:
    """True once :func:`write_meta` has run for this uuid (cache is usable)."""
    return (cache_dir(uuid) / META_FILENAME).is_file()


def load_meta(uuid: str) -> dict:
    """Read ``meta.json``; raises ``FileNotFoundError`` if the cache is absent."""
    return json.loads((cache_dir(uuid) / META_FILENAME).read_text(encoding="utf-8"))


def write_meta(uuid: str, meta: dict) -> None:
    """Persist ``meta.json`` (creating the cache dir). Written last by ingest so
    that ``cache_exists`` only flips true once every parquet is in place."""
    target = cache_dir(uuid)
    target.mkdir(parents=True, exist_ok=True)
    (target / META_FILENAME).write_text(
        json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def _raw_path(uuid: str, polo: str) -> Path:
    return cache_dir(uuid) / _RAW_DIR / f"{_slug(polo)}.parquet"


def save_raw(uuid: str, polo: str, df: pd.DataFrame) -> Path:
    """Persist one Polo file's normalized inspections as parquet."""
    path = _raw_path(uuid, polo)
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_parquet(path, index=False)
    return path


def load_raw(uuid: str, polo: str) -> pd.DataFrame:
    """Load one Polo's cached inspections. Raises ``KeyError`` if not cached."""
    path = _raw_path(uuid, polo)
    if not path.is_file():
        raise KeyError(polo)
    return pd.read_parquet(path)


def _agg_path(uuid: str, level: str, name: str) -> Path:
    stem = level if not name else f"{level}__{_slug(name)}"
    return cache_dir(uuid) / _AGG_DIR / f"{stem}.parquet"


def save_aggregation(uuid: str, level: str, name: str, df: pd.DataFrame) -> Path:
    """Persist a summary frame for ``level`` (``city``/``zone``/``municipality``).

    ``name`` is the zone label or Polo name; pass ``""`` for the singleton
    city-level aggregation.
    """
    path = _agg_path(uuid, level, name)
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_parquet(path, index=False)
    return path


def load_aggregation(uuid: str, level: str, name: str) -> pd.DataFrame:
    """Load a summary frame. Raises ``KeyError`` if that scope wasn't computed."""
    path = _agg_path(uuid, level, name)
    if not path.is_file():
        raise KeyError((level, name))
    return pd.read_parquet(path)


def invalidate(uuid: str) -> None:
    """Delete the whole cache for ``uuid``. No-op if it doesn't exist."""
    shutil.rmtree(cache_dir(uuid), ignore_errors=True)


def get_cache_size(uuid: str) -> int:
    """Total bytes on disk for this uuid's cache (0 if absent). For monitoring."""
    root = cache_dir(uuid)
    if not root.exists():
        return 0
    return sum(p.stat().st_size for p in root.rglob("*") if p.is_file())
