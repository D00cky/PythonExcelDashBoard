import time

from app.core.templates.sabesp_pimentas import (
    SabespPimentasTemplate,
    _inspections_cached,
    _read_service_sheets_cached,
)
from tests.fixtures.sabesp_minimal import make_minimal_sabesp


def _clear_caches():
    _inspections_cached.cache_clear()
    _read_service_sheets_cached.cache_clear()


def test_extract_inspections_returns_same_object_on_repeat_call(tmp_path):
    _clear_caches()
    path = make_minimal_sabesp(tmp_path, with_inspections=True)
    template = SabespPimentasTemplate()

    first = template.extract_inspections(path)
    second = template.extract_inspections(path)

    assert first is second


def test_extract_inspections_invalidates_when_file_mtime_changes(tmp_path):
    _clear_caches()
    path = make_minimal_sabesp(tmp_path, with_inspections=True)
    template = SabespPimentasTemplate()

    first = template.extract_inspections(path)
    # Force mtime to bump even on filesystems with coarse resolution.
    time.sleep(0.01)
    make_minimal_sabesp(tmp_path, with_inspections=True)

    second = template.extract_inspections(path)

    assert first is not second


def test_service_sheets_returns_cached_dict(tmp_path):
    _clear_caches()
    path = make_minimal_sabesp(tmp_path, with_inspections=True)
    template = SabespPimentasTemplate()

    first = template._service_sheets(path)
    second = template._service_sheets(path)

    assert first is second
