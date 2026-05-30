"""Cache + lazy-render behavior for the dashboard context builder.

The dashboard re-renders ~17 Plotly figures per request. Caching the
context dict by (path, mtime, filter, swap) cuts repeat hits down to a
dict lookup. These tests pin the contract so the cache cannot silently
degrade into a no-op or — worse — start serving stale data after the
underlying file changes.
"""

import os

from app.routes.main import _cached_sabesp_context, _defer_plotly_script
from tests.fixtures.sabesp_minimal import make_minimal_sabesp


def test_cache_reuses_context_for_same_path_and_filter(tmp_path):
    path = make_minimal_sabesp(tmp_path, with_inspections=True)
    _cached_sabesp_context.cache_clear()

    _cached_sabesp_context(str(path), path.stat().st_mtime_ns, "", "", False)
    _cached_sabesp_context(str(path), path.stat().st_mtime_ns, "", "", False)

    info = _cached_sabesp_context.cache_info()
    assert info.misses == 1
    assert info.hits == 1


def test_cache_treats_different_filters_as_different_entries(tmp_path):
    path = make_minimal_sabesp(tmp_path, with_inspections=True)
    _cached_sabesp_context.cache_clear()

    _cached_sabesp_context(str(path), path.stat().st_mtime_ns, "", "", False)
    _cached_sabesp_context(str(path), path.stat().st_mtime_ns, "2026-03-10", "", False)
    _cached_sabesp_context(str(path), path.stat().st_mtime_ns, "", "2026-03-20", False)
    _cached_sabesp_context(str(path), path.stat().st_mtime_ns, "", "", True)

    info = _cached_sabesp_context.cache_info()
    assert info.misses == 4
    assert info.hits == 0


def test_cache_invalidates_when_file_mtime_changes(tmp_path):
    """A re-uploaded file at the same path must NOT serve the previous cache entry."""
    path = make_minimal_sabesp(tmp_path, with_inspections=True)
    _cached_sabesp_context.cache_clear()

    first_mtime = path.stat().st_mtime_ns
    _cached_sabesp_context(str(path), first_mtime, "", "", False)

    # Simulate a new upload landing at the same path with a different mtime.
    os.utime(path, ns=(first_mtime + 1_000_000_000, first_mtime + 1_000_000_000))
    new_mtime = path.stat().st_mtime_ns
    assert new_mtime != first_mtime
    _cached_sabesp_context(str(path), new_mtime, "", "", False)

    info = _cached_sabesp_context.cache_info()
    assert info.misses == 2  # both calls were cache misses


def test_defer_plotly_script_swaps_only_the_first_script_tag():
    """Plotly emits exactly one inline <script> per chart; defer must mark that one only."""
    html = '<div id="x" class="plotly-graph-div"></div><script>Plotly.newPlot("x");</script>'

    deferred = _defer_plotly_script(html)

    assert deferred.count('<script type="text/plotly-defer">') == 1
    assert deferred.count("<script>") == 0
    # The body of the script is untouched — only the opening tag changes.
    assert "Plotly.newPlot" in deferred


def test_defer_plotly_script_is_idempotent_on_already_deferred_html():
    """If a chart HTML is somehow deferred twice, we don't end up with a broken tag."""
    html = '<div></div><script type="text/plotly-defer">Plotly.newPlot("x");</script>'

    assert _defer_plotly_script(html) == html


def test_dashboard_defers_only_per_service_chart_scripts(client, tmp_path):
    """Top-level charts must render eagerly; per-service charts deferred for lazy boot."""
    import io

    payload = make_minimal_sabesp(tmp_path, with_inspections=True).read_bytes()
    upload = client.post(
        "/upload",
        data={"file": (io.BytesIO(payload), "report.xlsx")},
        content_type="multipart/form-data",
    )

    body = client.get(upload.location).data.decode("utf-8")

    # 4 services × 2 charts each = 8 deferred scripts; everything else stays eager.
    assert body.count('<script type="text/plotly-defer">') == 8
    # Bootstrap is in place — without it the deferred charts would never render.
    assert "IntersectionObserver" in body
    assert "service-section" in body


def test_dashboard_route_caches_across_requests(client, tmp_path):
    """End-to-end: two identical GETs hit the figure-rendering code only once."""
    import io

    from tests.fixtures.sabesp_minimal import make_minimal_sabesp

    payload = make_minimal_sabesp(tmp_path, with_inspections=True).read_bytes()
    upload = client.post(
        "/upload",
        data={"file": (io.BytesIO(payload), "report.xlsx")},
        content_type="multipart/form-data",
    )
    dashboard_url = upload.location

    _cached_sabesp_context.cache_clear()
    client.get(dashboard_url)
    client.get(dashboard_url)

    info = _cached_sabesp_context.cache_info()
    assert info.misses == 1
    assert info.hits == 1
