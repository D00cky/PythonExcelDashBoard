import re
import unicodedata
import uuid
from functools import lru_cache
from pathlib import Path
from typing import Any
from urllib.parse import quote

import pandas as pd
import plotly.graph_objects as go
from flask import (
    Blueprint,
    Response,
    abort,
    current_app,
    jsonify,
    redirect,
    render_template,
    request,
    url_for,
)
from openpyxl import load_workbook

from app.core import dashboard_data
from app.core.aggregator import (
    MANIFEST_FILENAME,
    PoloBatch,
    date_bounds,
    discover_polo_file,
    filter_batch,
    format_period_pt,
    load_batch,
    write_manifest,
)
from app.core.aggregator import (
    ic_rows_from_inspections as _ic_rows_from_inspections,
)
from app.core.aggregator import (
    iqs_overall_from_inspections as _iqs_overall_from_inspections,
)
from app.core.aggregator import (
    iqs_rows_from_inspections as _iqs_rows_from_inspections,
)
from app.core.exporters import render_export
from app.core.jobs import (
    JOB_FAILED,
    JOB_PROCESSING,
    JOB_QUEUED,
    get_status,
    start_ingest_job,
)
from app.core.templates import recognize
from app.core.templates.pimentas import (
    PimentasTemplate,
    top_observations,
)

bp = Blueprint("main", __name__)


@bp.get("/")
def index() -> str:
    return render_template("index.html")


@bp.post("/upload")
def upload():
    files = request.files.getlist("file")
    files = [f for f in files if f and f.filename]
    if not files:
        abort(400)
    if not all(f.filename.lower().endswith(".xlsx") for f in files):
        abort(400)

    upload_id = uuid.uuid4().hex
    uploads_dir = Path(current_app.instance_path) / "uploads"
    uploads_dir.mkdir(parents=True, exist_ok=True)

    if len(files) == 1:
        # Legacy single-file layout: keep top-level path so existing /dashboard/<id>
        # links and tests continue to work without migration.
        files[0].save(uploads_dir / f"{upload_id}.xlsx")
        return redirect(url_for("main.dashboard", upload_id=upload_id), code=303)

    batch_dir = uploads_dir / upload_id
    batch_dir.mkdir(parents=True, exist_ok=True)
    polo_files = []
    for idx, f in enumerate(files):
        target = batch_dir / f"file_{idx:02d}.xlsx"
        f.save(target)
        try:
            polo_files.append(discover_polo_file(target))
        except ValueError as exc:
            current_app.logger.warning("skipping unrecognized upload %s: %s", f.filename, exc)
    if not polo_files:
        abort(400)
    write_manifest(batch_dir, polo_files)
    # Parse the batch into the parquet cache off the request path; the dashboard
    # shows a processing page (polling /api/status) until it's ready.
    start_ingest_job(upload_id)
    return redirect(url_for("main.dashboard", upload_id=upload_id), code=303)


@bp.get("/api/status/<upload_id>")
def api_status(upload_id: str) -> Response:
    """JSON ingestion status for the processing page's poller; 404 if unknown."""
    status = get_status(upload_id)
    if status is None:
        abort(404)
    return jsonify(status)


@bp.post("/retry/<upload_id>")
def retry(upload_id: str):
    """Re-run a failed (or any) ingestion and return to the dashboard."""
    _resolve_upload(upload_id)  # 404 if the upload is gone
    start_ingest_job(upload_id)
    return redirect(url_for("main.dashboard", upload_id=upload_id), code=303)


@bp.get("/dashboard/<upload_id>")
def dashboard(upload_id: str):
    kind, target = _resolve_upload(upload_id)

    # While the background ingestion runs (or after it failed), show the
    # processing page instead of the dashboard. A done job / warm cache falls
    # through to normal rendering.
    gated = _job_gate(upload_id)
    if gated is not None:
        return gated

    if kind == "batch":
        polo_arg = request.args.get("polo")
        zone_arg = request.args.get("zone")
        municipality_arg = request.args.get("municipality")
        if polo_arg and polo_arg != "__all":
            path = _batch_polo_path(target, polo_arg)
            return _render_single_polo_dashboard(upload_id, path, batch_dir=target, polo=polo_arg)
        return _render_batch_dashboard(
            upload_id,
            target,
            zone=zone_arg,
            municipality=municipality_arg,
        )

    path = target
    workbook = load_workbook(path, data_only=True, read_only=True)
    template = recognize(workbook.sheetnames)
    if not isinstance(template, PimentasTemplate):
        return render_template("dashboard_unknown.html", sheet_names=workbook.sheetnames)

    filter_start = _parse_iso_date(request.args.get("start", ""))
    filter_end = _parse_iso_date(request.args.get("end", ""))

    return render_template(
        "dashboard.html",
        download_action=url_for("main.download", upload_id=upload_id),
        **_cached_polo_context(
            str(path),
            path.stat().st_mtime_ns,
            _iso(filter_start),
            _iso(filter_end),
            _swap_arg(),
        ),
    )


@bp.get("/dashboard/<upload_id>/zona/<zone_slug>")
def dashboard_zone(upload_id: str, zone_slug: str):
    """Zone-scoped dashboard URL from the Phase 3 production navigation."""
    kind, target = _resolve_upload(upload_id)
    if kind != "batch":
        abort(404)
    gated = _job_gate(upload_id)
    if gated is not None:
        return gated

    batch = load_batch(target)
    zone = _zone_from_slug(batch, zone_slug)
    if zone is None:
        abort(404)
    return _render_batch_dashboard(upload_id, target, zone=zone)


@bp.get("/dashboard/<upload_id>/zona/<zone_slug>/municipio/<municipality_slug>")
def dashboard_municipality(upload_id: str, zone_slug: str, municipality_slug: str):
    """Municipality/Polo-scoped dashboard URL from the Phase 3 navigation."""
    kind, target = _resolve_upload(upload_id)
    if kind != "batch":
        abort(404)
    gated = _job_gate(upload_id)
    if gated is not None:
        return gated

    batch = load_batch(target)
    zone = _zone_from_slug(batch, zone_slug)
    if zone is None:
        abort(404)
    polo = _polo_from_slug(batch, municipality_slug, zone=zone)
    if polo is None:
        abort(404)
    path = _batch_polo_path(target, polo)
    return _render_single_polo_dashboard(upload_id, path, batch_dir=target, polo=polo)


def _job_gate(upload_id: str):
    """Return a processing/error response while ingestion is not usable yet."""
    job = get_status(upload_id)
    if job is not None and job["status"] in (JOB_QUEUED, JOB_PROCESSING):
        return render_template("processing.html", upload_id=upload_id, status=job)
    if job is not None and job["status"] == JOB_FAILED:
        return render_template("processing.html", upload_id=upload_id, status=job), 500
    return None


def _swap_arg() -> str:
    """Normalize ?swap= into a cache-friendly key. '' = auto-detect."""
    raw = request.args.get("swap")
    if raw == "1":
        return "1"
    if raw == "0":
        return "0"
    return ""


def _slug(value: str) -> str:
    folded = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "-", folded.lower()).strip("-")


def _zone_slug(zone: str) -> str:
    slug = _slug(zone)
    return slug.removeprefix("zona-") if slug.startswith("zona-") else slug


def _zone_from_slug(batch: PoloBatch, slug: str) -> str | None:
    from app.core.aggregator import polo_geography

    wanted = _zone_slug(slug)
    zones = {polo_geography(f)[1] for f in batch.files}
    for zone in zones:
        if _zone_slug(zone) == wanted or _slug(zone) == _slug(slug):
            return zone
    return None


def _polo_from_slug(batch: PoloBatch, slug: str, *, zone: str | None = None) -> str | None:
    from app.core.aggregator import polo_geography

    wanted = _slug(slug)
    for f in batch.files:
        if zone is not None and polo_geography(f)[1] != zone:
            continue
        if _slug(f.polo) == wanted:
            return f.polo
    return None


def _build_scope_sidebar(
    upload_id: str,
    batch: PoloBatch,
    *,
    active_zone: str | None,
    active_polo: str | None,
) -> dict[str, Any]:
    """Navigation model for the Phase 3 city → zone → municipality sidebar."""
    from app.core.aggregator import polo_geography

    zones: dict[str, list[dict[str, str]]] = {}
    for f in batch.files:
        municipality, zone = polo_geography(f)
        zones.setdefault(zone, []).append(
            {
                "label": f.polo.title(),
                "municipality": municipality or f.polo.title(),
                "href": _scope_url(upload_id, zone=zone, polo=f.polo),
                "active": active_polo == f.polo,
                "search_text": f"{f.polo} {municipality} {zone}".lower(),
            }
        )

    zone_items = []
    for zone, municipalities in sorted(zones.items()):
        sorted_municipalities = sorted(municipalities, key=lambda m: m["label"])
        has_active_municipality = any(m["active"] for m in sorted_municipalities)
        zone_items.append(
            {
                "label": zone,
                "href": _scope_url(upload_id, zone=zone),
                "active": active_zone == zone and active_polo in (None, "__all"),
                "expanded": active_zone == zone or has_active_municipality,
                "municipalities": sorted_municipalities,
                "count": len(sorted_municipalities),
            }
        )

    return {
        "city_href": _scope_url(upload_id),
        "city_active": active_zone is None and active_polo in (None, "__all"),
        "zones": zone_items,
        "stats": {
            "municipalities": len({f.polo for f in batch.files}),
            "zones": len(zone_items),
            "files": len(batch.files),
        },
    }


def _scope_url(upload_id: str, *, zone: str | None = None, polo: str | None = None) -> str:
    base = url_for("main.dashboard", upload_id=upload_id)
    if zone is None:
        return base
    path = f"{base}/zona/{quote(_zone_slug(zone))}"
    if polo is not None and polo != "__all":
        path += f"/municipio/{quote(_slug(polo))}"
    return path


def _breadcrumb(
    upload_id: str,
    *,
    zone: str | None = None,
    municipality: str | None = None,
    polo: str | None = None,
) -> list[dict[str, str]]:
    items = [{"label": "São Paulo", "href": _scope_url(upload_id)}]
    if zone:
        items.append({"label": zone, "href": _scope_url(upload_id, zone=zone)})
    label = polo if polo and polo != "__all" else municipality
    if label:
        items.append({"label": label.title(), "href": _scope_url(upload_id, zone=zone, polo=polo)})
    return items


def _swap_arg_to_bool(value: str) -> bool | None:
    return {"1": True, "0": False, "": None}.get(value)


def _batch_polo_path(batch_dir: Path, polo: str) -> Path:
    """Return the xlsx path for ``polo`` inside the batch, or abort 404."""
    batch = load_batch(batch_dir)
    for pf in batch.files:
        if pf.polo == polo:
            return pf.file_path
    abort(404)


def _polo_location(batch: PoloBatch, polo: str) -> tuple[str | None, str | None]:
    """Return ``(zone, municipality)`` for a Polo inside a batch."""
    from app.core.aggregator import polo_geography

    for f in batch.files:
        if f.polo == polo:
            municipality, zone = polo_geography(f)
            return zone, municipality or None
    return None, None


def _render_single_polo_dashboard(upload_id: str, path: Path, *, batch_dir: Path, polo: str) -> str:
    """Render the legacy single-file dashboard, but tagged as one tab in a batch."""
    _open_pimentas(path)  # 404 guard; cached context below re-opens & uses workbook

    batch = load_batch(batch_dir)
    derived_zone, derived_municipality = _polo_location(batch, polo)
    active_zone = request.args.get("zone") or derived_zone
    active_municipality = request.args.get("municipality") or derived_municipality
    filter_start = _parse_iso_date(request.args.get("start", ""))
    filter_end = _parse_iso_date(request.args.get("end", ""))

    context = _cached_polo_context(
        str(path),
        path.stat().st_mtime_ns,
        _iso(filter_start),
        _iso(filter_end),
        _swap_arg(),
    )
    hierarchical = _build_hierarchical_tabs(
        upload_id,
        batch,
        active_zone=active_zone,
        active_municipality=active_municipality,
        active_polo=polo,
    )
    return render_template(
        "dashboard.html",
        download_action=url_for("main.download", upload_id=upload_id),
        tabs=_build_tabs(upload_id, batch, active_polo=polo),
        zone_tabs=hierarchical["zone_tabs"],
        municipality_tabs=hierarchical["municipality_tabs"],
        polo_tabs=hierarchical["polo_tabs"],
        active_zone=active_zone,
        active_municipality=active_municipality,
        active_polo=polo,
        breadcrumbs=_breadcrumb(
            upload_id,
            zone=active_zone,
            municipality=active_municipality,
            polo=polo,
        ),
        scope_level="municipality",
        scope_sidebar=_build_scope_sidebar(
            upload_id,
            batch,
            active_zone=active_zone,
            active_polo=polo,
        ),
        polo_query=f"polo={polo}",
        **context,
    )


def _build_tabs(upload_id: str, batch: PoloBatch, *, active_polo: str | None) -> list[dict]:
    """Flat fallback tab strip (Todos + every polo). Kept for the integration
    tests that pin the legacy single-row behaviour; the hierarchical builder
    below is what production uses."""
    base = url_for("main.dashboard", upload_id=upload_id)
    tabs = [
        {
            "label": "Todos",
            "href": f"{base}?polo=__all",
            "active": active_polo in (None, "__all"),
        }
    ]
    for polo in batch.polos:
        tabs.append(
            {
                "label": polo.title(),
                "href": f"{base}?polo={polo}",
                "active": active_polo == polo,
            }
        )
    return tabs


def _build_hierarchical_tabs(
    upload_id: str,
    batch: PoloBatch,
    *,
    active_zone: str | None,
    active_municipality: str | None,
    active_polo: str | None,
) -> dict:
    """Build three tab rows: Zone → Município → Polo.

    Each row is suppressed (empty list) when the level has only one entry
    *and* no narrower filter has been applied yet — so a single-zone batch
    skips the zone row, etc. This keeps the UI uncluttered when there's
    nothing to choose at a level.

    Active level coalesces upward: passing only ``active_polo`` resolves
    its zone + município from the catalog so the right parent tab lights up.
    """
    from app.core.aggregator import polo_geography

    base = url_for("main.dashboard", upload_id=upload_id)

    # Catalogue: one entry per file with its zone + município.
    catalog: list[dict[str, str]] = []
    for f in batch.files:
        muni, zone = polo_geography(f)
        catalog.append({"polo": f.polo, "municipality": muni, "zone": zone})

    # If a polo is active without an explicit zone/município, derive parents.
    if active_polo and active_polo != "__all":
        for entry in catalog:
            if entry["polo"] == active_polo:
                active_zone = active_zone or entry["zone"]
                active_municipality = active_municipality or entry["municipality"]
                break

    zones = sorted({e["zone"] for e in catalog})
    zone_tabs = [
        {
            "label": "Todas as zonas",
            "href": _scope_url(upload_id),
            "active": active_zone is None and active_polo in (None, "__all"),
        }
    ]
    for z in zones:
        zone_tabs.append(
            {
                "label": z,
                "href": _scope_url(upload_id, zone=z),
                "active": active_zone == z and active_municipality is None,
            }
        )

    municipality_tabs: list[dict] = []
    if active_zone is not None:
        munis_in_zone = sorted(
            {e["municipality"] or "—" for e in catalog if e["zone"] == active_zone}
        )
        if len(munis_in_zone) > 1 or active_municipality is not None:
            municipality_tabs.append(
                {
                    "label": "Todas",
                    "href": _scope_url(upload_id, zone=active_zone),
                    "active": active_municipality is None,
                }
            )
            for m in munis_in_zone:
                municipality_tabs.append(
                    {
                        "label": m,
                        "href": f"{base}?zone={active_zone}&municipality={m}",
                        "active": active_municipality == m,
                    }
                )

    polo_tabs: list[dict] = []
    if active_municipality is not None or (active_zone is not None and active_polo):
        polos_in_scope = sorted(
            {
                e["polo"]
                for e in catalog
                if (active_zone is None or e["zone"] == active_zone)
                and (
                    active_municipality is None or (e["municipality"] or "—") == active_municipality
                )
            }
        )
        if len(polos_in_scope) > 1 or active_polo:
            for p in polos_in_scope:
                polo_tabs.append(
                    {
                        "label": p.title(),
                        "href": _scope_url(upload_id, zone=active_zone, polo=p)
                        if active_zone
                        else f"{base}?polo={p}",
                        "active": active_polo == p,
                    }
                )

    # Hide the zone row when the batch covers exactly one zone and no narrower
    # filter is in play — nothing for the user to switch to.
    if len(zones) == 1 and active_municipality is None and active_polo in (None, "__all"):
        zone_tabs = []

    return {
        "zone_tabs": zone_tabs,
        "municipality_tabs": municipality_tabs,
        "polo_tabs": polo_tabs,
    }


_MONTH_PT = {
    "01": "Janeiro",
    "02": "Fevereiro",
    "03": "Março",
    "04": "Abril",
    "05": "Maio",
    "06": "Junho",
    "07": "Julho",
    "08": "Agosto",
    "09": "Setembro",
    "10": "Outubro",
    "11": "Novembro",
    "12": "Dezembro",
}


def _period_options(batch: PoloBatch, view: str) -> list[dict[str, str]]:
    """Friendly labels for the period dropdown — date range for weeks,
    month-name for months. Falls back to the raw key on unexpected input."""
    if view == "weekly":
        options = []
        seen: set[str] = set()
        for f in sorted(batch.files, key=lambda x: x.period_start):
            if f.iso_week in seen:
                continue
            seen.add(f.iso_week)
            label = f"{f.iso_week} · {f.period_start:%m-%d} a {f.period_end:%m-%d-%Y}"
            options.append({"key": f.iso_week, "label": label})
        return options
    options = []
    for month in batch.months:
        year, mm = month.split("-", 1)
        options.append({"key": month, "label": f"{_MONTH_PT.get(mm, mm)} {year}"})
    return options


def _render_batch_dashboard(
    upload_id: str,
    batch_dir: Path,
    *,
    zone: str | None = None,
    municipality: str | None = None,
) -> str:
    from app.core.aggregator import polo_geography

    batch = load_batch(batch_dir)
    available_polos = batch.polos
    available_weeks = batch.iso_weeks
    available_months = batch.months

    view = request.args.get("view", "weekly")
    if view not in ("weekly", "monthly"):
        view = "weekly"
    period = request.args.get("period") or (
        (available_weeks[-1] if available_weeks else "")
        if view == "weekly"
        else (available_months[-1] if available_months else "")
    )

    # Narrow the batch by zone/município before selecting polos so the chart
    # context only sees the in-scope data. polo_geography is the source of
    # truth for both axes — reuse here.
    if zone is not None or municipality is not None:
        filtered_files = [
            f
            for f in batch.files
            if (zone is None or polo_geography(f)[1] == zone)
            and (municipality is None or (polo_geography(f)[0] or "—") == municipality)
        ]
        scope_polos = tuple({f.polo for f in filtered_files}) or tuple(available_polos)
    else:
        scope_polos = tuple(available_polos)

    selected_polos = tuple(request.args.getlist("polos")) or scope_polos
    selected_polos = tuple(p for p in selected_polos if p in scope_polos)
    if not selected_polos:
        selected_polos = scope_polos

    scope_level = "zone" if zone else "city"
    context = _build_batch_context(upload_id, batch, selected_polos, view, period, scope_level)
    hierarchical = _build_hierarchical_tabs(
        upload_id,
        batch,
        active_zone=zone,
        active_municipality=municipality,
        active_polo=None,
    )
    context.update(
        {
            "download_action": url_for("main.download", upload_id=upload_id),
            "tabs": _build_tabs(upload_id, batch, active_polo="__all"),
            "zone_tabs": hierarchical["zone_tabs"],
            "municipality_tabs": hierarchical["municipality_tabs"],
            "polo_tabs": hierarchical["polo_tabs"],
            "active_zone": zone,
            "active_municipality": municipality,
            "active_polo": "__all",
            "polo_query": "polo=__all",
            "available_polos": available_polos,
            "selected_polos": list(selected_polos),
            "available_weeks": available_weeks,
            "available_months": available_months,
            "available_periods": _period_options(batch, view),
            "selected_period": period,
            "view": view,
            "is_batch": True,
            "breadcrumbs": _breadcrumb(
                upload_id,
                zone=zone,
                municipality=municipality,
                polo=None,
            ),
            "scope_level": scope_level,
            "scope_sidebar": _build_scope_sidebar(
                upload_id,
                batch,
                active_zone=zone,
                active_polo="__all",
            ),
        }
    )
    return render_template("dashboard.html", **context)


def _build_chart_context(
    template: PimentasTemplate,
    inspections: pd.DataFrame,
    failures: pd.DataFrame,
    iqs_rows: list,
    ic_rows: list,
    iqs_overall: float | None,
) -> dict[str, Any]:
    """The polo-agnostic slice of the dashboard context — every Plotly chart
    HTML, the per-service sections, the top-observation lists, the simple
    counts. Single-Polo and batch dashboards layer their own polo_name,
    periodo, and filter metadata on top.

    Sharing this avoids the previous ~50-line cut-and-paste between
    _build_polo_context and _build_batch_context.
    """
    services = sorted(template.SERVICE_SHEETS)
    per_service_sections = []
    for idx, service in enumerate(services):
        team_html = _chart_html(
            template.build_team_conformity_for_service(inspections, service),
            div_id=f"conf-team-{idx}",
            filename=f"chart_conformidade_equipe_{idx}",
        )
        tss_html = _chart_html(
            template.build_tss_conformity_for_service(inspections, service),
            div_id=f"conf-tss-{idx}",
            filename=f"chart_conformidade_tss_{idx}",
        )
        per_service_sections.append(
            {
                "service": service,
                "team_chart": _defer_plotly_script(team_html),
                "tss_chart": _defer_plotly_script(tss_html),
            }
        )
    return {
        "iqs_overall": iqs_overall,
        "total_fotos": sum(r.fotos_avaliadas for r in iqs_rows),
        "total_inspections": len(inspections),
        "fig_ic_bar": _chart_html(
            template.build_ic_bar(ic_rows),
            div_id="ic-bar",
            filename="chart_ic_por_servico",
        ),
        "fig_iqs_bar": _chart_html(
            template.build_service_iqs_bar(iqs_rows),
            div_id="iqs-bar",
            filename="chart_iqs_por_servico",
        ),
        "fig_photos": _chart_html(
            template.build_photo_conformity_stacked(iqs_rows),
            div_id="photos",
            filename="chart_fotos_conformidade",
        ),
        "fig_team_service": _chart_html(
            template.build_team_service_stacked(inspections),
            div_id="team-service",
            filename="chart_equipe_servico",
        ),
        "fig_tss": _chart_html(
            template.build_tss_distribution(inspections),
            div_id="tss-distribution",
            filename="chart_tss_distribuicao",
        ),
        "fig_failing_stages": _chart_html(
            template.build_top_failing_stages(failures),
            div_id="failing-stages",
            filename="chart_etapas_falha",
        ),
        "fig_worst_teams": _chart_html(
            template.build_worst_teams(inspections),
            div_id="worst-teams",
            filename="chart_equipes_criticas",
        ),
        "scope_charts": [],
        "scope_rankings": [],
        "top_nc_observations": top_observations(failures, "NC"),
        "top_sf_observations": top_observations(failures, "SF"),
        "total_failing_os": int(inspections["nao_conforme_count"].sum())
        if not inspections.empty
        else 0,
        "per_service_sections": per_service_sections,
        "teams_sorted": sorted(inspections["team"].dropna().unique().tolist())
        if not inspections.empty
        else [],
    }


def _chart_html(fig, *, div_id: str, filename: str) -> str:
    """Render a Plotly figure with the dashboard's shared interaction config."""
    return fig.to_html(
        include_plotlyjs=False,
        full_html=False,
        div_id=div_id,
        config={
            "displaylogo": False,
            "modeBarButtonsToAdd": ["downloadImage"],
            "toImageButtonOptions": {
                "format": "png",
                "width": 1200,
                "height": 600,
                "scale": 2,
                "filename": filename,
            },
        },
    )


def _build_batch_context(
    uuid: str,
    batch: PoloBatch,
    polos: tuple[str, ...],
    view: str,
    period_key: str,
    scope_level: str = "city",
) -> dict[str, Any]:
    filtered = filter_batch(batch, polos=polos, view=view, period_key=period_key)
    inspections = dashboard_data.combined_inspections(uuid, filtered)
    failures = dashboard_data.combined_stage_failures(uuid, filtered)

    # Pick any concrete PimentasTemplate instance for SERVICE_SHEETS + build_* methods.
    template = PimentasTemplate()
    services = sorted(template.SERVICE_SHEETS)
    iqs_rows = _iqs_rows_from_inspections(inspections, services)
    ic_rows = _ic_rows_from_inspections(inspections, services)
    iqs_overall = _iqs_overall_from_inspections(inspections)

    polo_label = polos[0].title() if len(polos) == 1 else "Múltiplos Polos"
    context = _build_chart_context(template, inspections, failures, iqs_rows, ic_rows, iqs_overall)
    context.update(
        {
            "polo_name": polo_label,
            "periodo": format_period_pt(date_bounds(inspections)),
            "filter_start": "",
            "filter_end": "",
            "available_start": "",
            "available_end": "",
            "is_filtered": False,
            "swap_dates": False,
            "recomputed": True,
            "span_warning": None,
            "scope_charts": _scope_charts(inspections, scope_level=scope_level),
            "scope_rankings": _scope_rankings(inspections, scope_level=scope_level),
        }
    )
    return context


def _scope_charts(inspections: pd.DataFrame, *, scope_level: str) -> list[dict[str, str]]:
    """Comparison charts for city and zone dashboard scopes."""
    if scope_level == "city":
        return [
            *_comparison_charts(inspections, group_col="zone", label="Zona"),
            _trend_chart(inspections),
            _treemap_chart(inspections),
        ]
    if scope_level == "zone":
        return [
            *_comparison_charts(inspections, group_col="polo", label="Município"),
            _trend_chart(inspections),
        ]
    return []


def _scope_rankings(inspections: pd.DataFrame, *, scope_level: str) -> list[dict[str, Any]]:
    """Sortable ranking table rows for city/zone scope dashboards."""
    if scope_level == "city":
        return _ranking_rows(inspections, group_col="zone")
    if scope_level == "zone":
        return _ranking_rows(inspections, group_col="polo")
    return []


def _ranking_rows(inspections: pd.DataFrame, *, group_col: str) -> list[dict[str, Any]]:
    if inspections.empty or group_col not in inspections.columns:
        return []
    grouped = _scope_grouped_metrics(inspections, group_col=group_col)
    return [
        {
            "label": row[group_col],
            "ic_pct": row["ic"],
            "iqs_pct": row["iqs"],
            "inspections": int(row["inspecoes"]),
            "photos": int(row["fotos"]),
        }
        for row in grouped.sort_values("iqs", ascending=False).to_dict("records")
    ]


def _comparison_charts(
    inspections: pd.DataFrame,
    *,
    group_col: str,
    label: str,
) -> list[dict[str, str]]:
    if inspections.empty or group_col not in inspections.columns:
        return []
    grouped = _scope_grouped_metrics(inspections, group_col=group_col)
    grouped = grouped.sort_values("ic", ascending=False)
    return [
        {
            "title": f"Comparação de IC por {label}",
            "chart": _chart_html(
                _comparison_bar(grouped, group_col=group_col, metric="ic", label=label),
                div_id=f"scope-ic-{group_col}",
                filename=f"chart_ic_por_{_slug(label)}",
            ),
        },
        {
            "title": f"Ranking de IQS por {label}",
            "chart": _chart_html(
                _comparison_bar(grouped, group_col=group_col, metric="iqs", label=label),
                div_id=f"scope-iqs-{group_col}",
                filename=f"chart_iqs_por_{_slug(label)}",
            ),
        },
        {
            "title": f"Volume de inspeções por {label}",
            "chart": _chart_html(
                _volume_bar(grouped, group_col=group_col, label=label),
                div_id=f"scope-volume-{group_col}",
                filename=f"chart_volume_por_{_slug(label)}",
            ),
        },
    ]


def _scope_grouped_metrics(inspections: pd.DataFrame, *, group_col: str) -> pd.DataFrame:
    grouped = (
        inspections.groupby(group_col, dropna=False)
        .agg(
            inspecoes=("conforme_count", "size"),
            conforme=("conforme_count", "sum"),
            fotos=("photo_total", "sum"),
            fotos_conforme=("photo_conforme", "sum"),
        )
        .reset_index()
    )
    grouped[group_col] = grouped[group_col].fillna("Sem classificação").astype(str)
    grouped["ic"] = grouped["conforme"] / grouped["inspecoes"]
    grouped["iqs"] = grouped["fotos_conforme"] / grouped["fotos"].where(grouped["fotos"] > 0)
    grouped["iqs"] = grouped["iqs"].fillna(0)
    return grouped


def _comparison_bar(
    grouped: pd.DataFrame,
    *,
    group_col: str,
    metric: str,
    label: str,
) -> go.Figure:
    ordered = grouped.sort_values(metric, ascending=True).tail(20)
    title_metric = "IC" if metric == "ic" else "IQS"
    values = ordered[metric].tolist()
    names = ordered[group_col].tolist()
    return go.Figure(
        data=[
            go.Bar(
                x=values,
                y=names,
                orientation="h",
                marker_color="#f97316" if metric == "ic" else "#22c55e",
                text=[f"{v:.1%}" for v in values],
                textposition="outside",
            )
        ],
        layout=go.Layout(
            title=f"{title_metric} por {label}",
            xaxis={"title": title_metric, "tickformat": ".0%", "range": [0, 1.05]},
            yaxis={"title": label, "automargin": True},
            template="polo_dark",
            height=max(360, 34 * len(names) + 120),
            margin={"r": 90},
        ),
    )


def _trend_chart(inspections: pd.DataFrame) -> dict[str, str]:
    if inspections.empty or "start_date" not in inspections.columns:
        fig = _empty_scope_figure("Sem datas para tendência")
    else:
        dated = inspections.dropna(subset=["start_date"]).copy()
        if dated.empty:
            fig = _empty_scope_figure("Sem datas para tendência")
        else:
            dated["period"] = dated["start_date"].dt.to_period("D").dt.to_timestamp()
            grouped = (
                dated.groupby("period")
                .agg(
                    inspecoes=("conforme_count", "size"),
                    conforme=("conforme_count", "sum"),
                    fotos=("photo_total", "sum"),
                    fotos_conforme=("photo_conforme", "sum"),
                )
                .reset_index()
                .sort_values("period")
            )
            grouped["ic"] = grouped["conforme"] / grouped["inspecoes"]
            grouped["iqs"] = grouped["fotos_conforme"] / grouped["fotos"].where(
                grouped["fotos"] > 0
            )
            grouped["iqs"] = grouped["iqs"].fillna(0)
            fig = go.Figure(
                data=[
                    go.Scatter(
                        x=grouped["period"],
                        y=grouped["iqs"],
                        mode="lines+markers",
                        name="IQS",
                        line={"color": "#22c55e", "width": 3},
                    ),
                    go.Scatter(
                        x=grouped["period"],
                        y=grouped["ic"],
                        mode="lines+markers",
                        name="IC",
                        line={"color": "#f97316", "width": 3},
                    ),
                ],
                layout=go.Layout(
                    title="Tendência diária de IC e IQS",
                    xaxis={"title": "Data"},
                    yaxis={"title": "Percentual", "tickformat": ".0%", "range": [0, 1.05]},
                    template="polo_dark",
                    height=380,
                ),
            )
    return {
        "title": "Tendência do período",
        "chart": _chart_html(fig, div_id="scope-trend", filename="chart_tendencia_escopo"),
    }


def _treemap_chart(inspections: pd.DataFrame) -> dict[str, str]:
    if inspections.empty or not {"zone", "polo"}.issubset(inspections.columns):
        fig = _empty_scope_figure("Sem municípios para mapa")
    else:
        grouped = _scope_grouped_metrics(inspections, group_col="polo")
        zone_map = (
            inspections.groupby("polo")["zone"]
            .agg(lambda s: s.dropna().astype(str).mode().iloc[0] if not s.dropna().empty else "")
            .to_dict()
        )
        grouped["zone"] = grouped["polo"].map(zone_map).fillna("Sem classificação")
        fig = go.Figure(
            data=[
                go.Treemap(
                    labels=grouped["polo"],
                    parents=grouped["zone"],
                    values=grouped["inspecoes"],
                    marker={
                        "colors": grouped["iqs"],
                        "colorscale": "RdYlGn",
                        "cmin": 0,
                        "cmax": 1,
                        "colorbar": {"title": "IQS"},
                    },
                    texttemplate="%{label}<br>%{value} inspeções",
                )
            ],
            layout=go.Layout(
                title="Mapa de calor por município",
                template="polo_dark",
                height=460,
            ),
        )
    return {
        "title": "Mapa de calor por município",
        "chart": _chart_html(fig, div_id="scope-treemap", filename="chart_mapa_municipios"),
    }


def _empty_scope_figure(message: str) -> go.Figure:
    return go.Figure(
        layout=go.Layout(
            template="polo_dark",
            annotations=[
                {
                    "text": message,
                    "showarrow": False,
                    "xref": "paper",
                    "yref": "paper",
                    "x": 0.5,
                    "y": 0.5,
                    "font": {"size": 14, "color": "#a1a1aa"},
                }
            ],
        )
    )


def _volume_bar(grouped: pd.DataFrame, *, group_col: str, label: str) -> go.Figure:
    ordered = grouped.sort_values("inspecoes", ascending=True).tail(20)
    values = ordered["inspecoes"].astype(int).tolist()
    names = ordered[group_col].tolist()
    return go.Figure(
        data=[
            go.Bar(
                x=values,
                y=names,
                orientation="h",
                marker_color="#0ea5e9",
                text=values,
                textposition="outside",
            )
        ],
        layout=go.Layout(
            title=f"Inspeções por {label}",
            xaxis={"title": "Inspeções"},
            yaxis={"title": label, "automargin": True},
            template="polo_dark",
            height=max(360, 34 * len(names) + 120),
            margin={"r": 90},
        ),
    )


@bp.get("/report/<upload_id>")
def report(upload_id: str) -> str:
    """Formal "Gerar Relatório" view — formatted report with download buttons.

    Single Polo: ``?polo=<NAME>`` → per-Polo report (docx / pdf / pptx downloads).
    Combined: ``?polo=__all`` → cross-Polo summary (md / xlsx / html downloads).
    """
    kind, target = _resolve_upload(upload_id)
    polo_arg = request.args.get("polo", "__all")
    download_action = url_for("main.download", upload_id=upload_id)

    if kind == "legacy":
        path = target
        template, workbook = _open_pimentas(path)
        context = _build_polo_context(template, workbook, path)
        context["per_service_sections"] = _undefer_per_service_sections(
            context["per_service_sections"]
        )
        return render_template(
            "report.html",
            scope_label=template.polo_name.title(),
            scope_kind="single",
            download_action=download_action,
            polo_query="",
            **context,
        )

    batch_dir = target
    if polo_arg and polo_arg != "__all":
        path = _batch_polo_path(batch_dir, polo_arg)
        template, workbook = _open_pimentas(path)
        context = _build_polo_context(template, workbook, path)
        context["per_service_sections"] = _undefer_per_service_sections(
            context["per_service_sections"]
        )
        return render_template(
            "report.html",
            scope_label=polo_arg.title(),
            scope_kind="single",
            download_action=download_action,
            polo_query=f"polo={polo_arg}",
            **context,
        )

    batch = load_batch(batch_dir)
    selected_polos = tuple(batch.polos)
    view = request.args.get("view", "weekly")
    if view not in ("weekly", "monthly"):
        view = "weekly"
    period = request.args.get("period") or (
        (batch.iso_weeks[-1] if batch.iso_weeks else "")
        if view == "weekly"
        else (batch.months[-1] if batch.months else "")
    )
    context = _build_batch_context(upload_id, batch, selected_polos, view, period)
    context["polos_included"] = list(selected_polos)
    context["per_service_sections"] = _undefer_per_service_sections(context["per_service_sections"])
    return render_template(
        "report.html",
        scope_label="Múltiplos Polos",
        scope_kind="combined",
        download_action=download_action,
        polo_query="polo=__all",
        **context,
    )


@bp.get("/dashboard/<upload_id>/team")
def team_detail(upload_id: str) -> str:
    team_name = (request.args.get("name") or "").strip()
    if not team_name:
        abort(400)

    kind, target = _resolve_upload(upload_id)
    if kind == "batch":
        polo_arg = (request.args.get("polo") or "").strip()
        if not polo_arg or polo_arg == "__all":
            abort(400)
        path = _batch_polo_path(target, polo_arg)
        dashboard_url = url_for("main.dashboard", upload_id=upload_id) + f"?polo={polo_arg}"
    else:
        path = target
        dashboard_url = url_for("main.dashboard", upload_id=upload_id)

    template, _workbook = _open_pimentas(path)
    detail = template.extract_team_detail(path, team_name)
    if not detail:
        abort(404)
    return render_template(
        "team_detail.html",
        polo_name=template.polo_name.title(),
        team_name=team_name,
        dashboard_url=dashboard_url,
        detail=detail,
    )


_SUPPORTED_FORMATS = {"md", "xlsx", "pdf", "docx", "pptx"}
_BATCH_FORMATS = {"md", "xlsx", "html", "docx", "pdf"}
_DOCX_STYLES = {"sabesp_mensal", "sabesp_semanal", "generic"}


def _docx_style_arg() -> str | None:
    """Whitelist the ?style= query arg so callers can't smuggle paths into the loader."""
    raw = request.args.get("style", "").strip().lower()
    return raw if raw in _DOCX_STYLES else None


@bp.get("/download/<upload_id>")
def download(upload_id: str) -> Response:
    fmt = request.args.get("fmt", "md").lower()

    kind, target = _resolve_upload(upload_id)
    if kind == "batch":
        polo_arg = request.args.get("polo")
        if polo_arg and polo_arg != "__all":
            if fmt not in _SUPPORTED_FORMATS:
                abort(400)
            path = _batch_polo_path(target, polo_arg)
            template, workbook = _open_pimentas(path)
            body, mimetype = render_export(fmt, template, workbook, path, style=_docx_style_arg())
            response = Response(body, mimetype=mimetype)
            response.headers["Content-Disposition"] = (
                f'attachment; filename="dashboard-{polo_arg}-{upload_id[:8]}.{fmt}"'
            )
            return response

        if fmt not in _BATCH_FORMATS:
            abort(501)
        return _download_batch(upload_id, target, fmt)

    if fmt not in _SUPPORTED_FORMATS:
        abort(400)

    path = _upload_path(upload_id)
    template, workbook = _open_pimentas(path)

    body, mimetype = render_export(fmt, template, workbook, path, style=_docx_style_arg())
    response = Response(body, mimetype=mimetype)
    response.headers["Content-Disposition"] = f'attachment; filename="dashboard-{upload_id}.{fmt}"'
    return response


@bp.get("/templates/docx/<style>")
def docx_skeleton(style: str) -> Response:
    """Serve the raw, unbound docxtpl skeleton so users can hand-edit the layout.

    Whitelisted via ``_DOCX_STYLES`` — bare filename, no traversal possible.
    """
    if style not in _DOCX_STYLES:
        abort(404)
    from app.core.exporters import docx_sabesp

    skeleton_path = {
        "sabesp_mensal": docx_sabesp.MENSAL_SKELETON_PATH,
        "sabesp_semanal": docx_sabesp.SEMANAL_SKELETON_PATH,
    }[style]
    if not skeleton_path.exists():
        abort(404)
    body = skeleton_path.read_bytes()
    response = Response(
        body,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    response.headers["Content-Disposition"] = f'attachment; filename="{style}_template.docx"'
    return response


def _download_batch(upload_id: str, batch_dir: Path, fmt: str) -> Response:
    from app.core.exporters.batch import BatchSelection, render_batch_export

    batch = load_batch(batch_dir)
    view = request.args.get("view", "weekly")
    if view not in ("weekly", "monthly"):
        view = "weekly"
    period = request.args.get("period") or (
        (batch.iso_weeks[-1] if batch.iso_weeks else "")
        if view == "weekly"
        else (batch.months[-1] if batch.months else "")
    )
    polos = tuple(request.args.getlist("polos")) or tuple(batch.polos)
    polos = tuple(p for p in polos if p in batch.polos)
    if not polos:
        polos = tuple(batch.polos)

    selection = BatchSelection(polos=polos, view=view, period_key=period)
    body, mimetype = render_batch_export(fmt, batch, selection, style=_docx_style_arg())
    response = Response(body, mimetype=mimetype)
    response.headers["Content-Disposition"] = f'attachment; filename="dashboard-{upload_id}.{fmt}"'
    return response


def _upload_path(upload_id: str) -> Path:
    path = Path(current_app.instance_path) / "uploads" / f"{upload_id}.xlsx"
    if not path.exists():
        abort(404)
    return path


def _resolve_upload(upload_id: str) -> tuple[str, Path]:
    """Return either ("batch", batch_dir) or ("legacy", xlsx_path), 404 if neither."""
    uploads_dir = Path(current_app.instance_path) / "uploads"
    batch_dir = uploads_dir / upload_id
    if batch_dir.is_dir() and (batch_dir / MANIFEST_FILENAME).exists():
        return "batch", batch_dir
    legacy = uploads_dir / f"{upload_id}.xlsx"
    if legacy.exists():
        return "legacy", legacy
    abort(404)


def _open_pimentas(path: Path):
    """Open ``path`` read-only and recognize the template; ``abort(404)`` when it
    isn't a Pimentas workbook.

    Used by routes that don't need the unknown-template fallback page
    (``dashboard.html`` keeps its own branch since it renders ``dashboard_unknown.html``
    instead of 404'ing).
    """
    workbook = load_workbook(path, data_only=True, read_only=True)
    template = recognize(workbook.sheetnames)
    if not isinstance(template, PimentasTemplate):
        abort(404)
    return template, workbook


def _parse_iso_date(value: str) -> pd.Timestamp | None:
    """Parse YYYY-MM-DD from an <input type=date>; return None when invalid."""
    if not value:
        return None
    try:
        return pd.Timestamp(value)
    except (ValueError, TypeError):
        return None


_SUSPICIOUS_SPAN_DAYS = 60


def _build_polo_context(
    template: PimentasTemplate,
    workbook,
    path: Path,
    filter_start: pd.Timestamp | None = None,
    filter_end: pd.Timestamp | None = None,
    swap_dates: bool | None = None,
) -> dict[str, Any]:
    full_inspections = template.extract_inspections(path)
    full_failures = template.extract_stage_failures(path)

    # Auto-detect day/month swap unless the user explicitly opted in or out.
    if swap_dates is None:
        raw_start, raw_end = _date_bounds(full_inspections)
        if (
            raw_start is not None
            and raw_end is not None
            and (raw_end - raw_start).days > _SUSPICIOUS_SPAN_DAYS
        ):
            swap_dates = True
        else:
            swap_dates = False

    if swap_dates:
        full_inspections = _swap_day_month(full_inspections)
        full_failures = _swap_day_month(full_failures)

    available_start, available_end = _date_bounds(full_inspections)
    span_warning = _date_span_warning(available_start, available_end, swapped=swap_dates)
    inspections = _apply_date_filter(full_inspections, filter_start, filter_end)
    failures = _apply_date_filter(full_failures, filter_start, filter_end)
    is_filtered = filter_start is not None or filter_end is not None
    recomputed = is_filtered or swap_dates

    if recomputed:
        services = sorted(template.SERVICE_SHEETS)
        iqs_rows = _iqs_rows_from_inspections(inspections, services)
        ic_rows = _ic_rows_from_inspections(inspections, services)
        iqs_overall = _iqs_overall_from_inspections(inspections)
    else:
        iqs_rows = template.extract_iqs_by_service(workbook)
        ic_rows = template.extract_ic_by_service(workbook)
        iqs_overall = template.extract_iqs_overall(workbook)

    periodo = format_period_pt(date_bounds(inspections)) or template.extract_periodo(workbook)
    context = _build_chart_context(template, inspections, failures, iqs_rows, ic_rows, iqs_overall)
    context.update(
        {
            "polo_name": template.polo_name.title(),
            "periodo": periodo,
            "filter_start": _iso(filter_start),
            "filter_end": _iso(filter_end),
            "available_start": _iso(available_start),
            "available_end": _iso(available_end),
            "is_filtered": is_filtered,
            "swap_dates": swap_dates,
            "recomputed": recomputed,
            "span_warning": span_warning,
        }
    )
    return context


def _defer_plotly_script(chart_html: str) -> str:
    """Mark a Plotly chart's inline script so it does not execute on page load.

    The dashboard bootstrap (see dashboard.html) revives these scripts when
    the surrounding section scrolls into view. Plotly emits exactly one
    untyped <script> per chart, so a single replacement is sufficient and
    unambiguous.
    """
    return chart_html.replace("<script>", '<script type="text/plotly-defer">', 1)


def _undefer_plotly_script(chart_html: str) -> str:
    """Inverse of ``_defer_plotly_script`` — used by views that don't carry the
    IntersectionObserver bootstrap (e.g. report.html), where a deferred script
    would never execute and the chart would render as an empty <div>.
    """
    return chart_html.replace('<script type="text/plotly-defer">', "<script>", 1)


def _undefer_per_service_sections(sections: list[dict]) -> list[dict]:
    """Strip the dashboard's scroll-revive marker from per-service chart HTMLs.

    Section 7 (Detalhamento por Serviço) is the only place the report template
    consumes deferred charts; sections 1–6 use the raw chart HTML directly.
    """
    return [
        {
            **sec,
            "team_chart": _undefer_plotly_script(sec["team_chart"]),
            "tss_chart": _undefer_plotly_script(sec["tss_chart"]),
        }
        for sec in sections
    ]


@lru_cache(maxsize=64)
def _cached_polo_context(
    path_str: str,
    mtime_ns: int,  # noqa: ARG001 — cache key only; invalidates when the file changes
    filter_start_iso: str,
    filter_end_iso: str,
    swap_arg: str,
) -> dict[str, Any]:
    """Memoised dashboard context — same file + same filter → reuse rendered figures.

    ``swap_arg`` is the URL value: ``""`` = auto-detect, ``"1"`` = force on,
    ``"0"`` = force off. Kept as a string so the lru_cache key stays hashable
    and stable across requests.
    """
    path = Path(path_str)
    workbook = load_workbook(path, data_only=True, read_only=True)
    template = recognize(workbook.sheetnames)
    assert isinstance(template, PimentasTemplate)  # route guards this
    return _build_polo_context(
        template,
        workbook,
        path,
        filter_start=_parse_iso_date(filter_start_iso),
        filter_end=_parse_iso_date(filter_end_iso),
        swap_dates=_swap_arg_to_bool(swap_arg),
    )


def _date_bounds(df: pd.DataFrame) -> tuple[pd.Timestamp | None, pd.Timestamp | None]:
    if df.empty or "start_date" not in df.columns:
        return None, None
    dates = df["start_date"].dropna()
    if dates.empty:
        return None, None
    return dates.min().normalize(), dates.max().normalize()


def _apply_date_filter(
    df: pd.DataFrame,
    start: pd.Timestamp | None,
    end: pd.Timestamp | None,
) -> pd.DataFrame:
    if df.empty or "start_date" not in df.columns:
        return df
    mask = pd.Series(True, index=df.index)
    if start is not None:
        mask &= df["start_date"] >= start
    if end is not None:
        # End is inclusive on the day — bump to end of day.
        mask &= df["start_date"] < (end + pd.Timedelta(days=1))
    return df[mask]


def _date_span_warning(
    start: pd.Timestamp | None, end: pd.Timestamp | None, swapped: bool = False
) -> str | None:
    if start is None or end is None:
        return None
    span_days = (end - start).days
    if span_days <= _SUSPICIOUS_SPAN_DAYS:
        return None
    if swapped:
        return (
            f"Datas com dia/mês invertidos: {span_days} dias "
            f"({start:%m-%d-%Y} → {end:%m-%d-%Y}). O resultado ainda parece "
            "incorreto; verifique a planilha original."
        )
    return (
        f"Atenção: as datas das inspeções abrangem {span_days} dias "
        f"({start:%m-%d-%Y} → {end:%m-%d-%Y}). Isso pode indicar dia/mês "
        "trocados na planilha original."
    )


def _swap_day_month(df: pd.DataFrame) -> pd.DataFrame:
    """Flip day/month only when the swap moves a row into the target month.

    Polo reports occasionally store dates as MM/DD instead of DD/MM at
    data entry. Unambiguous rows (day > 12) reveal the file's actual
    target month; ambiguous rows (day ≤ 12) are swapped only when the
    swap puts them into that target month. This avoids breaking the
    correctly-entered dates that happen to have day ≤ 12.
    """
    if df.empty or "start_date" not in df.columns:
        return df
    dates = df["start_date"]
    unambiguous = dates.notna() & (dates.dt.day > 12)
    if not unambiguous.any():
        return df  # nothing tells us which month is "right"
    target_month = int(dates[unambiguous].dt.month.mode().iloc[0])

    df = df.copy()
    ambiguous = dates.notna() & (dates.dt.day <= 12)
    # Swap only when swap-day-to-month would yield the target month.
    swap_mask = ambiguous & (dates.dt.day == target_month) & (dates.dt.month != target_month)
    if not swap_mask.any():
        return df
    sub = dates[swap_mask]
    df.loc[swap_mask, "start_date"] = pd.to_datetime(
        {
            "year": sub.dt.year,
            "month": sub.dt.day,
            "day": sub.dt.month,
            "hour": sub.dt.hour,
            "minute": sub.dt.minute,
        }
    ).values
    return df


# Helpers _iqs_rows_from_inspections / _ic_rows_from_inspections /
# _iqs_overall_from_inspections moved to app.core.aggregator and re-imported
# above so the batch dashboard and single-file dashboard share one path.


def _iso(ts: pd.Timestamp | None) -> str:
    return ts.strftime("%Y-%m-%d") if ts is not None else ""
