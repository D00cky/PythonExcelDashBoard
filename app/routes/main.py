import uuid
from functools import lru_cache
from pathlib import Path
from typing import Any

import pandas as pd
from flask import (
    Blueprint,
    Response,
    abort,
    current_app,
    redirect,
    render_template,
    request,
    url_for,
)
from openpyxl import load_workbook

from app.core.aggregator import (
    MANIFEST_FILENAME,
    PoloBatch,
    combined_inspections,
    combined_stage_failures,
    discover_polo_file,
    filter_batch,
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
    return redirect(url_for("main.dashboard", upload_id=upload_id), code=303)


@bp.get("/dashboard/<upload_id>")
def dashboard(upload_id: str) -> str:
    kind, target = _resolve_upload(upload_id)
    if kind == "batch":
        polo_arg = request.args.get("polo")
        if polo_arg and polo_arg != "__all":
            path = _batch_polo_path(target, polo_arg)
            return _render_single_polo_dashboard(upload_id, path, batch_dir=target, polo=polo_arg)
        return _render_batch_dashboard(upload_id, target)

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


def _swap_arg() -> str:
    """Normalize ?swap= into a cache-friendly key. '' = auto-detect."""
    raw = request.args.get("swap")
    if raw == "1":
        return "1"
    if raw == "0":
        return "0"
    return ""


def _swap_arg_to_bool(value: str) -> bool | None:
    return {"1": True, "0": False, "": None}.get(value)


def _batch_polo_path(batch_dir: Path, polo: str) -> Path:
    """Return the xlsx path for ``polo`` inside the batch, or abort 404."""
    batch = load_batch(batch_dir)
    for pf in batch.files:
        if pf.polo == polo:
            return pf.file_path
    abort(404)


def _render_single_polo_dashboard(upload_id: str, path: Path, *, batch_dir: Path, polo: str) -> str:
    """Render the legacy single-file dashboard, but tagged as one tab in a batch."""
    workbook = load_workbook(path, data_only=True, read_only=True)
    template = recognize(workbook.sheetnames)
    if not isinstance(template, PimentasTemplate):
        abort(404)

    batch = load_batch(batch_dir)
    filter_start = _parse_iso_date(request.args.get("start", ""))
    filter_end = _parse_iso_date(request.args.get("end", ""))

    context = _cached_polo_context(
        str(path),
        path.stat().st_mtime_ns,
        _iso(filter_start),
        _iso(filter_end),
        _swap_arg(),
    )
    return render_template(
        "dashboard.html",
        download_action=url_for("main.download", upload_id=upload_id),
        tabs=_build_tabs(upload_id, batch, active_polo=polo),
        active_polo=polo,
        polo_query=f"polo={polo}",
        **context,
    )


def _build_tabs(upload_id: str, batch: PoloBatch, *, active_polo: str | None) -> list[dict]:
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
            label = f"{f.iso_week} · {f.period_start:%d/%m} a {f.period_end:%d/%m/%Y}"
            options.append({"key": f.iso_week, "label": label})
        return options
    options = []
    for month in batch.months:
        year, mm = month.split("-", 1)
        options.append({"key": month, "label": f"{_MONTH_PT.get(mm, mm)} {year}"})
    return options


def _render_batch_dashboard(upload_id: str, batch_dir: Path) -> str:
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
    selected_polos = tuple(request.args.getlist("polos")) or tuple(available_polos)
    selected_polos = tuple(p for p in selected_polos if p in available_polos)
    if not selected_polos:
        selected_polos = tuple(available_polos)

    context = _build_batch_context(batch, selected_polos, view, period)
    context.update(
        {
            "download_action": url_for("main.download", upload_id=upload_id),
            "tabs": _build_tabs(upload_id, batch, active_polo="__all"),
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
        }
    )
    return render_template("dashboard.html", **context)


def _build_batch_context(
    batch: PoloBatch,
    polos: tuple[str, ...],
    view: str,
    period_key: str,
) -> dict[str, Any]:
    filtered = filter_batch(batch, polos=polos, view=view, period_key=period_key)
    inspections = combined_inspections(filtered)
    failures = combined_stage_failures(filtered)

    # Pick any concrete PimentasTemplate instance for SERVICE_SHEETS + build_* methods.
    template = PimentasTemplate()
    services = sorted(template.SERVICE_SHEETS)
    iqs_rows = _iqs_rows_from_inspections(inspections, services)
    ic_rows = _ic_rows_from_inspections(inspections, services)
    iqs_overall = _iqs_overall_from_inspections(inspections)

    periodo = _periodo_from_inspections(inspections)
    teams_sorted = (
        sorted(inspections["team"].dropna().unique().tolist()) if not inspections.empty else []
    )
    polo_label = polos[0].title() if len(polos) == 1 else "Múltiplos Polos"

    per_service_sections = []
    for idx, service in enumerate(services):
        team_html = template.build_team_conformity_for_service(inspections, service).to_html(
            include_plotlyjs=False, full_html=False, div_id=f"conf-team-{idx}"
        )
        tss_html = template.build_tss_conformity_for_service(inspections, service).to_html(
            include_plotlyjs=False, full_html=False, div_id=f"conf-tss-{idx}"
        )
        per_service_sections.append(
            {
                "service": service,
                "team_chart": _defer_plotly_script(team_html),
                "tss_chart": _defer_plotly_script(tss_html),
            }
        )

    return {
        "polo_name": polo_label,
        "periodo": periodo,
        "iqs_overall": iqs_overall,
        "total_fotos": sum(r.fotos_avaliadas for r in iqs_rows),
        "total_inspections": len(inspections),
        "fig_ic_bar": template.build_ic_bar(ic_rows).to_html(
            include_plotlyjs=False, full_html=False, div_id="ic-bar"
        ),
        "fig_iqs_bar": template.build_service_iqs_bar(iqs_rows).to_html(
            include_plotlyjs=False, full_html=False, div_id="iqs-bar"
        ),
        "fig_photos": template.build_photo_conformity_stacked(iqs_rows).to_html(
            include_plotlyjs=False, full_html=False, div_id="photos"
        ),
        "fig_team_service": template.build_team_service_stacked(inspections).to_html(
            include_plotlyjs=False, full_html=False, div_id="team-service"
        ),
        "fig_tss": template.build_tss_distribution(inspections).to_html(
            include_plotlyjs=False, full_html=False, div_id="tss-distribution"
        ),
        "fig_failing_stages": template.build_top_failing_stages(failures).to_html(
            include_plotlyjs=False, full_html=False, div_id="failing-stages"
        ),
        "fig_worst_teams": template.build_worst_teams(inspections).to_html(
            include_plotlyjs=False, full_html=False, div_id="worst-teams"
        ),
        "top_nc_observations": top_observations(failures, "NC"),
        "top_sf_observations": top_observations(failures, "SF"),
        "total_failing_os": int(inspections["nao_conforme_count"].sum())
        if not inspections.empty
        else 0,
        "per_service_sections": per_service_sections,
        "teams_sorted": teams_sorted,
        "filter_start": "",
        "filter_end": "",
        "available_start": "",
        "available_end": "",
        "is_filtered": False,
        "swap_dates": False,
        "recomputed": True,
        "span_warning": None,
    }


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
        workbook = load_workbook(path, data_only=True, read_only=True)
        template = recognize(workbook.sheetnames)
        if not isinstance(template, PimentasTemplate):
            abort(404)
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
        workbook = load_workbook(path, data_only=True, read_only=True)
        template = recognize(workbook.sheetnames)
        if not isinstance(template, PimentasTemplate):
            abort(404)
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
    context = _build_batch_context(batch, selected_polos, view, period)
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

    workbook = load_workbook(path, data_only=True, read_only=True)
    template = recognize(workbook.sheetnames)
    if not isinstance(template, PimentasTemplate):
        abort(404)
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
_DOCX_STYLES = {"sabesp_mensal"}


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
            workbook = load_workbook(path, data_only=True, read_only=True)
            template = recognize(workbook.sheetnames)
            if not isinstance(template, PimentasTemplate):
                abort(404)
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
    workbook = load_workbook(path, data_only=True, read_only=True)
    template = recognize(workbook.sheetnames)
    if not isinstance(template, PimentasTemplate):
        abort(404)

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
        "sabesp_mensal": docx_sabesp.SKELETON_PATH,
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
    body, mimetype = render_batch_export(fmt, batch, selection)
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


def _periodo_from_inspections(df) -> str | None:
    if df.empty or "start_date" not in df.columns:
        return None
    dates = df["start_date"].dropna()
    if dates.empty:
        return None
    return f"{dates.min():%d/%m/%Y} à {dates.max():%d/%m/%Y}"


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

    periodo = _periodo_from_inspections(inspections) or template.extract_periodo(workbook)
    teams_sorted = (
        sorted(inspections["team"].dropna().unique().tolist()) if not inspections.empty else []
    )

    per_service_sections = []
    for idx, service in enumerate(sorted(template.SERVICE_SHEETS)):
        team_html = template.build_team_conformity_for_service(inspections, service).to_html(
            include_plotlyjs=False, full_html=False, div_id=f"conf-team-{idx}"
        )
        tss_html = template.build_tss_conformity_for_service(inspections, service).to_html(
            include_plotlyjs=False, full_html=False, div_id=f"conf-tss-{idx}"
        )
        per_service_sections.append(
            {
                "service": service,
                "team_chart": _defer_plotly_script(team_html),
                "tss_chart": _defer_plotly_script(tss_html),
            }
        )

    return {
        "polo_name": template.polo_name.title(),
        "periodo": periodo,
        "iqs_overall": iqs_overall,
        "total_fotos": sum(r.fotos_avaliadas for r in iqs_rows),
        "total_inspections": len(inspections),
        "fig_ic_bar": template.build_ic_bar(ic_rows).to_html(
            include_plotlyjs=False, full_html=False, div_id="ic-bar"
        ),
        "fig_iqs_bar": template.build_service_iqs_bar(iqs_rows).to_html(
            include_plotlyjs=False, full_html=False, div_id="iqs-bar"
        ),
        "fig_photos": template.build_photo_conformity_stacked(iqs_rows).to_html(
            include_plotlyjs=False, full_html=False, div_id="photos"
        ),
        "fig_team_service": template.build_team_service_stacked(inspections).to_html(
            include_plotlyjs=False, full_html=False, div_id="team-service"
        ),
        "fig_tss": template.build_tss_distribution(inspections).to_html(
            include_plotlyjs=False, full_html=False, div_id="tss-distribution"
        ),
        "fig_failing_stages": template.build_top_failing_stages(failures).to_html(
            include_plotlyjs=False, full_html=False, div_id="failing-stages"
        ),
        "fig_worst_teams": template.build_worst_teams(inspections).to_html(
            include_plotlyjs=False, full_html=False, div_id="worst-teams"
        ),
        "top_nc_observations": top_observations(failures, "NC"),
        "top_sf_observations": top_observations(failures, "SF"),
        "total_failing_os": int(inspections["nao_conforme_count"].sum())
        if not inspections.empty
        else 0,
        "per_service_sections": per_service_sections,
        "teams_sorted": teams_sorted,
        "filter_start": _iso(filter_start),
        "filter_end": _iso(filter_end),
        "available_start": _iso(available_start),
        "available_end": _iso(available_end),
        "is_filtered": is_filtered,
        "swap_dates": swap_dates,
        "recomputed": recomputed,
        "span_warning": span_warning,
    }


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
            f"({start:%d/%m/%Y} → {end:%d/%m/%Y}). O resultado ainda parece "
            "incorreto; verifique a planilha original."
        )
    return (
        f"Atenção: as datas das inspeções abrangem {span_days} dias "
        f"({start:%d/%m/%Y} → {end:%d/%m/%Y}). Isso pode indicar dia/mês "
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
