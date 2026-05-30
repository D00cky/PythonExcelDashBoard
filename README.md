# PythonExcelDashBoard

Flask web app that turns an uploaded weekly polo audit report (`.xlsx`) into an
interactive Plotly dashboard, with downloadable summaries in Markdown, Excel, PDF, and Word.

Built for the "Polo Pimentas" report template, but the template layer is pluggable —
adding support for another polo's layout is a single module under `app/core/templates/`.

## Features

- **Upload** a weekly `.xlsx`; uploads land in `instance/uploads/<uuid>.xlsx`.
- **Dashboard** with KPIs (IQS overall, fotos avaliadas, inspeções) and 13 Plotly figures
  (IC and IQS per service, photo conformity stacked, team × service, TSS distribution,
  per-service team / TSS conformity).
- **Date-range filter** (`?start=YYYY-MM-DD&end=YYYY-MM-DD`) narrows every chart, KPI, and
  the period header — KPIs are recomputed from raw inspection rows when the filter is
  active (CAPA's pre-aggregated numbers are used when the view is unfiltered).
- **Day/month swap action** (`?swap=1`) heuristically corrects xlsx files where dates
  were entered with day and month inverted at the data source. The swap only flips rows
  whose day matches the file's dominant month and whose stored month differs, so
  already-correct dates aren't broken.
- **Team drilldown** (`/dashboard/<id>/team?name=...`) reports per-service OS-level
  conformity (`conforme + nao_conforme == inspeções`), per-stage failure breakdown
  (with NC vs SF informational subsets), and lists every não-conforme OS with its
  Número OS, TSS, failed stages, and observation text.
- **Download** the dashboard as Markdown, XLSX, PDF, or DOCX (`/download/<id>?fmt=…`).
- **Swapped-date warning** banner fires automatically when the inspection span exceeds
  60 days — the typical signal that the source xlsx has inverted dates.

## Tech stack

| Layer | Library |
|---|---|
| Web | Flask 3, waitress (prod) |
| Data | pandas 2, python-calamine (Rust-backed xlsx parser) |
| Charts | Plotly (CDN-loaded client side) |
| Exporters | openpyxl, reportlab, python-docx, matplotlib |
| Lint / test | ruff, pytest, pytest-cov |

## Architecture

```
upload (POST /upload)
    └─→ file.save(instance/uploads/<uuid>.xlsx)
        └─→ 303 redirect to /dashboard/<uuid>

dashboard (GET /dashboard/<uuid>?start=&end=&swap=)
    └─→ load_workbook(read_only)  +  template.extract_inspections (calamine)
        ├─→ optional _swap_day_month
        ├─→ optional _apply_date_filter
        ├─→ recompute IQS/IC from rows  OR  template.extract_* from CAPA
        └─→ build_* Plotly figures
            └─→ render dashboard.html with Plotly CDN

team detail (GET /dashboard/<uuid>/team?name=…)
    └─→ template.extract_team_detail
        └─→ OS-level conformity, failing-OS drilldown with Número OS

download (GET /download/<uuid>?fmt=md|xlsx|pdf|docx)
    └─→ same pipeline → exporter writes the artefact
```

The xlsx parser (calamine) and the per-`(path, mtime_ns)` cache mean a warm
`/dashboard/<id>` round-trips in ~290 ms on the 1.9 MB Pimentas file (~6.5× faster than the
openpyxl-only baseline). Subsequent hits reuse the parsed inspections DataFrame so the
team-drilldown and downloads pay no parse cost.

### Adding a new polo template

1. Drop a module under `app/core/templates/`.
2. Expose:
   - `matches(sheet_names: Iterable[str]) -> bool`
   - extraction methods (`extract_inspections`, `extract_team_detail`, `extract_*`)
   - `build_*` methods returning `plotly.graph_objects.Figure`
3. Register it in `recognize()` in `app/core/templates/__init__.py`.

Routes don't change — they call into whichever template `recognize()` returns.

## Quickstart

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt -r requirements-dev.txt

# dev server (auto-reload)
.venv/bin/flask --app app --debug run --port 5000

# production server (waitress) — what Render runs
.venv/bin/python run.py
```

Open <http://127.0.0.1:5000>, upload an xlsx, follow the redirect to the dashboard.

## Tests

```bash
.venv/bin/pytest                              # fast suite (default)
.venv/bin/pytest --cov=app --cov-report=term-missing
.venv/bin/ruff check . && .venv/bin/ruff format --check .
```

Coverage gate is 90 %. Unit tests use openpyxl-built fixtures (see
`tests/fixtures/pimentas_minimal.py`); integration tests use the Flask test client with a
`tmp_path`-scoped instance folder (`tests/conftest.py`).

## Deploy

Render free tier via `render.yaml`. Push to `main` triggers a redeploy. First request
after 15 minutes of idle is a ~30 s cold start. `instance/uploads/` is ephemeral on
Render — acceptable for the demo since each dashboard is generated fresh per upload.

```bash
git push                    # triggers Render redeploy
```

## Project layout

```
app/
├── __init__.py             # Flask app factory
├── routes/main.py          # upload, dashboard, team detail, download routes
├── core/
│   ├── templates/
│   │   ├── __init__.py     # recognize()
│   │   └── pimentas.py
│   └── exporters/
│       ├── __init__.py     # render_export(fmt, ...)
│       ├── markdown.py
│       ├── xlsx.py
│       ├── pdf.py
│       └── docx.py
└── templates/
    ├── index.html          # upload page
    ├── dashboard.html      # main dashboard
    ├── dashboard_unknown.html
    └── team_detail.html    # /dashboard/<id>/team
tests/
├── conftest.py
├── fixtures/pimentas_minimal.py
├── unit/                   # extraction, figure builders, cache, swap helpers
└── integration/            # Flask test client hitting the real routes
Model/                      # real polo xlsx samples
instance/uploads/           # runtime upload destination (ephemeral on Render)
```
