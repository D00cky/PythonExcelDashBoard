# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

PythonExcelDashBoard — a Flask web app that turns an uploaded weekly polo audit-report
`.xlsx` into an interactive dashboard (Plotly), with a self-contained HTML download for
offline viewing.

Currently recognises the "Polo Pimentas" template (cover page `CAPA` + master sheet
`DADOS - PIMENTAS` + per-service sheets). Architecture supports adding more templates without
changing route code; see `app/core/templates/` and `app/core/templates/pimentas.py`.

## Workflow

XP / TDD / pair programming. **Tight red → green → refactor loops**, one commit per cycle.

- Tests live under `tests/`, mirroring `app/`. Unit tests use openpyxl-built fixtures
  (see `tests/fixtures/pimentas_minimal.py`), not the real `Model/` xlsx — fast and
  intent-revealing.
- Integration tests use the Flask test client + a `tmp_path`-scoped instance folder
  (see `tests/conftest.py`).
- Coverage gate: ≥ 90% per module (`pyproject.toml`).
- Lint: `ruff` (line 100, py312, E/F/I/B/UP/SIM).

## Commands

### Setup

```
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt -r requirements-dev.txt
```

### Test loop

```
.venv/bin/pytest                              # fast tests (default)
.venv/bin/pytest -m slow                      # @pytest.mark.slow tests
.venv/bin/pytest --cov=app --cov-report=term-missing
.venv/bin/ruff check . && .venv/bin/ruff format --check .
.venv/bin/ruff format .                       # apply formatting
```

If pytest spends >30 s on an assertion failure, the message is dumping a megabyte-class
HTML body — use `--tb=line` or tighten the assertion (see how `test_download_html_*` uses
`html.parser` instead of substring `in body`).

### Dev server

```
.venv/bin/flask --app app --debug run --port 5000
```

### Production server (same as Render runs)

```
.venv/bin/python run.py     # waitress, binds 0.0.0.0:$PORT or 8000
```

## Deploy — Render free tier

Already live. Push to `main`; Render rebuilds via `render.yaml` blueprint. First request
after 15 min idle cold-starts ~30 s.

```
git push                    # triggers Render redeploy
```

`instance/uploads/` is ephemeral on Render — uploaded files vanish on dyno restart.
Acceptable for the demo (TTL is short anyway; each upload generates a fresh dashboard).

## Architecture map

```
upload (POST /upload)
    → file.save(instance/uploads/<uuid>.xlsx)
    → 303 redirect to /dashboard/<uuid>

dashboard (GET /dashboard/<uuid>)
    → load_workbook(...) → recognize(sheet_names) → Template instance
    → template.extract_*  (periodo, IC, IQS, totals)
    → template.build_*    (Plotly Figures)
    → render dashboard.html with Plotly CDN (live page)

download (GET /download/<uuid>?fmt=html)
    → same pipeline, but Plotly inlined → self-contained HTML attachment
```

Adding a new template:
1. Drop a module under `app/core/templates/`.
2. Expose `matches(sheet_names: Iterable[str]) -> bool` and a `build_*` method per figure.
3. Add it to `recognize()` in `app/core/templates/__init__.py`.

## Sandbox note

This project runs inside [ai-jail](https://github.com/akitaonrails/ai-jail). `git push`
to GitHub requires SSH access that the sandbox blocks — run pushes from a shell outside
the jail. Regenerate `.ai-jail` with `ai-jail --clean --init` if needed.
