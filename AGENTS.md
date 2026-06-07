# Repository Guidelines

## Project Structure & Module Organization

`app/` contains the Flask application. `app/__init__.py` builds the app, `app/routes/main.py` owns upload/dashboard/report/download routes, and `app/core/` contains ingestion, caching, aggregation, geography, exporters, and template-specific parsing. XLSX template adapters live in `app/core/templates/`; add new polo layouts there and register them in `app/core/templates/__init__.py`.

Jinja templates live in `app/templates/`. Test code lives in `tests/`, with fast fixture builders in `tests/fixtures/`, unit tests in `tests/unit/`, and route-level integration tests in `tests/integration/`. Real sample workbooks are under `Model/` and should not be used as primary test fixtures.

## Build, Test, and Development Commands

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt -r requirements-dev.txt
```

Creates the local environment and installs runtime plus test tooling.

```bash
.venv/bin/flask --app app --debug run --port 5000
.venv/bin/python run.py
```

Runs the dev server, or the production-style Waitress server.

```bash
.venv/bin/pytest
.venv/bin/pytest --cov=app --cov-report=term-missing
.venv/bin/ruff check .
.venv/bin/ruff format --check .
```

Runs the default fast suite, coverage, lint, and formatting checks.

## Coding Style & Naming Conventions

Use Python 3.12 style with type hints on public functions. Prefer `pathlib.Path` for paths. Keep user-facing strings in Portuguese and code identifiers in English. Follow Ruff rules from `pyproject.toml`: line length 100, import sorting, and `E/F/I/B/UP/SIM` lint checks. Use concise docstrings for public APIs in `app/core/`.

## Testing Guidelines

Tests use `pytest`, `pytest-flask`, and openpyxl-generated fixtures. Name tests `test_<behavior>` and keep them close to the module they cover. Default test runs exclude `@pytest.mark.slow`; use `.venv/bin/pytest -m slow` only for real workbook or benchmark checks. Coverage gate is 90% for `app`.

## Commit & Pull Request Guidelines

Use conventional commits matching history, for example `feat(cache): add parquet caching layer`, `test(jobs): assert concurrent uploads are isolated by uuid`, or `fix(dashboard): guard failed ingestion state`.

PRs should include a short problem/solution summary, test commands run, and screenshots for dashboard or report-template UI changes. Keep each PR phase-scoped and avoid unrelated refactors.

## Security & Configuration Tips

Do not commit secrets or production uploads. Runtime files belong under `instance/` (`uploads/`, `cache/`, `jobs/`) and are disposable on Render. Networked services such as Redis or email must be configured through environment variables.
