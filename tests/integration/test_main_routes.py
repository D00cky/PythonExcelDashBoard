import io
from html.parser import HTMLParser
from pathlib import Path

_EXTERNAL_PREFIXES = ("http://", "https://", "//")


def _external_resource_attrs(html: str) -> list[tuple[str, str]]:
    """Return [(tag, url), ...] for every <script src> / <link href> that loads off-host."""

    class Finder(HTMLParser):
        def __init__(self) -> None:
            super().__init__()
            self.refs: list[tuple[str, str]] = []

        def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
            d = dict(attrs)
            if tag == "script":
                src = d.get("src") or ""
                if src.startswith(_EXTERNAL_PREFIXES):
                    self.refs.append((tag, src))
            elif tag == "link":
                href = d.get("href") or ""
                if href.startswith(_EXTERNAL_PREFIXES):
                    self.refs.append((tag, href))

    finder = Finder()
    finder.feed(html)
    return finder.refs


def test_index_route_serves_upload_form(client):
    response = client.get("/")

    assert response.status_code == 200
    assert b"<form" in response.data
    assert b'type="file"' in response.data


def test_upload_rejects_missing_file_with_400(client):
    response = client.post("/upload")

    assert response.status_code == 400


def test_upload_rejects_non_xlsx_extension_with_400(client):
    response = client.post(
        "/upload",
        data={"file": (io.BytesIO(b"not an xlsx"), "notes.txt")},
        content_type="multipart/form-data",
    )

    assert response.status_code == 400


def test_upload_persists_file_and_303_redirects_to_dashboard(client, app):
    payload = b"xlsx-bytes-stand-in"

    response = client.post(
        "/upload",
        data={"file": (io.BytesIO(payload), "report.xlsx")},
        content_type="multipart/form-data",
    )

    assert response.status_code == 303
    assert response.location.startswith("/dashboard/")

    upload_id = response.location.removeprefix("/dashboard/")
    saved = Path(app.instance_path) / "uploads" / f"{upload_id}.xlsx"
    assert saved.read_bytes() == payload


def test_dashboard_renders_sabesp_kpis_and_two_figures(client, tmp_path):
    from tests.fixtures.sabesp_minimal import make_minimal_sabesp

    payload = make_minimal_sabesp(tmp_path).read_bytes()
    upload_response = client.post(
        "/upload",
        data={"file": (io.BytesIO(payload), "report.xlsx")},
        content_type="multipart/form-data",
    )

    response = client.get(upload_response.location)

    assert response.status_code == 200
    body = response.data.decode("utf-8")
    assert "01/03/2026 à 31/03/2026" in body
    assert "66.1%" in body
    assert "Polo Pimentas" in body
    assert body.count('class="plotly-graph-div"') == 3
    assert "Listas de Verificação" in body
    assert "Fotos Avaliadas" in body


def test_dashboard_returns_404_for_unknown_id(client):
    response = client.get("/dashboard/" + "0" * 32)

    assert response.status_code == 404


def test_download_html_returns_self_contained_attachment(client, tmp_path):
    from tests.fixtures.sabesp_minimal import make_minimal_sabesp

    payload = make_minimal_sabesp(tmp_path).read_bytes()
    upload = client.post(
        "/upload",
        data={"file": (io.BytesIO(payload), "report.xlsx")},
        content_type="multipart/form-data",
    )
    upload_id = upload.location.removeprefix("/dashboard/")

    response = client.get(f"/download/{upload_id}?fmt=html")

    assert response.status_code == 200
    assert "attachment" in response.headers["Content-Disposition"]
    assert f"dashboard-{upload_id}.html" in response.headers["Content-Disposition"]
    body = response.data.decode("utf-8")
    refs = _external_resource_attrs(body)
    assert refs == [], f"downloaded HTML must not load external resources: {refs}"
    assert body.count('class="plotly-graph-div"') == 3
    assert "01/03/2026 à 31/03/2026" in body


def test_download_returns_400_for_unsupported_format(client, tmp_path):
    from tests.fixtures.sabesp_minimal import make_minimal_sabesp

    payload = make_minimal_sabesp(tmp_path).read_bytes()
    upload = client.post(
        "/upload",
        data={"file": (io.BytesIO(payload), "report.xlsx")},
        content_type="multipart/form-data",
    )
    upload_id = upload.location.removeprefix("/dashboard/")

    response = client.get(f"/download/{upload_id}?fmt=pdf")

    assert response.status_code == 400
