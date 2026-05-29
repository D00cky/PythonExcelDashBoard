import io
from pathlib import Path


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
