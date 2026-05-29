def test_index_route_serves_upload_form(client):
    response = client.get("/")

    assert response.status_code == 200
    assert b"<form" in response.data
    assert b'type="file"' in response.data


def test_upload_rejects_missing_file_with_400(client):
    response = client.post("/upload")

    assert response.status_code == 400
