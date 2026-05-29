def test_index_route_serves_upload_form(client):
    response = client.get("/")

    assert response.status_code == 200
    assert b"<form" in response.data
    assert b'type="file"' in response.data
