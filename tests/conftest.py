import pytest

from app import create_app


@pytest.fixture
def app(tmp_path):
    return create_app({"TESTING": True, "INSTANCE_PATH": str(tmp_path)})


@pytest.fixture
def client(app):
    return app.test_client()
