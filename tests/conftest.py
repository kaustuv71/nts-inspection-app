import os, json, base64, io, uuid
from pathlib import Path
from datetime import datetime
from PIL import Image
import pytest

os.environ["APP_PASSWORD"] = "testpass"
os.environ["SECRET_KEY"] = "test-secret-key"

from app import app as _app, db as _db, Inspection, Photo

TEST_PHOTO_DIR = Path(__file__).parent / "test_photos"

@pytest.fixture
def app():
    _app.config.update({
        "TESTING": True,
        "SQLALCHEMY_DATABASE_URI": "sqlite:///:memory:",
        "SERVER_NAME": "localhost",
    })
    with _app.app_context():
        _db.create_all()
        yield _app
        _db.drop_all()

@pytest.fixture
def client(app):
    return app.test_client()

@pytest.fixture
def auth_headers():
    return {"Content-Type": "application/json"}

@pytest.fixture
def logged_in_client(client):
    client.post("/api/login", json={"password": "testpass"})
    return client

@pytest.fixture
def sample_inspection_id(logged_in_client):
    resp = logged_in_client.post("/api/inspection/new", json={"product_type": "dryer_balls"})
    data = resp.get_json()
    return data["id"]

@pytest.fixture
def test_photo_base64():
    img = Image.new("RGB", (100, 100), color="red")
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=80)
    return base64.b64encode(buf.getvalue()).decode()
