import os
import sqlite3
import hashlib
import sys

import pytest
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import db
import utils


def setup_memory_db(monkeypatch):
    conn = sqlite3.connect(":memory:")
    monkeypatch.setattr(db, "conn", conn)
    monkeypatch.setattr(db, "cursor", conn.cursor())
    db.init_db()
    db.refresh_location_cache()
    return conn


def test_parse_port_valid():
    assert db._parse_port("1234") == 1234
    assert db._parse_port(" 65535 ") == 65535
    assert db._parse_port("00080") == 80
    # more than 5 digits keeps last 5
    assert db._parse_port("1234567") == 34567
    assert db._parse_port(None) is None


def test_parse_port_invalid():
    for val in ["0", "70000", "abc", "-1"]:
        with pytest.raises(ValueError):
            db._parse_port(val)


def test_password_hash_and_verify():
    h = db._hash_password("secret", _salt=b"\0" * 16)
    assert db._verify_password(h, "secret")
    assert not db._verify_password(h, "wrong")
    legacy = hashlib.sha256(b"x").hexdigest()
    assert db._verify_password(legacy, "x")
    assert not db._verify_password(legacy, "y")


def test_user_creation_and_login(monkeypatch):
    conn = setup_memory_db(monkeypatch)
    db.create_user("bob", "pass", role="seller")
    user = db.get_user("bob")
    assert user["role"] == "seller"
    assert db.check_login("bob", "pass")["username"] == "bob"
    assert db.check_login("bob", "bad") is None


def test_location_cache(monkeypatch):
    conn = setup_memory_db(monkeypatch)
    cur = conn.cursor()
    cur.execute("INSERT INTO locatii (city, county, address) VALUES ('A','B','C')")
    loc_id = cur.lastrowid
    conn.commit()
    db.refresh_location_cache()
    loc = db.get_location_by_id(loc_id)
    assert loc["city"] == "A"

    ts = db._cache_timestamp
    assert not db.maybe_refresh_location_cache(ttl=999999)
    assert db._cache_timestamp == ts
    # Force refresh by setting old timestamp
    db._cache_timestamp -= 1_000_000
    assert db.maybe_refresh_location_cache(ttl=0)


def test_firme_defaults(monkeypatch):
    conn = setup_memory_db(monkeypatch)
    rows = conn.execute("SELECT nume FROM firme ORDER BY id").fetchall()
    assert [r[0] for r in rows] == [
        "Focus Media Outdoor",
        "Excellence Media Production",
        "Michi Media Advertising",
    ]


def test_make_preview(monkeypatch, tmp_path):
    folder = tmp_path / "previews"
    folder.mkdir()
    img_path = folder / "p1.png"
    from PIL import Image
    Image.new("RGB", (100, 50)).save(img_path)
    monkeypatch.setattr(utils, "PREVIEW_FOLDER", str(folder))
    utils.make_preview.cache_clear()
    created = {}

    def fake_photoimage(img):
        created["size"] = img.size
        return "IMG"

    monkeypatch.setattr(utils.ImageTk, "PhotoImage", fake_photoimage)
    result = utils.make_preview("p1", max_w=50, max_h=50)
    assert result == "IMG"
    assert created["size"] == (50, 25)
    assert utils.make_preview("p1", max_w=50, max_h=50) == "IMG"
