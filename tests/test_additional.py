import os
import sqlite3
import pandas as pd
import utils
import db


def test_get_schita_path(tmp_path, monkeypatch):
    schite = tmp_path / "schite"
    schite.mkdir()
    file = schite / "CODE.png"
    file.write_text("data")
    monkeypatch.setattr(utils, "SCHITE_FOLDER", str(schite))
    assert utils.get_schita_path("CODE") == str(file)
    file.unlink()
    assert utils.get_schita_path("CODE") is None


def test_make_preview(monkeypatch, tmp_path):
    previews = tmp_path / "previews"
    previews.mkdir()
    img_path = previews / "img.png"
    from PIL import Image
    Image.new("RGB", (800, 600), color="red").save(img_path)
    monkeypatch.setattr(utils, "PREVIEW_FOLDER", str(previews))
    created = {}
    def fake_photo(image):
        created["size"] = image.size
        return "photo"
    monkeypatch.setattr(utils.ImageTk, "PhotoImage", fake_photo)
    result = utils.make_preview("img", max_w=400, max_h=300)
    assert result == "photo"
    assert created["size"] == (400, 300)
    img_path.unlink()
    assert utils.make_preview("img") is None


def test_pandas_conn_sqlite(monkeypatch):
    con = sqlite3.connect(":memory:")
    wrapper = db._ConnWrapper(con, False)
    monkeypatch.setattr(db, "conn", wrapper)
    assert db.pandas_conn() is con


def test_read_sql_query_mysql(monkeypatch):
    class DummyCursor:
        def __init__(self):
            self.sql = None
            self.params = None
            self.description = [("val",)]
        def execute(self, sql, params=()):
            self.sql = sql
            self.params = params
            return self
        def fetchall(self):
            return [("x",)]
    class DummyConn:
        def __init__(self):
            self.cur = DummyCursor()
        def cursor(self):
            return self.cur
    wrapper = db._ConnWrapper(DummyConn(), True)
    monkeypatch.setattr(db, "conn", wrapper)
    monkeypatch.setattr(db, "sqlalchemy", None)
    df = db.read_sql_query("SELECT val FROM t WHERE id=?", (1,))
    cur = wrapper.cursor()
    assert cur.sql == "SELECT val FROM t WHERE id=%s"
    assert cur.params == (1,)
    assert df.to_dict("list") == {"val": ["x"]}


def test_ensure_index_sqlite(monkeypatch):
    con = sqlite3.connect(":memory:")
    cur = con.cursor()
    cur.execute("CREATE TABLE foo (id INTEGER, data TEXT)")
    wrapper = db._ConnWrapper(con, False)
    monkeypatch.setattr(db, "conn", wrapper)
    monkeypatch.setattr(db, "cursor", wrapper.cursor())
    db.ensure_index("foo", "idx_data", "data")
    indices = [row[1] for row in cur.execute("PRAGMA index_list(foo)")]
    assert "idx_data" in indices


def test_ensure_index_mysql(monkeypatch):
    executed = []
    class DummyCursor:
        def execute(self, sql, params=()):
            executed.append(sql.strip())
            return self
        def fetchone(self):
            if executed[-1].startswith("SHOW INDEX"):
                return None
            if executed[-1].startswith("SHOW FIELDS"):
                return ("data", "TEXT")
    class DummyConn:
        def cursor(self):
            return DummyCursor()
    wrapper = db._ConnWrapper(DummyConn(), True)
    monkeypatch.setattr(db, "conn", wrapper)
    monkeypatch.setattr(db, "cursor", wrapper.cursor())
    db.ensure_index("foo", "idx", "data")
    assert any("CREATE INDEX idx ON foo(data(255))" in sql for sql in executed)
