import os
import sqlite3
import importlib

import db


def setup_module(module):
    # Use an in-memory SQLite database for testing
    test_conn = db._ConnWrapper(sqlite3.connect(":memory:"), False)
    db.conn = test_conn
    db.cursor = test_conn.cursor()
    db.init_db()
    db.refresh_location_cache()


def test_parse_port():
    assert db._parse_port("3306") == 3306
    assert db._parse_port("") is None
    assert db._parse_port(None) is None
    assert db._parse_port(" 123 ") == 123

    try:
        db._parse_port("abc")
    except ValueError:
        pass
    else:
        assert False, "Expected ValueError"


def test_hash_and_verify_password():
    pw = "secret"
    hashed = db._hash_password(pw)
    assert hashed != pw
    assert db._verify_password(hashed, pw)
    assert not db._verify_password(hashed, "other")


def test_create_and_get_user():
    username = "testuser"
    db.create_user(username, "pw123", role="seller")
    user = db.get_user(username)
    assert user["username"] == username
    assert db.check_login(username, "pw123")
    assert not db.check_login(username, "wrong")


def test_create_manager_user():
    username = "mgr"
    db.create_user(username, "pw123", role="manager")
    user = db.get_user(username)
    assert user["role"] == "manager"


def test_reconnect_on_lost_connection(monkeypatch):
    calls = {}

    class DummyCursor:
        def __init__(self, fail=False):
            self.fail = fail
            self.sql = None
            self.params = None
            self.call_count = 0

        def execute(self, sql, params=()):
            self.call_count += 1
            if self.fail and self.call_count == 1:
                raise db.mysql.connector.errors.OperationalError(msg="Lost", errno=2013)
            self.sql = sql
            self.params = params
            return self

        def executemany(self, sql, params):
            return self.execute(sql, params)

        def fetchall(self):
            return []

    class DummyConn:
        def __init__(self, cur):
            self.cur = cur

        def cursor(self):
            return self.cur

        def commit(self):
            pass

    failing_cursor = DummyCursor(fail=True)
    ok_cursor = DummyCursor()

    def dummy_create_connection():
        if not calls:
            calls["initial"] = True
            return db._ConnWrapper(DummyConn(failing_cursor), True)
        calls["reconnect"] = True
        return db._ConnWrapper(DummyConn(ok_cursor), True)

    monkeypatch.setattr(db, "_create_connection", dummy_create_connection)
    db.reconnect()
    db.cursor.execute("SELECT 1")
    assert calls.get("reconnect")
    assert ok_cursor.sql == "SELECT 1"


def test_needs_reconnect_additional_codes():
    codes = [2002, 2003, 2055]
    for code in codes:
        exc = db.mysql.connector.errors.OperationalError(msg="Lost", errno=code)
        assert db._needs_reconnect(exc)


def test_init_rezervari_mysql_add_columns():
    executed = []

    class DummyCursor:
        def __init__(self):
            self.last_sql = ""

        def execute(self, sql, params=()):
            executed.append(sql.strip())
            self.last_sql = sql.strip()
            return self

        def executemany(self, sql, params):
            return self.execute(sql, params)

        def fetchall(self):
            if self.last_sql.startswith("SHOW COLUMNS FROM rezervari"):
                return [("id",), ("loc_id",), ("client",)]
            return []

    class DummyConn:
        def __init__(self):
            self.cur = DummyCursor()

        def cursor(self):
            return self.cur

        def commit(self):
            executed.append("COMMIT")

    old_conn, old_cursor = db.conn, db.cursor
    db.conn = db._ConnWrapper(DummyConn(), True)
    db.cursor = db.conn.cursor()
    db.init_rezervari_table()
    db.conn = old_conn
    db.cursor = old_cursor

    assert any("ALTER TABLE rezervari ADD COLUMN firma_id" in sql for sql in executed)
    assert any("ALTER TABLE rezervari ADD COLUMN decor_cost" in sql for sql in executed)
    assert any("ALTER TABLE rezervari ADD COLUMN prod_cost" in sql for sql in executed)
    assert any("ALTER TABLE rezervari ADD COLUMN created_on" in sql for sql in executed)


def test_client_contacts_basic():
    old_conn, old_cursor = db.conn, db.cursor
    test_conn = db._ConnWrapper(sqlite3.connect(":memory:"), False)
    db.conn = test_conn
    db.cursor = test_conn.cursor()
    db.init_clienti_table()
    db.init_client_contacts_table()
    cur = db.conn.cursor()
    cur.execute("INSERT INTO clienti (nume) VALUES (?)", ("Test",))
    cid = cur.execute("SELECT last_insert_rowid()").fetchone()[0]
    db.add_client_contact(cid, "Pers", "Role", "e@test", "000")
    contacts = db.get_client_contacts(cid)
    assert len(contacts) == 1
    assert contacts[0]["nume"] == "Pers"
    db.conn = old_conn
    db.cursor = old_cursor


def test_update_and_delete_contact():
    old_conn, old_cursor = db.conn, db.cursor
    test_conn = db._ConnWrapper(sqlite3.connect(":memory:"), False)
    db.conn = test_conn
    db.cursor = test_conn.cursor()
    db.init_clienti_table()
    db.init_client_contacts_table()
    cur = db.conn.cursor()
    cur.execute("INSERT INTO clienti (nume) VALUES (?)", ("Test",))
    cid = cur.execute("SELECT last_insert_rowid()").fetchone()[0]
    db.add_client_contact(cid, "Pers", "Role", "e@test", "000")
    contact = db.get_client_contacts(cid)[0]
    db.update_client_contact(contact["id"], "Pers2", "R2", "m@test", "111")
    updated = db.get_client_contacts(cid)[0]
    assert updated["nume"] == "Pers2"
    db.delete_client_contact(contact["id"])
    assert db.get_client_contacts(cid) == []
    db.conn = old_conn
    db.cursor = old_cursor

