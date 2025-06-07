import importlib
import db
import os

class DummyCursor:
    def __init__(self):
        self.commands = []
    def execute(self, sql, params=None):
        self.commands.append(sql)
    def executemany(self, sql, params):
        self.commands.append(sql)
    def fetchone(self):
        return None
    def fetchall(self):
        return []

class DummyConnInner:
    def __init__(self):
        self.cur = DummyCursor()
    def cursor(self):
        return self.cur
    def commit(self):
        pass


def test_mysql_schema(monkeypatch):
    dummy_inner = DummyConnInner()
    conn = db._ConnWrapper(dummy_inner, True)
    monkeypatch.setattr(db, 'conn', conn)
    monkeypatch.setattr(db, 'cursor', conn.cursor())
    db.init_db()
    commands = dummy_inner.cur.commands
    create_table = next(c for c in commands if c.strip().startswith('CREATE TABLE IF NOT EXISTS locatii'))
    assert 'grup VARCHAR(255)' in create_table
    assert any('CREATE INDEX idx_locatii_grup' in c for c in commands)
