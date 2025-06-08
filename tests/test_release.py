import sqlite3
import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import db
import UI.dialogs as dialogs
from tkinter import messagebox


class DummyTop:
    def __init__(self, *a, **k):
        pass
    def title(self, *a, **k):
        pass
    def destroy(self):
        pass


class DummyListbox:
    def __init__(self, *a, **k):
        self.items = []
    def insert(self, *a):
        self.items.append(a[1])
    def grid(self, *a, **k):
        pass
    def curselection(self):
        return (0,)


class DummyButton:
    def __init__(self, *a, **k):
        cmd = k.get('command')
        if cmd:
            cmd()
    def grid(self, *a, **k):
        pass


def setup_db(monkeypatch):
    conn = sqlite3.connect(':memory:')
    monkeypatch.setattr(db, 'conn', conn)
    monkeypatch.setattr(db, 'cursor', conn.cursor())
    monkeypatch.setattr(dialogs, 'conn', conn)
    db.init_db()
    return conn


def base_setup(monkeypatch):
    conn = setup_db(monkeypatch)
    cur = conn.cursor()
    cur.execute("INSERT INTO clienti (nume) VALUES ('C')")
    cid = cur.lastrowid
    cur.execute("INSERT INTO locatii (city, county, address) VALUES ('A','B','C')")
    loc_id = cur.lastrowid
    start = datetime.date.today() - datetime.timedelta(days=5)
    end = datetime.date.today() + datetime.timedelta(days=5)
    cur.execute(
        "INSERT INTO rezervari (loc_id, client, client_id, data_start, data_end, suma)"
        " VALUES (?,?,?,?,?,100)",
        (loc_id, 'C', cid, start.isoformat(), end.isoformat()),
    )
    res_id = cur.lastrowid
    conn.commit()
    return conn, loc_id, res_id


def patch_widgets(monkeypatch):
    monkeypatch.setattr(dialogs.tk, 'Toplevel', DummyTop)
    monkeypatch.setattr(dialogs.tk, 'Listbox', DummyListbox)
    monkeypatch.setattr(dialogs.ttk, 'Button', DummyButton)
    monkeypatch.setattr(messagebox, 'showwarning', lambda *a, **k: None)
    monkeypatch.setattr(messagebox, 'showinfo', lambda *a, **k: None)


def test_release_delete(monkeypatch):
    conn, loc_id, res_id = base_setup(monkeypatch)
    patch_widgets(monkeypatch)
    monkeypatch.setattr(messagebox, 'askyesno', lambda *a, **k: True)
    calls = {'count': 0}
    def load_cb():
        calls['count'] += 1
    user = {'role': 'admin', 'username': 'x'}
    dialogs.open_release_window(object(), loc_id, load_cb, user)
    cur = conn.cursor()
    row = cur.execute('SELECT 1 FROM rezervari WHERE id=?', (res_id,)).fetchone()
    assert row is None
    db.update_statusuri_din_rezervari()
    status = cur.execute('SELECT status FROM locatii WHERE id=?', (loc_id,)).fetchone()[0]
    assert status == 'Disponibil'
    assert calls['count'] == 1


def test_release_opens_edit(monkeypatch):
    conn, loc_id, res_id = base_setup(monkeypatch)
    patch_widgets(monkeypatch)
    monkeypatch.setattr(messagebox, 'askyesno', lambda *a, **k: False)
    recorded = {}
    def fake_edit(root, rid, load_cb):
        recorded['rid'] = rid
    monkeypatch.setattr(dialogs, 'open_edit_rent_window', fake_edit)
    user = {'role': 'admin', 'username': 'x'}
    dialogs.open_release_window(object(), loc_id, lambda: None, user)
    assert recorded.get('rid') == res_id
