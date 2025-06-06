import sqlite3
import os
import sys
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import db
import UI.dialogs as dialogs
from tkinter import filedialog, messagebox


def test_export_vendor_report(tmp_path, monkeypatch):
    conn = sqlite3.connect(':memory:')
    monkeypatch.setattr(db, 'conn', conn)
    monkeypatch.setattr(db, 'cursor', conn.cursor())
    monkeypatch.setattr(dialogs, 'conn', conn)
    db.init_db()

    cur = conn.cursor()
    cur.execute("INSERT INTO users (username, password, role) VALUES ('v1','x','seller')")
    cur.execute("INSERT INTO users (username, password, role) VALUES ('v2','x','seller')")
    cur.execute("INSERT INTO clienti (nume) VALUES ('C')")
    cid = cur.lastrowid
    cur.execute("INSERT INTO locatii (city, county, address) VALUES ('A','B','Adr')")
    loc_id = cur.lastrowid
    cur.execute(
        "INSERT INTO rezervari (loc_id, client, client_id, data_start, data_end, suma, created_by)"
        " VALUES (?,?,?,?,?,?,?)",
        (loc_id, 'C', cid, '2025-06-01', '2025-06-30', 300, 'v1'),
    )
    conn.commit()

    out = tmp_path / 'vendors.xlsx'
    monkeypatch.setattr(filedialog, 'asksaveasfilename', lambda **k: str(out))
    monkeypatch.setattr(messagebox, 'showinfo', lambda *a, **k: None)

    sheet_names = []

    class DummySheet:
        def write(self, *a, **k):
            pass
        def write_row(self, *a, **k):
            pass
        def set_column(self, *a, **k):
            pass
        def merge_range(self, *a, **k):
            pass

    class DummyBook:
        def add_format(self, *a, **k):
            return object()
        def add_worksheet(self, name):
            sheet_names.append(name)
            return DummySheet()

    class DummyWriter:
        def __init__(self, *a, **k):
            self.book = DummyBook()
            self.sheets = {}
        def __enter__(self):
            return self
        def __exit__(self, exc_type, exc, tb):
            pass

    monkeypatch.setattr(pd, 'ExcelWriter', DummyWriter)

    def fake_to_excel(self, writer, *args, **kwargs):
        name = kwargs.get('sheet_name')
        sheet_names.append(name)
        writer.sheets[name] = writer.book.add_worksheet(name)
    monkeypatch.setattr(pd.DataFrame, 'to_excel', fake_to_excel)

    dialogs.export_vendor_report()

    assert 'v1' in sheet_names

