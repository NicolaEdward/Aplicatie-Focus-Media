import sqlite3
import os
import sys
import pandas as pd
import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import db
import UI.dialogs as dialogs
from tkinter import filedialog, messagebox


def test_export_sales_report(tmp_path, monkeypatch):
    conn = sqlite3.connect(':memory:')
    monkeypatch.setattr(db, 'conn', conn)
    monkeypatch.setattr(db, 'cursor', conn.cursor())
    monkeypatch.setattr(dialogs, 'conn', conn)
    monkeypatch.setattr(dialogs, 'update_statusuri_din_rezervari', db.update_statusuri_din_rezervari)
    db.init_db()

    cur = conn.cursor()
    cur.execute("INSERT INTO clienti (nume) VALUES ('Test')")
    cid = cur.lastrowid
    cur.execute("INSERT INTO locatii (city, county, address, pret_vanzare) VALUES ('A','B','C',100)")
    loc_id = cur.lastrowid
    start = datetime.date(2025, 6, 1)
    end = datetime.date(2025, 6, 10)
    cur.execute(
        "INSERT INTO rezervari (loc_id, client, client_id, data_start, data_end, suma)"
        " VALUES (?,?,?,?,?,?)",
        (loc_id, 'Test', cid, start.isoformat(), end.isoformat(), 1000)
    )
    conn.commit()

    out = tmp_path / 'report.xlsx'
    monkeypatch.setattr(filedialog, 'asksaveasfilename', lambda **k: str(out))
    monkeypatch.setattr(messagebox, 'showinfo', lambda *a, **k: None)

    captured = []

    class DummyBook:
        def add_format(self, *a, **k):
            return object()

    class DummySheet:
        def set_column(self, *a, **k):
            pass
        def write(self, *a, **k):
            pass
        def set_row(self, *a, **k):
            pass

    class DummyWriter:
        def __init__(self, *a, **k):
            self.book = DummyBook()
            self.sheets = {}
        def __enter__(self):
            return self
        def __exit__(self, exc_type, exc, tb):
            pass

    monkeypatch.setattr(pd, 'ExcelWriter', DummyWriter)

    def fake_to_excel(self, writer, sheet_name=None, *args, **kwargs):
        writer.sheets[sheet_name] = DummySheet()
        captured.append((sheet_name, self.copy()))
    monkeypatch.setattr(pd.DataFrame, 'to_excel', fake_to_excel)

    dialogs.export_sales_report()

    sheet_names = [s for s, _ in captured]
    assert 'Total' in sheet_names
    assert 'June' in sheet_names
