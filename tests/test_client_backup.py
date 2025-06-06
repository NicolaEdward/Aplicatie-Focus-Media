import sqlite3
import os
import sys
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import db
import UI.dialogs as dialogs
from tkinter import filedialog, messagebox


def test_export_client_backup(tmp_path, monkeypatch):
    conn = sqlite3.connect(':memory:')
    monkeypatch.setattr(db, 'conn', conn)
    monkeypatch.setattr(db, 'cursor', conn.cursor())
    monkeypatch.setattr(dialogs, 'conn', conn)
    db.init_db()

    cur = conn.cursor()
    cur.execute("INSERT INTO clienti (nume) VALUES ('Test')")
    cid = cur.lastrowid
    cur.execute("INSERT INTO locatii (city, county, address) VALUES ('A','B','C')")
    loc_id = cur.lastrowid
    cur.execute(
        "INSERT INTO rezervari (loc_id, client, client_id, data_start, data_end, suma)"
        " VALUES (?,?,?,?,?,?)",
        (loc_id, 'Test', cid, '2025-06-15', '2025-07-15', 1000),
    )
    conn.commit()

    out = tmp_path / 'out.xlsx'
    monkeypatch.setattr(filedialog, 'asksaveasfilename', lambda **k: str(out))
    monkeypatch.setattr(messagebox, 'showinfo', lambda *a, **k: None)

    captured = {}
    def fake_to_excel(self, path, index=False, engine=None):
        captured['df'] = self.copy()
    monkeypatch.setattr(pd.DataFrame, 'to_excel', fake_to_excel)

    dialogs.export_client_backup(6, 2025, cid)

    df = captured['df']
    assert round(df['Amount'].iloc[0], 2) == round(1000 * 16 / 30, 2)

