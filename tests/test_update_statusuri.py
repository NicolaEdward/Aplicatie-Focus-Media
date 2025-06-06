import sqlite3
import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import db


def test_update_statusuri(monkeypatch):
    conn = sqlite3.connect(':memory:')
    monkeypatch.setattr(db, 'conn', conn)
    monkeypatch.setattr(db, 'cursor', conn.cursor())
    db.init_db()

    cur = conn.cursor()
    cur.execute(
        "INSERT INTO clienti (nume) VALUES (?)",
        ('X',)
    )
    client_id = cur.lastrowid
    cur.execute(
        "INSERT INTO locatii (city, county, address) VALUES (?,?,?)",
        ('A', 'B', 'C')
    )
    loc_id = cur.lastrowid

    today = datetime.date.today()
    cur.execute(
        "INSERT INTO rezervari (loc_id, client, client_id, data_start, data_end, suma)"
        " VALUES (?, ?, ?, ?, ?, 100)",
        (loc_id, 'X', client_id, today.isoformat(), (today + datetime.timedelta(days=1)).isoformat())
    )
    conn.commit()

    db.update_statusuri_din_rezervari()

    status, cid = cur.execute(
        "SELECT status, client_id FROM locatii WHERE id=?", (loc_id,)
    ).fetchone()
    assert status == 'Închiriat'
    assert cid == client_id
