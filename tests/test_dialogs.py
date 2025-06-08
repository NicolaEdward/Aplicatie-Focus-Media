import pandas as pd
import UI.dialogs as dialogs
import db


def test_export_sales_report_year_no_reservations(monkeypatch, tmp_path):
    # Avoid database and UI interactions
    monkeypatch.setattr(db, "update_statusuri_din_rezervari", lambda: None)
    monkeypatch.setattr(dialogs, "choose_report_year", lambda parent=None: 2030)

    def dummy_read_sql_query(sql, params=None, **kwargs):
        if "FROM locatii" in sql:
            return pd.DataFrame({
                "id": [1],
                "city": ["c"],
                "county": ["co"],
                "address": ["a"],
                "type": ["t"],
                "size": ["s"],
                "sqm": [10],
                "illumination": ["i"],
                "ratecard": [100],
                "pret_vanzare": [1000],
                "grup": ["g"],
                "status": ["free"],
                "is_mobile": [0],
                "parent_id": [None],
            })
        return pd.DataFrame(
            columns=[
                "id",
                "grup",
                "city",
                "county",
                "address",
                "type",
                "size",
                "sqm",
                "illumination",
                "ratecard",
                "pret_vanzare",
                "is_mobile",
                "parent_id",
                "client",
                "data_start",
                "data_end",
                "suma",
            ]
        )

    monkeypatch.setattr(db, "read_sql_query", dummy_read_sql_query)

    saved = {}
    monkeypatch.setattr(dialogs.filedialog, "asksaveasfilename", lambda **k: str(tmp_path / "out.xlsx"))
    monkeypatch.setattr(dialogs.messagebox, "showinfo", lambda *a, **k: saved.setdefault("info", a))

    dialogs.export_sales_report()
    assert (tmp_path / "out.xlsx").exists()
    assert saved.get("info")


def test_export_client_backup(monkeypatch, tmp_path):
    row = (
        "Cli",
        "RO1",
        "AddrC",
        "Firma",
        "FCUI",
        "AddrF",
        "Camp",
        "City",
        "Addr",
        "CODE",
        "F1",
        "Billboard",
        "5x3",
        15,
        "2023-05-01",
        "2023-05-31",
        100.0,
        10.0,
        1,
        None,
        None,
    )

    class DummyCursor:
        def execute(self, sql, params=()):
            return self

        def fetchall(self):
            return [row]

    class DummyConn:
        def cursor(self):
            return DummyCursor()

    dummy = DummyConn()
    monkeypatch.setattr(dialogs, "conn", dummy)

    monkeypatch.setattr(dialogs.filedialog, "askdirectory", lambda **k: str(tmp_path))
    saved = {}
    monkeypatch.setattr(dialogs.messagebox, "showinfo", lambda *a, **k: saved.setdefault("info", a))

    dialogs.export_client_backup(5, 2023, 1)
    expected = tmp_path / "BKP Firma x Cli - Camp - May.xlsx"
    assert expected.exists()
    assert saved.get("info")


def test_export_all_backups(monkeypatch, tmp_path):
    rows = [
        (
            "Cli",
            "RO1",
            "AddrC",
            "Firma",
            "FCUI",
            "AddrF",
            "Camp1",
            "City",
            "Addr",
            "CODE",
            "F1",
            "Billboard",
            "5x3",
            15,
            "2023-05-01",
            "2023-05-31",
            100.0,
            10.0,
            1,
            None,
            None,
        ),
        (
            "Cli",
            "RO1",
            "AddrC",
            "Firma2",
            "FCUI",
            "AddrF",
            "Camp2",
            "City",
            "Addr",
            "CODE",
            "F1",
            "Billboard",
            "5x3",
            15,
            "2023-05-01",
            "2023-05-31",
            100.0,
            10.0,
            1,
            None,
            None,
        ),
    ]

    class DummyCursor:
        def execute(self, sql, params=()):
            return self

        def fetchall(self):
            return rows

    class DummyConn:
        def cursor(self):
            return DummyCursor()

    dummy = DummyConn()
    monkeypatch.setattr(dialogs, "conn", dummy)

    monkeypatch.setattr(dialogs.filedialog, "askdirectory", lambda **k: str(tmp_path))
    saved = {}
    monkeypatch.setattr(dialogs.messagebox, "showinfo", lambda *a, **k: saved.setdefault("info", a))

    dialogs.export_all_backups(5, 2023)
    assert (tmp_path / "Firma" / "BKP Firma x Cli - Camp1 - May.xlsx").exists()
    assert (tmp_path / "Firma2" / "BKP Firma2 x Cli - Camp2 - May.xlsx").exists()
    assert saved.get("info")

