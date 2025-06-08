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

