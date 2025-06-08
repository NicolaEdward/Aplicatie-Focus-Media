# UI/dialogs.py
import datetime
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import pandas as pd
from UI.date_picker import DatePicker
import xlsxwriter

from utils import PREVIEW_FOLDER, make_preview
from db import (
    conn,
    update_statusuri_din_rezervari,
    create_user,
    get_location_by_id,
    add_rental_with_decor,
)

def open_detail_window(tree, event):
    """Display extended information about the selected location."""
    rowid = tree.identify_row(event.y)
    if not rowid:
        return
    loc_id = int(rowid)     # folosim iid-ul (id-ul real din DB), nu valoarea din coloane
    # extrage detaliile locației și le afișează într-o fereastră dedicată

    data = get_location_by_id(loc_id)
    if not data:
        messagebox.showerror("Eroare", "Datele locației nu au fost găsite.")
        return

    city = data.get("city")
    county = data.get("county")
    address = data.get("address")
    type_ = data.get("type")
    gps = data.get("gps")
    code = data.get("code")
    size_ = data.get("size")
    photo_link = data.get("photo_link")
    sqm = data.get("sqm")
    illumination = data.get("illumination")
    ratecard = data.get("ratecard")
    decoration_cost = data.get("decoration_cost")
    pret_vanzare = data.get("pret_vanzare")
    pret_flotant = data.get("pret_flotant")
    observatii = data.get("observatii")
    status = data.get("status")
    client = data.get("client")
    ds = data.get("data_start")
    de = data.get("data_end")
    cur = conn.cursor()

    # fereastra
    win = tk.Toplevel(tree.master)
    win.title(f"Detalii locație #{loc_id}")

    frm = ttk.Frame(win, padding=10)
    frm.grid(sticky="nsew")
    win.columnconfigure(0, weight=1)
    win.rowconfigure(0, weight=1)

    # 1) Preview imagine
    img = make_preview(code)
    if img:
        lbl_img = ttk.Label(frm, image=img)
        lbl_img.image = img
    else:
        lbl_img = ttk.Label(frm, text="Fără preview")
    lbl_img.grid(row=0, column=0, columnspan=2, pady=(0,15))

    # helper pentru rânduri
    def add_field(r, label, widget):
        ttk.Label(frm, text=label+":", font=("Segoe UI", 9, "bold"))\
           .grid(row=r, column=0, sticky="e", padx=5, pady=2)
        widget.grid(row=r, column=1, sticky="w", padx=5, pady=2)

    r = 1
    add_field(r, "City",       ttk.Label(frm, text=city)); r += 1
    add_field(r, "County",     ttk.Label(frm, text=county)); r += 1
    add_field(r, "Address",    ttk.Label(frm, text=address)); r += 1
    add_field(r, "Type",       ttk.Label(frm, text=type_)); r += 1

    # GPS ca hyperlink către Google Maps, cu text "Google Maps"
    if gps:
        url_maps = f"https://www.google.com/maps/search/?api=1&query={gps}"
        lbl_gps = ttk.Label(frm, text="Google Maps", cursor="hand2", foreground="blue")
        lbl_gps.bind("<Button-1>", lambda e: webbrowser.open(url_maps))
    else:
        lbl_gps = ttk.Label(frm, text="-")
    add_field(r, "GPS", lbl_gps); r += 1

    add_field(r, "Code",      ttk.Label(frm, text=code)); r += 1
    add_field(r, "Size",      ttk.Label(frm, text=size_)); r += 1

    # Photo Link ca hyperlink
    if photo_link:
        href = photo_link.strip()
        if not href.lower().startswith(("http://","https://")):
            href = "https://" + href
        lbl_photo = ttk.Label(frm, text="Vezi poza", cursor="hand2", foreground="blue")
        lbl_photo.bind("<Button-1>", lambda e: webbrowser.open(href))
    else:
        lbl_photo = ttk.Label(frm, text="-")
    add_field(r, "Photo Link", lbl_photo); r += 1

    add_field(r, "SQM",       ttk.Label(frm, text=str(sqm))); r += 1
    add_field(r, "Illumination",
                         ttk.Label(frm, text=illumination)); r += 1
    add_field(r, "RateCard",  ttk.Label(frm, text=str(ratecard))); r += 1
    add_field(r, "Preț de vânzare",
                         ttk.Label(frm, text=str(pret_vanzare))); r += 1
    add_field(r, "Preț Flotant", ttk.Label(frm, text=str(pret_flotant))); r += 1
    add_field(r, "Preț de decorare",
                         ttk.Label(frm, text=str(decoration_cost))); r += 1

    add_field(r, "Observații",
                         ttk.Label(frm, text=observatii or "-")); r += 1
    add_field(r, "Status",    ttk.Label(frm, text=status)); r += 1

    # Client și perioadă doar dacă există
    if client:
        add_field(r, "Client",   ttk.Label(frm, text=client)); r += 1
        period = f"{ds} → {de}" if ds and de else "-"
        add_field(r, "Perioadă", ttk.Label(frm, text=period)); r += 1
        fee_row = None
        if ds and de:
            fee_row = cur.execute(
                "SELECT suma FROM rezervari WHERE loc_id=? AND data_start=? AND data_end=? ORDER BY data_start DESC LIMIT 1",
                (loc_id, ds, de)
            ).fetchone()
        if fee_row:
            add_field(r, "Sumă închiriere", ttk.Label(frm, text=str(fee_row[0]))); r += 1

    # face fereastra redimensionabilă
    for i in range(r):
        win.rowconfigure(i, weight=0)
    win.columnconfigure(1, weight=1)

def open_add_window(root, refresh_cb):
    win = tk.Toplevel(root)
    win.title("Adaugă locație")
    labels = [
        "City", "County", "Address", "Type", "GPS", "Code",
        "Size", "Photo Link", "SQM", "Illumination", "RateCard",
        "Preț Vânzare", "Preț Flotant", "Decoration cost",
        "Observații", "Grup", "Față"
    ]
    entries = {}
    for i, lbl in enumerate(labels):
        ttk.Label(win, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=5, pady=2)
        if lbl == "Față":
            e = ttk.Combobox(win, values=["Fața A", "Fața B"], state="readonly")
            e.current(0)
        else:
            e = ttk.Entry(win, width=40)
        e.grid(row=i, column=1, padx=5, pady=2)
        entries[lbl] = e

    def save():
        vals = {k: e.get().strip() for k, e in entries.items()}
        if not vals["City"]:
            messagebox.showwarning("Lipsește date", "Completează City.")
            return
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO locatii (
                city, county, address, type, gps, code, size,
                photo_link, sqm, illumination, ratecard,
                pret_vanzare, pret_flotant, decoration_cost,
                observatii, grup, face
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, [
            vals["City"], vals["County"], vals["Address"], vals["Type"],
            vals["GPS"], vals["Code"], vals["Size"],
            vals["Photo Link"], vals["SQM"], vals["Illumination"],
            vals["RateCard"],
            vals["Preț Vânzare"] or None,
            vals["Preț Flotant"] or None,
            vals["Decoration cost"] or None,
            vals["Observații"], vals["Grup"], vals["Față"]
        ])
        conn.commit()
        refresh_cb()
        win.destroy()

    ttk.Button(win, text="Salvează", command=save)\
        .grid(row=len(labels), column=0, columnspan=2, pady=10)


def open_edit_window(root, loc_id, load_cb, refresh_groups_cb):
    cur = conn.cursor()
    cur.execute("""
        SELECT
            city, county, address, type, gps, code, size,
            photo_link, sqm, illumination, ratecard,
            pret_vanzare, pret_flotant, decoration_cost,
            observatii, grup, face
        FROM locatii WHERE id=?
    """, (loc_id,))
    row = cur.fetchone()
    if not row:
        return

    labels = [
        "City", "County", "Address", "Type", "GPS", "Code",
        "Size", "Photo Link", "SQM", "Illumination", "RateCard",
        "Preț Vânzare", "Preț Flotant", "Decoration cost",
        "Observații", "Grup", "Față"
    ]

    win = tk.Toplevel(root)
    win.title(f"Editează locație #{loc_id}")
    entries = {}

    # Asigură-te că row are aceeași lungime cu labels
    row = list(row)
    while len(row) < len(labels):
        row.append("")

    for i, lbl in enumerate(labels):
        ttk.Label(win, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=5, pady=2)
        if lbl == "Față":
            e = ttk.Combobox(win, values=["Fața A", "Fața B"], state="readonly")
            e.set(row[i] if row[i] else "Fața A")
        else:
            e = ttk.Entry(win, width=40)
            e.insert(0, str(row[i]) if row[i] is not None else "")
        e.grid(row=i, column=1, padx=5, pady=2)
        entries[lbl] = e

    def save_edit():
        vals = {k: e.get().strip() for k, e in entries.items()}
        if not vals["City"]:
            messagebox.showwarning("Lipsește date", "Completează City.")
            return
        cur = conn.cursor()
        cur.execute("""
            UPDATE locatii SET
                city=?, county=?, address=?, type=?, gps=?, code=?, size=?,
                photo_link=?, sqm=?, illumination=?, ratecard=?,
                pret_vanzare=?, pret_flotant=?, decoration_cost=?,
                observatii=?, grup=?, face=?
            WHERE id=?
        """, [
            vals["City"], vals["County"], vals["Address"], vals["Type"],
            vals["GPS"], vals["Code"], vals["Size"],
            vals["Photo Link"], vals["SQM"], vals["Illumination"],
            vals["RateCard"],
            vals["Preț Vânzare"] or None,
            vals["Preț Flotant"] or None,
            vals["Decoration cost"] or None,
            vals["Observații"], vals["Grup"], vals["Față"],
            loc_id
        ])
        conn.commit()
        refresh_groups_cb()
        load_cb()
        win.destroy()

    ttk.Button(win, text="Salvează modificările", command=save_edit)\
        .grid(row=len(labels), column=0, columnspan=2, pady=10)


def cancel_reservation(root, loc_id, load_cb):
    if not messagebox.askyesno(
        "Confirmă anulare",
        "Sigur vrei să anulezi rezervarea/închirierea acestei locații?"
    ):
        return

    cur = conn.cursor()
    # Resetăm câmpurile în locatii
    cur.execute("""
        UPDATE locatii
           SET status     = 'Disponibil',
               client     = NULL,
               data_start = NULL,
               data_end   = NULL
         WHERE id = ?
    """, (loc_id,))

    # Ștergem și intrările din tabelul rezervari
    cur.execute("DELETE FROM rezervari WHERE loc_id=?", (loc_id,))

    conn.commit()

    load_cb()


def open_reserve_window(root, loc_id, load_cb, user):
    """Rezervă locația pentru 5 zile folosind doar numele clientului."""
    win = tk.Toplevel(root)
    win.title(f"Rezervă locația #{loc_id}")

    ttk.Label(win, text="Client:").grid(row=0, column=0, padx=5, pady=5)
    entry_client = ttk.Entry(win, width=30)
    entry_client.grid(row=0, column=1, padx=5, pady=5)

    def save():
        name = entry_client.get().strip()
        if not name:
            messagebox.showwarning("Lipsește client", "Completează numele clientului.")
            return
        start = datetime.date.today()
        end = start + datetime.timedelta(days=4)
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO rezervari (loc_id, client, data_start, data_end, created_by) VALUES (?, ?, ?, ?, ?)",
            (loc_id, name, start.isoformat(), end.isoformat(), user["username"]),
        )
        conn.commit()
        update_statusuri_din_rezervari()
        load_cb()
        win.destroy()

    ttk.Button(win, text="Rezervă", command=save).grid(row=1, column=0, columnspan=2, pady=10)

def open_rent_window(root, loc_id, load_cb, user):
    """Dialog pentru adăugarea unei închirieri în tabelul ``rezervari``.

    Perioada aleasă trebuie să nu se suprapună peste o rezervare sau o
    închiriere existentă pentru aceeași locație. După salvare statusurile sunt
    recalculate prin ``update_statusuri_din_rezervari``.
    """

    win = tk.Toplevel(root)
    win.title(f"Închiriază locația #{loc_id}")

    ttk.Label(win, text="Client:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    def client_list():
        return [r[0] for r in conn.cursor().execute("SELECT nume FROM clienti ORDER BY nume").fetchall()]
    cb_client = ttk.Combobox(win, values=client_list(), width=27)
    cb_client.grid(row=0, column=1, padx=5, pady=5)
    ttk.Button(win, text="+", command=lambda: (open_add_client_window(win, lambda: cb_client.configure(values=client_list())))).grid(row=0, column=2, padx=2, pady=5)

    ttk.Label(win, text="Client final (dacă agenție):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    entry_final = ttk.Entry(win, width=30)
    entry_final.grid(row=1, column=1, padx=5, pady=5)

    ttk.Label(win, text="Campanie:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    entry_campaign = ttk.Entry(win, width=30)
    entry_campaign.grid(row=2, column=1, padx=5, pady=5)

    ttk.Label(win, text="Firmă facturare:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
    def firm_list():
        return [r[0] for r in conn.cursor().execute("SELECT nume FROM firme ORDER BY nume").fetchall()]
    cb_firma = ttk.Combobox(win, values=firm_list(), width=27, state="readonly")
    cb_firma.current(0)
    cb_firma.grid(row=3, column=1, padx=5, pady=5)

    labels = ["Data start", "Data end", "Sumă finală"]
    entries = {}
    for i, lbl in enumerate(labels, start=4):
        ttk.Label(win, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=5, pady=5)
        if "Data" in lbl:
            e = DatePicker(win)
        else:
            e = ttk.Entry(win, width=30)
        e.grid(row=i, column=1, padx=5, pady=5)
        entries[lbl] = e

    deco_info = conn.cursor().execute("SELECT decoration_cost, sqm FROM locatii WHERE id=?", (loc_id,)).fetchone()
    deco_cost_default = deco_info[0] if deco_info else 0
    sqm_val = deco_info[1] if deco_info else 0

    ttk.Label(win, text="Data decorării:").grid(row=i+1, column=0, sticky="e", padx=5, pady=5)
    dp_deco = DatePicker(win)
    dp_deco.grid(row=i+1, column=1, padx=5, pady=5)

    def _sync_deco(*_):
        dp_deco.set_date(entries["Data start"].get_date())

    entries["Data start"].bind("<<DateEntrySelected>>", _sync_deco)
    _sync_deco()
    ttk.Label(win, text="Cost decorare (€):").grid(row=i+2, column=0, sticky="e", padx=5, pady=5)
    entry_deco = ttk.Entry(win, width=30)
    entry_deco.insert(0, str(deco_cost_default or 0))
    entry_deco.grid(row=i+2, column=1, padx=5, pady=5)
    var_prod = tk.BooleanVar(value=False)
    chk_prod = ttk.Checkbutton(win, text="Facem noi producția", variable=var_prod)
    chk_prod.grid(row=i+3, column=0, columnspan=2, sticky="w", padx=5, pady=2)
    ttk.Label(win, text="Preț producție (€):").grid(row=i+4, column=0, sticky="e", padx=5, pady=5)
    entry_prod = ttk.Entry(win, width=30)
    entry_prod.insert(0, f"{(sqm_val or 0)*7:.2f}")
    entry_prod.grid(row=i+4, column=1, padx=5, pady=5)
    def update_prod(*_):
        entry_prod.config(state="normal" if var_prod.get() else "disabled")
    var_prod.trace_add("write", update_prod)
    update_prod()

    def save_rent():
        client = cb_client.get().strip()
        if not client:
            messagebox.showwarning("Lipsește client", "Completează client.")
            return

        cur = conn.cursor()
        row = cur.execute("SELECT id, tip FROM clienti WHERE nume=?", (client,)).fetchone()
        if row:
            client_id, tip = row
        else:
            cur.execute("INSERT INTO clienti (nume) VALUES (?)", (client,))
            conn.commit()
            client_id = cur.lastrowid
            tip = "direct"

        final_client = entry_final.get().strip()
        client_display = client
        if tip == "agency":
            if not final_client:
                messagebox.showwarning("Lipsește client final", "Completează clientul final.")
                return
            client_display = f"{client} - {final_client}"

        start = entries["Data start"].get_date()
        end = entries["Data end"].get_date()
        if start > end:
            messagebox.showwarning(
                "Interval incorect",
                "«Data start» trebuie înainte de «Data end».",
            )
            return

        fee_txt = entries["Sumă finală"].get().strip()
        try:
            fee_val = float(fee_txt)
        except ValueError:
            messagebox.showwarning("Sumă invalidă", "Introdu o sumă numerică.")
            return

        cur = conn.cursor()
        campaign = entry_campaign.get().strip() or None
        firma_name = cb_firma.get().strip()
        firma_id = None
        if firma_name:
            fr = cur.execute("SELECT id FROM firme WHERE nume=?", (firma_name,)).fetchone()
            if fr:
                firma_id = fr[0]

        deco_date = dp_deco.get_date().isoformat()
        try:
            deco_cost = float(entry_deco.get() or 0)
        except ValueError:
            messagebox.showwarning("Date incorecte", "Cost decorare invalid.")
            return
        try:
            prod_cost = float(entry_prod.get() or 0) if var_prod.get() else 0.0
        except ValueError:
            messagebox.showwarning("Date incorecte", "Preț producție invalid.")
            return

        # verificăm suprapuneri cu alte perioade
        rows = cur.execute(
            "SELECT suma, created_by FROM rezervari WHERE loc_id=? AND NOT (data_end < ? OR data_start > ?)",
            (loc_id, start.isoformat(), end.isoformat()),
        ).fetchall()
        for suma, owner in rows:
            # Orice închiriere existentă blochează intervalul, iar o rezervare
            # făcută de alt vânzător nu poate fi suprascrisă.
            if suma is not None or owner != user["username"]:
                messagebox.showerror(
                    "Perioadă ocupată",
                    "Locația este deja rezervată sau închiriată în intervalul ales.",
                )
                return

        # Ștergem rezervările existente pentru intervalul ales
        cur.execute(
            "DELETE FROM rezervari WHERE loc_id=? AND suma IS NULL AND NOT (data_end < ? OR data_start > ?)",
            (loc_id, start.isoformat(), end.isoformat()),
        )

        # inserăm noua închiriere
        add_rental_with_decor(
            loc_id,
            client_display,
            client_id,
            start.isoformat(),
            end.isoformat(),
            fee_val,
            user["username"],
            campaign,
            firma_id,
            deco_date,
            deco_cost,
            prod_cost,
        )
        # actualizăm statusurile pe baza tuturor rezervărilor
        update_statusuri_din_rezervari()

        load_cb()
        win.destroy()

    ttk.Button(win, text="Confirmă închiriere", command=save_rent)\
        .grid(row=i+5, column=0, columnspan=3, pady=10)


def open_edit_rent_window(root, rid, load_cb):
    """Allow editing the rental period for reservation ``rid``."""
    cur = conn.cursor()
    row = cur.execute(
        "SELECT data_start, data_end FROM rezervari WHERE id=?",
        (rid,)
    ).fetchone()
    if not row:
        return

    ds, de = row
    win = tk.Toplevel(root)
    win.title(f"Modifică perioada #{rid}")

    ttk.Label(win, text="Data start:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    dp_start = DatePicker(win)
    dp_start.set_date(datetime.date.fromisoformat(ds))
    dp_start.grid(row=0, column=1, padx=5, pady=5)

    ttk.Label(win, text="Data end:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    dp_end = DatePicker(win)
    dp_end.set_date(datetime.date.fromisoformat(de))
    dp_end.grid(row=1, column=1, padx=5, pady=5)

    def save_edit():
        start = dp_start.get_date()
        end = dp_end.get_date()
        if start > end:
            messagebox.showwarning(
                "Interval incorect",
                "«Data start» trebuie înainte de «Data end».",
            )
            return
        cur.execute(
            "UPDATE rezervari SET data_start=?, data_end=? WHERE id=?",
            (start.isoformat(), end.isoformat(), rid),
        )
        conn.commit()
        update_statusuri_din_rezervari()
        load_cb()
        win.destroy()

    ttk.Button(win, text="Salvează", command=save_edit).grid(row=2, column=0, columnspan=2, pady=10)


def open_add_decor_window(root, rid):
    cur = conn.cursor()
    row = cur.execute(
        "SELECT loc_id FROM rezervari WHERE id=?",
        (rid,),
    ).fetchone()
    if not row:
        return
    loc_id = row[0]
    loc = cur.execute(
        "SELECT decoration_cost, sqm FROM locatii WHERE id=?",
        (loc_id,),
    ).fetchone()
    deco_default = loc[0] if loc else 0
    sqm_val = loc[1] if loc else 0
    win = tk.Toplevel(root)
    win.title("Adaugă decorare")

    ttk.Label(win, text="Data decorării:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    dp = DatePicker(win)
    dp.grid(row=0, column=1, padx=5, pady=5)

    ttk.Label(win, text="Cost decorare (€):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    entry_cost = ttk.Entry(win, width=20)
    entry_cost.insert(0, str(deco_default or 0))
    entry_cost.grid(row=1, column=1, padx=5, pady=5)

    var_prod = tk.BooleanVar(value=False)
    chk = ttk.Checkbutton(win, text="Facem noi producția", variable=var_prod)
    chk.grid(row=2, column=0, columnspan=2, sticky="w", padx=5, pady=2)

    ttk.Label(win, text="Preț producție (€):").grid(row=3, column=0, sticky="e", padx=5, pady=5)
    entry_prod = ttk.Entry(win, width=20)
    entry_prod.insert(0, f"{(sqm_val or 0)*7:.2f}")
    entry_prod.grid(row=3, column=1, padx=5, pady=5)

    def upd(*_):
        entry_prod.config(state="normal" if var_prod.get() else "disabled")
    var_prod.trace_add("write", upd)
    upd()

    def save():
        try:
            cost = float(entry_cost.get() or 0)
        except ValueError:
            messagebox.showwarning("Date", "Cost decorare invalid.")
            return
        try:
            prod = float(entry_prod.get() or 0) if var_prod.get() else 0.0
        except ValueError:
            messagebox.showwarning("Date", "Preț producție invalid.")
            return
        cur.execute(
            "INSERT INTO decorari (rez_id, data, cost, productie) VALUES (?,?,?,?)",
            (rid, dp.get_date().isoformat(), cost, prod),
        )
        conn.commit()
        win.destroy()

    ttk.Button(win, text="Salvează", command=save).grid(row=4, column=0, columnspan=2, pady=10)


def open_release_window(root, loc_id, load_cb, user):
    """Selectează și anulează o închiriere sau rezervare."""
    cur = conn.cursor()
    rows = cur.execute(
        "SELECT id, client, data_start, data_end, created_by, suma FROM rezervari WHERE loc_id=? ORDER BY data_start",
        (loc_id,)
    ).fetchall()
    if user.get("role") != "admin":
        rows = [r for r in rows if r[4] == user["username"] or r[5] is not None]

    if not rows:
        messagebox.showinfo("Eliberează", "Nu există închirieri pentru această locație.")
        return

    win = tk.Toplevel(root)
    win.title(f"Selectează închirierea #{loc_id}")

    lst = tk.Listbox(win, width=55, height=10)
    for rid, client, ds, de, *_ in rows:
        lst.insert("end", f"{client}: {ds} → {de}")
    lst.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

    def delete_selected():
        sel = lst.curselection()
        if not sel:
            messagebox.showwarning("Selectează", "Alege o închiriere.")
            return
        rid, _client, ds, _de, *_ = rows[sel[0]]

        if messagebox.askyesno(
            "Eliberează",
            "Ștergi complet perioada de închiriere?\nAlege 'Nu' pentru a modifica perioada.",
        ):
            cur.execute("DELETE FROM rezervari WHERE id=?", (rid,))
            conn.commit()
            update_statusuri_din_rezervari()
            load_cb()
            win.destroy()
        else:
            win.destroy()
            open_edit_rent_window(root, rid, load_cb)

    ttk.Button(win, text="Confirmă", command=delete_selected).grid(row=1, column=0, padx=5, pady=5)
    ttk.Button(win, text="Închide", command=win.destroy).grid(row=1, column=1, padx=5, pady=5)

    def add_decor():
        sel = lst.curselection()
        if not sel:
            messagebox.showwarning("Selectează", "Alege o închiriere.")
            return
        rid = rows[sel[0]][0]
        open_add_decor_window(win, rid)

    if hasattr(win, "tk"):
        ttk.Button(win, text="Adaugă decorare", command=add_decor).grid(row=1, column=2, padx=5, pady=5)


def export_available_excel(
    grup_filter, status_filter, search_term,
    ignore_dates, start_date, end_date
):
    import datetime
    import pandas as pd
    from tkinter import messagebox, filedialog
    from db import pandas_conn

    # 1) Construim WHERE identic cu load_locations()
    cond, params = [], []
    if grup_filter and grup_filter != "Toate":
        cond.append("grup = ?");       params.append(grup_filter)
    if status_filter and status_filter != "Toate":
        cond.append("status = ?");     params.append(status_filter)
    if search_term:
        cond.append("(city LIKE ? OR county LIKE ? OR address LIKE ?)")
        params += [f"%{search_term}%"] * 3
    if not ignore_dates:
        cond.append("""
            (data_start IS NULL
             OR data_end   IS NULL
             OR data_end   < ?
             OR data_start > ?)
        """)
        params += [start_date.isoformat(), end_date.isoformat()]

    where = ("WHERE " + " AND ".join(cond)) if cond else ""
    sql = f"""
        SELECT
          city, county, address, type,
          gps, photo_link,
          size, sqm, illumination,
          ratecard, decoration_cost,
          data_start, data_end,
          grup, status
        FROM locatii
        {where}
        ORDER BY grup, city, id
    """

    # 2) Citim datele
    df = pd.read_sql_query(
        sql, pandas_conn(), params=params,
        parse_dates=['data_start', 'data_end']
    )
    if df.empty:
        messagebox.showinfo("Export Excel", "Nu există locații pentru criteriile alese.")
        return

    # 3) Calculăm mesajul de Availability
    today = datetime.datetime.now().date()
    def avail_msg(r):
        ds, de = r['data_start'], r['data_end']
        if pd.notna(ds) and ds.date() > today:
            until = (ds.date() - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
            return f"Disponibil până la {until}"
        if pd.notna(de) and de.date() >= today:
            frm = (de.date() + datetime.timedelta(days=1)).strftime('%d.%m.%Y')
            return f"Disponibil din {frm}"
        if r['status'] != 'Disponibil' and pd.notna(de) and de.date() < today:
            frm = (de.date() + datetime.timedelta(days=1)).strftime('%d.%m.%Y')
            return f"Disponibil din {frm}"
        return "Disponibil"

    df['Availability'] = df.apply(avail_msg, axis=1)

    # 4) Coloane de export
    write_cols = [
        'city','county','address','type',
        'gps','photo_link',
        'size','sqm','illumination',
        'ratecard','decoration_cost',
        'Availability'
    ]

    # 5) Alege fișierul de salvat
    fp = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Salvează fișierul Excel"
    )
    if not fp:
        return

    # 6) Scriem Excel: câte o foaie per grup
    with pd.ExcelWriter(fp, engine='xlsxwriter') as writer:
        wb = writer.book
        center_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        money_fmt = wb.add_format({'num_format': '€#,##0.00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
        link_fmt = wb.add_format({'font_color': 'blue', 'underline': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        title_fmt = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
        hdr_fmt = wb.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})

        for grup, sub in df.groupby('grup'):
            grp_name = (grup or "").strip() or "FaraGrup"
            sheet = grp_name[:31]
            title = f"Locații {grp_name}"

            sub_df = sub.loc[:, write_cols].copy()
            sub_df.columns = [
                'City', 'County', 'Address', 'Type',
                'GPS', 'Photo Link',
                'Size', 'SQM', 'Illum',
                'Rate Card', 'Installation & Removal',
                'Availability'
            ]
            sub_df.insert(0, 'Nr', range(1, len(sub_df) + 1))

            startrow = 1
            ws = writer.book.add_worksheet(sheet)
            writer.sheets[sheet] = ws

            last_col = len(sub_df.columns) - 1
            ws.merge_range(0, 0, 0, last_col, title, title_fmt)

            for col_idx, name in enumerate(sub_df.columns):
                ws.write(startrow, col_idx, name, hdr_fmt)

            for row_idx, row in enumerate(sub_df.itertuples(index=False), start=startrow + 1):
                for col_idx, value in enumerate(row):
                    col_name = sub_df.columns[col_idx]
                    if pd.isna(value):
                        value = ""
                    if col_name in ('Rate Card', 'Installation & Removal'):
                        fmt = money_fmt
                    else:
                        fmt = center_fmt
                    ws.write(row_idx, col_idx, value, fmt)

            for idx, col_name in enumerate(sub_df.columns):
                if col_name in ('Rate Card', 'Installation & Removal'):
                    vals = pd.to_numeric(sub_df[col_name], errors='coerce').fillna(0)
                    formatted = [f"€{v:,.2f}" for v in vals]
                    max_len = max(len(col_name), *(len(v) for v in formatted))
                elif col_name == 'GPS':
                    max_len = max(len(col_name), len('Maps'))
                elif col_name == 'Photo Link':
                    max_len = max(len(col_name), len('Photo'))
                else:
                    max_len = max(len(col_name), sub_df[col_name].astype(str).map(len).max())
                ws.set_column(idx, idx, max_len + 2)

            gi = sub_df.columns.get_loc('GPS')
            for r, coord in enumerate(sub['gps'], start=startrow + 1):
                if coord:
                    url = f"https://www.google.com/maps/search/?api=1&query={coord}"
                    ws.write_url(r, gi, url, link_fmt, string="Maps")

            pi = sub_df.columns.get_loc('Photo Link')
            for r, u in enumerate(sub['photo_link'], start=startrow + 1):
                if u and u.strip():
                    url = u.strip()
                    if not url.lower().startswith(('http://', 'https://')):
                        url = 'https://' + url
                    ws.write_url(r, pi, url, link_fmt, string='Photo')
                else:
                    ws.write(r, pi, '', center_fmt)

    messagebox.showinfo("Export Excel", f"Am salvat locațiile în:\n{fp}")


def export_sales_report():
    """Exportă un raport structurat pe luni cu informații despre vânzări."""
    import pandas as pd
    import datetime
    from tkinter import messagebox, filedialog
    from db import pandas_conn, update_statusuri_din_rezervari

    update_statusuri_din_rezervari()

    df_loc = pd.read_sql_query(
        """
        SELECT id, city, county, address, type, size, sqm, illumination,
               ratecard, pret_vanzare, grup, status
          FROM locatii
         ORDER BY county, city, id
        """,
        pandas_conn(),
    )

    if df_loc.empty:
        messagebox.showinfo("Export Excel", "Nu există locații în baza de date.")
        return


    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        title="Salvează raportul Excel",
    )
    if not path:
        return


    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#4F81BD",
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        text_fmt = wb.add_format({"align": "center", "valign": "vcenter", "border": 1})
        money_fmt = wb.add_format({
            "num_format": "€#,##0.00",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        sold_text_fmt = wb.add_format({"align": "center", "valign": "vcenter", "border": 1, "bg_color": "#D9E1F2"})
        sold_money_fmt = wb.add_format({
            "num_format": "€#,##0.00",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#D9E1F2",
        })
        percent_fmt = wb.add_format({"num_format": "0.00%", "align": "center", "valign": "vcenter", "border": 1})

        stat_lbl_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#FFF2CC",
            "font_size": 14,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        stat_money_fmt = wb.add_format({
            "num_format": "€#,##0.00",
            "bold": True,
            "bg_color": "#FFF2CC",
            "font_size": 14,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        stat_money_pos_fmt = wb.add_format({
            "num_format": "€#,##0.00",
            "bold": True,
            "bg_color": "#FFF2CC",
            "font_size": 14,
            "font_color": "green",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        stat_money_neg_fmt = wb.add_format({
            "num_format": "€#,##0.00",
            "bold": True,
            "bg_color": "#FFF2CC",
            "font_size": 14,
            "font_color": "red",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        stat_percent_fmt = wb.add_format({
            "num_format": "0.00%",
            "bold": True,
            "bg_color": "#FFF2CC",
            "font_size": 14,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        stat_int_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#FFF2CC",
            "font_size": 14,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })

        money_cols = {"Ratecard/month", "PRET DE VANZARE", "PRET DE INCHIRIERE"}
        # Number of columns to span when merging statistic labels. A small value
        # keeps the merged cells narrow so the report prints nicely on A3 paper.
        STAT_MERGE_END = 3  # merge columns 0..STAT_MERGE_END

        def write_sheet(name, df_sheet):
            df_sheet = df_sheet.copy()
            df_sheet["Perioada"] = ""
            mask = df_sheet["data_start"].notna() & df_sheet["data_end"].notna()
            df_sheet.loc[mask, "Perioada"] = (
                df_sheet.loc[mask, "data_start"].dt.strftime("%d.%m.%Y")
                + " → "
                + df_sheet.loc[mask, "data_end"].dt.strftime("%d.%m.%Y")
            )
            df_sheet = df_sheet[[
                "city", "county", "address", "type", "size", "sqm", "illumination",
                "ratecard", "pret_vanzare", "suma", "client", "Perioada", "status"
            ]]
            df_sheet.columns = [
                "City", "County", "Address", "Type", "Size", "SQM", "Illum",
                "Ratecard/month", "PRET DE VANZARE", "PRET DE INCHIRIERE", "Client", "Perioada", "status"
            ]
            df_sheet.insert(0, "Nr", range(1, len(df_sheet) + 1))

            ws = wb.add_worksheet(name)
            writer.sheets[name] = ws

            for col_idx, col_name in enumerate(df_sheet.columns[:-1]):
                ws.write(0, col_idx, col_name, hdr_fmt)

            for row_idx, row in enumerate(df_sheet.itertuples(index=False), start=1):
                sold = row.status == "Închiriat"
                for col_idx, value in enumerate(row[:-1]):
                    col_name = df_sheet.columns[col_idx]
                    if pd.isna(value):
                        value = ""
                    if col_name in money_cols:
                        fmt = sold_money_fmt if sold else money_fmt
                    else:
                        fmt = sold_text_fmt if sold else text_fmt
                    ws.write(row_idx, col_idx, value, fmt)

            for idx, col_name in enumerate(df_sheet.columns[:-1]):
                if col_name in money_cols:
                    vals = pd.to_numeric(df_sheet[col_name], errors="coerce").fillna(0)
                    formatted = [f"€{v:,.2f}" for v in vals]
                    max_len = max(len(col_name), *(len(v) for v in formatted))
                else:
                    max_len = max(len(col_name), df_sheet[col_name].astype(str).map(len).max())
                ws.set_column(idx, idx, max_len + 2)

            sold_mask = df_sheet["status"] == "Închiriat"
            pct_sold = sold_mask.mean()
            pct_free = 1 - pct_sold
            count_sold = int(sold_mask.sum())
            count_free = int(len(df_sheet) - count_sold)
            sum_sale_total = pd.to_numeric(df_sheet["PRET DE VANZARE"], errors="coerce").fillna(0)
            sum_sale = sum_sale_total.sum()
            sum_sale_sold = sum_sale_total[sold_mask].sum()
            sum_sale_free = sum_sale_total[~sold_mask].sum()

            start = len(df_sheet) + 2
            ws.merge_range(start, 0, start, STAT_MERGE_END, "Locații vândute", stat_lbl_fmt)
            ws.merge_range(
                start,
                STAT_MERGE_END + 1,
                start,
                len(df_sheet.columns) - 1,
                f"{count_sold} ({pct_sold:.2%})",
                stat_int_fmt,
            )
            ws.merge_range(start + 1, 0, start + 1, STAT_MERGE_END, "Locații nevândute", stat_lbl_fmt)
            ws.merge_range(
                start + 1,
                STAT_MERGE_END + 1,
                start + 1,
                len(df_sheet.columns) - 1,
                f"{count_free} ({pct_free:.2%})",
                stat_int_fmt,
            )
            ws.merge_range(start + 2, 0, start + 2, STAT_MERGE_END, "Preț vânzare total", stat_lbl_fmt)
            ws.merge_range(
                start + 2,
                STAT_MERGE_END + 1,
                start + 2,
                len(df_sheet.columns) - 1,
                sum_sale,
                stat_money_fmt,
            )
            pct_sale_sold = sum_sale_sold / sum_sale if sum_sale else 0
            pct_sale_free = sum_sale_free / sum_sale if sum_sale else 0
            ws.merge_range(start + 3, 0, start + 3, STAT_MERGE_END, "Sumă locații vândute", stat_lbl_fmt)
            ws.merge_range(
                start + 3,
                STAT_MERGE_END + 1,
                start + 3,
                len(df_sheet.columns) - 1,
                f"€{sum_sale_sold:,.2f} ({pct_sale_sold:.2%})",
                stat_money_pos_fmt,
            )
            ws.merge_range(start + 4, 0, start + 4, STAT_MERGE_END, "Sumă locații nevândute", stat_lbl_fmt)
            ws.merge_range(
                start + 4,
                STAT_MERGE_END + 1,
                start + 4,
                len(df_sheet.columns) - 1,
                f"€{sum_sale_free:,.2f} ({pct_sale_free:.2%})",
                stat_money_neg_fmt,
            )

        current_year = datetime.date.today().year
        df_rez = pd.read_sql_query(
            """
            SELECT l.id, l.grup, l.city, l.county, l.address, l.type, l.size, l.sqm, l.illumination,
                   l.ratecard, l.pret_vanzare, r.client, r.data_start, r.data_end, r.suma
              FROM rezervari r
              JOIN locatii l ON r.loc_id = l.id
             ORDER BY r.data_start
            """,
            pandas_conn(),
            parse_dates=["data_start", "data_end"],
        )

        df_base = df_loc[[
            "id","city","county","address","type","size","sqm","illumination",
            "ratecard","pret_vanzare","grup","status"
        ]].copy()
        sold_count = {loc_id: 0 for loc_id in df_base["id"]}
        for month in range(1, 13):
            start_m = pd.Timestamp(current_year, month, 1)
            end_m = start_m + pd.offsets.MonthEnd(0)
            mask = (df_rez["data_end"] >= start_m) & (df_rez["data_start"] <= end_m)
            ids = df_rez.loc[mask, "id"].unique()
            for loc_id in ids:
                sold_count[loc_id] += 1

        # Sheet summarizing the entire year
        df_total = df_base.copy()
        df_total["Sold Months"] = df_total["id"].map(sold_count)
        df_total["% Year Sold"] = df_total["Sold Months"] / 12
        df_total["pret_vanzare"] = pd.to_numeric(df_total["pret_vanzare"], errors="coerce").fillna(0)
        grp_order = df_total.groupby("grup")["pret_vanzare"].max().sort_values(ascending=False).index
        order_map = {g: i for i, g in enumerate(grp_order)}
        df_total["__grp"] = df_total["grup"].map(order_map)
        df_total = df_total.sort_values(["__grp","pret_vanzare"], ascending=[True, False]).drop(columns="__grp")

        def write_total_sheet(df_sheet):
            df_sheet = df_sheet[[
                "city","county","address","type","size","sqm","illumination",
                "ratecard","pret_vanzare","Sold Months","% Year Sold"
            ]].copy()
            df_sheet.columns = [
                "City","County","Address","Type","Size","SQM","Illum",
                "Ratecard/month","PRET DE VANZARE","Luni vândută","% An vândut"
            ]
            df_sheet.insert(0, "Nr", range(1, len(df_sheet) + 1))
            ws = wb.add_worksheet("Total")
            writer.sheets["Total"] = ws
            for col_idx, col_name in enumerate(df_sheet.columns):
                ws.write(0, col_idx, col_name, hdr_fmt)
            for row_idx, row in enumerate(df_sheet.itertuples(index=False), start=1):
                for col_idx, value in enumerate(row):
                    fmt = money_fmt if col_idx == df_sheet.columns.get_loc("PRET DE VANZARE") else text_fmt
                    if col_idx == df_sheet.columns.get_loc("% An vândut"):
                        ws.write(row_idx, col_idx, value, percent_fmt)
                    else:
                        ws.write(row_idx, col_idx, value, fmt)
            for idx, col in enumerate(df_sheet.columns):
                if col == "PRET DE VANZARE":
                    vals = pd.to_numeric(df_sheet[col], errors="coerce").fillna(0)
                    formatted = [f"€{v:,.2f}" for v in vals]
                    width = max(len(col), *(len(v) for v in formatted)) + 2
                else:
                    width = max(len(col), df_sheet[col].astype(str).map(len).max()) + 2
                ws.set_column(idx, idx, width)
            overall_pct = df_sheet["% An vândut"].mean()
            start = len(df_sheet) + 2
            ws.merge_range(start, 0, start, STAT_MERGE_END, "% Total an", stat_lbl_fmt)
            ws.write(start, len(df_sheet.columns)-1, overall_pct, stat_percent_fmt)

        write_total_sheet(df_total)

        for month in range(1, 13):
            start_m = pd.Timestamp(current_year, month, 1)
            end_m = start_m + pd.offsets.MonthEnd(0)
            mask = (df_rez["data_end"] >= start_m) & (df_rez["data_start"] <= end_m)
            sub = df_rez.loc[mask]
            sub = sub.sort_values("data_start").groupby("id", as_index=False).first()
            df_month = df_base.merge(sub[["id","client","data_start","data_end","suma"]], on="id", how="left")
            df_month["status"] = df_month["client"].apply(lambda x: "Închiriat" if pd.notna(x) else "Disponibil")
            df_month["pret_vanzare"] = pd.to_numeric(df_month["pret_vanzare"], errors="coerce").fillna(0)
            grp_order = df_month.groupby("grup")["pret_vanzare"].max().sort_values(ascending=False).index
            order_map = {g: i for i, g in enumerate(grp_order)}
            df_month["__grp"] = df_month["grup"].map(order_map)
            df_month = df_month.sort_values(["__grp","pret_vanzare"], ascending=[True, False]).drop(columns="__grp")
            name = start_m.strftime("%B")
            write_sheet(name, df_month)

        messagebox.showinfo("Export Excel", f"Raport salvat:\n{path}")






def open_offer_window(tree):
    import pandas as pd
    import datetime
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    import os
    from utils import PREVIEW_FOLDER
    from db import pandas_conn

    # 1. Preluare selecție
    sel = tree.selection()
    if not sel:
        messagebox.showwarning("Nimic selectat", "Selectează cel puțin o locație.")
        return
    ids = [int(i) for i in sel]

    # 2. Fereastră setări ofertă
    win = tk.Toplevel(tree.master)
    win.title("Setări ofertă personalizată")

    # 2a. Parametri generali
    params = [
        ("Discount (%)", "20"),
        ("Decorare / m² (€)", "10"),
        ("Producție / m² (€)", "5")
    ]
    entries = {}
    for i, (label, default) in enumerate(params):
        ttk.Label(win, text=label + ":").grid(row=i, column=0, sticky="e", padx=5, pady=3)
        e = ttk.Entry(win, width=10)
        e.insert(0, default)
        e.grid(row=i, column=1, padx=5, pady=3)
        entries[label] = e

    # 2b. Alegere Preț de Bază
    price_var = tk.StringVar(value="ratecard")
    ttk.Label(win, text="Preț de bază:").grid(row=3, column=0, sticky="e", padx=5, pady=3)
    frm_price = ttk.Frame(win)
    frm_price.grid(row=3, column=1, sticky="w", padx=5, pady=3)
    ttk.Radiobutton(frm_price, text="Rate Card", variable=price_var, value="ratecard").pack(side="left")
    ttk.Radiobutton(frm_price, text="Preț Vânzare", variable=price_var, value="pret_vanzare").pack(side="left")

    # 2c. Checkbox Discount personalizat
    personal_var = tk.BooleanVar(value=False)
    chk_personal = ttk.Checkbutton(win, text="Discount personalizat per locație", variable=personal_var)
    chk_personal.grid(row=4, column=0, columnspan=2, pady=5)

    def export():
        # 3. Validare input
        try:
            disc_pct  = float(entries["Discount (%)"].get())
            cost_deco = float(entries["Decorare / m² (€)"].get())
            cost_prod = float(entries["Producție / m² (€)"].get())
        except ValueError:
            messagebox.showerror("Date invalide", "Introdu valori numerice valide.")
            return

        # 4. Citire date din DB (adăugăm address + pret_vanzare)
        sql = (
            f"SELECT id, city, county, address, gps, code, photo_link, sqm, type, "
            f"ratecard, pret_vanzare, data_start, data_end "
            f"FROM locatii WHERE id IN ({','.join(['?']*len(ids))})"
        )
        df = pd.read_sql_query(
            sql,
            pandas_conn(),
            params=ids,
            parse_dates=["data_start", "data_end"],
        )

        # 5. Calcul disponibilitate
        today = datetime.date.today()
        def avail(r):
            ds, de = r['data_start'], r['data_end']
            if pd.notna(ds) and ds.date() > today:
                return f"Până pe {(ds.date() - datetime.timedelta(days=1)).strftime('%d.%m.%Y')}"
            if pd.notna(de) and de.date() >= today:
                return f"Din {(de.date() + datetime.timedelta(days=1)).strftime('%d.%m.%Y')}"
            return "Disponibil"
        df['Availability'] = df.apply(avail, axis=1)

        # 6. Calcul costuri
        df['Installation & Removal'] = df['sqm'] * cost_deco
        df['Production']             = df['sqm'] * cost_prod

        # 7. Alege prețul de bază
        base_col = price_var.get()  # 'ratecard' sau 'pret_vanzare'
        df['Base Price'] = df[base_col].fillna(0).astype(float)

        # Funcție comună de scriere Excel
        def write_excel(df_export):
            df_export = df_export.sort_values(by=['City', 'CODE'])
            df_export.insert(0, 'No.', range(1, len(df_export) + 1))

            fp = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel', '*.xlsx')],
                title='Salvează oferta'
            )
            if not fp:
                return

            with pd.ExcelWriter(fp, engine='xlsxwriter') as writer:
                sheet = 'Ofertă'
                wb = writer.book
                ws = wb.add_worksheet(sheet)
                writer.sheets[sheet] = ws

                title_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_size': 14})
                hdr_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1})
                txt_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                money_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'num_format': '€#,##0.00', 'border': 1})
                link_fmt = wb.add_format({'font_color': 'blue', 'underline': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
                percent_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'num_format': '0.00%', 'border': 1})

                last_col = len(df_export.columns) - 1
                ws.merge_range(0, 0, 0, last_col, 'OFERTĂ PERSONALIZATĂ', title_fmt)

                for col_idx, hdr in enumerate(df_export.columns):
                    ws.write(1, col_idx, hdr, hdr_fmt)

                money_cols = {'Base Price', 'Final Price', 'Installation & Removal', 'Production'}
                percent_cols = {'% Discount'}

                for row_idx, row in enumerate(df_export.itertuples(index=False), start=2):
                    for col_idx, value in enumerate(row):
                        col_name = df_export.columns[col_idx]
                        if col_name in money_cols:
                            fmt = money_fmt
                        elif col_name in percent_cols:
                            fmt = percent_fmt
                        else:
                            fmt = txt_fmt
                        ws.write(row_idx, col_idx, value, fmt)

                for idx, col in enumerate(df_export.columns):
                    if col in money_cols:
                        vals = pd.to_numeric(df_export[col], errors='coerce').fillna(0)
                        width = max(len(col), *(len(f"€{v:,.2f}") for v in vals)) + 2
                    elif col == 'GPS':
                        width = max(len(col), len('Maps')) + 2
                    elif col == 'Photo Link':
                        width = max(len(col), len('Photo')) + 2
                    else:
                        width = max(len(col), df_export[col].astype(str).map(len).max()) + 2
                    ws.set_column(idx, idx, width)

                gps_idx = df_export.columns.get_loc('GPS')
                for r, coord in enumerate(df_export['GPS'], start=2):
                    if isinstance(coord, str) and coord.strip():
                        ws.write_url(r, gps_idx,
                                     f"https://www.google.com/maps/search/?api=1&query={coord.strip()}",
                                     link_fmt, string='Maps')

                photo_idx = df_export.columns.get_loc('Photo Link')
                for r, url in enumerate(df_export['Photo Link'], start=2):
                    if isinstance(url, str) and url.strip():
                        link = url.strip()
                        if not link.lower().startswith(('http://', 'https://')):
                            link = 'https://' + link
                        ws.write_url(r, photo_idx, link, link_fmt, string='Photo')

            messagebox.showinfo('Succes', f'Fișierul a fost salvat:\n{fp}')
            win.destroy()

        # 8. Branch pentru discount personalizat
        if personal_var.get():
            win2 = tk.Toplevel(win)
            win2.title("Discount pe locație")
            entry_map = {}
            ttk.Label(win2, text="Adresa").grid(row=0, column=0, padx=5, pady=3)
            ttk.Label(win2, text="Discount (%)").grid(row=0, column=1, padx=5, pady=3)
            for i, row in df.iterrows():
                ttk.Label(win2, text=row['address']).grid(row=i+1, column=0, padx=5, pady=2)
                e = ttk.Entry(win2, width=5)
                e.insert(0, "0")
                e.grid(row=i+1, column=1, padx=5, pady=2)
                entry_map[row['id']] = e
            def on_ok():
                # Aplic discount per locație (în fracțiune)
                discounts = {loc_id: float(ent.get()) for loc_id, ent in entry_map.items()}
                df['% Discount']      = df['id'].map(discounts) / 100
                df['Discount Amount'] = df['Base Price'] * df['% Discount']
                df['Final Price']     = df['Base Price'] - df['Discount Amount']

                df_export = df[[
                    'city','county','address','code','gps','photo_link','sqm','type',
                    'Base Price','% Discount','Final Price',
                    'Installation & Removal','Production','Availability'
                ]].copy()
                df_export.columns = [
                    'City','County','Address','CODE','GPS','Photo Link','SQM','Type',
                    'Base Price','% Discount','Final Price',
                    'Installation & Removal','Production','Availability'
                ]
                write_excel(df_export)
                win2.destroy()

            ttk.Button(win2, text="OK", command=on_ok)\
                .grid(row=len(df)+1, column=0, columnspan=2, pady=10)

        else:
            # Discount global
            # Discount global (în fracțiune, pentru format % corect)
            df['% Discount']      = disc_pct / 100
            df['Discount Amount'] = df['Base Price'] * df['% Discount']
            df['Final Price']     = df['Base Price'] - df['Discount Amount']

            df_export = df[[
                'city','county','address','code','gps','photo_link','sqm','type',
                'Base Price','% Discount','Final Price',
                'Installation & Removal','Production','Availability'
            ]].copy()
            df_export.columns = [
                'City','County','Address','CODE','GPS','Photo Link','SQM','Type',
                'Base Price','% Discount','Final Price',
                'Installation & Removal','Production','Availability'
            ]
            write_excel(df_export)

    # 9. Butonul de export (am mutat rândul la 5 pentru UI)
    ttk.Button(win, text='Generează Excel', command=export)\
        .grid(row=5, column=0, columnspan=2, pady=10)


def open_add_client_window(parent, refresh_cb=None):
    """Dialog pentru adăugarea unui client în tabela ``clienti``."""
    win = tk.Toplevel(parent)
    win.title("Adaugă client")

    labels = [
        "Nume companie",
        "Persoană contact",
        "Email",
        "Telefon",
        "CUI",
        "Adresă facturare",
        "Observații",
    ]
    entries = {}
    for i, lbl in enumerate(labels):
        ttk.Label(win, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=5, pady=2)
        e = ttk.Entry(win, width=40)
        e.grid(row=i, column=1, padx=5, pady=2)
        entries[lbl] = e

    ttk.Label(win, text="Tip:").grid(row=len(labels), column=0, sticky="e", padx=5, pady=2)
    cb_tip = ttk.Combobox(win, values=["direct", "agency"], state="readonly", width=37)
    cb_tip.current(0)
    cb_tip.grid(row=len(labels), column=1, padx=5, pady=2)

    def save():
        nume = entries["Nume companie"].get().strip()
        if not nume:
            messagebox.showwarning("Lipsește numele", "Completează numele companiei.")
            return
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO clienti (nume, contact, email, phone, cui, adresa, observatii, tip) VALUES (?,?,?,?,?,?,?,?)",
            (
                nume,
                entries["Persoană contact"].get().strip(),
                entries["Email"].get().strip(),
                entries["Telefon"].get().strip(),
                entries["CUI"].get().strip(),
                entries["Adresă facturare"].get().strip(),
                entries["Observații"].get().strip(),
                cb_tip.get() or "direct",
            ),
        )
        conn.commit()
        if refresh_cb:
            refresh_cb()
        win.destroy()

    ttk.Button(win, text="Salvează", command=save).grid(row=len(labels)+1, column=0, columnspan=2, pady=10)


def open_edit_client_window(parent, cid, refresh_cb=None):
    cur = conn.cursor()
    row = cur.execute(
        "SELECT nume, contact, email, phone, cui, adresa, observatii, tip FROM clienti WHERE id=?",
        (cid,),
    ).fetchone()
    if not row:
        return
    win = tk.Toplevel(parent)
    win.title("Editează client")
    labels = [
        "Nume companie",
        "Persoană contact",
        "Email",
        "Telefon",
        "CUI",
        "Adresă facturare",
        "Observații",
    ]
    entries = {}
    for i, lbl in enumerate(labels):
        ttk.Label(win, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=5, pady=2)
        e = ttk.Entry(win, width=40)
        e.insert(0, row[i] if row[i] is not None else "")
        e.grid(row=i, column=1, padx=5, pady=2)
        entries[lbl] = e

    ttk.Label(win, text="Tip:").grid(row=len(labels), column=0, sticky="e", padx=5, pady=2)
    cb_tip = ttk.Combobox(win, values=["direct", "agency"], state="readonly", width=37)
    cb_tip.set(row[7] or "direct")
    cb_tip.grid(row=len(labels), column=1, padx=5, pady=2)

    def save_edit():
        cur.execute(
            "UPDATE clienti SET nume=?, contact=?, email=?, phone=?, cui=?, adresa=?, observatii=?, tip=? WHERE id=?",
            (
                entries["Nume companie"].get().strip(),
                entries["Persoană contact"].get().strip(),
                entries["Email"].get().strip(),
                entries["Telefon"].get().strip(),
                entries["CUI"].get().strip(),
                entries["Adresă facturare"].get().strip(),
                entries["Observații"].get().strip(),
                cb_tip.get() or "direct",
                cid,
            ),
        )
        conn.commit()
        if refresh_cb:
            refresh_cb()
        win.destroy()

    ttk.Button(win, text="Salvează", command=save_edit).grid(row=len(labels)+1, column=0, columnspan=2, pady=10)


def export_client_backup(month, year, client_id=None):
    """Exportă un backup de facturare pentru luna dată, formatat în Excel."""
    import calendar
    import pandas as pd

    days_in_month = calendar.monthrange(year, month)[1]
    start_m = datetime.date(year, month, 1)
    end_m = datetime.date(year, month, days_in_month)

    cur = conn.cursor()
    sql = (
        "SELECT c.nume, l.city, l.address, l.code, l.face, l.type, l.size, l.sqm, "
        "r.data_start, r.data_end, r.suma, r.client_id, r.id "
        "FROM rezervari r "
        "JOIN locatii l ON r.loc_id = l.id "
        "JOIN clienti c ON r.client_id = c.id "
        "WHERE NOT (r.data_end < ? OR r.data_start > ?)"
    )
    params = [start_m.isoformat(), end_m.isoformat()]
    if client_id:
        sql += " AND r.client_id=?"
        params.append(client_id)

    rows = cur.execute(sql, params).fetchall()
    if not rows:
        messagebox.showinfo("Export", "Nu există închirieri pentru perioada aleasă.")
        return

    data = []
    for nume, city, addr, code, face, typ, size, sqm, ds, de, price, cid, rez_id in rows:
        ds_dt = datetime.date.fromisoformat(ds)
        de_dt = datetime.date.fromisoformat(de)
        ov_start = max(ds_dt, start_m)
        ov_end = min(de_dt, end_m)
        days = (ov_end - ov_start).days + 1
        frac = days / days_in_month
        amount = price * frac
        deco, prod = cur.execute(
            "SELECT COALESCE(SUM(cost),0), COALESCE(SUM(productie),0) FROM decorari WHERE rez_id=? AND data BETWEEN ? AND ?",
            (rez_id, start_m.isoformat(), end_m.isoformat()),
        ).fetchone()

        data.append([
            city,
            addr,
            code,
            face,
            size,
            typ,
            ds_dt,
            de_dt,
            frac,
            "EUR",
            price,
            amount,
            deco or 0.0,
            prod or 0.0,
            nume,
        ])

    df = pd.DataFrame(
        data,
        columns=[
            "City",
            "Address",
            "Code",
            "Face",
            "Size",
            "Type",
            "Data start",
            "Data end",
            "Perioada",
            "Valuta",
            "Preț chirie/lună",
            "Chirie net",
            "Preț Decorare",
            "Preț Producție",
            "Client",
        ],
    )

    df["Perioada"] = df["Perioada"].round(2)

    df.insert(0, "Nr. Crt", range(1, len(df) + 1))

    total_rent = df["Chirie net"].sum()
    total_deco = df["Preț Decorare"].sum()
    total_prod = df["Preț Producție"].sum()
    total_gen = total_rent + total_deco + total_prod

    path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
    if not path:
        return

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Backup", index=False, startrow=0)
        wb = writer.book
        ws = writer.sheets["Backup"]

        hdr_fmt = wb.add_format({"bold": True, "bg_color": "#4F81BD", "font_color": "white", "align": "center"})
        euro_fmt = wb.add_format({"num_format": "€#,##0.00", "align": "center"})
        center_fmt = wb.add_format({"align": "center"})

        money_cols = {"Preț chirie/lună", "Chirie net", "Preț Decorare", "Preț Producție"}
        for col_idx, col in enumerate(df.columns):
            width = max(len(str(col)), df[col].astype(str).map(len).max()) + 2
            fmt = euro_fmt if col in money_cols else center_fmt
            ws.set_column(col_idx, col_idx, width, fmt)
            ws.write(0, col_idx, col, hdr_fmt)

        row_tot = len(df) + 1
        ws.write(row_tot, 0, "Total", wb.add_format({"bold": True}))
        ws.write(row_tot, df.columns.get_loc("Chirie net"), total_rent, euro_fmt)
        ws.write(row_tot, df.columns.get_loc("Preț Decorare"), total_deco, euro_fmt)
        ws.write(row_tot, df.columns.get_loc("Preț Producție"), total_prod, euro_fmt)
        row_tot += 1
        ws.write(row_tot, 0, "Total General", wb.add_format({"bold": True}))
        ws.write(row_tot, df.columns.get_loc("Preț Producție"), total_gen, euro_fmt)

    messagebox.showinfo("Export", f"Backup salvat:\n{path}")


def open_clients_window(root):
    """Fereastră pentru administrarea clienților."""
    win = tk.Toplevel(root)
    win.title("Clienți")

    cols = ("Nume", "Tip", "Contact", "Email", "Telefon", "CUI", "Adresă", "Observații")
    tree = ttk.Treeview(win, columns=cols, show="headings")
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor="w")
    tree.grid(row=0, column=0, columnspan=4, sticky="nsew")

    vsb = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
    vsb.grid(row=0, column=4, sticky="ns")

    def on_double(event):
        sel = tree.selection()
        if not sel:
            return
        open_edit_client_window(win, int(sel[0]), refresh)

    tree.bind("<Double-1>", on_double)
    tree.configure(yscroll=vsb.set)

    search_var = tk.StringVar()
    ttk.Label(win, text="Caută:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    entry_search = ttk.Entry(win, textvariable=search_var, width=30)
    entry_search.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    def on_search(*_):
        refresh()
    entry_search.bind("<KeyRelease>", on_search)

    win.columnconfigure(0, weight=1)
    win.rowconfigure(0, weight=1)

    def refresh():
        tree.delete(*tree.get_children())
        term = search_var.get().strip()
        sql = "SELECT id, nume, tip, contact, email, phone, cui, adresa, observatii FROM clienti"
        params = []
        if term:
            sql += " WHERE nume LIKE ?"
            params.append(f"%{term}%")
        sql += " ORDER BY nume"
        rows = conn.cursor().execute(sql, params).fetchall()
        for cid, nume, tip, contact, email, phone, cui, adresa, obs in rows:
            tree.insert(
                "",
                "end",
                iid=str(cid),
                values=(nume, tip or "direct", contact, email, phone, cui or "", adresa or "", obs or ""),
            )

    def add_client():
        open_add_client_window(win, refresh)

    def delete_client():
        sel = tree.selection()
        if not sel:
            return
        if not messagebox.askyesno("Confirmă", "Ștergi clientul selectat?"):
            return
        cur = conn.cursor()
        cur.execute("DELETE FROM clienti WHERE id=?", (int(sel[0]),))
        conn.commit()
        refresh()

    def export_current():
        sel = tree.selection()
        cid = int(sel[0]) if sel else None
        today = datetime.date.today()
        year = simpledialog.askinteger("An", "Anul", initialvalue=today.year, parent=win)
        if year is None:
            return
        month = simpledialog.askinteger(
            "Luna", "Luna (1-12)", minvalue=1, maxvalue=12, initialvalue=today.month, parent=win
        )
        if month is None:
            return
        export_client_backup(month, year, cid)

    btn_add = ttk.Button(win, text="Adaugă", command=add_client)
    btn_del = ttk.Button(win, text="Șterge", command=delete_client)
    btn_export = ttk.Button(win, text="Export Backup", command=export_current)
    for i, b in enumerate((btn_add, btn_del, btn_export)):
        b.grid(row=2, column=i, padx=5, pady=5, sticky="w")

    refresh()

def export_vendor_report():
    """Exporta un raport detaliat pentru fiecare vânzător."""
    import pandas as pd
    from tkinter import filedialog, messagebox
    from db import pandas_conn

    users = pd.read_sql_query(
        "SELECT username, comune FROM users WHERE role='seller'",
        pandas_conn(),
    )
    if users.empty:
        messagebox.showinfo("Raport", "Nu există vânzători.")
        return

    df = pd.read_sql_query(
        """
        SELECT r.created_by, r.suma, r.data_start, r.data_end,
               l.city, l.county, l.address
          FROM rezervari r
          JOIN locatii l ON r.loc_id = l.id
        """,
        pandas_conn(),
        parse_dates=["data_start", "data_end"],
    )

    if df.empty:
        messagebox.showinfo("Raport", "Nu există închirieri.")
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        title="Salvează raportul",
    )
    if not path:
        return

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#4F81BD",
            "font_color": "white",
            "align": "center",
        })
        money_fmt = wb.add_format({"num_format": "€#,##0.00", "align": "center"})
        center_fmt = wb.add_format({"align": "center"})

        for _, row in users.iterrows():
            uname = row["username"]
            comune = {c for c in (row["comune"] or "").split(',') if c}
            sub = df[df["created_by"] == uname]
            if comune:
                sub = sub[sub["county"].isin(comune)]
            if sub.empty:
                continue

            sub = sub.copy()
            days = (sub["data_end"] - sub["data_start"]).dt.days + 1
            sub["Luni"] = days / 30
            sub["Chirie/lună"] = sub["suma"] / sub["Luni"]
            sub["Valoare"] = sub["suma"]

            df_det = sub[[
                "city",
                "county",
                "address",
                "Chirie/lună",
                "Luni",
                "Valoare",
            ]].copy()
            df_det.columns = [
                "Oraș",
                "Județ",
                "Adresă",
                "Chirie/lună",
                "Luni",
                "Valoare",
            ]
            df_det.insert(0, "Nr.crt", range(1, len(df_det) + 1))

            sheet = uname[:31]
            df_det.to_excel(writer, sheet_name=sheet, index=False, startrow=0)

            ws = writer.sheets[sheet]
            for col_idx, col in enumerate(df_det.columns):
                width = max(len(str(col)), df_det[col].astype(str).map(len).max()) + 2
                fmt = money_fmt if col in {"Chirie/lună", "Valoare"} else center_fmt
                ws.set_column(col_idx, col_idx, width, fmt)
                ws.write(0, col_idx, col, hdr_fmt)

            sub["Month"] = sub["data_start"].dt.to_period("M").dt.strftime("%B")
            stats = sub.groupby("Month")["suma"].sum().reset_index()
            stats.columns = ["Luna", "Total"]

            start = len(df_det) + 3
            ws.write(start - 1, 0, "Statistici lunare", hdr_fmt)
            ws.write_row(start, 0, stats.columns, hdr_fmt)
            for i, r in stats.iterrows():
                ws.write(start + 1 + i, 0, r["Luna"], center_fmt)
                ws.write(start + 1 + i, 1, r["Total"], money_fmt)
            ws.set_column(0, 1, 15)

    messagebox.showinfo("Raport", f"Raport salvat:\n{path}")

def open_users_window(root):
    """Admin window to manage user accounts."""
    import tkinter.simpledialog as simpledialog
    win = tk.Toplevel(root)
    win.title("Utilizatori")

    tree = ttk.Treeview(win, columns=("User", "Role", "Comune"), show="headings")
    for c in ("User", "Role", "Comune"):
        tree.heading(c, text=c)
    tree.grid(row=0, column=0, columnspan=3, sticky="nsew")
    win.columnconfigure(0, weight=1)
    win.rowconfigure(0, weight=1)

    def refresh():
        tree.delete(*tree.get_children())
        rows = conn.cursor().execute("SELECT username, role, comune FROM users").fetchall()
        for u, r, c in rows:
            tree.insert("", "end", values=(u, r, c or ""))

    def add_user():
        dlg = tk.Toplevel(win)
        dlg.title("Adaugă utilizator")

        ttk.Label(dlg, text="Utilizator:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        entry_user = ttk.Entry(dlg, width=30)
        entry_user.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(dlg, text="Parola:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        entry_pass = ttk.Entry(dlg, width=30, show="*")
        entry_pass.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(dlg, text="Rol:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        combo_role = ttk.Combobox(dlg, values=["admin", "seller"], state="readonly")
        combo_role.current(1)
        combo_role.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(dlg, text="Comune:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        entry_comune = ttk.Entry(dlg, width=30)
        entry_comune.grid(row=3, column=1, padx=5, pady=5)

        def on_ok():
            u = entry_user.get().strip()
            if not u:
                messagebox.showwarning("Date lipsă", "Completează utilizatorul.", parent=dlg)
                return
            p = entry_pass.get()
            if p is None:
                messagebox.showwarning("Date lipsă", "Completează parola.", parent=dlg)
                return
            role = combo_role.get() or "seller"
            comune = entry_comune.get().strip() or ""
            create_user(u, p, role, comune)
            dlg.destroy()
            refresh()

        ttk.Button(dlg, text="Salvează", command=on_ok).grid(row=4, column=0, columnspan=2, pady=10)
        dlg.grab_set()
        entry_user.focus()

    def delete_user():
        sel = tree.selection()
        if not sel:
            return
        user = tree.item(sel[0])["values"][0]
        if messagebox.askyesno("Confirmă", f"Ștergi utilizatorul {user}?"):
            conn.cursor().execute("DELETE FROM users WHERE username=?", (user,))
            conn.commit()
            refresh()

    ttk.Button(win, text="Adaugă", command=add_user).grid(row=1, column=0, pady=5)
    ttk.Button(win, text="Șterge", command=delete_user).grid(row=1, column=1, pady=5)
    ttk.Button(win, text="Închide", command=win.destroy).grid(row=1, column=2, pady=5)

    refresh()


def open_firme_window(root):
    """Admin window to manage billing companies."""
    win = tk.Toplevel(root)
    win.title("Firme de facturare")

    tree = ttk.Treeview(win, columns=("Nume", "CUI", "Adresă"), show="headings")
    for c in ("Nume", "CUI", "Adresă"):
        tree.heading(c, text=c)
    tree.grid(row=0, column=0, columnspan=3, sticky="nsew")
    win.columnconfigure(0, weight=1)
    win.rowconfigure(0, weight=1)

    def refresh():
        tree.delete(*tree.get_children())
        rows = conn.cursor().execute(
            "SELECT id, nume, cui, adresa FROM firme ORDER BY nume"
        ).fetchall()
        for fid, nume, cui, adresa in rows:
            tree.insert("", "end", iid=str(fid), values=(nume, cui or "", adresa or ""))

    def save_entry(fid=None):
        dlg = tk.Toplevel(win)
        dlg.title("Editează" if fid else "Adaugă firmă")
        lbls = ["Nume", "CUI", "Adresă"]
        ents = {}
        for i, l in enumerate(lbls):
            ttk.Label(dlg, text=l + ":").grid(row=i, column=0, padx=5, pady=5, sticky="e")
            e = ttk.Entry(dlg, width=40)
            e.grid(row=i, column=1, padx=5, pady=5)
            ents[l] = e
        if fid:
            r = conn.cursor().execute(
                "SELECT nume, cui, adresa FROM firme WHERE id=?", (fid,)
            ).fetchone()
            if r:
                for val, l in zip(r, lbls):
                    ents[l].insert(0, val or "")

        def ok():
            nume = ents["Nume"].get().strip()
            if not nume:
                messagebox.showwarning("Date lipsă", "Completează numele.", parent=dlg)
                return
            cui = ents["CUI"].get().strip()
            adresa = ents["Adresă"].get().strip()
            cur = conn.cursor()
            if fid:
                cur.execute(
                    "UPDATE firme SET nume=?, cui=?, adresa=? WHERE id=?",
                    (nume, cui, adresa, fid),
                )
            else:
                cur.execute(
                    "INSERT INTO firme (nume, cui, adresa) VALUES (?,?,?)",
                    (nume, cui, adresa),
                )
            conn.commit()
            dlg.destroy()
            refresh()

        ttk.Button(dlg, text="Salvează", command=ok).grid(row=3, column=0, columnspan=2, pady=10)
        dlg.grab_set()
        ents["Nume"].focus()

    def add_firm():
        save_entry()

    def edit_firm():
        sel = tree.selection()
        if not sel:
            return
        fid = int(sel[0])
        save_entry(fid)

    def delete_firm():
        sel = tree.selection()
        if not sel:
            return
        fid = int(sel[0])
        if messagebox.askyesno("Confirmă", "Ștergi firma selectată?"):
            conn.cursor().execute("DELETE FROM firme WHERE id=?", (fid,))
            conn.commit()
            refresh()

    ttk.Button(win, text="Adaugă", command=add_firm).grid(row=1, column=0, pady=5)
    ttk.Button(win, text="Editează", command=edit_firm).grid(row=1, column=1, pady=5)
    ttk.Button(win, text="Șterge", command=delete_firm).grid(row=1, column=2, pady=5)
    ttk.Button(win, text="Închide", command=win.destroy).grid(row=2, column=1, pady=5)

    refresh()
