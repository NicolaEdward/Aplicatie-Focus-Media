
 main

import os
import shutil
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry

from db import conn, cursor, update_statusuri_din_rezervari
from utils import make_preview, get_schita_path
from UI.dialogs import (
    open_detail_window,
    open_add_window,
    open_edit_window,

    open_reserve_window,
    open_rent_window,
 main
    cancel_reservation,
 main
    open_offer_window,
    export_available_excel,
    export_sales_report
)



 main
def start_app():
    root = tk.Tk()
    root.title("Gestionare Locații Publicitare")
    root.geometry("1200x600")

    # --- Top: filtre Grup, Status, Căutare, Interval ---
    frm_top = ttk.Frame(root, padding=10)
    frm_top.pack(fill="x", side="top")

    ttk.Label(frm_top, text="Grup:").pack(side="left")
    combo_group = ttk.Combobox(frm_top, values=[], state="readonly", width=20)
    combo_group.pack(side="left", padx=5)

    ttk.Label(frm_top, text="Status:").pack(side="left", padx=(20,5))
    combo_status = ttk.Combobox(
        frm_top,
        values=["Toate", "Disponibil", "Rezervat", "Închiriat"],
        state="readonly",
        width=12
    )
    combo_status.current(0)
    combo_status.pack(side="left", padx=5)

    ttk.Label(frm_top, text="Caută:").pack(side="left", padx=(20,5))
    search_var = tk.StringVar()
    ttk.Entry(frm_top, textvariable=search_var, width=20).pack(side="left", padx=5)

    ttk.Label(frm_top, text="Din:").pack(side="left", padx=(20,5))
    filter_start = DateEntry(frm_top, date_pattern="yyyy-mm-dd", width=12)
    filter_start.pack(side="left", padx=5)
    ttk.Label(frm_top, text="Până:").pack(side="left", padx=(10,5))
    filter_end = DateEntry(frm_top, date_pattern="yyyy-mm-dd", width=12)
    filter_end.pack(side="left", padx=5)

    var_ignore = tk.BooleanVar(value=True)
    ttk.Checkbutton(
        frm_top,
        text="Toate datele",
        variable=var_ignore,
        command=lambda: load_locations()
    ).pack(side="left", padx=(20,0))

    # --- Middle: TreeView + Detalii ---
    frm_mid = ttk.Frame(root, padding=10)
    frm_mid.pack(fill="both", expand=True)

    cols = ("NR.", "City", "County", "Address", "Type", "Status", "RateCard")
    tree = ttk.Treeview(frm_mid, columns=cols, show="headings", selectmode="extended")
    tree.heading("NR.", text="NR.");        tree.column("NR.", width=50, anchor="w")
    tree.heading("City", text="City");      tree.column("City", width=120, anchor="w")
    tree.heading("County", text="County");  tree.column("County", width=100, anchor="w")
    tree.heading("Address", text="Address")
    tree.column("Address", width=300, anchor="w", stretch=True)
    tree.heading("Type", text="Type");      tree.column("Type", width=100, anchor="w")
    tree.heading("Status", text="Status");  tree.column("Status", width=120, anchor="w")
    tree.heading("RateCard", text="RateCard")
    tree.column("RateCard", width=100, anchor="e")
    tree.pack(fill="both", expand=True, side="left")

    vsb = ttk.Scrollbar(frm_mid, orient="vertical", command=tree.yview)
    vsb.pack(side="left", fill="y")
    tree.configure(yscroll=vsb.set)

    tree.bind("<Double-1>", lambda e: open_detail_window(tree, e))
    tree.bind("<<TreeviewSelect>>", lambda e: on_tree_select())

    # --- Panoul detalii (dreapta) ---
    details = ttk.Frame(frm_mid, padding=10, relief="groove", width=400)
    details.pack(side="right", fill="y")
    details.pack_propagate(False)

    img_label = ttk.Label(details)
    img_label.pack(pady=(0,10))


    for w in (btn_add, btn_edit, btn_rent, btn_delete):
 main
    btn_download.pack(pady=(0,15))

    # etichete declarate fără pack()
    lbl_client_label       = ttk.Label(details, text="Client:")
    lbl_client_value       = ttk.Label(details, text="-")
    lbl_period_label       = ttk.Label(details, text="Perioadă:")
    lbl_period_value       = ttk.Label(details, text="-")
    lbl_ratecard_label     = ttk.Label(details, text="Rate Card:")
    lbl_ratecard_value     = ttk.Label(details, text="-")
    lbl_pret_vanz_label    = ttk.Label(details, text="Preț de vânzare:")
    lbl_pret_vanz_value    = ttk.Label(details, text="-")
    lbl_pret_flot_label    = ttk.Label(details, text="Preț flotant:")
    lbl_pret_flot_value    = ttk.Label(details, text="-")

 main

    # --- Bottom: butoane principale (stânga) și export (dreapta) ---
    frm_bot = ttk.Frame(root, padding=10)
    frm_bot.pack(fill="x", side="bottom")

    primary_frame = ttk.Frame(frm_bot)
    primary_frame.pack(side="left")
    btn_add     = ttk.Button(primary_frame, text="Adaugă",
                             command=lambda: open_add_window(root, load_locations))
    btn_edit    = ttk.Button(primary_frame, text="Editează", state="disabled",
                             command=lambda: open_edit_window(root, selected_id[0], load_locations, refresh_groups))

    btn_reserve = ttk.Button(primary_frame, text="Rezervă", state="disabled")
    btn_rent    = ttk.Button(primary_frame, text="Închiriază", state="disabled")
    btn_delete  = ttk.Button(primary_frame, text="Șterge", state="disabled",
                             command=lambda: delete_location())
    for w in (btn_add, btn_edit, btn_reserve, btn_rent, btn_delete):
        w.pack(side="left", padx=5)


 main
    export_frame = ttk.Frame(frm_bot)
    export_frame.pack(side="right")
    btn_xlsx  = ttk.Button(export_frame, text="Export Disponibil",
                           command=lambda: export_available_excel(
                               combo_group.get(), combo_status.get(),
                               search_var.get().strip(),
                               var_ignore.get(),
                               filter_start.get_date(),
                               filter_end.get_date()
                           ))
    btn_offer = ttk.Button(export_frame, text="Export Ofertă",
                           command=lambda: open_offer_window(tree))
    btn_report = ttk.Button(export_frame, text="Raport Vânzări",
                           command=lambda: export_sales_report())
    btn_xlsx.pack(side="left", padx=5)
    btn_offer.pack(side="left", padx=5)
    btn_report.pack(side="left", padx=5)



 main
    selected_id = [None]

    # --- Funcții auxiliare ---
    def refresh_groups():
        vals = ["Toate"] + [g[0] for g in cursor.execute(
            "SELECT DISTINCT grup FROM locatii").fetchall()]
        combo_group['values'] = vals
        if combo_group.get() not in vals:
            combo_group.current(0)

    def load_locations():
        # 1) Actualizează statusurile locațiilor pe baza rezervărilor
        update_statusuri_din_rezervari()

        # 2) Golește TreeView
        for iid in tree.get_children():
            tree.delete(iid)

        # 3) Construiește condițiile de filtrare pe Grup, Status și Căutare
        params, cond = [], []
        g = combo_group.get()
        if g and g != "Toate":
            cond.append("grup = ?");    params.append(g)
        s = combo_status.get()
        if s and s != "Toate":
            cond.append("status = ?");  params.append(s)
        term = search_var.get().strip()
        if term:
            cond.append("(city LIKE ? OR county LIKE ? OR address LIKE ?)")
            params += [f"%{term}%"]*3

        # 4) Citește intervalul Din–Până și normalizează-l
        start_dt = filter_start.get_date()
        end_dt   = filter_end.get_date()
        if end_dt < start_dt:
            end_dt = start_dt
        # îl folosim doar în availability()
        d0, d1 = start_dt.isoformat(), end_dt.isoformat()

 main
 main
        # 5) Interogarea inițială doar pe tabelă ``locatii``
        q = """
            SELECT id, city, county, address, type, ratecard
            FROM locatii
        """
        if cond:
            q += " WHERE " + " AND ".join(cond)


 main
        q += """
            ORDER BY
            CASE county
                WHEN 'Bucuresti Sectorul 1' THEN  1
                WHEN 'Bucuresti Sectorul 2' THEN  2
                WHEN 'Bucuresti Sectorul 3' THEN  3
                WHEN 'Bucuresti Sectorul 4' THEN  4
                WHEN 'Bucuresti Sectorul 5' THEN  5
                WHEN 'Bucuresti Sectorul 6' THEN  6
                WHEN 'Ilfov'                THEN  7
                WHEN 'Prahova'             THEN  8
                ELSE 9
            END ASC,
            county ASC,
            city ASC
        """
        rows = cursor.execute(q, params).fetchall()



        # 6) Funcție locală pentru a afla disponibilitatea pe interval
        def availability(loc_id):
            rez = cursor.execute(

                "SELECT data_start, data_end FROM rezervari WHERE loc_id=? ORDER BY data_start",
                (loc_id,)
            ).fetchall()
            periods = [(datetime.date.fromisoformat(ds), datetime.date.fromisoformat(de)) for ds,de in rez]
            overl = [(ds,de) for ds,de in periods if not (de < start_dt or ds > end_dt)]
            if not overl:
                return "Disponibil"
            first_ds = overl[0][0]
            if first_ds > start_dt:
                until = (first_ds - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
                return f"Disponibil până la {until}"
            last_de = overl[-1][1]
            if last_de < end_dt:
                frm = (last_de + datetime.timedelta(days=1)).strftime('%d.%m.%Y')
                return f"Disponibil din {frm}"
            return ""  # complet acoperit
 main

        # 7) Populează TreeView, aplicând filtrul de date doar când "Toate datele" NU e bifat
        display_index = 0
        for loc_id, city, county, addr, typ, rate in rows:
            if not var_ignore.get():
                avail = availability(loc_id)
                if not avail:
                    # nu se intersectează cu intervalul, deci nu-l afișăm
                    continue
                status_text = avail
                tag = "available" if avail.startswith("Disponibil") else ""
            else:
                # afișare fără filtrul de date
                status_row = cursor.execute(
                    "SELECT status FROM locatii WHERE id=?", (loc_id,)
                ).fetchone()[0]
                status_text = status_row
                tag = ("available","reserved","rented")[
                    ["Disponibil","Rezervat","Închiriat"].index(status_row)
                ] if status_row in ("Disponibil","Rezervat","Închiriat") else ""

            zebra = "evenrow" if display_index % 2 == 0 else "oddrow"
            display_index += 1

            tree.insert(
                "", "end",
                iid=str(loc_id),
                values=(display_index, city, county, addr, typ, status_text, rate),
                tags=(tag, zebra)
            )


    def on_tree_select():
        # ascundem toate etichetele
        for w in (
            lbl_client_label, lbl_client_value,
            lbl_period_label, lbl_period_value,
            lbl_ratecard_label, lbl_ratecard_value,
            lbl_pret_vanz_label, lbl_pret_vanz_value,

 main
        ):
            w.pack_forget()

        sel = tree.selection()
        if not sel:
            btn_edit.config(state='disabled')

 main
            btn_rent.config(state='disabled')
            btn_delete.config(state='disabled')
            img_label.config(image="", text="")
            btn_download.config(state='disabled')
            return

        loc_id = int(sel[0])
        selected_id[0] = loc_id

        code, client, ds, de, ratecard, pret_vanz, pret_flot = cursor.execute(
            "SELECT code, client, data_start, data_end, ratecard, pret_vanzare, pret_flotant "
            "FROM locatii WHERE id=?", (loc_id,)
        ).fetchone()

 main

        # actualizare valori
        lbl_client_value.config(text=client or "-")
        lbl_period_value.config(text=f"{ds} → {de}" if ds and de else "-")
        lbl_ratecard_value.config(text=str(ratecard))
        lbl_pret_vanz_value.config(text=str(pret_vanz) if pret_vanz is not None else "-")
        lbl_pret_flot_value.config(text=str(pret_flot) if pret_flot is not None else "-")

        # preview
        img = make_preview(code)
        if img:
            img_label.image = img
            img_label.config(image=img, text="")
        else:
            img_label.config(image="", text="Fără preview")

        btn_download.config(state='normal' if get_schita_path(code) else 'disabled')

        status = tree.item(sel[0])['values'][5]
        if status == "Închiriat":
            lbl_client_label.pack(anchor="center", pady=2)
            lbl_client_value.pack(anchor="center", pady=2)
            lbl_period_label.pack(anchor="center", pady=2)
            lbl_period_value.pack(anchor="center", pady=2)

 main
            lbl_pret_vanz_label.pack(anchor="center", pady=2)
            lbl_pret_vanz_value.pack(anchor="center", pady=2)
        else:
            lbl_ratecard_label.pack(anchor="center", pady=2)
            lbl_ratecard_value.pack(anchor="center", pady=2)
            lbl_pret_vanz_label.pack(anchor="center", pady=2)
            lbl_pret_vanz_value.pack(anchor="center", pady=2)
            lbl_pret_flot_label.pack(anchor="center", pady=2)
            lbl_pret_flot_value.pack(anchor="center", pady=2)

        btn_edit.config(state='normal')

        if status in ('Disponibil', 'Rezervat'):
            btn_rent.config(text="Închiriază", state='normal',
                            command=lambda: open_rent_window(root, loc_id, load_locations))
        else:
            btn_rent.config(text="Eliberează", state='normal',
                            command=lambda: release_and_refresh())

        btn_delete.config(state='normal')

    def download_schita():
        code = cursor.execute("SELECT code FROM locatii WHERE id=?", (selected_id[0],)).fetchone()[0]
        path = get_schita_path(code)
        if not path:
            messagebox.showerror("Eroare", "Nu există schița.")
            return
        dst = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG","*.png")])
        if dst:
            shutil.copy(path, dst)

    def delete_location():
        if messagebox.askyesno("Confirmă", "Ștergerea este irevocabilă!"):
            conn.cursor().execute("DELETE FROM locatii WHERE id=?", (selected_id[0],))
            conn.commit()
            load_locations()

    def release_and_refresh():
        if not messagebox.askyesno("Confirmă", "Încheie mai devreme închirierea?"):
            return
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM rezervari WHERE loc_id=? AND ? BETWEEN data_start AND data_end",
            (selected_id[0], datetime.date.today().isoformat()),
        )
        cur.execute(
            "UPDATE locatii SET status='Disponibil', client=NULL, client_id=NULL, data_start=NULL, data_end=NULL WHERE id=?",
            (selected_id[0],),
        )
        conn.commit()
        update_statusuri_din_rezervari()
        load_locations()

    def check_alerts():
        # implementare alerte...
        pass

    # bind filtre
    combo_group.bind("<<ComboboxSelected>>", lambda e: load_locations())
    combo_status.bind("<<ComboboxSelected>>", lambda e: load_locations())
    search_var.trace_add("write", lambda *a: load_locations())
    filter_start.bind("<<DateEntrySelected>>", lambda e: (var_ignore.set(False), load_locations()))
    filter_end.bind("<<DateEntrySelected>>", lambda e: (var_ignore.set(False), load_locations()))

    # inițializare
    refresh_groups()
    load_locations()
    check_alerts()
    root.mainloop()

if __name__ == "__main__":
    start_app()

 main

