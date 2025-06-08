

import os
import shutil
import datetime
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog

try:
    from ttkbootstrap import Style
    from ttkbootstrap import style as _ttkstyle
except Exception:
    Style = None
from tkcalendar import DateEntry as _DateEntry, Calendar as _Calendar
from UI.date_picker import DatePicker

# Work around a compatibility issue between ``tkcalendar.DateEntry`` and
# ``ttkbootstrap``.  The style patches applied by ``ttkbootstrap`` call the
# widget ``configure`` method before ``tkcalendar.DateEntry`` has initialised
# its internal ``_calendar`` attribute which results in an ``AttributeError``
# at construction time.  We intercept the call and delegate to ``ttk.Entry``
# until the calendar component is available.
from tkinter import ttk

def _safe_dateentry_config(self, cnf=None, **kw):
    if not hasattr(self, "_calendar"):
        return ttk.Entry.configure(self, cnf, **kw)
    if cnf is None:
        cnf = {}
    return _orig_dateentry_config(self, cnf, **kw)

_orig_dateentry_config = _DateEntry.configure
_DateEntry.configure = _safe_dateentry_config
_DateEntry.config = _safe_dateentry_config

# ``ttkbootstrap`` triggers ``Calendar.configure`` before the widget is fully
# initialised which causes ``_properties`` to be undefined.  Apply the same
# protective wrapper used for ``DateEntry``.
_orig_calendar_config = _Calendar.configure

def _safe_calendar_config(self, cnf=None, **kw):
    if not hasattr(self, "_properties"):
        return ttk.Frame.configure(self, cnf, **kw)
    if cnf is None:
        cnf = {}
    return _orig_calendar_config(self, cnf, **kw)

_Calendar.configure = _safe_calendar_config
_Calendar.config = _safe_calendar_config

DateEntry = _DateEntry

if Style:
    _orig_update_style = _ttkstyle.Bootstyle.update_ttk_widget_style

    def _safe_update_ttk_widget_style(widget=None, style_string=None, **kwargs):
        try:
            return _orig_update_style(widget, style_string, **kwargs)
        except Exception:
            return style_string

    _ttkstyle.Bootstyle.update_ttk_widget_style = staticmethod(_safe_update_ttk_widget_style)

from db import (
    conn,
    cursor,
    update_statusuri_din_rezervari,
    get_location_cache,
    maybe_refresh_location_cache,
    get_location_by_id,
    refresh_location_cache,
)
from utils import make_preview, get_schita_path
from UI.dialogs import (
    open_detail_window,
    open_add_window,
    open_edit_window,
    open_edit_rent_window,
    open_rent_window,
    open_reserve_window,
    open_release_window,
    cancel_reservation,
    open_offer_window,
    export_available_excel,
    export_sales_report,
    export_vendor_report,
    open_clients_window,
    open_users_window,
    open_firme_window,
    open_manage_window,
)



def start_app(user, root=None):
    """Launch the main application window for *user*.

    An existing ``Tk`` instance can be supplied via ``root``.  This allows the
    application to reuse a window created earlier (for example by the login
    dialog) and avoids initializing multiple ``Tk`` instances which can lead to
    errors on some platforms.
    """

    if root is None:
        if Style:
            style = Style("superhero")
            root = style.master
        else:
            root = tk.Tk()
            style = ttk.Style(root)
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass
    else:
        if Style:
            style = Style("superhero")
        else:
            style = ttk.Style(root)
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass

    root.title("Gestionare Locații Publicitare")

    BASE_WIDTH = 1920
    BASE_HEIGHT = 1080

    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(family="Segoe UI", size=12)
    root.option_add("*Font", default_font)

    menu_font = tkfont.nametofont("TkMenuFont")
    menu_font.configure(family="Segoe UI", size=14)

    style.configure("TButton", padding=(8, 4), font=("Segoe UI", 12), relief="solid", borderwidth=1)
    style.configure("Treeview.Heading", font=("Segoe UI", 12, "bold"))
    style.configure("Treeview", rowheight=28, font=("Segoe UI", 11))

    def apply_scale(scale):
        root.tk.call("tk", "scaling", scale)
        default_font.configure(size=int(12 * scale))
        menu_font.configure(size=int(14 * scale))
        style.configure(
            "TButton",
            padding=(int(8 * scale), int(4 * scale)),
            font=("Segoe UI", int(12 * scale)),
            relief="solid",
            borderwidth=1,
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", int(12 * scale), "bold"),
        )
        style.configure(
            "Treeview",
            rowheight=max(int(28 * scale), 18),
            font=("Segoe UI", int(11 * scale)),
        )

    def _set_geometry():
        """Configure the window geometry for a full HD layout."""
        root.geometry(f"{BASE_WIDTH}x{BASE_HEIGHT}")
        apply_scale(1.0)

    root.minsize(800, 600)

    _set_geometry()
    try:
        root.state('zoomed')
    except tk.TclError:
        root.attributes('-zoomed', True)

    def _maximize():
        """Maximize the application window."""
        try:
            root.state('zoomed')
        except tk.TclError:
            root.attributes('-zoomed', True)

    def _normalize():
        """Restore window to the default size."""
        try:
            root.state('normal')
        except tk.TclError:
            try:
                root.attributes('-zoomed', False)
            except tk.TclError:
                pass
        _set_geometry()

    def _is_zoomed():
        """Return ``True`` if the window is maximized."""
        try:
            return root.state() == 'zoomed' or bool(root.attributes('-zoomed'))
        except tk.TclError:
            return root.state() == 'zoomed'

    def toggle_fullscreen(event=None):
        """Toggle maximized mode."""
        if not _is_zoomed():
            _maximize()
        else:
            _normalize()

    root.bind('<F11>', toggle_fullscreen)
    root.bind('<Escape>', lambda e: _normalize())


    root.eval('tk::PlaceWindow . center')

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
    filter_start = DatePicker(frm_top, width=12)
    filter_start.pack(side="left", padx=5)
    ttk.Label(frm_top, text="Până:").pack(side="left", padx=(10,5))
    filter_end = DatePicker(frm_top, width=12)
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
    tree.heading("Status", text="Status");  tree.column("Status", width=120, anchor="center")
    tree.heading("RateCard", text="RateCard", anchor="e")
    tree.column("RateCard", width=100, anchor="e")
    tree.pack(fill="both", expand=True, side="left")

    vsb = ttk.Scrollbar(frm_mid, orient="vertical", command=tree.yview)
    vsb.pack(side="left", fill="y")
    tree.configure(yscroll=vsb.set)


    tree.tag_configure("evenrow", background="#f7f7f7", foreground="black")
    tree.tag_configure("oddrow", background="#ffffff", foreground="black")
    tree.tag_configure("available", background="#e8ffe8", foreground="black")
    tree.tag_configure("reserved", background="#fff5cc", foreground="black")
    tree.tag_configure("rented", background="#ffe8e8", foreground="black")

    drag_select = {"start": None}

    def on_tree_press(event):
        iid = tree.identify_row(event.y)
        if iid:
            # Allow default behaviour for Ctrl/Shift clicks so multiple
            # rows can be selected using the standard shortcuts.
            if event.state & 0x5:  # Ctrl or Shift pressed
                drag_select["start"] = None
                return
            drag_select["start"] = iid
            tree.selection_set(iid)

    def on_tree_drag(event):
        if not drag_select.get("start"):
            return
        iid = tree.identify_row(event.y)
        if not iid:
            return
        children = list(tree.get_children())
        i0 = children.index(drag_select["start"])
        i1 = children.index(iid)
        if i0 > i1:
            i0, i1 = i1, i0
        tree.selection_set(children[i0 : i1 + 1])
        tree.see(iid)

    tree.bind("<Button-1>", on_tree_press)
    tree.bind("<B1-Motion>", on_tree_drag)
    tree.bind("<ButtonRelease-1>", lambda e: drag_select.update(start=None))

    tree.bind("<Double-1>", lambda e: open_detail_window(tree, e))
    tree.bind("<<TreeviewSelect>>", lambda e: on_tree_select())

    # --- Panoul detalii (dreapta) ---
    details = ttk.Frame(frm_mid, padding=10, relief="groove", width=400)
    details.pack(side="right", fill="y")
    details.pack_propagate(False)

    img_label = ttk.Label(details)
    img_label.pack(pady=(0,10))

    btn_download = ttk.Button(
        details,
        text="Descarcă schița",
        state="disabled",
        command=lambda: download_schita(),
    )
    btn_download.pack(pady=(0, 15))

    # etichete declarate fără pack()
    lbl_client_label       = ttk.Label(details, text="Client:")
    lbl_client_value       = ttk.Label(details, text="-")
    lbl_period_label       = ttk.Label(details, text="Perioadă:")
    lbl_period_value       = ttk.Label(details, text="-")
    lbl_ratecard_label     = ttk.Label(details, text="Rate Card:")
    lbl_ratecard_value     = ttk.Label(details, text="-")
    lbl_pret_vanz_label    = ttk.Label(details, text="Preț de vânzare aprobat:")
    lbl_pret_vanz_value    = ttk.Label(details, text="-")
    lbl_pret_inch_label    = ttk.Label(details, text="Preț de închiriere:")
    lbl_pret_inch_value    = ttk.Label(details, text="-")
    lbl_pret_flot_label    = ttk.Label(details, text="Preț flotant:")
    lbl_pret_flot_value    = ttk.Label(details, text="-")
    lbl_res_by_label       = ttk.Label(details, text="Rezervată de:")
    lbl_res_by_value       = ttk.Label(details, text="-")
    lbl_next_rent_label    = ttk.Label(details, text="Următoarea închiriere:")
    lbl_next_rent_value    = ttk.Label(details, text="-")


    # --- Bottom: butoane principale (stânga) și export (dreapta) ---
    frm_bot = ttk.Frame(root, padding=10)
    frm_bot.pack(fill="x", side="bottom")
    frm_bot.columnconfigure(0, weight=1)
    frm_bot.columnconfigure(1, weight=1)

    primary_frame = ttk.Frame(frm_bot)
    primary_frame.grid(row=0, column=0, sticky="w")
    btn_add     = ttk.Button(primary_frame, text="Adaugă",
                             command=lambda: open_add_window(root, load_locations))
    btn_edit    = ttk.Button(primary_frame, text="Editează", state="disabled",
                             command=lambda: open_edit_window(root, selected_id[0], load_locations, refresh_groups))

    btn_rent    = ttk.Button(primary_frame, text="Închiriază", state="disabled")
    btn_release = ttk.Button(primary_frame, text="Eliberează", state="disabled")
    btn_extend  = ttk.Button(primary_frame, text="Extinde perioada", state="disabled")
    btn_reserve = ttk.Button(primary_frame, text="Rezervă", state="disabled")
    btn_delete  = ttk.Button(primary_frame, text="Șterge", state="disabled",
                             command=lambda: delete_location())
    btn_clients = ttk.Button(primary_frame, text="Clienți",
                             command=lambda: open_clients_window(root))
    btn_firme = ttk.Button(primary_frame, text="Firme",
                            command=lambda: open_firme_window(root))
    btn_users = ttk.Button(primary_frame, text="Utilizatori",
                           command=lambda: open_users_window(root))
    btn_manage = ttk.Button(primary_frame, text="Administrează",
                            command=lambda: open_manage_window(root))
    role = user.get("role")
    if role in ("admin", "manager"):
        for w in (btn_add, btn_edit, btn_delete):
            w.pack(side="left", padx=5, pady=5)
        if role == "admin":
            btn_manage.pack(side="left", padx=5, pady=5)
    for w in (btn_clients,):
        w.pack(side="left", padx=5, pady=5)
    if role != "admin":
        btn_firme.pack(side="left", padx=5, pady=5)
    if role != "manager":
        for w in (btn_rent, btn_release, btn_reserve):
            w.pack(side="left", padx=5, pady=5)


    export_frame = ttk.Frame(frm_bot)
    export_frame.grid(row=0, column=1, sticky="e")
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
    btn_vendor = ttk.Button(export_frame, text="Raport Vânzători",
                           command=lambda: export_vendor_report())
    btn_update = ttk.Button(export_frame, text="Update Database",
                           command=lambda: manual_refresh())
    btn_xlsx.pack(side="left", padx=5, pady=5)
    btn_offer.pack(side="left", padx=5, pady=5)
    btn_report.pack(side="left", padx=5, pady=5)
    if role in ("admin", "manager"):
        btn_vendor.pack(side="left", padx=5, pady=5)
    btn_update.pack(side="left", padx=5, pady=5)



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
        items = tree.get_children()
        if items:
            tree.delete(*items)

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

        # 5) Pornim de la versiunea în memorie a tabelului ``locatii``
        rows = get_location_cache()

        if cond:
            def match(row: dict) -> bool:
                for c, p in zip(cond, params):
                    if c == "grup = ?" and row.get("grup") != p:
                        return False
                    if c == "status = ?" and row.get("status") != p:
                        return False
                    if c.startswith("(city LIKE"):
                        term_lower = p.strip("%").lower()
                        if not (term_lower in (row.get("city") or "").lower() or
                                term_lower in (row.get("county") or "").lower() or
                                term_lower in (row.get("address") or "").lower()):
                            return False
                return True
            rows = [r for r in rows if match(r)]

        rows.sort(key=lambda r: (
            {
                'Bucuresti Sectorul 1': 1,
                'Bucuresti Sectorul 2': 2,
                'Bucuresti Sectorul 3': 3,
                'Bucuresti Sectorul 4': 4,
                'Bucuresti Sectorul 5': 5,
                'Bucuresti Sectorul 6': 6,
                'Ilfov': 7,
                'Prahova': 8,
            }.get(r.get('county'), 9),
            r.get('county'),
            r.get('city'),
        ))

        # 6) Preîncărcăm rezervările pentru perioada selectată
        reservations_by_loc: dict[int, list[tuple[datetime.date, datetime.date]]] = {}
        if not var_ignore.get():
            rows_res = cursor.execute(
                "SELECT loc_id, data_start, data_end FROM rezervari WHERE NOT (data_end < ? OR data_start > ?) ORDER BY data_start",
                (start_dt.isoformat(), end_dt.isoformat()),
            ).fetchall()
            for loc_id_r, ds, de in rows_res:
                reservations_by_loc.setdefault(loc_id_r, []).append(
                    (datetime.date.fromisoformat(ds), datetime.date.fromisoformat(de))
                )

        def availability(loc_id):
            periods = reservations_by_loc.get(loc_id, [])
            if not periods:
                return "Disponibil"
            overl = [(ds, de) for ds, de in periods if not (de < start_dt or ds > end_dt)]
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
            return ""


        # 7) Populează TreeView, aplicând filtrul de date doar când "Toate datele" NU e bifat
        display_index = 0
        for row in rows:
            if row.get("parent_id") and row.get("status") == "Expirat":
                continue
            loc_id = row["id"]
            city = row["city"]
            county = row["county"]
            addr = row["address"]
            typ = row["type"]
            rate = row["ratecard"]
            status = row["status"]
            if not var_ignore.get():
                avail = availability(loc_id)
                if not avail:
                    # nu se intersectează cu intervalul, deci nu-l afișăm
                    continue
                status_text = avail
                tag = "available" if avail.startswith("Disponibil") else ""
            else:
                # afișare fără filtrul de date
                status_text = status
                tag = ("available","reserved","rented")[
                    ["Disponibil","Rezervat","Închiriat"].index(status)
                ] if status in ("Disponibil","Rezervat","Închiriat") else ""

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
            lbl_pret_inch_label, lbl_pret_inch_value,
            lbl_pret_flot_label, lbl_pret_flot_value,
            lbl_res_by_label, lbl_res_by_value,
            lbl_next_rent_label, lbl_next_rent_value,

        ):
            w.pack_forget()

        sel = tree.selection()
        if not sel:
            btn_edit.config(state='disabled')
            btn_rent.config(state='disabled')
            btn_release.config(state='disabled')
            btn_delete.config(state='disabled')
            img_label.config(image="", text="")
            btn_download.config(state='disabled')
            return

        loc_id = int(sel[0])
        selected_id[0] = loc_id

        data = get_location_by_id(loc_id)
        if not data:
            return
        code = data.get("code")
        client = data.get("client")
        ds = data.get("data_start")
        de = data.get("data_end")
        ratecard = data.get("ratecard")
        pret_vanz = data.get("pret_vanzare")
        pret_flot = data.get("pret_flotant")

        rent_row = None
        if ds and de:
            rent_row = cursor.execute(
                "SELECT suma FROM rezervari WHERE loc_id=? AND data_start=? AND data_end=? ORDER BY id DESC LIMIT 1",
                (loc_id, ds, de)
            ).fetchone()
        rent_price = rent_row[0] if rent_row else None


        # actualizare valori
        lbl_client_value.config(text=client or "-")
        lbl_period_value.config(text=f"{ds} → {de}" if ds and de else "-")
        lbl_ratecard_value.config(text=str(ratecard))
        lbl_pret_vanz_value.config(text=str(pret_vanz) if pret_vanz is not None else "-")
        lbl_pret_inch_value.config(text=str(rent_price) if rent_price is not None else "-")
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
        reserved_info = None
        if status == "Rezervat":
            today = datetime.date.today().isoformat()
            reserved_info = cursor.execute(
                "SELECT created_by, data_end FROM rezervari "
                "WHERE loc_id=? AND ? BETWEEN data_start AND data_end "
                "AND suma IS NULL ORDER BY data_start DESC LIMIT 1",
                (loc_id, today),
            ).fetchone()

        if status == "Închiriat":
            lbl_client_label.pack(anchor="center", pady=2)
            lbl_client_value.pack(anchor="center", pady=2)
            lbl_period_label.pack(anchor="center", pady=2)
            lbl_period_value.pack(anchor="center", pady=2)

            lbl_pret_inch_label.pack(anchor="center", pady=2)
            lbl_pret_inch_value.pack(anchor="center", pady=2)
            lbl_pret_vanz_label.pack(anchor="center", pady=2)
            lbl_pret_vanz_value.pack(anchor="center", pady=2)
        else:
            lbl_ratecard_label.pack(anchor="center", pady=2)
            lbl_ratecard_value.pack(anchor="center", pady=2)
            lbl_pret_vanz_label.pack(anchor="center", pady=2)
            lbl_pret_vanz_value.pack(anchor="center", pady=2)
            lbl_pret_flot_label.pack(anchor="center", pady=2)
            lbl_pret_flot_value.pack(anchor="center", pady=2)
            if reserved_info:
                creator, end_d = reserved_info
                days_left = (datetime.date.fromisoformat(end_d) -
                             datetime.date.today()).days + 1
                lbl_res_by_value.config(
                    text=f"{creator} ({days_left} zile)")
                lbl_res_by_label.pack(anchor="center", pady=2)
                lbl_res_by_value.pack(anchor="center", pady=2)
            else:
                today = datetime.date.today().isoformat()
                next_rent = cursor.execute(
                    "SELECT client, data_start, data_end FROM rezervari "
                    "WHERE loc_id=? AND data_start>? AND suma IS NOT NULL "
                    "ORDER BY data_start LIMIT 1",
                    (loc_id, today),
                ).fetchone()
                if next_rent:
                    n_client, n_start, n_end = next_rent
                    lbl_next_rent_value.config(
                        text=f"{n_client}: {n_start} → {n_end}"
                    )
                    lbl_next_rent_label.pack(anchor="center", pady=2)
                    lbl_next_rent_value.pack(anchor="center", pady=2)

        if role in ("admin", "manager"):
            btn_edit.config(state='normal')
        
        # Butoane închiriere și eliberare
        if role != "manager":
            btn_rent.config(
                text="Închiriază",
                state='normal',
                command=lambda: open_rent_window(root, loc_id, load_locations, user),
            )
        else:
            btn_rent.config(state='disabled', command=lambda: None)

        has_rentals = cursor.execute(
            "SELECT 1 FROM rezervari WHERE loc_id=? AND suma IS NOT NULL LIMIT 1",
            (loc_id,),
        ).fetchone()
        if role != "manager":
            if has_rentals:
                btn_release.config(
                    state="normal",
                    command=lambda: open_release_window(
                        root, loc_id, load_locations, user
                    ),
                )
            else:
                btn_release.config(state="disabled", command=lambda: None)
        else:
            btn_release.config(state="disabled", command=lambda: None)

        if role != "manager":
            if data.get("is_mobile") and data.get("parent_id"):
                if not btn_extend.winfo_ismapped():
                    btn_extend.pack(side="left", padx=5, pady=5)
                rid_row = cursor.execute(
                    "SELECT id FROM rezervari WHERE loc_id=? AND ? BETWEEN data_start AND data_end AND suma IS NOT NULL ORDER BY id DESC LIMIT 1",
                    (loc_id, datetime.date.today().isoformat()),
                ).fetchone()
                if rid_row:
                    rid = rid_row[0]
                    btn_extend.config(
                        state="normal",
                        command=lambda r=rid, ds=data.get("data_start"), de=data.get("data_end"), pid=data.get("parent_id"): open_edit_rent_window(root, r, load_locations, parent=(pid, ds, de)),
                    )
                else:
                    btn_extend.config(state="disabled", command=lambda: None)
            else:
                if btn_extend.winfo_ismapped():
                    btn_extend.pack_forget()
                btn_extend.config(state="disabled", command=lambda: None)
        else:
            if btn_extend.winfo_ismapped():
                btn_extend.pack_forget()
            btn_extend.config(state="disabled", command=lambda: None)

        if role != "manager":
            if status == "Disponibil" and not (
                data.get("is_mobile") and not data.get("parent_id")
            ):
                btn_reserve.config(
                    text="Rezervă",
                    state='normal',
                    command=lambda: open_reserve_window(root, loc_id, load_locations, user)
                )
            elif status == "Rezervat" and reserved_info and reserved_info[0] == user["username"]:
                btn_reserve.config(
                    text="Anulează rezervarea",
                    state='normal',
                    command=lambda: cancel_reservation(root, loc_id, load_locations)
                )
            else:
                btn_reserve.config(text="Rezervă", state='disabled')
        else:
            btn_reserve.config(text="Rezervă", state='disabled')


        if role in ("admin", "manager"):
            btn_delete.config(state='normal')
        else:
            btn_delete.config(state='disabled')

    def download_schita():
        data = get_location_by_id(selected_id[0])
        if not data:
            messagebox.showerror("Eroare", "Nu există date pentru locație.")
            return
        code = data.get("code")
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

    def manual_refresh():
        """Force reload of the local cache from the database."""
        refresh_location_cache()
        load_locations()

    def check_alerts():
        """Placeholder for future alert functionality."""
        print("check_alerts: not yet implemented")

    def watch_updates():
        if maybe_refresh_location_cache():
            load_locations()
        root.after(300000, watch_updates)  # 5 minute

    # bind filtre
    combo_group.bind("<<ComboboxSelected>>", lambda e: load_locations())
    combo_status.bind("<<ComboboxSelected>>", lambda e: load_locations())

    # small delay for search updates for smoother typing
    _search_after = [None]

    def on_search_change(*args):
        if _search_after[0] is not None:
            root.after_cancel(_search_after[0])
        _search_after[0] = root.after(300, load_locations)

    search_var.trace_add("write", on_search_change)
    filter_start.bind("<<DateEntrySelected>>", lambda e: (var_ignore.set(False), load_locations()))
    filter_end.bind("<<DateEntrySelected>>", lambda e: (var_ignore.set(False), load_locations()))

    # inițializare
    refresh_groups()
    load_locations()
    check_alerts()
    watch_updates()
    root.mainloop()

if __name__ == "__main__":
    import db
    u = db.get_user("admin")
    start_app(u)


