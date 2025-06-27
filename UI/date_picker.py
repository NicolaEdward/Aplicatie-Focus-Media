import datetime
import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar

class DatePicker(ttk.Frame):
    """Entry with a popup calendar for selecting a date."""

    def __init__(self, master=None, date=None, width=12, **kwargs):
        super().__init__(master, **kwargs)
        self._var = tk.StringVar()
        self.entry = ttk.Entry(self, textvariable=self._var, width=width)
        self.entry.pack(side="left")
        ttk.Button(self, text="📅", width=2, command=self._open).pack(side="left")
        if date is None:
            date = datetime.date.today()
        self.set_date(date)
        self._top = None

    def _open(self):
        if self._top and self._top.winfo_exists():
            self._top.lift()
            return
        self._top = tk.Toplevel(self)
        self._top.transient(self)
        # Use a larger font for better visibility
        cal = Calendar(
            self._top,
            selectmode="day",
            date_pattern="yyyy-mm-dd",
            font=("Segoe UI", 14),
        )
        cal.pack(fill="both", expand=True)

        def _select(*_):
            self.set_date(datetime.date.fromisoformat(cal.get_date()))
            self._top.destroy()
            self.event_generate("<<DateEntrySelected>>")

        cal.bind("<<CalendarSelected>>", _select)

    def get_date(self):
        try:
            return datetime.date.fromisoformat(self._var.get())
        except ValueError:
            return datetime.date.today()

    def set_date(self, date: datetime.date):
        self._var.set(date.isoformat())
