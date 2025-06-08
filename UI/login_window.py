import tkinter as tk
from tkinter import ttk, messagebox
import db


def show_login(root=None):
    """Display a login dialog and return the authenticated user object.

    If *root* is provided, the dialog is opened as a child window and the
    function will return once the dialog is closed.  Otherwise a temporary
    ``Tk`` instance is created and destroyed when finished, preserving the
    previous behaviour.
    """

    owns_root = root is None
    if owns_root:
        root = tk.Tk()
        win = root
    else:
        win = tk.Toplevel(root)

    win.title("Login")

    ttk.Label(win, text="Username:").grid(row=0, column=0, padx=5, pady=5)
    user_entry = ttk.Entry(win, width=30)
    user_entry.grid(row=0, column=1, padx=5, pady=5)

    ttk.Label(win, text="Password:").grid(row=1, column=0, padx=5, pady=5)
    pass_entry = ttk.Entry(win, show="*", width=30)
    pass_entry.grid(row=1, column=1, padx=5, pady=5)

    result = {"user": None}

    def attempt():
        u = user_entry.get().strip()
        p = pass_entry.get().strip()
        user = db.check_login(u, p)
        if user:
            result["user"] = user
            win.destroy()
        else:
            messagebox.showerror("Login", "Credențiale invalide")

    ttk.Button(win, text="Login", command=attempt).grid(
        row=2, column=0, columnspan=2, pady=10
    )

    if owns_root:
        win.mainloop()
    else:
        root.wait_window(win)

    if owns_root and root.winfo_exists():
        root.destroy()

    return result["user"]

if __name__ == "__main__":
    print(show_login())
