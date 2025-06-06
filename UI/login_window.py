import tkinter as tk
from tkinter import ttk, messagebox
import db

def show_login():
    root = tk.Tk()
    root.title("Login")

    ttk.Label(root, text="Username:").grid(row=0, column=0, padx=5, pady=5)
    user_entry = ttk.Entry(root, width=30)
    user_entry.grid(row=0, column=1, padx=5, pady=5)

    ttk.Label(root, text="Password:").grid(row=1, column=0, padx=5, pady=5)
    pass_entry = ttk.Entry(root, show="*", width=30)
    pass_entry.grid(row=1, column=1, padx=5, pady=5)

    result = {"user": None}

    def attempt():
        u = user_entry.get().strip()
        p = pass_entry.get().strip()
        user = db.check_login(u, p)
        if user:
            result["user"] = user
            root.destroy()
        else:
            messagebox.showerror("Login", "Credențiale invalide")

    ttk.Button(root, text="Login", command=attempt).grid(row=2, column=0, columnspan=2, pady=10)
    root.mainloop()
    return result["user"]

if __name__ == "__main__":
    print(show_login())
