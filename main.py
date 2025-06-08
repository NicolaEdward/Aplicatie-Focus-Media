# main.py
import tkinter as tk
from UI.main_window import start_app
from UI.login_window import show_login

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    user = show_login(root)
    if user:
        root.deiconify()
        start_app(user, root)
