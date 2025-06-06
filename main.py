# main.py
from UI.main_window import start_app
from UI.login_window import show_login

if __name__ == "__main__":
    user = show_login()
    if user:
        start_app(user)
