import sqlite3, os

def get_db_connection():
    base = os.path.dirname(os.path.dirname(__file__))  # coboară din UI/ în proiect
    path = os.path.join(base, "locatii.db")
    return sqlite3.connect(path)
