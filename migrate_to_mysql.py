import os
import sqlite3
from db import get_db_path

try:
    from dotenv import load_dotenv  # type: ignore
except Exception:
    load_dotenv = lambda *a, **k: None

load_dotenv()

try:
    import mysql.connector  # type: ignore
except Exception:
    raise SystemExit("mysql-connector-python is required for migration")


def connect_mysql():
    try:
        return mysql.connector.connect(
            host=os.environ.get("MYSQL_HOST", "localhost"),
            port=int(os.environ.get("MYSQL_PORT", 3306)),
            user=os.environ.get("MYSQL_USER", "root"),
            password=os.environ.get("MYSQL_PASSWORD", ""),
            database=os.environ.get("MYSQL_DATABASE", "focus_media"),
        )
    except mysql.connector.Error as exc:
        raise SystemExit(
            "Failed to connect to MySQL. Check environment variables "
            "MYSQL_HOST, MYSQL_USER, MYSQL_PASSWORD and MYSQL_DATABASE.\n" +
            str(exc)
        )


def create_tables(cur):
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS locatii (
            id INT AUTO_INCREMENT PRIMARY KEY,
            city TEXT,
            county TEXT,
            address TEXT,
            type TEXT,
            gps TEXT,
            code TEXT,
            size TEXT,
            photo_link TEXT,
            sqm DOUBLE,
            illumination TEXT,
            ratecard DOUBLE,
            pret_vanzare DOUBLE,
            pret_flotant DOUBLE,
            decoration_cost DOUBLE,
            observatii TEXT,
            status TEXT DEFAULT 'Disponibil',
            client TEXT,
            client_id INT,
            data_start TEXT,
            data_end TEXT,
            grup TEXT,
            face TEXT DEFAULT 'Fața A'
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS clienti (
            id INT AUTO_INCREMENT PRIMARY KEY,
            nume TEXT UNIQUE NOT NULL,
            contact TEXT,
            email TEXT,
            phone TEXT,
            observatii TEXT,
            tip TEXT DEFAULT 'direct'
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS rezervari (
            id INT AUTO_INCREMENT PRIMARY KEY,
            loc_id INT NOT NULL,
            client TEXT NOT NULL,
            client_id INT,
            data_start TEXT NOT NULL,
            data_end TEXT NOT NULL,
            suma DOUBLE,
            created_by TEXT,
            FOREIGN KEY(loc_id) REFERENCES locatii(id),
            FOREIGN KEY(client_id) REFERENCES clienti(id)
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INT AUTO_INCREMENT PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            comune TEXT
        )
        """
    )


def copy_table(src_cur, dst_cur, table):
    try:
        columns = [row[1] for row in src_cur.execute(f"PRAGMA table_info({table})")]
    except sqlite3.OperationalError:
        return
    col_list = ", ".join(columns)
    placeholders = ", ".join(["%s"] * len(columns))
    rows = src_cur.execute(f"SELECT {col_list} FROM {table}").fetchall()
    if rows:
        dst_cur.executemany(
            f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})",
            rows,
        )


def main():
    sqlite_conn = sqlite3.connect(get_db_path())
    sqlite_cur = sqlite_conn.cursor()

    mysql_conn = connect_mysql()
    mysql_cur = mysql_conn.cursor()

    create_tables(mysql_cur)
    mysql_conn.commit()

    for table in ["locatii", "clienti", "rezervari", "users"]:
        copy_table(sqlite_cur, mysql_cur, table)

    mysql_conn.commit()
    mysql_conn.close()
    sqlite_conn.close()

    print("Migrated data from SQLite to MySQL database.")


if __name__ == "__main__":
    main()
