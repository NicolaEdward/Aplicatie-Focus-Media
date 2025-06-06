import os
import sqlite3
import datetime


def get_db_path():
    """Return the absolute path to the local ``locatii.db`` database."""
    base_dir = os.path.dirname(__file__)
    return os.path.join(base_dir, "locatii.db")

conn   = sqlite3.connect(get_db_path())
cursor = conn.cursor()

def init_db():
    # (1) Creăm tabelul cu toate coloanele, inclusiv noile preturi
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS locatii (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        city TEXT,
        county TEXT,
        address TEXT,
        type TEXT,
        gps TEXT,
        code TEXT,
        size TEXT,
        photo_link TEXT,
        sqm REAL,
        illumination TEXT,
        ratecard REAL,
        pret_vanzare REAL,
        pret_flotant REAL,
        decoration_cost REAL,
        observatii TEXT,
        status TEXT DEFAULT 'Disponibil',
        client TEXT,
        client_id INTEGER,
        data_start TEXT,
        data_end TEXT,
        grup TEXT,
        face TEXT DEFAULT 'Fața A'
    )
    """)
    init_clienti_table()
    init_rezervari_table()
    conn.commit()
    
    # (2) Verificăm ce coloane există deja
    existing = {col[1] for col in cursor.execute("PRAGMA table_info(locatii)").fetchall()}

    # (3) Pentru fiecare coloană nouă, adăugăm dacă lipsește
    to_add = {
        "pret_vanzare": "REAL",
        "pret_flotant": "REAL",
        "client_id":    "INTEGER",
        "face":         "TEXT DEFAULT 'Fața A'"
    }

    for col, definition in to_add.items():
        if col not in existing:
            cursor.execute(f"ALTER TABLE locatii ADD COLUMN {col} {definition}")
            conn.commit()

def init_clienti_table():
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS clienti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nume TEXT UNIQUE NOT NULL
        )
    """)
    conn.commit()

def init_rezervari_table():
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS rezervari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        loc_id INTEGER NOT NULL,
        client TEXT NOT NULL,
        client_id INTEGER,
        data_start TEXT NOT NULL,
        data_end TEXT NOT NULL,
        suma REAL,
        FOREIGN KEY(loc_id) REFERENCES locatii(id),
        FOREIGN KEY(client_id) REFERENCES clienti(id)
    )
    """)
    # Adăugare coloana client_id dacă lipsește (migrare simplă)
    cols = {c[1] for c in cursor.execute("PRAGMA table_info(rezervari)").fetchall()}
    if "client_id" not in cols:
        cursor.execute("ALTER TABLE rezervari ADD COLUMN client_id INTEGER")
        conn.commit()
    conn.commit()
def update_statusuri_din_rezervari():
    today = datetime.date.today().isoformat()
    cur = conn.cursor()

    # 1) Resetăm totul la Disponibil
    cur.execute("""
        UPDATE locatii
        SET status='Disponibil',
            client=NULL,
            client_id=NULL,
            data_start=NULL,
            data_end=NULL
    """)


    # 2) Nu mai marcăm rezervările viitoare ca "Rezervat" pentru a permite
    #     programarea mai multor închirieri. Locațiile rămân "Disponibil" până
    #     începe perioada efectivă de închiriere.

    # 3) Marcăm închirierile curente ca 'Închiriat'
    cur.execute("""
        UPDATE locatii
        SET status      = 'Închiriat',
            client      = (
                SELECT client
                FROM rezervari
                WHERE rezervari.loc_id = locatii.id
                  AND ? BETWEEN data_start AND data_end
                ORDER BY data_start DESC
                LIMIT 1
            ),
            client_id   = (
                SELECT client_id
                FROM rezervari
                WHERE rezervari.loc_id = locatii.id
                  AND ? BETWEEN data_start AND data_end
                ORDER BY data_start DESC
                LIMIT 1
            ),
            data_start  = (
                SELECT data_start
                FROM rezervari
                WHERE rezervari.loc_id = locatii.id
                  AND ? BETWEEN data_start AND data_end
                ORDER BY data_start DESC
                LIMIT 1
            ),
            data_end    = (
                SELECT data_end
                FROM rezervari
                WHERE rezervari.loc_id = locatii.id
                  AND ? BETWEEN data_start AND data_end
                ORDER BY data_start DESC
                LIMIT 1
            )
        WHERE EXISTS (
            SELECT 1
            FROM rezervari
            WHERE rezervari.loc_id = locatii.id
              AND ? BETWEEN data_start AND data_end
        )
    """, (today, today, today, today, today))

    conn.commit()

# Initialize DB on import
init_db()
