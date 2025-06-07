import os
import datetime
import hashlib
import hmac
import sqlite3

try:
    from dotenv import load_dotenv  # type: ignore
except Exception:  # pragma: no cover - optional dep
    load_dotenv = lambda *a, **k: None

load_dotenv()

try:
    import mysql.connector  # type: ignore
except Exception:  # pragma: no cover - optional dep
    mysql = None


class _CursorWrapper:
    """Cursor wrapper translating ``?`` placeholders for MySQL."""

    def __init__(self, cur, mysql_mode: bool):
        self._cur = cur
        self._mysql = mysql_mode

    def execute(self, sql, params=None):
        if self._mysql:
            sql = sql.replace("?", "%s")
        return self._cur.execute(sql, params or ())

    def executemany(self, sql, params):
        if self._mysql:
            sql = sql.replace("?", "%s")
        return self._cur.executemany(sql, params)

    def __getattr__(self, name):
        return getattr(self._cur, name)


class _ConnWrapper:
    def __init__(self, conn, mysql_mode: bool):
        self._conn = conn
        self._mysql = mysql_mode

    def cursor(self):
        return _CursorWrapper(self._conn.cursor(), self._mysql)

    @property
    def mysql(self):
        return self._mysql

    def __getattr__(self, name):
        return getattr(self._conn, name)


def get_db_path():
    """Return the absolute path to the local ``locatii.db`` database."""
    base_dir = os.path.dirname(__file__)
    return os.path.join(base_dir, "locatii.db")


def _create_connection():
    host = os.environ.get("MYSQL_HOST")
    if host and mysql is not None:
        conn = mysql.connector.connect(
            host=host,
            port=int(os.environ.get("MYSQL_PORT", 3306)),
            user=os.environ.get("MYSQL_USER", "root"),
            password=os.environ.get("MYSQL_PASSWORD", ""),
            database=os.environ.get("MYSQL_DATABASE", "focus_media"),
        )
        return _ConnWrapper(conn, True)
    return _ConnWrapper(sqlite3.connect(get_db_path()), False)


conn = _create_connection()
cursor = conn.cursor()


def ensure_index(table: str, index_name: str, column: str) -> None:
    """Create *index_name* on *table* if it is missing."""
    if getattr(conn, "mysql", False):
        cur = conn.cursor()
        cur.execute(f"SHOW INDEX FROM {table} WHERE Key_name=?", (index_name,))
        if not cur.fetchone():
            cur.execute(f"CREATE INDEX {index_name} ON {table}({column})")
    else:
        cursor.execute(
            f"CREATE INDEX IF NOT EXISTS {index_name} ON {table}({column})"
        )

def init_db():
    if getattr(conn, "mysql", False):
        cursor.execute(
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
                status VARCHAR(32) DEFAULT 'Disponibil',
                client TEXT,
                client_id INT,
                data_start TEXT,
                data_end TEXT,
                grup VARCHAR(255),
                face VARCHAR(32) DEFAULT 'Fața A'
            )
            """
        )
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS clienti (
                id INT AUTO_INCREMENT PRIMARY KEY,
                nume VARCHAR(255) UNIQUE NOT NULL,
                contact TEXT,
                email TEXT,
                phone TEXT,
                observatii TEXT,
                tip VARCHAR(32) DEFAULT 'direct'
            )
            """
        )
        cursor.execute(
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
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INT AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(255) UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL,
                comune TEXT
            )
            """
        )
    else:
        # (1) Creăm tabelul cu toate coloanele, inclusiv noile preturi
        cursor.execute(
            """
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
        """
        )
        init_clienti_table()
        init_rezervari_table()
        init_users_table()
        conn.commit()

    # Indexuri pentru o interogare mai rapidă
    ensure_index("locatii", "idx_locatii_grup", "grup")
    ensure_index("locatii", "idx_locatii_status", "status")
    ensure_index("rezervari", "idx_rezervari_loc", "loc_id")
    conn.commit()

    if not getattr(conn, "mysql", False):
        existing = {col[1] for col in cursor.execute("PRAGMA table_info(locatii)").fetchall()}

        to_add = {
            "rental_fee":    "REAL DEFAULT 0",
            "pret_vanzare":  "REAL",
            "pret_flotant":  "REAL",
            "client_id":    "INTEGER",
        }

        existing = {col[1] for col in cursor.execute("PRAGMA table_info(locatii)").fetchall()}
        if "face" not in existing:
            cursor.execute("ALTER TABLE locatii ADD COLUMN face TEXT DEFAULT 'Fața A'")
            conn.commit()

        for col, definition in to_add.items():
            if col not in existing:
                cursor.execute(f"ALTER TABLE locatii ADD COLUMN {col} {definition}")
                conn.commit()

def init_clienti_table():
    if getattr(conn, "mysql", False):
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS clienti (
                id INT AUTO_INCREMENT PRIMARY KEY,
                nume VARCHAR(255) UNIQUE NOT NULL,
                contact TEXT,
                email TEXT,
                phone TEXT,
                observatii TEXT,
                tip TEXT DEFAULT 'direct'
            )
            """
        )
    else:
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS clienti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nume TEXT UNIQUE NOT NULL,
            contact TEXT,
            email TEXT,
            phone TEXT,
            observatii TEXT,
            tip TEXT DEFAULT 'direct'
        )
        """
        )
    if not getattr(conn, "mysql", False):
        cols = {c[1] for c in cursor.execute("PRAGMA table_info(clienti)").fetchall()}
        to_add = {
            "contact": "TEXT",
            "email": "TEXT",
            "phone": "TEXT",
            "observatii": "TEXT",
            "tip": "TEXT DEFAULT 'direct'",
        }
        for col, definition in to_add.items():
            if col not in cols:
                cursor.execute(f"ALTER TABLE clienti ADD COLUMN {col} {definition}")
                conn.commit()
        conn.commit()

def init_rezervari_table():
    if getattr(conn, "mysql", False):
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS rezervari (
                id INT AUTO_INCREMENT PRIMARY KEY,
                loc_id INT NOT NULL,
                client TEXT NOT NULL,
                client_id INT,
                data_start TEXT NOT NULL,
                data_end TEXT NOT NULL,
                suma DOUBLE,
                FOREIGN KEY(loc_id) REFERENCES locatii(id),
                FOREIGN KEY(client_id) REFERENCES clienti(id)
            )
            """
        )
    else:
        cursor.execute(
            """
    CREATE TABLE IF NOT EXISTS rezervari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        loc_id INTEGER NOT NULL,
        client TEXT NOT NULL,
        client_id INTEGER,
        data_start TEXT NOT NULL,
        data_end TEXT NOT NULL,
        suma REAL,
        created_by TEXT,
        FOREIGN KEY(loc_id) REFERENCES locatii(id),
        FOREIGN KEY(client_id) REFERENCES clienti(id)
    )
    """
        )
        cols = {c[1] for c in cursor.execute("PRAGMA table_info(rezervari)").fetchall()}
        if "client_id" not in cols:
            cursor.execute("ALTER TABLE rezervari ADD COLUMN client_id INTEGER")
            conn.commit()
        if "created_by" not in cols:
            cursor.execute("ALTER TABLE rezervari ADD COLUMN created_by TEXT")
            conn.commit()
        conn.commit()

def init_users_table():
    if getattr(conn, "mysql", False):
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INT AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(255) UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL,
                comune TEXT
            )
            """
        )
    else:
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            comune TEXT
        )
        """
        )
    # create default admin if table empty
    if not cursor.execute("SELECT COUNT(*) FROM users").fetchone()[0]:
        cursor.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, 'admin')",
            ("admin", _hash_password("admin")),
        )
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


def _hash_password(pw: str, *, _salt: bytes | None = None) -> str:
    """Return a salted PBKDF2 hash of *pw* suitable for storage."""
    if _salt is None:
        _salt = os.urandom(16)
    hashed = hashlib.pbkdf2_hmac("sha256", pw.encode(), _salt, 100_000)
    return f"{_salt.hex()}${hashed.hex()}"


def _verify_password(stored: str, pw: str) -> bool:
    """Verify *pw* against the stored hash."""
    if "$" in stored:
        salt_hex, hash_hex = stored.split("$", 1)
        salt = bytes.fromhex(salt_hex)
        hashed = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt, 100_000)
        return hmac.compare_digest(hash_hex, hashed.hex())
    # legacy unsalted sha256
    return stored == hashlib.sha256(pw.encode()).hexdigest()


def create_user(username: str, password: str, role: str = "seller", comune: str = ""):
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO users (username, password, role, comune) VALUES (?, ?, ?, ?)",
        (username, _hash_password(password), role, comune),
    )
    conn.commit()


def get_user(username: str):
    cur = conn.cursor()
    row = cur.execute(
        "SELECT username, password, role, comune FROM users WHERE username=?",
        (username,),
    ).fetchone()
    if row:
        return {
            "username": row[0],
            "password": row[1],
            "role": row[2],
            "comune": row[3] or "",
        }
    return None


def check_login(username: str, password: str):
    user = get_user(username)
    if not user:
        return None
    if _verify_password(user["password"], password):
        return user
    return None

# Initialize DB on import
init_db()
