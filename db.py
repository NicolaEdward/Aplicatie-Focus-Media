import os
import time
import datetime
import hashlib
import hmac
import sqlite3

try:
    from dotenv import load_dotenv  # type: ignore
except Exception:  # pragma: no cover - optional dep
    load_dotenv = lambda *a, **k: None

# Load variables from a `.env` file next to this module when available so the
# application behaves the same regardless of the current working directory.
load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

try:
    import mysql.connector  # type: ignore
except Exception:  # pragma: no cover - optional dep
    mysql = None

try:
    import sqlalchemy  # type: ignore
except Exception:  # pragma: no cover - optional dep
    sqlalchemy = None


class _CursorWrapper:
    """Cursor wrapper translating ``?`` placeholders for MySQL."""

    def __init__(self, cur, mysql_mode: bool):
        self._cur = cur
        self._mysql = mysql_mode

    def execute(self, sql, params=None):
        if self._mysql:
            sql = sql.replace("?", "%s")
        try:
            # ``mysql.connector`` returns ``None`` from ``execute`` instead of
            # the cursor instance like ``sqlite3`` does.  Since a lot of the code
            # relies on chaining calls like ``cursor.execute(...).fetchall()``,
            # always return ``self`` so the wrapper mimics the sqlite behaviour.
            self._cur.execute(sql, params or ())
        except Exception as exc:
            if _needs_reconnect(exc):
                reconnect()
                self._cur = cursor._cur
                self._mysql = cursor._mysql
                if self._mysql:
                    sql = sql.replace("?", "%s")
                self._cur.execute(sql, params or ())
            else:
                raise
        return self

    def executemany(self, sql, params):
        if self._mysql:
            sql = sql.replace("?", "%s")
        try:
            # See comment in ``execute`` above about return value.
            self._cur.executemany(sql, params)
        except Exception as exc:
            if _needs_reconnect(exc):
                reconnect()
                self._cur = cursor._cur
                self._mysql = cursor._mysql
                if self._mysql:
                    sql = sql.replace("?", "%s")
                self._cur.executemany(sql, params)
            else:
                raise
        return self

    def __getattr__(self, name):
        return getattr(self._cur, name)


class _ConnWrapper:
    def __init__(self, conn, mysql_mode: bool):
        self._conn = conn
        self._mysql = mysql_mode

    def cursor(self):
        try:
            return _CursorWrapper(self._conn.cursor(), self._mysql)
        except Exception as exc:
            if _needs_reconnect(exc):
                reconnect()
                self._conn = conn._conn
                self._mysql = conn._mysql
                return _CursorWrapper(self._conn.cursor(), self._mysql)
            raise

    @property
    def mysql(self):
        return self._mysql

    def commit(self):
        try:
            self._conn.commit()
        except Exception as exc:
            if _needs_reconnect(exc):
                reconnect()
                self._conn = conn._conn
                self._mysql = conn._mysql
                self._conn.commit()
            else:
                raise
        try:
            refresh_location_cache()
        except Exception:
            pass

    def __getattr__(self, name):
        return getattr(self._conn, name)


def get_db_path() -> str:
    """Return the absolute path to the bundled ``locatii.db`` database."""
    base_dir = os.path.dirname(__file__)
    return os.path.join(base_dir, "locatii.db")


def _parse_port(value: str | None) -> int | None:
    """Return a validated port number from *value* or ``None``."""
    if not value:
        return None
    value = value.strip()
    if value.isdigit():
        # Ports are at most 5 digits; take the last chunk if more were provided
        if len(value) > 5:
            value = value[-5:]
        port = int(value)
        if 0 < port <= 65535:
            return port
    raise ValueError(f"Invalid MYSQL_PORT value: {value!r}")


def _create_connection():
    host = os.environ.get("MYSQL_HOST")
    if host and mysql is not None:
        try:
            port = _parse_port(os.environ.get("MYSQL_PORT"))
        except Exception:
            # invalid port configured; ignore and fall back to SQLite
            return _ConnWrapper(sqlite3.connect(get_db_path()), False)

        # Allow specifying the port as part of the host, e.g. ``HOST=example:3306``
        if ":" in host:
            host_part, host_port = host.rsplit(":", 1)
            if host_port.isdigit():
                host = host_part
                if not port:
                    try:
                        port = _parse_port(host_port)
                    except Exception:
                        return _ConnWrapper(sqlite3.connect(get_db_path()), False)

        params = {
            "host": host,
            "user": os.environ.get("MYSQL_USER"),
            "password": os.environ.get("MYSQL_PASSWORD"),
            "database": os.environ.get("MYSQL_DATABASE"),
        }
        if port:
            params["port"] = port
        # mysql.connector uses port 3306 by default when not provided.
        try:
            conn = mysql.connector.connect(**params)
            return _ConnWrapper(conn, True)
        except Exception:
            pass  # fall back to bundled SQLite database
    return _ConnWrapper(sqlite3.connect(get_db_path()), False)


def _needs_reconnect(exc: Exception) -> bool:
    """Return ``True`` if *exc* indicates a lost MySQL connection."""
    if mysql is None:
        return False
    err = getattr(exc, "errno", None)
    return err in (2006, 2013)


def reconnect() -> None:
    """Recreate the global connection and cursor."""
    global conn, cursor
    conn = _create_connection()
    cursor = conn.cursor()
    try:
        refresh_location_cache()
    except Exception:
        pass


conn = _create_connection()
cursor = conn.cursor()

# --- simple in-memory cache for the locatii table ---
_location_cache: list[dict] | None = None
_cache_timestamp: float = 0.0


def refresh_location_cache() -> None:
    """Load all rows from ``locatii`` into memory."""
    global _location_cache, _cache_timestamp
    cur = conn.cursor()
    cur.execute("SELECT * FROM locatii")
    cols = [d[0] for d in cur.description]
    _location_cache = [dict(zip(cols, row)) for row in cur.fetchall()]
    _cache_timestamp = time.time()


def get_location_cache() -> list[dict]:
    """Return the cached locations, loading them on first use."""
    if _location_cache is None:
        refresh_location_cache()
    return list(_location_cache)


def maybe_refresh_location_cache(ttl: int = 300) -> bool:
    """Refresh cache if more than ``ttl`` seconds elapsed since last update."""
    if time.time() - _cache_timestamp >= ttl:
        refresh_location_cache()
        return True
    return False


def get_location_by_id(loc_id: int) -> dict | None:
    """Return location data from cache for the given ``loc_id``."""
    for row in get_location_cache():
        if row.get("id") == loc_id:
            return dict(row)
    return None


def pandas_conn():
    """Return a connection/engine suitable for ``pandas.read_sql_query``."""
    if getattr(conn, "mysql", False):
        host = os.environ.get("MYSQL_HOST")
        port = _parse_port(os.environ.get("MYSQL_PORT"))
        if host and ":" in host and not port:
            host, host_port = host.rsplit(":", 1)
            if host_port.isdigit():
                port = _parse_port(host_port)

        user = os.environ.get("MYSQL_USER", "")
        password = os.environ.get("MYSQL_PASSWORD", "")
        db_name = os.environ.get("MYSQL_DATABASE", "")

        url = f"mysql+mysqlconnector://{user}:{password}@{host}"
        if port:
            url += f":{port}"
        url += f"/{db_name}"

        if sqlalchemy is not None:
            if not hasattr(pandas_conn, "_engine"):
                pandas_conn._engine = sqlalchemy.create_engine(url)
            return pandas_conn._engine

        # ``pandas.read_sql_query`` also accepts an SQLAlchemy URL string
        # and will create the engine on demand if SQLAlchemy is available.
        # Returning the URL avoids passing the raw ``mysql.connector``
        # connection which triggers a UserWarning.
        return url
    # ``pandas.read_sql_query`` expects a raw DB-API connection or SQLAlchemy
    # engine.  When using SQLite we wrap the connection in ``_ConnWrapper``
    # so translate it back to the underlying connection object.
    if isinstance(conn, _ConnWrapper):
        return conn._conn
    return conn


def read_sql_query(sql: str, params=None, **kwargs):
    """Return a DataFrame using ``pandas.read_sql_query`` with adapted placeholders.

    Pandas expects ``%s`` style placeholders when working with MySQL connections
    regardless of whether an SQLAlchemy engine or a raw ``mysql.connector``
    connection is used.  Replace the ``?`` placeholders accordingly so queries
    work with both backends.
    """
    import pandas as pd

    if getattr(conn, "mysql", False):
        sql = sql.replace("?", "%s")

        # Avoid ``UserWarning`` when ``sqlalchemy`` is missing by falling back
        # to manual fetches instead of passing the raw DB-API connection.
        if sqlalchemy is None:
            cur = conn.cursor()
            cur.execute(sql, params or ())
            cols = [d[0] for d in cur.description]
            df = pd.DataFrame(cur.fetchall(), columns=cols)

            parse_dates = kwargs.pop("parse_dates", None)
            if parse_dates:
                for col in parse_dates:
                    df[col] = pd.to_datetime(df[col])
            return df

    return pd.read_sql_query(sql, pandas_conn(), params=params, **kwargs)


def ensure_index(table: str, index_name: str, column: str) -> None:
    """Create *index_name* on *table* if it is missing."""
    if getattr(conn, "mysql", False):
        cur = conn.cursor()
        cur.execute(f"SHOW INDEX FROM {table} WHERE Key_name=?", (index_name,))
        if not cur.fetchone():
            length = ""
            cur.execute(f"SHOW FIELDS FROM {table} WHERE Field=?", (column,))
            field = cur.fetchone()
            if field:
                ctype = str(field[1]).lower()
                if "text" in ctype or "blob" in ctype:
                    length = "(255)"
            cur.execute(f"CREATE INDEX {index_name} ON {table}({column}{length})")
    else:
        cursor.execute(f"CREATE INDEX IF NOT EXISTS {index_name} ON {table}({column})")


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
                face VARCHAR(32) DEFAULT 'Fața A',
                is_mobile TINYINT(1) DEFAULT 0,
                parent_id INT
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
            face TEXT DEFAULT 'Fața A',
            is_mobile INTEGER DEFAULT 0,
            parent_id INTEGER
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
        existing = {
            col[1] for col in cursor.execute("PRAGMA table_info(locatii)").fetchall()
        }

        to_add = {
            "rental_fee": "REAL DEFAULT 0",
            "pret_vanzare": "REAL",
            "pret_flotant": "REAL",
            "client_id": "INTEGER",
            "is_mobile": "INTEGER DEFAULT 0",
            "parent_id": "INTEGER",
        }

        existing = {
            col[1] for col in cursor.execute("PRAGMA table_info(locatii)").fetchall()
        }
        if "face" not in existing:
            cursor.execute("ALTER TABLE locatii ADD COLUMN face TEXT DEFAULT 'Fața A'")
            conn.commit()

        for col, definition in to_add.items():
            if col not in existing:
                cursor.execute(f"ALTER TABLE locatii ADD COLUMN {col} {definition}")
                conn.commit()
    else:
        cur = conn.cursor()
        cur.execute("SHOW COLUMNS FROM locatii")
        existing = {row[0] for row in cur.fetchall()}

        to_add = {
            "pret_vanzare": "DOUBLE",
            "pret_flotant": "DOUBLE",
            "client_id": "INT",
            "is_mobile": "TINYINT(1) DEFAULT 0",
            "parent_id": "INT",
        }

        if "face" not in existing:
            cur.execute("ALTER TABLE locatii ADD COLUMN face VARCHAR(32) DEFAULT 'Fața A'")
            conn.commit()

        for col, definition in to_add.items():
            if col not in existing:
                cur.execute(f"ALTER TABLE locatii ADD COLUMN {col} {definition}")
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
    cursor.execute("SELECT COUNT(*) FROM users")
    if not cursor.fetchone()[0]:
        cursor.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, 'admin')",
            ("admin", _hash_password("admin")),
        )
    conn.commit()


def update_statusuri_din_rezervari():
    today = datetime.date.today().isoformat()
    cur = conn.cursor()

    # Ștergem rezervările expirate (fără sumă) care nu au fost anulate
    cur.execute(
        "DELETE FROM rezervari WHERE data_end < ? AND suma IS NULL",
        (today,),
    )

    # 1) Resetăm totul la Disponibil
    cur.execute(
        """
        UPDATE locatii
        SET status='Disponibil',
            client=NULL,
            client_id=NULL,
            data_start=NULL,
            data_end=NULL
    """
    )

    # 2) Marcăm rezervările curente fără sumă ca 'Rezervat'
    cur.execute(
        """
        UPDATE locatii
        SET status     = 'Rezervat',
            client     = (
                SELECT client
                  FROM rezervari
                 WHERE rezervari.loc_id = locatii.id
                   AND ? BETWEEN data_start AND data_end
                   AND suma IS NULL
                 ORDER BY data_start DESC
                 LIMIT 1
            ),
            data_start = (
                SELECT data_start FROM rezervari
                 WHERE rezervari.loc_id = locatii.id
                   AND ? BETWEEN data_start AND data_end
                   AND suma IS NULL
                 ORDER BY data_start DESC
                 LIMIT 1
            ),
            data_end   = (
                SELECT data_end FROM rezervari
                 WHERE rezervari.loc_id = locatii.id
                   AND ? BETWEEN data_start AND data_end
                   AND suma IS NULL
                 ORDER BY data_start DESC
                 LIMIT 1
            )
        WHERE EXISTS (
            SELECT 1 FROM rezervari
             WHERE rezervari.loc_id = locatii.id
               AND ? BETWEEN data_start AND data_end
               AND suma IS NULL
        )
    """,
        (today, today, today, today),
    )

    # 3) Marcăm închirierile curente ca 'Închiriat'
    cur.execute(
        """
        UPDATE locatii
        SET status      = 'Închiriat',
            client      = (
                SELECT client
                FROM rezervari
                WHERE rezervari.loc_id = locatii.id
                  AND ? BETWEEN data_start AND data_end
                  AND suma IS NOT NULL
                ORDER BY data_start DESC
                LIMIT 1
            ),
            client_id   = (
                SELECT client_id
                FROM rezervari
                WHERE rezervari.loc_id = locatii.id
                  AND ? BETWEEN data_start AND data_end
                  AND suma IS NOT NULL
                ORDER BY data_start DESC
                LIMIT 1
            ),
            data_start  = (
                SELECT data_start
                FROM rezervari
                WHERE rezervari.loc_id = locatii.id
                  AND ? BETWEEN data_start AND data_end
                  AND suma IS NOT NULL
                ORDER BY data_start DESC
                LIMIT 1
            ),
            data_end    = (
                SELECT data_end
                FROM rezervari
                WHERE rezervari.loc_id = locatii.id
                  AND ? BETWEEN data_start AND data_end
                  AND suma IS NOT NULL
                ORDER BY data_start DESC
                LIMIT 1
            )
        WHERE EXISTS (
            SELECT 1
            FROM rezervari
            WHERE rezervari.loc_id = locatii.id
              AND ? BETWEEN data_start AND data_end
              AND suma IS NOT NULL
        )
    """,
        (today, today, today, today, today),
    )

    # Mark expired mobile instances as hidden
    cur.execute(
        """
        UPDATE locatii
           SET status='Expirat'
         WHERE is_mobile=1 AND parent_id IS NOT NULL
           AND data_end IS NOT NULL AND data_end < ?
        """,
        (today,),
    )

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
    cur.execute(
        "SELECT username, password, role, comune FROM users WHERE username=?",
        (username,),
    )
    row = cur.fetchone()
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
refresh_location_cache()
