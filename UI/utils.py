import os
from dotenv import load_dotenv
import mysql.connector  # type: ignore

# Load variables from the `.env` file placed in the project root so utilities
# function correctly when imported from various modules.
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env"))


def _parse_port(value: str | None) -> int | None:
    """Return a validated port number from *value* or ``None``."""
    if not value:
        return None
    value = value.strip()
    if value.isdigit():
        if len(value) > 5:
            value = value[-5:]
        port = int(value)
        if 0 < port <= 65535:
            return port
    raise ValueError(f"Invalid MYSQL_PORT value: {value!r}")

def get_db_connection():
    """Return a new connection to the MySQL database."""
    host = os.environ.get("MYSQL_HOST")
    port = _parse_port(os.environ.get("MYSQL_PORT"))

    if ":" in host:
        host_part, host_port = host.rsplit(":", 1)
        if host_port.isdigit():
            host = host_part
            if not port:
                port = _parse_port(host_port)

    params = {
        "host": host,
        "user": os.environ.get("MYSQL_USER"),
        "password": os.environ.get("MYSQL_PASSWORD"),
        "database": os.environ.get("MYSQL_DATABASE"),
    }
    if port:
        params["port"] = port

    # ``mysql.connector`` defaults the port to 3306 when not provided,
    # so only include it if ``MYSQL_PORT`` or the host string specified it.
    return mysql.connector.connect(**params)
