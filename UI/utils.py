import os
from dotenv import load_dotenv
import mysql.connector  # type: ignore

# Load variables from the `.env` file placed in the project root so utilities
# function correctly when imported from various modules.
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env"))

def get_db_connection():
    """Return a new connection to the MySQL database."""
    host = os.environ.get("MYSQL_HOST")
    port = os.environ.get("MYSQL_PORT")

    if ":" in host:
        host_part, host_port = host.rsplit(":", 1)
        if host_port.isdigit():
            host = host_part
            if not port:
                port = host_port

    params = {
        "host": host,
        "user": os.environ.get("MYSQL_USER"),
        "password": os.environ.get("MYSQL_PASSWORD"),
        "database": os.environ.get("MYSQL_DATABASE"),
    }
    if port:
        params["port"] = int(port)

    # ``mysql.connector`` defaults the port to 3306 when not provided,
    # so only include it if ``MYSQL_PORT`` or the host string specified it.
    return mysql.connector.connect(**params)
