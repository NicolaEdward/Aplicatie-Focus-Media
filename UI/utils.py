import os
from dotenv import load_dotenv
import mysql.connector  # type: ignore

# Load variables from the `.env` file placed in the project root so utilities
# function correctly when imported from various modules.
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env"))

def get_db_connection():
    """Return a new connection to the MySQL database."""
    host = os.environ.get("MYSQL_HOST", "localhost")
    port = os.environ.get("MYSQL_PORT")

    if ":" in host:
        host_part, host_port = host.rsplit(":", 1)
        if host_port.isdigit():
            host = host_part
            if not port:
                port = host_port

    return mysql.connector.connect(
        host=host,
        port=int(port or 3306),
        user=os.environ.get("MYSQL_USER", "root"),
        password=os.environ.get("MYSQL_PASSWORD", ""),
        database=os.environ.get("MYSQL_DATABASE", "focus_media"),
    )
