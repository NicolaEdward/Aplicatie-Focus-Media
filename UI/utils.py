import os
from dotenv import load_dotenv
import mysql.connector  # type: ignore

load_dotenv()

def get_db_connection():
    """Return a new connection to the MySQL database."""
    return mysql.connector.connect(
        host=os.environ.get("MYSQL_HOST", "localhost"),
        port=int(os.environ.get("MYSQL_PORT", 3306)),
        user=os.environ.get("MYSQL_USER", "root"),
        password=os.environ.get("MYSQL_PASSWORD", ""),
        database=os.environ.get("MYSQL_DATABASE", "focus_media"),
    )
