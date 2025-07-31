import time
import oracledb
from config import (
    DB_HOST, DB_PORT, DB_SERVICE_NAME, DB_USER, DB_PASSWORD,
    MAX_RETRIES, RETRY_DELAY
)


class OracleService:
    def __init__(self):
        self.connection = None

    def connect(self):
        dsn = f"{DB_HOST}:{DB_PORT}/{DB_SERVICE_NAME}"
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                self.connection = oracledb.connect(
                    user=DB_USER, password=DB_PASSWORD, dsn=dsn
                )
                print("Connected to Oracle Database")
                return
            except Exception as e:
                print(f"Failed to connect (attempt {attempt}): {e}")
                self.connection = None
                if attempt < MAX_RETRIES:
                    time.sleep(RETRY_DELAY)
        raise Exception("Could not connect to Oracle Database after retries.")

    def disconnect(self):
        if self.connection:
            try:
                self.connection.close()
                print("Disconnected from Oracle Database")
            finally:
                self.connection = None

    def is_connection_alive(self):
        try:
            if self.connection:
                self.connection.ping()
                return True
        except Exception:
            return False
        return False

    def ensure_connection(self):
        if self.connection is None or not self.is_connection_alive():
            print("Reconnecting to Oracle Database...")
            self.connect()

    def execute_query(self, query, params=None):
        self.ensure_connection()
        with self.connection.cursor() as cursor:
            cursor.execute(query, params or {})
            return cursor.fetchall()

    def execute_non_query(self, query, params=None):
        self.ensure_connection()
        with self.connection.cursor() as cursor:
            cursor.execute(query, params or {})
        self.connection.commit()

    def start_transaction(self):
        self.ensure_connection()
        self.connection.begin()

    def commit(self):
        self.ensure_connection()
        self.connection.commit()

    def rollback(self):
        self.ensure_connection()
        self.connection.rollback()
