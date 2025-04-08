import oracledb
from config import DB_HOST, DB_PORT, DB_SERVICE_NAME, DB_USER, DB_PASSWORD


class OracleService:
    def __init__(self):
        self.connection = None

    def connect(self):
        try:
            # Usar las constantes definidas en config.py
            dsn = f"{DB_HOST}:{DB_PORT}/{DB_SERVICE_NAME}"
            self.connection = oracledb.connect(
                user=DB_USER, password=DB_PASSWORD, dsn=dsn
            )
            print("Connected to Oracle Database")
        except Exception as e:
            print(f"Failed to connect to Oracle Database: {e}")
            raise

    def disconnect(self):
        if self.connection:
            self.connection.close()
            print("Disconnected from Oracle Database")

    def execute_query(self, query, params=None):
        # Ejecuta consultas que devuelven filas (SELECT)
        if not self.connection:
            raise Exception("No active database connection")
        with self.connection.cursor() as cursor:
            cursor.execute(query, params or {})
            return cursor.fetchall()

    def execute_non_query(self, query, params=None):
        # Ejecuta consultas que no devuelven filas (UPDATE, INSERT, DELETE)
        if not self.connection:
            raise Exception("No active database connection")
        with self.connection.cursor() as cursor:
            cursor.execute(query, params or {})
            self.connection.commit()  # Confirmar los cambios

    def start_transaction(self):
        if not self.connection:
            raise Exception("No active database connection")
        self.connection.begin()

    def commit(self):
        if not self.connection:
            raise Exception("No active database connection")
        self.connection.commit()

    def rollback(self):
        if not self.connection:
            raise Exception("No active database connection")
        self.connection.rollback()
