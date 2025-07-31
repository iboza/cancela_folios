import os
import logging
from domain.database_port import DatabasePort
from infrastructure.oracle_service import OracleService

log_directory = "c:\\logs_python\\cancela_folios\\"
os.makedirs(log_directory, exist_ok=True)
log_file_path = os.path.join(log_directory, "cancela_folios.log")
logging.basicConfig(
    filename=log_file_path,
    filemode="w",
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)


class OracleAdapter(DatabasePort):
    def __init__(self):
        self.oracle_service = OracleService()
        self.logger = logging.getLogger(__name__)

    def connect(self):
        try:
            self.logger.info("Conectando a la base de datos...")
            self.oracle_service.connect()
            self.logger.info("Conexión establecida.")
        except Exception as e:
            self.logger.error(f"Error al conectar a la base de datos: {e}")
            raise

    def disconnect(self):
        try:
            self.logger.info("Desconectando de la base de datos...")
            self.oracle_service.disconnect()
            self.logger.info("Conexión cerrada.")
        except Exception as e:
            self.logger.error(f"Error al desconectar de la base de datos: {e}")
            raise

    def execute_query(self, query, params=None):
        try:
            result = self.oracle_service.execute_query(query, params)
            self.logger.info(f"Consulta ejecutada: {query} | Params: {params}")
            return result
        except Exception as e:
            self.logger.error(
                f"Error al ejecutar consulta: {query} | Error: {e}")
            raise

    def execute_non_query(self, query, params=None):
        try:
            self.oracle_service.execute_non_query(query, params)
            self.logger.info(
                f"Operación ejecutada: {query} | Params: {params}")
        except Exception as e:
            self.logger.error(
                f"Error ejecutando operación: {query} | Error: {e}")
            raise

    def start_transaction(self):
        try:
            self.logger.info("Iniciando transacción...")
            self.oracle_service.start_transaction()
            self.logger.info("Transacción iniciada.")
        except Exception as e:
            self.logger.error(f"Error al iniciar la transacción: {e}")
            raise

    def commit(self):
        try:
            self.logger.info("Confirmando transacción...")
            self.oracle_service.commit()
            self.logger.info("Transacción confirmada.")
        except Exception as e:
            self.logger.error(f"Error al confirmar la transacción: {e}")
            raise

    def rollback(self):
        try:
            self.logger.info("Revirtiendo transacción...")
            self.oracle_service.rollback()
            self.logger.info("Transacción revertida.")
        except Exception as e:
            self.logger.error(f"Error al revertir la transacción: {e}")
            raise

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()
        if exc_type or exc_val or exc_tb:
            self.logger.error("Error durante el bloque 'with':",
                              exc_info=(exc_type, exc_val, exc_tb))
