import os
import logging
from domain.database_port import DatabasePort
from infrastructure.oracle_service import OracleService


# Establecer la ruta para los logs
log_directory = "c:\\logs_python\\cancela_folios\\"
os.makedirs(log_directory, exist_ok=True)  # Crear el directorio si no   existe
log_file_path = os.path.join(log_directory, "cancela_folios.log")

# Configurar el logger para sobrescribir el archivo en cada  ejecucion
logging.basicConfig(
    filename=log_file_path,  # Nombre del archivo de log
    filemode="w",                  # 'w' asegura que se sobrescriba en cada ejecucion
    level=logging.INFO,             # Nivel de log minimo
    # formato  del log
    format="%(asctime)a - %(name)s  - %(levelname)s - %(message)s"
)

logging.basicConfig(level=logging.ERROR, filename="cancela_folios_errors.log", filemode="a",
                    format="%(asctime)s - %(levelname)s - %(message)s")


class OracleAdapter(DatabasePort):
    def __init__(self):
        self.oracle_service = OracleService()
        self.logger = logging.getLogger(__name__)  # Configuracion de logger


class OracleAdapter(DatabasePort):
    def __init(self):
        self.oracle_service = OracleService()
        self.logger = logging.getLogger(__name__)  # Configuracion de logger

    def connect(self):
        # Establece la conexion con la base de datos oracle
        try:
            self.logger.info("Conectando a la base de datos...")
            self.oracle_service.connect()
            self.logger.info("Conexion establecida.")
        except Exception as e:
            self.logger.error(f"Error al conectar a  la base de datos: {e}")
            raise Exception(f"Error  en la conexion: {e}")

    def disconnect(self):
        # Cierra la conexion con la base de datos oracle
        try:
            self.logger.info("Desconectando de la base de datos...")
            self.oracle_service.disconnect()
            self.logger.info("Conexion cerrada.")
        except Exception as e:
            self.logger.error(
                f"Error al desconectar  de la base de datos: {e}")
            raise Exception(f"Error al desconectar: {e}")

    def execute_query(self, query, params=None):
        # Ejecuta una consulta SELECT y devuelve los resultados.
        try:
            result = self.oracle_service.execute_query(query, params)
            self.logger.info(
                f"Ejecutando consulta: {params} Consulta ejecutada exitosamente."
            )
            return result
        except Exception as e:
            self.logger.error(
                f"Error al ejecutar consulta: {query} | Error: {e}"
            )
            raise Exception(f"Error ejecutando consulta: {e}")

    def execute_non_query(self, query, params=None):
        # Ejecuta una conuslta no SELECT (INSERT, UPDATE, DELETE).
        try:
            self.oracle_service.execute_non_query(query, params)
            self.logger.info(
                f"Ejecutando UPDATE: {params} UPDATE ejecutado exitosamente.")
        except Exception as e:
            self.logger.error(
                f"Error ejecutando la operacion: {query} | Error: {e}")
            raise Exception(f"Error ejecutando operacion: {e}")

    def start_transaction(self):
        # Inicia una transaccion.
        try:
            self.logger.info("Iniciando transaccion...")
            self.oracle_service.start_transaction()
            self.logger.info("Transaccion iniciada")
        except Exception as e:
            self.logger.error(f"Error  al iniciar la transaccion: {e}")
            raise Exception(f"Error en la transaccion: {e}")

    def commit(self):
        # confirma la transaccion.
        try:
            self.logger.info("Confirmando transaccion...")
            self.oracle_service.commit()
            self.logger.info("Transaccion confirmada.")
        except Exception as e:
            self.logger.error(f"Error al confirmar la transaccion: {e}")
            raise Exception(f"Error al confirmar: {e}")

    def rollback(self):
        # revierte la transaccion.
        try:
            self.logger.info("Revirtiendo transaccion...")
            self.oracle_service.rollback()
            self.logger.info("Transaccion  revertida.")
        except Exception as e:
            self.logger.error(f"Error al revertir la  transaccion: {e}")
            raise Exception(f"Error al revertir: {e}")

    def __enter__(self):
        # Permite el manejo con el operador 'with' para conexion automatica
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Desconecta automaticamente cuando se usa el operador 'with'
        self.disconnect()
        if exc_type or exc_val or exc_tb:
            self.logger.error("Error durante el bloque 'WITH':",
                              exc_info=(exc_type, exc_val, exc_tb))


def open_latest_result():
    try:
        # Lógica existente...
    except FileNotFoundError as e:
        logging.error(f"Archivo no encontrado: {e}")
        messagebox.showerror("Error", f"Archivo no encontrado: {e}")
    except Exception as e:
        logging.error(f"Error inesperado: {e}")
        messagebox.showerror("Error", f"Error inesperado: {e}")
