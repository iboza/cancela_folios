import os
from infrastructure.excel_service import ExcelService
from application.oracle_adapter import OracleAdapter


class ExcelController:
    # Controlador para procesar archivos Excel y actualizar la base de datos.

    def __init__(self):
        self.excel_service = ExcelService()
        self.database_adapter = OracleAdapter()

    def process_file(self, file_path, messagebox):

        # Procesa un archivo Excel y actualiza los registros en la base de datos.

        try:
            # Leer el archivo Excel
            data = self.excel_service.read_excel(file_path)

            # Validar que el archivo tiene al menos una columna
            if data.shape[1] < 1:
                raise ValueError(
                    "El archivo Excel no tiene columnas suficientes.")

            # validar que el archivo tiene contenido en la columna A
            folios = data.iloc[:, 0].dropna().tolist()

            # limpiar los folios eliminando espacios en blanco
            folios = [str(folio).strip() for folio in folios]

            if not folios:
                messagebox.showinfo(
                    "Informacion", "El archivo Excel no contiene folios validos.")
                return
        except Exception as e:
            messagebox.showerror(
                "Error", f"Error al leer el archivo Excel:  {e}")
            return

        # conectar a la base de datos y procesar los folios
        query_select = "SELECT FOLIO_APM FROM HIP_SOLICITUD WHERE FOLIO_APM = :value"
        query_update = "UPDATE HIP_SOLICITUD SET ID_FLUJO =  8 WHERE FOLIO_APM = :value"

        try:
            self.database_adapter.connect()
            # Lista para almacenar resultaddos ("ACTUALIZADO" o "NO ENCONTRADO")
            estados = []

            for folio in folios:
                try:
                    select_result = self.database_adapter.execute_query(
                        query_select, {"value": folio})
                    if select_result:
                        self.database_adapter.execute_non_query(
                            query_update, {"value": folio})
                        estados.append("ACTUALIZADO")
                    else:
                        estados.append("NO ENCONTRADO")
                except Exception as query_error:
                    estados.append("NO ENCONTRADO")
                    print(
                        f"Error al procesar FOLIO_APM='{folio}': {query_error}")
        except Exception as conn_error:
            messagebox.showerror(
                "ERROR", f"Error en la conexion o actualizacion: {conn_error}")
            return
        finally:
            self.database_adapter.disconnect()

        # crear y/o sobreescribir el archivo "RESULTADOS.xlsx" con los datos actualizados
        try:
            if len(data.columns) < 2:
                data.insert(1, 'RESULTADOS', '')

            data.iloc[:, 1] = estados

            output_file_path = os.path.join(
                os.path.dirname(
                    file_path), f"RESULTADOS_{os.path.basename(file_path)}"
            )

            self.excel_service.write_excel(output_file_path, data)
            messagebox.showinfo(
                "Exito", f"Archivo generado correctamente con los resultados: {output_file_path}")
        except Exception as e:
            messagebox.showerror(
                "Error", f"Error al generar el archivo excel: {e}. Cerrar  el archivo")
