import os
from infrastructure.excel_service import ExcelService
from application.oracle_adapter import OracleAdapter


class ExcelController:
    def __init__(self):
        self.excel_service = ExcelService()
        self.database_adapter = OracleAdapter()

    def process_file(self, file_path, messagebox, progress_callback):
        try:
            data = self.excel_service.read_excel(file_path)
            if data.shape[1] < 1:
                raise ValueError(
                    "El archivo Excel no tiene columnas suficientes.")
            folios = data.iloc[:, 0].dropna().astype(str).str.strip().tolist()
            if not folios:
                messagebox.showinfo(
                    "Información", "El archivo Excel no contiene folios válidos.")
                return
        except Exception as e:
            messagebox.showerror(
                "Error", f"Error al leer el archivo Excel: {e}")
            return

        batch_size = 1000
        total = len(folios)
        estados = []

        try:
            self.database_adapter.connect()
            for i in range(0, total, batch_size):
                batch = folios[i:i+batch_size]
                placeholders = ",".join(
                    [f":folio{j}" for j in range(len(batch))])
                update_query = f"UPDATE HIP_SOLICITUD SET ID_FLUJO = 8 WHERE FOLIO_APM IN ({placeholders})"
                params = {f"folio{j}": folio for j, folio in enumerate(batch)}
                self.database_adapter.execute_non_query(update_query, params)
                estados.extend(["ACTUALIZADO"] * len(batch))
                progress_callback(i + len(batch), total)
        except Exception as conn_error:
            messagebox.showerror(
                "Error", f"Error al conectar o actualizar en la base de datos: {conn_error}")
            return
        finally:
            self.database_adapter.disconnect()

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
                "Éxito", f"Archivo generado correctamente con los resultados: {output_file_path}")
        except Exception as e:
            messagebox.showerror(
                "Error", f"Error al generar el archivo Excel: {e}. Cierre el archivo si está abierto.")
