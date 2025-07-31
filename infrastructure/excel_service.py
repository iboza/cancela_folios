import pandas as pd


class ExcelService:
    def read_excel(self, file_path):
        try:
            data = pd.read_excel(file_path, engine="openpyxl")
            if data.empty:
                raise ValueError("El archivo Excel está vacío.")
            return data.dropna(how="all")
        except Exception as e:
            print(f"Error al leer el archivo Excel: {e}")
            raise

    def write_excel(self, file_path, data):
        try:
            data.to_excel(file_path, index=False, engine="openpyxl")
        except Exception as e:
            print(f"Error al escribir en el archivo Excel: {e}")
            raise
