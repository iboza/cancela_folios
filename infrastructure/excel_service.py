import pandas as pd


class ExcelService:
    def read_excel(self, file_path):
        try:
            """
            lee el archivo excel mientras se conservan todas las columnas
            y evita leer la primera columna
            """
            data = pd.read_excel(file_path, engine="openpyxl")
            if data.empty:
                raise ValueError("El archivo Excel está vacío.")
            return data.dropna()  # remueve columnas vacías
        except Exception as e:
            print(f"Error al leer el archivo Excel: {e}")
            raise

    def write_excel(self, file_path, data):
        try:
            data.to_excel(file_path, index=False, engine="openpyxl")
        except Exception as e:
            print(f"Error al sobreescribir en el archivo Excel: {e}")
            raise
