import os
import webbrowser
from tkinter import messagebox
from config import RESULT_FILE_PREFIX, EXCEL_EXTENSIONS
from application.business_logic import get_latest_result_file, open_file_in_browser


def get_latest_result_file(directory):
    """
    Obtiene el archivo de resultados más reciente en el directorio especificado.
    """
    result_files = [f for f in os.listdir(directory)
                    if f.startswith(RESULT_FILE_PREFIX) and f.endswith(tuple(EXCEL_EXTENSIONS))]

    if not result_files:
        return None

    latest_file = max(result_files, key=lambda f: os.path.getmtime(
        os.path.join(directory, f)))
    return latest_file


def open_file_in_browser(file_path):
    """
    Abre un archivo en el navegador predeterminado del sistema.
    """
    try:
        webbrowser.open(file_path)
        return True
    except Exception as e:
        raise e


def open_latest_result():
    try:
        if not loaded_file_directory:
            messagebox.showwarning(
                "Advertencia", "No se ha seleccionado ningún archivo previamente")
            return

        latest_file = get_latest_result_file(loaded_file_directory)

        if not latest_file:
            messagebox.showwarning(
                "Advertencia", "No se encontró ningún archivo de resultados.")
            return

        open_file_in_browser(os.path.join(loaded_file_directory, latest_file))
        messagebox.showinfo("Éxito", f"Abriendo archivo: {latest_file}")
    except FileNotFoundError as e:
        messagebox.showerror("Error", f"Archivo no encontrado: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {e}")
