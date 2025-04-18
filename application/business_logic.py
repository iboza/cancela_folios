import os
import webbrowser
from tkinter import messagebox
from config import RESULT_FILE_PREFIX, EXCEL_EXTENSIONS
from application.business_logic import get_latest_result_file


def get_latest_result_file(directory):
    result_files = [f for f in os.listdir(directory)
                    if f.startswith(RESULT_FILE_PREFIX) and f.endswith(tuple(EXCEL_EXTENSIONS))]
    if not result_files:
        return None
    return max(result_files, key=lambda f: os.path.getmtime(os.path.join(directory, f)))


def open_latest_result():
    if not app_state.loaded_file_directory or not os.path.isdir(app_state.loaded_file_directory):
        messagebox.showwarning("Advertencia", "No se ha seleccionado ningún archivo previamente")
        return

    try:
        latest_file = get_latest_result_file(app_state.loaded_file_directory)
        if not latest_file:
            messagebox.showwarning("Advertencia", "No se encontró ningún archivo de resultados.")
            return

        webbrowser.open(os.path.join(app_state.loaded_file_directory, latest_file))
        messagebox.showinfo("Éxito", f"Abriendo archivo: {latest_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {e}")
