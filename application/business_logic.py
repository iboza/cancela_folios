import os
import webbrowser
from tkinter import messagebox
from config import RESULT_FILE_PREFIX, EXCEL_EXTENSIONS

def get_latest_result_file(directory):
    result_files = [
        f for f in os.listdir(directory)
        if f.startswith(RESULT_FILE_PREFIX) and f.endswith(tuple(EXCEL_EXTENSIONS))
    ]
    if not result_files:
        return None
    return max(result_files, key=lambda f: os.path.getmtime(os.path.join(directory, f)))

def open_latest_result(app_state):
    directory = app_state.loaded_file_directory
    if not directory or not os.path.isdir(directory):
        messagebox.showwarning("Advertencia", "No se ha seleccionado ningún archivo previamente")
        return

    latest_file = get_latest_result_file(directory)
    if not latest_file:
        messagebox.showwarning("Advertencia", "No se encontró ningún archivo de resultados.")
        return

    try:
        webbrowser.open(os.path.join(directory, latest_file))
        messagebox.showinfo("Éxito", f"Abriendo archivo: {latest_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {e}")