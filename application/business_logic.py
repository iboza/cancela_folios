import os
import webbrowser
from tkinter import messagebox
from config import RESULT_FILE_PREFIX, EXCEL_EXTENSIONS
from application.business_logic import get_latest_result_file, open_file_in_browser


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
