import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import webbrowser
from application.excel_controller import ExcelController
from config import IMAGE_PATH, WINDOW_WIDTH, WINDOW_HEIGHT, BUTTON_BG_COLOR, BUTTON_FG_COLOR, RESULT_FILE_PREFIX, EXCEL_EXTENSIONS

# Variable global para almacenar la ruta del archivo cargado
loaded_file_directory = None


def exit_app():
    root.destroy()


def open_latest_result():
    try:
        if not loaded_file_directory:
            messagebox.showwarning(
                "Advertencia", "No se ha seleccionado ningún archivo previamente")
            return

        # Filtrar archivos que comienzan con el prefijo y terminan con extensiones válidas
        result_files = [f for f in os.listdir(loaded_file_directory)
                        if f.startswith(RESULT_FILE_PREFIX) and f.endswith(tuple(EXCEL_EXTENSIONS))]

        if not result_files:
            messagebox.showwarning(
                "Advertencia", "No se encontró ningún archivo de resultados.")
            return

        # Identificar el archivo más reciente por fecha de modificación
        latest_file = max(result_files, key=lambda f: os.path.getmtime(
            os.path.join(loaded_file_directory, f)))

        # Abrir el archivo utilizando el navegador por defecto asociado al SO
        webbrowser.open(os.path.join(loaded_file_directory, latest_file))
        messagebox.showinfo("Éxito", f"Abriendo archivo: {latest_file}")
    except FileNotFoundError as e:
        messagebox.showerror("Error", f"Archivo no encontrado: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {e}")


def select_file():
    global loaded_file_directory
    file_path = filedialog.askopenfilename(
        title="Cancelación de Folios",
        filetypes=[("Archivos de Excel", "*.xlsx;*.xls")]
    )
    if not file_path:
        messagebox.showwarning(
            "Advertencia", "No se seleccionó ningún archivo.")
        return

    if not os.path.exists(file_path):
        messagebox.showerror("Error", "El archivo seleccionado no existe.")
        return

    # Guardar la ruta del directorio del archivo seleccionado
    loaded_file_directory = os.path.dirname(file_path)

    # Llama al método process_file de ExcelController
    excel_controller.process_file(file_path, messagebox)


# Crear la ventana principal
root = tk.Tk()
root.title("Cancelación de Folios")
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
root.configure(bg="#f0f0f0")
root.resizable(False, False)

# Centrar la ventana en pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = (screen_height // 2) - (WINDOW_HEIGHT // 2)
position_right = (screen_width // 2) - (WINDOW_WIDTH // 2)
root.geometry(
    f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{position_right}+{position_top}")

# Cargar y mostrar el logo
image = Image.open(IMAGE_PATH)
image = image.resize((200, 200), Image.Resampling.LANCZOS)
photo = ImageTk.PhotoImage(image)

label_image = tk.Label(root, image=photo)
label_image.pack(pady=20)

# Crear una instancia de ExcelController
excel_controller = ExcelController()

# Botón para seleccionar archivo
btn_select_file = tk.Button(
    root, text="Seleccionar Archivo",
    command=select_file,
    bg=BUTTON_BG_COLOR,
    fg=BUTTON_FG_COLOR,
    font=("Arial", 12),
    width=20,
    height=2
)
btn_select_file.pack(pady=10)

# Botón para abrir archivo más reciente
btn_open_result = tk.Button(
    root, text="Resultado",
    command=open_latest_result,
    bg=BUTTON_BG_COLOR,
    fg=BUTTON_FG_COLOR,
    font=("Arial", 12),
    width=20,
    height=2
)
btn_open_result.pack(pady=10)

# Botón para salir
btn_exit = tk.Button(
    root, text="Salir",
    command=exit_app,
    bg=BUTTON_BG_COLOR,
    fg=BUTTON_FG_COLOR,
    font=("Arial", 12),
    width=20,
    height=2
)
btn_exit.pack(pady=10)

root.mainloop()
