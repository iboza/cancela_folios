import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import webbrowser
from application.excel_controller import ExcelController
from config import IMAGE_PATH, WINDOW_WIDTH, WINDOW_HEIGHT, BUTTON_BG_COLOR, BUTTON_FG_COLOR, RESULT_FILE_PREFIX, EXCEL_EXTENSIONS, APP_TITLE, BUTTON_FONT, BUTTON_WIDTH, BUTTON_HEIGHT


class AppState:
    def __init__(self):
        self.loaded_file_directory = None


app_state = AppState()


def exit_app():
    root.destroy()


def open_latest_result():
    if not app_state.loaded_file_directory or not os.path.isdir(app_state.loaded_file_directory):
        messagebox.showwarning(
            "Advertencia", "No se ha seleccionado ningún archivo previamente")
        return

    result_files = [f for f in os.listdir(app_state.loaded_file_directory)
                    if f.startswith(RESULT_FILE_PREFIX) and f.endswith(tuple(EXCEL_EXTENSIONS))]

    if not result_files:
        messagebox.showwarning(
            "Advertencia", "No se encontró ningún archivo de resultados.")
        return

    try:
        latest_file = max(result_files, key=lambda f: os.path.getmtime(
            os.path.join(app_state.loaded_file_directory, f)))
        webbrowser.open(os.path.join(app_state.loaded_file_directory, latest_file))
        messagebox.showinfo("Éxito", f"Abriendo archivo: {latest_file}")
    except FileNotFoundError as e:
        messagebox.showerror("Error", f"Archivo no encontrado: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {e}")


def select_file():
    file_path = filedialog.askopenfilename(
        title=APP_TITLE,
        filetypes=[("Archivos de Excel", "*.xlsx;*.xls")]
    )
    if not file_path:
        messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
        return

    if not os.path.exists(file_path):
        messagebox.showerror("Error", "El archivo seleccionado no existe.")
        return

    app_state.loaded_file_directory = os.path.dirname(file_path)
    try:
        excel_controller.process_file(file_path, messagebox)
    except Exception as e:
        messagebox.showerror("Error", f"Error al procesar el archivo: {e}")


def load_image(path, size=(200, 200)):
    try:
        image = Image.open(path)
        image = image.resize(size, Image.LANCZOS)
        return ImageTk.PhotoImage(image)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar la imagen: {e}")
        return None


def create_main_window():
    # Configura y devuelve la ventana principal de la aplicación.

    root = tk.Tk()
    root.title(APP_TITLE)
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

    # Establecer el icono de la ventana
    icon_path = os.path.join("sources", "favion.ico")
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except Exception as e:
            messagebox.showwarning("Advertencia", f"No se pudo cargar el ícono: {e}")
    else:
        messagebox.showwarning("Advertencia", "El ícono no se encontró en la ruta especificada.")

    return root


# Crear la ventana principal
root = create_main_window()

# Cargar y mostrar el logo
photo = load_image(IMAGE_PATH)
if photo:
    label_image = tk.Label(root, image=photo)
    label_image.pack(pady=20)

# Crear una instancia de ExcelController
excel_controller = ExcelController()

# Botón para seleccionar archivo
btn_select_file = tk.Button(
    root, text="Seleccionar Archivo",
    command=lambda: select_file(app_state),
    bg=BUTTON_BG_COLOR,
    fg=BUTTON_FG_COLOR,
    font=BUTTON_FONT,
    width=BUTTON_WIDTH,
    height=BUTTON_HEIGHT
)
btn_select_file.pack(pady=10)

# Botón para abrir archivo más reciente
btn_open_result = tk.Button(
    root, text="Resultado",
    command=open_latest_result,
    bg=BUTTON_BG_COLOR,
    fg=BUTTON_FG_COLOR,
    font=BUTTON_FONT,
    width=BUTTON_WIDTH,
    height=BUTTON_HEIGHT
)
btn_open_result.pack(pady=10)

# Botón para salir
btn_exit = tk.Button(
    root, text="Salir",
    command=exit_app,
    bg=BUTTON_BG_COLOR,
    fg=BUTTON_FG_COLOR,
    font=BUTTON_FONT,
    width=BUTTON_WIDTH,
    height=BUTTON_HEIGHT
)
btn_exit.pack(pady=10)

root.mainloop()
