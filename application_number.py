import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from application.excel_controller import ExcelController
from application.business_logic import open_latest_result
from config import (
    IMAGE_PATH, WINDOW_WIDTH, WINDOW_HEIGHT, BUTTON_BG_COLOR, BUTTON_FG_COLOR,
    APP_TITLE, BUTTON_FONT, BUTTON_WIDTH, BUTTON_HEIGHT
)


class AppState:
    def __init__(self):
        self.loaded_file_directory = None


app_state = AppState()
excel_controller = ExcelController()


def exit_app():
    root.destroy()


def select_file():
    file_path = filedialog.askopenfilename(
        title=APP_TITLE,
        filetypes=[("Archivos de Excel", "*.xlsx;*.xls")]
    )
    if not file_path or not os.path.exists(file_path):
        messagebox.showwarning(
            "Advertencia", "No se seleccionó ningún archivo válido.")
        return

    app_state.loaded_file_directory = os.path.dirname(file_path)

    progress_window = tk.Toplevel(root)
    progress_window.title("Procesando...")
    progress_window.geometry("300x100")
    progress_window.resizable(False, False)
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = (screen_height // 2) - 50
    position_right = (screen_width // 2) - 150
    progress_window.geometry(f"300x100+{position_right}+{position_top}")

    progress_bar = ttk.Progressbar(
        progress_window, orient="horizontal", length=270, mode="determinate")
    progress_bar.pack(pady=20)
    progress_label = tk.Label(progress_window, text="Iniciando...")
    progress_label.pack()

    def update_progress(current, total):
        progress_bar["value"] = (current / total) * 100
        progress_label.config(text=f"Procesando {current} de {total} folios.")
        progress_window.update_idletasks()

    try:
        excel_controller.process_file(file_path, messagebox, update_progress)
    except Exception as e:
        messagebox.showerror("Error", f"Error al procesar el archivo: {e}")
    finally:
        progress_window.destroy()


def load_image(path, size=(200, 200)):
    try:
        image = Image.open(path)
        image = image.resize(size, Image.LANCZOS)
        return ImageTk.PhotoImage(image)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar la imagen: {e}")
        return None


def create_main_window():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
    root.configure(bg="#f0f0f0")
    root.resizable(False, False)
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = (screen_height // 2) - (WINDOW_HEIGHT // 2)
    position_right = (screen_width // 2) - (WINDOW_WIDTH // 2)
    root.geometry(
        f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{position_right}+{position_top}")

    icon_path = os.path.abspath(os.path.join("sources", "favion.ico"))
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except Exception as e:
            messagebox.showwarning(
                "Advertencia", f"No se pudo cargar el ícono: {e}")
    else:
        messagebox.showwarning(
            "Advertencia", "El ícono no se encontró en la ruta especificada.")
    return root


root = create_main_window()
photo = load_image(IMAGE_PATH)
if photo:
    tk.Label(root, image=photo).pack(pady=20)

tk.Button(
    root, text="Seleccionar Archivo", command=select_file,
    bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
    width=BUTTON_WIDTH, height=BUTTON_HEIGHT
).pack(pady=10)

tk.Button(
    root, text="Resultado", command=lambda: open_latest_result(app_state),
    bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
    width=BUTTON_WIDTH, height=BUTTON_HEIGHT
).pack(pady=10)

tk.Button(
    root, text="Salir", command=exit_app,
    bg=BUTTON_BG_COLOR, fg=BUTTON_FG_COLOR, font=BUTTON_FONT,
    width=BUTTON_WIDTH, height=BUTTON_HEIGHT
).pack(pady=10)

root.mainloop()
