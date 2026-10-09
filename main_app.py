# main_app.py (CORREGIDO)
import tkinter as tk
from ttkbootstrap.constants import *
import ttkbootstrap as ttk
from gui_manager import AppGUI
from utils import create_lock, release_lock, LOCK_FILE
import psutil
import os
import sys

# Importación para Drag and Drop
from tkinterdnd2 import TkinterDnD

# --- INICIO DE LA CORRECCIÓN ---
# Envolvemos la clase de la ventana para que soporte DND de forma correcta
class TkinterDnDWrapper(TkinterDnD.Tk):
    def __init__(self, **kwargs):
        # Primero, extraemos el argumento 'themename' que es solo para ttkbootstrap
        themename = kwargs.pop('themename', 'litera')
        
        # Luego, inicializamos la ventana base de TkinterDnD con los argumentos restantes
        super().__init__(**kwargs)

        # Finalmente, aplicamos el estilo de ttkbootstrap a la ventana ya creada
        self.style = ttk.Style(theme=themename)
# --- FIN DE LA CORRECCIÓN ---


def check_single_instance():
    """Verifica si ya hay una instancia del programa en ejecución."""
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r") as f:
                pid = int(f.read().strip())
            if psutil.pid_exists(pid):
                root = tk.Tk()
                root.withdraw()
                ttk.dialogs.Messagebox.show_error(
                    title="Error de Ejecución",
                    message=f"AquaCheck ya está en ejecución (PID: {pid}).\n\nPor favor, cierre la instancia anterior antes de abrir una nueva."
                )
                root.destroy()
                return False
            else:
                os.remove(LOCK_FILE)
        except (IOError, ValueError):
             os.remove(LOCK_FILE)
    return True


if __name__ == "__main__":
    if check_single_instance():
        lock_pid = create_lock()
        try:
            # Usamos nuestra clase envuelta corregida
            root = TkinterDnDWrapper(themename="flatly")
            
            app = AppGUI(root)
            root.mainloop()
        finally:
            release_lock(lock_pid)