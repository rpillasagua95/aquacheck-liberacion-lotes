# utils.py
import os
import sys

def obtener_directorio_app():
    if hasattr(sys, 'frozen'):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

APP_DIR = obtener_directorio_app()
LOCK_FILE = os.path.join(APP_DIR, "aquacheck.lock")

def create_lock():
    """Crea el archivo .lock con el PID del proceso actual."""
    pid = os.getpid()
    with open(LOCK_FILE, "w") as f:
        f.write(str(pid))
    return pid

def release_lock(pid):
    """Libera el archivo .lock si el PID coincide."""
    try:
        if os.path.exists(LOCK_FILE):
            with open(LOCK_FILE, "r") as f:
                lock_pid = int(f.read().strip())
            if lock_pid == pid:
                os.remove(LOCK_FILE)
    except (IOError, ValueError):
        # Si hay error, es mejor intentar borrarlo
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)