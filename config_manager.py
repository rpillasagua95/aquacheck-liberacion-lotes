# -*- coding: utf-8 -*-
# config_manager.py (Reestructurado para mayor estabilidad)

import os
import json

# --- CONSTANTES DE RUTA ---
# Define el nombre de la carpeta de la aplicación en AppData
APP_NAME = "AquaCheck"
# Crea la ruta completa al directorio de la aplicación
APP_DIR = os.path.join(os.getenv('APPDATA'), APP_NAME)

# Define las rutas completas para los archivos de configuración y datos
CONFIG_FILE = os.path.join(APP_DIR, "config.json")
IMAGENES_GUARDADAS_FILE = os.path.join(APP_DIR, "imagenes_guardadas.json")
MENSAJES_ENVIADOS_FILE = os.path.join(APP_DIR, "mensajes_enviados.json")

# Asegura que el directorio de la aplicación exista al iniciar
os.makedirs(APP_DIR, exist_ok=True)

DEFAULT_CONFIG = {
    "carpeta_excel": "",
    "carpeta_fotos": os.path.join(APP_DIR, "capturas_temporales"),
    "carpeta_liberaciones": "",
    "primera_configuracion": True
}

def cargar_configuracion():
    """Carga la configuración desde el archivo JSON."""
    if not os.path.exists(CONFIG_FILE):
        # Si no existe, crea la carpeta de capturas temporales por defecto
        os.makedirs(DEFAULT_CONFIG["carpeta_fotos"], exist_ok=True)
        return DEFAULT_CONFIG.copy()
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (IOError, json.JSONDecodeError):
        return DEFAULT_CONFIG.copy()

def guardar_configuracion(config_data):
    """Guarda la configuración en el archivo JSON."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
        return True
    except IOError:
        return False

def verificar_configuracion():
    """Verifica si la configuración de carpetas está completa."""
    config = cargar_configuracion()
    if config.get("primera_configuracion", True):
        return False, "Es necesario configurar las carpetas por primera vez."
    
    if not config.get("carpeta_excel"):
        return False, "La carpeta de archivos Excel no está configurada."
            
    if not os.path.isdir(config["carpeta_excel"]):
        return False, f"La carpeta de Excel configurada no existe:\n{config['carpeta_excel']}"

    # La verificación de Green API se hace ahora directamente en whatsapp_service.py
    return True, "Configuración correcta."