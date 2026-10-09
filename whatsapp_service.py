# whatsapp_service.py (COMPLETO Y ACTUALIZADO)

import os
import requests
from dotenv import load_dotenv
from whatsapp_api_client_python import API
import time
import webbrowser
from tkinter import messagebox

# Carga las variables de entorno desde el archivo .env
load_dotenv()

# Ahora todas las variables se cargan desde el archivo .env
ID_INSTANCE = os.getenv('ID_INSTANCE')
API_TOKEN_INSTANCE = os.getenv('API_TOKEN_INSTANCE')
GROUP_CHAT_ID = os.getenv('GROUP_CHAT_ID') 
# Servidor de tu instancia de Green API (sale en su panel)
GREEN_API_HOST = os.getenv('GREEN_API_HOST', 'api.green-api.com')

# La verificación ahora incluye el GROUP_CHAT_ID
if not all([ID_INSTANCE, API_TOKEN_INSTANCE, GROUP_CHAT_ID]):
    messagebox.showerror(
        "Error de Configuración", 
        "No se encontraron 'ID_INSTANCE', 'API_TOKEN_INSTANCE' o 'GROUP_CHAT_ID' en el archivo .env."
    )
    exit()

# Inicializa la API con las credenciales cargadas
greenAPI = API.GreenApi(ID_INSTANCE, API_TOKEN_INSTANCE)

def verificar_estado_green_api():
    """Verifica el estado de la instancia de Green API. Devuelve el resultado crudo."""
    try:
        url = f"https://{GREEN_API_HOST}/waInstance{ID_INSTANCE}/getStateInstance/{API_TOKEN_INSTANCE}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return True, response.json().get("stateInstance", "unknown")
        else:
            return False, "error_http"
    except requests.RequestException:
        return False, "connection_error"
    except Exception:
        return False, "unknown_error"

def abrir_panel_green_api():
    """Abre el panel de control de Green API en el navegador web."""
    webbrowser.open(f"https://console.green-api.com/instanceList/{ID_INSTANCE}")

def cerrar_sesion_whatsapp_gui():
    """Intenta cerrar la sesión de WhatsApp a través de la API."""
    success, state = verificar_estado_green_api()
    if not success or state == "notAuthorized":
        return False, "No hay una sesión activa para cerrar."
    try:
        url = f"https://{GREEN_API_HOST}/waInstance{ID_INSTANCE}/logout/{API_TOKEN_INSTANCE}"
        response = requests.get(url, timeout=15)
        if response.status_code == 200 and response.json().get("isLoggedOut"):
            return True, "Sesión de WhatsApp cerrada correctamente."
        else:
            return False, f"El servidor no pudo cerrar la sesión. Código: {response.status_code}"
    except Exception as e:
        return False, f"Error de conexión al intentar cerrar sesión: {e}"

def enviar_mensaje_con_imagenes_gui(lote, muestra, mensaje, image_paths):
    """Envía un mensaje con una o más imágenes al grupo de WhatsApp configurado."""
    try:
        enviados_count = 0
        if not image_paths:
            return False, {"error": "No se proporcionaron imágenes para enviar."}
            
        # Envía la primera imagen con el mensaje como caption
        response = greenAPI.sending.sendFileByUpload(
            GROUP_CHAT_ID, 
            image_paths[0], 
            os.path.basename(image_paths[0]), 
            caption=mensaje
        )
        if "idMessage" in response.data: 
            enviados_count += 1
        else: 
            raise Exception(f"Fallo al enviar la primera imagen: {response.error}")
        
        time.sleep(1)
        
        # Envía las imágenes restantes sin caption
        for image_path in image_paths[1:]:
            response = greenAPI.sending.sendFileByUpload(
                GROUP_CHAT_ID, 
                image_path, 
                os.path.basename(image_path)
            )
            if "idMessage" in response.data: 
                enviados_count += 1
            time.sleep(1)
            
        return True, {"enviados": enviados_count, "totales": len(image_paths)}
        
    except Exception as e:
        return False, {"error": str(e)}