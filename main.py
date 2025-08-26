import imaplib
import email
import requests
import base64
import time
import json
import os
from email.header import decode_header
from urllib.parse import quote
from dotenv import load_dotenv

# Cargar variables de entorno desde el archivo .env
load_dotenv()

# ===== CONFIGURACIÓN DESDE .env =====
IMAP_SERVER = os.getenv("IMAP_SERVER", "imap.gmail.com")
IMAP_USER = os.getenv("IMAP_USER")
IMAP_PASS = os.getenv("IMAP_PASS")
AZURE_ORG = os.getenv("AZURE_ORG")
AZURE_PROJECT = os.getenv("AZURE_PROJECT")
AZURE_PAT = os.getenv("AZURE_PAT")
AZURE_DEVOPS_SENDER = os.getenv("AZURE_DEVOPS_SENDER", "azuredevops@microsoft.com")
LOG_FILE = os.getenv("LOG_FILE", "azure_devops_mail_processor.log")

# Validar que las variables requeridas estén presentes
required_vars = ["IMAP_USER", "IMAP_PASS", "AZURE_ORG", "AZURE_PROJECT", "AZURE_PAT"]
missing_vars = [var for var in required_vars if not os.getenv(var)]

if missing_vars:
    raise ValueError(f"Faltan variables de entorno requeridas: {', '.join(missing_vars)}")

# ⚡ CONFIGURACIÓN DE TU TABLERO - BASADO EN TU CONFIGURACIÓN REAL
TABLERO_CONFIG = {
    # Mapeo de columnas a estados (según tu tablero "Issues")
    "mapeo_columnas_estados": {
        "Bugs creados": "To Do",  # Columna → Estado
        "En revision": "Doing",  # Columna → Estado
        "Ejecucion existosa": "Done"  # Columna → Estado
    },

    # Asignación automática según tipo de correo
    "asignacion_correos": {
        "failed": "Bugs creados",  # FAILED → Columna "Bugs creados"
        "succeeded": "Ejecucion existosa"  # SUCCEEDED → Columna "Ejecucion existosa"
    }
}


# ===== FUNCIONES PRINCIPALES =====
def log_message(message):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] {message}\n")
    print(f"[{timestamp}] {message}")


def get_work_item_types():
    """Obtiene los tipos de work items disponibles"""
    try:
        project_encoded = quote(AZURE_PROJECT)
        url = f"{AZURE_ORG}/{project_encoded}/_apis/wit/workitemtypes?api-version=6.0"

        headers = {
            "Authorization": "Basic " + base64.b64encode((":" + AZURE_PAT).encode()).decode()
        }

        response = requests.get(url, headers=headers, timeout=30)

        if response.status_code == 200:
            work_item_types = response.json()
            available_types = [wit['name'] for wit in work_item_types['value']]
            log_message(f"📋 Tipos de Work Items disponibles: {available_types}")
            return available_types
        else:
            log_message(f"❌ Error obteniendo tipos: {response.status_code}")
            return ["Issue", "Task"]
    except Exception as e:
        log_message(f"❌ Excepción al obtener tipos: {e}")
        return ["Issue", "Task"]


def get_work_item_states(work_item_type):
    """Obtiene los estados disponibles para un tipo de work item"""
    try:
        project_encoded = quote(AZURE_PROJECT)
        url = f"{AZURE_ORG}/{project_encoded}/_apis/wit/workitemtypes/{work_item_type}/states?api-version=6.0"

        headers = {
            "Authorization": "Basic " + base64.b64encode((":" + AZURE_PAT).encode()).decode()
        }

        response = requests.get(url, headers=headers, timeout=30)

        if response.status_code == 200:
            states = response.json()
            available_states = [state['name'] for state in states['value']]
            log_message(f"🎯 Estados disponibles para '{work_item_type}': {available_states}")
            return available_states
        else:
            log_message(f"❌ Error obteniendo estados: {response.status_code}")
            return ["To Do", "Doing", "Done"]
    except Exception as e:
        log_message(f"❌ Excepción al obtener estados: {e}")
        return ["To Do", "Doing", "Done"]


def create_work_item(title, work_item_type, target_column):
    """Crea un work item y lo asigna a la columna específica"""
    try:
        project_encoded = quote(AZURE_PROJECT)
        url = f"{AZURE_ORG}/{project_encoded}/_apis/wit/workitems/${work_item_type}?api-version=6.0"

        headers = {
            "Content-Type": "application/json-patch+json",
            "Authorization": "Basic " + base64.b64encode((":" + AZURE_PAT).encode()).decode()
        }

        # Obtener el estado correspondiente a la columna
        target_state = TABLERO_CONFIG["mapeo_columnas_estados"].get(target_column, "To Do")

        # Verificar que el estado existe
        available_states = get_work_item_states(work_item_type)
        if target_state not in available_states:
            log_message(f"⚠️ Estado '{target_state}' no disponible. Usando 'To Do'")
            target_state = "To Do" if "To Do" in available_states else available_states[0]

        # Payload para crear el work item
        payload = [
            {"op": "add", "path": "/fields/System.Title", "value": title},
            {"op": "add", "path": "/fields/System.Description",
             "value": f"Creado automáticamente desde correo de Azure DevOps"},
            {"op": "add", "path": "/fields/System.State", "value": target_state},
            {"op": "add", "path": "/fields/System.Tags", "value": "Auto-Generated"}
        ]

        response = requests.post(url, headers=headers, json=payload, timeout=30)

        if response.status_code in [200, 201]:
            work_item_id = response.json().get('id', 'N/A')
            work_item_url = f"{AZURE_ORG}/{AZURE_PROJECT}/_workitems/edit/{work_item_id}"
            log_message(f"✅ Work Item creado en columna '{target_column}': #{work_item_id}")
            log_message(f"   📌 Título: {title}")
            log_message(f"   🎯 Estado: {target_state}")
            log_message(f"   🔗 URL: {work_item_url}")
            return True
        else:
            log_message(f"❌ Error creando work item: {response.status_code} - {response.text}")
            return False

    except Exception as e:
        log_message(f"❌ Excepción al crear work item: {e}")
        return False


def decode_subject(encoded_subject):
    """Decodifica el asunto del correo"""
    try:
        decoded_parts = decode_header(encoded_subject)
        subject = ""
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                if encoding:
                    subject += part.decode(encoding)
                else:
                    subject += part.decode('utf-8', errors='ignore')
            else:
                subject += part
        return subject
    except:
        return str(encoded_subject)


def process_mail(mail, msg_id):
    """Procesa un correo de Azure DevOps y crea el work item correspondiente"""
    try:
        status, data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK":
            return

        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)
        subject = msg["subject"]
        decoded_subject = decode_subject(subject)

        log_message(f"📧 Procesando correo: {decoded_subject}")

        # Marcar como leído
        mail.store(msg_id, '+FLAGS', '\\Seen')

        # Determinar tipo de work item y columna destino
        available_types = get_work_item_types()

        if "failed" in decoded_subject.lower():
            # FAILED → Issue en columna "Bugs creados"
            work_item_type = "Issue" if "Issue" in available_types else available_types[0]
            target_column = "Bugs creados"
            title = f"❌ Error en pipeline: {decoded_subject}"

        elif "succeeded" in decoded_subject.lower():
            # SUCCEEDED → Issue en columna "Ejecucion existosa"
            work_item_type = "Issue" if "Issue" in available_types else available_types[0]
            target_column = "Ejecucion existosa"
            title = f"✅ Pipeline exitoso: {decoded_subject}"

        else:
            # Otros correos → Issue en primera columna disponible
            work_item_type = "Issue" if "Issue" in available_types else available_types[0]
            target_column = list(TABLERO_CONFIG["mapeo_columnas_estados"].keys())[0]
            title = f"📋 Notificación pipeline: {decoded_subject}"
            log_message(f"📨 Correo procesado (sin acción específica): {decoded_subject}")
            return

        # Crear el work item en la columna correspondiente
        success = create_work_item(title, work_item_type, target_column)

        if success:
            log_message(f"🎯 Work Item asignado a columna: {target_column}")
        else:
            log_message("❌ No se pudo crear el Work Item")

    except Exception as e:
        log_message(f"❌ Error procesando correo: {e}")


def connect_mail():
    """Conecta al servidor IMAP"""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(IMAP_USER, IMAP_PASS)
        mail.select("inbox")
        log_message("✅ Conexión IMAP exitosa")
        return mail
    except Exception as e:
        log_message(f"❌ Error conectando al servidor IMAP: {e}")
        return None


def check_azure_devops_mails(mail):
    """Busca correos no leídos de Azure DevOps"""
    try:
        search_criteria = f'(UNSEEN FROM "{AZURE_DEVOPS_SENDER}")'
        status, messages = mail.search(None, search_criteria)

        if status == "OK":
            return messages[0].split()
        else:
            log_message("❌ Error buscando correos")
            return []
    except Exception as e:
        log_message(f"❌ Error buscando correos de Azure DevOps: {e}")
        return []


def main_loop():
    """Bucle principal del script"""
    log_message("🚀 Iniciando procesador de correos para TU tablero")
    log_message("🔍 Obteniendo información de tu configuración...")

    # Obtener información real de tu proyecto
    available_types = get_work_item_types()

    log_message("🎯 Configuración actual del script:")
    log_message(f"   - Mapeo columnas→estados: {TABLERO_CONFIG['mapeo_columnas_estados']}")
    log_message(f"   - Asignación correos→columnas: {TABLERO_CONFIG['asignacion_correos']}")
    log_message("✅ Configuración verificada. Iniciando monitoreo de correos...")

    while True:
        try:
            mail = connect_mail()
            if mail:
                azure_mails = check_azure_devops_mails(mail)

                if azure_mails:
                    log_message(f"📬 Encontrados {len(azure_mails)} correos nuevos de Azure DevOps")
                    for msg_id in azure_mails:
                        process_mail(mail, msg_id)
                else:
                    log_message("📭 No hay correos nuevos de Azure DevOps")

                mail.close()
                mail.logout()

            time.sleep(60)  # Esperar 1 minuto

        except Exception as e:
            log_message(f"❌ Error en el bucle principal: {e}")
            time.sleep(60)


if __name__ == "__main__":
    main_loop()