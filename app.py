"""
Script para obtener el estado y calificaciones de múltiples entregables de un curso en Moodle,
y exportarlos a Google Sheets de manera organizada.
"""

import os
import re
import time
from datetime import datetime
from requests.exceptions import RequestException
import gspread
import requests
from bs4 import BeautifulSoup
# IGNORAR IMPORT ERROR. Bug de Pylint. Funciona aunque marca error.
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

# ===============================
# 1. Cargar configuraciones .env
# ===============================
load_dotenv()

USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")
COURSE_ID = int(os.getenv("COURSE_ID"))
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME")
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME")
CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE")

# Leer múltiples entregables desde .env
ids_raw = os.getenv("DELIVERABLE_IDS", "")
labels_raw = os.getenv("DELIVERABLE_LABELS", "")

# Convertir a listas eliminando espacios
ENTREGABLE_IDS = [e.strip() for e in ids_raw.split(",") if e.strip().isdigit()]
ENTREGABLE_LABELS = [l.strip() for l in labels_raw.split(",") if l.strip()]

# Validación de longitud
if len(ENTREGABLE_IDS) != len(ENTREGABLE_LABELS):
    raise ValueError(
        "DELIVERABLE_IDS y DELIVERABLE_LABELS deben tener la misma cantidad de elementos.")

# ==============================
# 2. Establecer sesión en Moodle
# ==============================
HEADERS = {"User-Agent": "Mozilla/5.0"}
session = requests.Session()

# Obtener token de login
LOGIN_URL = "https://prodep.capacitacioncontinua.mx/login/index.php"
res = session.get(LOGIN_URL, headers=HEADERS, verify=False)
res.raise_for_status()
match = re.search(r'name="logintoken" value="(\w+)"', res.text)
if not match:
    raise ValueError(
        "No se encontró el logintoken en la página de inicio de sesión.")
logintoken = match.group(1)

# Enviar login
res = session.post(LOGIN_URL, data={
    "username": USERNAME,
    "password": PASSWORD,
    "anchor": "",
    "logintoken": logintoken
}, headers=HEADERS, verify=False, allow_redirects=False)
res.raise_for_status()

# Visitar página principal para estabilizar sesión
res = session.get("https://prodep.capacitacioncontinua.mx/my/",
                  headers=HEADERS, verify=False)
res.raise_for_status()

# Obtener sesskey para llamadas AJAX
sesskey_match = re.search(
    r'sesskey["\']*\s*:?\s*["\']([a-zA-Z0-9]+)', res.text)
if not sesskey_match:
    raise ValueError("No se encontró el sesskey tras el login.")
sesskey = sesskey_match.group(1)

# ===============================
# 3. Obtener usuarios del curso
# ===============================
AJAX_URL = "https://prodep.capacitacioncontinua.mx/lib/ajax/service.php"
payload_users = [{
    "index": 0,
    "methodname": "gradereport_grader_get_users_in_report",
    "args": {"courseid": COURSE_ID}
}]
resp = session.post(f"{AJAX_URL}?sesskey={sesskey}&info=gradereport_grader_get_users_in_report",
                    json=payload_users, headers=HEADERS, verify=False)
resp.raise_for_status()

user_list = []
user_info = {}
data_users = resp.json()[0].get("data", {}).get("users", [])
for u in data_users:
    uid = u.get("id")
    name = u.get("fullname")
    email = u.get("email")
    if uid and email:
        user_list.append(uid)
        user_info[email] = {"name": name, "email": email}

# ==================================================
# 4. Obtener datos de cada entregable por separado
# ==================================================
status_data = {}

for entregable_id, label in zip(ENTREGABLE_IDS, ENTREGABLE_LABELS):
    URL = f"https://prodep.capacitacioncontinua.mx/mod/assign/view.php?id={entregable_id}&action=grading"

    print(f"🔄 Procesando entregable: {label} (ID {entregable_id})")

    # Reintentos automáticos
    SUCCESS = False
    for attempt in range(1, 4):  # 3 intentos máximo
        try:
            time.sleep(2)  # Esperar 2 segundos para no saturar el servidor
            res = session.get(URL, headers=HEADERS, verify=False, timeout=15)
            res.raise_for_status()
            SUCCESS = True
            break
        except RequestException as e:
            print(
                f"⏳ Intento {attempt} fallido para {label} (ID {entregable_id}): {e}")
            if attempt < 3:
                print("🔁 Reintentando...")
            else:
                print(
                    f"❌ Falló definitivamente al intentar acceder a {label}. Pasando al siguiente.")

    if not SUCCESS:
        continue

    soup = BeautifulSoup(res.text, "html.parser")
    table = soup.find("table", class_="generaltable")
    if not table:
        print(
            f"⚠️ Advertencia: No se encontró la tabla de entregas para el entregable {label} (ID {entregable_id}).")
        continue

    for row in table.select("tbody tr"):
        try:
            name_cell = row.select_one("td.c2 a")
            email_cell = row.select_one("td.c3.email")
            status_cell = row.select_one("td.c4")
            grade_cell = row.select_one("td.c5")
            modified_cell = row.select_one("td.c7")

            if not name_cell or not email_cell:
                continue

            full_name = name_cell.text.strip()
            email = email_cell.text.strip()

            # Asegurar entrada única por email
            if email not in status_data:
                status_data[email] = {
                    "name": full_name,
                    "email": email
                }

            # Obtener estatus
            divs_status = status_cell.select("div")
            STATUS_TEXT = "Sin entrega"
            for div in divs_status:
                class_list = div.get("class", [])
                if "submissionstatusdraft" in class_list:
                    STATUS_TEXT = "Borrador"
                elif "submissionstatussubmitted" in class_list:
                    STATUS_TEXT = "Entregado"
                elif "submissionstatus" in class_list:
                    STATUS_TEXT = "Sin entrega"
                if "submissiongraded" in class_list:
                    STATUS_TEXT += " y Calificado"

            # Calificación
            GRADE_TEXT = ""
            if grade_cell:
                for part in grade_cell.stripped_strings:
                    if "/" in part:
                        GRADE_TEXT = part.strip()
                        break

            # Última modificación
            modified = modified_cell.text.strip() if modified_cell else ""

            # Guardar con nombre del entregable
            status_data[email][f"{label}_status"] = STATUS_TEXT
            status_data[email][f"{label}_grade"] = GRADE_TEXT
            status_data[email][f"{label}_modified"] = modified

        except ValueError as e:
            print(f"⚠️ Error procesando fila de {label}: {e}")

# =====================================
# 5. Preparar datos para Google Sheets
# =====================================
headers = ["Nombre completo", "Email"]
for label in ENTREGABLE_LABELS:
    headers += [
        f"{label} Estatus",
        f"{label} Calificación",
        f"{label} Última modificación"
    ]

table_data = [headers]
for email, data in status_data.items():
    row = [data.get("name", ""), email]
    for label in ENTREGABLE_LABELS:
        row += [
            data.get(f"{label}_status", ""),
            data.get(f"{label}_grade", ""),
            data.get(f"{label}_modified", "")
        ]
    table_data.append(row)

# =======================================
# 6. Subir resultados a Google Sheets
# =======================================
if not CREDENTIALS_FILE:
    raise ValueError("GOOGLE_CREDENTIALS_FILE no especificado en .env")

creds = Credentials.from_service_account_file(
    CREDENTIALS_FILE,
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
)

client = gspread.authorize(creds)
sheet = client.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)

# Verificar que hay datos válidos
if len(table_data) <= 1:
    raise ValueError(
        "No se generaron filas válidas para actualizar la hoja de cálculo.")

# Limpiar hoja y agregar timestamp en A1
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
sheet.clear()
sheet.update("A1", [[f"Actualizado el: {timestamp}"]])
sheet.update("B1", table_data)

# Registrar historial
timestamp_log = ["Ejecución registrada el:", timestamp]
try:
    log_sheet = client.open(SPREADSHEET_NAME).worksheet("Historial")
except gspread.exceptions.WorksheetNotFound:
    log_sheet = client.open(SPREADSHEET_NAME).add_worksheet(
        title="Historial", rows="100", cols="2")
log_sheet.append_row(timestamp_log)

print("✅ Datos exportados correctamente a Google Sheets.")
