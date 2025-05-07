"""
Este script automatiza el proceso de extracción de respuestas de la encuesta de valoración 
de las acciones de formación de los cursos en Moodle. 

El flujo de trabajo es el siguiente:
1. Inicia sesión en Moodle utilizando las credenciales proporcionadas.
2. Extrae los datos de los usuarios (nombre completo, email) que completaron la encuesta, a través de la llamada a la API de Moodle.
3. Descarga el archivo CSV de las respuestas de la encuesta utilizando Selenium para interactuar con la página de Moodle y hacer clic en el botón de descarga.
4. Procesa el contenido del archivo CSV descargado, extrayendo los datos relevantes (nombre, email y respuestas a las preguntas).
5. Combina los datos de los usuarios con sus respuestas a la encuesta.
6. Sube los datos combinados a una hoja de Google Sheets para su análisis y visualización.

Este script se conecta a la plataforma Moodle, realiza la autenticación, obtiene los datos de los usuarios,
descarga las respuestas de la encuesta y las carga en Google Sheets utilizando la API de Google.

Dependencias necesarias:
- `requests`: para interactuar con la API de Moodle.
- `selenium`: para automatizar la descarga del archivo CSV desde la interfaz web de Moodle.
- `gspread`: para interactuar con Google Sheets.
- `google.oauth2`: para la autenticación con la API de Google.
- `dotenv`: para cargar las variables de configuración desde un archivo `.env`.
"""

import re
import os
import time
from datetime import datetime
import requests
from dotenv import load_dotenv
import gspread
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from google.oauth2.service_account import Credentials

# ===============================
# 1. Cargar configuración desde .env
# ===============================
load_dotenv()
USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")
COURSE_ID = int(os.getenv("COURSE_ID"))
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME")
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME")
GOOGLE_CRED_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE")
SURVEY_ASSIGN_ID = os.getenv("SURVEY_ASSIGN_ID")
SURVEY_CSV_PARAM = os.getenv("SURVEY_CSV_PARAM", "")

HEADERS = {"User-Agent": "Mozilla/5.0"}


def get_selenium_driver():
    """
    Configura y devuelve una instancia de WebDriver para Chrome con opciones de ejecución en segundo plano.
    Este método configura el WebDriver de Selenium para usar Google Chrome en modo 'headless', es decir, sin abrir una ventana del navegador, lo que permite ejecutar el script en entornos sin interfaz gráfica. 
    El WebDriver es configurado con las opciones necesarias para la ejecución en segundo plano y luego se devuelve para ser utilizado en otras funciones.

    Returns:
        webdriver.Chrome: Instancia del WebDriver de Chrome configurada para ejecutar en segundo plano.
    """
    chrome_options = Options()
    # Ejecutar en segundo plano sin abrir el navegador
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)
    return driver


def download_survey_with_selenium():
    """
    Utiliza Selenium para abrir el navegador, autenticar, y hacer clic en el botón de descarga
    para obtener el archivo CSV de la encuesta.
    """
    # Inicializar WebDriver (asegúrate de tener el driver para tu navegador)
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    driver = webdriver.Chrome(service=Service(
        executable_path="C:/chromedriver-win64/chromedriver-win64/chromedriver.exe"), options=options)

    # Accede a la página de la encuesta
    driver.get(
        f"https://prodep.capacitacioncontinua.mx/mod/feedback/show_entries.php?id={SURVEY_ASSIGN_ID}")
    time.sleep(3)  # Esperar que la página cargue

    # Iniciar sesión con las credenciales
    driver.find_element(By.NAME, "username").send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.ID, "loginbtn").click()

    time.sleep(3)  # Esperar a que el login se complete

    # Esperar que la página cargue completamente después del login
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//form[contains(@class, "dataformatselector")]')))

    # Identificar el formulario para la descarga usando su id dinámico
    form_id = driver.find_element(
        By.XPATH, '//form[contains(@class, "dataformatselector")]')

    # Esperar que la opción por defecto (CSV) esté seleccionada. Si no, seleccionamos CSV manualmente
    select_element = form_id.find_element(By.ID, "downloadtype_download")
    current_value = select_element.get_attribute('value')

    if current_value != 'csv':
        select_element.click()
        option_csv = select_element.find_element(
            By.XPATH, "//option[@value='csv']")
        option_csv.click()
        time.sleep(2)  # Esperar a que se cambie la opción a CSV

    # Ahora hacer clic en el botón de descarga (usando el id dinámico)
    download_button = form_id.find_element(
        By.XPATH, '//button[@type="submit"]')
    download_button.click()

    # Esperar a que la descarga se complete (ajustar según sea necesario)
    time.sleep(5)

    # Path del archivo descargado (asegúrate de que el archivo se guarde en esta ubicación)
    download_path = r"C:\Users\rbueno\Downloads\Encuesta de valoración de las acciones de formación 2024.csv"

    # Verificar que el archivo existe antes de leerlo
    if not os.path.exists(download_path):
        print(
            f"Error: El archivo no fue encontrado en {download_path}. Asegúrate de que la descarga se haya completado correctamente.")
        driver.quit()
        return None

    # Leer el contenido del archivo descargado
    with open(download_path, "r", encoding="utf-8") as file:
        csv_text = file.read()

    driver.quit()  # Cerrar el navegador

    return csv_text


def login() -> requests.Session:
    """
    Esta función autentica al usuario en Moodle, obtiene las cookies necesarias para la sesión,
    y extrae el sesskey necesario para realizar las llamadas AJAX posteriores.
    Devuelve un objeto `session` que puede usarse para realizar las solicitudes posteriores.
    """
    session = requests.Session()

    # 1. Obtener página de login para extraer logintoken
    res = session.get(
        "https://prodep.capacitacioncontinua.mx/login/index.php", headers=HEADERS, verify=False)
    res.raise_for_status()
    token = re.search(r'name="logintoken" value="(\w+)"', res.text)
    if not token:
        raise RuntimeError("No se pudo obtener logintoken")
    logintoken = token.group(1)

    # 2. Enviar credenciales para iniciar sesión
    payload = {
        "username": USERNAME,
        "password": PASSWORD,
        "anchor": "",
        "logintoken": logintoken
    }
    res = session.post("https://prodep.capacitacioncontinua.mx/login/index.php",
                       data=payload, headers=HEADERS, allow_redirects=False, verify=False)
    res.raise_for_status()

    # 3. Forzar visita al dashboard para estabilizar la sesión y extraer sesskey
    res = session.get("https://prodep.capacitacioncontinua.mx/my/",
                      headers=HEADERS, verify=False)
    res.raise_for_status()
    m = re.search(r'sesskey["\']*\s*:?\s*["\']([a-zA-Z0-9]+)', res.text)
    if not m:
        raise RuntimeError("No se encontró sesskey tras login")
    session.sesskey = m.group(1)

    return session


def get_all_users(session: requests.Session) -> dict:
    """
    Esta función obtiene todos los usuarios de un curso de Moodle mediante la llamada AJAX
    "gradereport_grader_get_users_in_report". 
    Devuelve un diccionario donde las claves son los correos electrónicos y los valores son
    un diccionario con el nombre completo y el correo electrónico del usuario.
    """
    ajax = "https://prodep.capacitacioncontinua.mx/lib/ajax/service.php"
    payload = [{
        "index": 0,
        "methodname": "gradereport_grader_get_users_in_report",
        "args": {"courseid": COURSE_ID}
    }]
    url = f"{ajax}?sesskey={session.sesskey}&info=gradereport_grader_get_users_in_report"
    res = session.post(url, json=payload, headers=HEADERS, verify=False)
    res.raise_for_status()
    data = res.json()[0]["data"].get("users", [])

    users = {}
    for u in data:
        email = u.get("email")
        if email:
            users[email] = {"name": u.get("fullname", ""), "email": email}

    return users


def parse_survey_csv(csv_text: str):
    """
    Esta función procesa el contenido del CSV de respuestas de la encuesta. 
    Extrae el nombre, correo electrónico, fecha y las respuestas a las preguntas de la encuesta.
    Devuelve una lista de diccionarios con los datos de los usuarios y sus respuestas.
    """
    lines = csv_text.strip().split("\n")
    headers = lines[0].split(",")  # Encabezados

    # Eliminar comillas dobles si es necesario
    headers = [header.strip().replace('"', '') for header in headers]

    print(f"Encabezados encontrados: {headers}")

    try:
        # Buscamos la columna 'Dirección Email' y 'Fecha'
        idx_email = headers.index("Dirección Email")
        idx_date = headers.index("Fecha")  # Columna de fecha
    except ValueError:
        print("No se encontró la columna 'Dirección Email' o 'Fecha' en el archivo CSV.")
        return []  # Si no se encuentra alguna de las columnas, retornamos una lista vacía

    survey_data = []

    # Procesar las filas del CSV (saltando la primera fila de encabezado)
    for line in lines[1:]:
        data = [field.strip().replace('"', '') for field in line.split(",")]

        email = data[idx_email].strip()  # Extraemos el email
        name = data[0].strip()  # Suponemos que la primera columna es el nombre

        # Combinamos las celdas de la fecha (día, fecha, hora) en una sola celda
        if len(data) > idx_date + 2:  # Aseguramos que hay 3 celdas para la fecha
            # Combinamos el día, fecha y hora
            date_combined = f"{data[idx_date]} {data[idx_date + 1]} {data[idx_date + 2]}"
            # Reemplazar las celdas de la fecha con la celda combinada
            data[idx_date] = date_combined
            data.pop(idx_date + 1)  # Eliminamos la celda duplicada de la fecha
            data.pop(idx_date + 1)  # Eliminamos la celda duplicada de la hora

        # Las respuestas a las preguntas están en las demás columnas después del nombre y email
        answers = data[1:idx_date] + data[idx_date + 1:]

        # Almacenamos la fecha combinada para cada usuario
        survey_data.append({
            "name": name,
            "email": email,
            "answers": answers,
            "date": data[idx_date]  # Agregar la fecha combinada
        })

    return survey_data


def merge_data(all_users: dict, survey: list) -> list:
    """
    Combina los usuarios con las respuestas de la encuesta.
    Devuelve una lista con las filas para agregar a Google Sheets:
    1. Nombre completo
    2. Email
    3. Respuestas a las preguntas (según lo obtenido en el CSV)
    """
    # Cabecera: nombre, email + preguntas
    if not survey:
        return []  # Si survey está vacío, retornar una lista vacía.

    # Obtener las preguntas (basado en la primera fila de respuestas)
    sample_qs = [f"Pregunta {i+1}" for i in range(len(survey[0]["answers"]))]

    header = ["Nombre completo", "Email", "Fecha", "Grupo"] + sample_qs
    table = [header]

    # Recorrer todos los usuarios
    for user in all_users.values():
        # Incluye el email en una nueva celda "Email de Encuestados"
        row = [user["name"], user["email"]]

        # Buscar las respuestas de este usuario
        answers = next(
            (entry["answers"] for entry in survey if entry["email"] == user["email"]), [])

        # Agregar la fecha combinada de este usuario en la fila
        date = next(
            (entry["date"] for entry in survey if entry["email"] == user["email"]), "")
        row.append(date)  # Fecha combinada de este usuario

        row.extend(answers)  # Añadir las respuestas a la fila
        table.append(row)

    return table


def upload_to_sheets(table: list):
    """
    Esta función sube los datos procesados a Google Sheets.
    Limpiará la hoja, agregará el timestamp en A1, y escribirá los datos desde B1.
    También añadirá un registro en la hoja 'Historial'.
    """
    creds = Credentials.from_service_account_file(GOOGLE_CRED_FILE, scopes=[
                                                  "https://www.googleapis.com/auth/spreadsheets",
                                                  "https://www.googleapis.com/auth/drive"])
    client = gspread.authorize(creds)
    sh = client.open(SPREADSHEET_NAME)

    ws = sh.worksheet(WORKSHEET_NAME)
    ws.clear()  # Limpiar hoja

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.update("A1", [[f"Actualizado el: {ts}"]])
    ws.update("B1", table)

    # Historial
    ts_log = ["Ejecución registrada el:", ts]
    try:
        hist = sh.worksheet("Historial")
    except gspread.exceptions.WorksheetNotFound:
        hist = sh.add_worksheet(title="Historial", rows="100", cols="2")
    hist.append_row(ts_log)


def main():
    """
    Función principal que coordina todo el flujo de datos:
    1. Inicia sesión en Moodle.
    2. Obtiene todos los usuarios del curso.
    3. Descarga y procesa los datos de la encuesta.
    4. Combina las respuestas con los usuarios.
    5. Sube los datos a Google Sheets.
    """
    session = login()
    users = get_all_users(session)
    csv_text = download_survey_with_selenium()  # Usar Selenium para obtener el CSV
    survey = parse_survey_csv(csv_text)
    table_data = merge_data(users, survey)
    upload_to_sheets(table_data)
    print("✅ Encuesta exportada correctamente a Google Sheets.")


if __name__ == "__main__":
    main()
