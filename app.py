"""
Este script automatiza el proceso de extracción de respuestas de la encuesta de valoración
de las acciones de formación de los cursos en Moodle.

El flujo de trabajo es el siguiente:
1. Inicia sesión en Moodle utilizando las credenciales proporcionadas.
2. Extrae los datos de los usuarios (nombre completo, email) que completaron la encuesta,
   a través de la llamada a la API de Moodle.
3. Descarga el archivo CSV de las respuestas de la encuesta utilizando Selenium para interactuar
   con la página de Moodle y hacer clic en el botón de descarga.
4. Procesa el contenido del archivo CSV descargado, extrayendo los datos relevantes (nombre, email
   y respuestas a las preguntas).
5. Combina los datos de los usuarios con sus respuestas a la encuesta.
6. Sube los datos combinados a una hoja de Google Sheets para su análisis y visualización.

El script se conecta a la plataforma Moodle, realiza autenticación, obtiene datos de los usuarios,
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
import csv
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

# Ignorar error E0401 de Pylint. Programa corre correctamente

# ===============================
# 1. Cargar configuración desde .env
# ===============================

load_dotenv()  # Cargar variables de entorno desde el archivo .env

# Verificar que todas las variables de entorno necesarias están presentes
required_env_vars = ["USERNAME", "PASSWORD", "COURSE_ID", "SPREADSHEET_NAME",
                     "WORKSHEET_NAME", "GOOGLE_CREDENTIALS_FILE", "SURVEY_ID"]
for var in required_env_vars:
    if not os.getenv(var):
        raise EnvironmentError(
            f"La variable de entorno {var} no está definida.")

# Asignar las variables de entorno a las variables locales
USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")
COURSE_ID = int(os.getenv("COURSE_ID"))
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME")
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME")
GOOGLE_CRED_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE")
SURVEY_ID = os.getenv("SURVEY_ID")

HEADERS = {"User-Agent": "Mozilla/5.0"}

# ==================================================================
# Función para configurar el WebDriver de Selenium en modo headless
# ==================================================================


def get_selenium_driver():
    """
    Configura y devuelve una instancia de WebDriver para Chrome en modo 'headless'.
    'headless' significa que el navegador se ejecutará sin abrir una ventana visible.
    Esto es útil para ejecutar el script en entornos sin interfaz gráfica.

    Returns:
        webdriver.Chrome: Instancia del WebDriver de Chrome
        configurada para ejecutar en segundo plano.
    """
    chrome_options = Options()
    # Ejecutar en segundo plano sin abrir el navegador
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)
    return driver


def download_survey_with_selenium():
    """
    Utiliza Selenium para abrir el navegador, autenticarse en Moodle
    y descargar el archivo CSV con las respuestas.

    Retorna:
        str: El contenido del archivo CSV descargado.
    """
    try:
        # Configuración del WebDriver en segundo plano (sin interfaz gráfica)
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")

        # Inicializamos el WebDriver con el path al ejecutable de ChromeDriver
        driver = webdriver.Chrome(service=Service(
            executable_path="C:/chromedriver-win64/chromedriver-win64/chromedriver.exe"),
            options=options)
        driver.get(
            f"https://prodep.capacitacioncontinua.mx/mod/feedback/show_entries.php?id={SURVEY_ID}")
        time.sleep(3)  # Esperar que la página cargue completamente

        # Iniciar sesión con las credenciales de usuario
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "loginbtn").click()
        time.sleep(3)  # Esperar a que el login se complete

        # Esperar a que la página cargue completamente después del login
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//form[contains(@class, "dataformatselector")]'))
        )

        form_id = driver.find_element(
            By.XPATH, '//form[contains(@class, "dataformatselector")]')
        select_element = form_id.find_element(By.ID, "downloadtype_download")
        current_value = select_element.get_attribute('value')

        # Si la opción CSV no está seleccionada, la cambiamos a CSV
        if current_value != 'csv':
            select_element.click()
            option_csv = select_element.find_element(
                By.XPATH, "//option[@value='csv']")
            option_csv.click()
            time.sleep(2)  # Esperar a que se cambie la opción a CSV

        # Hacer clic en el botón de descarga
        download_button = form_id.find_element(
            By.XPATH, '//button[@type="submit"]')
        download_button.click()

        time.sleep(5)  # Esperar a que la descarga se complete

        # Ruta del archivo CSV descargado
        download_path = r"C:\Users\rbueno\Downloads\Encuesta de valoración de las acciones de formación 2024.csv"

        # Verificar que el archivo existe antes de leerlo
        if not os.path.exists(download_path):
            raise FileNotFoundError(
                f"Archivo no encontrado en {download_path}. Asegurar la descarga correctamente.")

        with open(download_path, "r", encoding="utf-8") as file:
            csv_text = file.read()

        driver.quit()  # Cerrar el navegador
        return csv_text
    except Exception as e:
        print(f"Error al descargar el archivo con Selenium: {e}")
        raise


def login() -> requests.Session:
    """
    Esta función autentica al usuario en Moodle, obtiene las cookies necesarias para la sesión,
    y extrae el sesskey necesario para realizar las llamadas AJAX posteriores.
    Devuelve un objeto `session` que puede usarse para realizar las solicitudes posteriores.

    Returns:
        requests.Session: Objeto de sesión autenticado para realizar solicitudes a Moodle.
    """
    session = requests.Session()

    # Obtener página de login para extraer logintoken
    res = session.get(
        "https://prodep.capacitacioncontinua.mx/login/index.php", headers=HEADERS, verify=False)
    res.raise_for_status()
    token = re.search(r'name="logintoken" value="(\w+)"', res.text)
    if not token:
        raise RuntimeError("No se pudo obtener logintoken")
    logintoken = token.group(1)

    # Enviar credenciales para iniciar sesión
    payload = {
        "username": USERNAME,
        "password": PASSWORD,
        "anchor": "",
        "logintoken": logintoken
    }
    res = session.post("https://prodep.capacitacioncontinua.mx/login/index.php",
                       data=payload, headers=HEADERS, allow_redirects=False, verify=False)
    res.raise_for_status()

    # Forzar visita al dashboard para estabilizar la sesión y extraer sesskey
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

    Returns:
        dict: Diccionario con los datos de los usuarios (correo electrónico, nombre completo).
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

    Args:
        csv_text (str): El contenido completo del archivo CSV con las respuestas de la encuesta.

    Returns:
        list: Lista de diccionarios con datos de los usuarios (nombre, email, respuestas, fecha).
    """
    lines = csv_text.strip().split("\n")
    reader = csv.reader(lines)

    headers = next(reader)  # Obtener los encabezados

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
    for data in reader:
        data = [field.strip().replace('"', '') for field in data]

        email = data[idx_email].strip()  # Extraemos el email
        name = data[0].strip()  # Suponemos que la primera columna es el nombre

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

    Args:
        all_users (dict): Diccionario con los datos de los usuarios (email, nombre).
        survey (list): Lista de diccionarios con las respuestas de la encuesta.

    Returns:
        list: Tabla de datos combinados (usuario, grupo, fecha, respuestas) para Google Sheets.
    """
    if not survey:
        return []  # Si survey está vacío, retornar una lista vacía.

    # Obtener las preguntas
    sample_qs = [f"Pregunta {i+1}" for i in range(len(survey[0]["answers"]))]

    # Crear la cabecera
    header = ["Nombre Completo", "Grupos", "Fecha de Encuesta",
              "Email de Encuestados"] + sample_qs
    table = [header]

    for user in all_users.values():
        row = [user["name"]]

        # Verificar si el usuario respondió la encuesta
        survey_entry = next(
            (entry for entry in survey if entry["email"] == user["email"]), None)

        if survey_entry:
            # Tomar el grupo del primer campo de la respuesta
            group = survey_entry["answers"][0]
            row.append(group)  # Agregar el grupo
        else:
            row.append("")  # Si no respondió, dejar vacío el grupo

        # Obtener la fecha combinada
        date = next(
            (entry["date"] for entry in survey if entry["email"] == user["email"]), "")
        row.append(date)  # Fecha combinada de este usuario

        # Respuestas del usuario
        answers = survey_entry["answers"] if survey_entry else []
        row.extend(answers[1:])  # Añadir las respuestas, omitiendo el correo

        table.append(row)

    return table


def upload_to_sheets(table: list):
    """
    Esta función sube los datos procesados a Google Sheets.
    Limpiará la hoja, agregará el timestamp en A1, y escribirá los datos desde B1.
    También añadirá un registro en la hoja 'Historial'.

    Args:
        table (list): Lista de datos para escribir en Google Sheets.
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
    try:
        session = login()
        users = get_all_users(session)
        csv_text = download_survey_with_selenium()  # Usar Selenium para obtener el CSV
        survey = parse_survey_csv(csv_text)
        table_data = merge_data(users, survey)
        upload_to_sheets(table_data)
        print("✅ Encuesta exportada correctamente a Google Sheets.")
    except Exception as e:
        print(f"Error en el flujo principal: {e}")
        raise


if __name__ == "__main__":
    main()
