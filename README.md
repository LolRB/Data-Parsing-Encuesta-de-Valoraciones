# 📊 Automatización de Exportación de Estatus y Calificaciones de Entregables de Moodle a Google Sheets
Este proyecto permite automatizar la extracción de calificaciones de un curso en Moodle y exportarlas a una hoja de cálculo en Google Sheets. Está diseñado específicamente para plataformas Moodle como `https://prodep.capacitacioncontinua.mx`.

## 🚀 Características

- 🔐 **Inicio de sesión automático en Moodle** (como `https://prodep.capacitacioncontinua.mx`) con credenciales seguras desde archivo `.env`.

- 🧠 **Lectura inteligente de múltiples entregables** desde variables configurables (`DELIVERABLE_IDS` y `DELIVERABLE_LABELS`).

- 📥 **Obtención de datos completos por entregable:**

    - Nombre del usuario

    - Correo electrónico

    - Estatus de la entrega (Entregado, Borrador, Sin entrega)

    - Calificación obtenida

    - Fecha y hora de la última modificación

- 📊 **Unificación de los datos en una única tabla,** sin duplicar usuarios, incluso si los entregables están en diferente orden.

- 📤 **Exportación directa a Google Sheets,** formateado desde la celda B1.

- 🕒 **Registro automático de la fecha de ejecución** en A1 y en una hoja adicional llamada "Historial".

- 🔁 **Manejo de errores con reintentos automáticos:** si hay una desconexión temporal al consultar un entregable, el script intenta hasta 3 veces antes de continuar.

- ⚙️ **Configuración sencilla y segura** mediante archivo `.env`.

- ⏰ **Compatible con automatización** mediante `Programador de Tareas` (Windows) o `cron` (Linux/macOS).

## 📂 Estructura del proyecto

```
├── app.py              # Script principal 
├── .env                # Variables de entorno sensibles (NO subir al repositorio)
├── .env.example        # Archivo de ejemplo para configurar las variables de entorno
├── credentials.json    # Clave de servicio de Google (añadir al .gitignore)
├── README.md           # Documentación del proyecto
├── .gitignore          # Archivo para excluir archivos sensibles o irrelevantes del repositorio
```

## 🛡️ Recomendaciones de seguridad

### Para proteger tus credenciales y entorno de desarrollo:

- Nunca subas tu archivo `.env` ni `credentials.json` al repositorio.

- Usa un archivo `.env.example` para compartir la estructura de las variables necesarias sin exponer datos sensibles.

- Asegúrate de incluir los siguientes archivos en tu archivo `.gitignore`:

```
# Archivos sensibles:
.env
# |->Variables de entorno
credentials.json
# |->Google credentials

# Entorno virtual de Python
venv/
__pycache__/

# Archivos del sistema (OS junk)
.DS_Store
Thumbs.db

# Archivos de configuración de editores
.vscode/
.idea/
```

Esto evitará que información confidencial sea accidentalmente publicada o compartida.

## 🔎 Requisitos

- Python 3.8 o superior

- Cuenta de servicio de Google Cloud y archivo `credentials.json` con permisos de Sheets y Drive

- Acceso al curso en Moodle con credenciales válidas

## 🔧 Instalación

### 1. Clona el repositorio

Para clonar este repositorio, asegúrate de tener acceso autorizado en GitHub.

- SSH (recomendado) si tienes configurada tu clave SSH:

```bash
git clone git@github.com:LolRB/Data-Parsing-Reporte-Final.git
cd Data-Parsing-Reporte-Final
```

- HTTPS (te pedirá usuario y contraseña o token personal):

```bash
git clone https://github.com/LolRB/Data-Parsing-Reporte-Final.git
cd Data-Parsing-Reporte-Final
```

🔒 Nota: Si usas HTTPS, GitHub puede solicitar un token de acceso personal en lugar de tu contraseña.

### 2. Crea y activa un entorno virtual (opcional pero recomendado):

```bash
python -m venv venv
venv\Scripts\activate  # Windows
# ó
source venv/bin/activate  # macOS/Linux
```

### 3. Instala las dependencias:

```bash
pip install requests beautifulsoup4 gspread google-auth python-dotenv
```

## 📄 Google Sheets API Setup

1. Entra a **Google Cloud Console**

2. Crea un nuevo proyecto y habilita:

    - **Google Sheets API**

    - **Google Drive API**

3. Cree una cuenta de servicio, genere una clave **JSON** y descargue el archivo `.json`.

4. Guarde el archivo como `credentials.json` en la raíz del proyecto.

5. Comparta su hoja de cálculo de Google de destino con el correo electrónico de la cuenta de servicio (que se encuentra en el archivo JSON).

## ✏️ Configuración del archivo `.env`

Este proyecto utiliza variables de entorno para manejar credenciales y parámetros de forma segura. Antes de ejecutar el script, crea un archivo `.env` en la raíz del proyecto siguiendo el formato de `.env.example`.

### 1. Copia el archivo de ejemplo:
```bash
cp .env.example .env
```

### 2. Edita el archivo `.env` y reemplaza los valores con tus datos:

|   Variable                  |   Descripción                                                                |
|:----------------------------|:-----------------------------------------------------------------------------|
|   `USERNAME`                |   Usuario de Moodle                                                          |
|   `PASSWORD`                |   Contraseña del usuario en Moodle                                           |
|   `COURSE_ID`               |   ID numérico del curso en Moodle                                            |
|   `SPREADSHEET_NAME`        |   Nombre de tu hoja de cálculo en Google Sheets                              |
|   `WORKSHEET_NAME`          |   Nombre de la pestaña donde se exportarán los datos                         |
|   `GOOGLE_CREDENTIALS_FILE` |   Nombre del archivo .json con las credenciales de la cuenta de servicio     |
|   `DELIVERABLE_IDS`         |   Comas separadas con los IDs de los entregables (parte de la URL en Moodle)   |
|   `DELIVERABLE_LABELS`      |   Comas separadas con nombres legibles para los entregables                    |

🧠 Ejemplo de entregables:

```.env
DELIVERABLE_IDS=842,843,844
DELIVERABLE_LABELS=Entregable 1,Entregable 2,Entregable 3
```
⚠️ **Importante:** Asegúrate de que `DELIVERABLE_IDS` y `DELIVERABLE_LABELS` tengan **la misma cantidad de elementos y en el mismo orden**, ya que se asocian entre sí directamente.

### 🔒 Seguridad
No subas tu archivo `.env` ni `credentials.json` a ningún repositorio público. Añádelos a tu archivo `.gitignore`:

```gitignore
.env
credentials.json
```

## ▶️ Ejecuta el Script

### Lanza el script con:

```bash
python app.py
```
### ¿Qué hace en cada ejecución?

- 🔐 Inicia sesión automáticamente en Moodle con las credenciales del archivo `.env`.

- 📥 Visita las páginas de **cada entregable configurado** en `DELIVERABLE_IDS`.

- 🔎 Extrae los siguientes datos por usuario y por entregable:

    - Nombre completo

    - Correo electrónico

    - Estatus del entregable (Entregado, Borrador, Sin entrega)

    - Calificación obtenida (si aplica)

    - Última modificación (fecha y hora de entrega)

- 🧩 Combina y alinea la información por usuario, evitando duplicados aunque los entregables estén en diferente orden.

- 💾 Limpia y actualiza la hoja de Google Sheets:

    - Agrega un timestamp en la celda **A1**

    - Coloca la tabla de datos a partir de **B1**

- 🕒 Registra cada ejecución en una hoja adicional llamada **"Historial"**.

- 🔁 Si ocurre una desconexión o error temporal al acceder a Moodle, el script realiza **hasta 3 reintentos automáticos** antes de continuar.

## 🕒 Automatización (opcional)

Puedes usar:

- 🪟 Windows: Usa el Programador de tareas con un `.bat`.. que ejecute el script.

- 🐧 Linux/macOS:Usa `.cron`. para lanzar el script con un `.sh`.

## 🛠 Tecnologías utilizadas

- Python 3.x

- Requests (peticiones HTTP)

- BeautifulSoup (parseo HTML)

- gspread + Google API (acceso a hojas de cálculo)

- dotenv (variables de entorno)

## 📌 Notas

- Este script fue probado en plataformas Moodle personalizadas, por lo que podrían requerirse ajustes si cambia la estructura HTML.

- El `verify=False` está activo para ignorar advertencias de certificados SSL. Se recomienda desactivarlo si cuentas con certificados válidos.

## 🧑‍💻 Author

Proyecto desarrollado por **Rodrigo Bueno**.

Para dudas o mejoras, contáctame por correo:

📧 [ztmsiul79@gmail.com](mailto:ztmsiul79@gmail.com).

🐈‍⬛ [Github](https://github.com/LolRB).