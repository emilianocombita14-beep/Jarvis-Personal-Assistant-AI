import sys
import os
import subprocess
import webbrowser
import time
import ctypes
import win32gui
import win32con
import wmi
import pyttsx3
import whisper
import sounddevice as sd
import numpy as np
import wave
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QTextBrowser, QLineEdit, QPushButton, QHBoxLayout
from PySide6.QtGui import QFont, QColor, QPalette
import requests
import re
import pygame
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import os.path
from datetime import datetime, timedelta
import unicodedata
from word2number_es import w2n
import keyboard



BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)


def normalizar_comando(texto):
    texto = texto.lower().strip()
    if texto in ["apaga", "apaga el pc", "shutdown", "apaga el computador"]:
        return "apagar"
    elif texto in ["reinicia", "restart", "reiniciar el sistema"]:
        return "reiniciar"
    elif texto in ["bloquea", "lock", "bloquear pantalla"]:
        return "bloquear"
    else:
        return None


def normalizar_texto(texto):
    texto = texto.lower()

    correcciones = {
        "setiembre":"septiembre",
        "octube":"octubre",
        "dosiembre":"diciembre",
        "febreo":"febrero",
        "enero ":" enero ",
        "alas ":"a las ",
        "ala ":"a las "
    }

    for mal,bien in correcciones.items():
        texto = texto.replace(mal,bien)

    return texto


model = None

def cargar_whisper():
    global model
    if model is None:
        print("Cargando IA de voz...")
        model = whisper.load_model("small")
        print("✅ Whisper listo")



MESES = {
    "enero":"01","febrero":"02","marzo":"03","abril":"04",
    "mayo":"05","junio":"06","julio":"07","agosto":"08",
    "septiembre":"09","octubre":"10","noviembre":"11","diciembre":"12"
}

def convertir_fecha_natural(texto):
    patron = r'(\d{1,2}) de (\w+) del (\d{4})'
    match = re.search(patron, texto)

    if match:
        dia = match.group(1).zfill(2)
        mes_nombre = match.group(2)
        anio = match.group(3)

        if mes_nombre in MESES:
            mes = MESES[mes_nombre]
            fecha = f"{dia}/{mes}/{anio}"

            texto = texto.replace(match.group(0), fecha)

    return texto


def convertir_hora_natural(texto):
    patron = r'a las (\d{1,2})'
    match = re.search(patron, texto)

    if match:
        hora = match.group(1).zfill(2)
        texto = texto.replace(match.group(0), f"a las {hora}")

    return texto



NUMEROS = {
    "cero":0,"uno":1,"una":1,"dos":2,"tres":3,"cuatro":4,"cinco":5,"seis":6,
    "siete":7,"ocho":8,"nueve":9,"diez":10,"once":11,"doce":12,"trece":13,
    "catorce":14,"quince":15,"dieciseis":16,"dieciséis":16,"diecisiete":17,
    "dieciocho":18,"diecinueve":19,"veinte":20,"veintiuno":21,"veintidos":22,
    "veintitrés":23,"veinticuatro":24,"veinticinco":25,"treinta":30,
    "cuarenta":40,"cincuenta":50,"sesenta":60,"setenta":70,"ochenta":80,
    "noventa":90,"cien":100,"ciento":100,"doscientos":200,"trescientos":300,
    "cuatrocientos":400,"quinientos":500,"seiscientos":600,"setecientos":700,
    "ochocientos":800,"novecientos":900,"mil":1000,"millón":1000000,
    "millon":1000000,"millones":1000000
}


def palabras_a_numeros(texto):

    palabras = texto.split()
    resultado = []

    for palabra in palabras:
        try:
            num = w2n.word_to_num(palabra)
            resultado.append(str(num))
        except:
            resultado.append(palabra)
    
    texto = " ".join(resultado)
    
    meses = {
        "enero":"01","febrero":"02","marzo":"03","abril":"04",
        "mayo":"05","junio":"06","julio":"07","agosto":"08",
        "septiembre":"09","octubre":"10",
        "noviembre":"11","diciembre":"12"
    }

    for mes,num in meses.items():
        texto = texto.replace(mes,num)

    return texto


def corregir_voz(texto):

    texto = texto.lower()

    correcciones = {

        # meses mal escritos
        "enero": "enero", "hen ero":"enero",
        "febrero":"febrero","fevrero":"febrero","febreo":"febrero",
        "marso":"marzo","marzo":"marzo",
        "abril":"abril","abrill":"abril",
        "mallo":"mayo","mayo":"mayo",
        "junio":"junio","hunio":"junio",
        "julio":"julio",
        "agosto":"agosto","agosto":"agosto",
        "septiembre":"septiembre","setiembre":"septiembre",
        "octubre":"octubre","oktubre":"octubre","octube":"octubre",
        "noviembre":"noviembre","nobiembre":"noviembre",
        "diciembre":"diciembre","disiembre":"diciembre",

        # numeros mal escritos
        "sero":"cero","uno":"uno","una":"uno",
        "dos":"dos","tres":"tres",
        "cuatro":"cuatro","kcuatro":"cuatro",
        "cinco":"cinco","sinco":"cinco","sinko":"cinco",
        "seis":"seis",
        "siete":"siete",
        "ocho":"ocho",
        "nueve":"nueve",
        "dies":"diez","diez":"diez",
        "onse":"once","once":"once",
        "dose":"doce","doce":"doce",
        "trese":"trece","trece":"trece",
        "catorse":"catorce","catorce":"catorce",
        "quinse":"quince","quince":"quince",
        "veinte":"veinte","biente":"veinte",
        "treinta":"treinta",
        "cuarenta":"cuarenta",
        "cincuenta":"cincuenta",

        # años
        "dosmil":"dos mil",
        "veintiseis":"veintiseis",
        "ventiseis":"veintiseis",
        "veinti seis":"veintiseis",
    }

    for mal,bien in correcciones.items():
        texto = texto.replace(mal,bien)

    return texto


def quitar_tildes(texto):
    return ''.join(
c for c in unicodedata.normalize('NFD', texto)
if unicodedata.category(c) != 'Mn'
)

def normalizar_comando(texto):
    texto = quitar_tildes(texto.lower())
    texto = quitar_tildes(texto.lower())
    texto = corregir_voz(texto)
    texto = palabras_a_numeros(texto)
    texto = convertir_fecha_natural(texto)
    texto = convertir_hora_natural(texto)

    reemplazos = {

        # ================= SALIDA =================
        "adio": "adios", "adiós": "adios", "adioz": "adios", "adioss": "adios",
        "a dios": "adios", "hasta luego": "adios", "nos vemos": "adios",

        "sal": "salir", "sali": "salir", "salí": "salir", "salte": "salir",
        "salga": "salir", "salir de aqui": "salir",

        "gracias": "gracias por todo", "grasias": "gracias por todo",
        "gracis": "gracias por todo", "grax": "gracias por todo",
        "grasia": "gracias por todo",

        # ================= BUSCAR =================
        "busk": "busca", "buska": "busca", "vusca": "busca",
        "buscame": "busca", "búscame": "busca", "buskar": "busca",
        "busque": "busca", "buscar": "busca",

        # ================= YOUTUBE =================
        "youtube": "en youtube", "yutub": "en youtube", "yutube": "en youtube",
        "yutu": "en youtube", "you tuve": "en youtube", "llutu": "en youtube",

        # ================= GOOGLE =================
        "google": "en google", "gogle": "en google", "gugle": "en google",
        "googel": "en google", "googl": "en google",

        # ================= ARCHIVOS =================
        "archivo": "en archivos", "archibos": "en archivos",
        "mis archivos": "en archivos", "carpetas": "en archivos",
        "documentos": "en archivos",

        # ================= MODO =================
        "modos": "modo", "moodo": "modo", "mod": "modo",
        "modo jarvis": "modo",

        # ================= INICIA =================
        "iniciar": "inicia", "inicie": "inicia", "inisa": "inicia",
        "arranca": "inicia", "arrancar": "inicia", "prende": "inicia",

        # ================= EJECUTA =================
        "ejecutar": "ejecuta", "ejecute": "ejecuta",
        "ejcuta": "ejecuta", "ejekuta": "ejecuta",

        # ================= ABRE =================
        "abrir": "abre", "abri": "abre", "abreme": "abre",
        "abrelo": "abre", "habre": "abre", "aber": "abre",
        "aver": "abre", "abres": "abre",

        # ================= CIERRA =================
        "cerrar": "cierra", "cierre": "cierra", "sierra": "cierra",
        "cierralo": "cierra", "serrar": "cierra",

        # ================= MIN/MAX =================
        "minimizar": "minimiza", "minimise": "minimiza", "minimice": "minimiza",
        "maximizar": "maximiza", "maximise": "maximiza", "maximice": "maximiza",

        # ================= VOLUMEN =================
        "subir volumen": "sube el volumen", "sube volumen": "sube el volumen",
        "subele volumen": "sube el volumen", "subir el volumen": "sube el volumen",

        "bajar volumen": "baja el volumen", "baja volumen": "baja el volumen",
        "bajale volumen": "baja el volumen", "bajar el volumen": "baja el volumen",

        # ================= BRILLO =================
        "subir brillo": "sube el brillo", "sube brillo": "sube el brillo",
        "subir el brillo": "sube el brillo",

        "bajar brillo": "baja el brillo", "baja brillo": "baja el brillo",
        "bajar el brillo": "baja el brillo",

        # ================= CLIMA =================
        "el clima": "clima", "clim": "clima", "climma": "clima",

        # ================= TEMPERATURA =================
        "temperatura actual": "temperatura", "temp": "temperatura",

        # ================= MAPAS =================
        "mapas": "maps",
        "map": "mapa",
        "mapita": "mapa",
        "googlemap": "google maps",
        "gogle maps": "google maps",

        # ================= EVENTOS =================
        "ver eventos": "mostrar eventos",
        "mis eventos": "mostrar eventos",

        "ver agenda": "mostrar agenda",
        "mi agenda": "mostrar agenda",

        "ver reuniones": "mostrar reuniones",
        "mis reuniones": "mostrar reuniones",

        "nueva reunion": "crear reunion",
        "agendar reunion": "crear reunion",

        # ================= TAREAS =================
        "nueva tarea": "crear tarea",
        "añade tarea": "añadir tarea",

        "mis tareas": "mostrar tareas",
        "ver tareas": "mostrar tareas",

        "lista tareas": "lista de tareas",

        "borrar tarea": "eliminar tarea",
    }

    for mal, bien in reemplazos.items():
        if mal in texto:
            texto = texto.replace(mal, bien)

    return texto


def formatear_fecha(fecha_iso):
    dt = datetime.fromisoformat(fecha_iso)
    return dt.strftime("%d/%m/%Y a las %I:%M %p")

SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks"
]



def obtener_servicio():
    creds = None

    ruta_token = os.path.join(BASE_DIR, "token.json")
    ruta_credentials = os.path.join(BASE_DIR, "credentials.json")

    if os.path.exists(ruta_token):
        creds = Credentials.from_authorized_user_file(ruta_token, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(ruta_credentials):
                raise FileNotFoundError("No se encontró credentials.json en la carpeta de Jarvis")
            flow = InstalledAppFlow.from_client_secrets_file(ruta_credentials, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(ruta_token, "w") as token:
            token.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)


def obtener_servicio_tasks():
    creds = None

    ruta_token = os.path.join(BASE_DIR, "token.json")
    ruta_credentials = os.path.join(BASE_DIR, "credentials.json")

    if os.path.exists(ruta_token):
        creds = Credentials.from_authorized_user_file(ruta_token, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(ruta_credentials):
                raise FileNotFoundError("No se encontró credentials.json en la carpeta de Jarvis")
            flow = InstalledAppFlow.from_client_secrets_file(ruta_credentials, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(ruta_token, "w") as token:
            token.write(creds.to_json())

    return build("tasks", "v1", credentials=creds)


def delete_event(service,    event_id):
    service.events().delete(calendarId="primary", eventId=event_id).execute()
    return True



def get_events(service):
    events_result = service.events().list(calendarId="primary", maxResults=10,
    singleEvents=True, orderBy="startTime").execute()
    return events_result.get("items", [])

def create_event(service, subject, start, end, timezone="America/Bogota"):
    event = {
        "summary": subject,
        "start": {"dateTime": start, "timeZone": timezone},
        "end": {"dateTime": end, "timeZone": timezone},
    }
    event = service.events().insert(calendarId="primary", body=event).execute()
    return event

def update_event(service, event_id, nombre=None, start=None, end=None):
    # Recuperar el evento actual
    evento_actual = service.events().get(calendarId="primary", eventId=event_id).execute()

    # Construir el cuerpo con los datos existentes
    evento = {"summary": evento_actual["summary"]}

    if nombre:   # solo cambia si se da un nombre nuevo
        evento["summary"] = nombre
    
    if start and end:
        evento["start"] = {"dateTime": start, "timeZone": "America/Bogota"}
        evento["end"] = {"dateTime": end, "timeZone": "America/Bogota"}
    return service.events().update(calendarId="primary", eventId=event_id, body=evento).execute()


tareas = []  # lista global para almacenar tareas

# Crear tarea
def create_task(service_tasks, titulo, fecha=None):
    tarea = {"title": titulo}
    
    if fecha:
        # Google Tasks solo acepta fecha límite en formato YYYY-MM-DDT00:00:00Z
        tarea["due"] = fecha.replace(
            hour=0, minute=0, second=0, microsecond=0
        ).isoformat() + "Z"
    
    return service_tasks.tasks().insert(
        tasklist="@default", 
        body=tarea
    ).execute()

    # Mostrar tareas
def get_tasks(service_tasks, max_results=10):
    tasks_result = service_tasks.tasks().list(tasklist="@default", maxResults=max_results).execute()
    return tasks_result.get("items", [])


# Modificar tarea
def update_task(service_tasks, task_id, titulo=None, fecha=None):
    tarea = {}
    if titulo:
        tarea["title"] = titulo
    if fecha:
        tarea["due"] = fecha.date().isoformat() + "T00:00:00Z"

    return service_tasks.tasks().update(
        tasklist="@default",
        task=task_id,   # 👈 correcto
        
        body=tarea
    ).execute()

        # Eliminar tarea
def delete_task(service_tasks, task_id):
    service_tasks.tasks().delete(tasklist="@default", task=task_id).execute()
    return True


try:
    pygame.mixer.init()
except:
    print("Audio no disponible")



def resource_path(relative_path):
    """ Devuelve la ruta correcta de un recurso, incluso dentro del exe """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

json_path = resource_path("jarvis.json")
apps_path = resource_path("memoria_apps.txt")
micro_path = resource_path("memoria_microfono.txt")

if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)  # carpeta del .exe
else:
    base_path = os.path.abspath(".")  # si es .py

    apps_file = os.path.join(base_path, "memoria_apps.txt")
    micro_file = os.path.join(base_path, "memoria_microfono.txt")


def buscar_archivos(termino, carpeta=None):
    if carpeta is None:
        carpeta = os.path.expanduser("~")  # Carpeta Home real
    resultados = []
    for root, dirs, files in os.walk(carpeta):
        for file in files:
            if termino.lower() in file.lower():
                resultados.append(os.path.join(root, file))
    return resultados


def abrir_maps(lugar, chat_widget=None):
    url = f"https://www.google.com/maps/search/{lugar}"
    webbrowser.open(url)
    hablar(f"Mostrando {lugar} en Google Maps", chat_widget)


def obtener_clima(ciudad, chat_widget=None):
    api_key = "d26db5c0be2d034422f043e6711865b1"  # tu clave de OpenWeatherMap
    url = f"http://api.openweathermap.org/data/2.5/weather?q={ciudad}&appid=d26db5c0be2d034422f043e6711865b1&units=metric&lang=es"
    resp = requests.get(url)

    print("DEBUG:", resp.status_code, resp.text)  # <-- depuración

    if resp.status_code == 200:
        data = resp.json()
        temp = data["main"]["temp"]
        desc = data["weather"][0]["description"]
        hablar(f"En {ciudad} hay {desc} con {temp}°C", chat_widget)
    else:
        hablar("No pude obtener el clima", chat_widget)


service_calendar = obtener_servicio()
service_tasks = obtener_servicio_tasks()


def grabar_audio(nombre_archivo="input.wav", duracion=5, fs=16000):
    print("Grabando...")
    audio = sd.rec(int(duracion * fs), samplerate=fs, channels=1, dtype='int16')
    sd.wait()
    print("Grabación terminada.")

    # Guardar como WAV
    with wave.open(nombre_archivo, 'wb') as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(fs)
        wf.writeframes(audio.tobytes())

        return nombre_archivo

def reconocer_comando():
    archivo = grabar_audio()
    result = model.transcribe(archivo, language="es")
    texto = result["text"].strip()
    print("Transcripción:", texto)
    return texto


        
def convertir_numeros(texto):
    palabras = texto.lower().split()
    resultado = []
    for p in palabras:
        if p in NUMEROS:
            resultado.append(str(NUMEROS[p]))
        else:
            resultado.append(p)
    return " ".join(resultado)

# =============================
# MOTOR DE VOZ
# =============================
engine = pyttsx3.init()
engine.setProperty('rate', 170)
engine.setProperty('volume', 1.0)



def hablar(texto, chat_widget=None):
    if chat_widget:
        try:
            chat_widget.parent().agregar_jarvis(texto)
        except:
            chat_widget.append("🤖 " + texto)

    engine.say(texto)
    engine.runAndWait()


# =============================
# MEMORIA DE APPS
# =============================
apps = {}


# ====================================
# Modos de trabajo, gaming o estudio
# ====================================
modos = {
    "estudio": ["teams", "chrome", "colegio", "office"],
    "gaming 1": ["steam", "msi afterburner"],
    "gaming 2": ["epic", "msi afterburner", "bakkesmod"],
    "programacion": ["vscode", "copilot", "chatgpt"],
}
if os.path.exists(apps_path):
    with open(apps_path, "r", encoding="utf-8") as f:
        for linea in f:
            linea = linea.strip()
            if linea == "" or "|" not in linea:
                continue
            nombre, ruta = linea.split("|", 1)
            apps[nombre.lower()] = ruta




def guardar_memoria(nombre, ruta, chat_widget=None):
    nombre = nombre.lower()
    apps[nombre] = ruta
    with open(apps_file, "a", encoding="utf-8") as f:
        f.write(f"{nombre}|{ruta}\n")
    
    hablar(f"He aprendido la aplicación {nombre}.", chat_widget)



# =============================
# ESCUCHAR VOZ
# =============================

def escuchar(chat_widget=None):
    if chat_widget:
        # Mostrar mensaje dentro del div de chat
        chat_widget.insertHtml(f"""
    <table width="100%" style="margin-top:10px;">
    <tr>
    <td width="25%"></td>
    <td width="50%" align="left">
    <table cellspacing="0" cellpadding="0">
    <tr>
    <td style="
    background:#0f1928;
        border-radius:18px;
        padding:10px 16px;
        color:white;
        font-size:15px;
        max-width:350px;">
        🎤 Escuchando... (presiona ENTER para detener)
        </td>
        </tr>
        </table>
        </td>
        <td width="25%"></td>
        </tr>
        </table>
        """)
        chat_widget.verticalScrollBar().setValue(chat_widget.verticalScrollBar().maximum())
        chat_widget.append("🎤 Escuchando... (presiona ENTER para detener)")

    fs = 16000
    audio_total = []
    
    def callback(indata, frames, time, status):
        audio_total.append(indata.copy())

    # stream continuo
    with sd.InputStream(samplerate=fs, channels=1, dtype='int16', callback=callback):
        while True:
            if keyboard.is_pressed("enter"):
                break

    # unir audio
    audio_np = np.concatenate(audio_total, axis=0)

    archivo = "input.wav"
    with wave.open(archivo, 'wb') as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(fs)
        wf.writeframes(audio_np.tobytes())

    if chat_widget:
        chat_widget.append("🧠 Procesando voz...")

    result = model.transcribe(archivo, language="es")
    comando = result["text"].strip().lower()

    print("Transcripción:", comando)

    if comando and chat_widget:
        chat_widget.append(f"Tú (voz): {comando}")

    return comando

    


def interpretar_comando(comando):
    comando = comando.lower()
    if "alarma" in comando:
        return "alarma"
    elif "cronómetro" in comando or "timer" in comando:
        return "cronometro"
    else:
        return None

def extraer_datos(comando, tipo):
    if tipo == "alarma":
        match = re.search(r'(\d{1,2})(?::(\d{2}))?', comando)
        if match:
            hora = int(match.group(1))
            minuto = int(match.group(2)) if match.group(2) else 0
            return datetime.time(hora, minuto)
        elif tipo == "cronometro":
            match = re.search(r'(\d+)\s*(minutos|min|segundos|seg)', comando)
            if match:
                cantidad = int(match.group(1))
                unidad = match.group(2)
                if "min" in unidad:
                    return cantidad * 60
                else:
                    return cantidad
        return None


# =============================
# CONTROL DE VENTANAS
# =============================
def encontrar_ventana(nombre):
    hwnd = win32gui.FindWindow(None, nombre)
    if hwnd == 0:
        def enumHandler(hwnd, lParam):
            if nombre.lower() in win32gui.GetWindowText(hwnd).lower():
                lParam.append(hwnd)
        hwnds = []
        win32gui.EnumWindows(enumHandler, hwnds)
        if hwnds:
            return hwnds[0]
        return None
    return hwnd

def minimizar_app(nombre, chat_widget=None):
    hwnd = encontrar_ventana(nombre)
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        hablar(f"{nombre} minimizada", chat_widget)
    else:
        hablar(f"No encontré la ventana de {nombre}", chat_widget)

def maximizar_app(nombre, chat_widget=None):
    hwnd = encontrar_ventana(nombre)
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        hablar(f"{nombre} maximizada", chat_widget)
    else:
        hablar(f"No encontré la ventana de {nombre}", chat_widget)

def cerrar_app(nombre, chat_widget=None):
    hwnd = encontrar_ventana(nombre)
    if hwnd:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        hablar(f"{nombre} cerrada", chat_widget)
    else:
        hablar(f"No encontré la ventana de {nombre}", chat_widget)



# =============================
# CONTROL DE VOLUMEN
# =============================
VK_VOLUME_MUTE = 0xAD
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP = 0xAF

def ajustar_volumen_relativo(porcentaje, chat_widget=None):
    pasos = int(abs(porcentaje) // 2)
    tecla = VK_VOLUME_UP if porcentaje > 0 else VK_VOLUME_DOWN
    for _ in range(pasos):
        ctypes.windll.user32.keybd_event(tecla, 0, 0, 0)
    hablar(f"Volumen ajustado en {porcentaje}%", chat_widget)

def subir_volumen(porcentaje=10, chat_widget=None):
    ajustar_volumen_relativo(porcentaje, chat_widget)

def bajar_volumen(porcentaje=10, chat_widget=None):
    ajustar_volumen_relativo(-porcentaje, chat_widget)

def mute_volumen(chat_widget=None):
    ctypes.windll.user32.keybd_event(VK_VOLUME_MUTE, 0, 0, 0)
    hablar("Volumen silenciado o activado", chat_widget)

# =============================
# CONTROL DE BRILLO
# =============================
w = wmi.WMI(namespace='wmi')

def get_brightness():
    for b in w.WmiMonitorBrightness():
        return b.CurrentBrightness
    return 50

def set_brightness(valor, chat_widget=None):
    valor = max(0, min(100, valor))
    for b in w.WmiMonitorBrightnessMethods():
        b.WmiSetBrightness(valor, 0)
    hablar(f"Brillo ajustado a {valor}%", chat_widget)

def ajustar_brillo_relativo(porcentaje, chat_widget=None):
    current = get_brightness()
    set_brightness(current + porcentaje, chat_widget)

def subir_brillo(porcentaje=10, chat_widget=None):
    ajustar_brillo_relativo(porcentaje, chat_widget)

def bajar_brillo(porcentaje=10, chat_widget=None):
    ajustar_brillo_relativo(-porcentaje, chat_widget)
    
# =============================
# ABRIR APP
# =============================
def abrir_app(nombre, chat_widget=None):
    nombre = nombre.lower()
    if nombre not in apps:
        hablar(f"No conozco la app {nombre}", chat_widget)
        return
    ruta = apps[nombre]
    try:
        if "!" in ruta:
            comando = f'explorer shell:appsFolder\\{ruta}'
            subprocess.Popen(comando)
        elif ruta.lower().startswith("http"):
            webbrowser.open(ruta)
        else:
            os.startfile(ruta)
        hablar(f"Abriendo {nombre}", chat_widget)
    except Exception as e:
        hablar(f"Error abriendo app: {e}", chat_widget)



def activar_modo(nombre_modo, chat_widget=None):
    nombre_modo = nombre_modo.lower().strip()

    if nombre_modo not in modos:
        hablar(f"No existe el modo {nombre_modo}", chat_widget)
        return

    hablar(f"Activando modo {nombre_modo}", chat_widget)

    for app in modos[nombre_modo]:
        abrir_app(app, chat_widget)
        time.sleep(1)  # evita que se abra todo al mismo tiempo


def texto_a_numero(texto):
    texto = texto.lower()

    unidades = {
        "cero":0,"uno":1,"una":1,"un":1,"dos":2,"tres":3,"cuatro":4,"cinco":5,
        "seis":6,"siete":7,"ocho":8,"nueve":9,"diez":10,"once":11,"doce":12,
        "trece":13,"catorce":14,"quince":15,"dieciseis":16,"dieciséis":16,
        "diecisiete":17,"dieciocho":18,"diecinueve":19,"veinte":20,
        "veintiuno":21,"veintidos":22,"veintidós":22,"veintitres":23,"veintitrés":23,
        "veinticuatro":24,"veinticinco":25,"veintiseis":26,"veintiséis":26,
        "veintisiete":27,"veintiocho":28,"veintinueve":29
    }

    decenas = {
        "treinta":30,"cuarenta":40,"cincuenta":50,"sesenta":60,
        "setenta":70,"ochenta":80,"noventa":90
    }

    centenas = {
        "cien":100,"ciento":100,"doscientos":200,"trescientos":300,
        "cuatrocientos":400,"quinientos":500,"seiscientos":600,
        "setecientos":700,"ochocientos":800,"novecientos":900
    }

    texto = texto.replace(" y ", " ")
    palabras = re.findall(r'\w+', texto)

    total = 0
    actual = 0

    for palabra in palabras:

        if palabra in unidades:
            actual += unidades[palabra]

        elif palabra in decenas:
            actual += decenas[palabra]

        elif palabra in centenas:
            actual += centenas[palabra]

        elif palabra == "mil":
            if actual == 0:
                actual = 1
            total += actual * 1000
            actual = 0

        elif palabra == "millon" or palabra == "millón":
            if actual == 0:
                actual = 1
            total += actual * 1000000
            actual = 0

        elif palabra == "millones":
            if actual == 0:
                actual = 1
            total += actual * 1000000
            actual = 0

    total += actual
    return total if total > 0 else None
                    

# =============================
# EXTRAER NUMERO DEL TEXTO
# =============================
def extraer_numero(texto, defecto=10):
    # primero números normales
    numeros = re.findall(r'\d+', texto)
    if numeros:
        return int(numeros[0])

    # luego números por voz
    n = texto_a_numero(texto)
    if n:
        return n

    return defecto



def extraer_nombre_fecha_hora(texto, tipo="evento"):
    """
    Extrae nombre, fecha y hora de un comando tipo 'crear tarea' o 'crear evento'.
    Retorna: nombre (str), fecha (datetime), hora (int)
    """
    texto = texto.lower()

    # 1️⃣ Detectar fecha (15/02/2026, 15-02-2026, 15 02 2026)
    match_fecha = re.search(r"(\d{1,2})[\/\-\s](\d{1,2})[\/\-\s](\d{4})", texto)
    if match_fecha:
        dia, mes, anio = map(int, match_fecha.groups())
        fecha = datetime(anio, mes, dia)
    elif "mañana" in texto:
        fecha = datetime.now() + timedelta(days=1)
    elif "hoy" in texto:
        fecha = datetime.now()
    else:
        fecha = datetime.now()  # por defecto

    # 2️⃣ Detectar hora (a las 17)
    match_hora = re.search(r"a las (\d{1,2})", texto)
    hora = int(match_hora.group(1)) if match_hora else 8

    # Ajustar hora en la fecha
    fecha = fecha.replace(hour=hora, minute=0)

    # 3️⃣ Extraer nombre → eliminar palabras clave y la fecha/hora
    nombre = texto
    if tipo == "evento":
        nombre = re.sub(r"crear (reunión|reunion|evento)", "", nombre)
    else:  # tarea
        nombre = re.sub(r"crear (tarea|agregar tarea|añadir tarea)", "", nombre)

    # quitar fecha y hora
    nombre = re.sub(r"(\d{1,2}[\/\-\s]\d{1,2}[\/\-\s]\d{4})", "", nombre)
    nombre = re.sub(r"a las \d{1,2}", "", nombre)

    # quitar artículos innecesarios como "el", "la", "los", "las"
    nombre = re.sub(r"\b(el|la|los|las|un|una|unos|unas)\b", "", nombre)

    # limpiar espacios extra
    nombre = nombre.strip()
    if not nombre:
        nombre = "Reunión Jarvis" if tipo=="evento" else "Tarea sin nombre"

    return nombre, fecha, hora


estado_confirmacion = None  # inicialización global

def apagar_pc():
    os.system("shutdown /s /t 1")

def reiniciar_pc():
    os.system("shutdown /r /t 1")

def bloquear_pc():
    os.system("rundll32.exe user32.dll,LockWorkStation")

# =============================
# INTERPRETAR COMANDOS
# =============================
def interpretar(comando, chat_widget=None):
    global estado_confirmacion

    if not comando:  # maneja None o ""
        if chat_widget:
            chat_widget.append("⚠️ No se recibió texto para interpretar.")
        return
    
    comando = comando.strip().lower()
    
    if estado_confirmacion:
        if comando in ["sí", "si", "claro", "vale", "ok"]:
            if estado_confirmacion == "apagar":
                hablar("Apagando el sistema...", chat_widget)
                apagar_pc()
            elif estado_confirmacion == "reiniciar":
                hablar("Reiniciando el sistema...", chat_widget)
                reiniciar_pc()
            estado_confirmacion = None
            return
        elif comando in ["no", "cancelar", "mejor no"]:
            hablar("Acción cancelada.", chat_widget)
            estado_confirmacion = None
            return
        else:
            hablar("Responde con 'sí' o 'no'.", chat_widget)
            return

    # Comandos de poder
    if comando in ["apagar", "apaga el pc", "shutdown", "apaga el computador"]:
        hablar("¿Seguro que quieres apagar el sistema? (sí/no)", chat_widget)
        estado_confirmacion = "apagar"
        return

    if comando in ["reiniciar", "restart", "reinicia el sistema"]:
        hablar("¿Seguro que quieres reiniciar el sistema? (sí/no)", chat_widget)
        estado_confirmacion = "reiniciar"
        return

    if comando in ["bloquear", "lock", "bloquear pantalla"]:
        hablar("Bloqueando la pantalla...", chat_widget)
        bloquear_pc()
        return



    if comando.strip() in ["adios", "salir", "gracias por todo"]:
        hablar("Jarvis apagándose...", chat_widget)
        QApplication.quit()  # cierra la app de forma controlada
        return
    
    if "busca" in comando and "en youtube" in comando:
        termino = comando.replace("busca","").replace("en youtube","").strip()
        url = f"https://www.youtube.com/results?search_query={termino}"
        webbrowser.open(url)
        hablar(f"Buscando {termino} en YouTube", chat_widget)
        return

    if "busca" in comando and "en google" in comando:
        termino = comando.replace("busca","").replace("en google","").strip()
        url = f"https://www.google.com/search?q={termino}"
        webbrowser.open(url)
        hablar(f"Buscando {termino} en Google", chat_widget)
        return
    
    if "busca" in comando and "en archivos" in comando:
        termino = comando.replace("busca","").replace("en archivos","").strip()
        resultados = buscar_archivos(termino)
        if resultados:
            hablar(f"Encontré {len(resultados)} archivos con '{termino}' en Home:", chat_widget)
            for r in resultados:
                if chat_widget:
                    chat_widget.insertHtml(f"""
            <table width="100%" style="margin-top:8px;">
            <tr>
            <td width="25%"></td>
            <td width="50%" align="left">
            <table cellspacing="0" cellpadding="0">
            <tr>
            <td style="
            background:#0f1928;
                border-radius:18px;
                padding:10px 16px;
                color:white;
                font-size:15px;
                max-width:350px;
                word-wrap:break-word;">
                📂 <a href='file:///{r}' style='color:#00BFFF; text-decoration:none;'>
                    {os.path.basename(r)}
                    </a>
                    </td>
                    </tr>
                    </table>
                    </td>
                    <td width="25%"></td>
                    </tr>
                    </table>
                    """)  
        else:
            hablar(f"No encontré archivos con '{termino}' en Home.", chat_widget)
        return
    
    
    if "modo" in comando:
        nombre = comando.replace("modo","").strip()
        activar_modo(nombre, chat_widget)
        return

    if "abre" in comando or "inicia" in comando or "ejecuta" in comando:
        nombre = comando.replace("abre","").replace("inicia","").replace("ejecuta","").strip()
        abrir_app(nombre, chat_widget)
        return

    if "cierra" in comando:
        nombre = comando.replace("cierra","").strip()
        cerrar_app(nombre, chat_widget)
        return
    if "minimiza" in comando:
        nombre = comando.replace("minimiza","").strip()
        minimizar_app(nombre, chat_widget)
        return
    if "maximiza" in comando:
        nombre = comando.replace("maximiza","").strip()
        maximizar_app(nombre, chat_widget)
        return

    if "sube el volumen" in comando:
        porcentaje = extraer_numero(comando, 10)
        subir_volumen(porcentaje, chat_widget)
        return
    
    if "baja el volumen" in comando:
        porcentaje = extraer_numero(comando, 10)
        bajar_volumen(porcentaje, chat_widget)
        return
    
    if "sube el brillo" in comando:
        porcentaje = extraer_numero(comando, 10)
        subir_brillo(porcentaje, chat_widget)
        return
    
    if "baja el brillo" in comando:
        porcentaje = extraer_numero(comando, 10)
        bajar_brillo(porcentaje, chat_widget)
        return
        
    if "clima" in comando or "temperatura" in comando:
        ciudad = comando.replace("clima","").replace("temperatura","").replace("en","").strip()
        if not ciudad:
            ciudad = "Bogotá"  # valor por defecto
        
        obtener_clima(ciudad, chat_widget)
        return
    

    if "mapa" in comando or "maps" in comando or "google maps" in comando:
        lugar = comando.replace("mapa","").replace("maps","").replace("google maps","").strip()
        if not lugar:
            lugar = "Cota, Cundinamarca"  # valor por defecto
        abrir_maps(lugar, chat_widget)
        return
    
    service = obtener_servicio()
    

    # Mostrar eventos
    if "mostrar eventos" in comando or "mostrar agenda" in comando or "mostrar reuniones" in comando:
        eventos = get_events(service)
        if eventos:
            for e in eventos:
                inicio = formatear_fecha(e["start"]["dateTime"])
                hablar(f"📅 {e['summary']} → {inicio}", chat_widget)
        else:
            hablar("No encontré eventos próximos.", chat_widget)
        return


    # Crear reunión / evento
    elif "crear reunión" in comando or "crear reunion" in comando or "crear evento" in comando:
        nombre, fecha, hora = extraer_nombre_fecha_hora(comando, tipo="evento")
        start = fecha.isoformat()
        end = (fecha + timedelta(hours=1)).isoformat()
        nuevo = create_event(service, nombre, start, end)
        hablar(f"✅ He creado tu reunión: {nuevo['summary']} el {formatear_fecha(nuevo['start']['dateTime'])}", chat_widget)
        return

    
    # Eliminar reunión
    elif "eliminar reunión" in comando or "borrar reunión" in comando or "cancelar reunión" in comando or "eliminar reunion" in comando or "borrar reunion" in comando or "cancelar reunion" in comando or "emilinar evento" in comando or "borrar evento" in comando or "cancelar evento" in comando:
        eventos = get_events(service)
        if eventos:
            nombre_buscar = comando.replace("eliminar reunión", "").replace("borrar reunión", "").replace("cancelar reunión", "").replace("eliminar reunion", "").replace("borrar reunion", "").replace("cancelar reunion", "").replace("emilinar evento", "").replace("borrar evento", "").replace("cancelar evento", "").strip()
            nombre_buscar = re.sub(r"(mañana|hoy|a las \d{1,2})", "", nombre_buscar).strip()

            evento_id = None
            for e in eventos:
                if nombre_buscar and nombre_buscar.lower() in e["summary"].lower():
                    evento_id = e["id"]
                    break
            if not evento_id:
                evento_id = eventos[0]["id"]

            delete_event(service, evento_id)
            hablar(f"🗑️ He eliminado la reunión {nombre_buscar if nombre_buscar else 'seleccionada'}.", chat_widget)
        else:
            hablar("No encontré reuniones para eliminar.", chat_widget)
        return
    
    # Modificar reunión
    elif "modificar reunión" in comando or "editar reunión" in comando or "modificar evento" in comando:
        eventos = get_events(service)
        if eventos:
            nombre_buscar = comando.replace("modificar reunión", "").replace("editar reunión", "").replace("modificar evento", "").strip()
            nombre_buscar = re.sub(r"(mañana|hoy|a las \d{1,2})", "", nombre_buscar).strip()
            
            event_id = None
            nombre_original = None
            for e in eventos:
                if nombre_buscar and nombre_buscar.lower() in e["summary"].lower():
                    event_id = e["id"]
                    nombre_original = e["summary"]
                    break

            if not event_id:
                event_id = eventos[0]["id"]
                nombre_original = eventos[0]["summary"]

            # Detectar nueva fecha/hora
            fecha = datetime.now()
            match_fecha = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", comando)
            if match_fecha:
                fecha = datetime.strptime(match_fecha.group(1), "%d/%m/%Y")

            match_hora = re.search(r"a las (\d{1,2})", comando)
            hora = int(match_hora.group(1)) if match_hora else 8

            start = fecha.replace(hour=hora, minute=0).isoformat()
            end = fecha.replace(hour=hora+1, minute=0).isoformat()

            # 👇 aquí pasamos nombre=None para conservar el original
            actualizada = update_event(service, event_id, nombre=None, start=start, end=end)
            hablar(f"✏️ He modificado la reunión: {actualizada['summary']} el {formatear_fecha(actualizada['start']['dateTime'])}", chat_widget)
            return
        else:
            hablar("No encontré reuniones para modificar.", chat_widget)
            return
        
    # Crear tarea    
    elif "crear tarea" in comando or "agregar tarea" in comando or "añadir tarea" in comando:
        # Extraer nombre, fecha y hora usando la función que ya parsea
        nombre, fecha, hora = extraer_nombre_fecha_hora(comando, tipo="tarea")

        # Si se indicó hora, aplicarla
        if hora is not None:
            fecha = fecha.replace(hour=hora, minute=0)
        else:
            fecha = fecha.replace(hour=8, minute=0)  # hora por defecto

        # Si no quedó nombre, usar uno por defecto
        if not nombre:
            nombre = "Tarea sin nombre"

        # Crear la tarea
        nueva_tarea = create_task(service_tasks, nombre, fecha)
        if nueva_tarea:
            hablar(f"📝 He agregado la tarea: {nueva_tarea['title']} para {fecha.strftime('%d/%m/%Y')} a las {fecha.hour}", chat_widget)
        else:
            hablar("⚠️ No pude crear la tarea.", chat_widget)
        return



    # Mostrar tareas
    elif "mostrar tareas" in comando or "lista de tareas" in comando:
        tareas = get_tasks(service_tasks)
        if tareas:
            for t in tareas:
                fecha = t.get("due", None)
                if fecha:
                    hablar(f"📝 {t['title']} → vence {fecha}", chat_widget)
                else:
                    hablar(f"📝 {t['title']} → sin fecha límite", chat_widget)
        else:
            hablar("No encontré tareas pendientes.", chat_widget)
        return



        # Modificar tarea
    elif "modificar tarea" in comando or "editar tarea" in comando:
        tareas = get_tasks(service_tasks)
        if tareas:
            nombre_buscar = comando.replace("modificar tarea", "").replace("editar tarea", "").strip()
            nombre_buscar = re.sub(r"(mañana|hoy|a las \d{1,2})", "", nombre_buscar).strip()
            
            tarea_id = None
            titulo_original = None
            for t in tareas:
                if nombre_buscar and nombre_buscar.lower() in t["title"].lower():
                    tarea_id = t["id"]
                    titulo_original = t["title"]
                    break

            # fallback: si no encontró coincidencia, usa la primera tarea
            if not tarea_id:
                tarea_id = tareas[0]["id"]
                titulo_original = tareas[0]["title"]

            # Detectar nueva fecha
            fecha = None
            match_fecha = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", comando)
            if match_fecha:
                fecha = datetime.strptime(match_fecha.group(1), "%d/%m/%Y")

                actualizada = update_task(service_tasks, tarea_id, titulo=titulo_original, fecha=fecha)
                hablar(f"✏️ He modificado la tarea: {actualizada['title']} {f'para {fecha.date()}' if fecha else ''}", chat_widget)
                return
            else:
                hablar("No encontré tareas para modificar.", chat_widget)
                return


    # Eliminar tarea
    elif "eliminar tarea" in comando or "borrar tarea" in comando or "cancelar tarea" in comando:
        tareas = get_tasks(service_tasks)
        if tareas:
            nombre_buscar = comando.replace("eliminar tarea", "").replace("borrar tarea", "").replace("cancelar tarea", "").strip()
            nombre_buscar = re.sub(r"(mañana|hoy|a las \d{1,2})", "", nombre_buscar).strip()

            tarea_id = None
            for t in tareas:
                if nombre_buscar and nombre_buscar.lower() in t["title"].lower():
                    tarea_id = t["id"]
                    break
            if not tarea_id:
                tarea_id = tareas[0]["id"]

            delete_task(service_tasks, tarea_id)
            hablar(f"🗑️ He eliminado la tarea {nombre_buscar if nombre_buscar else 'seleccionada'}.", chat_widget)
        else:
            hablar("No encontré tareas para eliminar.", chat_widget)
            return
    
    # Otros comandos
    else:
        hablar("No entendí el comando, intenta con 'mostrar eventos', 'crear reunión', 'modificar reunión' o 'eliminar reunión'.", chat_widget)



def interpretar_multiple(comando, chat_widget=None):
    if not comando:
        if chat_widget:
            chat_widget.append("⚠️ No se recibió texto.")
        return

    texto = comando.lower().strip()

    # dividir por comas o " y "
    partes = re.split(r',|\s+y\s+', texto)
    partes = [p.strip() for p in partes if p.strip()]

    if not partes:
        return

    # detectar verbo principal del primero
    verbo = None
    if partes[0].startswith(("abre", "inicia", "ejecuta")):
        verbo = "abre"
    elif partes[0].startswith("cierra"):
        verbo = "cierra"
    elif partes[0].startswith("minimiza"):
        verbo = "minimiza"
    elif partes[0].startswith("maximiza"):
        verbo = "maximiza"

    for i, parte in enumerate(partes):
        #si es el primero → normal
        if i == 0:
            interpretar(parte, chat_widget)
        else:
            # si no tiene verbo, hereda
            if verbo and not parte.startswith(("abre","cierra","minimiza","maximiza","sube","baja","busca")):
                parte = verbo + " " + parte

            interpretar(parte, chat_widget)



# =============================
# INTERFAZ GRÁFICA
# =============================

class JarvisUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Jarvis")
        self.resize(900, 600)

        # ===== FONDO =====
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor("#101523"))
        palette.setColor(QPalette.Base, QColor("#101523"))
        palette.setColor(QPalette.Text, QColor("#FFFFFF"))
        self.setPalette(palette)

        # ===== LAYOUT PRINCIPAL =====
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0,0,0,0)
        main_layout.setSpacing(0)

        # ===== CHAT =====
        self.chat = QTextBrowser()
        self.chat.setOpenExternalLinks(False)   # Para manejar clics manualmente
        self.chat.setOpenLinks(False)   # 👈 evita que muestre el contenido del enlace
        self.chat.anchorClicked.connect(self.abrir_archivo)
        self.chat.setFont(QFont("Segoe UI", 14))
        self.chat.setStyleSheet("""
QTextBrowser {
    border: none;
    padding: 10px;
    background-color: #101523;
        color: white;
    }
    """)
        main_layout.addWidget(self.chat, 1)

        # ===== BARRA CENTRADA =====
        contenedor = QHBoxLayout()
        contenedor.addStretch()
        
        barra = QHBoxLayout()


        # INPUT
        self.input = QLineEdit()
        self.input.setPlaceholderText("Escribe tu comando...")
        self.input.setFixedWidth(500)
        self.input.setFixedHeight(50)
        self.input.setStyleSheet("""
        QLineEdit {
            border-radius: 25px;
            padding-left: 15px;
            background-color: #1a2130;
                    
            color: white;
            font-size: 16px;
            border: 2px solid #6b748f;
        }
        """)
        self.input.returnPressed.connect(self.enviar)
        barra.addWidget(self.input)

        # BOTON ENVIAR
        self.btn = QPushButton("➤")
        self.btn.setFixedSize(50, 50)
        self.btn.clicked.connect(self.enviar)
        self.btn.setStyleSheet("""
        QPushButton {
            border-radius: 25px;
            background: #1a2130;
            
            color: white;
            font-size: 18px;
        }
        QPushButton:hover {
            background: #2B4FD6;
        }
        """)
        barra.addWidget(self.btn)

        # BOTON VOZ
        self.btn_voice = QPushButton("🎤")
        self.btn_voice.setFixedSize(50, 50)
        self.btn_voice.clicked.connect(lambda: interpretar_multiple(escuchar(chat_widget=self.chat), self.chat))
        self.btn_voice.setStyleSheet("""
        QPushButton {
            border-radius: 25px;
            background: #1a2130;
                
            color: white;
            font-size: 18px;
        }
        QPushButton:hover {
            background: #0055ff;
        }
        """)
        barra.addWidget(self.btn_voice)
        contenedor.addLayout(barra)
        contenedor.addStretch()
        main_layout.addLayout(contenedor)
        main_layout.addSpacing(50)   # cambia número si quieres

        self.setLayout(main_layout)
        
    def abrir_archivo(self, url):
        ruta = url.toLocalFile() if url.isLocalFile() else url.toString()
        try:
            # Si es un archivo local, abrir con la aplicación predeterminada
            if ruta.lower().endswith((".docx", ".xlsx", ".pptx", ".pdf", ".txt")):
                os.startfile(ruta)  # abre con Word, Excel, PowerPoint, etc.
            else:
                # Si es un enlace web, abrir en el navegador
                webbrowser.open(ruta)
        except Exception as e:
            self.chat.append(f"⚠️ No pude abrir {ruta}: {e}")


    
    def agregar_usuario(self, texto):
        self.chat.insertHtml(f"""
    <table width="100%" style="margin-top:10px;">
    <tr>

    <td width="25%"></td>

    <td width="50%" align="right">

    <table cellspacing="0" cellpadding="0">
    <tr>
    <td style="
    background:#1d2539;
        border-radius:18px;
        padding:10px 16px;
        color:white;
        font-size:15px;
        max-width:350px;">
        {texto}
        </td>
        </tr>
        </table>

        </td>

        <td width="25%"></td>

        </tr>
        </table>
        """)
        self.chat.verticalScrollBar().setValue(self.chat.verticalScrollBar().maximum())
    
    def agregar_jarvis(self, texto):
        self.chat.insertHtml(f"""
    <table width="100%" style="margin-top:10px;">
    <tr>

    <td width="25%"></td>

    <td width="50%" align="left">

    <table cellspacing="0" cellpadding="0">
    <tr>
    <td style="
    background:#0f1928;
        border-radius:18px;
        padding:10px 16px;
        color:white;
        font-size:15px;
        max-width:350px;">
        {texto}
        </td>
        </tr>
        </table>

        </td>

        <td width="25%"></td>

        </tr>
        </table>
        """)
        self.chat.verticalScrollBar().setValue(self.chat.verticalScrollBar().maximum())



    def enviar(self):
        texto = self.input.text().strip()
        if texto:
            self.agregar_usuario(texto)
            interpretar_multiple(texto, self.chat)   # 👈 ahora usa la nueva función
            self.input.clear()


# =============================
# MAIN
# =============================
if __name__ == "__main__":
    cargar_whisper()
    app = QApplication(sys.argv)
    ui = JarvisUI()
    ui.show()
    sys.exit(app.exec())