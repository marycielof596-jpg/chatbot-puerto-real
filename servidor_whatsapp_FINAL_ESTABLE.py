from flask import Flask, request
from openpyxl import load_workbook, Workbook
from datetime import datetime
from zoneinfo import ZoneInfo

import requests
import os
import threading
import queue
import time
import re


# ==========================================
# CONFIGURACIÓN
# ==========================================

app = Flask(__name__)

PHONE_NUMBER_ID = "1248995224968397"

IMAGEN_UBICACION = "PIMENTEL.jpg"

IMAGEN_CASA = "CASA PUERTO REAL 2026.jpeg"

TOKEN_VERIFICACION = "puerto_real_2026"

ARCHIVO_EXCEL = "clientes_puerto_real.xlsx"

GOOGLE_SHEETS_URL = "https://script.google.com/macros/s/AKfycbzMSUCODNw21Jsv4cUiBpZosSBIRbHFQThC-JipvhJ1r6ok1So0cde7D4GbuFx6r4Jp/exec"

# ==========================================
# MEMORIA DE CONVERSACIONES
# ==========================================

historiales = {}


# ==========================================
# CONTROL DE DUPLICADOS Y COLA DE TRABAJO
# ==========================================

ARCHIVO_MENSAJES_PROCESADOS = "mensajes_procesados.txt"

mensajes_procesados = set()
lock_mensajes = threading.Lock()
cola_mensajes = queue.Queue()
ultimo_mensaje_por_numero = {}

ultimos_mensajes_por_cliente = {}
VENTANA_DUPLICADO_SEGUNDOS = 12


def cargar_mensajes_procesados():
    if not os.path.exists(ARCHIVO_MENSAJES_PROCESADOS):
        return

    try:
        with open(ARCHIVO_MENSAJES_PROCESADOS, "r", encoding="utf-8") as archivo:
            for linea in archivo:
                mensaje_id = linea.strip()
                if mensaje_id:
                    mensajes_procesados.add(mensaje_id)
    except Exception as error:
        print("No se pudo cargar mensajes procesados:", error)


def registrar_mensaje_procesado(mensaje_id):
    if not mensaje_id:
        return True

    with lock_mensajes:
        if mensaje_id in mensajes_procesados:
            return False

        mensajes_procesados.add(mensaje_id)

        try:
            with open(ARCHIVO_MENSAJES_PROCESADOS, "a", encoding="utf-8") as archivo:
                archivo.write(mensaje_id + "\n")
        except Exception as error:
            print("No se pudo guardar el ID del mensaje:", error)

    return True


# ==========================================
# INFORMACIÓN DE PUERTO REAL
# ==========================================

informacion_puerto_real = """
PROYECTO: PUERTO REAL

Ubicación:
Pimentel, Chiclayo.
