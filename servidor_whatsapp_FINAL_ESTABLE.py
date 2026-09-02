from flask import Flask, request
from openai import OpenAI
from openpyxl import load_workbook, Workbook
from datetime import datetime
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
client = OpenAI()

PHONE_NUMBER_ID = "1248995224968397"
TOKEN_VERIFICACION = "puerto_real_2026"
ARCHIVO_EXCEL = "clientes_puerto_real.xlsx"


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

Tipo:
Viviendas.

Características:
- 2 pisos.
- 3 habitaciones.
- 2 baños.

Condominio:
- Condominio privado.
- Cerco perimétrico de concreto.
- Incluye 4 parques.

Construcción:
- Sistema constructivo tradicional de ladrillo y cemento.

Áreas:
- Área total: 60 m².
- Área construida: 60 m².
- Primer piso: 33 m².
- Segundo piso: 27 m².

Facilidades económicas:
- Inicial: S/ 5,000.
- Bono de Techo Propio: S/ 62,700.
- Financiamiento: 48 meses sin intereses.

Servicios básicos:
- Luz.
- Agua.
- Desagüe.

Entrega:
- Se entrega en 30 meses.
"""


# ==========================================
# EXCEL
# ==========================================

ENCABEZADOS_EXCEL = [
    "Fecha",
    "Hora",
    "Nombre",
    "WhatsApp",
    "ID WhatsApp",
    "Username",
    "Motivo de compra",
    "Número de personas",
    "Inicial",
    "Bono",
    "Ubicación",
    "Financiamiento",
    "Solicitud de asesor",
    "Horario de contacto",
    "Nivel del lead",
    "Resumen",
]


def preparar_excel():
    if not os.path.exists(ARCHIVO_EXCEL):
        libro = Workbook()
        hoja = libro.active
        hoja.title = "Clientes"
        hoja.append(ENCABEZADOS_EXCEL)
        libro.save(ARCHIVO_EXCEL)
        return

    # Asegura que la fila de encabezados tenga las 16 columnas actuales.
    try:
        libro = load_workbook(ARCHIVO_EXCEL)
        hoja = libro.active
        for columna, encabezado in enumerate(ENCABEZADOS_EXCEL, start=1):
            hoja.cell(row=1, column=columna).value = encabezado
        libro.save(ARCHIVO_EXCEL)
    except PermissionError:
        print("ERROR: CIERRA EL EXCEL ANTES DE INICIAR EL SERVIDOR.")


def dato_valido(valor):
    if valor is None:
        return False

    valor = str(valor).strip()

    if valor == "":
        return False

    if valor.lower() in [
        "no especificado",
        "no especificada",
        "desconocido",
        "desconocida",
    ]:
        return False

    return True


def guardar_cliente_excel(
    nombre,
    whatsapp,
    id_whatsapp,
    username,
    motivo_compra,
    numero_personas,
    inicial,
    bono,
    ubicacion,
    financiamiento,
    solicitud_asesor,
    horario_contacto,
    nivel_lead,
    resumen,
):
    preparar_excel()

    ahora = datetime.now()
    fecha = ahora.strftime("%d/%m/%Y")
    hora = ahora.strftime("%H:%M")

    try:
        libro = load_workbook(ARCHIVO_EXCEL)
    except PermissionError:
        print()
        print("ERROR: CIERRA EL EXCEL ANTES DE GUARDAR.")
        print()
        return

    hoja = libro.active
    fila_encontrada = None

    for fila in range(2, hoja.max_row + 1):
        whatsapp_excel = hoja.cell(row=fila, column=4).value
        id_excel = hoja.cell(row=fila, column=5).value

        if (
            str(id_excel).strip() == str(id_whatsapp).strip()
            and str(id_whatsapp).strip() not in ["", "None", "No especificado"]
        ):
            fila_encontrada = fila
            break

        if (
            str(whatsapp_excel).strip() == str(whatsapp).strip()
            and str(whatsapp).strip() not in ["", "None", "No especificado"]
        ):
            fila_encontrada = fila
            break

    nuevos_datos = {
        3: nombre,
        4: whatsapp,
        5: id_whatsapp,
        6: username,
        7: motivo_compra,
        8: numero_personas,
        9: inicial,
        10: bono,
        11: ubicacion,
        12: financiamiento,
        13: solicitud_asesor,
        14: horario_contacto,
        15: nivel_lead,
        16: resumen,
    }

    if fila_encontrada:
        hoja.cell(row=fila_encontrada, column=1).value = fecha
        hoja.cell(row=fila_encontrada, column=2).value = hora

        for columna, valor in nuevos_datos.items():
            if dato_valido(valor):
                hoja.cell(row=fila_encontrada, column=columna).value = valor

        print("CLIENTE ACTUALIZADO EN EXCEL")

    else:
        hoja.append([
            fecha,
            hora,
            nombre,
            whatsapp,
            id_whatsapp,
            username,
            motivo_compra,
            numero_personas,
            inicial,
            bono,
            ubicacion,
            financiamiento,
            solicitud_asesor,
            horario_contacto,
            nivel_lead,
            resumen,
        ])

        print("CLIENTE GUARDADO EN EXCEL")

    libro.save(ARCHIVO_EXCEL)


# ==========================================
# CONVERTIR RESPUESTA DE IA EN CAMPOS
# ==========================================


def convertir_datos_cliente(texto):
    datos = {
        "Nombre": "No especificado",
        "Motivo de compra": "No especificado",
        "Número de personas": "No especificado",
        "Inicial": "No especificado",
        "Bono": "No especificado",
        "Ubicación": "No especificado",
        "Financiamiento": "No especificado",
        "Solicitud de asesor": "No especificado",
        "Horario de contacto": "No especificado",
        "Nivel del lead": "No especificado",
        "Resumen": "No especificado",
    }

    for linea in texto.splitlines():
        if ":" not in linea:
            continue

        clave, valor = linea.split(":", 1)
        clave = clave.strip()
        valor = valor.strip()

        if clave in datos:
            datos[clave] = valor

    return datos


# ==========================================
# EXTRAER DATOS DE TODA LA CONVERSACIÓN
# ==========================================

def extraer_datos_cliente(numero):
    historial = historiales.get(numero, [])

    texto_historial = ""

    for mensaje in historial:
        texto_historial += (
            f"{mensaje['role']}: "
            f"{mensaje['content']}\n"
        )

    respuesta = client.responses.create(
        model="gpt-5.6-luna",
        reasoning={"effort": "none"},
        instructions="""
Analiza la conversación de un posible cliente interesado en Puerto Real.

Devuelve SOLO estos datos exactamente en este formato:

Nombre:
Motivo de compra:
Número de personas:
Inicial:
Bono:
Ubicación:
Financiamiento:
Solicitud de asesor:
Horario de contacto:
Nivel del lead:
Resumen:

REGLAS:
- No inventes información.
- Si no conoces un dato, escribe: No especificado
- Nivel del lead solo puede ser: FRÍO, TIBIO o CALIENTE.
- El resumen debe ser breve y útil para el equipo comercial.
- No confundas información proporcionada por el asistente con información confirmada por el cliente.
""",
        input=texto_historial,
        max_output_tokens=600,
    )

    return respuesta.output_text
# ==========================================
# VERIFICACIÓN DEL WEBHOOK
# ==========================================


@app.route("/webhook", methods=["GET"])
def verificar_webhook():
    modo = request.args.get("hub.mode")
    token = request.args.get("hub.verify_token")
    desafio = request.args.get("hub.challenge")

    print()
    print("===================================")
    print("VERIFICACIÓN DE WEBHOOK")
    print("===================================")

    if modo == "subscribe" and token == TOKEN_VERIFICACION:
        print("Webhook verificado correctamente")
        return desafio, 200

    print("Token de verificación incorrecto")
    return "Token incorrecto", 403


# ==========================================
# RECIBIR MENSAJES DE WHATSAPP
# ==========================================


@app.route("/webhook", methods=["POST"])
def recibir_mensaje():
    datos = request.get_json(silent=True) or {}

    try:
        value = datos["entry"][0]["changes"][0]["value"]
    except (KeyError, IndexError, TypeError):
        return "EVENT_RECEIVED", 200

    # Meta envía eventos de enviado, entregado, leído, etc.
    # No son mensajes del cliente y se ignoran silenciosamente.
    if "statuses" in value:
        return "EVENT_RECEIVED", 200

    mensajes = value.get("messages")
    if not mensajes:
        return "EVENT_RECEIVED", 200

    mensaje = mensajes[0]

    # Por ahora procesamos únicamente mensajes de texto.
    if mensaje.get("type") != "text":
        return "EVENT_RECEIVED", 200

    texto = mensaje.get("text", {}).get("body", "").strip()
    if not texto:
        return "EVENT_RECEIVED", 200

    contacto = value.get("contacts", [{}])[0]

    numero = (
        contacto.get("wa_id")
        or contacto.get("phone_number")
        or mensaje.get("from")
        or mensaje.get("from_user_id")
        or contacto.get("user_id")
    )

    if not numero:
        return "EVENT_RECEIVED", 200

    numero = str(numero)
    mensaje_id = mensaje.get("id")
    id_whatsapp = contacto.get("user_id", numero)

    perfil = contacto.get("profile", {})
    username = perfil.get("username", "No especificado")

    # Se registra antes de trabajar para bloquear reintentos de Meta.
    if not registrar_mensaje_procesado(mensaje_id):
        print("Mensaje duplicado ignorado:", mensaje_id)
        return "EVENT_RECEIVED", 200

    ahora = time.time()
    texto_normalizado = texto.lower()
    ultimo = ultimos_mensajes_por_cliente.get(numero)

    if ultimo:
        mismo_texto = ultimo["texto"] == texto_normalizado
        dentro_ventana = (ahora - ultimo["tiempo"]) <= VENTANA_DUPLICADO_SEGUNDOS

        if mismo_texto and dentro_ventana:
            print("Duplicado reciente ignorado:", numero, texto)
            return "EVENT_RECEIVED", 200

    ultimos_mensajes_por_cliente[numero] = {
        "texto": texto_normalizado,
        "tiempo": ahora,
    }

    print()
    print("===================================")
    print("MENSAJE RECIBIDO")
    print("===================================")
    print("Número:", numero)
    print("Mensaje:", texto)

    ultimo_mensaje_por_numero[numero] = mensaje_id

    threading.Thread(
        target=procesar_mensaje_cliente,
        args=(
            numero,
            texto,
            mensaje_id,
            id_whatsapp,
            username,
        ),
        daemon=True,
    ).start()

    return "EVENT_RECEIVED", 200


# ==========================================
# PROCESAR MENSAJE DEL CLIENTE
# ==========================================


def procesar_mensaje_cliente(
    numero,
    texto,
    mensaje_id,
    id_whatsapp,
    username,
):
    if numero not in historiales:
        historiales[numero] = []

    historiales[numero].append({
        "role": "user",
        "content": texto,
    })

    # ======================================
    # RESPUESTA DEL ASESOR IA
    # ======================================

    try:
        respuesta = client.responses.create(
            model="gpt-5",
            instructions=f"""
            
Eres el asesor virtual de WhatsApp de Puerto Real, proyecto inmobiliario ubicado en Pimentel, Chiclayo.

IDENTIFICADOR DEL CLIENTE:
{numero}

BASE DE CONOCIMIENTO:
{informacion_puerto_real}

OBJETIVO:
Conversar como un asesor humano por WhatsApp, resolver dudas de manera breve y natural y, cuando exista interés real, registrar los datos necesarios para que el equipo comercial continúe la atención.

FORMA DE HABLAR:
- Español natural de Perú.
- Cercano, amable y profesional.
- Respuestas cortas: normalmente entre 1 y 4 líneas.
- Ve directo al punto.
- No envíes listas largas salvo que el cliente las pida.
- Usa emojis de manera natural: 😊🏡✨👍📲
- Usa 1 o 2 emojis cuando encajen; no exageres.
- No seas seco ni robótico.
- No repitas saludos si la conversación ya comenzó.
- Responde primero exactamente a lo que el cliente acaba de decir.
- Haz máximo UNA pregunta por mensaje.
- No es obligatorio terminar cada mensaje con una pregunta.
- No repitas información que ya explicaste.
- Recuerda todo lo dicho anteriormente en el historial.

NOMBRE DEL CLIENTE:
- Si el cliente dice su nombre, recuérdalo durante toda la conversación.
- Llámalo por su nombre de manera natural de vez en cuando.
- Ejemplos: "Claro, Daniela 😊", "Perfecto, María 👍".
- No uses el nombre en absolutamente todos los mensajes.
- Si todavía no sabes su nombre y solicita hablar con un asesor, pregunta únicamente su nombre.

DATOS DEL CLIENTE:
El sistema debe intentar identificar de manera natural:
- Nombre.
- Motivo de compra.
- Número de personas.
- Interés en la inicial.
- Interés en Bono Techo Propio.
- Ubicación.
- Interés en financiamiento.
- Solicitud de asesor.
- Horario de contacto.

WHATSAPP:
- El sistema obtiene automáticamente el número o identificador de WhatsApp.
- NUNCA le pidas al cliente su número de celular.
- No preguntes si prefiere WhatsApp o llamada salvo que sea estrictamente necesario.

ASESOR HUMANO:
Si el cliente dice que quiere hablar con un asesor, ser contactado, comprar, separar, cotizar, visitar o avanzar:
1. Reconoce inmediatamente su intención.
2. No vuelvas a enviarle todas las características del proyecto.
3. Si no sabes su nombre, pregunta únicamente su nombre.
4. Si ya sabes su nombre pero no sabes su horario, pregunta únicamente el horario.
5. Si ya tienes nombre y horario, confirma brevemente que la solicitud quedó registrada.

Ejemplo:
Cliente: "Quiero hablar con un asesor."
Respuesta si no sabes su nombre: "¡Claro! 😊 Para registrar tu solicitud, ¿me indicas tu nombre?"

Cliente: "Soy Daniela Alcántara."
Respuesta: "Gracias, Daniela 😊 ¿En qué horario te queda mejor que el equipo comercial te escriba?"

Cliente: "A las 4 pm."
Respuesta: "Perfecto, Daniela 👍 Quedó registrada tu solicitud para las 4:00 pm. El equipo comercial podrá contactarte por WhatsApp."

IMPORTANTE:
- No digas que ya hablaste con un asesor.
- No digas que ya lo agendaste.
- No digas que ya coordinaste la cita.
- Di únicamente que la solicitud quedó registrada para el equipo comercial.

BONO TECHO PROPIO:
- Puedes informar únicamente que el Bono Techo Propio indicado para Puerto Real es de S/ 62,700.
- No determines si una persona califica al bono.
- No inventes requisitos.
- Si pregunta si califica, explica brevemente que debe ser evaluado por el equipo comercial.

INFORMACIÓN FINANCIERA CONFIRMADA:
- Inicial: S/ 5,000.
- Bono Techo Propio: S/ 62,700.
- Financiamiento: hasta 48 meses sin intereses.

MUY IMPORTANTE:
- NO inventes fórmulas.
- NO calcules cuotas si no existe información suficiente.
- NO hagas operaciones como (precio - bono - inicial) / 48.
- NO inventes precios.
- NO inventes disponibilidad.
- NO inventes promociones.
- NO inventes requisitos.
- NO inventes fechas diferentes a las confirmadas.
- Si falta información, dilo brevemente y ofrece registrarlo para que lo confirme el equipo comercial.

UBICACIÓN Y UBICACIONES:

- Si el cliente pregunta "¿dónde queda Puerto Real?", "¿dónde está ubicado?", "¿dónde queda la casa?" o algo similar:
  responde únicamente:
  "Puerto Real está ubicado en Pimentel, Chiclayo 😊🏡"

- Si el cliente pregunta por "ubicaciones", "qué ubicaciones tienen", "qué zonas manejan", "qué lugares tienen", "dónde tienen proyectos" o algo similar:
  NO inventes otras ubicaciones.
  Indica que esa información la puede brindar mejor un asesor comercial.

- Si no sabes el nombre del cliente:
  "Claro 😊 Para brindarte información sobre las ubicaciones disponibles, ¿me indicas tu nombre para registrar tu solicitud con un asesor?"

- Si ya sabes su nombre pero no sabes el horario:
  "Claro, Daniela 😊 Un asesor puede detallarte las ubicaciones disponibles. ¿En qué horario te queda mejor que te contacten?"

- Si ya tienes nombre y horario:
  confirma brevemente que la solicitud quedó registrada.

- NO ofrezcas mapas, dirección exacta, Google Maps ni enlaces de ubicación si esa información no está confirmada.
RESPUESTAS GENERALES:
- Si solamente dice "Hola": responde de forma breve y amable.
  Ejemplo: "¡Hola! 😊 Soy el asesor virtual de Puerto Real 🏡 ¿En qué puedo ayudarte?"
- Si pide información general, responde breve y primero pregunta qué desea conocer.
  Ejemplo: "¡Claro! 😊 Con gusto te ayudo con Puerto Real 🏡 ¿Qué te gustaría conocer: ubicación, características de la casa o facilidades de compra?"

- Si pregunta algo específico, responde SOLO ese punto.
- No aproveches una pregunta simple para enviar toda la información del proyecto.
- No agregues datos extra si el cliente no los pidió.

- Si agradece, responde brevemente y no fuerces otra pregunta.

No hagas una pregunta por costumbre. Pregunta solamente cuando ayude a continuar naturalmente la conversación.
""",
            input=historiales[numero],
            max_output_tokens=600,
        )

        respuesta_texto = respuesta.output_text.strip()

        # ======================================
        # RESPALDO SI OPENAI DEVUELVE VACÍO
        # ======================================

        if not respuesta_texto:

            texto_normalizado = texto.lower().strip()

            # VARIAS UBICACIONES / ZONAS
            if any(palabra in texto_normalizado for palabra in [
                "ubicaciones",
                "qué ubicaciones",
                "que ubicaciones",
                "zonas",
                "lugares",
                "dónde tienen proyectos",
                "donde tienen proyectos"
            ]):
                respuesta_texto = (
                    "Claro 😊 Esa información te la puede detallar mejor "
                    "uno de nuestros asesores comerciales. "
                    "¿Me indicas tu nombre para registrar tu solicitud?"
                )

            # UBICACIÓN DE PUERTO REAL
            elif any(palabra in texto_normalizado for palabra in [
                "donde queda",
                "dónde queda",
                "donde está",
                "dónde está",
                "ubicado",
                "ubicada"
            ]):
                respuesta_texto = (
                    "Puerto Real está ubicado en Pimentel, Chiclayo 😊🏡"
                )

            # INICIAL
            elif "inicial" in texto_normalizado:
                respuesta_texto = (
                    "La inicial para Puerto Real es de S/ 5,000 😊"
                )

            # BONO
            elif "bono" in texto_normalizado:
                respuesta_texto = (
                    "Puerto Real cuenta con el Bono Techo Propio de S/ 62,700 🏡✨"
                )

            # FINANCIAMIENTO
            elif any(palabra in texto_normalizado for palabra in [
                "financiamiento",
                "financiar",
                "cuotas",
                "cuota"
            ]):
                respuesta_texto = (
                    "Tenemos financiamiento hasta 48 meses sin intereses 😊"
                )

            # HABITACIONES
            elif any(palabra in texto_normalizado for palabra in [
                "habitaciones",
                "habitacion",
                "habitación",
                "cuartos",
                "dormitorios"
            ]):
                respuesta_texto = (
                    "Las casas cuentan con 3 habitaciones y 2 baños 🏡"
                )

            # ASESOR
            elif "asesor" in texto_normalizado:
                respuesta_texto = (
                    "¡Claro! 😊 Ya tengo registrada tu solicitud. "
                    "El equipo comercial podrá contactarte en el horario indicado."
                )

            # PRECIO
            elif any(palabra in texto_normalizado for palabra in [
                "precio",
                "costo",
                "cuanto cuesta",
                "cuánto cuesta"
            ]):
                respuesta_texto = (
                    "El precio final no está confirmado en la información que manejo 😊 "
                    "Puedo registrar tu consulta para que un asesor te lo confirme."
                )

            # CIERRE / CONFIRMACIÓN
            elif texto_normalizado in [
                "ok",
                "okey",
                "okay",
                "esta bien",
                "está bien",
                "esta bien esa informacion",
                "está bien esa información",
                "gracias",
                "muchas gracias"
            ]:
                respuesta_texto = (
                    "Perfecto 😊 Si deseas continuar con un asesor, "
                    "tu solicitud ya puede quedar registrada."
                )

            # MENSAJE SOLO CON SIGNOS
            elif texto_normalizado in ["?", "??", "???"]:
                respuesta_texto = (
                    "Aquí estoy 😊 Si quieres, puedo ayudarte con alguna duda de Puerto Real."
                )

            # RESPUESTA GENERAL
            else:
                respuesta_texto = (
                    "¡Claro! 😊 Dime qué deseas saber sobre Puerto Real."
                )

    except Exception as error:
        print("ERROR DE OPENAI:", error)
        return        

    # ======================================
    # GUARDAR RESPUESTA EN MEMORIA
    # ======================================

    historiales[numero].append({
        "role": "assistant",
        "content": respuesta_texto,
    })

    # ======================================
    # EXTRAER DATOS PARA EL CRM
    # ======================================

    try:
        datos_extraidos = extraer_datos_cliente(numero)
        datos_cliente = convertir_datos_cliente(datos_extraidos)

        telefono_detectado = re.search(
            r"\b(?:51)?9\d{8}\b",
            texto.replace(" ", ""),
        )

        whatsapp_para_excel = numero

        if telefono_detectado:
            telefono_real = telefono_detectado.group()

            if telefono_real.startswith("51"):
                telefono_real = telefono_real[2:]

            whatsapp_para_excel = telefono_real
            print("TELÉFONO REAL PARA EXCEL:", telefono_real)

        guardar_cliente_excel(
            datos_cliente["Nombre"],
            whatsapp_para_excel,
            id_whatsapp,
            username,
            datos_cliente["Motivo de compra"],
            datos_cliente["Número de personas"],
            datos_cliente["Inicial"],
            datos_cliente["Bono"],
            datos_cliente["Ubicación"],
            datos_cliente["Financiamiento"],
            datos_cliente["Solicitud de asesor"],
            datos_cliente["Horario de contacto"],
            datos_cliente["Nivel del lead"],
            datos_cliente["Resumen"],
        )

    except Exception as error:
        print("ERROR GUARDANDO CLIENTE:", error)

    # ======================================
    # MOSTRAR RESPUESTA
    # ======================================

    print()
    print("RESPUESTA DE OPENAI:")
    print(respuesta_texto)
    print()

    # ======================================
    # RESPONDER POR WHATSAPP
    # ======================================

    token_whatsapp_actual = os.getenv("WHATSAPP_TOKEN")

    if not token_whatsapp_actual:
        print("WHATSAPP_TOKEN no está configurado.")
        return

    url = f"https://graph.facebook.com/v26.0/{PHONE_NUMBER_ID}/messages"

    headers = {
        "Authorization": f"Bearer {token_whatsapp_actual}",
        "Content-Type": "application/json",
    }

    payload = {
        "messaging_product": "whatsapp",
        "type": "text",
        "text": {"body": respuesta_texto},
    }

    # La cuenta actual acepta IDs PE.* mediante recipient y números mediante to.
    if "." in str(numero):
        payload["recipient"] = numero
    else:
        payload["to"] = numero

    try:
        envio = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=20,
        )

        print("RESPUESTA DE WHATSAPP:")
        print(envio.status_code)
        print(envio.text)

    except Exception as error:
        print("ERROR ENVIANDO A WHATSAPP:", error)


# ==========================================
# TRABAJADOR DE LA COLA
# ==========================================


def trabajador_mensajes():
    while True:
        numero, texto, mensaje_id, id_whatsapp, username = cola_mensajes.get()

        try:
            procesar_mensaje_cliente(
                numero,
                texto,
                mensaje_id,
                id_whatsapp,
                username,
            )
        except Exception as error:
            print("ERROR PROCESANDO MENSAJE:", error)
        finally:
            cola_mensajes.task_done()


# ==========================================
# INICIAR TRABAJADOR
# ==========================================

preparar_excel()
cargar_mensajes_procesados()

hilo_trabajador = threading.Thread(
    target=trabajador_mensajes,
    daemon=True,
)

hilo_trabajador.start()


# ==========================================
# INICIAR SERVIDOR LOCAL
# ==========================================

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 5001))

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False,
        threaded=True,
    )
