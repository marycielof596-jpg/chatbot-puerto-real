        ]):
            respuesta_texto = (
                "Son 48 cuotas de S/ 938 mensuales 😊 "
                "¿Podrías manejar ese monto de cuota mensual?"
            )
        else:
            print("ENVIANDO MENSAJE A OPENAI...")

            respuesta_texto = llamar_openai_responses(
                instructions=instructions,
                input_text=input_actual,
                max_output_tokens=220,
            )

            print("OPENAI RESPONDIÓ")

    except Exception as error:
        print("ERROR DE OPENAI:", error)
        respuesta_texto = (
            "Disculpa, en este momento estoy teniendo una demora para responder. "
            "Puedes intentar nuevamente en unos segundos 😊"
        )

    historiales[numero].append({
        "role": "assistant",
        "content": respuesta_texto,
    })

    # Primero enviamos el texto para que la conversación se sienta natural.
    enviar_texto_whatsapp(numero, respuesta_texto)

    # Después, si corresponde, enviamos UNA de las dos imágenes.
    imagen = detectar_imagen_para_mensaje(texto)
    if imagen:
        url_imagen, descripcion = imagen
        enviar_imagen_whatsapp(numero, url_imagen, descripcion)

    # CRM: un error aquí nunca debe impedir responder al cliente.
    try:
        datos_extraidos = extraer_datos_cliente(numero)
        datos_cliente = convertir_datos_cliente(datos_extraidos)

        # La columna WhatsApp contiene solo un celular real o queda vacía.
        # El ID (incluido PE...) se conserva por separado en id_whatsapp.
        whatsapp_para_excel = telefono_real_cliente or ""

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

        guardar_cliente_google_sheets(
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
        print("CLIENTE PROCESADO EN CRM")
    except Exception as error:
        print("ERROR GUARDANDO CLIENTE:", error)

    print("RESPUESTA FINAL:", respuesta_texto)


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


# ==========================================
# INICIAR SERVIDOR LOCAL
# ==========================================

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 5001))

    app.run(
        host="0.0.0.0",
        port=port,
