/**
 * Cloudflare Pages Function
 * Ruta: /analyze-event
 *
 * Recibe el contenido de un correo de Outlook y utiliza Gemini
 * para extraer los datos de un evento en formato JSON estructurado.
 *
 * La clave GEMINI_API_KEY se obtiene desde Cloudflare:
 * Settings -> Variables and Secrets
 */

export async function onRequestPost(context) {
    try {
        // ---------------------------------------------------------
        // 1. Comprobar API KEY
        // ---------------------------------------------------------
        const apiKey = context.env.GEMINI_API_KEY;

        if (!apiKey) {
            return jsonResponse(
                {
                    ok: false,
                    error: "No está configurada la variable GEMINI_API_KEY en Cloudflare."
                },
                500
            );
        }

        // ---------------------------------------------------------
        // 2. Leer petición de Outlook
        // ---------------------------------------------------------
        let body;

        try {
            body = await context.request.json();
        } catch (error) {
            return jsonResponse(
                {
                    ok: false,
                    error: "La petición recibida no contiene JSON válido."
                },
                400
            );
        }

        const subject = cleanText(body.subject || "");
        const sender = cleanText(body.sender || "");
        const emailBody = cleanText(body.body || "");

        if (!subject && !emailBody) {
            return jsonResponse(
                {
                    ok: false,
                    error: "No se recibió contenido suficiente del correo para analizarlo."
                },
                400
            );
        }

        // ---------------------------------------------------------
        // 3. Limitar tamaño para evitar enviar contenido innecesario
        // ---------------------------------------------------------
        const MAX_BODY_LENGTH = 30000;

        const limitedBody = emailBody.substring(0, MAX_BODY_LENGTH);

        // ---------------------------------------------------------
        // 4. Fecha actual para ayudar a interpretar fechas sin año
        // ---------------------------------------------------------
        const timezone = cleanText(body.timezone || "Europe/Paris");
        const receivedLocal = cleanText(body.receivedLocal || "");
        const receivedAt = cleanText(body.receivedAt || new Date().toISOString());

        const currentDate = receivedLocal
            ? receivedLocal + " (" + receivedAt.substring(0, 10) + ")"
            : receivedAt.substring(0, 10);

        // ---------------------------------------------------------
        // 5. Prompt de extracción
        // ---------------------------------------------------------
        const prompt = `
Eres un extractor profesional de información de eventos procedente
de correos electrónicos.

Tu trabajo es analizar el correo y extraer ÚNICAMENTE los datos del
evento que una persona necesitaría para añadirlo a un calendario.

FECHA EN QUE SE RECIBIÓ EL CORREO (úsala para interpretar "mañana",
"el próximo jueves", "esta semana", etc., y para deducir el año):
${currentDate}

ZONA HORARIA DEL USUARIO:
${timezone}

ASUNTO DEL CORREO:
${subject}

REMITENTE:
${sender}

CONTENIDO DEL CORREO:
---------------------
${limitedBody}
---------------------

REGLAS IMPORTANTES:

1. Identifica el NOMBRE REAL DEL EVENTO.
   - No confundas el nombre del destinatario, remitente, empresa,
     organizador o persona con el nombre del evento.
   - El título debe representar el evento.
   - Si el asunto contiene claramente el nombre del evento, puede
     utilizarse como ayuda.
   - No incluy saludos ni nombres personales en el título.

2. FECHAS:
   - Detecta fechas explícitas.
   - Debes distinguir SIEMPRE entre fecha de inicio y fecha de finalización.
   - Ejemplos:
       "del 10 al 13 de noviembre"
       inicio = 10 de noviembre
       fin = 13 de noviembre

       "10-13 noviembre"
       inicio = 10 de noviembre
       fin = 13 de noviembre

       "del 15 de noviembre hasta el 17 de noviembre"
       inicio = 15 de noviembre
       fin = 17 de noviembre

   - Si se menciona un año, respétalo.
   - Si no se menciona el año, utiliza el año más razonable basándote
     en la fecha actual y en el contexto del correo.
   - NO inventes una segunda fecha cuando solamente existe una.

3. HORAS:
   - Detecta hora de inicio y hora de finalización.
   - Ejemplos:
       "09:00 - 17:00"
       inicio = 09:00
       fin = 17:00

       "a las 10:00 hasta las 18:00"
       inicio = 10:00
       fin = 18:00

       "de 9 a 17"
       inicio = 09:00
       fin = 17:00

   - Convierte las horas al formato 24 horas HH:mm cuando sea posible.
   - Si no hay una hora claramente indicada, devuelve null.

4. EVENTOS DE VARIOS DÍAS:
   - Conserva las fechas diferentes de inicio y final.
   - NO conviertas un evento de varios días en un evento de un solo día.

5. TODO EL DÍA:
   - true solamente si el correo indica claramente "todo el día",
     "all day", "jornada completa" o equivalente, o si el evento
     claramente carece de horario y el contexto indica que es un
     evento de día completo.
   - Si existen horas concretas, normalmente debe ser false.

6. LUGAR:
   - Extrae el lugar físico, edificio, sala, dirección o ubicación
     claramente asociada al evento.
   - No confundas el lugar con la dirección del remitente.

7. DESCRIPCIÓN:
   - Incluye la información útil del evento.
   - Excluye firmas, saludos, avisos legales y texto claramente
     ajeno al evento cuando sea posible.

8. NO INVENTES DATOS.
   - Si un dato no aparece o no puede determinarse razonablemente,
     devuelve null o una cadena vacía según corresponda.

9. Si existen varias fechas o varios eventos en el correo, determina
   cuál es el evento principal al que se refiere el correo.
   No mezcles fechas de eventos diferentes.

10. ZONA HORARIA:
   - Si el correo indica una zona horaria (CET, CEST, GMT, UTC, "hora de
     Madrid", "hora de Lima", "EST", "PT"...), devuelve su nombre IANA
     (ej. "Europe/Madrid", "America/Lima", "America/New_York") y deja
     las horas TAL COMO aparecen en el correo, sin convertirlas.
   - Si el correo da varias zonas, usa la que coincide con ${timezone}
     si existe; si no, la primera.
   - Si no se indica ninguna, devuelve "${timezone}".

11. ENLACE ONLINE:
   - Si es una reunión online (Teams, Zoom, Google Meet, Webex...),
     devuelve el enlace para unirse en "enlace_online".
   - Si no hay enlace, devuelve null.

12. CONFIRMACIONES:
   - Los correos de confirmación de inscripción, registro o
     participación también son eventos: extrae los datos igual.

13. es_evento = false solo si el correo no contiene ningún evento,
    reunión, cita ni actividad con fecha.

14. Devuelve exclusivamente el objeto JSON solicitado por el esquema.
`;

        // ---------------------------------------------------------
        // 6. Esquema JSON que Gemini debe respetar
        // ---------------------------------------------------------
        const responseSchema = {
            type: "OBJECT",
            properties: {
                es_evento: {
                    type: "BOOLEAN",
                    description: "true si el correo contiene un evento con fecha."
                },
                titulo: {
                    type: "STRING",
                    description: "Nombre real del evento."
                },
                fecha_inicio: {
                    type: "STRING",
                    nullable: true,
                    description: "Fecha de inicio en formato YYYY-MM-DD."
                },
                hora_inicio: {
                    type: "STRING",
                    nullable: true,
                    description: "Hora de inicio en formato HH:mm (24 h)."
                },
                fecha_fin: {
                    type: "STRING",
                    nullable: true,
                    description: "Fecha de finalización en formato YYYY-MM-DD."
                },
                hora_fin: {
                    type: "STRING",
                    nullable: true,
                    description: "Hora de finalización en formato HH:mm (24 h)."
                },
                todo_el_dia: {
                    type: "BOOLEAN",
                    description: "Indica si el evento es de todo el día."
                },
                zona_horaria: {
                    type: "STRING",
                    description: "Zona horaria IANA de las horas indicadas."
                },
                lugar: {
                    type: "STRING",
                    nullable: true,
                    description: "Lugar físico o dirección del evento."
                },
                enlace_online: {
                    type: "STRING",
                    nullable: true,
                    description: "Enlace para unirse si es online."
                },
                descripcion: {
                    type: "STRING",
                    nullable: true,
                    description: "Descripción breve y útil del evento."
                }
            },
            required: [
                "es_evento",
                "titulo",
                "fecha_inicio",
                "hora_inicio",
                "fecha_fin",
                "hora_fin",
                "todo_el_dia",
                "zona_horaria",
                "lugar",
                "enlace_online",
                "descripcion"
            ],
            propertyOrdering: [
                "es_evento", "titulo", "fecha_inicio", "hora_inicio",
                "fecha_fin", "hora_fin", "todo_el_dia", "zona_horaria",
                "lugar", "enlace_online", "descripcion"
            ]
        };

        // ---------------------------------------------------------
        // 7. Llamada a Gemini
        // ---------------------------------------------------------
        //
        // Utilizamos generateContent mediante REST.
        // No necesitamos instalar ninguna librería npm.
        //
        const model = "gemini-2.5-flash-lite";

        const geminiUrl =
            "https://generativelanguage.googleapis.com/v1beta/models/" +
            model +
            ":generateContent";

        const geminiPayload = {
            contents: [
                {
                    role: "user",
                    parts: [
                        {
                            text: prompt
                        }
                    ]
                }
            ],
            generationConfig: {
                temperature: 0.1,
                responseMimeType: "application/json",
                responseSchema: responseSchema
            }
        };

        const geminiResponse = await fetch(geminiUrl, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "x-goog-api-key": apiKey
            },
            body: JSON.stringify(geminiPayload)
        });

        // ---------------------------------------------------------
        // 8. Comprobar respuesta de Gemini
        // ---------------------------------------------------------
        if (!geminiResponse.ok) {
            const errorText = await geminiResponse.text();

            console.error("Gemini API error:", errorText);

            return jsonResponse(
                {
                    ok: false,
                    error:
                        "Gemini no pudo analizar el correo.",
                    details:
                        geminiResponse.status +
                        " " +
                        errorText.substring(0, 1000)
                },
                502
            );
        }

        const geminiData = await geminiResponse.json();

        // ---------------------------------------------------------
        // 9. Obtener texto JSON generado
        // ---------------------------------------------------------
        const generatedText =
            geminiData?.candidates?.[0]?.content?.parts?.[0]?.text;

        if (!generatedText) {
            console.error(
                "Respuesta Gemini sin texto:",
                JSON.stringify(geminiData)
            );

            return jsonResponse(
                {
                    ok: false,
                    error: "Gemini no devolvió datos estructurados."
                },
                502
            );
        }

        // ---------------------------------------------------------
        // 10. Convertir JSON de Gemini a objeto
        // ---------------------------------------------------------
        let eventData;

        try {
            eventData = JSON.parse(generatedText);
        } catch (error) {
            console.error(
                "Gemini devolvió JSON no válido:",
                generatedText
            );

            return jsonResponse(
                {
                    ok: false,
                    error: "Gemini devolvió una respuesta que no pudo convertirse en JSON."
                },
                502
            );
        }

        // ---------------------------------------------------------
        // 11. Validación básica
        // ---------------------------------------------------------
        eventData = normalizeEvent(eventData, timezone);

        // ---------------------------------------------------------
        // 12. Devolver resultado a Outlook
        // ---------------------------------------------------------
        return jsonResponse({
            ok: true,
            event: eventData
        });

    } catch (error) {
        console.error("Unexpected error:", error);

        return jsonResponse(
            {
                ok: false,
                error:
                    "Error interno del servidor.",
                details:
                    error?.message || String(error)
            },
            500
        );
    }
}


// ================================================================
// FUNCIONES AUXILIARES
// ================================================================

function cleanText(value) {
    if (value === null || value === undefined) {
        return "";
    }

    return String(value)
        .replace(/\u00A0/g, " ")
        .replace(/\r\n/g, "\n")
        .replace(/\r/g, "\n")
        .trim();
}


function normalizeEvent(event, defaultTimezone) {
    const fechaInicio = isValidDateString(event.fecha_inicio)
        ? event.fecha_inicio
        : null;

    return {
        es_evento:
            event.es_evento !== false,

        zona_horaria:
            typeof event.zona_horaria === "string" &&
            /^[A-Za-z_]+(\/[A-Za-z_\-]+)*$/.test(event.zona_horaria.trim())
                ? event.zona_horaria.trim()
                : defaultTimezone,

        enlace_online:
            typeof event.enlace_online === "string" &&
            /^https?:\/\//.test(event.enlace_online.trim())
                ? event.enlace_online.trim()
                : null,

        titulo:
            typeof event.titulo === "string"
                ? event.titulo.trim()
                : "",

        fecha_inicio: fechaInicio,

        hora_inicio:
            isValidTimeString(event.hora_inicio)
                ? event.hora_inicio.padStart(5, "0")
                : null,

        fecha_fin:
            isValidDateString(event.fecha_fin)
                ? event.fecha_fin
                : fechaInicio,

        hora_fin:
            isValidTimeString(event.hora_fin)
                ? event.hora_fin.padStart(5, "0")
                : null,

        todo_el_dia:
            event.todo_el_dia === true,

        lugar:
            typeof event.lugar === "string" &&
            event.lugar.trim()
                ? event.lugar.trim()
                : null,

        descripcion:
            typeof event.descripcion === "string" &&
            event.descripcion.trim()
                ? event.descripcion.trim()
                : null
    };
}


function isValidDateString(value) {
    if (typeof value !== "string") {
        return false;
    }

    return /^\d{4}-\d{2}-\d{2}$/.test(value);
}


function isValidTimeString(value) {
    if (typeof value !== "string") {
        return false;
    }

    // Aceptar "9:00" además de "09:00"
    if (/^\d:\d\d$/.test(value)) {
        value = "0" + value;
    }

    return /^([01]\d|2[0-3]):[0-5]\d$/.test(value);
}


function jsonResponse(data, status = 200) {
    return new Response(
        JSON.stringify(data),
        {
            status,
            headers: {
                "Content-Type": "application/json; charset=UTF-8",
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Headers": "Content-Type",
                "Access-Control-Allow-Methods": "POST, OPTIONS"
            }
        }
    );
}


// ---------------------------------------------------------------
// Permitir la petición OPTIONS del navegador
// ---------------------------------------------------------------

export function onRequestOptions() {
    return new Response(null, {
        status: 204,
        headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Headers": "Content-Type",
            "Access-Control-Allow-Methods": "POST, OPTIONS"
        }
    });
}
