```javascript
/* ============================================================
   TASKPANE.JS
   Outlook → Cloudflare Worker → Gemini
   → extracción de evento → revisión manual
   ============================================================ */

const API_URL =
    "https://outlook-calendar.louisportre.workers.dev/analyze-event";

let datosCorreoOriginales = null;


// ============================================================
// INICIO
// ============================================================

Office.onReady(function (info) {

    if (info.host !== Office.HostType.Outlook) {
        return;
    }

    console.log("Outlook Add-in iniciado");

    // IDs REALES de taskpane.html
    const analyzeButton =
        document.getElementById("analyzeButton");

    const reanalyzeButton =
        document.getElementById("reanalyzeButton");

    const googleCalendarButton =
        document.getElementById("googleCalendarButton");


    // Analizar evento
    if (analyzeButton) {

        analyzeButton.addEventListener(
            "click",
            function () {
                analizarEvento(false);
            }
        );
    }


    // Volver a analizar
    if (reanalyzeButton) {

        reanalyzeButton.addEventListener(
            "click",
            function () {
                analizarEvento(true);
            }
        );
    }


    // Google Calendar
    if (googleCalendarButton) {

        googleCalendarButton.addEventListener(
            "click",
            prepararGoogleCalendar
        );
    }


    // Leer automáticamente el correo
    cargarCorreo();
});


// ============================================================
// CARGAR CORREO DE OUTLOOK
// ============================================================

function cargarCorreo() {

    const item =
        Office.context.mailbox.item;


    if (!item) {

        mostrarEstado(
            "No se pudo acceder al correo.",
            "error"
        );

        return;
    }


    // Mostrar asunto inmediatamente
    const emailSubject =
        document.getElementById("emailSubject");

    if (emailSubject) {

        emailSubject.textContent =
            item.subject || "—";
    }


    // Obtener remitente
    const emailSender =
        document.getElementById("emailSender");

    if (emailSender) {

        const sender =
            obtenerRemitente(item);

        emailSender.textContent =
            sender || "—";
    }


    // Leer cuerpo
    item.body.getAsync(
        Office.CoercionType.Html,
        function (result) {

            if (
                result.status !==
                Office.AsyncResultStatus.Succeeded
            ) {

                mostrarEstado(
                    "No se pudo leer el contenido del correo.",
                    "error"
                );

                return;
            }


            const html =
                result.value || "";


            const texto =
                limpiarHTML(html);


            datosCorreoOriginales = {

                subject:
                    item.subject || "",

                sender:
                    obtenerRemitente(item),

                body:
                    texto
            };


            console.log(
                "Correo cargado correctamente:",
                datosCorreoOriginales
            );


            mostrarEstado(
                "Correo leído. Pulsa «Analizar evento».",
                "normal"
            );
        }
    );
}


// ============================================================
// OBTENER REMITENTE
// ============================================================

function obtenerRemitente(item) {

    try {

        if (
            item.from &&
            item.from.emailAddress
        ) {

            return item.from.emailAddress;
        }

    } catch (error) {

        console.log(
            "No se pudo obtener el remitente:",
            error
        );
    }


    return "";
}


// ============================================================
// LIMPIAR HTML
// ============================================================

function limpiarHTML(html) {

    const parser =
        new DOMParser();


    const doc =
        parser.parseFromString(
            html,
            "text/html"
        );


    // Eliminar elementos innecesarios
    doc.querySelectorAll(
        "script, style, noscript"
    ).forEach(function (elemento) {

        elemento.remove();
    });


    let texto =
        doc.body
            ? doc.body.innerText
            : html;


    // Normalizar espacios
    texto = texto
        .replace(/\r/g, "")
        .replace(/\n{3,}/g, "\n\n")
        .replace(/[ \t]{2,}/g, " ")
        .trim();


    return texto;
}


// ============================================================
// ANALIZAR EVENTO
// ============================================================

async function analizarEvento(releerCorreo) {

    try {

        bloquearBotones(true);


        mostrarEstado(
            "Analizando el correo con IA...",
            "loading"
        );


        // Si se pulsa "Volver a analizar",
        // volvemos a leer el correo.
        if (
            releerCorreo ||
            !datosCorreoOriginales
        ) {

            await cargarCorreoAsync();
        }


        if (!datosCorreoOriginales) {

            throw new Error(
                "No se pudo obtener el contenido del correo."
            );
        }


        console.log(
            "Enviando correo a Cloudflare:",
            datosCorreoOriginales
        );


        // ====================================================
        // CLOUDFLARE → GEMINI
        // ====================================================

        const response =
            await fetch(
                API_URL,
                {
                    method: "POST",

                    headers: {
                        "Content-Type": "application/json"
                    },

                    body: JSON.stringify(
                        datosCorreoOriginales
                    )
                }
            );


        console.log(
            "Respuesta HTTP Cloudflare:",
            response.status
        );


        // ====================================================
        // ERROR HTTP
        // ====================================================

        if (!response.ok) {

            const textoError =
                await response.text();


            throw new Error(
                "Error del servidor (" +
                response.status +
                "): " +
                textoError
            );
        }


        // ====================================================
        // LEER JSON
        // ====================================================

        const resultado =
            await response.json();


        console.log(
            "Respuesta de Cloudflare/Gemini:",
            resultado
        );


        if (!resultado.ok) {

            throw new Error(
                resultado.error ||
                "La IA no pudo analizar el correo."
            );
        }


        if (!resultado.event) {

            throw new Error(
                "La IA no devolvió los datos del evento."
            );
        }


        // ====================================================
        // RELLENAR FORMULARIO ORIGINAL
        // ====================================================

        rellenarFormulario(
            resultado.event
        );


        // Mostrar formulario
        const eventForm =
            document.getElementById("eventForm");

        if (eventForm) {

            eventForm.hidden = false;
        }


        mostrarEstado(
            "Evento detectado. Revísalo y modifica lo que necesites antes de añadirlo al calendario.",
            "success"
        );


    } catch (error) {

        console.error(
            "Error analizando evento:",
            error
        );


        mostrarEstado(
            "Error: " + error.message,
            "error"
        );


    } finally {

        bloquearBotones(false);
    }
}


// ============================================================
// CARGAR CORREO COMO PROMESA
// ============================================================

function cargarCorreoAsync() {

    return new Promise(
        function (resolve, reject) {

            const item =
                Office.context.mailbox.item;


            if (!item) {

                reject(
                    new Error(
                        "No se pudo acceder al correo."
                    )
                );

                return;
            }


            item.body.getAsync(
                Office.CoercionType.Html,
                function (result) {

                    if (
                        result.status !==
                        Office.AsyncResultStatus.Succeeded
                    ) {

                        reject(
                            new Error(
                                "No se pudo leer el correo."
                            )
                        );

                        return;
                    }


                    const texto =
                        limpiarHTML(
                            result.value || ""
                        );


                    datosCorreoOriginales = {

                        subject:
                            item.subject || "",

                        sender:
                            obtenerRemitente(item),

                        body:
                            texto
                    };


                    resolve();
                }
            );
        }
    );
}


// ============================================================
// RELLENAR FORMULARIO ORIGINAL
// ============================================================

function rellenarFormulario(evento) {

    // TÍTULO
    establecerValor(
        "eventTitle",
        evento.titulo
    );


    // FECHA INICIO
    establecerValor(
        "startDate",
        evento.fecha_inicio
    );


    // HORA INICIO
    establecerValor(
        "startTime",
        evento.hora_inicio
    );


    // FECHA FIN
    establecerValor(
        "endDate",
        evento.fecha_fin
    );


    // HORA FIN
    establecerValor(
        "endTime",
        evento.hora_fin
    );


    // LUGAR
    establecerValor(
        "eventLocation",
        evento.lugar
    );


    // DESCRIPCIÓN
    establecerValor(
        "eventDescription",
        evento.descripcion
    );


    // TODO EL DÍA
    const allDay =
        document.getElementById("allDay");


    if (allDay) {

        allDay.checked =
            evento.todo_el_dia === true;
    }
}


// ============================================================
// ESTABLECER VALOR
// ============================================================

function establecerValor(id, valor) {

    const elemento =
        document.getElementById(id);


    if (!elemento) {
        return;
    }


    elemento.value =
        valor !== null &&
        valor !== undefined
            ? valor
            : "";
}


// ============================================================
// GOOGLE CALENDAR
// ============================================================

function prepararGoogleCalendar() {

    const evento = {

        titulo:
            obtenerValor("eventTitle"),

        fechaInicio:
            obtenerValor("startDate"),

        horaInicio:
            obtenerValor("startTime"),

        fechaFin:
            obtenerValor("endDate"),

        horaFin:
            obtenerValor("endTime"),

        todoElDia:
            obtenerCheckbox("allDay"),

        lugar:
            obtenerValor("eventLocation"),

        descripcion:
            obtenerValor("eventDescription")
    };


    console.log(
        "Evento preparado para Google Calendar:",
        evento
    );


    const googleStatus =
        document.getElementById("googleStatus");


    if (googleStatus) {

        googleStatus.hidden = false;

        googleStatus.textContent =
            "Evento revisado y preparado. La conexión con Google Calendar se añadirá en el siguiente paso.";
    }


    mostrarEstado(
        "Evento preparado correctamente.",
        "success"
    );
}


// ============================================================
// OBTENER VALOR
// ============================================================

function obtenerValor(id) {

    const elemento =
        document.getElementById(id);


    return elemento
        ? elemento.value.trim()
        : "";
}


// ============================================================
// OBTENER CHECKBOX
// ============================================================

function obtenerCheckbox(id) {

    const elemento =
        document.getElementById(id);


    return elemento
        ? elemento.checked
        : false;
}


// ============================================================
// BLOQUEAR / DESBLOQUEAR BOTONES
// ============================================================

function bloquearBotones(bloquear) {

    const botones = [

        "analyzeButton",
        "reanalyzeButton",
        "googleCalendarButton"

    ];


    botones.forEach(function (id) {

        const boton =
            document.getElementById(id);


        if (boton) {

            boton.disabled =
                bloquear;
        }
    });
}


// ============================================================
// ESTADO
// ============================================================

function mostrarEstado(mensaje, tipo) {

    const estado =
        document.getElementById("status");


    if (!estado) {
        return;
    }


    estado.textContent =
        mensaje;


    estado.className =
        "status " + (tipo || "");
}
```
