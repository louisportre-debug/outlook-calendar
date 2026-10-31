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

    const btnAnalizar = document.getElementById("btnAnalizar");
    const btnReanalizar = document.getElementById("btnReanalizar");
    const btnCalendario = document.getElementById("btnCalendario");

    if (btnAnalizar) {
        btnAnalizar.addEventListener("click", function () {
            analizarEvento(false);
        });
    }

    if (btnReanalizar) {
        btnReanalizar.addEventListener("click", function () {
            analizarEvento(true);
        });
    }

    if (btnCalendario) {
        btnCalendario.addEventListener("click", prepararGoogleCalendar);
    }

    // Cargar automáticamente el correo
    cargarCorreo();
});


// ============================================================
// CARGAR CORREO DE OUTLOOK
// ============================================================

function cargarCorreo() {

    const item = Office.context.mailbox.item;

    if (!item) {
        mostrarEstado("No se pudo acceder al correo.", "error");
        return;
    }

    item.body.getAsync(
        Office.CoercionType.Html,
        function (result) {

            if (result.status !== Office.AsyncResultStatus.Succeeded) {

                mostrarEstado(
                    "No se pudo leer el contenido del correo.",
                    "error"
                );

                return;
            }

            const html = result.value || "";

            const texto = limpiarHTML(html);

            datosCorreoOriginales = {

                subject:
                    item.subject || "",

                sender:
                    obtenerRemitente(item),

                body:
                    texto

            };

            console.log(
                "Correo cargado:",
                datosCorreoOriginales
            );

            mostrarEstado(
                "Correo cargado. Pulsa «Analizar evento».",
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

    } catch (e) {
        console.log("No se pudo obtener remitente:", e);
    }

    return "";
}


// ============================================================
// LIMPIAR HTML DEL CORREO
// ============================================================

function limpiarHTML(html) {

    const parser = new DOMParser();

    const doc = parser.parseFromString(
        html,
        "text/html"
    );

    // Eliminar elementos que no aportan contenido
    doc.querySelectorAll(
        "script, style, noscript"
    ).forEach(function (el) {
        el.remove();
    });

    let texto = doc.body
        ? doc.body.innerText
        : html;

    // Limpiar espacios excesivos
    texto = texto
        .replace(/\r/g, "")
        .replace(/\n{3,}/g, "\n\n")
        .replace(/[ \t]{2,}/g, " ")
        .trim();

    return texto;
}


// ============================================================
// ANALIZAR EVENTO CON GEMINI
// ============================================================

async function analizarEvento(releerCorreo) {

    const btnAnalizar =
        document.getElementById("btnAnalizar");

    const btnReanalizar =
        document.getElementById("btnReanalizar");

    try {

        bloquearBotones(true);

        mostrarEstado(
            "Analizando el correo con IA...",
            "loading"
        );


        // ----------------------------------------------------
        // VOLVER A LEER EL CORREO
        // ----------------------------------------------------

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


        // ----------------------------------------------------
        // ENVIAR A CLOUDFLARE
        // ----------------------------------------------------

        const response = await fetch(
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


        // ----------------------------------------------------
        // COMPROBAR RESPUESTA HTTP
        // ----------------------------------------------------

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


        const resultado =
            await response.json();


        // ----------------------------------------------------
        // COMPROBAR RESULTADO
        // ----------------------------------------------------

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


        console.log(
            "Evento detectado:",
            resultado.event
        );


        // ----------------------------------------------------
        // PASAR RESULTADO AL FORMULARIO
        // ----------------------------------------------------

        rellenarFormulario(
            resultado.event
        );


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

    return new Promise(function (resolve, reject) {

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
                    limpiarHTML(result.value || "");


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
    });
}


// ============================================================
// RELLENAR FORMULARIO
// ============================================================

function rellenarFormulario(evento) {

    establecerValor(
        "titulo",
        evento.titulo
    );

    establecerValor(
        "fechaInicio",
        evento.fecha_inicio
    );

    establecerValor(
        "horaInicio",
        evento.hora_inicio
    );

    establecerValor(
        "fechaFin",
        evento.fecha_fin
    );

    establecerValor(
        "horaFin",
        evento.hora_fin
    );

    establecerValor(
        "lugar",
        evento.lugar
    );

    establecerValor(
        "descripcion",
        evento.descripcion
    );


    const checkbox =
        document.getElementById("todoElDia");

    if (checkbox) {

        checkbox.checked =
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
        valor || "";
}


// ============================================================
// GOOGLE CALENDAR
// ============================================================

function prepararGoogleCalendar() {

    const evento = {

        titulo:
            obtenerValor("titulo"),

        fechaInicio:
            obtenerValor("fechaInicio"),

        horaInicio:
            obtenerValor("horaInicio"),

        fechaFin:
            obtenerValor("fechaFin"),

        horaFin:
            obtenerValor("horaFin"),

        todoElDia:
            obtenerCheckbox("todoElDia"),

        lugar:
            obtenerValor("lugar"),

        descripcion:
            obtenerValor("descripcion")
    };


    console.log(
        "Evento preparado para Google Calendar:",
        evento
    );


    mostrarEstado(
        "Evento revisado y preparado. La conexión con Google Calendar se añadirá en el siguiente paso.",
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

        "btnAnalizar",
        "btnReanalizar",
        "btnCalendario"

    ];

    botones.forEach(function (id) {

        const boton =
            document.getElementById(id);

        if (boton) {
            boton.disabled = bloquear;
        }
    });
}


// ============================================================
// ESTADO
// ============================================================

function mostrarEstado(mensaje, tipo) {

    const estado =
        document.getElementById("estado");

    if (!estado) {
        return;
    }

    estado.textContent =
        mensaje;

    estado.className =
        "estado " + (tipo || "");
}
```
