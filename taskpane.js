/* ============================================================
   OUTLOOK → GOOGLE CALENDAR
   Analizador inteligente con Gemini + Cloudflare
   ============================================================ */

let datosCorreoOriginales = {
    subject: "",
    sender: "",
    body: ""
};

let ultimoAnalisis = null;


/* ============================================================
   INICIO
   ============================================================ */

Office.onReady(async (info) => {

    if (info.host !== Office.HostType.Outlook) {
        mostrarEstado(
            "⚠️ Este complemento debe ejecutarse dentro de Outlook.",
            "error"
        );
        return;
    }

    await cargarCorreo();

    configurarEventos();

    // Analizar automáticamente al abrir el complemento
    await analizarEvento();
});


/* ============================================================
   CARGAR CORREO
   ============================================================ */

async function cargarCorreo() {

    try {

        const item = Office.context.mailbox.item;

        if (!item) {
            throw new Error("No se pudo acceder al correo actual.");
        }

        const subject = item.subject || "";

        let sender = "";

        if (item.from) {
            sender =
                item.from.displayName ||
                item.from.emailAddress ||
                "";
        }

        // --------------------------------------------------------
        // IMPORTANTE:
        // Utilizamos HTML porque Outlook Web puede conservar mejor
        // tablas, listas, enlaces y estructura del mensaje.
        // --------------------------------------------------------

        const body = await obtenerCuerpoHTML(item);

        datosCorreoOriginales = {
            subject: subject,
            sender: sender,
            body: limpiarHTML(body)
        };

        console.log(
            "Correo cargado:",
            datosCorreoOriginales
        );

    } catch (error) {

        console.error(
            "Error cargando correo:",
            error
        );

        mostrarEstado(
            "⚠️ No se pudo leer el contenido del correo.",
            "error"
        );
    }
}


/* ============================================================
   OBTENER HTML DEL CORREO
   ============================================================ */

function obtenerCuerpoHTML(item) {

    return new Promise((resolve, reject) => {

        item.body.getAsync(
            Office.CoercionType.Html,
            (result) => {

                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value || "");
                } else {
                    reject(
                        new Error(
                            result.error?.message ||
                            "No se pudo obtener el cuerpo del correo."
                        )
                    );
                }
            }
        );
    });
}


/* ============================================================
   LIMPIAR HTML
   ============================================================ */

function limpiarHTML(html) {

    if (!html) {
        return "";
    }

    const parser = new DOMParser();

    const doc = parser.parseFromString(
        html,
        "text/html"
    );

    // Eliminar elementos que no aportan información
    doc.querySelectorAll(
        "script, style, noscript, svg"
    ).forEach(el => el.remove());

    // Convertir saltos visuales en saltos de texto
    doc.querySelectorAll("br").forEach(br => {
        br.replaceWith("\n");
    });

    doc.querySelectorAll("p, div, li, tr").forEach(el => {
        el.insertAdjacentText("afterend", "\n");
    });

    let text = doc.body?.innerText || doc.body?.textContent || "";

    // Normalización
    text = text
        .replace(/\u00A0/g, " ")
        .replace(/\r/g, "")
        .replace(/[ \t]+\n/g, "\n")
        .replace(/\n{3,}/g, "\n\n")
        .trim();

    return text;
}


/* ============================================================
   CONFIGURAR BOTONES
   ============================================================ */

function configurarEventos() {

    const btnAnalizar =
        document.getElementById("btnAnalizar");

    const btnReanalizar =
        document.getElementById("btnReanalizar");

    const btnCalendario =
        document.getElementById("btnCalendario");

    if (btnAnalizar) {
        btnAnalizar.addEventListener(
            "click",
            analizarEvento
        );
    }

    if (btnReanalizar) {
        btnReanalizar.addEventListener(
            "click",
            async () => {

                // Volvemos a leer el correo antes de analizar.
                // Esto hace que "Volver a analizar" sea realmente
                // una nueva lectura + nueva llamada a Gemini.

                await cargarCorreo();

                await analizarEvento(true);
            }
        );
    }

    if (btnCalendario) {
        btnCalendario.addEventListener(
            "click",
            prepararGoogleCalendar
        );
    }
}


/* ============================================================
   ANALIZAR EVENTO CON GEMINI
   ============================================================ */

async function analizarEvento(esReanalisis = false) {

    try {

        if (
            !datosCorreoOriginales.subject &&
            !datosCorreoOriginales.body
        ) {
            await cargarCorreo();
        }

        if (
            !datosCorreoOriginales.subject &&
            !datosCorreoOriginales.body
        ) {
            throw new Error(
                "No hay contenido del correo para analizar."
            );
        }

        mostrarEstado(
            esReanalisis
                ? "🔄 Volviendo a analizar el correo..."
                : "🔎 Analizando el evento...",
            "loading"
        );

        bloquearBotones(true);

        // --------------------------------------------------------
        // URL del backend de Cloudflare
        //
        // IMPORTANTE:
        // Esta URL se sustituirá por la URL real que Cloudflare
        // te haya dado para tu proyecto.
        // --------------------------------------------------------

        const backendUrl =
            "https://outlook-calendar.pages.dev/analyze-event";

        const response = await fetch(
            backendUrl,
            {
                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify({
                    subject:
                        datosCorreoOriginales.subject,

                    sender:
                        datosCorreoOriginales.sender,

                    body:
                        datosCorreoOriginales.body
                })
            }
        );

        let result;

        try {
            result = await response.json();
        } catch (error) {
            throw new Error(
                "El servidor no devolvió una respuesta JSON válida."
            );
        }

        if (!response.ok || !result.ok) {

            console.error(
                "Respuesta del backend:",
                result
            );

            throw new Error(
                result.error ||
                "No se pudo analizar el evento."
            );
        }

        ultimoAnalisis = result.event;

        rellenarFormulario(
            ultimoAnalisis
        );

        mostrarEstado(
            esReanalisis
                ? "✅ Evento vuelto a analizar y actualizado."
                : "✅ Evento analizado correctamente.",
            "success"
        );

    } catch (error) {

        console.error(
            "Error analizando evento:",
            error
        );

        mostrarEstado(
            "⚠️ " +
            (
                error?.message ||
                "No se pudo analizar el evento."
            ),
            "error"
        );

    } finally {

        bloquearBotones(false);
    }
}


/* ============================================================
   RELLENAR FORMULARIO
   ============================================================ */

function rellenarFormulario(evento) {

    ponerValor(
        "titulo",
        evento.titulo || ""
    );

    ponerValor(
        "fechaInicio",
        evento.fecha_inicio || ""
    );

    ponerValor(
        "horaInicio",
        evento.hora_inicio || ""
    );

    ponerValor(
        "fechaFin",
        evento.fecha_fin || ""
    );

    ponerValor(
        "horaFin",
        evento.hora_fin || ""
    );

    ponerValor(
        "lugar",
        evento.lugar || ""
    );

    ponerValor(
        "descripcion",
        evento.descripcion || ""
    );

    const todoElDia =
        document.getElementById("todoElDia");

    if (todoElDia) {
        todoElDia.checked =
            evento.todo_el_dia === true;
    }

    actualizarVisibilidadHoras();
}


/* ============================================================
   HELPERS FORMULARIO
   ============================================================ */

function ponerValor(id, valor) {

    const elemento =
        document.getElementById(id);

    if (!elemento) {
        console.warn(
            "No existe el elemento:",
            id
        );
        return;
    }

    elemento.value = valor;
}


function obtenerValor(id) {

    const elemento =
        document.getElementById(id);

    return elemento
        ? elemento.value.trim()
        : "";
}


/* ============================================================
   TODO EL DÍA
   ============================================================ */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        const checkbox =
            document.getElementById("todoElDia");

        if (checkbox) {

            checkbox.addEventListener(
                "change",
                actualizarVisibilidadHoras
            );
        }

        actualizarVisibilidadHoras();
    }
);


function actualizarVisibilidadHoras() {

    const checkbox =
        document.getElementById("todoElDia");

    const horaInicio =
        document.getElementById("horaInicio");

    const horaFin =
        document.getElementById("horaFin");

    if (!checkbox) {
        return;
    }

    const ocultar =
        checkbox.checked;

    if (horaInicio) {
        horaInicio.disabled = ocultar;
    }

    if (horaFin) {
        horaFin.disabled = ocultar;
    }
}


/* ============================================================
   PREPARAR GOOGLE CALENDAR
   ============================================================ */

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
            document.getElementById("todoElDia")?.checked || false,

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
        "✅ Evento preparado. La conexión con Google Calendar se añadirá en la siguiente fase.",
        "success"
    );
}


/* ============================================================
   ESTADO VISUAL
   ============================================================ */

function mostrarEstado(
    mensaje,
    tipo = ""
) {

    const estado =
        document.getElementById("estado");

    if (!estado) {
        console.log(mensaje);
        return;
    }

    estado.textContent = mensaje;

    estado.className =
        "estado " +
        (tipo || "");
}


/* ============================================================
   BLOQUEAR BOTONES
   ============================================================ */

function bloquearBotones(bloquear) {

    const ids = [
        "btnAnalizar",
        "btnReanalizar"
    ];

    ids.forEach(id => {

        const button =
            document.getElementById(id);

        if (button) {
            button.disabled =
                bloquear;
        }
    });
}
