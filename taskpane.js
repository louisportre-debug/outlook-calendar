/* ============================================================
   TASKPANE.JS
   Outlook → Cloudflare → Gemini → revisión manual → Google Calendar
   ============================================================ */

const API_URL =
    "https://outlook-calendar.louisportre.workers.dev/analyze-event";

// Tiempo máximo de espera a la IA (milisegundos)
const TIEMPO_MAXIMO = 30000;

// Cuenta de Google donde se guardan los eventos
const CUENTA_GOOGLE = "lizzet.tito@gmail.com";

// Zona horaria del ordenador (ej. "Europe/Paris")
const ZONA_LOCAL =
    Intl.DateTimeFormat().resolvedOptions().timeZone || "Europe/Paris";

let datosCorreo = null;


// ============================================================
// INICIO
// ============================================================

Office.onReady(function (info) {

    if (info.host !== Office.HostType.Outlook) {
        mostrarEstado("Este complemento solo funciona en Outlook.", "error");
        return;
    }

    document.getElementById("analyzeButton")
        .addEventListener("click", function () { analizarEvento(); });

    document.getElementById("reanalyzeButton")
        .addEventListener("click", function () { analizarEvento(); });

    document.getElementById("googleCalendarButton")
        .addEventListener("click", abrirGoogleCalendar);

    document.getElementById("allDay")
        .addEventListener("change", actualizarTodoElDia);

    document.getElementById("eventTimezone").value = ZONA_LOCAL;

    // Si el panel está fijado y cambias de correo, volver a analizar
    try {
        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            function () {
                datosCorreo = null;
                document.getElementById("eventForm").hidden = true;
                analizarEvento();
            }
        );
    } catch (e) {
        console.log("ItemChanged no disponible:", e);
    }

    // Leer y analizar automáticamente al abrir
    analizarEvento();
});


// ============================================================
// LEER EL CORREO
// ============================================================

function leerCorreo() {

    return new Promise(function (resolve, reject) {

        const item = Office.context.mailbox.item;

        if (!item) {
            reject(new Error("No hay ningún correo abierto."));
            return;
        }

        // Texto plano: conserva los saltos de línea (mejor para la IA)
        item.body.getAsync(Office.CoercionType.Text, function (result) {

            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error(
                    "No se pudo leer el correo: " +
                    (result.error ? result.error.message : "error desconocido")
                ));
                return;
            }

            const fechaRecibido = item.dateTimeCreated
                ? new Date(item.dateTimeCreated)
                : new Date();

            resolve({
                subject: item.subject || "",
                sender: obtenerRemitente(item),
                body: (result.value || "")
                    .replace(/\r/g, "")
                    .replace(/\n{3,}/g, "\n\n")
                    .replace(/[ \t]{2,}/g, " ")
                    .trim(),
                // Fecha del correo: para entender "mañana", "el jueves"...
                receivedAt: fechaRecibido.toISOString(),
                receivedLocal: fechaRecibido.toLocaleString("es-ES", {
                    timeZone: ZONA_LOCAL,
                    weekday: "long", year: "numeric", month: "long",
                    day: "numeric", hour: "2-digit", minute: "2-digit"
                }),
                timezone: ZONA_LOCAL
            });
        });
    });
}


function obtenerRemitente(item) {
    try {
        if (item.from) {
            const nombre = item.from.displayName || "";
            const email = item.from.emailAddress || "";
            return nombre && email ? nombre + " <" + email + ">" : (email || nombre);
        }
    } catch (e) { /* sin remitente */ }
    return "";
}


// ============================================================
// ANALIZAR EVENTO
// ============================================================

async function analizarEvento() {

    bloquearBotones(true);
    ocultarGoogleStatus();

    try {

        mostrarEstado("Leyendo el correo...", "loading");

        datosCorreo = await leerCorreo();

        document.getElementById("emailSubject").textContent =
            datosCorreo.subject || "—";
        document.getElementById("emailSender").textContent =
            datosCorreo.sender || "—";

        if (!datosCorreo.body && !datosCorreo.subject) {
            throw new Error("El correo está vacío.");
        }

        mostrarEstado("Analizando el correo con IA...", "loading");

        // Cancelar si tarda demasiado (evita que se quede colgado)
        const controlador = new AbortController();
        const temporizador = setTimeout(function () {
            controlador.abort();
        }, TIEMPO_MAXIMO);

        let response;

        try {
            response = await fetch(API_URL, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(datosCorreo),
                signal: controlador.signal
            });
        } catch (errorRed) {
            if (errorRed.name === "AbortError") {
                throw new Error("La IA tardó más de 30 segundos. Pulsa «Volver a analizar».");
            }
            throw new Error(
                "No se pudo conectar con el servidor de Cloudflare (" +
                errorRed.message + ")."
            );
        } finally {
            clearTimeout(temporizador);
        }

        let resultado;
        try {
            resultado = await response.json();
        } catch (e) {
            throw new Error("El servidor respondió con algo que no es JSON (HTTP " + response.status + ").");
        }

        console.log("Respuesta del servidor:", resultado);

        if (!response.ok || !resultado.ok) {
            throw new Error(
                (resultado.error || "Error del servidor") +
                (resultado.details ? " — " + resultado.details : "")
            );
        }

        rellenarFormulario(resultado.event);
        document.getElementById("eventForm").hidden = false;

        if (resultado.event.es_evento === false) {
            mostrarEstado(
                "No parece que este correo contenga un evento. Puedes completar los datos a mano.",
                "error"
            );
        } else if (!resultado.event.fecha_inicio) {
            mostrarEstado(
                "No se encontró la fecha. Complétala antes de añadirlo.",
                "error"
            );
        } else {
            mostrarEstado(
                "Evento detectado. Revisa los datos antes de añadirlo.",
                "success"
            );
        }

    } catch (error) {

        console.error("Error analizando evento:", error);
        mostrarEstado("Error: " + error.message, "error");

        // Mostrar el formulario vacío para poder rellenarlo a mano
        document.getElementById("eventForm").hidden = false;

    } finally {
        bloquearBotones(false);
    }
}


// ============================================================
// FORMULARIO
// ============================================================

function rellenarFormulario(ev) {

    establecerValor("eventTitle", ev.titulo || (datosCorreo && datosCorreo.subject));
    establecerValor("startDate", ev.fecha_inicio);
    establecerValor("startTime", ev.hora_inicio);
    establecerValor("endDate", ev.fecha_fin || ev.fecha_inicio);
    establecerValor("endTime", ev.hora_fin);
    establecerValor("eventLocation", ev.lugar || ev.enlace_online);
    establecerValor("eventTimezone", ev.zona_horaria || ZONA_LOCAL);

    let descripcion = ev.descripcion || "";
    if (ev.enlace_online && ev.lugar) {
        descripcion = "Enlace: " + ev.enlace_online + "\n\n" + descripcion;
    }
    if (datosCorreo && datosCorreo.sender) {
        descripcion += (descripcion ? "\n\n" : "") +
            "Correo de: " + datosCorreo.sender;
    }
    establecerValor("eventDescription", descripcion);

    document.getElementById("allDay").checked = ev.todo_el_dia === true;
    actualizarTodoElDia();
}


function actualizarTodoElDia() {
    const todoElDia = document.getElementById("allDay").checked;
    document.querySelectorAll(".time-field").forEach(function (campo) {
        campo.hidden = todoElDia;
    });
}


function establecerValor(id, valor) {
    const el = document.getElementById(id);
    if (el) {
        el.value = valor !== null && valor !== undefined ? valor : "";
    }
}

function obtenerValor(id) {
    const el = document.getElementById(id);
    return el ? el.value.trim() : "";
}


// ============================================================
// GOOGLE CALENDAR
// Abre Google Calendar con el evento ya rellenado.
// Tú lo revisas ahí y pulsas «Guardar».
// ============================================================

function abrirGoogleCalendar() {

    const titulo = obtenerValor("eventTitle");
    const fechaInicio = obtenerValor("startDate");
    let horaInicio = obtenerValor("startTime");
    let fechaFin = obtenerValor("endDate") || fechaInicio;
    let horaFin = obtenerValor("endTime");
    const todoElDia = document.getElementById("allDay").checked;
    const zona = obtenerValor("eventTimezone") || ZONA_LOCAL;

    if (!titulo || !fechaInicio) {
        mostrarGoogleStatus("Falta el título o la fecha de inicio.", true);
        return;
    }

    if (!todoElDia && !horaInicio) {
        mostrarGoogleStatus("Falta la hora de inicio (o marca «Todo el día»).", true);
        return;
    }

    let fechas;

    if (todoElDia) {
        // En Google, la fecha final de un evento de todo el día es exclusiva
        fechas = compactarFecha(fechaInicio) + "/" +
                 compactarFecha(sumarDias(fechaFin, 1));
    } else {
        // Sin hora de fin → 1 hora de duración
        if (!horaFin) {
            const fin = sumarMinutos(fechaInicio, horaInicio, 60);
            fechaFin = fin.fecha;
            horaFin = fin.hora;
        }

        if (fechaFin + horaFin < fechaInicio + horaInicio) {
            mostrarGoogleStatus("La finalización es anterior al inicio. Revísalo.", true);
            return;
        }

        fechas = compactarFecha(fechaInicio) + "T" + compactarHora(horaInicio) + "/" +
                 compactarFecha(fechaFin) + "T" + compactarHora(horaFin);
    }

    const params = new URLSearchParams({
        authuser: CUENTA_GOOGLE,
        action: "TEMPLATE",
        text: titulo,
        dates: fechas,
        details: obtenerValor("eventDescription").substring(0, 1500),
        location: obtenerValor("eventLocation")
    });

    if (!todoElDia) {
        params.set("ctz", zona);
    }

    const url = "https://calendar.google.com/calendar/render?" + params.toString();

    console.log("Google Calendar URL:", url);

    try {
        if (Office.context.ui && Office.context.ui.openBrowserWindow) {
            Office.context.ui.openBrowserWindow(url);
        } else {
            window.open(url, "_blank");
        }
        mostrarGoogleStatus(
            "Se abrió Google Calendar en otra pestaña. Revisa el evento y pulsa «Guardar».",
            false
        );
    } catch (e) {
        mostrarGoogleStatus(
            "No se pudo abrir la ventana. Copia este enlace en el navegador:\n" + url,
            true
        );
    }
}


// "2026-11-12" → "20261112"
function compactarFecha(fecha) {
    return fecha.replace(/-/g, "");
}

// "09:30" → "093000"
function compactarHora(hora) {
    return hora.replace(":", "").substring(0, 4) + "00";
}

function sumarDias(fecha, dias) {
    const d = new Date(fecha + "T12:00:00Z");
    d.setUTCDate(d.getUTCDate() + dias);
    return d.toISOString().substring(0, 10);
}

function sumarMinutos(fecha, hora, minutos) {
    const d = new Date(fecha + "T" + hora + ":00Z");
    d.setUTCMinutes(d.getUTCMinutes() + minutos);
    const iso = d.toISOString();
    return { fecha: iso.substring(0, 10), hora: iso.substring(11, 16) };
}


// ============================================================
// UTILIDADES DE INTERFAZ
// ============================================================

function bloquearBotones(bloquear) {
    ["analyzeButton", "reanalyzeButton", "googleCalendarButton"]
        .forEach(function (id) {
            const b = document.getElementById(id);
            if (b) b.disabled = bloquear;
        });
}

function mostrarEstado(mensaje, tipo) {
    const estado = document.getElementById("status");
    if (!estado) return;
    estado.textContent = mensaje;
    estado.className = "status " + (tipo || "");
}

function mostrarGoogleStatus(mensaje, esError) {
    const el = document.getElementById("googleStatus");
    if (!el) return;
    el.hidden = false;
    el.textContent = mensaje;
    el.className = "google-status" + (esError ? " error" : "");
}

function ocultarGoogleStatus() {
    const el = document.getElementById("googleStatus");
    if (el) el.hidden = true;
}
