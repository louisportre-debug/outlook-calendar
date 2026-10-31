/* global Office */

let cuerpoCorreo = "";
let asuntoCorreo = "";
let remitenteCorreo = "";


/* ============================================================
   INICIO DEL COMPLEMENTO
   ============================================================ */

Office.onReady(function (info) {

    if (info.host !== Office.HostType.Outlook) {

        mostrarError(
            "Este complemento debe ejecutarse en Outlook."
        );

        return;
    }


    cargarCorreo();


    document
        .getElementById("analyzeButton")
        .addEventListener(
            "click",
            analizarEvento
        );


    document
        .getElementById("reanalyzeButton")
        .addEventListener(
            "click",
            analizarEvento
        );


    document
        .getElementById("allDay")
        .addEventListener(
            "change",
            cambiarTodoElDia
        );


    document
        .getElementById("googleCalendarButton")
        .addEventListener(
            "click",
            prepararGoogleCalendar
        );

});


/* ============================================================
   LEER CORREO
   ============================================================ */

function cargarCorreo() {

    const item =
        Office.context.mailbox.item;


    if (!item) {

        mostrarError(
            "No se ha podido acceder al correo actual."
        );

        return;
    }


    asuntoCorreo =
        item.subject || "";


    document
        .getElementById("emailSubject")
        .textContent =
        asuntoCorreo || "(sin asunto)";


    if (item.from) {

        const nombre =
            item.from.displayName || "";


        const email =
            item.from.emailAddress || "";


        remitenteCorreo =
            nombre +
            (
                email
                    ? " <" + email + ">"
                    : ""
            );


        document
            .getElementById("emailSender")
            .textContent =
            remitenteCorreo;

    } else {

        document
            .getElementById("emailSender")
            .textContent =
            "(remitente no disponible)";
    }


    item.body.getAsync(

        Office.CoercionType.Text,

        function (result) {

            if (
                result.status ===
                Office.AsyncResultStatus.Succeeded
            ) {

                cuerpoCorreo =
                    result.value || "";


                document
                    .getElementById("status")
                    .textContent =
                    "✅ Correo leído correctamente.";

            } else {

                mostrarError(
                    "No se pudo leer el cuerpo del correo: " +
                    result.error.message
                );
            }

        }

    );

}


/* ============================================================
   ANALIZAR EVENTO
   ============================================================ */

function analizarEvento() {

    if (
        !cuerpoCorreo ||
        cuerpoCorreo.trim() === ""
    ) {

        mostrarError(
            "El correo no contiene texto."
        );

        return;
    }


    const datos =
        detectarEvento(
            cuerpoCorreo,
            asuntoCorreo
        );


    document
        .getElementById("eventForm")
        .hidden = false;


    document
        .getElementById("eventTitle")
        .value =
        datos.titulo;


    document
        .getElementById("startDate")
        .value =
        datos.fechaInicio;


    document
        .getElementById("endDate")
        .value =
        datos.fechaFin;


    document
        .getElementById("startTime")
        .value =
        datos.horaInicio;


    document
        .getElementById("endTime")
        .value =
        datos.horaFin;


    document
        .getElementById("eventLocation")
        .value =
        datos.lugar;


    document
        .getElementById("eventDescription")
        .value =
        datos.descripcion;


    document
        .getElementById("allDay")
        .checked =
        datos.todoElDia;


    cambiarTodoElDia();


    document
        .getElementById("status")
        .textContent =
        "✅ Evento detectado. Revisa los datos.";

}


/* ============================================================
   DETECTOR DE EVENTOS
   ============================================================ */

function detectarEvento(texto, asunto) {

    const textoCompleto =
        (asunto + "\n" + texto)
            .replace(/\r/g, "")
            .trim();


    let titulo =
        asunto ||
        "Nuevo evento";


    let fechaInicio = "";
    let fechaFin = "";

    let horaInicio = "";
    let horaFin = "";

    let todoElDia = false;


    /* --------------------------------------------------------
       DETECTAR FECHAS DD/MM/YYYY
       -------------------------------------------------------- */

    const fechas =
        textoCompleto.match(
            /\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\b/g
        );


    if (fechas && fechas.length > 0) {

        const primera =
            convertirFecha(
                fechas[0]
            );


        if (primera) {

            fechaInicio =
                primera;

            fechaFin =
                primera;
        }


        if (fechas.length >= 2) {

            const segunda =
                convertirFecha(
                    fechas[1]
                );


            if (segunda) {

                fechaFin =
                    segunda;
            }

        }

    }


    /* --------------------------------------------------------
       DETECTAR FECHAS TIPO "3 DE NOVIEMBRE"
       -------------------------------------------------------- */

    if (!fechaInicio) {

        const meses = {
            enero: 1,
            febrero: 2,
            marzo: 3,
            abril: 4,
            mayo: 5,
            junio: 6,
            julio: 7,
            agosto: 8,
            septiembre: 9,
            octubre: 10,
            noviembre: 11,
            diciembre: 12
        };


        const regexMes =
            /\b(\d{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)(?:\s+de\s+(\d{4}))?\b/gi;


        const encontrados =
            [
                ...textoCompleto.matchAll(
                    regexMes
                )
            ];


        if (encontrados.length > 0) {

            const actual =
                new Date()
                    .getFullYear();


            const f =
                encontrados[0];


            const dia =
                parseInt(f[1], 10);


            const mes =
                meses[
                    f[2].toLowerCase()
                ];


            const año =
                f[3]
                    ? parseInt(f[3], 10)
                    : actual;


            fechaInicio =
                formatoFecha(
                    año,
                    mes,
                    dia
                );


            fechaFin =
                fechaInicio;


            if (encontrados.length >= 2) {

                const f2 =
                    encontrados[1];


                const dia2 =
                    parseInt(
                        f2[1],
                        10
                    );


                const mes2 =
                    meses[
                        f2[2].toLowerCase()
                    ];


                const año2 =
                    f2[3]
                        ? parseInt(f2[3], 10)
                        : año;


                fechaFin =
                    formatoFecha(
                        año2,
                        mes2,
                        dia2
                    );
            }

        }

    }


    /* --------------------------------------------------------
       DETECTAR HORAS
       -------------------------------------------------------- */

    const horas =
        textoCompleto.match(
            /\b([01]?\d|2[0-3])[:.]([0-5]\d)\b/g
        );


    if (horas && horas.length > 0) {

        horaInicio =
            normalizarHora(
                horas[0]
            );


        if (horas.length >= 2) {

            horaFin =
                normalizarHora(
                    horas[1]
                );
        }

    }


    /* --------------------------------------------------------
       DETECTAR "TODO EL DÍA"
       -------------------------------------------------------- */

    if (
        /todo\s+el\s+d[ií]a/i.test(
            textoCompleto
        )
    ) {

        todoElDia = true;

    }


    /* --------------------------------------------------------
       SI NO HAY HORAS
       -------------------------------------------------------- */

    if (
        !horaInicio &&
        !horaFin
    ) {

        todoElDia = true;
    }


    /* --------------------------------------------------------
       LUGAR
       -------------------------------------------------------- */

    let lugar = "";


    const lugarMatch =
        textoCompleto.match(
            /(?:lugar|ubicaci[oó]n|direcci[oó]n|venue)\s*[:\-]\s*(.+)/i
        );


    if (lugarMatch) {

        lugar =
            lugarMatch[1]
                .split("\n")[0]
                .trim();

    }


    /* --------------------------------------------------------
       DESCRIPCIÓN
       -------------------------------------------------------- */

    const descripcion =
        texto
            .trim();


    return {

        titulo:
            titulo,

        fechaInicio:
            fechaInicio,

        fechaFin:
            fechaFin,

        horaInicio:
            horaInicio,

        horaFin:
            horaFin,

        todoElDia:
            todoElDia,

        lugar:
            lugar,

        descripcion:
            descripcion

    };

}


/* ============================================================
   CONVERSIÓN DE FECHAS
   ============================================================ */

function convertirFecha(valor) {

    const partes =
        valor.split(
            /[\/\-]/
        );


    if (partes.length !== 3) {
        return "";
    }


    let dia =
        parseInt(
            partes[0],
            10
        );


    let mes =
        parseInt(
            partes[1],
            10
        );


    let año =
        parseInt(
            partes[2],
            10
        );


    if (año < 100) {
        año += 2000;
    }


    return formatoFecha(
        año,
        mes,
        dia
    );

}


function formatoFecha(
    año,
    mes,
    dia
) {

    return (
        año +
        "-" +
        String(mes).padStart(2, "0") +
        "-" +
        String(dia).padStart(2, "0")
    );

}


/* ============================================================
   NORMALIZAR HORA
   ============================================================ */

function normalizarHora(valor) {

    const partes =
        valor
            .replace(".", ":")
            .split(":");


    if (partes.length !== 2) {
        return "";
    }


    return (
        String(
            parseInt(
                partes[0],
                10
            )
        ).padStart(2, "0") +
        ":" +
        partes[1]
    );

}


/* ============================================================
   TODO EL DÍA
   ============================================================ */

function cambiarTodoElDia() {

    const todoElDia =
        document
            .getElementById("allDay")
            .checked;


    const camposHora =
        document.querySelectorAll(
            ".time-field"
        );


    camposHora.forEach(
        function (campo) {

            campo.style.display =
                todoElDia
                    ? "none"
                    : "block";

        }
    );

}


/* ============================================================
   GOOGLE CALENDAR
   ============================================================ */

function prepararGoogleCalendar() {

    const evento =
        obtenerEventoFormulario();


    if (!evento.titulo) {

        mostrarGoogleStatus(
            "⚠️ El evento necesita un título."
        );

        return;
    }


    if (!evento.fechaInicio) {

        mostrarGoogleStatus(
            "⚠️ Debes indicar una fecha de inicio."
        );

        return;
    }


    if (!evento.fechaFin) {

        evento.fechaFin =
            evento.fechaInicio;
    }


    /*
       ESTA FUNCIÓN TODAVÍA NO ENVÍA NADA A GOOGLE.

       En la siguiente fase añadiremos OAuth
       y Google Calendar API.
    */


    mostrarGoogleStatus(
        "✅ Evento preparado para Google Calendar. " +
        "La conexión con Google se añadirá en la siguiente fase."
    );

}


/* ============================================================
   OBTENER DATOS DEL FORMULARIO
   ============================================================ */

function obtenerEventoFormulario() {

    return {

        titulo:
            document
                .getElementById("eventTitle")
                .value
                .trim(),

        fechaInicio:
            document
                .getElementById("startDate")
                .value,

        horaInicio:
            document
                .getElementById("startTime")
                .value,

        fechaFin:
            document
                .getElementById("endDate")
                .value,

        horaFin:
            document
                .getElementById("endTime")
                .value,

        todoElDia:
            document
                .getElementById("allDay")
                .checked,

        lugar:
            document
                .getElementById("eventLocation")
                .value
                .trim(),

        descripcion:
            document
                .getElementById("eventDescription")
                .value
                .trim()

    };

}


/* ============================================================
   MENSAJES
   ============================================================ */

function mostrarGoogleStatus(
    mensaje
) {

    const elemento =
        document.getElementById(
            "googleStatus"
        );


    elemento.hidden =
        false;


    elemento.textContent =
        mensaje;

}


function mostrarError(
    mensaje
) {

    document
        .getElementById("status")
        .textContent =
        "❌ " + mensaje;

}
