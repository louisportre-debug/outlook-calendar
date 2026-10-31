/* global Office */

Office.onReady(function (info) {

    if (info.host !== Office.HostType.Outlook) {

        document.getElementById("status").textContent =
            "Este complemento debe ejecutarse en Outlook.";

        return;
    }

    cargarCorreo();

    document
        .getElementById("analyzeButton")
        .addEventListener(
            "click",
            analizarEvento
        );

});


function cargarCorreo() {

    const item =
        Office.context.mailbox.item;

    if (!item) {

        mostrarError(
            "No se ha podido acceder al correo actual."
        );

        return;
    }


    // ASUNTO

    const subject =
        item.subject || "(sin asunto)";

    document.getElementById("subject")
        .textContent = subject;


    // REMITENTE

    if (item.from) {

        const senderName =
            item.from.displayName || "";

        const senderEmail =
            item.from.emailAddress || "";

        document.getElementById("sender")
            .textContent =
                senderName +
                (
                    senderEmail
                        ? " <" + senderEmail + ">"
                        : ""
                );

    } else {

        document.getElementById("sender")
            .textContent =
            "(remitente no disponible)";
    }


    // CUERPO

    item.body.getAsync(
        Office.CoercionType.Text,

        function (result) {

            if (
                result.status ===
                Office.AsyncResultStatus.Succeeded
            ) {

                document.getElementById("body")
                    .value =
                    result.value;

                document.getElementById("status")
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


function analizarEvento() {

    const body =
        document.getElementById("body")
            .value;

    if (!body || body.trim() === "") {

        mostrarError(
            "El correo no contiene texto."
        );

        return;
    }


    document.getElementById("result")
        .hidden = false;


    document.getElementById("eventResult")
        .textContent =
        "El correo está preparado para analizarse.";


    document.getElementById("status")
        .textContent =
        "✅ Datos preparados.";
}


function mostrarError(mensaje) {

    document.getElementById("status")
        .textContent =
        "❌ " + mensaje;

}
