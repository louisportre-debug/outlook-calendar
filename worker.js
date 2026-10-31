/**
 * Cloudflare Worker - punto de entrada
 * Dirige /analyze-event al código de functions/analyze-event.js
 */

import { onRequestPost, onRequestOptions } from "./functions/analyze-event.js";

const CORS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS"
};

export default {
    async fetch(request, env, ctx) {
        const url = new URL(request.url);

        if (url.pathname === "/analyze-event") {
            if (request.method === "OPTIONS") {
                return onRequestOptions();
            }
            if (request.method === "POST") {
                return onRequestPost({ request, env, ctx });
            }
            return new Response("Método no permitido", { status: 405, headers: CORS });
        }

        return new Response(
            "Servicio del complemento de Outlook activo.",
            { headers: { "Content-Type": "text/plain; charset=UTF-8", ...CORS } }
        );
    }
};
