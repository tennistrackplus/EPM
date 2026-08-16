/**
 * ============================================================
 * DRACO PLANNING · Proxy CORS opcional para Snowflake
 * ============================================================
 * Solo necesario SI, al conectar Draco Planning a Snowflake, ves en
 * la consola del navegador errores de tipo "CORS" / "Failed to
 * fetch". El dominio *.snowflakecomputing.com no siempre devuelve
 * cabeceras Access-Control-Allow-Origin para peticiones hechas
 * directamente desde un navegador en otro dominio; este Worker se
 * limita a reenviar la petición tal cual y añadir esas cabeceras.
 *
 * Despliegue (gratis, unos minutos):
 *   1. Crea una cuenta en https://workers.cloudflare.com
 *   2. `npx wrangler init draco-sf-proxy` y pega este archivo como
 *      `src/index.js` (o usa el editor web de Cloudflare).
 *   3. Configura la variable de entorno SF_ACCOUNT con tu
 *      identificador de cuenta de Snowflake (ej. "xy12345.eu-west-1").
 *   4. Despliega (`npx wrangler deploy`) y copia la URL resultante
 *      (ej. https://draco-sf-proxy.tuusuario.workers.dev).
 *   5. En js/snowflake.js, cambia `apiBase: ""` por esa URL.
 *   6. En Snowflake, actualiza OAUTH_REDIRECT_URI si publicas Draco
 *      Planning bajo un nuevo dominio (no hace falta cambiarlo por
 *      usar este proxy, el redirect sigue siendo tu propia app).
 *
 * El proxy NO ve tu token en claro más de lo necesario para
 * reenviarlo: no lo guarda ni lo registra.
 */

const ALLOWED_ORIGIN = "*"; // en producción, pon aquí el dominio exacto de tu app

export default {
    async fetch(request, env) {
        const url = new URL(request.url);

        if (request.method === "OPTIONS") {
            return new Response(null, { headers: corsHeaders() });
        }

        if (!env.SF_ACCOUNT) {
            return new Response("Falta configurar la variable de entorno SF_ACCOUNT.", { status: 500 });
        }

        const targetUrl = `https://${env.SF_ACCOUNT}.snowflakecomputing.com${url.pathname}${url.search}`;

        const init = {
            method: request.method,
            headers: stripHopByHop(request.headers),
            body: ["GET", "HEAD"].includes(request.method) ? undefined : await request.arrayBuffer()
        };

        const upstream = await fetch(targetUrl, init);
        const response = new Response(upstream.body, upstream);
        Object.entries(corsHeaders()).forEach(([k, v]) => response.headers.set(k, v));
        return response;
    }
};

function corsHeaders() {
    return {
        "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Authorization, Content-Type, Accept, X-Snowflake-Authorization-Token-Type",
        "Access-Control-Max-Age": "86400"
    };
}

function stripHopByHop(headers) {
    const h = new Headers(headers);
    h.delete("host");
    h.delete("origin");
    return h;
}
