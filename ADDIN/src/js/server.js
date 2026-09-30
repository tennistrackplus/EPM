/**
 * ============================================================
 * DRACO BI — GATEWAY DE MICROSOFT FABRIC (HTTP -> TDS)
 * ============================================================
 * El SQL endpoint de Fabric (Lakehouse y Warehouse) solo habla TDS
 * (protocolo de SQL Server, TCP 1433). El add-in de Excel corre en un
 * navegador/WebView que no puede abrir sockets TCP, así que este servicio
 * hace de puente:
 *
 *   Excel (add-in) --HTTPS + token Entra del usuario--> gateway --TDS--> Fabric
 *
 * Principios de seguridad:
 *   - SIN credenciales propias: cada consulta usa el token del USUARIO
 *     (cabecera Authorization), así que Fabric aplica sus permisos reales.
 *   - Nada se persiste ni se registra: ni tokens ni SQL ni resultados.
 *   - Solo conecta a hosts de Fabric (*.datawarehouse.fabric.microsoft.com
 *     y *.database.fabric.microsoft.com): no es un proxy abierto.
 *   - CORS restringido a los orígenes del add-in (ALLOWED_ORIGINS).
 *
 * Variables de entorno:
 *   PORT              Puerto HTTP (Cloud Run lo inyecta; por defecto 8080)
 *   ALLOWED_ORIGINS   Orígenes permitidos, separados por comas
 *                     (por defecto https://tennistrackplus.github.io)
 *   MAX_ROWS          Máximo de filas devueltas por consulta (por defecto 200000)
 *   QUERY_TIMEOUT_MS  Timeout de cada consulta (por defecto 120000)
 */
"use strict";

const express = require("express");
const { Connection, Request } = require("tedious");

const PORT = parseInt(process.env.PORT || "8080", 10);
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || "https://tennistrackplus.github.io")
    .split(",").map(s => s.trim()).filter(Boolean);
const MAX_ROWS = parseInt(process.env.MAX_ROWS || "200000", 10);
const QUERY_TIMEOUT_MS = parseInt(process.env.QUERY_TIMEOUT_MS || "120000", 10);
const MAX_SQL_LENGTH = 1000000;

const SERVER_RE = /^[a-z0-9-]+\.(datawarehouse|database)\.fabric\.microsoft\.com$/i;
// Nombre de item (Lakehouse/Warehouse): hasta 256 caracteres, sin caracteres de control
const DATABASE_RE = /^[^\u0000-\u001f]{0,256}$/;

const app = express();
app.disable("x-powered-by");
app.use(express.json({ limit: "2mb" }));

// ---------------------------------------------------------------
// CORS
// ---------------------------------------------------------------
app.use((req, res, next) => {
    const origin = req.headers.origin;
    if (origin) {
        if (!ALLOWED_ORIGINS.includes(origin)) {
            return res.status(403).json({ error: "Origen no permitido." });
        }
        res.setHeader("Access-Control-Allow-Origin", origin);
        res.setHeader("Vary", "Origin");
        res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
        res.setHeader("Access-Control-Allow-Headers", "Authorization, Content-Type");
        res.setHeader("Access-Control-Max-Age", "600");
    }
    if (req.method === "OPTIONS") return res.status(204).end();
    next();
});

// ---------------------------------------------------------------
// Conversión de valores a algo serializable y estable en JSON
// ---------------------------------------------------------------
function formatValue(value, typeName) {
    if (value === null || value === undefined) return null;
    if (value instanceof Date) {
        const iso = value.toISOString(); // useUTC: true -> sin desplazamientos locales
        switch (typeName) {
            case "Date": return iso.slice(0, 10);
            case "Time": return iso.slice(11, 23);
            case "DateTimeOffset": return iso;
            default: return iso.replace("T", " ").replace("Z", ""); // DateTime2, DateTime...
        }
    }
    if (Buffer.isBuffer(value)) return value.toString("base64");
    if (typeof value === "bigint") return value.toString();
    return value;
}

// ---------------------------------------------------------------
// Ejecución de una consulta con el token del usuario
// ---------------------------------------------------------------
function runQuery({ server, database, token, sql }) {
    return new Promise((resolve, reject) => {
        const connection = new Connection({
            server,
            authentication: {
                type: "azure-active-directory-access-token",
                options: { token }
            },
            options: {
                database: database || undefined,
                port: 1433,
                encrypt: true,
                trustServerCertificate: false,
                connectTimeout: 30000,
                requestTimeout: QUERY_TIMEOUT_MS,
                rowCollectionOnRequestCompletion: false,
                useColumnNames: false,
                useUTC: true,
                appName: "Draco-BI-Fabric-Gateway"
            }
        });

        let columns = [];
        let rows = [];
        let truncated = false;
        let finished = false;

        const finish = (err, result) => {
            if (finished) return;
            finished = true;
            try { connection.close(); } catch (e) { /* ya cerrada */ }
            if (err) reject(err); else resolve(result);
        };

        connection.on("error", err => finish(err));

        connection.connect(err => {
            if (err) return finish(err);

            const request = new Request(sql, err2 => {
                // Si hemos cancelado nosotros por exceso de filas, no es un error
                if (err2 && !(truncated && err2.code === "ECANCEL")) return finish(err2);
                finish(null, { columns, rows, truncated });
            });

            // Si el lote devuelve varios resultados, nos quedamos con el último
            request.on("columnMetadata", meta => {
                const list = Array.isArray(meta) ? meta : Object.values(meta);
                columns = list.map(c => ({ name: c.colName, type: c.type && c.type.name }));
                rows = [];
            });

            request.on("row", cols => {
                if (rows.length >= MAX_ROWS) {
                    if (!truncated) {
                        truncated = true;
                        connection.cancel();
                    }
                    return;
                }
                rows.push(cols.map(c => formatValue(c.value, c.metadata && c.metadata.type && c.metadata.type.name)));
            });

            connection.execSql(request);
        });
    });
}

// ---------------------------------------------------------------
// Rutas
// ---------------------------------------------------------------
app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/sql/query", async (req, res) => {
    const auth = String(req.headers.authorization || "");
    const token = auth.startsWith("Bearer ") ? auth.slice(7).trim() : "";
    if (!token || token.split(".").length !== 3 || token.length > 20000) {
        return res.status(401).json({ error: "Falta el token de Microsoft Entra ID (Authorization: Bearer ...)." });
    }

    const { server, database, sql } = req.body || {};
    if (!SERVER_RE.test(String(server || ""))) {
        return res.status(400).json({ error: "El servidor debe ser un SQL endpoint de Microsoft Fabric (*.datawarehouse.fabric.microsoft.com)." });
    }
    if (database !== undefined && database !== null && !DATABASE_RE.test(String(database))) {
        return res.status(400).json({ error: "Nombre de Lakehouse/Warehouse no válido." });
    }
    if (typeof sql !== "string" || !sql.trim() || sql.length > MAX_SQL_LENGTH) {
        return res.status(400).json({ error: "Consulta SQL vacía o demasiado larga." });
    }

    const startedAt = Date.now();
    try {
        const result = await runQuery({ server, database: database || "", token, sql });
        res.json({ ...result, elapsedMs: Date.now() - startedAt });
    } catch (err) {
        const code = err && err.code;
        const message = (err && err.message) || "Error desconocido.";
        // Sin tokens ni SQL en los logs: solo el tipo de error
        console.error(`[gateway] ${code || "ERROR"} tras ${Date.now() - startedAt} ms`);
        if (code === "ELOGIN") {
            return res.status(401).json({ error: "Fabric ha rechazado el inicio de sesión: " + message });
        }
        if (code === "ESOCKET" || /Failed to connect/i.test(message)) {
            return res.status(502).json({ error: "No se pudo conectar con el SQL endpoint de Fabric (¿nombre correcto? ¿salida al puerto 1433?): " + message });
        }
        if (code === "ETIMEOUT") {
            return res.status(504).json({ error: "La consulta ha superado el tiempo máximo: " + message });
        }
        // Errores de SQL (sintaxis, objeto inexistente, permisos...): útiles para el usuario
        return res.status(400).json({ error: message });
    }
});

// ---------------------------------------------------------------
// Respaldo para la API REST de Fabric (metadatos de modelos semánticos)
// ---------------------------------------------------------------
// El add-in llama a api.fabric.microsoft.com directamente; solo si el
// navegador lo bloquea (CORS) o no puede leer las cabeceras de una
// operación asíncrona, repite la llamada aquí. Lista blanca estricta:
// solo listar workspaces/modelos, leer la definición y consultar
// operaciones. Nada de escritura.
const GUID = "[0-9a-fA-F-]{36}";
const FABRIC_API_ROUTES = [
    { method: "GET", re: new RegExp("^/v1/workspaces(\\?continuationToken=[^&]+)?$") },
    { method: "GET", re: new RegExp(`^/v1/workspaces/${GUID}/semanticModels(\\?continuationToken=[^&]+)?$`) },
    { method: "POST", re: new RegExp(`^/v1/workspaces/${GUID}/semanticModels/${GUID}/getDefinition\\?format=TMSL$`) },
    { method: "GET", re: new RegExp(`^/v1/operations/${GUID}(/result)?$`) }
];

app.post("/fabric/api", async (req, res) => {
    const auth = String(req.headers.authorization || "");
    if (!auth.startsWith("Bearer ") || auth.length > 20000) {
        return res.status(401).json({ error: "Falta el token de Microsoft Entra ID." });
    }
    const { method, path } = req.body || {};
    const allowed = FABRIC_API_ROUTES.some(r => r.method === method && r.re.test(String(path || "")));
    if (!allowed) {
        return res.status(400).json({ error: "Operación de la API de Fabric no permitida por el gateway." });
    }
    try {
        const upstream = await fetch("https://api.fabric.microsoft.com" + path, {
            method,
            headers: { "Authorization": auth, "Content-Type": "application/json" },
            body: method === "POST" ? "{}" : undefined
        });
        const text = await upstream.text();
        let body = null;
        try { body = text ? JSON.parse(text) : null; } catch (e) { body = null; }
        res.json({
            status: upstream.status,
            headers: {
                location: upstream.headers.get("location"),
                operationId: upstream.headers.get("x-ms-operation-id"),
                retryAfter: upstream.headers.get("retry-after")
            },
            body
        });
    } catch (err) {
        console.error("[gateway] Error llamando a la API de Fabric");
        res.status(502).json({ error: "No se pudo contactar con la API de Fabric." });
    }
});

// ---------------------------------------------------------------------
// Respaldo para la API executeQueries de Power BI (consultas DAX contra
// modelos semánticos, ver ADDIN/src/js/powerBiQuery.js). El add-in llama
// directamente a api.powerbi.com; solo si el navegador lo bloquea repite
// la llamada aquí. Solo se reenvía executeQueries (lectura), con el token
// del propio usuario: Power BI aplica sus permisos y su RLS.
// ---------------------------------------------------------------------
const GUID_RE = new RegExp(`^${GUID}$`);

app.post("/powerbi/executeQueries", async (req, res) => {
    const auth = String(req.headers.authorization || "");
    if (!auth.startsWith("Bearer ") || auth.length > 20000) {
        return res.status(401).json({ error: "Falta el token de Microsoft Entra ID." });
    }
    const { workspaceId, datasetId, query } = req.body || {};
    if (!GUID_RE.test(String(workspaceId || "")) || !GUID_RE.test(String(datasetId || ""))) {
        return res.status(400).json({ error: "workspaceId o datasetId no válidos." });
    }
    if (typeof query !== "string" || query.length === 0 || query.length > 200000) {
        return res.status(400).json({ error: "Consulta DAX no válida." });
    }
    try {
        const upstream = await fetch(
            `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/datasets/${datasetId}/executeQueries`,
            {
                method: "POST",
                headers: { "Authorization": auth, "Content-Type": "application/json" },
                body: JSON.stringify({ queries: [{ query }], serializerSettings: { includeNulls: true } })
            }
        );
        const text = await upstream.text();
        res.status(upstream.status).type("application/json").send(text);
    } catch (err) {
        console.error("[gateway] Error llamando a la API de Power BI");
        res.status(502).json({ error: "No se pudo contactar con la API de Power BI." });
    }
});

app.use((req, res) => res.status(404).json({ error: "Ruta no encontrada." }));

app.listen(PORT, () => {
    console.log(`[gateway] Escuchando en :${PORT}. Orígenes permitidos: ${ALLOWED_ORIGINS.join(", ")}`);
});
