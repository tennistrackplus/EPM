/**
 * ============================================================
 * EPM ADD-IN — MODELOS SEMÁNTICOS DE MICROSOFT FABRIC
 * ============================================================
 * Permite "Abrir modelo semántico" desde Fabric/Power BI:
 *   1. Lista los workspaces y los modelos semánticos del usuario
 *      (API REST de Fabric: api.fabric.microsoft.com).
 *   2. Descarga la definición del modelo (getDefinition, formato TMSL:
 *      model.bim) con tablas, columnas, jerarquías, medidas y relaciones.
 *   3. La convierte al MISMO objeto { fact, fields } que genera
 *      LkmlImport.parseContent / semantic_model.js, así que
 *      SemanticModelStore y el taskpane lo muestran sin cambios:
 *        - Tabla de hechos  = la tabla del lado "varios" de las relaciones.
 *        - Dimensiones      = cada tabla relacionada con la de hechos
 *                             (nombre = columna de la tabla de hechos,
 *                             igual que en los modelos de BigQuery).
 *        - Atributos        = columnas visibles de cada dimensión (+ la clave).
 *        - Jerarquías       = jerarquías definidas en el modelo de Power BI.
 *        - Medidas          = medidas DAX visibles del modelo.
 *
 * Todavía NO se genera ninguna consulta (ni SQL ni DAX) contra estos
 * modelos: solo se importa su estructura. El objeto guardado lleva además
 * "source" con los identificadores del workspace y del modelo, para poder
 * consultarlo más adelante con DAX (executeQueries).
 *
 * Permisos: listar requiere rol Visor en el workspace; getDefinition
 * requiere permiso de escritura sobre el modelo semántico (rol
 * Colaborador o superior en el workspace).
 *
 * Llamadas directas desde el navegador; si el navegador las bloquea (CORS)
 * o no puede leer las cabeceras de la operación asíncrona, se repiten a
 * través del gateway de Draco (ruta /fabric/api, solo lectura de metadatos).
 */
(function () {

    const API_BASE = "https://api.fabric.microsoft.com";

    // -----------------------------------------------------------------
    // Llamadas HTTP (directas, con respaldo vía gateway)
    // -----------------------------------------------------------------
    async function directCall(method, path, token) {
        const response = await fetch(API_BASE + path, {
            method,
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: method === "POST" ? "{}" : undefined
        });
        const text = await response.text();
        let body = null;
        try { body = text ? JSON.parse(text) : null; } catch (e) { body = null; }
        return {
            status: response.status,
            headers: {
                location: response.headers.get("location"),
                operationId: response.headers.get("x-ms-operation-id"),
                retryAfter: response.headers.get("retry-after")
            },
            body
        };
    }

    async function gatewayCall(method, path, token) {
        const response = await fetch(FB.gatewayUrl() + "/fabric/api", {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({ method, path })
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok && data.error) {
            throw new Error(data.error);
        }
        return data; // { status, headers:{location, operationId, retryAfter}, body }
    }

    let useGateway = false; // se activa si el navegador bloquea la API

    async function call(method, path) {
        const token = await FB.getApiToken();
        if (!useGateway) {
            try {
                return await directCall(method, path, token);
            } catch (err) {
                // TypeError = bloqueo CORS o de red: probamos por el gateway
                if (!(err instanceof TypeError)) throw err;
                console.warn("[Fabric API] Llamada directa bloqueada, se usa el gateway:", err.message);
                useGateway = true;
            }
        }
        return gatewayCall(method, path, token);
    }

    function apiError(res, what) {
        const b = res.body || {};
        const msg = b.message || (b.error && (b.error.message || b.error.code)) || b.errorCode || ("HTTP " + res.status);
        if (res.status === 401) return new Error(`Sin autorización para ${what}. Vuelve a conectar tu conexión de Microsoft Fabric.`);
        if (res.status === 403) return new Error(`No tienes permisos para ${what}. ${msg}`);
        return new Error(`Error al ${what}: ${msg}`);
    }

    // -----------------------------------------------------------------
    // Listados
    // -----------------------------------------------------------------
    async function listAll(pathBase, what) {
        const items = [];
        let path = pathBase;
        for (let page = 0; page < 20 && path; page++) {
            const res = await call("GET", path);
            if (res.status !== 200) throw apiError(res, what);
            const body = res.body || {};
            items.push(...(body.value || []));
            path = body.continuationToken
                ? pathBase + (pathBase.includes("?") ? "&" : "?") + "continuationToken=" + encodeURIComponent(body.continuationToken)
                : null;
        }
        return items;
    }

    async function listWorkspaces() {
        const list = await listAll("/v1/workspaces", "listar los workspaces");
        return list
            .map(w => ({ id: w.id, name: w.displayName, type: w.type }))
            .sort((a, b) => a.name.localeCompare(b.name));
    }

    async function listSemanticModels(workspaceId) {
        const list = await listAll(`/v1/workspaces/${encodeURIComponent(workspaceId)}/semanticModels`, "listar los modelos semánticos");
        return list
            .map(m => ({ id: m.id, name: m.displayName, description: m.description || "" }))
            .sort((a, b) => a.name.localeCompare(b.name));
    }

    // -----------------------------------------------------------------
    // Definición del modelo (operación asíncrona)
    // -----------------------------------------------------------------
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));

    function operationIdFrom(res) {
        const h = res.headers || {};
        if (h.operationId) return h.operationId;
        const m = String(h.location || "").match(/operations\/([0-9a-f-]{36})/i);
        return m ? m[1] : null;
    }

    async function getDefinitionParts(workspaceId, modelId, onStatus) {
        const path = `/v1/workspaces/${encodeURIComponent(workspaceId)}/semanticModels/${encodeURIComponent(modelId)}/getDefinition?format=TMSL`;
        let res = await call("POST", path);

        if (res.status === 200 && res.body && res.body.definition) {
            return res.body.definition.parts || [];
        }
        if (res.status !== 202) throw apiError(res, "leer la definición del modelo");

        let opId = operationIdFrom(res);
        if (!opId && !useGateway) {
            // El navegador no deja leer las cabeceras de la operación: repetimos por el gateway
            useGateway = true;
            res = await call("POST", path);
            if (res.status === 200 && res.body && res.body.definition) return res.body.definition.parts || [];
            opId = operationIdFrom(res);
        }
        if (!opId) throw new Error("Fabric no ha devuelto el identificador de la operación de lectura del modelo.");

        let wait = Math.min(Math.max(parseInt((res.headers && res.headers.retryAfter) || "2", 10), 1), 10) * 1000;
        for (let i = 0; i < 60; i++) {
            await sleep(wait);
            wait = 2000;
            const st = await call("GET", `/v1/operations/${opId}`);
            if (st.status !== 200) throw apiError(st, "consultar el estado de la lectura del modelo");
            const status = st.body && st.body.status;
            if (onStatus) onStatus(status);
            if (status === "Succeeded") break;
            if (status === "Failed") {
                const e = st.body.error || {};
                throw new Error("Fabric no ha podido leer el modelo: " + (e.message || e.errorCode || "error desconocido"));
            }
            if (i === 59) throw new Error("La lectura del modelo está tardando demasiado. Inténtalo de nuevo.");
        }

        const result = await call("GET", `/v1/operations/${opId}/result`);
        if (result.status !== 200) throw apiError(result, "descargar la definición del modelo");
        return (result.body && result.body.definition && result.body.definition.parts) || [];
    }

    function decodeBase64Utf8(b64) {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        return new TextDecoder("utf-8").decode(bytes).replace(/^\uFEFF/, "");
    }

    function extractModelBim(parts) {
        const part = (parts || []).find(p => /(^|\/)model\.bim$/i.test(p.path || ""));
        if (!part) {
            throw new Error("La definición del modelo no incluye model.bim (formato TMSL).");
        }
        return JSON.parse(decodeBase64Utf8(part.payload));
    }

    // -----------------------------------------------------------------
    // Conversión TMSL (model.bim) -> modelo Draco { fact, fields }
    // -----------------------------------------------------------------
    function mapTmslType(t) {
        switch (String(t || "").toLowerCase()) {
            case "string": return "STRING";
            case "int64": return "INTEGER";
            case "double": return "FLOAT";
            case "decimal": return "NUMERIC";
            case "boolean": return "BOOLEAN";
            case "datetime": return "DATETIME";
            default: return String(t || "").toUpperCase();
        }
    }

    function textOf(expr) {
        return Array.isArray(expr) ? expr.join("\n") : (expr || "");
    }

    /** Tabla física de origen (Direct Lake / Import) si se puede deducir */
    function sourceTableOf(table) {
        const p = (table.partitions || [])[0];
        const src = p && p.source;
        if (src && src.entityName) {
            return { schema: src.schemaName || "", table: src.entityName };
        }
        return null;
    }

    /**
     * @param {object} bim  JSON de model.bim
     * @param {object} meta { workspaceId, workspaceName, modelId, modelName }
     * @returns {{ fact, fields, source, warnings }}
     */
    function convertTmslToDraco(bim, meta) {
        const model = (bim && bim.model) || {};
        const tables = (model.tables || []).filter(t => t && t.name);
        const byName = {};
        tables.forEach(t => { byName[t.name] = t; });

        const relationships = (model.relationships || [])
            .filter(r => r.isActive !== false)
            // En TMSL, por defecto fromCardinality = "many" y toCardinality = "one"
            .filter(r => (r.toCardinality || "one") === "one" && (r.fromCardinality || "many") === "many");

        const warnings = [];

        // Tabla de hechos: la que más veces aparece en el lado "varios";
        // si no hay relaciones, la que tiene más medidas.
        const manySideCount = {};
        relationships.forEach(r => { manySideCount[r.fromTable] = (manySideCount[r.fromTable] || 0) + 1; });
        let factName = Object.keys(manySideCount).sort((a, b) => manySideCount[b] - manySideCount[a])[0];
        if (!factName) {
            const withMeasures = tables.slice().sort((a, b) => (b.measures || []).length - (a.measures || []).length);
            factName = withMeasures[0] ? withMeasures[0].name : "";
            warnings.push("El modelo no tiene relaciones: solo se importan las medidas.");
        }

        const fields = [];

        // Dimensiones (una por relación activa desde la tabla de hechos)
        relationships.filter(r => r.fromTable === factName).forEach(r => {
            const dim = byName[r.toTable];
            if (!dim) return;

            const attributes = (dim.columns || [])
                .filter(c => c.type !== "rowNumber" && !/^RowNumber-/.test(c.name))
                .filter(c => !c.isHidden || c.name === r.toColumn)
                .map(c => ({
                    name: c.name,
                    alias: c.name,
                    dataType: mapTmslType(c.dataType),
                    isKey: c.name === r.toColumn,
                    enabled: true
                }));

            const hierarchies = (dim.hierarchies || [])
                .filter(h => !h.isHidden)
                .map(h => ({
                    name: h.name,
                    levels: (h.levels || [])
                        .slice()
                        .sort((a, b) => (a.ordinal || 0) - (b.ordinal || 0))
                        .map(l => ({ attribute: l.column }))
                }))
                .filter(h => h.levels.length > 0);

            const physical = sourceTableOf(dim);

            fields.push({
                enabled: true,
                type: "DIMENSION",
                name: r.fromColumn,               // columna de la tabla de hechos (igual que en BigQuery)
                relProject: meta.workspaceName,
                relDataset: meta.modelName,
                relTable: dim.name,               // tabla del modelo semántico
                sourceTable: physical,            // tabla física (Lakehouse), para uso futuro
                attributes,
                hierarchies
            });
        });

        // Medidas visibles de todo el modelo (suelen estar en la tabla de hechos
        // o en una tabla de medidas)
        tables.forEach(t => {
            (t.measures || []).filter(m => !m.isHidden).forEach(m => {
                fields.push({
                    enabled: true,
                    type: "MEASURE",
                    name: m.name,
                    aggregation: "dax",
                    format: m.formatString || "",
                    expression: textOf(m.expression),
                    table: t.name
                });
            });
        });

        if (!fields.some(f => f.type === "MEASURE")) {
            warnings.push("El modelo no tiene medidas visibles. Crea al menos una medida (p. ej. SUM del importe) en Fabric.");
        }
        if (!fields.some(f => f.type === "DIMENSION")) {
            warnings.push("No se han encontrado tablas de atributos relacionadas con la tabla de hechos.");
        }

        return {
            fact: {
                project: meta.workspaceName,
                dataset: meta.modelName,
                table: factName,
                sourceTable: factName && byName[factName] ? sourceTableOf(byName[factName]) : null
            },
            fields,
            source: {
                type: "fabric-semantic-model",
                workspaceId: meta.workspaceId,
                workspaceName: meta.workspaceName,
                semanticModelId: meta.modelId,
                semanticModelName: meta.modelName,
                importedAt: new Date().toISOString()
            },
            warnings
        };
    }

    /** Flujo completo: descarga + conversión */
    async function importModel(meta, onStatus) {
        if (onStatus) onStatus("Descargando definición…");
        const parts = await getDefinitionParts(meta.workspaceId, meta.modelId, s => onStatus && onStatus("Leyendo modelo… (" + s + ")"));
        if (onStatus) onStatus("Generando modelo…");
        const bim = extractModelBim(parts);
        return convertTmslToDraco(bim, meta);
    }

    const api = { listWorkspaces, listSemanticModels, importModel, convertTmslToDraco, extractModelBim };
    if (typeof window !== "undefined") window.FabricSemanticModel = api;
    if (typeof module !== "undefined") module.exports = api;

})();
