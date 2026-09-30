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
 *      LkmlImport.parseContent / semantic_model.js, replicando la lista
 *      de campos de Power BI:
 *        - Tabla de hechos  = la tabla del lado "varios" de las relaciones.
 *        - Cajas            = una por tabla visible, INCLUIDA la de hechos
 *                             (nombre = nombre de la tabla en Power BI).
 *        - Atributos        = columnas visibles sin Σ (+ la clave de la relación).
 *        - Jerarquías       = jerarquías definidas en el modelo de Power BI.
 *        - Medidas          = columnas con Σ (con su agregación: SUM, AVG...)
 *                             y medidas DAX visibles del modelo.
 *
 * Estos modelos NO se consultan con SQL: el objeto guardado lleva
 * source.type = "fabric-semantic-model" y los identificadores del workspace
 * y del modelo, para consultarlo con DAX (API executeQueries de Power BI).
 * Por eso aquí no se imponen las reglas del generador SQL de Draco
 * (dimensión = columna de la tabla de hechos, medidas solo en la tabla de
 * hechos...): se replica el modelo tal cual lo muestra Power BI, y cada
 * atributo y medida guarda su tabla y columna reales.
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

        // Sondeo rápido. Fabric responde con "Retry-After: 20" (antes se
        // respetaba hasta un máximo de 10 s, y era casi todo el tiempo de
        // importación), pero leer la definición de un modelo normal tarda
        // bastante menos de un segundo: se consulta enseguida y se va
        // espaciando poco a poco (0,3 s, 0,5 s, 0,8 s, 1 s... hasta 3 s).
        const schedule = [300, 500, 800, 1000, 1500, 2000];
        const startedAt = Date.now();
        for (let i = 0; ; i++) {
            await sleep(i < schedule.length ? schedule[i] : 3000);
            const st = await call("GET", `/v1/operations/${opId}`);
            if (st.status === 429) continue; // demasiadas consultas seguidas: se reintenta con la siguiente espera
            if (st.status !== 200) throw apiError(st, "consultar el estado de la lectura del modelo");
            const status = st.body && st.body.status;
            if (onStatus) onStatus(status);
            if (status === "Succeeded") break;
            if (status === "Failed") {
                const e = st.body.error || {};
                throw new Error("Fabric no ha podido leer el modelo: " + (e.message || e.errorCode || "error desconocido"));
            }
            if (Date.now() - startedAt > 120000) throw new Error("La lectura del modelo está tardando demasiado. Inténtalo de nuevo.");
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
    // Se replica la lista de campos de Power BI:
    //   - Una caja (DIMENSION) por cada tabla visible del modelo, incluida
    //     la tabla de hechos, con sus columnas como atributos.
    //   - Las columnas numéricas que Power BI resume (símbolo Σ, es decir,
    //     summarizeBy distinto de "none") se importan como MEDIDAS con su
    //     agregación, y no como atributos.
    //   - Las medidas DAX del modelo se importan como medidas "dax".
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

    const NUMERIC_TYPES = ["int64", "double", "decimal"];

    // summarizeBy de TMSL -> agregación SQL de Draco (columna AGGREGATION)
    const AGGREGATION_BY_SUMMARIZE = {
        sum: "SUM",
        average: "AVG",
        min: "MIN",
        max: "MAX",
        count: "COUNT",
        distinctcount: "COUNT_DISTINCT"
    };

    /**
     * Agregación con la que Power BI resume la columna (la que muestra con Σ),
     * o null si la columna no se resume. En TMSL, si summarizeBy falta o vale
     * "default", Power BI suma las columnas numéricas.
     */
    function aggregationOf(column) {
        const s = String(column.summarizeBy || "default").toLowerCase();
        if (s === "none") return null;
        if (s === "default") {
            return NUMERIC_TYPES.includes(String(column.dataType || "").toLowerCase()) ? "SUM" : null;
        }
        return AGGREGATION_BY_SUMMARIZE[s] || null;
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

    /** Tablas internas que Power BI no enseña en la lista de campos */
    function isSystemTable(t) {
        return !!t.isHidden || !!t.isPrivate ||
            /^(LocalDateTable_|DateTableTemplate_)/.test(t.name || "");
    }

    /** Columnas que Power BI enseña en la lista de campos */
    function isVisibleColumn(c) {
        return c && c.name && !c.isHidden && c.type !== "rowNumber" && !/^RowNumber-/.test(c.name);
    }

    /**
     * @param {object} bim  JSON de model.bim
     * @param {object} meta { workspaceId, workspaceName, modelId, modelName }
     * @returns {{ fact, fields, source, warnings }}
     */
    function convertTmslToDraco(bim, meta) {
        const model = (bim && bim.model) || {};
        const tables = (model.tables || []).filter(t => t && t.name && !isSystemTable(t));
        const byName = {};
        tables.forEach(t => { byName[t.name] = t; });

        const relationships = (model.relationships || [])
            .filter(r => r.isActive !== false)
            // En TMSL, por defecto fromCardinality = "many" y toCardinality = "one"
            .filter(r => (r.toCardinality || "one") === "one" && (r.fromCardinality || "many") === "many");

        const warnings = [];

        // Tabla de hechos: la que más veces aparece en el lado "varios";
        // si no hay relaciones, la que tiene más medidas y columnas con Σ.
        const manySideCount = {};
        relationships.forEach(r => {
            if (byName[r.fromTable]) manySideCount[r.fromTable] = (manySideCount[r.fromTable] || 0) + 1;
        });
        let factName = Object.keys(manySideCount).sort((a, b) => manySideCount[b] - manySideCount[a])[0];
        if (!factName) {
            const score = t => (t.measures || []).length + (t.columns || []).filter(c => isVisibleColumn(c) && aggregationOf(c)).length;
            const ranked = tables.slice().sort((a, b) => score(b) - score(a));
            factName = ranked[0] ? ranked[0].name : "";
            warnings.push("El modelo no tiene relaciones: las tablas se importan sin enlazar con la de hechos.");
        }

        // Relación de cada tabla con la de hechos (clave en la dimensión)
        const relToFact = {};
        relationships.filter(r => r.fromTable === factName).forEach(r => {
            if (!relToFact[r.toTable]) relToFact[r.toTable] = r;
        });

        // Orden de las cajas: hechos, dimensiones relacionadas y el resto
        const ordered = [];
        if (byName[factName]) ordered.push(byName[factName]);
        tables.filter(t => relToFact[t.name]).forEach(t => ordered.push(t));
        tables.filter(t => t.name !== factName && !relToFact[t.name]).forEach(t => ordered.push(t));

        const dimensions = [];
        const measures = [];
        const usedMeasureNames = new Set();
        const uniqueMeasureName = (name, table) => {
            let n = name;
            if (usedMeasureNames.has(n)) n = `${name} (${table})`;
            let i = 2;
            while (usedMeasureNames.has(n)) n = `${name} (${table}) ${i++}`;
            usedMeasureNames.add(n);
            return n;
        };

        ordered.forEach(t => {
            const isFact = t.name === factName;
            const rel = relToFact[t.name] || null;
            const columns = (t.columns || []).filter(c => c && c.name && c.type !== "rowNumber" && !/^RowNumber-/.test(c.name));

            // Atributos: columnas visibles sin Σ (+ la clave de la relación,
            // aunque esté oculta, porque Draco la necesita para enlazar)
            const attributes = columns
                .filter(c => (isVisibleColumn(c) && !aggregationOf(c)) || (rel && c.name === rel.toColumn))
                .map(c => ({
                    name: c.name,
                    alias: c.name,
                    dataType: mapTmslType(c.dataType),
                    isKey: !!rel && c.name === rel.toColumn,
                    enabled: true
                }));

            // Medidas: columnas visibles con Σ, con su agregación
            columns.filter(c => isVisibleColumn(c) && aggregationOf(c)).forEach(c => {
                measures.push({
                    enabled: true,
                    type: "MEASURE",
                    name: uniqueMeasureName(c.name, t.name),
                    aggregation: aggregationOf(c),
                    field: c.name,                  // columna real (FACT_FIELD)
                    table: t.name,
                    format: c.formatString || "",
                    dataType: mapTmslType(c.dataType)
                });
            });

            // Medidas DAX de la tabla
            (t.measures || []).filter(m => !m.isHidden).forEach(m => {
                measures.push({
                    enabled: true,
                    type: "MEASURE",
                    name: uniqueMeasureName(m.name, t.name),
                    aggregation: "dax",
                    daxName: m.name,                // nombre real de la medida en Power BI ([Medida])
                    format: m.formatString || "",
                    expression: textOf(m.expression),
                    table: t.name
                });
            });

            if (attributes.length === 0) return; // p. ej. una tabla solo de medidas

            const hierarchies = (t.hierarchies || [])
                .filter(h => !h.isHidden)
                .map(h => ({
                    name: h.name,
                    levels: (h.levels || [])
                        .slice()
                        .sort((a, b) => (a.ordinal || 0) - (b.ordinal || 0))
                        .map(l => ({ attribute: l.column }))
                }))
                .filter(h => h.levels.length > 0);

            dimensions.push({
                enabled: true,
                type: "DIMENSION",
                name: t.name,                         // igual que en la lista de campos de Power BI
                relProject: meta.workspaceName,
                relDataset: meta.modelName,
                relTable: t.name,                     // tabla del modelo semántico
                isFactTable: isFact,                  // caja de la propia tabla de hechos
                joinColumn: rel ? rel.fromColumn : null,  // columna de la tabla de hechos que enlaza
                keyColumn: rel ? rel.toColumn : null,     // clave en esta tabla
                related: isFact || !!rel,
                sourceTable: sourceTableOf(t),        // tabla física (Lakehouse), para uso futuro
                attributes,
                hierarchies
            });

            if (!isFact && !rel) {
                warnings.push(`La tabla "${t.name}" no está relacionada con "${factName}": se importa, pero no filtrará los datos.`);
            }
        });

        const fields = [...dimensions, ...measures];

        if (measures.length === 0) {
            warnings.push("El modelo no tiene medidas ni columnas numéricas resumibles (Σ).");
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
