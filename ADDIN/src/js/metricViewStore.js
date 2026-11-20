/**
 * DracoMetricView
 * ------------------------------------------------------------------------
 * Abre en Draco una Metric View de Databricks CREADA POR DRACO.
 *
 * Databricks guarda el YAML de la Metric View sin los comentarios «#», así
 * que los datos propios de Draco (jerarquías, textos de clave, formatos,
 * agrupación de atributos...) viajan en las PROPIEDADES de la vista:
 *
 *   draco.generator    = 'Draco'
 *   draco.model.format = 'json-gzip-base64-v1'   (o 'json-base64-v1')
 *   draco.model.name   = nombre del modelo semántico
 *   draco.model        = modelo semántico completo de Draco (JSON,
 *                        comprimido con gzip y en base64)
 *
 * La Metric View solo sirve para CREAR el modelo semántico: los informes
 * siguen lanzando el SQL de siempre contra las tablas, nunca contra la
 * Metric View. Una Metric View sin draco.model (hecha a mano o por otra
 * herramienta) se abre también: se monta el modelo desde su descripción y,
 * solo en ese caso, los informes consultan la Metric View con MEASURE().
 *
 * SEGURIDAD (F-07): draco.model lo puede escribir cualquiera con permiso
 * ALTER sobre la vista, así que se trata como un fichero importado: se
 * reconstruye el modelo campo a campo (solo las claves conocidas) y se
 * validan todos los nombres de catálogo, esquema, tabla, columna y medida.
 * No se copia SQL libre al modelo.
 */
(function () {

    const GENERATOR = "Draco";
    const FORMATS = ["json-gzip-base64-v1", "json-base64-v1"];
    const MAX_VALUE = 2 * 1024 * 1024;   // 2 MB de texto en la propiedad, como mucho

    const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]{0,254}$/;
    const CATALOG_RE = /^[A-Za-z_][A-Za-z0-9_-]{0,254}$/;
    // MEASURE: medida de una Metric View ajena (la calcula Databricks con MEASURE())
    const AGGS = ["SUM", "COUNT", "AVG", "MIN", "MAX", "COUNT_DISTINCT", "MEASURE"];
    const KINDS = ["FILTERED", "CALC"];
    const OPS = ["IN", "NOT_IN"];

    /* ---------------- descompresión ---------------- */

    function b64ToBytes(s) {
        if (!/^[A-Za-z0-9+/=\s]+$/.test(s)) throw new Error("draco.model no es base64 válido.");
        const bin = atob(String(s).replace(/\s+/g, ""));
        const out = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
        return out;
    }

    async function gunzip(bytes) {
        if (typeof DecompressionStream === "undefined") {
            throw new Error("Esta versión de Excel no puede descomprimir el modelo (falta DecompressionStream). Actualiza Office.");
        }
        const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"));
        return await new Response(stream).text();
    }

    async function decode(format, value) {
        if (!FORMATS.includes(format)) throw new Error(`Formato de draco.model no soportado: «${String(format).slice(0, 40)}».`);
        if (!value || value.length > MAX_VALUE) throw new Error("draco.model está vacío o es demasiado grande.");
        const bytes = b64ToBytes(value);
        const text = format === "json-gzip-base64-v1" ? await gunzip(bytes) : new TextDecoder().decode(bytes);
        try { return JSON.parse(text); } catch (e) { throw new Error("draco.model no contiene un modelo válido."); }
    }

    /* ---------------- validación / reconstrucción ---------------- */

    function ident(v, what, re) {
        const s = String(v == null ? "" : v);
        if (!(re || IDENT_RE).test(s)) throw new Error(`El modelo de la Metric View tiene un nombre de ${what} no válido: "${s.slice(0, 80)}".`);
        return s;
    }
    const optIdent = (v, what, re) => (v === undefined || v === null || v === "") ? "" : ident(v, what, re);
    const str = (v, max) => String(v == null ? "" : v).slice(0, max || 1000);
    const isObj = (v) => v && typeof v === "object" && !Array.isArray(v);

    function cleanAttribute(a) {
        if (!isObj(a)) throw new Error("Atributo con formato no válido en draco.model.");
        const out = { name: ident(a.name, "atributo"), alias: str(a.alias || a.name, 255), dataType: str(a.dataType, 40),
            isKey: a.isKey === true, enabled: a.enabled !== false };
        if (a.calc === true) { out.calc = true; out.expression = str(a.expression, 20000); }
        return out;
    }

    function cleanDimension(d) {
        const out = {
            enabled: d.enabled !== false, type: "DIMENSION", name: ident(d.name, "campo"),
            relProject: optIdent(d.relProject, "catálogo", CATALOG_RE),
            relDataset: optIdent(d.relDataset, "esquema"),
            relTable: optIdent(d.relTable, "tabla"),
            attributes: (Array.isArray(d.attributes) ? d.attributes : []).map(cleanAttribute),
            hierarchies: (Array.isArray(d.hierarchies) ? d.hierarchies : []).map(h => {
                if (!isObj(h)) throw new Error("Jerarquía con formato no válido en draco.model.");
                const hh = { name: str(h.name, 255), levels: (Array.isArray(h.levels) ? h.levels : []).map(l => ({
                    attribute: ident(l && l.attribute, "nivel de jerarquía"),
                    textAttribute: optIdent(l && l.textAttribute, "atributo de texto")
                })) };
                if (h.emptyValue !== undefined && h.emptyValue !== "") hh.emptyValue = str(h.emptyValue, 255);
                return hh;
            }),
            textAttribute: optIdent(d.textAttribute, "atributo de texto"),
            // Cálculos entre miembros: datos del lenguaje de fórmulas de Draco (como en un .lkml importado)
            calcMembers: isObj(d.calcMembers) ? JSON.parse(JSON.stringify(d.calcMembers)) : {}
        };
        if (d.isGroup === true) out.isGroup = true;   // grupo de columnas de la tabla de hechos (editor de modelos)
        return out;
    }

    function cleanMeasure(m) {
        const out = { enabled: m.enabled !== false, type: "MEASURE", name: ident(m.name, "medida"),
            alias: str(m.alias || m.name, 255),
            aggregation: AGGS.includes(String(m.aggregation || "SUM").toUpperCase()) ? String(m.aggregation || "SUM").toUpperCase() : "SUM",
            format: str(m.format || "Auto", 40) };
        if (m.dataType) out.dataType = str(m.dataType, 40);
        if (m.table) out.table = ident(m.table, "tabla");
        if (m.field) out.field = ident(m.field, "columna");
        if (m.measureKind) {
            if (!KINDS.includes(m.measureKind)) throw new Error(`Tipo de medida no válido en draco.model: «${str(m.measureKind, 40)}».`);
            out.measureKind = m.measureKind;
            out.formula = str(m.formula, 20000);
            out.base = optIdent(m.base, "medida base");
            out.conditions = (Array.isArray(m.conditions) ? m.conditions : []).map(c => ({
                column: ident(c && c.column, "columna de filtro"),
                dataType: str(c && c.dataType, 40),
                op: OPS.includes(c && c.op) ? c.op : "IN",
                values: (Array.isArray(c && c.values) ? c.values : []).map(v => (typeof v === "number" ? v : str(v, 1000)))
            }));
        }
        return out;
    }

    /** Reconstruye y valida el modelo leído de draco.model (solo claves conocidas). */
    function sanitizeModel(raw) {
        if (!isObj(raw) || !isObj(raw.fact) || !Array.isArray(raw.fields)) {
            throw new Error("draco.model no tiene el formato de un modelo semántico de Draco.");
        }
        const fact = {
            project: optIdent(raw.fact.project, "catálogo", CATALOG_RE),
            dataset: optIdent(raw.fact.dataset, "esquema"),
            table: ident(raw.fact.table, "tabla de hechos")
        };
        const fields = raw.fields.map(f => {
            if (!isObj(f)) throw new Error("Campo con formato no válido en draco.model.");
            if (f.type === "DIMENSION") return cleanDimension(f);
            if (f.type === "MEASURE") return cleanMeasure(f);
            throw new Error(`Tipo de campo no válido en draco.model: «${str(f.type, 40)}».`);
        });
        const model = { fact, fields };
        if (raw.dracoPlanning === true) model.dracoPlanning = true;
        // Modelo montado sobre una Metric View ajena: los informes la consultan con MEASURE()
        if (isObj(raw.metricView)) model.metricView = {
            catalog: ident(raw.metricView.catalog, "catálogo", CATALOG_RE),
            schema: ident(raw.metricView.schema, "esquema"),
            view: ident(raw.metricView.view, "vista")
        };
        return model;
    }

    /* ---------------- compresión (al guardar) ---------------- */

    function bytesToB64(bytes) {
        let bin = "";
        for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
        return btoa(bin);
    }

    async function encode(model) {
        const json = JSON.stringify(model);
        const bytes = new TextEncoder().encode(json);
        if (typeof CompressionStream !== "undefined") {
            const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream("gzip"));
            const gz = new Uint8Array(await new Response(stream).arrayBuffer());
            return { format: "json-gzip-base64-v1", value: bytesToB64(gz) };
        }
        return { format: "json-base64-v1", value: bytesToB64(bytes) };
    }

    /* ---------------- Metric Views ajenas ---------------- */

    /**
     * Modelo semántico a partir de la descripción de una Metric View que no
     * es de Draco (DESCRIBE TABLE EXTENDED ... AS JSON). La tabla de hechos
     * es la propia Metric View: cada campo es una columna y cada medida se
     * pide con MEASURE(). Los campos que vienen de un mismo join se agrupan
     * en una dimensión cuando la Metric View expone su clave.
     */
    function buildExternalModel(desc, catalog, schema, view) {
        const warnings = [];
        const cols = Array.isArray(desc.columns) ? desc.columns : [];
        const normType = (t) => (window.DBX && DBX.normalizeType) ? DBX.normalizeType(String((t && t.name) || "")) : String((t && t.name) || "").toUpperCase();
        const ok = (name, what) => {
            if (IDENT_RE.test(String(name))) return true;
            warnings.push(`${what} «${String(name).slice(0, 60)}»: Draco no admite ese nombre (espacios o caracteres especiales); se omite.`);
            return false;
        };

        // Del YAML (sin comentarios): de qué join viene cada campo y la clave de cada join
        const fieldJoin = {}, joinKey = {};
        try {
            const y = (window.jsyaml && desc.view_text) ? window.jsyaml.load(String(desc.view_text)) : null;
            ((y && y.joins) || []).forEach(j => {
                const on = String(j.on || j["on"] || j[true] || "");
                const m = /source\.(\w+)\s*=\s*(\w+)\.(\w+)/.exec(on) || /(\w+)\.(\w+)\s*=\s*source\.(\w+)/.exec(on);
                if (m && j.name) joinKey[j.name] = m[0].startsWith("source.") ? m[3] : m[2];
            });
            ((y && (y.dimensions || y.fields)) || []).forEach(f => {
                const m = /^\s*(\w+)\.(\w+)\s*$/.exec(String(f.expr || ""));
                if (m && m[1] !== "source" && f.name) fieldJoin[f.name] = { join: m[1], col: m[2] };
            });
        } catch (e) { warnings.push("No se ha podido leer el YAML de la Metric View: cada campo será una dimensión suelta."); }

        const dimCols = cols.filter(c => c && !c.is_measure && ok(c.name, "Campo"));
        const attrOf = (c, isKey) => ({ name: c.name, alias: (c.metadata && c.metadata.display_name) || c.name,
            dataType: normType(c.type), isKey: !!isKey, enabled: true });

        const dims = [], used = new Set();
        // 1) Una dimensión por join con su clave expuesta
        Object.keys(joinKey).forEach(jn => {
            const fields = dimCols.filter(c => fieldJoin[c.name] && fieldJoin[c.name].join === jn);
            const key = fields.find(c => fieldJoin[c.name].col === joinKey[jn]);
            if (!key) return;
            fields.forEach(c => used.add(c.name));
            dims.push({ enabled: true, type: "DIMENSION", name: key.name, relProject: "", relDataset: "", relTable: "",
                attributes: [attrOf(key, true), ...fields.filter(c => c !== key).map(c => attrOf(c, false))],
                hierarchies: [], textAttribute: "", calcMembers: {} });
        });
        // 2) El resto, cada campo su dimensión
        dimCols.filter(c => !used.has(c.name)).forEach(c => dims.push({ enabled: true, type: "DIMENSION", name: c.name,
            relProject: "", relDataset: "", relTable: "", attributes: [attrOf(c, true)], hierarchies: [], textAttribute: "", calcMembers: {} }));
        // Mismo orden que en la Metric View
        const pos = new Map(cols.map((c, i) => [c.name, i]));
        dims.sort((a, b) => pos.get(a.attributes[0].name) - pos.get(b.attributes[0].name));

        const measures = cols.filter(c => c && c.is_measure && ok(c.name, "Medida")).map(c => ({
            enabled: true, type: "MEASURE", name: c.name, alias: (c.metadata && c.metadata.display_name) || c.name,
            dataType: normType(c.type), aggregation: "MEASURE", format: "Auto", field: c.name }));
        if (!measures.length) warnings.push("La Metric View no tiene medidas que Draco pueda usar.");

        const model = sanitizeModel({ fact: { project: catalog, dataset: schema, table: view },
            fields: [...dims, ...measures], metricView: { catalog, schema, view } });
        return { model, warnings };
    }

    /* ---------------- lectura desde Databricks ---------------- */

    /**
     * Lee el modelo semántico de la Metric View catalog.schema.view:
     *  - creada por Draco (propiedad draco.model): el modelo completo, y los
     *    informes consultan las tablas como siempre;
     *  - ajena: se monta desde su descripción y los informes consultan la
     *    Metric View con MEASURE().
     * Devuelve { model, modelName, kind: "draco"|"external", warnings }.
     */
    async function readModel(catalog, schema, view) {
        ident(catalog, "catálogo", CATALOG_RE); ident(schema, "esquema"); ident(view, "vista");
        if (!window.DBX) throw new Error("Draco no tiene cargado el conector de Databricks.");
        const desc = await DBX.describeJson(catalog, schema, view);
        if (String(desc.type || "").toUpperCase() !== "METRIC_VIEW") throw new Error(`«${view}» no es una Metric View.`);
        const props = (desc.table_properties && typeof desc.table_properties === "object") ? desc.table_properties : {};
        if (props["draco.generator"] === GENERATOR && props["draco.model"]) {
            const raw = await decode(props["draco.model.format"] || "", props["draco.model"]);
            return { model: sanitizeModel(raw), modelName: str(props["draco.model.name"] || "", 255), kind: "draco", warnings: [] };
        }
        const r = buildExternalModel(desc, catalog, schema, view);
        return { model: r.model, modelName: "", kind: "external", warnings: r.warnings };
    }

    /* ---------------- guardar: crear la Metric View ---------------- */

    function viewNameFor(modelName) {
        let n = String(modelName || "modelo").toLowerCase().replace(/[^a-z0-9_]+/g, "_").replace(/^_+|_+$/g, "");
        if (!/^[a-z_]/.test(n)) n = "m_" + n;
        return (n || "modelo").slice(0, 240) + "_metrics";
    }

    const lit = (v) => "'" + String(v).replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'";

    async function run(sql, kind) {
        if (window.DracoQueryJournal) return window.DracoQueryJournal.track(sql, () => DBX.execute(sql), { kind: kind || "Metric View", engine: "Databricks" });
        return DBX.execute(sql);
    }

    /**
     * Crea o actualiza la Metric View del modelo (catalog.schema.<modelo>_metrics,
     * en el catálogo y esquema de la tabla de hechos) y guarda el modelo de
     * Draco en sus propiedades. No sobrescribe nada que no haya creado Draco.
     * Devuelve el nombre completo de la vista.
     */
    async function publish(modelName, model) {
        if (!window.DBX || !window.MetricViewExport) throw new Error("Draco no tiene cargado lo necesario para crear la Metric View.");
        const clean = sanitizeModel(model);   // lo mismo que luego se leerá
        if (clean.metricView) throw new Error("Este modelo está montado sobre una Metric View: no se crea otra.");
        const catalog = clean.fact.project, schema = clean.fact.dataset;
        if (!catalog || !schema) throw new Error("La tabla de hechos no tiene catálogo y esquema.");
        const view = viewNameFor(modelName);
        const ref = [catalog, schema, view].map(x => DBX.quoteIdent(x)).join(".");

        // ¿Ya existe algo con ese nombre que no sea de Draco? Entonces no se toca.
        let existing = null;
        try { existing = await DBX.describeJson(catalog, schema, view); }
        catch (e) {
            if (!/TABLE_OR_VIEW_NOT_FOUND|not found|no existe|does not exist/i.test(String(e && e.message))) throw e;
        }
        if (existing) {
            const props = existing.table_properties || {};
            if (String(existing.type || "").toUpperCase() !== "METRIC_VIEW" || props["draco.generator"] !== GENERATOR) {
                throw new Error(`Ya existe ${catalog}.${schema}.${view} y no la creó Draco: no se sobrescribe.`);
            }
        }

        // YAML sin comentarios (Databricks no los guarda y así no puede aparecer «$$»)
        const yaml = window.MetricViewExport.buildContent(modelName, clean).content
            .split("\n").filter(l => !/^\s*#/.test(l)).join("\n").trim();
        if (yaml.indexOf("$$") !== -1) throw new Error("El modelo contiene «$$» y no se puede crear la Metric View.");
        await run("CREATE OR REPLACE VIEW " + ref + "\nWITH METRICS\nLANGUAGE YAML\nAS $$\n" + yaml + "\n$$");

        const enc = await encode(clean);
        await run("ALTER VIEW " + ref + " SET TBLPROPERTIES (\n"
            + "  'draco.generator' = " + lit(GENERATOR) + ",\n"
            + "  'draco.model.format' = " + lit(enc.format) + ",\n"
            + "  'draco.model.name' = " + lit(String(modelName).slice(0, 255)) + ",\n"
            + "  'draco.model' = " + lit(enc.value) + "\n)");
        return catalog + "." + schema + "." + view;
    }

    window.DracoMetricView = { readModel, publish, viewNameFor, buildExternalModel, sanitizeModel, decode, encode, GENERATOR, FORMATS };
})();
