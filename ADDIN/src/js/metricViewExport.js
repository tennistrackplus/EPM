/**
 * MetricViewExport
 * ------------------------------------------------------------------------
 * Genera la Metric View (YAML) de Databricks de un modelo semántico de
 * Draco, a partir de su objeto { fact, fields } (el mismo que guarda
 * SemanticModelStore y que produce LkmlImport). Es el equivalente
 * Databricks de LkmlExport y la función inversa de MetricViewImport.
 *
 *   model.fact   = { project: catálogo, dataset: esquema, table }
 *   model.fields = dimensiones y medidas (ver js/lkmlExport.js)
 *
 * El YAML es una Metric View válida (especificación 1.1): se puede crear
 * tal cual en Databricks con
 *   CREATE OR REPLACE VIEW catalogo.esquema.nombre WITH METRICS LANGUAGE YAML AS $$ ... $$
 * y la usan Databricks SQL, los dashboards y Genie. Lo que la Metric View no
 * sabe representar (jerarquías, texto de la clave, tipos, definición exacta
 * de las medidas de Draco) va en comentarios «# draco-…» / «# Jerarquía …»,
 * como en el .lkml: Databricks los ignora y MetricViewImport los recupera.
 *
 * Correspondencia:
 *   tabla de hechos            -> source
 *   dimensión con tabla propia -> join (name = tabla; 'on' = FK de hechos = clave)
 *                                 + un campo por atributo (expr: join.columna)
 *   dimensión de hechos        -> campo (expr: source.columna)
 *   medida simple              -> SUM(col) / COUNT(col) / AVG / MIN / MAX / COUNT(DISTINCT col)
 *   medida filtrada            -> AGG(base) FILTER (WHERE col IN (...))
 *   medida calculada           -> MEASURE(a) - MEASURE(b), (x) / NULLIF(y, 0)…
 *
 * Función PURA (no toca Excel ni Office).
 */
(function () {
    "use strict";

    const AGG = { SUM: "SUM", COUNT: "COUNT", AVG: "AVG", MIN: "MIN", MAX: "MAX", COUNT_DISTINCT: "COUNT_DISTINCT" };
    const SIMPLE_ID = /^[A-Za-z_][A-Za-z0-9_]*$/;

    /** Cadena YAML entre comillas dobles (sintaxis JSON, válida en YAML). */
    const q = (s) => JSON.stringify(String(s == null ? "" : s));
    /** Identificador SQL de Databricks (con comillas invertidas si hace falta). */
    const id = (s) => SIMPLE_ID.test(String(s)) ? String(s) : "`" + String(s).replace(/`/g, "``") + "`";
    const tableRef = (c, s, t) => [c, s, t].filter(Boolean).map(id).join(".");
    /** Literal de texto de Databricks (escapa con barra invertida). */
    const lit = (v) => "'" + String(v).replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'";
    const b64 = (obj) => (typeof btoa === "function"
        ? btoa(unescape(encodeURIComponent(JSON.stringify(obj))))
        : Buffer.from(JSON.stringify(obj), "utf8").toString("base64"));
    const isNumericType = (t) => /^(INT|INTEGER|INT64|BIGINT|SMALLINT|TINYINT|FLOAT|FLOAT64|DOUBLE|DECIMAL|NUMERIC|BIGNUMERIC|LONG)/i.test(String(t || ""));

    function aggExpr(aggregation, column) {
        const a = AGG[String(aggregation || "SUM").toUpperCase()] || "SUM";
        const c = id(column);
        return a === "COUNT_DISTINCT" ? `COUNT(DISTINCT ${c})` : `${a}(${c})`;
    }

    /**
     * @param {string} modelName
     * @param {{fact, fields}} model
     * @param {{catalog?: string, schema?: string}} [opts] sustituyen a
     *        project/dataset de TODAS las tablas (p. ej. al llevar a
     *        Databricks un modelo que venía de BigQuery).
     * @returns {{ content: string, warnings: string[] }}
     */
    function buildContent(modelName, model, opts) {
        const o = opts || {};
        const fact = (model && model.fact) || {};
        const fields = (model && model.fields) || [];
        const warnings = [];
        const cat = (p) => o.catalog || p || "";
        const sch = (d) => o.schema || d || "";

        const dims = fields.filter(f => f && f.enabled !== false && f.type === "DIMENSION");
        const measures = fields.filter(f => f && f.enabled !== false && f.type === "MEASURE");
        const factTable = tableRef(cat(fact.project), sch(fact.dataset), fact.table);
        if (!fact.table) throw new Error("El modelo no tiene tabla de hechos.");

        // Nombres únicos en toda la Metric View (campos y medidas, sin mayúsculas)
        const used = new Set();
        const unique = (base) => {
            let n = String(base).replace(/[^A-Za-z0-9_]+/g, "_").replace(/^_+|_+$/g, "") || "campo";
            if (!/^[A-Za-z_]/.test(n)) n = "c_" + n;
            let k = n, i = 2;
            while (used.has(k.toLowerCase())) k = n + "_" + (i++);
            used.add(k.toLowerCase());
            return k;
        };
        // Las medidas conservan su nombre de Draco (MEASURE(nombre) en las calculadas)
        measures.forEach(m => used.add(String(m.name).toLowerCase()));

        // Nombre de join por dimensión: la tabla de la dimensión si es única, si no el nombre de la dimensión
        const tableCount = {};
        dims.filter(d => d.relTable).forEach(d => { tableCount[d.relTable] = (tableCount[d.relTable] || 0) + 1; });
        const joinNames = new Set();
        const joinOf = (d) => {
            let n = tableCount[d.relTable] === 1 ? d.relTable : d.name;
            n = String(n).replace(/[^A-Za-z0-9_]+/g, "_");
            let k = n, i = 2;
            while (joinNames.has(k.toLowerCase()) || k.toLowerCase() === "source") k = n + "_" + (i++);
            joinNames.add(k.toLowerCase());
            return k;
        };

        const L = [];
        L.push(`# Modelo semántico: ${modelName}`);
        L.push("# Metric View de Databricks generada por Draco. Las líneas «# draco-…» y «# Jerarquía …»");
        L.push("# son metadatos de Draco: Databricks las ignora y Draco las usa al volver a abrir el fichero.");
        L.push("#");
        L.push("# Para crearla en Databricks (SQL Editor):");
        L.push(`#   CREATE OR REPLACE VIEW ${tableRef(cat(fact.project), sch(fact.dataset), String(modelName || "modelo").replace(/[^A-Za-z0-9_]+/g, "_") + "_metrics")}`);
        L.push("#   WITH METRICS LANGUAGE YAML AS $$ <contenido de este fichero> $$");
        if (model && model.dracoPlanning === true) L.push("# draco: planning");
        L.push("");
        L.push("version: 1.1");
        L.push(`comment: ${q("Modelo semántico " + modelName + " (Draco)")}`);
        L.push(`source: ${factTable}`);
        L.push("");

        const dimMeta = (d, indent) => {
            const p = " ".repeat(indent) + "# ";
            const out = [];
            out.push(p + `draco-dimension: ${d.name}`);
            if (d.textAttribute) out.push(p + `texto de la clave: ${d.textAttribute}`);
            if (d.calcMembers && Object.keys(d.calcMembers).length) out.push(p + `draco-calculos: ${b64(d.calcMembers)}`);
            (d.hierarchies || []).forEach(h => {
                if (!h.levels || !h.levels.length) return;
                const ev = h.emptyValue ? ` (vacío: "${String(h.emptyValue).replace(/"/g, "'")}")` : "";
                out.push(p + `Jerarquía "${h.name}"${ev}:`);
                h.levels.forEach((l, i) => out.push(p + `  ${i + 1}. ${l.attribute}${l.textAttribute ? " | texto: " + l.textAttribute : ""}`));
            });
            return out;
        };

        /* ---------- joins ---------- */
        const joined = dims.filter(d => d.relTable);
        const joinByDim = new Map();
        if (joined.length) {
            L.push("joins:");
            joined.forEach(d => {
                const jn = joinOf(d);
                joinByDim.set(d, jn);
                const attrs = (d.attributes || []).filter(a => a.enabled !== false);
                const key = attrs.find(a => a.isKey) || { name: d.name };
                L.push(`  - name: ${jn}`);
                L.push(`    source: ${tableRef(cat(d.relProject || fact.project), sch(d.relDataset || fact.dataset), d.relTable)}`);
                L.push(`    'on': ${q(`source.${id(d.name)} = ${jn}.${id(key.name)}`)}`);
                dimMeta(d, 4).forEach(x => L.push(x));
            });
            L.push("");
        }

        /* ---------- campos (atributos) ---------- */
        L.push("dimensions:");
        dims.forEach(d => {
            const jn = joinByDim.get(d) || null;
            let attrs = (d.attributes || []).filter(a => a.enabled !== false);
            if (!attrs.length) attrs = [{ name: d.name, isKey: true }];
            attrs.forEach((a, i) => {
                const fname = unique(jn ? jn + "_" + a.name : a.name);
                L.push(`  - name: ${fname}`);
                if (a.calc) {
                    // Atributo calculado de Draco: su fórmula no es SQL; Databricks
                    // ve una columna vacía y Draco recupera la fórmula al importar.
                    L.push(`    expr: ${q("CAST(NULL AS STRING)")}`);
                    L.push(`    # draco-atributo: ${b64({ name: a.name, expression: a.expression || "" })}`);
                    warnings.push(`Atributo calculado «${d.name}.${a.name}»: en Databricks aparece vacío (su fórmula es de Draco).`);
                } else {
                    L.push(`    expr: ${jn ? jn + "." + id(a.name) : "source." + id(a.name)}`);
                }
                if (a.alias && a.alias !== a.name) L.push(`    display_name: ${q(a.alias)}`);
                if (!jn && i === 0) dimMeta(d, 4).forEach(x => L.push(x));   // dimensión de hechos: metadatos en su campo
                if (a.dataType) L.push(`    # draco-tipo: ${a.dataType}`);
            });
        });
        L.push("");

        /* ---------- medidas (simples, filtradas y, al final, calculadas) ---------- */
        const byName = new Map(measures.map(m => [String(m.name).toUpperCase(), m]));
        const colOf = (m) => m.name;   // medida simple: su columna es su nombre
        const condSql = (c) => {
            const vals = (c.values || []).map(v => isNumericType(c.dataType) && /^-?\d+(\.\d+)?$/.test(String(v)) ? String(v) : lit(v));
            if (!vals.length) return null;
            return `${id(c.column)} ${c.op === "NOT_IN" ? "NOT IN" : "IN"} (${vals.join(", ")})`;
        };
        const ordered = [
            ...measures.filter(m => m.measureKind !== "FILTERED" && m.measureKind !== "CALC"),
            ...measures.filter(m => m.measureKind === "FILTERED")
        ];
        // Calculadas en orden de dependencias (Databricks exige definir antes las referenciadas)
        const calcs = measures.filter(m => m.measureKind === "CALC");
        const refsOf = (m) => (window.DracoMeasureFormula ? window.DracoMeasureFormula.refs(m.formula || "") : []).map(r => String(r).toUpperCase());
        const placed = new Set(ordered.map(m => String(m.name).toUpperCase()));
        let guard = calcs.length + 1;
        while (calcs.length && guard-- > 0) {
            for (let i = 0; i < calcs.length; i++) {
                if (refsOf(calcs[i]).every(r => placed.has(r) || !byName.has(r))) {
                    placed.add(String(calcs[i].name).toUpperCase());
                    ordered.push(calcs.splice(i, 1)[0]);
                    i--;
                }
            }
        }
        calcs.forEach(m => ordered.push(m));   // ciclo: se dejan al final (fallarán en Databricks)

        L.push("measures:");
        ordered.forEach(m => {
            let expr;
            if (m.measureKind === "FILTERED") {
                const base = byName.get(String(m.base || "").toUpperCase());
                const col = base ? colOf(base) : m.base;
                const where = (m.conditions || []).map(condSql).filter(Boolean);
                expr = aggExpr(m.aggregation, col) + (where.length ? ` FILTER (WHERE ${where.join(" AND ")})` : "");
            } else if (m.measureKind === "CALC") {
                try {
                    if (!window.DracoMeasureFormula) throw new Error("sin traductor de fórmulas");
                    expr = window.DracoMeasureFormula.toSql(m.formula || "", (ref) => {
                        const r = byName.get(String(ref).toUpperCase());
                        return `MEASURE(${id(r ? r.name : ref)})`;
                    });
                    // La división de Draco ya sale como «x / NULLIF(y, 0)», válida en Databricks
                } catch (e) {
                    expr = "SUM(CAST(NULL AS DOUBLE))";
                    warnings.push(`Medida calculada «${m.name}»: no se pudo traducir su fórmula (${e.message}); en Databricks aparece vacía.`);
                }
            } else {
                expr = aggExpr(m.aggregation, colOf(m));
            }
            L.push(`  - name: ${m.name}`);
            L.push(`    expr: ${q(expr)}`);
            if (m.alias && m.alias !== m.name) L.push(`    display_name: ${q(m.alias)}`);
            if (m.measureKind === "FILTERED" || m.measureKind === "CALC") {
                L.push(`    # draco-medida: ${b64({ name: m.name, alias: m.alias, measureKind: m.measureKind, formula: m.formula || "",
                    base: m.base || "", conditions: m.conditions || [], aggregation: m.aggregation, format: m.format })}`);
            } else if (m.format && m.format !== "Auto") {
                L.push(`    # draco-formato: ${m.format}`);
            }
        });

        return { content: L.join("\n") + "\n", warnings };
    }

    window.MetricViewExport = { buildContent };
})();
